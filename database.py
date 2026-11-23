import ast
import json
import os
import re
import shutil
import sqlite3
import threading
import zipfile


def _parse_date_tuple(date_str):
    """Парсит дату ДД.ММ.ГГГГ (или ГГГГ-ММ-ДД / ДД/ММ/ГГГГ) в кортеж (год, месяц, день) или None."""
    s = str(date_str or "").strip()
    if not s:
        return None
    try:
        if "." in s:
            parts = [p.strip() for p in s.split(".")]
            if len(parts) == 3:
                d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
                if y < 100:
                    y += 2000
                return (y, m, d)
        elif "-" in s:
            parts = [p.strip() for p in s.split("-")]
            if len(parts) == 3:
                y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
                if y < 100:
                    y += 2000
                return (y, m, d)
        elif "/" in s:
            parts = [p.strip() for p in s.split("/")]
            if len(parts) == 3:
                d, m, y = int(parts[0]), int(parts[1]), int(parts[2])
                if y < 100:
                    y += 2000
                return (y, m, d)
    except Exception:
        pass
    return None


def _parse_time_mins(time_str):
    """Парсит время ЧЧ:ММ в количество минут от полуночи для сортировки в пределах одной даты."""
    s = str(time_str or "").strip()
    if not s:
        return 0
    try:
        m = re.search(r"(\d{1,2})\s*[:.]\s*(\d{1,2})", s)
        if m:
            return int(m.group(1)) * 60 + int(m.group(2))
    except Exception:
        pass
    return 0


def _wb_chrono_key(date_str, wb_num, row_id, time_out=None):
    """Хронологический ключ сортировки путевых листов: (год, месяц, день, время выезда, числовой № ПЛ, id)."""
    dt_t = _parse_date_tuple(date_str)
    y, m, d = dt_t if dt_t else (0, 0, 0)
    t_mins = _parse_time_mins(time_out) if time_out is not None else 0
    nums = re.findall(r"\d+", str(wb_num or ""))
    wb_int = int(nums[-1]) if nums else 0
    return (y, m, d, t_mins, wb_int, int(row_id or 0))


def _configure_sqlite_conn(conn):
    """Оптимизирует соединение SQLite для высокой скорости работы на Windows NTFS."""
    try:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute("PRAGMA cache_size=-8000")
        conn.execute("PRAGMA temp_store=MEMORY")
        conn.execute("PRAGMA busy_timeout=5000")
    except Exception:
        pass
    return conn


def _checkpoint_db_file(db_path):
    """Сбрасывает WAL-журнал обратно в основной .db файл перед резервным копированием."""
    if not db_path or not os.path.exists(db_path):
        return
    try:
        conn = sqlite3.connect(db_path, timeout=5.0)
        try:
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
            conn.commit()
        finally:
            conn.close()
    except Exception:
        pass


class _BaseSQLiteManager:
    """Общая высокопроизводительная база с пулом соединений по потокам и пакетными транзакциями."""

    def __init__(self, db_name):
        self.db_name = db_name
        self._local = threading.local()
        self._all_conns = []
        self._conn_lock = threading.Lock()

    def get_conn(self):
        conn = getattr(self._local, "conn", None)
        if conn is not None:
            return conn
        conn = _configure_sqlite_conn(sqlite3.connect(self.db_name, timeout=5.0))
        self._local.conn = conn
        with self._conn_lock:
            self._all_conns.append(conn)
        return conn

    def close_connections(self):
        """Безопасно закрывает открытые соединения и сбрасывает WAL (перед бэкапом/восстановлением)."""
        with self._conn_lock:
            conns = list(self._all_conns)
            self._all_conns.clear()
        for c in conns:
            try:
                c.execute("PRAGMA wal_checkpoint(TRUNCATE)")
            except Exception:
                pass
            try:
                c.close()
            except Exception:
                pass
        if hasattr(self._local, "conn"):
            self._local.conn = None

    def checkpoint(self):
        try:
            conn = self.get_conn()
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        except Exception:
            _checkpoint_db_file(self.db_name)

    def query(self, sql, params=()):
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute(sql, params)
        # Автоматический commit только для изменяющих запросов (SELECT без лишней записи)
        sql_head = sql.lstrip()[:7].upper()
        if not sql_head.startswith(("SELECT", "PRAGMA")):
            conn.commit()
        return cur.fetchall()

    def query_many(self, sql, seq_of_params):
        if not seq_of_params:
            return
        conn = self.get_conn()
        cur = conn.cursor()
        cur.executemany(sql, seq_of_params)
        conn.commit()

    def get_list(self, table):
        return self.query(f"SELECT * FROM {table}")

    def get_cached_route(self, cache_key):
        try:
            rows = self.query(
                "SELECT dist, coords FROM route_cache WHERE key=?", (cache_key,)
            )
            if rows:
                dist_val, coords_json = rows[0]
                coords = json.loads(coords_json) if coords_json else None
                if coords:
                    coords = [tuple(pt) for pt in coords]
                return (dist_val, coords)
        except Exception:
            pass
        return None

    def save_cached_route(self, cache_key, dist_val, coords):
        try:
            coords_json = json.dumps(coords, separators=(",", ":")) if coords else "[]"
            self.query(
                "INSERT OR REPLACE INTO route_cache (key, dist, coords) VALUES (?, ?, ?)",
                (cache_key, float(dist_val) if dist_val is not None else None, coords_json),
            )
        except Exception:
            pass

    def renumber_history_by_date_detailed(self, renumber_wb_num=True):
        """
        Сортирует записи в таблице history строго по дате составления (хронологически),
        перенумеровывает ID (1, 2, 3, ...) и (при renumber_wb_num=True) автоматически
        перенумеровывает № ПЛ (wb_num = "1", "2", "3", ...).
        Возвращает кортеж (id_map, renamed_records), где:
          - id_map: {старый_id: новый_id}
          - renamed_records: список словарей {old_id, new_id, old_wb, new_wb, date, veh}
            для записей, у которых изменился № ПЛ.
        """
        conn = self.get_conn()
        cur = conn.cursor()
        rows = cur.execute(
            "SELECT id, wb_num, date, time_out, veh FROM history"
        ).fetchall()
        if not rows:
            return {}, []

        rows_sorted = sorted(
            rows, key=lambda r: _wb_chrono_key(r[2], r[1], r[0], r[3])
        )
        id_map = {}
        new_wb_by_old_id = {}
        renamed_records = []
        needs_id_update = False
        needs_wb_update = False

        for idx, r in enumerate(rows_sorted, start=1):
            old_id, old_wb, d_str, _, veh = r[0], str(r[1] or "").strip(), str(r[2] or ""), r[3], str(r[4] or "")
            new_id = idx
            new_wb = str(idx) if renumber_wb_num else old_wb
            id_map[old_id] = new_id
            new_wb_by_old_id[old_id] = new_wb
            if old_id != new_id:
                needs_id_update = True
            if old_wb != new_wb:
                needs_wb_update = True
                renamed_records.append(
                    {
                        "old_id": old_id,
                        "new_id": new_id,
                        "old_wb": old_wb,
                        "new_wb": new_wb,
                        "date": d_str,
                        "veh": veh,
                    }
                )

        if not needs_id_update and not needs_wb_update:
            return id_map, []

        try:
            if needs_id_update:
                max_id = max(r[0] for r in rows)
                offset = max_id + len(rows) + 1000
                cur.execute("UPDATE history SET id = id + ?", (offset,))
                for old_id, new_id in id_map.items():
                    if renumber_wb_num:
                        cur.execute(
                            "UPDATE history SET id = ?, wb_num = ? WHERE id = ?",
                            (new_id, new_wb_by_old_id[old_id], old_id + offset),
                        )
                    else:
                        cur.execute(
                            "UPDATE history SET id = ? WHERE id = ?",
                            (new_id, old_id + offset),
                        )
            elif needs_wb_update:
                for old_id, new_wb in new_wb_by_old_id.items():
                    cur.execute(
                        "UPDATE history SET wb_num = ? WHERE id = ?",
                        (new_wb, old_id),
                    )
            try:
                cur.execute(
                    "UPDATE sqlite_sequence SET seq = ? WHERE name = 'history'",
                    (len(rows),),
                )
            except Exception:
                pass
            conn.commit()
        except Exception:
            try:
                conn.rollback()
            except Exception:
                pass
            return {r[0]: r[0] for r in rows}, []
        return id_map, renamed_records

    def renumber_history_ids_by_date(self, renumber_wb_num=True):
        """
        Перенумеровывает ID и (по умолчанию) № ПЛ в таблице history по порядку (1, 2, 3, ...)
        строго по хронологии даты составления путевого листа.
        Возвращает словарь {старый_id: новый_id}.
        """
        id_map, _ = self.renumber_history_by_date_detailed(
            renumber_wb_num=renumber_wb_num
        )
        return id_map


class DatabaseManager(_BaseSQLiteManager):
    """База данных для легковых автомобилей (waybills.db) — полностью изолирована от тракторов."""

    def __init__(self, db_name="waybills.db"):
        super().__init__(db_name)
        self.init_db()

    def init_db(self, auto_sync_history=True):
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS orgs (id INTEGER PRIMARY KEY, name TEXT UNIQUE)")
        cur.execute("CREATE TABLE IF NOT EXISTS clients (id INTEGER PRIMARY KEY, name TEXT UNIQUE)")
        cur.execute("CREATE TABLE IF NOT EXISTS staff (id INTEGER PRIMARY KEY, fio TEXT, role TEXT, UNIQUE(fio, role))")
        self._migrate_staff_table(cur)
        cur.execute(
            """CREATE TABLE IF NOT EXISTS vehicles (
                id INTEGER PRIMARY KEY, mark TEXT UNIQUE, reg TEXT, gar TEXT,
                f_brand TEXT, norm REAL, f_brand_gas TEXT, norm_gas REAL
            )"""
        )
        cur.execute(
            """CREATE TABLE IF NOT EXISTS drivers (
                id INTEGER PRIMARY KEY, fio TEXT UNIQUE, cat TEXT, exp TEXT, tab TEXT, ser TEXT, num TEXT
            )"""
        )
        cur.execute(
            """CREATE TABLE IF NOT EXISTS locs (
                id INTEGER PRIMARY KEY, name TEXT UNIQUE, lat REAL, lon REAL
            )"""
        )
        cur.execute(
            """CREATE TABLE IF NOT EXISTS route_cache (
                key TEXT PRIMARY KEY, dist REAL, coords TEXT
            )"""
        )

        cur.execute("PRAGMA table_info(history)")
        cols = [info[1] for info in cur.fetchall()]
        if cols and "g_s" not in cols:
            cur.execute("ALTER TABLE history RENAME TO history_old")
            cur.execute(
                """CREATE TABLE history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, wb_num TEXT, date TEXT,
                    time_out TEXT, time_in TEXT, org TEXT, client TEXT, veh TEXT, drv TEXT,
                    odo_s REAL, odo_e REAL,
                    f_s REAL, f_a REAL, f_e REAL, f_norm REAL,
                    g_s REAL, g_a REAL, g_e REAL, g_norm REAL,
                    route TEXT, dist REAL
                )"""
            )
            cur.execute(
                """INSERT INTO history (
                    id, wb_num, date, time_out, time_in, org, client, veh, drv,
                    odo_s, odo_e, f_s, f_a, f_e, f_norm, g_s, g_a, g_e, g_norm, route, dist
                )
                SELECT id, wb_num, date, '08:00', '17:00', org, client, veh, drv,
                       odo_s, odo_e, f_s, f_a, f_e, f_norm, 0, 0, 0, 0, route, dist
                FROM history_old"""
            )
            cur.execute("DROP TABLE history_old")
        else:
            cur.execute(
                """CREATE TABLE IF NOT EXISTS history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT, wb_num TEXT, date TEXT,
                    time_out TEXT, time_in TEXT, org TEXT, client TEXT, veh TEXT, drv TEXT,
                    odo_s REAL, odo_e REAL,
                    f_s REAL, f_a REAL, f_e REAL, f_norm REAL,
                    g_s REAL, g_a REAL, g_e REAL, g_norm REAL,
                    route TEXT, dist REAL
                )"""
            )
        cur.execute("CREATE INDEX IF NOT EXISTS idx_history_veh ON history(veh)")
        conn.commit()
        if auto_sync_history:
            self.renumber_history_ids_by_date(renumber_wb_num=True)
            self.recalculate_all_vehicles()

    def _migrate_staff_table(self, cur):
        """Снимает старое ограничение fio TEXT UNIQUE, чтобы один сотрудник мог быть и Механиком, и Диспетчером."""
        cur.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='staff'")
        row = cur.fetchone()
        if row and row[0] and "fio TEXT UNIQUE" in row[0]:
            cur.execute("ALTER TABLE staff RENAME TO staff_old")
            cur.execute(
                "CREATE TABLE staff (id INTEGER PRIMARY KEY, fio TEXT, role TEXT, UNIQUE(fio, role))"
            )
            cur.execute("INSERT OR IGNORE INTO staff (id, fio, role) SELECT id, fio, role FROM staff_old")
            cur.execute("DROP TABLE staff_old")

    def add_driver(self, fio, cat, exp, tab, ser, num):
        self.query(
            "INSERT OR REPLACE INTO drivers (fio, cat, exp, tab, ser, num) VALUES (?, ?, ?, ?, ?, ?)",
            (fio, cat, exp, tab, ser, num),
        )

    def update_driver(self, d_id, fio, cat, exp, tab, ser, num):
        self.query(
            "UPDATE drivers SET fio=?, cat=?, exp=?, tab=?, ser=?, num=? WHERE id=?",
            (fio, cat, exp, tab, ser, num, d_id),
        )

    def update_vehicle(self, v_id, mark, reg, gar, fb, norm, fbg, normg):
        self.query(
            "UPDATE vehicles SET mark=?, reg=?, gar=?, f_brand=?, norm=?, f_brand_gas=?, norm_gas=? WHERE id=?",
            (mark, reg, gar, fb, norm, fbg, normg, v_id),
        )

    def update_org(self, o_id, name):
        try:
            self.query("UPDATE orgs SET name=? WHERE id=?", (name, o_id))
            return True
        except sqlite3.IntegrityError:
            return False

    def update_client(self, c_id, name):
        try:
            self.query("UPDATE clients SET name=? WHERE id=?", (name, c_id))
            return True
        except sqlite3.IntegrityError:
            return False

    def add_staff(self, fio, role):
        self.query(
            "INSERT OR IGNORE INTO staff (fio, role) VALUES (?, ?)",
            (fio, role),
        )

    def update_staff(self, s_id, fio, role):
        try:
            self.query(
                "UPDATE staff SET fio=?, role=? WHERE id=?",
                (fio, role, s_id),
            )
            return True
        except sqlite3.IntegrityError:
            return False

    def update_wb(
        self,
        w_id,
        wb_num,
        date,
        t_out,
        t_in,
        org,
        client,
        veh,
        drv,
        os_val,
        oe_val,
        fs,
        fa,
        fe,
        fn,
        gs,
        ga,
        ge,
        gn,
        route,
        dist,
    ):
        self.query(
            """UPDATE history SET
                wb_num=?, date=?, time_out=?, time_in=?, org=?, client=?, veh=?, drv=?,
                odo_s=?, odo_e=?, f_s=?, f_a=?, f_e=?, f_norm=?,
                g_s=?, g_a=?, g_e=?, g_norm=?, route=?, dist=?
               WHERE id=?""",
            (
                wb_num,
                date,
                t_out,
                t_in,
                org,
                client,
                veh,
                drv,
                os_val,
                oe_val,
                fs,
                fa,
                fe,
                fn,
                gs,
                ga,
                ge,
                gn,
                route,
                dist,
                w_id,
            ),
        )

    def save_wb(
        self,
        wb_num,
        date,
        t_out,
        t_in,
        org,
        client,
        veh,
        drv,
        os_val,
        oe_val,
        fs,
        fa,
        fe,
        fn,
        gs,
        ga,
        ge,
        gn,
        route,
        dist,
    ):
        self.query(
            """INSERT INTO history (
                wb_num, date, time_out, time_in, org, client, veh, drv,
                odo_s, odo_e, f_s, f_a, f_e, f_norm, g_s, g_a, g_e, g_norm, route, dist
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                wb_num,
                date,
                t_out,
                t_in,
                org,
                client,
                veh,
                drv,
                os_val,
                oe_val,
                fs,
                fa,
                fe,
                fn,
                gs,
                ga,
                ge,
                gn,
                route,
                dist,
            ),
        )

    def get_last_veh(self, veh):
        if not veh:
            return None
        rows = self.query(
            "SELECT id, wb_num, date, odo_e, f_e, g_e, time_out FROM history WHERE veh=?",
            (veh,),
        )
        if not rows:
            return None
        rows.sort(key=lambda r: _wb_chrono_key(r[2], r[1], r[0], r[6]))
        last = rows[-1]
        return (last[3], last[4], last[5])

    def get_record_by_id(self, r_id):
        rows = self.query("SELECT * FROM history WHERE id=?", (r_id,))
        return rows[0] if rows else None

    def find_by_veh_and_date(self, veh, date_str, wb_num=None):
        """Находит существующий путевой лист для данного ТС на указанную дату (приоритет совпадению № ПЛ)."""
        if not veh or not date_str:
            return None
        rows = self.query(
            "SELECT * FROM history WHERE veh=? AND date=? ORDER BY id DESC",
            (str(veh).strip(), str(date_str).strip()),
        )
        if not rows:
            return None
        if wb_num is not None:
            wb_s = str(wb_num).strip()
            exact = next((r for r in rows if str(r[1]).strip() == wb_s), None)
            if exact:
                return exact
        return rows[0]

    def recalculate_chain(self, veh):
        """
        Хронологически (от начала года до конца) пересчитывает цепочку одометра и остатков ГСМ
        для автомобиля за одну пакетную транзакцию и возвращает список изменившихся ID.
        """
        if not veh:
            return []
        rows = list(self.query("SELECT * FROM history WHERE veh=?", (veh,)))
        if not rows:
            return []
        rows.sort(key=lambda r: _wb_chrono_key(r[2], r[1], r[0], r[3]))

        updated_ids = []
        updates_batch = []

        # Проверяем и первую запись (исправление арифметики одометра/остатков)
        first = rows[0]
        f_odo_s = float(first[9] or 0.0)
        f_dist = float(first[20] or 0.0)
        f_odo_e = round(f_odo_s + f_dist, 2)
        f_f_s = float(first[11] or 0.0)
        f_f_a = float(first[12] or 0.0)
        f_f_norm = float(first[14] or 0.0)
        f_f_e = round(f_f_s + f_f_a - f_f_norm, 2)
        f_g_s = float(first[15] or 0.0)
        f_g_a = float(first[16] or 0.0)
        f_g_norm = float(first[18] or 0.0)
        f_g_e = round(f_g_s + f_g_a - f_g_norm, 2)
        old_first = (
            float(first[9] or 0.0),
            float(first[10] or 0.0),
            float(first[11] or 0.0),
            float(first[13] or 0.0),
            float(first[15] or 0.0),
            float(first[17] or 0.0),
        )
        new_first = (f_odo_s, f_odo_e, f_f_s, f_f_e, f_g_s, f_g_e)
        if any(abs(a - b) > 1e-4 for a, b in zip(old_first, new_first)):
            updates_batch.append((f_odo_s, f_odo_e, f_f_s, f_f_e, f_g_s, f_g_e, first[0]))
            lst0 = list(rows[0])
            lst0[9], lst0[10] = f_odo_s, f_odo_e
            lst0[11], lst0[13] = f_f_s, f_f_e
            lst0[15], lst0[17] = f_g_s, f_g_e
            rows[0] = tuple(lst0)
            updated_ids.append(first[0])

        for i in range(1, len(rows)):
            prev = rows[i - 1]
            curr = rows[i]
            prev_odo_e = float(prev[10] or 0.0)
            new_odo_s = prev_odo_e if prev_odo_e > 0.001 else float(curr[9] or 0.0)
            new_f_s = float(prev[13] or 0.0)
            new_g_s = float(prev[17] or 0.0)

            dist = float(curr[20] or 0.0)
            f_a, f_norm = float(curr[12] or 0.0), float(curr[14] or 0.0)
            g_a, g_norm = float(curr[16] or 0.0), float(curr[18] or 0.0)

            new_odo_e = round(new_odo_s + dist, 2)
            new_f_e = round(new_f_s + f_a - f_norm, 2)
            new_g_e = round(new_g_s + g_a - g_norm, 2)

            old_vals = (
                float(curr[9] or 0.0),
                float(curr[10] or 0.0),
                float(curr[11] or 0.0),
                float(curr[13] or 0.0),
                float(curr[15] or 0.0),
                float(curr[17] or 0.0),
            )
            new_vals = (new_odo_s, new_odo_e, new_f_s, new_f_e, new_g_s, new_g_e)
            changed = any(abs(a - b) > 1e-4 for a, b in zip(old_vals, new_vals))

            if changed:
                updates_batch.append(
                    (new_odo_s, new_odo_e, new_f_s, new_f_e, new_g_s, new_g_e, curr[0])
                )
                lst = list(rows[i])
                lst[9], lst[10] = new_odo_s, new_odo_e
                lst[11], lst[13] = new_f_s, new_f_e
                lst[15], lst[17] = new_g_s, new_g_e
                rows[i] = tuple(lst)
                updated_ids.append(curr[0])

        if updates_batch:
            self.query_many(
                "UPDATE history SET odo_s=?, odo_e=?, f_s=?, f_e=?, g_s=?, g_e=? WHERE id=?",
                updates_batch,
            )
        return updated_ids

    def recalculate_all_vehicles(self):
        """Пересчитывает хронологические цепочки всех автомобилей в базе."""
        vehs = [r[0] for r in self.query("SELECT DISTINCT veh FROM history WHERE veh IS NOT NULL AND veh != ''")]
        all_ids = []
        for v in vehs:
            all_ids.extend(self.recalculate_chain(v))
        return all_ids


class TractorDatabaseManager(_BaseSQLiteManager):
    """
    Отдельная изолированная база данных для тракторов МТЗ-82 и экскаваторов на базе МТЗ-82 (tractors.db).
    Включает справочник видов работ (work_types), нормы ДТ по километражу (л/100 км) и моточасам (л/м·ч),
    и непрерывную цепочку показаний одометра (км), счётчика моточасов (м/ч) и остатка ДТ.
    """

    def __init__(self, db_name="tractors.db"):
        super().__init__(db_name)
        self.init_db()

    def init_db(self, auto_sync_history=True):
        conn = self.get_conn()
        cur = conn.cursor()
        cur.execute("CREATE TABLE IF NOT EXISTS orgs (id INTEGER PRIMARY KEY, name TEXT UNIQUE)")
        cur.execute("CREATE TABLE IF NOT EXISTS clients (id INTEGER PRIMARY KEY, name TEXT UNIQUE)")
        cur.execute("CREATE TABLE IF NOT EXISTS staff (id INTEGER PRIMARY KEY, fio TEXT, role TEXT, UNIQUE(fio, role))")
        self._migrate_staff_table(cur)
        cur.execute(
            """CREATE TABLE IF NOT EXISTS vehicles (
                id INTEGER PRIMARY KEY,
                mark TEXT UNIQUE,
                reg TEXT,
                gar TEXT,
                eq_type TEXT,
                f_brand TEXT,
                norm_km REAL,
                norm_mh REAL
            )"""
        )
        cur.execute(
            """CREATE TABLE IF NOT EXISTS drivers (
                id INTEGER PRIMARY KEY,
                fio TEXT UNIQUE,
                cat TEXT,
                exp TEXT,
                tab TEXT,
                ser TEXT,
                num TEXT
            )"""
        )
        cur.execute(
            """CREATE TABLE IF NOT EXISTS work_types (
                id INTEGER PRIMARY KEY,
                name TEXT UNIQUE,
                norm_mh REAL,
                unit TEXT
            )"""
        )
        cur.execute(
            """CREATE TABLE IF NOT EXISTS locs (
                id INTEGER PRIMARY KEY,
                name TEXT UNIQUE,
                lat REAL,
                lon REAL
            )"""
        )
        cur.execute(
            """CREATE TABLE IF NOT EXISTS route_cache (
                key TEXT PRIMARY KEY, dist REAL, coords TEXT
            )"""
        )
        cur.execute(
            """CREATE TABLE IF NOT EXISTS history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                wb_num TEXT,
                date TEXT,
                time_out TEXT,
                time_in TEXT,
                org TEXT,
                client TEXT,
                veh TEXT,
                drv TEXT,
                work_type TEXT,
                norm_km REAL,
                norm_mh REAL,
                odo_s REAL,
                odo_e REAL,
                moto_s REAL,
                moto_e REAL,
                moto_hours REAL,
                f_s REAL,
                f_a REAL,
                f_e REAL,
                f_norm_km REAL,
                f_norm_mh REAL,
                f_norm_total REAL,
                f_act REAL,
                route TEXT,
                dist REAL,
                form_type TEXT
            )"""
        )
        cur.execute("CREATE INDEX IF NOT EXISTS idx_tr_history_veh ON history(veh)")
        self._cleanup_demo_locs(cur)
        conn.commit()
        if auto_sync_history:
            self.renumber_history_ids_by_date(renumber_wb_num=True)
            self.recalculate_all_vehicles()

    def _cleanup_demo_locs(self, cur):
        """Удаляет предустановленные демонстрационные адреса из справочника тракторов, не затрагивая пользовательские."""
        demo_locs = (
            "Беляевка, ул. Ленинская (стройплощадка)",
            "Буртинский (База МТП)",
            "Буртинский, ул. Советская (участок водопровода)",
            "Оренбург, ул. Родимцева, 16 (объект ДорСтрой)",
            "Саракташ (коммунальный участок)",
        )
        placeholders = ",".join("?" for _ in demo_locs)
        cur.execute(f"DELETE FROM locs WHERE name IN ({placeholders})", demo_locs)

    def _migrate_staff_table(self, cur):
        """Снимает старое ограничение fio TEXT UNIQUE в tractors.db."""
        cur.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='staff'")
        row = cur.fetchone()
        if row and row[0] and "fio TEXT UNIQUE" in row[0]:
            cur.execute("ALTER TABLE staff RENAME TO staff_old")
            cur.execute(
                "CREATE TABLE staff (id INTEGER PRIMARY KEY, fio TEXT, role TEXT, UNIQUE(fio, role))"
            )
            cur.execute("INSERT OR IGNORE INTO staff (id, fio, role) SELECT id, fio, role FROM staff_old")
            cur.execute("DROP TABLE staff_old")

    def get_last_veh(self, veh):
        if not veh:
            return None
        rows = self.query(
            "SELECT id, wb_num, date, odo_e, moto_e, f_e, time_out FROM history WHERE veh=?",
            (veh,),
        )
        if not rows:
            return None
        rows.sort(key=lambda r: _wb_chrono_key(r[2], r[1], r[0], r[6]))
        last = rows[-1]
        return (last[3], last[4], last[5])

    def get_record_by_id(self, r_id):
        rows = self.query("SELECT * FROM history WHERE id=?", (r_id,))
        return rows[0] if rows else None

    def find_by_veh_and_date(self, veh, date_str, wb_num=None):
        """Находит существующий путевой лист трактора на указанную дату (приоритет совпадению № ПЛ)."""
        if not veh or not date_str:
            return None
        rows = self.query(
            "SELECT * FROM history WHERE veh=? AND date=? ORDER BY id DESC",
            (str(veh).strip(), str(date_str).strip()),
        )
        if not rows:
            return None
        if wb_num is not None:
            wb_s = str(wb_num).strip()
            exact = next((r for r in rows if str(r[1]).strip() == wb_s), None)
            if exact:
                return exact
        return rows[0]

    def add_driver(self, fio, cat, exp, tab, ser, num):
        self.query(
            "INSERT OR REPLACE INTO drivers (fio, cat, exp, tab, ser, num) VALUES (?, ?, ?, ?, ?, ?)",
            (fio, cat, exp, tab, ser, num),
        )

    def update_driver(self, d_id, fio, cat, exp, tab, ser, num):
        self.query(
            "UPDATE drivers SET fio=?, cat=?, exp=?, tab=?, ser=?, num=? WHERE id=?",
            (fio, cat, exp, tab, ser, num, d_id),
        )

    def add_vehicle(self, mark, reg, gar, eq_type, f_brand, norm_km, norm_mh):
        self.query(
            "INSERT OR REPLACE INTO vehicles (mark, reg, gar, eq_type, f_brand, norm_km, norm_mh) VALUES (?,?,?,?,?,?,?)",
            (mark, reg, gar, eq_type, f_brand, float(norm_km or 0.0), float(norm_mh or 0.0)),
        )

    def update_vehicle(self, v_id, mark, reg, gar, eq_type, f_brand, norm_km, norm_mh):
        self.query(
            "UPDATE vehicles SET mark=?, reg=?, gar=?, eq_type=?, f_brand=?, norm_km=?, norm_mh=? WHERE id=?",
            (mark, reg, gar, eq_type, f_brand, float(norm_km or 0.0), float(norm_mh or 0.0), v_id),
        )

    def update_org(self, o_id, name):
        try:
            self.query("UPDATE orgs SET name=? WHERE id=?", (name, o_id))
            return True
        except sqlite3.IntegrityError:
            return False

    def update_client(self, c_id, name):
        try:
            self.query("UPDATE clients SET name=? WHERE id=?", (name, c_id))
            return True
        except sqlite3.IntegrityError:
            return False

    def add_staff(self, fio, role):
        self.query(
            "INSERT OR IGNORE INTO staff (fio, role) VALUES (?, ?)",
            (fio, role),
        )

    def update_staff(self, s_id, fio, role):
        try:
            self.query(
                "UPDATE staff SET fio=?, role=? WHERE id=?",
                (fio, role, s_id),
            )
            return True
        except sqlite3.IntegrityError:
            return False

    def add_work_type(self, name, norm_mh, unit="м/ч"):
        self.query(
            "INSERT OR REPLACE INTO work_types (name, norm_mh, unit) VALUES (?, ?, ?)",
            (name, float(norm_mh or 0.0), unit or "м/ч"),
        )

    def update_work_type(self, w_id, name, norm_mh, unit="м/ч"):
        try:
            self.query(
                "UPDATE work_types SET name=?, norm_mh=?, unit=? WHERE id=?",
                (name, float(norm_mh or 0.0), unit or "м/ч", w_id),
            )
            return True
        except sqlite3.IntegrityError:
            return False

    def save_wb(
        self,
        wb_num,
        date,
        t_out,
        t_in,
        org,
        client,
        veh,
        drv,
        work_type,
        norm_km,
        norm_mh,
        odo_s,
        odo_e,
        moto_s,
        moto_e,
        moto_hours,
        f_s,
        f_a,
        f_e,
        f_norm_km,
        f_norm_mh,
        f_norm_total,
        f_act,
        route,
        dist,
        form_type="412-АПК",
    ):
        self.query(
            """INSERT INTO history (
                wb_num, date, time_out, time_in, org, client, veh, drv,
                work_type, norm_km, norm_mh, odo_s, odo_e, moto_s, moto_e, moto_hours,
                f_s, f_a, f_e, f_norm_km, f_norm_mh, f_norm_total, f_act,
                route, dist, form_type
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                wb_num,
                date,
                t_out,
                t_in,
                org,
                client,
                veh,
                drv,
                work_type,
                norm_km,
                norm_mh,
                odo_s,
                odo_e,
                moto_s,
                moto_e,
                moto_hours,
                f_s,
                f_a,
                f_e,
                f_norm_km,
                f_norm_mh,
                f_norm_total,
                f_act,
                route,
                dist,
                form_type,
            ),
        )

    def update_wb(
        self,
        w_id,
        wb_num,
        date,
        t_out,
        t_in,
        org,
        client,
        veh,
        drv,
        work_type,
        norm_km,
        norm_mh,
        odo_s,
        odo_e,
        moto_s,
        moto_e,
        moto_hours,
        f_s,
        f_a,
        f_e,
        f_norm_km,
        f_norm_mh,
        f_norm_total,
        f_act,
        route,
        dist,
        form_type="412-АПК",
    ):
        self.query(
            """UPDATE history SET
                wb_num=?, date=?, time_out=?, time_in=?, org=?, client=?, veh=?, drv=?,
                work_type=?, norm_km=?, norm_mh=?, odo_s=?, odo_e=?, moto_s=?, moto_e=?, moto_hours=?,
                f_s=?, f_a=?, f_e=?, f_norm_km=?, f_norm_mh=?, f_norm_total=?, f_act=?,
                route=?, dist=?, form_type=?
               WHERE id=?""",
            (
                wb_num,
                date,
                t_out,
                t_in,
                org,
                client,
                veh,
                drv,
                work_type,
                norm_km,
                norm_mh,
                odo_s,
                odo_e,
                moto_s,
                moto_e,
                moto_hours,
                f_s,
                f_a,
                f_e,
                f_norm_km,
                f_norm_mh,
                f_norm_total,
                f_act,
                route,
                dist,
                form_type,
                w_id,
            ),
        )

    def recalculate_chain(self, veh):
        """
        Хронологически (от начала года до конца) пересчитывает цепочку одометра, моточасов и остатков ДТ
        для трактора/экскаватора за одну пакетную транзакцию и возвращает список изменившихся ID.
        """
        if not veh:
            return []
        rows = list(self.query("SELECT * FROM history WHERE veh=?", (veh,)))
        if not rows:
            return []
        rows.sort(key=lambda r: _wb_chrono_key(r[2], r[1], r[0], r[3]))

        # Карта справочных норм л/100км техники (на случай, если в старой записи norm_km была 0)
        veh_rows = self.query("SELECT norm_km FROM vehicles WHERE mark=?", (veh,))
        veh_ref_norm_km = float(veh_rows[0][0] or 0.0) if veh_rows else 0.0

        def _recalc_row_fuel(rec):
            route_raw = str(rec[24] or "")
            is_only_mh = (
                "'only_motohours': True" in route_raw
                or '"only_motohours": True' in route_raw
                or '"only_motohours": true' in route_raw
            )
            eff_dist = 0.0 if is_only_mh else float(rec[25] or 0.0)
            norm_km_v = float(rec[10] or 0.0) or veh_ref_norm_km
            norm_mh_v = float(rec[11] or 0.0)
            moto_h_v = float(rec[16] or 0.0)
            f_n_mh = float(rec[21] or 0.0)
            if moto_h_v > 0.0 and f_n_mh <= 0.001 and norm_mh_v > 0.0:
                f_n_mh = round(moto_h_v * norm_mh_v, 2)

            if is_only_mh or eff_dist <= 0.001:
                f_n_km = 0.0
            else:
                f_n_km = float(rec[20] or 0.0)
                if f_n_km <= 0.001 and norm_km_v > 0.0:
                    f_n_km = round((eff_dist / 100.0) * norm_km_v, 2)

            old_tot = float(rec[22] or 0.0)
            f_n_tot = round(f_n_km + f_n_mh, 2)
            f_act_v = float(rec[23] or 0.0)
            if f_act_v <= 0.001 or abs(f_act_v - old_tot) < 0.05:
                f_act_v = f_n_tot
            return eff_dist, f_n_km, f_n_mh, f_n_tot, f_act_v

        updated_ids = []
        updates_batch = []

        # Проверяем первую запись
        first = rows[0]
        f_odo_s = float(first[12] or 0.0)
        f_dist, f_n_km0, f_n_mh0, f_n_tot0, f_f_act = _recalc_row_fuel(first)
        f_odo_e = round(f_odo_s + f_dist, 2)
        f_moto_s = float(first[14] or 0.0)
        f_moto_h = float(first[16] or 0.0)
        f_moto_e = round(f_moto_s + f_moto_h, 2)
        f_f_s = float(first[17] or 0.0)
        f_f_a = float(first[18] or 0.0)
        f_f_e = round(f_f_s + f_f_a - f_f_act, 2)
        old_first = (
            float(first[12] or 0.0),
            float(first[13] or 0.0),
            float(first[14] or 0.0),
            float(first[15] or 0.0),
            float(first[17] or 0.0),
            float(first[19] or 0.0),
            float(first[20] or 0.0),
            float(first[21] or 0.0),
            float(first[22] or 0.0),
            float(first[23] or 0.0),
            float(first[25] or 0.0),
        )
        new_first = (
            f_odo_s,
            f_odo_e,
            f_moto_s,
            f_moto_e,
            f_f_s,
            f_f_e,
            f_n_km0,
            f_n_mh0,
            f_n_tot0,
            f_f_act,
            f_dist,
        )
        if any(abs(a - b) > 1e-4 for a, b in zip(old_first, new_first)):
            updates_batch.append(
                (
                    f_odo_s,
                    f_odo_e,
                    f_moto_s,
                    f_moto_e,
                    f_f_s,
                    f_f_e,
                    f_n_km0,
                    f_n_mh0,
                    f_n_tot0,
                    f_f_act,
                    f_dist,
                    first[0],
                )
            )
            lst0 = list(rows[0])
            lst0[12], lst0[13] = f_odo_s, f_odo_e
            lst0[14], lst0[15] = f_moto_s, f_moto_e
            lst0[17], lst0[19] = f_f_s, f_f_e
            lst0[20], lst0[21], lst0[22], lst0[23] = f_n_km0, f_n_mh0, f_n_tot0, f_f_act
            lst0[25] = f_dist
            rows[0] = tuple(lst0)
            updated_ids.append(first[0])

        for i in range(1, len(rows)):
            prev = rows[i - 1]
            curr = rows[i]
            prev_odo_e = float(prev[13] or 0.0)
            prev_moto_e = float(prev[15] or 0.0)
            new_odo_s = prev_odo_e if prev_odo_e > 0.001 else float(curr[12] or 0.0)
            new_moto_s = prev_moto_e if prev_moto_e > 0.001 else float(curr[14] or 0.0)
            new_f_s = float(prev[19] or 0.0)

            dist, f_n_km_c, f_n_mh_c, f_n_tot_c, f_act = _recalc_row_fuel(curr)
            moto_h = float(curr[16] or 0.0)
            f_a = float(curr[18] or 0.0)

            new_odo_e = round(new_odo_s + dist, 2)
            new_moto_e = round(new_moto_s + moto_h, 2)
            new_f_e = round(new_f_s + f_a - f_act, 2)

            old_vals = (
                float(curr[12] or 0.0),
                float(curr[13] or 0.0),
                float(curr[14] or 0.0),
                float(curr[15] or 0.0),
                float(curr[17] or 0.0),
                float(curr[19] or 0.0),
                float(curr[20] or 0.0),
                float(curr[21] or 0.0),
                float(curr[22] or 0.0),
                float(curr[23] or 0.0),
                float(curr[25] or 0.0),
            )
            new_vals = (
                new_odo_s,
                new_odo_e,
                new_moto_s,
                new_moto_e,
                new_f_s,
                new_f_e,
                f_n_km_c,
                f_n_mh_c,
                f_n_tot_c,
                f_act,
                dist,
            )
            changed = any(abs(a - b) > 1e-4 for a, b in zip(old_vals, new_vals))

            if changed:
                updates_batch.append(
                    (
                        new_odo_s,
                        new_odo_e,
                        new_moto_s,
                        new_moto_e,
                        new_f_s,
                        new_f_e,
                        f_n_km_c,
                        f_n_mh_c,
                        f_n_tot_c,
                        f_act,
                        dist,
                        curr[0],
                    )
                )
                lst = list(rows[i])
                lst[12], lst[13] = new_odo_s, new_odo_e
                lst[14], lst[15] = new_moto_s, new_moto_e
                lst[17], lst[19] = new_f_s, new_f_e
                lst[20], lst[21], lst[22], lst[23] = f_n_km_c, f_n_mh_c, f_n_tot_c, f_act
                lst[25] = dist
                rows[i] = tuple(lst)
                updated_ids.append(curr[0])

        if updates_batch:
            self.query_many(
                "UPDATE history SET odo_s=?, odo_e=?, moto_s=?, moto_e=?, f_s=?, f_e=?, f_norm_km=?, f_norm_mh=?, f_norm_total=?, f_act=?, dist=? WHERE id=?",
                updates_batch,
            )
        return updated_ids

    def recalculate_all_vehicles(self):
        """Пересчитывает хронологические цепочки всех тракторов в базе."""
        vehs = [r[0] for r in self.query("SELECT DISTINCT veh FROM history WHERE veh IS NOT NULL AND veh != ''")]
        all_ids = []
        for v in vehs:
            all_ids.extend(self.recalculate_chain(v))
        return all_ids

    def apply_work_type_norm(
        self,
        work_name,
        new_norm_mh,
        date_from=None,
        date_to=None,
        update_reference=True,
    ):
        """
        Массово обновляет норму расхода ДТ (л/м·ч) ТОЛЬКО для указанного вида работ
        во всех сохранённых путевых листах или за выбранный период (date_from..date_to),
        не затрагивая остальные виды работ, затем каскадно пересчитывает цепочку остатков ДТ.
        """
        target_clean = str(work_name or "").strip()
        target_low = target_clean.lower()
        new_norm = round(float(new_norm_mh or 0.0), 3)
        if not target_clean or new_norm <= 0:
            return {"matched_ids": [], "cascaded_ids": [], "affected_vehicles": []}

        if update_reference:
            self.query(
                "UPDATE work_types SET norm_mh=? WHERE LOWER(TRIM(name))=LOWER(TRIM(?))",
                (new_norm, target_clean),
            )

        d_from_t = _parse_date_tuple(date_from) if date_from else None
        d_to_t = _parse_date_tuple(date_to) if date_to else None

        rows = list(self.query("SELECT * FROM history"))
        updates_batch = []
        matched_ids = []
        affected_vehs = set()

        for rec in rows:
            rec_id = rec[0]
            rec_date_t = _parse_date_tuple(rec[2])
            if d_from_t and (not rec_date_t or rec_date_t < d_from_t):
                continue
            if d_to_t and (not rec_date_t or rec_date_t > d_to_t):
                continue

            old_norm_mh = float(rec[11] or 0.0)
            moto_h = float(rec[16] or 0.0)
            f_norm_km = float(rec[20] or 0.0)
            old_f_norm_mh = float(rec[21] or 0.0)
            old_f_norm_tot = float(rec[22] or 0.0)
            old_f_act = float(rec[23] or 0.0)
            route_raw = rec[24] or ""

            try:
                raw_items = ast.literal_eval(route_raw) if route_raw else []
                if not isinstance(raw_items, list):
                    raw_items = []
            except Exception:
                raw_items = []

            work_items = [
                it for it in raw_items if isinstance(it, dict) and it.get("entry_type") == "work"
            ]

            changed = False
            new_rec_norm_mh = old_norm_mh
            new_f_norm_mh = old_f_norm_mh
            new_route_raw = route_raw

            if work_items:
                matched_here = False
                for w in work_items:
                    w_name = str(w.get("name", "")).strip().lower()
                    if w_name == target_low:
                        w_mh = float(w.get("mh", 0.0) or 0.0)
                        w_old_norm = float(w.get("norm_mh", 0.0) or 0.0)
                        w_old_dt = float(w.get("dt_l", 0.0) or 0.0)
                        mult = 1.0
                        if w_mh > 0 and w_old_norm > 0 and w_old_dt > 0:
                            ratio = w_old_dt / (w_mh * w_old_norm)
                            if 0.99 <= ratio <= 1.65:
                                mult = round(ratio, 4)
                        w["norm_mh"] = round(new_norm, 2)
                        w["dt_l"] = round(w_mh * new_norm * mult, 2)
                        matched_here = True
                if matched_here:
                    changed = True
                    new_f_norm_mh = round(
                        sum(float(w.get("dt_l", 0.0) or 0.0) for w in work_items), 2
                    )
                    total_w_mh = sum(float(w.get("mh", 0.0) or 0.0) for w in work_items)
                    if all(str(w.get("name", "")).strip().lower() == target_low for w in work_items):
                        new_rec_norm_mh = round(new_norm, 2)
                    elif total_w_mh > 0:
                        new_rec_norm_mh = round(
                            sum(
                                float(w.get("mh", 0.0) or 0.0) * float(w.get("norm_mh", 0.0) or 0.0)
                                for w in work_items
                            )
                            / total_w_mh,
                            2,
                        )
                    new_route_raw = repr(raw_items)
            else:
                wt_str = str(rec[9] or "").strip().lower()
                if wt_str == target_low or wt_str.startswith(target_low + " ("):
                    changed = True
                    mult = 1.0
                    if moto_h > 0 and old_norm_mh > 0 and old_f_norm_mh > 0:
                        ratio = old_f_norm_mh / (moto_h * old_norm_mh)
                        if 0.99 <= ratio <= 1.65:
                            mult = round(ratio, 4)
                    new_rec_norm_mh = round(new_norm, 2)
                    new_f_norm_mh = round(moto_h * new_norm * mult, 2)

            if changed:
                new_f_norm_tot = round(f_norm_km + new_f_norm_mh, 2)
                if abs(old_f_act - old_f_norm_tot) < 0.05 or old_f_act <= 0.001:
                    new_f_act = new_f_norm_tot
                else:
                    new_f_act = old_f_act
                updates_batch.append(
                    (
                        new_rec_norm_mh,
                        new_f_norm_mh,
                        new_f_norm_tot,
                        new_f_act,
                        new_route_raw,
                        rec_id,
                    )
                )
                matched_ids.append(rec_id)
                if rec[7]:
                    affected_vehs.add(str(rec[7]))

        if updates_batch:
            self.query_many(
                "UPDATE history SET norm_mh=?, f_norm_mh=?, f_norm_total=?, f_act=?, route=? WHERE id=?",
                updates_batch,
            )

        cascaded_ids = []
        for v in sorted(affected_vehs):
            cascaded_ids.extend(self.recalculate_chain(v) or [])

        return {
            "matched_ids": matched_ids,
            "cascaded_ids": list(dict.fromkeys(matched_ids + cascaded_ids)),
            "affected_vehicles": sorted(affected_vehs),
        }


_BACKUP_LOCK = threading.Lock()


def create_fast_db_backup(
    cars_db="waybills.db",
    tractors_db="tractors.db",
    backup_dir="Авто_резервные_копии",
):
    """
    Сверхбыстрое неблокирующее резервное копирование баз SQLite через встроенный online backup API
    (без WAL TRUNCATE блокировок и без пережатия всех Excel-файлов — выполняется за ~10–20 мс).
    """
    if not _BACKUP_LOCK.acquire(blocking=False):
        return None
    try:
        os.makedirs(backup_dir, exist_ok=True)
        for db_path in (cars_db, tractors_db):
            if db_path and os.path.exists(db_path):
                dst_path = os.path.join(backup_dir, os.path.basename(db_path))
                tmp_path = dst_path + ".tmp"
                src_conn = sqlite3.connect(db_path, timeout=2.0)
                try:
                    dst_conn = sqlite3.connect(tmp_path)
                    try:
                        src_conn.backup(dst_conn, pages=64, sleep=0.005)
                    finally:
                        dst_conn.close()
                finally:
                    src_conn.close()
                os.replace(tmp_path, dst_path)
        for mem_f in ("memory.json", "memory_tractors.json"):
            if os.path.exists(mem_f):
                shutil.copy2(mem_f, os.path.join(backup_dir, mem_f))
        return backup_dir
    except Exception:
        return None
    finally:
        _BACKUP_LOCK.release()


def create_unified_backup(
    cars_db="waybills.db",
    tractors_db="tractors.db",
    unified_sqlite="Единая_база_Легковые_и_Трактора.db",
    unified_zip="Резервная_копия_Легковые_и_Трактора.zip",
):
    """
    Объединяет изолированную базу легковых авто (waybills.db) и базу тракторов/экскаваторов (tractors.db)
    в единую резервную базу SQLite (с префиксами cars_* и tractors_*) и формирует общий архив .zip
    со всеми базами, настройками, шаблонами и готовыми путевыми листами.
    """
    with _BACKUP_LOCK:
        _checkpoint_db_file(cars_db)
        _checkpoint_db_file(tractors_db)

        if os.path.exists(unified_sqlite):
            os.remove(unified_sqlite)

        with sqlite3.connect(unified_sqlite) as uconn:
            ucur = uconn.cursor()
            if os.path.exists(cars_db):
                ucur.execute("ATTACH DATABASE ? AS cars_src", (os.path.abspath(cars_db),))
                for tbl in ("orgs", "clients", "staff", "vehicles", "drivers", "locs", "history"):
                    ucur.execute(f"CREATE TABLE cars_{tbl} AS SELECT * FROM cars_src.{tbl}")
                uconn.commit()
                ucur.execute("DETACH DATABASE cars_src")

            if os.path.exists(tractors_db):
                ucur.execute("ATTACH DATABASE ? AS tr_src", (os.path.abspath(tractors_db),))
                for tbl in ("orgs", "clients", "staff", "vehicles", "drivers", "work_types", "locs", "history"):
                    ucur.execute(f"CREATE TABLE tractors_{tbl} AS SELECT * FROM tr_src.{tbl}")
                uconn.commit()
                ucur.execute("DETACH DATABASE tr_src")

        with zipfile.ZipFile(unified_zip, "w", zipfile.ZIP_DEFLATED, compresslevel=1) as zf:
            for f in (cars_db, tractors_db, unified_sqlite, "memory.json", "memory_tractors.json"):
                if os.path.exists(f):
                    zf.write(f, arcname=f)
            for folder in ("templates", "Готовые_Путевые_Листы"):
                if os.path.exists(folder):
                    for root, _dirs, files in os.walk(folder):
                        for file in files:
                            if file.endswith(".xlsx") and not file.startswith("~$"):
                                full_p = os.path.join(root, file)
                                rel_p = os.path.relpath(full_p, ".")
                                zf.write(full_p, arcname=rel_p)

        return {
            "unifiedSqlite": unified_sqlite,
            "unifiedZip": unified_zip,
            "unified_db_path": unified_sqlite,
            "zip_path": unified_zip,
        }


def create_mode_backup(mode="cars", out_zip_path=None):
    """
    Создаёт изолированную резервную копию отдельно для легковых авто (waybills.db + memory.json + ПЛ легковых)
    или отдельно для тракторов МТЗ-82 / экскаваторов (tractors.db + memory_tractors.json + ПЛ тракторов).
    """
    with _BACKUP_LOCK:
        if mode == "tractors":
            db_file = "tractors.db"
            mem_file = "memory_tractors.json"
            zip_name = out_zip_path or "Резервная_копия_Трактора_МТЗ82.zip"
            pl_folder = os.path.join("Готовые_Путевые_Листы", "Трактора_МТЗ82")
        else:
            db_file = "waybills.db"
            mem_file = "memory.json"
            zip_name = out_zip_path or "Резервная_копия_Легковые_авто.zip"
            pl_folder = "Готовые_Путевые_Листы"

        _checkpoint_db_file(db_file)

        with zipfile.ZipFile(zip_name, "w", zipfile.ZIP_DEFLATED, compresslevel=1) as zf:
            for f in (db_file, mem_file):
                if os.path.exists(f):
                    zf.write(f, arcname=f)
            if os.path.isdir(pl_folder):
                if mode == "tractors":
                    for root, _dirs, files in os.walk(pl_folder):
                        for file in files:
                            if file.endswith(".xlsx") and not file.startswith("~$"):
                                full_p = os.path.join(root, file)
                                zf.write(full_p, arcname=os.path.relpath(full_p, "."))
                else:
                    for file in os.listdir(pl_folder):
                        full_p = os.path.join(pl_folder, file)
                        if (
                            os.path.isfile(full_p)
                            and file.endswith(".xlsx")
                            and not file.startswith("~$")
                        ):
                            zf.write(full_p, arcname=os.path.relpath(full_p, "."))

        return {
            "mode": mode,
            "db_path": db_file,
            "zip_path": zip_name,
        }
