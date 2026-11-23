import ast
import json
import os
import shutil
import sys

if getattr(sys, "frozen", False):
    _BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if __name__ == "__main__":
    os.chdir(_BASE_DIR)
_INTERNAL_DIR = os.path.join(_BASE_DIR, "_internal")
if os.path.isdir(_INTERNAL_DIR) and _INTERNAL_DIR not in sys.path:
    sys.path.insert(0, _INTERNAL_DIR)

from database import DatabaseManager, TractorDatabaseManager, create_unified_backup
from logic_gsm import calculate_tractor_diesel_consumption
from parser_f3 import format_ru_date, generate_f3_from_excel, generate_tractor_wb_excel

OUT_DIR = "Готовые_Путевые_Листы"
TRACTOR_OUT_DIR = "Готовые_Путевые_Листы/Трактора_МТЗ82"


def load_memory(filename="memory.json"):
    if os.path.exists(filename):
        try:
            with open(filename, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_memory(data, filename="memory.json"):
    current = load_memory(filename)
    current.update(data)
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(current, f, ensure_ascii=False, indent=2)
    return current


def get_latest_excel_path(folder=OUT_DIR):
    if not os.path.exists(folder):
        return None
    files = [
        f
        for f in os.listdir(folder)
        if os.path.isfile(os.path.join(folder, f))
        and f.lower().endswith(".xlsx")
        and not f.startswith("~$")
    ]
    if not files:
        return None
    files.sort(
        key=lambda name: os.path.getmtime(os.path.join(folder, name)),
        reverse=True,
    )
    return f"{folder}/{files[0]}"


def ensure_orenburg_workspace(db: DatabaseManager):
    """
    Replaces any old synthetic Tobolsk seed data with Orenburg / Буртинский data
    and ensures a populated waybill exists in Готовые_Путевые_Листы/.
    """
    if os.path.exists("Путевые_листы"):
        shutil.rmtree("Путевые_листы", ignore_errors=True)

    orgs = db.get_list("orgs")
    has_tobolsk = any("Тобольск" in (r[1] or "") for r in orgs)
    locs = db.get_list("locs")
    has_tobolsk_locs = any("Тобольск" in (r[1] or "") or "Абалак" in (r[1] or "") for r in locs)

    if has_tobolsk or has_tobolsk_locs or not orgs:
        db.query("DELETE FROM orgs")
        db.query("DELETE FROM clients")
        db.query("DELETE FROM staff")
        db.query("DELETE FROM vehicles")
        db.query("DELETE FROM drivers")
        db.query("DELETE FROM locs")
        db.query("DELETE FROM history")

        db.query(
            "INSERT INTO orgs (name) VALUES (?)",
            ("ООО «ОренбургАвтоТранс», г. Оренбург, ул. Родимцева, 16, тел. 8(3532) 44-12-00",),
        )
        db.query(
            "INSERT INTO orgs (name) VALUES (?)",
            ("ИП Коваленко С.В., Оренбургская обл., пос. Буртинский",),
        )

        db.query("INSERT INTO clients (name) VALUES (?)", ("ООО «ОренбургАвтоТранс»",))
        db.query("INSERT INTO clients (name) VALUES (?)", ("Администрация Беляевского района",))

        db.query("INSERT INTO staff (fio, role) VALUES (?,?)", ("Смирнов А.В.", "Механик"))
        db.query("INSERT INTO staff (fio, role) VALUES (?,?)", ("Иванова Е.С.", "Диспетчер"))

        db.query(
            "INSERT INTO vehicles (mark, reg, gar, f_brand, norm, f_brand_gas, norm_gas) VALUES (?,?,?,?,?,?,?)",
            ("ВАЗ 21214 Нива", "О 456 РВ 56", "101", "АИ-92", 10.8, "", 0.0),
        )
        db.query(
            "INSERT INTO vehicles (mark, reg, gar, f_brand, norm, f_brand_gas, norm_gas) VALUES (?,?,?,?,?,?,?)",
            ("Lada Granta", "К 156 МР 56", "102", "АИ-95", 8.2, "ПБА", 9.8),
        )
        db.query(
            "INSERT INTO vehicles (mark, reg, gar, f_brand, norm, f_brand_gas, norm_gas) VALUES (?,?,?,?,?,?,?)",
            ("УАЗ Патриот", "А 777 АА 56", "105", "АИ-92", 12.5, "ПБА", 14.5),
        )

        db.add_driver("Коваленко Сергей Викторович", "B, C", "14 лет", "0014", "56 18", "458912")
        db.add_driver("Морозов Дмитрий Игоревич", "B, BE", "9 лет", "0021", "56 20", "771234")

        orenburg_locs = [
            ("Буртинский", 51.3683131, 55.9932570),
            ("Оренбург, ул. Родимцева, 16", 51.8358182, 55.1535769),
            ("Беляевка, Оренбургская обл.", 51.3976000, 56.4172000),
            ("Оренбург, ж/д вокзал", 51.7782000, 55.0604000),
            ("Оренбург, ул. Чкалова, 32", 51.7681990, 55.1268000),
            ("Саракташ, Оренбургская обл.", 51.7876000, 56.3601000),
        ]
        for name, lat, lon in orenburg_locs:
            db.query("INSERT INTO locs (name, lat, lon) VALUES (?,?,?)", (name, lat, lon))

        segments = [
            {
                "from": "Буртинский",
                "to": "Оренбург, ул. Родимцева, 16",
                "t_out": "08:00",
                "t_in": "12:30",
                "km": 97.7,
            },
            {
                "from": "Оренбург, ул. Родимцева, 16",
                "to": "Буртинский",
                "t_out": "13:30",
                "t_in": "17:00",
                "km": 97.7,
            },
        ]
        db.save_wb(
            "001",
            "16.11.2026",
            "08:00",
            "17:00",
            "ООО «ОренбургАвтоТранс», г. Оренбург, ул. Родимцева, 16, тел. 8(3532) 44-12-00",
            "ООО «ОренбургАвтоТранс»",
            "ВАЗ 21214 Нива",
            "Коваленко Сергей Викторович",
            45120.0,
            45315.0,
            18.0,
            25.0,
            21.9,
            21.1,
            0.0,
            0.0,
            0.0,
            0.0,
            repr(segments),
            195.4,
        )

        save_memory(
            {
                "org": "ООО «ОренбургАвтоТранс», г. Оренбург, ул. Родимцева, 16, тел. 8(3532) 44-12-00",
                "client": "ООО «ОренбургАвтоТранс»",
                "pickup": "Буртинский",
                "veh": "ВАЗ 21214 Нива",
                "vehicle": "ВАЗ 21214 Нива",
                "drv": "Коваленко Сергей Викторович",
                "driver": "Коваленко Сергей Викторович",
                "mech": "Смирнов А.В.",
                "mechanic": "Смирнов А.В.",
                "disp": "Иванова Е.С.",
                "dispatcher": "Иванова Е.С.",
                "fuel_type": "Бензин",
                "winter": 10.0,
                "city": 0.0,
                "dirt": 0.0,
                "trailer": 0.0,
                "create_excel": True,
            },
            "memory.json",
        )

    os.makedirs(OUT_DIR, exist_ok=True)
    latest_car_xlsx = get_latest_excel_path(OUT_DIR)
    if not latest_car_xlsx or os.path.getsize(latest_car_xlsx) < 12000:
        excel_data = {
            "wb_num": "001",
            "date": "16.11.2026",
            "ru_date": format_ru_date("16.11.2026"),
            "org": "ООО «ОренбургАвтоТранс», г. Оренбург, ул. Родимцева, 16, тел. 8(3532) 44-12-00",
            "okpo": "48291034",
            "client": "ООО «ОренбургАвтоТранс»",
            "pickup": "Буртинский",
            "drv": "Коваленко Сергей Викторович",
            "tab": "0014",
            "ser": "56 18",
            "num": "458912",
            "cat": "B, C",
            "car": "ВАЗ 21214 Нива",
            "reg": "О 456 РВ 56",
            "gar": "101",
            "mech": "Смирнов А.В.",
            "disp": "Иванова Е.С.",
            "odo_s": 45120,
            "odo_e": 45315,
            "time_out": "08:00",
            "time_in": "17:00",
            "time_tot": "9.0",
            "f_brand": "АИ-92",
            "f_s": "18.0",
            "f_a": "25.0",
            "f_e": "21.9",
            "spent": "21.1",
            "f_act": "21.1",
            "dist": 195.4,
            "msg_type": "пригородное / междугородное",
            "trans_type": "служебная",
            "trailer": "без прицепа",
            "seg_from": ["Буртинский", "Оренбург, ул. Родимцева, 16"],
            "seg_to": ["Оренбург, ул. Родимцева, 16", "Буртинский"],
            "seg_t_out": ["08:00", "13:30"],
            "seg_t_in": ["12:30", "17:00"],
            "seg_km": [97.7, 97.7],
        }
        out_rel = f"{OUT_DIR}/ПЛ_001_16.11.2026_ВАЗ_21214_Нива.xlsx"
        generate_f3_from_excel("templates/put-list-f3.xlsx", out_rel, excel_data)


def ensure_tractor_workspace(tdb: TractorDatabaseManager):
    """
    Создаёт полностью отдельную базу тракторов МТЗ-82 и экскаваторов на базе МТЗ-82 (tractors.db),
    справочник видов работ с нормами ДТ по моточасам и готовые путевые листы на официальных бланках
    Форма № 412-АПК и Форма № ЭСМ-2 в Готовые_Путевые_Листы/Трактора_МТЗ82/.
    """
    orgs = tdb.get_list("orgs")
    if not orgs:
        tdb.query(
            "INSERT INTO orgs (name) VALUES (?)",
            ("МУП «ЖКХ Беляевского района», Оренбургская обл., пос. Буртинский, ул. Центральная, 10",),
        )
        tdb.query(
            "INSERT INTO orgs (name) VALUES (?)",
            ("ООО «ОренбургДорСтрой», г. Оренбург, ул. Родимцева, 16",),
        )

        tdb.query("INSERT INTO clients (name) VALUES (?)", ("Администрация пос. Буртинский",))
        tdb.query("INSERT INTO clients (name) VALUES (?)", ("МУП «Водоканал Оренбург»",))
        tdb.query("INSERT INTO clients (name) VALUES (?)", ("ООО «ОренбургДорСтрой»",))

        tdb.query("INSERT INTO staff (fio, role) VALUES (?,?)", ("Петров Н.И.", "Механик спецтехники"))
        tdb.query("INSERT INTO staff (fio, role) VALUES (?,?)", ("Сидорова М.В.", "Диспетчер"))

        # Спецтехника: Трактор МТЗ-82 и Экскаваторы на базе МТЗ-82
        tdb.add_vehicle(
            "Экскаватор ЭО-2621 (база МТЗ-82)",
            "56 ОА 4412",
            "Т-01",
            "Экскаватор на базе МТЗ-82",
            "ДТ (дизельное)",
            16.5,
            5.2,
        )
        tdb.add_vehicle(
            "Трактор МТЗ-82.1 Беларус",
            "56 ОА 4415",
            "Т-02",
            "Трактор МТЗ-82",
            "ДТ (дизельное)",
            15.5,
            4.5,
        )
        tdb.add_vehicle(
            "Экскаватор-погрузчик ЭО-2626 (МТЗ-82)",
            "56 ОА 4419",
            "Т-03",
            "Экскаватор на базе МТЗ-82",
            "ДТ (дизельное)",
            17.0,
            5.5,
        )

        # Трактористы-машинисты (Удостоверение тракториста-машиниста Гостехнадзора)
        tdb.add_driver("Захаров Николай Петрович", "B, C, D, E (Экскаваторщик 5 разр.)", "18 лет", "Т-004", "АІ", "654321")
        tdb.add_driver("Громов Алексей Васильевич", "C, D, E (Тракторист-машинист)", "12 лет", "Т-007", "АІ", "654890")

        # Виды работ для трактора МТЗ-82 и экскаватора на базе МТЗ-82 с нормами расхода ДТ (л/моточас)
        default_work_types = [
            ("Экскавация грунта I-II кат. (обратная лопата)", 5.2, "м³"),
            ("Рытьё траншей водопровода (грунт III кат.)", 5.8, "п.м."),
            ("Погрузочно-разгрузочные работы (фронтальный ковш)", 5.0, "т"),
            ("Планировка участка / засыпка траншей отвалом", 4.6, "м²"),
            ("Очистка дорог от снега (отвал + щётка МТЗ-82)", 5.5, "км"),
            ("Транспортные работы с прицепом 2ПТС-4", 6.2, "рейс"),
            ("Работа гидромолотом (вскрытие асфальта / мёрзлый грунт)", 7.5, "м²"),
            ("Перегон своим ходом / холостой ход", 4.0, "км"),
        ]
        for w_name, w_norm, w_unit in default_work_types:
            tdb.add_work_type(w_name, w_norm, w_unit)

        # Предустановленные демонстрационные адреса отключены — список адресов заполняется пользователем

        segments = [
            {
                "from": "Буртинский (База МТП)",
                "to": "Буртинский, ул. Советская (участок водопровода)",
                "t_out": "08:00",
                "t_in": "12:00",
                "km": 6.0,
                "mh": 3.5,
                "work_type": "Экскавация грунта I-II кат. (обратная лопата)",
            },
            {
                "from": "Буртинский, ул. Советская (участок водопровода)",
                "to": "Буртинский (База МТП)",
                "t_out": "13:00",
                "t_in": "17:00",
                "km": 6.0,
                "mh": 3.0,
                "work_type": "Планировка участка / засыпка траншей отвалом",
            },
        ]
        tdb.save_wb(
            "Т-001",
            "16.11.2026",
            "08:00",
            "17:00",
            "МУП «ЖКХ Беляевского района», Оренбургская обл., пос. Буртинский, ул. Центральная, 10",
            "Администрация пос. Буртинский",
            "Экскаватор ЭО-2621 (база МТЗ-82)",
            "Захаров Николай Петрович",
            "Экскавация грунта I-II кат. (обратная лопата)",
            16.5,
            5.2,
            3420.0,
            3432.0,
            1280.5,
            1287.0,
            6.5,
            40.0,
            30.0,
            34.22,
            1.98,
            33.80,
            35.78,
            35.78,
            repr(segments),
            12.0,
            "ЭО-2621 (Форма ЭСМ-2 / 412-АПК)",
        )

        save_memory(
            {
                "org": "МУП «ЖКХ Беляевского района», Оренбургская обл., пос. Буртинский, ул. Центральная, 10",
                "client": "Администрация пос. Буртинский",
                "pickup": "Буртинский (База МТП)",
                "veh": "Экскаватор ЭО-2621 (база МТЗ-82)",
                "vehicle": "Экскаватор ЭО-2621 (база МТЗ-82)",
                "drv": "Захаров Николай Петрович",
                "driver": "Захаров Николай Петрович",
                "mech": "Петров Н.И.",
                "mechanic": "Петров Н.И.",
                "disp": "Сидорова М.В.",
                "dispatcher": "Сидорова М.В.",
                "work_type": "Экскавация грунта I-II кат. (обратная лопата)",
                "form_type": "ЭО-2621 (Форма ЭСМ-2 / 412-АПК)",
                "fuel_type": "ДТ (Дизтопливо)",
                "winter": 0.0,
                "city": 0.0,
                "dirt": 0.0,
                "trailer": 0.0,
                "create_excel": True,
            },
            "memory_tractors.json",
        )

    os.makedirs(TRACTOR_OUT_DIR, exist_ok=True)
    latest_tr_xlsx = get_latest_excel_path(TRACTOR_OUT_DIR)
    if not latest_tr_xlsx or os.path.getsize(latest_tr_xlsx) < 12000:
        for fn in os.listdir(TRACTOR_OUT_DIR):
            if fn.endswith(".xlsx"):
                try:
                    os.remove(os.path.join(TRACTOR_OUT_DIR, fn))
                except Exception:
                    pass
        excel_data_esm2 = {
            "wb_num": "Т-001",
            "date": "16.11.2026",
            "ru_date": format_ru_date("16.11.2026"),
            "form_type": "ЭО-2621 (Форма ЭСМ-2 / 412-АПК)",
            "eq_type": "Экскаватор на базе МТЗ-82",
            "org": "МУП «ЖКХ Беляевского района», Оренбургская обл., пос. Буртинский, ул. Центральная, 10",
            "okpo": "56104921",
            "client": "Администрация пос. Буртинский",
            "pickup": "Буртинский (База МТП)",
            "drv": "Захаров Николай Петрович",
            "tab": "Т-004",
            "ser": "АІ",
            "num": "654321",
            "cat": "B, C, D, E (Экскаваторщик 5 разр.)",
            "car": "Экскаватор ЭО-2621 (база МТЗ-82)",
            "reg": "56 ОА 4412",
            "gar": "Т-01",
            "mech": "Петров Н.И.",
            "disp": "Сидорова М.В.",
            "work_type": "Экскавация грунта I-II кат. (обратная лопата)",
            "trailer": "Обратная лопата 0.25 м³ + отвал",
            "msg_type": "перегон своим ходом / работа на объекте",
            "odo_s": 3420.0,
            "odo_e": 3432.0,
            "dist": 12.0,
            "moto_s": 1280.5,
            "moto_e": 1287.0,
            "moto_hours": 6.5,
            "norm_km": 16.5,
            "norm_mh": 5.2,
            "f_norm_km": 1.98,
            "f_norm_mh": 33.80,
            "f_norm_total": 35.78,
            "f_act": 35.78,
            "f_brand": "ДТ (дизельное)",
            "f_s": 40.0,
            "f_a": 30.0,
            "f_e": 34.22,
            "time_out": "08:00",
            "time_in": "17:00",
            "time_tot": "9.0",
            "segments": [
                {
                    "from": "Буртинский (База МТП)",
                    "to": "Буртинский, ул. Советская (участок водопровода)",
                    "t_out": "08:00",
                    "t_in": "12:00",
                    "km": 6.0,
                    "mh": 3.5,
                    "work_type": "Экскавация грунта I-II кат. (обратная лопата)",
                },
                {
                    "from": "Буртинский, ул. Советская (участок водопровода)",
                    "to": "Буртинский (База МТП)",
                    "t_out": "13:00",
                    "t_in": "17:00",
                    "km": 6.0,
                    "mh": 3.0,
                    "work_type": "Планировка участка / засыпка траншей отвалом",
                },
            ],
        }
        out_rel_esm2 = f"{TRACTOR_OUT_DIR}/ПЛ_Т-001_16.11.2026_Экскаватор_ЭО-2621_база_МТЗ-82.xlsx"
        generate_tractor_wb_excel(out_rel_esm2, excel_data_esm2, tpl_path="templates/put-list-esm2.xlsx")

        excel_data_412 = dict(excel_data_esm2)
        excel_data_412.update({
            "wb_num": "Т-002",
            "form_type": "Форма № 412-АПК (Трактор МТЗ-82)",
            "eq_type": "Трактор МТЗ-82",
            "car": "Трактор МТЗ-82.1 Беларус",
            "reg": "56 ОА 4415",
            "gar": "Т-02",
            "drv": "Громов Алексей Васильевич",
            "tab": "Т-007",
            "num": "654890",
            "cat": "C, D, E (Тракторист-машинист)",
            "work_type": "Очистка дорог от снега (отвал + щётка МТЗ-82)",
            "trailer": "Отвал коммунальный КО-4 + щётка МТЗ-82",
        })
        out_rel_412 = f"{TRACTOR_OUT_DIR}/ПЛ_Т-002_16.11.2026_Трактор_МТЗ-82.1_Беларус.xlsx"
        generate_tractor_wb_excel(out_rel_412, excel_data_412, tpl_path="templates/put-list-412-apk.xlsx")
        # Make Т-001 the newest mtime so default latest matches history row #1
        generate_tractor_wb_excel(out_rel_esm2, excel_data_esm2, tpl_path="templates/put-list-esm2.xlsx")



def get_tractor_state(tdb: TractorDatabaseManager):
    ensure_tractor_workspace(tdb)
    orgs = [{"id": r[0], "name": r[1]} for r in tdb.get_list("orgs")]
    clients = [{"id": r[0], "name": r[1]} for r in tdb.get_list("clients")]
    staff = [{"id": r[0], "fio": r[1], "role": r[2]} for r in tdb.get_list("staff")]
    vehicles = [
        {
            "id": r[0],
            "mark": r[1],
            "reg": r[2],
            "gar": r[3],
            "eq_type": r[4] or "Трактор МТЗ-82",
            "f_brand": r[5] or "ДТ (дизельное)",
            "norm_km": float(r[6] or 16.0),
            "norm_mh": float(r[7] or 5.0),
        }
        for r in tdb.get_list("vehicles")
    ]
    drivers = [
        {
            "id": r[0],
            "fio": r[1],
            "cat": r[2],
            "exp": r[3],
            "tab": r[4],
            "ser": r[5],
            "num": r[6],
        }
        for r in tdb.get_list("drivers")
    ]
    work_types = [
        {
            "id": r[0],
            "name": r[1],
            "norm_mh": float(r[2] or 5.0),
            "unit": r[3] or "м³",
        }
        for r in tdb.get_list("work_types")
    ]
    locs = [
        {"id": r[0], "name": r[1], "lat": float(r[2]), "lon": float(r[3])}
        for r in tdb.get_list("locs")
    ]

    history_rows = tdb.get_list("history")
    history = []
    for r in history_rows:
        segs = []
        try:
            parsed = ast.literal_eval(r[24]) if r[24] else []
            if isinstance(parsed, list):
                segs = parsed
        except Exception:
            segs = []
        history.append(
            {
                "id": r[0],
                "wb_num": r[1],
                "date": r[2],
                "time_out": r[3],
                "time_in": r[4],
                "org": r[5],
                "client": r[6],
                "veh": r[7],
                "drv": r[8],
                "work_type": r[9] or "",
                "norm_km": float(r[10] or 0.0),
                "norm_mh": float(r[11] or 0.0),
                "odo_s": float(r[12] or 0.0),
                "odo_e": float(r[13] or 0.0),
                "moto_s": float(r[14] or 0.0),
                "moto_e": float(r[15] or 0.0),
                "moto_hours": float(r[16] or 0.0),
                "f_s": float(r[17] or 0.0),
                "f_a": float(r[18] or 0.0),
                "f_e": float(r[19] or 0.0),
                "f_norm_km": float(r[20] or 0.0),
                "f_norm_mh": float(r[21] or 0.0),
                "f_norm_total": float(r[22] or 0.0),
                "f_act": float(r[23] or 0.0),
                "route": r[24],
                "segments": segs,
                "dist": float(r[25] or 0.0),
                "form_type": r[26] if len(r) > 26 else "Форма № 412-АПК (МТЗ-82)",
            }
        )

    return {
        "orgs": orgs,
        "clients": clients,
        "staff": staff,
        "vehicles": vehicles,
        "drivers": drivers,
        "work_types": work_types,
        "locs": locs,
        "history": history,
        "memory": load_memory("memory_tractors.json"),
        "latestExcelPath": get_latest_excel_path(TRACTOR_OUT_DIR),
    }


def get_full_state(db: DatabaseManager, tdb: TractorDatabaseManager = None):
    if tdb is None:
        tdb = TractorDatabaseManager("tractors.db")
    ensure_orenburg_workspace(db)
    tractor_state = get_tractor_state(tdb)

    # Ensure unified backup files exist on disk so the user can download them anytime
    if not os.path.exists("Резервная_копия_Легковые_и_Трактора.zip") or not os.path.exists(
        "Единая_база_Легковые_и_Трактора.db"
    ):
        create_unified_backup()

    orgs = [{"id": r[0], "name": r[1]} for r in db.get_list("orgs")]
    clients = [{"id": r[0], "name": r[1]} for r in db.get_list("clients")]
    staff = [{"id": r[0], "fio": r[1], "role": r[2]} for r in db.get_list("staff")]
    vehicles = [
        {
            "id": r[0],
            "mark": r[1],
            "reg": r[2],
            "gar": r[3],
            "f_brand": r[4],
            "norm": float(r[5] or 0.0),
            "f_brand_gas": r[6] or "",
            "norm_gas": float(r[7] or 0.0),
        }
        for r in db.get_list("vehicles")
    ]
    drivers = [
        {
            "id": r[0],
            "fio": r[1],
            "cat": r[2],
            "exp": r[3],
            "tab": r[4],
            "ser": r[5],
            "num": r[6],
        }
        for r in db.get_list("drivers")
    ]
    locs = [
        {"id": r[0], "name": r[1], "lat": float(r[2]), "lon": float(r[3])}
        for r in db.get_list("locs")
    ]

    history_rows = db.get_list("history")
    history = []
    for r in history_rows:
        segs = []
        try:
            parsed = ast.literal_eval(r[19]) if r[19] else []
            if isinstance(parsed, list):
                segs = parsed
        except Exception:
            segs = []
        history.append(
            {
                "id": r[0],
                "wb_num": r[1],
                "date": r[2],
                "time_out": r[3],
                "time_in": r[4],
                "org": r[5],
                "client": r[6],
                "veh": r[7],
                "drv": r[8],
                "odo_s": float(r[9] or 0.0),
                "odo_e": float(r[10] or 0.0),
                "f_s": float(r[11] or 0.0),
                "f_a": float(r[12] or 0.0),
                "f_e": float(r[13] or 0.0),
                "f_norm": float(r[14] or 0.0),
                "g_s": float(r[15] or 0.0),
                "g_a": float(r[16] or 0.0),
                "g_e": float(r[17] or 0.0),
                "g_norm": float(r[18] or 0.0),
                "route": r[19],
                "segments": segs,
                "dist": float(r[20] or 0.0),
            }
        )

    return {
        "orgs": orgs,
        "clients": clients,
        "staff": staff,
        "vehicles": vehicles,
        "drivers": drivers,
        "locs": locs,
        "history": history,
        "memory": load_memory("memory.json"),
        "latestExcelPath": get_latest_excel_path(OUT_DIR),
        "tractors": tractor_state,
        "unifiedBackup": {
            "zipPath": "Резервная_копия_Легковые_и_Трактора.zip",
            "dbPath": "Единая_база_Легковые_и_Трактора.db",
        },
    }


def main():
    os.chdir(_BASE_DIR)
    action = sys.argv[1] if len(sys.argv) > 1 else "get_state"
    payload = json.loads(sys.argv[2]) if len(sys.argv) > 2 else {}
    db = DatabaseManager("waybills.db")
    tdb = TractorDatabaseManager("tractors.db")

    if action in ("get_state", "init_demo"):
        print(json.dumps(get_full_state(db, tdb), ensure_ascii=False))
        return

    if action == "create_unified_backup":
        res = create_unified_backup()
        print(
            json.dumps(
                {
                    "backup": res,
                    "state": get_full_state(db, tdb),
                },
                ensure_ascii=False,
            )
        )
        return

    if action == "save_memory":
        mode = payload.get("mode", "cars")
        data = {k: v for k, v in payload.items() if k != "mode"}
        if mode == "tractors":
            save_memory(data, "memory_tractors.json")
        else:
            save_memory(data, "memory.json")
        print(json.dumps(get_full_state(db, tdb), ensure_ascii=False))
        return

    if action == "crud":
        mode = payload.get("mode", "cars")
        target_db = tdb if mode == "tractors" else db
        op = payload.get("op")
        table = payload.get("table")

        if op == "add_driver":
            target_db.add_driver(
                payload.get("fio", ""),
                payload.get("cat", ""),
                payload.get("exp", ""),
                payload.get("tab", ""),
                payload.get("ser", ""),
                payload.get("num", ""),
            )
        elif op == "update_driver":
            target_db.update_driver(
                payload.get("id"),
                payload.get("fio", ""),
                payload.get("cat", ""),
                payload.get("exp", ""),
                payload.get("tab", ""),
                payload.get("ser", ""),
                payload.get("num", ""),
            )
        elif op == "add_vehicle":
            if mode == "tractors":
                tdb.add_vehicle(
                    payload.get("mark", ""),
                    payload.get("reg", ""),
                    payload.get("gar", ""),
                    payload.get("eq_type", "Трактор МТЗ-82"),
                    payload.get("f_brand", "ДТ (дизельное)"),
                    float(payload.get("norm_km") or 16.0),
                    float(payload.get("norm_mh") or 5.0),
                )
            else:
                db.query(
                    "INSERT INTO vehicles (mark, reg, gar, f_brand, norm, f_brand_gas, norm_gas) VALUES (?,?,?,?,?,?,?)",
                    (
                        payload.get("mark", ""),
                        payload.get("reg", ""),
                        payload.get("gar", ""),
                        payload.get("f_brand", "АИ-92"),
                        float(payload.get("norm") or 0.0),
                        payload.get("f_brand_gas", ""),
                        float(payload.get("norm_gas") or 0.0),
                    ),
                )
        elif op == "update_vehicle":
            if mode == "tractors":
                tdb.update_vehicle(
                    payload.get("id"),
                    payload.get("mark", ""),
                    payload.get("reg", ""),
                    payload.get("gar", ""),
                    payload.get("eq_type", "Трактор МТЗ-82"),
                    payload.get("f_brand", "ДТ (дизельное)"),
                    float(payload.get("norm_km") or 16.0),
                    float(payload.get("norm_mh") or 5.0),
                )
            else:
                db.update_vehicle(
                    payload.get("id"),
                    payload.get("mark", ""),
                    payload.get("reg", ""),
                    payload.get("gar", ""),
                    payload.get("f_brand", "АИ-92"),
                    float(payload.get("norm") or 0.0),
                    payload.get("f_brand_gas", ""),
                    float(payload.get("norm_gas") or 0.0),
                )
        elif op == "add_work_type" and mode == "tractors":
            tdb.add_work_type(
                payload.get("name", ""),
                float(payload.get("norm_mh") or 5.0),
                payload.get("unit", "м³"),
            )
        elif op == "update_work_type" and mode == "tractors":
            tdb.update_work_type(
                payload.get("id"),
                payload.get("name", ""),
                float(payload.get("norm_mh") or 5.0),
                payload.get("unit", "м³"),
            )
        elif op == "add_simple":
            if table in ("orgs", "clients"):
                target_db.query(f"INSERT INTO {table} (name) VALUES (?)", (payload.get("name", ""),))
        elif op == "add_staff":
            target_db.query(
                "INSERT INTO staff (fio, role) VALUES (?,?)",
                (payload.get("fio", ""), payload.get("role", "Механик")),
            )
        elif op == "add_loc":
            name = (payload.get("name") or "").strip()
            lat = float(payload.get("lat") or 51.766)
            lon = float(payload.get("lon") or 55.1)
            if name:
                existing = target_db.query("SELECT id FROM locs WHERE LOWER(name)=LOWER(?)", (name,))
                if existing:
                    target_db.query("UPDATE locs SET lat=?, lon=? WHERE id=?", (lat, lon, existing[0][0]))
                else:
                    target_db.query("INSERT INTO locs (name, lat, lon) VALUES (?,?,?)", (name, lat, lon))
        elif op == "delete":
            allowed_tables = ("drivers", "vehicles", "orgs", "clients", "staff", "locs", "history", "work_types")
            if table in allowed_tables:
                target_db.query(f"DELETE FROM {table} WHERE id=?", (payload.get("id"),))

        create_unified_backup()
        print(json.dumps(get_full_state(db, tdb), ensure_ascii=False))
        return

    if action == "save_waybill":
        mode = payload.get("mode", "cars")

        if mode == "tractors":
            wb_num = str(payload.get("wb_num", "Т-001")).strip() or "Т-001"
            d_str = str(payload.get("date", "")).strip()
            t_out = str(payload.get("time_out", "08:00")).strip()
            t_in = str(payload.get("time_in", "17:00")).strip()
            org = str(payload.get("org", "")).strip()
            client = str(payload.get("client", "")).strip()
            pickup = str(payload.get("pickup", "")).strip()
            veh = str(payload.get("veh", "")).strip()
            drv = str(payload.get("drv", "")).strip()
            mech = str(payload.get("mech", "")).strip()
            disp = str(payload.get("disp", "")).strip()
            work_type = str(payload.get("work_type", "Экскавация грунта I-II кат. (обратная лопата)")).strip()
            form_type = str(payload.get("form_type", "Форма № 412-АПК (МТЗ-82)")).strip()

            norm_km = float(payload.get("norm_km") or 16.0)
            norm_mh = float(payload.get("norm_mh") or 5.0)
            odo_s = float(payload.get("odo_s") or 0.0)
            dist = float(payload.get("dist") or 0.0)
            odo_e = float(payload.get("odo_e") or (odo_s + dist))

            moto_s = float(payload.get("moto_s") or 0.0)
            moto_hours = float(payload.get("moto_hours") or 0.0)
            moto_e = float(payload.get("moto_e") or (moto_s + moto_hours))

            extra_coef = float(payload.get("extra_coef") or 0.0)
            calc_km, calc_mh, calc_tot = calculate_tractor_diesel_consumption(
                norm_km, dist, norm_mh, moto_hours, extra_coef
            )
            f_norm_km = float(payload.get("f_norm_km") if payload.get("f_norm_km") is not None else calc_km)
            f_norm_mh = float(payload.get("f_norm_mh") if payload.get("f_norm_mh") is not None else calc_mh)
            f_norm_total = float(payload.get("f_norm_total") if payload.get("f_norm_total") is not None else calc_tot)

            f_s = float(payload.get("f_s") or 0.0)
            f_a = float(payload.get("f_a") or 0.0)
            f_act = float(payload.get("f_act") if payload.get("f_act") is not None else f_norm_total)
            f_e = float(payload.get("f_e") if payload.get("f_e") is not None else round(f_s + f_a - f_act, 2))

            segments = payload.get("segments", []) or []
            route_repr = repr(segments)
            editing_id = payload.get("editing_id")

            if editing_id:
                tdb.update_wb(
                    editing_id,
                    wb_num,
                    d_str,
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
                    route_repr,
                    dist,
                    form_type,
                )
            else:
                tdb.save_wb(
                    wb_num,
                    d_str,
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
                    route_repr,
                    dist,
                    form_type,
                )

            if veh:
                tdb.recalculate_chain(veh)

            full_state = get_full_state(db, tdb)
            t_state = full_state["tractors"]
            drv_row = next((d for d in t_state["drivers"] if d["fio"] == drv), {})
            veh_row = next((v for v in t_state["vehicles"] if v["mark"] == veh), {})

            excel_data = {
                "wb_num": wb_num,
                "date": d_str,
                "ru_date": format_ru_date(d_str),
                "form_type": form_type,
                "eq_type": veh_row.get("eq_type", "Трактор МТЗ-82"),
                "org": org,
                "okpo": str(payload.get("okpo", "56104921")).strip(),
                "client": client,
                "pickup": pickup or (segments[0].get("from", "") if segments else ""),
                "drv": drv,
                "tab": drv_row.get("tab", ""),
                "ser": drv_row.get("ser", ""),
                "num": drv_row.get("num", ""),
                "cat": drv_row.get("cat", ""),
                "car": veh,
                "reg": veh_row.get("reg", ""),
                "gar": veh_row.get("gar", ""),
                "mech": mech,
                "disp": disp,
                "work_type": work_type,
                "trailer": str(payload.get("trailer", "Обратная лопата / Отвал МТЗ-82")).strip(),
                "msg_type": str(payload.get("msg_type", "перегон своим ходом / работа на объекте")).strip(),
                "trans_type": str(payload.get("trans_type", work_type)).strip(),
                "odo_s": round(odo_s, 1),
                "odo_e": round(odo_e, 1),
                "dist": round(dist, 1),
                "moto_s": round(moto_s, 1),
                "moto_e": round(moto_e, 1),
                "moto_hours": round(moto_hours, 1),
                "norm_km": round(norm_km, 2),
                "norm_mh": round(norm_mh, 2),
                "f_norm_km": round(f_norm_km, 2),
                "f_norm_mh": round(f_norm_mh, 2),
                "f_norm_total": round(f_norm_total, 2),
                "f_act": round(f_act, 2),
                "f_brand": veh_row.get("f_brand", "ДТ (дизельное)"),
                "f_s": round(f_s, 2),
                "f_a": round(f_a, 2),
                "f_e": round(f_e, 2),
                "time_out": t_out,
                "time_in": t_in,
                "time_tot": payload.get("time_tot", "9.0"),
                "segments": segments,
            }

            excel_rel_path = None
            excel_error = None
            if payload.get("create_excel", True):
                os.makedirs(TRACTOR_OUT_DIR, exist_ok=True)
                safe_veh = (
                    (veh or "МТЗ-82")
                    .replace(" ", "_")
                    .replace("/", "_")
                    .replace("(", "")
                    .replace(")", "")
                )
                out_rel = f"{TRACTOR_OUT_DIR}/ПЛ_{wb_num}_{d_str}_{safe_veh}.xlsx"
                ok, msg = generate_tractor_wb_excel(out_rel, excel_data)
                if ok:
                    excel_rel_path = out_rel
                else:
                    excel_error = msg

            create_unified_backup()
            print(
                json.dumps(
                    {
                        "state": get_full_state(db, tdb),
                        "excelRelPath": excel_rel_path,
                        "excelError": excel_error,
                        "previewData": excel_data,
                    },
                    ensure_ascii=False,
                )
            )
            return

        # Режим легковых автомобилей (Форма № 3 — waybills.db)
        wb_num = str(payload.get("wb_num", "1")).strip() or "1"
        d_str = str(payload.get("date", "")).strip()
        t_out = str(payload.get("time_out", "08:00")).strip()
        t_in = str(payload.get("time_in", "17:00")).strip()
        org = str(payload.get("org", "")).strip()
        client = str(payload.get("client", "")).strip()
        pickup = str(payload.get("pickup", "")).strip()
        veh = str(payload.get("veh", "")).strip()
        drv = str(payload.get("drv", "")).strip()
        mech = str(payload.get("mech", "")).strip()
        disp = str(payload.get("disp", "")).strip()
        odo_s = float(payload.get("odo_s") or 0.0)
        odo_e = float(payload.get("odo_e") or 0.0)
        f_s = float(payload.get("f_s") or 0.0)
        f_a = float(payload.get("f_a") or 0.0)
        f_e = float(payload.get("f_e") or 0.0)
        f_n = float(payload.get("f_norm") or 0.0)
        g_s = float(payload.get("g_s") or 0.0)
        g_a = float(payload.get("g_a") or 0.0)
        g_e = float(payload.get("g_e") or 0.0)
        g_n = float(payload.get("g_norm") or 0.0)
        dist = float(payload.get("dist") or 0.0)
        segments = payload.get("segments", []) or []
        route_repr = repr(segments)
        editing_id = payload.get("editing_id")

        if editing_id:
            db.update_wb(
                editing_id,
                wb_num,
                d_str,
                t_out,
                t_in,
                org,
                client,
                veh,
                drv,
                odo_s,
                odo_e,
                f_s,
                f_a,
                f_e,
                f_n,
                g_s,
                g_a,
                g_e,
                g_n,
                route_repr,
                dist,
            )
        else:
            db.save_wb(
                wb_num,
                d_str,
                t_out,
                t_in,
                org,
                client,
                veh,
                drv,
                odo_s,
                odo_e,
                f_s,
                f_a,
                f_e,
                f_n,
                g_s,
                g_a,
                g_e,
                g_n,
                route_repr,
                dist,
            )

        if veh:
            db.recalculate_chain(veh)

        state = get_full_state(db, tdb)
        drv_row = next((d for d in state["drivers"] if d["fio"] == drv), {})
        veh_row = next((v for v in state["vehicles"] if v["mark"] == veh), {})

        fuel_type = payload.get("fuel_type", "Бензин")
        if fuel_type == "Газ":
            brand_str = veh_row.get("f_brand_gas") or "ПБА"
            f_s_str, f_a_str, f_e_str = str(g_s), str(g_a), str(g_e)
            spent_str = str(g_n)
            act_str = str(payload.get("g_act") or spent_str)
        elif fuel_type == "Бензин + Газ":
            b1 = veh_row.get("f_brand") or "АИ-92"
            b2 = veh_row.get("f_brand_gas") or "ПБА"
            brand_str = f"{b1} / {b2}"
            f_s_str = f"{f_s} / {g_s}"
            f_a_str = f"{f_a} / {g_a}"
            f_e_str = f"{f_e} / {g_e}"
            spent_str = f"{f_n} / {g_n}"
            act_str = spent_str
        else:
            brand_str = veh_row.get("f_brand") or "АИ-92"
            f_s_str, f_a_str, f_e_str = str(f_s), str(f_a), str(f_e)
            spent_str = str(f_n)
            act_str = str(payload.get("f_act") or spent_str)

        excel_data = {
            "wb_num": wb_num,
            "date": d_str,
            "ru_date": format_ru_date(d_str),
            "org": org,
            "okpo": str(payload.get("okpo", "48291034")).strip(),
            "client": client,
            "pickup": pickup or (segments[0].get("from", "") if segments else ""),
            "drv": drv,
            "tab": drv_row.get("tab", ""),
            "ser": drv_row.get("ser", ""),
            "num": drv_row.get("num", ""),
            "cat": drv_row.get("cat", ""),
            "car": veh,
            "reg": veh_row.get("reg", ""),
            "gar": veh_row.get("gar", ""),
            "mech": mech,
            "disp": disp,
            "trailer": str(payload.get("trailer", "без прицепа")).strip(),
            "msg_type": str(payload.get("msg_type", "городское / пригородное")).strip(),
            "trans_type": str(payload.get("trans_type", "служебная")).strip(),
            "odo_s": int(round(odo_s)),
            "odo_e": int(round(odo_e)),
            "time_out": t_out,
            "time_in": t_in,
            "time_tot": payload.get("time_tot", "9.0"),
            "f_brand": brand_str,
            "f_s": f_s_str,
            "f_a": f_a_str,
            "f_e": f_e_str,
            "spent": spent_str,
            "f_act": act_str,
            "dist": dist,
            "seg_from": [s.get("from", "") for s in segments],
            "seg_to": [s.get("to", "") for s in segments],
            "seg_t_out": [s.get("t_out", t_out) for s in segments],
            "seg_t_in": [s.get("t_in", t_in) for s in segments],
            "seg_km": [s.get("km", 0.0) for s in segments],
        }

        excel_rel_path = None
        excel_error = None
        if payload.get("create_excel", True):
            os.makedirs(OUT_DIR, exist_ok=True)
            safe_veh = (veh or "Авто").replace(" ", "_").replace("/", "_")
            out_rel = f"{OUT_DIR}/ПЛ_{wb_num}_{d_str}_{safe_veh}.xlsx"
            ok, msg = generate_f3_from_excel(
                "templates/put-list-f3.xlsx", out_rel, excel_data
            )
            if ok:
                excel_rel_path = out_rel
            else:
                excel_error = msg

        create_unified_backup()
        print(
            json.dumps(
                {
                    "state": get_full_state(db, tdb),
                    "excelRelPath": excel_rel_path,
                    "excelError": excel_error,
                    "previewData": excel_data,
                },
                ensure_ascii=False,
            )
        )
        return


if __name__ == "__main__":
    main()
