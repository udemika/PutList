import os
import sqlite3
import sys

_INTERNAL_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_internal")
if os.path.isdir(_INTERNAL_DIR) and _INTERNAL_DIR not in sys.path:
    sys.path.insert(0, _INTERNAL_DIR)

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def safe_float(val):
    """Безопасное приведение к числу для корректной работы формул и форматов Excel"""
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def export_history_to_excel(history_data, out_path="Реестр_Путевых_Листов.xlsx"):
    """Генерирует коммерческий аналитический реестр всех путевых листов (легковые авто и трактора МТЗ-82)"""
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Реестр Путевых Листов"

        font_h = Font(name="Helvetica", size=10, bold=True, color="FFFFFF")
        font_b = Font(name="Helvetica", size=10, color="1D1D1F")
        fill_h = PatternFill(start_color="007AFF", end_color="007AFF", fill_type="solid")
        fill_zebra = PatternFill(start_color="F5F5F7", end_color="F5F5F7", fill_type="solid")
        border_cell = Border(
            left=Side(border_style="thin", color="E5E5EA"),
            right=Side(border_style="thin", color="E5E5EA"),
            top=Side(border_style="thin", color="E5E5EA"),
            bottom=Side(border_style="thin", color="E5E5EA"),
        )

        is_tractor = bool(history_data and len(history_data[0]) >= 25)
        if is_tractor:
            headers = [
                "ID",
                "№ Путевки",
                "Дата",
                "Выезд",
                "Возврат",
                "Организация",
                "Заказчик",
                "Трактор / Экскаватор",
                "Тракторист-машинист",
                "Вид работ",
                "Одометр Старт",
                "Одометр Конец",
                "Пробег (км)",
                "Моточасы Старт",
                "Моточасы Конец",
                "Отработано (м/ч)",
                "Остаток ДТ (л)",
                "Выдано ДТ (л)",
                "Остаток ДТ Кон. (л)",
                "Расход ДТ Норма (л)",
                "Расход ДТ Факт (л)",
            ]
        else:
            headers = [
                "ID",
                "№ Путевки",
                "Дата",
                "Выезд",
                "Возврат",
                "Организация",
                "Заказчик",
                "Автомобиль",
                "Водитель",
                "Одометр Старт",
                "Одометр Конец",
                "Дистанция (км)",
                "Остаток Б (л)",
                "Выдано Б",
                "Конец Б",
                "Расход Б Норма",
                "Остаток Г (л)",
                "Выдано Г",
                "Конец Г",
                "Расход Г Норма",
            ]
        ws.append(headers)
        ws.row_dimensions[1].height = 28

        for col_num, _ in enumerate(headers, 1):
            c = ws.cell(row=1, column=col_num)
            c.font = font_h
            c.fill = fill_h
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row_idx, r in enumerate(history_data, start=2):
            if is_tractor:
                row_data = [
                    r[0],
                    r[1],
                    r[2],
                    r[3],
                    r[4],
                    r[5],
                    r[6],
                    r[7],
                    r[8],
                    r[9],
                    safe_float(r[12]),
                    safe_float(r[13]),
                    safe_float(r[25]),
                    safe_float(r[14]),
                    safe_float(r[15]),
                    safe_float(r[16]),
                    safe_float(r[17]),
                    safe_float(r[18]),
                    safe_float(r[19]),
                    safe_float(r[22]),
                    safe_float(r[23]),
                ]
            else:
                dist_val = safe_float(r[20]) if len(r) > 20 else round(safe_float(r[10]) - safe_float(r[9]), 1)
                row_data = [
                    r[0],
                    r[1],
                    r[2],
                    r[3],
                    r[4],
                    r[5],
                    r[6],
                    r[7],
                    r[8],
                    safe_float(r[9]),
                    safe_float(r[10]),
                    dist_val,
                    safe_float(r[11]),
                    safe_float(r[12]),
                    safe_float(r[13]),
                    safe_float(r[14]),
                    safe_float(r[15]),
                    safe_float(r[16]),
                    safe_float(r[17]),
                    safe_float(r[18]),
                ]
            ws.append(row_data)
            ws.row_dimensions[row_idx].height = 22
            is_zebra = row_idx % 2 == 0
            for col_num in range(1, len(row_data) + 1):
                c = ws.cell(row=row_idx, column=col_num)
                c.font = font_b
                c.border = border_cell
                if is_zebra:
                    c.fill = fill_zebra
                if isinstance(c.value, float):
                    c.number_format = "#,##0.0"

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

        wb.save(out_path)
        return (True, out_path)
    except Exception as e:
        return (False, str(e))


def generate_word_report(history_data, out_path="Журнал_Учета_Путевых_Листов_Форма8.xlsx"):
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Форма № 8"
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

        is_tractor = bool(history_data and len(history_data[0]) >= 25)
        db_file = "tractors.db" if is_tractor else "waybills.db"

        vehicles_dict = {}
        drivers_dict = {}
        try:
            conn = sqlite3.connect(db_file)
            cur = conn.cursor()
            for v in cur.execute("SELECT * FROM vehicles").fetchall():
                vehicles_dict[v[1]] = v
            for d in cur.execute("SELECT * FROM drivers").fetchall():
                drivers_dict[d[1]] = d
            conn.close()
        except Exception:
            pass

        ws.merge_cells("A1:K1")
        ws["A1"] = "Типовая межотраслевая форма № 8"
        ws["A1"].font = Font(name="Arial", size=9, italic=True)
        ws["A1"].alignment = Alignment(horizontal="right")

        ws.merge_cells("A2:K2")
        ws["A2"] = "ЖУРНАЛ УЧЕТА ДВИЖЕНИЯ ПУТЕВЫХ ЛИСТОВ"
        ws["A2"].font = Font(name="Arial", size=12, bold=True)
        ws["A2"].alignment = Alignment(horizontal="center")

        headers = [
            "№ путевого листа",
            "Дата выдачи",
            "Водитель / Тракторист (ФИО)",
            "Табельный номер",
            "Марка ТС / Трактора",
            "Гос. рег. знак",
            "Гаражный номер",
            "Время выезда",
            "Время возвращения",
            "Тип ГСМ",
            "Примечание / Вид работ",
        ]
        ws.append([])
        ws.append(headers)

        border_thin = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        for idx, h in enumerate(headers, 1):
            c = ws.cell(row=4, column=idx)
            c.font = Font(name="Arial", size=9, bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border_thin

        for r_idx, r in enumerate(history_data, start=5):
            car_mark = r[7]
            drv_fio = r[8]
            v_info = vehicles_dict.get(car_mark)
            d_info = drivers_dict.get(drv_fio)
            car_reg = v_info[2] if v_info else ""
            car_gar = v_info[3] if v_info else ""
            drv_tab = d_info[4] if d_info else ""

            if is_tractor:
                fuel_label = "ДТ (дизельное)"
                note_str = f"{r[9]} ({r[16]} м/ч)"
            else:
                f_norm = safe_float(r[14])
                g_norm = safe_float(r[18])
                if f_norm > 0 and g_norm > 0:
                    fuel_label = "Бенз/Газ"
                elif g_norm > 0:
                    fuel_label = "Газ"
                else:
                    fuel_label = "Бензин"
                note_str = ""

            row_vals = [
                r[1],
                r[2],
                drv_fio,
                drv_tab,
                car_mark,
                car_reg,
                car_gar,
                r[3],
                r[4],
                fuel_label,
                note_str,
            ]
            ws.append(row_vals)
            for col_i in range(1, 12):
                c = ws.cell(row=r_idx, column=col_i)
                c.font = Font(name="Arial", size=9)
                c.border = border_thin

        wb.save(out_path)
        return (True, out_path)
    except Exception as e:
        return (False, str(e))

