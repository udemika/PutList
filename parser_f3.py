import io
import os
import re
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

# Автоматическое подключение локальной папки _internal (где установлен openpyxl 3.1.5)
_INTERNAL_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_internal")
if os.path.isdir(_INTERNAL_DIR) and _INTERNAL_DIR not in sys.path:
    sys.path.insert(0, _INTERNAL_DIR)

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

_TEMPLATE_BYTES_CACHE = {}


def _load_workbook_cached(tpl_path):
    """Загружает .xlsx шаблон из кеша байтов в памяти (без повторного чтения с диска)."""
    abs_p = os.path.abspath(tpl_path)
    mtime = os.path.getmtime(abs_p)
    cached = _TEMPLATE_BYTES_CACHE.get(abs_p)
    if not cached or cached[0] != mtime:
        with open(abs_p, "rb") as f:
            raw = f.read()
        _TEMPLATE_BYTES_CACHE[abs_p] = (mtime, raw)
    else:
        raw = cached[1]
    return load_workbook(io.BytesIO(raw))


def _build_merged_index(sheet):
    """Строит O(1)-индекс объединённых ячеек {(row, col): (min_row, min_col, max_row, max_col)}."""
    idx = {}
    for mr in sheet.merged_cells.ranges:
        bounds = (mr.min_row, mr.min_col, mr.max_row, mr.max_col)
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                idx[(r, c)] = bounds
    return idx


def get_master(sheet, cell, merged_index=None):
    if merged_index is not None:
        bounds = merged_index.get((cell.row, cell.column))
        if bounds is not None:
            return sheet.cell(row=bounds[0], column=bounds[1])
        return cell
    for m in sheet.merged_cells.ranges:
        if cell.coordinate in m:
            return sheet.cell(row=m.min_row, column=m.min_col)
    return cell


def set_cell_master(sheet, coord, val, merged_index=None):
    """Безопасная запись значения в главную ячейку объединённого диапазона по координатам."""
    if val is None:
        return
    val_str = str(val).strip()
    if val_str == "":
        return
    get_master(sheet, sheet[coord], merged_index=merged_index).value = val_str


def split_hh_mm(time_str):
    """Разделяет строку времени '08:30' на часы ('08') и минуты ('30') для граф ч. и мин. на Листе 2."""
    s = str(time_str or "").strip()
    if ":" in s:
        parts = s.split(":", 1)
        return (parts[0].strip() or "00", parts[1].strip() or "00")
    return (s or "00", "00")


def write_right(sheet, lbl_cell, val, skip=0):
    if val is None or str(val).strip() == "":
        return
    r = lbl_cell.row
    c = lbl_cell.column
    for m in sheet.merged_cells.ranges:
        if lbl_cell.coordinate in m:
            c = m.max_col
            break

    found = 0
    skip_cols_until = -1
    for col in range(c + 1, sheet.max_column + 1):
        if col <= skip_cols_until:
            continue
        target = sheet.cell(row=r, column=col)
        is_merged = False
        for m in sheet.merged_cells.ranges:
            if target.coordinate in m:
                is_merged = True
                skip_cols_until = m.max_col
                break
        b = target.border
        is_line = bool(b and b.bottom and b.bottom.style)
        if is_merged or is_line:
            if found == skip:
                get_master(sheet, target).value = str(val)
                return
            found += 1


def write_below(sheet, lbl_cell, val):
    """Функция для записи строго ПОД указанной ячейкой (для марки ГСМ)"""
    if val is None or str(val).strip() == "":
        return
    r = lbl_cell.row + 1
    c = lbl_cell.column
    for m in sheet.merged_cells.ranges:
        if lbl_cell.coordinate in m:
            r = m.max_row + 1
            break
    get_master(sheet, sheet.cell(row=r, column=c)).value = str(val)


def write_above(sheet, lbl_cell, val):
    if val is None or str(val).strip() == "":
        return
    r = lbl_cell.row - 1
    c = lbl_cell.column
    get_master(sheet, sheet.cell(row=r, column=c)).value = str(val)


def write_list(sheet, lbl_cell, vals):
    if not vals:
        return
    r = lbl_cell.row + 1
    c = lbl_cell.column
    while r <= sheet.max_row:
        c_val = get_master(sheet, sheet.cell(row=r, column=c)).value
        if c_val is None or str(c_val).strip() == "":
            break
        r += 1

    for val in vals:
        if r > sheet.max_row:
            break
        step = 1
        for m in sheet.merged_cells.ranges:
            if sheet.cell(row=r, column=c).coordinate in m:
                step = m.max_row - m.min_row + 1
                break
        if val is not None and str(val).strip() != "":
            get_master(sheet, sheet.cell(row=r, column=c)).value = str(val)
        r += step


def format_ru_date(d_str):
    try:
        d, m, y = d_str.split(".")
        mos = {
            "01": "января",
            "02": "февраля",
            "03": "марта",
            "04": "апреля",
            "05": "мая",
            "06": "июня",
            "07": "июля",
            "08": "августа",
            "09": "сентября",
            "10": "октября",
            "11": "ноября",
            "12": "декабря",
        }
        return f"{d} {mos.get(m, '')} {y} г."
    except Exception:
        return d_str


def clean_fio(v, drv):
    if not isinstance(v, str):
        return v
    v = re.sub(r"[Аа]ксарин\s*[А-ЯA-Z]\.?[IИ]?\.?", "", v)
    v = re.sub(r"ксарин\s*", "", v)
    v = re.sub(r"\s+", " ", v).strip()
    return f"{v} {drv}".strip()


def _split_coord(coord):
    m = re.match(r"^([A-Z]+)(\d+)$", coord)
    if not m:
        return ("A", 1)
    return (m.group(1), int(m.group(2)))


def _apply_template_zip_fallback(tpl_path, out_file, sheet1_updates, sheet2_updates, drv=""):
    """Резервная запись в официальный шаблон .xlsx через стандартную библиотеку zipfile + xml."""
    os.makedirs(os.path.dirname(os.path.abspath(out_file)), exist_ok=True)
    NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ET.register_namespace("", NS_MAIN)

    def _apply_inline_strings(xml_bytes, cell_map):
        root = ET.fromstring(xml_bytes)
        sheet_data = root.find(f"{{{NS_MAIN}}}sheetData")
        if sheet_data is None:
            return xml_bytes
        rows_by_num = {}
        for r_el in sheet_data.findall(f"{{{NS_MAIN}}}row"):
            r_attr = r_el.attrib.get("r")
            if r_attr and r_attr.isdigit():
                rows_by_num[int(r_attr)] = r_el

        for coord, val_str in cell_map.items():
            if val_str is None or str(val_str).strip() == "":
                continue
            _col_s, r_num = _split_coord(coord)
            r_el = rows_by_num.get(r_num)
            if r_el is None:
                r_el = ET.SubElement(sheet_data, f"{{{NS_MAIN}}}row", {"r": str(r_num)})
                rows_by_num[r_num] = r_el
            target_c = None
            for c_el in r_el.findall(f"{{{NS_MAIN}}}c"):
                if c_el.attrib.get("r") == coord:
                    target_c = c_el
                    break
            if target_c is None:
                target_c = ET.SubElement(r_el, f"{{{NS_MAIN}}}c", {"r": coord})
            target_c.attrib["t"] = "inlineStr"
            for child in list(target_c):
                target_c.remove(child)
            is_el = ET.SubElement(target_c, f"{{{NS_MAIN}}}is")
            t_el = ET.SubElement(is_el, f"{{{NS_MAIN}}}t")
            t_el.text = str(val_str)

        return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(tpl_path, "r") as zin, zipfile.ZipFile(out_file, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            raw = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                raw = _apply_inline_strings(raw, sheet1_updates)
            elif item.filename == "xl/worksheets/sheet2.xml":
                raw = _apply_inline_strings(raw, sheet2_updates)
            elif item.filename == "xl/sharedStrings.xml" and drv:
                txt_xml = raw.decode("utf-8", errors="ignore")
                txt_xml = re.sub(r"[Аа]ксарин\s*[А-ЯA-Z]\.?[IИ]?\.?", drv, txt_xml)
                txt_xml = txt_xml.replace("{ФамилияИО}", drv).replace("{фамилияио}", drv)
                raw = txt_xml.encode("utf-8")
            zout.writestr(item, raw)

    return (True, out_file)


def _build_f3_cell_maps(data):
    """Формирует точные карты координат всех ячеек Листа 1 и Листа 2 для Формы № 3 (put-list-f3.xlsx)."""
    ru_date = format_ru_date(data.get("date", ""))
    drv = str(data.get("drv", "")).strip()
    date_raw = str(data.get("date", "")).strip()
    t_out = str(data.get("time_out", "08:00")).strip()
    t_in = str(data.get("time_in", "17:00")).strip()
    mech = str(data.get("mech", "")).strip()
    disp = str(data.get("disp", "")).strip()

    s1_map = {
        "AZ4": str(data.get("wb_ser", "")),
        "BX4": str(data.get("wb_num", "")),
        "L5": "от",
        "M5": str(ru_date),
        "Q8": str(data.get("org", "")),
        "BP8": str(data.get("okpo", "")),
        "T10": str(data.get("car", "")),
        "AD11": str(data.get("reg", "")),
        "BP11": str(data.get("gar", "")),
        "M12": drv,
        "BP12": str(data.get("tab", "")),
        "S14": f"{data.get('ser', '')} {data.get('num', '')}".strip(),
        "BO14": str(data.get("cat", "")),
        "Z15": str(data.get("lic_card", "стандартная")),
        "BO15": str(data.get("snils", "")),
        "AR17": str(data.get("ser", "")),
        "BI17": str(data.get("num", "")),
        "M18": str(data.get("trailer", "без прицепа")),
        "BU19": str(data.get("odo_s", "")),
        "Q20": str(data.get("client", "")),
        "A22": str(data.get("client_org", data.get("client", ""))),
        "BN22": mech,
        "R26": str(data.get("pickup", "")),
        "BP26": drv,
        "BF29": str(data.get("f_brand", "")),
        "BT29": str(data.get("f_code", "")),
        "AE30": t_out,
        "AD32": disp,
        "BT34": str(data.get("f_a", "")),
        "AE35": t_in,
        "AD36": disp,
        "BT37": str(data.get("f_s", "")),
        "BT38": str(data.get("f_e", "")),
        "BT39": str(data.get("spent", "")),
        "BT40": str(data.get("f_act", "")),
        "BT41": str(data.get("f_econ", "")),
        "BT42": str(data.get("f_over", "")),
        "Q42": str(data.get("msg_type", "городское / пригородное")),
        "Q43": str(data.get("trans_type", "служебная")),
        "AD45": drv,
        "BT45": str(data.get("odo_e", "")),
        "BN46": mech,
        "B49": date_raw,
        "Q49": t_out,
        "AT49": date_raw,
        "BI49": t_in,
        "AG54": f"{date_raw} {t_out}".strip(),
        "BB54": mech,
    }

    s2_map = {
        "E33": str(data.get("time_tot", "")),
        "E35": str(data.get("dist", "")),
        "K36": drv,
        "D39": "Диспетчер",
        "K39": disp,
    }

    seg_from = data.get("seg_from", []) or []
    seg_to = data.get("seg_to", []) or []
    seg_t_out = data.get("seg_t_out", []) or []
    seg_t_in = data.get("seg_t_in", []) or []
    seg_km = data.get("seg_km", []) or []

    if not seg_from and data.get("segments"):
        segs = data.get("segments") or []
        seg_from = [s.get("from", "") for s in segs]
        seg_to = [s.get("to", "") for s in segs]
        seg_t_out = [s.get("t_out", t_out) for s in segs]
        seg_t_in = [s.get("t_in", t_in) for s in segs]
        seg_km = [s.get("km", 0.0) for s in segs]

    for idx in range(min(len(seg_from), 26)):
        r = 5 + idx
        s2_map[f"B{r}"] = str(idx + 1)
        s2_map[f"C{r}"] = str(data.get("client_code", "01"))
        s2_map[f"E{r}"] = str(seg_from[idx])
        if idx < len(seg_to):
            s2_map[f"H{r}"] = str(seg_to[idx])
        if idx < len(seg_t_out):
            h_o, m_o = split_hh_mm(seg_t_out[idx])
            s2_map[f"J{r}"] = h_o
            s2_map[f"L{r}"] = m_o
        if idx < len(seg_t_in):
            h_i, m_i = split_hh_mm(seg_t_in[idx])
            s2_map[f"M{r}"] = h_i
            s2_map[f"N{r}"] = m_i
        if idx < len(seg_km):
            s2_map[f"O{r}"] = str(seg_km[idx])

    return s1_map, s2_map


def _build_tractor_cell_maps(data):
    """Формирует точные карты координат всех ячеек Листа 1 и Листа 2 для бланков трактора 412-АПК и экскаватора ЭСМ-2."""
    ru_date = format_ru_date(data.get("date", ""))
    drv = str(data.get("drv", "")).strip()
    date_raw = str(data.get("date", "")).strip()
    t_out = str(data.get("time_out", "08:00")).strip()
    t_in = str(data.get("time_in", "17:00")).strip()
    mech = str(data.get("mech", "")).strip()
    disp = str(data.get("disp", "")).strip()
    work_type = str(data.get("work_type", "")).strip()
    try:
        dist_f = float(data.get("dist", 0.0) or 0.0)
    except Exception:
        dist_f = 0.0
    try:
        norm_km_f = float(data.get("norm_km", 0.0) or 0.0)
    except Exception:
        norm_km_f = 0.0
    try:
        f_norm_km_f = float(data.get("f_norm_km", 0.0) or 0.0)
    except Exception:
        f_norm_km_f = 0.0
    try:
        f_norm_mh_f = float(data.get("f_norm_mh", 0.0) or 0.0)
    except Exception:
        f_norm_mh_f = 0.0
    try:
        spent_total_f = float(data.get("f_norm_total", data.get("spent", 0.0)) or 0.0)
    except Exception:
        spent_total_f = 0.0
    try:
        f_act_f = float(data.get("f_act", spent_total_f) or 0.0)
    except Exception:
        f_act_f = spent_total_f
    try:
        f_s_f = float(data.get("f_s", 0.0) or 0.0)
        f_a_f = float(data.get("f_a", 0.0) or 0.0)
        f_e_f = float(data.get("f_e", 0.0) or 0.0)
    except Exception:
        f_s_f, f_a_f, f_e_f = 0.0, 0.0, 0.0

    only_mh = bool(data.get("only_motohours", False))

    # Защита от 0.0 л за перегон при наличии пробега и нормы из справочника (если не включён режим «только моточасы»)
    if (not only_mh) and dist_f > 0.0 and norm_km_f > 0.0 and f_norm_km_f <= 0.001:
        f_norm_km_f = round((dist_f / 100.0) * norm_km_f, 2)
        old_total = spent_total_f
        spent_total_f = round(f_norm_km_f + f_norm_mh_f, 2)
        if abs(f_act_f - old_total) < 0.05 or f_act_f <= 0.001:
            f_act_f = spent_total_f
            f_e_f = round(f_s_f + f_a_f - f_act_f, 2)
    elif only_mh:
        f_norm_km_f = 0.0

    spent_total = round(spent_total_f, 2)
    f_act = round(f_act_f, 2)
    f_norm_km_val = round(f_norm_km_f, 2)
    f_norm_mh_val = round(f_norm_mh_f, 2)
    f_e_val = round(f_e_f, 2)

    odo_s = data.get("odo_s", 0)
    odo_e = data.get("odo_e", 0)
    moto_s = data.get("moto_s", 0)
    moto_e = data.get("moto_e", 0)
    dist = data.get("dist", 0)
    moto_hours = data.get("moto_hours", 0)

    works = data.get("works") or []
    if works:
        work_summary = "; ".join(
            f"{w.get('name', '')} ({w.get('mh', 0)} м/ч)"
            for w in works
            if str(w.get("name", "")).strip()
        )
        if work_summary:
            work_type = work_summary
        norms_set = []
        for w in works:
            nm_s = str(w.get("norm_mh", "")).strip()
            if nm_s and nm_s not in norms_set:
                norms_set.append(nm_s)
        norms_mh_str = "/".join(norms_set) if norms_set else str(data.get("norm_mh", 5.2))
    else:
        norms_mh_str = str(data.get("norm_mh", 5.2))

    s1_map = {
        "AZ4": str(data.get("wb_ser", "МТЗ")),
        "BX4": str(data.get("wb_num", "Т-001")),
        "L5": "от",
        "M5": str(ru_date),
        "Q8": str(data.get("org", "")),
        "BP8": str(data.get("okpo", "")),
        "T10": str(data.get("car", "")),
        "AD11": str(data.get("reg", "")),
        "BP11": str(data.get("gar", "")),
        "M12": drv,
        "BP12": str(data.get("tab", "")),
        "S14": f"{data.get('ser', '')} {data.get('num', '')}".strip(),
        "BO14": str(data.get("cat", "")),
        "Z15": work_type or "Навесное оборудование МТЗ-82",
        "BO15": str(data.get("snils", "")),
        "AR17": str(data.get("ser", "")),
        "BI17": str(data.get("num", "")),
        "M18": str(data.get("trailer", "Ковш / Отвал МТЗ-82")),
        "BU19": f"{odo_s} км / {moto_s} м/ч",
        "Q20": str(data.get("client", "")),
        "A22": f"{data.get('client', '')} ({work_type})".strip(" ()"),
        "BN22": mech,
        "R26": str(data.get("pickup", "")),
        "BP26": drv,
        "BF29": str(data.get("f_brand", "ДТ (дизельное)")),
        "BT29": (
            f"{norms_mh_str}л/мч (только м/ч)"
            if only_mh
            else f"{norm_km_f if norm_km_f > 0 else data.get('norm_km', 16.5)}л/100км; {norms_mh_str}л/мч"
        ),
        "AE30": t_out,
        "AD32": disp,
        "BT34": str(data.get("f_a", "0.0")),
        "AE35": t_in,
        "AD36": disp,
        "BT37": str(data.get("f_s", "0.0")),
        "BT38": str(f_e_val),
        "BT39": str(spent_total),
        "BT40": str(f_act),
        "BT41": "0.0 л (без учёта км)" if only_mh else f"{f_norm_km_val} л ({dist} км)",
        "BT42": f"{f_norm_mh_val} л ({moto_hours} м/ч)",
        "Q42": str(data.get("msg_type", "перегон своим ходом / работа на объекте")),
        "Q43": work_type or str(data.get("trans_type", "земляные / тракторные работы")),
        "AD45": drv,
        "BT45": f"{odo_e} км / {moto_e} м/ч",
        "BN46": mech,
        "B49": date_raw,
        "Q49": t_out,
        "AT49": date_raw,
        "BI49": t_in,
        "AG54": f"{date_raw} {t_out}".strip(),
        "BB54": mech,
    }

    s2_map = {
        "E33": str(data.get("time_tot", "9.0")),
        "E35": str(dist),
        "P33": f"{moto_hours} м/ч",
        "P35": f"{spent_total} л",
        "P38": f"{f_act} л",
        "K36": drv,
        "D39": "Диспетчер",
        "K39": disp,
    }

    segments = data.get("segments") or []
    # Формируем строки Листа 2: сначала перегоны (если есть), затем каждый выбранный вид работ с его м/ч и ДТ
    row_entries = []
    for seg in segments:
        if isinstance(seg, dict) and seg.get("entry_type") in ("work", "meta"):
            continue
        row_entries.append(
            {
                "work": str(seg.get("work_type") or ("Перегон своим ходом" if works else work_type) or "Работа МТЗ-82"),
                "from": str(seg.get("from", "")),
                "to": str(seg.get("to", "")),
                "t_out": seg.get("t_out", t_out),
                "t_in": seg.get("t_in", t_in),
                "km": "0" if only_mh else str(seg.get("km", 0.0)),
                "mh_txt": f"{seg.get('mh', 0.0)} м/ч" if not works else ("без км" if only_mh else "перегон"),
            }
        )

    if works:
        pickup_str = str(data.get("pickup", "") or (segments[0].get("from", "") if segments else ""))
        client_str = str(data.get("client", "") or "Объект работ")
        for w in works:
            w_name = str(w.get("name", "")).strip() or "Работа МТЗ-82"
            w_mh = w.get("mh", 0.0)
            w_norm = w.get("norm_mh", 0.0)
            w_dt = w.get("dt_l", round(float(w_mh or 0.0) * float(w_norm or 0.0), 2))
            row_entries.append(
                {
                    "work": w_name,
                    "from": pickup_str,
                    "to": client_str,
                    "t_out": t_out,
                    "t_in": t_in,
                    "km": "0",
                    "mh_txt": f"{w_mh} м/ч ({w_dt} л)",
                }
            )
    elif not row_entries:
        row_entries.append(
            {
                "work": work_type or "Работа спецтехники",
                "from": str(data.get("pickup", "")),
                "to": str(data.get("client", "Объект работ")),
                "t_out": t_out,
                "t_in": t_in,
                "km": str(dist),
                "mh_txt": f"{moto_hours} м/ч",
            }
        )

    for idx, entry in enumerate(row_entries[:26]):
        r = 5 + idx
        h_o, m_o = split_hh_mm(entry["t_out"])
        h_i, m_i = split_hh_mm(entry["t_in"])
        s2_map[f"B{r}"] = str(idx + 1)
        s2_map[f"C{r}"] = entry["work"]
        s2_map[f"E{r}"] = entry["from"]
        s2_map[f"H{r}"] = entry["to"]
        s2_map[f"J{r}"] = h_o
        s2_map[f"L{r}"] = m_o
        s2_map[f"M{r}"] = h_i
        s2_map[f"N{r}"] = m_i
        s2_map[f"O{r}"] = entry["km"]
        s2_map[f"Q{r}"] = entry["mh_txt"]

    return s1_map, s2_map


def _build_col_widths_map(sheet):
    """Возвращает словарь {col_index: width} с учётом диапазонов column_dimensions."""
    from openpyxl.utils import column_index_from_string

    widths = {}
    for key, dim in sheet.column_dimensions.items():
        if dim.width is None:
            continue
        if dim.min is not None and dim.max is not None:
            for c_idx in range(dim.min, dim.max + 1):
                widths[c_idx] = dim.width
        else:
            try:
                widths[column_index_from_string(key)] = dim.width
            except Exception:
                pass
    return widths


def _merged_span_width(sheet, cell, col_widths, merged_index=None):
    """Вычисляет суммарную ширину объединённой ячейки (или одной ячейки) за O(1)."""
    c1, c2 = cell.column, cell.column
    if merged_index is not None:
        bounds = merged_index.get((cell.row, cell.column))
        if bounds is not None:
            c1, c2 = bounds[1], bounds[3]
    else:
        for mr in sheet.merged_cells.ranges:
            if cell.coordinate in mr:
                c1, c2 = mr.min_col, mr.max_col
                break
    return sum(col_widths.get(c, 1.5) for c in range(c1, c2 + 1))


def _estimate_wrapped_lines(text, width_units, font_size):
    """Оценка количества строк текста при переносе по словам в Excel (с запасом под кириллицу Times New Roman)."""
    if not text:
        return 1
    chars_per_line = max(4.0, width_units * (8.2 / max(font_size, 6.0)))
    total_lines = 0
    for para in str(text).split("\n"):
        if not para:
            total_lines += 1
            continue
        cur_len = 0
        lines_in_para = 1
        for word in para.split(" "):
            w_len = len(word)
            if cur_len == 0:
                if w_len > chars_per_line:
                    lines_in_para += int((w_len - 1) // chars_per_line)
                    cur_len = w_len % int(chars_per_line) or int(chars_per_line)
                else:
                    cur_len = w_len
            elif cur_len + 1 + w_len <= chars_per_line:
                cur_len += 1 + w_len
            else:
                lines_in_para += 1
                if w_len > chars_per_line:
                    lines_in_para += int((w_len - 1) // chars_per_line)
                    cur_len = w_len % int(chars_per_line) or int(chars_per_line)
                else:
                    cur_len = w_len
        total_lines += lines_in_para
    return max(1, total_lines)


def _copy_font(font, size=None):
    return Font(
        name=(font.name if font and font.name else "Times New Roman"),
        size=(size if size is not None else (font.size if font and font.size else 9)),
        bold=(font.bold if font else False),
        italic=(font.italic if font else False),
        underline=(font.underline if font else None),
        color=(font.color if font else None),
    )


def _format_and_fit_workbook(wb, s1_keys=None, s2_keys=None, m1=None, m2=None):
    """
    Гарантирует, что ВСЕ надписи и заполненные данные полностью помещаются внутри ячеек
    без обрезания на Листе 1 и Листе 2, а оба листа строго вписываются в формат 1 страница А4 (fitToPage 1x1).
    """
    if not HAS_OPENPYXL:
        return

    # ---------------- ЛИСТ 1 (Лицевая сторона) ----------------
    if len(wb.worksheets) >= 1:
        s1 = wb.worksheets[0]
        if m1 is None:
            m1 = _build_merged_index(s1)
        col_w1 = _build_col_widths_map(s1)

        wrap_coords = {"Q8", "Z15", "M18", "A22", "R26", "BO14", "BT29", "Q42", "Q43"}
        shrink_coords = {
            "AZ4", "BX4", "M5", "BP8", "T10", "AD11", "BP11", "M12", "BP12",
            "S14", "BO15", "AR17", "BI17", "BU19", "Q20", "BN22", "BP26",
            "BF29", "AE30", "AD32", "BT34", "AE35", "AD36", "BT37", "BT38",
            "BT39", "BT40", "BT41", "BT42", "AD45", "BT45", "BN46", "AG54", "BB54",
        }

        for coord in (s1_keys or []):
            if coord in ("L5",):
                continue
            cell = get_master(s1, s1[coord], merged_index=m1)
            val = str(cell.value or "").strip()
            if not val:
                continue
            avail_w = _merged_span_width(s1, cell, col_w1, merged_index=m1)
            base_fs = float(cell.font.size if cell.font and cell.font.size else 9.0)
            h_align = (cell.alignment.horizontal if cell.alignment and cell.alignment.horizontal else "center")

            if coord in wrap_coords:
                cap_single = avail_w * (8.5 / max(base_fs, 6.0))
                if len(val) > cap_single:
                    target_fs = 8.0 if len(val) <= avail_w * 2.2 else 7.5
                    cell.font = _copy_font(cell.font, size=target_fs)
                    cell.alignment = Alignment(horizontal=h_align, vertical="center", wrap_text=True, shrink_to_fit=False)
                    lines = _estimate_wrapped_lines(val, avail_w, target_fs)
                    bounds = m1.get((cell.row, cell.column))
                    if bounds is not None:
                        row_span = bounds[2] - bounds[0] + 1
                        r_start = bounds[0]
                    else:
                        row_span = 1
                        r_start = cell.row
                    needed_h = lines * (target_fs * 1.48) + 5.0
                    cur_h = sum((s1.row_dimensions[rr].height or 12.75) for rr in range(r_start, r_start + row_span))
                    if needed_h > cur_h:
                        extra = min(needed_h - cur_h, 18.0)
                        s1.row_dimensions[r_start].height = (s1.row_dimensions[r_start].height or 12.75) + extra
                else:
                    cell.alignment = Alignment(horizontal=h_align, vertical="center", wrap_text=False, shrink_to_fit=True)
            elif coord in shrink_coords:
                cap_single = avail_w * (8.5 / max(base_fs, 6.0))
                if len(val) > cap_single:
                    target_fs = 8.0 if base_fs > 8.0 else base_fs
                    cell.font = _copy_font(cell.font, size=target_fs)
                cell.alignment = Alignment(horizontal=h_align, vertical="center", wrap_text=False, shrink_to_fit=True)

        if (s1.row_dimensions[18].height or 0) < 23.0:
            s1.row_dimensions[18].height = 23.0
        if (s1.row_dimensions[48].height or 0) < 25.0:
            s1.row_dimensions[48].height = 25.0
        if (s1.row_dimensions[53].height or 0) < 15.0:
            s1.row_dimensions[53].height = 15.0
        if (s1.row_dimensions[54].height or 0) < 14.0:
            s1.row_dimensions[54].height = 14.0

        s1.print_area = "A1:CH56"
        s1.page_margins.left = 0.20
        s1.page_margins.right = 0.20
        s1.page_margins.top = 0.20
        s1.page_margins.bottom = 0.20
        s1.print_options.horizontalCentered = True
        s1.print_options.verticalCentered = True
        s1.sheet_properties.pageSetUpPr.fitToPage = True
        s1.page_setup.fitToWidth = 1
        s1.page_setup.fitToHeight = 1

    # ---------------- ЛИСТ 2 (Оборотная сторона — пропорции строго совпадают с Листом 1 под A5/A4) ----------------
    if len(wb.worksheets) >= 2:
        s2 = wb.worksheets[1]
        if m2 is None:
            m2 = _build_merged_index(s2)

        # Суммарная ширина ~92 ед. (~690px) при высоте ~750pt (~1000px) -> пропорция W/H ≈ 0.69 (точно как у Листа 1 / A5)
        s2_widths = {
            "A": 0.8,
            "B": 3.8,
            "C": 9.5,
            "D": 6.8,
            "E": 17.5,
            "F": 1.0,
            "G": 1.0,
            "H": 17.5,
            "I": 1.4,
            "J": 1.4,
            "K": 2.8,
            "L": 3.6,
            "M": 3.8,
            "N": 3.8,
            "O": 2.2,
            "P": 4.8,
            "Q": 9.2,
            "R": 1.8,
            "S": 0.6,
        }
        for col_letter, w_val in s2_widths.items():
            s2.column_dimensions[col_letter].width = w_val

        s2.row_dimensions[1].height = 20.0
        s2.row_dimensions[2].height = 18.0
        s2.row_dimensions[3].height = 18.0
        s2.row_dimensions[4].height = 24.0
        align_center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
        align_center_shrink = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)

        for hdr_coord in ["B2", "C2", "E2", "E3", "H3", "J2", "J3", "M3", "J4", "L4", "M4", "N4", "O2", "Q2"]:
            hc = get_master(s2, s2[hdr_coord], merged_index=m2)
            hc.font = _copy_font(hc.font, size=8.0)
            hc.alignment = align_center_wrap

        filled_rows = 0
        empty_rows = []
        row_cols_align = [
            ("B", align_center_wrap), ("C", align_left_wrap), ("E", align_left_wrap), ("H", align_left_wrap),
            ("J", align_center_wrap), ("L", align_center_wrap), ("M", align_center_wrap), ("N", align_center_wrap),
            ("O", align_center_wrap), ("Q", align_center_wrap),
        ]
        for r in range(5, 31):
            c_txt = str(get_master(s2, s2[f"C{r}"], merged_index=m2).value or "").strip()
            e_txt = str(get_master(s2, s2[f"E{r}"], merged_index=m2).value or "").strip()
            h_txt = str(get_master(s2, s2[f"H{r}"], merged_index=m2).value or "").strip()
            q_txt = str(get_master(s2, s2[f"Q{r}"], merged_index=m2).value or "").strip()
            b_txt = str(get_master(s2, s2[f"B{r}"], merged_index=m2).value or "").strip()

            has_data = bool(c_txt or e_txt or h_txt or q_txt or b_txt)
            if has_data:
                filled_rows += 1
                for col_l, al_obj in row_cols_align:
                    cell = get_master(s2, s2[f"{col_l}{r}"], merged_index=m2)
                    cell.font = _copy_font(cell.font, size=8.0)
                    cell.alignment = al_obj
                l_c = _estimate_wrapped_lines(c_txt, 16.0, 8.0)
                l_e = _estimate_wrapped_lines(e_txt, 19.5, 8.0)
                l_h = _estimate_wrapped_lines(h_txt, 19.5, 8.0)
                l_q = _estimate_wrapped_lines(q_txt, 11.0, 8.0)
                max_l = max(l_c, l_e, l_h, l_q, 1)
                s2.row_dimensions[r].height = max(21.0, max_l * 12.2 + 7.0)
            else:
                empty_rows.append(r)
                s2.row_dimensions[r].height = 19.5

        if filled_rows > 8:
            for r in empty_rows:
                s2.row_dimensions[r].height = 15.0

        s2.row_dimensions[31].height = 24.0
        s2.row_dimensions[33].height = 26.0
        s2.row_dimensions[35].height = 26.0
        s2.row_dimensions[38].height = 22.0
        s2.row_dimensions[39].height = 22.0
        for box_coord in ["E33", "E35", "P33", "P35", "P38", "K36", "D39", "K39"]:
            bc = get_master(s2, s2[box_coord], merged_index=m2)
            bc.font = _copy_font(bc.font, size=9.0)
            bc.alignment = align_center_shrink

        s2.print_area = "A1:S40"
        s2.page_margins.left = 0.20
        s2.page_margins.right = 0.20
        s2.page_margins.top = 0.20
        s2.page_margins.bottom = 0.20
        s2.print_options.horizontalCentered = True
        s2.print_options.verticalCentered = True
        s2.sheet_properties.pageSetUpPr.fitToPage = True
        s2.page_setup.fitToWidth = 1
        s2.page_setup.fitToHeight = 1


def _resolve_template_path(tpl_path):
    """Ищет шаблон .xlsx в текущей папке, рядом с .exe, в _MEIPASS (PyInstaller) и рядом с parser_f3.py."""
    if tpl_path and os.path.exists(tpl_path):
        return tpl_path
    candidates = []
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        candidates.append(os.path.join(exe_dir, tpl_path))
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            candidates.append(os.path.join(meipass, tpl_path))
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates.append(os.path.join(script_dir, tpl_path))
    for cand in candidates:
        if os.path.exists(cand):
            return cand
    return None


def generate_f3_from_excel(tpl_path, out_file, data):
    """
    Заполняет официальный бланк путевого листа легкового автомобиля (Типовая межотраслевая форма № 3,
    шаблон templates/put-list-f3.xlsx) по всем реквизитам Листа 1 и Листа 2.
    """
    resolved_tpl = _resolve_template_path(tpl_path)
    if not resolved_tpl:
        return (False, f"Шаблон не найден: {tpl_path}")
    tpl_path = resolved_tpl

    s1_map, s2_map = _build_f3_cell_maps(data)
    drv = str(data.get("drv", "")).strip()
    ru_date = format_ru_date(data.get("date", ""))

    if not HAS_OPENPYXL:
        try:
            return _apply_template_zip_fallback(tpl_path, out_file, s1_map, s2_map, drv=drv)
        except Exception as e:
            return (False, f"Ошибка при заполнении: {str(e)}")

    try:
        wb = _load_workbook_cached(tpl_path)
        m1 = _build_merged_index(wb.worksheets[0]) if len(wb.worksheets) >= 1 else None
        m2 = _build_merged_index(wb.worksheets[1]) if len(wb.worksheets) >= 2 else None

        # 1. Быстрый обход только заполненных ячеек (без пустых ячеек 86x56)
        for s_idx, sheet in enumerate(wb.worksheets):
            m_idx = m1 if s_idx == 0 else (m2 if s_idx == 1 else None)
            for cell in list(sheet._cells.values()):
                val_raw = cell.value
                if val_raw is None or not isinstance(val_raw, str):
                    continue
                txt = val_raw.lower()
                if "аксарин" in txt or "ксарин" in txt or "{фамилияио}" in txt:
                    v = val_raw
                    if "аксарин" in txt or "ксарин" in txt:
                        v = clean_fio(v, drv)
                    v = v.replace("{ФамилияИО}", drv).replace("{фамилияио}", drv)
                    get_master(sheet, cell, merged_index=m_idx).value = v
                elif "срок действия" in txt:
                    get_master(sheet, cell, merged_index=m_idx).value = f"Срок действия с {ru_date} по {ru_date}"

        # 2. Точная адресная простановка всех полей Листа 1 и Листа 2 по координатам бланка Госкомстата
        if len(wb.worksheets) >= 1:
            s1 = wb.worksheets[0]
            for coord, val in s1_map.items():
                set_cell_master(s1, coord, val, merged_index=m1)

        if len(wb.worksheets) >= 2:
            s2 = wb.worksheets[1]
            for coord, val in s2_map.items():
                set_cell_master(s2, coord, val, merged_index=m2)

        # 3. Автоподгонка размеров, переноса текста и вписывание в 1 страницу A4 на каждый лист
        _format_and_fit_workbook(wb, s1_keys=s1_map.keys(), s2_keys=s2_map.keys(), m1=m1, m2=m2)

        os.makedirs(os.path.dirname(os.path.abspath(out_file)), exist_ok=True)
        wb.save(out_file)
        return (True, out_file)
    except PermissionError:
        return (
            False,
            "Файл открыт в Excel! Пожалуйста, закройте путевой лист и попробуйте выписать снова.",
        )
    except Exception as e:
        return (False, f"Ошибка при заполнении: {str(e)}")


def generate_tractor_wb_excel(out_file, data, tpl_path=None):
    """
    Формирует заполненный официальный бланк путевого листа для:
    - Трактора МТЗ-82 (Специализированная форма № 412-АПК -> templates/put-list-412-apk.xlsx)
    - Экскаватора на базе МТЗ-82 (Типовая межотраслевая форма № ЭСМ-2 -> templates/put-list-esm2.xlsx)
    с сохранением всей 86-колоночной разметки, рамок, штампов медосмотра/техконтроля и настроек печати.
    """
    form_type = str(data.get("form_type", "412-АПК"))
    is_esm2 = "ЭСМ" in form_type.upper()

    if not tpl_path:
        tpl_name = "put-list-esm2.xlsx" if is_esm2 else "put-list-412-apk.xlsx"
        tpl_path = _resolve_template_path(os.path.join("templates", tpl_name)) or _resolve_template_path(
            os.path.join("templates", "put-list-f3.xlsx")
        )
    else:
        tpl_path = _resolve_template_path(tpl_path)

    if not tpl_path or not os.path.exists(tpl_path):
        return (False, f"Шаблон путевого листа трактора не найден: {tpl_path}")

    s1_map, s2_map = _build_tractor_cell_maps(data)
    drv = str(data.get("drv", "")).strip()

    if not HAS_OPENPYXL:
        try:
            return _apply_template_zip_fallback(tpl_path, out_file, s1_map, s2_map, drv=drv)
        except Exception as e:
            return (False, f"Ошибка при формировании бланка трактора/экскаватора: {str(e)}")

    try:
        wb = _load_workbook_cached(tpl_path)
        m1 = _build_merged_index(wb.worksheets[0]) if len(wb.worksheets) >= 1 else None
        m2 = _build_merged_index(wb.worksheets[1]) if len(wb.worksheets) >= 2 else None

        if len(wb.worksheets) >= 1:
            s1 = wb.worksheets[0]
            for coord, val in s1_map.items():
                set_cell_master(s1, coord, val, merged_index=m1)

        if len(wb.worksheets) >= 2:
            s2 = wb.worksheets[1]
            for coord, val in s2_map.items():
                set_cell_master(s2, coord, val, merged_index=m2)

        _format_and_fit_workbook(wb, s1_keys=s1_map.keys(), s2_keys=s2_map.keys(), m1=m1, m2=m2)

        os.makedirs(os.path.dirname(os.path.abspath(out_file)), exist_ok=True)
        wb.save(out_file)
        return (True, out_file)
    except PermissionError:
        return (
            False,
            "Файл открыт в Excel! Пожалуйста, закройте путевой лист трактора и попробуйте снова.",
        )
    except Exception as e:
        return (False, f"Ошибка при формировании бланка трактора/экскаватора: {str(e)}")
