from datetime import datetime, timedelta
import math
import re


def haversine(lat1, lon1, lat2, lon2):
    """Вычисляет расстояние между GPS координатами (с учетом изгиба дорог)"""
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(math.radians(lat1))
        * math.cos(math.radians(lat2))
        * math.sin(dlon / 2) ** 2
    )
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return round(R * c * 1.3, 1)


def calculate_distance(lat1, lon1, lat2, lon2):
    return haversine(lat1, lon1, lat2, lon2)


def calculate_fuel_consumption(base_norm, mileage, winter_coef=0.0, city_coef=0.0):
    try:
        wc = float(winter_coef or 0.0)
        cc = float(city_coef or 0.0)
        # Поддержка как долей (0.10 = +10%), так и процентов (10.0 = +10%)
        if abs(wc) >= 1.0:
            wc = wc / 100.0
        if abs(cc) >= 1.0:
            cc = cc / 100.0
        return round(
            float(base_norm) * 0.01 * float(mileage) * (1.0 + wc + cc),
            2,
        )
    except Exception:
        return 0.0


def calculate_tractor_diesel_consumption(
    norm_km,
    mileage_km,
    norm_mh,
    moto_hours,
    extra_coef_pct=0.0,
):
    """
    Расчёт расхода дизельного топлива (ДТ) для трактора МТЗ-82 и экскаватора на базе МТЗ-82:
    1) Расход ДТ на перегон своим ходом (по километражу, л/100 км)
    2) Расход ДТ на работу спецоборудования / двигателя под нагрузкой (по моточасам, л/м·ч)
    Возвращает кортеж: (f_norm_km, f_norm_mh, f_norm_total)
    """
    try:
        k = 1.0 + (float(extra_coef_pct or 0.0) / 100.0)
        f_norm_km = round((float(mileage_km or 0.0) / 100.0) * float(norm_km or 0.0) * k, 2)
        f_norm_mh = round(float(moto_hours or 0.0) * float(norm_mh or 0.0) * k, 2)
        f_norm_total = round(f_norm_km + f_norm_mh, 2)
        return f_norm_km, f_norm_mh, f_norm_total
    except Exception:
        return 0.0, 0.0, 0.0


def calculate_tractor_multi_works_diesel(
    norm_km,
    mileage_km,
    works_list,
    extra_coef_pct=0.0,
):
    """
    Расчёт расхода ДТ для трактора МТЗ-82 при выполнении нескольких видов работ за одну смену.
    works_list: список словарей {'name': str, 'mh': float, 'norm_mh': float, 'unit': str}
    Возвращает кортеж: (f_norm_km, f_norm_mh, f_norm_total, total_mh, per_work_dts)
    """
    try:
        k = 1.0 + (float(extra_coef_pct or 0.0) / 100.0)
        f_norm_km = round((float(mileage_km or 0.0) / 100.0) * float(norm_km or 0.0) * k, 2)
        per_work_dts = []
        total_mh = 0.0
        total_mh_dt = 0.0
        for w in works_list or []:
            mh = float(w.get("mh", 0.0) or 0.0)
            nm = float(w.get("norm_mh", 0.0) or 0.0)
            w_dt = round(mh * nm * k, 2)
            per_work_dts.append(w_dt)
            total_mh += mh
            total_mh_dt += w_dt
        total_mh = round(total_mh, 2)
        f_norm_mh = round(total_mh_dt, 2)
        f_norm_total = round(f_norm_km + f_norm_mh, 2)
        return f_norm_km, f_norm_mh, f_norm_total, total_mh, per_work_dts
    except Exception:
        return 0.0, 0.0, 0.0, 0.0, []


def calculate_fuel_balance(fuel_start, fuel_added, fuel_spent):
    try:
        return round(float(fuel_start) + float(fuel_added) - float(fuel_spent), 2)
    except Exception:
        return 0.0


def parse_time_intervals(intervals_text):
    total_hours = 0.0
    try:
        times = re.findall(r"(\d{1,2}:\d{2})", intervals_text)
        if len(times) >= 2:
            t_out = datetime.strptime(times[0], "%H:%M")
            t_in = datetime.strptime(times[-1], "%H:%M")
            if t_in < t_out:
                t_in += timedelta(days=1)
            total_hours = round((t_in - t_out).total_seconds() / 3600.0, 1)
    except Exception:
        pass
    return total_hours if total_hours > 0 else 8.0
