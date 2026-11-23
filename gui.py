import ast
import json
import math
import os
import queue
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox, ttk

_INTERNAL_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_internal")
if os.path.isdir(_INTERNAL_DIR) and _INTERNAL_DIR not in sys.path:
    sys.path.insert(0, _INTERNAL_DIR)

import customtkinter as ctk
import requests
import tkintermapview
from tkcalendar import Calendar

from database import (
    DatabaseManager,
    TractorDatabaseManager,
    _wb_chrono_key,
    create_fast_db_backup,
    create_mode_backup,
    create_unified_backup,
)
from emulator_cli import ensure_orenburg_workspace, ensure_tractor_workspace
from export import export_history_to_excel, generate_word_report
from logic_gsm import (
    calculate_distance,
    calculate_fuel_consumption,
    calculate_tractor_diesel_consumption,
    calculate_tractor_multi_works_diesel,
)
from parser_f3 import generate_f3_from_excel, generate_tractor_wb_excel

ctk.set_appearance_mode("light")


def safe_float(val, default=0.0):
    """Безопасно преобразует значение из настроек/полей ввода в float (поддерживает запятые и строки вида 'Трасса (0%)' / 'Зима (+10%)')."""
    if val is None or isinstance(val, bool):
        return float(default)
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", ".")
    if not s:
        return float(default)
    try:
        return float(s)
    except (ValueError, TypeError):
        match = re.search(r"[-+]?\d+(?:\.\d+)?", s)
        if match:
            try:
                return float(match.group(0))
            except (ValueError, TypeError):
                pass
        return float(default)


def safe_str(val, default=""):
    if val is None:
        return str(default)
    text = str(val)
    return text if text != "" else str(default)


# Константы проекции Яндекс Карт (эллиптический Меркатор EPSG:3395 vs сферический EPSG:3857 в TkinterMapView)
YANDEX_WGS84_E = 0.08181919084262149

MAP_PROVIDERS = {
    "Яндекс Карты": {
        "key": "yandex",
        "url": "https://core-renderer-tiles.maps.yandex.net/tiles?l=map&x={x}&y={y}&z={z}&scale=1&lang=ru_RU",
        "max_zoom": 19,
        "elliptical": True,
    },
    "OpenStreetMap": {
        "key": "osm",
        "url": "https://a.tile.openstreetmap.org/{z}/{x}/{y}.png",
        "max_zoom": 19,
        "elliptical": False,
    },
    "Google Карты": {
        "key": "google",
        "url": "https://mt0.google.com/vt/lyrs=m&hl=ru&x={x}&y={y}&z={z}&s=Ga",
        "max_zoom": 20,
        "elliptical": False,
    },
}


def wgs84_to_yandex_display_lat(lat_deg):
    """Переводит истинную широту WGS84 в широту отображения для TkinterMapView (EPSG:3857) поверх тайлов Яндекс Карт (EPSG:3395)."""
    lat_clamped = max(-85.051128, min(85.051128, float(lat_deg)))
    phi = math.radians(lat_clamped)
    sin_phi = math.sin(phi)
    ecc_term = ((1.0 - YANDEX_WGS84_E * sin_phi) / (1.0 + YANDEX_WGS84_E * sin_phi)) ** (
        YANDEX_WGS84_E / 2.0
    )
    theta = math.tan(math.pi / 4.0 + phi / 2.0) * ecc_term
    disp_phi = 2.0 * math.atan(theta) - math.pi / 2.0
    return math.degrees(disp_phi)


def yandex_display_to_wgs84_lat(disp_lat_deg):
    """Обратное преобразование: переводит широту клика в TkinterMapView поверх Яндекс Карт обратно в истинную широту WGS84."""
    disp_clamped = max(-85.051128, min(85.051128, float(disp_lat_deg)))
    disp_phi = math.radians(disp_clamped)
    exp_merc_y = math.tan(math.pi / 4.0 + disp_phi / 2.0)
    phi = disp_phi
    for _ in range(6):
        sin_phi = math.sin(phi)
        ecc_term = ((1.0 - YANDEX_WGS84_E * sin_phi) / (1.0 + YANDEX_WGS84_E * sin_phi)) ** (
            YANDEX_WGS84_E / 2.0
        )
        phi = 2.0 * math.atan(exp_merc_y / ecc_term) - math.pi / 2.0
    return math.degrees(phi)


class CustomInputDialog(ctk.CTkToplevel):
    def __init__(
        self,
        parent,
        title="Ввод данных",
        prompt="Введите значение:",
        initial_value="",
    ):
        super().__init__(parent)
        self.title(title)
        self.geometry("450x220")
        self.resizable(False, False)
        self.configure(fg_color="#FFFFFF")
        self.transient(parent)
        self.grab_set()
        self.result = None

        ctk.CTkLabel(
            self,
            text=prompt,
            font=("Helvetica", 14, "bold"),
            text_color="#1D1D1F",
        ).pack(pady=(25, 10), padx=25, anchor="w")

        self.entry = ctk.CTkEntry(
            self,
            width=400,
            height=38,
            font=("Helvetica", 14),
        )
        self.entry.pack(pady=5, padx=25)
        if initial_value:
            self.entry.insert(0, initial_value)
        self.entry.focus_set()

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=20, padx=25, fill="x")

        ctk.CTkButton(
            btn_frame,
            text="Отмена",
            fg_color="#E5E7EB",
            hover_color="#D1D5DB",
            text_color="#374151",
            width=120,
            height=36,
            command=self.cancel,
        ).pack(side="right", padx=(10, 0))

        ctk.CTkButton(
            btn_frame,
            text="Сохранить",
            fg_color="#007AFF",
            hover_color="#0056B3",
            width=120,
            height=36,
            command=self.confirm,
        ).pack(side="right")

        self.bind("<Return>", lambda e: self.confirm())
        self.bind("<Escape>", lambda e: self.cancel())

    def confirm(self):
        self.result = self.entry.get().strip()
        self.destroy()

    def cancel(self):
        self.result = None
        self.destroy()


class BackupChoiceDialog(ctk.CTkToplevel):
    """Окно выбора типа резервной копии: отдельно Легковые, отдельно Трактора МТЗ-82 или Единая."""

    def __init__(self, parent, active_mode="cars"):
        super().__init__(parent)
        self.title("Сохранение резервной копии")
        self.geometry("500x310")
        self.resizable(False, False)
        self.configure(fg_color="#FFFFFF")
        self.transient(parent)
        self.grab_set()
        self.choice = None

        ctk.CTkLabel(
            self,
            text="Выберите, какую базу данных сохранить в резервную копию:",
            font=("Helvetica", 14, "bold"),
            text_color="#1D1D1F",
            wraplength=450,
            justify="left",
        ).pack(pady=(20, 12), padx=24, anchor="w")

        ctk.CTkButton(
            self,
            text="🚗  Только Легковые авто (waybills.db + настройки + ПЛ)",
            fg_color="#007AFF" if active_mode == "cars" else "#3B82F6",
            hover_color="#0056B3",
            height=42,
            font=("Helvetica", 13, "bold"),
            command=lambda: self._pick("cars"),
        ).pack(fill="x", padx=24, pady=6)

        ctk.CTkButton(
            self,
            text="🚜  Только Трактора МТЗ-82 / Экскаваторы (tractors.db + ПЛ)",
            fg_color="#D97706" if active_mode == "tractors" else "#F59E0B",
            hover_color="#B45309",
            height=42,
            font=("Helvetica", 13, "bold"),
            command=lambda: self._pick("tractors"),
        ).pack(fill="x", padx=24, pady=6)

        ctk.CTkButton(
            self,
            text="📦  Единая копия: Легковые + Трактора вместе (.zip + .db)",
            fg_color="#10B981",
            hover_color="#059669",
            height=42,
            font=("Helvetica", 13, "bold"),
            command=lambda: self._pick("unified"),
        ).pack(fill="x", padx=24, pady=6)

        ctk.CTkButton(
            self,
            text="Отмена",
            fg_color="#E5E7EB",
            hover_color="#D1D5DB",
            text_color="#374151",
            height=34,
            command=self.destroy,
        ).pack(pady=(12, 16), padx=24, anchor="e")

        self.bind("<Escape>", lambda e: self.destroy())

    def _pick(self, mode):
        self.choice = mode
        self.destroy()


class WorkNormMassUpdateDialog(ctk.CTkToplevel):
    """Окно массового пересчёта нормы расхода ДТ по конкретному виду работ (по всем ПЛ или за период)."""

    def __init__(
        self,
        parent,
        work_types,
        preselect_work=None,
        preselect_norm=None,
        default_from="",
        default_to="",
    ):
        super().__init__(parent)
        self.title("Изменение нормы расхода по виду работ в путевых листах")
        self.geometry("560x420")
        self.resizable(False, False)
        self.configure(fg_color="#FFFFFF")
        self.transient(parent)
        self.grab_set()
        self.result = None

        self._norm_map = {str(r[1]): float(r[2] or 0.0) for r in work_types if r and r[1]}
        wt_names = list(self._norm_map.keys()) or ["Очистка дорог от снежного покрова"]

        ctk.CTkLabel(
            self,
            text="Пересчёт нормы ДТ по виду работ в сохранённых путёвках",
            font=("Helvetica", 15, "bold"),
            text_color="#1D1D1F",
        ).pack(pady=(16, 4), padx=22, anchor="w")

        ctk.CTkLabel(
            self,
            text="Изменяется расход ТОЛЬКО по выбранному виду работ; остальные виды работ\nв путевых листах остаются без изменений, остатки ДТ пересчитываются по цепочке.",
            font=("Helvetica", 11),
            text_color="#4B5563",
            justify="left",
        ).pack(pady=(0, 10), padx=22, anchor="w")

        ctk.CTkLabel(
            self, text="Вид работ:", font=("Helvetica", 12, "bold"), text_color="#111827"
        ).pack(padx=22, anchor="w")
        self.cb_work = ctk.CTkComboBox(
            self, values=wt_names, width=516, command=self._on_work_selected
        )
        self.cb_work.pack(padx=22, pady=(2, 8), fill="x")
        init_work = (
            preselect_work
            if preselect_work and preselect_work in wt_names
            else wt_names[0]
        )
        self.cb_work.set(init_work)

        f_norm = ctk.CTkFrame(self, fg_color="transparent")
        f_norm.pack(fill="x", padx=22, pady=(2, 8))
        ctk.CTkLabel(
            f_norm,
            text="Новая норма расхода ДТ (л/м·ч):",
            font=("Helvetica", 12, "bold"),
        ).pack(side="left")
        self.en_norm = ctk.CTkEntry(f_norm, width=110)
        self.en_norm.pack(side="left", padx=10)
        init_norm = (
            preselect_norm
            if preselect_norm is not None
            else self._norm_map.get(init_work, 5.2)
        )
        self.en_norm.insert(0, f"{float(init_norm):g}")

        self.var_scope = tk.StringVar(value="all")
        f_scope = ctk.CTkFrame(self, fg_color="#F8FAFC", corner_radius=8)
        f_scope.pack(fill="x", padx=22, pady=(2, 8))

        ctk.CTkRadioButton(
            f_scope,
            text="По всем сохранённым путевым листам",
            variable=self.var_scope,
            value="all",
            font=("Helvetica", 12, "bold"),
            command=self._toggle_dates,
        ).pack(anchor="w", padx=12, pady=(8, 4))

        ctk.CTkRadioButton(
            f_scope,
            text="За определённый период (по дате составления ПЛ):",
            variable=self.var_scope,
            value="period",
            font=("Helvetica", 12, "bold"),
            command=self._toggle_dates,
        ).pack(anchor="w", padx=12, pady=(2, 4))

        f_dates = ctk.CTkFrame(f_scope, fg_color="transparent")
        f_dates.pack(fill="x", padx=28, pady=(0, 10))
        ctk.CTkLabel(f_dates, text="С даты:").pack(side="left")
        self.en_from = ctk.CTkEntry(f_dates, width=120, placeholder_text="01.01.2025")
        self.en_from.pack(side="left", padx=(6, 16))
        self.en_from.insert(0, default_from or datetime.now().strftime("01.01.%Y"))

        ctk.CTkLabel(f_dates, text="По дату:").pack(side="left")
        self.en_to = ctk.CTkEntry(f_dates, width=120, placeholder_text="31.12.2025")
        self.en_to.pack(side="left", padx=6)
        self.en_to.insert(0, default_to or datetime.now().strftime("%d.%m.%Y"))

        self.var_update_ref = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            self,
            text="Также обновить эту норму в справочнике видов работ",
            variable=self.var_update_ref,
            font=("Helvetica", 12),
        ).pack(anchor="w", padx=22, pady=(4, 2))

        self.var_rebuild_excel = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            self,
            text="Переформировать существующие Excel-файлы (в фоне)",
            variable=self.var_rebuild_excel,
            font=("Helvetica", 12),
        ).pack(anchor="w", padx=22, pady=(2, 8))

        f_btns = ctk.CTkFrame(self, fg_color="transparent")
        f_btns.pack(fill="x", padx=22, pady=(6, 14))
        ctk.CTkButton(
            f_btns,
            text="Применить и пересчитать",
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            height=36,
            command=self._confirm,
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(
            f_btns,
            text="Отмена",
            fg_color="#E5E7EB",
            hover_color="#D1D5DB",
            text_color="#1F2937",
            width=110,
            height=36,
            command=self.destroy,
        ).pack(side="right")

        self._toggle_dates()
        self.bind("<Escape>", lambda e: self.destroy())

    def _on_work_selected(self, chosen):
        if chosen in self._norm_map:
            self.en_norm.delete(0, "end")
            self.en_norm.insert(0, f"{self._norm_map[chosen]:g}")

    def _toggle_dates(self):
        st = "normal" if self.var_scope.get() == "period" else "disabled"
        self.en_from.configure(state=st)
        self.en_to.configure(state=st)

    def _confirm(self):
        work_name = self.cb_work.get().strip()
        new_norm = safe_float(self.en_norm.get(), 0.0)
        if not work_name:
            messagebox.showwarning("Внимание", "Выберите вид работ.", parent=self)
            return
        if new_norm <= 0:
            messagebox.showwarning("Внимание", "Укажите корректную норму ДТ (> 0 л/м·ч).", parent=self)
            return
        d_from = None
        d_to = None
        if self.var_scope.get() == "period":
            d_from = self.en_from.get().strip()
            d_to = self.en_to.get().strip()
            try:
                dt1 = datetime.strptime(d_from, "%d.%m.%Y")
                dt2 = datetime.strptime(d_to, "%d.%m.%Y")
                if dt1 > dt2:
                    messagebox.showwarning(
                        "Внимание", "Дата начала периода больше даты окончания.", parent=self
                    )
                    return
            except Exception:
                messagebox.showwarning(
                    "Внимание", "Введите даты периода в формате ДД.ММ.ГГГГ (например, 01.01.2025).", parent=self
                )
                return
        self.result = {
            "work_name": work_name,
            "new_norm": new_norm,
            "scope": self.var_scope.get(),
            "date_from": d_from,
            "date_to": d_to,
            "update_ref": bool(self.var_update_ref.get()),
            "rebuild_excel": bool(self.var_rebuild_excel.get()),
        }
        self.destroy()


class PrintFlipDialog(ctk.CTkToplevel):
    """Окно паузы после печати 1-й страницы (лицевой) для одиночной и массовой печати."""

    def __init__(
        self,
        parent,
        count,
        orientation_label,
        is_landscape,
        rotate_180_default=True,
        sheets_count=None,
    ):
        super().__init__(parent)
        self.title("Переворот листов — Оборотная сторона (стр. 2)")
        self.geometry("560x325")
        self.resizable(False, False)
        self.configure(fg_color="#FFFFFF")
        self.transient(parent)
        self.grab_set()
        self.confirmed = False
        self.var_rotate_180 = tk.BooleanVar(value=bool(rotate_180_default))

        sh_cnt = sheets_count if sheets_count is not None else count
        ctk.CTkLabel(
            self,
            text=f"Лицевая сторона (стр. 1) отправлена: {count} путёв. ({sh_cnt} лист. A4)",
            font=("Helvetica", 15, "bold"),
            text_color="#1D1D1F",
        ).pack(pady=(18, 6), padx=24, anchor="w")

        ctk.CTkLabel(
            self,
            text=f"Формат печати: {orientation_label}",
            font=("Helvetica", 12, "bold"),
            text_color="#2563EB",
        ).pack(pady=(0, 8), padx=24, anchor="w")

        if is_landscape:
            hint_text = (
                "1. Дождитесь выхода всех листов A4 с 1-й стороной из принтера.\n"
                "   • 1 путёвка печатается на 1-й (левой) половине листа A4 (вторая половина остаётся чистой).\n"
                "   • 2 и более путёвок автоматически раскладываются по 2 шт. на один лист A4.\n"
                "2. Переверните листы чистой стороной вверх (вдоль длинного края, той же кромкой вперёд) и вставьте в лоток.\n"
                "3. Нажмите кнопку ниже — 2-я сторона напечатается ровно того же размера A5 с заглавия (не вверх тормашками)."
            )
        else:
            hint_text = (
                "1. Дождитесь выхода всех листов 1-й страницы из принтера.\n"
                "2. Переверните стопку чистой стороной для печати оборота (заглавие остаётся вверху) и вставьте в лоток.\n"
                "3. Нажмите кнопку ниже — 2-я страница напечатается ровно в том же размере A4."
            )

        ctk.CTkLabel(
            self,
            text=hint_text,
            font=("Helvetica", 12),
            text_color="#374151",
            justify="left",
            wraplength=510,
        ).pack(pady=(2, 10), padx=24, anchor="w")

        if is_landscape:
            ctk.CTkCheckBox(
                self,
                text="Авто-разворот 2-й стр. на 180° (стандарт для переворота вдоль длинного края)",
                variable=self.var_rotate_180,
                font=("Helvetica", 12, "bold"),
                text_color="#111827",
            ).pack(pady=(0, 10), padx=24, anchor="w")

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=24, pady=(8, 16))

        ctk.CTkButton(
            btn_row,
            text="Отмена (не печатать оборот)",
            fg_color="#E5E7EB",
            hover_color="#D1D5DB",
            text_color="#374151",
            width=190,
            height=38,
            command=self._cancel,
        ).pack(side="right", padx=(10, 0))

        ctk.CTkButton(
            btn_row,
            text="🖨️ Печатать 2-ю страницу (оборот)",
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            width=260,
            height=38,
            font=("Helvetica", 13, "bold"),
            command=self._confirm,
        ).pack(side="right")

        self.bind("<Return>", lambda e: self._confirm())
        self.bind("<Escape>", lambda e: self._cancel())

    def _confirm(self):
        self.confirmed = True
        self.destroy()

    def _cancel(self):
        self.confirmed = False
        self.destroy()


class WaybillApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Про.Путевые Листы - ERP Система")
        self.geometry("1400x900")
        self.configure(fg_color="#F5F5F7")

        self.db_cars = DatabaseManager("waybills.db")
        self.db_tractors = TractorDatabaseManager("tractors.db")
        try:
            ensure_orenburg_workspace(self.db_cars)
            ensure_tractor_workspace(self.db_tractors)
        except Exception:
            pass
        self.active_mode = "cars"
        self.db = self.db_cars
        self.memory_file = "memory.json"
        self.settings = self.load_memory()

        self.txt_route = None
        self.txt_route_main = None
        self.route_markers = []
        self.route_paths = []
        self.marker_metadata = []
        self.selected_marker_index = None
        saved_provider = safe_str(self.settings.get("map_provider"), "Яндекс Карты")
        self.map_provider_name = (
            saved_provider if saved_provider in MAP_PROVIDERS else "Яндекс Карты"
        )
        self.compact_labels_var = tk.BooleanVar(
            value=bool(self.settings.get("compact_map_labels", True))
        )
        saved_orient = safe_str(
            self.settings.get("print_orientation"), "Поперёк (A5 — 2 на лист A4)"
        )
        if "Поперёк" in saved_orient or "Альбом" in saved_orient:
            saved_orient = "Поперёк (A5 — 2 на лист A4)"
        else:
            saved_orient = "На весь лист (Книжная A4)"
        self.print_orientation_var = tk.StringVar(value=saved_orient)
        self.print_rotate_180 = bool(self.settings.get("print_rotate_180", True))
        self.route_segments = self.settings.get("segments", [])
        self.tractor_works = list(self.settings.get("tractor_works", []) or [])
        self._route_mem_cache = {}
        self._recalc_after_id = None
        self._map_fetch_gen = 0
        self._map_dirty = True
        self._bg_backup_after_id = None
        self._last_works_render_sig = None
        self._excel_bg_queue = queue.Queue()
        self._excel_bg_worker_started = False

        self.editing_id = None
        self.old_filename = None
        self._editing_old_veh = None
        self.editing_crud_id = {
            "drivers": None,
            "vehicles": None,
            "orgs": None,
            "clients": None,
            "staff": None,
            "work_types": None,
        }
        self.last_spent_p = 0.0
        self.last_spent_g = 0.0
        self.last_dt_km = 0.0
        self.last_dt_mh = 0.0
        self.last_dt_total = 0.0
        self._tractor_norm_km = 16.5

        try:
            self.current_user = os.getlogin()
        except Exception:
            self.current_user = "User"

        self.create_widgets()
        self.load_vehicle_data()
        self.recalc_live()
        self.log_action("Вход в систему")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def log_action(self, action_text):
        try:
            timestamp = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
            with open("audit.log", "a", encoding="utf-8") as f:
                f.write(
                    f"[{timestamp}] Пользователь: {self.current_user} | Действие: {action_text}\n"
                )
        except Exception:
            pass

    def load_memory(self):
        if os.path.exists(self.memory_file):
            try:
                with open(self.memory_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                demo_locs = {
                    "Беляевка, ул. Ленинская (стройплощадка)",
                    "Буртинский (База МТП)",
                    "Буртинский, ул. Советская (участок водопровода)",
                    "Оренбург, ул. Родимцева, 16 (объект ДорСтрой)",
                    "Саракташ (коммунальный участок)",
                }
                if isinstance(data, dict) and data.get("pickup") in demo_locs:
                    data["pickup"] = ""
                return data
            except Exception:
                return {}
        return {}

    def save_memory(self, *args):
        try:
            mem_segments = [
                {
                    "from": s.get("from", ""),
                    "to": s.get("to", ""),
                    "km": s.get("km", 0.0),
                }
                for s in self.route_segments
            ]
            drv_val = self.cb_drv.get() if hasattr(self, "cb_drv") else ""
            veh_val = self.cb_veh.get() if hasattr(self, "cb_veh") else ""
            mech_val = self.cb_mech.get() if hasattr(self, "cb_mech") else ""
            disp_val = self.cb_disp.get() if hasattr(self, "cb_disp") else ""
            data = {
                "org": self.cb_org.get() if hasattr(self, "cb_org") else "",
                "client": self.cb_client.get() if hasattr(self, "cb_client") else "",
                "pickup": self.cb_pickup.get() if hasattr(self, "cb_pickup") else "",
                "driver": drv_val,
                "drv": drv_val,
                "vehicle": veh_val,
                "veh": veh_val,
                "mechanic": mech_val,
                "mech": mech_val,
                "dispatcher": disp_val,
                "disp": disp_val,
                "okpo": self.en_okpo.get() if hasattr(self, "en_okpo") else "48291034",
                "trailer_name": self.en_trailer.get() if hasattr(self, "en_trailer") else "без прицепа",
                "msg_type": self.cb_msg_type.get() if hasattr(self, "cb_msg_type") else "городское / пригородное",
                "trans_type": self.cb_trans_type.get() if hasattr(self, "cb_trans_type") else "служебная",
                "form_type": self.cb_form_type.get() if hasattr(self, "cb_form_type") else "ЭО-2621 (Форма ЭСМ-2 / 412-АПК)",
                "work_type": self.cb_work_type.get() if hasattr(self, "cb_work_type") else "",
                "winter": safe_float(self.var_winter.get(), 0.0) if hasattr(self, "var_winter") else 0.0,
                "city": safe_float(self.var_city.get(), 0.0) if hasattr(self, "var_city") else 0.0,
                "dirt": safe_float(self.var_dirt.get(), 0.0) if hasattr(self, "var_dirt") else 0.0,
                "trailer": safe_float(self.var_trailer.get(), 0.0) if hasattr(self, "var_trailer") else 0.0,
                "wb_num": self.en_wb_num.get() if hasattr(self, "en_wb_num") else "1",
                "date": self.en_date.get() if hasattr(self, "en_date") else "",
                "t_out": self.en_t_out.get() if hasattr(self, "en_t_out") else "08:00",
                "t_in": self.en_t_in.get() if hasattr(self, "en_t_in") else "17:00",
                "f_a": self.en_f_a.get() if hasattr(self, "en_f_a") else "0",
                "f_act": self.en_f_act.get() if hasattr(self, "en_f_act") else "",
                "g_a": self.en_g_a.get() if hasattr(self, "en_g_a") else "0",
                "g_act": self.en_g_act.get() if hasattr(self, "en_g_act") else "",
                "fuel_type": (
                    self.var_fuel_type.get()
                    if hasattr(self, "var_fuel_type")
                    else "Бензин"
                ),
                "map_provider": getattr(self, "map_provider_name", "Яндекс Карты"),
                "compact_map_labels": (
                    bool(self.compact_labels_var.get())
                    if hasattr(self, "compact_labels_var")
                    else True
                ),
                "print_orientation": (
                    self.print_orientation_var.get()
                    if hasattr(self, "print_orientation_var")
                    else "Поперёк (A5 — 2 на лист A4)"
                ),
                "print_rotate_180": bool(getattr(self, "print_rotate_180", True)),
                "only_motohours": (
                    bool(self.var_only_motohours.get())
                    if hasattr(self, "var_only_motohours")
                    else False
                ),
                "segments": mem_segments,
                "tractor_works": list(getattr(self, "tractor_works", []) or []),
                "saved_templates": self.settings.get("saved_templates", {}),
            }
            self.settings.update(data)
            with open(self.memory_file, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _run_bg_backup(self):
        """
        Отложенное неблокирующее авто-копирование баз SQLite (без пережатия .xlsx и без зависаний UI).
        Если идёт серия сохранений подряд, запросы объединяются в одно сохранение через 2.5 сек.
        """
        if getattr(self, "_bg_backup_after_id", None):
            try:
                self.after_cancel(self._bg_backup_after_id)
            except Exception:
                pass

        def _kick():
            self._bg_backup_after_id = None
            threading.Thread(target=create_fast_db_backup, daemon=True).start()

        try:
            self._bg_backup_after_id = self.after(2500, _kick)
        except Exception:
            threading.Thread(target=create_fast_db_backup, daemon=True).start()

    def on_closing(self):
        self.save_memory()
        try:
            self.db_cars.close_connections()
            self.db_tractors.close_connections()
        except Exception:
            pass
        self.destroy()

    def get_names(self, table):
        return [r[1] for r in self.db.get_list(table)]

    def get_unique_locs(self):
        names = [r[1] for r in self.db.get_list("locs") if r[1]]
        names = list(set(names))
        names.sort()
        return names

    def get_sorted_history(self):
        rows = self.db.get_list("history")
        return sorted(
            rows,
            key=lambda r: _wb_chrono_key(
                r[2], r[1], r[0], r[3] if len(r) > 3 else None
            ),
        )

    def _handle_renamed_wb_excels(self, renamed_records, mode, exclude_ids=None):
        """Удаляет старые Excel-файлы с устаревшим № ПЛ и возвращает ID записей, чьи файлы нужно пересобрать с новым № ПЛ."""
        if not renamed_records:
            return []
        excl = set(exclude_ids or ())
        rebuild_ids = []
        for rec in renamed_records:
            new_id = rec.get("new_id")
            old_wb = rec.get("old_wb")
            new_wb = rec.get("new_wb")
            d_str = rec.get("date")
            veh = rec.get("veh")
            if not old_wb or old_wb == new_wb:
                continue
            old_path = self._resolve_waybill_excel_path(old_wb, d_str, veh, mode=mode)
            new_path = self._resolve_waybill_excel_path(new_wb, d_str, veh, mode=mode)
            if os.path.exists(old_path):
                if os.path.abspath(old_path) != os.path.abspath(new_path):
                    try:
                        os.remove(old_path)
                    except Exception:
                        pass
                if new_id not in excl:
                    rebuild_ids.append(new_id)
        return list(dict.fromkeys(rebuild_ids))

    def renumber_and_sort_history(self):
        self.db.renumber_history_ids_by_date(renumber_wb_num=True)

    def open_work_norm_mass_update(self, preselect_work=None, preselect_norm=None):
        """Открывает диалог массового обновления нормы вида работ по всем ПЛ или за выбранный период."""
        wts = self.db_tractors.get_list("work_types")
        if not wts:
            messagebox.showwarning(
                "Внимание", "Справочник видов работ пуст. Добавьте хотя бы один вид работ."
            )
            return
        hist_rows = self.db_tractors.get_list("history")
        dates_parsed = []
        for r in hist_rows:
            try:
                dates_parsed.append(datetime.strptime(str(r[2] or "").strip(), "%d.%m.%Y"))
            except Exception:
                pass
        def_from = min(dates_parsed).strftime("%d.%m.%Y") if dates_parsed else datetime.now().strftime("01.01.%Y")
        def_to = max(dates_parsed).strftime("%d.%m.%Y") if dates_parsed else datetime.now().strftime("%d.%m.%Y")

        dlg = WorkNormMassUpdateDialog(
            self,
            wts,
            preselect_work=preselect_work,
            preselect_norm=preselect_norm,
            default_from=def_from,
            default_to=def_to,
        )
        self.wait_window(dlg)
        res = dlg.result
        if not res:
            return

        _, renamed_records = self.db_tractors.renumber_history_by_date_detailed(
            renumber_wb_num=True
        )
        outcome = self.db_tractors.apply_work_type_norm(
            res["work_name"],
            res["new_norm"],
            date_from=res["date_from"],
            date_to=res["date_to"],
            update_reference=res["update_ref"],
        )
        matched_ids = list(outcome.get("matched_ids", []))
        cascaded_ids = list(outcome.get("cascaded_ids", []))
        renamed_excel_ids = self._handle_renamed_wb_excels(
            renamed_records, mode="tractors"
        )

        if res["rebuild_excel"]:
            if renamed_excel_ids:
                self._rebuild_excel_batch_bg(
                    renamed_excel_ids, mode="tractors", only_existing=False
                )
            rem_cascaded = [i for i in cascaded_ids if i not in set(renamed_excel_ids)]
            if rem_cascaded:
                self._rebuild_excel_batch_bg(
                    rem_cascaded, mode="tractors", only_existing=True
                )

        # Обновляем выбранные работы в текущей форме, если там присутствует этот вид работ
        target_low = res["work_name"].strip().lower()
        for w in self.tractor_works:
            if isinstance(w, dict) and str(w.get("name", "")).strip().lower() == target_low:
                w["norm_mh"] = round(float(res["new_norm"]), 2)
        self._last_works_render_sig = None

        self.refresh_all_tables()
        self.refresh_history()
        self.load_vehicle_data()
        self.recalc_live()
        self._run_bg_backup()

        period_txt = (
            f"за период {res['date_from']} – {res['date_to']}"
            if res["scope"] == "period"
            else "по всем путевым листам"
        )
        excel_note = (
            "\nСуществующие Excel-файлы обновляются в фоновом режиме."
            if (res["rebuild_excel"] and cascaded_ids)
            else ""
        )
        messagebox.showinfo(
            "Норма пересчитана",
            f"Вид работ: «{res['work_name']}» → {res['new_norm']:g} л/м·ч ({period_txt}).\n"
            f"Обновлено ПЛ с этим видом работ: {len(matched_ids)}\n"
            f"Всего пересчитано ПЛ по цепочке остатков: {len(cascaded_ids)}{excel_note}",
        )

    def do_backup(self):
        dialog = BackupChoiceDialog(self, active_mode=self.active_mode)
        self.wait_window(dialog)
        choice = dialog.choice
        if not choice:
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            self.save_memory()
            if choice == "cars":
                out_path = filedialog.asksaveasfilename(
                    title="Сохранить резервную копию — Только Легковые авто",
                    defaultextension=".zip",
                    initialfile=f"Резервная_копия_Легковые_авто_{ts}.zip",
                    filetypes=[("ZIP Archive", "*.zip")],
                )
                if not out_path:
                    return
                create_mode_backup(mode="cars", out_zip_path=out_path)
                self.log_action(f"Создана резервная копия Легковые авто: {out_path}")
                messagebox.showinfo(
                    "Успех",
                    f"Резервная копия базы Легковых автомобилей успешно сохранена!\n\n"
                    f"Файл: {out_path}\n"
                    f"Содержимое: waybills.db, memory.json и легковые путевые листы.",
                )
            elif choice == "tractors":
                out_path = filedialog.asksaveasfilename(
                    title="Сохранить резервную копию — Только Трактора МТЗ-82",
                    defaultextension=".zip",
                    initialfile=f"Резервная_копия_Трактора_МТЗ82_{ts}.zip",
                    filetypes=[("ZIP Archive", "*.zip")],
                )
                if not out_path:
                    return
                create_mode_backup(mode="tractors", out_zip_path=out_path)
                self.log_action(f"Создана резервная копия Трактора МТЗ-82: {out_path}")
                messagebox.showinfo(
                    "Успех",
                    f"Резервная копия базы Тракторов МТЗ-82 / Экскаваторов успешно сохранена!\n\n"
                    f"Файл: {out_path}\n"
                    f"Содержимое: tractors.db, memory_tractors.json и тракторные путевые листы.",
                )
            else:
                res = create_unified_backup()
                out_path = filedialog.asksaveasfilename(
                    title="Сохранить единую резервную копию (Легковые + Трактора МТЗ-82)",
                    defaultextension=".zip",
                    initialfile=f"Резервная_копия_Легковые_и_Трактора_{ts}.zip",
                    filetypes=[("ZIP Archive", "*.zip")],
                )
                if not out_path:
                    return
                shutil.copyfile(res["zip_path"], out_path)
                self.log_action("Создана единая резервная копия Легковые + Трактора (.zip + .db)")
                messagebox.showinfo(
                    "Успех",
                    f"Единая резервная копия (Легковые + Трактора МТЗ-82) успешно сохранена!\n"
                    f"ZIP-архив: {out_path}\n"
                    f"Единая SQLite БД: {res['unified_db_path']}",
                )
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def _get_staff_by_role(self, role_substr):
        """Возвращает список ФИО сотрудников для указанной роли (поддерживает 'Механик / Диспетчер')."""
        rows = self.db.get_list("staff")
        substr = str(role_substr or "").strip().lower()
        result = []
        seen = set()
        for r in rows:
            fio = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ""
            role = str(r[2]).strip().lower() if len(r) > 2 and r[2] is not None else ""
            if not fio:
                continue
            if not substr or substr in role:
                if fio not in seen:
                    seen.add(fio)
                    result.append(fio)
        if not result:
            for r in rows:
                fio = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ""
                if fio and fio not in seen:
                    seen.add(fio)
                    result.append(fio)
        return result or [""]

    def _get_tractor_norm_km(self, mark=None):
        """Всегда берёт норму расхода на перегон (л/100 км) автоматически из справочника выбранного трактора."""
        veh_name = (mark if mark is not None else (self.cb_veh.get().strip() if hasattr(self, "cb_veh") else "")).strip()
        try:
            v_list = self.db_tractors.get_list("vehicles")
            v_info = next((r for r in v_list if str(r[1]).strip() == veh_name), None)
            if not v_info and v_list:
                v_info = v_list[0]
            if v_info and len(v_info) > 6:
                val = safe_float(v_info[6], 0.0)
                if val > 0:
                    self._tractor_norm_km = val
                    return val
        except Exception:
            pass
        return safe_float(getattr(self, "_tractor_norm_km", 16.5), 16.5)

    def _update_mode_visibility(self):
        """Показывает/скрывает элементы формы в зависимости от режима ('cars' vs 'tractors')."""
        is_tr = self.active_mode == "tractors"

        # Адаптивный размер правой колонки: просторный классический вид для Легковых и компактный без прокрутки для Тракторов
        try:
            if hasattr(self, "route_frame"):
                self.route_frame.pack_configure(
                    padx=12 if is_tr else 14,
                    pady=(4, 3) if is_tr else (10, 8),
                )
            if hasattr(self, "pickup_row"):
                self.pickup_row.pack_configure(
                    padx=10 if is_tr else 12,
                    pady=(5, 2) if is_tr else (10, 4),
                )
            if hasattr(self, "lbl_pickup_title"):
                self.lbl_pickup_title.configure(
                    text="Адрес подачи:" if is_tr else "Адрес подачи (Начало маршрута):"
                )
            if hasattr(self, "add_frame"):
                self.add_frame.pack_configure(
                    padx=10 if is_tr else 12,
                    pady=2 if is_tr else 4,
                )
            if hasattr(self, "btn_r_frame"):
                self.btn_r_frame.pack_configure(
                    padx=10 if is_tr else 12,
                    pady=2 if is_tr else 4,
                )
            for b_name in ("btn_undo_rt", "btn_clear_rt", "btn_save_tmpl", "btn_load_tmpl", "btn_del_tmpl"):
                if hasattr(self, b_name):
                    getattr(self, b_name).configure(height=26 if is_tr else 30)
            if hasattr(self, "tmpl_frame"):
                self.tmpl_frame.pack_configure(
                    padx=10 if is_tr else 12,
                    pady=2 if is_tr else 4,
                )
            if hasattr(self, "txt_route_main"):
                self.txt_route_main.configure(height=68 if is_tr else 150)
                self.txt_route_main.pack_configure(
                    padx=10 if is_tr else 12,
                    pady=2 if is_tr else 6,
                )
            if hasattr(self, "f_dist"):
                self.f_dist.pack_configure(
                    padx=10 if is_tr else 12,
                    pady=(2, 5) if is_tr else (4, 10),
                )
            if hasattr(self, "odo_frame"):
                self.odo_frame.pack_configure(
                    padx=12 if is_tr else 14,
                    pady=3 if is_tr else 8,
                )
            if hasattr(self, "lbl_odo_title"):
                self.lbl_odo_title.pack_configure(pady=5 if is_tr else 10)
            if hasattr(self, "fuel_type_seg"):
                self.fuel_type_seg.pack_configure(
                    padx=12 if is_tr else 14,
                    pady=3 if is_tr else 8,
                )
            if hasattr(self, "grid_frame"):
                self.grid_frame.pack_configure(
                    padx=12 if is_tr else 14,
                    pady=3 if is_tr else 8,
                )
            if hasattr(self, "btn_action"):
                self.btn_action.pack_configure(
                    padx=12 if is_tr else 14,
                    pady=(4, 4) if is_tr else (12, 10),
                )
            if hasattr(self, "btn_save_db"):
                self.btn_save_db.configure(height=36 if is_tr else 42)
            if hasattr(self, "btn_gen"):
                self.btn_gen.configure(height=36 if is_tr else 42)
        except Exception:
            pass

        if hasattr(self, "tractor_frame") and hasattr(self, "fuel_type_seg"):
            try:
                self.tractor_frame.pack_forget()
            except Exception:
                pass
            if is_tr:
                self.tractor_frame.pack(
                    fill="x", padx=12, pady=3, before=self.fuel_type_seg
                )

        if hasattr(self, "chk_only_motohours"):
            try:
                self.chk_only_motohours.pack_forget()
            except Exception:
                pass
            if is_tr:
                self.chk_only_motohours.pack(side="left", padx=(10, 4))

        if hasattr(self, "lbl_veh_title"):
            self.lbl_veh_title.configure(
                text="Трактор / Спецтехника:" if is_tr else "Автомобиль:"
            )
        if hasattr(self, "lbl_trailer_title"):
            self.lbl_trailer_title.configure(
                text="Прицеп / Агрегат:" if is_tr else "Прицеп:"
            )
        if hasattr(self, "lbl_fuel_row1"):
            self.lbl_fuel_row1.configure(text="ДТ (дизель)" if is_tr else "Бензин")

        # В режиме тракторов скрываем столбец «Норма л/100км» — расход на перегон считается автоматически из справочника
        for norm_w_name in ("lbl_fuel_hdr_norm", "en_f_norm"):
            if hasattr(self, norm_w_name):
                try:
                    w = getattr(self, norm_w_name)
                    if is_tr:
                        w.grid_remove()
                    else:
                        w.grid()
                except Exception:
                    pass

        if hasattr(self, "gas_row_widgets"):
            for w in self.gas_row_widgets:
                try:
                    if is_tr:
                        w.grid_remove()
                    else:
                        w.grid()
                except Exception:
                    pass

        if hasattr(self, "fuel_type_seg") and hasattr(self, "var_fuel_type"):
            if is_tr:
                self.fuel_type_seg.configure(values=["ДТ (Дизтопливо)"])
                self.var_fuel_type.set("ДТ (Дизтопливо)")
            else:
                car_fuels = ["Бензин", "Газ", "Бензин + Газ"]
                self.fuel_type_seg.configure(values=car_fuels)
                if self.var_fuel_type.get() not in car_fuels:
                    self.var_fuel_type.set("Бензин")

        if hasattr(self, "cb_msg_type"):
            if is_tr:
                tr_msgs = [
                    "перегон своим ходом / работа на объекте",
                    "городское / пригородное",
                    "внутрихозяйственное",
                ]
                self.cb_msg_type.configure(values=tr_msgs)
                cur_msg = self.cb_msg_type.get().strip()
                saved_msg = safe_str(self.settings.get("msg_type"), "")
                if saved_msg in tr_msgs:
                    self.cb_msg_type.set(saved_msg)
                elif not cur_msg or cur_msg not in tr_msgs:
                    self.cb_msg_type.set(tr_msgs[0])
            else:
                car_msgs = [
                    "городское / пригородное",
                    "междугородное",
                    "служебная поездка",
                ]
                self.cb_msg_type.configure(values=car_msgs)
                cur_msg = self.cb_msg_type.get().strip()
                saved_msg = safe_str(self.settings.get("msg_type"), "")
                if saved_msg in car_msgs:
                    self.cb_msg_type.set(saved_msg)
                elif not cur_msg or "перегон" in cur_msg.lower() or "объекте" in cur_msg.lower():
                    self.cb_msg_type.set(car_msgs[0])

        if hasattr(self, "cb_trans_type"):
            if is_tr:
                tr_trans = [
                    "коммунальные / строительные работы",
                    "земляные работы / погрузка",
                    "транспортные работы с прицепом",
                ]
                self.cb_trans_type.configure(values=tr_trans)
                cur_tr = self.cb_trans_type.get().strip()
                saved_tr = safe_str(self.settings.get("trans_type"), "")
                if saved_tr in tr_trans:
                    self.cb_trans_type.set(saved_tr)
                elif not cur_tr or cur_tr not in tr_trans:
                    self.cb_trans_type.set(tr_trans[0])
            else:
                car_trans = [
                    "служебная",
                    "перевозка пассажиров и багажа",
                    "хозяйственные нужды",
                ]
                self.cb_trans_type.configure(values=car_trans)
                cur_tr = self.cb_trans_type.get().strip()
                saved_tr = safe_str(self.settings.get("trans_type"), "")
                if saved_tr in car_trans:
                    self.cb_trans_type.set(saved_tr)
                elif (
                    not cur_tr
                    or "коммунальные" in cur_tr.lower()
                    or "земляные" in cur_tr.lower()
                    or "прицепом" in cur_tr.lower()
                ):
                    self.cb_trans_type.set(car_trans[0])

        if hasattr(self, "lbl_fuel_hint"):
            if is_tr:
                self.lbl_fuel_hint.configure(
                    text="Расход на перегон и работы считается автоматически из справочника; если «Факт (л)» пусто — списывается ИТОГО ДТ"
                )
            else:
                self.lbl_fuel_hint.configure(
                    text="Норма = л/100 км из справочника; если «Факт (л)» пусто — списывается расчётный расход по норме"
                )

    def switch_browser_tab(self, selected_label):
        self.save_memory()
        self.editing_id = None
        self.old_filename = None
        self._editing_old_veh = None
        if hasattr(self, "btn_gen"):
            self.btn_gen.configure(text="Выписать путевой лист")
        if hasattr(self, "btn_save_db"):
            self.btn_save_db.configure(text="Внести в журнал")
        if "Трактора" in selected_label:
            self.active_mode = "tractors"
            self.db = self.db_tractors
            self.memory_file = "memory_tractors.json"
        else:
            self.active_mode = "cars"
            self.db = self.db_cars
            self.memory_file = "memory.json"
        self.settings = self.load_memory()
        self.route_segments = self.settings.get("segments", [])
        self.tractor_works = list(self.settings.get("tractor_works", []) or [])
        self._update_mode_visibility()
        self.refresh_all_tables()
        self.refresh_history()
        self._refresh_tractor_works_tree()
        try:
            if hasattr(self, "cb_org"):
                self.cb_org.set(safe_str(self.settings.get("org"), ""))
            if hasattr(self, "cb_client"):
                self.cb_client.set(safe_str(self.settings.get("client"), ""))
            if hasattr(self, "cb_pickup"):
                self.cb_pickup.set(safe_str(self.settings.get("pickup"), ""))
            if hasattr(self, "cb_drv"):
                self.cb_drv.set(safe_str(self.settings.get("driver") or self.settings.get("drv"), ""))
            if hasattr(self, "cb_veh"):
                self.cb_veh.set(safe_str(self.settings.get("vehicle") or self.settings.get("veh"), ""))
            if hasattr(self, "cb_mech"):
                self.cb_mech.set(safe_str(self.settings.get("mechanic") or self.settings.get("mech"), ""))
            if hasattr(self, "cb_disp"):
                self.cb_disp.set(safe_str(self.settings.get("dispatcher") or self.settings.get("disp"), ""))
            if hasattr(self, "en_wb_num"):
                self.en_wb_num.delete(0, "end")
                self.en_wb_num.insert(0, safe_str(self.settings.get("wb_num"), "1"))
            if hasattr(self, "en_date"):
                self.en_date.delete(0, "end")
                self.en_date.insert(
                    0,
                    safe_str(self.settings.get("date"), datetime.now().strftime("%d.%m.%Y")),
                )
            if hasattr(self, "var_only_motohours"):
                self.var_only_motohours.set(
                    bool(self.settings.get("only_motohours", False)) if self.active_mode == "tractors" else False
                )
            if hasattr(self, "var_winter"):
                self.var_winter.set(safe_float(self.settings.get("winter", self.settings.get("season", 0.0)), 0.0))
            if hasattr(self, "var_city"):
                self.var_city.set(safe_float(self.settings.get("city", 0.0), 0.0))
            if hasattr(self, "var_dirt"):
                self.var_dirt.set(safe_float(self.settings.get("dirt", 0.0), 0.0))
            if hasattr(self, "var_trailer"):
                self.var_trailer.set(safe_float(self.settings.get("trailer", 0.0), 0.0))
        except Exception:
            pass
        self.load_vehicle_data()
        # Синхронизируем дистанцию под маршруты новой вкладки, чтобы не переносилось значение из другого режима
        if hasattr(self, "en_dist"):
            new_tab_km = sum(safe_float(s.get("km", 0.0), 0.0) for s in (self.route_segments or []))
            self.en_dist.delete(0, "end")
            self.en_dist.insert(0, f"{new_tab_km:.1f}")
        self.update_route_displays()
        self.after(60, self.redraw_map_from_segments)
        self.recalc_live()

    def restore_backup(self):
        import zipfile

        file_path = filedialog.askopenfilename(
            title="Восстановление",
            filetypes=[
                ("Archives & Databases", "*.zip *.db"),
                ("Database", "*.db"),
            ],
        )
        if not file_path:
            return
        if not messagebox.askyesno(
            "ВНИМАНИЕ!",
            "Текущая база данных и все ваши шаблоны маршрутов будут полностью заменены.\nПродолжить?",
        ):
            return
        try:
            self.db_cars.close_connections()
            self.db_tractors.close_connections()
            # Удаляем устаревшие WAL/SHM файлы перед заменой БД, чтобы старые транзакции не наложились на восстановленную базу
            for stale_f in (
                "waybills.db-wal",
                "waybills.db-shm",
                "tractors.db-wal",
                "tractors.db-shm",
            ):
                if os.path.exists(stale_f):
                    try:
                        os.remove(stale_f)
                    except Exception:
                        pass
            if file_path.endswith(".zip"):
                with zipfile.ZipFile(file_path, "r") as zipf:
                    zipf.extractall(".")
            else:
                target_db_file = "tractors.db" if self.active_mode == "tractors" else "waybills.db"
                shutil.copyfile(file_path, target_db_file)
            # Запускаем миграцию схемы, затем хронологическую перенумерацию ID и № ПЛ (1..N) и полный пересчёт цепочек
            self.db_cars.init_db(auto_sync_history=False)
            self.db_tractors.init_db(auto_sync_history=False)
            _, cars_renamed = self.db_cars.renumber_history_by_date_detailed(renumber_wb_num=True)
            cars_recalc = self.db_cars.recalculate_all_vehicles() or []
            _, tr_renamed = self.db_tractors.renumber_history_by_date_detailed(renumber_wb_num=True)
            tr_recalc = self.db_tractors.recalculate_all_vehicles() or []

            cars_ren_ids = self._handle_renamed_wb_excels(cars_renamed, mode="cars")
            if cars_ren_ids:
                self._rebuild_excel_batch_bg(cars_ren_ids, mode="cars", only_existing=False)
            cars_rem = [i for i in cars_recalc if i not in set(cars_ren_ids)]
            if cars_rem:
                self._rebuild_excel_batch_bg(cars_rem, mode="cars", only_existing=True)

            tr_ren_ids = self._handle_renamed_wb_excels(tr_renamed, mode="tractors")
            if tr_ren_ids:
                self._rebuild_excel_batch_bg(tr_ren_ids, mode="tractors", only_existing=False)
            tr_rem = [i for i in tr_recalc if i not in set(tr_ren_ids)]
            if tr_rem:
                self._rebuild_excel_batch_bg(tr_rem, mode="tractors", only_existing=True)

            self._route_mem_cache.clear()
            self.settings = self.load_memory()
            self.route_segments = self.settings.get("segments", [])
            self.tractor_works = list(self.settings.get("tractor_works", []) or [])
            self.selected_marker_index = None
            self._update_mode_visibility()
            self.refresh_all_tables()
            self.refresh_history()
            try:
                if hasattr(self, "cb_org"):
                    self.cb_org.set(safe_str(self.settings.get("org"), ""))
                if hasattr(self, "cb_client"):
                    self.cb_client.set(safe_str(self.settings.get("client"), ""))
                if hasattr(self, "cb_pickup"):
                    self.cb_pickup.set(safe_str(self.settings.get("pickup"), ""))
                if hasattr(self, "cb_drv"):
                    self.cb_drv.set(safe_str(self.settings.get("driver") or self.settings.get("drv"), ""))
                if hasattr(self, "cb_veh"):
                    self.cb_veh.set(safe_str(self.settings.get("vehicle") or self.settings.get("veh"), ""))
                if hasattr(self, "cb_mech"):
                    self.cb_mech.set(safe_str(self.settings.get("mechanic") or self.settings.get("mech"), ""))
                if hasattr(self, "cb_disp"):
                    self.cb_disp.set(safe_str(self.settings.get("dispatcher") or self.settings.get("disp"), ""))
                if hasattr(self, "en_wb_num"):
                    next_num = str(len(self.db.get_list("history")) + 1)
                    self.en_wb_num.delete(0, "end")
                    self.en_wb_num.insert(0, next_num)
            except Exception:
                pass
            self._refresh_tractor_works_tree()
            self.load_vehicle_data()
            self.update_route_displays()
            self.after(60, lambda: self.redraw_map_from_segments(auto_fit=True))
            self.recalc_live()
            messagebox.showinfo(
                "Успех",
                "Данные успешно восстановлены!\n"
                "Журнал автоматически отсортирован по датам, номера ПЛ (1..N), расход/остаток топлива, одометры и моточасы пересчитаны.",
            )
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def exp_w(self):
        history_data = self.get_sorted_history()
        if not history_data:
            messagebox.showwarning(
                "Внимание",
                "Журнал путевых листов пуст. Нет данных для формирования отчета.",
            )
            return
        ok, res = generate_word_report(history_data)
        if ok:
            messagebox.showinfo(
                "Успех", f"Журнал (Форма № 8) успешно сформирован:\n{res}"
            )
            if hasattr(os, "startfile"):
                os.startfile(os.path.abspath(res))
        else:
            messagebox.showerror(
                "Ошибка", f"Не удалось сформировать Форму № 8:\n{res}"
            )

    def exp_xl(self):
        history_data = self.get_sorted_history()
        if not history_data:
            messagebox.showwarning(
                "Внимание",
                "Журнал путевых листов пуст. Нет данных для формирования отчета.",
            )
            return
        ok, res = export_history_to_excel(history_data)
        if ok:
            messagebox.showinfo("Успех", f"Реестр ГСМ успешно сформирован:\n{res}")
            if hasattr(os, "startfile"):
                os.startfile(os.path.abspath(res))
        else:
            messagebox.showerror(
                "Ошибка", f"Не удалось сформировать Реестр ГСМ:\n{res}"
            )

    def _on_main_tab_change(self):
        """Отрисовывает карту только при переходе на вкладку «Маршрут на карте» (убирает фоновые тормоза)."""
        try:
            if self.tabview.get() == "Маршрут на карте" and getattr(self, "_map_dirty", False):
                self._map_dirty = False
                self.after(30, lambda: self.redraw_map_from_segments(auto_fit=True, force=True))
        except Exception:
            pass

    def create_widgets(self):
        # Верхняя панель вкладок как в браузере (Легковые автомобили / Трактора МТЗ-82 и Экскаваторы)
        browser_bar = ctk.CTkFrame(self, fg_color="#E2E8F0", height=36, corner_radius=0)
        browser_bar.pack(fill="x", padx=0, pady=0)

        self.var_browser_tab = tk.StringVar(
            value="🚗 Легковые автомобили (Форма № 3 — waybills.db)"
        )
        self.browser_tabs = ctk.CTkSegmentedButton(
            browser_bar,
            values=[
                "🚗 Легковые автомобили (Форма № 3 — waybills.db)",
                "🚜 Трактора МТЗ-82 и Экскаваторы (Форма 412-АПК / ЭСМ-2 — tractors.db)",
            ],
            variable=self.var_browser_tab,
            command=self.switch_browser_tab,
            font=("Helvetica", 12, "bold"),
            selected_color="#2563EB",
            selected_hover_color="#1D4ED8",
        )
        self.browser_tabs.pack(side="left", padx=12, pady=3)

        header = ctk.CTkFrame(self, fg_color="#FFFFFF", height=44, corner_radius=0)
        header.pack(fill="x", padx=0, pady=0)

        ctk.CTkLabel(
            header,
            text="Про.Путевые Листы",
            font=("Helvetica", 18, "bold"),
            text_color="#1D1D1F",
        ).pack(side="left", padx=16, pady=6)

        btn_frame_top = ctk.CTkFrame(header, fg_color="transparent")
        btn_frame_top.pack(side="right", padx=14, pady=5)

        ctk.CTkButton(
            btn_frame_top,
            text="⚙️ Нормы работ → ПЛ",
            fg_color="#6366F1",
            hover_color="#4F46E5",
            width=165,
            height=30,
            command=self.open_work_norm_mass_update,
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            btn_frame_top,
            text="Восстановить базу",
            fg_color="#F59E0B",
            hover_color="#D97706",
            width=142,
            height=30,
            command=self.restore_backup,
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            btn_frame_top,
            text="📦 Резервная копия (Легковые / Трактора)",
            fg_color="#34C759",
            hover_color="#28A745",
            width=250,
            height=30,
            command=self.do_backup,
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            btn_frame_top,
            text="Форма № 8 (Печать)",
            fg_color="#007AFF",
            hover_color="#005CFF",
            width=142,
            height=30,
            command=self.exp_w,
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            btn_frame_top,
            text="Реестр ГСМ (Excel)",
            fg_color="#10B981",
            hover_color="#059669",
            width=142,
            height=30,
            command=self.exp_xl,
        ).pack(side="left", padx=4)

        self.tabview = ctk.CTkTabview(
            self, fg_color="#FFFFFF", corner_radius=10, command=self._on_main_tab_change
        )
        self.tabview.pack(fill="both", expand=True, padx=10, pady=(3, 5))

        self.tab_req = self.tabview.add("Заполнение путевого листа")
        self.tab_map = self.tabview.add("Маршрут на карте")
        self.tab_hist = self.tabview.add("Журнал путевых листов")
        self.tab_crud = self.tabview.add("Справочники системы")

        self.build_req_tab()
        self.build_map_tab()
        self.build_hist_tab()
        self.build_crud_tab()

        self.update_route_displays()
        self.calc_time()

    def build_req_tab(self):
        scroll = ctk.CTkScrollableFrame(self.tab_req, fg_color="transparent")
        scroll.pack(fill="both", expand=True)
        scroll.grid_columnconfigure(0, weight=1)
        scroll.grid_columnconfigure(1, weight=1)

        col1 = ctk.CTkFrame(scroll, fg_color="#F9FAFB", corner_radius=10)
        col1.grid(row=0, column=0, sticky="nsew", padx=6, pady=3)

        col2 = ctk.CTkFrame(scroll, fg_color="#F9FAFB", corner_radius=10)
        col2.grid(row=0, column=1, sticky="nsew", padx=6, pady=3)

        # Левая колонка: классические просторные реквизиты путевого листа (единый вид для легковых и тракторов)
        f_top = ctk.CTkFrame(col1, fg_color="transparent")
        f_top.pack(fill="x", padx=14, pady=(6, 2))

        ctk.CTkLabel(f_top, text="№ ПЛ:", font=("Helvetica", 13, "bold")).pack(
            side="left"
        )
        self.en_wb_num = ctk.CTkEntry(f_top, width=85)
        self.en_wb_num.pack(side="left", padx=8)
        self.en_wb_num.insert(0, safe_str(self.settings.get("wb_num"), "1"))

        ctk.CTkLabel(f_top, text="Дата:", font=("Helvetica", 13, "bold")).pack(
            side="left", padx=(14, 5)
        )
        self.en_date = ctk.CTkEntry(f_top, width=112)
        self.en_date.pack(side="left", padx=5)
        default_date = safe_str(
            self.settings.get("date"), datetime.now().strftime("%d.%m.%Y")
        )
        self.en_date.insert(0, default_date)

        ctk.CTkButton(
            f_top, text="📅", width=36, command=self.open_calendar
        ).pack(side="left", padx=4)

        f_time = ctk.CTkFrame(col1, fg_color="transparent")
        f_time.pack(fill="x", padx=15, pady=5)

        ctk.CTkLabel(f_time, text="Выезд:").pack(side="left")
        self.en_t_out = ctk.CTkEntry(f_time, width=72)
        self.en_t_out.pack(side="left", padx=5)
        self.en_t_out.insert(0, safe_str(self.settings.get("t_out"), "08:00"))
        self.en_t_out.bind("<KeyRelease>", self.calc_time)

        ctk.CTkLabel(f_time, text="Возврат:").pack(side="left", padx=(10, 0))
        self.en_t_in = ctk.CTkEntry(f_time, width=72)
        self.en_t_in.pack(side="left", padx=5)
        self.en_t_in.insert(0, safe_str(self.settings.get("t_in"), "17:00"))
        self.en_t_in.bind("<KeyRelease>", self.calc_time)

        self.lbl_t_tot = ctk.CTkLabel(
            f_time, text="Итого: 9.0 ч", font=("Helvetica", 13, "bold")
        )
        self.lbl_t_tot.pack(side="left", padx=15)

        # Выпадающие списки (с поддержкой правого клика для переименования/удаления)
        ctk.CTkLabel(col1, text="Организация:").pack(anchor="w", padx=15, pady=(6, 2))
        self.cb_org = ctk.CTkComboBox(
            col1, values=self.get_names("orgs"), width=460
        )
        self.cb_org.pack(fill="x", padx=15)
        self.cb_org.set(safe_str(self.settings.get("org"), ""))
        self.cb_org.bind("<Button-3>", lambda e: self.show_context_menu(e, self.cb_org))

        ctk.CTkLabel(col1, text="Заказчик (В распоряжение):").pack(
            anchor="w", padx=15, pady=(6, 2)
        )
        self.cb_client = ctk.CTkComboBox(
            col1, values=self.get_names("clients"), width=460
        )
        self.cb_client.pack(fill="x", padx=15)
        self.cb_client.set(safe_str(self.settings.get("client"), ""))
        self.cb_client.bind("<Button-3>", lambda e: self.show_context_menu(e, self.cb_client))

        ctk.CTkLabel(col1, text="Водитель:").pack(anchor="w", padx=15, pady=(6, 2))
        self.cb_drv = ctk.CTkComboBox(
            col1,
            values=self.get_names("drivers"),
            width=460,
            command=self.auto_predict_driver_data,
        )
        self.cb_drv.pack(fill="x", padx=15)
        self.cb_drv.set(safe_str(self.settings.get("driver") or self.settings.get("drv"), ""))

        self.lbl_veh_title = ctk.CTkLabel(col1, text="Автомобиль:")
        self.lbl_veh_title.pack(anchor="w", padx=15, pady=(6, 2))
        self.cb_veh = ctk.CTkComboBox(
            col1,
            values=self.get_names("vehicles"),
            width=460,
            command=self.load_vehicle_data,
        )
        self.cb_veh.pack(fill="x", padx=15)
        self.cb_veh.set(safe_str(self.settings.get("vehicle") or self.settings.get("veh"), ""))

        ctk.CTkLabel(col1, text="Механик:").pack(anchor="w", padx=15, pady=(6, 2))
        staff_mechs = self._get_staff_by_role("механ")
        self.cb_mech = ctk.CTkComboBox(col1, values=staff_mechs, width=460)
        self.cb_mech.pack(fill="x", padx=15)
        self.cb_mech.set(safe_str(self.settings.get("mechanic") or self.settings.get("mech"), ""))

        ctk.CTkLabel(col1, text="Диспетчер:").pack(anchor="w", padx=15, pady=(6, 2))
        staff_disps = self._get_staff_by_role("диспетч")
        self.cb_disp = ctk.CTkComboBox(col1, values=staff_disps, width=460)
        self.cb_disp.pack(fill="x", padx=15)
        self.cb_disp.set(safe_str(self.settings.get("dispatcher") or self.settings.get("disp"), ""))

        # Дополнительные реквизиты путевого листа (ОКПО, Прицеп/Агрегат, Вид сообщения, Вид перевозки)
        f_extra_req = ctk.CTkFrame(col1, fg_color="transparent")
        f_extra_req.pack(fill="x", padx=15, pady=(8, 2))
        f_extra_req.grid_columnconfigure(3, weight=1)
        ctk.CTkLabel(f_extra_req, text="Код ОКПО:").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=3)
        self.en_okpo = ctk.CTkEntry(f_extra_req, width=110)
        self.en_okpo.grid(row=0, column=1, sticky="w", padx=(0, 12), pady=3)
        self.en_okpo.insert(0, safe_str(self.settings.get("okpo"), "48291034"))

        self.lbl_trailer_title = ctk.CTkLabel(f_extra_req, text="Прицеп:")
        self.lbl_trailer_title.grid(row=0, column=2, sticky="w", padx=(0, 6), pady=3)
        self.en_trailer = ctk.CTkEntry(f_extra_req, width=190)
        self.en_trailer.grid(row=0, column=3, sticky="ew", pady=3)
        self.en_trailer.insert(0, safe_str(self.settings.get("trailer_name"), "без прицепа"))

        ctk.CTkLabel(f_extra_req, text="Вид сообщения:").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=3)
        self.cb_msg_type = ctk.CTkComboBox(
            f_extra_req,
            values=[
                "городское / пригородное",
                "междугородное",
                "служебная поездка",
            ],
            width=340,
        )
        self.cb_msg_type.grid(row=1, column=1, columnspan=3, sticky="ew", pady=3)
        self.cb_msg_type.set(safe_str(self.settings.get("msg_type"), "городское / пригородное"))

        ctk.CTkLabel(f_extra_req, text="Вид перевозки:").grid(row=2, column=0, sticky="w", padx=(0, 6), pady=3)
        self.cb_trans_type = ctk.CTkComboBox(
            f_extra_req,
            values=[
                "служебная",
                "перевозка пассажиров и багажа",
                "хозяйственные нужды",
            ],
            width=340,
        )
        self.cb_trans_type.grid(row=2, column=1, columnspan=3, sticky="ew", pady=3)
        self.cb_trans_type.set(safe_str(self.settings.get("trans_type"), "служебная"))

        # Коэффициенты расхода
        f_coef = ctk.CTkFrame(col1, fg_color="transparent")
        f_coef.pack(fill="x", padx=15, pady=10)

        self.var_winter = tk.DoubleVar(
            value=safe_float(self.settings.get("winter", self.settings.get("season", 0.0)), 0.0)
        )
        self.var_city = tk.DoubleVar(value=safe_float(self.settings.get("city", 0.0), 0.0))
        self.var_dirt = tk.DoubleVar(value=safe_float(self.settings.get("dirt", 0.0), 0.0))
        self.var_trailer = tk.DoubleVar(value=safe_float(self.settings.get("trailer", 0.0), 0.0))

        ctk.CTkCheckBox(
            f_coef,
            text="Зимняя надбавка (+10%)",
            variable=self.var_winter,
            onvalue=10.0,
            offvalue=0.0,
            command=self.recalc_live,
        ).grid(row=0, column=0, sticky="w", padx=6, pady=4)

        ctk.CTkCheckBox(
            f_coef,
            text="Городской режим (+5%)",
            variable=self.var_city,
            onvalue=5.0,
            offvalue=0.0,
            command=self.recalc_live,
        ).grid(row=0, column=1, sticky="w", padx=6, pady=4)

        ctk.CTkCheckBox(
            f_coef,
            text="Грунт / Бездорожье (+10%)",
            variable=self.var_dirt,
            onvalue=10.0,
            offvalue=0.0,
            command=self.recalc_live,
        ).grid(row=1, column=0, sticky="w", padx=6, pady=4)

        ctk.CTkCheckBox(
            f_coef,
            text="Работа с прицепом (+20%)",
            variable=self.var_trailer,
            onvalue=20.0,
            offvalue=0.0,
            command=self.recalc_live,
        ).grid(row=1, column=1, sticky="w", padx=6, pady=4)

        # Правая колонка: маршрут и ГСМ (динамически подстраивается: просторная для Легковых, компактная без прокрутки для Тракторов)
        self.route_frame = ctk.CTkFrame(col2, fg_color="#FFFFFF", corner_radius=8)
        self.route_frame.pack(fill="x", padx=14, pady=(10, 8))

        self.pickup_row = ctk.CTkFrame(self.route_frame, fg_color="transparent")
        self.pickup_row.pack(fill="x", padx=12, pady=(10, 4))
        self.lbl_pickup_title = ctk.CTkLabel(
            self.pickup_row,
            text="Адрес подачи (Начало маршрута):",
            font=("Helvetica", 12, "bold"),
        )
        self.lbl_pickup_title.pack(side="left", padx=(0, 8))

        locs = self.get_unique_locs()
        self.cb_pickup = ctk.CTkComboBox(self.pickup_row, values=locs, width=400)
        self.cb_pickup.pack(side="left", fill="x", expand=True)
        self.cb_pickup.set(safe_str(self.settings.get("pickup"), ""))
        self.cb_pickup.bind("<Button-3>", lambda e: self.show_context_menu(e, self.cb_pickup))

        self.add_frame = ctk.CTkFrame(self.route_frame, fg_color="transparent")
        self.add_frame.pack(fill="x", padx=12, pady=4)

        self.cb_route_add = ctk.CTkComboBox(self.add_frame, values=locs, width=300)
        self.cb_route_add.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.cb_route_add.bind("<Button-3>", lambda e: self.show_context_menu(e, self.cb_route_add))

        ctk.CTkButton(
            self.add_frame,
            text="+ Добавить точку",
            width=130,
            command=self.add_rt_point,
        ).pack(side="left")

        self.btn_r_frame = ctk.CTkFrame(self.route_frame, fg_color="transparent")
        self.btn_r_frame.pack(fill="x", padx=12, pady=4)

        self.btn_undo_rt = ctk.CTkButton(
            self.btn_r_frame,
            text="Отменить шаг",
            width=112,
            height=30,
            fg_color="#6B7280",
            command=self.undo_rt_point,
        )
        self.btn_undo_rt.pack(side="left", padx=(0, 5))

        self.btn_clear_rt = ctk.CTkButton(
            self.btn_r_frame,
            text="Очистить маршрут",
            width=138,
            height=30,
            fg_color="#EF4444",
            command=self.clear_map,
        )
        self.btn_clear_rt.pack(side="left", padx=5)

        self.btn_save_tmpl = ctk.CTkButton(
            self.btn_r_frame,
            text="Сохранить шаблон",
            width=138,
            height=30,
            fg_color="#10B981",
            command=self.save_route_tmpl,
        )
        self.btn_save_tmpl.pack(side="left", padx=5)

        self.tmpl_frame = ctk.CTkFrame(self.route_frame, fg_color="transparent")
        self.tmpl_frame.pack(fill="x", padx=12, pady=4)
        tmpl_names = list(self.settings.get("saved_templates", {}).keys()) or [""]
        self.cb_tmpl = ctk.CTkComboBox(self.tmpl_frame, values=tmpl_names, width=240)
        self.cb_tmpl.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.btn_load_tmpl = ctk.CTkButton(
            self.tmpl_frame, text="Загрузить", width=92, height=30, command=self.load_route_tmpl
        )
        self.btn_load_tmpl.pack(side="left", padx=3)
        self.btn_del_tmpl = ctk.CTkButton(
            self.tmpl_frame,
            text="Удалить",
            width=86,
            height=30,
            fg_color="#EF4444",
            command=self.delete_route_tmpl,
        )
        self.btn_del_tmpl.pack(side="left", padx=3)

        self.txt_route_main = ctk.CTkTextbox(self.route_frame, height=150)
        self.txt_route_main.pack(fill="x", padx=12, pady=6)

        self.f_dist = ctk.CTkFrame(self.route_frame, fg_color="transparent")
        self.f_dist.pack(fill="x", padx=12, pady=(4, 10))
        ctk.CTkLabel(
            self.f_dist, text="Дистанция:", font=("Helvetica", 12, "bold")
        ).pack(side="left")
        self.en_dist = ctk.CTkEntry(self.f_dist, width=90)
        self.en_dist.pack(side="left", padx=6)
        self.en_dist.insert(0, "0.0")
        self.en_dist.bind("<KeyRelease>", self._schedule_recalc)

        self.var_only_motohours = tk.BooleanVar(
            value=bool(self.settings.get("only_motohours", False))
        )
        self.chk_only_motohours = ctk.CTkCheckBox(
            self.f_dist,
            text="Не учитывать км (только моточасы)",
            variable=self.var_only_motohours,
            font=("Helvetica", 12, "bold"),
            command=self.recalc_live,
        )
        self.chk_only_motohours.pack(side="left", padx=(10, 4))

        self.lbl_dist_main = ctk.CTkLabel(
            self.f_dist, text="0.0 км", font=("Helvetica", 13, "bold"), text_color="#007AFF"
        )
        self.lbl_dist_main.pack(side="right")

        # Одометр
        self.odo_frame = ctk.CTkFrame(col2, fg_color="#FFFFFF", corner_radius=8)
        self.odo_frame.pack(fill="x", padx=14, pady=8)

        self.lbl_odo_title = ctk.CTkLabel(
            self.odo_frame, text="Одометр СТАРТ:", font=("Helvetica", 12, "bold")
        )
        self.lbl_odo_title.pack(side="left", padx=10, pady=10)
        self.en_odo_s = ctk.CTkEntry(self.odo_frame, width=105)
        self.en_odo_s.pack(side="left", padx=6)
        self.en_odo_s.insert(0, "0")
        self.en_odo_s.bind("<KeyRelease>", self._schedule_recalc)

        self.lbl_odo_e = ctk.CTkLabel(
            self.odo_frame,
            text="КОНЕЦ: 0.0 км",
            font=("Helvetica", 12, "bold"),
            text_color="#10B981",
        )
        self.lbl_odo_e.pack(side="right", padx=10)

        # Блок учёта моточасов и нескольких видов работ для Трактора МТЗ-82 / Экскаватора
        self.tractor_frame = ctk.CTkFrame(col2, fg_color="#FEF3C7", corner_radius=8)
        self.tractor_frame.pack(fill="x", padx=12, pady=3)

        tr_top = ctk.CTkFrame(self.tractor_frame, fg_color="transparent")
        tr_top.pack(fill="x", padx=10, pady=(5, 2))
        ctk.CTkLabel(
            tr_top,
            text="🚜 Бланк и виды работ (Трактор МТЗ-82 / Экскаватор):",
            font=("Helvetica", 12, "bold"),
            text_color="#92400E",
        ).pack(side="left")

        self.cb_form_type = ctk.CTkComboBox(
            tr_top,
            values=[
                "Форма № 412-АПК (Трактор МТЗ-82)",
                "ЭО-2621 (Форма ЭСМ-2 / 412-АПК)",
            ],
            width=250,
        )
        self.cb_form_type.pack(side="right")
        self.cb_form_type.set(
            safe_str(self.settings.get("form_type"), "ЭО-2621 (Форма ЭСМ-2 / 412-АПК)")
        )

        wt_names = [r[1] for r in self.db_tractors.get_list("work_types")] or [
            "Экскавация грунта I-II кат. (обратная лопата)"
        ]
        tr_add_row = ctk.CTkFrame(self.tractor_frame, fg_color="transparent")
        tr_add_row.pack(fill="x", padx=10, pady=2)
        self.cb_work_type = ctk.CTkComboBox(
            tr_add_row,
            values=wt_names,
            width=310,
            command=self.on_work_type_change,
        )
        self.cb_work_type.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.cb_work_type.set(safe_str(self.settings.get("work_type"), wt_names[0]))

        ctk.CTkLabel(tr_add_row, text="м/ч:").pack(side="left", padx=(2, 2))
        self.en_work_mh = ctk.CTkEntry(tr_add_row, width=55)
        self.en_work_mh.pack(side="left", padx=(0, 6))
        self.en_work_mh.insert(0, "2.0")

        ctk.CTkButton(
            tr_add_row,
            text="+ Добавить работу",
            width=128,
            height=26,
            fg_color="#D97706",
            hover_color="#B45309",
            command=self.add_tractor_work,
        ).pack(side="left")

        # Список выбранных работ за смену (5 видимых строк + скролл)
        tr_list_wrap = ctk.CTkFrame(self.tractor_frame, fg_color="#FFFFFF", corner_radius=6)
        tr_list_wrap.pack(fill="x", padx=10, pady=(2, 3))
        self.tr_works = ttk.Treeview(
            tr_list_wrap,
            columns=("Вид работ", "М/ч", "Норма", "ДТ (л)"),
            show="headings",
            height=5,
        )
        self.tr_works.heading("Вид работ", text="Выбранные работы (попадут в ПЛ)")
        self.tr_works.heading("М/ч", text="М/ч")
        self.tr_works.heading("Норма", text="л/м·ч")
        self.tr_works.heading("ДТ (л)", text="ДТ (л)")
        self.tr_works.column("Вид работ", width=310, anchor="w")
        self.tr_works.column("М/ч", width=55, anchor="center")
        self.tr_works.column("Норма", width=55, anchor="center")
        self.tr_works.column("ДТ (л)", width=65, anchor="center")
        tr_works_sb = ttk.Scrollbar(tr_list_wrap, orient="vertical", command=self.tr_works.yview)
        self.tr_works.configure(yscrollcommand=tr_works_sb.set)
        self.tr_works.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=3)
        tr_works_sb.pack(side="left", fill="y", pady=3)

        tr_w_btns = ctk.CTkFrame(tr_list_wrap, fg_color="transparent")
        tr_w_btns.pack(side="right", padx=5, pady=4)
        ctk.CTkButton(
            tr_w_btns,
            text="Удалить",
            width=74,
            height=26,
            fg_color="#EF4444",
            hover_color="#DC2626",
            font=("Helvetica", 12),
            command=self.remove_tractor_work,
        ).pack(pady=(0, 5))
        ctk.CTkButton(
            tr_w_btns,
            text="Очистить",
            width=74,
            height=26,
            fg_color="#6B7280",
            hover_color="#4B5563",
            font=("Helvetica", 12),
            command=self.clear_tractor_works,
        ).pack()

        tr_moto = ctk.CTkFrame(self.tractor_frame, fg_color="transparent")
        tr_moto.pack(fill="x", padx=10, pady=(2, 2))
        ctk.CTkLabel(tr_moto, text="Моточасы СТАРТ:").pack(side="left")
        self.en_moto_s = ctk.CTkEntry(tr_moto, width=76)
        self.en_moto_s.pack(side="left", padx=4)
        self.en_moto_s.insert(0, "0.0")
        self.en_moto_s.bind("<KeyRelease>", self._schedule_recalc)

        ctk.CTkLabel(tr_moto, text="Итого м/ч:").pack(side="left", padx=(6, 0))
        self.en_moto_h = ctk.CTkEntry(tr_moto, width=58)
        self.en_moto_h.pack(side="left", padx=4)
        self.en_moto_h.insert(0, "6.5")
        self.en_moto_h.bind("<KeyRelease>", self._schedule_recalc)

        ctk.CTkLabel(tr_moto, text="Норма л/м·ч:").pack(side="left", padx=(6, 0))
        self.en_norm_mh = ctk.CTkEntry(tr_moto, width=56)
        self.en_norm_mh.pack(side="left", padx=4)
        self.en_norm_mh.insert(0, "5.2")
        self.en_norm_mh.bind("<KeyRelease>", self._schedule_recalc)

        self.lbl_moto_e = ctk.CTkLabel(
            tr_moto,
            text="КОН: 0.0 м/ч",
            font=("Helvetica", 12, "bold"),
            text_color="#B45309",
        )
        self.lbl_moto_e.pack(side="right")

        self.lbl_dt_summary = ctk.CTkLabel(
            self.tractor_frame,
            text="Расход ДТ: Перегон 0.0 л + Моточасы 0.0 л = ИТОГО 0.0 л",
            font=("Helvetica", 11, "bold"),
            text_color="#78350F",
        )
        self.lbl_dt_summary.pack(anchor="w", padx=10, pady=(1, 5))
        self._refresh_tractor_works_tree()

        self.var_fuel_type = tk.StringVar(
            value=safe_str(self.settings.get("fuel_type"), "Бензин")
        )
        self.fuel_type_seg = ctk.CTkSegmentedButton(
            col2,
            values=["Бензин", "Газ", "Бензин + Газ"],
            variable=self.var_fuel_type,
            command=self.recalc_live,
        )
        self.fuel_type_seg.pack(fill="x", padx=12, pady=3)

        self.grid_frame = ctk.CTkFrame(col2, fg_color="#FFFFFF", corner_radius=8)
        self.grid_frame.pack(fill="x", padx=14, pady=8)

        headers = ["Топливо", "Выезд", "Заправка", "Норма л/100км", "Факт (л)", "Возврат"]
        for c_idx, h_text in enumerate(headers):
            hdr_lbl = ctk.CTkLabel(
                self.grid_frame, text=h_text, font=("Helvetica", 12, "bold")
            )
            hdr_lbl.grid(row=0, column=c_idx, padx=6, pady=3)
            if c_idx == 3:
                self.lbl_fuel_hdr_norm = hdr_lbl

        # Строка Бензин / ДТ
        self.lbl_fuel_row1 = ctk.CTkLabel(self.grid_frame, text="Бензин")
        self.lbl_fuel_row1.grid(row=1, column=0, padx=6, pady=3)
        self.en_f_s = ctk.CTkEntry(self.grid_frame, width=75)
        self.en_f_s.grid(row=1, column=1, padx=5, pady=3)
        self.en_f_s.insert(0, "0")
        self.en_f_s.bind("<KeyRelease>", self._schedule_recalc)

        self.en_f_a = ctk.CTkEntry(self.grid_frame, width=75)
        self.en_f_a.grid(row=1, column=2, padx=5, pady=3)
        self.en_f_a.insert(0, safe_str(self.settings.get("f_a"), "0"))
        self.en_f_a.bind("<KeyRelease>", self._schedule_recalc)

        self.en_f_norm = ctk.CTkEntry(self.grid_frame, width=75)
        self.en_f_norm.grid(row=1, column=3, padx=5, pady=3)
        self.en_f_norm.insert(0, "10.0")
        self.en_f_norm.bind("<KeyRelease>", self._schedule_recalc)

        self.en_f_act = ctk.CTkEntry(self.grid_frame, width=75)
        self.en_f_act.grid(row=1, column=4, padx=5, pady=3)
        self.en_f_act.bind("<KeyRelease>", self._schedule_recalc)

        self.en_f_e = ctk.CTkEntry(self.grid_frame, width=80)
        self.en_f_e.grid(row=1, column=5, padx=5, pady=3)

        # Строка Газ
        lbl_gas = ctk.CTkLabel(self.grid_frame, text="Газ")
        lbl_gas.grid(row=2, column=0, padx=6, pady=3)
        self.en_g_s = ctk.CTkEntry(self.grid_frame, width=75)
        self.en_g_s.grid(row=2, column=1, padx=5, pady=3)
        self.en_g_s.insert(0, "0")
        self.en_g_s.bind("<KeyRelease>", self._schedule_recalc)

        self.en_g_a = ctk.CTkEntry(self.grid_frame, width=75)
        self.en_g_a.grid(row=2, column=2, padx=5, pady=3)
        self.en_g_a.insert(0, safe_str(self.settings.get("g_a"), "0"))
        self.en_g_a.bind("<KeyRelease>", self._schedule_recalc)

        self.en_g_norm = ctk.CTkEntry(self.grid_frame, width=75)
        self.en_g_norm.grid(row=2, column=3, padx=5, pady=3)
        self.en_g_norm.insert(0, "0.0")
        self.en_g_norm.bind("<KeyRelease>", self._schedule_recalc)

        self.en_g_act = ctk.CTkEntry(self.grid_frame, width=75)
        self.en_g_act.grid(row=2, column=4, padx=5, pady=3)
        self.en_g_act.bind("<KeyRelease>", self._schedule_recalc)

        self.en_g_e = ctk.CTkEntry(self.grid_frame, width=80)
        self.en_g_e.grid(row=2, column=5, padx=5, pady=3)

        self.lbl_fuel_hint = ctk.CTkLabel(
            self.grid_frame,
            text="Норма = л/100 км из справочника; если «Факт (л)» пусто — списывается расчётный расход",
            font=("Helvetica", 10),
            text_color="#6B7280",
        )
        self.lbl_fuel_hint.grid(row=3, column=0, columnspan=6, sticky="w", padx=8, pady=(0, 4))

        self.gas_row_widgets = [
            lbl_gas,
            self.en_g_s,
            self.en_g_a,
            self.en_g_norm,
            self.en_g_act,
            self.en_g_e,
        ]

        # Кнопки сохранения и выписки ПЛ
        self.btn_action = ctk.CTkFrame(col2, fg_color="transparent")
        self.btn_action.pack(fill="x", padx=14, pady=(12, 10))

        self.btn_save_db = ctk.CTkButton(
            self.btn_action,
            text="Внести в журнал",
            fg_color="#10B981",
            hover_color="#059669",
            height=42,
            command=lambda: self.generate_main(create_excel=False),
        )
        self.btn_save_db.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self.btn_gen = ctk.CTkButton(
            self.btn_action,
            text="Выписать путевой лист",
            fg_color="#007AFF",
            hover_color="#005CFF",
            height=42,
            command=lambda: self.generate_main(create_excel=True),
        )
        self.btn_gen.pack(side="left", fill="x", expand=True, padx=(6, 0))

        # Сразу применяем видимость элементов под текущий режим ('cars' или 'tractors')
        self._update_mode_visibility()

    def build_map_tab(self):
        top = ctk.CTkFrame(self.tab_map, fg_color="transparent")
        top.pack(fill="x", padx=10, pady=(6, 2))

        left_ctrl = ctk.CTkFrame(top, fg_color="transparent")
        left_ctrl.pack(side="left", fill="x")

        ctk.CTkLabel(
            left_ctrl,
            text="Карта:",
            font=("Helvetica", 12, "bold"),
            text_color="#1F2937",
        ).pack(side="left", padx=(4, 6))

        self.map_provider_seg = ctk.CTkSegmentedButton(
            left_ctrl,
            values=list(MAP_PROVIDERS.keys()),
            command=self.on_map_provider_change,
        )
        self.map_provider_seg.pack(side="left", padx=(0, 10))
        self.map_provider_seg.set(self.map_provider_name)

        self.chk_compact_labels = ctk.CTkCheckBox(
            left_ctrl,
            text="Компактные подписи",
            variable=self.compact_labels_var,
            command=self.on_toggle_compact_labels,
            font=("Helvetica", 12),
            checkbox_width=18,
            checkbox_height=18,
        )
        self.chk_compact_labels.pack(side="left", padx=6)

        search_box = ctk.CTkFrame(top, fg_color="transparent")
        search_box.pack(side="right")

        self.entry_search = ctk.CTkEntry(
            search_box, placeholder_text="Поиск (напр. Буртинский, Оренбург Родимцева 16)...", width=260
        )
        self.entry_search.pack(side="left", padx=4)
        self.entry_search.bind("<Return>", lambda e: self.search_map(add_to_route=True))

        ctk.CTkButton(
            search_box,
            text="Найти и в маршрут",
            width=128,
            command=lambda: self.search_map(add_to_route=True),
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            search_box,
            text="Очистить",
            width=84,
            fg_color="#FEE2E2",
            hover_color="#FECACA",
            text_color="#DC2626",
            command=self.clear_map,
        ).pack(side="left", padx=4)

        # Вторая строка: быстрый выбор маркеров маршрута и информация о выбранной точке
        marker_bar = ctk.CTkFrame(self.tab_map, fg_color="#F3F4F6", corner_radius=8)
        marker_bar.pack(fill="x", padx=10, pady=(2, 4))

        ctk.CTkLabel(
            marker_bar,
            text="📍 Маркеры:",
            font=("Helvetica", 12, "bold"),
            text_color="#374151",
        ).pack(side="left", padx=(10, 6), pady=5)

        self.cb_map_markers = ctk.CTkComboBox(
            marker_bar,
            values=["Нет точек на карте"],
            width=290,
            command=self.focus_map_marker_by_label,
        )
        self.cb_map_markers.pack(side="left", padx=(0, 8), pady=5)
        self.cb_map_markers.set("Нет точек на карте")

        ctk.CTkButton(
            marker_bar,
            text="Весь маршрут",
            width=110,
            height=28,
            fg_color="#E5E7EB",
            hover_color="#D1D5DB",
            text_color="#1F2937",
            command=self.fit_map_to_route,
        ).pack(side="left", padx=(0, 10), pady=5)

        self.lbl_selected_marker = ctk.CTkLabel(
            marker_bar,
            text="Правый клик по карте — добавить точку; левый клик по маркеру — показать полное название и GPS.",
            font=("Helvetica", 12),
            text_color="#4B5563",
            anchor="w",
        )
        self.lbl_selected_marker.pack(side="left", fill="x", expand=True, padx=(4, 10), pady=5)

        self.map_widget = tkintermapview.TkinterMapView(
            self.tab_map, corner_radius=10
        )
        self.map_widget.pack(fill="both", expand=True, padx=10, pady=4)
        self._apply_map_provider(self.map_provider_name, preserve_view=False)
        init_lat, init_lon = self._wgs84_to_map(51.766, 55.1)
        self.map_widget.set_position(init_lat, init_lon)
        self.map_widget.set_zoom(9)
        self.map_widget.add_right_click_menu_command(
            label="Добавить адрес в маршрут",
            command=self.add_map_point,
            pass_coords=True,
        )

        bot = ctk.CTkFrame(self.tab_map, fg_color="#F9FAFB", height=90)
        bot.pack(fill="x", padx=10, pady=6)

        ctk.CTkLabel(
            bot, text="Маршрут следования:", font=("Helvetica", 13, "bold")
        ).pack(anchor="w", padx=10, pady=(6, 2))

        self.txt_route = ctk.CTkTextbox(bot, height=55)
        self.txt_route.pack(side="left", fill="x", expand=True, padx=10, pady=6)

        self.lbl_dist_show = ctk.CTkLabel(
            bot, text="0.0 км", font=("Helvetica", 16, "bold"), text_color="#007AFF"
        )
        self.lbl_dist_show.pack(side="right", padx=18)

    def _is_elliptical_map(self):
        return bool(MAP_PROVIDERS.get(self.map_provider_name, {}).get("elliptical", False))

    def _wgs84_to_map(self, lat, lon):
        lat_f = safe_float(lat, 51.766)
        lon_f = safe_float(lon, 55.1)
        if self._is_elliptical_map():
            return (wgs84_to_yandex_display_lat(lat_f), lon_f)
        return (lat_f, lon_f)

    def _map_to_wgs84(self, lat, lon):
        lat_f = safe_float(lat, 51.766)
        lon_f = safe_float(lon, 55.1)
        if self._is_elliptical_map():
            return (yandex_display_to_wgs84_lat(lat_f), lon_f)
        return (lat_f, lon_f)

    def _apply_map_provider(self, provider_name, preserve_view=True):
        if provider_name not in MAP_PROVIDERS:
            provider_name = "Яндекс Карты"
        cur_wgs = None
        cur_zoom = None
        if preserve_view and hasattr(self, "map_widget"):
            try:
                pos = self.map_widget.get_position()
                if pos and len(pos) == 2:
                    cur_wgs = self._map_to_wgs84(pos[0], pos[1])
                cur_zoom = getattr(self.map_widget, "zoom", None)
            except Exception:
                cur_wgs = None

        self.map_provider_name = provider_name
        cfg = MAP_PROVIDERS[provider_name]
        if hasattr(self, "map_widget"):
            try:
                self.map_widget.set_tile_server(cfg["url"], max_zoom=cfg["max_zoom"])
            except Exception:
                pass
            if cur_wgs is not None:
                try:
                    new_lat, new_lon = self._wgs84_to_map(cur_wgs[0], cur_wgs[1])
                    self.map_widget.set_position(new_lat, new_lon)
                    if cur_zoom is not None:
                        self.map_widget.set_zoom(int(round(cur_zoom)))
                except Exception:
                    pass

    def on_map_provider_change(self, selected_provider):
        self._apply_map_provider(selected_provider, preserve_view=True)
        self.redraw_map_from_segments(auto_fit=False)
        self.save_memory()

    def on_toggle_compact_labels(self):
        self.redraw_map_from_segments(auto_fit=False)
        self.save_memory()

    def _create_styled_marker(self, disp_lat, disp_lon, text, meta, is_selected=False):
        cb = lambda _marker=None, m=meta: self.on_marker_click(m)
        try:
            if is_selected:
                return self.map_widget.set_marker(
                    disp_lat,
                    disp_lon,
                    text=text,
                    command=cb,
                    marker_color_circle="#FFFFFF",
                    marker_color_outside="#007AFF",
                    text_color="#0047AB",
                    font=("Helvetica", 11, "bold"),
                )
            return self.map_widget.set_marker(
                disp_lat,
                disp_lon,
                text=text,
                command=cb,
                marker_color_circle="#9B261E",
                marker_color_outside="#C5542D",
                text_color="#4A1D12",
                font=("Helvetica", 10, "bold"),
            )
        except TypeError:
            return self.map_widget.set_marker(disp_lat, disp_lon, text=text, command=cb)

    def on_marker_click(self, meta):
        if not meta:
            return
        self.selected_marker_index = meta["index"]
        label_entry = f"#{meta['index']}. {meta['name']}"
        if hasattr(self, "cb_map_markers"):
            self.cb_map_markers.set(label_entry)
        if hasattr(self, "lbl_selected_marker"):
            self.lbl_selected_marker.configure(
                text=f"Выбрана точка #{meta['index']}: {meta['name']}  |  GPS (WGS84): {meta['lat']:.5f}, {meta['lon']:.5f}",
                text_color="#0056B3",
            )
        self.redraw_map_from_segments(auto_fit=False)
        if hasattr(self, "map_widget"):
            d_lat, d_lon = self._wgs84_to_map(meta["lat"], meta["lon"])
            self.map_widget.set_position(d_lat, d_lon)
            try:
                cur_z = getattr(self.map_widget, "zoom", 11)
                if cur_z < 13:
                    self.map_widget.set_zoom(13)
            except Exception:
                pass

    def focus_map_marker_by_label(self, selected_label):
        for meta in self.marker_metadata:
            if f"#{meta['index']}. {meta['name']}" == selected_label:
                self.on_marker_click(meta)
                return

    def fit_map_to_route(self):
        if not hasattr(self, "map_widget") or not self.marker_metadata:
            return
        disp_coords = [self._wgs84_to_map(m["lat"], m["lon"]) for m in self.marker_metadata]
        lats = [p[0] for p in disp_coords]
        lons = [p[1] for p in disp_coords]
        if len(disp_coords) == 1 or (max(lats) == min(lats) and max(lons) == min(lons)):
            self.map_widget.set_position(lats[0], lons[0])
            self.map_widget.set_zoom(13)
            return
        lat_span = max(max(lats) - min(lats), 0.01)
        lon_span = max(max(lons) - min(lons), 0.01)
        pad_lat = lat_span * 0.14
        pad_lon = lon_span * 0.14
        self.map_widget.fit_bounding_box(
            (max(lats) + pad_lat, min(lons) - pad_lon),
            (min(lats) - pad_lat, max(lons) + pad_lon),
        )

    def build_hist_tab(self):
        tree_frame = ctk.CTkFrame(self.tab_hist, fg_color="transparent")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=8)

        cols = (
            "ID",
            "№ ПЛ",
            "Дата",
            "Выезд",
            "Возврат",
            "Автомобиль",
            "Водитель",
            "Пробег",
            "Одометр",
            "Бензин",
            "Газ",
        )
        self.tr_hist = ttk.Treeview(
            tree_frame, columns=cols, show="headings", height=18
        )
        for col in cols:
            self.tr_hist.heading(col, text=col)
            self.tr_hist.column(col, width=100, anchor="center")
        self.tr_hist.column("ID", width=45)
        self.tr_hist.column("Автомобиль", width=150)
        self.tr_hist.column("Водитель", width=160)
        self.tr_hist.bind("<B1-Motion>", self.drag_select)

        scroll_y = ctk.CTkScrollbar(tree_frame, command=self.tr_hist.yview)
        self.tr_hist.configure(yscrollcommand=scroll_y.set)
        self.tr_hist.pack(side="left", fill="both", expand=True)
        scroll_y.pack(side="right", fill="y")

        print_bar = ctk.CTkFrame(self.tab_hist, fg_color="#F3F4F6", corner_radius=8)
        print_bar.pack(fill="x", padx=10, pady=(2, 4))

        ctk.CTkLabel(
            print_bar,
            text="🖨️ Формат автоматической печати Windows:",
            font=("Helvetica", 12, "bold"),
            text_color="#1F2937",
        ).pack(side="left", padx=(12, 8), pady=6)

        self.print_orient_seg = ctk.CTkSegmentedButton(
            print_bar,
            values=["Поперёк (A5 — 2 на лист A4)", "На весь лист (Книжная A4)"],
            variable=self.print_orientation_var,
            command=lambda _v: self.save_memory(),
        )
        self.print_orient_seg.pack(side="left", padx=4, pady=6)

        ctk.CTkLabel(
            print_bar,
            text="(Поперёк: 1 шт. — на 1-й половине A4, 2 шт. — на одном A4; 1-я и 2-я стр. точно совпадают)",
            font=("Helvetica", 11),
            text_color="#4B5563",
        ).pack(side="left", padx=10, pady=6)

        bot_panel = ctk.CTkFrame(self.tab_hist, fg_color="transparent")
        bot_panel.pack(fill="x", padx=10, pady=6)

        ctk.CTkButton(
            bot_panel,
            text="Редактировать",
            command=self.load_history_for_edit,
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            bot_panel,
            text="Сформировать ПЛ",
            fg_color="#10B981",
            command=self.generate_excel_from_hist,
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            bot_panel,
            text="Открыть Excel",
            fg_color="#0EA5E9",
            hover_color="#0284C7",
            command=self.open_selected_excel,
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            bot_panel,
            text="🔄 Пересчитать цепочку",
            fg_color="#F59E0B",
            hover_color="#D97706",
            text_color="#FFFFFF",
            command=self.recalculate_selected_or_all_chains,
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            bot_panel,
            text="Одиночная печать",
            fg_color="#8B5CF6",
            hover_color="#7C3AED",
            command=self.single_print_selected,
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            bot_panel,
            text="Массовая печать",
            fg_color="#6366F1",
            command=self.mass_print_selected,
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            bot_panel,
            text="Удалить запись",
            fg_color="#EF4444",
            command=self.delete_history,
        ).pack(side="right", padx=4)

        self.refresh_history()

    def drag_select(self, event):
        row_id = self.tr_hist.identify_row(event.y)
        if row_id:
            self.tr_hist.selection_add(row_id)

    def build_crud_tab(self):
        self.c_tabs = ctk.CTkTabview(self.tab_crud)
        self.c_tabs.pack(fill="both", expand=True, padx=10, pady=6)

        t_drv = self.c_tabs.add("Водители")
        t_veh = self.c_tabs.add("Автомобили")
        t_org = self.c_tabs.add("Организации")
        t_stf = self.c_tabs.add("Персонал")

        # Водители
        f_d = ctk.CTkFrame(t_drv, fg_color="transparent")
        f_d.pack(fill="x", padx=10, pady=6)
        self.en_d_fio = ctk.CTkEntry(f_d, placeholder_text="ФИО водителя", width=200)
        self.en_d_fio.pack(side="left", padx=4)
        self.en_d_cat = ctk.CTkEntry(f_d, placeholder_text="Категории (B, C)", width=110)
        self.en_d_cat.pack(side="left", padx=4)
        self.en_d_exp = ctk.CTkEntry(f_d, placeholder_text="Годен до (ДД.ММ.ГГГГ)", width=140)
        self.en_d_exp.pack(side="left", padx=4)
        self.en_d_tab = ctk.CTkEntry(f_d, placeholder_text="Таб. №", width=90)
        self.en_d_tab.pack(side="left", padx=4)
        self.en_d_ser = ctk.CTkEntry(f_d, placeholder_text="Серия ВУ", width=90)
        self.en_d_ser.pack(side="left", padx=4)
        self.en_d_num = ctk.CTkEntry(f_d, placeholder_text="Номер ВУ", width=110)
        self.en_d_num.pack(side="left", padx=4)
        self.btn_d_add = ctk.CTkButton(f_d, text="Добавить", width=110, command=self.add_drv)
        self.btn_d_add.pack(side="left", padx=6)

        self.tr_d = ttk.Treeview(
            t_drv,
            columns=("ID", "ФИО", "Кат.", "Годен", "Таб. №", "Серия", "Номер"),
            show="headings",
        )
        for col in ("ID", "ФИО", "Кат.", "Годен", "Таб. №", "Серия", "Номер"):
            self.tr_d.heading(col, text=col)
        self.tr_d.pack(fill="both", expand=True, padx=10, pady=6)
        self.tr_d.bind("<Double-1>", lambda e: self.load_crud_edit("drivers", self.tr_d))
        f_d_btns = ctk.CTkFrame(t_drv, fg_color="transparent")
        f_d_btns.pack(pady=6)
        ctk.CTkButton(
            f_d_btns,
            text="Редактировать",
            fg_color="#2563EB",
            command=lambda: self.load_crud_edit("drivers", self.tr_d),
        ).pack(side="left", padx=6)
        ctk.CTkButton(
            f_d_btns,
            text="Удалить выбранного водителя",
            fg_color="#EF4444",
            command=lambda: self.del_item("drivers", self.tr_d),
        ).pack(side="left", padx=6)

        # Автомобили
        f_v = ctk.CTkFrame(t_veh, fg_color="transparent")
        f_v.pack(fill="x", padx=10, pady=6)
        self.en_v_mark = ctk.CTkEntry(f_v, placeholder_text="Марка авто", width=160)
        self.en_v_mark.pack(side="left", padx=4)
        self.en_v_reg = ctk.CTkEntry(f_v, placeholder_text="Госномер", width=110)
        self.en_v_reg.pack(side="left", padx=4)
        self.en_v_gar = ctk.CTkEntry(f_v, placeholder_text="Гаражный №", width=100)
        self.en_v_gar.pack(side="left", padx=4)
        self.en_v_f = ctk.CTkEntry(f_v, placeholder_text="Бензин (АИ-92)", width=110)
        self.en_v_f.pack(side="left", padx=4)
        self.en_v_fn = ctk.CTkEntry(f_v, placeholder_text="Норма бенз.", width=100)
        self.en_v_fn.pack(side="left", padx=4)
        self.en_v_g = ctk.CTkEntry(f_v, placeholder_text="Газ (ПБА)", width=100)
        self.en_v_g.pack(side="left", padx=4)
        self.en_v_gn = ctk.CTkEntry(f_v, placeholder_text="Норма газа", width=100)
        self.en_v_gn.pack(side="left", padx=4)
        self.btn_v_add = ctk.CTkButton(f_v, text="Добавить", width=110, command=self.add_veh)
        self.btn_v_add.pack(side="left", padx=6)

        self.tr_v = ttk.Treeview(
            t_veh,
            columns=(
                "ID",
                "Марка",
                "Госномер",
                "Гараж",
                "Бензин (марка)",
                "Норма бенз.",
                "Газ (марка)",
                "Норма газа",
            ),
            show="headings",
        )
        for col in (
            "ID",
            "Марка",
            "Госномер",
            "Гараж",
            "Бензин (марка)",
            "Норма бенз.",
            "Газ (марка)",
            "Норма газа",
        ):
            self.tr_v.heading(col, text=col)
        self.tr_v.pack(fill="both", expand=True, padx=10, pady=6)
        self.tr_v.bind("<Double-1>", lambda e: self.load_crud_edit("vehicles", self.tr_v))
        f_v_btns = ctk.CTkFrame(t_veh, fg_color="transparent")
        f_v_btns.pack(pady=6)
        ctk.CTkButton(
            f_v_btns,
            text="Редактировать",
            fg_color="#2563EB",
            command=lambda: self.load_crud_edit("vehicles", self.tr_v),
        ).pack(side="left", padx=6)
        ctk.CTkButton(
            f_v_btns,
            text="Удалить выбранный автомобиль",
            fg_color="#EF4444",
            command=lambda: self.del_item("vehicles", self.tr_v),
        ).pack(side="left", padx=6)

        # Организации и Заказчики
        f_o1 = ctk.CTkFrame(t_org, fg_color="transparent")
        f_o1.pack(side="left", fill="both", expand=True, padx=8, pady=6)
        f_o1_top = ctk.CTkFrame(f_o1, fg_color="transparent")
        f_o1_top.pack(fill="x", pady=4)
        self.en_org = ctk.CTkEntry(f_o1_top, placeholder_text="Наименование организации", width=280)
        self.en_org.pack(side="left", padx=4)
        self.btn_o_add = ctk.CTkButton(
            f_o1_top,
            text="Добавить",
            width=100,
            command=lambda: self.add_simple("orgs", self.en_org),
        )
        self.btn_o_add.pack(side="left", padx=4)
        self.tr_org = ttk.Treeview(f_o1, columns=("ID", "Организации"), show="headings")
        self.tr_org.heading("ID", text="ID")
        self.tr_org.heading("Организации", text="Организации")
        self.tr_org.column("ID", width=55, anchor="center")
        self.tr_org.pack(fill="both", expand=True, pady=4)
        self.tr_org.bind("<Double-1>", lambda e: self.load_crud_edit("orgs", self.tr_org))
        f_o1_btns = ctk.CTkFrame(f_o1, fg_color="transparent")
        f_o1_btns.pack(pady=4)
        ctk.CTkButton(
            f_o1_btns,
            text="Редактировать",
            fg_color="#2563EB",
            width=120,
            command=lambda: self.load_crud_edit("orgs", self.tr_org),
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            f_o1_btns,
            text="Удалить",
            fg_color="#EF4444",
            width=100,
            command=lambda: self.del_item("orgs", self.tr_org),
        ).pack(side="left", padx=4)

        f_o2 = ctk.CTkFrame(t_org, fg_color="transparent")
        f_o2.pack(side="right", fill="both", expand=True, padx=8, pady=6)
        f_o2_top = ctk.CTkFrame(f_o2, fg_color="transparent")
        f_o2_top.pack(fill="x", pady=4)
        self.en_cli = ctk.CTkEntry(f_o2_top, placeholder_text="Заказчик", width=280)
        self.en_cli.pack(side="left", padx=4)
        self.btn_c_add = ctk.CTkButton(
            f_o2_top,
            text="Добавить",
            width=100,
            command=lambda: self.add_simple("clients", self.en_cli),
        )
        self.btn_c_add.pack(side="left", padx=4)
        self.tr_cli = ttk.Treeview(f_o2, columns=("ID", "Заказчики"), show="headings")
        self.tr_cli.heading("ID", text="ID")
        self.tr_cli.heading("Заказчики", text="Заказчики")
        self.tr_cli.column("ID", width=55, anchor="center")
        self.tr_cli.pack(fill="both", expand=True, pady=4)
        self.tr_cli.bind("<Double-1>", lambda e: self.load_crud_edit("clients", self.tr_cli))
        f_o2_btns = ctk.CTkFrame(f_o2, fg_color="transparent")
        f_o2_btns.pack(pady=4)
        ctk.CTkButton(
            f_o2_btns,
            text="Редактировать",
            fg_color="#2563EB",
            width=120,
            command=lambda: self.load_crud_edit("clients", self.tr_cli),
        ).pack(side="left", padx=4)
        ctk.CTkButton(
            f_o2_btns,
            text="Удалить",
            fg_color="#EF4444",
            width=100,
            command=lambda: self.del_item("clients", self.tr_cli),
        ).pack(side="left", padx=4)

        # Персонал (поддержка совмещения Механик / Диспетчер + редактирование)
        f_s = ctk.CTkFrame(t_stf, fg_color="transparent")
        f_s.pack(fill="x", padx=10, pady=6)
        self.en_s_fio = ctk.CTkEntry(f_s, placeholder_text="ФИО сотрудника", width=260)
        self.en_s_fio.pack(side="left", padx=4)
        self.cb_s_role = ctk.CTkComboBox(
            f_s, values=["Механик", "Диспетчер", "Механик / Диспетчер"], width=210
        )
        self.cb_s_role.pack(side="left", padx=4)
        self.btn_s_add = ctk.CTkButton(
            f_s, text="Добавить", width=110, command=self.add_staff
        )
        self.btn_s_add.pack(side="left", padx=6)
        ctk.CTkButton(
            f_s,
            text="Сброс",
            width=80,
            fg_color="#E5E7EB",
            hover_color="#D1D5DB",
            text_color="#374151",
            command=self.reset_crud_buttons,
        ).pack(side="left", padx=4)

        self.tr_stf = ttk.Treeview(
            t_stf, columns=("ID", "ФИО", "Должность"), show="headings"
        )
        for col in ("ID", "ФИО", "Должность"):
            self.tr_stf.heading(col, text=col)
        self.tr_stf.column("ID", width=60, anchor="center")
        self.tr_stf.pack(fill="both", expand=True, padx=10, pady=6)
        self.tr_stf.bind("<Double-1>", lambda e: self.load_crud_edit("staff", self.tr_stf))
        f_stf_btns = ctk.CTkFrame(t_stf, fg_color="transparent")
        f_stf_btns.pack(pady=6)
        ctk.CTkButton(
            f_stf_btns,
            text="Редактировать сотрудника",
            fg_color="#2563EB",
            command=lambda: self.load_crud_edit("staff", self.tr_stf),
        ).pack(side="left", padx=6)
        ctk.CTkButton(
            f_stf_btns,
            text="Удалить сотрудника",
            fg_color="#EF4444",
            command=lambda: self.del_item("staff", self.tr_stf),
        ).pack(side="left", padx=6)

        # Виды работ и нормы ДТ (для Трактора МТЗ-82 и Экскаватора) + редактирование
        t_wt = self.c_tabs.add("Виды работ (ДТ)")
        f_wt = ctk.CTkFrame(t_wt, fg_color="transparent")
        f_wt.pack(fill="x", padx=10, pady=6)
        self.en_wt_name = ctk.CTkEntry(
            f_wt, placeholder_text="Наименование вида работ (МТЗ-82 / ЭО-2621)", width=340
        )
        self.en_wt_name.pack(side="left", padx=4)
        self.en_wt_norm = ctk.CTkEntry(
            f_wt, placeholder_text="Норма ДТ (л/м·ч)", width=130
        )
        self.en_wt_norm.pack(side="left", padx=4)
        self.en_wt_unit = ctk.CTkEntry(
            f_wt, placeholder_text="Ед. изм. (м³, га, км)", width=130
        )
        self.en_wt_unit.pack(side="left", padx=4)
        self.btn_wt_add = ctk.CTkButton(
            f_wt, text="Добавить вид работ", width=160, command=self.add_wt
        )
        self.btn_wt_add.pack(side="left", padx=6)
        ctk.CTkButton(
            f_wt,
            text="Сброс",
            width=80,
            fg_color="#E5E7EB",
            hover_color="#D1D5DB",
            text_color="#374151",
            command=self.reset_crud_buttons,
        ).pack(side="left", padx=4)

        self.tr_wt = ttk.Treeview(
            t_wt,
            columns=("ID", "Наименование вида работ", "Норма ДТ (л/м·ч)", "Ед. изм."),
            show="headings",
        )
        for col in ("ID", "Наименование вида работ", "Норма ДТ (л/м·ч)", "Ед. изм."):
            self.tr_wt.heading(col, text=col)
        self.tr_wt.column("ID", width=60, anchor="center")
        self.tr_wt.pack(fill="both", expand=True, padx=10, pady=6)
        self.tr_wt.bind("<Double-1>", lambda e: self.load_crud_edit("work_types", self.tr_wt))
        f_wt_btns = ctk.CTkFrame(t_wt, fg_color="transparent")
        f_wt_btns.pack(pady=6)
        ctk.CTkButton(
            f_wt_btns,
            text="Редактировать вид работ",
            fg_color="#2563EB",
            command=lambda: self.load_crud_edit("work_types", self.tr_wt),
        ).pack(side="left", padx=6)
        ctk.CTkButton(
            f_wt_btns,
            text="⚙️ Применить норму к путевым листам",
            fg_color="#6366F1",
            hover_color="#4F46E5",
            command=self._apply_selected_wt_norm,
        ).pack(side="left", padx=6)
        ctk.CTkButton(
            f_wt_btns,
            text="Удалить выбранный вид работ",
            fg_color="#EF4444",
            command=lambda: self.del_wt_item(),
        ).pack(side="left", padx=6)

        self.refresh_all_tables()

    def show_context_menu(self, event, combobox):
        current_val = combobox.get().strip()
        if not current_val:
            return
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(
            label="Переименовать",
            command=lambda: self.rename_list_item(combobox, current_val),
        )
        menu.add_command(
            label="Удалить из списка",
            command=lambda: self.delete_list_item(combobox, current_val),
        )
        menu.tk_popup(event.x_root, event.y_root)

    def rename_list_item(self, combobox, old_name):
        dialog = CustomInputDialog(
            self,
            title="Переименование",
            prompt="Введите новое название:",
            initial_value=old_name,
        )
        self.wait_window(dialog)
        new_name = dialog.result
        if not new_name or new_name == old_name:
            return
        self.db.query("UPDATE orgs SET name=? WHERE name=?", (new_name, old_name))
        self.db.query("UPDATE clients SET name=? WHERE name=?", (new_name, old_name))
        self.db.query("UPDATE locs SET name=? WHERE name=?", (new_name, old_name))
        self.refresh_all_tables()
        combobox.set(new_name)
        self.log_action(f"Переименование пункта '{old_name}' в '{new_name}'")

    def delete_list_item(self, combobox, name_to_delete):
        if not messagebox.askyesno(
            "Удаление",
            f"Вы действительно хотите удалить '{name_to_delete}' из списков?",
        ):
            return
        self.db.query("DELETE FROM orgs WHERE name=?", (name_to_delete,))
        self.db.query("DELETE FROM clients WHERE name=?", (name_to_delete,))
        self.db.query("DELETE FROM locs WHERE name=?", (name_to_delete,))
        self.refresh_all_tables()
        combobox.set("")
        self.log_action(f"Удаление пункта '{name_to_delete}' из списков")

    def open_calendar(self):
        top = ctk.CTkToplevel(self)
        top.title("Выберите дату")
        top.geometry("300x250")
        top.transient(self)
        top.grab_set()
        cal = Calendar(top, selectmode="day", date_pattern="dd.mm.yyyy", locale="ru_RU")
        cal.pack(fill="both", expand=True, padx=10, pady=10)

        def set_date():
            self.en_date.delete(0, "end")
            self.en_date.insert(0, cal.get_date())
            top.destroy()

        ctk.CTkButton(top, text="Выбрать", command=set_date).pack(pady=6)

    def calc_time(self, *args):
        try:
            h1, m1 = map(int, self.en_t_out.get().strip().split(":"))
            h2, m2 = map(int, self.en_t_in.get().strip().split(":"))
            tot = round(((h2 * 60 + m2) - (h1 * 60 + m1)) / 60.0, 1)
            if tot < 0:
                tot += 24.0
            self.lbl_t_tot.configure(text=f"Итого: {tot} ч")
            self.update_route_displays()
            return str(tot)
        except Exception:
            self.lbl_t_tot.configure(text="Итого: 8.0 ч")
            return "8.0"

    def _schedule_recalc(self, *args):
        """Отложенный (debounce 80 мс) пересчёт расхода и спидометра без подвисаний при наборе текста."""
        if getattr(self, "_recalc_after_id", None):
            try:
                self.after_cancel(self._recalc_after_id)
            except Exception:
                pass
        self._recalc_after_id = self.after(80, self.recalc_live)

    def _refresh_tractor_works_tree(self):
        if not hasattr(self, "tr_works"):
            return
        total_extra_coef = (
            safe_float(self.var_winter.get(), 0.0)
            + safe_float(self.var_city.get(), 0.0)
            + safe_float(self.var_dirt.get(), 0.0)
            + safe_float(self.var_trailer.get(), 0.0)
        )
        mult = 1.0 + total_extra_coef / 100.0
        rows_to_render = []
        for w in self.tractor_works:
            mh = safe_float(w.get("mh"), 0.0)
            nm = safe_float(w.get("norm_mh"), 5.2)
            dt_l = round(mh * nm * mult, 2)
            w["dt_l"] = dt_l
            rows_to_render.append((w.get("name", ""), f"{mh:g}", f"{nm:g}", f"{dt_l:.2f}"))

        sig = tuple(rows_to_render)
        if sig == getattr(self, "_last_works_render_sig", None):
            return
        self._last_works_render_sig = sig

        for item in self.tr_works.get_children():
            self.tr_works.delete(item)
        for idx, vals in enumerate(rows_to_render):
            self.tr_works.insert("", "end", iid=str(idx), values=vals)

    def add_tractor_work(self):
        if self.active_mode != "tractors":
            return
        wt_name = self.cb_work_type.get().strip() if hasattr(self, "cb_work_type") else ""
        if not wt_name:
            messagebox.showwarning("Внимание", "Выберите вид работ из справочника.")
            return
        mh = safe_float(self.en_work_mh.get(), 0.0) if hasattr(self, "en_work_mh") else 0.0
        if mh <= 0:
            messagebox.showwarning("Внимание", "Укажите отработанные моточасы (> 0) для добавляемой работы.")
            return
        wt_rows = self.db_tractors.get_list("work_types")
        wt_info = next((r for r in wt_rows if r[1] == wt_name), None)
        norm_mh = safe_float(wt_info[2], 0.0) if wt_info else safe_float(self.en_norm_mh.get(), 5.2)
        if norm_mh <= 0:
            norm_mh = safe_float(self.en_norm_mh.get(), 5.2)
        self.tractor_works.append(
            {
                "name": wt_name,
                "mh": round(mh, 2),
                "norm_mh": round(norm_mh, 2),
            }
        )
        self._refresh_tractor_works_tree()
        self.recalc_live()

    def remove_tractor_work(self):
        if not hasattr(self, "tr_works"):
            return
        sel = self.tr_works.selection()
        if not sel:
            return
        idxs = sorted([int(i) for i in sel if str(i).isdigit()], reverse=True)
        for idx in idxs:
            if 0 <= idx < len(self.tractor_works):
                self.tractor_works.pop(idx)
        self._refresh_tractor_works_tree()
        self.recalc_live()

    def clear_tractor_works(self):
        self.tractor_works.clear()
        self._refresh_tractor_works_tree()
        self.recalc_live()

    def on_work_type_change(self, selected_wt):
        wt_rows = self.db_tractors.get_list("work_types")
        wt_info = next((r for r in wt_rows if r[1] == selected_wt), None)
        if wt_info and hasattr(self, "en_norm_mh"):
            self.en_norm_mh.delete(0, "end")
            self.en_norm_mh.insert(0, str(wt_info[2] or 5.2))
        self.recalc_live()

    def load_vehicle_data(self, *args):
        mark = self.cb_veh.get().strip() if hasattr(self, "cb_veh") else ""
        v_list = self.db.get_list("vehicles")
        v_info = next((r for r in v_list if r[1] == mark), None)
        if not v_info and v_list:
            v_info = v_list[0]
            mark = str(v_info[1])
            if hasattr(self, "cb_veh"):
                self.cb_veh.set(mark)

        if self.active_mode == "tractors":
            if v_info:
                tr_norm_km = safe_float(v_info[6], 16.5)
                self._tractor_norm_km = tr_norm_km
                self.en_f_norm.delete(0, "end")
                self.en_f_norm.insert(0, str(tr_norm_km))
                if hasattr(self, "en_norm_mh"):
                    # Если выбран вид работ из справочника — берём норму из справочника видов работ, иначе из техники
                    wt_sel = self.cb_work_type.get().strip() if hasattr(self, "cb_work_type") else ""
                    wt_rows = self.db_tractors.get_list("work_types")
                    wt_info = next((r for r in wt_rows if r[1] == wt_sel), None)
                    mh_norm_val = (wt_info[2] if wt_info and wt_info[2] else None) or v_info[7] or 5.2
                    self.en_norm_mh.delete(0, "end")
                    self.en_norm_mh.insert(0, str(mh_norm_val))
            last = self.db.get_last_veh(mark)
            if last:
                odo_e, moto_e, f_e = last
                self.en_odo_s.delete(0, "end")
                self.en_odo_s.insert(0, str(odo_e or 0.0))
                if hasattr(self, "en_moto_s"):
                    self.en_moto_s.delete(0, "end")
                    self.en_moto_s.insert(0, str(moto_e or 0.0))
                self.en_f_s.delete(0, "end")
                self.en_f_s.insert(0, str(f_e or 0.0))
            self.var_fuel_type.set("ДТ (Дизтопливо)")
        else:
            if v_info:
                p_norm = safe_float(v_info[5], 0.0)
                g_norm = safe_float(v_info[7], 0.0)
                f_brand_str = str(v_info[4] or "").upper()
                self.en_f_norm.delete(0, "end")
                self.en_f_norm.insert(0, str(p_norm))
                self.en_g_norm.delete(0, "end")
                self.en_g_norm.insert(0, str(g_norm))
                # Автовыбор типа топлива из справочника автомобиля
                if "ДТ" in f_brand_str or "ДИЗ" in f_brand_str:
                    self.var_fuel_type.set("ДТ (Дизтопливо)")
                elif p_norm > 0 and g_norm > 0:
                    if self.var_fuel_type.get() not in ("Бензин", "Газ", "Бензин + Газ"):
                        self.var_fuel_type.set("Бензин + Газ")
                elif g_norm > 0 and p_norm <= 0:
                    self.var_fuel_type.set("Газ")
                else:
                    self.var_fuel_type.set("Бензин")

            last = self.db.get_last_veh(mark)
            if last:
                odo_e, f_e, g_e = last
                self.en_odo_s.delete(0, "end")
                self.en_odo_s.insert(0, str(odo_e or 0.0))
                self.en_f_s.delete(0, "end")
                self.en_f_s.insert(0, str(f_e or 0.0))
                self.en_g_s.delete(0, "end")
                self.en_g_s.insert(0, str(g_e or 0.0))
        self.recalc_live()

    def recalc_live(self, *args):
        self._recalc_after_id = None
        try:
            dist = safe_float(self.en_dist.get(), 0.0)
            odo_s = safe_float(self.en_odo_s.get(), 0.0)
            only_mh = (
                self.active_mode == "tractors"
                and hasattr(self, "var_only_motohours")
                and bool(self.var_only_motohours.get())
            )
            eff_dist = 0.0 if only_mh else dist
            odo_e = round(odo_s + eff_dist, 1)
            self.lbl_odo_e.configure(
                text=f"КОНЕЦ: {odo_e} км (только м/ч)" if only_mh else f"КОНЕЦ: {odo_e} км"
            )
            if hasattr(self, "lbl_dist_main"):
                self.lbl_dist_main.configure(
                    text=f"{dist:.1f} км (не уч.)" if only_mh else f"{dist:.1f} км"
                )

            total_extra_coef = (
                safe_float(self.var_winter.get(), 0.0)
                + safe_float(self.var_city.get(), 0.0)
                + safe_float(self.var_dirt.get(), 0.0)
                + safe_float(self.var_trailer.get(), 0.0)
            )
            norm_p = safe_float(self.en_f_norm.get(), 0.0)

            if self.active_mode == "tractors":
                norm_p = self._get_tractor_norm_km()
                moto_s = safe_float(self.en_moto_s.get(), 0.0) if hasattr(self, "en_moto_s") else 0.0
                if self.tractor_works:
                    dt_km, dt_mh, dt_tot, total_mh, per_work_dts = calculate_tractor_multi_works_diesel(
                        norm_p, eff_dist, self.tractor_works, total_extra_coef
                    )
                    for w_item, w_dt in zip(self.tractor_works, per_work_dts):
                        if isinstance(w_item, dict):
                            w_item["dt_l"] = w_dt
                    moto_h = total_mh
                    if hasattr(self, "en_moto_h"):
                        cur_mh_txt = self.en_moto_h.get().strip()
                        new_mh_txt = f"{moto_h:g}"
                        if cur_mh_txt != new_mh_txt:
                            self.en_moto_h.delete(0, "end")
                            self.en_moto_h.insert(0, new_mh_txt)
                    self._refresh_tractor_works_tree()
                    works_cnt = len(self.tractor_works)
                    if only_mh:
                        summary_txt = (
                            f"ДТ (работ: {works_cnt}, {moto_h:g} м/ч, без км): "
                            f"Работы {dt_mh} л = ИТОГО {dt_tot} л"
                        )
                    else:
                        summary_txt = (
                            f"ДТ (работ: {works_cnt}, {moto_h:g} м/ч): "
                            f"Перегон {dt_km} л ({dist:g} км × {norm_p:g}) + Работы {dt_mh} л = ИТОГО {dt_tot} л"
                        )
                else:
                    moto_h = safe_float(self.en_moto_h.get(), 0.0) if hasattr(self, "en_moto_h") else 0.0
                    norm_mh = safe_float(self.en_norm_mh.get(), 5.2) if hasattr(self, "en_norm_mh") else 5.2
                    dt_km, dt_mh, dt_tot = calculate_tractor_diesel_consumption(
                        norm_p, eff_dist, norm_mh, moto_h, total_extra_coef
                    )
                    if only_mh:
                        summary_txt = (
                            f"Расход ДТ (только моточасы, без км): Моточасы {dt_mh} л = ИТОГО {dt_tot} л"
                        )
                    else:
                        summary_txt = (
                            f"Расход ДТ: Перегон {dt_km} л ({dist:g} км × {norm_p:g}) + Моточасы {dt_mh} л = ИТОГО {dt_tot} л"
                        )

                moto_e = round(moto_s + moto_h, 1)
                if hasattr(self, "lbl_moto_e"):
                    self.lbl_moto_e.configure(text=f"КОН: {moto_e} м/ч")
                self.last_dt_km = dt_km
                self.last_dt_mh = dt_mh
                self.last_dt_total = dt_tot
                if hasattr(self, "lbl_dt_summary"):
                    self.lbl_dt_summary.configure(text=summary_txt)

                f_s = safe_float(self.en_f_s.get(), 0.0)
                f_a = safe_float(self.en_f_a.get(), 0.0)
                self.last_spent_p = dt_tot
                f_act_str = self.en_f_act.get().strip()
                act_p = safe_float(f_act_str, dt_tot) if f_act_str else dt_tot
                f_e = round(f_s + f_a - act_p, 2)
                self.en_f_e.delete(0, "end")
                self.en_f_e.insert(0, f"{f_e:.2f}")
                return

            active_fuel = self.var_fuel_type.get()
            coef_dec = total_extra_coef / 100.0

            # Бензин / ДТ
            f_s = safe_float(self.en_f_s.get(), 0.0)
            f_a = safe_float(self.en_f_a.get(), 0.0)
            spent_p = (
                calculate_fuel_consumption(dist, norm_p, coef_dec)
                if ("Бензин" in active_fuel or "ДТ" in active_fuel)
                else 0.0
            )
            self.last_spent_p = spent_p
            f_act_str = self.en_f_act.get().strip()
            act_p = safe_float(f_act_str, spent_p) if f_act_str else spent_p
            f_e = round(f_s + f_a - act_p, 2)
            self.en_f_e.delete(0, "end")
            self.en_f_e.insert(0, f"{f_e:.2f}")

            # Газ
            g_s = safe_float(self.en_g_s.get(), 0.0)
            g_a = safe_float(self.en_g_a.get(), 0.0)
            norm_g = safe_float(self.en_g_norm.get(), 0.0)
            spent_g = (
                calculate_fuel_consumption(dist, norm_g, coef_dec)
                if "Газ" in active_fuel
                else 0.0
            )
            self.last_spent_g = spent_g
            g_act_str = self.en_g_act.get().strip()
            act_g = safe_float(g_act_str, spent_g) if g_act_str else spent_g
            g_e = round(g_s + g_a - act_g, 2)
            self.en_g_e.delete(0, "end")
            self.en_g_e.insert(0, f"{g_e:.2f}")
        except Exception:
            pass

    def geocode_address(self, address_name):
        """
        Автоматически находит координаты адреса и сохраняет точку в таблицу locs (waybills.db),
        чтобы все введённые точки запоминались.
        """
        clean_name = str(address_name).strip()
        if not clean_name:
            return None
        locs = self.db.get_list("locs")
        existing = next(
            (r for r in locs if str(r[1]).strip().lower() == clean_name.lower()),
            None,
        )
        if existing:
            return (float(existing[2]), float(existing[3]))
        try:
            url = f"https://nominatim.openstreetmap.org/search?q={clean_name}&format=json&limit=1&accept-language=ru"
            resp = requests.get(
                url, headers={"User-Agent": "WaybillApp/1.0"}, timeout=5
            )
            if resp.status_code == 200:
                res = resp.json()
                if res:
                    lat = float(res[0]["lat"])
                    lon = float(res[0]["lon"])
                    self.db.query(
                        "INSERT OR IGNORE INTO locs (name, lat, lon) VALUES (?, ?, ?)",
                        (clean_name, lat, lon),
                    )
                    self.refresh_all_tables()
                    return (lat, lon)
        except Exception:
            pass
        return None

    def _route_cache_key(self, lat1, lon1, lat2, lon2):
        return f"{round(float(lat1), 5)},{round(float(lon1), 5)}->{round(float(lat2), 5)},{round(float(lon2), 5)}"

    def get_road_distance(self, lat1, lon1, lat2, lon2, allow_network=True):
        key = self._route_cache_key(lat1, lon1, lat2, lon2)
        if key in self._route_mem_cache:
            return self._route_mem_cache[key]

        try:
            cached = self.db.get_cached_route(key)
            if cached:
                dist_c, coords_json = cached
                coords_c = json.loads(coords_json) if coords_json else None
                if coords_c:
                    coords_c = [(float(c[0]), float(c[1])) for c in coords_c]
                    res = (round(float(dist_c), 1), coords_c)
                    self._route_mem_cache[key] = res
                    return res
        except Exception:
            pass

        if not allow_network:
            return (None, None)

        dist = None
        path_coords = None
        try:
            osrm_url = (
                f"http://router.project-osrm.org/route/v1/driving/"
                f"{lon1},{lat1};{lon2},{lat2}?overview=full&geometries=geojson"
            )
            response = requests.get(osrm_url, timeout=2.5)
            if response.status_code == 200:
                data = response.json()
                if data.get("code") == "Ok" and data.get("routes"):
                    dist = round(data["routes"][0]["distance"] / 1000.0, 1)
                    coords = data["routes"][0]["geometry"]["coordinates"]
                    path_coords = [(float(lat), float(lon)) for lon, lat in coords]
        except Exception:
            pass

        if dist is None:
            dist = calculate_distance(lat1, lon1, lat2, lon2)
        if not path_coords:
            path_coords = [(float(lat1), float(lon1)), (float(lat2), float(lon2))]

        res = (dist, path_coords)
        self._route_mem_cache[key] = res
        try:
            self.db.save_cached_route(key, dist, json.dumps(path_coords))
        except Exception:
            pass
        return res

    def _fetch_missing_routes_bg(self, pending_items, gen):
        updated = False
        for seg, p1, p2 in pending_items:
            if gen != self._map_fetch_gen:
                return
            _, coords = self.get_road_distance(p1[0], p1[1], p2[0], p2[1], allow_network=True)
            if coords and len(coords) >= 2:
                seg["coords"] = coords
                updated = True
        if updated and gen == self._map_fetch_gen:
            try:
                self.after(0, lambda: self.redraw_map_from_segments(auto_fit=False))
            except Exception:
                pass

    def redraw_map_from_segments(self, auto_fit=True, force=False):
        if not force and hasattr(self, "tabview"):
            try:
                if self.tabview.get() != "Маршрут на карте":
                    self._map_dirty = True
                    return
            except Exception:
                pass
        self._map_dirty = False

        for m in self.route_markers:
            try:
                m.delete()
            except Exception:
                pass
        self.route_markers.clear()

        for p in self.route_paths:
            try:
                p.delete()
            except Exception:
                pass
        self.route_paths.clear()
        self.marker_metadata.clear()

        if not hasattr(self, "map_widget"):
            return

        locs = self.db.get_list("locs")
        loc_map = {r[1]: r for r in locs}
        ordered_points = []
        seen_coords = set()
        pending_route_fetches = []

        def _register_point(name, lat, lon):
            lat_f, lon_f = safe_float(lat, 0.0), safe_float(lon, 0.0)
            coord_key = (round(lat_f, 5), round(lon_f, 5))
            if coord_key not in seen_coords:
                seen_coords.add(coord_key)
                ordered_points.append({"name": str(name), "lat": lat_f, "lon": lon_f})

        if not self.route_segments:
            # Если поставлена первая стартовая точка (например, Буртинский), отображаем её маркер на карте
            pickup_name = self.cb_pickup.get().strip() if hasattr(self, "cb_pickup") else ""
            sp = loc_map.get(pickup_name) if pickup_name else None
            if sp:
                _register_point(pickup_name, sp[2], sp[3])
            elif locs and auto_fit:
                first_lat, first_lon = self._wgs84_to_map(locs[0][2], locs[0][3])
                self.map_widget.set_position(first_lat, first_lon)
        else:
            for seg in self.route_segments:
                sp = loc_map.get(seg.get("from"))
                tp = loc_map.get(seg.get("to"))
                if not sp or not tp:
                    continue
                p1 = (safe_float(sp[2], 0.0), safe_float(sp[3], 0.0))
                p2 = (safe_float(tp[2], 0.0), safe_float(tp[3], 0.0))
                _register_point(seg.get("from", ""), p1[0], p1[1])
                _register_point(seg.get("to", ""), p2[0], p2[1])

                coords = seg.get("coords")
                if not coords:
                    # Сначала мгновенно проверяем кэш (память/SQLite) без блокировки интерфейса
                    _, cached_coords = self.get_road_distance(
                        p1[0], p1[1], p2[0], p2[1], allow_network=False
                    )
                    if cached_coords:
                        coords = cached_coords
                        seg["coords"] = coords
                    else:
                        coords = [p1, p2]
                        pending_route_fetches.append((seg, p1, p2))

                disp_path = [self._wgs84_to_map(c[0], c[1]) for c in coords]
                if len(disp_path) >= 2:
                    path_obj = self.map_widget.set_path(disp_path, color="#007AFF", width=4)
                    self.route_paths.append(path_obj)

        if pending_route_fetches:
            self._map_fetch_gen += 1
            gen = self._map_fetch_gen
            threading.Thread(
                target=self._fetch_missing_routes_bg,
                args=(pending_route_fetches, gen),
                daemon=True,
            ).start()

        # Определяем плотные кластеры близких точек (~1.8 км), чтобы подписи не накладывались друг на друга
        compact_mode = bool(self.compact_labels_var.get()) if hasattr(self, "compact_labels_var") else True
        for i, pt in enumerate(ordered_points):
            is_dense = False
            cos_lat = math.cos(math.radians(pt["lat"]))
            for j, other in enumerate(ordered_points):
                if i == j:
                    continue
                d_deg = math.hypot(pt["lat"] - other["lat"], (pt["lon"] - other["lon"]) * cos_lat)
                if d_deg < 0.016:
                    is_dense = True
                    break
            idx = i + 1
            is_selected = self.selected_marker_index == idx
            d_lat, d_lon = self._wgs84_to_map(pt["lat"], pt["lon"])
            meta = {
                "index": idx,
                "name": pt["name"],
                "lat": pt["lat"],
                "lon": pt["lon"],
                "disp_lat": d_lat,
                "disp_lon": d_lon,
                "dense": is_dense,
            }
            if is_selected:
                label_txt = f"#{idx} {pt['name']}"
            elif compact_mode and is_dense:
                label_txt = f"#{idx}"
            elif compact_mode:
                short_name = pt["name"] if len(pt["name"]) <= 24 else pt["name"][:23] + "…"
                label_txt = f"#{idx} {short_name}"
            else:
                label_txt = f"#{idx} {pt['name']}"

            marker_obj = self._create_styled_marker(
                d_lat, d_lon, label_txt, meta, is_selected=is_selected
            )
            meta["marker"] = marker_obj
            self.route_markers.append(marker_obj)
            self.marker_metadata.append(meta)

        if hasattr(self, "cb_map_markers"):
            combo_vals = [f"#{m['index']}. {m['name']}" for m in self.marker_metadata]
            if combo_vals:
                self.cb_map_markers.configure(values=combo_vals)
                if self.selected_marker_index and 1 <= self.selected_marker_index <= len(combo_vals):
                    self.cb_map_markers.set(combo_vals[self.selected_marker_index - 1])
                else:
                    self.cb_map_markers.set(combo_vals[0])
            else:
                self.cb_map_markers.configure(values=["Нет точек на карте"])
                self.cb_map_markers.set("Нет точек на карте")
                if hasattr(self, "lbl_selected_marker"):
                    self.lbl_selected_marker.configure(
                        text="Правый клик по карте — добавить точку; левый клик по маркеру — показать полное название и GPS.",
                        text_color="#4B5563",
                    )

        if auto_fit and self.marker_metadata:
            self.fit_map_to_route()

    def update_route_displays(self):
        if not self.route_segments:
            pickup_name = self.cb_pickup.get().strip() if hasattr(self, "cb_pickup") else ""
            init_txt = f"Старт: {pickup_name} (добавьте следующую точку маршрута)" if pickup_name else ""
            for box in (self.txt_route, self.txt_route_main):
                if box:
                    box.delete("1.0", "end")
                    if init_txt:
                        box.insert("1.0", init_txt)
            return

        try:
            h1, m1 = map(int, self.en_t_out.get().strip().split(":"))
            h2, m2 = map(int, self.en_t_in.get().strip().split(":"))
            start_mins = h1 * 60 + m1
            end_mins = h2 * 60 + m2
            if end_mins < start_mins:
                end_mins += 24 * 60
        except Exception:
            start_mins = 8 * 60
            end_mins = 17 * 60

        total_km = sum(safe_float(s.get("km", 0.0), 0.0) for s in self.route_segments)
        total_duration = max(end_mins - start_mins, len(self.route_segments) * 10)
        current_mins = start_mins

        lines = []
        for idx, seg in enumerate(self.route_segments):
            km = safe_float(seg.get("km", 0.0), 0.0)
            if idx == len(self.route_segments) - 1:
                next_mins = end_mins
            else:
                share = (
                    (km / total_km)
                    if total_km > 0
                    else (1.0 / len(self.route_segments))
                )
                duration = max(int(total_duration * share), 5)
                next_mins = min(current_mins + duration, end_mins)

            t_out_str = f"{(current_mins // 60) % 24:02d}:{current_mins % 60:02d}"
            t_in_str = f"{(next_mins // 60) % 24:02d}:{next_mins % 60:02d}"
            seg["t_out"] = t_out_str
            seg["t_in"] = t_in_str
            current_mins = next_mins

            lines.append(
                f"{idx + 1}. {seg.get('from')} -> {seg.get('to')} | {t_out_str} - {t_in_str} | {km:.1f} км"
            )

        text_val = "\n".join(lines)
        for box in (self.txt_route, self.txt_route_main):
            if box:
                box.delete("1.0", "end")
                box.insert("1.0", text_val)

        if hasattr(self, "en_dist") and total_km > 0:
            cur_d = safe_float(self.en_dist.get(), 0.0)
            if cur_d <= 0.001:
                self.en_dist.delete(0, "end")
                self.en_dist.insert(0, f"{total_km:.1f}")
        if hasattr(self, "lbl_dist_main"):
            only_mh = (
                self.active_mode == "tractors"
                and hasattr(self, "var_only_motohours")
                and bool(self.var_only_motohours.get())
            )
            self.lbl_dist_main.configure(
                text=f"{total_km:.1f} км (не уч.)" if only_mh else f"{total_km:.1f} км"
            )
        if hasattr(self, "lbl_dist_show"):
            self.lbl_dist_show.configure(text=f"{total_km:.1f} км")

    def _append_point_and_build_route(self, loc_name, lat=None, lon=None):
        """
        Универсальное добавление точки маршрута:
        1) Сохраняет точку и координаты в locs (waybills.db) — точка запоминается навсегда.
        2) Если это 1-я точка (нет адреса подачи и маршрут пуст) — ставит её как стартовую (Адрес подачи).
        3) Если стартовая точка уже есть — автоматически строит маршрут от предыдущей точки до новой!
        """
        loc_name = str(loc_name).strip()
        if not loc_name:
            return

        if lat is None or lon is None:
            geocoded = self.geocode_address(loc_name)
            if geocoded:
                lat, lon = geocoded
            else:
                messagebox.showwarning(
                    "Адрес не найден",
                    f"Точка '{loc_name}' не найдена в базе адресов.\nКликните правой кнопкой мыши по карте, чтобы добавить её.",
                )
                return
        else:
            self.db.query(
                "INSERT OR IGNORE INTO locs (name, lat, lon) VALUES (?, ?, ?)",
                (loc_name, float(lat), float(lon)),
            )
            self.refresh_all_tables()

        start_point = (
            self.route_segments[-1]["to"]
            if self.route_segments
            else self.cb_pickup.get().strip()
        )

        # Если это первая точка (например, Буртинский) — назначаем её стартом маршрута
        if not start_point:
            self.cb_pickup.set(loc_name)
            self.update_route_displays()
            self.redraw_map_from_segments()
            self.save_memory()
            return

        if start_point == loc_name:
            messagebox.showwarning("Внимание", "Эта точка уже выбрана последней!")
            return

        locs = self.db.get_list("locs")
        sp = next((r for r in locs if r[1] == start_point), None)
        if not sp:
            g_sp = self.geocode_address(start_point)
            if not g_sp:
                messagebox.showwarning(
                    "Внимание", f"Стартовая точка '{start_point}' не найдена в базе."
                )
                return
            lat1, lon1 = g_sp
        else:
            lat1, lon1 = float(sp[2]), float(sp[3])

        lat2, lon2 = float(lat), float(lon)
        road_km, path_coords = self.get_road_distance(lat1, lon1, lat2, lon2)
        dist = (
            road_km
            if road_km is not None
            else round(calculate_distance(lat1, lon1, lat2, lon2), 1)
        )
        if not path_coords:
            path_coords = [(lat1, lon1), (lat2, lon2)]

        self.route_segments.append(
            {
                "from": start_point,
                "to": loc_name,
                "km": dist,
                "coords": path_coords,
            }
        )

        cur_d = round(sum(safe_float(s.get("km", 0.0), 0.0) for s in self.route_segments), 1)
        self.en_dist.delete(0, "end")
        self.en_dist.insert(0, str(cur_d))

        self.update_route_displays()
        self.redraw_map_from_segments()
        self.recalc_live()
        self.save_memory()

    def add_rt_point(self):
        to_point = self.cb_route_add.get().strip()
        if not to_point:
            messagebox.showwarning("Внимание", "Выберите пункт назначения!")
            return
        # Если адрес подачи тоже введён вручную, сначала геокодируем и запоминаем его
        pickup_val = self.cb_pickup.get().strip()
        if pickup_val:
            self.geocode_address(pickup_val)
        self._append_point_and_build_route(to_point)

    def undo_rt_point(self):
        if not self.route_segments:
            return
        self.route_segments.pop()
        new_d = round(sum(safe_float(s.get("km", 0.0), 0.0) for s in self.route_segments), 1)
        self.en_dist.delete(0, "end")
        self.en_dist.insert(0, str(new_d))
        self.update_route_displays()
        self.redraw_map_from_segments()
        self.recalc_live()
        self.save_memory()

    def clear_map(self):
        self.route_segments.clear()
        self.selected_marker_index = None
        self.en_dist.delete(0, "end")
        self.en_dist.insert(0, "0.0")
        self.update_route_displays()
        self.redraw_map_from_segments()
        self.recalc_live()
        self.save_memory()

    def search_map(self, add_to_route=True):
        q = self.entry_search.get().strip()
        if not q:
            return
        try:
            url = f"https://nominatim.openstreetmap.org/search?q={q}&format=json&limit=1&accept-language=ru"
            resp = requests.get(
                url, headers={"User-Agent": "WaybillApp/1.0"}, timeout=5
            )
            if resp.status_code == 200 and resp.json():
                item = resp.json()[0]
                lat = float(item["lat"])
                lon = float(item["lon"])
                d_lat, d_lon = self._wgs84_to_map(lat, lon)
                self.map_widget.set_position(d_lat, d_lon)
                self.map_widget.set_zoom(14)
                if add_to_route:
                    self._append_point_and_build_route(q, lat, lon)
            else:
                messagebox.showinfo("Поиск", "Адрес не найден!")
        except Exception:
            messagebox.showerror("Ошибка", "Ошибка соединения с сервером карт.")

    def add_map_point(self, coords):
        # Конвертируем координаты клика с экрана обратно в истинный WGS84 (важно для Яндекс Карт EPSG:3395)
        lat, lon = self._map_to_wgs84(coords[0], coords[1])
        suggested_address = ""
        try:
            url = f"https://nominatim.openstreetmap.org/reverse?lat={lat}&lon={lon}&format=json&accept-language=ru"
            resp = requests.get(
                url, headers={"User-Agent": "WaybillApp/1.0"}, timeout=5
            )
            if resp.status_code == 200:
                addr = resp.json().get("address", {})
                city = (
                    addr.get("city")
                    or addr.get("town")
                    or addr.get("village")
                    or ""
                )
                street = addr.get("road") or ""
                house = addr.get("house_number") or ""
                parts = [p for p in (city, street, house) if p]
                suggested_address = ", ".join(parts)
        except Exception:
            suggested_address = ""

        dialog = CustomInputDialog(
            self,
            title="Данные маршрута",
            prompt="Организация или адрес:",
            initial_value=suggested_address,
        )
        self.wait_window(dialog)
        loc_name = dialog.result
        if not loc_name:
            return

        # Автоматически запоминает точку в locs и строит маршрут (1-я точка -> старт, 2-я -> маршрут)
        self._append_point_and_build_route(loc_name, lat, lon)

    def save_route_tmpl(self):
        if not self.route_segments:
            return
        dialog = CustomInputDialog(
            self,
            title="Шаблон",
            prompt="Введите название шаблона маршрута:",
        )
        self.wait_window(dialog)
        t_name = dialog.result
        if not t_name:
            return
        tmpls = self.settings.get("saved_templates", {})
        clean_segs = [
            {"from": s["from"], "to": s["to"], "km": s["km"]}
            for s in self.route_segments
        ]
        tmpls[t_name] = clean_segs
        self.settings["saved_templates"] = tmpls
        self.save_memory()
        self.cb_tmpl.configure(values=list(tmpls.keys()))
        self.cb_tmpl.set(t_name)
        messagebox.showinfo(
            "Шаблон", f"Текущий маршрут сохранен под именем:\n'{t_name}'"
        )

    def load_route_tmpl(self):
        t_name = self.cb_tmpl.get().strip()
        tmpls = self.settings.get("saved_templates", {})
        if t_name in tmpls:
            self.route_segments = [dict(s) for s in tmpls[t_name]]
            if self.route_segments and not self.cb_pickup.get().strip():
                self.cb_pickup.set(self.route_segments[0].get("from", ""))
            total_d = round(
                sum(safe_float(s.get("km", 0.0), 0.0) for s in self.route_segments), 1
            )
            self.en_dist.delete(0, "end")
            self.en_dist.insert(0, str(total_d))
            self.update_route_displays()
            self.redraw_map_from_segments()
            self.recalc_live()
            messagebox.showinfo("Шаблон", f"Шаблон '{t_name}' загружен!")
        else:
            messagebox.showwarning(
                "Шаблон", "Этот шаблон пуст или не существует!"
            )

    def delete_route_tmpl(self):
        t_name = self.cb_tmpl.get().strip()
        tmpls = self.settings.get("saved_templates", {})
        if t_name in tmpls:
            if messagebox.askyesno(
                "Удаление",
                f"Вы действительно хотите удалить шаблон '{t_name}'?",
            ):
                del tmpls[t_name]
                self.settings["saved_templates"] = tmpls
                self.save_memory()
                new_values = list(tmpls.keys()) or [""]
                self.cb_tmpl.configure(values=new_values)
                self.cb_tmpl.set(new_values[0])
                messagebox.showinfo("Удаление", f"Шаблон '{t_name}' удален.")
        else:
            messagebox.showwarning("Удаление", "Такой шаблон не найден.")

    def _rebuild_excel_batch_bg(self, record_ids, mode=None, only_existing=True):
        """
        Пересобирает затронутые пересчётом Excel-файлы:
        - До 20 файлов обновляет сразу (синхронно, ~0.3 сек), чтобы на диске гарантированно были свежие данные.
        - Если файлов больше 20 — первые 10 обновляет сразу, остальные в фоне без зависания интерфейса.
        """
        if not record_ids:
            return 0
        target_mode = mode or self.active_mode
        snapshot = {
            "okpo": self.en_okpo.get().strip() if hasattr(self, "en_okpo") else ("56104921" if target_mode == "tractors" else "48291034"),
            "pickup": self.cb_pickup.get().strip() if hasattr(self, "cb_pickup") else "",
            "mech": self.cb_mech.get().strip() if hasattr(self, "cb_mech") else "",
            "disp": self.cb_disp.get().strip() if hasattr(self, "cb_disp") else "",
            "trailer": self.en_trailer.get().strip() if hasattr(self, "en_trailer") else ("Ковш / Отвал" if target_mode == "tractors" else "без прицепа"),
            "msg_type": self.cb_msg_type.get().strip() if hasattr(self, "cb_msg_type") else "городское / пригородное",
            "trans_type": self.cb_trans_type.get().strip() if hasattr(self, "cb_trans_type") else "служебная",
            "time_tot": self.calc_time(),
        }
        ids_copy = list(dict.fromkeys(record_ids))
        db_inst = self.db_tractors if target_mode == "tractors" else self.db_cars
        drv_map = {r[1]: r for r in db_inst.get_list("drivers")}
        veh_map = {r[1]: r for r in db_inst.get_list("vehicles")}

        self._ensure_excel_bg_worker()
        for rid in ids_copy:
            self._excel_bg_queue.put(
                (rid, target_mode, drv_map, veh_map, snapshot, only_existing)
            )
        return len(ids_copy)

    def _ensure_excel_bg_worker(self):
        if getattr(self, "_excel_bg_worker_started", False):
            return
        self._excel_bg_worker_started = True

        def _worker_loop():
            while True:
                job = self._excel_bg_queue.get()
                try:
                    rid, target_mode, drv_map, veh_map, snapshot, only_existing = job
                    self.rebuild_excel_for_id(
                        rid,
                        mode=target_mode,
                        drv_map=drv_map,
                        veh_map=veh_map,
                        ui_snapshot=snapshot,
                        only_if_exists=only_existing,
                    )
                except Exception:
                    pass
                finally:
                    self._excel_bg_queue.task_done()

        threading.Thread(target=_worker_loop, daemon=True).start()

    def rebuild_excel_for_id(
        self,
        r_id,
        mode=None,
        drv_map=None,
        veh_map=None,
        ui_snapshot=None,
        only_if_exists=False,
    ):
        target_mode = mode or self.active_mode
        db_inst = self.db_tractors if target_mode == "tractors" else self.db_cars
        d = db_inst.get_record_by_id(r_id)
        if not d:
            return (False, "Запись не найдена")

        if drv_map is None:
            drv_map = {r[1]: r for r in db_inst.get_list("drivers")}
        if veh_map is None:
            veh_map = {r[1]: r for r in db_inst.get_list("vehicles")}
        snap = ui_snapshot or {
            "okpo": self.en_okpo.get().strip() if hasattr(self, "en_okpo") else ("56104921" if target_mode == "tractors" else "48291034"),
            "pickup": self.cb_pickup.get().strip() if hasattr(self, "cb_pickup") else "",
            "mech": self.cb_mech.get().strip() if hasattr(self, "cb_mech") else "",
            "disp": self.cb_disp.get().strip() if hasattr(self, "cb_disp") else "",
            "trailer": self.en_trailer.get().strip() if hasattr(self, "en_trailer") else ("Ковш / Отвал" if target_mode == "tractors" else "без прицепа"),
            "msg_type": self.cb_msg_type.get().strip() if hasattr(self, "cb_msg_type") else "городское / пригородное",
            "trans_type": self.cb_trans_type.get().strip() if hasattr(self, "cb_trans_type") else "служебная",
            "time_tot": self.calc_time(),
        }

        if target_mode == "tractors":
            wb_num, d_str, t_out, t_in, org, client, veh, drv, work_type = (
                d[1],
                d[2],
                d[3],
                d[4],
                d[5],
                d[6],
                d[7],
                d[8],
                d[9],
            )
            norm_km, norm_mh, odo_s, odo_e, moto_s, moto_e, moto_hours = (
                d[10],
                d[11],
                d[12],
                d[13],
                d[14],
                d[15],
                d[16],
            )
            f_s, f_a, f_e, f_norm_km, f_norm_mh, f_norm_total, f_act, route_str, dist = (
                d[17],
                d[18],
                d[19],
                d[20],
                d[21],
                d[22],
                d[23],
                d[24],
                d[25],
            )
            form_type = d[26] if len(d) > 26 else "Форма № 412-АПК (МТЗ-82)"
            try:
                raw_items = ast.literal_eval(route_str) if route_str else []
                if not isinstance(raw_items, list):
                    raw_items = []
            except Exception:
                raw_items = []
            works_list = [
                it for it in raw_items if isinstance(it, dict) and it.get("entry_type") == "work"
            ]
            meta_items = [
                it for it in raw_items if isinstance(it, dict) and it.get("entry_type") == "meta"
            ]
            only_mh_flag = any(bool(it.get("only_motohours")) for it in meta_items)
            segs = [
                it
                for it in raw_items
                if isinstance(it, dict) and it.get("entry_type") not in ("work", "meta")
            ]

            drv_row = drv_map.get(drv)
            veh_row = veh_map.get(veh)
            os.makedirs("Готовые_Путевые_Листы/Трактора_МТЗ82", exist_ok=True)
            safe_veh = (
                str(veh)
                .replace(" ", "_")
                .replace("/", "_")
                .replace("(", "")
                .replace(")", "")
            )
            out_file = os.path.join(
                "Готовые_Путевые_Листы/Трактора_МТЗ82",
                f"ПЛ_{wb_num}_{d_str}_{safe_veh}.xlsx",
            )
            existing_resolved = self._resolve_waybill_excel_path(wb_num, d_str, veh, mode="tractors")
            if only_if_exists and not os.path.exists(out_file) and not os.path.exists(existing_resolved):
                return (False, out_file)
            if os.path.exists(existing_resolved) and os.path.abspath(existing_resolved) != os.path.abspath(out_file):
                try:
                    os.remove(existing_resolved)
                except Exception:
                    pass
            data = {
                "wb_num": wb_num,
                "date": d_str,
                "form_type": form_type,
                "org": org,
                "okpo": snap["okpo"],
                "client": client,
                "pickup": segs[0].get("from", "") if segs else snap["pickup"],
                "drv": drv,
                "tab": drv_row[4] if drv_row else "",
                "ser": drv_row[5] if drv_row else "",
                "num": drv_row[6] if drv_row else "",
                "cat": drv_row[2] if drv_row else "",
                "car": veh,
                "reg": veh_row[2] if veh_row else "",
                "gar": veh_row[3] if veh_row else "",
                "mech": snap["mech"],
                "disp": snap["disp"],
                "work_type": work_type,
                "works": works_list,
                "trailer": snap["trailer"],
                "msg_type": snap["msg_type"],
                "trans_type": work_type,
                "odo_s": odo_s,
                "odo_e": odo_e,
                "dist": dist,
                "moto_s": moto_s,
                "moto_e": moto_e,
                "moto_hours": moto_hours,
                "norm_km": norm_km,
                "norm_mh": norm_mh,
                "f_norm_km": f_norm_km,
                "f_norm_mh": f_norm_mh,
                "f_norm_total": f_norm_total,
                "f_act": f_act,
                "f_brand": veh_row[5] if veh_row and len(veh_row) > 5 else "ДТ (дизельное)",
                "f_s": f_s,
                "f_a": f_a,
                "f_e": f_e,
                "time_out": t_out,
                "time_in": t_in,
                "time_tot": snap["time_tot"],
                "segments": segs,
                "only_motohours": only_mh_flag,
            }
            return generate_tractor_wb_excel(out_file, data)

        wb_num, d_str, t_out, t_in, org, client, veh, drv = (
            d[1],
            d[2],
            d[3],
            d[4],
            d[5],
            d[6],
            d[7],
            d[8],
        )
        odo_s, odo_e, f_s, f_a, f_e, f_n, g_s, g_a, g_e, g_n, route_str, dist = (
            d[9],
            d[10],
            d[11],
            d[12],
            d[13],
            d[14],
            d[15],
            d[16],
            d[17],
            d[18],
            d[19],
            d[20],
        )
        try:
            segs = ast.literal_eval(route_str) if route_str else []
            if not isinstance(segs, list):
                segs = []
        except Exception:
            segs = []

        drv_row = drv_map.get(drv)
        veh_row = veh_map.get(veh)

        os.makedirs("Готовые_Путевые_Листы", exist_ok=True)
        safe_veh = str(veh).replace(" ", "_").replace("/", "_")
        out_file = os.path.join(
            "Готовые_Путевые_Листы", f"ПЛ_{wb_num}_{d_str}_{safe_veh}.xlsx"
        )
        existing_resolved = self._resolve_waybill_excel_path(wb_num, d_str, veh, mode="cars")
        if only_if_exists and not os.path.exists(out_file) and not os.path.exists(existing_resolved):
            return (False, out_file)
        if os.path.exists(existing_resolved) and os.path.abspath(existing_resolved) != os.path.abspath(out_file):
            try:
                os.remove(existing_resolved)
            except Exception:
                pass

        f_n_f = safe_float(f_n, 0.0)
        g_n_f = safe_float(g_n, 0.0)
        if g_n_f > 0 and f_n_f <= 0:
            brand_str = veh_row[6] if veh_row and len(veh_row) > 6 and veh_row[6] else "ПБА"
            f_s_str, f_a_str, f_e_str, spent_str = str(g_s), str(g_a), str(g_e), str(g_n)
        elif g_n_f > 0 and f_n_f > 0:
            b1 = veh_row[4] if veh_row else "АИ-92"
            b2 = veh_row[6] if veh_row and len(veh_row) > 6 and veh_row[6] else "ПБА"
            brand_str = f"{b1} / {b2}"
            f_s_str = f"{f_s} / {g_s}"
            f_a_str = f"{f_a} / {g_a}"
            f_e_str = f"{f_e} / {g_e}"
            spent_str = f"{f_n} / {g_n}"
        else:
            brand_str = veh_row[4] if veh_row else "АИ-92"
            f_s_str, f_a_str, f_e_str, spent_str = str(f_s), str(f_a), str(f_e), str(f_n)

        data = {
            "wb_num": wb_num,
            "date": d_str,
            "org": org,
            "okpo": snap["okpo"],
            "client": client,
            "pickup": segs[0].get("from", "") if segs else snap["pickup"],
            "drv": drv,
            "tab": drv_row[4] if drv_row else "",
            "ser": drv_row[5] if drv_row else "",
            "num": drv_row[6] if drv_row else "",
            "cat": drv_row[2] if drv_row else "",
            "car": veh,
            "reg": veh_row[2] if veh_row else "",
            "gar": veh_row[3] if veh_row else "",
            "mech": snap["mech"],
            "disp": snap["disp"],
            "trailer": snap["trailer"],
            "msg_type": snap["msg_type"],
            "trans_type": snap["trans_type"],
            "odo_s": odo_s,
            "odo_e": odo_e,
            "time_out": t_out,
            "time_in": t_in,
            "time_tot": snap["time_tot"],
            "f_brand": brand_str,
            "f_s": f_s_str,
            "f_a": f_a_str,
            "f_e": f_e_str,
            "spent": spent_str,
            "f_act": spent_str,
            "dist": dist,
            "seg_from": [s.get("from", "") for s in segs],
            "seg_to": [s.get("to", "") for s in segs],
            "seg_t_out": [s.get("t_out", t_out) for s in segs],
            "seg_t_in": [s.get("t_in", t_in) for s in segs],
            "seg_km": [s.get("km", 0.0) for s in segs],
        }
        return generate_f3_from_excel("templates/put-list-f3.xlsx", out_file, data)

    def _increment_wb_num(self, wb_num_str):
        """Увеличивает номер путевого листа на +1 (поддерживает 1->2, 001->002, Т-005->Т-006)."""
        s = str(wb_num_str or "").strip()
        m = re.search(r"(\d+)(\D*)$", s)
        if m:
            num_str = m.group(1)
            next_num = int(num_str) + 1
            if len(num_str) > 1 and num_str.startswith("0"):
                padded = str(next_num).zfill(len(num_str))
            else:
                padded = str(next_num)
            return s[: m.start(1)] + padded + m.group(2)
        return "2" if not s else f"{s}-2"

    def _increment_date_str(self, d_str):
        """Возвращает следующую календарную дату в формате ДД.ММ.ГГГГ."""
        s = str(d_str or "").strip()
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                dt = datetime.strptime(s, fmt)
                return (dt + timedelta(days=1)).strftime("%d.%m.%Y")
            except Exception:
                continue
        return (datetime.now() + timedelta(days=1)).strftime("%d.%m.%Y")

    def _advance_next_waybill(self, wb_num, d_str):
        """Автоматически переводит форму на следующий № ПЛ (по хронологии журнала) и следующую дату после выписки."""
        try:
            total_rows = len(self.db.get_list("history"))
            next_num = str(total_rows + 1) if total_rows > 0 else self._increment_wb_num(wb_num)
        except Exception:
            next_num = self._increment_wb_num(wb_num)
        next_date = self._increment_date_str(d_str)
        if hasattr(self, "en_wb_num"):
            self.en_wb_num.delete(0, "end")
            self.en_wb_num.insert(0, next_num)
        if hasattr(self, "en_date"):
            self.en_date.delete(0, "end")
            self.en_date.insert(0, next_date)
        return next_num, next_date

    def generate_main(self, create_excel=True):
        try:
            # Сначала синхронизируем маршрут и дистанцию, затем выполняем свежий расчёт ДТ/ГСМ
            self.update_route_displays()
            self.recalc_live()

            wb_num = self.en_wb_num.get().strip() or "1"
            d_str = self.en_date.get().strip()
            t_out = self.en_t_out.get().strip()
            t_in = self.en_t_in.get().strip()
            org = self.cb_org.get().strip()
            client = self.cb_client.get().strip()
            pickup = self.cb_pickup.get().strip()
            drv = self.cb_drv.get().strip()
            veh = self.cb_veh.get().strip()
            mech = self.cb_mech.get().strip()
            disp = self.cb_disp.get().strip()

            dist = safe_float(self.en_dist.get(), 0.0)
            odo_s = safe_float(self.en_odo_s.get(), 0.0)
            odo_e = round(odo_s + dist, 1)

            f_s = safe_float(self.en_f_s.get(), 0.0)
            f_a = safe_float(self.en_f_a.get(), 0.0)
            f_e = safe_float(self.en_f_e.get(), 0.0)
            f_n = self.last_spent_p

            g_s = safe_float(self.en_g_s.get(), 0.0)
            g_a = safe_float(self.en_g_a.get(), 0.0)
            g_e = safe_float(self.en_g_e.get(), 0.0)
            g_n = self.last_spent_g

            # Определяем: редактируем существующую запись (через «Редактировать» или по совпадению даты/ТС) или создаём новую
            target_edit_id = self.editing_id
            old_file_to_replace = self.old_filename if self.editing_id else None
            old_veh_for_chain = self._editing_old_veh if self.editing_id else None
            replaced_existing = bool(self.editing_id)

            if not target_edit_id:
                existing_row = self.db.find_by_veh_and_date(veh, d_str, wb_num)
                if existing_row:
                    replace_ok = messagebox.askyesno(
                        "Замена путевого листа",
                        f"Путевой лист для «{veh}» на дату {d_str} (№ {existing_row[1]}) уже есть в журнале.\n\n"
                        "Заменить старый путевой лист обновлённым и автоматически пересчитать все последующие путёвки?",
                    )
                    if replace_ok:
                        target_edit_id = existing_row[0]
                        old_file_to_replace = self._resolve_waybill_excel_path(
                            existing_row[1], existing_row[2], existing_row[7], mode=self.active_mode
                        )
                        old_veh_for_chain = str(existing_row[7])
                        replaced_existing = True

            if self.active_mode == "tractors":
                only_mh = (
                    hasattr(self, "var_only_motohours")
                    and bool(self.var_only_motohours.get())
                )
                eff_dist = 0.0 if only_mh else dist
                odo_e = round(odo_s + eff_dist, 1)

                if self.tractor_works:
                    work_type = "; ".join(
                        f"{w.get('name', '')} ({safe_float(w.get('mh'), 0.0):g} м/ч)"
                        for w in self.tractor_works
                    )
                else:
                    work_type = self.cb_work_type.get().strip() if hasattr(self, "cb_work_type") else ""
                form_type = self.cb_form_type.get().strip() if hasattr(self, "cb_form_type") else "Форма № 412-АПК"
                norm_km = self._get_tractor_norm_km(veh)
                norm_mh = safe_float(self.en_norm_mh.get(), 5.2)
                moto_s = safe_float(self.en_moto_s.get(), 0.0)
                moto_hours = safe_float(self.en_moto_h.get(), 0.0)
                moto_e = round(moto_s + moto_hours, 1)
                f_act_raw = self.en_f_act.get().strip()
                f_act = safe_float(f_act_raw, self.last_dt_total) if f_act_raw else self.last_dt_total

                works_payload = [
                    {
                        "entry_type": "work",
                        "name": str(w.get("name", "")).strip(),
                        "mh": round(safe_float(w.get("mh"), 0.0), 2),
                        "norm_mh": round(safe_float(w.get("norm_mh"), norm_mh), 2),
                        "dt_l": round(safe_float(w.get("dt_l"), 0.0), 2),
                    }
                    for w in self.tractor_works
                ]
                meta_payload = (
                    [{"entry_type": "meta", "only_motohours": True}] if only_mh else []
                )
                db_segments = [
                    {
                        "from": s.get("from", ""),
                        "to": s.get("to", ""),
                        "t_out": s.get("t_out", t_out),
                        "t_in": s.get("t_in", t_in),
                        "km": s.get("km", 0.0),
                        "mh": s.get("mh", 0.0 if works_payload else (moto_hours if idx == 0 else 0.0)),
                        "work_type": s.get("work_type", work_type),
                    }
                    for idx, s in enumerate(self.route_segments)
                ]
                route_repr = repr(db_segments + works_payload + meta_payload)

                if target_edit_id:
                    saved_id = target_edit_id
                    self.db_tractors.update_wb(
                        saved_id,
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
                        self.last_dt_km,
                        self.last_dt_mh,
                        self.last_dt_total,
                        f_act,
                        route_repr,
                        eff_dist,
                        form_type,
                    )
                else:
                    self.db_tractors.save_wb(
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
                        self.last_dt_km,
                        self.last_dt_mh,
                        self.last_dt_total,
                        f_act,
                        route_repr,
                        eff_dist,
                        form_type,
                    )
                    row_last = self.db_tractors.query("SELECT MAX(id) FROM history")
                    saved_id = row_last[0][0] if row_last and row_last[0] else None

                self.editing_id = None
                self.old_filename = None
                self._editing_old_veh = None
                self.btn_gen.configure(text="Выписать путевой лист")
                self.btn_save_db.configure(text="Внести в журнал")

                id_map, renamed_records = self.db_tractors.renumber_history_by_date_detailed(
                    renumber_wb_num=True
                )
                saved_id = id_map.get(saved_id, saved_id)
                updated_ids = list(self.db_tractors.recalculate_chain(veh) or [])
                if old_veh_for_chain and old_veh_for_chain != veh:
                    updated_ids.extend(self.db_tractors.recalculate_chain(old_veh_for_chain) or [])
                other_ids = [
                    u for u in dict.fromkeys(updated_ids) if u != saved_id
                ]
                renamed_excel_ids = self._handle_renamed_wb_excels(
                    renamed_records, mode="tractors", exclude_ids={saved_id}
                )

                excel_ok, excel_res = (True, "")
                if create_excel:
                    excel_ok, excel_res = self.rebuild_excel_for_id(
                        saved_id, mode="tractors", only_if_exists=False
                    )
                    if excel_ok and old_file_to_replace and os.path.exists(old_file_to_replace):
                        if os.path.abspath(old_file_to_replace) != os.path.abspath(excel_res):
                            try:
                                os.remove(old_file_to_replace)
                            except Exception:
                                pass
                elif replaced_existing and old_file_to_replace and os.path.exists(old_file_to_replace):
                    # Если старый Excel существовал, обновляем и заменяем его даже при кнопке «Только сохранить»
                    ok_rep, res_rep = self.rebuild_excel_for_id(
                        saved_id, mode="tractors", only_if_exists=False
                    )
                    if ok_rep and os.path.abspath(old_file_to_replace) != os.path.abspath(res_rep):
                        try:
                            os.remove(old_file_to_replace)
                        except Exception:
                            pass

                if renamed_excel_ids:
                    self._rebuild_excel_batch_bg(
                        renamed_excel_ids, mode="tractors", only_existing=False
                    )
                rem_other_ids = [u for u in other_ids if u not in set(renamed_excel_ids)]
                if rem_other_ids:
                    self._rebuild_excel_batch_bg(
                        rem_other_ids, mode="tractors", only_existing=True
                    )

                next_note = ""
                if not replaced_existing:
                    next_num, next_date = self._advance_next_waybill(wb_num, d_str)
                    next_note = f"\nФорма переведена на следующий ПЛ: № {next_num} от {next_date}."

                self.refresh_history()
                self.save_memory()
                self._run_bg_backup()
                self.load_vehicle_data()

                chain_note = (
                    f"\nАвтоматически пересчитано последующих ПЛ по цепочке: {len(other_ids)} (Excel-файлы обновлены)."
                    if other_ids
                    else ""
                )
                replace_note = " (старый ПЛ заменён)" if replaced_existing else ""
                if not create_excel:
                    messagebox.showinfo(
                        "Успех",
                        f"Запись спецтехники (МТЗ-82) сохранена в журнал{replace_note}.{next_note}{chain_note}",
                    )
                    return

                if excel_ok:
                    messagebox.showinfo(
                        "Успех",
                        f"Путевой лист спецтехники (МТЗ-82) сформирован{replace_note} без открытия Excel:\n{excel_res}{next_note}{chain_note}",
                    )
                else:
                    messagebox.showerror("Ошибка", excel_res)
                return

            db_segments = [
                {
                    "from": s.get("from", ""),
                    "to": s.get("to", ""),
                    "t_out": s.get("t_out", t_out),
                    "t_in": s.get("t_in", t_in),
                    "km": s.get("km", 0.0),
                }
                for s in self.route_segments
            ]
            route_repr = repr(db_segments)

            if target_edit_id:
                saved_id = target_edit_id
                self.db.update_wb(
                    saved_id,
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
                self.db.save_wb(
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
                row_last = self.db.query("SELECT MAX(id) FROM history")
                saved_id = row_last[0][0] if row_last and row_last[0] else None

            self.editing_id = None
            self.old_filename = None
            self._editing_old_veh = None
            self.btn_gen.configure(text="Выписать путевой лист")
            self.btn_save_db.configure(text="Внести в журнал")

            id_map, renamed_records = self.db.renumber_history_by_date_detailed(
                renumber_wb_num=True
            )
            saved_id = id_map.get(saved_id, saved_id)
            updated_ids = list(self.db.recalculate_chain(veh) or [])
            if old_veh_for_chain and old_veh_for_chain != veh:
                updated_ids.extend(self.db.recalculate_chain(old_veh_for_chain) or [])
            other_ids = [u for u in dict.fromkeys(updated_ids) if u != saved_id]
            renamed_excel_ids = self._handle_renamed_wb_excels(
                renamed_records, mode="cars", exclude_ids={saved_id}
            )

            excel_ok, excel_res = (True, "")
            if create_excel:
                excel_ok, excel_res = self.rebuild_excel_for_id(
                    saved_id, mode="cars", only_if_exists=False
                )
                if excel_ok and old_file_to_replace and os.path.exists(old_file_to_replace):
                    if os.path.abspath(old_file_to_replace) != os.path.abspath(excel_res):
                        try:
                            os.remove(old_file_to_replace)
                        except Exception:
                            pass
            elif replaced_existing and old_file_to_replace and os.path.exists(old_file_to_replace):
                ok_rep, res_rep = self.rebuild_excel_for_id(
                    saved_id, mode="cars", only_if_exists=False
                )
                if ok_rep and os.path.abspath(old_file_to_replace) != os.path.abspath(res_rep):
                    try:
                        os.remove(old_file_to_replace)
                    except Exception:
                        pass

            if renamed_excel_ids:
                self._rebuild_excel_batch_bg(
                    renamed_excel_ids, mode="cars", only_existing=False
                )
            rem_other_ids = [u for u in other_ids if u not in set(renamed_excel_ids)]
            if rem_other_ids:
                self._rebuild_excel_batch_bg(
                    rem_other_ids, mode="cars", only_existing=True
                )

            next_note = ""
            if not replaced_existing:
                next_num, next_date = self._advance_next_waybill(wb_num, d_str)
                next_note = f"\nФорма переведена на следующий ПЛ: № {next_num} от {next_date}."

            self.refresh_history()
            self.save_memory()
            self._run_bg_backup()
            self.load_vehicle_data()

            chain_note = (
                f"\nАвтоматически пересчитано последующих ПЛ по цепочке: {len(other_ids)} (Excel-файлы обновлены)."
                if other_ids
                else ""
            )
            replace_note = " (старый ПЛ заменён)" if replaced_existing else ""
            if not create_excel:
                messagebox.showinfo(
                    "Успех", f"Запись успешно сохранена в журнал{replace_note}.{next_note}{chain_note}"
                )
                return

            if excel_ok:
                messagebox.showinfo(
                    "Успех",
                    f"Документ сформирован{replace_note} и сохранён в журнал (без открытия Excel):\n{excel_res}{next_note}{chain_note}",
                )
            else:
                messagebox.showerror("Ошибка", excel_res)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Проверьте данные: {str(e)}")

    def refresh_history(self):
        if not hasattr(self, "tr_hist"):
            return
        for item in self.tr_hist.get_children():
            self.tr_hist.delete(item)
        for r in self.get_sorted_history():
            if self.active_mode == "tractors" and len(r) >= 26:
                odo_str = f"{r[12]}->{r[13]} км | {r[14]}->{r[15]} м/ч"
                dt_str = f"ДТ: {r[17]} + {r[18]} - {r[23]} = {r[19]} л"
                wt_str = f"{r[9]} ({r[16]} м/ч)"
                self.tr_hist.insert(
                    "",
                    "end",
                    values=(
                        r[0],
                        r[1],
                        r[2],
                        r[3],
                        r[4],
                        r[7],
                        r[8],
                        f"{r[25]} км / {r[16]} м/ч",
                        odo_str,
                        dt_str,
                        wt_str,
                    ),
                )
            else:
                odo_str = f"{r[9]} -> {r[10]}"
                petrol_str = f"{r[11]} + {r[12]} - {r[14]} = {r[13]}"
                gas_str = f"{r[15]} + {r[16]} - {r[18]} = {r[17]}"
                self.tr_hist.insert(
                    "",
                    "end",
                    values=(
                        r[0],
                        r[1],
                        r[2],
                        r[3],
                        r[4],
                        r[7],
                        r[8],
                        r[20],
                        odo_str,
                        petrol_str,
                        gas_str,
                    ),
                )

    def delete_history(self):
        sel = self.tr_hist.selection()
        if not sel:
            return
        affected_vehs = set()
        for item in sel:
            vals = self.tr_hist.item(item, "values")
            wb_id = vals[0]
            if len(vals) > 5 and vals[5]:
                affected_vehs.add(str(vals[5]))
            self.db.query("DELETE FROM history WHERE id=?", (wb_id,))

        _, renamed_records = self.db.renumber_history_by_date_detailed(
            renumber_wb_num=True
        )
        all_updated = []
        for v_mark in affected_vehs:
            all_updated.extend(self.db.recalculate_chain(v_mark) or [])
        renamed_excel_ids = self._handle_renamed_wb_excels(
            renamed_records, mode=self.active_mode
        )
        if renamed_excel_ids:
            self._rebuild_excel_batch_bg(
                renamed_excel_ids, mode=self.active_mode, only_existing=False
            )
        rem_updated = [u for u in dict.fromkeys(all_updated) if u not in set(renamed_excel_ids)]
        if rem_updated:
            self._rebuild_excel_batch_bg(
                rem_updated, mode=self.active_mode, only_existing=True
            )

        self.refresh_history()
        self.load_vehicle_data()
        self._run_bg_backup()

    def generate_excel_from_hist(self):
        sel = self.tr_hist.selection()
        if not sel:
            messagebox.showwarning(
                "Внимание", "Выберите запись для формирования путевого листа."
            )
            return
        drv_map = {r[1]: r for r in self.db.get_list("drivers")}
        veh_map = {r[1]: r for r in self.db.get_list("vehicles")}
        count = 0
        for item in sel:
            r_id = self.tr_hist.item(item, "values")[0]
            ok, _ = self.rebuild_excel_for_id(
                r_id, mode=self.active_mode, drv_map=drv_map, veh_map=veh_map, only_if_exists=False
            )
            if ok:
                count += 1
        messagebox.showinfo(
            "Успех",
            f"Сформировано путевых листов (без открытия Excel): {count}\nПапка: 'Готовые_Путевые_Листы'.",
        )

    def open_selected_excel(self):
        """Открывает выбранный путевой лист в Excel вручную (по желанию пользователя)."""
        sel = self.tr_hist.selection()
        if not sel:
            messagebox.showwarning("Внимание", "Выберите запись в журнале для открытия Excel.")
            return
        vals = self.tr_hist.item(sel[0], "values")
        r_id, wb_num, date_str, v_mark = vals[0], vals[1], vals[2], vals[5]
        fpath = self._resolve_waybill_excel_path(wb_num, date_str, v_mark)
        if not os.path.exists(fpath):
            ok, res = self.rebuild_excel_for_id(r_id, mode=self.active_mode, only_if_exists=False)
            if ok and os.path.exists(res):
                fpath = os.path.abspath(res)
            else:
                messagebox.showerror("Ошибка", f"Не удалось создать файл Excel:\n{res}")
                return
        if hasattr(os, "startfile"):
            os.startfile(os.path.abspath(fpath))
        else:
            messagebox.showinfo("Файл готов", f"Путевой лист:\n{fpath}")

    def recalculate_selected_or_all_chains(self):
        """Полный хронологический пересчёт цепочки спидометра, моточасов и топлива (с начала года)."""
        sel = self.tr_hist.selection()
        vehs = set()
        for item in sel:
            vals = self.tr_hist.item(item, "values")
            if len(vals) > 5 and vals[5]:
                vehs.add(str(vals[5]))

        _, renamed_records = self.db.renumber_history_by_date_detailed(
            renumber_wb_num=True
        )
        updated_ids = []
        if vehs:
            for v in vehs:
                updated_ids.extend(self.db.recalculate_chain(v) or [])
            scope_txt = ", ".join(sorted(vehs))
        else:
            updated_ids = self.db.recalculate_all_vehicles() or []
            scope_txt = "всех ТС в журнале"

        renamed_excel_ids = self._handle_renamed_wb_excels(
            renamed_records, mode=self.active_mode
        )
        if renamed_excel_ids:
            self._rebuild_excel_batch_bg(
                renamed_excel_ids, mode=self.active_mode, only_existing=False
            )
        rem_updated = [u for u in dict.fromkeys(updated_ids) if u not in set(renamed_excel_ids)]
        if rem_updated:
            self._rebuild_excel_batch_bg(
                rem_updated, mode=self.active_mode, only_existing=True
            )
        self.refresh_history()
        self.load_vehicle_data()
        self._run_bg_backup()
        messagebox.showinfo(
            "Пересчёт завершён",
            f"Журнал отсортирован по датам, номера ПЛ (1..N) и цепочка пересчитаны для: {scope_txt}.\n"
            f"Перенумеровано ПЛ: {len(renamed_records)}, обновлено по цепочке: {len(updated_ids)} (существующие Excel-файлы обновляются в фоне).",
        )

    def load_history_for_edit(self):
        sel = self.tr_hist.selection()
        if not sel:
            return
        r_id = self.tr_hist.item(sel[0], "values")[0]
        d = self.db.get_record_by_id(r_id)
        if not d:
            return
        self.editing_id = d[0]
        self._editing_old_veh = str(d[7])
        self.old_filename = self._resolve_waybill_excel_path(d[1], d[2], d[7], mode=self.active_mode)

        self.en_wb_num.delete(0, "end")
        self.en_wb_num.insert(0, str(d[1]))
        self.en_date.delete(0, "end")
        self.en_date.insert(0, str(d[2]))
        self.en_t_out.delete(0, "end")
        self.en_t_out.insert(0, str(d[3]))
        self.en_t_in.delete(0, "end")
        self.en_t_in.insert(0, str(d[4]))
        self.cb_org.set(str(d[5]))
        self.cb_client.set(str(d[6]))
        self.cb_veh.set(str(d[7]))
        self.cb_drv.set(str(d[8]))

        self.en_f_act.delete(0, "end")
        self.en_g_act.delete(0, "end")
        self.tractor_works = []
        if self.active_mode == "tractors" and len(d) >= 26:
            if hasattr(self, "cb_work_type"):
                self.cb_work_type.set(str(d[9] or ""))
            if hasattr(self, "cb_form_type") and len(d) > 26 and d[26]:
                self.cb_form_type.set(str(d[26]))
            self._tractor_norm_km = safe_float(d[10], 0.0) or self._get_tractor_norm_km(str(d[7]))
            self.en_f_norm.delete(0, "end")
            self.en_f_norm.insert(0, str(self._tractor_norm_km))
            if hasattr(self, "en_norm_mh"):
                self.en_norm_mh.delete(0, "end")
                self.en_norm_mh.insert(0, str(d[11] or 5.2))
            self.en_odo_s.delete(0, "end")
            self.en_odo_s.insert(0, str(d[12] or 0.0))
            if hasattr(self, "en_moto_s"):
                self.en_moto_s.delete(0, "end")
                self.en_moto_s.insert(0, str(d[14] or 0.0))
            if hasattr(self, "en_moto_h"):
                self.en_moto_h.delete(0, "end")
                self.en_moto_h.insert(0, str(d[16] or 0.0))
            self.en_f_s.delete(0, "end")
            self.en_f_s.insert(0, str(d[17] or 0.0))
            self.en_f_a.delete(0, "end")
            self.en_f_a.insert(0, str(d[18] or 0.0))
            f_norm_tot_v = safe_float(d[22], 0.0)
            f_act_v = safe_float(d[23], f_norm_tot_v)
            if abs(f_act_v - f_norm_tot_v) > 0.01:
                self.en_f_act.insert(0, f"{f_act_v:g}")
            self.en_dist.delete(0, "end")
            self.en_dist.insert(0, str(d[25] or 0.0))
            route_raw = d[24]
        else:
            self.en_odo_s.delete(0, "end")
            self.en_odo_s.insert(0, str(d[9]))
            self.en_f_s.delete(0, "end")
            self.en_f_s.insert(0, str(d[11]))
            self.en_f_a.delete(0, "end")
            self.en_f_a.insert(0, str(d[12]))
            self.en_g_s.delete(0, "end")
            self.en_g_s.insert(0, str(d[15]))
            self.en_g_a.delete(0, "end")
            self.en_g_a.insert(0, str(d[16]))
            self.en_dist.delete(0, "end")
            self.en_dist.insert(0, str(d[20]))
            route_raw = d[19]

        only_mh_flag = False
        try:
            raw_items = ast.literal_eval(route_raw) if route_raw else []
            if isinstance(raw_items, list):
                self.tractor_works = [
                    it for it in raw_items if isinstance(it, dict) and it.get("entry_type") == "work"
                ]
                meta_items = [
                    it for it in raw_items if isinstance(it, dict) and it.get("entry_type") == "meta"
                ]
                only_mh_flag = any(bool(it.get("only_motohours")) for it in meta_items)
                self.route_segments = [
                    it
                    for it in raw_items
                    if isinstance(it, dict) and it.get("entry_type") not in ("work", "meta")
                ]
            else:
                self.route_segments = []
        except Exception:
            self.route_segments = []

        if hasattr(self, "var_only_motohours"):
            self.var_only_motohours.set(
                bool(only_mh_flag) if self.active_mode == "tractors" else False
            )

        if self.route_segments:
            self.cb_pickup.set(self.route_segments[0].get("from", ""))
        self._refresh_tractor_works_tree()
        self.btn_gen.configure(text="Сохранить и заменить ПЛ")
        self.btn_save_db.configure(text="Только сохранить изменения")
        self.update_route_displays()
        self.after(60, self.redraw_map_from_segments)
        self.recalc_live()
        self.tabview.set("Заполнение путевого листа")

    def _resolve_waybill_excel_path(self, wb_num, date_str, v_mark, mode=None):
        target_mode = mode or self.active_mode
        base_dir = (
            os.path.join("Готовые_Путевые_Листы", "Трактора_МТЗ82")
            if target_mode == "tractors"
            else "Готовые_Путевые_Листы"
        )
        safe_veh = (
            str(v_mark)
            .replace(" ", "_")
            .replace("/", "_")
            .replace("(", "")
            .replace(")", "")
        )
        direct = os.path.abspath(os.path.join(base_dir, f"ПЛ_{wb_num}_{date_str}_{safe_veh}.xlsx"))
        if os.path.exists(direct):
            return direct
        alt_veh = str(v_mark).replace(" ", "_").replace("/", "_")
        direct2 = os.path.abspath(os.path.join(base_dir, f"ПЛ_{wb_num}_{date_str}_{alt_veh}.xlsx"))
        if os.path.exists(direct2):
            return direct2
        if os.path.isdir(base_dir):
            prefix = f"ПЛ_{wb_num}_{date_str}_"
            for fn in os.listdir(base_dir):
                if fn.startswith(prefix) and fn.endswith(".xlsx") and not fn.startswith("~$"):
                    return os.path.abspath(os.path.join(base_dir, fn))
        return direct

    def single_print_selected(self):
        sel = self.tr_hist.selection()
        if not sel:
            messagebox.showwarning(
                "Внимание", "Выберите документ для одиночной печати."
            )
            return
        vals = self.tr_hist.item(sel[0], "values")
        wb_num, date_str, v_mark = vals[1], vals[2], vals[5]
        fpath = self._resolve_waybill_excel_path(wb_num, date_str, v_mark)
        if os.path.exists(fpath):
            self.print_excel_files([fpath], duplex=True)
        else:
            messagebox.showwarning(
                "Внимание", "Файл не найден. Сначала нажмите «Сформировать ПЛ»."
            )

    def mass_print_selected(self):
        sel = self.tr_hist.selection()
        if not sel:
            messagebox.showwarning(
                "Внимание", "Выберите документы для вывода на печать."
            )
            return
        file_paths = []
        for item in sel:
            vals = self.tr_hist.item(item, "values")
            wb_num, date_str, v_mark = vals[1], vals[2], vals[5]
            fpath = self._resolve_waybill_excel_path(wb_num, date_str, v_mark)
            if os.path.exists(fpath):
                file_paths.append(fpath)
        if not file_paths:
            messagebox.showwarning(
                "Внимание", "Файлы не найдены. Сначала сформируйте их."
            )
            return
        self.print_excel_files(file_paths, duplex=True)

    def _prepare_print_copy(self, src_path, tmp_dir, idx=0, for_pdf=True, is_landscape=False):
        """Готовит временную копию Excel: выравнивает пропорции 2-й страницы под 1-ю и задаёт точный масштаб 1x1."""
        import openpyxl
        from openpyxl.utils import get_column_letter

        dst_path = os.path.join(tmp_dir, f"wb_{idx:03d}.xlsx")
        wb = openpyxl.load_workbook(src_path)
        orient = "portrait" if for_pdf else ("landscape" if is_landscape else "portrait")

        # Выравниваем пропорции 2-й страницы (для совместимости со старыми файлами) под вертикальный формат 1-й
        if len(wb.worksheets) >= 2:
            ws2 = wb.worksheets[1]
            s2_widths = {
                "A": 3.5, "B": 4.0, "C": 6.0, "D": 6.5, "E": 6.5, "F": 6.5,
                "G": 5.5, "H": 5.5, "I": 5.5, "J": 5.5, "K": 4.2, "L": 4.2,
                "M": 4.2, "N": 4.2, "O": 4.2, "P": 4.2, "Q": 4.5, "R": 4.5, "S": 4.5,
            }
            cur_total_w = sum(
                (ws2.column_dimensions[get_column_letter(c)].width or 8.43)
                for c in range(1, 20)
            )
            if cur_total_w > 108.0:
                for col_l, w_val in s2_widths.items():
                    ws2.column_dimensions[col_l].width = w_val
                for r_idx in range(1, 41):
                    cur_h = ws2.row_dimensions[r_idx].height
                    if r_idx in (1, 2, 3, 4):
                        ws2.row_dimensions[r_idx].height = max(cur_h or 19.0, 21.0)
                    elif 5 <= r_idx <= 30:
                        ws2.row_dimensions[r_idx].height = max(cur_h or 16.0, 18.0)
                    else:
                        ws2.row_dimensions[r_idx].height = max(cur_h or 18.0, 20.0)

        for s_idx, ws in enumerate(wb.worksheets[:2]):
            ws.page_setup.orientation = orient
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 1
            ws.print_options.horizontalCentered = True
            ws.print_options.verticalCentered = True
            ws.page_margins.left = 0.15
            ws.page_margins.right = 0.15
            ws.page_margins.top = 0.15
            ws.page_margins.bottom = 0.15
            ws.page_margins.header = 0.0
            ws.page_margins.footer = 0.0
            if s_idx == 0:
                ws.print_area = "A1:CH56"
            elif s_idx == 1:
                ws.print_area = "A1:S40"
        wb.save(dst_path)
        return os.path.abspath(dst_path)

    def _export_workbooks_to_pdfs_windows(self, prepared_paths, tmp_dir):
        """Экспортирует все путевые листы в PDF за ОДИН сеанс Excel (без буфера обмена и без зависаний)."""
        pdf_paths = [os.path.abspath(os.path.join(tmp_dir, f"wb_{i:03d}.pdf")) for i in range(len(prepared_paths))]
        try:
            import win32com.client

            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                for src, pdf in zip(prepared_paths, pdf_paths):
                    wb = excel.Workbooks.Open(src, ReadOnly=True)
                    try:
                        wb.ExportAsFixedFormat(0, pdf)
                    finally:
                        wb.Close(SaveChanges=False)
            finally:
                try:
                    excel.Quit()
                except Exception:
                    pass
            if all(os.path.exists(p) and os.path.getsize(p) > 500 for p in pdf_paths):
                return pdf_paths
        except Exception:
            pass

        manifest = [
            {"src": os.path.abspath(src), "pdf": pdf}
            for src, pdf in zip(prepared_paths, pdf_paths)
        ]
        manifest_path = os.path.join(tmp_dir, "export_manifest.json")
        with open(manifest_path, "w", encoding="utf-8") as mf:
            json.dump(manifest, mf, ensure_ascii=False)

        ps_path = os.path.join(tmp_dir, "export_pdfs.ps1")
        ps_code = (
            "$ErrorActionPreference = 'Stop'\n"
            f"$items = Get-Content -Raw -Encoding UTF8 '{manifest_path.replace(chr(39), chr(39)*2)}' | ConvertFrom-Json\n"
            "if ($items -isnot [System.Array]) { $items = @($items) }\n"
            "$excel = New-Object -ComObject Excel.Application\n"
            "$excel.Visible = $false\n"
            "$excel.DisplayAlerts = $false\n"
            "try {\n"
            "  foreach ($it in $items) {\n"
            "    $wb = $excel.Workbooks.Open($it.src, $null, $true)\n"
            "    try { $wb.ExportAsFixedFormat(0, $it.pdf) } finally { $wb.Close($false) }\n"
            "  }\n"
            "} finally {\n"
            "  $excel.Quit()\n"
            "  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null\n"
            "}\n"
        )
        with open(ps_path, "w", encoding="utf-8-sig") as pf:
            pf.write(ps_code)

        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps_path],
            check=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if not all(os.path.exists(p) and os.path.getsize(p) > 500 for p in pdf_paths):
            raise RuntimeError("Не удалось экспортировать Excel в PDF.")
        return pdf_paths

    def _render_pdfs_to_pngs_windows(self, pdf_paths, tmp_dir):
        """Рендерит страницы 1 и 2 всех PDF в PNG за ОДИН вызов встроенного WinRT Windows.Data.Pdf (Windows 10/11)."""
        manifest = []
        expected_pairs = []
        for idx, pdf in enumerate(pdf_paths):
            p1 = os.path.abspath(os.path.join(tmp_dir, f"page_{idx:03d}_1.png"))
            p2 = os.path.abspath(os.path.join(tmp_dir, f"page_{idx:03d}_2.png"))
            manifest.append({"pdf": os.path.abspath(pdf), "p1": os.path.basename(p1), "p2": os.path.basename(p2)})
            expected_pairs.append((p1, p2))

        manifest_path = os.path.join(tmp_dir, "render_manifest.json")
        with open(manifest_path, "w", encoding="utf-8") as mf:
            json.dump(manifest, mf, ensure_ascii=False)

        ps_path = os.path.join(tmp_dir, "render_pdfs.ps1")
        ps_code = (
            "$ErrorActionPreference = 'Stop'\n"
            "Add-Type -AssemblyName System.Runtime.WindowsRuntime\n"
            "$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]\n"
            "Function Await($WinRtTask, $ResultType) {\n"
            "  $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)\n"
            "  $netTask = $asTask.Invoke($null, @($WinRtTask))\n"
            "  $netTask.Wait(-1) | Out-Null\n"
            "  $netTask.Result\n"
            "}\n"
            "Function AwaitAction($WinRtAction) {\n"
            "  $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]\n"
            "  $netTask = $asTask.Invoke($null, @($WinRtAction))\n"
            "  $netTask.Wait(-1) | Out-Null\n"
            "}\n"
            "[Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType=WindowsRuntime] | Out-Null\n"
            "[Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType=WindowsRuntime] | Out-Null\n"
            "[Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime] | Out-Null\n"
            "[Windows.Storage.StorageFolder, Windows.Storage, ContentType=WindowsRuntime] | Out-Null\n"
            "[Windows.Storage.CreationCollisionOption, Windows.Storage, ContentType=WindowsRuntime] | Out-Null\n"
            "[Windows.Storage.FileAccessMode, Windows.Storage, ContentType=WindowsRuntime] | Out-Null\n"
            "[Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType=WindowsRuntime] | Out-Null\n"
            f"$outDir = '{os.path.abspath(tmp_dir).replace(chr(39), chr(39)*2)}'\n"
            "$folder = Await ([Windows.Storage.StorageFolder]::GetFolderFromPathAsync($outDir)) ([Windows.Storage.StorageFolder])\n"
            f"$items = Get-Content -Raw -Encoding UTF8 '{manifest_path.replace(chr(39), chr(39)*2)}' | ConvertFrom-Json\n"
            "if ($items -isnot [System.Array]) { $items = @($items) }\n"
            "foreach ($it in $items) {\n"
            "  $sf = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($it.pdf)) ([Windows.Storage.StorageFile])\n"
            "  $doc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($sf)) ([Windows.Data.Pdf.PdfDocument])\n"
            "  $targets = @($it.p1, $it.p2)\n"
            "  for ($pi = 0; $pi -lt [Math]::Min(2, [int]$doc.PageCount); $pi++) {\n"
            "    $page = $doc.GetPage([uint32]$pi)\n"
            "    $outFile = Await ($folder.CreateFileAsync($targets[$pi], [Windows.Storage.CreationCollisionOption]::ReplaceExisting)) ([Windows.Storage.StorageFile])\n"
            "    $stream = Await ($outFile.OpenAsync([Windows.Storage.FileAccessMode]::ReadWrite)) ([Windows.Storage.Streams.IRandomAccessStream])\n"
            "    $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions\n"
            "    $opts.DestinationWidth = 2100\n"
            "    AwaitAction ($page.RenderToStreamAsync($stream, $opts))\n"
            "    $stream.Dispose()\n"
            "    $page.Dispose()\n"
            "  }\n"
            "}\n"
        )
        with open(ps_path, "w", encoding="utf-8-sig") as pf:
            pf.write(ps_code)

        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps_path],
            check=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        result_pairs = []
        for p1, p2 in expected_pairs:
            if not os.path.exists(p1) or os.path.getsize(p1) < 500:
                raise RuntimeError("WinRT PDF-рендеринг не создал PNG 1-й страницы.")
            result_pairs.append((p1, p2 if (os.path.exists(p2) and os.path.getsize(p2) > 500) else None))
        return result_pairs

    def _crop_form_image(self, img):
        """Аккуратно обрезает лишние белые поля вокруг бланка, оставляя небольшой равномерный отступ."""
        from PIL import Image, ImageChops

        rgb = img.convert("RGB")
        bg = Image.new("RGB", rgb.size, (255, 255, 255))
        diff = ImageChops.difference(rgb, bg)
        # Игнорируем едва заметный шум рендера (< 12)
        bbox = diff.point(lambda p: 255 if p > 12 else 0).getbbox()
        if not bbox:
            return rgb
        pad = 14
        x0 = max(0, bbox[0] - pad)
        y0 = max(0, bbox[1] - pad)
        x1 = min(rgb.width, bbox[2] + pad)
        y1 = min(rgb.height, bbox[3] + pad)
        return rgb.crop((x0, y0, x1, y1))

    def _compose_print_sheets_pil(self, page_pairs, tmp_dir, is_landscape=True, rotate_180=True):
        """
        Компонует готовые листы A4 (300 DPI) для лицевой и оборотной стороны:
        - В режиме «Поперёк (A5 — 2 на лист A4)»:
          * 1 путевой лист печатается строго на ПЕРВОЙ (левой) половине A4 (правая половина чистая под 2-й ПЛ).
          * 2+ путевых листа автоматически размещаются по 2 на один лист A4.
          * 1-я и 2-я страницы одного ПЛ имеют ТОЧНО СОВПАДАЮЩИЙ размер области печати и точное совмещение при перевороте!
        - В режиме «На весь лист (Книжная A4)»:
          * 1-я и 2-я страницы печатаются на весь лист A4 с одинаковым размером области печати.
        """
        from PIL import Image, ImageDraw

        resample = getattr(Image, "Resampling", Image).LANCZOS
        front_sheets = []
        back_sheets = []

        def _fit_size(src_w, src_h, box_w, box_h):
            scale = min(box_w / float(src_w), box_h / float(src_h))
            return (max(1, int(round(src_w * scale))), max(1, int(round(src_h * scale))))

        def _paste_in_slot(canvas, img_cropped, slot_x, slot_y, box_w, box_h, target_wh=None, rotate=False):
            if img_cropped is None:
                return None
            if target_wh is None:
                dw, dh = _fit_size(img_cropped.width, img_cropped.height, box_w, box_h)
            else:
                # Если пропорции 2-й страницы близки к 1-й (в пределах 25%), совпадаем 1-в-1 миллиметр в миллиметр
                src_ratio = img_cropped.width / float(img_cropped.height)
                tgt_ratio = target_wh[0] / float(target_wh[1])
                if 0.75 <= (src_ratio / tgt_ratio) <= 1.33:
                    dw, dh = target_wh
                else:
                    dw, dh = _fit_size(img_cropped.width, img_cropped.height, target_wh[0], target_wh[1])
            resized = img_cropped.resize((dw, dh), resample)
            if rotate:
                resized = resized.rotate(180)
            px = slot_x + (box_w - dw) // 2
            py = slot_y + (box_h - dh) // 2
            canvas.paste(resized, (px, py))
            return (dw, dh)

        if is_landscape:
            canvas_w, canvas_h = 3508, 2480  # A4 Landscape 300 DPI
            margin = 70
            center_gap = 70
            slot_w = (canvas_w - 2 * margin - center_gap) // 2  # 1649 px (A5)
            slot_h = canvas_h - 2 * margin                      # 2340 px (A5)
            left_slot = (margin, margin)
            right_slot = (margin + slot_w + center_gap, margin)

            for pair_idx in range(0, len(page_pairs), 2):
                wb_left = page_pairs[pair_idx]
                wb_right = page_pairs[pair_idx + 1] if (pair_idx + 1) < len(page_pairs) else None

                front_img = Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255))
                back_img = Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255))

                l_p1 = self._crop_form_image(Image.open(wb_left[0]))
                l_p2 = self._crop_form_image(Image.open(wb_left[1])) if wb_left[1] else None
                l_wh = _paste_in_slot(front_img, l_p1, left_slot[0], left_slot[1], slot_w, slot_h)

                r_wh = None
                r_p2 = None
                if wb_right is not None:
                    r_p1 = self._crop_form_image(Image.open(wb_right[0]))
                    r_p2 = self._crop_form_image(Image.open(wb_right[1])) if wb_right[1] else None
                    r_wh = _paste_in_slot(front_img, r_p1, right_slot[0], right_slot[1], slot_w, slot_h)

                # Небольшие метки линии реза сверху и снизу строго по центру A4 (чистая правая половина не пачкается)
                draw_f = ImageDraw.Draw(front_img)
                cx = canvas_w // 2
                draw_f.line([(cx, 18), (cx, 54)], fill=(170, 170, 170), width=2)
                draw_f.line([(cx, canvas_h - 54), (cx, canvas_h - 18)], fill=(170, 170, 170), width=2)

                has_back = (l_p2 is not None) or (r_p2 is not None)
                if has_back:
                    if rotate_180:
                        # Переворот листа по длинной стороне (книжный переворот портретного листа):
                        # Левая половина остаётся слева, но развёрнута на 180° — заголовок строго за заголовком!
                        if l_p2 is not None:
                            _paste_in_slot(
                                back_img, l_p2, left_slot[0], left_slot[1], slot_w, slot_h, target_wh=l_wh, rotate=True
                            )
                        if r_p2 is not None:
                            _paste_in_slot(
                                back_img, r_p2, right_slot[0], right_slot[1], slot_w, slot_h, target_wh=r_wh, rotate=True
                            )
                    else:
                        # Переворот листа через верхний/нижний край: половины меняются местами без поворота на 180°
                        if l_p2 is not None:
                            _paste_in_slot(
                                back_img, l_p2, right_slot[0], right_slot[1], slot_w, slot_h, target_wh=l_wh, rotate=False
                            )
                        if r_p2 is not None:
                            _paste_in_slot(
                                back_img, r_p2, left_slot[0], left_slot[1], slot_w, slot_h, target_wh=r_wh, rotate=False
                            )

                f_out = os.path.abspath(os.path.join(tmp_dir, f"sheet_front_{pair_idx // 2:03d}.png"))
                front_img.save(f_out, "PNG", dpi=(300, 300))
                front_sheets.append(f_out)

                if has_back:
                    b_out = os.path.abspath(os.path.join(tmp_dir, f"sheet_back_{pair_idx // 2:03d}.png"))
                    back_img.save(b_out, "PNG", dpi=(300, 300))
                    back_sheets.append(b_out)
        else:
            canvas_w, canvas_h = 2480, 3508  # A4 Portrait 300 DPI
            margin = 80
            slot_w = canvas_w - 2 * margin
            slot_h = canvas_h - 2 * margin

            for idx, (p1_path, p2_path) in enumerate(page_pairs):
                front_img = Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255))
                p1_crop = self._crop_form_image(Image.open(p1_path))
                p1_wh = _paste_in_slot(front_img, p1_crop, margin, margin, slot_w, slot_h)

                f_out = os.path.abspath(os.path.join(tmp_dir, f"sheet_front_{idx:03d}.png"))
                front_img.save(f_out, "PNG", dpi=(300, 300))
                front_sheets.append(f_out)

                if p2_path:
                    back_img = Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255))
                    p2_crop = self._crop_form_image(Image.open(p2_path))
                    _paste_in_slot(back_img, p2_crop, margin, margin, slot_w, slot_h, target_wh=p1_wh, rotate=False)
                    b_out = os.path.abspath(os.path.join(tmp_dir, f"sheet_back_{idx:03d}.png"))
                    back_img.save(b_out, "PNG", dpi=(300, 300))
                    back_sheets.append(b_out)

        return (front_sheets, back_sheets)

    def _print_image_batch_windows(self, sheet_pngs, tmp_dir, is_landscape=True, job_name="Waybills"):
        """Отправляет всю пачку готовых листов A4 в ОДНОМ стабильном задании .NET PrintDocument."""
        if not sheet_pngs:
            return
        list_path = os.path.join(tmp_dir, f"print_list_{job_name}.json")
        with open(list_path, "w", encoding="utf-8") as lf:
            json.dump([os.path.abspath(p) for p in sheet_pngs], lf, ensure_ascii=False)

        land_ps = "$true" if is_landscape else "$false"
        ps_path = os.path.join(tmp_dir, f"print_batch_{job_name}.ps1")
        ps_code = (
            "$ErrorActionPreference = 'Stop'\n"
            "Add-Type -AssemblyName System.Drawing\n"
            f"$files = Get-Content -Raw -Encoding UTF8 '{list_path.replace(chr(39), chr(39)*2)}' | ConvertFrom-Json\n"
            "if ($files -isnot [System.Array]) { $files = @($files) }\n"
            "$script:pageIdx = 0\n"
            "$pd = New-Object System.Drawing.Printing.PrintDocument\n"
            f"$pd.DocumentName = '{job_name}'\n"
            "$pd.PrintController = New-Object System.Drawing.Printing.StandardPrintController\n"
            f"$pd.DefaultPageSettings.Landscape = {land_ps}\n"
            "$pd.OriginAtMargins = $false\n"
            "$pd.DefaultPageSettings.Margins = New-Object System.Drawing.Printing.Margins(0, 0, 0, 0)\n"
            "if ($pd.PrinterSettings.CanDuplex) { $pd.PrinterSettings.Duplex = [System.Drawing.Printing.Duplex]::Simplex }\n"
            "$handler = [System.Drawing.Printing.PrintPageEventHandler]{\n"
            "  param($sender, $e)\n"
            "  $img = [System.Drawing.Image]::FromFile($files[$script:pageIdx])\n"
            "  try {\n"
            "    $bounds = $e.PageBounds\n"
            "    $ratioX = $bounds.Width / [double]$img.Width\n"
            "    $ratioY = $bounds.Height / [double]$img.Height\n"
            "    $ratio = [Math]::Min($ratioX, $ratioY)\n"
            "    $drawW = [int]([Math]::Round([double]$img.Width * $ratio))\n"
            "    $drawH = [int]([Math]::Round([double]$img.Height * $ratio))\n"
            "    $drawX = $bounds.Left + [int](($bounds.Width - $drawW) / 2)\n"
            "    $drawY = $bounds.Top + [int](($bounds.Height - $drawH) / 2)\n"
            "    $e.Graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic\n"
            "    $e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality\n"
            "    $e.Graphics.DrawImage($img, $drawX, $drawY, $drawW, $drawH)\n"
            "  } finally {\n"
            "    $img.Dispose()\n"
            "  }\n"
            "  $script:pageIdx++\n"
            "  $e.HasMorePages = ($script:pageIdx -lt $files.Count)\n"
            "}\n"
            "$pd.add_PrintPage($handler)\n"
            "$pd.Print()\n"
            "$pd.Dispose()\n"
        )
        with open(ps_path, "w", encoding="utf-8-sig") as pf:
            pf.write(ps_code)

        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps_path],
            check=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )

    def _print_excel_fallback_single_session(self, prepared_paths, sheet_index, is_landscape, tmp_dir):
        """Резервная пакетная печать через ОДИН сеанс Excel PrintOut, если WinRT PDF недоступен."""
        orient_num = 2 if is_landscape else 1
        list_path = os.path.join(tmp_dir, f"excel_fb_{sheet_index}.json")
        with open(list_path, "w", encoding="utf-8") as lf:
            json.dump([os.path.abspath(p) for p in prepared_paths], lf, ensure_ascii=False)
        ps_path = os.path.join(tmp_dir, f"excel_fb_{sheet_index}.ps1")
        ps_code = (
            "$ErrorActionPreference = 'Stop'\n"
            f"$files = Get-Content -Raw -Encoding UTF8 '{list_path.replace(chr(39), chr(39)*2)}' | ConvertFrom-Json\n"
            "if ($files -isnot [System.Array]) { $files = @($files) }\n"
            "$excel = New-Object -ComObject Excel.Application\n"
            "$excel.Visible = $false; $excel.DisplayAlerts = $false\n"
            "try {\n"
            "  foreach ($f in $files) {\n"
            "    $wb = $excel.Workbooks.Open($f, $null, $true)\n"
            "    try {\n"
            f"      if ($wb.Worksheets.Count -ge {sheet_index}) {{\n"
            f"        $ws = $wb.Worksheets.Item({sheet_index})\n"
            f"        $ws.PageSetup.Orientation = {orient_num}\n"
            "        $ws.PageSetup.Zoom = $false\n"
            "        $ws.PageSetup.FitToPagesWide = 1\n"
            "        $ws.PageSetup.FitToPagesTall = 1\n"
            "        $ws.PageSetup.CenterHorizontally = $true\n"
            "        $ws.PrintOut()\n"
            "      }\n"
            "    } finally { $wb.Close($false) }\n"
            "  }\n"
            "} finally { $excel.Quit() }\n"
        )
        with open(ps_path, "w", encoding="utf-8-sig") as pf:
            pf.write(ps_code)
        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps_path],
            check=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )

    def print_excel_files(self, file_paths, duplex=True):
        orientation_label = (
            self.print_orientation_var.get()
            if hasattr(self, "print_orientation_var")
            else "Поперёк (A5 — 2 на лист A4)"
        )
        is_landscape = "Поперёк" in orientation_label or "Альбом" in orientation_label or "A5" in orientation_label
        tmp_dir = tempfile.mkdtemp(prefix="waybill_print_")
        try:
            prepared_paths = [
                self._prepare_print_copy(p, tmp_dir, idx=i, for_pdf=True, is_landscape=is_landscape)
                for i, p in enumerate(file_paths)
            ]

            # Основной стабильный конвейер Windows 11: Excel -> PDF -> WinRT PNG -> Pillow A5 2-up -> .NET PrintDocument
            page_pairs = None
            try:
                pdf_paths = self._export_workbooks_to_pdfs_windows(prepared_paths, tmp_dir)
                page_pairs = self._render_pdfs_to_pngs_windows(pdf_paths, tmp_dir)
            except Exception:
                page_pairs = None

            if page_pairs:
                # ШАГ 1: Компонуем и печатаем 1-е страницы (в A5: 1 ПЛ строго на первой/левой половине A4, 2+ ПЛ — по 2 на лист)
                front_sheets, _ = self._compose_print_sheets_pil(
                    page_pairs, tmp_dir, is_landscape=is_landscape, rotate_180=self.print_rotate_180
                )
                self._print_image_batch_windows(
                    front_sheets, tmp_dir, is_landscape=is_landscape, job_name="Waybills_Page1"
                )

                # ШАГ 2: Окно паузы для переворота листов (и для одиночной, и для массовой печати!)
                if duplex and any(p2 is not None for _, p2 in page_pairs):
                    dlg = PrintFlipDialog(
                        self,
                        count=len(page_pairs),
                        orientation_label=orientation_label,
                        is_landscape=is_landscape,
                        rotate_180_default=self.print_rotate_180,
                        sheets_count=len(front_sheets),
                    )
                    self.wait_window(dlg)
                    if is_landscape:
                        self.print_rotate_180 = bool(dlg.var_rotate_180.get())
                        self.save_memory()
                    if not dlg.confirmed:
                        return

                    _, back_sheets = self._compose_print_sheets_pil(
                        page_pairs, tmp_dir, is_landscape=is_landscape, rotate_180=self.print_rotate_180
                    )
                    self._print_image_batch_windows(
                        back_sheets, tmp_dir, is_landscape=is_landscape, job_name="Waybills_Page2"
                    )
                return

            # Резервный режим, если WinRT PDF недоступен
            fb_paths = [
                self._prepare_print_copy(p, tmp_dir, idx=i, for_pdf=False, is_landscape=is_landscape)
                for i, p in enumerate(file_paths)
            ]
            self._print_excel_fallback_single_session(fb_paths, 1, is_landscape, tmp_dir)
            if duplex:
                dlg = PrintFlipDialog(
                    self,
                    count=len(fb_paths),
                    orientation_label=orientation_label,
                    is_landscape=is_landscape,
                    rotate_180_default=self.print_rotate_180,
                    sheets_count=len(fb_paths),
                )
                self.wait_window(dlg)
                if not dlg.confirmed:
                    return
                self._print_excel_fallback_single_session(fb_paths, 2, is_landscape, tmp_dir)
        except Exception as e:
            try:
                if hasattr(os, "startfile"):
                    for p in file_paths:
                        os.startfile(os.path.abspath(p), "print")
                    return
            except Exception:
                pass
            messagebox.showerror(
                "Ошибка печати", f"Не удалось выполнить печать средствами Windows:\n{str(e)}"
            )
        finally:
            try:
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception:
                pass

    def auto_predict_driver_data(self, selected_drv):
        drv_row = next(
            (r for r in self.db.get_list("drivers") if r[1] == selected_drv), None
        )
        if drv_row and drv_row[3]:
            try:
                exp_date = datetime.strptime(str(drv_row[3]).strip(), "%d.%m.%Y")
                days_left = (exp_date - datetime.now()).days
                if days_left < 0:
                    messagebox.showwarning(
                        "Системный контроль",
                        f"Внимание! У водителя {selected_drv} истек срок действия водительского удостоверения ({drv_row[3]}). Выпуск на рейс не допускается.",
                    )
                elif days_left <= 30:
                    messagebox.showwarning(
                        "Системный контроль",
                        f"У водителя {selected_drv} срок действия водительского удостоверения истекает через {days_left} дн.",
                    )
            except Exception:
                pass

    def reset_crud_buttons(self):
        self.editing_crud_id = {
            "drivers": None,
            "vehicles": None,
            "orgs": None,
            "clients": None,
            "staff": None,
            "work_types": None,
        }
        self._editing_old_staff_fio = None
        if hasattr(self, "btn_d_add"):
            self.btn_d_add.configure(text="Добавить")
        if hasattr(self, "btn_v_add"):
            self.btn_v_add.configure(text="Добавить")
        if hasattr(self, "btn_o_add"):
            self.btn_o_add.configure(text="Добавить")
        if hasattr(self, "btn_c_add"):
            self.btn_c_add.configure(text="Добавить")
        if hasattr(self, "btn_s_add"):
            self.btn_s_add.configure(text="Добавить")
        if hasattr(self, "btn_wt_add"):
            self.btn_wt_add.configure(text="Добавить вид работ")
        for en_name in ("en_s_fio", "en_org", "en_cli", "en_wt_name", "en_wt_norm", "en_wt_unit"):
            if hasattr(self, en_name):
                getattr(self, en_name).delete(0, "end")

    def load_crud_edit(self, table, tree):
        sel = tree.selection()
        if not sel:
            return
        vals = tree.item(sel[0], "values")
        item_id = vals[0]
        self.editing_crud_id[table] = item_id
        if table == "drivers":
            self.en_d_fio.delete(0, "end")
            self.en_d_fio.insert(0, vals[1])
            self.en_d_cat.delete(0, "end")
            self.en_d_cat.insert(0, vals[2])
            self.en_d_exp.delete(0, "end")
            self.en_d_exp.insert(0, vals[3])
            self.en_d_tab.delete(0, "end")
            self.en_d_tab.insert(0, vals[4])
            self.en_d_ser.delete(0, "end")
            self.en_d_ser.insert(0, vals[5])
            self.en_d_num.delete(0, "end")
            self.en_d_num.insert(0, vals[6])
            self.btn_d_add.configure(text="Сохранить")
        elif table == "vehicles":
            self.en_v_mark.delete(0, "end")
            self.en_v_mark.insert(0, vals[1])
            self.en_v_reg.delete(0, "end")
            self.en_v_reg.insert(0, vals[2])
            self.en_v_gar.delete(0, "end")
            self.en_v_gar.insert(0, vals[3])
            self.en_v_f.delete(0, "end")
            self.en_v_f.insert(0, vals[4])
            self.en_v_fn.delete(0, "end")
            self.en_v_fn.insert(0, vals[5])
            self.en_v_g.delete(0, "end")
            self.en_v_g.insert(0, vals[6])
            self.en_v_gn.delete(0, "end")
            self.en_v_gn.insert(0, vals[7])
            self.btn_v_add.configure(text="Сохранить")
        elif table == "orgs":
            self.en_org.delete(0, "end")
            self.en_org.insert(0, vals[1])
            self.btn_o_add.configure(text="Сохранить")
        elif table == "clients":
            self.en_cli.delete(0, "end")
            self.en_cli.insert(0, vals[1])
            self.btn_c_add.configure(text="Сохранить")
        elif table == "staff":
            self._editing_old_staff_fio = str(vals[1])
            self.en_s_fio.delete(0, "end")
            self.en_s_fio.insert(0, vals[1])
            if len(vals) > 2:
                self.cb_s_role.set(str(vals[2]))
            self.btn_s_add.configure(text="Сохранить")
        elif table == "work_types":
            self.en_wt_name.delete(0, "end")
            self.en_wt_name.insert(0, vals[1])
            self.en_wt_norm.delete(0, "end")
            self.en_wt_norm.insert(0, str(vals[2]) if len(vals) > 2 else "5.2")
            self.en_wt_unit.delete(0, "end")
            self.en_wt_unit.insert(0, str(vals[3]) if len(vals) > 3 else "м³")
            self.btn_wt_add.configure(text="Сохранить вид работ")

    def add_drv(self):
        fio = self.en_d_fio.get().strip()
        if not fio:
            return
        cat = self.en_d_cat.get().strip()
        exp = self.en_d_exp.get().strip()
        tab = self.en_d_tab.get().strip()
        ser = self.en_d_ser.get().strip()
        num = self.en_d_num.get().strip()
        if self.editing_crud_id["drivers"]:
            self.db.update_driver(
                self.editing_crud_id["drivers"], fio, cat, exp, tab, ser, num
            )
            self.reset_crud_buttons()
        else:
            self.db.add_driver(fio, cat, exp, tab, ser, num)
        for en in (
            self.en_d_fio,
            self.en_d_cat,
            self.en_d_exp,
            self.en_d_tab,
            self.en_d_ser,
            self.en_d_num,
        ):
            en.delete(0, "end")
        self.refresh_all_tables()

    def add_veh(self):
        mark = self.en_v_mark.get().strip()
        if not mark:
            return
        reg = self.en_v_reg.get().strip()
        gar = self.en_v_gar.get().strip()
        f_b = self.en_v_f.get().strip()
        norm = safe_float(self.en_v_fn.get(), 0.0)
        g_b = self.en_v_g.get().strip()
        norm_g = safe_float(self.en_v_gn.get(), 0.0)
        if self.active_mode == "tractors":
            eq_t = g_b or "Трактор МТЗ-82"
            if self.editing_crud_id["vehicles"]:
                self.db_tractors.update_vehicle(
                    self.editing_crud_id["vehicles"],
                    mark,
                    reg,
                    gar,
                    eq_t,
                    f_b or "ДТ (дизельное)",
                    norm or 16.0,
                    norm_g or 5.0,
                )
                self.reset_crud_buttons()
            else:
                self.db_tractors.add_vehicle(
                    mark,
                    reg,
                    gar,
                    eq_t,
                    f_b or "ДТ (дизельное)",
                    norm or 16.0,
                    norm_g or 5.0,
                )
        else:
            if self.editing_crud_id["vehicles"]:
                self.db.update_vehicle(
                    self.editing_crud_id["vehicles"],
                    mark,
                    reg,
                    gar,
                    f_b,
                    norm,
                    g_b,
                    norm_g,
                )
                self.reset_crud_buttons()
            else:
                self.db.query(
                    "INSERT INTO vehicles (mark, reg, gar, f_brand, norm, f_brand_gas, norm_gas) VALUES (?,?,?,?,?,?,?)",
                    (mark, reg, gar, f_b, norm, g_b, norm_g),
                )
        for en in (
            self.en_v_mark,
            self.en_v_reg,
            self.en_v_gar,
            self.en_v_f,
            self.en_v_fn,
            self.en_v_g,
            self.en_v_gn,
        ):
            en.delete(0, "end")
        self.refresh_all_tables()
        self.load_vehicle_data()

    def _apply_selected_wt_norm(self):
        sel = self.tr_wt.selection() if hasattr(self, "tr_wt") else ()
        if sel:
            vals = self.tr_wt.item(sel[0], "values")
            self.open_work_norm_mass_update(
                preselect_work=str(vals[1]),
                preselect_norm=safe_float(vals[2], 5.2),
            )
        else:
            self.open_work_norm_mass_update()

    def add_wt(self):
        name = self.en_wt_name.get().strip()
        if not name:
            return
        norm_mh = safe_float(self.en_wt_norm.get(), 5.2)
        unit = self.en_wt_unit.get().strip() or "м³"
        norm_changed = False
        if self.editing_crud_id.get("work_types"):
            wt_id = self.editing_crud_id["work_types"]
            old_rows = self.db_tractors.query("SELECT norm_mh FROM work_types WHERE id=?", (wt_id,))
            old_norm = float(old_rows[0][0] or 0.0) if old_rows else norm_mh
            norm_changed = abs(old_norm - norm_mh) > 1e-4
            self.db_tractors.update_work_type(wt_id, name, norm_mh, unit)
            self.reset_crud_buttons()
        else:
            self.db_tractors.add_work_type(name, norm_mh, unit)
            self.en_wt_name.delete(0, "end")
            self.en_wt_norm.delete(0, "end")
            self.en_wt_unit.delete(0, "end")
        self.refresh_all_tables()
        self.load_vehicle_data()
        if norm_changed:
            if messagebox.askyesno(
                "Пересчёт путевых листов",
                f"Норма для вида работ «{name}» изменена на {norm_mh:g} л/м·ч.\n\n"
                "Применить новую норму к уже сохранённым путевым листам (всем или за период)?",
            ):
                self.open_work_norm_mass_update(preselect_work=name, preselect_norm=norm_mh)

    def del_wt_item(self):
        sel = self.tr_wt.selection()
        if not sel:
            return
        for item in sel:
            item_id = self.tr_wt.item(item, "values")[0]
            self.db_tractors.query("DELETE FROM work_types WHERE id=?", (item_id,))
        self.reset_crud_buttons()
        self.refresh_all_tables()

    def add_simple(self, table, entry_widget):
        val = entry_widget.get().strip()
        if not val:
            return
        edit_id = self.editing_crud_id.get(table)
        if edit_id:
            if table == "orgs":
                self.db.update_org(edit_id, val)
            elif table == "clients":
                self.db.update_client(edit_id, val)
            self.reset_crud_buttons()
        else:
            self.db.query(f"INSERT OR IGNORE INTO {table} (name) VALUES (?)", (val,))
            entry_widget.delete(0, "end")
        self.refresh_all_tables()

    def add_staff(self):
        fio = self.en_s_fio.get().strip()
        role = self.cb_s_role.get().strip() or "Механик"
        if not fio:
            return
        old_fio = getattr(self, "_editing_old_staff_fio", None)
        if self.editing_crud_id.get("staff"):
            self.db.update_staff(self.editing_crud_id["staff"], fio, role)
            if old_fio:
                if hasattr(self, "cb_mech") and self.cb_mech.get().strip() == old_fio:
                    self.cb_mech.set(fio)
                if hasattr(self, "cb_disp") and self.cb_disp.get().strip() == old_fio:
                    self.cb_disp.set(fio)
            self.reset_crud_buttons()
        else:
            self.db.add_staff(fio, role)
            self.en_s_fio.delete(0, "end")
        self.refresh_all_tables()
        self.save_memory()

    def refresh_all_tables(self):
        for tree, table in (
            (self.tr_d, "drivers"),
            (self.tr_v, "vehicles"),
            (self.tr_org, "orgs"),
            (self.tr_cli, "clients"),
            (self.tr_stf, "staff"),
        ):
            if hasattr(self, "tr_d"):
                for item in tree.get_children():
                    tree.delete(item)
                for row in self.db.get_list(table):
                    tree.insert("", "end", values=row)

        if hasattr(self, "tr_wt"):
            for item in self.tr_wt.get_children():
                self.tr_wt.delete(item)
            for row in self.db_tractors.get_list("work_types"):
                self.tr_wt.insert("", "end", values=row)

        unique_locs = self.get_unique_locs()
        if hasattr(self, "cb_drv"):
            self.cb_drv.configure(values=self.get_names("drivers") or [""])
        if hasattr(self, "cb_veh"):
            self.cb_veh.configure(values=self.get_names("vehicles") or [""])
        if hasattr(self, "cb_org"):
            self.cb_org.configure(values=self.get_names("orgs") or [""])
        if hasattr(self, "cb_client"):
            self.cb_client.configure(values=self.get_names("clients") or [""])
        if hasattr(self, "cb_mech"):
            mechs = self._get_staff_by_role("механ")
            self.cb_mech.configure(values=mechs)
            cur_m = self.cb_mech.get().strip()
            if mechs and mechs[0] and (not cur_m or cur_m not in mechs):
                self.cb_mech.set(mechs[0])
        if hasattr(self, "cb_disp"):
            disps = self._get_staff_by_role("диспетч")
            self.cb_disp.configure(values=disps)
            cur_d = self.cb_disp.get().strip()
            if disps and disps[0] and (not cur_d or cur_d not in disps):
                self.cb_disp.set(disps[0])
        if hasattr(self, "cb_pickup"):
            self.cb_pickup.configure(values=unique_locs or [""])
        if hasattr(self, "cb_route_add"):
            self.cb_route_add.configure(values=unique_locs or [""])
        if hasattr(self, "cb_work_type"):
            wt_vals = [r[1] for r in self.db_tractors.get_list("work_types")]
            if wt_vals:
                self.cb_work_type.configure(values=wt_vals)

    def del_item(self, table, tree):
        sel = tree.selection()
        if not sel:
            return
        for item in sel:
            item_id = tree.item(item, "values")[0]
            self.db.query(f"DELETE FROM {table} WHERE id=?", (item_id,))
        self.reset_crud_buttons()
        self.refresh_all_tables()


if __name__ == "__main__":
    app = WaybillApp()
    app.mainloop()
