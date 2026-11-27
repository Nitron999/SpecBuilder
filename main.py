import ctypes
import json
import os
import re
import shutil
import subprocess
import threading
import time
from datetime import datetime
from difflib import SequenceMatcher
from tkinter import messagebox

import customtkinter as ctk
import openpyxl
import pythoncom
import win32api
import win32clipboard
import win32com.client
import win32con
import xlrd
from num2words import num2words
from pywinauto import Desktop, mouse


# =============================================================================
# НАСТРОЙКИ ПО УМОЛЧАНИЮ
# =============================================================================
CFG = {
    "company_name": 'ООО "ПЕЧАТНЫЕ ПЛАТЫ"',
    "employee_name": "Мамонтова Людмила Давыдовна",
    "1c_exe": r"C:\Program Files\1cv8\common\1cestart.exe",
    "1c_base_args": 'ENTERPRISE /S"DB01\\pp_buh_new" /N"МамонтоваЛД" /P"mamon123"',
    "specs_root": r"\\Fs01\obmen\ВП\Спецификации 2019 Новая форма",
    "calc_root": r"\\Fs01\obmen\РАСЧЕТЫ",
    "folder_mapping": {"Электроприбор": "Elektropribor_K"},
}

SETTINGS_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "SpecBuilder")
SETTINGS_FILE = os.path.join(SETTINGS_DIR, "settings.json")


# =============================================================================
# 7 ТЕМ ОФОРМЛЕНИЯ
# =============================================================================
THEMES = {
    "Тёмная ночь": {
        "mode": "dark", "bg_app": "#0f1216", "bg_sidebar": "#13171d", "bg_card": "#181d24",
        "bg_input": "#10141a", "border": "#262d38", "accent": "#10b981", "accent_hover": "#059669",
        "accent_text": "#052e16", "sel_bg": "#163529", "sel_text": "#34d399", "hover": "#1f2937",
        "title_accent1": "#f59e0b", "title_accent2": "#38bdf8", "title_accent3": "#10b981",
        "text_main": "#f1f5f9", "text_muted": "#94a3b8", "badge_bg": "#1d2430",
        "toggle_off_bg": "#1e3a5f", "toggle_off_hover": "#254b7a", "toggle_off_border": "#38bdf8",
        "toggle_off_text": "#e0f2fe", "toggle_on_bg": "#d97706", "toggle_on_hover": "#b45309",
        "toggle_on_border": "#fbbf24", "toggle_on_text": "#ffffff",
    },
    "Светлая классика": {
        "mode": "light", "bg_app": "#f1f5f9", "bg_sidebar": "#ffffff", "bg_card": "#ffffff",
        "bg_input": "#f8fafc", "border": "#cbd5e1", "accent": "#16a34a", "accent_hover": "#15803d",
        "accent_text": "#ffffff", "sel_bg": "#dcfce7", "sel_text": "#166534", "hover": "#e2e8f0",
        "title_accent1": "#d97706", "title_accent2": "#0284c7", "title_accent3": "#16a34a",
        "text_main": "#0f172a", "text_muted": "#475569", "badge_bg": "#f1f5f9",
        "toggle_off_bg": "#e0f2fe", "toggle_off_hover": "#bae6fd", "toggle_off_border": "#0284c7",
        "toggle_off_text": "#0c4a6e", "toggle_on_bg": "#d97706", "toggle_on_hover": "#b45309",
        "toggle_on_border": "#b45309", "toggle_on_text": "#ffffff",
    },
    "Изумруд СДЭК": {
        "mode": "dark", "bg_app": "#0b1410", "bg_sidebar": "#0f1c16", "bg_card": "#13241c",
        "bg_input": "#0c1712", "border": "#1f3b2e", "accent": "#1ab248", "accent_hover": "#15963c",
        "accent_text": "#04210c", "sel_bg": "#163d25", "sel_text": "#4ade80", "hover": "#1a3327",
        "title_accent1": "#4ade80", "title_accent2": "#34d399", "title_accent3": "#1ab248",
        "text_main": "#ecfdf5", "text_muted": "#86efac", "badge_bg": "#163024",
        "toggle_off_bg": "#1e3a5f", "toggle_off_hover": "#254b7a", "toggle_off_border": "#38bdf8",
        "toggle_off_text": "#e0f2fe", "toggle_on_bg": "#d97706", "toggle_on_hover": "#b45309",
        "toggle_on_border": "#fbbf24", "toggle_on_text": "#ffffff",
    },
    "Глубокий океан": {
        "mode": "dark", "bg_app": "#0a1320", "bg_sidebar": "#0e1a2b", "bg_card": "#132238",
        "bg_input": "#0b1626", "border": "#1e3554", "accent": "#0ea5e9", "accent_hover": "#0284c7",
        "accent_text": "#082f49", "sel_bg": "#13344f", "sel_text": "#38bdf8", "hover": "#1b2e4a",
        "title_accent1": "#38bdf8", "title_accent2": "#818cf8", "title_accent3": "#22d3ee",
        "text_main": "#f0f9ff", "text_muted": "#94a3b8", "badge_bg": "#172942",
        "toggle_off_bg": "#312e81", "toggle_off_hover": "#3730a3", "toggle_off_border": "#818cf8",
        "toggle_off_text": "#e0e7ff", "toggle_on_bg": "#d97706", "toggle_on_hover": "#b45309",
        "toggle_on_border": "#fbbf24", "toggle_on_text": "#ffffff",
    },
    "Графит и золото": {
        "mode": "dark", "bg_app": "#17171a", "bg_sidebar": "#1e1e22", "bg_card": "#242429",
        "bg_input": "#19191d", "border": "#36363d", "accent": "#eab308", "accent_hover": "#ca8a04",
        "accent_text": "#1c1917", "sel_bg": "#383018", "sel_text": "#facc15", "hover": "#2d2d34",
        "title_accent1": "#facc15", "title_accent2": "#fbbf24", "title_accent3": "#eab308",
        "text_main": "#fafaf9", "text_muted": "#a8a29e", "badge_bg": "#2b2b31",
        "toggle_off_bg": "#1e3a5f", "toggle_off_hover": "#254b7a", "toggle_off_border": "#38bdf8",
        "toggle_off_text": "#e0f2fe", "toggle_on_bg": "#16a34a", "toggle_on_hover": "#15803d",
        "toggle_on_border": "#4ade80", "toggle_on_text": "#ffffff",
    },
    "Чёрное золото": {
        "mode": "dark", "bg_app": "#09090b", "bg_sidebar": "#111114", "bg_card": "#16161a",
        "bg_input": "#0d0d10", "border": "#2e2718", "accent": "#d4af37", "accent_hover": "#b8952a",
        "accent_text": "#09090b", "sel_bg": "#2b2412", "sel_text": "#f3d267", "hover": "#221e15",
        "title_accent1": "#f3d267", "title_accent2": "#d4af37", "title_accent3": "#eab308",
        "text_main": "#fefce8", "text_muted": "#a1a1aa", "badge_bg": "#1d1a12",
        "toggle_off_bg": "#1e3a5f", "toggle_off_hover": "#254b7a", "toggle_off_border": "#38bdf8",
        "toggle_off_text": "#e0f2fe", "toggle_on_bg": "#16a34a", "toggle_on_hover": "#15803d",
        "toggle_on_border": "#4ade80", "toggle_on_text": "#ffffff",
    },
    "Тёплый песок": {
        "mode": "light", "bg_app": "#f5efe6", "bg_sidebar": "#fbf8f3", "bg_card": "#fffdf9",
        "bg_input": "#f7f1e8", "border": "#dfd3c3", "accent": "#b45309", "accent_hover": "#92400e",
        "accent_text": "#ffffff", "sel_bg": "#f3e5d0", "sel_text": "#78350f", "hover": "#ede2d2",
        "title_accent1": "#b45309", "title_accent2": "#0f766e", "title_accent3": "#92400e",
        "text_main": "#292524", "text_muted": "#78716c", "badge_bg": "#f2e8da",
        "toggle_off_bg": "#e0f2fe", "toggle_off_hover": "#bae6fd", "toggle_off_border": "#0284c7",
        "toggle_off_text": "#0c4a6e", "toggle_on_bg": "#15803d", "toggle_on_hover": "#166534",
        "toggle_on_border": "#16a34a", "toggle_on_text": "#ffffff",
    },
}


def load_user_settings() -> dict:
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
    except Exception:
        pass
    return {"theme": "Тёмная ночь", "folder_mapping": dict(CFG["folder_mapping"])}


def save_user_settings(data: dict):
    try:
        os.makedirs(SETTINGS_DIR, exist_ok=True)
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# =============================================================================
# КЛАВИАТУРА И БУФЕР ОБМЕНА
# =============================================================================
class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort), ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong), ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class _INPUT_UNION(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT), ("pad", ctypes.c_ubyte * 28)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong), ("union", _INPUT_UNION)]


def type_unicode_text(text: str):
    extra = ctypes.c_ulong(0)
    for ch in text:
        code = ord(ch)
        inp_down = INPUT(type=1)
        inp_down.union.ki = KEYBDINPUT(0, code, 0x0004, 0, ctypes.pointer(extra))
        inp_up = INPUT(type=1)
        inp_up.union.ki = KEYBDINPUT(0, code, 0x0004 | 0x0002, 0, ctypes.pointer(extra))
        ctypes.windll.user32.SendInput(1, ctypes.byref(inp_down), ctypes.sizeof(INPUT))
        ctypes.windll.user32.SendInput(1, ctypes.byref(inp_up), ctypes.sizeof(INPUT))
        time.sleep(0.03)


def press_ctrl_key(vk_code: int):
    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    time.sleep(0.05)
    win32api.keybd_event(vk_code, 0, 0, 0)
    time.sleep(0.05)
    win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.05)
    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)


def press_ctrl_f4():
    press_ctrl_key(win32con.VK_F4)


def press_vk(vk_code: int):
    win32api.keybd_event(vk_code, 0, 0, 0)
    time.sleep(0.05)
    win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)


def get_clipboard_text() -> str:
    for _ in range(5):
        try:
            win32clipboard.OpenClipboard()
            try:
                if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                    return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT) or ""
                return ""
            finally:
                win32clipboard.CloseClipboard()
        except Exception:
            time.sleep(0.15)
    return ""


def clear_clipboard():
    for _ in range(5):
        try:
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                return
            finally:
                win32clipboard.CloseClipboard()
        except Exception:
            time.sleep(0.15)


# =============================================================================
# ТРАНСЛИТЕРАЦИЯ И АВТОПОДБОР ПАПКИ В «РАСЧЕТЫ»
# =============================================================================
TRANSLIT_MAP = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e",
    "ж": "zh", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
    "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "h", "ц": "c", "ч": "ch", "ш": "sh", "щ": "sch",
    "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya",
}


def translit_ru(text: str) -> str:
    s = text.lower().replace("кс", "ks")
    return "".join(TRANSLIT_MAP.get(ch, ch) for ch in s)


def normalize_latin(text: str) -> str:
    s = translit_ru(text)
    s = s.replace("x", "ks").replace("w", "v").replace("q", "k")
    s = s.replace("ph", "f").replace("kh", "h").replace("ts", "c").replace("tz", "c")
    s = s.replace("shch", "sch").replace("yu", "u").replace("ju", "u").replace("ya", "a").replace("ja", "a")
    s = s.replace("yo", "e").replace("je", "e").replace("ye", "e")
    s = s.replace("j", "i").replace("y", "i")
    return re.sub(r"[^a-z0-9]+", "", s)


def normalize_raw(text: str) -> str:
    return re.sub(r"[^a-zа-я0-9]+", "", text.lower().replace("ё", "е"))


def score_folder_match(client_name: str, folder_name: str) -> float:
    c_raw, f_raw = normalize_raw(client_name), normalize_raw(folder_name)
    if not c_raw or not f_raw:
        return 0.0
    if c_raw == f_raw:
        return 1.0
    c_lat, f_lat = normalize_latin(client_name), normalize_latin(folder_name)
    if not c_lat or not f_lat:
        return 0.0
    if c_lat == f_lat:
        return 0.99
    if f_lat.startswith(c_lat) and len(c_lat) >= 3:
        return 0.90 + 0.08 * (len(c_lat) / len(f_lat))
    if c_lat.startswith(f_lat) and len(f_lat) >= 3:
        return 0.85 + 0.08 * (len(f_lat) / len(c_lat))
    if c_lat in f_lat and len(c_lat) >= 4:
        return 0.82 + 0.06 * (len(c_lat) / len(f_lat))
    if f_lat in c_lat and len(f_lat) >= 4:
        return 0.80 + 0.06 * (len(f_lat) / len(c_lat))
    return SequenceMatcher(None, c_lat, f_lat).ratio()


def guess_calc_folder(client_name: str, calc_folders: list, mapping: dict) -> str:
    mapped = mapping.get(client_name) or CFG.get("folder_mapping", {}).get(client_name)
    if mapped and mapped in calc_folders:
        return mapped
    best_folder, best_score = "(не выбрана)", 0.0
    for folder in calc_folders:
        sc = score_folder_match(client_name, folder)
        if sc > best_score:
            best_score, best_folder = sc, folder
    return best_folder if best_score >= 0.72 else "(не выбрана)"


# =============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =============================================================================
def resolve_1c_exe(configured_path: str) -> str:
    for path in (
        configured_path,
        r"C:\Program Files\1cv8\common\1cestart.exe",
        r"C:\Program Files (x86)\1cv8\common\1cestart.exe",
    ):
        if path and os.path.exists(path):
            return path
    raise FileNotFoundError("Не найден 1cestart.exe на этом компьютере.")


def clean_invoice_number(raw: str) -> str:
    return re.sub(r"^(?:сч[её]т|сч\.?|№)\s*", "", raw.strip(), flags=re.IGNORECASE).strip()


def format_1c_invoice_search(invoice_num: str) -> str:
    s = str(invoice_num).strip()
    return f"0000-{s.zfill(6)}" if re.fullmatch(r"\d{1,6}", s) else s


def parse_number(val) -> float:
    if isinstance(val, (int, float)):
        return float(val)
    if val is None:
        return 0.0
    s = str(val).replace("\xa0", "").replace(" ", "").replace("руб", "")
    if s.count(".") > 1:
        s = s.replace(".", "", s.count(".") - 1)
    s = s.replace(",", ".")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group(0)) if m else 0.0


def looks_like_number(val: str) -> bool:
    s = str(val).replace("\xa0", "").replace(" ", "").replace(",", ".")
    return bool(re.fullmatch(r"-?\d+(?:\.\d+)?", s))


def format_rub_text(amount: float, with_suffix: bool = True) -> str:
    formatted = f"{amount:,.2f}".replace(",", "\xa0").replace(".", ",")
    return f"{formatted}\xa0руб" if with_suffix else formatted


def decline(n: int, form1: str, form2: str, form5: str) -> str:
    n = abs(n) % 100
    if 11 <= n <= 19:
        return form5
    n %= 10
    if n == 1:
        return form1
    if 2 <= n <= 4:
        return form2
    return form5


def amount_in_words(amount: float) -> str:
    rub = int(amount)
    kop = int(round((amount - rub) * 100))
    if kop == 100:
        rub += 1
        kop = 0
    rub_words = num2words(rub, lang="ru").capitalize()
    return f"{rub_words} {decline(rub, 'рубль', 'рубля', 'рублей')} {kop:02d} {decline(kop, 'копейка', 'копейки', 'копеек')}"


def split_board_and_ps(raw_name: str):
    s = raw_name.strip()
    m = re.search(r"\bps\s*(\d+)\s+(.+)$", s, flags=re.IGNORECASE)
    if m:
        return m.group(2).strip(), m.group(1).strip()
    m2 = re.search(r"\bps\s*(\d+)\b(.*)$", s, flags=re.IGNORECASE)
    if m2:
        rest = m2.group(2).strip(" -–—")
        return (rest if rest else s), m2.group(1).strip()
    return s, ""


def find_latest_spec(spec_folder: str):
    patterns = [
        re.compile(r"сп\.?\s*(?:№\s*)?(\d+)", re.IGNORECASE),
        re.compile(r"спец(?:ификация)?\.?\s*(?:№\s*)?(\d+)", re.IGNORECASE),
    ]
    matched, all_excel = [], []
    for fname in os.listdir(spec_folder):
        if fname.startswith("~$") or not fname.lower().endswith((".xls", ".xlsx")):
            continue
        full_path = os.path.join(spec_folder, fname)
        if not os.path.isfile(full_path):
            continue
        mtime = os.path.getmtime(full_path)
        year = datetime.fromtimestamp(mtime).year
        all_excel.append((mtime, full_path))
        for pat in patterns:
            m = pat.search(fname)
            if m:
                matched.append((year, int(m.group(1)), mtime, full_path))
                break
    if matched:
        matched.sort(key=lambda x: (x[0], x[1], x[2]), reverse=True)
        return matched[0][1], matched[0][3]
    if all_excel:
        all_excel.sort(key=lambda x: x[0], reverse=True)
        return 0, all_excel[0][1]
    raise FileNotFoundError(f"В папке {spec_folder} нет файлов Excel (.xls / .xlsx).")


# =============================================================================
# РАБОТА С 1С (ПОИСК, СНЯТИЕ ПЛОЩАДИ ЧЕРЕЗ «ВЫВЕСТИ СПИСОК» + ИГК, ПЕЧАТЬ 7.7)
# =============================================================================
def find_1c_window():
    for w in Desktop(backend="uia").windows():
        title = w.window_text()
        if "БП_рабочая база" in title or ("1С:Предприятие" in title and "Запуск" not in title):
            return w
    return None


def get_el_text_all(el) -> str:
    parts = []
    try:
        t = el.window_text()
        if t:
            parts.append(str(t))
    except Exception:
        pass
    try:
        v = el.legacy_properties().get("Value")
        if v:
            parts.append(str(v))
    except Exception:
        pass
    return " ".join(parts).strip()


def has_exact_invoice_doc_open(win_1c, invoice_number: str) -> bool:
    num_clean = str(invoice_number).strip().lstrip("0") or "0"
    full_num = format_1c_invoice_search(invoice_number).lower()
    num_pat = re.compile(rf"(?:^|[-0]){re.escape(num_clean)}(?:\b|\s|от)")
    has_doc_toolbar = has_matching_header = False
    try:
        win_rect = win_1c.rectangle()
        for el in win_1c.descendants():
            txt = get_el_text_all(el).lower()
            if not txt:
                continue
            r = el.rectangle()
            if r.width() < 15 or r.height() < 10:
                continue
            if txt == "провести и закрыть":
                has_doc_toolbar = True
            if win_rect.top + 48 < r.top < win_rect.top + 220:
                if ("счет покупателю" in txt and num_pat.search(txt)) or txt == full_num:
                    has_matching_header = True
    except Exception:
        pass
    return has_doc_toolbar and has_matching_header


def has_open_print_form(win_1c) -> bool:
    try:
        for el in win_1c.descendants():
            txt = (el.window_text() or "").strip().lower()
            if txt == "печать документа" or txt.startswith("копий"):
                r = el.rectangle()
                if r.width() > 15 and r.height() > 10:
                    return True
    except Exception:
        pass
    return False


def click_1c_element_by_name(win_1c, exact_or_part: str, timeout: float = 8.0, exact_only: bool = False) -> bool:
    deadline = time.time() + timeout
    needle = exact_or_part.lower().replace("ё", "е")
    while time.time() < deadline:
        try:
            for el in win_1c.descendants():
                txt = (el.window_text() or "").strip().lower().replace("ё", "е")
                matched = (txt == needle) if exact_only else (txt == needle or (len(needle) >= 5 and needle in txt))
                if matched:
                    rect = el.rectangle()
                    if rect.width() > 5 and rect.height() > 5:
                        el.click_input()
                        return True
        except Exception:
            pass
        time.sleep(0.4)
    return False


def click_1c_popup_item(win_1c, exact_text: str, timeout: float = 2.5) -> bool:
    deadline = time.time() + timeout
    needle = exact_text.lower().replace("ё", "е")
    pid_1c = win_1c.process_id()
    desktop = Desktop(backend="uia")
    while time.time() < deadline:
        try:
            for w in desktop.windows(process=pid_1c):
                for el in w.descendants():
                    txt = (el.window_text() or "").strip().lower().replace("ё", "е")
                    if txt == needle:
                        rect = el.rectangle()
                        if rect.width() > 5 and rect.height() > 5:
                            el.click_input()
                            return True
        except Exception:
            pass
        time.sleep(0.25)
    return False


def switch_to_open_invoice_tab(win_1c, invoice_number: str) -> bool:
    num_clean = str(invoice_number).strip().lstrip("0") or "0"
    num_pat = re.compile(rf"(?:^|[-0]){re.escape(num_clean)}(?:\b|\s|от)")
    try:
        win_rect = win_1c.rectangle()
        for el in win_1c.descendants():
            txt = (el.window_text() or "").strip().lower()
            if "счет покупателю" in txt and num_pat.search(txt):
                r = el.rectangle()
                if r.width() > 20 and r.height() > 10 and r.top < win_rect.top + 95:
                    el.click_input()
                    time.sleep(0.7)
                    return True
    except Exception:
        pass
    return False


def parse_areas_from_tsv(txt: str) -> dict:
    areas = {}
    area_col_idx = None
    for line in txt.splitlines():
        cols = [c.strip() for c in line.split("\t")]
        low_cols = [c.lower() for c in cols]
        if any("площад" in c for c in low_cols):
            for ci, c in enumerate(low_cols):
                if "площад" in c:
                    area_col_idx = ci
                    break
            continue
        if cols and cols[0].isdigit():
            row_num = int(cols[0])
            if area_col_idx is not None and area_col_idx < len(cols):
                val = parse_number(cols[area_col_idx])
                if val > 0:
                    areas[row_num] = val
            elif len(cols) >= 4 and looks_like_number(cols[3]):
                val = parse_number(cols[3])
                if val > 0:
                    areas[row_num] = val
    return areas


def grab_doc_extras_from_open_invoice(win_1c, log_cb) -> dict:
    """
    Забирает из открытого документа «Счет покупателю» в 1С:
      - ИГК (20-25 цифр)
      - колонку «Площадь» через ПКМ по таблице -> «Вывести список...» -> Enter -> Ctrl+A/C -> Ctrl+F4
    """
    result = {"areas": {}, "igk": ""}
    try:
        win_rect = win_1c.rectangle()
        add_btn_bottom = None

        for el in win_1c.descendants():
            txt = get_el_text_all(el)
            if not txt:
                continue
            r = el.rectangle()
            if r.width() < 15 or r.height() < 10:
                continue
            low = txt.lower().strip()
            if low == "добавить" and r.top > win_rect.top + 150:
                add_btn_bottom = r.bottom
            if not result["igk"] and win_rect.top + 70 < r.top < win_rect.top + 380:
                m_igk = re.search(r"\b(\d{20,25})\b", txt)
                if m_igk:
                    result["igk"] = m_igk.group(1)

        if result["igk"]:
            log_cb(f"Подтянули ИГК из 1С: {result['igk']}")

        if add_btn_bottom is None:
            return result

        log_cb("Снимаем колонку «Площадь» из таблицы документа 1С...")
        tbl_x = win_rect.left + int(win_rect.width() * 0.25)
        tbl_y = add_btn_bottom + 48
        mouse.click(coords=(tbl_x, tbl_y))
        time.sleep(0.25)
        mouse.right_click(coords=(tbl_x, tbl_y))
        time.sleep(0.6)

        if click_1c_popup_item(win_1c, "Вывести список...", timeout=2.5) or click_1c_popup_item(win_1c, "Вывести список", timeout=1.0):
            time.sleep(0.6)
            press_vk(win32con.VK_RETURN)
            time.sleep(1.3)

            clear_clipboard()
            time.sleep(0.15)
            mouse.click(coords=(win_rect.left + int(win_rect.width() * 0.45), win_rect.top + int(win_rect.height() * 0.45)))
            time.sleep(0.25)
            press_ctrl_key(0x41)
            time.sleep(0.35)
            press_ctrl_key(0x43)
            time.sleep(0.6)

            txt = get_clipboard_text()
            if txt.strip():
                result["areas"] = parse_areas_from_tsv(txt)

            press_ctrl_f4()
            time.sleep(0.5)

        if result["areas"]:
            log_cb(f"Подтянули «Площадь» из 1С ({len(result['areas'])} поз.).")
    except Exception:
        pass
    return result


def focus_list_search_box_and_filter(win_1c, search_query: str, log_cb):
    win_rect = win_1c.rectangle()
    toolbar_top = toolbar_bottom = None
    try:
        for el in win_1c.descendants():
            if (el.window_text() or "").strip().lower() == "создать":
                r = el.rectangle()
                if r.width() > 10 and r.top > win_rect.top + 50:
                    toolbar_top, toolbar_bottom = r.top, r.bottom
                    break
    except Exception:
        pass

    search_edit = None
    try:
        candidates = []
        for el in win_1c.descendants(control_type="Edit"):
            r = el.rectangle()
            if r.width() < 40 or r.height() < 12 or r.top < win_rect.top + 60:
                continue
            name = (el.window_text() or "").lower()
            is_right = r.left > (win_rect.left + win_rect.width() * 0.40)
            near_toolbar = (toolbar_top is None) or (abs(r.top - toolbar_top) < 65)
            if "поиск" in name and near_toolbar:
                candidates.insert(0, el)
            elif is_right and near_toolbar:
                candidates.append(el)
        if candidates:
            search_edit = candidates[0]
            if toolbar_bottom is None:
                toolbar_bottom = search_edit.rectangle().bottom
    except Exception:
        pass

    if search_edit is not None:
        log_cb(f"Вводим «{search_query}» в поле поиска справа...")
        search_edit.click_input()
        time.sleep(0.3)
        press_ctrl_key(0x41)
        time.sleep(0.1)
        press_vk(win32con.VK_BACK)
        time.sleep(0.15)
        type_unicode_text(search_query)
        time.sleep(0.35)
        press_vk(win32con.VK_RETURN)
        time.sleep(2.2)
        return toolbar_bottom, False

    log_cb(f"Вводим «{search_query}» через Ctrl+F...")
    press_ctrl_key(0x46)
    time.sleep(0.9)
    type_unicode_text(search_query)
    time.sleep(0.35)
    press_vk(win32con.VK_RETURN)
    time.sleep(2.2)
    return toolbar_bottom, True


def open_filtered_invoice_row(win_1c, invoice_number: str, toolbar_bottom, used_ctrl_f: bool, log_cb) -> bool:
    if has_exact_invoice_doc_open(win_1c, invoice_number):
        return True

    num_str = str(invoice_number).strip().lstrip("0") or "0"
    win_rect = win_1c.rectangle()
    pattern = re.compile(rf"(?:^|[-0]){re.escape(num_str)}\b")
    try:
        for el in win_1c.descendants():
            txt = get_el_text_all(el)
            if not txt:
                continue
            r = el.rectangle()
            if r.top < win_rect.top + 120 or r.width() < 20 or r.height() < 10:
                continue
            if r.left > win_rect.left + win_rect.width() * 0.55:
                continue
            if pattern.search(txt):
                log_cb(f"Найдена строка «{txt}» — открываем...")
                el.double_click_input()
                time.sleep(2.0)
                if has_exact_invoice_doc_open(win_1c, invoice_number):
                    return True
    except Exception:
        pass

    if used_ctrl_f:
        log_cb("Открываем найденную строку (Enter)...")
        press_vk(win32con.VK_RETURN)
        time.sleep(2.0)
        if has_exact_invoice_doc_open(win_1c, invoice_number):
            return True

    log_cb("Переходим в таблицу списка (↓ → Enter)...")
    press_vk(win32con.VK_DOWN)
    time.sleep(0.4)
    press_vk(win32con.VK_RETURN)
    time.sleep(2.0)
    if has_exact_invoice_doc_open(win_1c, invoice_number):
        return True

    try:
        for ct in ("Table", "DataGrid"):
            for tbl in win_1c.descendants(control_type=ct):
                r = tbl.rectangle()
                if r.width() > 300 and r.height() > 120 and r.top > win_rect.top + 80:
                    abs_x = r.left + min(180, r.width() // 4)
                    for offset_y in (42, 58, 30):
                        mouse.double_click(coords=(abs_x, r.top + offset_y))
                        time.sleep(2.0)
                        if has_exact_invoice_doc_open(win_1c, invoice_number):
                            return True
    except Exception:
        pass

    base_y = toolbar_bottom if toolbar_bottom else (win_rect.top + int(win_rect.height() * 0.14))
    abs_x = win_rect.left + int(win_rect.width() * 0.22)
    for dy in (65, 90, 115):
        mouse.double_click(coords=(abs_x, base_y + dy))
        time.sleep(2.0)
        if has_exact_invoice_doc_open(win_1c, invoice_number):
            return True
    return False


def find_document_print_button(win_1c):
    win_rect = win_1c.rectangle()
    doc_toolbar_y = None
    try:
        for el in win_1c.descendants():
            if (el.window_text() or "").strip().lower() == "провести и закрыть":
                r = el.rectangle()
                if r.width() > 20 and r.height() > 10:
                    doc_toolbar_y = r.top
                    break
    except Exception:
        pass

    best_btn, best_dist = None, 999999
    try:
        for el in win_1c.descendants():
            txt = (el.window_text() or "").strip().lower()
            if txt == "печать" or txt.startswith("печать "):
                r = el.rectangle()
                if r.width() > 20 and r.height() > 10 and r.top > win_rect.top + 50:
                    dist = abs(r.top - doc_toolbar_y) if doc_toolbar_y is not None else r.top
                    if dist < best_dist:
                        best_dist, best_btn = dist, el
    except Exception:
        pass
    return best_btn


def click_document_print_77(win_1c, log_cb):
    print_btn = find_document_print_button(win_1c)
    if print_btn is None:
        raise RuntimeError("Не найдена кнопка «Печать» в открытом счёте.")

    btn_rect = print_btn.rectangle()
    btn_h = max(24, btn_rect.height())

    log_cb("Нажимаем «Печать» → «Счёт в формате 7.7»...")
    print_btn.click_input()
    time.sleep(0.8)

    if click_1c_popup_item(win_1c, "Счёт в формате 7.7", timeout=2.5):
        for _ in range(12):
            time.sleep(0.5)
            if has_open_print_form(win_1c):
                time.sleep(0.7)
                return

    if not has_open_print_form(win_1c):
        log_cb("Выбираем 4-й пункт меню («Счёт в формате 7.7»)...")
        mouse.click(coords=(btn_rect.left + 65, btn_rect.bottom + int(btn_h * 3.85)))
        for _ in range(10):
            time.sleep(0.5)
            if has_open_print_form(win_1c):
                time.sleep(0.7)
                return

    if not has_open_print_form(win_1c):
        log_cb("Повторно открываем «Печать» и выбираем стрелками...")
        print_btn.click_input()
        time.sleep(0.7)
        for _ in range(4):
            press_vk(win32con.VK_DOWN)
            time.sleep(0.15)
        press_vk(win32con.VK_RETURN)
        for _ in range(12):
            time.sleep(0.5)
            if has_open_print_form(win_1c):
                time.sleep(0.7)
                return

    raise RuntimeError(
        "Не удалось открыть печатную форму «Счёт в формате 7.7».\n"
        "Открой её в 1С вручную и включи кнопку «Печатная форма уже открыта»."
    )


def grab_text_from_1c_print_form(cfg: dict, invoice_number: str, use_open_form: bool, log_cb):
    log_cb("Ищем окно 1С:Предприятие...")
    win_1c = find_1c_window()
    just_launched = False
    extras = {"areas": {}, "igk": ""}

    if win_1c is None:
        if use_open_form:
            raise RuntimeError("Окно 1С:Предприятие не найдено на этом компьютере.")
        log_cb("Запускаем 1С и входим в базу...")
        exe_1c = resolve_1c_exe(cfg.get("1c_exe", ""))
        subprocess.Popen(f'"{exe_1c}" {cfg["1c_base_args"]}', shell=True)
        just_launched = True
        for _ in range(45):
            time.sleep(2.0)
            win_1c = find_1c_window()
            if win_1c:
                break

    if win_1c is None:
        raise RuntimeError("Не удалось найти или запустить окно 1С:Предприятие.")

    if just_launched:
        log_cb("Ждём полную загрузку интерфейса 1С (6 сек)...")
        time.sleep(6.0)

    try:
        if win_1c.is_minimized():
            win_1c.restore()
        win_1c.maximize()
    except Exception:
        pass
    win_1c.set_focus()
    time.sleep(0.7)

    if not use_open_form:
        if not has_exact_invoice_doc_open(win_1c, invoice_number):
            log_cb("Переходим в «Счета покупателям»...")
            if not click_1c_element_by_name(win_1c, "Счета покупателям", timeout=2.5, exact_only=True):
                if click_1c_element_by_name(win_1c, "Продажи", timeout=8.0, exact_only=True):
                    time.sleep(0.9)
                    click_1c_element_by_name(win_1c, "Счета покупателям", timeout=6.0)
            time.sleep(1.5)

            full_search = format_1c_invoice_search(invoice_number)
            toolbar_bottom, used_ctrl_f = focus_list_search_box_and_filter(win_1c, full_search, log_cb)

            if not open_filtered_invoice_row(win_1c, invoice_number, toolbar_bottom, used_ctrl_f, log_cb):
                raise RuntimeError(f"В списке «Счета покупателям» не удалось открыть счёт {full_search}.")
            time.sleep(1.0)

        extras = grab_doc_extras_from_open_invoice(win_1c, log_cb)
        click_document_print_77(win_1c, log_cb)

    log_cb("Копируем таблицу из печатной формы (Ctrl+A → Ctrl+C)...")
    rect = win_1c.rectangle()
    raw_text = ""
    for fx, fy in ((0.45, 0.55), (0.35, 0.62)):
        clear_clipboard()
        time.sleep(0.2)
        try:
            mouse.click(coords=(int(rect.left + rect.width() * fx), int(rect.top + rect.height() * fy)))
            time.sleep(0.4)
        except Exception:
            pass
        press_ctrl_key(0x41)
        time.sleep(0.5)
        press_ctrl_key(0x43)
        time.sleep(0.9)
        raw_text = get_clipboard_text()
        if raw_text.strip():
            break

    if not raw_text.strip():
        raise RuntimeError(
            "Буфер обмена остался пустым после копирования из 1С.\n\n"
            "Быстрый обход: открой «Счёт в формате 7.7» в 1С, нажми Ctrl+A и Ctrl+C, "
            "а затем нажми кнопку «📋 Вставить из буфера»."
        )

    num_clean = str(invoice_number).strip().lstrip("0") or "0"
    m_inv = re.search(r"Сч[её]т\s*(?:на\s+оплату\s*)?№\s*0*(\d+)", raw_text, flags=re.IGNORECASE)
    if m_inv and m_inv.group(1) != num_clean:
        raise RuntimeError(
            f"1С открыла не тот счёт: в печатной форме № {m_inv.group(1)}, а нужен № {num_clean}."
        )

    if use_open_form and not extras["areas"]:
        if switch_to_open_invoice_tab(win_1c, invoice_number):
            extras = grab_doc_extras_from_open_invoice(win_1c, log_cb)

    return raw_text, extras


def parse_invoice_clipboard_text(raw_text: str) -> dict:
    invoice_date = datetime.now().strftime("%d.%m.%Y")
    items = []
    totals = {"no_nds": 0.0, "nds": 0.0, "with_nds": 0.0}
    vat_rate = "22"
    igk_from_text = ""

    date_re = re.compile(r"от\s+(\d{1,2})\s+([а-яА-Я]+)\s+(\d{4})", re.IGNORECASE)
    months = {
        "января": "01", "февраля": "02", "марта": "03", "апреля": "04",
        "мая": "05", "июня": "06", "июля": "07", "августа": "08",
        "сентября": "09", "октября": "10", "ноября": "11", "декабря": "12",
    }

    in_table = False
    for line in raw_text.splitlines():
        cells = [c.strip() for c in line.split("\t") if c.strip() != ""]
        if not cells:
            continue
        joined = " ".join(cells)

        m = date_re.search(joined)
        if m:
            invoice_date = f"{m.group(1).zfill(2)}.{months.get(m.group(2).lower(), '01')}.{m.group(3)}"

        if not igk_from_text and ("игк" in joined.lower() or "основание" in joined.lower()):
            m_igk = re.search(r"\b(\d{20,25})\b", joined)
            if m_igk:
                igk_from_text = m_igk.group(1)

        if ("Товар" in joined or "Наименование" in joined) and "Цена" in joined:
            in_table = True
            continue

        if in_table:
            if joined.startswith("Итого") or "Итого" in cells[0]:
                in_table = False
                totals["no_nds"] = parse_number(cells[-1])
                continue
            if cells[0].isdigit() and len(cells) >= 6:
                board_name, ps_num = split_board_and_ps(cells[1])
                if looks_like_number(cells[2]) and not looks_like_number(cells[3]):
                    qty_val, unit_val = parse_number(cells[2]), cells[3]
                else:
                    unit_val, qty_val = cells[2], parse_number(cells[3])
                items.append({
                    "raw_name": cells[1], "name": board_name, "ps_num": ps_num,
                    "unit": unit_val, "qty": qty_val, "price": parse_number(cells[4]),
                    "sum": parse_number(cells[5]), "term": cells[6] if len(cells) >= 7 else "",
                })

        if "НДС" in joined and totals["nds"] == 0.0 and "без НДС" not in joined.lower():
            m_vat = re.search(r"НДС\s*\(?(\d{1,2})\s*%", joined, flags=re.IGNORECASE)
            if m_vat:
                vat_rate = m_vat.group(1)
            totals["nds"] = parse_number(cells[-1])
        elif "Всего к оплате" in joined:
            totals["with_nds"] = parse_number(cells[-1])

    if totals["no_nds"] == 0.0 and items:
        totals["no_nds"] = round(sum(i["sum"] for i in items), 2)
    if totals["nds"] == 0.0 and totals["no_nds"] > 0:
        totals["nds"] = round(totals["no_nds"] * (int(vat_rate) / 100.0), 2)
    if totals["with_nds"] == 0.0 and totals["no_nds"] > 0:
        totals["with_nds"] = round(totals["no_nds"] + totals["nds"], 2)

    return {
        "date": invoice_date,
        "year": invoice_date.split(".")[-1] if "." in invoice_date else str(datetime.now().year),
        "vat_rate": vat_rate,
        "igk": igk_from_text,
        "items": items,
        "totals": totals,
    }


# =============================================================================
# ПОИСК БЛАНКА И НИЖНЕЙ СТРОКИ В «РАСЧЕТЫ»
# =============================================================================
def extract_board_tokens(board_name: str) -> list:
    stop_words = {"печ.платы", "печатные", "платы", "плата", "ps", "пп", "шт"}
    return [
        t.strip().lower() for t in re.split(r"[\s()\[\],;]+", board_name)
        if len(t.strip()) >= 3 and t.strip().lower() not in stop_words
    ]


def is_ignored_calc_filename(fname: str) -> bool:
    low = fname.lower()
    return low.startswith("~$") or any(b in low for b in ("счет", "счёт", "расчет", "расчёт", "прайс"))


def find_invoice_calc_folder(calc_folder: str, invoice_num: str) -> str:
    if not calc_folder or not os.path.isdir(calc_folder) or not invoice_num:
        return ""
    num_clean = str(invoice_num).strip().lstrip("0") or "0"
    pat = re.compile(rf"(?:^|[^0-9])0*{re.escape(num_clean)}(?:[^0-9]|$)")
    search_roots = [calc_folder]
    try:
        for entry in os.listdir(calc_folder):
            sub = os.path.join(calc_folder, entry)
            if os.path.isdir(sub):
                if pat.search(entry):
                    return sub
                search_roots.append(sub)
    except Exception:
        pass
    for root_dir in search_roots[1:]:
        try:
            for entry in os.listdir(root_dir):
                sub = os.path.join(root_dir, entry)
                if os.path.isdir(sub) and pat.search(entry):
                    return sub
        except Exception:
            pass
    return ""


def extract_6_tech_cells_from_rows(rows: list) -> list:
    for r in range(len(rows) - 1, -1, -1):
        row = rows[r]
        joined = " ".join(str(v or "").lower() for v in row)
        if "маска" in joined and ("сло" in joined or "покр" in joined or "марк" in joined):
            if r + 1 < len(rows):
                start_col = next((c for c, val in enumerate(row) if "сло" in str(val or "").lower()), 0)
                data_row = rows[r + 1]
                out = []
                for c in range(start_col, start_col + 6):
                    v = data_row[c] if c < len(data_row) else ""
                    if isinstance(v, float) and v.is_integer():
                        v = int(v)
                    out.append("" if v is None else v)
                return out
    for r in range(len(rows) - 1, -1, -1):
        non_empty = [v for v in rows[r] if v is not None and str(v).strip() != ""]
        if len(non_empty) >= 3:
            cleaned = [int(v) if isinstance(v, float) and v.is_integer() else v for v in non_empty]
            return (cleaned + [""] * 6)[:6]
    return [""] * 6


def read_tech_cells_from_excel(filepath: str) -> list:
    rows = []
    if filepath.lower().endswith(".xls"):
        wb = xlrd.open_workbook(filepath)
        sh = wb.sheet_by_index(0)
        for r in range(sh.nrows):
            rows.append([sh.cell_value(r, c) for c in range(sh.ncols)])
    else:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        for r in range(1, ws.max_row + 1):
            rows.append([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])
    return extract_6_tech_cells_from_rows(rows)


def pick_best_blank_in_folder(folder: str, tokens: list, item_idx: int = 0) -> str:
    if not folder or not os.path.isdir(folder):
        return ""
    files = [
        f for f in os.listdir(folder)
        if f.lower().endswith((".xls", ".xlsx")) and not is_ignored_calc_filename(f)
    ]
    best_file, best_hits = "", 0
    for fname in files:
        hits = sum(1 for tok in tokens if tok in fname.lower())
        if hits > best_hits:
            best_hits, best_file = hits, os.path.join(folder, fname)
    if best_file:
        return best_file
    if item_idx > 0:
        for ext in (".xls", ".xlsx"):
            for fname in files:
                if fname.lower() == f"{item_idx}{ext}":
                    return os.path.join(folder, fname)
    return ""


def find_calc_bottom_row(calc_folder: str, board_name: str, raw_name: str = "", invoice_num: str = "", item_idx: int = 0) -> list:
    if not calc_folder or not os.path.isdir(calc_folder):
        return [""] * 6
    tokens = extract_board_tokens(board_name + " " + raw_name)
    target_file = ""
    inv_folder = find_invoice_calc_folder(calc_folder, invoice_num)
    if inv_folder:
        target_file = pick_best_blank_in_folder(inv_folder, tokens, item_idx=item_idx)
    if not target_file and tokens:
        target_file = pick_best_blank_in_folder(calc_folder, tokens, item_idx=0)
    if not target_file:
        return [""] * 6
    try:
        return read_tech_cells_from_excel(target_file)
    except Exception:
        return [""] * 6


# =============================================================================
# СБОРКА EXCEL-ФАЙЛА СПЕЦИФИКАЦИИ
# =============================================================================
def build_specification(cfg: dict, client_name: str, calc_folder_name: str, invoice_num: str,
                        new_spec_num: int, use_open_form: bool, from_clipboard_only: bool, log_cb) -> str:
    spec_folder = os.path.join(cfg["specs_root"], client_name)
    calc_folder = (
        os.path.join(cfg["calc_root"], calc_folder_name)
        if calc_folder_name and calc_folder_name != "(не выбрана)" else ""
    )

    log_cb("Ищем последнюю спецификацию и делаем копию...")
    _, last_file = find_latest_spec(spec_folder)
    ext = os.path.splitext(last_file)[1]
    new_file = os.path.join(spec_folder, f"сп.{new_spec_num} сч.{invoice_num}{ext}")
    shutil.copy2(last_file, new_file)

    extras = {"areas": {}, "igk": ""}
    if from_clipboard_only:
        log_cb("Читаем таблицу из буфера обмена...")
        raw_text = get_clipboard_text()
        if not raw_text.strip():
            raise RuntimeError("Буфер обмена пуст. Скопируй «Счёт в формате 7.7» в 1С (Ctrl+A, Ctrl+C).")
    else:
        raw_text, extras = grab_text_from_1c_print_form(cfg, invoice_num, use_open_form, log_cb)

    log_cb("Разбираем строки счёта...")
    inv = parse_invoice_clipboard_text(raw_text)
    items = inv["items"]
    if not items:
        raise RuntimeError("Не удалось распознать товары в скопированном тексте 1С.")

    areas_map = extras.get("areas", {})
    final_igk = extras.get("igk") or inv.get("igk", "")
    inv_year = inv["year"]
    vat_rate = inv["vat_rate"]

    inv_calc_dir = find_invoice_calc_folder(calc_folder, invoice_num)
    if inv_calc_dir:
        log_cb(f"Найдена папка счёта в РАСЧЕТЫ: {os.path.basename(inv_calc_dir)}")

    log_cb(f"Найдено позиций: {len(items)}. Заполняем Excel (дата счёта: {inv['date']})...")
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(new_file)
        ws = wb.Sheets(1)

        for r in range(1, 4):
            for c in range(1, 17):
                val = ws.Cells(r, c).Value
                if isinstance(val, str) and "приложение" in val.lower():
                    ws.Cells(r, c).Value = re.sub(r"(Приложение\s*№\s*)\d+", rf"\g<1>{new_spec_num}", val, flags=re.IGNORECASE)

        spec_title_updated = False
        for r in range(2, 7):
            for c in range(1, 5):
                val = ws.Cells(r, c).Value
                if isinstance(val, str) and "спецификация" in val.lower():
                    ws.Cells(r, c).Value = f"Спецификация № {new_spec_num} от {inv['date']} г."
                    spec_title_updated = True
                    break
            if spec_title_updated:
                break
        if not spec_title_updated:
            ws.Range("A4").Value = f"Спецификация № {new_spec_num} от {inv['date']} г."

        header_row = 8
        for r in range(5, 13):
            if "наименование" in str(ws.Cells(r, 2).Value or "").lower():
                header_row = r
                break
        start_row = header_row + 1

        d_hdr = str(ws.Cells(header_row, 4).Value or "").lower()
        e_hdr = str(ws.Cells(header_row, 5).Value or "").lower()
        unit_col, term_col = (4, 5) if ("ед" in d_hdr and "срок" in e_hdr) else (5, 4)

        total_row = start_row
        while total_row < 150:
            if "итого" in " ".join(str(ws.Cells(total_row, c).Value or "") for c in range(1, 8)).lower():
                break
            total_row += 1

        existing_rows, needed_rows = total_row - start_row, len(items)
        if needed_rows > existing_rows:
            for _ in range(needed_rows - existing_rows):
                ws.Rows(start_row + 1).Insert()
                ws.Rows(start_row).Copy(ws.Rows(start_row + 1))
        elif needed_rows < existing_rows:
            for _ in range(existing_rows - needed_rows):
                ws.Rows(start_row + needed_rows).Delete()

        for idx, item in enumerate(items, start=1):
            r = start_row + idx - 1
            ws.Cells(r, 1).Value = idx
            ws.Cells(r, 2).Value = item["name"]
            ws.Cells(r, 3).Value = item["ps_num"]
            ws.Cells(r, term_col).Value = item["term"]
            ws.Cells(r, unit_col).Value = item["unit"]
            ws.Cells(r, 6).Value = item["qty"]
            ws.Cells(r, 7).Value = item["price"]
            ws.Cells(r, 8).Value = format_rub_text(item["sum"], False)
            ws.Cells(r, 9).Value = areas_map.get(idx, "")

            tech = find_calc_bottom_row(calc_folder, item["name"], item["raw_name"], invoice_num=invoice_num, item_idx=idx)
            for offset, val in enumerate(tech):
                ws.Cells(r, 10 + offset).Value = val

        new_total_row = start_row + needed_rows
        for r in range(new_total_row, new_total_row + 22):
            ag_low = " ".join(str(ws.Cells(r, c).Value or "") for c in range(1, 8)).lower()
            ab_low = " ".join(str(ws.Cells(r, c).Value or "") for c in range(1, 3)).lower()

            if "итого с ндс" in ag_low:
                ws.Cells(r, 8).Value = format_rub_text(inv["totals"]["with_nds"], True)
            elif "итого без ндс" in ag_low or ("итого" in ag_low and "ндс" not in ag_low):
                ws.Cells(r, 8).Value = format_rub_text(inv["totals"]["no_nds"], True)
            elif "в том числе ндс" in ab_low:
                for c in (1, 2):
                    v = ws.Cells(r, c).Value
                    if isinstance(v, str) and "ндс" in v.lower():
                        ws.Cells(r, c).Value = re.sub(r"\(\d{1,2}%\)", f"({vat_rate}%)", v)
                ws.Cells(r, 3).Value = amount_in_words(inv["totals"]["nds"])
            elif "сумма" in ab_low:
                ws.Cells(r, 3).Value = amount_in_words(inv["totals"]["with_nds"])
            elif "ндс" in ag_low and "итого" not in ag_low:
                for c in range(1, 8):
                    v = ws.Cells(r, c).Value
                    if isinstance(v, str) and "ндс" in v.lower():
                        ws.Cells(r, c).Value = re.sub(r"\(\d{1,2}%\)", f"({vat_rate}%)", v)
                ws.Cells(r, 8).Value = format_rub_text(inv["totals"]["nds"], True)

            for c in range(1, 16):
                v = ws.Cells(r, c).Value
                if not isinstance(v, str):
                    continue
                if final_igk and ("игк" in v.lower() or "гк №" in v.lower()):
                    ws.Cells(r, c).Value = f"ИГК {final_igk}"
                elif re.search(r"20\d{2}\s*г\.?", v) and "договор" not in v.lower():
                    ws.Cells(r, c).Value = re.sub(r"20\d{2}(\s*г\.?)", rf"{inv_year}\1", v)

        wb.Save()
    finally:
        pythoncom.CoUninitialize()
    return new_file


# =============================================================================
# ПРИЛОЖЕНИЕ (КОМПАКТНОЕ ОКНО, ЗАКРЕПЛЁННЫЕ КНОПКИ ВНИЗУ, ЯРКАЯ КНОПКА 1С)
# =============================================================================
class SpecBuilderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.user_settings = load_user_settings()
        self.current_theme_name = self.user_settings.get("theme", "Тёмная ночь")
        if self.current_theme_name not in THEMES:
            self.current_theme_name = "Тёмная ночь"
        self.folder_mapping = self.user_settings.get("folder_mapping", dict(CFG["folder_mapping"]))
        self.t = THEMES[self.current_theme_name]

        ctk.set_appearance_mode(self.t["mode"])
        ctk.set_default_color_theme("green")

        self.title(f'{CFG["company_name"]} — Сборка спецификаций')
        win_w, win_h = 1160, 650
        pos_x = max(20, (self.winfo_screenwidth() - win_w) // 2)
        pos_y = max(20, (self.winfo_screenheight() - win_h) // 2 - 20)
        self.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
        self.minsize(1000, 600)

        self.all_clients, self.all_calc_folders = [], []
        self.selected_client = None
        self.selected_calc_folder = "(не выбрана)"
        self.calc_picker_open = False
        self.use_open_form = False
        self.client_buttons, self.calc_buttons = {}, {}

        self._load_folders()
        self._build_layout()
        self._apply_theme(self.current_theme_name, save=False)
        self._render_client_list("")
        self._set_calc_picker_visible(False)
        self._update_open_form_btn_style()

    def _load_folders(self):
        try:
            self.all_clients = sorted(
                d for d in os.listdir(CFG["specs_root"])
                if os.path.isdir(os.path.join(CFG["specs_root"], d))
            )
        except Exception as e:
            messagebox.showerror("Нет доступа к папке Спецификаций", f"{CFG['specs_root']}\n\n{e}")
            self.all_clients = []
        try:
            self.all_calc_folders = sorted(
                d for d in os.listdir(CFG["calc_root"])
                if os.path.isdir(os.path.join(CFG["calc_root"], d))
            )
        except Exception:
            self.all_calc_folders = []

    def _build_layout(self):
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0)
        self.grid_rowconfigure(1, weight=1)

        # === ШАПКА ===
        self.header = ctk.CTkFrame(self, height=58, corner_radius=0, border_width=1)
        self.header.grid(row=0, column=0, columnspan=3, sticky="ew")
        self.header.grid_columnconfigure(1, weight=1)

        brand_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        brand_frame.grid(row=0, column=0, padx=16, pady=8, sticky="w")
        self.logo_badge = ctk.CTkLabel(
            brand_frame, text="ПП", width=38, height=38, corner_radius=10,
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
        )
        self.logo_badge.pack(side="left", padx=(0, 10))

        title_box = ctk.CTkFrame(brand_frame, fg_color="transparent")
        title_box.pack(side="left")
        self.company_lbl = ctk.CTkLabel(
            title_box, text=CFG["company_name"], font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold")
        )
        self.company_lbl.pack(anchor="w")
        self.subtitle_lbl = ctk.CTkLabel(
            title_box, text="Автоматическая сборка спецификаций (1С + Расчёты)", font=ctk.CTkFont(family="Segoe UI", size=11)
        )
        self.subtitle_lbl.pack(anchor="w")

        right_header = ctk.CTkFrame(self.header, fg_color="transparent")
        right_header.grid(row=0, column=2, padx=16, pady=8, sticky="e")

        self.theme_menu = ctk.CTkOptionMenu(
            right_header, values=list(THEMES.keys()), command=self._on_theme_change,
            width=170, height=32, corner_radius=8, font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
        )
        self.theme_menu.set(self.current_theme_name)
        self.theme_menu.pack(side="left", padx=(0, 10))

        self.user_badge = ctk.CTkFrame(right_header, corner_radius=10, border_width=1)
        self.user_badge.pack(side="left")
        self.user_dot = ctk.CTkLabel(self.user_badge, text="●", font=ctk.CTkFont(size=12))
        self.user_dot.pack(side="left", padx=(10, 6), pady=6)
        self.user_lbl = ctk.CTkLabel(
            self.user_badge, text=CFG["employee_name"], font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        )
        self.user_lbl.pack(side="left", padx=(0, 12), pady=6)

        # === ЛЕВОЕ МЕНЮ ===
        self.left_nav = ctk.CTkFrame(self, width=180, corner_radius=0, border_width=1)
        self.left_nav.grid(row=1, column=0, sticky="ns")
        self.left_nav.grid_propagate(False)

        self.nav_title = ctk.CTkLabel(self.left_nav, text="РАЗДЕЛЫ", font=ctk.CTkFont(size=11, weight="bold"))
        self.nav_title.pack(anchor="w", padx=16, pady=(14, 6))

        self.nav_active_btn = ctk.CTkButton(
            self.left_nav, text="📄  Спецификация", anchor="w", height=38, font=ctk.CTkFont(size=13, weight="bold")
        )
        self.nav_active_btn.pack(fill="x", padx=12, pady=3)

        self.nav_spec_btn = ctk.CTkButton(
            self.left_nav, text="📁  Папка Спец.", anchor="w", height=36, fg_color="transparent",
            font=ctk.CTkFont(size=13), command=self._open_selected_spec_folder,
        )
        self.nav_spec_btn.pack(fill="x", padx=12, pady=3)

        self.nav_calc_btn = ctk.CTkButton(
            self.left_nav, text="📊  Папка Расчёты", anchor="w", height=36, fg_color="transparent",
            font=ctk.CTkFont(size=13), command=self._open_selected_calc_folder,
        )
        self.nav_calc_btn.pack(fill="x", padx=12, pady=3)

        self.nav_reload_btn = ctk.CTkButton(
            self.left_nav, text="🔄  Обновить списки", anchor="w", height=36, fg_color="transparent",
            font=ctk.CTkFont(size=12), command=self._reload_all_folders,
        )
        self.nav_reload_btn.pack(fill="x", padx=12, pady=(12, 4))

        self.status_badge = ctk.CTkLabel(
            self.left_nav, text="● Готов к работе", font=ctk.CTkFont(size=12, weight="bold"),
            wraplength=150, justify="left",
        )
        self.status_badge.pack(side="bottom", anchor="w", padx=14, pady=14)

        # === ЦЕНТРАЛЬНАЯ КОЛОНКА ===
        self.center_container = ctk.CTkFrame(self, fg_color="transparent")
        self.center_container.grid(row=1, column=1, sticky="nsew", padx=12, pady=10)
        self.center_container.grid_columnconfigure(0, weight=1)
        self.center_container.grid_rowconfigure(0, weight=1)
        self.center_container.grid_rowconfigure(1, weight=0)

        self.center = ctk.CTkScrollableFrame(self.center_container, corner_radius=0)
        self.center.grid(row=0, column=0, sticky="nsew")

        # Карточка 1
        self.card1 = ctk.CTkFrame(self.center, corner_radius=12, border_width=1)
        self.card1.pack(fill="x", pady=(0, 10))
        self.c1_title = ctk.CTkLabel(self.card1, text="1. ФИРМА, ШАБЛОН И ПАПКА В «РАСЧЕТЫ»", font=ctk.CTkFont(size=12, weight="bold"))
        self.c1_title.pack(anchor="w", padx=14, pady=(10, 4))

        firm_row = ctk.CTkFrame(self.card1, fg_color="transparent")
        firm_row.pack(fill="x", padx=14, pady=(0, 8))
        self.selected_firm_lbl = ctk.CTkLabel(firm_row, text="Выберите фирму в списке справа →", font=ctk.CTkFont(size=16, weight="bold"))
        self.selected_firm_lbl.pack(side="left")
        self.last_file_lbl = ctk.CTkLabel(firm_row, text="Последний файл: —", font=ctk.CTkFont(size=12))
        self.last_file_lbl.pack(side="right")

        self.calc_bar = ctk.CTkFrame(self.card1, corner_radius=10, border_width=1)
        self.calc_bar.pack(fill="x", padx=14, pady=(0, 12))
        self.calc_bar_left = ctk.CTkFrame(self.calc_bar, fg_color="transparent")
        self.calc_bar_left.pack(side="left", fill="x", expand=True, padx=12, pady=6)
        self.calc_title_lbl = ctk.CTkLabel(self.calc_bar_left, text="Папка в «РАСЧЕТЫ»:", font=ctk.CTkFont(size=11))
        self.calc_title_lbl.pack(anchor="w")
        self.calc_selected_lbl = ctk.CTkLabel(self.calc_bar_left, text="(не выбрана)", font=ctk.CTkFont(size=13, weight="bold"))
        self.calc_selected_lbl.pack(anchor="w")

        self.calc_toggle_btn = ctk.CTkButton(
            self.calc_bar, text="🔍 Изменить папку", width=150, height=32, corner_radius=8,
            border_width=1, font=ctk.CTkFont(size=12, weight="bold"), command=self._toggle_calc_picker,
        )
        self.calc_toggle_btn.pack(side="right", padx=10, pady=6)

        self.calc_picker_frame = ctk.CTkFrame(self.card1, fg_color="transparent")
        self.calc_search_var = ctk.StringVar()
        self.calc_search_var.trace_add("write", lambda *_: self._render_calc_list(self.calc_search_var.get()))
        self.calc_search_entry = ctk.CTkEntry(
            self.calc_picker_frame, textvariable=self.calc_search_var,
            placeholder_text="🔍 Введите название папки в «РАСЧЕТЫ» (по-русски или латиницей)...", height=32,
        )
        self.calc_search_entry.pack(fill="x", padx=14, pady=(0, 6))
        self.calc_scroll = ctk.CTkScrollableFrame(self.calc_picker_frame, height=105, corner_radius=8, border_width=1)
        self.calc_scroll.pack(fill="x", padx=14, pady=(0, 12))

        # Карточка 2
        self.card2 = ctk.CTkFrame(self.center, corner_radius=12, border_width=1)
        self.card2.pack(fill="x", pady=(0, 10))
        self.c2_title = ctk.CTkLabel(self.card2, text="2. ПАРАМЕТРЫ СПЕЦИФИКАЦИИ И СЧЁТА 1С", font=ctk.CTkFont(size=12, weight="bold"))
        self.c2_title.pack(anchor="w", padx=14, pady=(10, 8))

        grid2 = ctk.CTkFrame(self.card2, fg_color="transparent")
        grid2.pack(fill="x", padx=14, pady=(0, 10))
        grid2.grid_columnconfigure(0, weight=1)
        grid2.grid_columnconfigure(1, weight=1)

        self.lbl_spec_num = ctk.CTkLabel(grid2, text="Новый № спецификации:", font=ctk.CTkFont(size=12))
        self.lbl_spec_num.grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 4))
        self.spec_num_var = ctk.StringVar()
        self.spec_entry = ctk.CTkEntry(grid2, textvariable=self.spec_num_var, height=36, font=ctk.CTkFont(size=14, weight="bold"))
        self.spec_entry.grid(row=1, column=0, sticky="ew", padx=(0, 8))

        self.lbl_inv_num = ctk.CTkLabel(grid2, text="Номер счёта в 1С (дата подтянется из счёта):", font=ctk.CTkFont(size=12))
        self.lbl_inv_num.grid(row=0, column=1, sticky="w", padx=(8, 0), pady=(0, 4))
        self.inv_var = ctk.StringVar()
        self.inv_entry = ctk.CTkEntry(grid2, textvariable=self.inv_var, placeholder_text="412", height=36, font=ctk.CTkFont(size=14, weight="bold"))
        self.inv_entry.grid(row=1, column=1, sticky="ew", padx=(8, 0))

        self.open_form_btn = ctk.CTkButton(
            self.card2, height=44, corner_radius=10, border_width=2,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            command=self._toggle_open_form_mode,
        )
        self.open_form_btn.pack(fill="x", padx=14, pady=(0, 12))

        # Карточка 3
        self.card3 = ctk.CTkFrame(self.center, corner_radius=12, border_width=1)
        self.card3.pack(fill="both", expand=True, pady=(0, 4))
        self.c3_title = ctk.CTkLabel(self.card3, text="3. ХОД ВЫПОЛНЕНИЯ", font=ctk.CTkFont(size=12, weight="bold"))
        self.c3_title.pack(anchor="w", padx=14, pady=(10, 4))
        self.log_box = ctk.CTkTextbox(self.card3, height=95, border_width=1, font=ctk.CTkFont(family="Consolas", size=12))
        self.log_box.pack(fill="both", expand=True, padx=14, pady=(0, 12))
        self._append_log("Готово. Выберите фирму справа — папка в «РАСЧЕТЫ» подберётся автоматически.")

        # Нижние закреплённые кнопки
        btn_row = ctk.CTkFrame(self.center_container, fg_color="transparent")
        btn_row.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        btn_row.grid_columnconfigure(0, weight=3)
        btn_row.grid_columnconfigure(1, weight=1)

        self.build_btn = ctk.CTkButton(
            btn_row, text="СОБРАТЬ СПЕЦИФИКАЦИЮ И ОТКРЫТЬ EXCEL", height=46, corner_radius=10,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            command=lambda: self._on_start_build(from_clipboard_only=False),
        )
        self.build_btn.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.clip_btn = ctk.CTkButton(
            btn_row, text="📋 Вставить из буфера", height=46, corner_radius=10, border_width=1,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            command=lambda: self._on_start_build(from_clipboard_only=True),
        )
        self.clip_btn.grid(row=0, column=1, sticky="ew")

        # === ПРАВАЯ ПАНЕЛЬ ===
        self.right_panel = ctk.CTkFrame(self, width=290, corner_radius=0, border_width=1)
        self.right_panel.grid(row=1, column=2, sticky="ns")
        self.right_panel.grid_propagate(False)

        self.firms_header_lbl = ctk.CTkLabel(
            self.right_panel, text=f"Справочник фирм ({len(self.all_clients)})", font=ctk.CTkFont(size=13, weight="bold")
        )
        self.firms_header_lbl.pack(anchor="w", padx=14, pady=(12, 6))

        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._render_client_list(self.search_var.get()))
        self.search_entry = ctk.CTkEntry(self.right_panel, textvariable=self.search_var, placeholder_text="🔍 Поиск фирмы...", height=34)
        self.search_entry.pack(fill="x", padx=14, pady=(0, 8))

        self.clients_scroll = ctk.CTkScrollableFrame(self.right_panel, fg_color="transparent")
        self.clients_scroll.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def _set_calc_picker_visible(self, visible: bool):
        self.calc_picker_open = visible
        if visible:
            self.calc_picker_frame.pack(fill="x")
            self.calc_toggle_btn.configure(text="▲ Свернуть список")
            self._render_calc_list(self.calc_search_var.get())
        else:
            self.calc_picker_frame.pack_forget()
            self.calc_toggle_btn.configure(text="🔍 Изменить папку")

    def _toggle_calc_picker(self):
        self._set_calc_picker_visible(not self.calc_picker_open)
        if self.calc_picker_open:
            self.calc_search_entry.focus_set()

    def _toggle_open_form_mode(self):
        self.use_open_form = not self.use_open_form
        self._update_open_form_btn_style()

    def _update_open_form_btn_style(self):
        t = self.t
        if self.use_open_form:
            self.open_form_btn.configure(
                text="✅   ПЕЧАТНАЯ ФОРМА «СЧЁТ В ФОРМАТЕ 7.7» УЖЕ ОТКРЫТА В 1С  —  ЗАБЕРЁМ СРАЗУ",
                fg_color=t["toggle_on_bg"], hover_color=t["toggle_on_hover"],
                border_color=t["toggle_on_border"], text_color=t["toggle_on_text"],
            )
        else:
            self.open_form_btn.configure(
                text="⚡   Печатная форма «Счёт в формате 7.7» уже открыта в 1С  (нажмите, чтобы забрать сразу)",
                fg_color=t["toggle_off_bg"], hover_color=t["toggle_off_hover"],
                border_color=t["toggle_off_border"], text_color=t["toggle_off_text"],
            )

    def _on_theme_change(self, theme_name: str):
        self._apply_theme(theme_name, save=True)
        self._render_client_list(self.search_var.get())
        if self.calc_picker_open:
            self._render_calc_list(self.calc_search_var.get())

    def _apply_theme(self, theme_name: str, save: bool = True):
        self.current_theme_name = theme_name
        t = THEMES[theme_name]
        self.t = t
        ctk.set_appearance_mode(t["mode"])
        if save:
            self.user_settings["theme"] = theme_name
            self.user_settings["folder_mapping"] = self.folder_mapping
            save_user_settings(self.user_settings)

        self.configure(fg_color=t["bg_app"])
        self.header.configure(fg_color=t["bg_sidebar"], border_color=t["border"])
        self.logo_badge.configure(fg_color=t["accent"], text_color=t["accent_text"])
        self.company_lbl.configure(text_color=t["text_main"])
        self.subtitle_lbl.configure(text_color=t["text_muted"])
        self.theme_menu.configure(
            fg_color=t["badge_bg"], button_color=t["accent"], button_hover_color=t["accent_hover"],
            text_color=t["text_main"], dropdown_fg_color=t["bg_card"],
            dropdown_hover_color=t["hover"], dropdown_text_color=t["text_main"],
        )
        self.user_badge.configure(fg_color=t["badge_bg"], border_color=t["border"])
        self.user_dot.configure(text_color=t["accent"])
        self.user_lbl.configure(text_color=t["text_main"])
        self.left_nav.configure(fg_color=t["bg_sidebar"], border_color=t["border"])
        self.nav_title.configure(text_color=t["text_muted"])
        self.nav_active_btn.configure(fg_color=t["sel_bg"], hover_color=t["hover"], text_color=t["sel_text"])
        self.nav_spec_btn.configure(hover_color=t["hover"], text_color=t["text_main"])
        self.nav_calc_btn.configure(hover_color=t["hover"], text_color=t["text_main"])
        self.nav_reload_btn.configure(hover_color=t["hover"], text_color=t["text_muted"])
        self.status_badge.configure(text_color=t["accent"])
        self.center.configure(fg_color=t["bg_app"])
        for card in (self.card1, self.card2, self.card3):
            card.configure(fg_color=t["bg_card"], border_color=t["border"])
        self.c1_title.configure(text_color=t["title_accent1"])
        self.c2_title.configure(text_color=t["title_accent2"])
        self.c3_title.configure(text_color=t["title_accent3"])
        self.selected_firm_lbl.configure(text_color=t["text_main"])
        self.last_file_lbl.configure(text_color=t["text_muted"])
        self._update_calc_bar_style()
        self.calc_toggle_btn.configure(
            fg_color=t["bg_card"], hover_color=t["hover"], border_color=t["border"], text_color=t["text_main"]
        )
        for ent in (self.calc_search_entry, self.spec_entry, self.inv_entry, self.search_entry):
            ent.configure(fg_color=t["bg_input"], border_color=t["border"], text_color=t["text_main"])
        self.calc_scroll.configure(fg_color=t["bg_input"], border_color=t["border"])
        self.lbl_spec_num.configure(text_color=t["text_muted"])
        self.lbl_inv_num.configure(text_color=t["text_muted"])
        self._update_open_form_btn_style()
        self.log_box.configure(fg_color=t["bg_input"], border_color=t["border"], text_color=t["text_main"])
        self.build_btn.configure(fg_color=t["accent"], hover_color=t["accent_hover"], text_color=t["accent_text"])
        self.clip_btn.configure(
            fg_color=t["bg_card"], hover_color=t["hover"], border_color=t["border"], text_color=t["text_main"]
        )
        self.right_panel.configure(fg_color=t["bg_sidebar"], border_color=t["border"])
        self.firms_header_lbl.configure(text_color=t["title_accent1"])

    def _append_log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{ts}] {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _set_status(self, msg: str, color: str = None):
        c = color or self.t["accent"]
        self.after(0, lambda: self.status_badge.configure(text=f"● {msg}", text_color=c))
        self.after(0, lambda: self._append_log(msg))

    def _render_client_list(self, query: str):
        for widget in self.clients_scroll.winfo_children():
            widget.destroy()
        self.client_buttons.clear()
        t = self.t
        q = query.strip().lower()
        filtered = [c for c in self.all_clients if q in c.lower()]
        self.firms_header_lbl.configure(text=f"Справочник фирм ({len(filtered)})")
        for name in filtered:
            is_sel = name == self.selected_client
            btn = ctk.CTkButton(
                self.clients_scroll, text=name, anchor="w", height=32, corner_radius=8,
                fg_color=t["sel_bg"] if is_sel else t["bg_card"], hover_color=t["hover"],
                border_width=1, border_color=t["accent"] if is_sel else t["border"],
                text_color=t["sel_text"] if is_sel else t["text_main"],
                font=ctk.CTkFont(size=12, weight="bold" if is_sel else "normal"),
                command=lambda n=name: self._select_client(n),
            )
            btn.pack(fill="x", pady=2, padx=4)
            self.client_buttons[name] = btn

    def _render_calc_list(self, query: str):
        for widget in self.calc_scroll.winfo_children():
            widget.destroy()
        self.calc_buttons.clear()
        t = self.t
        q = query.strip().lower()
        q_lat = normalize_latin(q) if q else ""
        if q:
            items = ["(не выбрана)"] + [
                f for f in self.all_calc_folders if q in f.lower() or (q_lat and q_lat in normalize_latin(f))
            ]
        else:
            items = []
            if self.selected_calc_folder and self.selected_calc_folder != "(не выбрана)":
                items.append(self.selected_calc_folder)
            items.append("(не выбрана)")
            items.extend(f for f in self.all_calc_folders if f != self.selected_calc_folder)
        for name in items[:80]:
            is_sel = name == self.selected_calc_folder
            btn = ctk.CTkButton(
                self.calc_scroll, text=name, anchor="w", height=28, corner_radius=6,
                fg_color=t["sel_bg"] if is_sel else "transparent", hover_color=t["hover"],
                text_color=t["sel_text"] if is_sel else t["text_main"],
                font=ctk.CTkFont(size=12, weight="bold" if is_sel else "normal"),
                command=lambda n=name: self._select_calc_folder(n),
            )
            btn.pack(fill="x", pady=1, padx=2)
            self.calc_buttons[name] = btn

    def _update_calc_bar_style(self):
        t = self.t
        is_ok = self.selected_calc_folder and self.selected_calc_folder != "(не выбрана)"
        self.calc_title_lbl.configure(text_color=t["text_muted"])
        self.calc_selected_lbl.configure(
            text=f"📁  {self.selected_calc_folder}" if is_ok else "(не выбрана — нажмите «Изменить папку»)",
            text_color=t["sel_text"] if is_ok else t["text_muted"],
        )
        self.calc_bar.configure(
            fg_color=t["sel_bg"] if is_ok else t["bg_input"],
            border_color=t["accent"] if is_ok else t["border"],
        )

    def _select_calc_folder(self, folder_name: str):
        self.selected_calc_folder = folder_name
        self._update_calc_bar_style()
        if self.selected_client and folder_name != "(не выбрана)":
            self.folder_mapping[self.selected_client] = folder_name
            self.user_settings["folder_mapping"] = self.folder_mapping
            save_user_settings(self.user_settings)
        self._set_calc_picker_visible(False)

    def _select_client(self, client_name: str):
        self.selected_client = client_name
        self.selected_firm_lbl.configure(text=client_name)
        t = self.t
        for name, btn in self.client_buttons.items():
            is_sel = name == client_name
            btn.configure(
                fg_color=t["sel_bg"] if is_sel else t["bg_card"],
                border_color=t["accent"] if is_sel else t["border"],
                text_color=t["sel_text"] if is_sel else t["text_main"],
                font=ctk.CTkFont(size=12, weight="bold" if is_sel else "normal"),
            )
        spec_folder = os.path.join(CFG["specs_root"], client_name)
        try:
            last_num, last_file = find_latest_spec(spec_folder)
            self.spec_num_var.set(str(last_num + 1))
            self.last_file_lbl.configure(text=f"Последний файл: {os.path.basename(last_file)}", text_color=t["text_muted"])
        except Exception as e:
            self.spec_num_var.set("1")
            self.last_file_lbl.configure(text=f"Внимание: {e}", text_color="#ef4444")

        guessed = guess_calc_folder(client_name, self.all_calc_folders, self.folder_mapping)
        self.selected_calc_folder = guessed
        self.calc_search_var.set("")
        self._update_calc_bar_style()
        self._set_calc_picker_visible(guessed == "(не выбрана)")

    def _open_selected_spec_folder(self):
        folder = os.path.join(CFG["specs_root"], self.selected_client) if self.selected_client else CFG["specs_root"]
        if os.path.isdir(folder):
            os.startfile(folder)

    def _open_selected_calc_folder(self):
        calc = self.selected_calc_folder
        folder = os.path.join(CFG["calc_root"], calc) if calc and calc != "(не выбрана)" else CFG["calc_root"]
        if os.path.isdir(folder):
            os.startfile(folder)

    def _reload_all_folders(self):
        self._load_folders()
        self._render_client_list(self.search_var.get())
        if self.calc_picker_open:
            self._render_calc_list(self.calc_search_var.get())
        self._append_log("Списки папок обновлены.")

    def _on_start_build(self, from_clipboard_only: bool = False):
        if not self.selected_client:
            messagebox.showwarning("Внимание", "Сначала выберите фирму в списке справа.")
            return
        inv_num = clean_invoice_number(self.inv_var.get())
        if not inv_num:
            messagebox.showwarning("Внимание", "Введите номер счёта из 1С (например, 412).")
            return
        try:
            new_spec_num = int(self.spec_num_var.get().strip())
        except ValueError:
            messagebox.showwarning("Внимание", "Номер спецификации должен быть числом.")
            return

        calc_folder_name = self.selected_calc_folder
        use_open = self.use_open_form
        self.build_btn.configure(state="disabled", text="ВЫПОЛНЯЕТСЯ СБОРКА...")
        self.clip_btn.configure(state="disabled")
        if not from_clipboard_only:
            self.iconify()

        def worker():
            try:
                res = build_specification(
                    CFG, self.selected_client, calc_folder_name, inv_num, new_spec_num,
                    use_open, from_clipboard_only, lambda m: self._set_status(m, self.t["title_accent1"]),
                )
                self._set_status("Готово! Спецификация открыта в Excel.", self.t["accent"])
                self.after(0, self.deiconify)
                self.after(0, lambda: messagebox.showinfo("Готово", f"Спецификация собрана и открыта в Excel:\n{res}"))
            except Exception as e:
                self._set_status(f"Ошибка: {e}", "#ef4444")
                self.after(0, self.deiconify)
                self.after(0, lambda err=str(e): messagebox.showerror("Ошибка", err))
            finally:
                self.after(0, lambda: self.build_btn.configure(state="normal", text="СОБРАТЬ СПЕЦИФИКАЦИЮ И ОТКРЫТЬ EXCEL"))
                self.after(0, lambda: self.clip_btn.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = SpecBuilderApp()
    app.mainloop()
