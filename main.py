import ctypes
import json
import os
import re
import shutil
import subprocess
import threading
import time
from datetime import datetime
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
from pywinauto import Desktop


# =============================================================================
# НАСТРОЙКИ ПО УМОЛЧАНИЮ (зашиты в EXE)
# =============================================================================
CFG = {
    "company_name": 'ООО "ПЕЧАТНЫЕ ПЛАТЫ"',
    "employee_name": "Мамонтова Людмила Давыдовна",
    "1c_exe": r"C:\Program Files\1cv8\common\1cestart.exe",
    "1c_base_args": 'ENTERPRISE /S"DB01\\pp_buh_new" /N"МамонтоваЛД" /P"mamon123"',
    "specs_root": r"\\Fs01\obmen\ВП\Спецификации 2019 Новая форма",
    "calc_root": r"\\Fs01\obmen\РАСЧЕТЫ",
    "folder_mapping": {
        "Электроприбор": "Elektropribor_K",
    },
}

SETTINGS_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "SpecBuilder")
SETTINGS_FILE = os.path.join(SETTINGS_DIR, "settings.json")


# =============================================================================
# 7 ТЕМ ОФОРМЛЕНИЯ
# =============================================================================
THEMES = {
    "Тёмная ночь": {
        "mode": "dark",
        "bg_app": "#0f1216",
        "bg_sidebar": "#13171d",
        "bg_card": "#181d24",
        "bg_input": "#10141a",
        "border": "#262d38",
        "accent": "#10b981",
        "accent_hover": "#059669",
        "accent_text": "#052e16",
        "sel_bg": "#163529",
        "sel_text": "#34d399",
        "hover": "#1f2937",
        "title_accent1": "#f59e0b",
        "title_accent2": "#38bdf8",
        "title_accent3": "#10b981",
        "text_main": "#f1f5f9",
        "text_muted": "#94a3b8",
        "badge_bg": "#1d2430",
    },
    "Светлая классика": {
        "mode": "light",
        "bg_app": "#f1f5f9",
        "bg_sidebar": "#ffffff",
        "bg_card": "#ffffff",
        "bg_input": "#f8fafc",
        "border": "#cbd5e1",
        "accent": "#16a34a",
        "accent_hover": "#15803d",
        "accent_text": "#ffffff",
        "sel_bg": "#dcfce7",
        "sel_text": "#166534",
        "hover": "#e2e8f0",
        "title_accent1": "#d97706",
        "title_accent2": "#0284c7",
        "title_accent3": "#16a34a",
        "text_main": "#0f172a",
        "text_muted": "#475569",
        "badge_bg": "#f1f5f9",
    },
    "Изумруд СДЭК": {
        "mode": "dark",
        "bg_app": "#0b1410",
        "bg_sidebar": "#0f1c16",
        "bg_card": "#13241c",
        "bg_input": "#0c1712",
        "border": "#1f3b2e",
        "accent": "#1ab248",
        "accent_hover": "#15963c",
        "accent_text": "#04210c",
        "sel_bg": "#163d25",
        "sel_text": "#4ade80",
        "hover": "#1a3327",
        "title_accent1": "#4ade80",
        "title_accent2": "#34d399",
        "title_accent3": "#1ab248",
        "text_main": "#ecfdf5",
        "text_muted": "#86efac",
        "badge_bg": "#163024",
    },
    "Глубокий океан": {
        "mode": "dark",
        "bg_app": "#0a1320",
        "bg_sidebar": "#0e1a2b",
        "bg_card": "#132238",
        "bg_input": "#0b1626",
        "border": "#1e3554",
        "accent": "#0ea5e9",
        "accent_hover": "#0284c7",
        "accent_text": "#082f49",
        "sel_bg": "#13344f",
        "sel_text": "#38bdf8",
        "hover": "#1b2e4a",
        "title_accent1": "#38bdf8",
        "title_accent2": "#818cf8",
        "title_accent3": "#22d3ee",
        "text_main": "#f0f9ff",
        "text_muted": "#94a3b8",
        "badge_bg": "#172942",
    },
    "Графит и золото": {
        "mode": "dark",
        "bg_app": "#17171a",
        "bg_sidebar": "#1e1e22",
        "bg_card": "#242429",
        "bg_input": "#19191d",
        "border": "#36363d",
        "accent": "#eab308",
        "accent_hover": "#ca8a04",
        "accent_text": "#1c1917",
        "sel_bg": "#383018",
        "sel_text": "#facc15",
        "hover": "#2d2d34",
        "title_accent1": "#facc15",
        "title_accent2": "#fbbf24",
        "title_accent3": "#eab308",
        "text_main": "#fafaf9",
        "text_muted": "#a8a29e",
        "badge_bg": "#2b2b31",
    },
    "Чёрное золото": {
        "mode": "dark",
        "bg_app": "#09090b",
        "bg_sidebar": "#111114",
        "bg_card": "#16161a",
        "bg_input": "#0d0d10",
        "border": "#2e2718",
        "accent": "#d4af37",
        "accent_hover": "#b8952a",
        "accent_text": "#09090b",
        "sel_bg": "#2b2412",
        "sel_text": "#f3d267",
        "hover": "#221e15",
        "title_accent1": "#f3d267",
        "title_accent2": "#d4af37",
        "title_accent3": "#eab308",
        "text_main": "#fefce8",
        "text_muted": "#a1a1aa",
        "badge_bg": "#1d1a12",
    },
    "Тёплый песок": {
        "mode": "light",
        "bg_app": "#f5efe6",
        "bg_sidebar": "#fbf8f3",
        "bg_card": "#fffdf9",
        "bg_input": "#f7f1e8",
        "border": "#dfd3c3",
        "accent": "#b45309",
        "accent_hover": "#92400e",
        "accent_text": "#ffffff",
        "sel_bg": "#f3e5d0",
        "sel_text": "#78350f",
        "hover": "#ede2d2",
        "title_accent1": "#b45309",
        "title_accent2": "#0f766e",
        "title_accent3": "#92400e",
        "text_main": "#292524",
        "text_muted": "#78716c",
        "badge_bg": "#f2e8da",
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
# УТИЛИТЫ КЛАВИАТУРЫ И БУФЕРА (БЕЗ ОШИБОК ДЕСКРИПТОРА И РАСКЛАДКИ)
# =============================================================================
class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort),
        ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class _INPUT_UNION(ctypes.Union):
    _fields_ = [("ki", KEYBDINPUT), ("pad", ctypes.c_ubyte * 28)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong), ("union", _INPUT_UNION)]


def type_unicode_text(text: str):
    """Вводит любой текст через SendInput напрямую — без буфера обмена и раскладки."""
    INPUT_KEYBOARD = 1
    KEYEVENTF_UNICODE = 0x0004
    KEYEVENTF_KEYUP = 0x0002
    extra = ctypes.c_ulong(0)

    for ch in text:
        code = ord(ch)
        inp_down = INPUT(type=INPUT_KEYBOARD)
        inp_down.union.ki = KEYBDINPUT(0, code, KEYEVENTF_UNICODE, 0, ctypes.pointer(extra))
        inp_up = INPUT(type=INPUT_KEYBOARD)
        inp_up.union.ki = KEYBDINPUT(0, code, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, 0, ctypes.pointer(extra))
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
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =============================================================================
def resolve_1c_exe(configured_path: str) -> str:
    candidates = [
        configured_path,
        r"C:\Program Files\1cv8\common\1cestart.exe",
        r"C:\Program Files (x86)\1cv8\common\1cestart.exe",
    ]
    for path in candidates:
        if path and os.path.exists(path):
            return path
    raise FileNotFoundError("Не найден 1cestart.exe на этом компьютере.")


def clean_invoice_number(raw: str) -> str:
    s = raw.strip()
    s = re.sub(r"^(?:сч[её]т|сч\.?|№)\s*", "", s, flags=re.IGNORECASE).strip()
    return s


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
    rub_unit = decline(rub, "рубль", "рубля", "рублей")
    kop_unit = decline(kop, "копейка", "копейки", "копеек")
    return f"{rub_words} {rub_unit} {kop:02d} {kop_unit}"


def split_board_and_ps(raw_name: str):
    """
    Разбирает строку из 1С вида:
      'Печ.платы ps 412468 6U7102201v2 (6Ю7.102.201v2)'
      'Печ.платы ps 638725 bvp1'
    Возвращает кортеж: (наименование_платы, зав_номер_заказа).
    """
    s = raw_name.strip()
    m = re.search(r"\bps\s+(\d+)\s+(.+)$", s, flags=re.IGNORECASE)
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
    matched = []
    all_excel = []

    for fname in os.listdir(spec_folder):
        if fname.startswith("~$"):
            continue
        if not fname.lower().endswith((".xls", ".xlsx")):
            continue
        full_path = os.path.join(spec_folder, fname)
        if not os.path.isfile(full_path):
            continue
        mtime = os.path.getmtime(full_path)
        all_excel.append((mtime, full_path))

        for pat in patterns:
            m = pat.search(fname)
            if m:
                matched.append((int(m.group(1)), mtime, full_path))
                break

    if matched:
        matched.sort(key=lambda x: (x[0], x[1]), reverse=True)
        return matched[0][0], matched[0][2]

    if all_excel:
        all_excel.sort(key=lambda x: x[0], reverse=True)
        return 0, all_excel[0][1]

    raise FileNotFoundError(f"В папке {spec_folder} нет файлов Excel (.xls / .xlsx).")


def guess_calc_folder(client_name: str, calc_folders: list, mapping: dict) -> str:
    mapped = mapping.get(client_name) or CFG.get("folder_mapping", {}).get(client_name)
    if mapped and mapped in calc_folders:
        return mapped

    low = client_name.lower()
    for folder in calc_folders:
        if folder.lower() == low:
            return folder
    for folder in calc_folders:
        if low in folder.lower() or folder.lower() in low:
            return folder
    return "(не выбрана)"


# =============================================================================
# РАБОТА С 1С
# =============================================================================
def find_1c_window():
    desktop = Desktop(backend="uia")
    for w in desktop.windows():
        title = w.window_text()
        if "БП_рабочая база" in title or ("1С:Предприятие" in title and "Запуск" not in title):
            return w
    return None


def click_1c_element_by_name(win_1c, exact_or_part: str, timeout: float = 8.0) -> bool:
    deadline = time.time() + timeout
    needle = exact_or_part.lower()
    while time.time() < deadline:
        try:
            for el in win_1c.descendants():
                txt = (el.window_text() or "").strip().lower()
                if txt == needle or (len(needle) >= 5 and needle in txt):
                    rect = el.rectangle()
                    if rect.width() > 5 and rect.height() > 5:
                        el.click_input()
                        return True
        except Exception:
            pass
        time.sleep(0.5)
    return False


def grab_text_from_1c_print_form(cfg: dict, invoice_number: str, use_open_form: bool, log_cb) -> str:
    log_cb("Ищем окно 1С:Предприятие...")
    win_1c = find_1c_window()
    just_launched = False

    if win_1c is None:
        if use_open_form:
            raise RuntimeError(
                "Окно 1С:Предприятие не найдено на этом компьютере.\n"
                "Открой 1С и печатную форму счёта на этом же ПК."
            )
        log_cb("Запускаем 1С и входим в базу...")
        exe_1c = resolve_1c_exe(cfg.get("1c_exe", ""))
        cmd = f'"{exe_1c}" {cfg["1c_base_args"]}'
        subprocess.Popen(cmd, shell=True)
        just_launched = True
        for _ in range(45):
            time.sleep(2.0)
            win_1c = find_1c_window()
            if win_1c:
                break

    if win_1c is None:
        raise RuntimeError(
            "Не удалось найти или запустить окно 1С:Предприятие.\n"
            "Важно: программа должна запускаться на том же компьютере, где установлена 1С."
        )

    if just_launched:
        log_cb("Ждём полную загрузку интерфейса 1С (6 сек)...")
        time.sleep(6.0)

    win_1c.set_focus()
    time.sleep(0.7)

    if not use_open_form:
        log_cb("Нажимаем раздел «Продажи»...")
        if click_1c_element_by_name(win_1c, "Продажи", timeout=10.0):
            time.sleep(1.0)
            log_cb("Открываем «Счета покупателям»...")
            if click_1c_element_by_name(win_1c, "Счета покупателям", timeout=6.0):
                time.sleep(2.0)
            else:
                log_cb("Пункт «Счета покупателям» не найден (возможно, уже открыт).")
        else:
            log_cb("Раздел «Продажи» не найден — пробуем поиск в текущем окне 1С.")

        log_cb(f"Ищем счёт № {invoice_number} (Ctrl+F)...")
        press_ctrl_key(0x46)  # Ctrl+F
        time.sleep(0.9)

        # Вводим номер напрямую через SendInput (без буфера обмена!)
        type_unicode_text(str(invoice_number))
        time.sleep(0.35)
        press_vk(win32con.VK_RETURN)
        time.sleep(1.6)

        log_cb("Открываем найденный счёт (Enter)...")
        press_vk(win32con.VK_RETURN)
        time.sleep(2.5)

        log_cb("Открываем печатную форму «Счет покупателю»...")
        if click_1c_element_by_name(win_1c, "Печать", timeout=6.0):
            time.sleep(0.9)
            if not click_1c_element_by_name(win_1c, "Счет покупателю", timeout=2.5):
                press_vk(win32con.VK_RETURN)
        else:
            press_ctrl_key(0x50)  # Ctrl+P
        time.sleep(2.5)

    log_cb("Копируем таблицу из печатной формы (Ctrl+A → Ctrl+C)...")
    clear_clipboard()
    time.sleep(0.25)

    try:
        rect = win_1c.rectangle()
        cx = int((rect.left + rect.right) * 0.52)
        cy = int((rect.top + rect.bottom) * 0.45)
        win_1c.click_input(coords=(cx - rect.left, cy - rect.top))
        time.sleep(0.4)
    except Exception:
        pass

    press_ctrl_key(0x41)  # Ctrl+A
    time.sleep(0.5)
    press_ctrl_key(0x43)  # Ctrl+C
    time.sleep(0.9)

    raw_text = get_clipboard_text()
    if not raw_text.strip():
        raise RuntimeError(
            "Буфер обмена остался пустым после копирования из 1С.\n\n"
            "Быстрый обход: открой в 1С печатную форму счёта, нажми там Ctrl+A и Ctrl+C, "
            "а затем нажми кнопку «📋 Вставить из буфера»."
        )
    return raw_text


def parse_invoice_clipboard_text(raw_text: str) -> dict:
    invoice_date = datetime.now().strftime("%d.%m.%Y")
    items = []
    totals = {"no_nds": 0.0, "nds": 0.0, "with_nds": 0.0}

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
            day = m.group(1).zfill(2)
            mon = months.get(m.group(2).lower(), "01")
            year = m.group(3)
            invoice_date = f"{day}.{mon}.{year}"

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
                term = cells[6] if len(cells) >= 7 else ""
                items.append({
                    "raw_name": cells[1],
                    "name": board_name,
                    "ps_num": ps_num,
                    "unit": cells[2],
                    "qty": parse_number(cells[3]),
                    "price": parse_number(cells[4]),
                    "sum": parse_number(cells[5]),
                    "term": term,
                })

        if "НДС" in joined and totals["nds"] == 0.0 and "без НДС" not in joined.lower():
            totals["nds"] = parse_number(cells[-1])
        elif "Всего к оплате" in joined:
            totals["with_nds"] = parse_number(cells[-1])

    if totals["no_nds"] == 0.0 and items:
        totals["no_nds"] = round(sum(i["sum"] for i in items), 2)
    if totals["nds"] == 0.0 and totals["no_nds"] > 0:
        totals["nds"] = round(totals["no_nds"] * 0.22, 2)
    if totals["with_nds"] == 0.0 and totals["no_nds"] > 0:
        totals["with_nds"] = round(totals["no_nds"] + totals["nds"], 2)

    return {"date": invoice_date, "items": items, "totals": totals}


# =============================================================================
# ПОИСК НИЖНЕЙ СТРОКИ В ПАПКЕ "РАСЧЕТЫ" (ТОЧНО ПОД СТРОКОЙ "Слои | Класс | Маска | Марк | Покр | Доп")
# =============================================================================
def extract_board_tokens(board_name: str) -> list:
    stop_words = {"печ.платы", "печатные", "платы", "плата", "ps", "пп", "шт"}
    raw_tokens = re.split(r"[\s()\[\],;]+", board_name)
    tokens = []
    for t in raw_tokens:
        clean = t.strip().lower()
        if len(clean) >= 3 and clean not in stop_words:
            tokens.append(clean)
    return tokens


def extract_6_tech_cells_from_rows(rows: list) -> list:
    """
    Ищет строку-заголовок с 'Маска' и ('Слои' или 'Покр') в бланке заказа ПП
    и берёт ровно 6 ячеек из следующей строки (сохраняя пустые ячейки на своих местах).
    """
    for r in range(len(rows) - 1, -1, -1):
        row = rows[r]
        joined = " ".join(str(v or "").lower() for v in row)
        if "маска" in joined and ("сло" in joined or "покр" in joined or "марк" in joined):
            if r + 1 < len(rows):
                start_col = 0
                for c, val in enumerate(row):
                    if "сло" in str(val or "").lower():
                        start_col = c
                        break
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


def find_calc_bottom_row(calc_folder: str, board_name: str, raw_name: str = "") -> list:
    if not calc_folder or not os.path.isdir(calc_folder):
        return [""] * 6

    tokens = extract_board_tokens(board_name + " " + raw_name)
    if not tokens:
        return [""] * 6

    target_file = None
    best_hits = 0

    for fname in os.listdir(calc_folder):
        if fname.startswith("~$"):
            continue
        if not fname.lower().endswith((".xls", ".xlsx")):
            continue
        low_name = fname.lower()
        hits = sum(1 for tok in tokens if tok in low_name)
        if hits > best_hits:
            best_hits = hits
            target_file = os.path.join(calc_folder, fname)

    if not target_file:
        return [""] * 6

    rows = []
    if target_file.lower().endswith(".xls"):
        wb = xlrd.open_workbook(target_file)
        sh = wb.sheet_by_index(0)
        for r in range(sh.nrows):
            rows.append([sh.cell_value(r, c) for c in range(sh.ncols)])
    else:
        wb = openpyxl.load_workbook(target_file, data_only=True)
        ws = wb.active
        for r in range(1, ws.max_row + 1):
            rows.append([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])

    return extract_6_tech_cells_from_rows(rows)


# =============================================================================
# СБОРКА ИТОГОВОГО EXCEL-ФАЙЛА СПЕЦИФИКАЦИИ
# =============================================================================
def build_specification(
    cfg: dict,
    client_name: str,
    calc_folder_name: str,
    invoice_num: str,
    new_spec_num: int,
    use_open_form: bool,
    from_clipboard_only: bool,
    log_cb,
) -> str:
    spec_folder = os.path.join(cfg["specs_root"], client_name)
    calc_folder = (
        os.path.join(cfg["calc_root"], calc_folder_name)
        if calc_folder_name and calc_folder_name != "(не выбрана)"
        else ""
    )

    log_cb("Ищем последнюю спецификацию и делаем копию...")
    _, last_file = find_latest_spec(spec_folder)
    ext = os.path.splitext(last_file)[1]
    new_file = os.path.join(spec_folder, f"сп.{new_spec_num} сч.{invoice_num}{ext}")
    shutil.copy2(last_file, new_file)

    if from_clipboard_only:
        log_cb("Читаем таблицу из буфера обмена...")
        raw_text = get_clipboard_text()
        if not raw_text.strip():
            raise RuntimeError(
                "Буфер обмена пуст. Сначала скопируй печатную форму счёта в 1С (Ctrl+A, Ctrl+C)."
            )
    else:
        raw_text = grab_text_from_1c_print_form(cfg, invoice_num, use_open_form, log_cb)

    log_cb("Разбираем строки счёта...")
    inv = parse_invoice_clipboard_text(raw_text)
    items = inv["items"]
    if not items:
        raise RuntimeError(
            "Не удалось распознать товары в скопированном тексте 1С.\n"
            "Убедись, что открыта именно печатная форма «Счет покупателю»."
        )

    log_cb(f"Найдено позиций: {len(items)}. Заполняем Excel и подтягиваем РАСЧЕТЫ...")
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb = excel.Workbooks.Open(new_file)
        ws = wb.Sheets(1)

        ws.Range("O1").Value = f"Приложение № {new_spec_num}"
        ws.Range("A4").Value = f"Спецификация № {new_spec_num} от {inv['date']} г."

        start_row = 9
        total_row = start_row
        while total_row < 150:
            val = str(ws.Cells(total_row, 1).Value or "")
            if "Итого без НДС" in val:
                break
            total_row += 1

        existing_rows = total_row - start_row
        needed_rows = len(items)

        if needed_rows > existing_rows:
            for _ in range(needed_rows - existing_rows):
                ws.Rows(start_row + 1).Insert()
                ws.Rows(start_row).Copy(ws.Rows(start_row + 1))
        elif needed_rows < existing_rows:
            for _ in range(existing_rows - needed_rows):
                ws.Rows(start_row + needed_rows).Delete()

        for idx, item in enumerate(items, start=1):
            r = start_row + idx - 1
            ws.Cells(r, 1).Value = idx                                   # A: № п/п
            ws.Cells(r, 2).Value = item["name"]                          # B: Наименование платы
            ws.Cells(r, 3).Value = item["ps_num"]                        # C: Зав. № заказа (Пэсочка)
            ws.Cells(r, 4).Value = item["term"]                          # D: Срок пост.
            ws.Cells(r, 5).Value = item["unit"]                          # E: Ед. изм.
            ws.Cells(r, 6).Value = item["qty"]                           # F: Кол-во
            ws.Cells(r, 7).Value = item["price"]                         # G: Цена
            ws.Cells(r, 8).Value = format_rub_text(item["sum"], False)   # H: Сумма
            ws.Cells(r, 9).Value = ""                                    # I: Площадь (с заводского счёта)

            tech = find_calc_bottom_row(calc_folder, item["name"], item["raw_name"])
            for offset, val in enumerate(tech):
                ws.Cells(r, 10 + offset).Value = val

        new_total_row = start_row + needed_rows
        ws.Cells(new_total_row, 8).Value = format_rub_text(inv["totals"]["no_nds"], True)
        ws.Cells(new_total_row + 2, 8).Value = format_rub_text(inv["totals"]["nds"], True)
        ws.Cells(new_total_row + 3, 8).Value = format_rub_text(inv["totals"]["with_nds"], True)
        ws.Cells(new_total_row + 5, 3).Value = amount_in_words(inv["totals"]["with_nds"])
        ws.Cells(new_total_row + 6, 3).Value = amount_in_words(inv["totals"]["nds"])

        wb.Save()
    finally:
        pythoncom.CoUninitialize()

    return new_file


# =============================================================================
# ПРИЛОЖЕНИЕ С 7 ТЕМАМИ И ЖИВЫМ ПОИСКОМ ПО «СПЕЦИФИКАЦИЯМ» И «РАСЧЁТАМ»
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
        self.geometry("1220x740")
        self.minsize(1020, 650)

        self.all_clients = []
        self.all_calc_folders = []
        self.selected_client = None
        self.selected_calc_folder = "(не выбрана)"
        self.client_buttons = {}
        self.calc_buttons = {}

        self._load_folders()
        self._build_layout()
        self._apply_theme(self.current_theme_name, save=False)
        self._render_client_list("")
        self._render_calc_list("")

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

        # === ВЕРХНЯЯ ШАПКА ===
        self.header = ctk.CTkFrame(self, height=64, corner_radius=0, border_width=1)
        self.header.grid(row=0, column=0, columnspan=3, sticky="ew")
        self.header.grid_columnconfigure(1, weight=1)

        brand_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        brand_frame.grid(row=0, column=0, padx=18, pady=10, sticky="w")

        self.logo_badge = ctk.CTkLabel(
            brand_frame,
            text="ПП",
            width=40,
            height=40,
            corner_radius=10,
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
        )
        self.logo_badge.pack(side="left", padx=(0, 12))

        title_box = ctk.CTkFrame(brand_frame, fg_color="transparent")
        title_box.pack(side="left")
        self.company_lbl = ctk.CTkLabel(
            title_box,
            text=CFG["company_name"],
            font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
        )
        self.company_lbl.pack(anchor="w")
        self.subtitle_lbl = ctk.CTkLabel(
            title_box,
            text="Автоматическая сборка спецификаций (1С + Расчёты)",
            font=ctk.CTkFont(family="Segoe UI", size=11),
        )
        self.subtitle_lbl.pack(anchor="w")

        # СПРАВА В ШАПКЕ: Выбор темы + Мамонтова Людмила Давыдовна
        right_header = ctk.CTkFrame(self.header, fg_color="transparent")
        right_header.grid(row=0, column=2, padx=18, pady=10, sticky="e")

        self.theme_menu = ctk.CTkOptionMenu(
            right_header,
            values=list(THEMES.keys()),
            command=self._on_theme_change,
            width=175,
            height=34,
            corner_radius=8,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
        )
        self.theme_menu.set(self.current_theme_name)
        self.theme_menu.pack(side="left", padx=(0, 12))

        self.user_badge = ctk.CTkFrame(right_header, corner_radius=10, border_width=1)
        self.user_badge.pack(side="left")
        self.user_dot = ctk.CTkLabel(self.user_badge, text="●", font=ctk.CTkFont(size=12))
        self.user_dot.pack(side="left", padx=(12, 6), pady=7)
        self.user_lbl = ctk.CTkLabel(
            self.user_badge,
            text=CFG["employee_name"],
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
        )
        self.user_lbl.pack(side="left", padx=(0, 14), pady=7)

        # === ЛЕВАЯ ПАНЕЛЬ НАВИГАЦИИ ===
        self.left_nav = ctk.CTkFrame(self, width=185, corner_radius=0, border_width=1)
        self.left_nav.grid(row=1, column=0, sticky="ns")
        self.left_nav.grid_propagate(False)

        self.nav_title = ctk.CTkLabel(self.left_nav, text="РАЗДЕЛЫ", font=ctk.CTkFont(size=11, weight="bold"))
        self.nav_title.pack(anchor="w", padx=16, pady=(18, 8))

        self.nav_active_btn = ctk.CTkButton(
            self.left_nav,
            text="📄  Спецификация",
            anchor="w",
            height=40,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.nav_active_btn.pack(fill="x", padx=12, pady=4)

        self.nav_spec_btn = ctk.CTkButton(
            self.left_nav,
            text="📁  Папка Спец.",
            anchor="w",
            height=38,
            fg_color="transparent",
            font=ctk.CTkFont(size=13),
            command=self._open_selected_spec_folder,
        )
        self.nav_spec_btn.pack(fill="x", padx=12, pady=4)

        self.nav_calc_btn = ctk.CTkButton(
            self.left_nav,
            text="📊  Папка Расчёты",
            anchor="w",
            height=38,
            fg_color="transparent",
            font=ctk.CTkFont(size=13),
            command=self._open_selected_calc_folder,
        )
        self.nav_calc_btn.pack(fill="x", padx=12, pady=4)

        self.nav_reload_btn = ctk.CTkButton(
            self.left_nav,
            text="🔄  Обновить списки",
            anchor="w",
            height=38,
            fg_color="transparent",
            font=ctk.CTkFont(size=12),
            command=self._reload_all_folders,
        )
        self.nav_reload_btn.pack(fill="x", padx=12, pady=(16, 4))

        self.status_badge = ctk.CTkLabel(
            self.left_nav,
            text="● Готов к работе",
            font=ctk.CTkFont(size=12, weight="bold"),
            wraplength=155,
            justify="left",
        )
        self.status_badge.pack(side="bottom", anchor="w", padx=14, pady=16)

        # === ЦЕНТРАЛЬНАЯ РАБОЧАЯ ОБЛАСТЬ ===
        self.center = ctk.CTkScrollableFrame(self, corner_radius=0)
        self.center.grid(row=1, column=1, sticky="nsew", padx=14, pady=12)

        # Карточка 1: Фирма и поиск папки в РАСЧЕТЫ
        self.card1 = ctk.CTkFrame(self.center, corner_radius=12, border_width=1)
        self.card1.pack(fill="x", pady=(0, 12))

        self.c1_title = ctk.CTkLabel(
            self.card1, text="1. ФИРМА, ШАБЛОН И ПАПКА В «РАСЧЕТЫ»",
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.c1_title.pack(anchor="w", padx=16, pady=(14, 6))

        self.selected_firm_lbl = ctk.CTkLabel(
            self.card1, text="Выберите фирму в списке справа →",
            font=ctk.CTkFont(size=16, weight="bold"),
        )
        self.selected_firm_lbl.pack(anchor="w", padx=16, pady=(0, 2))

        self.last_file_lbl = ctk.CTkLabel(
            self.card1, text="Последний файл: —",
            font=ctk.CTkFont(size=12),
        )
        self.last_file_lbl.pack(anchor="w", padx=16, pady=(0, 10))

        calc_top = ctk.CTkFrame(self.card1, fg_color="transparent")
        calc_top.pack(fill="x", padx=16, pady=(0, 4))
        self.calc_title_lbl = ctk.CTkLabel(
            calc_top, text="Папка фирмы в «РАСЧЕТЫ» (выбрано: ", font=ctk.CTkFont(size=12),
        )
        self.calc_title_lbl.pack(side="left")
        self.calc_selected_lbl = ctk.CTkLabel(
            calc_top, text="(не выбрана)", font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.calc_selected_lbl.pack(side="left")
        self.calc_close_lbl = ctk.CTkLabel(calc_top, text=")", font=ctk.CTkFont(size=12))
        self.calc_close_lbl.pack(side="left")

        self.calc_search_var = ctk.StringVar()
        self.calc_search_var.trace_add("write", lambda *_: self._render_calc_list(self.calc_search_var.get()))
        self.calc_search_entry = ctk.CTkEntry(
            self.card1,
            textvariable=self.calc_search_var,
            placeholder_text="🔍 Поиск папки фирмы в «РАСЧЕТЫ»...",
            height=34,
        )
        self.calc_search_entry.pack(fill="x", padx=16, pady=(0, 6))

        self.calc_scroll = ctk.CTkScrollableFrame(self.card1, height=110, corner_radius=8, border_width=1)
        self.calc_scroll.pack(fill="x", padx=16, pady=(0, 14))

        # Карточка 2: Номер спецификации и счёт
        self.card2 = ctk.CTkFrame(self.center, corner_radius=12, border_width=1)
        self.card2.pack(fill="x", pady=(0, 12))

        self.c2_title = ctk.CTkLabel(
            self.card2, text="2. ПАРАМЕТРЫ СПЕЦИФИКАЦИИ И СЧЁТА 1С",
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.c2_title.pack(anchor="w", padx=16, pady=(14, 10))

        grid2 = ctk.CTkFrame(self.card2, fg_color="transparent")
        grid2.pack(fill="x", padx=16, pady=(0, 12))
        grid2.grid_columnconfigure(0, weight=1)
        grid2.grid_columnconfigure(1, weight=1)

        self.lbl_spec_num = ctk.CTkLabel(grid2, text="Новый № спецификации:", font=ctk.CTkFont(size=12))
        self.lbl_spec_num.grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 4))
        self.lbl_inv_num = ctk.CTkLabel(grid2, text="Номер счёта в 1С (например, 319):", font=ctk.CTkFont(size=12))
        self.lbl_inv_num.grid(row=0, column=1, sticky="w", padx=(8, 0), pady=(0, 4))

        self.spec_num_var = ctk.StringVar()
        self.spec_entry = ctk.CTkEntry(
            grid2, textvariable=self.spec_num_var, height=38, font=ctk.CTkFont(size=14, weight="bold")
        )
        self.spec_entry.grid(row=1, column=0, sticky="ew", padx=(0, 8))

        self.inv_var = ctk.StringVar()
        self.inv_entry = ctk.CTkEntry(
            grid2, textvariable=self.inv_var, placeholder_text="319", height=38, font=ctk.CTkFont(size=14, weight="bold")
        )
        self.inv_entry.grid(row=1, column=1, sticky="ew", padx=(8, 0))

        self.use_open_form_var = ctk.BooleanVar(value=False)
        self.open_form_chk = ctk.CTkCheckBox(
            self.card2,
            text="Печатная форма счёта уже открыта в 1С (забрать сразу, без поиска)",
            variable=self.use_open_form_var,
            font=ctk.CTkFont(size=12),
        )
        self.open_form_chk.pack(anchor="w", padx=16, pady=(0, 14))

        # Карточка 3: Журнал выполнения
        self.card3 = ctk.CTkFrame(self.center, corner_radius=12, border_width=1)
        self.card3.pack(fill="both", expand=True, pady=(0, 12))

        self.c3_title = ctk.CTkLabel(
            self.card3, text="3. ХОД ВЫПОЛНЕНИЯ",
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.c3_title.pack(anchor="w", padx=16, pady=(14, 6))

        self.log_box = ctk.CTkTextbox(
            self.card3, height=115, border_width=1, font=ctk.CTkFont(family="Consolas", size=12)
        )
        self.log_box.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        self._append_log("Готово. Выберите фирму справа, при необходимости найдите папку в «РАСЧЕТЫ» и введите № счёта.")

        # Кнопки запуска
        btn_row = ctk.CTkFrame(self.center, fg_color="transparent")
        btn_row.pack(fill="x", pady=(4, 8))
        btn_row.grid_columnconfigure(0, weight=3)
        btn_row.grid_columnconfigure(1, weight=1)

        self.build_btn = ctk.CTkButton(
            btn_row,
            text="СОБРАТЬ СПЕЦИФИКАЦИЮ И ОТКРЫТЬ EXCEL",
            height=48,
            corner_radius=10,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            command=lambda: self._on_start_build(from_clipboard_only=False),
        )
        self.build_btn.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.clip_btn = ctk.CTkButton(
            btn_row,
            text="📋 Вставить из буфера",
            height=48,
            corner_radius=10,
            border_width=1,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            command=lambda: self._on_start_build(from_clipboard_only=True),
        )
        self.clip_btn.grid(row=0, column=1, sticky="ew")

        # === ПРАВАЯ ПАНЕЛЬ: СПРАВОЧНИК ФИРМ ===
        self.right_panel = ctk.CTkFrame(self, width=300, corner_radius=0, border_width=1)
        self.right_panel.grid(row=1, column=2, sticky="ns")
        self.right_panel.grid_propagate(False)

        self.firms_header_lbl = ctk.CTkLabel(
            self.right_panel,
            text=f"Справочник фирм ({len(self.all_clients)})",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.firms_header_lbl.pack(anchor="w", padx=14, pady=(14, 8))

        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._render_client_list(self.search_var.get()))
        self.search_entry = ctk.CTkEntry(
            self.right_panel,
            textvariable=self.search_var,
            placeholder_text="🔍 Поиск фирмы...",
            height=36,
        )
        self.search_entry.pack(fill="x", padx=14, pady=(0, 10))

        self.clients_scroll = ctk.CTkScrollableFrame(self.right_panel, fg_color="transparent")
        self.clients_scroll.pack(fill="both", expand=True, padx=8, pady=(0, 10))

    # =========================================================================
    # ПРИМЕНЕНИЕ ТЕМЫ
    # =========================================================================
    def _on_theme_change(self, theme_name: str):
        self._apply_theme(theme_name, save=True)
        self._render_client_list(self.search_var.get())
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
            fg_color=t["badge_bg"],
            button_color=t["accent"],
            button_hover_color=t["accent_hover"],
            text_color=t["text_main"],
            dropdown_fg_color=t["bg_card"],
            dropdown_hover_color=t["hover"],
            dropdown_text_color=t["text_main"],
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
        self.calc_title_lbl.configure(text_color=t["text_muted"])
        self.calc_selected_lbl.configure(text_color=t["accent"])
        self.calc_close_lbl.configure(text_color=t["text_muted"])

        for ent in (self.calc_search_entry, self.spec_entry, self.inv_entry, self.search_entry):
            ent.configure(fg_color=t["bg_input"], border_color=t["border"], text_color=t["text_main"])

        self.calc_scroll.configure(fg_color=t["bg_input"], border_color=t["border"])
        self.lbl_spec_num.configure(text_color=t["text_muted"])
        self.lbl_inv_num.configure(text_color=t["text_muted"])

        self.open_form_chk.configure(
            text_color=t["text_main"], fg_color=t["accent"], hover_color=t["accent_hover"]
        )
        self.log_box.configure(fg_color=t["bg_input"], border_color=t["border"], text_color=t["text_main"])

        self.build_btn.configure(
            fg_color=t["accent"], hover_color=t["accent_hover"], text_color=t["accent_text"]
        )
        self.clip_btn.configure(
            fg_color=t["bg_card"], hover_color=t["hover"], border_color=t["border"], text_color=t["text_main"]
        )

        self.right_panel.configure(fg_color=t["bg_sidebar"], border_color=t["border"])
        self.firms_header_lbl.configure(text_color=t["title_accent1"])

    # =========================================================================
    # СПИСКИ ФИРМ И ПАПОК «РАСЧЕТЫ»
    # =========================================================================
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
                self.clients_scroll,
                text=name,
                anchor="w",
                height=34,
                corner_radius=8,
                fg_color=t["sel_bg"] if is_sel else t["bg_card"],
                hover_color=t["hover"],
                border_width=1,
                border_color=t["accent"] if is_sel else t["border"],
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
        items = ["(не выбрана)"] + [f for f in self.all_calc_folders if q in f.lower()]
        for name in items[:80]:
            is_sel = name == self.selected_calc_folder
            btn = ctk.CTkButton(
                self.calc_scroll,
                text=name,
                anchor="w",
                height=28,
                corner_radius=6,
                fg_color=t["sel_bg"] if is_sel else "transparent",
                hover_color=t["hover"],
                text_color=t["sel_text"] if is_sel else t["text_main"],
                font=ctk.CTkFont(size=12, weight="bold" if is_sel else "normal"),
                command=lambda n=name: self._select_calc_folder(n),
            )
            btn.pack(fill="x", pady=1, padx=2)
            self.calc_buttons[name] = btn

    def _select_calc_folder(self, folder_name: str):
        self.selected_calc_folder = folder_name
        self.calc_selected_lbl.configure(text=folder_name)

        if self.selected_client and folder_name != "(не выбрана)":
            self.folder_mapping[self.selected_client] = folder_name
            self.user_settings["folder_mapping"] = self.folder_mapping
            save_user_settings(self.user_settings)

        t = self.t
        for name, btn in self.calc_buttons.items():
            is_sel = name == folder_name
            btn.configure(
                fg_color=t["sel_bg"] if is_sel else "transparent",
                text_color=t["sel_text"] if is_sel else t["text_main"],
                font=ctk.CTkFont(size=12, weight="bold" if is_sel else "normal"),
            )

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
            self.last_file_lbl.configure(
                text=f"Последний файл: {os.path.basename(last_file)}",
                text_color=t["text_muted"],
            )
        except Exception as e:
            self.spec_num_var.set("1")
            self.last_file_lbl.configure(text=f"Внимание: {e}", text_color="#ef4444")

        guessed = guess_calc_folder(client_name, self.all_calc_folders, self.folder_mapping)
        self.selected_calc_folder = guessed
        self.calc_selected_lbl.configure(text=guessed)
        self.calc_search_var.set("")
        self._render_calc_list("")

    def _open_selected_spec_folder(self):
        if not self.selected_client:
            os.startfile(CFG["specs_root"])
            return
        folder = os.path.join(CFG["specs_root"], self.selected_client)
        if os.path.isdir(folder):
            os.startfile(folder)

    def _open_selected_calc_folder(self):
        calc = self.selected_calc_folder
        if calc and calc != "(не выбрана)":
            folder = os.path.join(CFG["calc_root"], calc)
            if os.path.isdir(folder):
                os.startfile(folder)
                return
        if os.path.isdir(CFG["calc_root"]):
            os.startfile(CFG["calc_root"])

    def _reload_all_folders(self):
        self._load_folders()
        self._render_client_list(self.search_var.get())
        self._render_calc_list(self.calc_search_var.get())
        self._append_log("Списки папок обновлены.")

    def _on_start_build(self, from_clipboard_only: bool = False):
        if not self.selected_client:
            messagebox.showwarning("Внимание", "Сначала выберите фирму в списке справа.")
            return

        inv_num = clean_invoice_number(self.inv_var.get())
        if not inv_num:
            messagebox.showwarning("Внимание", "Введите номер счёта из 1С (например, 319).")
            return

        try:
            new_spec_num = int(self.spec_num_var.get().strip())
        except ValueError:
            messagebox.showwarning("Внимание", "Номер спецификации должен быть числом.")
            return

        calc_folder_name = self.selected_calc_folder
        use_open = self.use_open_form_var.get()
        self.build_btn.configure(state="disabled", text="ВЫПОЛНЯЕТСЯ СБОРКА...")
        self.clip_btn.configure(state="disabled")

        def worker():
            try:
                res = build_specification(
                    CFG,
                    self.selected_client,
                    calc_folder_name,
                    inv_num,
                    new_spec_num,
                    use_open,
                    from_clipboard_only,
                    lambda m: self._set_status(m, self.t["title_accent1"]),
                )
                self._set_status("Готово! Спецификация открыта в Excel.", self.t["accent"])
                self.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Готово", f"Спецификация собрана и открыта в Excel:\n{res}"
                    ),
                )
            except Exception as e:
                self._set_status(f"Ошибка: {e}", "#ef4444")
                self.after(0, lambda err=str(e): messagebox.showerror("Ошибка", err))
            finally:
                self.after(
                    0,
                    lambda: self.build_btn.configure(
                        state="normal", text="СОБРАТЬ СПЕЦИФИКАЦИЮ И ОТКРЫТЬ EXCEL"
                    ),
                )
                self.after(0, lambda: self.clip_btn.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = SpecBuilderApp()
    app.mainloop()
