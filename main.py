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
import win32clipboard
import win32com.client
import xlrd
from num2words import num2words
from pywinauto import Desktop, keyboard


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

# Цветовая палитра (тёмный стиль)
BG_APP = "#0f1216"
BG_SIDEBAR = "#13171d"
BG_CARD = "#181d24"
BG_INPUT = "#10141a"
BORDER_CARD = "#262d38"
ACCENT_GREEN = "#10b981"
ACCENT_GREEN_HOVER = "#059669"
ACCENT_AMBER = "#f59e0b"
ACCENT_BLUE = "#38bdf8"
TEXT_MAIN = "#f1f5f9"
TEXT_MUTED = "#94a3b8"


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


def get_clipboard_text() -> str:
    text = ""
    try:
        win32clipboard.OpenClipboard()
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
            text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass
    return text or ""


def clear_clipboard():
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass


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


def guess_calc_folder(client_name: str, calc_folders: list) -> str:
    mapped = CFG.get("folder_mapping", {}).get(client_name)
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
# РАБОТА С 1С (ЧЕРЕЗ БУФЕР ОБМЕНА ПЕЧАТНОЙ ФОРМЫ — БЕЗ ОШИБОК .MXL/.XLSX)
# =============================================================================
def find_1c_window():
    desktop = Desktop(backend="uia")
    for w in desktop.windows():
        title = w.window_text()
        if "БП_рабочая база" in title or ("1С:Предприятие" in title and "Запуск" not in title):
            return w
    return None


def grab_text_from_1c_print_form(cfg: dict, invoice_number: str, use_open_form: bool, log_cb) -> str:
    log_cb("Ищем окно 1С:Предприятие...")
    win_1c = find_1c_window()

    if win_1c is None:
        if use_open_form:
            raise RuntimeError(
                "Окно 1С:Предприятие не найдено на этом компьютере.\n"
                "Убедись, что программа запущена на том же ПК, где открыта 1С (не через AnyDesk)."
            )
        log_cb("Запускаем 1С и входим в базу...")
        exe_1c = resolve_1c_exe(cfg.get("1c_exe", ""))
        cmd = f'"{exe_1c}" {cfg["1c_base_args"]}'
        subprocess.Popen(cmd, shell=True)
        for _ in range(45):
            time.sleep(2.0)
            win_1c = find_1c_window()
            if win_1c:
                break

    if win_1c is None:
        raise RuntimeError(
            "Не удалось найти или запустить окно 1С:Предприятие.\n"
            "Важно: SpecBuilder.exe должен запускаться на том же компьютере, где стоит 1С."
        )

    win_1c.set_focus()
    time.sleep(0.6)

    if not use_open_form:
        log_cb("Открываем «Продажи → Счета покупателям»...")
        try:
            win_1c.child_window(title="Продажи", control_type="TabItem").click_input()
            time.sleep(0.5)
            win_1c.child_window(title="Счета покупателям").click_input()
            time.sleep(1.2)
        except Exception:
            pass

        log_cb(f"Ищем счёт № {invoice_number} в 1С...")
        keyboard.send_keys("^f")
        time.sleep(0.6)
        keyboard.send_keys(str(invoice_number), with_spaces=True)
        keyboard.send_keys("{ENTER}")
        time.sleep(1.2)

        keyboard.send_keys("{ENTER}")
        time.sleep(1.8)

        log_cb("Открываем печатную форму «Счет покупателю»...")
        try:
            win_1c.child_window(title_re=".*Печать.*").click_input()
            time.sleep(0.6)
            keyboard.send_keys("{ENTER}")
        except Exception:
            keyboard.send_keys("^p")
        time.sleep(2.0)

    log_cb("Считываем таблицу из печатной формы 1С...")
    clear_clipboard()
    time.sleep(0.2)

    try:
        rect = win_1c.rectangle()
        cx = (rect.left + rect.right) // 2
        cy = (rect.top + rect.bottom) // 2
        win_1c.click_input(coords=(cx - rect.left, cy - rect.top))
        time.sleep(0.3)
    except Exception:
        pass

    keyboard.send_keys("^a")
    time.sleep(0.4)
    keyboard.send_keys("^c")
    time.sleep(0.7)

    raw_text = get_clipboard_text()
    if not raw_text.strip():
        raise RuntimeError(
            "Не удалось скопировать таблицу из 1С.\n"
            "Попробуй открыть печатную форму счёта в 1С вручную и поставить галочку "
            "«Печатная форма счёта уже открыта в 1С»."
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
                term = cells[6] if len(cells) >= 7 else ""
                items.append({
                    "name": cells[1],
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
# ПОИСК НИЖНЕЙ СТРОКИ В ПАПКЕ "РАСЧЕТЫ"
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


def find_calc_bottom_row(calc_folder: str, board_name: str) -> list:
    if not calc_folder or not os.path.isdir(calc_folder):
        return [""] * 6

    tokens = extract_board_tokens(board_name)
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

    if target_file.lower().endswith(".xls"):
        wb = xlrd.open_workbook(target_file)
        sh = wb.sheet_by_index(0)
        for r in range(sh.nrows - 1, -1, -1):
            row_vals = [sh.cell_value(r, c) for c in range(sh.ncols)]
            non_empty = [v for v in row_vals if str(v).strip() != ""]
            if len(non_empty) >= 3:
                return (non_empty + [""] * 6)[:6]
    else:
        wb = openpyxl.load_workbook(target_file, data_only=True)
        ws = wb.active
        for r in range(ws.max_row, 0, -1):
            row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            non_empty = [v for v in row_vals if v is not None and str(v).strip() != ""]
            if len(non_empty) >= 3:
                return (non_empty + [""] * 6)[:6]

    return [""] * 6


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

    raw_text = grab_text_from_1c_print_form(cfg, invoice_num, use_open_form, log_cb)

    log_cb("Разбираем строки счёта...")
    inv = parse_invoice_clipboard_text(raw_text)
    items = inv["items"]
    if not items:
        raise RuntimeError(
            "Не удалось распознать товары в скопированном тексте 1С.\n"
            "Убедись, что в 1С открыта именно печатная форма «Счет покупателю»."
        )

    log_cb("Заполняем Excel и подтягиваем параметры из РАСЧЕТЫ...")
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
            ws.Cells(r, 2).Value = item["name"]                          # B: Наименование
            ws.Cells(r, 3).Value = ""                                    # C: Зав. № заказа
            ws.Cells(r, 4).Value = item["term"]                          # D: Срок пост.
            ws.Cells(r, 5).Value = item["unit"]                          # E: Ед. изм.
            ws.Cells(r, 6).Value = item["qty"]                           # F: Кол-во
            ws.Cells(r, 7).Value = item["price"]                         # G: Цена
            ws.Cells(r, 8).Value = format_rub_text(item["sum"], False)   # H: Сумма
            ws.Cells(r, 9).Value = ""                                    # I: Площадь

            tech = find_calc_bottom_row(calc_folder, item["name"])
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
# СОВРЕМЕННЫЙ ТЁМНЫЙ ИНТЕРФЕЙС (CUSTOMTKINTER)
# =============================================================================
class SpecBuilderApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        self.title(f'{CFG["company_name"]} — Сборка спецификаций')
        self.geometry("1160x700")
        self.minsize(980, 620)
        self.configure(fg_color=BG_APP)

        self.all_clients = []
        self.all_calc_folders = []
        self.selected_client = None
        self.client_buttons = {}

        self._load_folders()
        self._build_layout()
        self._render_client_list("")

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
        header = ctk.CTkFrame(
            self, height=64, fg_color=BG_SIDEBAR, corner_radius=0,
            border_width=1, border_color=BORDER_CARD
        )
        header.grid(row=0, column=0, columnspan=3, sticky="ew")
        header.grid_columnconfigure(1, weight=1)

        brand_frame = ctk.CTkFrame(header, fg_color="transparent")
        brand_frame.grid(row=0, column=0, padx=18, pady=10, sticky="w")

        ctk.CTkLabel(
            brand_frame,
            text="ПП",
            width=40,
            height=40,
            corner_radius=10,
            fg_color=ACCENT_GREEN,
            text_color="#052e16",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
        ).pack(side="left", padx=(0, 12))

        title_box = ctk.CTkFrame(brand_frame, fg_color="transparent")
        title_box.pack(side="left")
        ctk.CTkLabel(
            title_box,
            text=CFG["company_name"],
            font=ctk.CTkFont(family="Segoe UI", size=17, weight="bold"),
            text_color=TEXT_MAIN,
        ).pack(anchor="w")
        ctk.CTkLabel(
            title_box,
            text="Автоматическая сборка спецификаций (1С + Расчёты)",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=TEXT_MUTED,
        ).pack(anchor="w")

        # СПРАВА: Мамонтова Людмила Давыдовна
        user_badge = ctk.CTkFrame(
            header,
            fg_color="#1d2430",
            corner_radius=10,
            border_width=1,
            border_color="#2f3a4c",
        )
        user_badge.grid(row=0, column=2, padx=18, pady=10, sticky="e")
        ctk.CTkLabel(
            user_badge, text="●", text_color=ACCENT_GREEN, font=ctk.CTkFont(size=12)
        ).pack(side="left", padx=(12, 6), pady=8)
        ctk.CTkLabel(
            user_badge,
            text=CFG["employee_name"],
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=TEXT_MAIN,
        ).pack(side="left", padx=(0, 14), pady=8)

        # === ЛЕВАЯ ПАНЕЛЬ НАВИГАЦИИ ===
        left_nav = ctk.CTkFrame(
            self, width=185, fg_color=BG_SIDEBAR, corner_radius=0,
            border_width=1, border_color=BORDER_CARD
        )
        left_nav.grid(row=1, column=0, sticky="ns")
        left_nav.grid_propagate(False)

        ctk.CTkLabel(
            left_nav, text="РАЗДЕЛЫ", font=ctk.CTkFont(size=11, weight="bold"), text_color=TEXT_MUTED
        ).pack(anchor="w", padx=16, pady=(18, 8))

        ctk.CTkButton(
            left_nav,
            text="📄  Спецификация",
            anchor="w",
            height=40,
            fg_color="#163529",
            hover_color="#1d4737",
            text_color="#34d399",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(fill="x", padx=12, pady=4)

        ctk.CTkButton(
            left_nav,
            text="📁  Папка Спец.",
            anchor="w",
            height=38,
            fg_color="transparent",
            hover_color="#1e2530",
            text_color=TEXT_MAIN,
            font=ctk.CTkFont(size=13),
            command=self._open_selected_spec_folder,
        ).pack(fill="x", padx=12, pady=4)

        ctk.CTkButton(
            left_nav,
            text="📊  Папка Расчёты",
            anchor="w",
            height=38,
            fg_color="transparent",
            hover_color="#1e2530",
            text_color=TEXT_MAIN,
            font=ctk.CTkFont(size=13),
            command=self._open_selected_calc_folder,
        ).pack(fill="x", padx=12, pady=4)

        ctk.CTkButton(
            left_nav,
            text="🔄  Обновить списки",
            anchor="w",
            height=38,
            fg_color="transparent",
            hover_color="#1e2530",
            text_color=TEXT_MUTED,
            font=ctk.CTkFont(size=12),
            command=self._reload_all_folders,
        ).pack(fill="x", padx=12, pady=(16, 4))

        self.status_badge = ctk.CTkLabel(
            left_nav,
            text="● Готов к работе",
            text_color=ACCENT_GREEN,
            font=ctk.CTkFont(size=12, weight="bold"),
            wraplength=155,
            justify="left",
        )
        self.status_badge.pack(side="bottom", anchor="w", padx=14, pady=16)

        # === ЦЕНТРАЛЬНАЯ РАБОЧАЯ ОБЛАСТЬ ===
        self.center = ctk.CTkScrollableFrame(self, fg_color=BG_APP, corner_radius=0)
        self.center.grid(row=1, column=1, sticky="nsew", padx=14, pady=12)

        # Карточка 1: Фирма и шаблон
        card1 = ctk.CTkFrame(
            self.center, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=BORDER_CARD
        )
        card1.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(
            card1, text="1. ФИРМА И ШАБЛОН СПЕЦИФИКАЦИИ",
            font=ctk.CTkFont(size=12, weight="bold"), text_color=ACCENT_AMBER
        ).pack(anchor="w", padx=16, pady=(14, 6))

        self.selected_firm_lbl = ctk.CTkLabel(
            card1, text="Выберите фирму в списке справа →",
            font=ctk.CTkFont(size=16, weight="bold"), text_color=TEXT_MAIN
        )
        self.selected_firm_lbl.pack(anchor="w", padx=16, pady=(0, 4))

        self.last_file_lbl = ctk.CTkLabel(
            card1, text="Последний файл: —",
            font=ctk.CTkFont(size=12), text_color=TEXT_MUTED
        )
        self.last_file_lbl.pack(anchor="w", padx=16, pady=(0, 12))

        ctk.CTkLabel(
            card1, text="Папка фирмы в «РАСЧЕТЫ» (для подстановки нижней строки J–O):",
            font=ctk.CTkFont(size=12), text_color=TEXT_MUTED
        ).pack(anchor="w", padx=16, pady=(0, 4))

        self.calc_var = ctk.StringVar(value="(не выбрана)")
        self.calc_combo = ctk.CTkComboBox(
            card1,
            variable=self.calc_var,
            values=["(не выбрана)"] + self.all_calc_folders,
            height=36,
            fg_color=BG_INPUT,
            border_color=BORDER_CARD,
            button_color="#2a3342",
            dropdown_fg_color=BG_CARD,
        )
        self.calc_combo.pack(fill="x", padx=16, pady=(0, 16))

        # Карточка 2: Номер спецификации и счёт
        card2 = ctk.CTkFrame(
            self.center, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=BORDER_CARD
        )
        card2.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(
            card2, text="2. ПАРАМЕТРЫ СПЕЦИФИКАЦИИ И СЧЁТА 1С",
            font=ctk.CTkFont(size=12, weight="bold"), text_color=ACCENT_BLUE
        ).pack(anchor="w", padx=16, pady=(14, 10))

        grid2 = ctk.CTkFrame(card2, fg_color="transparent")
        grid2.pack(fill="x", padx=16, pady=(0, 14))
        grid2.grid_columnconfigure(0, weight=1)
        grid2.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            grid2, text="Новый № спецификации:", font=ctk.CTkFont(size=12), text_color=TEXT_MUTED
        ).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 4))
        ctk.CTkLabel(
            grid2, text="Номер счёта в 1С (например, 319):", font=ctk.CTkFont(size=12), text_color=TEXT_MUTED
        ).grid(row=0, column=1, sticky="w", padx=(8, 0), pady=(0, 4))

        self.spec_num_var = ctk.StringVar()
        ctk.CTkEntry(
            grid2, textvariable=self.spec_num_var, height=38, fg_color=BG_INPUT,
            border_color=BORDER_CARD, font=ctk.CTkFont(size=14, weight="bold")
        ).grid(row=1, column=0, sticky="ew", padx=(0, 8))

        self.inv_var = ctk.StringVar()
        ctk.CTkEntry(
            grid2, textvariable=self.inv_var, placeholder_text="319", height=38, fg_color=BG_INPUT,
            border_color=BORDER_CARD, font=ctk.CTkFont(size=14, weight="bold")
        ).grid(row=1, column=1, sticky="ew", padx=(8, 0))

        self.use_open_form_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            card2,
            text="Печатная форма счёта уже открыта в 1С (забрать сразу, без поиска)",
            variable=self.use_open_form_var,
            font=ctk.CTkFont(size=12),
            text_color=TEXT_MAIN,
            fg_color=ACCENT_GREEN,
            hover_color=ACCENT_GREEN_HOVER,
        ).pack(anchor="w", padx=16, pady=(0, 16))

        # Карточка 3: Журнал выполнения
        card3 = ctk.CTkFrame(
            self.center, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=BORDER_CARD
        )
        card3.pack(fill="both", expand=True, pady=(0, 12))

        ctk.CTkLabel(
            card3, text="3. ХОД ВЫПОЛНЕНИЯ",
            font=ctk.CTkFont(size=12, weight="bold"), text_color=ACCENT_GREEN
        ).pack(anchor="w", padx=16, pady=(14, 6))

        self.log_box = ctk.CTkTextbox(
            card3,
            height=135,
            fg_color=BG_INPUT,
            border_width=1,
            border_color=BORDER_CARD,
            font=ctk.CTkFont(family="Consolas", size=12),
            text_color="#cbd5e1",
        )
        self.log_box.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        self._append_log("Готово. Выберите фирму справа, введите номер счёта и нажмите зелёную кнопку.")

        self.build_btn = ctk.CTkButton(
            self.center,
            text="СОБРАТЬ СПЕЦИФИКАЦИЮ И ОТКРЫТЬ EXCEL",
            height=48,
            corner_radius=10,
            fg_color=ACCENT_GREEN,
            hover_color=ACCENT_GREEN_HOVER,
            text_color="#052e16",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            command=self._on_start_build,
        )
        self.build_btn.pack(fill="x", pady=(4, 8))

        # === ПРАВАЯ ПАНЕЛЬ: СПРАВОЧНИК ФИРМ ===
        right_panel = ctk.CTkFrame(
            self, width=300, fg_color=BG_SIDEBAR, corner_radius=0,
            border_width=1, border_color=BORDER_CARD
        )
        right_panel.grid(row=1, column=2, sticky="ns")
        right_panel.grid_propagate(False)

        self.firms_header_lbl = ctk.CTkLabel(
            right_panel,
            text=f"Справочник фирм ({len(self.all_clients)})",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=ACCENT_AMBER,
        )
        self.firms_header_lbl.pack(anchor="w", padx=14, pady=(14, 8))

        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._render_client_list(self.search_var.get()))
        ctk.CTkEntry(
            right_panel,
            textvariable=self.search_var,
            placeholder_text="🔍 Поиск фирмы...",
            height=36,
            fg_color=BG_INPUT,
            border_color=BORDER_CARD,
        ).pack(fill="x", padx=14, pady=(0, 10))

        self.clients_scroll = ctk.CTkScrollableFrame(right_panel, fg_color="transparent")
        self.clients_scroll.pack(fill="both", expand=True, padx=8, pady=(0, 10))

    def _append_log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{ts}] {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _set_status(self, msg: str, color: str = ACCENT_GREEN):
        self.after(0, lambda: self.status_badge.configure(text=f"● {msg}", text_color=color))
        self.after(0, lambda: self._append_log(msg))

    def _render_client_list(self, query: str):
        for widget in self.clients_scroll.winfo_children():
            widget.destroy()
        self.client_buttons.clear()

        q = query.strip().lower()
        filtered = [c for c in self.all_clients if q in c.lower()]
        self.firms_header_lbl.configure(text=f"Справочник фирм ({len(filtered)})")

        for name in filtered:
            is_sel = name == self.selected_client
            btn = ctk.CTkButton(
                self.clients_scroll,
                text=name,
                anchor="w",
                height=36,
                corner_radius=8,
                fg_color="#163529" if is_sel else BG_CARD,
                hover_color="#1f2937",
                border_width=1,
                border_color=ACCENT_GREEN if is_sel else BORDER_CARD,
                text_color="#34d399" if is_sel else TEXT_MAIN,
                font=ctk.CTkFont(size=12, weight="bold" if is_sel else "normal"),
                command=lambda n=name: self._select_client(n),
            )
            btn.pack(fill="x", pady=3, padx=4)
            self.client_buttons[name] = btn

    def _select_client(self, client_name: str):
        self.selected_client = client_name
        self.selected_firm_lbl.configure(text=client_name)

        for name, btn in self.client_buttons.items():
            is_sel = name == client_name
            btn.configure(
                fg_color="#163529" if is_sel else BG_CARD,
                border_color=ACCENT_GREEN if is_sel else BORDER_CARD,
                text_color="#34d399" if is_sel else TEXT_MAIN,
                font=ctk.CTkFont(size=12, weight="bold" if is_sel else "normal"),
            )

        spec_folder = os.path.join(CFG["specs_root"], client_name)
        try:
            last_num, last_file = find_latest_spec(spec_folder)
            self.spec_num_var.set(str(last_num + 1))
            self.last_file_lbl.configure(
                text=f"Последний файл: {os.path.basename(last_file)}",
                text_color="#94a3b8",
            )
        except Exception as e:
            self.spec_num_var.set("1")
            self.last_file_lbl.configure(text=f"Внимание: {e}", text_color="#f87171")

        self.calc_var.set(guess_calc_folder(client_name, self.all_calc_folders))

    def _open_selected_spec_folder(self):
        if not self.selected_client:
            os.startfile(CFG["specs_root"])
            return
        folder = os.path.join(CFG["specs_root"], self.selected_client)
        if os.path.isdir(folder):
            os.startfile(folder)

    def _open_selected_calc_folder(self):
        calc = self.calc_var.get().strip()
        if calc and calc != "(не выбрана)":
            folder = os.path.join(CFG["calc_root"], calc)
            if os.path.isdir(folder):
                os.startfile(folder)
                return
        if os.path.isdir(CFG["calc_root"]):
            os.startfile(CFG["calc_root"])

    def _reload_all_folders(self):
        self._load_folders()
        self.calc_combo.configure(values=["(не выбрана)"] + self.all_calc_folders)
        self._render_client_list(self.search_var.get())
        self._append_log("Списки папок обновлены.")

    def _on_start_build(self):
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

        calc_folder_name = self.calc_var.get().strip()
        use_open = self.use_open_form_var.get()
        self.build_btn.configure(state="disabled", text="ВЫПОЛНЯЕТСЯ СБОРКА...")

        def worker():
            try:
                res = build_specification(
                    CFG,
                    self.selected_client,
                    calc_folder_name,
                    inv_num,
                    new_spec_num,
                    use_open,
                    lambda m: self._set_status(m, ACCENT_AMBER),
                )
                self._set_status("Готово! Спецификация открыта в Excel.", ACCENT_GREEN)
                self.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Готово", f"Спецификация собрана и открыта в Excel:\n{res}"
                    ),
                )
            except Exception as e:
                self._set_status(f"Ошибка: {e}", "#f87171")
                self.after(0, lambda err=str(e): messagebox.showerror("Ошибка", err))
            finally:
                self.after(
                    0,
                    lambda: self.build_btn.configure(
                        state="normal", text="СОБРАТЬ СПЕЦИФИКАЦИЮ И ОТКРЫТЬ EXCEL"
                    ),
                )

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = SpecBuilderApp()
    app.mainloop()
