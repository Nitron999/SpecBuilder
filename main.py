import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk

import openpyxl
import win32com.client
import xlrd
from num2words import num2words
from pywinauto import Desktop, keyboard


def get_app_dir() -> str:
    """Возвращает папку, в которой лежит SpecBuilder.exe (или main.py)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(get_app_dir(), "config.json")


def load_config() -> dict:
    if not os.path.exists(CONFIG_PATH):
        raise FileNotFoundError(
            f"Не найден файл настроек:\n{CONFIG_PATH}\n\n"
            "Положи config.json в ту же папку, где лежит SpecBuilder.exe."
        )
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def resolve_1c_exe(configured_path: str) -> str:
    """Проверяет путь к 1cestart.exe и при необходимости подхватывает Program Files (x86)."""
    candidates = [
        configured_path,
        r"C:\Program Files\1cv8\common\1cestart.exe",
        r"C:\Program Files (x86)\1cv8\common\1cestart.exe",
    ]
    for path in candidates:
        if path and os.path.exists(path):
            return path
    raise FileNotFoundError(
        "Не найден 1cestart.exe. Проверь путь в поле '1c_exe' внутри config.json."
    )


def parse_number(val) -> float:
    """Преобразует число или строку вида '10 171,20 руб' в float."""
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
    """Форматирует число в строку вида '10\xa0171,20\xa0руб'."""
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
    """Преобразует сумму в рубли и копейки прописью."""
    rub = int(amount)
    kop = int(round((amount - rub) * 100))
    if kop == 100:
        rub += 1
        kop = 0

    rub_words = num2words(rub, lang="ru").capitalize()
    rub_unit = decline(rub, "рубль", "рубля", "рублей")
    kop_unit = decline(kop, "копейка", "копейки", "копеек")
    return f"{rub_words} {rub_unit} {kop:02d} {kop_unit}"


def find_latest_spec(spec_folder: str):
    """Находит файл сп.<N>сч.<M>.xls с максимальным номером спецификации N."""
    pattern = re.compile(r"сп\.(\d+)[сc]ч\.(\d+)", re.IGNORECASE)
    best_num = -1
    best_file = None

    for fname in os.listdir(spec_folder):
        if not fname.lower().endswith((".xls", ".xlsx")):
            continue
        m = pattern.search(fname)
        if m:
            num = int(m.group(1))
            if num > best_num:
                best_num = num
                best_file = os.path.join(spec_folder, fname)

    if not best_file:
        raise FileNotFoundError(
            f"В папке {spec_folder} не найдено файлов вида сп.520сч.730.xls"
        )
    return best_num, best_file


def find_1c_window():
    desktop = Desktop(backend="uia")
    for w in desktop.windows():
        title = w.window_text()
        if "БП_рабочая база" in title or ("1С:Предприятие" in title and "Запуск" not in title):
            return w
    return None


def export_invoice_from_1c(cfg: dict, invoice_number: str) -> str:
    """Подключается к 1С (или запускает её), находит счёт и сохраняет печатную форму в .xlsx."""
    win_1c = find_1c_window()

    if win_1c is None:
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
            "Не удалось дождаться открытия окна 1С:Предприятие.\n"
            "Проверь логин, пароль и адрес базы в config.json."
        )

    win_1c.set_focus()
    time.sleep(0.7)

    # Переходим в Продажи -> Счета покупателям (если ещё не там)
    try:
        win_1c.child_window(title="Продажи", control_type="TabItem").click_input()
        time.sleep(0.6)
        win_1c.child_window(title="Счета покупателям").click_input()
        time.sleep(1.2)
    except Exception:
        pass

    # Ищем счёт по номеру через Ctrl+F
    keyboard.send_keys("^f")
    time.sleep(0.6)
    keyboard.send_keys(str(invoice_number), with_spaces=True)
    keyboard.send_keys("{ENTER}")
    time.sleep(1.2)

    # Открываем найденный счёт
    keyboard.send_keys("{ENTER}")
    time.sleep(1.8)

    # Открываем печатную форму
    try:
        win_1c.child_window(title_re=".*Печать.*").click_input()
        time.sleep(0.6)
        keyboard.send_keys("{ENTER}")
    except Exception:
        keyboard.send_keys("^p")

    time.sleep(2.2)

    # Сохраняем печатную форму счёта во временный .xlsx через Ctrl+S
    tmp_xlsx = os.path.join(tempfile.gettempdir(), f"invoice_{invoice_number}.xlsx")
    if os.path.exists(tmp_xlsx):
        os.remove(tmp_xlsx)

    keyboard.send_keys("^s")
    time.sleep(1.2)
    keyboard.send_keys(tmp_xlsx, with_spaces=True)
    time.sleep(0.4)
    keyboard.send_keys("{ENTER}")
    time.sleep(2.5)

    if not os.path.exists(tmp_xlsx):
        raise RuntimeError(
            "Не удалось сохранить печатную форму счёта в .xlsx.\n"
            "Убедись, что счёт с таким номером есть в «Счета покупателям»."
        )
    return tmp_xlsx


def parse_invoice_xlsx(xlsx_path: str) -> dict:
    """Парсит сохранённую из 1С печатную форму счёта."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active

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
    for row in ws.iter_rows(values_only=True):
        vals = [v for v in row if v is not None and str(v).strip() != ""]
        if not vals:
            continue
        joined = " ".join(str(v) for v in vals)

        m = date_re.search(joined)
        if m:
            day = m.group(1).zfill(2)
            mon = months.get(m.group(2).lower(), "01")
            year = m.group(3)
            invoice_date = f"{day}.{mon}.{year}"

        if "Товары" in joined or ("Наименование" in joined and "Цена" in joined):
            in_table = True
            continue

        if in_table:
            if "Итого" in joined:
                in_table = False
                totals["no_nds"] = parse_number(vals[-1])
                continue
            if isinstance(vals[0], (int, float)) or str(vals[0]).strip().isdigit():
                if len(vals) >= 6:
                    items.append({
                        "name": str(vals[1]).strip(),
                        "unit": str(vals[2]).strip(),
                        "qty": parse_number(vals[3]),
                        "price": parse_number(vals[4]),
                        "sum": parse_number(vals[5]),
                    })

        if "НДС" in joined and totals["nds"] == 0.0 and "без НДС" not in joined.lower():
            totals["nds"] = parse_number(vals[-1])
        elif "Всего к оплате" in joined:
            totals["with_nds"] = parse_number(vals[-1])

    # Если итоги не считались из текста — считаем сами по позициям (ставка НДС 22%)
    if totals["no_nds"] == 0.0 and items:
        totals["no_nds"] = round(sum(i["sum"] for i in items), 2)
    if totals["nds"] == 0.0 and totals["no_nds"] > 0:
        totals["nds"] = round(totals["no_nds"] * 0.22, 2)
    if totals["with_nds"] == 0.0 and totals["no_nds"] > 0:
        totals["with_nds"] = round(totals["no_nds"] + totals["nds"], 2)

    return {"date": invoice_date, "items": items, "totals": totals}


def find_calc_bottom_row(calc_folder: str, board_name: str) -> list:
    """Ищет файл расчёта по коду платы и забирает нижнюю заполненную строку (6 ячеек J-O)."""
    if not os.path.isdir(calc_folder):
        return [""] * 6

    code = board_name.split()[0].strip().lower()
    target_file = None

    for fname in os.listdir(calc_folder):
        if fname.lower().endswith((".xls", ".xlsx")) and code in fname.lower():
            target_file = os.path.join(calc_folder, fname)
            break

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


def build_specification(cfg: dict, client_name: str, invoice_num: str) -> str:
    spec_folder = os.path.join(cfg["specs_root"], client_name)
    calc_folder_name = cfg.get("folder_mapping", {}).get(client_name, client_name)
    calc_folder = os.path.join(cfg["calc_root"], calc_folder_name)

    last_num, last_file = find_latest_spec(spec_folder)
    new_num = last_num + 1
    ext = os.path.splitext(last_file)[1]
    new_file = os.path.join(spec_folder, f"сп.{new_num}сч.{invoice_num}{ext}")
    shutil.copy2(last_file, new_file)

    tmp_xlsx = export_invoice_from_1c(cfg, invoice_num)
    inv = parse_invoice_xlsx(tmp_xlsx)
    items = inv["items"]

    if not items:
        raise RuntimeError("Не удалось распознать позиции товаров в печатной форме счёта.")

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    wb = excel.Workbooks.Open(new_file)
    ws = wb.Sheets(1)

    # Обновляем шапку
    ws.Range("O1").Value = f"Приложение № {new_num}"
    ws.Range("A4").Value = f"Спецификация № {new_num} от {inv['date']} г."

    # Находим строку "Итого без НДС:"
    start_row = 9
    total_row = start_row
    while total_row < 150:
        val = str(ws.Cells(total_row, 1).Value or "")
        if "Итого без НДС" in val:
            break
        total_row += 1

    existing_rows = total_row - start_row
    needed_rows = len(items)

    # Подгоняем количество строк под число позиций в счёте
    if needed_rows > existing_rows:
        for _ in range(needed_rows - existing_rows):
            ws.Rows(start_row + 1).Insert()
            ws.Rows(start_row).Copy(ws.Rows(start_row + 1))
    elif needed_rows < existing_rows:
        for _ in range(existing_rows - needed_rows):
            ws.Rows(start_row + needed_rows).Delete()

    # Заполняем позиции (столбцы A-O)
    for idx, item in enumerate(items, start=1):
        r = start_row + idx - 1
        ws.Cells(r, 1).Value = idx                                   # A: № п/п
        ws.Cells(r, 2).Value = item["name"]                          # B: Наименование
        ws.Cells(r, 3).Value = ""                                    # C: Зав. № заказа
        ws.Cells(r, 4).Value = ""                                    # D: Срок пост.
        ws.Cells(r, 5).Value = item["unit"]                          # E: Ед. изм.
        ws.Cells(r, 6).Value = item["qty"]                           # F: Кол-во
        ws.Cells(r, 7).Value = item["price"]                         # G: Цена
        ws.Cells(r, 8).Value = format_rub_text(item["sum"], False)   # H: Сумма
        ws.Cells(r, 9).Value = ""                                    # I: Площадь

        tech = find_calc_bottom_row(calc_folder, item["name"])
        for offset, val in enumerate(tech):
            ws.Cells(r, 10 + offset).Value = val

    # Обновляем итоги и суммы прописью
    new_total_row = start_row + needed_rows
    ws.Cells(new_total_row, 8).Value = format_rub_text(inv["totals"]["no_nds"], True)
    ws.Cells(new_total_row + 2, 8).Value = format_rub_text(inv["totals"]["nds"], True)
    ws.Cells(new_total_row + 3, 8).Value = format_rub_text(inv["totals"]["with_nds"], True)
    ws.Cells(new_total_row + 5, 3).Value = amount_in_words(inv["totals"]["with_nds"])
    ws.Cells(new_total_row + 6, 3).Value = amount_in_words(inv["totals"]["nds"])

    wb.Save()
    return new_file


def run_gui():
    try:
        cfg = load_config()
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Ошибка конфигурации", str(e))
        return

    try:
        clients = sorted(
            d for d in os.listdir(cfg["specs_root"])
            if os.path.isdir(os.path.join(cfg["specs_root"], d))
        )
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Нет доступа к сетевой папке", f"{cfg['specs_root']}\n\n{e}")
        return

    root = tk.Tk()
    root.title("Сборка спецификации (1С + Расчёты)")
    root.geometry("440x200")
    root.resizable(False, False)

    tk.Label(root, text="Контрагент (папка в Спецификациях):").pack(anchor="w", padx=15, pady=(14, 2))
    client_var = tk.StringVar(value=clients[0] if clients else "")
    ttk.Combobox(root, textvariable=client_var, values=clients, width=55).pack(padx=15)

    tk.Label(root, text="Номер счёта в 1С:").pack(anchor="w", padx=15, pady=(12, 2))
    inv_entry = tk.Entry(root, width=58)
    inv_entry.pack(padx=15)

    def on_start():
        client = client_var.get().strip()
        inv_num = inv_entry.get().strip()
        if not client or not inv_num:
            messagebox.showwarning("Внимание", "Выбери контрагента и введи номер счёта.")
            return
        try:
            res = build_specification(cfg, client, inv_num)
            messagebox.showinfo("Готово", f"Спецификация собрана и открыта в Excel:\n{res}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    tk.Button(root, text="Собрать спецификацию", command=on_start, height=2, width=28).pack(pady=16)
    root.mainloop()


if __name__ == "__main__":
    run_gui()
