import os
import re
import shutil
import subprocess
import tempfile
import threading
import time
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk

import openpyxl
import pythoncom
import win32com.client
import xlrd
from num2words import num2words
from pywinauto import Desktop, keyboard


# Все настройки, логин, пароль и пути зашиты прямо здесь
CFG = {
    "1c_exe": r"C:\Program Files\1cv8\common\1cestart.exe",
    "1c_base_args": 'ENTERPRISE /S"DB01\\pp_buh_new" /N"МамонтоваЛД" /P"mamon123"',
    "specs_root": r"\\Fs01\obmen\ВП\Спецификации 2019 Новая форма",
    "calc_root": r"\\Fs01\obmen\РАСЧЕТЫ",
    "folder_mapping": {
        "Электроприбор": "Elektropribor_K",
    },
}


def resolve_1c_exe(configured_path: str) -> str:
    """Находит 1cestart.exe в Program Files или Program Files (x86)."""
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
    """Превращает 'счет 40', 'сч. 730', '№ 0083' в чистый номер ('40', '730', '0083')."""
    s = raw.strip()
    s = re.sub(r"^(?:сч[её]т|сч\.?|№)\s*", "", s, flags=re.IGNORECASE).strip()
    return s


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
    """
    Находит последнюю спецификацию в папке.
    Понимает форматы: 'сп.520сч.730.xls', 'сп. 1 счет 40.xls', 'сп.2счет 1083.xls',
    'Спецификация 5.xls' и др. Если по номеру не нашлось — берёт самый свежий .xls/.xlsx.
    """
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
        # Берём максимальный номер; если номера одинаковые — самый свежий по дате
        matched.sort(key=lambda x: (x[0], x[1]), reverse=True)
        best_num, _, best_file = matched[0]
        return best_num, best_file

    if all_excel:
        all_excel.sort(key=lambda x: x[0], reverse=True)
        return 0, all_excel[0][1]

    raise FileNotFoundError(f"В папке {spec_folder} вообще нет файлов Excel (.xls / .xlsx).")


def guess_calc_folder(client_name: str, calc_folders: list) -> str:
    """Подбирает папку в РАСЧЕТЫ по маппингу или похожему названию."""
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


def find_1c_window():
    desktop = Desktop(backend="uia")
    for w in desktop.windows():
        title = w.window_text()
        if "БП_рабочая база" in title or ("1С:Предприятие" in title and "Запуск" not in title):
            return w
    return None


def export_invoice_from_1c(cfg: dict, invoice_number: str, log_cb) -> str:
    """Подключается к 1С (или запускает её), находит счёт и сохраняет печатную форму в .xlsx."""
    log_cb("Ищем открытое окно 1С...")
    win_1c = find_1c_window()

    if win_1c is None:
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
        raise RuntimeError("Не удалось дождаться открытия окна 1С:Предприятие.")

    win_1c.set_focus()
    time.sleep(0.7)

    log_cb("Открываем «Продажи → Счета покупателям»...")
    try:
        win_1c.child_window(title="Продажи", control_type="TabItem").click_input()
        time.sleep(0.6)
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

    log_cb("Открываем печатную форму и выгружаем в Excel...")
    try:
        win_1c.child_window(title_re=".*Печать.*").click_input()
        time.sleep(0.6)
        keyboard.send_keys("{ENTER}")
    except Exception:
        keyboard.send_keys("^p")

    time.sleep(2.2)

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
            f"Не удалось сохранить печатную форму счёта № {invoice_number} в .xlsx.\n"
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

    if totals["no_nds"] == 0.0 and items:
        totals["no_nds"] = round(sum(i["sum"] for i in items), 2)
    if totals["nds"] == 0.0 and totals["no_nds"] > 0:
        totals["nds"] = round(totals["no_nds"] * 0.22, 2)
    if totals["with_nds"] == 0.0 and totals["no_nds"] > 0:
        totals["with_nds"] = round(totals["no_nds"] + totals["nds"], 2)

    return {"date": invoice_date, "items": items, "totals": totals}


def find_calc_bottom_row(calc_folder: str, board_name: str) -> list:
    """Ищет файл расчёта по коду платы и забирает нижнюю заполненную строку (6 ячеек J-O)."""
    if not calc_folder or not os.path.isdir(calc_folder):
        return [""] * 6

    code = board_name.split()[0].strip().lower()
    target_file = None

    for fname in os.listdir(calc_folder):
        if fname.startswith("~$"):
            continue
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


def build_specification(
    cfg: dict,
    client_name: str,
    calc_folder_name: str,
    invoice_num: str,
    new_spec_num: int,
    log_cb,
) -> str:
    spec_folder = os.path.join(cfg["specs_root"], client_name)
    calc_folder = (
        os.path.join(cfg["calc_root"], calc_folder_name)
        if calc_folder_name and calc_folder_name != "(не выбрана)"
        else ""
    )

    log_cb("Копируем шаблон последней спецификации...")
    _, last_file = find_latest_spec(spec_folder)
    ext = os.path.splitext(last_file)[1]
    new_file = os.path.join(spec_folder, f"сп.{new_spec_num} сч.{invoice_num}{ext}")
    shutil.copy2(last_file, new_file)

    tmp_xlsx = export_invoice_from_1c(cfg, invoice_num, log_cb)

    log_cb("Разбираем позиции из счёта 1С...")
    inv = parse_invoice_xlsx(tmp_xlsx)
    items = inv["items"]
    if not items:
        raise RuntimeError("Не удалось распознать позиции товаров в печатной форме счёта.")

    log_cb("Заполняем спецификацию в Excel...")
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
            ws.Cells(r, 1).Value = idx
            ws.Cells(r, 2).Value = item["name"]
            ws.Cells(r, 3).Value = ""
            ws.Cells(r, 4).Value = ""
            ws.Cells(r, 5).Value = item["unit"]
            ws.Cells(r, 6).Value = item["qty"]
            ws.Cells(r, 7).Value = item["price"]
            ws.Cells(r, 8).Value = format_rub_text(item["sum"], False)
            ws.Cells(r, 9).Value = ""

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


def run_gui():
    try:
        all_clients = sorted(
            d for d in os.listdir(CFG["specs_root"])
            if os.path.isdir(os.path.join(CFG["specs_root"], d))
        )
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Нет доступа к сетевой папке", f"{CFG['specs_root']}\n\n{e}")
        return

    try:
        all_calc_folders = sorted(
            d for d in os.listdir(CFG["calc_root"])
            if os.path.isdir(os.path.join(CFG["calc_root"], d))
        )
    except Exception:
        all_calc_folders = []

    root = tk.Tk()
    root.title("Сборка спецификации (1С + Расчёты)")
    root.geometry("560x540")
    root.minsize(520, 500)

    pad = {"padx": 14}

    # --- 1. Поиск и выбор фирмы ---
    tk.Label(root, text="Поиск фирмы (в Спецификациях):", font=("Segoe UI", 9, "bold")).pack(
        anchor="w", pady=(12, 2), **pad
    )
    search_var = tk.StringVar()
    search_entry = tk.Entry(root, textvariable=search_var, font=("Segoe UI", 10))
    search_entry.pack(fill="x", **pad)

    list_frame = tk.Frame(root)
    list_frame.pack(fill="both", expand=True, pady=(6, 6), **pad)

    scrollbar = tk.Scrollbar(list_frame)
    scrollbar.pack(side="right", fill="y")

    listbox = tk.Listbox(
        list_frame,
        height=8,
        font=("Segoe UI", 10),
        yscrollcommand=scrollbar.set,
        exportselection=False,
    )
    listbox.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=listbox.yview)

    # Информация о найденном шаблоне
    info_var = tk.StringVar(value="Выбери фирму из списка выше")
    tk.Label(root, textvariable=info_var, fg="#333333", anchor="w", justify="left").pack(
        fill="x", pady=(0, 8), **pad
    )

    # --- 2. Папка в РАСЧЕТЫ ---
    tk.Label(root, text="Папка фирмы в РАСЧЕТЫ:", font=("Segoe UI", 9, "bold")).pack(
        anchor="w", pady=(2, 2), **pad
    )
    calc_var = tk.StringVar(value="(не выбрана)")
    calc_combo = ttk.Combobox(
        root,
        textvariable=calc_var,
        values=["(не выбрана)"] + all_calc_folders,
        font=("Segoe UI", 10),
    )
    calc_combo.pack(fill="x", **pad)

    # --- 3. Номер спецификации и номер счёта ---
    row_frame = tk.Frame(root)
    row_frame.pack(fill="x", pady=(10, 6), **pad)

    left_col = tk.Frame(row_frame)
    left_col.pack(side="left", fill="x", expand=True, padx=(0, 8))
    tk.Label(left_col, text="Новый № спецификации:", font=("Segoe UI", 9, "bold")).pack(anchor="w")
    spec_num_var = tk.StringVar()
    tk.Entry(left_col, textvariable=spec_num_var, font=("Segoe UI", 10)).pack(fill="x", pady=(2, 0))

    right_col = tk.Frame(row_frame)
    right_col.pack(side="left", fill="x", expand=True, padx=(8, 0))
    tk.Label(right_col, text="Номер счёта в 1С:", font=("Segoe UI", 9, "bold")).pack(anchor="w")
    inv_var = tk.StringVar()
    inv_entry = tk.Entry(right_col, textvariable=inv_var, font=("Segoe UI", 10))
    inv_entry.pack(fill="x", pady=(2, 0))

    # --- 4. Кнопки и статус ---
    btn_frame = tk.Frame(root)
    btn_frame.pack(fill="x", pady=(10, 6), **pad)

    start_btn = tk.Button(
        btn_frame,
        text="Собрать спецификацию",
        font=("Segoe UI", 10, "bold"),
        bg="#1e7e34",
        fg="white",
        height=2,
    )
    start_btn.pack(side="left", fill="x", expand=True, padx=(0, 8))

    def open_selected_folder():
        sel = listbox.curselection()
        if not sel:
            return
        folder = os.path.join(CFG["specs_root"], listbox.get(sel[0]))
        if os.path.isdir(folder):
            os.startfile(folder)

    tk.Button(btn_frame, text="Открыть папку", command=open_selected_folder, height=2, width=16).pack(
        side="right"
    )

    status_var = tk.StringVar(value="Готов к работе.")
    tk.Label(root, textvariable=status_var, fg="#0055aa", anchor="w").pack(fill="x", pady=(2, 10), **pad)

    # --- Логика фильтрации списка и автоподстановки ---
    def refresh_list(*_):
        query = search_var.get().strip().lower()
        listbox.delete(0, tk.END)
        for name in all_clients:
            if query in name.lower():
                listbox.insert(tk.END, name)
        if listbox.size() == 1:
            listbox.selection_set(0)
            on_select_client()

    def on_select_client(_event=None):
        sel = listbox.curselection()
        if not sel:
            return
        client = listbox.get(sel[0])
        spec_folder = os.path.join(CFG["specs_root"], client)

        try:
            last_num, last_file = find_latest_spec(spec_folder)
            spec_num_var.set(str(last_num + 1))
            info_var.set(f"Последний файл: {os.path.basename(last_file)}")
        except Exception as e:
            spec_num_var.set("1")
            info_var.set(f"Внимание: {e}")

        calc_var.set(guess_calc_folder(client, all_calc_folders))

    search_var.trace_add("write", refresh_list)
    listbox.bind("<<ListboxSelect>>", on_select_client)
    refresh_list()

    def on_start():
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("Внимание", "Выбери фирму из списка.")
            return
        client = listbox.get(sel[0])
        inv_num = clean_invoice_number(inv_var.get())
        if not inv_num:
            messagebox.showwarning("Внимание", "Введи номер счёта из 1С (например, 40 или 730).")
            return
        try:
            new_spec_num = int(spec_num_var.get().strip())
        except ValueError:
            messagebox.showwarning("Внимание", "Номер спецификации должен быть числом.")
            return

        calc_folder_name = calc_var.get().strip()
        start_btn.config(state="disabled")

        def set_status(msg: str):
            root.after(0, lambda: status_var.set(msg))

        def worker():
            try:
                res = build_specification(
                    CFG, client, calc_folder_name, inv_num, new_spec_num, set_status
                )
                set_status("Готово!")
                root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Готово", f"Спецификация собрана и открыта в Excel:\n{res}"
                    ),
                )
            except Exception as e:
                set_status("Ошибка.")
                root.after(0, lambda err=str(e): messagebox.showerror("Ошибка", err))
            finally:
                root.after(0, lambda: start_btn.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    start_btn.config(command=on_start)
    search_entry.focus_set()
    root.mainloop()


if __name__ == "__main__":
    run_gui()
