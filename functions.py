from datetime import datetime, timedelta
from typing import Optional
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from constants import *
from tkinter import Frame, Scrollbar, Canvas, messagebox, Button, END
import os, sqlite3
from openpyxl import Workbook, load_workbook

#Center Window According to Screen
def center_window(window, w, h):
    window.resizable(False, False)
    window.update_idletasks()
    screen_w = window.winfo_screenwidth()
    screen_h = window.winfo_screenheight()
    x = (screen_w // 2) - (w // 2)
    y = (screen_h // 2) - (h // 2)
    window.geometry(f"{w}x{h}+{x}+{y}")
    window.grab_set()

#Create Scrollable Frame
def create_scrollable_frame(parent):
    contents = Frame(parent, bg=C4)
    contents.pack(fill="both", expand=True, pady=(5, 0))
    canvas = Canvas(contents, bg=C4, highlightthickness=0)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar = Scrollbar(contents, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")
    canvas.configure(yscrollcommand=scrollbar.set)
    scroll_frame = Frame(canvas, bg=C4)
    scroll_window = canvas.create_window((0, 0), window=scroll_frame, anchor="n")
    def update_scroll_region(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(scroll_window, width=canvas.winfo_width())
    scroll_frame.bind("<Configure>", update_scroll_region)
    def _on_mousewheel(event):
        canvas.yview_scroll(-1 * int(event.delta / 120), "units")
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    return contents, canvas, scroll_frame, scrollbar

#Update Time for Entry
def update_time(time_entry):
    now = datetime.now().strftime("%H:%M")
    time_entry.delete(0, END)
    time_entry.insert(0, now)

#Create Excel for Vials
def create_excel_for_vial(excel_path):
    if os.path.exists(excel_path):
        return
    wb = Workbook()
    ws_info = wb.active
    ws_info.title = "Vial Info"
    ws_info.append(["Calibration Date", "Calibration Time", "Activity(mCi)", "Volume(ml)", "Concentration(mCi/ml)", "Expiration Date", "Stored Date", "Disposal Date"])
    ws_admin = wb.create_sheet("Administrations")
    ws_admin.append(["ID", "Date", "Time", "Patient Name", "Concentration(mCi/ml)", "Dose(mCi)", "Volume(ml)", "Volume Left(ml)"])
    wb.save(excel_path)

#Create Excel for Tc99m Gen
def create_excel_for_tc99m(excel_path):
    if not os.path.exists(excel_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Gen Info"
        ws.append(["Gen ID", "Calibration Date", "Calibration Time", "Mo99 Activity (mCi)", "Start Date", "Expiration Date", "Stored Date", "Disposal Date"])
        ws2 = wb.create_sheet("Elutions")
        ws2.append(["", "Date", "Time", "Activity(mCi)", "Expected(mCi)", "Div(%)", "Vol(ml)", "Conc(mCi/ml)"])
        ws3 = wb.create_sheet("Kits")
        ws3.append(["Kit ID", "Patient ID", "Date", "Time", "Kit", "Volume", "Activity(mCi)", "Conc(mCi/ml)", "Dose(mCi)", "Dose Volume(ml)", "Volume Left(ml)", "Patient Name"])
        wb.save(excel_path)

#Create Excel for Ga68 Gen
def create_excel_for_ga68(excel_path):
    if not os.path.exists(excel_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Gen Info"
        ws.append(["Gen ID", "Model", "Start Date", "Calibration Date", "Calibration Time", "Activity (MBq)", "Expiration Date", "Disposal Date"])
        ws2 = wb.create_sheet("Elutions")
        ws2.append(["", "Date", "Time", "Activity(mCi)"])
        ws3 = wb.create_sheet("DOTATOC")
        ws3.append(["", "Date", "Patient", "Weight (kg)", "Admin Time", "Dose (mCi)", "Volume (ml)", "Concentration (mCi/ml)", "ITLC(<2%)"])
        wb.save(excel_path)

#Excel Helpers for Ga68
def get_dotatoc_excel_path(dbfile):
    folder = os.path.dirname(dbfile)
    return os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
def update_dotatoc_excel(dbfile, row_id, new_dose, new_volume, new_concentration, new_itlc):
    excel_path = get_dotatoc_excel_path(dbfile)
    wb = load_workbook(excel_path)
    ws = wb["DOTATOC"]
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=1).value).strip() == str(row_id).strip():
            ws.cell(row=r, column=6, value=new_dose)
            ws.cell(row=r, column=7, value=new_volume)
            ws.cell(row=r, column=8, value=new_concentration)
            ws.cell(row=r, column=9, value=new_itlc)
            break
    wb.save(excel_path)

#Add New Row to Excel Sheets
def append_row_to_sheet(excel_path, sheet_name, row_values):
   if not os.path.exists(excel_path):
       if "tc" in excel_path.lower():
           create_excel_for_tc99m(excel_path)
       elif "ga" in excel_path.lower():
           create_excel_for_ga68(excel_path)
       elif "vial" in excel_path.lower():
           create_excel_for_vial(excel_path)
       else:
           raise ValueError(f"Cannot determine generator type from path: {excel_path}.")
   wb = load_workbook(excel_path)
   if sheet_name not in wb.sheetnames:
       ws = wb.create_sheet(sheet_name)
   else:
       ws = wb[sheet_name]
   ws.append(row_values)
   wb.save(excel_path)

#Find the Patient in Excel and Insert New Data
def find_patient_insert_row(ws, parent_id):
    last_child_row = None
    parent_row = None
    for r in range(2, ws.max_row + 1):
        cell_val = ws.cell(row=r, column=1).value
        if cell_val is None:
            continue
        cell_val = str(cell_val)
        if cell_val == str(parent_id):
            parent_row = r
        elif cell_val.startswith(f"{parent_id}."):
            last_child_row = r
    if last_child_row:
        return last_child_row + 1
    elif parent_row:
        return parent_row + 1
    else:
        return ws.max_row + 1

#Renumber Child Rows After Delete
def renumber_children(conn, ws, tree, parent_iid):
    cur = conn.cursor()
    cur.execute("""SELECT id FROM kits WHERE parent_id=? ORDER BY time""", (parent_iid,))
    children = [row[0] for row in cur.fetchall()]
    max_idx = 0
    for cid in children:
        parts = cid.split(".")
        if len(parts) == 2 and parts[0] == str(parent_iid):
            try:
                seq = int(parts[1])
                if seq > max_idx:
                    max_idx = seq
            except ValueError:
                continue
    next_idx = max_idx + 1
    for old_id in children:
        pass
    for r in range(2, ws.max_row + 1):
        cell_val = ws.cell(row=r, column=1).value
        if cell_val and str(cell_val).startswith(f"{parent_iid}."):
            ws.cell(row=r, column=1).value = cell_val
            ws.parent_id = parent_iid
    children_iids = list(tree.get_children(parent_iid))
    for old_iid in children_iids:
        new_iid = old_iid
        values = tree.item(old_iid, "values")
        index = tree.index(old_iid)
        tree.delete(old_iid)
        tree.insert(parent_iid, index, iid=new_iid, values=values)
    conn.commit()
    return next_idx

#Find Last Folder When Opening Old File
def find_last_folder(base_dir, subfolder=None):
    root = os.path.join(base_dir, subfolder) if subfolder else base_dir
    now = datetime.now()
    preferred = os.path.join(root, now.strftime("%Y"), now.strftime("%m"))
    if os.path.exists(preferred):
        return preferred
    latest_path = None
    latest_time = 0
    if os.path.exists(root):
        for dirpath, dirnames, _ in os.walk(root):
            for d in dirnames:
                full_path = os.path.join(dirpath, d)
                try:
                    ctime = os.path.getctime(full_path)
                except OSError:
                    continue
                if ctime > latest_time:
                    latest_time = ctime
                    latest_path = full_path
    return latest_path or root

#Store Gens
def store_gen(*, conn, dbfile, excel_sheet="Gen Info", date_format=DATE_FORMAT, on_store_callback=None):
    if not messagebox.askyesno("Store Generator", "Are you sure you want to store this generator?\nThis action cannot be undone."):
        return False
    stored_date = datetime.now().strftime(date_format)
    cur = conn.cursor()
    cur.execute("UPDATE generator_info SET stored_date=?", (stored_date,))
    conn.commit()
    folder = os.path.dirname(dbfile)
    excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
    wb = load_workbook(excel_path)
    ws = wb[excel_sheet]
    ws.cell(row=2, column=7).value = stored_date
    wb.save(excel_path)
    messagebox.showinfo("Stored", f"Generator stored on {stored_date}.")
    if on_store_callback:
        on_store_callback()
    return True

#Dispose Gens
def dispose_gen(*, conn, dbfile, excel_sheet="Gen Info", date_format="%d-%m-%Y", on_disposed_callback=None):
    if not messagebox.askyesno("Dispose Generator", "Are you sure you want to dispose this generator?\nThis action cannot be undone."):
        return False
    disposal_date = datetime.now().strftime(date_format)
    cur = conn.cursor()
    cur.execute("UPDATE generator_info SET disposal_date=?", (disposal_date,))
    conn.commit()
    folder = os.path.dirname(dbfile)
    excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
    wb = load_workbook(excel_path)
    ws = wb[excel_sheet]
    ws.cell(row=2, column=6).value = disposal_date
    wb.save(excel_path)
    messagebox.showinfo("Disposed", f"Generator disposed on {disposal_date}.")
    if on_disposed_callback:
        on_disposed_callback()
    return True

#Disable Buttons after Stored or Expired
def disable_buttons(parent, exempt_texts=None):
    if exempt_texts is None:
        exempt_texts = []
    for widget in parent.winfo_children():
        if isinstance(widget, Button):
            if widget.cget("text") not in exempt_texts:
                widget.config(state="disabled")
        elif widget.winfo_children():
            disable_buttons(widget, exempt_texts)

#Update Headers and Disable Buttons after Stored or Expired
def update_header_and_disable(cur, header, tab, is_stored=False, is_disposed=False, is_expired=False):
    cur.execute("SELECT stored_date, disposal_date FROM generator_info ORDER BY rowid DESC LIMIT 1")
    stored_date, disposal_date = cur.fetchone()
    if is_disposed:
        header.config(text=f"⚠ GENERATOR DISPOSED ({disposal_date}) – NO FURTHER ACTIONS ALLOWED", fg="#660000", highlightthickness=1)
        disable_buttons(tab, exempt_texts=["Back", "Load"])
    elif is_stored:
        header.config(text=f"⚠ GENERATOR STORED ({stored_date}) – NO FURTHER ACTIONS ALLOWED", fg="#CC0000", highlightthickness=1)
        disable_buttons(tab, exempt_texts=["Back", "Load", "✗Dispose Gen✗"])
    elif is_expired:
        header.config(text="⚠ GENERATOR EXPIRED – NO FURTHER ACTIONS ALLOWED", fg="#FF8000", highlightthickness=1)
        disable_buttons(tab, exempt_texts=["Back", "Load", "✗Store Gen✗"])

def ensure_dir(path):
    os.makedirs(path, exist_ok=True)

#=====DECAY AND DATES CALCULATIONS=====
def bq_to_mci(bq):
    return float(bq) / 3.7e7

def mci_to_bq(mci):
    return float(mci) * 3.7e7

def decay_activity(activity_mci, half_life_hours, delta_hours):
    lambda_ = math.log(2) / half_life_hours
    return float(activity_mci) * math.exp(-lambda_ * float(delta_hours))

def activity_now(radionuclide, stored_at_str, activity0):
    now_dt = datetime.now()
    stored_dt = datetime.strptime(stored_at_str, DATE_FORMAT)
    delta_h = (now_dt - stored_dt).total_seconds() / 3600.0
    if delta_h < 0:
        delta_h = 0.0
    half_life = next(hl for name, hl in VIAL_DATA if name == radionuclide)
    return float(decay_activity(float(activity0), float(half_life), float(delta_h)))

def calc_date_below_limit(activity_mci, half_life_hours, limit_bq, start_date):
    activity_mci = float(activity_mci)
    half_life_hours = float(half_life_hours)
    if half_life_hours <= 0:
        return start_date.strftime(DATE_FORMAT)
    if limit_bq is None:
        return start_date.strftime(DATE_FORMAT)
    limit_mci = round(float(limit_bq) / 3.7e7, 2)
    if limit_mci <= 0:
        return start_date.strftime(DATE_FORMAT)
    if float(activity_mci) <= limit_mci:
        return start_date.strftime(DATE_FORMAT)
    lambda_ = math.log(2) / half_life_hours
    ratio = activity_mci / limit_mci
    if ratio <= 1:
        return start_date.strftime(DATE_FORMAT)
    t_hours = math.log(ratio) / lambda_
    return (start_date + timedelta(hours=t_hours)).strftime(DATE_FORMAT)

def calc_recommended_and_permitted_date(radionuclide, activity_mci, stored_at, safety_factor=0.1):
    start_date = datetime.strptime(stored_at, DATE_FORMAT)
    half_life = next(hl for name, hl in VIAL_DATA if name == radionuclide)
    limit_bq = DISPOSAL_LIMITS_BQ[radionuclide]
    if limit_bq in (None, 0):
        recommended = (start_date + timedelta(days=60)).strftime(DATE_FORMAT)
        permitted = (start_date + timedelta(days=60)).strftime(DATE_FORMAT)
        return recommended, permitted, None
    permitted = calc_date_below_limit(activity_mci, half_life, limit_bq, start_date)
    recommended = calc_date_below_limit(activity_mci, half_life, limit_bq * safety_factor, start_date)
    return recommended, permitted, float(limit_bq)

#=====TREE STATUS=====
def disposal_status(recommended_date, permitted_date):
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    rec = datetime.strptime(recommended_date, DATE_FORMAT)
    perm = datetime.strptime(permitted_date, DATE_FORMAT)
    if today < perm:
        return "STORED"
    elif perm < today < rec:
        return "WAIT"
    return "READY"

#=====DISPOSAL SUMMARY=====
def disposal_summary(rows):
    total_vials = len(rows)
    now_dt = datetime.now()
    total_activity_now = 0.0
    ready_count = 0
    for r in rows:
        radionuclide = r[1]
        stored_at_str = r[3]
        activity0 = float(r[4])
        half_life = next(hl for name, hl in VIAL_DATA if name == radionuclide)
        stored_dt = datetime.strptime(stored_at_str, DATE_FORMAT)
        delta_hours = (now_dt - stored_dt).total_seconds() / 3600
        if delta_hours < 0:
            delta_hours = 0.0
        activity_now = decay_activity(activity0, half_life, delta_hours)
        total_activity_now += activity_now
        permitted = r[5]
        recommended = r[6]
        if disposal_status(recommended, permitted) == "READY":
            ready_count += 1
    return total_vials, round(total_activity_now, 2), ready_count

#=====DAILY LOG EXCEL+SQLite (COMMON)=====
def get_disposed_by_date_dir(for_date: Optional[datetime] = None) -> str:
    if for_date is None:
        for_date = datetime.now()
    year = for_date.strftime("%Y")
    month = for_date.strftime("%m")
    path = os.path.join(DAILY_DISPOSALS_DIR, year, month)
    ensure_dir(path)
    return path

def get_daily_disposal_excel_path(disposal_date_str: str) -> str:
    dt = datetime.strptime(disposal_date_str, DATE_FORMAT)
    folder = get_disposed_by_date_dir(dt)
    return os.path.join(folder, f"disposal__{disposal_date_str}.xlsx")

def get_ready_vials_pdf_path(pdf_date_str: str) -> str:
    dt = datetime.strptime(pdf_date_str, DATE_FORMAT)
    folder = get_disposed_by_date_dir(dt)
    return os.path.join(folder, f"ready_vials__{pdf_date_str}.pdf")

def ensure_daily_log_workbook(xlsx_path: str):
    if os.path.exists(xlsx_path):
        wb = load_workbook(xlsx_path)
    else:
        wb = Workbook()
        ws0 = wb.active
        wb.remove(ws0)
    if "Vials" not in wb.sheetnames:
        ws = wb.create_sheet("Vials")
        ws.append(["Disposal Date", "Disposal Time", "Radionuclide", "Calibration Date", "Stored At", "Activity(mCi)",
                   "Permitted Date", "Recommended Date", "Limit(mCi)", "Limit(Bq)"])
    if "Tc99m" not in wb.sheetnames:
        wb.create_sheet("Tc99m")
    wb.save(xlsx_path)
    return wb

def ensure_daily_log_sqlite():
    ensure_dir(DAILY_DISPOSALS_DIR)
    conn  = sqlite3.connect(DAILY_LOG_DB)
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS disposed_vials (id INTEGER PRIMARY KEY AUTOINCREMENT,
                                                              disposal_date TEXT NOT NULL,
                                                              disposal_time TEXT NOT NULL,
                                                              radionuclide TEXT NOT NULL,
                                                              calibration_date TEXT,
                                                              stored_at TEXT NOT NULL,
                                                              activity_mci REAL NOT NULL,
                                                              permitted_date TEXT,
                                                              recommended_date TEXT,
                                                              limit_mci REAL,
                                                              limit_bq REAL)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS disposed_tc99m_batches (id INTEGER PRIMARY KEY AUTOINCREMENT,
                                                                      batch_id INTEGER NOT NULL,
                                                                      item_id TEXT NOT NULL,
                                                                      stored_at TEXT NOT NULL,
                                                                      activity_mci REAL NOT NULL,
                                                                      permitted_date TEXT,
                                                                      recommended_date TEXT,
                                                                      limit_mci REAL)""")
    conn.commit()
    return conn

def log_vials_disposal(vials_full_rows):
    disp_date = datetime.now().strftime(DATE_FORMAT)
    disp_time = datetime.now().strftime(HOUR_FORMAT)
    xlsx_path = get_daily_disposal_excel_path(disp_date)
    wb = ensure_daily_log_workbook(xlsx_path)
    ws = wb["Vials"]
    conn = ensure_daily_log_sqlite()
    cur = conn.cursor()
    for r in vials_full_rows:
        (rid, radionuclide, source_db, calibration_date, stored_at, activity_mci, permitted_date, recommended_date, limit_bq, limit_mci) = r
        ws.append([disp_date, disp_time, radionuclide, calibration_date, stored_at, float(activity_mci), permitted_date, recommended_date,
                   "" if limit_mci is None else float(limit_mci), "" if limit_bq is None else float(limit_bq)])
        cur.execute("INSERT INTO disposed_vials (disposal_date, disposal_time, radionuclide, calibration_date, stored_at, activity_mci, permitted_date, recommended_date, limit_mci, limit_bq) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?)", (disp_date, disp_time, radionuclide, calibration_date, stored_at, float(activity_mci), permitted_date, recommended_date,
                                                      "" if limit_mci is None else float(limit_mci), "" if limit_bq is None else float(limit_bq)))
        mark_vial_as_disposed(source_db=source_db, calibration_date=calibration_date, activity=activity_mci,  disposed_date=disp_date)
    wb.save(xlsx_path)
    conn.commit()
    conn.close()

def append_tc99m_batch(ws, *, batch_name, finalized_at, disposed_dt_str, items_rows, limit_mci=bq_to_mci(1e7)):
    conn = ensure_daily_log_sqlite()
    cur = conn.cursor()
    start_row = ws.max_row + 1
    if ws.max_row == 1 and ws.cell(1, 1).value is None:
        start_row = 1
    title = f"BATCH: {batch_name} | Finalized: {finalized_at or '-'} | Disposed: {disposed_dt_str}"
    ws.cell(row=start_row, column=1, value=title)
    ws.cell(row=start_row, column=1).font = Font("bold")
    ws.cell(row=start_row, column=1).alignment = Alignment(horizontal="left")
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=9)
    ws.cell(row=start_row + 1, column=1, value=f"(radionuclide={TC99M_NUCLIDE}, limit(mCi)={limit_mci})")
    ws.cell(row=start_row + 1, column=1).alignment = Alignment(horizontal="left")
    ws.merge_cells(start_row=start_row + 1, start_column=1, end_row=start_row + 1, end_column=9)
    headers = ["ID", "Stored At", "Activity(mCi)", "Permitted Date", "Recommended Date"]
    header_row = start_row + 2
    for c,h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.font = Font("bold")
        cell.alignment = Alignment(horizontal="center")
    r = header_row + 1
    for item in items_rows:
        iid, stored_at, activity_mci, permitted, recommended = item
        ws.cell(row=r, column=1, value=iid)
        ws.cell(row=r, column=2, value=stored_at)
        ws.cell(row=r, column=3, value=float(activity_mci))
        ws.cell(row=r, column=4, value=permitted)
        ws.cell(row=r, column=5, value=recommended)
        ws.cell(row=r, column=6, value=float(limit_mci))
        r += 1
    ws.append([])
    widths = [8, 16, 12, 16, 16, 8]
    for i,w in enumerate(widths, start=1):
        col = get_column_letter(i)
        cw = ws.column_dimensions[col].width
        if cw is None or cw < w:
            ws.column_dimensions[col].width = w
    for item in items_rows:
        iid, stored_at, activity_mci, permitted, recommended = item
        cur.execute("INSERT INTO disposed_tc99m_batches (batch_id, item_id, stored_at, activity_mci, permitted_date, recommended_date, limit_mci)"
                    "VALUES (?,?,?,?,?,?,?)", (batch_name, iid, stored_at, float(activity_mci), permitted, recommended, float(limit_mci)))
    conn.commit()
    conn.close()

def log_tc99m_batch_disposal(batch_path: str, finalized_at: str, items_rows, limit_mci=bq_to_mci(1e7)):
    disp_date = datetime.now().strftime(DATE_FORMAT)
    disp_time = datetime.now().strftime(HOUR_FORMAT)
    disp_dt_str = f"{disp_date} {disp_time}"
    batch_name = os.path.basename(batch_path)
    xlsx_path = get_daily_disposal_excel_path(disp_date)
    wb = ensure_daily_log_workbook(xlsx_path)
    ws = wb["Tc99m"]
    append_tc99m_batch(ws, batch_name=batch_name, finalized_at=finalized_at, disposed_dt_str=disp_dt_str, items_rows=items_rows, limit_mci=limit_mci)
    wb.save(xlsx_path)

#=====MARK VIAL AS DISPOSED=====
def mark_vial_as_disposed(source_db, calibration_date, activity, disposed_date):
    conn = sqlite3.connect(source_db)
    cur = conn.cursor()
    cur.execute("UPDATE vial_info SET disposal_date=? WHERE cal_date=? AND activity=?", (disposed_date, calibration_date, float(activity)))
    conn.commit()
    conn.close()
    source_excel = os.path.splitext(source_db)[0] + ".xlsx"
    if os.path.exists(source_excel):
        wb = load_workbook(source_excel)
        ws = wb["Vial Info"]
        for row in range(2, ws.max_row + 1):
            cal_date_cell = ws.cell(row=row, column=1).value
            activity_cell = ws.cell(row=row, column=3).value
            try:
                same_activity = float(activity_cell) == float(activity)
            except:
                same_activity = False
            if (
                str(cal_date_cell) == str(calibration_date)
                and same_activity
            ):
                ws.cell(row=row, column=8, value=disposed_date)
                break
        wb.save(source_excel)

#=====VIALS SQLITE + (LIVE) EXCEL=====
def init_vials_storage():
    ensure_dir(VIALS_STORAGE_DIR)
    conn = sqlite3.connect(VIALS_DB)
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS stored_vials (id INTEGER PRIMARY KEY AUTOINCREMENT,
                                                            radionuclide TEXT,
                                                            calibration_date TEXT,
                                                            stored_at TEXT,
                                                            activity_mci REAL,
                                                            permitted_date TEXT,
                                                            recommended_date TEXT,
                                                            limit_mci REAL,
                                                            limit_bq REAL,
                                                            source_db TEXT)""")
    conn.commit()
    conn.close()
    if not os.path.exists(VIALS_XLSX):
        wb = Workbook()
        ws = wb.active
        ws.title = "Stored Vials"
        ws.append(["ID", "Radionuclide", "Calibration Date", "Stored At", "Activity(mCi)", "Permitted Date",
                   "Recommended Date", "Limit(mCi)", "Limit(Bq)", "Source DB"])
        wb.save(VIALS_XLSX)

#=====STORE VIAL IN SQLITE + EXCEL=====
def store_vial(radionuclide, source_db, calibration_date, stored_at, activity_mci, permitted_date=None, recommended_date=None, limit_bq=None):
    if permitted_date is None or recommended_date is None:
        recommended_date, permitted_date, limit_bq = calc_recommended_and_permitted_date(radionuclide, float(activity_mci), stored_at)
    if recommended_date is None or permitted_date is None:
        return None, None
    limit_mci = None
    if limit_bq not in (None, 0):
        limit_mci = round(bq_to_mci(limit_bq), 2)
    conn = sqlite3.connect(VIALS_DB)
    cur = conn.cursor()
    cur.execute("INSERT INTO stored_vials (radionuclide, calibration_date, stored_at, activity_mci, permitted_date, recommended_date, limit_mci, limit_bq, source_db) VALUES (?,?,?,?,?,?,?,?,?)",
                (radionuclide, calibration_date, stored_at, float(activity_mci), permitted_date, recommended_date, None if limit_mci is None else float(limit_mci), None if limit_bq is None else float(limit_bq), source_db))
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    wb = load_workbook(VIALS_XLSX)
    ws = wb["Stored Vials"]
    ws.append([new_id, radionuclide, calibration_date, stored_at, float(activity_mci), permitted_date, recommended_date, "" if limit_mci is None else float(limit_mci), "" if limit_bq is None else float(limit_bq), source_db])
    wb.save(VIALS_XLSX)
    return new_id

#=====READ STORED VIALS FROM BATCH=====
def read_stored_vials():
    init_vials_storage()
    conn = sqlite3.connect(VIALS_DB)
    cur = conn.cursor()
    rows = cur.execute("SELECT id, radionuclide, calibration_date, stored_at, activity_mci, permitted_date, recommended_date, limit_mci FROM stored_vials ORDER BY id").fetchall()
    conn.close()
    return rows

#=====READ VIALS IDS=====
def read_vials_full_ids(ids):
    if not ids:
        return []
    init_vials_storage()
    conn = sqlite3.connect(VIALS_DB)
    cur = conn.cursor()
    placeholders = ",".join(["?"] * len(ids))
    rows = cur.execute(f"SELECT id, radionuclide, source_db, calibration_date, stored_at, activity_mci, permitted_date, recommended_date, limit_bq, limit_mci FROM stored_vials WHERE id IN ({placeholders}) ORDER BY id", ids).fetchall()
    conn.close()
    return rows

#=====DELETE VIALS BY IDS=====
def delete_vials_by_ids(ids):
    if not ids:
        return
    init_vials_storage()
    conn = sqlite3.connect(VIALS_DB)
    cur = conn.cursor()
    placeholder = ",".join(["?"] * len(ids))
    cur.execute(f"DELETE FROM stored_vials WHERE id IN ({placeholder})", ids)
    conn.commit()
    conn.close()
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Stored Vials"
        ws.append(["ID", "Radionuclide", "Calibration Date", "Stored At", "Activity(mCi)", "Permitted Date",
                    "Recommended Date", "Limit(mCi)", "Limit(Bq)", "Source DB"])
        conn2 = sqlite3.connect(VIALS_DB)
        cur2 = conn2.cursor()
        rows = cur2.execute("SELECT id, radionuclide, calibration_date, stored_at, activity_mci, permitted_date, recommended_date, limit_mci, limit_bq, source_db FROM stored_vials ORDER BY id").fetchall()
        conn2.close()
        for r in rows:
            ws.append(list(r))
        wb.save(VIALS_XLSX)
    except Exception:
        pass

#=====TC99M SQLITE CREATE TABLES=====
def init_tc99m_registry():
    ensure_dir(TC99M_DIR)
    conn = sqlite3.connect(TC99M_REGISTRY_DB)
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS batches (id INTEGER PRIMARY KEY AUTOINCREMENT, folder_path TEXT, created_at TEXT, finalized_at TEXT, disposed_at TEXT)""")
    cols = [r[1] for r in cur.execute("PRAGMA table_info(batches)").fetchall()]
    if "disposed_at" not in cols:
        cur.execute("ALTER TABLE batches ADD COLUMN disposed_at TEXT")
    conn.commit()
    conn.close()

#=====CREATE NEW BATCH=====
def create_new_batch_folder():
    init_tc99m_registry()
    year = datetime.now().strftime("%Y")
    creation_date = datetime.now().strftime(DATE_FORMAT)
    year_dir = os.path.join(TC99M_DIR, year)
    ensure_dir(year_dir)
    batch_folder = f"Batch__{creation_date}"
    batch_path = os.path.join(year_dir, batch_folder)
    if not os.path.exists(batch_path):
        ensure_dir(batch_path)
    conn = sqlite3.connect(TC99M_REGISTRY_DB)
    cur = conn.cursor()
    cur.execute("SELECT 1 FROM batches WHERE folder_path=?", (batch_path,))
    exists = cur.fetchone()
    if not exists:
        cur.execute("INSERT INTO batches (folder_path, created_at, finalized_at) VALUES (?,?,NULL)", (batch_path, creation_date))
    cur.execute("INSERT OR REPLACE INTO settings (key,value) VALUES ('active_batch',?)", (batch_path,))
    conn.commit()
    conn.close()
    return batch_path

#=====GET ACTIVE BATCH (IF IT DOESN'T EXIST CREATE NEW)=====
def get_active_batch():
    init_tc99m_registry()
    conn = sqlite3.connect(TC99M_REGISTRY_DB)
    cur = conn.cursor()
    row = cur.execute("SELECT value FROM settings WHERE key='active_batch'").fetchone()
    conn.close()
    if row and os.path.isdir(row[0]):
        return row[0]
    return create_new_batch_folder()

#=====FINALIZE CURRENT ACTIVE BATCH AND START NEW=====
def finalize_active_batch():
    init_tc99m_registry()
    old_batch = get_active_batch()
    finalized_date = datetime.now().strftime(DATE_FORMAT)
    conn = sqlite3.connect(TC99M_REGISTRY_DB)
    cur = conn.cursor()
    cur.execute("UPDATE batches SET finalized_at=? WHERE folder_path=? AND finalized_at IS NULL", (finalized_date, old_batch))
    conn.commit()
    conn.close()
    new_batch = create_new_batch_folder()
    return old_batch, new_batch

#=====DISPOSE BATCH=====
def dispose_batch(batch_path):
    init_tc99m_registry()
    disposed_date = datetime.now().strftime(DATE_FORMAT)
    conn = sqlite3.connect(TC99M_REGISTRY_DB)
    cur = conn.cursor()
    cur.execute("UPDATE batches SET disposed_at=? WHERE folder_path=? AND disposed_at IS NULL", (disposed_date, batch_path))
    conn.commit()
    conn.close()

#=====READ BATCH DATE INFO=====
def read_batch_info(batch_path):
    init_tc99m_registry()
    conn = sqlite3.connect(TC99M_REGISTRY_DB)
    cur = conn.cursor()
    row = cur.execute("SELECT created_at, finalized_at, disposed_at FROM batches WHERE folder_path=?", (batch_path,)).fetchone()
    conn.close()
    if not row:
        return None, None, None
    return row[0], row[1], row[2]

#=====CREATE SQLITE+XLSX FILES INSIDE BATCH FOLDER=====
def init_storage_files(batch_path):
    db_path = os.path.join(batch_path, "storage.sqlite")
    xlsx_path = os.path.join(batch_path, "storage.xlsx")
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS stored_items (id INTEGER PRIMARY KEY AUTOINCREMENT,
                                                            stored_at TEXT,
                                                            activity_mci REAL,
                                                            permitted_date TEXT,
                                                            recommended_date TEXT)""")
    conn.commit()
    conn.close()
    if not os.path.exists(xlsx_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Stored Items"
        ws.append(["ID", "Stored at", "Activity(mCi)", "Permitted Date", "Recommended Date"])
        wb.save(xlsx_path)
    return db_path, xlsx_path

#=====STORE TC99M ITEM IN SQLITE+XLSX=====
def store_tc99m_item(stored_at, activity_mci, permitted_date=None, recommended_date=None):
    batch_path = get_active_batch()
    db_path, xlsx_path = init_storage_files(batch_path)
    if permitted_date is None or recommended_date is None:
        recommended_date, permitted_date, _ = calc_recommended_and_permitted_date(radionuclide=TC99M_NUCLIDE, activity_mci=float(activity_mci), stored_at=stored_at)
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("INSERT INTO stored_items (stored_at, activity_mci, permitted_date, recommended_date) VALUES (?,?,?,?)",
                (stored_at, float(activity_mci), permitted_date, recommended_date))
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    wb = load_workbook(xlsx_path)
    ws = wb["Stored Items"]
    ws.append([new_id, stored_at, float(activity_mci), permitted_date, recommended_date])
    wb.save(xlsx_path)
    return new_id, batch_path

#=====READ ITEMS INFO=====
def read_tc99m_items(batch_path=None):
    if batch_path is None:
        batch_path = get_active_batch()
    db_path = os.path.join(batch_path, "storage.sqlite")
    if not os.path.exists(db_path):
        return []
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    rows = cur.execute("SELECT id, stored_at, activity_mci, permitted_date, recommended_date FROM stored_items ORDER BY id").fetchall()
    conn.close()
    return rows