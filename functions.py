from constants import *
from tkinter import Frame, Scrollbar, Canvas, messagebox, Button, END
import datetime, os
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
    now = datetime.datetime.now().strftime("%H:%M")
    time_entry.delete(0, END)
    time_entry.insert(0, now)

#Create Excel for Vials
def create_excel_for_vial(excel_path):
    if os.path.exists(excel_path):
        return
    wb = Workbook()
    ws_info = wb.active
    ws_info.title = "Vial Info"
    ws_info.append(["Date", "Time", "Activity(mCi)", "Volume(ml)", "Concentration(mCi/ml)", "Expiration Date", "Stored Date"])
    ws_admin = wb.create_sheet("Administrations")
    ws_admin.append(["ID", "Date", "Time", "Patient Name", "Concentration(mCi/ml)", "Dose(mCi)", "Volume(ml)", "Volume Left(ml)"])
    wb.save(excel_path)

#Create Excel for Tc99m Gen
def create_excel_for_tc99m(excel_path):
    if not os.path.exists(excel_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Gen Info"
        ws.append(["Gen ID", "Calibration Date", "Calibration Time", "Mo99 Activity (mCi)", "Start Date", "Expiration Date", "Disposal Date"])
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
        ws3.append(["Date", "Patient", "Weight (kg)", "Admin Time", "Dose (mCi)", "Concentration (mCi/ml)", "Volume (ml)", "Real Dose (mCi)", "ITLC(<2%)", "Residual (mCi)"])
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
    now = datetime.datetime.now()
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

#Dispose Gens
def dispose_gen(*, conn, dbfile, excel_sheet="Gen Info", date_format="%d-%m-%Y", on_disposed_callback=None):
    if not messagebox.askyesno("Dispose Generator", "Are you sure you want to dispose this generator?\nThis action cannot be undone."):
        return False
    disposal_date = datetime.datetime.now().strftime(date_format)
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
def update_header_and_disable(header, tab, is_disposed=False, is_expired=False):
    if is_disposed:
        header.config(text="⚠ GENERATOR DISPOSED – NO FURTHER ACTIONS ALLOWED", fg="red")
        disable_buttons(tab, exempt_texts=["Back", "Load"])
    elif is_expired:
        header.config(text="⚠ GENERATOR EXPIRED – NO FURTHER ACTIONS ALLOWED", fg="orange")
        disable_buttons(tab, exempt_texts=["Back", "Load", "✗Dispose Gen✗"])
