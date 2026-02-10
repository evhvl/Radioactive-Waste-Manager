import tkinter
from tkinter import *
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import math, os, sqlite3, pandas
from tkcalendar import DateEntry
from openpyxl import load_workbook, Workbook
from pathlib import Path
from constants import *
import disposal
from disposal import decay_activity


#=======================================================================================================================================================================#
def center_window(window, w, h):
    window.resizable(False, False)
    window.update_idletasks()
    screen_w = window.winfo_screenwidth()
    screen_h = window.winfo_screenheight()
    x = (screen_w // 2) - (w // 2)
    y = (screen_h // 2) - (h // 2)
    window.geometry(f"{w}x{h}+{x}+{y}")
    window.grab_set()

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

def update_time(time_entry):
    now = datetime.now().strftime("%H:%M")
    time_entry.delete(0, END)
    time_entry.insert(0, now)

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

def create_excel_for_ga68(excel_path):
    if not os.path.exists(excel_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Gen Info"
        ws.append(["Gen ID", "Model", "Start Date", "Activity (MBq)", "Expiration Date", "Disposal Date"])
        ws2 = wb.create_sheet("Elutions")
        ws2.append(["", "Date", "Time", "Activity(mCi)"])
        ws3 = wb.create_sheet("DOTATOC")
        ws3.append(["Date", "Patient", "Weight (kg)", "Admin Time", "Dose (mCi)", "Concentration (mCi/ml)", "Volume (ml)", "Real Dose (mCi)", "ITLC(<2%)", "Residual (mCi)"])
        wb.save(excel_path)

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

def disable_buttons(parent, exempt_texts=None):
    if exempt_texts is None:
        exempt_texts = []
    for widget in parent.winfo_children():
        if isinstance(widget, Button):
            if widget.cget("text") not in exempt_texts:
                widget.config(state="disabled")
        elif widget.winfo_children():
            disable_buttons(widget, exempt_texts)

def update_header_and_disable(header, tab, is_disposed=False, is_expired=False):
    if is_disposed:
        header.config(text="⚠ GENERATOR DISPOSED – NO FURTHER ACTIONS ALLOWED", fg="red")
    elif is_expired:
        header.config(text="⚠ GENERATOR EXPIRED – NO FURTHER ACTIONS ALLOWED", fg="orange")
    disable_buttons(tab, exempt_texts=["Load"])
#=======================================================================================================================================================================#
class Functions:

    def __init__(self, window, tabs_frame, main_tab):
        self.window = window
        self.tabs_frame = tabs_frame
        self.main_tab = main_tab

    def back_to_main(self, tab_name):
        self.tabs_frame.forget(tab_name)
        self.tabs_frame.select(self.main_tab)

    #=====CREATE TABS=====
    def create_vials_tab(self, isotope):
        for tab in self.tabs_frame.tabs():
            if self.tabs_frame.tab(tab, "text") == isotope:
                self.tabs_frame.select(tab)
                return
        new_tab = Frame(self.tabs_frame, bg=C4)
        self.tabs_frame.add(new_tab, text=isotope)
        self.tabs_frame.select(new_tab)

    def create_new_tab(self, tab_name):
        for tab in self.tabs_frame.tabs():
            if self.tabs_frame.tab(tab, "text") == tab_name:
                self.tabs_frame.select(tab)
                return
        new_tab = Frame(self.tabs_frame, bg=C4)
        self.tabs_frame.add(new_tab, text=tab_name)
        self.tabs_frame.select(new_tab)
        if tab_name == "Vials":
            self._tab_vials(new_tab)
        elif tab_name == "Generators":
            self._tab_generators(new_tab)
        elif tab_name == "I131":
            self._tab_i131(new_tab)
        elif tab_name == "Tc99m Gen":
            self._tab_tc99m(new_tab)
        elif tab_name == "Ga68 Gen":
            self._tab_ga68(new_tab)
        elif tab_name in [name for name, _ in VIAL_DATA]:
            self._tab_vial(new_tab, tab_name)
        elif tab_name == "Disposal":
            self._tab_disposal(new_tab)

    #=====CUSTOMIZE EACH TAB=====
    def _tab_vials(self, tab):
        content_frame = Frame(tab, bg=C4)
        content_frame.place(relx=0.5, rely=0.5, anchor="center")
        Label(content_frame, text="Select New Radionuclide:", bg=C4, fg="white", font=(FONT_NAME, 24, "bold")).grid(columnspan=3, column=0, row=0, pady=(0, 20))
        for idx, (name,_) in enumerate(VIAL_DATA):
            r = idx // 3 + 1
            c = idx % 3
            Button(content_frame, text=name, **TAB_BUTTON_STYLE, command=lambda t=name: self.create_new_tab(t)).grid(column=c, row=r, padx=10, pady=10)
        Button(content_frame, text="Back", **TAB_BUTTON_STYLE, command=lambda nt=tab: self.back_to_main(nt)).grid(columnspan=3, column=0, row=5, pady=(30, 0))

    def _tab_generators(self, tab):
        frame = Frame(tab, bg=C4)
        frame.pack(expand=True)
        Button(frame, text="Tc99m Gen", **GEN_BUTTON_STYLE, command=lambda: self.create_new_tab("Tc99m Gen")).grid(row=0, column=0,sticky="e", padx=20, pady=10)
        Button(frame, text="Ga68 Gen", **GEN_BUTTON_STYLE, command=lambda : self.create_new_tab("Ga68 Gen")).grid(row=0, column=1, sticky="w", padx=20, pady=10)
        Button(tab, text="Back", **TAB_BUTTON_STYLE, command=lambda nt=tab: self.back_to_main(nt)).pack(pady=40)

    def _tab_disposal(self, tab):
        disposal.build_disposal_tab(tab, on_back=lambda: self.back_to_main(tab))

#=========================================================================================VIALS====================================================================================
    def _tab_vial(self, tab, vial_name):
        #Choose New or Old File
        def select_vial_file():
            popup_window = Toplevel(self.window)
            popup_window.title("Choose File")
            popup_window.config(bg=C4)
            center_window(window=popup_window, w=350, h=250)
            Label(popup_window, text=f"Select {vial_name} File:", **TEXT_COLORS, font=(FONT_NAME, 16, "bold")).pack(pady=20)
            def create_new():
                popup_window.destroy()
                new_vial_file()
            def open_existing():
                popup_window.destroy()
                existing_vial_file()
            Button(popup_window, text="New File", **TAB_BUTTON_STYLE, command=create_new).pack(pady=10)
            Button(popup_window, text="Old File", **TAB_BUTTON_STYLE, command=open_existing).pack()
        #Create New File
        def new_vial_file():
            popup = Toplevel(self.window)
            popup.title(f"New {vial_name} Vial")
            popup.config(bg=C4)
            center_window(window=popup, w=350, h=290)
            Label(popup, text=f"New {vial_name} Vial Info", **TEXT_COLORS, font=(FONT_NAME, 16, "bold")).pack(pady=10)
            info_frame = Frame(popup, bg=C4)
            info_frame.pack(pady=10)
            Label(info_frame, text="Date:", **TEXT_COLORS).grid(row=0, column=0, sticky="e", padx=5, pady=5)
            date_entry = DateEntry(info_frame, width=12, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
            date_entry.grid(row=0, column=1, pady=5)
            time_field = Frame(info_frame, bg="white", highlightbackground="black", highlightthickness=0)
            time_field.grid(row=1, column=1, padx=5)
            Label(info_frame, text="Time:", **TEXT_COLORS).grid(row=1, column=0, sticky="e", padx=5, pady=5)
            time_entry = Entry(time_field, width=9, bd=0, font=(FONT_NAME, 10))
            time_entry.pack(side="left", padx=(3, 0), pady=2)
            update_time(time_entry)
            refresh_time_button = Button(time_field, text="↻", command=lambda nt=time_entry: update_time(nt), bg="white", fg="black", bd=0,
                                         padx=3, pady=0, font=(FONT_NAME, 10), cursor="hand2")
            refresh_time_button.pack(side="right", padx=3)
            Label(info_frame, text="Activity (mCi):", **TEXT_COLORS).grid(row=2, column=0, sticky="e", padx=5, pady=5)
            activity_entry = Entry(info_frame, width=14)
            activity_entry.grid(row=2, column=1, pady=5)
            Label(info_frame, text="Volume (ml):", **TEXT_COLORS).grid(row=3, column=0, sticky="e", padx=5, pady=5)
            volume_entry = Entry(info_frame, width=14)
            volume_entry.grid(row=3, column=1, pady=5)
            Label(info_frame, text="Expiration Date:", **TEXT_COLORS).grid(row=4, column=0, sticky="e", padx=5, pady=5)
            expiration_entry = DateEntry(info_frame, width=12, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
            expiration_entry.grid(row=4, column=1, pady=5)
            def save_new_vial_file():
                fields = {
                    "Date": date_entry,
                    "Time": time_entry,
                    "Activity(mCi)": activity_entry,
                    "Volume(ml)": volume_entry,
                    "Expiration Date": expiration_entry
                }
                values = {}
                for name, entry in fields.items():
                    val = entry.get().strip()
                    if not val:
                        messagebox.showerror("Error", f"Please enter {name}")
                        return
                    values[name] = val
                try:
                    activity = float(values["Activity(mCi)"])
                    volume = float(values["Volume(ml)"])
                except ValueError:
                    messagebox.showerror("Error", "Activity & Volume must be numbers!")
                    return
                conc = round(activity / volume, 2)
                date = date_entry.get()
                time = time_entry.get()
                exp_date = expiration_entry.get()
                base_dir = "Vials"
                dt = datetime.strptime(date, "%d-%m-%Y")
                year = dt.strftime("%Y")
                month = dt.strftime("%m")
                vial_dir = os.path.join(base_dir, vial_name, year, month, f"{vial_name}__{date}")
                os.makedirs(vial_dir, exist_ok=True)
                db_path = os.path.join(vial_dir, f"{vial_name}__{date}.sqlite")
                excel_path = os.path.join(vial_dir, f"{vial_name}__{date}.xlsx")
                conn = sqlite3.connect(db_path)
                cur = conn.cursor()
                cur.execute("""CREATE TABLE IF NOT EXISTS vial_info(date TEXT, time TEXT, activity REAL, volume REAL, concentration REAL, expiration_date TEXT, stored_date TEXT)""")
                cur.execute("""CREATE TABLE IF NOT EXISTS patient_info(id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, time TEXT, patient_name TEXT, concentration REAL, dose_planned REAL, volume_planned REAL, dose_actual REAL, volume_actual REAL, volume_left REAL)""")
                cur.execute("""INSERT INTO vial_info VALUES (?,?,?,?,?,?,NULL)""", (date, time, activity, volume, conc, exp_date))
                conn.commit()
                conn.close()
                create_excel_for_vial(excel_path)
                append_row_to_sheet(excel_path, "Vial Info", [date, time, activity, volume, conc, exp_date, ""])
                popup.destroy()
                load_vial(db_path)
            bttn_frame = Frame(popup, bg=C4)
            bttn_frame.pack()
            Button(bttn_frame, text="Save File",
                   **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=8,
                   height=1, font=(FONT_NAME, 10, "bold"), command=save_new_vial_file).grid(row=0, column=0, padx=10, pady=10)
            Button(bttn_frame, text="Back", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=8, height=1, font=(FONT_NAME, 10, "bold"),
                   command=lambda: (popup.destroy(), self.tabs_frame.forget(tab), self.create_new_tab("Vials"))).grid(row=0, column=1, padx=10, pady=10)
        #Open Old File
        def existing_vial_file():
            popup = Toplevel(self.window)
            popup.title(f"Open Existing {vial_name} Vial File")
            popup.config(bg=C4)
            center_window(window=popup, w=360, h=130)
            Label(popup, text=f"Select Existing {vial_name} Folder", **TEXT_COLORS, font=(FONT_NAME, 17, "bold")).pack(pady=10)
            def open_folder():
                base_dir = Path(__file__).resolve().parent / "Vials"
                initial_dir = find_last_folder(base_dir=base_dir, subfolder=vial_name)
                folder = filedialog.askdirectory(title="Select Vial Folder", initialdir=initial_dir)
                if not folder:
                    return
                sqlite_files = [f for f in os.listdir(folder) if f.endswith(".sqlite")]
                if not sqlite_files:
                    messagebox.showerror("Error", "No .sqlite file found.")
                    return
                popup.destroy()
                load_vial(os.path.join(folder, sqlite_files[0]))
            button_frame = Frame(popup, bg=C4)
            button_frame.pack()
            Button(button_frame, text="Open File 🗁", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                   width=12, height=2, font=(FONT_NAME, 12, "bold"), command=open_folder).grid(row=0, column=0, padx=10, pady=10)
            Button(button_frame, text="Back", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=12, height=2, font=(FONT_NAME, 12, "bold"),
                   command=lambda: (popup.destroy(), self.tabs_frame.forget(tab),self.create_new_tab("Vials"))).grid(row=0, column=1, padx=10, pady=10)
        def load_vial(dbfile):
            for widget in tab.winfo_children():
                widget.destroy()
            header = Label(tab, text=f"{vial_name} Log Sheet", **TEXT_COLORS, font=(FONT_NAME, 18, "bold"))
            header.pack(pady=(5,0), fill="x")
            conn = sqlite3.connect(dbfile)
            cur = conn.cursor()
            date, time, activity, volume, conc, exp_date, stored_date = cur.execute("SELECT * FROM vial_info").fetchone()
            exp_dt = datetime.strptime(exp_date, "%d-%m-%Y")
            today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            is_stored = False
            if stored_date:
                stored_dt = datetime.strptime(stored_date, "%d-%m-%Y")
                if today >= stored_dt:
                    is_stored = True
            is_expired = today > exp_dt
            if is_stored or is_expired:
                messagebox.showerror("Vial Disposed/Expired","This vial has reached its disposal/expiration date.\nNo further administrations are allowed.")
            half_life = next(hl for name, hl in VIAL_DATA if name == vial_name)
            info_frame = Frame(tab, bg=C4)
            info_frame.pack(anchor="center", pady=20)
            Label(info_frame, text=f"Date: {date}", **TEXT_COLORS, font=(FONT_NAME, 10)).grid(row=0, column=0, padx=6, pady=6)
            Label(info_frame, text=f"Activity (mCi): {activity}", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=1, column=0, padx=6, pady=6)
            Label(info_frame, text=f"Conc (mCi/ml): {conc}", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=2, column=0, padx=6, pady=6)
            Label(info_frame, text=f"T1/2 {vial_name} (HR): {half_life}", **TEXT_COLORS, font=(FONT_NAME, 10)).grid(row=0, column=1, padx=6, pady=6)
            Label(info_frame, text=f"Expiration Date: {exp_date}", **TEXT_COLORS, font=(FONT_NAME, 10)).grid(row=1, column=1, padx=6, pady=6)
            #Store Vial
            def store_current_vial():
                if not messagebox.askyesno("Store Vial", "Are you sure you want to store this vial for disposal?\nNo further administrations will be allowed."):
                    return
                last_left = cur.execute("SELECT volume_left FROM patient_info ORDER BY id DESC LIMIT 1").fetchone()
                volume_left_now = float(last_left[0]) if last_left and last_left[0] is not None else float(volume)
                if volume_left_now <= 0:
                    messagebox.showerror("Store Vial", "No volume left in vial.")
                    return
                vial_dt = datetime.strptime(f"{date} {time}", "%d-%m-%Y %H:%M")
                now_dt = datetime.now()
                delta_minutes = (now_dt - vial_dt).total_seconds() / 60
                if delta_minutes < 0:
                    delta_minutes = 0
                decay_factor = math.exp(-math.log(2) * delta_minutes / (half_life * 60))
                current_conc = float(conc) * decay_factor
                current_activity_mci = round(current_conc * volume_left_now, 3)
                stored_at = datetime.now().strftime("%d-%m-%Y")
                recommended, permitted, limit_bq = disposal.calc_recommended_and_permitted_date(vial_name, float(current_activity_mci), stored_at)
                disposal.store_vial(radionuclide=vial_name, source_db=dbfile, calibration_date=date, stored_at=stored_at, activity_mci=current_activity_mci, permitted_date=permitted, recommended_date=recommended, limit_bq=limit_bq)
                cur.execute("UPDATE vial_info SET stored_date=?", (stored_at,))
                conn.commit()
                folder = os.path.dirname(dbfile)
                excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                wb = load_workbook(excel_path)
                ws = wb["Vial Info"]
                ws.cell(row=2, column=7, value=stored_at)
                wb.save(excel_path)
                messagebox.showinfo("Vial Stored", f"{vial_name} vial stored successfully.\n\nRecommended disposal after: {recommended}\nPermitted disposal after: {permitted}")
            dispose_button = Button(info_frame, text="✗Store Vial✗", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=14, height=1, font=(FONT_NAME, 10, "bold"), command=store_current_vial)
            dispose_button.grid(row=2, column=1, padx=6, pady=6)
            #Table
            columns = [("date", "Date", 120), ("time", "Time", 100), ("patient_name", "Patient", 160), ("concentration", "Conc(mCi/ml)", 120), ("dose", "Dose(mCi)", 100), ("volume", "Vol(ml)", 90), ("volume_left", "Vol Left(ml)", 100)]
            tree = ttk.Treeview(tab, columns=[c[0] for c in columns], show="headings", height=8)
            tree.pack(pady=10)
            for col_id, col_title, col_width in columns:
                tree.heading(col_id, text=col_title)
                tree.column(col_id, width=col_width, anchor="center")
            style = ttk.Style()
            style.theme_use("default")
            style.configure("Treeview", background=C2, fieldbackground=C2, foreground="black", rowheight=26, borderwidth=1, bordercolor="black", relief="solid")
            style.configure("Treeview.Heading", background=C3, foreground="white", font=(FONT_NAME, 11, "bold"), relief="solid")
            style.map("Treeview", background=[("selected", "#8FAADC"), ("!selected", C2)], foreground=[("selected", "black")])
            style.layout("Treeview", [("Treeview.treearea", {"sticky": "nsew"})])
            #Load Data
            rows = cur.execute("SELECT id,date,time,patient_name,concentration,dose_planned,volume_planned,dose_actual,volume_actual,volume_left FROM patient_info").fetchall()
            for r in rows:
                row_id = r[0]
                (date, time, patient, conc, dose_p, vol_p, dose_a, vol_a, vol_left) = r[1:]
                dose_txt = f"{dose_p:.2f}" if dose_p is not None else "-"
                vol_txt = f"{vol_p:.2f}" if vol_p is not None else "-"
                if dose_a is not None:
                    dose_txt += f" → {dose_a:.2f}"
                if vol_a is not None:
                    vol_txt += f" → {vol_a:.2f}"
                tree.insert("", "end", iid=row_id, values=(date, time, patient, f"{conc:.2f}" if conc is not None else "-", dose_txt, vol_txt, f"{vol_left:.2f}" if vol_left is not None else "-"))
            # Add New Data
            add_frame = Frame(tab, bg=C4)
            add_frame.pack(pady=10)
            Label(add_frame, text="Date:", **TEXT_COLORS).grid(row=0, column=0)
            admin_date_entry = DateEntry(add_frame, width=10, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
            admin_date_entry.grid(row=0, column=1, padx=5)
            time_field = Frame(add_frame, bg="white", highlightbackground="black", highlightthickness=0)
            time_field.grid(row=0, column=3, padx=5)
            Label(add_frame, text="Admin Time:", **TEXT_COLORS).grid(row=0, column=2, sticky="e", padx=5)
            admin_time_entry = Entry(time_field, width=5, bd=0, font=(FONT_NAME, 10))
            admin_time_entry.pack(side="left", padx=(3, 0), pady=2)
            update_time(admin_time_entry)
            refresh_time_button = Button(time_field, text="↻", command=lambda nt=admin_time_entry: update_time(nt), bg="white",
                                         fg="black", bd=0, padx=3, pady=0, font=(FONT_NAME, 10), cursor="hand2")
            refresh_time_button.pack(side="right", padx=3)
            Label(add_frame, text="Patient Name:", **TEXT_COLORS).grid(row=0, column=4, sticky="e", padx=5)
            patient_name_entry = Entry(add_frame, width=18)
            patient_name_entry.insert(0, "-")
            patient_name_entry.grid(row=0, column=5, padx=5)
            Label(add_frame, text="Dose (mCi):", **TEXT_COLORS).grid(row=0, column=6, sticky="e", padx=5)
            dose_entry = Entry(add_frame, width=8)
            dose_entry.grid(row=0, column=7, padx=5)
            def add_record():
                if not patient_name_entry.get().strip():
                    messagebox.showerror("Error", "Please enter Patient Name.")
                    return
                try:
                    dose = float(dose_entry.get())
                    if dose <= 0:
                        raise ValueError
                except ValueError:
                    messagebox.showerror("Error", "Enter valid Dose (mCi).")
                    return
                admin_date = admin_date_entry.get()
                admin_time = admin_time_entry.get()
                admin_dt = datetime.strptime(f"{admin_date} {admin_time}", "%d-%m-%Y %H:%M")
                vial_dt = datetime.strptime(f"{date} {time}", "%d-%m-%Y %H:%M")
                delta_minutes = (admin_dt - vial_dt).total_seconds() / 60
                if delta_minutes < 0:
                    messagebox.showerror("Error", "Administration Time cannot be before Vial Calibration.")
                    return
                decay_factor = math.exp(-math.log(2) * delta_minutes / (half_life * 60))
                updated_conc = round(conc * decay_factor, 2)
                dose_volume = round(dose / updated_conc, 1)
                prev = cur.execute("SELECT volume_left FROM patient_info ORDER BY id DESC LIMIT 1").fetchone()
                if prev:
                    prev_volume_left = prev[0]
                else:
                    prev_volume_left = volume
                volume_left = round(prev_volume_left - dose_volume, 1)
                if volume_left < 0:
                    messagebox.showerror("Error", "Not enough volume left in Vial.")
                    return
                cur.execute("""INSERT INTO patient_info (date, time, patient_name, concentration, dose_planned, volume_planned, dose_actual, volume_actual, volume_left) VALUES (?,?,?,?,?,?,?,?,?)""",
                            (admin_date, admin_time, patient_name_entry.get().strip(), updated_conc, dose, dose_volume, None, None, volume_left))
                conn.commit()
                row_id = cur.lastrowid
                tree.insert("", "end", iid=row_id, values=(admin_date, admin_time, patient_name_entry.get().strip(), f"{updated_conc:.2f}", dose, f"{dose_volume:.1f}", f"{volume_left:.1f}"))
                def on_double_click(event):
                    selected = tree.selection()
                    if not selected:
                        return
                    row_id = int(selected[0])
                    popup = Toplevel(tab)
                    popup.title("Insert Actual Administration Values")
                    popup.config(bg=C4, pady=15)
                    center_window(popup, 250, 150)
                    Label(popup, text="Actual Dose (mCi):", **TEXT_COLORS, font=(FONT_NAME,10,"bold")).grid(row=0, column=0, padx=10, pady=5)
                    dose_actual_entry = Entry(popup, width=8)
                    dose_actual_entry.insert(0, f"{dose}")
                    dose_actual_entry.grid(row=0, column=1)
                    Label(popup, text="Actual Volume (ml):", **TEXT_COLORS, font=(FONT_NAME,10,"bold")).grid(row=1, column=0, padx=10, pady=5)
                    volume_actual_entry = Entry(popup, width=8)
                    volume_actual_entry.insert(0, f"{dose_volume}")
                    volume_actual_entry.grid(row=1, column=1)
                    def save_actual():
                        try:
                            dose_actual = float(dose_actual_entry.get())
                            volume_actual = float(volume_actual_entry.get())
                        except ValueError:
                            messagebox.showerror("Error", "Invalid values.")
                            return
                        prev = cur.execute("""SELECT volume_left FROM patient_info WHERE id < ? ORDER BY id DESC LIMIT 1""", (row_id,)).fetchone()
                        prev_left = prev[0] if prev else volume
                        new_left = round(prev_left - volume_actual, 1)
                        if new_left < 0:
                            messagebox.showerror("Error", "Not enough volume left.")
                            return
                        cur.execute("""UPDATE patient_info SET dose_actual=?, volume_actual=?, volume_left=? WHERE id=?""", (dose_actual, volume_actual, new_left, row_id))
                        conn.commit()
                        planned = cur.execute("""SELECT dose_planned, volume_planned FROM patient_info WHERE id=?""", (row_id,)).fetchone()
                        tree.item(row_id, values=(tree.item(row_id, "values")[0], tree.item(row_id, "values")[1], tree.item(row_id, "values")[2], tree.item(row_id, "values")[3],
                                                  f"{planned[0]:.2f} → {dose_actual:.2f}", f"{planned[1]:.2f} → {volume_actual:.2f}", f"{new_left:.2f}"))
                        folder = os.path.dirname(dbfile)
                        excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                        append_row_to_sheet(excel_path, "Administrations",[row_id, admin_date, admin_time, patient_name_entry.get(), updated_conc, f"{dose} → {dose_actual}", f"{dose_volume} → {volume_actual}", new_left])
                        popup.destroy()
                    Button(popup, text="OK", command=save_actual, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                           width=10, height=2, font=(FONT_NAME, 10, "bold")).grid(row=2, column=0, pady=10, padx=6)
                    Button(popup, text="Cancel", command=popup.destroy, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['bg', 'width', 'height', 'font']},
                           bg=C4, width=10, height=2, font=(FONT_NAME, 10, "bold")).grid(row=2, column=1, pady=10, padx=6)
                tree.bind("<Double-1>", on_double_click)
                dose_entry.delete(0, "end")
            add_button = Button(add_frame, text="Add", command=add_record, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                   width=5, height=1, font=(FONT_NAME, 10, "bold"))
            add_button.grid(row=0, column=8, padx=6)
            if is_expired:
                messagebox.showerror("Expired Vial", "This vial has expired. Administration is not allowed.")
                add_button.config(state="disabled")
                Label(tab, text="⚠ VIAL EXPIRED – NO ADMINISTRATION ALLOWED", fg="red", bg=C4, font=(FONT_NAME, 11, "bold")).pack(pady=(5,0))
            elif is_stored:
                Label(tab, text=f"⚠ VIAL STORED / DISPOSAL DATE REACHED ({stored_date})\nNO FURTHER ACTIONS ALLOWED", fg="red", bg=C4, font=(FONT_NAME, 11, "bold"), justify="center").pack(pady=(5, 0))
                disable_buttons(tab, exempt_texts=["Back"])
            #Delete Data
            def update_volume_after_delete():
                start_volume = float(cur.execute("SELECT volume FROM vial_info").fetchone()[0])
                rows = cur.execute("""SELECT id, volume_planned, volume_actual FROM patient_info ORDER BY id""").fetchall()
                current_vol_left = start_volume
                for rid, vol_p_db, vol_a_db in rows:
                    used = vol_a_db if vol_a_db is not None else vol_p_db
                    used = float(used) if used is not None else 0.0
                    current_vol_left = round(current_vol_left - used, 1)
                    if current_vol_left < 0:
                        current_vol_left = 0.0
                    cur.execute("UPDATE patient_info SET volume_left=? WHERE id=?", (current_vol_left, rid))
                conn.commit()
                folder = os.path.dirname(dbfile)
                excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                wb = load_workbook(excel_path)
                ws = wb["Administrations"]
                excel_row_map = {row[0].value: row[0].row for row in ws.iter_rows(min_row=2)}
                for rid, _, _ in rows:
                    left_val = cur.execute("SELECT volume_left FROM patient_info WHERE id=?", (rid,)).fetchone()[0]
                    if rid in excel_row_map:
                        ws.cell(row=excel_row_map[rid], column=8, value=left_val)
                wb.save(excel_path)
            def delete_record():
                selected = tree.selection()
                if not selected:
                    messagebox.showerror("Error", "No row selected.")
                    return
                row_id = int(selected[0])
                if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected record?"):
                    return
                cur.execute("DELETE FROM patient_info WHERE id=?", (row_id,))
                conn.commit()
                folder = os.path.dirname(dbfile)
                excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                wb = load_workbook(excel_path)
                ws = wb["Administrations"]
                for row in ws.iter_rows(min_row=2):
                    if row[0].value == row_id:
                        ws.delete_rows(row[0].row)
                        break
                wb.save(excel_path)
                update_volume_after_delete()
                tree.delete(*tree.get_children())
                rows = cur.execute("""SELECT id, date, time, patient_name, concentration, dose_planned, volume_planned, dose_actual, volume_actual, volume_left FROM patient_info ORDER BY id""").fetchall()
                for r in rows:
                    row_id = r[0]
                    (date, time, patient, conc, dose_p, vol_p, dose_a, vol_a, vol_left) = r[1:]
                    dose_txt = f"{dose_p:.1f}" if dose_p is not None else "-"
                    vol_txt = f"{vol_p:.1f}" if vol_p is not None else "-"
                    if dose_a is not None:
                        dose_txt += f" → {dose_a:.1f}"
                    if vol_a is not None:
                        vol_txt += f" → {vol_a:.1f}"
                    tree.insert("", "end", iid=row_id, values=(date, time, patient, f"{conc:.2f}" if conc is not None else "-", dose_txt, vol_txt, f"{vol_left:.2f}" if vol_left is not None else "-"))
            Button(add_frame, text="🗑", command=delete_record, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                   width=5, height=1, font=(FONT_NAME, 10, "bold")).grid(row=0, column=9, padx=6)
            #----
            if is_expired:
                Label(tab, text="⚠ VIAL EXPIRED – NO ADMINISTRATION ALLOWED", fg="red", bg=C4, font=(FONT_NAME, 11, "bold")).pack(pady=(5, 0))
                disable_buttons(tab, exempt_texts=["Back"])
            if stored_date:
                Label(tab, text=f"⚠ VIAL STORED ({stored_date})\nNO FURTHER ACTIONS ALLOWED", fg="red", bg=C4, font=(FONT_NAME, 11, "bold"), justify="center").pack(pady=(5, 0))
                disable_buttons(tab, exempt_texts=["Back"])
            # Buttons
            btn_frame = Frame(tab, bg=C4)
            btn_frame.pack(pady=(20,0))
            Button(btn_frame, text="Back", **TAB_BUTTON_STYLE, command=lambda nt=tab: self.back_to_main(nt)).pack()
        #RUN
        select_vial_file()

    # ==========================================================================================TC99M================================================================================

    def _tab_tc99m(self, tab):
        # Choose New or Old File
        def select_file():
            popup_window = Toplevel(self.window)
            popup_window.title("Choose File")
            popup_window.config(bg=C4)
            center_window(window=popup_window, w=350, h=250)
            Label(popup_window, text=f"Select Tc99m Generator File:", **TEXT_COLORS, font=(FONT_NAME, 16, "bold")).pack(pady=20)
            def create_new():
                popup_window.destroy()
                new_generator_file()
            def open_existing():
                popup_window.destroy()
                existing_generator_file()
            Button(popup_window, text="New File", **TAB_BUTTON_STYLE, command=create_new).pack(pady=10)
            Button(popup_window, text="Old File", **TAB_BUTTON_STYLE, command=open_existing).pack()
        # Create New File
        def new_generator_file():
            popup_window = Toplevel(self.window)
            popup_window.title("New Tc99m Generator")
            popup_window.config(bg=C4)
            center_window(window=popup_window, w=420, h=320)
            Label(popup_window, text="New Generator Info", **TEXT_COLORS, font=(FONT_NAME, 16, "bold")).pack(pady=10)
            info_frame = Frame(popup_window, bg=C4)
            info_frame.pack(pady=10)
            Label(info_frame, text="Generator ID:", **TEXT_COLORS).grid(row=0, column=0, sticky="e", padx=5, pady=5)
            gen_id_entry = Entry(info_frame, width=18)
            gen_id_entry.insert(0, "-")
            gen_id_entry.grid(row=0, column=1, pady=5)
            Label(info_frame, text="Calibration Date:", **TEXT_COLORS).grid(row=1, column=0, sticky="e", padx=5, pady=5)
            cal_date_entry = DateEntry(info_frame, width=16, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
            cal_date_entry.grid(row=1, column=1, pady=5)
            Label(info_frame, text="Calibration Time:", **TEXT_COLORS).grid(row=2, column=0, sticky="e", padx=5, pady=5)
            time_field = Frame(info_frame, bg="white", highlightbackground="black", highlightthickness=0)
            time_field.grid(row=2, column=1, padx=5)
            cal_time_entry = Entry(time_field, width=13, bd=0, font=(FONT_NAME, 10))
            cal_time_entry.pack(side="left", padx=(3, 0), pady=2)
            update_time(cal_time_entry)
            refresh_time_button = Button(time_field, text="↻", command=lambda nt=cal_time_entry: update_time(nt), bg="white", fg="black", bd=0,
                                         padx=3, pady=0, font=(FONT_NAME, 10), cursor="hand2")
            refresh_time_button.pack(side="right", padx=3)
            Label(info_frame, text="Mo99 Activity (mCi):", **TEXT_COLORS).grid(row=3, column=0, sticky="e", padx=5, pady=5)
            activity_entry = Entry(info_frame, width=18)
            activity_entry.grid(row=3, column=1, pady=5)
            Label(info_frame, text="Start Date:", **TEXT_COLORS).grid(row=4, column=0, sticky="e", padx=5, pady=5)
            start_date_entry = DateEntry(info_frame, width=16, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
            start_date_entry.grid(row=4, column=1, pady=5)
            Label(info_frame, text="Expiration Date:", **TEXT_COLORS).grid(row=5, column=0, sticky="e", padx=5, pady=5)
            expiration_date_entry = DateEntry(info_frame, width=16, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
            expiration_date_entry.grid(row=5, column=1, pady=5)
            def save_new_file():
                fields = {"Generator ID": gen_id_entry,
                          "Calibration Date": cal_date_entry,
                          "Calibration Time": cal_time_entry,
                          "Mo Activity (mCi)": activity_entry,
                          "Start Date": start_date_entry,
                          "Expiration Date": expiration_date_entry}
                values = {}
                for name, entry in fields.items():
                    val = entry.get().strip()
                    if not val:
                        messagebox.showerror("Error", f"Enter {name}")
                        return
                    values[name] = val
                try:
                    values["Mo Activity (mCi)"] = float(values["Mo Activity (mCi)"])
                except ValueError:
                    messagebox.showerror("Error", "Enter valid Activity (mCi)")
                    return
                gen_id = values["Generator ID"]
                activity = values["Mo Activity (mCi)"]
                start_date = start_date_entry.get()
                cal_date = cal_date_entry.get()
                cal_time = cal_time_entry.get()
                expiration_date = expiration_date_entry.get()
                base_dir = "Tc99m_Generators"
                dt = datetime.strptime(start_date, "%d-%m-%Y")
                year = dt.strftime("%Y")
                month = dt.strftime("%m")
                year_dir = os.path.join(base_dir, year)
                month_dir = os.path.join(year_dir, month)
                folder_name = f"Tc99m_Generator__{start_date}"
                gen_dir = os.path.join(month_dir, folder_name)
                os.makedirs(gen_dir, exist_ok=True)
                db_path = os.path.join(gen_dir, f"{folder_name}.sqlite")
                conn = sqlite3.connect(db_path)
                cur = conn.cursor()
                cur.execute("""CREATE TABLE IF NOT EXISTS generator_info(id TEXT PRIMARY KEY, cal_date TEXT, cal_time TEXT, mo_activity REAL, start_date TEXT, expiration_date TEXT, disposal_date TEXT)""")
                cur.execute("""CREATE TABLE IF NOT EXISTS elutions(id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, time TEXT, tc_activity REAL, expected_activity REAL, 
                                                                   div REAL, volume REAL, concentration REAL, mo_activity REAL)""")
                cur.execute("""CREATE TABLE IF NOT EXISTS kits(id TEXT PRIMARY KEY, parent_id TEXT, date TEXT, time TEXT, kit TEXT, volume REAL, activity REAL,
                                                                concentration REAL, dose REAL, dose_volume REAL, volume_left REAL, patient_name TEXT)""")
                cur.execute("INSERT INTO generator_info VALUES (?,?,?,?,?,?,?)",
                            (gen_id, cal_date, cal_time, activity, start_date_entry.get(), expiration_date_entry.get(), None))
                conn.commit()
                excel_path = os.path.join(gen_dir, f"{folder_name}.xlsx")
                create_excel_for_tc99m(excel_path)
                append_row_to_sheet(excel_path, "Gen Info",[gen_id, cal_date, cal_time, activity, start_date, expiration_date, ""])
                conn.close()
                popup_window.destroy()
                load_generator(db_path)
            bttn_frame = Frame(popup_window, bg=C4)
            bttn_frame.pack()
            Button(bttn_frame, text="Save File", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=8,
                   height=1, font=(FONT_NAME, 10, "bold"), command=save_new_file).grid(row=0, column=0, padx=10, pady=10)
            Button(bttn_frame, text="Back", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=8, height=1,
                   font=(FONT_NAME, 10, "bold"), command=lambda: (popup_window.destroy(), self.tabs_frame.forget(tab), self.create_new_tab("Generators"))).grid(row=0, column=1, padx=10, pady=10)
        # Open Existing File
        def existing_generator_file():
            popup_window = Toplevel(self.window)
            popup_window.title("Open Existing Tc99m Generator")
            popup_window.config(bg=C4)
            center_window(window=popup_window, w=360, h=130)
            Label(popup_window, text="Select Existing Generator Folder", **TEXT_COLORS, font=(FONT_NAME, 17, "bold")).pack(pady=10)
            def open_folder():
                ga68_root = Path(__file__).resolve().parent / "Tc99m_Generators"
                initial_dir = max((p for p in ga68_root.rglob("*") if p.is_dir()), key=lambda p: p.stat().st_mtime).parent
                folder = filedialog.askdirectory(title="Select Tc99m Generator Folder", initialdir=initial_dir)
                if not folder:
                    return
                sqlite_files = [f for f in os.listdir(folder) if f.endswith(".sqlite")]
                if not sqlite_files:
                    messagebox.showerror("Error", "No .sqlite file found in selected folder.")
                    return
                sqlite_path = os.path.join(folder, sqlite_files[0])
                excel_files = [f for f in os.listdir(folder) if f.endswith(".xlsx")]
                if excel_files:
                    excel_path = os.path.join(folder, excel_files[0])
                else:
                    excel_path = os.path.join(folder, os.path.basename(folder) + ".xlsx")
                    create_excel_for_tc99m(excel_path)
                popup_window.destroy()
                load_generator(sqlite_path)
            button_frame = Frame(popup_window, bg=C4)
            button_frame.pack()
            Button(button_frame, text="Open File 🗁", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                   width=12, height=2, font=(FONT_NAME, 12, "bold"), command=open_folder).grid(row=0, column=0, padx=10, pady=10)
            Button(button_frame, text="Back", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=12, height=2,
                   font=(FONT_NAME, 12, "bold"), command=lambda: (popup_window.destroy(), self.tabs_frame.forget(tab), self.create_new_tab("Generators"))).grid(row=0, column=1, padx=10, pady=10)
        # Load Selected Generator
        def load_generator(dbfile):
            for widget in tab.winfo_children():
                widget.destroy()
            conn = sqlite3.connect(dbfile)
            cur = conn.cursor()
            gen_id, cal_date, cal_time, activity, start_date, expiration_date, disposal_date = cur.execute("SELECT * FROM generator_info").fetchone()
            is_disposed = disposal_date is not None
            today = datetime.now().date()
            exp_date = datetime.strptime(expiration_date, "%d-%m-%Y").date()
            is_expired = today > exp_date
            header = Label(tab, text="Daily Tc99m Generator Elution Log Sheet", fg="white", bg=C4, font=(FONT_NAME, 18, "bold"))
            header.pack(pady=(5, 0), fill="x")
            #Scrollable Canvas/Frame
            contents, canvas, scroll_frame, scrollbar = create_scrollable_frame(tab)
            #Selected Generator Info
            info_frame = Frame(scroll_frame, bg=C4)
            info_frame.pack(anchor="center", pady=20)
            Label(info_frame, text=f"Generator ID: {gen_id}", **TEXT_COLORS, font=(FONT_NAME, 10)).grid(row=0, column=0, padx=6, pady=6)
            Label(info_frame, text=f"Calibration on: {cal_date} {cal_time}", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=1, column=0, padx=6, pady=6)
            Label(info_frame, text=f"Mo Activity (mCi): {activity}", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=2, column=0, padx=6, pady=6)
            Label(info_frame, text=f"Start Date: {start_date}", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=3, column=0, padx=6, pady=6)
            Label(info_frame, text=f"T1/2 Mo99 (HR): {T12_MO99}", **TEXT_COLORS, font=(FONT_NAME, 10)).grid(row=0, column=1, padx=6, pady=6)
            Label(info_frame, text=f"T1/2 Tc99m (HR): {T12_TC99M}", **TEXT_COLORS, font=(FONT_NAME, 10)).grid(row=1, column=1, padx=6, pady=6)
            Label(info_frame, text=f"Expiration Date: {expiration_date}", **TEXT_COLORS, font=(FONT_NAME, 10)).grid(row=2, column=1, padx=6, pady=6)
            dispose_button = Button(info_frame, text="✗Dispose Gen✗", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                                    width=14, height=1, font=(FONT_NAME, 10, "bold"), command=lambda: dispose_gen(conn=conn, dbfile=dbfile, on_disposed_callback=update_header_and_disable(header=header, tab=tab)))
            dispose_button.grid(row=3, column=1, padx=6, pady=6)
            if is_disposed or is_expired:
                update_header_and_disable(header=header, tab=tab, is_disposed=is_disposed, is_expired=is_expired)
            #Elutions Table
            columns = [("date", "Date", 120), ("time", "Time", 100), ("activity", "Activity(mCi)", 120), ("expected_activity", "Expected(mCi)", 120),
                       ("div", "Div(%)", 100), ("volume", "Vol(ml)", 120), ("concentration", "Conc(mCi/ml)", 120)]
            tree = ttk.Treeview(scroll_frame, columns=[c[0] for c in columns], show="headings")
            tree.pack(pady=15)
            for col_id, col_title, col_width in columns:
                tree.heading(col_id, text=col_title)
                tree.column(col_id, width=col_width, anchor="center")
            style = ttk.Style()
            style.theme_use("default")
            style.configure("Treeview", background=C2, fieldbackground=C2, foreground="black", rowheight=26, borderwidth=1, bordercolor="black", relief="solid")
            style.configure("Treeview.Heading", background=C3, foreground="white", font=(FONT_NAME, 11, "bold"), relief="solid")
            style.map("Treeview", background=[("selected", "#8FAADC"), ("!selected", C2)], foreground=[("selected", "black")])
            style.layout("Treeview", [("Treeview.treearea", {"sticky": "nsew"})])
            # Load Data
            rows = cur.execute("SELECT id, date, time, tc_activity, expected_activity, div, volume, concentration FROM elutions").fetchall()
            for r in rows:
                row_id = r[0]
                data = r[1:]
                tree.insert("", "end", iid=row_id, values=data)
            # Add New Elution
            add_frame = Frame(scroll_frame, bg=C4)
            add_frame.pack(pady=10)
            Label(add_frame, text="Date:", **TEXT_COLORS).grid(row=0, column=0)
            elution_date_entry = DateEntry(add_frame, width=10, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
            elution_date_entry.grid(row=0, column=1, padx=5)
            time_field = Frame(add_frame, bg="white", highlightbackground="black", highlightthickness=0)
            time_field.grid(row=0, column=3, padx=5)
            Label(add_frame, text="Time of Elution:", **TEXT_COLORS).grid(row=0, column=2, sticky="e", padx=5)
            elution_time_entry = Entry(time_field, width=7, bd=0, font=(FONT_NAME, 10))
            elution_time_entry.pack(side="left", padx=(3, 0), pady=2)
            update_time(elution_time_entry)
            refresh_time_button = Button(time_field, text="↻", command=lambda nt=elution_time_entry: update_time(nt),bg="white",
                                         fg="black", bd=0, padx=3, pady=0, font=(FONT_NAME, 10), cursor="hand2")
            refresh_time_button.pack(side="right", padx=3)
            Label(add_frame, text="Activity (mCi):", **TEXT_COLORS).grid(row=0, column=4, sticky="e", padx=5)
            elution_activity_entry = Entry(add_frame, width=10)
            elution_activity_entry.grid(row=0, column=5, padx=5)
            Label(add_frame, text="Volume (ml):", **TEXT_COLORS).grid(row=0, column=6, sticky="e", padx=5)
            elution_volume_entry = Entry(add_frame, width=8)
            elution_volume_entry.grid(row=0, column=7, padx=5)
            def add_record():
                def mo_activity(A0, dt_hours):
                    return A0 * math.exp(-LAMBDA_MO * dt_hours)
                # "Tc99_lab" helper (χωρίς yield/efficiency)
                def tc99_lab_from_mo(A_mo_now, dt_hours):
                    dlam = (LAMBDA_TC - LAMBDA_MO)
                    if abs(dlam) < 1e-12:
                        return A_mo_now * LAMBDA_TC * dt_hours
                    return A_mo_now * (LAMBDA_TC / dlam) * (1 - math.exp(-dlam * dt_hours))
                try:
                    a = round(float(elution_activity_entry.get()), 2)
                    v = round(float(elution_volume_entry.get()), 1)
                    conc = round(a / v, 2) if v > 0 else 0.0
                    d = elution_date_entry.get()
                    t = elution_time_entry.get()
                    elution_dt = datetime.strptime(f"{d} {t}", "%d-%m-%Y %H:%M")
                    elution_iso = elution_dt.strftime("%Y-%m-%d %H:%M")
                    # --- calibration ---
                    cur.execute(
                        "SELECT cal_date, cal_time, mo_activity FROM generator_info ORDER BY rowid DESC LIMIT 1")
                    cal_date_str, cal_time_str, cal_mo_activity = cur.fetchone()
                    cal_dt = datetime.strptime(f"{cal_date_str} {cal_time_str}", "%d-%m-%Y %H:%M")
                    # --- last elution (DD-MM-YYYY safe order) ---
                    iso_expr = ("(substr(date,7,4) || '-' || substr(date,4,2) || '-' || substr(date,1,2) "
                                "|| ' ' || time)")
                    cur.execute(f"SELECT date, time FROM elutions WHERE {iso_expr} <= ? ORDER BY {iso_expr} DESC LIMIT 1",(elution_iso,))
                    last_elution = cur.fetchone()
                    # --- Δt build since last elution (or calibration) ---
                    if last_elution:
                        last_date, last_time = last_elution
                        ref_dt = datetime.strptime(f"{last_date} {last_time}", "%d-%m-%Y %H:%M")
                        dt_build = (elution_dt - ref_dt).total_seconds() / 3600.0
                    else:
                        ref_dt = cal_dt
                        dt_build = (elution_dt - cal_dt).total_seconds() / 3600.0
                    if dt_build < 0:
                        dt_build = abs(dt_build)
                    # --- Mo now from calibration ---
                    hours_from_cal = (elution_dt - cal_dt).total_seconds() / 3600.0
                    mo_now = mo_activity(cal_mo_activity, hours_from_cal)
                    # --- helper "Tc99_lab" ---
                    tc99_lab = tc99_lab_from_mo(mo_now, dt_build)
                    # --- βοηθητικοί παράγοντες ---
                    K_RED = 0.8742  # αυτό θα το πειράζεις για να φέρνεις σταθερή απόκλιση
                    CAL_FACTOR = 0.88
                    tc_expected = round(tc99_lab * K_RED * CAL_FACTOR, 2)
                    div = round(((a - tc_expected) / tc_expected) * 100, 1) if tc_expected > 0 else 0
                    cur.execute("INSERT INTO elutions (date, time, tc_activity, expected_activity, div, volume, concentration, mo_activity) VALUES (?,?,?,?,?,?,?,?)",
                                (d, t, a, tc_expected, div, v, conc, round(mo_now, 2)))
                    conn.commit()
                    row_id = cur.lastrowid
                    tree.insert("", "end", iid=row_id, values=(d, t, f"{a:.2f}", f"{tc_expected:.2f}", f"{div:.1f}", f"{v:.1f}", f"{conc:.2f}"))
                    folder = os.path.dirname(dbfile)
                    excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                    append_row_to_sheet(excel_path, "Elutions", [row_id, d, t, a, tc_expected, div, v, conc])
                    elution_activity_entry.delete(0, END)
                    elution_volume_entry.delete(0, END)
                except Exception as e:
                    messagebox.showerror("Error", str(e))
            Button(add_frame, text="Add", command=add_record, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                   width=5, height=1, font=(FONT_NAME, 10, "bold")).grid(row=0, column=8, padx=6)
            #Delete Record
            def delete_record():
                selected = tree.selection()
                if not selected:
                    messagebox.showerror("Error", "No row selected.")
                    return
                row_id = int(selected[0])
                if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected record?"):
                    return
                cur.execute("DELETE FROM Elutions WHERE id=?", (row_id,))
                conn.commit()
                folder = os.path.dirname(dbfile)
                excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                wb = load_workbook(excel_path)
                ws = wb["Elutions"]
                for row in ws.iter_rows(min_row=2):
                    if row[0].value == row_id:
                        ws.delete_rows(row[0].row)
                        break
                wb.save(excel_path)
                tree.delete(selected)
            Button(add_frame, text="🗑", command=delete_record, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                   width=5, height=1, font=(FONT_NAME, 10, "bold")).grid(row=0, column=9, padx=6)
            # Kits
            Label(scroll_frame, text="Select Radiopharmaceutical Kit:", **TEXT_COLORS, font=(FONT_NAME, 14, "bold")).pack(pady=(40,20))
            kits_frame = Frame(scroll_frame, bg=C4)
            kits_frame.pack(pady=2)
            def open_kit_popup(kit_name, conn):
                cfg = KIT_CONFIG.get(kit_name)
                if cfg is None:
                    messagebox.showerror("Config Error", f"No configuration found for {kit_name}.")
                    return
                cur = conn.cursor()
                today = datetime.now().strftime("%d-%m-%Y")
                cur.execute("""SELECT date, time, tc_activity, volume, concentration FROM elutions WHERE date=? ORDER BY rowid DESC""", (today,))
                elutions = cur.fetchall()
                elution_map = {}
                elution_options = []
                for el in elutions:
                    time_str = el[1]
                    elution_options.append(time_str)
                    elution_map[time_str] = el
                if elutions is None:
                    messagebox.showerror("Error", "No elution found for this date.")
                    return
                popup = Toplevel()
                popup.title(f"{kit_name}")
                card = Frame(popup, bg=C3, bd=2, relief="solid")
                card.pack(expand=True, fill="both", padx=20, pady=20)
                container = Frame(card, bg=C3)
                container.pack(anchor="center", pady=10)
                row = 0
                Label(container, text=cfg["title"], **TEXT_COLORS_KITS, font=(FONT_NAME,16,"bold")).grid(row=row, columnspan=2, column=0, pady=(10,5))
                row += 1
                Label(container, text=f"Radiopharmaceutical Kit: {kit_name}", **TEXT_COLORS_KITS, font=(FONT_NAME,14,"italic")).grid(row=row,columnspan=2, column=0, pady=(0,10))
                container.grid_columnconfigure(0, weight=1)
                container.grid_columnconfigure(1, weight=1)
                row += 1
                Label(container, text="Select Elution:", **TEXT_COLORS_KITS, font=(FONT_NAME, 13, "bold underline")).grid(row=row, column=0, sticky="e", padx=10, pady=(15,5))
                elution_options = [f"{el[1]}" for el in elutions]
                selected_elution = tkinter.StringVar()
                selected_elution.set(elution_options[0])
                dropdown = OptionMenu(container, selected_elution, *elution_options)
                dropdown.config(bg="white", fg="black", width=4, height=1, highlightthickness=0)
                dropdown.grid(row=row, column=1, sticky="w", padx=(5,20), pady=(15,5))
                row+=1
                Label(container, text="Segmentation Time:", **TEXT_COLORS_KITS, font=(FONT_NAME, 12, "bold")).grid(row=row, column=0, padx=8, pady=10, sticky="e")
                t_field = Frame(container, bg="white", highlightbackground="black", highlightthickness=0)
                t_field.grid(row=row, column=1, padx=5, sticky="w")
                time_entry = Entry(t_field, width=5, bd=0, font=(FONT_NAME, 8))
                time_entry.pack(side="left", padx=(2, 0), pady=2)
                update_time(time_entry)
                refresh_button = Button(t_field, text="↻", command=lambda nt=time_entry: update_time(nt), bg="white",
                                        fg="black", bd=0, padx=3, pady=0, font=(FONT_NAME, 8), cursor="hand2")
                refresh_button.pack(side="right", padx=3)
                row += 1
                Label(container, text="Required Activity (mCi):", **TEXT_COLORS_KITS, font=(FONT_NAME, 12, "bold")).grid(row=row, column=0, padx=5, pady=10, sticky="e")
                req_activity = Entry(container, width=9, font=(FONT_NAME, 8))
                req_activity.grid(row=row, column=1, sticky="w", padx=5)
                req_activity.insert(0, str(cfg["default_activity"]))
                row += 1
                Label(container, text="Required Volume:", **TEXT_COLORS_KITS, font=(FONT_NAME, 12, "bold")).grid(row=row, column=0, padx=10, pady=10, sticky="e")
                result_frame = Frame(container, bg=BG, highlightthickness=0)
                result_frame.grid(row=row, column=1, padx=10, pady=10, sticky="w")
                required_volume_lbl = Label(result_frame, text="-- ml", bg=BG, fg="white", font=(FONT_NAME, 12, "bold"))
                required_volume_lbl.pack()
                row += 1
                Label(container, text="Preparation Steps:", **TEXT_COLORS_KITS, font=(FONT_NAME,12,"bold")).grid(row=row, columnspan=2, column=0, pady=(10,0), padx=6)
                row += 1
                skip_keys = {"title", "default_activity", "final_volume"}
                for key, value in cfg.items():
                    if key in skip_keys:
                        continue
                    pretty_key = key.replace("_", " ").capitalize()
                    pretty_value = value.capitalize()
                    Label(container, text=f"{pretty_key} {pretty_value}", **TEXT_COLORS_KITS, font=(FONT_NAME,12), anchor="center").grid(row=row, columnspan=2, column=0, padx=6, pady=6)
                    row += 1
                Label(container, text=f"Final Volume: {cfg["final_volume"]} ml", **TEXT_COLORS_KITS, font=(FONT_NAME,12,"bold"), anchor="center").grid(row=row, columnspan=2, column=0, padx=10, pady=(15,5))
                row +=1
                def calculate_volume(*args):
                    try:
                        required_activity = float(req_activity.get())
                        selected_time = selected_elution.get()
                        el_date, el_time, el_act, el_vol, el_conc = elution_map[selected_time]
                        elution_dt = datetime.strptime(el_date + " " + el_time, "%d-%m-%Y %H:%M")
                        labeling_dt = datetime.strptime(datetime.now().strftime("%d-%m-%Y") + " " + time_entry.get(),"%d-%m-%Y %H:%M")
                        delta_mins = (labeling_dt - elution_dt).total_seconds() / 60
                        decay_factor = math.exp(-math.log(2) * delta_mins / (T12_TC99M * 60))
                        concentration_now = el_conc * decay_factor
                        volume_needed = required_activity / concentration_now
                        required_volume_lbl.config(text=f"{volume_needed:.2f} ml")
                    except ValueError:
                        required_volume_lbl.config(text="-- ml")
                time_entry.bind("<KeyRelease>", calculate_volume)
                req_activity.bind("<KeyRelease>", calculate_volume)
                selected_elution.trace_add("write", lambda name, index, mode: calculate_volume())
                def save_kit_data():
                    try:
                        time_val = time_entry.get()
                        kit_val = kit_name
                        volume_val = float(required_volume_lbl.cget("text").replace(" ml", "").replace("--", "0"))
                        activity_val = float(req_activity.get())
                        kit_config = KIT_CONFIG.get(kit_val, {})
                        dilution_cfg = kit_config.get("dilution", "0ml")
                        if isinstance(dilution_cfg, str) and "ml" in dilution_cfg:
                            dilution_val = float(dilution_cfg.replace("ml", "").strip())
                        elif isinstance(dilution_cfg, (int, float)):
                            dilution_val = float(dilution_cfg)
                        else:
                            dilution_val = volume_val
                        if volume_val < dilution_val:
                            volume_left_val = dilution_val
                        else:
                            volume_left_val = volume_val
                        concentration_val = round(float(activity_val / volume_left_val),2) if volume_val > 0 else 0
                        date_val = datetime.now().strftime("%d-%m-%Y")
                        cur = conn.cursor()
                        cur.execute("""SELECT MAX(CAST(id AS INTEGER)) FROM kits WHERE parent_id IS NULL""")
                        max_id = cur.fetchone()[0]
                        kit_id = str(int(max_id or 0) + 1)
                        cur.execute("""INSERT INTO kits (id, parent_id, date, time, kit, volume, activity, concentration, dose, dose_volume, volume_left, patient_name) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                                    (kit_id, None, date_val, time_val, kit_val, volume_val, activity_val, concentration_val, None, None, volume_left_val, None))
                        conn.commit()
                        tv_id = kit_tree.insert("", "end", iid=kit_id, values=(time_val, kit_val, f"{volume_val:.2f}", f"{activity_val:.2f}", f"{concentration_val:.2f}", "", "", f"{volume_left_val:.2f}"))
                        kit_tree.kit_ids[tv_id] = kit_id
                        folder = os.path.dirname(dbfile)
                        excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                        append_row_to_sheet(excel_path=excel_path, sheet_name="Kits", row_values=[kit_id, "", date_val, time_val, kit_val, volume_val, activity_val, concentration_val, "", "", volume_left_val, ""])
                        popup.destroy()
                    except Exception as e:
                        messagebox.showerror("Error", str(e))
                b_frame = Frame(card, bg=C3)
                b_frame.pack(anchor="center", pady=10)
                Button(b_frame, text="Calculate", command=calculate_volume, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['bg', 'width', 'height', 'font']},
                       bg=C4, width=10, height=2, font=(FONT_NAME, 10, "bold")).grid(row=row, column=0, pady=10, padx=6)
                Button(b_frame, text="OK", command=save_kit_data, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                       width=10, height=2, font=(FONT_NAME, 10, "bold")).grid(row=row, column=1, pady=10, padx=6)
                Button(b_frame, text="Cancel", command=popup.destroy, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['bg', 'width', 'height', 'font']},
                       bg=C4, width=10, height=2, font=(FONT_NAME, 10, "bold")).grid(row=row, column=2, pady=10, padx=6)
                popup.update_idletasks()
                width = popup.winfo_reqwidth() + 100
                height = popup.winfo_reqheight() + 40
                popup.geometry(f"{width}x{height}")
                popup.resizable(False, False)
                popup.configure(bg=C3)
                center_window(popup, width, height)
            kits = ["MDP", "CERETEC", "MAG-3", "CEA-SCAN", "DTPA", "LEUKOSCAN", "MAASCINT", "BIDA", "DMSA","CARDIOLITE", "MYOVIEW", "NEOSPECT", "PHYTATE", "--", "HIG"]
            for idx, text in enumerate(kits):
                r = idx // 5
                c = idx % 5
                Button(kits_frame, text=text, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                       width=12, height=1, font=(FONT_NAME, 10, "bold"), command=lambda t=text: open_kit_popup(t, conn)).grid(column=c, row=r, padx=10, pady=10)
            #Kit Table
            date_frame = Frame(scroll_frame, bg=C4)
            date_frame.pack(pady=20)
            Label(date_frame, text="Select Date:", font=(FONT_NAME,10,"bold"), **TEXT_COLORS).grid(column=0, row=0, padx=5)
            select_date = DateEntry(date_frame, date_pattern="dd-mm-yyyy", width=12)
            select_date.grid(column=1, row=0, padx=5)
            Button(date_frame, text="Load", command=lambda: load_kits_by_date(select_date.get()), **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=6, height=1).grid(column=2, row=0, padx=10)
            table_frame = Frame(scroll_frame, bg=C4)
            table_frame.pack(pady=10)
            columns = ["time", "kit", "volume", "activity", "concentration", "dose", "dose_volume", "volume_left"]
            kit_tree = ttk.Treeview(table_frame, columns=columns, show="tree headings")
            kit_tree.pack(pady=10)
            kit_tree.kit_ids = {}
            kit_tree.column("#0", width=0, stretch=False)
            kit_tree.heading("#0", text="")
            headers = [("time", "Time", 60), ('kit', "Kit", 60), ("volume", "Vol(ml)", 100), ("activity", "Activity(mCi)", 120), ("concentration", "Conc(mCi/ml)", 110),
                       ("dose", "Dose(mCi)", 110), ("dose_volume", "Dose Vol(ml)", 110), ("volume_left", "Vol Left(ml)", 100)]
            for col, title, width in headers:
                kit_tree.heading(col, text=title)
                kit_tree.column(col, width=width, anchor="center")
            # ACTUAL VALUES FOR PARENT
            def open_actual_parent_popup(parent_id):
                popup = Toplevel(tab)
                popup.title("Insert Actual Values")
                popup.config(bg=C4, pady=15)
                center_window(popup, 250, 160)
                values = list(kit_tree.item(parent_id, "values"))
                planned_activity = values[3].split("→")[0].strip()
                planned_volume = values[2].split("→")[0].strip()
                Label(popup, text="Actual Activity (mCi):", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=0, column=0, padx=5, pady=10)
                actual_activity_entry = Entry(popup, width=10)
                actual_activity_entry.insert(0, f"{planned_activity}")
                actual_activity_entry.grid(row=0, column=1, padx=5, pady=10)
                Label(popup, text="Actual Volume (ml):", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=1, column=0, padx=5, pady=10)
                actual_volume_entry = Entry(popup, width=10)
                actual_volume_entry.insert(0, f"{planned_volume}")
                actual_volume_entry.grid(row=1, column=1, padx=5, pady=10)
                def save():
                    try:
                        actual_activity = float(actual_activity_entry.get())
                        actual_volume = float(actual_volume_entry.get())
                    except ValueError:
                        messagebox.showerror("Error", "Invalid values.")
                        return
                    cfg = KIT_CONFIG.get(values[1], {})
                    dilution_cfg = cfg.get("dilution", "0ml")
                    if isinstance(dilution_cfg, str):
                        if "ml" in dilution_cfg:
                            dilution_val = float(dilution_cfg.replace("ml", "").strip())
                        else:
                            try:
                                dilution_val = float(dilution_cfg)
                            except ValueError:
                                dilution_val = float(values[2].split("→")[0].strip())
                    elif isinstance(dilution_cfg, (int, float)):
                        dilution_val = float(dilution_cfg)
                    else:
                        dilution_val = float(values[2].split("→")[0].strip())
                    volume_left_parent = max(actual_volume, dilution_val)
                    values[3] = f"{planned_activity} → {actual_activity:.2f}"
                    values[2] = f"{planned_volume} → {actual_volume:.2f}"
                    values[7] = f"{volume_left_parent:.2f}"
                    kit_tree.item(parent_id, values=values)
                    cur.execute("UPDATE kits SET activity=?, volume=?, volume_left=? WHERE id=?",
                                (actual_activity, actual_volume, volume_left_parent, parent_id))
                    conn.commit()
                    children = kit_tree.get_children(parent_id)
                    running_vol_left = volume_left_parent
                    for child_id in children:
                        child_vals = list(kit_tree.item(child_id, "values"))
                        dose_volume = float(child_vals[6].split("→")[-1].strip())
                        running_vol_left = round(running_vol_left - dose_volume, 2)
                        child_vals[7] = f"{running_vol_left:.2f}"
                        kit_tree.item(child_id, values=child_vals)
                        cur.execute("UPDATE kits SET volume_left=? WHERE id=?", (running_vol_left, child_id))
                    conn.commit()
                    folder = os.path.dirname(dbfile)
                    excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                    wb = load_workbook(excel_path)
                    ws = wb["Kits"]
                    for row in ws.iter_rows(min_row=2):
                        if str(row[0].value) == str(parent_id):
                            row[5].value = actual_volume
                            row[6].value = actual_activity
                            row[10].value = volume_left_parent
                            break
                    running_vol_left = volume_left_parent
                    for child_id in children:
                        for row in ws.iter_rows(min_row=2):
                            if str(row[0].value) == str(child_id):
                                dose_vol = float(row[9].value)
                                running_vol_left = round(running_vol_left - dose_vol, 2)
                                row[10].value = running_vol_left
                                break
                    wb.save(excel_path)
                    popup.destroy()
                Button(popup, text="OK", command=save, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                       width=10, height=2, font=(FONT_NAME, 10, "bold")).grid(row=2, column=0, pady=10, padx=6)
                Button(popup, text="Cancel", command=popup.destroy, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['bg', 'width', 'height', 'font']},
                       bg=C4, width=10, height=2, font=(FONT_NAME, 10, "bold")).grid(row=2, column=1, pady=10, padx=6)
            # ACTUAL VALUES FOR CHILD
            def open_actual_child_popup(child_id):
                popup = Toplevel(tab)
                popup.title("Insert Actual Values")
                popup.config(bg=C4, pady=15)
                center_window(popup, 250, 160)
                values = list(kit_tree.item(child_id, "values"))
                planned_dose = values[5].split("→")[0].strip()
                planned_dose_vol = values[6].split("→")[0].strip()
                Label(popup, text="Actual Dose (mCi):", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=0, column=0, padx=5, pady=10)
                actual_dose_entry = Entry(popup, width=10)
                actual_dose_entry.insert(0, f"{planned_dose}")
                actual_dose_entry.grid(row=0, column=1, padx=5, pady=10)
                Label(popup, text="Actual Volume (ml):", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=1, column=0, padx=5, pady=10)
                actual_dose_vol_entry = Entry(popup, width=10)
                actual_dose_vol_entry.insert(0, f"{planned_dose_vol}")
                actual_dose_vol_entry.grid(row=1, column=1, padx=5, pady=10)
                def save():
                    try:
                        actual_dose = float(actual_dose_entry.get())
                        actual_dose_vol = float(actual_dose_vol_entry.get())
                    except ValueError:
                        messagebox.showerror("Error", "Invalid values.")
                        return
                    parent_id = kit_tree.parent(child_id)
                    siblings = list(kit_tree.get_children(parent_id))
                    child_index = siblings.index(child_id)
                    values[5] = f"{planned_dose} → {actual_dose:.2f}"
                    values[6] = f"{planned_dose_vol} → {actual_dose_vol:.2f}"
                    if child_index == 0:
                        parent_vals = kit_tree.item(parent_id, "values")
                        running_vol_left = float(parent_vals[7])
                    else:
                        prev_vals = kit_tree.item(siblings[child_index - 1], "values")
                        running_vol_left = float(prev_vals[7])
                    running_vol_left = round(running_vol_left - actual_dose_vol, 2)
                    values[7] = f"{running_vol_left:.2f}"
                    kit_tree.item(child_id, values=values)
                    cur.execute("UPDATE kits SET dose=?, dose_volume=?, volume_left=? WHERE id=?",
                                (actual_dose, actual_dose_vol, running_vol_left, child_id))
                    for next_child in siblings[child_index + 1:]:
                        next_vals = list(kit_tree.item(next_child, "values"))
                        dv = float(next_vals[6].split("→")[-1].strip())
                        running_vol_left = round(running_vol_left - dv, 2)
                        next_vals[7] = f"{running_vol_left:.2f}"
                        kit_tree.item(next_child, values=next_vals)
                        cur.execute("UPDATE kits SET volume_left=? WHERE id=?", (running_vol_left, next_child))
                    conn.commit()
                    folder = os.path.dirname(dbfile)
                    excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                    wb = load_workbook(excel_path)
                    ws = wb["Kits"]
                    excel_rows = {str(ws.cell(row=r, column=1).value): r for r in range(2, ws.max_row + 1)}
                    for cid in siblings[child_index:]:
                        if cid in excel_rows:
                            r = excel_rows[cid]
                            vals = kit_tree.item(cid, "values")
                            ws.cell(row=r, column=9, value=vals[5])
                            ws.cell(row=r, column=10, value=vals[6])
                            ws.cell(row=r, column=11, value=vals[7])
                    wb.save(excel_path)
                    popup.destroy()
                Button(popup, text="Save", command=save, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ["width", "height", "font"]},
                       width=10, height=1).grid(row=2, column=0, padx=10, pady=10)
                Button(popup, text="Cancel", command=popup.destroy, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ["width", "height", "font"]},
                       width=10, height=1).grid(row=2, column=1, padx=10, pady=10)
            #DOUBLE CLICK FOR PATIENT OR ACTUAL VALUES (PARENT)
            def on_tree_double_click(event):
                row_id = kit_tree.identify_row(event.y)
                col_id = kit_tree.identify_column(event.x)
                if not row_id or not col_id:
                    return
                col_index = int(col_id.replace("#", "")) - 1
                is_child = bool(kit_tree.parent(row_id))
                if not is_child:
                    if col_index in (3,2):
                        open_actual_parent_popup(row_id)
                    else:
                        open_patient_popup(row_id)
                else:
                    if col_index in (5,6):
                        open_actual_child_popup(row_id)
            kit_tree.bind("<Double-1>", on_tree_double_click)
            #LOAD OLD KITS DATA
            def load_kits_by_date(selected_date):
                for item in kit_tree.get_children():
                    kit_tree.delete(item)
                kit_tree.kit_ids = {}
                cur = conn.cursor()
                cur.execute("""SELECT id, parent_id, time, kit, volume, activity, concentration, dose, dose_volume, volume_left, patient_name FROM kits WHERE date=? ORDER BY id""", (selected_date,))
                rows = cur.fetchall()
                parents = {}
                children = []
                for row in rows:
                    id, parent_id, time_val, kit_val, volume, activity, conc, dose, dose_vol, vol_left, name = row
                    if parent_id is None:
                        tv_id = kit_tree.insert("", "end", iid=id, values=(time_val, kit_val, f"{volume:.2f}" if volume else "", f"{activity:.2f}" if activity else "",
                                                                                        f"{conc:.2f}" if conc else "", "", "", f"{vol_left:.2f}" if vol_left else ""))
                        kit_tree.kit_ids[tv_id] = id
                        parents[id] = tv_id
                    else:
                        children.append((id, parent_id, time_val, dose, dose_vol, conc, vol_left, name))
                for id, parent_id, time_val, dose, dose_vol, conc, vol_left, name in children:
                    if parent_id in parents:
                        kit_tree.insert(parents[parent_id], "end", iid=id, values=(time_val, "", "", "", f"{conc:.2f}" if conc else "", f"{dose:.2f}" if dose else "", f"{dose_vol:.2f}" if dose_vol else "", f"{vol_left:.2f}" if vol_left else ""))
                        kit_tree.item(parents[parent_id], open=True)
            #ADD NEW PATIENT TO KIT
            def open_patient_popup(kit_row_id):
                popup = Toplevel()
                popup.title("Patient Data")
                popup.configure(bg=C4)
                center_window(popup, 240, 130)
                frame = Frame(popup, bg=C4)
                frame.pack(expand=True, fill="both", anchor="center", padx=10, pady=10)
                Label(frame, text="Patient Name: ", font=(FONT_NAME,10), **TEXT_COLORS).grid(row=0, column=0, pady=5)
                name_entry = Entry(frame, width=18)
                name_entry.insert(0, "-")
                name_entry.grid(row=0, column=1, pady=5)
                Label(frame, text="Dose(mCi):", font=(FONT_NAME,10), **TEXT_COLORS).grid(row=1, column=0, pady=5)
                dose_entry = Entry(frame, width=10)
                dose_entry.grid(row=1, column=1, pady=5)
                def save_patient():
                    name = name_entry.get().strip()
                    try:
                        dose = float(dose_entry.get().strip())
                    except ValueError:
                        messagebox.showerror("Error", "Please enter a valid Dose in mCi.")
                        return
                    parent_id = kit_row_id
                    parent_vals = kit_tree.item(parent_id, "values")
                    kit_val = parent_vals[1]
                    parent_time = parent_vals[0]
                    initial_activity = float(parent_vals[3])
                    initial_conc = float(parent_vals[4])
                    initial_volume = float(parent_vals[7])
                    date_val = datetime.now().strftime("%d-%m-%Y")
                    parent_datetime = datetime.strptime(f"{date_val} {parent_time}", "%d-%m-%Y %H:%M")
                    now_dt = datetime.now()
                    delta_h = (now_dt - parent_datetime).total_seconds() / 3600
                    decay_factor = math.exp(-math.log(2) * delta_h / T12_TC99M)
                    activity_now = round(initial_activity * decay_factor, 2)
                    children = list(kit_tree.get_children(parent_id))
                    given_activity = 0.0
                    for child in children:
                        vals = kit_tree.item(child, "values")
                        if vals[5]:
                            given_activity += float(vals[5])
                    activity_left = round(activity_now - given_activity, 2)
                    if activity_left <= 0:
                        messagebox.showerror("Error", "No activity left in vial.")
                        return
                    if dose > activity_left:
                        messagebox.showerror("Error", f"Not enough activity left.\nAvailable: {activity_left:.2f} mCi")
                        return
                    current_volume_left = (float(kit_tree.item(children[-1], "values")[-1]) if children else initial_volume)
                    current_conc = round(activity_left / current_volume_left, 2)
                    dose_volume = round(dose / current_conc, 2)
                    new_volume_left = round(current_volume_left - dose_volume, 2)
                    if new_volume_left < 0:
                        messagebox.showerror("Error", "Not enough volume left.")
                        return
                    max_seq = 0
                    for child in children:
                        try:
                            seq = int(str(child).split(".")[1])
                            if seq > max_seq:
                                max_seq = seq
                        except IndexError:
                            continue
                    sequence = max_seq + 1
                    patient_id = f"{parent_id}.{sequence}"
                    now_time = datetime.now().strftime("%H:%M")
                    cur = conn.cursor()
                    cur.execute("""INSERT INTO kits (id, parent_id, date, time, kit, volume, activity, concentration, dose, dose_volume, volume_left, patient_name) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                                (patient_id, parent_id, date_val, now_time, kit_val, None, None, current_conc, dose, dose_volume, new_volume_left, name))
                    conn.commit()
                    kit_tree.insert(parent_id, "end", iid=patient_id, values=(now_time, "", "", "", f"{current_conc:.2f}", f"{dose:.2f}", f"{dose_volume:.2f}", f"{new_volume_left:.2f}"))
                    kit_tree.item(parent_id, open=True)
                    folder = os.path.dirname(dbfile)
                    excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                    wb = load_workbook(excel_path)
                    ws = wb["Kits"]
                    insert_row = find_patient_insert_row(ws, parent_id)
                    ws.insert_rows(insert_row)
                    row_values = [patient_id, parent_id, "", now_time, "", "", "", f"{current_conc:.2f}", f"{dose:.2f}", f"{dose_volume:.2f}", f"{new_volume_left:.2f}", name]
                    for col, value in enumerate(row_values, start=1):
                        ws.cell(row=insert_row, column=col, value=value)
                    wb.save(excel_path)
                    popup.destroy()
                Button(frame, text="Save", command=save_patient, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                       width=12, height=1).grid(row=2, column=0, pady=10, padx=5)
                Button(frame, text="Cancel", command=popup.destroy, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                       width=12, height=1).grid(row=2, column=1, pady=10, padx=5)
            def delete_selected_kit_or_patient():
                selected = kit_tree.selection()
                if not selected:
                    messagebox.showwarning("Failed", "Please select a row to delete.")
                    return
                item_id = selected[0]
                folder = os.path.dirname(dbfile)
                excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                wb = load_workbook(excel_path)
                ws = wb["Kits"]
                cur = conn.cursor()
                #Delete Child row(Patient)
                if "." in item_id:
                    if not messagebox.askyesno("Confirm", "Are you sure you want to delete the selected patient?"):
                        return
                    cur.execute("""DELETE FROM kits WHERE id=?""", (item_id,))
                    conn.commit()
                    child_id = item_id
                    for row_idx in reversed(range(2, ws.max_row + 1)):
                        cell_val = ws.cell(row=row_idx, column=1).value
                        if cell_val is None:
                            continue
                        cell_val = str(cell_val).strip()
                        if str(cell_val).strip() == child_id:
                            ws.delete_rows(row_idx)
                    wb.save(excel_path)
                    parent_id = kit_tree.parent(child_id)
                    kit_tree.delete(child_id)
                    renumber_children(conn, ws, kit_tree, parent_id)
                    update_volume_after_delete(parent_id)
                    return
                #Delete Parent + Children rows(Kit+Patients)
                if not messagebox.askyesno("Confirm", "Are you sure you want to delete the selected Kit AND ALL patients?"):
                    return
                cur.execute("""DELETE FROM kits WHERE id=? OR parent_id=?""", (item_id, item_id))
                conn.commit()
                parent_id = item_id
                for row_idx in reversed(range(2, ws.max_row + 1)):
                    cell_val = ws.cell(row=row_idx, column=1).value
                    if cell_val is None:
                        continue
                    cell_val = str(cell_val).strip()
                    if cell_val == parent_id or cell_val.startswith(f"{parent_id}."):
                        ws.delete_rows(row_idx)
                wb.save(excel_path)
                kit_tree.delete(item_id)
            #UPDATE VOLUME AFTER DELETE
            def update_volume_after_delete(parent_id):
                children = kit_tree.get_children(parent_id)
                if not children:
                    return
                parent_vals = list(kit_tree.item(parent_id, "values"))
                last_vol = float(parent_vals[-1])
                folder = os.path.dirname(dbfile)
                excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                wb = load_workbook(excel_path)
                ws = wb["Kits"]
                excel_rows = {}
                for r in range(2, ws.max_row + 1):
                    cell_kit_id = ws.cell(row=r, column=1).value
                    if cell_kit_id:
                        excel_rows[str(cell_kit_id)] = r
                for child_iid in children:
                    vals = kit_tree.item(child_iid, "values")
                    dose_volume = float(vals[6])
                    new_vol_left = round(last_vol - dose_volume, 2)
                    kit_tree.set(child_iid, column=7, value=f"{new_vol_left:.2f}")
                    cur.execute("UPDATE kits SET volume_left=? WHERE id=?", (new_vol_left, child_iid))
                    conn.commit()
                    row_idx = excel_rows.get(child_iid)
                    if row_idx:
                        ws.cell(row=row_idx, column=11, value=new_vol_left)
                    last_vol = new_vol_left
                wb.save(excel_path)
            Button(date_frame, text="🗑", command=delete_selected_kit_or_patient, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                       width=5, height=1, font=(FONT_NAME, 10, "bold")).grid(row=0, column=3, padx=5)
            # Buttons
            btn_frame = Frame(tab, bg=C4)
            btn_frame.pack(pady=10)
            Button(btn_frame, text="Back", **TAB_BUTTON_STYLE, command=lambda nt=tab: self.back_to_main(nt)).pack()
        #RUN
        select_file()

#============================================================================================GA68=====================================================================================

    def _tab_ga68(self, tab):
        # Choose New or Old File
        def select_file():
            popup_window = Toplevel(self.window)
            popup_window.title("Choose File")
            popup_window.config(bg=C4)
            center_window(window=popup_window, w=350, h=250)
            Label(popup_window, text=f"Select Ga68 Generator File:", **TEXT_COLORS, font=(FONT_NAME, 16, "bold")).pack(pady=20)
            def create_new():
                popup_window.destroy()
                new_generator_file()
            def open_existing():
                popup_window.destroy()
                existing_generator_file()
            Button(popup_window, text="New File", **TAB_BUTTON_STYLE, command=create_new).pack(pady=10)
            Button(popup_window, text="Old File", **TAB_BUTTON_STYLE, command=open_existing).pack()
        # Create New File
        def new_generator_file():
            popup_window = Toplevel(self.window)
            popup_window.title("New Ga68 Generator")
            popup_window.config(bg=C4)
            center_window(window=popup_window, w=420, h=280)
            Label(popup_window, text="New Generator Info", **TEXT_COLORS, font=(FONT_NAME,16,"bold")).pack(pady=10)
            info_frame = Frame(popup_window, bg=C4)
            info_frame.pack(pady=10)
            Label(info_frame, text="Generator Model:", **TEXT_COLORS).grid(row=0, column=0, sticky="e", padx=5, pady=5)
            gen_model_entry = Entry(info_frame, width=18)
            gen_model_entry.grid(row=0, column=1, pady=5)
            gen_model_entry.insert(0, "Galli-Ad")
            Label(info_frame, text="Generator ID:", **TEXT_COLORS).grid(row=1, column=0, sticky="e", padx=5, pady=5)
            gen_id_entry = Entry(info_frame, width=18)
            gen_id_entry.grid(row=1, column=1, pady=5)
            Label(info_frame, text="Activity (MBq):", **TEXT_COLORS).grid(row=2, column=0, sticky="e", padx=5, pady=5)
            activity_entry = Entry(info_frame, width=18)
            activity_entry.grid(row=2, column=1, pady=5)
            Label(info_frame, text="Start Date:", **TEXT_COLORS).grid(row=3, column=0, sticky="e", padx=5, pady=5)
            start_date_entry = DateEntry(info_frame, width=16, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
            start_date_entry.grid(row=3, column=1, pady=5)
            Label(info_frame, text="Expiration Date:", **TEXT_COLORS).grid(row=4, column=0, sticky="e", padx=5, pady=5)
            expiration_date_entry = DateEntry(info_frame, width=16, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
            expiration_date_entry.grid(row=4, column=1, pady=5)
            def save_new_file():
                fields = {"Generator ID": gen_id_entry,
                          "Generator Model": gen_model_entry,
                          "Activity (MBq)": activity_entry,
                          "Start Date": start_date_entry,
                          "Expiration Date": expiration_date_entry}
                values = {}
                for name, entry in fields.items():
                    val = entry.get().strip()
                    if not val:
                        messagebox.showerror("Error", f"Enter {name}")
                        return
                    values[name] = val
                try:
                    values["Activity (MBq)"] = float(values["Activity (MBq)"])
                except ValueError:
                    messagebox.showerror("Error", "Enter valid Activity (MBq)")
                    return
                gen_id = values["Generator ID"]
                gen_model = values["Generator Model"]
                activity = values["Activity (MBq)"]
                start_date = start_date_entry.get()
                expiration_date = expiration_date_entry.get()
                base_dir = "Ga68_Generators"
                dt = datetime.strptime(start_date, "%d-%m-%Y")
                year = dt.strftime("%Y")
                year_dir = os.path.join(base_dir, year)
                folder_name = f"Ga68_Generator__{start_date}"
                gen_dir = os.path.join(year_dir, folder_name)
                os.makedirs(gen_dir, exist_ok=True)
                db_path = os.path.join(gen_dir, f"{folder_name}.sqlite")
                conn = sqlite3.connect(db_path)
                cur = conn.cursor()
                cur.execute(
                    """CREATE TABLE IF NOT EXISTS generator_info(id TEXT PRIMARY KEY, model TEXT, start_date TEXT, activity REAL, expiration_date TEXT, disposal_date TEXT)""")
                cur.execute(
                    """CREATE TABLE IF NOT EXISTS elutions(id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, time TEXT, activity REAL)""")
                cur.execute(
                    """CREATE TABLE IF NOT EXISTS dotatoc (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, patient TEXT,
                     weight REAL, admin_time TEXT, dose REAL, concentration REAL, volume REAL, real_dose REAL, itlc REAL, residual REAL)""")
                cur.execute("INSERT INTO generator_info VALUES (?,?,?,?,?,?)",
                            (gen_id, gen_model, start_date_entry.get(), activity, expiration_date_entry.get(), None))
                conn.commit()
                excel_path = os.path.join(gen_dir, f"{folder_name}.xlsx")
                create_excel_for_ga68(excel_path)
                append_row_to_sheet(excel_path, "Gen Info", [gen_id, gen_model, start_date, activity, expiration_date, ""])
                conn.close()
                popup_window.destroy()
                load_generator(db_path)
            btn_frame = Frame(popup_window, bg=C4)
            btn_frame.pack()
            Button(btn_frame, text="Save File", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=8, height=1,
                    font=(FONT_NAME,10,"bold"), command=save_new_file).grid(row=0, column=0, padx=10, pady=10)
            Button(btn_frame, text="Back", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=8, height=1,
                    font=(FONT_NAME,10,"bold"), command=lambda: (popup_window.destroy(),self.tabs_frame.forget(tab), self.create_new_tab("Generators"))).grid(row=0, column=1, padx=10, pady=10)
        # Open Existing File
        def existing_generator_file():
            popup_window = Toplevel(self.window)
            popup_window.title("Open Existing Ga68 Generator")
            popup_window.config(bg=C4)
            center_window(window=popup_window, w=360, h=130)
            Label(popup_window, text="Select Existing Generator Folder", **TEXT_COLORS, font=(FONT_NAME, 17, "bold")).pack(pady=10)
            def open_folder():
                ga68_root = Path(__file__).resolve().parent / "Ga68_Generators"
                initial_dir = max((p for p in ga68_root.rglob("*") if p.is_dir()), key=lambda p: p.stat().st_mtime).parent
                folder = filedialog.askdirectory(title="Select Ga68 Generator Folder", initialdir=initial_dir)
                if not folder:
                    return
                sqlite_files = [f for f in os.listdir(folder) if f.endswith(".sqlite")]
                if not sqlite_files:
                    messagebox.showerror("Error", "No .sqlite file found in selected folder.")
                    return
                sqlite_path = os.path.join(folder, sqlite_files[0])
                excel_files = [f for f in os.listdir(folder) if f.endswith(".xlsx")]
                if excel_files:
                    excel_path = os.path.join(folder, excel_files[0])
                else:
                    excel_path = os.path.join(folder, os.path.basename(folder) + ".xlsx")
                    create_excel_for_ga68(excel_path)
                popup_window.destroy()
                load_generator(sqlite_path)
            btn_frame = Frame(popup_window, bg=C4)
            btn_frame.pack()
            Button(btn_frame, text="Open File 🗁", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=12,
                   height=2, font=(FONT_NAME, 12, "bold"), command=open_folder).grid(row=0, column=0, padx=10, pady=10)
            Button(btn_frame, text="Back", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=12, height=2,
                   font=(FONT_NAME, 12, "bold"), command=lambda: (popup_window.destroy(), self.tabs_frame.forget(tab), self.create_new_tab("Generators"))).grid(row=0, column=1, padx=10, pady=10)
        # Load Selected Generator
        def load_generator(dbfile):
            for widget in tab.winfo_children():
                widget.destroy()
            conn = sqlite3.connect(dbfile)
            cur = conn.cursor()
            gen_id, gen_model, start_date, activity, expiration_date, disposal_date = cur.execute("SELECT * FROM generator_info").fetchone()
            is_disposed = disposal_date is not None
            today = datetime.now().date()
            exp_date = datetime.strptime(expiration_date, "%d-%m-%Y").date()
            is_expired = today > expiration_date
            header = Label(tab, text="Daily Ga68 Generator Elution Log Sheet", fg="white", bg=C4, font=(FONT_NAME, 18, "bold"))
            header.pack(pady=10)
            #Scrollable Canvas/Frame
            contents, canvas, scroll_frame, scrollbar = create_scrollable_frame(tab)
            # Selected Generator Info Frame
            info_frame = Frame(scroll_frame, bg=C4)
            info_frame.pack(pady=5)
            Label(info_frame, text=f"Generator ID: {gen_id}", **TEXT_COLORS, font=(FONT_NAME, 10)).grid(row=0, column=0, padx=6, pady=6)
            Label(info_frame, text=f"Activity (MBq): {activity}", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=1, column=0, padx=6, pady=6)
            Label(info_frame, text=f"Start Date: {start_date}", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=2, column=0, padx=6, pady=6)
            Label(info_frame, text=f"T1/2 Ge68 (D): {T12_GE68}", **TEXT_COLORS, font=(FONT_NAME, 10)).grid(row=0, column=1, padx=6, pady=6)
            Label(info_frame, text=f"Expiration Date: {expiration_date}", **TEXT_COLORS, font=(FONT_NAME, 10)).grid(row=1, column=1, padx=6, pady=6)
            dispose_button = Button(info_frame, text="✗Dispose Gen✗", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                                    width=14, height=1, font=(FONT_NAME, 10, "bold"), command=lambda: dispose_gen(conn=conn, dbfile=dbfile, on_disposed_callback=update_header_and_disable(header=header, tab=tab)))
            dispose_button.grid(row=2, column=1, padx=6, pady=6)
            if is_disposed or is_expired:
                update_header_and_disable(header=header, tab=tab, is_disposed=is_disposed, is_expired=is_expired)
            # Table
            columns = [("date", "Date", 120), ("time", "Time", 100), ("activity", "Activity(mCi)", 120)]
            tree = ttk.Treeview(scroll_frame, columns=[c[0] for c in columns], show="headings")
            tree.pack(pady=15)
            for col_id, col_title, col_width in columns:
                tree.heading(col_id, text=col_title)
                tree.column(col_id, width=col_width, anchor="center")
            style = ttk.Style()
            style.theme_use("default")
            style.configure("Treeview", background=C2, fieldbackground=C2, foreground="black", rowheight=26, borderwidth=1, bordercolor="black", relief="solid")
            style.configure("Treeview.Heading", background=C3, foreground="white", font=(FONT_NAME,11,"bold"), relief="solid")
            style.map("Treeview", background=[("selected","#8FAADC"),("!selected",C2)], foreground=[("selected","black")])
            style.layout("Treeview",[("Treeview.treearea",{"sticky":"nsew"})])
            # Load Data
            rows = cur.execute("SELECT id,date,time,activity FROM elutions").fetchall()
            for r in rows:
                row_id = r[0]
                data = r[1:]
                tree.insert("", "end", iid=row_id, values=data)
            # Add New Elution
            add_frame = Frame(scroll_frame, bg=C4)
            add_frame.pack(pady=10)
            Label(add_frame, text="Date:", **TEXT_COLORS).grid(row=0, column=0)
            elution_date_entry = DateEntry(add_frame, width=10, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
            elution_date_entry.grid(row=0, column=1, padx=5)
            time_field = Frame(add_frame, bg="white", highlightbackground="black", highlightthickness=0)
            time_field.grid(row=0, column=3, padx=5)
            Label(add_frame, text="Time of Elution:", **TEXT_COLORS).grid(row=0, column=2, sticky="e", padx=5)
            elution_time_entry = Entry(time_field, width=7, bd=0, font=(FONT_NAME, 10))
            elution_time_entry.pack(side="left", padx=(3, 0), pady=2)
            update_time(elution_time_entry)
            refresh_time_button = Button(time_field, text="↻", command=lambda nt=elution_time_entry: update_time(nt),
                                             bg="white", fg="black", bd=0, padx=3, pady=0, font=(FONT_NAME, 10), cursor="hand2")
            refresh_time_button.pack(side="right", padx=3)
            Label(add_frame, text="Activity (mCi):", **TEXT_COLORS).grid(row=0, column=4, sticky="e", padx=5)
            elution_activity_entry = Entry(add_frame, width=10)
            elution_activity_entry.grid(row=0, column=5, padx=5)
            def add_record():
                try:
                    a = round(float(elution_activity_entry.get()), 2)
                except ValueError:
                    messagebox.showerror("Error", "Invalid Activity")
                    return
                d = elution_date_entry.get()
                t = elution_time_entry.get()
                cur.execute("INSERT INTO elutions(date,time,activity) VALUES (?,?,?)",
                            (d, t, a))
                conn.commit()
                row_id = cur.lastrowid
                tree.insert("", "end", iid=row_id, values=(d, t, f"{a:.2f}"))
                folder = os.path.dirname(dbfile)
                excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                append_row_to_sheet(excel_path, "Elutions", [row_id, d, t, a])
                elution_activity_entry.delete(0, END)
            Button(add_frame, text="Add", command=add_record, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                   width=5, height=1, font=(FONT_NAME, 10, "bold")).grid(row=0, column=6, padx=6)
            #Delete Record
            def delete_record():
                selected = tree.selection()
                if not selected:
                    messagebox.showerror("Error", "No row selected.")
                    return
                row_id = int(selected[0])
                if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected record?"):
                    return
                cur.execute("DELETE FROM Elutions WHERE id=?",(row_id,))
                conn.commit()
                folder = os.path.dirname(dbfile)
                excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                wb = load_workbook(excel_path)
                ws = wb["Elutions"]
                for row in ws.iter_rows(min_row=2):
                    if row[0].value == row_id:
                        ws.delete_rows(row[0].row)
                        break
                wb.save(excel_path)
                tree.delete(selected)
            Button(add_frame, text="🗑", command=delete_record, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                   width=5, height=1, font=(FONT_NAME, 10, "bold")).grid(row=0, column=7, padx=6)

            #Ga68-DOTATOC
            dotatoc_frame = Frame(scroll_frame, bg=C4)
            dotatoc_frame.pack(pady=30)
            Label(dotatoc_frame, text="DOTATOC Dose Calculator", **TEXT_COLORS, font=(FONT_NAME, 16, "bold")).grid(row=0, column=0, columnspan=6, pady=(20,20))
            Label(dotatoc_frame, text="Date:", **TEXT_COLORS).grid(row=1, column=0, sticky="e", padx=10, pady=5)
            date_entry = DateEntry(dotatoc_frame, width=10, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
            date_entry.grid(row=1, column=1, padx=10, pady=5)
            Label(dotatoc_frame, text="Patient Name:", **TEXT_COLORS).grid(row=1, column=2, sticky="e", padx=10, pady=5)
            patient_entry = Entry(dotatoc_frame, width=20)
            patient_entry.insert(0, "-")
            patient_entry.grid(row=1, column=3, padx=10, pady=5)
            Label(dotatoc_frame, text="Weight (kg):", **TEXT_COLORS).grid(row=1, column=4, sticky="e", padx=10, pady=5)
            weight_entry = Entry(dotatoc_frame, width=6)
            weight_entry.grid(row=1, column=5, padx=10, pady=5)
            time_field = Frame(dotatoc_frame, bg="white", highlightbackground="black", highlightthickness=0)
            time_field.grid(row=2, column=1, padx=10, pady=5)
            Label(dotatoc_frame, text="Segmentation\nTime:", **TEXT_COLORS).grid(row=2, column=0, sticky="e", padx=10, pady=5)
            time_entry = Entry(time_field, width=6, bd=0, font=(FONT_NAME, 10))
            time_entry.pack(side="left", padx=(3, 0), pady=2)
            update_time(time_entry)
            refresh_time_button = Button(time_field, text="↻", command=lambda nt=time_entry: update_time(nt),
                                         bg="white", fg="black", bd=0, padx=3, pady=0, font=(FONT_NAME, 10), cursor="hand2")
            refresh_time_button.pack(side="right", padx=3)
            Label(dotatoc_frame, text="Administration\nTime:", **TEXT_COLORS).grid(row=2, column=2, sticky="e", padx=10, pady=5)
            admin_time_field = Frame(dotatoc_frame, bg="white", highlightbackground="black", highlightthickness=0)
            admin_time_field.grid(row=2, column=3, padx=10, pady=5)
            admin_time_entry = Entry(admin_time_field, width=6, bd=0, font=(FONT_NAME,10))
            admin_time_entry.pack(side="left", padx=(3,0), pady=2)
            update_time(admin_time_entry)
            refresh_time_button = Button(admin_time_field, text="↻", command=lambda nt=admin_time_entry: update_time(nt),
                                         bg="white", fg="black", bd=0, padx=3, pady=0, font=(FONT_NAME, 10), cursor="hand2")
            refresh_time_button.pack(side="right", padx=3)
            #Table
            date_frame = Frame(scroll_frame, bg=C4)
            date_frame.pack(pady=5)
            Label(date_frame, text="Select Date:", font=(FONT_NAME, 10, "bold"), **TEXT_COLORS).grid(column=0, row=0, padx=5)
            select_date = DateEntry(date_frame, date_pattern="dd-mm-yyyy", width=12)
            select_date.grid(column=1, row=0, padx=5)
            Button(date_frame, text="Load", command=lambda: load_dotatoc_by_date(select_date.get()), **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=6, height=1).grid(column=2, row=0, padx=10)
            Button(date_frame, text="🗑", command=lambda: delete_dotatoc_row(), **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=5, height=1, font=(FONT_NAME, 10, "bold")).grid(row=0, column=3, padx=5)
            dotatoc_tree = ttk.Treeview(scroll_frame, columns=("Date","Patient","Weight (kg)","Admin Time","Dose (mCi)","Conc (mCi/ml)","Vol (ml)"), show="headings")
            dotatoc_tree.pack(pady=10)
            for col in ("Date","Patient","Weight (kg)","Admin Time","Dose (mCi)","Conc (mCi/ml)","Vol (ml)"):
                dotatoc_tree.heading(col, text=col.capitalize())
                dotatoc_tree.column(col, width=110, anchor="center")
            tree_style = ttk.Style()
            tree_style.configure("Treeview", background=C2, fieldbackground=C2, foreground="black", rowheight=26)
            tree_style.configure("Treeview.Heading", background=C3, foreground="white", font=(FONT_NAME,11,"bold"))
            #Load Data
            def load_dotatoc_by_date(selected_date):
                for item in dotatoc_tree.get_children():
                    dotatoc_tree.delete(item)
                rows = cur.execute("""SELECT id, date, patient, weight, admin_time, dose, concentration, volume FROM dotatoc WHERE date=? ORDER BY admin_time""", (selected_date,)).fetchall()
                for r in rows:
                    row_id = r[0]
                    data = r[1:]
                    dotatoc_tree.insert("", "end", iid=row_id, values=data)
            #Calculate & Add Row
            def dotatoc_calc():
                try:
                    date = date_entry.get()
                    patient = patient_entry.get()
                    weight = int(weight_entry.get())
                    time = time_entry.get()
                    admin_time = admin_time_entry.get()
                    dose = round(max(weight * 0.067 + 0.2 + (0.5 if weight > 90 else 0), 5.2), 2)
                    cur.execute("SELECT activity, time FROM elutions WHERE date=?",
                                (date,))
                    row = cur.fetchone()
                    if not row:
                        messagebox.showerror("Error", "No elution found for this date.")
                        return
                    el_activity, el_pet_time = row
                    def decay(tmin):
                        return math.exp(-(math.log(2)/67.84) * tmin)
                    fmt = "%H:%M"
                    dt1 = (datetime.strptime(time, fmt) - datetime.strptime(el_pet_time, fmt)).total_seconds()/60
                    conc = round((el_activity / 5.2) * decay(dt1), 2)
                    dt2 = (datetime.strptime(admin_time, fmt) - datetime.strptime(time, fmt)).total_seconds()/60
                    vol = round(dose / (conc * decay(dt2)))
                    dotatoc_tree.insert("", "end", values=(date, patient, weight, admin_time, f"{dose:.1f}", f"{conc:.2f}", f"{vol:.1f}"))
                except Exception as e:
                    messagebox.showerror("Error", str(e))
                    patient_entry.delete(0, END)
                    weight_entry.delete(0, END)
            Button(dotatoc_frame, text="Calculate", command=dotatoc_calc, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                   width=10, height=1, font=(FONT_NAME,10,"bold")).grid(row=2, column=4, padx=20, pady=5)
            #Delete Row
            def delete_dotatoc_row():
                selected = dotatoc_tree.selection()
                if not selected:
                    messagebox.showwarning("Error", "No row selected.")
                    return
                row_id = selected[0]
                if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected row?"):
                    return
                cur.execute("DELETE FROM dotatoc WHERE ID=? ",(row_id,))
                conn.commit()
                folder = os.path.dirname(dbfile)
                excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                wb = load_workbook(excel_path)
                ws = wb["DOTATOC"]
                for r in ws.iter_rows(min_row=2):
                    if r[0].value == row_id:
                        ws.delete_rows(r[0].row)
                        break
                wb.save(excel_path)
                dotatoc_tree.delete(selected[0])
            Button(dotatoc_frame, text="🗑", command=delete_dotatoc_row, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                   width=5, height=1, font=(FONT_NAME,9,"bold")).grid(row=2, column=5, padx=6)
            # Last Entries & SAVE
            def open_last_entries_popup(event):
                selected = dotatoc_tree.selection()
                if not selected:
                    return
                row_id = selected[0]
                values = dotatoc_tree.item(row_id, "values")
                date, patient, weight, admin_time, dose, conc, vol = values[:7]
                popup = Toplevel()
                popup.title("Last Entries")
                popup.geometry("250x200")
                popup.configure(bg=BG)
                center_window(popup, 250, 200)
                frame = Frame(popup, bg=BG)
                frame.pack(expand=True, fill="both", padx=10, pady=10)
                Label(frame, text="Real Dose (mCi):", bg=BG, fg="white", font=(FONT_NAME,10,"bold")).grid(row=0, column=0, padx=5, pady=10)
                real_dose_entry = Entry(frame, width=10)
                real_dose_entry.grid(row=0, column=1, padx=5, pady=10)
                Label(frame, text="ITLC (<2%):", bg=BG, fg="white", font=(FONT_NAME,10,"bold")).grid(row=1, column=0, padx=5, pady=10)
                itlc_entry = Entry(frame, width=10)
                itlc_entry.grid(row=1, column=1, padx=5, pady=10)
                Label(frame, text="Residual (mCi):", bg=BG, fg="white", font=(FONT_NAME,10,"bold")).grid(row=2, column=0, padx=5, pady=10)
                residual_entry = Entry(frame, width=10)
                residual_entry.grid(row=2, column=1, padx=5, pady=10)
                def save_to_dotatoc():
                    real_dose = real_dose_entry.get().strip()
                    itlc = itlc_entry.get().strip()
                    residual = residual_entry.get().strip()
                    if not patient or not weight or not real_dose or not itlc or not residual:
                        messagebox.showerror("Error", "Please fill all the fields.")
                        return
                    cur.execute("""INSERT INTO dotatoc (date, patient, weight, admin_time, dose, concentration, volume,real_dose, itlc, residual) VALUES (?,?,?,?,?,?,?,?,?,?)""",
                            (date, patient, weight, admin_time, dose, conc, vol, real_dose, itlc, residual))
                    conn.commit()
                    row_id = cur.lastrowid
                    folder = os.path.dirname(dbfile)
                    excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                    append_row_to_sheet(excel_path, "DOTATOC", [date, patient, weight, admin_time, dose, conc, vol, real_dose, itlc, residual])
                    messagebox.showinfo("Saved", "Data saved successfully!")
                    real_dose_entry.delete(0, END)
                    itlc_entry.delete(0, END)
                    residual_entry.delete(0, END)
                    popup.destroy()
                Button(frame, text="Save", command=save_to_dotatoc, **{k: v for k, v in BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                       width=8, height=1, font=(FONT_NAME, 12, "bold")).grid(row=3, column=0, padx=10, pady=10)
                Button(frame, text="Cancel", command=popup.destroy, **{k: v for k, v in BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                       width=8, height=1, font=(FONT_NAME, 12, "bold")).grid(row=3, column=1, padx=10, pady=10)
            dotatoc_tree.bind("<Double-1>", open_last_entries_popup)
            #Buttons
            btn_frame = Frame(tab, bg=C4)
            btn_frame.pack(pady=10)
            Button(btn_frame, text="Back", **TAB_BUTTON_STYLE, command=lambda nt=tab: self.back_to_main(nt)).pack()
        # RUN
        select_file()


#==========================================================================================I131=====================================================================================

    def _tab_i131(self, tab):
        header = Label(tab, text="I-131 Calculator", **TEXT_COLORS, font=(FONT_NAME, 24, "bold"))
        header.pack(pady=10)
        frm = Frame(tab, bg=C4)
        frm.pack(pady=10)
        # Labels + Entries
        Label(frm, text="Calibration Date :", **TEXT_COLORS).grid(row=0, column=0, sticky="e", padx=6, pady=6)
        calib_entry = DateEntry(frm, width=16, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
        calib_entry.grid(row=0, column=1, padx=6, pady=6)
        Label(frm, text="Calibration Date Activity (MBq):", **TEXT_COLORS).grid(row=1, column=0, sticky="e", padx=6, pady=6)
        activity_entry = Entry(frm, width=18)
        activity_entry.grid(row=1, column=1, padx=6, pady=6)
        Label(frm, text="Administration Date :", **TEXT_COLORS).grid(row=2, column=0, sticky="e", padx=6, pady=6)
        admin_entry = DateEntry(frm, width=16, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
        admin_entry.grid(row=2, column=1, padx=6, pady=6)
        Label(frm, text="Patient Name :", **TEXT_COLORS).grid(row=3, column=0, sticky="e", padx=6, pady=6)
        name_entry = Entry(frm, width=18)
        name_entry.insert(0, "-")
        name_entry.grid(row=3, column=1, padx=6, pady=6)
        Label(frm, text="Serial Number:", **TEXT_COLORS).grid(row=4, column=0, sticky="e", padx=6, pady=6)
        serial_number_entry = Entry(frm, width=18)
        serial_number_entry.grid(row=4, column=1, padx=6, pady=6)
        # Result
        result_txt = Text(tab, height=14, width=60, bg=BG, fg="white")
        result_txt.pack(pady=12)
        result_txt.configure(state="disabled")
        # Functions
        def mbq_to_mci(mbq):
            return mbq / 37.0
        def calculate_and_show():
            try:
                patient_name = name_entry.get().strip()
                cal_date = calib_entry.get_date()
                A0_mbq =float(activity_entry.get().strip())
                if not A0_mbq:
                    raise ValueError("Please enter the Initial Activity (MBq)")
                admin_date = admin_entry.get_date()
                days_diff = (admin_date - cal_date).days
                df = math.exp(-math.log(2) * days_diff/T12_I131)
                A_admin_mbq = A0_mbq * df
                A_admin_mci = mbq_to_mci(A_admin_mbq)
                serial_number = serial_number_entry.get().strip()
                #Show Output
                result_txt.configure(state="normal")
                result_txt.delete("1.0", END)
                result_txt.tag_configure("normal", font=("Courier",9,"normal"))
                result_txt.tag_configure("bold", font=("Courier",11,"bold"))
                result_txt.insert(END, f"Patient Name: {patient_name}\n\n", "bold")
                result_txt.insert(END, f"Calibration Date: {cal_date.strftime("%d-%m-%Y")}\n\n", "normal")
                result_txt.insert(END,f"Administration Date: {admin_date.strftime("%d-%m-%Y")} (offset {days_diff} days)\n\n","normal")
                result_txt.insert(END, f"Initial Activity: {A0_mbq:.2f} MBq ({mbq_to_mci(A0_mbq):.2f} mCi)\n\n","normal")
                result_txt.insert(END, f"Decay Factor for {days_diff:.2f} days: {df:.2f}\n\n", "normal")
                result_txt.insert(END, f"Activity at Administration: {A_admin_mbq:.2f} MBq = {A_admin_mci:.2f} mCi\n\n","bold")
                result_txt.insert(END, f"Serial Number: {serial_number}\n\n", "normal")
            except Exception as e:
                messagebox.showerror("Calculation Error", str(e))
        def save_record():
            try:
                patient_name = name_entry.get().strip()
                if not patient_name:
                    messagebox.showerror("Save Error", "Please enter Patient Name.")
                cal_date = calib_entry.get_date()
                A0_mbq = float(activity_entry.get().strip())
                admin_date = admin_entry.get_date()
                serial_number = serial_number_entry.get().strip()
                days_diff = (admin_date - cal_date).days
                df = math.exp(-math.log(2) * days_diff / T12_I131)
                A_admin_mbq = A0_mbq * df
                A_admin_mci = mbq_to_mci(A_admin_mbq)
                base_dir = "I131"
                year_dir = os.path.join(base_dir, str(admin_date.year))
                month_dir = os.path.join(year_dir, admin_date.strftime("%m"))
                os.makedirs(month_dir, exist_ok=True)
                sqlite_path = os.path.join(month_dir, f"I131_{admin_date.strftime("%m")}.sqlite")
                conn = sqlite3.connect(sqlite_path)
                cur = conn.cursor()
                cur.execute("""CREATE TABLE IF NOT EXISTS i131_records (id INTEGER PRIMARY KEY AUTOINCREMENT, patient_name TEXT,
                                                   calibration_date TEXT, initial_mbq REAL, admin_date TEXT, admin_mci REAL, serial_number TEXT)""")
                cur.execute("""INSERT INTO i131_records (patient_name, calibration_date, initial_mbq, admin_date, admin_mci, serial_number) VALUES (?,?,?,?,?,?)""",
                            (patient_name, cal_date.strftime("%d-%m-%Y"), A0_mbq, admin_date.strftime("%d-%m-%Y"), A_admin_mci, serial_number))
                conn.commit()
                conn.close()
                excel_path = os.path.join(month_dir, f"I131_{admin_date.strftime("%m")}.xlsx")
                record = {"Patient": patient_name,
                          "Calibration Date": cal_date.strftime("%d-%m-%Y"),
                          "Initial Activity (MBq)": A0_mbq,
                          "Administration Date": admin_date.strftime("%d-%m-%Y"),
                          "Activity at Admin (mCi)": round(A_admin_mci,3),
                          "Serial Number": serial_number}
                df_new = pandas.DataFrame([record])
                if os.path.exists(excel_path):
                    df_existing = pandas.read_excel(excel_path)
                    df_total = pandas.concat([df_existing, df_new], ignore_index=True)
                else:
                    df_total = df_new
                df_total.to_excel(excel_path, index=False)
                messagebox.showinfo("Saved", "Record saved successfully.")
            except Exception as e:
                messagebox.showerror("Save Error", f"Invalid Input: {e}")
                return
        # Buttons
        btn_frame = Frame(tab, bg=C4)
        btn_frame.pack(pady=10)
        calc_button = Button(btn_frame, text="Calculate", **TAB_BUTTON_STYLE, command=calculate_and_show)
        calc_button.grid(row=0, column=0, padx=10)
        save_button = Button(btn_frame, text="Save", **TAB_BUTTON_STYLE, command=save_record)
        save_button.grid(row=0, column=1, padx=10)
        back_button = Button(btn_frame, text="Back", **TAB_BUTTON_STYLE, command=lambda nt=tab: self.back_to_main(nt))
        back_button.grid(row=0, column=2, padx=10)

