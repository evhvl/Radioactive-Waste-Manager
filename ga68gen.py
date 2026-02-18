import sqlite3
from tkinter import *
from tkinter import filedialog, ttk
from tkcalendar import DateEntry
from functions import *
from datetime import datetime
from pathlib import Path

def build_tab(app, tab):

    # Choose New or Old File
    def select_file():
        popup_window = Toplevel(app.window)
        popup_window.title("Choose File")
        popup_window.config(bg=C4)
        center_window(window=popup_window, w=350, h=250)
        Label(popup_window, text=f"Select Ga68 Generator File:", **TEXT_COLORS, font=(FONT_NAME, 16, "bold")).pack \
            (pady=20)
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
        popup_window = Toplevel(app.window)
        popup_window.title("New Ga68 Generator")
        popup_window.config(bg=C4)
        center_window(window=popup_window, w=420, h=340)
        Label(popup_window, text="New Generator Info", **TEXT_COLORS, font=(FONT_NAME ,16 ,"bold")).pack(pady=10)
        info_frame = Frame(popup_window, bg=C4)
        info_frame.pack(pady=10)
        Label(info_frame, text="Generator Model:", **TEXT_COLORS).grid(row=0, column=0, sticky="e", padx=5, pady=5)
        gen_model_entry = Entry(info_frame, width=18)
        gen_model_entry.grid(row=0, column=1, pady=5)
        gen_model_entry.insert(0, "Galli-Ad")
        Label(info_frame, text="Generator ID:", **TEXT_COLORS).grid(row=1, column=0, sticky="e", padx=5, pady=5)
        gen_id_entry = Entry(info_frame, width=18)
        gen_id_entry.grid(row=1, column=1, pady=5)
        gen_id_entry.insert(0, "-")
        Label(info_frame, text="Calibration Date:", **TEXT_COLORS).grid(row=2, column=0, sticky="e", padx=5, pady=5)
        cal_date_entry = DateEntry(info_frame, width=16, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
        cal_date_entry.grid(row=2, column=1, pady=5)
        Label(info_frame, text="Calibration Time:", **TEXT_COLORS).grid(row=3, column=0, sticky="e", padx=5, pady=5)
        time_field = Frame(info_frame, bg="white", highlightbackground="black", highlightthickness=0)
        time_field.grid(row=3, column=1, padx=5)
        cal_time_entry = Entry(time_field, width=13, bd=0, font=(FONT_NAME, 10))
        cal_time_entry.pack(side="left", padx=(3, 0), pady=2)
        update_time(cal_time_entry)
        refresh_time_button = Button(time_field, text="↻", command=lambda nt=cal_time_entry: update_time(nt), bg="white", fg="black", bd=0,
                                     padx=3, pady=0, font=(FONT_NAME, 10), cursor="hand2")
        refresh_time_button.pack(side="right", padx=3)
        Label(info_frame, text="Activity (MBq):", **TEXT_COLORS).grid(row=4, column=0, sticky="e", padx=5, pady=5)
        activity_entry = Entry(info_frame, width=18)
        activity_entry.grid(row=4, column=1, pady=5)
        Label(info_frame, text="Start Date:", **TEXT_COLORS).grid(row=5, column=0, sticky="e", padx=5, pady=5)
        start_date_entry = DateEntry(info_frame, width=16, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
        start_date_entry.grid(row=5, column=1, pady=5)
        Label(info_frame, text="Expiration Date:", **TEXT_COLORS).grid(row=6, column=0, sticky="e", padx=5, pady=5)
        expiration_date_entry = DateEntry(info_frame, width=16, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
        expiration_date_entry.grid(row=6, column=1, pady=5)

        def save_new_file():
            fields = {"Generator ID": gen_id_entry,
                      "Generator Model": gen_model_entry,
                      "Calibration Date": cal_date_entry,
                      "Calibration Time": cal_time_entry,
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
            cal_date = cal_date_entry.get()
            cal_time = cal_time_entry.get()
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
                """CREATE TABLE IF NOT EXISTS generator_info(id TEXT PRIMARY KEY, model TEXT, start_date TEXT, cal_date TEXT, cal_time TEXT, activity REAL, expiration_date TEXT, disposal_date TEXT)""")
            cur.execute(
                """CREATE TABLE IF NOT EXISTS elutions(id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, time TEXT, activity REAL)""")
            cur.execute(
                """CREATE TABLE IF NOT EXISTS dotatoc (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, patient TEXT,
                 weight REAL, admin_time TEXT, dose REAL, concentration REAL, volume REAL, real_dose REAL, itlc REAL, residual REAL)""")
            cur.execute("INSERT INTO generator_info VALUES (?,?,?,?,?,?,?,?)",
                        (gen_id, gen_model, start_date_entry.get(), cal_date, cal_time, activity, expiration_date_entry.get(), None))
            conn.commit()
            excel_path = os.path.join(gen_dir, f"{folder_name}.xlsx")
            create_excel_for_ga68(excel_path)
            append_row_to_sheet(excel_path, "Gen Info", [gen_id, gen_model, start_date, cal_date, cal_time, activity, expiration_date, ""])
            conn.close()
            popup_window.destroy()
            load_generator(db_path)
        btn_frame = Frame(popup_window, bg=C4)
        btn_frame.pack()
        Button(btn_frame, text="Save File", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=8, height=1,
               font=(FONT_NAME ,10 ,"bold"), command=save_new_file).grid(row=0, column=0, padx=10, pady=10)
        Button(btn_frame, text="Back", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=8, height=1,
               font=(FONT_NAME ,10 ,"bold"), command=lambda: (popup_window.destroy() , app.tabs_frame.forget(tab), app.create_new_tab("Generators"))).grid(row=0, column=1, padx=10, pady=10)

    # Open Existing File
    def existing_generator_file():
        popup_window = Toplevel(app.window)
        popup_window.title("Open Existing Ga68 Generator")
        popup_window.config(bg=C4)
        center_window(window=popup_window, w=360, h=130)
        Label(popup_window, text="Select Existing Generator Folder", **TEXT_COLORS, font=(FONT_NAME, 17, "bold")).pack \
            (pady=10)
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
               font=(FONT_NAME, 12, "bold"), command=lambda: (popup_window.destroy(), app.tabs_frame.forget(tab), app.create_new_tab("Generators"))).grid(row=0, column=1, padx=10, pady=10)

    # Load Selected Generator
    def load_generator(dbfile):
        for widget in tab.winfo_children():
            widget.destroy()
        conn = sqlite3.connect(dbfile)
        cur = conn.cursor()
        gen_id, gen_model, start_date, cal_date, cal_time, activity, expiration_date, disposal_date = cur.execute \
            ("SELECT * FROM generator_info").fetchone()
        is_disposed = disposal_date is not None
        today = datetime.now().date()
        expiration_date = datetime.strptime(expiration_date, "%d-%m-%Y").date()
        is_expired = today > expiration_date
        header = Label(tab, text="Daily Ga68 Generator Elution Log Sheet", fg="white", bg=C4, font=(FONT_NAME, 18, "bold"))
        header.pack(pady=10)
        # Scrollable Canvas/Frame
        contents, canvas, scroll_frame, scrollbar = create_scrollable_frame(tab)
        # Selected Generator Info Frame
        info_frame = Frame(scroll_frame, bg=C4)
        info_frame.pack(pady=5)
        Label(info_frame, text=f"Generator ID: {gen_id}", **TEXT_COLORS, font=(FONT_NAME, 10)).grid(row=0, column=0, padx=6, pady=6)
        Label(info_frame, text=f"Calibration on: {cal_date} {cal_time}", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=1, column=0, padx=6, pady=6)
        Label(info_frame, text=f"Activity (MBq): {activity}", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=2, column=0, padx=6, pady=6)
        Label(info_frame, text=f"Start Date: {start_date}", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=3, column=0, padx=6, pady=6)
        Label(info_frame, text=f"T1/2 Ge68 (D): {T12_GE68}", **TEXT_COLORS, font=(FONT_NAME, 10)).grid(row=0, column=1, padx=6, pady=6)
        Label(info_frame, text=f"T1/2 Ga68 (MIN): {T12_GA68}", **TEXT_COLORS, font=(FONT_NAME, 10)).grid(row=1, column=1, padx=6, pady=6)
        Label(info_frame, text=f"Expiration Date: {expiration_date}", **TEXT_COLORS, font=(FONT_NAME, 10)).grid(row=2, column=1, padx=6, pady=6)
        dispose_button = Button(info_frame, text="✗Dispose Gen✗", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                                width=14, height=1, font=(FONT_NAME, 10, "bold"), command=lambda: dispose_gen(conn=conn, dbfile=dbfile, on_disposed_callback=update_header_and_disable
                                                                (header=header, tab=tab)))
        dispose_button.grid(row=3, column=1, padx=6, pady=6)
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
        style.configure("Treeview.Heading", background=C3, foreground="white", font=(FONT_NAME ,11 ,"bold"), relief="solid")
        style.map("Treeview", background=[("selected" ,"#8FAADC") ,("!selected" ,C2)], foreground=[("selected" ,"black")])
        style.layout("Treeview" ,[("Treeview.treearea" ,{"sticky" :"nsew"})])
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

        #Save New Data
        def add_record():
            try:
                a = round(float(elution_activity_entry.get()), 2)
            except ValueError:
                messagebox.showerror("Error", "Invalid Activity")
                return
            d = elution_date_entry.get()
            t = elution_time_entry.get()
            cur.execute("INSERT INTO elutions(date,time,activity) VALUES (?,?,?)" ,(d, t, a))
            conn.commit()
            dotatoc_frame.after(50, refresh_elution_dropdown)
            row_id = cur.lastrowid
            tree.insert("", "end", iid=row_id, values=(d, t, f"{a:.2f}"))
            folder = os.path.dirname(dbfile)
            excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
            append_row_to_sheet(excel_path, "Elutions", [row_id, d, t, a])
            elution_activity_entry.delete(0, END)
        Button(add_frame, text="Add", command=add_record, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
               width=5, height=1, font=(FONT_NAME, 10, "bold")).grid(row=0, column=6, padx=6)

        # Delete Record
        def delete_record():
            selected = tree.selection()
            if not selected:
                messagebox.showerror("Error", "No row selected.")
                return
            row_id = int(selected[0])
            if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected record?"):
                return
            cur.execute("DELETE FROM Elutions WHERE id=?" ,(row_id,))
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

        # Ga68-DOTATOC
        dotatoc_frame = Frame(scroll_frame, bg=C4)
        dotatoc_frame.pack(pady=30)
        Label(dotatoc_frame, text="DOTATOC Dose Calculator", **TEXT_COLORS, font=(FONT_NAME, 16, "bold")).grid(row=0, column=0, columnspan=6, pady=(20, 20))
        elution_select_frame = Frame(dotatoc_frame, bg=C4)
        elution_select_frame.grid(row=1, column=0, columnspan=6, pady=10)
        Label(elution_select_frame, text="Select Elution:", **TEXT_COLORS, font=(FONT_NAME, 12, "bold underline")).pack \
            (side="left", padx=5, pady=(0 ,10))
        Label(dotatoc_frame, text="Date:", **TEXT_COLORS).grid(row=2, column=0, sticky="e", padx=10, pady=5)
        date_entry = DateEntry(dotatoc_frame, width=10, bg=C3, fg="white", date_pattern="dd-mm-yyyy")
        date_entry.set_date(datetime.now())
        date_entry.grid(row=2, column=1, padx=10, pady=5)
        selected_elution = StringVar(dotatoc_frame)
        def get_elution_times(sel_date: str):
            cur.execute \
                ("SELECT TRIM(time) AS t FROM elutions WHERE date=? GROUP BY TRIM(time) ORDER BY MAX(rowid) DESC", (sel_date,))
            return [r[0] for r in cur.fetchall()]
        def refresh_elution_dropdown():
            sel_date = date_entry.get().strip()
            times = get_elution_times(sel_date)
            menu = dropdown["menu"]
            menu.delete(0, "end")
            if not times:
                selected_elution.set("-")
                menu.add_command(label="-", command=lambda v="-": selected_elution.set(v))
                return
            current = selected_elution.get().strip()
            if current not in times:
                selected_elution.set(times[0])
            for t in times:
                menu.add_command(label=t, command=lambda v=t: selected_elution.set(v))
        selected_elution.set("-")
        dropdown = OptionMenu(elution_select_frame, selected_elution, "-")
        dropdown.config(bg="white", fg="black", width=6, height=1, highlightthickness=0)
        dropdown.pack(side="left", padx=5, pady=(0 ,10))
        refresh_elution_dropdown()
        date_entry.bind("<<DateEntrySelected>>", lambda e: refresh_elution_dropdown())
        Label(dotatoc_frame, text="Patient Name:", **TEXT_COLORS).grid(row=2, column=2, sticky="e", padx=10, pady=5)
        patient_entry = Entry(dotatoc_frame, width=20)
        patient_entry.insert(0, "-")
        patient_entry.grid(row=2, column=3, padx=10, pady=5)
        Label(dotatoc_frame, text="Weight (kg):", **TEXT_COLORS).grid(row=2, column=4, sticky="e", padx=10, pady=5)
        weight_entry = Entry(dotatoc_frame, width=6)
        weight_entry.grid(row=2, column=5, padx=10, pady=5)
        time_field = Frame(dotatoc_frame, bg="white", highlightbackground="black", highlightthickness=0)
        time_field.grid(row=3, column=1, padx=10, pady=5)
        Label(dotatoc_frame, text="Segmentation\nTime:", **TEXT_COLORS).grid(row=3, column=0, sticky="e", padx=10, pady=5)
        time_entry = Entry(time_field, width=6, bd=0, font=(FONT_NAME, 10))
        time_entry.pack(side="left", padx=(3, 0), pady=2)
        update_time(time_entry)
        refresh_time_button = Button(time_field, text="↻", command=lambda nt=time_entry: update_time(nt),
                                     bg="white", fg="black", bd=0, padx=3, pady=0, font=(FONT_NAME, 10), cursor="hand2")
        refresh_time_button.pack(side="right", padx=3)
        Label(dotatoc_frame, text="Administration\nTime:", **TEXT_COLORS).grid(row=3, column=2, sticky="e", padx=10, pady=5)
        admin_time_field = Frame(dotatoc_frame, bg="white", highlightbackground="black", highlightthickness=0)
        admin_time_field.grid(row=3, column=3, padx=10, pady=5)
        admin_time_entry = Entry(admin_time_field, width=6, bd=0, font=(FONT_NAME ,10))
        admin_time_entry.pack(side="left", padx=(3 ,0), pady=2)
        update_time(admin_time_entry)
        refresh_time_button = Button(admin_time_field, text="↻", command=lambda nt=admin_time_entry: update_time(nt),
                                     bg="white", fg="black", bd=0, padx=3, pady=0, font=(FONT_NAME, 10), cursor="hand2")
        refresh_time_button.pack(side="right", padx=3)
        def _on_date_change(_event=None):
            refresh_elution_dropdown()
        date_entry.bind("<<DateEntrySelected>>", _on_date_change)
        # Table
        date_frame = Frame(scroll_frame, bg=C4)
        date_frame.pack(pady=5)
        Label(date_frame, text="Select Date:", font=(FONT_NAME, 10, "bold"), **TEXT_COLORS).grid(column=0, row=0, padx=5)
        select_date = DateEntry(date_frame, date_pattern="dd-mm-yyyy", width=12)
        select_date.grid(column=1, row=0, padx=5)
        Button(date_frame, text="Load", command=lambda: load_dotatoc_by_date(select_date.get()), **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=6, height=1).grid(column=2, row=0, padx=10)
        Button(date_frame, text="🗑", command=lambda: delete_dotatoc_row(), **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=5, height=1, font=(FONT_NAME, 10, "bold")).grid(row=0, column=3, padx=5)
        dotatoc_tree = ttk.Treeview(scroll_frame, columns=("Date" ,"Patient" ,"Weight (kg)" ,"Admin Time" ,"Dose (mCi)"
                                                           ,"Conc (mCi/ml)" ,"Vol (ml)"), show="headings")
        dotatoc_tree.pack(pady=10)
        for col in ("Date" ,"Patient" ,"Weight (kg)" ,"Admin Time" ,"Dose (mCi)" ,"Conc (mCi/ml)" ,"Vol (ml)"):
            dotatoc_tree.heading(col, text=col.capitalize())
            dotatoc_tree.column(col, width=110, anchor="center")
        tree_style = ttk.Style()
        tree_style.configure("Treeview", background=C2, fieldbackground=C2, foreground="black", rowheight=26)
        tree_style.configure("Treeview.Heading", background=C3, foreground="white", font=(FONT_NAME ,11 ,"bold"))

        # Load Old DOTATOC Data
        def load_dotatoc_by_date(selected_date):
            for item in dotatoc_tree.get_children():
                dotatoc_tree.delete(item)
            rows = cur.execute \
                ("""SELECT id, date, patient, weight, admin_time, dose, concentration, volume FROM dotatoc WHERE date=? ORDER BY admin_time""", (selected_date,)).fetchall()
            for r in rows:
                row_id = r[0]
                data = r[1:]
                dotatoc_tree.insert("", "end", iid=row_id, values=data)

        # Calculate & Add Row
        def dotatoc_calc():
            try:
                date_str = date_entry.get().strip()
                el_time_str = selected_elution.get().strip()
                seg_time_str = time_entry.get().strip()
                admin_time_str = admin_time_entry.get().strip()
                if not el_time_str or el_time_str == "-":
                    messagebox.showerror("Error", "Please select an elution from the dropdown.")
                    return
                patient = patient_entry.get().strip()
                weight = int(weight_entry.get())
                dose = round(max(weight * 0.067 + 0.2 + (0.5 if weight > 90 else 0), 5.2), 2)
                row = cur.execute \
                    ("SELECT activity, date, time FROM elutions WHERE date=? AND time=? ORDER BY rowid DESC LIMIT 1", (date_str, el_time_str,)).fetchone()
                if not row:
                    messagebox.showerror("Error", "Selected elution not found for this date.")
                    return
                el_activity, el_date_str, el_time_str_db = row
                fmt_dt = "%d-%m-%Y %H:%M"
                el_dt = datetime.strptime(f"{el_date_str} {el_time_str_db}", fmt_dt)
                seg_dt = datetime.strptime(f"{date_str} {seg_time_str}", fmt_dt)
                admin_dt = datetime.strptime(f"{date_str} {admin_time_str}", fmt_dt)
                def decay_minutes(dt_minutes: float) -> float:
                    return math.exp(-(math.log(2) / T12_GA68) * dt_minutes)
                dt1_min = (seg_dt - el_dt).total_seconds() / 60.0
                conc_seg = round((el_activity / 5.2) * decay_minutes(dt1_min), 2)
                dt2_min = (admin_dt - seg_dt).total_seconds() / 60.0
                conc_admin = round(conc_seg * decay_minutes(dt2_min), 2)
                if conc_admin <= 0:
                    messagebox.showerror("Error", "Calculated concentration is not valid.")
                    return
                vol = round(dose / conc_admin, 1)
                dotatoc_tree.insert("", "end", values=(date_str, patient, weight, admin_time_str, f"{dose:.1f}", f"{conc_admin:.2f}", f"{vol:.1f}"))
            except Exception as e:
                messagebox.showerror("Error", str(e))
        Button(dotatoc_frame, text="Calculate", command=dotatoc_calc, **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
               width=10, height=1, font=(FONT_NAME ,10 ,"bold")).grid(row=3, column=4, padx=20, pady=5)

        # Delete Row
        def delete_dotatoc_row():
            selected = dotatoc_tree.selection()
            if not selected:
                messagebox.showwarning("Error", "No row selected.")
                return
            row_id = selected[0]
            if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected row?"):
                return
            cur.execute("DELETE FROM dotatoc WHERE ID=? " ,(row_id,))
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
               width=5, height=1, font=(FONT_NAME ,9 ,"bold")).grid(row=3, column=5, padx=6)

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
            Label(frame, text="Real Dose (mCi):", bg=BG, fg="white", font=(FONT_NAME ,10 ,"bold")).grid(row=0, column=0, padx=5, pady=10)
            real_dose_entry = Entry(frame, width=10)
            real_dose_entry.grid(row=0, column=1, padx=5, pady=10)
            Label(frame, text="ITLC (<2%):", bg=BG, fg="white", font=(FONT_NAME ,10 ,"bold")).grid(row=1, column=0, padx=5, pady=10)
            itlc_entry = Entry(frame, width=10)
            itlc_entry.grid(row=1, column=1, padx=5, pady=10)
            Label(frame, text="Residual (mCi):", bg=BG, fg="white", font=(FONT_NAME ,10 ,"bold")).grid(row=2, column=0, padx=5, pady=10)
            residual_entry = Entry(frame, width=10)
            residual_entry.grid(row=2, column=1, padx=5, pady=10)
            def save_to_dotatoc():
                real_dose = real_dose_entry.get().strip()
                itlc = itlc_entry.get().strip()
                residual = residual_entry.get().strip()
                if not patient or not weight or not real_dose or not itlc or not residual:
                    messagebox.showerror("Error", "Please fill all the fields.")
                    return
                cur.execute \
                    ("""INSERT INTO dotatoc (date, patient, weight, admin_time, dose, concentration, volume,real_dose, itlc, residual) VALUES (?,?,?,?,?,?,?,?,?,?)""",
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

        # Main Buttons
        btn_frame = Frame(tab, bg=C4)
        btn_frame.pack(pady=10)
        Button(btn_frame, text="Back", **TAB_BUTTON_STYLE, command=lambda nt=tab: app.back_to_main(nt)).pack()

    # RUN
    select_file()