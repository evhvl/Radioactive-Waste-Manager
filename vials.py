import sqlite3, disposal
from functions import *
from constants import *
from tkinter import *
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from datetime import datetime
from pathlib import Path
from typing import Optional

def build_tab(app, tab, vial_name):

    # Choose New or Old File
    def select_vial_file():
        popup_window = Toplevel(app.window)
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

    # Create New File
    def new_vial_file():
        popup = Toplevel(app.window)
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
        refresh_time_button = Button(time_field, text="↻", command=lambda nt=time_entry: update_time(nt), bg="white",
                                     fg="black", bd=0,
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
            cur.execute(
                """CREATE TABLE IF NOT EXISTS vial_info(date TEXT, time TEXT, activity REAL, volume REAL, concentration REAL, expiration_date TEXT, stored_date TEXT)""")
            cur.execute(
                """CREATE TABLE IF NOT EXISTS patient_info(id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT, time TEXT, patient_name TEXT, concentration REAL, dose_planned REAL, volume_planned REAL, dose_actual REAL, volume_actual REAL, volume_left REAL)""")
            cur.execute("""INSERT INTO vial_info VALUES (?,?,?,?,?,?,NULL)""",
                        (date, time, activity, volume, conc, exp_date))
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
        Button(bttn_frame, text="Back",
               **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=8, height=1,
               font=(FONT_NAME, 10, "bold"),
               command=lambda: (popup.destroy(), app.tabs_frame.forget(tab), app.create_new_tab("Vials"))).grid(row=0, column=1, padx=10, pady=10)

    # Open Old File
    def existing_vial_file():
        popup = Toplevel(app.window)
        popup.title(f"Open Existing {vial_name} Vial File")
        popup.config(bg=C4)
        center_window(window=popup, w=360, h=130)
        Label(popup, text=f"Select Existing {vial_name} Folder", **TEXT_COLORS, font=(FONT_NAME, 17, "bold")).pack(
            pady=10)

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
        Button(button_frame, text="Open File 🗁",
               **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
               width=12, height=2, font=(FONT_NAME, 12, "bold"), command=open_folder).grid(row=0, column=0, padx=10,
                                                                                           pady=10)
        Button(button_frame, text="Back",
               **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=12,
               height=2, font=(FONT_NAME, 12, "bold"),
               command=lambda: (popup.destroy(), app.tabs_frame.forget(tab), app.create_new_tab("Vials"))).grid(row=0, column=1, padx=10, pady=10)

    #Load Tab
    def load_vial(dbfile):
        for widget in tab.winfo_children():
            widget.destroy()
        header = Label(tab, text=f"{vial_name} Log Sheet", **TEXT_COLORS, font=(FONT_NAME, 18, "bold"))
        header.pack(pady=(5, 0), fill="x")
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

        # Store Vial
        def store_current_vial():
            if not messagebox.askyesno("Store Vial","Are you sure you want to store this vial for disposal?\nNo further administrations will be allowed."):
                return
            last_left = cur.execute("SELECT volume_left FROM patient_info ORDER BY id DESC LIMIT 1").fetchone()
            volume_left_now = float(last_left[0]) if last_left and last_left[0] is not None else float(volume)
            if volume_left_now < 0:
                messagebox.showerror("Error", "No volume left in vial.")
                return
            vial_dt = datetime.strptime(f"{date} {time}", "%d-%m-%Y %H:%M")
            now_dt = datetime.now()
            delta_minutes = (now_dt - vial_dt).total_seconds() / 60
            if delta_minutes < 0:
                delta_minutes = 0
            decay_factor = math.exp(-math.log(2) * delta_minutes / (half_life * 60))
            current_conc = float(conc) * decay_factor
            calculated_activity_mci = round(current_conc * volume_left_now, 2)
            #(Residual Popup)
            popup = Toplevel(tab)
            popup.title("Enter Residual")
            popup.config(bg=C4, pady=12, padx=13)
            center_window(popup, 300, 140)
            Label(popup, text="Residual Activity (mCi):", **TEXT_COLORS, font=(FONT_NAME, 11, "bold")).pack(pady=(5,6))
            res_var = StringVar(value=f"{calculated_activity_mci:.2f}")
            res_entry = Entry(popup, width=12, textvariable=res_var, font=(FONT_NAME, 12))
            res_entry.pack(pady=(0,10))
            res_entry.focus_set()
            res_entry.selection_range(0, "end")
            final_activity: Optional[float] = None
            def on_ok():
                nonlocal final_activity
                try:
                    v = float(res_var.get().replace(",","."))
                    if v < 0:
                        raise ValueError
                except ValueError:
                    messagebox.showerror("Error", "Please enter a valid residual activity in mCi.")
                    return
                final_activity = round(v, 2)
                popup.destroy()
            def on_cancel():
                popup.destroy()
            buttons_frame = Frame(popup, bg=C4)
            buttons_frame.pack(pady=(6,0))
            Button(buttons_frame, text="OK", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                   width=10, height=1, font=(FONT_NAME, 10, "bold"), command=on_ok).grid(row=0, column=0, padx=8)
            Button(buttons_frame, text="Cancel", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                   width=10, height=1, font=(FONT_NAME, 10, "bold"), command=on_cancel).grid(row=0, column=1, padx=8)
            popup.transient(tab)
            popup.grab_set()
            tab.wait_window(popup)
            if final_activity is None:
                return
            current_activity_mci = final_activity
            stored_at = datetime.now().strftime("%d-%m-%Y")
            recommended, permitted, limit_bq = disposal.calc_recommended_and_permitted_date(vial_name, current_activity_mci, stored_at)
            if recommended is None or permitted is None:
                return
            disposal.store_vial(radionuclide=vial_name, source_db=dbfile, calibration_date=date, stored_at=stored_at,
                                activity_mci=current_activity_mci, permitted_date=permitted,
                                recommended_date=recommended, limit_bq=limit_bq)
            cur.execute("UPDATE vial_info SET stored_date=?", (stored_at,))
            conn.commit()
            folder = os.path.dirname(dbfile)
            excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
            wb = load_workbook(excel_path)
            ws = wb["Vial Info"]
            ws.cell(row=2, column=7, value=stored_at)
            wb.save(excel_path)
            messagebox.showinfo("Vial Stored",f"{vial_name} vial stored successfully.\n\nRecommended disposal after: {recommended}\nPermitted disposal after: {permitted}")

        dispose_button = Button(info_frame, text="✗Store Vial✗",**{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                                width=14, height=1, font=(FONT_NAME, 10, "bold"), command=store_current_vial)
        dispose_button.grid(row=2, column=1, padx=6, pady=6)

        # Table
        columns = [("date", "Date", 120), ("time", "Time", 100), ("patient_name", "Patient", 160),
                   ("concentration", "Conc(mCi/ml)", 120), ("dose", "Dose(mCi)", 100), ("volume", "Vol(ml)", 90),
                   ("volume_left", "Vol Left(ml)", 100)]
        tree = ttk.Treeview(tab, columns=[c[0] for c in columns], show="headings", height=8)
        tree.pack(pady=10)
        for col_id, col_title, col_width in columns:
            tree.heading(col_id, text=col_title)
            tree.column(col_id, width=col_width, anchor="center")
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background=C2, fieldbackground=C2, foreground="black", rowheight=26, borderwidth=1,
                        bordercolor="black", relief="solid")
        style.configure("Treeview.Heading", background=C3, foreground="white", font=(FONT_NAME, 11, "bold"),
                        relief="solid")
        style.map("Treeview", background=[("selected", "#8FAADC"), ("!selected", C2)],
                  foreground=[("selected", "black")])
        style.layout("Treeview", [("Treeview.treearea", {"sticky": "nsew"})])

        # Load Old Data
        rows = cur.execute(
            "SELECT id,date,time,patient_name,concentration,dose_planned,volume_planned,dose_actual,volume_actual,volume_left FROM patient_info").fetchall()
        for r in rows:
            row_id = r[0]
            (date, time, patient, conc, dose_p, vol_p, dose_a, vol_a, vol_left) = r[1:]
            dose_txt = f"{dose_p:.2f}" if dose_p is not None else "-"
            vol_txt = f"{vol_p:.2f}" if vol_p is not None else "-"
            if dose_a is not None:
                dose_txt += f" → {dose_a:.2f}"
            if vol_a is not None:
                vol_txt += f" → {vol_a:.2f}"
            tree.insert("", "end", iid=row_id, values=(date, time, patient, f"{conc:.2f}" if conc is not None else "-", dose_txt, vol_txt,
                                                                    f"{vol_left:.2f}" if vol_left is not None else "-"))
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

        #Save New Data
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
            cur.execute("INSERT INTO patient_info (date, time, patient_name, concentration, dose_planned, volume_planned, dose_actual, volume_actual, volume_left)"
                        " VALUES (?,?,?,?,?,?,?,?,?)",
                (admin_date, admin_time, patient_name_entry.get().strip(), updated_conc, dose, dose_volume, None, None,
                 volume_left))
            conn.commit()
            row_id = cur.lastrowid
            tree.insert("", "end", iid=row_id, values=(admin_date, admin_time, patient_name_entry.get().strip(), f"{updated_conc:.2f}", dose,
                                                                    f"{dose_volume:.1f}", f"{volume_left:.1f}"))

            #Enter + Save Corrected Dose + Volume
            def on_double_click(event):
                selected = tree.selection()
                if not selected:
                    return
                row_id = int(selected[0])
                popup = Toplevel(tab)
                popup.title("Insert Actual Administration Values")
                popup.config(bg=C4, pady=15)
                center_window(popup, 250, 150)
                Label(popup, text="Actual Dose (mCi):", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=0, column=0, padx=10, pady=5)
                dose_actual_entry = Entry(popup, width=8)
                dose_actual_entry.insert(0, f"{dose}")
                dose_actual_entry.grid(row=0, column=1)
                Label(popup, text="Actual Volume (ml):", **TEXT_COLORS, font=(FONT_NAME, 10, "bold")).grid(row=1, column=0, padx=10, pady=5)
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
                    prev = cur.execute("""SELECT volume_left FROM patient_info WHERE id < ? ORDER BY id DESC LIMIT 1""",
                                       (row_id,)).fetchone()
                    prev_left = prev[0] if prev else volume
                    new_left = round(prev_left - volume_actual, 1)
                    if new_left < 0:
                        messagebox.showerror("Error", "Not enough volume left.")
                        return
                    cur.execute("""UPDATE patient_info SET dose_actual=?, volume_actual=?, volume_left=? WHERE id=?""",
                                (dose_actual, volume_actual, new_left, row_id))
                    conn.commit()
                    planned = cur.execute("""SELECT dose_planned, volume_planned FROM patient_info WHERE id=?""",
                                          (row_id,)).fetchone()
                    tree.item(row_id, values=(tree.item(row_id, "values")[0], tree.item(row_id, "values")[1],
                                              tree.item(row_id, "values")[2], tree.item(row_id, "values")[3],
                                              f"{planned[0]:.2f} → {dose_actual:.2f}",
                                              f"{planned[1]:.2f} → {volume_actual:.2f}", f"{new_left:.2f}"))
                    folder = os.path.dirname(dbfile)
                    excel_path = os.path.join(folder, f"{os.path.basename(folder)}.xlsx")
                    append_row_to_sheet(excel_path, "Administrations",
                                        [row_id, admin_date, admin_time, patient_name_entry.get(), updated_conc,
                                         f"{dose} → {dose_actual}", f"{dose_volume} → {volume_actual}", new_left])
                    popup.destroy()

                Button(popup, text="OK", command=save_actual,
                       **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                       width=10, height=2, font=(FONT_NAME, 10, "bold")).grid(row=2, column=0, pady=10, padx=6)
                Button(popup, text="Cancel", command=popup.destroy,
                       **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['bg', 'width', 'height', 'font']},
                       bg=C4, width=10, height=2, font=(FONT_NAME, 10, "bold")).grid(row=2, column=1, pady=10, padx=6)

            tree.bind("<Double-1>", on_double_click)
            dose_entry.delete(0, "end")

        add_button = Button(add_frame, text="Add", command=add_record,
                            **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
                            width=5, height=1, font=(FONT_NAME, 10, "bold"))
        add_button.grid(row=0, column=8, padx=6)

        # Update Volume Left After Delete
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

        #Delete Data
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
            rows = cur.execute("SELECT id, date, time, patient_name, concentration, dose_planned, volume_planned, dose_actual, volume_actual, volume_left FROM patient_info ORDER BY id").fetchall()
            for r in rows:
                row_id = r[0]
                (date, time, patient, conc, dose_p, vol_p, dose_a, vol_a, vol_left) = r[1:]
                dose_txt = f"{dose_p:.1f}" if dose_p is not None else "-"
                vol_txt = f"{vol_p:.1f}" if vol_p is not None else "-"
                if dose_a is not None:
                    dose_txt += f" → {dose_a:.1f}"
                if vol_a is not None:
                    vol_txt += f" → {vol_a:.1f}"
                tree.insert("", "end", iid=row_id, values=(date, time, patient, f"{conc:.2f}" if conc is not None else "-", dose_txt, vol_txt,
                                                                        f"{vol_left:.2f}" if vol_left is not None else "-"))

        Button(add_frame, text="🗑", command=delete_record,
               **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
               width=5, height=1, font=(FONT_NAME, 10, "bold")).grid(row=0, column=9, padx=6)

        # ----
        if is_expired:
            Label(tab, text="⚠ VIAL EXPIRED – NO ADMINISTRATION ALLOWED", fg="red", bg=C4,
                  font=(FONT_NAME, 11, "bold")).pack(pady=(5, 0))
            disable_buttons(tab, exempt_texts=["Back", "✗Store Vial✗"])
        if stored_date:
            Label(tab, text=f"⚠ VIAL STORED ({stored_date})-NO FURTHER ACTIONS ALLOWED", fg="red", bg=C4,
                  font=(FONT_NAME, 11, "bold"), justify="center").pack(pady=(5, 0))
            disable_buttons(tab, exempt_texts=["Back"])

        # Main Buttons
        btn_frame = Frame(tab, bg=C4)
        btn_frame.pack(pady=(20, 0))
        Button(btn_frame, text="Back", **TAB_BUTTON_STYLE, command=lambda nt=tab: app.back_to_main(nt)).pack()

    # RUN
    select_vial_file()