from tkinter import *
from tkcalendar import DateEntry
from constants import *
from tkinter import messagebox
import os, pandas, sqlite3

def build_tab(app, tab):

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
            A0_mbq = float(activity_entry.get().strip())
            if not A0_mbq:
                raise ValueError("Please enter the Initial Activity (MBq)")
            admin_date = admin_entry.get_date()
            days_diff = (admin_date - cal_date).days
            df = math.exp(-math.log(2) * days_diff / T12_I131)
            A_admin_mbq = A0_mbq * df
            A_admin_mci = mbq_to_mci(A_admin_mbq)
            serial_number = serial_number_entry.get().strip()
            # Show Output
            result_txt.configure(state="normal")
            result_txt.delete("1.0", END)
            result_txt.tag_configure("normal", font=("Courier", 9, "normal"))
            result_txt.tag_configure("bold", font=("Courier", 11, "bold"))
            result_txt.insert(END, f"Patient Name: {patient_name}\n\n", "bold")
            result_txt.insert(END, f"Calibration Date: {cal_date.strftime("%d-%m-%Y")}\n\n", "normal")
            result_txt.insert(END,
                              f"Administration Date: {admin_date.strftime("%d-%m-%Y")} (offset {days_diff} days)\n\n",
                              "normal")
            result_txt.insert(END, f"Initial Activity: {A0_mbq:.2f} MBq ({mbq_to_mci(A0_mbq):.2f} mCi)\n\n", "normal")
            result_txt.insert(END, f"Decay Factor for {days_diff:.2f} days: {df:.2f}\n\n", "normal")
            result_txt.insert(END, f"Activity at Administration: {A_admin_mbq:.2f} MBq = {A_admin_mci:.2f} mCi\n\n",
                              "bold")
            result_txt.insert(END, f"Serial Number: {serial_number}\n\n", "normal")
        except Exception as e:
            messagebox.showerror("Calculation Error", str(e))

    #Save New Data
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
            cur.execute("INSERT INTO i131_records (patient_name, calibration_date, initial_mbq, admin_date, admin_mci, serial_number) VALUES (?,?,?,?,?,?)",
                        (patient_name, cal_date.strftime("%d-%m-%Y"), A0_mbq, admin_date.strftime("%d-%m-%Y"), A_admin_mci, serial_number))
            conn.commit()
            conn.close()
            excel_path = os.path.join(month_dir, f"I131_{admin_date.strftime("%m")}.xlsx")
            record = {"Patient": patient_name,
                      "Calibration Date": cal_date.strftime("%d-%m-%Y"),
                      "Initial Activity (MBq)": A0_mbq,
                      "Administration Date": admin_date.strftime("%d-%m-%Y"),
                      "Activity at Admin (mCi)": round(A_admin_mci, 3),
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

    # Main Buttons
    btn_frame = Frame(tab, bg=C4)
    btn_frame.pack(pady=10)
    calc_button = Button(btn_frame, text="Calculate", **TAB_BUTTON_STYLE, command=calculate_and_show)
    calc_button.grid(row=0, column=0, padx=10)
    save_button = Button(btn_frame, text="Save", **TAB_BUTTON_STYLE, command=save_record)
    save_button.grid(row=0, column=1, padx=10)
    back_button = Button(btn_frame, text="Back", **TAB_BUTTON_STYLE, command=lambda nt=tab: app.back_to_main(nt))
    back_button.grid(row=0, column=2, padx=10)