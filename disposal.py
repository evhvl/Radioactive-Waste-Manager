from tkinter import Label, filedialog, ttk, Toplevel
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from functions import *

#====VIALS DISPOSAL TAB=====
def build_vials_disposal_tab(parent_tab, *, on_back=None):
    for w in parent_tab.winfo_children():
        w.destroy()
    header = Label(parent_tab, text="VIALS DISPOSAL (LIVE STORAGE)", **TEXT_COLORS, font=(FONT_NAME, 18, "bold"))
    header.pack(pady=(10,5), fill="x")
    btns = Frame(parent_tab, bg=C4)
    btns.pack(pady=(10,15))
    #Tree
    columns = [("Vials", 75), ("Cal Date", 95), ("Stored at", 95), ("Activity(mCi)", 110), ("Permitted Date", 120), ("Recommended Date", 150), ("Limit(mCi)", 90), ("Status", 80),]
    tree = ttk.Treeview(parent_tab, columns=[c[0] for c in columns], show="headings", height=12)
    tree.pack(pady=(10,0))
    for col_name, col_width in columns:
        tree.heading(col_name, text=col_name)
        tree.column(col_name, anchor="center", width=col_width, stretch=False)
    style = ttk.Style()
    style.theme_use("default")
    style.configure("Treeview", background=C2, fieldbackground=C2, foreground="black", rowheight=26, borderwidth=1, bordercolor="black", relief="solid")
    style.configure("Treeview.Heading", background=C3, foreground="white", font=(FONT_NAME, 11, "bold"), relief="solid")
    style.map("Treeview", background=[("selected", "#8FAADC"), ("!selected", C2)], foreground=[("selected", "black")])
    style.layout("Treeview", [("Treeview.treearea", {"sticky": "nsew"})])
    tree.tag_configure("STORED", background="#CDECCF")
    tree.tag_configure("WAIT", background="#FFE9B3")
    tree.tag_configure("READY", background="#F6B3B3")
    summary_frame = Frame(parent_tab, bg=C4)
    summary_frame.pack(pady=(15,15))
    summary_label = Label(summary_frame, text="", **TEXT_COLORS, font=(FONT_NAME, 12, "bold"))
    summary_label.pack()
    #Functions
    def load_live_storage():
        rows = read_stored_vials()
        tree.delete(*tree.get_children())
        summary_rows = []
        for r in rows:
            rid = r[0]
            radionuclide = r[1]
            cal_date = r[2]
            stored_at = r[3]
            activity0 = r[4]
            permitted = r[5]
            recommended = r[6]
            limit_mci = r[7]
            status = disposal_status(recommended, permitted)
            tree.insert("", "end", iid=rid, values=[radionuclide,cal_date, stored_at, float(activity0), permitted if permitted is not None else "",
                        recommended if recommended is not None else "", "" if limit_mci is None else float(limit_mci), status], tags=(status,))
            summary_rows.append((rid, radionuclide, stored_at, activity0, permitted, recommended, limit_mci))
        total_vials, total_activity, ready_count = disposal_summary(rows)
        summary_label.config(text=f"Total vials: {total_vials}   |   Total activity ATM: {total_activity} mCi   |   READY: {ready_count}")
    def refresh():
        load_live_storage()
    def dispose_selected_vials():
        sel = tree.selection()
        if not sel:
            messagebox.showerror("Error", "Select 1 or more vials to dispose.")
            return
        ids = [int(x) for x in sel]
        if not messagebox.askyesno("Dispose Vials", f"Dispose selected vials: {len(ids)}?"):
            return
        full_rows =  read_vials_full_ids(ids)
        log_vials_disposal(full_rows)
        delete_vials_by_ids(ids)
        refresh()
        messagebox.showinfo("Disposed", f"Disposed {len(ids)} vials (logged to daily disposal file).")
    def check_ready_vials():
        rows = read_stored_vials()
        if not rows:
            messagebox.showinfo("Check READY", "No vials in live storage.")
            return
        today0 = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        ready_items = []
        for r in rows:
            rid, radionuclide, cal_date, stored_at, activity0, permitted, recommended, limit_mci = r
            if permitted is None:
                continue
            if not isinstance(permitted, str):
                permitted = str(permitted)
            try:
                perm_dt = datetime.strptime(permitted, DATE_FORMAT)
            except Exception:
                continue
            if today0 >= perm_dt:
                try:
                    status = disposal_status(recommended, permitted)
                except Exception:
                    continue
                if status != "READY":
                    continue
                a_now = activity_now(radionuclide, stored_at, activity0)
                ready_items.append({"id": int(rid),
                                    "radionuclide": radionuclide,
                                    "cal_date": cal_date,
                                    "stored_at": stored_at,
                                    "activity0": float(activity0),
                                    "activity_now": float(a_now),
                                    "permitted": permitted,
                                    "limit_mci": None if limit_mci is None else float(limit_mci)})
        if not ready_items:
            messagebox.showinfo("Check READY", "No READY vials found today.")
            return
        groups = {}
        for it in ready_items:
            nucl = it["radionuclide"]
            g = groups.setdefault(nucl, {"radionuclide": nucl,
                                         "limit_mci": it["limit_mci"],
                                         "total_now": 0.0,
                                         "items": []})
            g["total_now"] += it["activity_now"]
            g["items"].append(it)
        for g in groups.values():
            g["total_now"] = round(g["total_now"], 2)
            g["items"].sort(key=lambda x: (str(x["cal_date"]), x["id"]))
        eligible = {}
        for nucl, g in groups.items():
            lim = g["limit_mci"]
            if lim is None:
                eligible[nucl] = g
            else:
                if g["total_now"] <= float(lim):
                    eligible[nucl] = g
        if not eligible:
            messagebox.showinfo("Check READY", "READY vials exist today, but none are eligible to dispose (limit exceeded).")
            return
        lines = ["Eligible radionuclides (can be disposed today):"]
        for nucl, g in sorted(eligible.items(), key=lambda kv: kv[0]):
            lim_txt = "-" if g["limit_mci"] is None else f"{g['limit_mci']:.2f}"
            lines.append(f"- {nucl}: {g['total_now']:.2f} mCi (limit {lim_txt}) | vials: {len(g['items'])}")
        messagebox.showinfo("Check READY", "\n".join(lines))
        popup = Toplevel(parent_tab)
        popup.title("Eligible READY Vials (by radionuclide)")
        Label(popup, text="Select radionuclides to dispose", **TEXT_COLORS, font=(FONT_NAME, 18, "bold")).pack(pady=(0, 10))
        cols = ("Radionuclide", "Total A ATM(mCi)", "Limit(mCi)", "Count")
        t_groups = ttk.Treeview(popup, columns=cols, show="headings", height=6, selectmode="extended")
        for c in cols:
            t_groups.heading(c, text=c)
            t_groups.column(c, anchor="center", width=160, stretch=False)
        t_groups.column("Radionuclide", width=160)
        t_groups.column("Count", width=90)
        t_groups.pack(pady=(0, 10))
        det_cols = ("ID", "Nuclide", "Cal Date", "Stored at", "A0(mCi)", "A Now(mCi)", "Permitted")
        t_det = ttk.Treeview(popup, columns=det_cols, show="headings", height=10)
        widths = [70, 80, 95, 95, 90, 95, 95]
        for c, w in zip(det_cols, widths):
            t_det.heading(c, text=c)
            t_det.column(c, anchor="center", width=w, stretch=False)
        t_det.pack(pady=(0, 12))
        for nucl, g in sorted(eligible.items(), key=lambda kv: kv[0]):
            lim_txt = "-" if g["limit_mci"] is None else f"{g['limit_mci']:.2f}"
            t_groups.insert("", "end", iid=nucl, values=(nucl, f"{g['total_now']:.2f}", lim_txt, str(len(g["items"]))))
        def refresh_details(_evt=None):
            t_det.delete(*t_det.get_children())
            sel = t_groups.selection()
            if not sel:
                return
            for nucl in sel:
                g = eligible.get(nucl)
                if not g:
                    continue
                for it in g["items"]:
                    t_det.insert("", "end", values=(it["id"], it["radionuclide"], it["cal_date"], it["stored_at"],
                                                    f"{it['activity0']:.2f}", f"{it['activity_now']:.2f}", it["permitted"]))
        t_groups.bind("<<TreeviewSelect>>", refresh_details)
        t_groups.selection_set(t_groups.get_children())
        refresh_details()
        # ---- PDF export ----
        def export_pdf():
            try:
                disp_date = datetime.now().strftime(DATE_FORMAT)
                pdf_path = get_ready_vials_pdf_path(disp_date)
                styles = getSampleStyleSheet()
                doc = SimpleDocTemplate(pdf_path, pagesize=A4)
                story = []
                story.append(Paragraph(f"READY Vials Eligible for Disposal - {disp_date}", styles["Title"]))
                story.append(Spacer(1, 10))
                sel = list(t_groups.selection())
                if not sel:
                    sel = sorted(list(eligible.keys()))
                for nucl in sel:
                    g = eligible.get(nucl)
                    if not g:
                        continue
                    lim_txt = "-" if g["limit_mci"] is None else f"{g['limit_mci']:.2f} mCi"
                    story.append(Paragraph(f"{nucl} | Total Activity ATM (mCi): {g['total_now']:.2f} mCi | Limit: {lim_txt} | Count: {len(g['items'])}",
                                            styles["Heading2"]))
                    story.append(Spacer(1, 6))
                    data = [["Cal Date", "Stored at", "A0 (mCi)", "Activity Now (mCi)", "Permitted"]]
                    for it in g["items"]:
                        data.append([str(it["cal_date"]),
                                     str(it["stored_at"]),
                                     f"{it['activity0']:.2f}",
                                     f"{it['activity_now']:.2f}",
                                     str(it["permitted"])])
                    tbl = Table(data, hAlign="LEFT")
                    tbl.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                                            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                                            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                                            ("ALIGN", (3, 1), (-2, -1), "RIGHT")]))
                    story.append(tbl)
                    story.append(Spacer(1, 14))
                doc.build(story)
                messagebox.showinfo("PDF Export", f"PDF created:\n{pdf_path}")
                os.startfile(pdf_path)
            except Exception as e:
                messagebox.showerror("PDF Export Error", str(e))
        def dispose_selected_groups():
            sel = list(t_groups.selection())
            if not sel:
                messagebox.showerror("Error", "Select at least 1 radionuclide group.")
                return
            ids = []
            for nucl in sel:
                g = eligible.get(nucl)
                if not g:
                    continue
                ids.extend([it["id"] for it in g["items"]])
            if not ids:
                return
            if not messagebox.askyesno("Dispose", f"Dispose {len(ids)} vials from selected radionuclides?"):
                return
            full_rows = read_vials_full_ids(ids)
            log_vials_disposal(full_rows)
            delete_vials_by_ids(ids)
            popup.destroy()
            refresh()
            messagebox.showinfo("Disposed", f"Disposed {len(ids)} vials and logged them to sheet 'Vials'.")
        btn_row = Frame(popup, bg=C4)
        btn_row.pack(pady=(15, 0))
        Button(btn_row, text="Print PDF", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
               width=10, height=2, font=(FONT_NAME, 10, "bold"), command=export_pdf).grid(row=0, column=0, padx=6)
        Button(btn_row, text="OK", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
               width=10, height=2, font=(FONT_NAME, 10, "bold"), command=dispose_selected_groups).grid(row=0, column=1, padx=6)
        Button(btn_row, text="Close", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
               width=10, height=2, font=(FONT_NAME, 10, "bold"), command=popup.destroy).grid(row=0, column=2, padx=6)
        popup.update_idletasks()
        width = popup.winfo_reqwidth() + 100
        height = popup.winfo_reqheight() + 40
        popup.geometry(f"{width}x{height}")
        popup.resizable(False, False)
        popup.configure(bg=C4)
        center_window(popup, width, height)
    Button(btns, text="Refresh", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
            width=10, height=1, font=(FONT_NAME, 10, "bold"), command=refresh).grid(row=0, column=0, padx=6)
    Button(btns, text="Check READY", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
           width=14, height=1, font=(FONT_NAME, 10, "bold"), command=check_ready_vials).grid(row=0, column=1, padx=6)
    Button(btns, text="✗Dispose Selected✗", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']},
            width=18, height=1, font=(FONT_NAME, 10, "bold"), command=dispose_selected_vials).grid(row=0, column=2, padx=6)
    if on_back:
        Button(parent_tab, text="Back", **TAB_BUTTON_STYLE, command=on_back).pack(pady=(0, 10))
    def auto_refresh():
        refresh()
        parent_tab.after(6000, auto_refresh)
    auto_refresh()
    refresh()

#=====TC99M DISPOSAL TAB=====
def build_tc99m_disposal_tab(parent_tab, *, on_back=None):
    for w in parent_tab.winfo_children():
        w.destroy()
    state = {"batch_path": None, "read_only": False}
    header = Label(parent_tab, text="TC99M DISPOSAL BATCH", **TEXT_COLORS, font=(FONT_NAME, 18, "bold"))
    header.pack(pady=(10,5), fill="x")
    frame = Frame(parent_tab, bg=C4)
    frame.pack(pady=5)
    batch_label = Label(frame, text="", **TEXT_COLORS, font=(FONT_NAME, 12, "bold"))
    batch_label.grid(row=0, column=0, padx=10)
    mode_label = Label(frame, text="", bg=C4, fg="orange", font=(FONT_NAME, 10, "bold"))
    mode_label.grid(row=1, column=0, padx=10)
    action_btn = None
    #Fuctions
    def dispose_current_batch():
        if not state["batch_path"]:
            return
        created_at, finalized_at, disposed_at = read_batch_info(state["batch_path"])
        if finalized_at is None:
            messagebox.showwarning("Warning", "This batch is not Finalized.")
            return
        if disposed_at is not None:
            messagebox.showinfo("Info", "This batch is already disposed.")
            return
        if not messagebox.askyesno("Dispose Batch", "Dispose this FINALIZED Batch?\nThis will be logged and the batch will be marked as disposed."):
            return
        dispose_batch(state["batch_path"])
        messagebox.showinfo("Disposed", "Batch marked as disposed.")
        refresh()
    def load_batch(batch_path, read_only=False):
        state["batch_path"] = batch_path
        state["read_only"] = read_only
        rows = read_tc99m_items(batch_path)
        batch_label.config(text=f"Viewing: {os.path.basename(batch_path)}")
        if read_only:
            mode_label.config(text="READ ONLY (Finalized/Old Batch)")
        else:
            mode_label.config(text="ACTIVE BATCH")
        tree.delete(*tree.get_children())
        summary_rows = []
        for r in rows:
            iid, stored_at, activity_mci, permitted, recommended = r
            status = disposal_status(recommended, permitted)
            tree.insert("", "end", iid=iid, values=[iid, stored_at, float(activity_mci), permitted, recommended, status], tags=(status,))
            summary_rows.append((iid, TC99M_NUCLIDE, stored_at, float(activity_mci), permitted, recommended, None))
            total_items, total_activity, ready_count = disposal_summary(summary_rows)
            summary_label.config(text=f"Total vials: {total_items}   |   Total activity ATM: {total_activity} mCi   |   READY: {ready_count}")
            created_at, finalized_at, disposed_at = read_batch_info(batch_path)
            if not read_only:
                action_btn.config(text="✗Finalize Batch✗", command=finalize_batch, state="normal")
                return
            if finalized_at is not None and disposed_at is None:
                action_btn.config(text="✗Dispose Batch✗", command=dispose_current_batch, state="normal")
            else:
                action_btn.config(text="✗Dispose Batch✗", command=lambda: None, state="disabled")
    def refresh():
        if not state["batch_path"]:
            load_active()
            return
        load_batch(state["batch_path"], read_only=state["read_only"])
    def load_active():
        active = get_active_batch()
        load_batch(active, read_only=False)
    def open_old_batch():
        folder = filedialog.askdirectory(title="Select Old (Finalized) Batch Folder", initialdir=TC99M_DIR)
        if not folder:
            return
        load_batch(folder, read_only=True)
    def finalize_batch():
        if state["read_only"]:
            return
        if not messagebox.askyesno("Finalize Batch", "Are you sure you want to finalize this ACTIVE batch?\nAfter finalize, new stored items will go into a NEW BATCH."):
            return
        old_batch, new_batch = finalize_active_batch()
        messagebox.showinfo("Batch Finalized", f"Old batch closed:\n{os.path.basename(old_batch)}\n\nNew active batch:\n\n{os.path.basename(new_batch)}")
        load_batch(new_batch, read_only=False)
    btns = Frame(parent_tab, bg=C4)
    btns.pack(pady=(10,15))
    Button(btns, text="Refresh", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=10, height=1, font=(FONT_NAME, 10, "bold"), command=refresh).grid(row=0, column=0, padx=6)
    Button(btns, text="Open Old Batch",**{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=14, height=1, font=(FONT_NAME, 10, "bold"), command=open_old_batch).grid(row=0, column=1, padx=6)
    Button(btns, text="Back to Active", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=14, height=1, font=(FONT_NAME, 10, "bold"), command=load_active).grid(row=0, column=2, padx=6)
    action_btn = Button(btns, text="✗Finalize Batch✗", **{k: v for k, v in TAB_BUTTON_STYLE.items() if k not in ['width', 'height', 'font']}, width=16, height=1, font=(FONT_NAME, 10, "bold"), command=finalize_batch)
    action_btn.grid(row=0, column=3, padx=6)
    #Tree
    columns = [("Item", 80), ("Stored at", 95), ("Activity (mCi)", 110), ("Permitted Date", 120), ("Recommended Date", 150), ("Status", 85),]
    tree = ttk.Treeview(parent_tab, columns=[c[0] for c in columns], show="headings", height=11)
    tree.pack(pady=(10,0))
    for col_name, col_width in columns:
        tree.heading(col_name, text=col_name)
        tree.column(col_name, anchor="center", width=col_width, stretch=False)
    style = ttk.Style()
    style.theme_use("default")
    style.configure("Treeview", background=C2, fieldbackground=C2, foreground="black", rowheight=26, borderwidth=1, bordercolor="black", relief="solid")
    style.configure("Treeview.Heading", background=C3, foreground="white", font=(FONT_NAME, 11, "bold"), relief="solid")
    style.map("Treeview", background=[("selected", "#8FAADC"), ("!selected", C2)], foreground=[("selected", "black")])
    style.layout("Treeview", [("Treeview.treearea", {"sticky": "nsew"})])
    tree.tag_configure("STORED", background="#CDECCF")
    tree.tag_configure("WAIT", background="#FFE9B3")
    tree.tag_configure("READY", background="#F6B3B3")
    summary_frame = Frame(parent_tab, bg=C4)
    summary_frame.pack(pady=(5,10))
    summary_label = Label(summary_frame, text="", **TEXT_COLORS, font=(FONT_NAME, 12, "bold"))
    summary_label.pack()
    if on_back:
        Button(parent_tab, text="Back", **TAB_BUTTON_STYLE, command=on_back).pack(pady=(0,10))
    def auto_refresh():
        refresh()
        parent_tab.after(6000, auto_refresh)
    auto_refresh()
    load_active()