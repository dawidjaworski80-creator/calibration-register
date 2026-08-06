import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import os
import importlib


class CalibrationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Measuring Equipment Calibration Register")
        self.root.state("zoomed")
        self.root.configure(bg="#1e2a3a")

        self._setup_styles()
        self.init_db()
        self.certificates_folder = self.get_setting("certificates_folder", "")
        self._calendar_class = None

        self.categories = {
            "External": "external",
            "Internal": "internal",
            "Gauges": "gauges",
            "Other": "other",
        }

        # Header bar
        header = tk.Frame(self.root, bg="#1e2a3a", pady=10)
        header.pack(fill="x")
        tk.Label(
            header,
            text="Measuring Equipment Calibration Register",
            bg="#1e2a3a",
            fg="#e0e8f0",
            font=("Segoe UI", 16, "bold"),
        ).pack()

        self.notebook = ttk.Notebook(self.root, style="Custom.TNotebook")
        self.notebook.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.tabs = {}
        self.trees = {}
        self.cert_preview_widgets = {}

        for tab_name, cat_code in self.categories.items():
            tab_frame = ttk.Frame(self.notebook, style="Custom.TFrame")
            self.notebook.add(tab_frame, text=f"  {tab_name}  ")
            self.tabs[cat_code] = tab_frame
            self.build_tab_ui(tab_frame, cat_code)

        tool_db_frame = ttk.Frame(self.notebook, style="Custom.TFrame")
        self.notebook.add(tool_db_frame, text="  Tool Database  ")
        self.build_tool_db_tab(tool_db_frame)

    def _setup_styles(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")

        BG = "#1e2a3a"
        TAB_BG = "#253447"
        TAB_ACTIVE = "#2e86c1"
        TAB_FG = "#a0b8cc"
        TAB_ACTIVE_FG = "#ffffff"
        FRAME_BG = "#253447"
        BTN_ADD = "#27ae60"
        BTN_EDIT = "#2e86c1"
        BTN_DEL = "#c0392b"
        BTN_FG = "#ffffff"
        TREE_BG = "#1e2a3a"
        TREE_FG = "#d0e4f0"
        TREE_HEAD_BG = "#2e86c1"
        TREE_HEAD_FG = "#ffffff"
        TREE_SEL = "#2e86c1"
        ROW_ODD = "#253447"
        ROW_EVEN = "#1e2a3a"

        style.configure("Custom.TNotebook", background=BG, borderwidth=0)
        style.configure(
            "Custom.TNotebook.Tab",
            background=TAB_BG,
            foreground=TAB_FG,
            padding=[18, 8],
            font=("Segoe UI", 11, "bold"),
            borderwidth=0,
        )
        style.map(
            "Custom.TNotebook.Tab",
            background=[("selected", TAB_ACTIVE), ("active", "#34495e")],
            foreground=[("selected", TAB_ACTIVE_FG), ("active", "#ffffff")],
        )

        style.configure("Custom.TFrame", background=FRAME_BG)
        style.configure(
            "Custom.TLabel",
            background=FRAME_BG,
            foreground=TREE_FG,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Custom.TEntry",
            fieldbackground="#2c3e50",
            foreground=TREE_FG,
            insertcolor=TREE_FG,
        )

        for name, bg in (("Add.TButton", BTN_ADD), ("Edit.TButton", BTN_EDIT), ("Del.TButton", BTN_DEL)):
            style.configure(
                name,
                background=bg,
                foreground=BTN_FG,
                font=("Segoe UI", 10, "bold"),
                padding=[14, 6],
                borderwidth=0,
                relief="flat",
            )
            style.map(name, background=[("active", bg)])

        style.configure(
            "Calendar.TButton",
            background=BTN_EDIT,
            foreground=BTN_FG,
            font=("Segoe UI", 9, "bold"),
            padding=[8, 4],
            borderwidth=0,
            relief="flat",
        )
        style.map("Calendar.TButton", background=[("active", BTN_EDIT)])

        style.configure(
            "Custom.Treeview",
            background=TREE_BG,
            foreground=TREE_FG,
            fieldbackground=TREE_BG,
            rowheight=28,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Custom.Treeview.Heading",
            background=TREE_HEAD_BG,
            foreground=TREE_HEAD_FG,
            font=("Segoe UI", 10, "bold"),
            relief="flat",
        )
        style.map("Custom.Treeview", background=[("selected", TREE_SEL)])

        self._row_odd = ROW_ODD
        self._row_even = ROW_EVEN

    def init_db(self):
        self.conn = sqlite3.connect("calibration_data.db")
        self.cursor = self.conn.cursor()
        self.cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS equipment (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                category TEXT,
                name TEXT,
                serial_number TEXT,
                last_calibration TEXT,
                next_calibration TEXT,
                measurement_range TEXT,
                certificate_path TEXT
            )
        """
        )
        # Lightweight migration for existing databases created before measurement_range was added.
        self.cursor.execute("PRAGMA table_info(equipment)")
        equipment_columns = [row[1] for row in self.cursor.fetchall()]
        if "measurement_range" not in equipment_columns:
            self.cursor.execute("ALTER TABLE equipment ADD COLUMN measurement_range TEXT")
        if "certificate_path" not in equipment_columns:
            self.cursor.execute("ALTER TABLE equipment ADD COLUMN certificate_path TEXT")

        self.cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS app_settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """
        )
        self.cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS tool_types (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                description TEXT,
                measurement_range TEXT
            )
        """
        )
        # Lightweight migration for existing databases created before measurement_range was added.
        self.cursor.execute("PRAGMA table_info(tool_types)")
        tool_type_columns = [row[1] for row in self.cursor.fetchall()]
        if "measurement_range" not in tool_type_columns:
            self.cursor.execute("ALTER TABLE tool_types ADD COLUMN measurement_range TEXT")
        defaults = [
            ("Micrometer", "External micrometer"),
            ("Caliper", "Vernier / digital caliper"),
            ("Groove Caliper", "Caliper for internal grooves"),
            ("Depth Micrometer", "Micrometer for depth measurements"),
            ("Bore Gauge", "Internal bore measuring gauge"),
            ("Dial Indicator", "Dial test indicator"),
            ("Height Gauge", "Precision height measuring gauge"),
            ("Feeler Gauge", "Thickness / gap feeler gauge"),
            ("Thread Gauge", "Go / No-Go thread gauge"),
            ("Radius Gauge", "Fillet and radius gauge set"),
        ]
        self.cursor.executemany(
            "INSERT OR IGNORE INTO tool_types (name, description) VALUES (?, ?)", defaults
        )
        self.conn.commit()

    def get_setting(self, key, default=""):
        self.cursor.execute("SELECT value FROM app_settings WHERE key=?", (key,))
        row = self.cursor.fetchone()
        return row[0] if row else default

    def set_setting(self, key, value):
        self.cursor.execute(
            "INSERT OR REPLACE INTO app_settings (key, value) VALUES (?, ?)",
            (key, value),
        )
        self.conn.commit()

    def choose_certificates_folder(self):
        selected = filedialog.askdirectory(title="Select certificates folder")
        if not selected:
            return
        self.certificates_folder = selected
        self.set_setting("certificates_folder", selected)
        for category in self.categories.values():
            label = self.cert_folder_labels.get(category)
            if label is not None:
                label.config(text=f"Folder: {selected}")

    def get_certificate_path_for_equipment(self, equipment_id):
        self.cursor.execute(
            "SELECT certificate_path FROM equipment WHERE id=?",
            (equipment_id,),
        )
        row = self.cursor.fetchone()
        return row[0] if row and row[0] else ""

    def set_certificate_path_for_equipment(self, equipment_id, cert_path):
        self.cursor.execute(
            "UPDATE equipment SET certificate_path=? WHERE id=?",
            (cert_path, equipment_id),
        )
        self.conn.commit()

    def resolve_certificate_path_for_equipment(self, equipment_id, serial_number):
        saved_path = self.get_certificate_path_for_equipment(equipment_id)
        if saved_path and os.path.isfile(saved_path):
            return saved_path

        # For legacy records or moved files, re-link by serial number from selected folder.
        discovered_path = self.find_certificate_for_serial(serial_number)
        if discovered_path:
            if discovered_path != saved_path:
                self.set_certificate_path_for_equipment(equipment_id, discovered_path)
            return discovered_path

        return saved_path

    def open_selected_certificate(self, category):
        widgets = self.cert_preview_widgets.get(category)
        if widgets is None:
            return
        cert_path = widgets.get("current_path", "")
        if not cert_path or not os.path.isfile(cert_path):
            messagebox.showwarning("Open certificate", "No certificate file available for this item.")
            return
        try:
            os.startfile(cert_path)
        except OSError as exc:
            messagebox.showerror("Open certificate", f"Could not open certificate file.\n{exc}")

    def update_certificate_preview(self, category):
        tree = self.trees.get(category)
        widgets = self.cert_preview_widgets.get(category)
        if tree is None or widgets is None:
            return

        selected = tree.selection()
        if not selected:
            widgets["title"].config(text="No equipment selected")
            widgets["serial"].config(text="Serial number: -")
            widgets["path"].config(text="Certificate path: -")
            widgets["status"].config(text="Status: -")
            widgets["current_path"] = ""
            widgets["open_button"].config(state="disabled")
            return

        values = tree.item(selected[0]).get("values", [])
        if not values:
            return

        equipment_id = values[0]
        equipment_name = values[1] if len(values) > 1 else "-"
        serial_number = values[2] if len(values) > 2 else "-"
        cert_path = self.resolve_certificate_path_for_equipment(equipment_id, str(serial_number))
        if cert_path and os.path.isfile(cert_path):
            status = "Status: linked"
            path_text = f"Certificate path: {cert_path}"
            widgets["open_button"].config(state="normal")
        elif cert_path:
            status = "Status: missing file"
            path_text = f"Certificate path: {cert_path}"
            widgets["open_button"].config(state="disabled")
        else:
            status = "Status: not linked"
            path_text = "Certificate path: -"
            widgets["open_button"].config(state="disabled")

        widgets["title"].config(text=f"{equipment_name}")
        widgets["serial"].config(text=f"Serial number: {serial_number}")
        widgets["path"].config(text=path_text)
        widgets["status"].config(text=status)
        widgets["current_path"] = cert_path

    def _on_equipment_select(self, category, _event=None):
        self.update_certificate_preview(category)

    def find_certificate_for_serial(self, serial_number):
        serial = serial_number.strip()
        if not serial:
            return None
        if not self.certificates_folder:
            return None
        docx_path = os.path.join(self.certificates_folder, f"{serial}.docx")
        doc_path = os.path.join(self.certificates_folder, f"{serial}.doc")
        if os.path.isfile(docx_path):
            return docx_path
        if os.path.isfile(doc_path):
            return doc_path
        return None

    def build_tool_db_tab(self, frame):
        btn_frame = ttk.Frame(frame, style="Custom.TFrame")
        btn_frame.pack(fill="x", padx=10, pady=10)
        ttk.Button(btn_frame, text="+ Add Tool", style="Add.TButton",
                   command=self.add_tool_type).pack(side="left", padx=(0, 6))
        ttk.Button(btn_frame, text="Edit", style="Edit.TButton",
                   command=self.edit_tool_type).pack(side="left", padx=(0, 6))
        ttk.Button(btn_frame, text="Delete", style="Del.TButton",
                   command=self.delete_tool_type).pack(side="left")

        columns = ("ID", "Tool Name", "Description", "Measurement Range")
        self.tool_tree = ttk.Treeview(frame, columns=columns, show="headings", style="Custom.Treeview")
        self.tool_tree.heading("ID", text="ID")
        self.tool_tree.heading("Tool Name", text="Tool Name")
        self.tool_tree.heading("Description", text="Description")
        self.tool_tree.column("ID", width=40, anchor="center")
        self.tool_tree.column("Tool Name", width=220)
        self.tool_tree.column("Description", width=350)
        self.tool_tree.heading("Measurement Range", text="Measurement Range")
        self.tool_tree.column("Measurement Range", width=220)
        self.tool_tree.tag_configure("odd", background=self._row_odd)
        self.tool_tree.tag_configure("even", background=self._row_even)
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.tool_tree.yview)
        self.tool_tree.configure(yscrollcommand=sb.set)
        self.tool_tree.bind("<Double-1>", self._on_tool_double_click)
        self.tool_tree.pack(fill="both", expand=True, padx=10, pady=(0, 10), side="left")
        sb.pack(fill="y", pady=(0, 10), side="left")
        self.load_tool_types()

    def _on_tool_double_click(self, event):
        row_id = self.tool_tree.identify_row(event.y)
        if not row_id:
            return
        self.tool_tree.selection_set(row_id)
        self.tool_tree.focus(row_id)
        self.edit_tool_type()

    def load_tool_types(self):
        for row in self.tool_tree.get_children():
            self.tool_tree.delete(row)
        self.cursor.execute(
            "SELECT id, name, description, measurement_range FROM tool_types ORDER BY name"
        )
        for i, row in enumerate(self.cursor.fetchall()):
            self.tool_tree.insert("", "end", values=row, tags=("odd" if i % 2 else "even",))

    def _get_tool_names(self):
        self.cursor.execute("SELECT name FROM tool_types ORDER BY name")
        return [r[0] for r in self.cursor.fetchall()]

    def open_tool_window(self, title, data=None, on_save=None):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("460x250")
        win.configure(bg="#253447")
        win.grab_set()
        win.resizable(False, False)
        ttk.Label(win, text="Tool Name:", style="Custom.TLabel").grid(
            row=0, column=0, padx=16, pady=12, sticky="w")
        e_name = ttk.Entry(win, style="Custom.TEntry", width=28)
        e_name.grid(row=0, column=1, padx=16, pady=12)
        ttk.Label(win, text="Description:", style="Custom.TLabel").grid(
            row=1, column=0, padx=16, pady=12, sticky="w")
        e_desc = ttk.Entry(win, style="Custom.TEntry", width=28)
        e_desc.grid(row=1, column=1, padx=16, pady=12)
        ttk.Label(win, text="Measurement Range:", style="Custom.TLabel").grid(
            row=2, column=0, padx=16, pady=12, sticky="w")
        e_range = ttk.Entry(win, style="Custom.TEntry", width=28)
        e_range.grid(row=2, column=1, padx=16, pady=12)
        if data:
            e_name.insert(0, data[1])
            e_desc.insert(0, data[2] or "")
            e_range.insert(0, data[3] or "")

        def save():
            name = e_name.get().strip()
            if not name:
                messagebox.showwarning("Error", "Tool name is required!", parent=win)
                return
            on_save(name, e_desc.get().strip(), e_range.get().strip())
            win.destroy()

        ttk.Button(win, text="Save", style="Add.TButton", command=save).grid(
            row=3, column=0, columnspan=2, pady=16)

    def add_tool_type(self):
        def save_action(name, desc, measurement_range):
            try:
                self.cursor.execute(
                    "INSERT INTO tool_types (name, description, measurement_range) VALUES (?, ?, ?)",
                    (name, desc, measurement_range),
                )
                self.conn.commit()
                self.load_tool_types()
            except sqlite3.IntegrityError:
                messagebox.showwarning("Error", f"Tool '{name}' already exists.")
        self.open_tool_window("Add Tool", on_save=save_action)

    def edit_tool_type(self):
        selected = self.tool_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a tool to edit.")
            return
        data = self.tool_tree.item(selected[0])["values"]
        tool_id = data[0]

        def save_action(name, desc, measurement_range):
            try:
                self.cursor.execute(
                    "UPDATE tool_types SET name=?, description=?, measurement_range=? WHERE id=?",
                    (name, desc, measurement_range, tool_id),
                )
                self.conn.commit()
                self.load_tool_types()
            except sqlite3.IntegrityError:
                messagebox.showwarning("Error", f"Tool '{name}' already exists.")
        self.open_tool_window("Edit Tool", data=data, on_save=save_action)

    def delete_tool_type(self):
        selected = self.tool_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a tool to delete.")
            return
        data = self.tool_tree.item(selected[0])["values"]
        if messagebox.askyesno("Confirm", f"Delete tool '{data[1]}'?"):
            self.cursor.execute("DELETE FROM tool_types WHERE id=?", (data[0],))
            self.conn.commit()
            self.load_tool_types()

    def build_tab_ui(self, frame, category):
        btn_frame = ttk.Frame(frame, style="Custom.TFrame")
        btn_frame.pack(fill="x", padx=10, pady=10)

        ttk.Button(
            btn_frame,
            text="+ Add Equipment",
            style="Add.TButton",
            command=lambda: self.add_equipment(category),
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            btn_frame,
            text="✎  Edit",
            style="Edit.TButton",
            command=lambda: self.edit_equipment(category),
        ).pack(side="left", padx=(0, 6))
        ttk.Button(
            btn_frame,
            text="✕  Delete",
            style="Del.TButton",
            command=lambda: self.delete_equipment(category),
        ).pack(side="left")

        ttk.Button(
            btn_frame,
            text="Certificates Folder",
            style="Edit.TButton",
            command=self.choose_certificates_folder,
        ).pack(side="left", padx=(12, 0))

        if not hasattr(self, "cert_folder_labels"):
            self.cert_folder_labels = {}
        cert_text = (
            f"Folder: {self.certificates_folder}"
            if self.certificates_folder
            else "Folder: not selected"
        )
        folder_label = ttk.Label(frame, text=cert_text, style="Custom.TLabel")
        folder_label.pack(fill="x", padx=12, pady=(0, 6), anchor="w")
        self.cert_folder_labels[category] = folder_label

        content_frame = ttk.Frame(frame, style="Custom.TFrame")
        content_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        if category == "external":
            columns = ("ID", "Name", "Serial Number", "Last Calib.", "Next Calib.", "Range")
        else:
            columns = ("ID", "Name", "Serial Number", "Last Calib.", "Next Calib.")

        table_frame = ttk.Frame(content_frame, style="Custom.TFrame")
        table_frame.pack(side="left", fill="both", expand=True)

        tree = ttk.Treeview(table_frame, columns=columns, show="headings", style="Custom.Treeview")

        tree.heading("ID", text="ID")
        tree.heading("Name", text="Equipment Name")
        tree.heading("Serial Number", text="Serial Number")
        tree.heading("Last Calib.", text="Last Calibration")
        tree.heading("Next Calib.", text="Next Calibration")
        if category == "external":
            tree.heading("Range", text="Range")

        tree.column("ID", width=40, anchor="center")
        tree.column("Name", width=200)
        tree.column("Serial Number", width=150)
        tree.column("Last Calib.", width=120, anchor="center")
        tree.column("Next Calib.", width=120, anchor="center")
        if category == "external":
            tree.column("Range", width=140, anchor="center")

        tree.tag_configure("odd", background=self._row_odd)
        tree.tag_configure("even", background=self._row_even)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.bind("<Double-1>", lambda event, c=category: self._on_equipment_double_click(event, c))
        tree.bind("<<TreeviewSelect>>", lambda event, c=category: self._on_equipment_select(c, event))
        tree.pack(fill="both", expand=True, side="left")
        scrollbar.pack(fill="y", side="left")

        preview_frame = ttk.Frame(content_frame, style="Custom.TFrame")
        preview_frame.pack(side="left", fill="y", padx=(12, 0))
        preview_frame.configure(width=420)
        preview_frame.pack_propagate(False)
        preview_title = ttk.Label(preview_frame, text="Certificate", style="Custom.TLabel")
        preview_title.pack(anchor="w", pady=(4, 10))
        preview_name = ttk.Label(preview_frame, text="No equipment selected", style="Custom.TLabel", wraplength=320)
        preview_name.pack(anchor="w", pady=(0, 8))
        preview_serial = ttk.Label(preview_frame, text="Serial number: -", style="Custom.TLabel", wraplength=320)
        preview_serial.pack(anchor="w", pady=(0, 8))
        preview_path = ttk.Label(preview_frame, text="Certificate path: -", style="Custom.TLabel", wraplength=320, justify="left")
        preview_path.pack(anchor="w", pady=(0, 8))
        preview_status = ttk.Label(preview_frame, text="Status: -", style="Custom.TLabel")
        preview_status.pack(anchor="w", pady=(0, 8))
        open_button = ttk.Button(
            preview_frame,
            text="Open Certificate",
            style="Edit.TButton",
            command=lambda c=category: self.open_selected_certificate(c),
            state="disabled",
        )
        open_button.pack(anchor="w", pady=(0, 10))

        self.cert_preview_widgets[category] = {
            "title": preview_name,
            "serial": preview_serial,
            "path": preview_path,
            "status": preview_status,
            "open_button": open_button,
            "current_path": "",
        }
        self.trees[category] = tree

        self.load_data(category)
        self.update_certificate_preview(category)

    def _on_equipment_double_click(self, event, category):
        tree = self.trees[category]
        row_id = tree.identify_row(event.y)
        if not row_id:
            return
        tree.selection_set(row_id)
        tree.focus(row_id)
        self.edit_equipment(category)

    def load_data(self, category):
        tree = self.trees[category]
        for row in tree.get_children():
            tree.delete(row)

        if category == "external":
            self.cursor.execute(
                """
                SELECT id, name, serial_number, last_calibration, next_calibration, measurement_range
                FROM equipment
                WHERE category=?
                """,
                (category,),
            )
        else:
            self.cursor.execute(
                "SELECT id, name, serial_number, last_calibration, next_calibration FROM equipment WHERE category=?",
                (category,),
            )
        for i, row in enumerate(self.cursor.fetchall()):
            tag = "odd" if i % 2 else "even"
            tree.insert("", "end", values=row, tags=(tag,))

        items = tree.get_children()
        if items:
            tree.selection_set(items[0])
            tree.focus(items[0])

        self.update_certificate_preview(category)

    def open_window(self, title, category, data=None, on_save=None):
        win = tk.Toplevel(self.root)
        win.title(title)
        if category == "external":
            win.geometry("420x350")
        else:
            win.geometry("420x300")
        win.configure(bg="#253447")
        win.grab_set()
        win.resizable(False, False)

        ttk.Label(win, text="Equipment Name:", style="Custom.TLabel").grid(
            row=0, column=0, padx=16, pady=8, sticky="w")
        tool_names = self._get_tool_names()
        e_name = ttk.Combobox(win, values=tool_names, width=22, state="readonly")
        e_name.grid(row=0, column=1, padx=16, pady=8)

        ttk.Label(win, text="Serial Number:", style="Custom.TLabel").grid(
            row=1, column=0, padx=16, pady=8, sticky="w")
        e_serial = ttk.Entry(win, style="Custom.TEntry", width=24)
        e_serial.grid(row=1, column=1, padx=16, pady=8)

        ttk.Label(win, text="Last Calibration:", style="Custom.TLabel").grid(
            row=2, column=0, padx=16, pady=8, sticky="w")
        last_frame = ttk.Frame(win, style="Custom.TFrame")
        last_frame.grid(row=2, column=1, padx=16, pady=8, sticky="w")
        e_last = ttk.Entry(last_frame, style="Custom.TEntry", width=20)
        e_last.pack(side="left")
        ttk.Button(
            last_frame,
            text="📅",
            style="Calendar.TButton",
            command=lambda: self.open_calendar_picker(win, e_last),
        ).pack(side="left", padx=(6, 0))

        ttk.Label(win, text="Next Calibration:", style="Custom.TLabel").grid(
            row=3, column=0, padx=16, pady=8, sticky="w")
        next_frame = ttk.Frame(win, style="Custom.TFrame")
        next_frame.grid(row=3, column=1, padx=16, pady=8, sticky="w")
        e_next = ttk.Entry(next_frame, style="Custom.TEntry", width=20)
        e_next.pack(side="left")
        ttk.Button(
            next_frame,
            text="📅",
            style="Calendar.TButton",
            command=lambda: self.open_calendar_picker(win, e_next),
        ).pack(side="left", padx=(6, 0))

        e_range = None
        if category == "external":
            ttk.Label(win, text="Range:", style="Custom.TLabel").grid(
                row=4, column=0, padx=16, pady=8, sticky="w")
            e_range = ttk.Entry(win, style="Custom.TEntry", width=24)
            e_range.grid(row=4, column=1, padx=16, pady=8)

        if data:
            if data[1] in tool_names:
                e_name.set(data[1])
            e_serial.insert(0, data[2])
            e_last.insert(0, data[3])
            e_next.insert(0, data[4])
            if category == "external" and e_range is not None and len(data) > 5:
                e_range.insert(0, data[5] or "")

        def save():
            name = e_name.get().strip()
            if not name:
                messagebox.showwarning("Error", "Please select an equipment name!", parent=win)
                return
            measurement_range = e_range.get().strip() if e_range is not None else ""
            on_save(name, e_serial.get(), e_last.get(), e_next.get(), measurement_range)
            win.destroy()

        ttk.Button(win, text="Save", style="Add.TButton", command=save).grid(
            row=5 if category == "external" else 4, column=0, columnspan=2, pady=18
        )

    def open_calendar_picker(self, parent, target_entry):
        def open_basic_date_picker(selected_date):
            picker = tk.Toplevel(parent)
            picker.title("Select Date")
            picker.configure(bg="#253447")
            picker.resizable(False, False)
            picker.grab_set()

            form = ttk.Frame(picker, style="Custom.TFrame")
            form.pack(padx=12, pady=12)

            ttk.Label(form, text="Day:", style="Custom.TLabel").grid(row=0, column=0, sticky="w", pady=4)
            day_var = tk.IntVar(value=selected_date.day)
            day_spin = tk.Spinbox(form, from_=1, to=31, width=6, textvariable=day_var)
            day_spin.grid(row=0, column=1, padx=(8, 0), pady=4)

            ttk.Label(form, text="Month:", style="Custom.TLabel").grid(row=1, column=0, sticky="w", pady=4)
            month_var = tk.IntVar(value=selected_date.month)
            month_spin = tk.Spinbox(form, from_=1, to=12, width=6, textvariable=month_var)
            month_spin.grid(row=1, column=1, padx=(8, 0), pady=4)

            ttk.Label(form, text="Year:", style="Custom.TLabel").grid(row=2, column=0, sticky="w", pady=4)
            year_var = tk.IntVar(value=selected_date.year)
            year_spin = tk.Spinbox(form, from_=1900, to=2100, width=8, textvariable=year_var)
            year_spin.grid(row=2, column=1, padx=(8, 0), pady=4)

            def apply_date_fallback():
                try:
                    chosen = datetime(year_var.get(), month_var.get(), day_var.get())
                except ValueError:
                    messagebox.showwarning("Invalid date", "Selected date is invalid.", parent=picker)
                    return
                target_entry.delete(0, tk.END)
                target_entry.insert(0, chosen.strftime("%d.%m.%Y"))
                picker.destroy()

            ttk.Button(picker, text="Use Date", style="Add.TButton", command=apply_date_fallback).pack(
                pady=(0, 12)
            )

        current_value = target_entry.get().strip()
        try:
            selected = datetime.strptime(current_value, "%d.%m.%Y")
        except ValueError:
            try:
                selected = datetime.strptime(current_value, "%Y-%m-%d")
            except ValueError:
                try:
                    selected = datetime.strptime(current_value, "%d-%m-%Y")
                except ValueError:
                    selected = datetime.today()

        if self._calendar_class is None:
            try:
                tkcalendar_module = importlib.import_module("tkcalendar")
                self._calendar_class = tkcalendar_module.Calendar
            except ModuleNotFoundError:
                open_basic_date_picker(selected)
                return

        picker = tk.Toplevel(parent)
        picker.title("Select Date")
        picker.configure(bg="#253447")
        picker.resizable(False, False)
        picker.grab_set()

        cal = self._calendar_class(
            picker,
            selectmode="day",
            date_pattern="dd.mm.yyyy",
            year=selected.year,
            month=selected.month,
            day=selected.day,
        )
        cal.pack(padx=12, pady=12)

        def apply_date():
            target_entry.delete(0, tk.END)
            target_entry.insert(0, cal.get_date())
            picker.destroy()

        ttk.Button(picker, text="Use Date", style="Add.TButton", command=apply_date).pack(
            pady=(0, 12)
        )

    def add_equipment(self, category):
        def save_action(name, serial, last_cal, next_cal, measurement_range):
            if not name:
                messagebox.showwarning(
                    "Error", "Equipment name is required!"
                )
                return
            cert_path = self.find_certificate_for_serial(serial) or ""
            self.cursor.execute(
                """
                INSERT INTO equipment (category, name, serial_number, last_calibration, next_calibration, measurement_range, certificate_path)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
                (category, name, serial, last_cal, next_cal, measurement_range, cert_path),
            )
            self.conn.commit()
            self.load_data(category)

        self.open_window("Add Equipment", category, on_save=save_action)

    def edit_equipment(self, category):
        tree = self.trees[category]
        selected = tree.selection()
        if not selected:
            messagebox.showwarning(
                "Warning", "Please select an item from the list to edit."
            )
            return

        item_data = tree.item(selected[0])["values"]
        item_id = item_data[0]

        def save_action(name, serial, last_cal, next_cal, measurement_range):
            cert_path = self.find_certificate_for_serial(serial) or ""
            self.cursor.execute(
                """
                UPDATE equipment 
                SET name=?, serial_number=?, last_calibration=?, next_calibration=?, measurement_range=?, certificate_path=?
                WHERE id=?
            """,
                (name, serial, last_cal, next_cal, measurement_range, cert_path, item_id),
            )
            self.conn.commit()
            self.load_data(category)

        self.open_window(
            "Edit Equipment", category, data=item_data, on_save=save_action
        )

    def delete_equipment(self, category):
        tree = self.trees[category]
        selected = tree.selection()
        if not selected:
            messagebox.showwarning(
                "Warning", "Please select an item from the list to delete."
            )
            return

        item_id = tree.item(selected[0])["values"][0]

        if messagebox.askyesno(
            "Confirm", "Are you sure you want to delete this item?"
        ):
            self.cursor.execute(
                "DELETE FROM equipment WHERE id=?", (item_id,)
            )
            self.conn.commit()
            self.load_data(category)


if __name__ == "__main__":
    root = tk.Tk()
    app = CalibrationApp(root)
    root.mainloop()