import sqlite3
import tkinter as tk
from tkinter import messagebox, ttk


class CalibrationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Measuring Equipment Calibration Register")
        self.root.state("zoomed")
        self.root.configure(bg="#1e2a3a")

        self._setup_styles()
        self.init_db()

        self.categories = {
            "External": "external",
            "Internal": "internal",
            "Gauges": "gauges",
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
                next_calibration TEXT
            )
        """
        )
        self.cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS tool_types (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                description TEXT
            )
        """
        )
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

    def build_tool_db_tab(self, frame):
        btn_frame = ttk.Frame(frame, style="Custom.TFrame")
        btn_frame.pack(fill="x", padx=10, pady=10)
        ttk.Button(btn_frame, text="+ Add Tool", style="Add.TButton",
                   command=self.add_tool_type).pack(side="left", padx=(0, 6))
        ttk.Button(btn_frame, text="Edit", style="Edit.TButton",
                   command=self.edit_tool_type).pack(side="left", padx=(0, 6))
        ttk.Button(btn_frame, text="Delete", style="Del.TButton",
                   command=self.delete_tool_type).pack(side="left")

        columns = ("ID", "Tool Name", "Description")
        self.tool_tree = ttk.Treeview(frame, columns=columns, show="headings", style="Custom.Treeview")
        self.tool_tree.heading("ID", text="ID")
        self.tool_tree.heading("Tool Name", text="Tool Name")
        self.tool_tree.heading("Description", text="Description")
        self.tool_tree.column("ID", width=40, anchor="center")
        self.tool_tree.column("Tool Name", width=220)
        self.tool_tree.column("Description", width=450)
        self.tool_tree.tag_configure("odd", background=self._row_odd)
        self.tool_tree.tag_configure("even", background=self._row_even)
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.tool_tree.yview)
        self.tool_tree.configure(yscrollcommand=sb.set)
        self.tool_tree.pack(fill="both", expand=True, padx=10, pady=(0, 10), side="left")
        sb.pack(fill="y", pady=(0, 10), side="left")
        self.load_tool_types()

    def load_tool_types(self):
        for row in self.tool_tree.get_children():
            self.tool_tree.delete(row)
        self.cursor.execute("SELECT id, name, description FROM tool_types ORDER BY name")
        for i, row in enumerate(self.cursor.fetchall()):
            self.tool_tree.insert("", "end", values=row, tags=("odd" if i % 2 else "even",))

    def _get_tool_names(self):
        self.cursor.execute("SELECT name FROM tool_types ORDER BY name")
        return [r[0] for r in self.cursor.fetchall()]

    def open_tool_window(self, title, data=None, on_save=None):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("420x200")
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
        if data:
            e_name.insert(0, data[1])
            e_desc.insert(0, data[2] or "")

        def save():
            name = e_name.get().strip()
            if not name:
                messagebox.showwarning("Error", "Tool name is required!", parent=win)
                return
            on_save(name, e_desc.get().strip())
            win.destroy()

        ttk.Button(win, text="Save", style="Add.TButton", command=save).grid(
            row=2, column=0, columnspan=2, pady=16)

    def add_tool_type(self):
        def save_action(name, desc):
            try:
                self.cursor.execute(
                    "INSERT INTO tool_types (name, description) VALUES (?, ?)", (name, desc))
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

        def save_action(name, desc):
            try:
                self.cursor.execute(
                    "UPDATE tool_types SET name=?, description=? WHERE id=?",
                    (name, desc, tool_id))
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

        columns = ("ID", "Name", "Serial Number", "Last Calib.", "Next Calib.")
        tree = ttk.Treeview(frame, columns=columns, show="headings", style="Custom.Treeview")

        tree.heading("ID", text="ID")
        tree.heading("Name", text="Equipment Name")
        tree.heading("Serial Number", text="Serial Number")
        tree.heading("Last Calib.", text="Last Calibration")
        tree.heading("Next Calib.", text="Next Calibration")

        tree.column("ID", width=40, anchor="center")
        tree.column("Name", width=200)
        tree.column("Serial Number", width=150)
        tree.column("Last Calib.", width=120, anchor="center")
        tree.column("Next Calib.", width=120, anchor="center")

        tree.tag_configure("odd", background=self._row_odd)
        tree.tag_configure("even", background=self._row_even)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(fill="both", expand=True, padx=10, pady=(0, 10), side="left")
        scrollbar.pack(fill="y", pady=(0, 10), side="left")
        self.trees[category] = tree

        self.load_data(category)

    def load_data(self, category):
        tree = self.trees[category]
        for row in tree.get_children():
            tree.delete(row)

        self.cursor.execute(
            "SELECT id, name, serial_number, last_calibration, next_calibration FROM equipment WHERE category=?",
            (category,),
        )
        for i, row in enumerate(self.cursor.fetchall()):
            tag = "odd" if i % 2 else "even"
            tree.insert("", "end", values=row, tags=(tag,))

    def open_window(self, title, data=None, on_save=None):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("420x300")
        win.configure(bg="#253447")
        win.grab_set()
        win.resizable(False, False)

        ttk.Label(win, text="Equipment Name:", style="Custom.TLabel").grid(
            row=0, column=0, padx=16, pady=8, sticky="w")
        tool_names = self._get_tool_names()
        e_name = ttk.Combobox(win, values=tool_names, width=22, state="readonly")
        e_name.grid(row=0, column=1, padx=16, pady=8)

        labels = ["Serial Number:", "Last Calibration:", "Next Calibration:"]
        entries = []
        for i, label in enumerate(labels, start=1):
            ttk.Label(win, text=label, style="Custom.TLabel").grid(
                row=i, column=0, padx=16, pady=8, sticky="w")
            e = ttk.Entry(win, style="Custom.TEntry", width=24)
            e.grid(row=i, column=1, padx=16, pady=8)
            entries.append(e)
        e_serial, e_last, e_next = entries

        if data:
            if data[1] in tool_names:
                e_name.set(data[1])
            e_serial.insert(0, data[2])
            e_last.insert(0, data[3])
            e_next.insert(0, data[4])

        def save():
            name = e_name.get().strip()
            if not name:
                messagebox.showwarning("Error", "Please select an equipment name!", parent=win)
                return
            on_save(name, e_serial.get(), e_last.get(), e_next.get())
            win.destroy()

        ttk.Button(win, text="Save", style="Add.TButton", command=save).grid(
            row=4, column=0, columnspan=2, pady=18
        )

    def add_equipment(self, category):
        def save_action(name, serial, last_cal, next_cal):
            if not name:
                messagebox.showwarning(
                    "Error", "Equipment name is required!"
                )
                return
            self.cursor.execute(
                """
                INSERT INTO equipment (category, name, serial_number, last_calibration, next_calibration)
                VALUES (?, ?, ?, ?, ?)
            """,
                (category, name, serial, last_cal, next_cal),
            )
            self.conn.commit()
            self.load_data(category)

        self.open_window("Add Equipment", on_save=save_action)

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

        def save_action(name, serial, last_cal, next_cal):
            self.cursor.execute(
                """
                UPDATE equipment 
                SET name=?, serial_number=?, last_calibration=?, next_calibration=?
                WHERE id=?
            """,
                (name, serial, last_cal, next_cal, item_id),
            )
            self.conn.commit()
            self.load_data(category)

        self.open_window(
            "Edit Equipment", data=item_data, on_save=save_action
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