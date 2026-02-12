import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import os
import shutil
import json
import fitz  # PyMuPDF
from PIL import Image, ImageTk

# ===============================
# CONFIG & CONSTANTS
# ===============================
CONFIG_FILE = "config.json"
COLUMNS_FILE = "columns.json"
DEFAULT_CONFIG = {
    "excel_path": "excel/diagram_list.xlsx",
    "pdf_dir": "pdf",
    "language": "Japanese"
}
DEFAULT_COLUMNS = {
    "english": ["Search No","Reference model","Contents","Before correction","After correction","Type 1","Type 2","Type 3","Type 4","Type 5"],
    "japanese": ["検索No.","参考機種","内容","訂正前","訂正後","分類1","分類2","分類3","分類4","分類5"]
}

def load_config():
    if not os.path.exists(CONFIG_FILE):
        return DEFAULT_CONFIG.copy()
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=4, ensure_ascii=False)

def load_columns():
    if not os.path.exists(COLUMNS_FILE):
        with open(COLUMNS_FILE, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_COLUMNS, f, indent=4, ensure_ascii=False)
        return DEFAULT_COLUMNS.copy()
    with open(COLUMNS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_columns(columns):
    with open(COLUMNS_FILE, "w", encoding="utf-8") as f:
        json.dump(columns, f, indent=4, ensure_ascii=False)

config = load_config()
columns_data = load_columns()

EXCEL_PATH = config["excel_path"]
PDF_DIR = config["pdf_dir"]
DEFAULT_LANG = config.get("language", "Japanese")
os.makedirs(PDF_DIR, exist_ok=True)

COLUMNS = columns_data["english"]
JAPANESE_COLUMNS = columns_data["japanese"]

# ===============================
# LANGUAGE TEXT
# ===============================
with open("lang.json", "r", encoding="utf-8") as f:
    LANG_TEXT = json.load(f)

# ===============================
# EXCEL HANDLING
# ===============================
def load_excel():
    if not os.path.exists(EXCEL_PATH):
        return pd.DataFrame(columns=COLUMNS)
    df = pd.read_excel(EXCEL_PATH, dtype=str).fillna("")
    clean_df = pd.DataFrame({col: df[col] if col in df.columns else "" for col in COLUMNS})
    return clean_df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

def save_excel(df):
    os.makedirs(os.path.dirname(EXCEL_PATH), exist_ok=True)
    df[COLUMNS].to_excel(EXCEL_PATH, index=False)

def export_excel(df, lang):
    headers = JAPANESE_COLUMNS if lang=="Japanese" else COLUMNS
    file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
    if file:
        df_out = df.copy()
        df_out.columns = headers
        df_out.to_excel(file, index=False)
        messagebox.showinfo(LANG_TEXT[lang]["export_done"],
                            LANG_TEXT[lang]["export_msg"].format(file=file))

# ===============================
# PDF HANDLING
# ===============================
def find_pdf(search_no):
    search_no_norm = str(search_no).zfill(3)
    for f in os.listdir(PDF_DIR):
        if f.lower().endswith(".pdf") and f"検索no.{search_no_norm}" in f.lower():
            return os.path.join(PDF_DIR, f)
    return None

def generate_pdf_thumbnail(pdf_path, width=700):
    if not pdf_path or not os.path.exists(pdf_path):
        return None
    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    pix = page.get_pixmap(matrix=fitz.Matrix(2,2))
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    ratio = width / img.width
    img = img.resize((int(img.width*ratio), int(img.height*ratio)))
    doc.close()
    return ImageTk.PhotoImage(img)

# ===============================
# MULTI-SELECT DROPDOWN
# ===============================
class MultiSelectDropdown(tk.Frame):

    def __init__(self, parent, values, label="All", width=20, callback=None):
        super().__init__(parent)

        self.values = values
        self.selected = []
        self.callback = callback
        self.width = width
        self.default_label = label

        self.button = tk.Button(
            self,
            text=self.default_label,
            width=self.width,
            command=self.toggle_menu
        )
        self.button.pack()

    def toggle_menu(self):
        menu = tk.Toplevel(self)
        menu.wm_overrideredirect(True)

        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        menu.geometry(f"+{x}+{y}")

        self.vars = {}

        for v in self.values:
            var = tk.BooleanVar(value=v in self.selected)

            chk = tk.Checkbutton(
                menu,
                text=v,
                variable=var,
                command=self.update_selection,
                anchor="w"
            )
            chk.pack(fill="x", padx=5)

            self.vars[v] = var

        menu.bind("<FocusOut>", lambda e: menu.destroy())
        menu.focus_set()

    def update_selection(self):
        self.selected = [
            v for v, var in self.vars.items()
            if var.get()
        ]

        # Update button text
        if not self.selected:
            self.button.config(text=self.default_label)
        elif len(self.selected) == 1:
            self.button.config(text=self.selected[0])
        else:
            self.button.config(text=f"{len(self.selected)} selected")

        if self.callback:
            self.callback()

    def get_selected(self):
        return self.selected

    def clear_selection(self):
        self.selected = []
        self.button.config(text=self.default_label)

        if self.callback:
            self.callback()


# ===============================
# MAIN APP
# ===============================
class DiagramApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.lang = DEFAULT_LANG
        self.text = LANG_TEXT[self.lang]
        self.df = load_excel()
        self.columns_data = columns_data
        self.title(self.text["app_title"])
        self.geometry("1800x900")
        self.create_styles()
        self.create_ui()
        self.update_headers()
        self.refresh_table(self.df)

    def t(self, key):
        return LANG_TEXT[self.lang][key]

    # ---------- Styles ----------
    def create_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview", rowheight=28)

    def stripe_rows(self):
        for i, item in enumerate(self.tree.get_children()):
            current_tags = list(self.tree.item(item, "tags"))
            if "hover" in current_tags:
                continue
            if i % 2 == 0:
                if "even" not in current_tags:
                    current_tags.append("even")
                if "odd" in current_tags:
                    current_tags.remove("odd")
            else:
                if "odd" not in current_tags:
                    current_tags.append("odd")
                if "even" in current_tags:
                    current_tags.remove("even")
            self.tree.item(item, tags=current_tags)

        self.tree.tag_configure("even", background="#d6e6f2")   # soft light blue
        self.tree.tag_configure("odd", background="#eaf1f8")    # very soft blue/white
        self.tree.tag_configure("exists", foreground="#27ae60") # green
        self.tree.tag_configure("missing", foreground="#c0392b")# red
        self.tree.tag_configure("hover", background="#084a72")  # hover effect

    def create_ui(self):
        # ===============================
        # Header
        # ===============================
        header = tk.Frame(self, height=60)  # keep default bg
        header.pack(fill="x")
        header.pack_propagate(False)

        # Title label
        self.title_lbl = tk.Label(
            header,
            text=self.t("app_title") if hasattr(self, "t") else "Document Manager",
            fg="#005f99",  # light blue text
            font=("Segoe UI", 20, "bold")
        )
        self.title_lbl.pack(side="left", padx=25)

        # Three-dot menu
        self.more_btn = tk.Button(
            header,
            text="⋮",
            font=("Segoe UI", 18),
            fg="#005f99",  # light blue text
            activebackground="#66b3ff",
            activeforeground="white",
            bd=0,
            cursor="hand2",
            command=self.show_header_menu
        )
        self.more_btn.pack(side="right", padx=20)

        # ===============================
        # Filter Section
        # ===============================
        self.filter_frame = tk.LabelFrame(
            self,
            text="Filters",
            font=("Segoe UI", 10, "bold"),
            bg="#f0f4f8",
            fg="#005f99",
            bd=2,
            relief="groove",
            labelanchor="n"
        )
        self.filter_frame.pack(fill="x", padx=20, pady=10)
        self.create_filters()

        # ===============================
        # Table Actions (Add Entry button)
        # ===============================
        table_actions = tk.Frame(self, bg="#f0f4f8")
        table_actions.pack(fill="x", padx=20, pady=(0, 5))  # small space above table

        add_btn = tk.Button(
            table_actions,
            text=self.t("add_entry") if hasattr(self, "t") else "Add Entry",
            fg="white",
            bg="#005f99",           # light blue button
            activebackground="#3399cc",
            activeforeground="white",
            font=("Segoe UI", 10, "bold"),
            bd=0,
            padx=10,
            pady=5,
            cursor="hand2",
            command=self.open_add_window
        )
        add_btn.pack(side="right")

        # ===============================
        # Table Section
        # ===============================
        container = tk.Frame(self, bg="#f0f4f8")
        container.pack(fill="both", expand=True, padx=20, pady=10)

        style = ttk.Style()
        style.theme_use("default")

        # Treeview style
        style.configure(
            "Treeview",
            font=("Segoe UI", 11),
            rowheight=32,
            background="#ffffff",
            fieldbackground="#ffffff",
            bordercolor="#c1d4e6",
            borderwidth=1,
            relief="solid"
        )
        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 11, "bold"),
            background="#005f99",
            foreground="white",
            relief="raised",
            borderwidth=1
        )
        style.map(
            "Treeview",
            background=[("selected", "#66b3ff")],
            foreground=[("selected", "white")]
        )

        self.tree = ttk.Treeview(container, columns=COLUMNS + ["PDF"], show="headings")

        # Scrollbars
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # Hover effect
        def on_hover(event):
            row_id = self.tree.identify_row(event.y)
            self.stripe_rows()
            if row_id:
                current_tags = list(self.tree.item(row_id, "tags"))
                if "hover" not in current_tags:
                    current_tags.append("hover")
                    self.tree.item(row_id, tags=current_tags)
        self.tree.bind("<Motion>", on_hover)
        self.tree.bind("<Leave>", lambda e: self.stripe_rows())

        # Double click preview
        self.tree.bind("<Double-1>", self.open_pdf_preview)

        # Right-click context menu
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Edit", command=self.edit_selected_row)
        self.menu.add_command(label="Delete", command=self.delete_selected_row)
        self.tree.bind("<Button-3>", self.show_context_menu)

        # Columns setup
        for col in COLUMNS:
            self.tree.heading(col, text=col, anchor="center")
            self.tree.column(col, anchor="center", width=140, stretch=False)
        self.tree.heading("PDF", text="PDF", anchor="center")
        self.tree.column("PDF", anchor="center", width=100, stretch=False)

        # ===============================
        # Status Bar
        # ===============================
        self.status_bar = tk.Frame(self, height=28)
        self.status_bar.pack(fill="x", side="bottom")
        self.status_bar.pack_propagate(False)

        self.result_label = tk.Label(
            self.status_bar,
            text="0 Results",
            anchor="w",
            fg="#005f99",  # light blue text
            font=("Segoe UI", 10)
        )
        self.result_label.pack(side="left", padx=15)


    # This show the header menu after creating UI
    def show_header_menu(self):
        menu = tk.Menu(self, tearoff=0)

        menu.add_command(label=self.t("add_entry"), command=self.open_add_window)
        menu.add_command(label=self.t("settings"), command=self.open_settings)
        menu.add_separator()
        menu.add_command(label=self.t("export_excel"),
                        command=lambda: export_excel(self.df, self.lang))

        x = self.more_btn.winfo_rootx()
        y = self.more_btn.winfo_rooty() + self.more_btn.winfo_height()

        menu.tk_popup(x, y)


    # ---------- Filters ----------
    def create_filters(self):
        for w in self.filter_frame.winfo_children():
            w.destroy()

        labels = (
            self.columns_data["japanese"]
            if self.lang == "Japanese"
            else self.columns_data["english"]
        )

        # Make filter frame responsive
        self.filter_frame.grid_columnconfigure(0, weight=1)

        container = tk.Frame(self.filter_frame, bg="#f0f4f8")
        container.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        container.grid_columnconfigure(0, weight=1)

        # ==========================
        # FILTERS AREA (WRAPS)
        # ==========================
        filters_container = tk.Frame(container, bg="#f0f4f8")
        filters_container.grid(row=0, column=0, sticky="w")

        row = 0
        col = 0

        # ---------- Search No ----------
        self.search_var = tk.StringVar()
        tk.Label(filters_container, text=labels[0], fg="#005f99", bg="#f0f4f8").grid(row=row, column=col, sticky="w")
        col += 1

        tk.Entry(filters_container, textvariable=self.search_var, width=12, relief="solid", bd=1, highlightbackground="#c1d4e6")\
            .grid(row=row, column=col, padx=5, pady=2)
        self.search_var.trace_add("write", lambda *_: self.apply_filters())
        col += 1

        # ---------- Reference Model ----------
        tk.Label(filters_container, text=labels[1], fg="#005f99", bg="#f0f4f8")\
            .grid(row=row, column=col, padx=(20, 0), sticky="w")
        col += 1

        models = sorted(self.df["Reference model"].dropna().unique()) \
            if "Reference model" in self.df.columns else []

        self.model_filter = MultiSelectDropdown(
            filters_container,
            models,
            width=14,
            callback=self.apply_filters
        )
        self.model_filter.grid(row=row, column=col, padx=5)
        col += 1

        # ---------- Type Filters ----------
        self.type_filters = {}
        for i in range(1, 6):
            type_col = f"Type {i}"
            if type_col not in self.df.columns:
                continue

            # Auto wrap to next row if too wide
            if col > 8:
                row += 1
                col = 0

            tk.Label(filters_container, text=labels[i + 4], fg="#005f99", bg="#f0f4f8")\
                .grid(row=row, column=col, padx=(20 if col > 0 else 0, 0), sticky="w")
            col += 1

            types = sorted(self.df[type_col].dropna().unique())

            msd = MultiSelectDropdown(
                filters_container,
                types,
                width=14,
                callback=self.apply_filters
            )
            msd.grid(row=row, column=col, padx=5)
            col += 1

            self.type_filters[i] = msd

        # ==========================
        # ACTION BUTTONS (RIGHT)
        # ==========================
        action_container = tk.Frame(container, bg="#f0f4f8")
        action_container.grid(row=0, column=1, sticky="e", padx=10)

        # Export Button (light blue, modern style)
        export_btn = tk.Button(
            action_container,
            text=self.t("export_filtered") if hasattr(self, "t") else "Export",
            fg="white",               # text color
            bg="#005f99",             # light blue background
            activebackground="#3399cc",
            activeforeground="white",
            font=("Segoe UI", 10, "bold"),
            bd=0,                     # no border
            padx=10,
            pady=5,
            cursor="hand2",
            command=self.export_filtered
        )
        export_btn.pack(side="right", padx=5)


       # Clear Button (modern gray style)
        clear_btn = tk.Button(
            action_container,
            text=self.t("clear_filters") if hasattr(self, "t") else "Clear",
            fg="white",               # text color
            bg="#888888",             # gray background
            activebackground="#666666",
            activeforeground="white",
            font=("Segoe UI", 10, "bold"),
            bd=0,                     # no border
            padx=10,
            pady=5,
            cursor="hand2",
            command=self.clear_all_filters
        )
        clear_btn.pack(side="right", padx=5)

    def clear_all_filters(self):
        self.search_var.set("")
        if hasattr(self, "model_filter"):
            self.model_filter.clear_selection()

        for msd in getattr(self, "type_filters", {}).values():
            msd.clear_selection()

        self.apply_filters()

    def apply_filters(self):
        df = self.df.copy()
        search = self.search_var.get().strip()
        if search:
            df = df[df["Search No"].astype(str).str.contains(search, na=False)]

        # ---------- Get current selections ----------
        selected_models = self.model_filter.get_selected() if self.model_filter else []

        selected_types = {}
        for i in self.type_filters:
            selected_types[i] = self.type_filters[i].get_selected()

        # ---------- Filter DataFrame ----------
        if selected_models:
            df = df[df["Reference model"].isin(selected_models)]

        for i in self.type_filters:
            if selected_types[i]:
                df = df[df[f"Type {i}"].isin(selected_types[i])]

        # ---------- Update dropdown options dynamically ----------

        # Update Reference Model options based on selected Types
        if self.model_filter:
            if any(selected_types.values()):
                # Only show reference models that match the selected types
                mask = pd.Series([True] * len(self.df))
                for i, vals in selected_types.items():
                    if vals:
                        mask &= self.df[f"Type {i}"].isin(vals)
                models = sorted(self.df[mask]["Reference model"].dropna().unique())
            else:
                models = sorted(self.df["Reference model"].dropna().unique())
            self.model_filter.values = models
            self.model_filter.selected = [m for m in self.model_filter.selected if m in models]

        # Update Type 1-5 options based on selected Reference Models
        for i in range(1, 6):
            type_col = f"Type {i}"
            if type_col in self.df.columns and i in self.type_filters:
                mask = pd.Series([True] * len(self.df))
                if selected_models:
                    mask &= self.df["Reference model"].isin(selected_models)
                for j, vals in selected_types.items():
                    if j != i and vals:
                        mask &= self.df[f"Type {j}"].isin(vals)
                available = sorted(self.df[mask][type_col].dropna().unique())
                self.type_filters[i].values = available
                self.type_filters[i].selected = [t for t in self.type_filters[i].selected if t in available]

        # ---------- Refresh Table ----------
        self.refresh_table(df)

    def export_filtered(self):
        if not hasattr(self, "filtered_df") or self.filtered_df.empty:
            messagebox.showinfo(self.t("error"), self.t("no_data_to_export"))
            return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.filtered_df.to_excel(file_path, index=False)
            messagebox.showinfo(self.t("export_done"), self.t("export_msg").format(file=file_path))

    def update_headers(self):
        self.text = LANG_TEXT[self.lang]

        # Update window title
        self.title(self.text["app_title"])
        self.title_lbl.config(text=self.text["app_title"])

        # Update filter frame title
        self.filter_frame.config(text=self.text["filters"])

        # Reset Treeview columns
        self.tree["columns"] = COLUMNS + ["PDF"]

        # Language-based column headers
        headers = (
            self.columns_data["japanese"]
            if self.lang == "Japanese"
            else self.columns_data["english"]
        )

        # Update headings
        for i, col in enumerate(COLUMNS):
            self.tree.heading(col, text=headers[i])
            self.tree.column(col, width=140, anchor="center")

        # PDF column (fixed)
        self.tree.heading("PDF", text="PDF", anchor="center")
        self.tree.column("PDF", width=100, anchor="center")

    # Initial Load and reload ito....
    def refresh_table(self, df):

        # Sort by numeric Search No
        if "Search No" in df.columns:
            df = df.copy()
            df["Search No"] = pd.to_numeric(df["Search No"], errors="coerce").fillna(0)
            df = df.sort_values("Search No")  # ascending order

        # Clear existing rows
        self.tree.delete(*self.tree.get_children())
        self.filtered_df = df.copy()

        # Update columns dynamically
        headers = COLUMNS + ["PDF"]
        self.tree["columns"] = headers

        # Localized headers
        labels = (
            self.columns_data["japanese"]
            if self.lang == "Japanese"
            else self.columns_data["english"]
        )

        for i, col in enumerate(COLUMNS):
            self.tree.heading(col, text=labels[i], anchor="center")
            self.tree.column(col, width=140, anchor="center")

        self.tree.heading("PDF", text=self.t("pdf_header"), anchor="center")
        self.tree.column("PDF", width=100, anchor="center")

        # Insert rows
        for _, row in df.iterrows():
            pdf = find_pdf(row.get("Search No", ""))
            status = self.t("pdf_exists") if pdf else self.t("pdf_missing")
            tag = "exists" if pdf else "missing"

            self.tree.insert(
                "",
                "end",
                values=[row.get(c, "") for c in COLUMNS] + [status],
                tags=(tag,)
            )

        # Apply row striping
        self.stripe_rows()

        # Update Result Counter
        result_count = len(df)
        total_count = len(self.df)

        if result_count == total_count:
            self.result_label.config(text=f"{result_count} Results")
        else:
            self.result_label.config(text=f"{result_count} of {total_count} Results")

    # ---------- Right-click ----------
    def show_context_menu(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self.tree.selection_set(row_id)
            self.menu.tk_popup(event.x_root, event.y_root)

    def delete_selected_row(self):
        sel = self.tree.selection()
        if not sel:
            return
        values = self.tree.item(sel, "values")
        search_no = values[0]
        if messagebox.askyesno(self.t("delete_title"), self.t("delete_confirm")):
            idx = self.df[self.df["Search No"]==search_no].index
            self.df.drop(idx, inplace=True)
            save_excel(self.df)
            pdf_path = find_pdf(search_no)
            if pdf_path and os.path.exists(pdf_path):
                os.remove(pdf_path)
            self.refresh_table(self.df)

    # ---------- Edit ----------
    def edit_selected_row(self):
        sel = self.tree.selection()
        if not sel:
            return
        values = self.tree.item(sel, "values")
        original_search_no = values[0]

        win = tk.Toplevel(self)
        win.title("Edit Entry")
        win.geometry("900x500")
        win.minsize(700, 400)

        # ---------- Paned Window ----------
        paned = tk.PanedWindow(win, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=5, pady=5)

        # ---------- Left: Scrollable column entries ----------
        left_frame_outer = tk.Frame(paned)
        paned.add(left_frame_outer, stretch="always")

        left_canvas = tk.Canvas(left_frame_outer)
        left_scrollbar = tk.Scrollbar(left_frame_outer, orient="vertical", command=left_canvas.yview)
        left_inner = tk.Frame(left_canvas)

        left_inner.bind(
            "<Configure>",
            lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        )

        left_canvas.create_window((0, 0), window=left_inner, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set, height=400)

        left_canvas.pack(side="left", fill="both", expand=True)
        left_scrollbar.pack(side="right", fill="y")

        labels = self.columns_data["japanese"] if self.lang == "Japanese" else self.columns_data["english"]
        fields = {}

        for i, col in enumerate(COLUMNS):
            tk.Label(left_inner, text=labels[i]).pack(anchor="w", padx=10, pady=(5, 0))
            var = tk.StringVar(value=values[i])
            ent = tk.Entry(left_inner, textvariable=var, width=80)
            ent.pack(fill="x", padx=10, pady=(0, 5))
            fields[col] = var
            if col == "Search No":
                ent.config(state="disabled")

        # ---------- Right: PDF preview ----------
        right_frame_outer = tk.Frame(paned)
        paned.add(right_frame_outer, stretch="never")

        right_canvas = tk.Canvas(right_frame_outer, width=300)
        right_scrollbar = tk.Scrollbar(right_frame_outer, orient="vertical", command=right_canvas.yview)
        right_inner = tk.Frame(right_canvas, padx=20, pady=20)  # padding for nicer look

        right_inner.bind("<Configure>", lambda e: right_canvas.configure(scrollregion=right_canvas.bbox("all")))
        right_canvas.create_window((0, 0), window=right_inner, anchor="nw")
        right_canvas.configure(yscrollcommand=right_scrollbar.set)

        right_canvas.pack(side="left", fill="both", expand=True)
        right_scrollbar.pack(side="right", fill="y")

        # Layout with grid for centering
        right_inner.grid_columnconfigure(0, weight=1)

        pdf_var = tk.StringVar()
        existing_pdf = find_pdf(original_search_no)

        pdf_label = tk.Label(right_inner,
            text=os.path.basename(existing_pdf) if existing_pdf else self.t("no_pdf"),
            fg="green" if existing_pdf else "red"
        )
        pdf_label.grid(row=0, column=0, pady=(0,5), sticky="n")

        preview_label = tk.Label(right_inner)
        preview_label.grid(row=1, column=0, pady=5, sticky="n")

        if existing_pdf:
            thumb = generate_pdf_thumbnail(existing_pdf, width=200)
            if thumb:
                preview_label.config(image=thumb)
                preview_label.image = thumb

        def select_new_pdf():
            p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
            if p:
                pdf_var.set(p)
                pdf_label.config(text=os.path.basename(p), fg="green")
                thumb = generate_pdf_thumbnail(p, width=200)
                if thumb:
                    preview_label.config(image=thumb)
                    preview_label.image = thumb
                messagebox.showinfo(self.t("success"), self.t("pdf_replaced"))

        ttk.Button(right_inner, text=self.t("replace_pdf"), command=select_new_pdf).grid(row=2, column=0, pady=5,padx=100, sticky="n")

        # ---------- Bottom buttons ----------
        btn_frame = tk.Frame(win)
        btn_frame.pack(fill="x", pady=5, padx=5)
        ttk.Button(
            btn_frame, text=self.t("cancel"), command=win.destroy
        ).pack(side="right", padx=(0,10))
        ttk.Button(
            btn_frame, text=self.t("save_changes"),
            command=lambda: self.save_edited_entry(win, fields, pdf_var, original_search_no)
        ).pack(side="right")
        
    def save_edited_entry(self, win, fields, pdf_var, original_search_no):
        idx = self.df[self.df["Search No"] == original_search_no].index
        if idx.empty:
            return

        # Update DataFrame values
        for col in COLUMNS:
            self.df.loc[idx, col] = fields[col].get()

        # Handle PDF replacement
        if pdf_var.get():
            if not os.path.exists(PDF_DIR):
                os.makedirs(PDF_DIR)

            # Remove old PDF if it exists
            old_pdf = find_pdf(original_search_no)
            if old_pdf and os.path.exists(old_pdf):
                os.remove(old_pdf)

            # Normalize Search No
            search_no_norm = str(fields["Search No"].get()).zfill(3)

            # Get Type1 and Type2 values from the entry fields
            type1 = fields.get("Type 1", tk.StringVar(value="")).get().strip()
            type2 = fields.get("Type 2", tk.StringVar(value="")).get().strip()

            # Build filename: 検索No.004_WR18_MOTOR ASSY.pdf
            new_pdf_name = f"検索No.{search_no_norm}_{type1}_{type2}.pdf"
            new_pdf_path = os.path.join(PDF_DIR, new_pdf_name)

            # Copy and rename
            shutil.copy(pdf_var.get(), new_pdf_path)

        # Save Excel and refresh UI
        save_excel(self.df)
        self.update_headers()
        self.create_filters()
        self.refresh_table(self.df)
        win.destroy()

    # ---------- PDF Preview ----------
    def open_pdf_preview(self, event):
        sel = self.tree.selection()
        if not sel: return
        values = self.tree.item(sel, "values")
        pdf = find_pdf(values[0])
        if not pdf:
            messagebox.showerror(self.t("error"), self.t("pdf_not_found"))
            return

        win = tk.Toplevel(self)
        win.title(self.t("pdf_preview"))
        win.geometry("800x600")

        canvas = tk.Canvas(win, bg="white")
        canvas.pack(fill="both", expand=True)

        scrollbar = tk.Scrollbar(win, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        # state variables
        self.zoom_level = 1.0
        self.current_img = None
        self.img_id = None

        def render_image():
            # regenerate thumbnail at current zoom level
            base_width = int(700 * self.zoom_level)
            thumb = generate_pdf_thumbnail(pdf, width=base_width)
            if thumb:
                if self.img_id:
                    canvas.delete(self.img_id)
                self.img_id = canvas.create_image(0, 0, image=thumb, anchor="nw")
                canvas.image = thumb  # keep reference
                canvas.config(scrollregion=canvas.bbox("all"))

        # initial render
        render_image()

        # pan with mouse drag
        canvas.bind("<ButtonPress-1>", lambda e: canvas.scan_mark(e.x, e.y))
        canvas.bind("<B1-Motion>", lambda e: canvas.scan_dragto(e.x, e.y, gain=1))

        # zoom with mouse wheel
        def zoom(event):
            if event.delta > 0:
                self.zoom_level *= 1.2
            else:
                self.zoom_level /= 1.2
            render_image()
        canvas.bind("<MouseWheel>", zoom)

        # buttons frame
        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text=self.t("open_pdf"), command=lambda: os.startfile(pdf)).pack(side="left", padx=5)
        tk.Button(btn_frame, text=self.t("close"), command=win.destroy).pack(side="left", padx=5)

    # ---------- Settings ----------
    def open_settings(self):
        win = tk.Toplevel(self)
        win.title(self.t("settings_title"))
        win.geometry("500x350")
        excel_var = tk.StringVar(value=EXCEL_PATH)
        pdf_var = tk.StringVar(value=PDF_DIR)
        lang_var = tk.StringVar(value=self.lang)

        def update_text(*args):
            lang = lang_var.get()
            excel_browse_btn.config(text=LANG_TEXT[lang]["browse_excel"])
            pdf_browse_btn.config(text=LANG_TEXT[lang]["browse_pdf"])
            excel_label.config(text=LANG_TEXT[lang]["excel_file"])
            pdf_label.config(text=LANG_TEXT[lang]["pdf_folder"])
            lang_label.config(text=LANG_TEXT[lang]["default_language"])
            save_btn.config(text=LANG_TEXT[lang]["save_settings"])
            columns_btn.config(text=LANG_TEXT[lang]["manage_columns"])
            win.title(LANG_TEXT[lang]["settings_title"])

        # Excel
        excel_label = tk.Label(win, text=LANG_TEXT[lang_var.get()]["excel_file"])
        excel_label.pack(anchor="w", padx=10)
        excel_frame = tk.Frame(win)
        excel_frame.pack(fill="x", padx=10)
        tk.Entry(excel_frame, textvariable=excel_var).pack(side="left", fill="x", expand=True)
        excel_browse_btn = tk.Button(
            excel_frame,
            text=LANG_TEXT[lang_var.get()]["browse_excel"],
            command=lambda: excel_var.set(filedialog.askopenfilename(
                title=LANG_TEXT[lang_var.get()]["browse_excel"],
                filetypes=[("Excel files", "*.xlsx *.xls")]
            ))
        )
        excel_browse_btn.pack(side="left", padx=5)

        # PDF folder
        pdf_label = tk.Label(win, text=LANG_TEXT[lang_var.get()]["pdf_folder"])
        pdf_label.pack(anchor="w", padx=10, pady=(10,0))
        pdf_frame = tk.Frame(win)
        pdf_frame.pack(fill="x", padx=10)
        tk.Entry(pdf_frame, textvariable=pdf_var).pack(side="left", fill="x", expand=True)
        pdf_browse_btn = tk.Button(
            pdf_frame,
            text=LANG_TEXT[lang_var.get()]["browse_pdf"],
            command=lambda: pdf_var.set(filedialog.askdirectory(
                title=LANG_TEXT[lang_var.get()]["browse_pdf"]
            ))
        )
        pdf_browse_btn.pack(side="left", padx=5)

        # Language
        lang_label = tk.Label(win, text=LANG_TEXT[lang_var.get()]["default_language"])
        lang_label.pack(anchor="w", padx=10, pady=(10,0))
        lang_combobox = ttk.Combobox(win, textvariable=lang_var, values=list(LANG_TEXT.keys()))
        lang_combobox.pack(anchor="w", padx=10)

        # Manage columns
        columns_btn = ttk.Button(win, text=LANG_TEXT[lang_var.get()]["manage_columns"], command=self.manage_columns)
        columns_btn.pack(pady=5)

        # Save
        save_btn = ttk.Button(win, text=LANG_TEXT[lang_var.get()]["save_settings"],
            command=lambda:self.save_settings(win, excel_var, pdf_var, lang_var))
        save_btn.pack(pady=20)
        lang_var.trace_add("write", update_text)

    def manage_columns(self):
        win = tk.Toplevel(self)
        win.title(self.t("manage_columns"))
        win.geometry("400x400")

        listbox = tk.Listbox(win)
        listbox.pack(fill="both", expand=True, padx=10, pady=10)
        # Show headers in the current language
        headers_to_show = JAPANESE_COLUMNS if self.lang == "Japanese" else COLUMNS
        for header in headers_to_show:
            listbox.insert("end", header)

        def add_column():
            eng_name = simpledialog.askstring(self.t("add_column"), self.t("enter_column_name_english"))
            if not eng_name:
                messagebox.showerror(self.t("error"), self.t("invalid_column_name"))
                return

            jpn_name = simpledialog.askstring(self.t("add_column"), self.t("enter_column_name_japanese"))
            if not jpn_name:
                messagebox.showerror(self.t("error"), self.t("invalid_column_name"))
                return

            # Check if column already exists
            if eng_name in COLUMNS or jpn_name in JAPANESE_COLUMNS:
                messagebox.showerror(self.t("error"), self.t("column_exists").format(col=eng_name))
                return

            try:
                COLUMNS.append(eng_name)
                JAPANESE_COLUMNS.append(jpn_name)
                self.df[eng_name] = ""

                listbox.insert("end", eng_name)

                save_columns({"english": COLUMNS, "japanese": JAPANESE_COLUMNS})
                save_excel(self.df)

                self.update_headers()
                self.refresh_table(self.df)

                messagebox.showinfo(self.t("success"), self.t("column_added").format(col=eng_name))
            except Exception as e:
                messagebox.showerror(self.t("error"), f"{self.t('add_failed')}: {e}")

        def remove_column():
            sel = listbox.curselection()
            if not sel:
                messagebox.showerror(self.t("error"), self.t("no_selection"))
                return

            col = listbox.get(sel)
            if col in ["Search No", "Reference model"]:
                messagebox.showerror(self.t("error"), self.t("cannot_remove"))
                return

            try:
                idx = COLUMNS.index(col)
                del COLUMNS[idx]
                del JAPANESE_COLUMNS[idx]
                self.df.drop(columns=[col], inplace=True)

                listbox.delete(sel)

                save_columns({"english": COLUMNS, "japanese": JAPANESE_COLUMNS})
                save_excel(self.df)

                self.update_headers()
                self.refresh_table(self.df)

                messagebox.showinfo(self.t("success"), self.t("column_removed").format(col=col))
            except Exception as e:
                messagebox.showerror(self.t("error"), f"{self.t('remove_failed')}: {e}")

        tk.Button(win, text=self.t("add_column"), command=add_column).pack(side="left", padx=10, pady=5)
        tk.Button(win, text=self.t("remove_column"), command=remove_column).pack(side="right", padx=10, pady=5)

    def save_settings(self, win, excel_var, pdf_var, lang_var):
        global EXCEL_PATH, PDF_DIR
        EXCEL_PATH = excel_var.get()
        PDF_DIR = pdf_var.get()
        self.lang = lang_var.get()
        save_config({"excel_path":EXCEL_PATH,"pdf_dir":PDF_DIR,"language":self.lang})
        self.df = load_excel()
        self.update_headers()
        self.refresh_table(self.df)
        self.create_filters()
        win.destroy()

    # ---------- Add Entry ----------
    def open_add_window(self):
        win = tk.Toplevel(self)
        win.title(self.t("add_title"))
        win.geometry("900x500")
        win.minsize(700, 400)

        # ---------- Paned Window ----------
        paned = tk.PanedWindow(win, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=5, pady=5)

        # ---------- Left: Scrollable column entries ----------
        left_frame_outer = tk.Frame(paned)
        paned.add(left_frame_outer, stretch="always")

        left_canvas = tk.Canvas(left_frame_outer)
        left_scrollbar = tk.Scrollbar(left_frame_outer, orient="vertical", command=left_canvas.yview)
        left_inner = tk.Frame(left_canvas)

        left_inner.bind(
            "<Configure>",
            lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        )

        left_canvas.create_window((0, 0), window=left_inner, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)

        left_canvas.pack(side="left", fill="both", expand=True)
        left_scrollbar.pack(side="right", fill="y")

        labels = self.columns_data["japanese"] if self.lang == "Japanese" else self.columns_data["english"]
        fields = {}

        for i, col in enumerate(COLUMNS):
            tk.Label(left_inner, text=labels[i]).pack(anchor="w", padx=10, pady=(5, 0))
            var = tk.StringVar()
            ent = tk.Entry(left_inner, textvariable=var, width=80)
            ent.pack(fill="x", padx=10, pady=(0, 5))
            fields[col] = var

        # ---------- Right: PDF preview ----------
        right_frame_outer = tk.Frame(paned)
        paned.add(right_frame_outer, stretch="never")

        right_canvas = tk.Canvas(right_frame_outer, width=300)
        right_scrollbar = tk.Scrollbar(right_frame_outer, orient="vertical", command=right_canvas.yview)
        right_inner = tk.Frame(right_canvas, padx=20, pady=20)

        right_inner.bind(
            "<Configure>",
            lambda e: right_canvas.configure(scrollregion=right_canvas.bbox("all"))
        )

        right_canvas.create_window((0, 0), window=right_inner, anchor="nw")
        right_canvas.configure(yscrollcommand=right_scrollbar.set)

        right_canvas.pack(side="left", fill="both", expand=True)
        right_scrollbar.pack(side="right", fill="y")

        right_inner.grid_columnconfigure(0, weight=1)

        pdf_var = tk.StringVar()

        pdf_label = tk.Label(
            right_inner,
            text=self.t("no_pdf"),
            fg="red"
        )
        pdf_label.grid(row=0, column=0, pady=(0, 5), sticky="n")

        preview_label = tk.Label(right_inner)
        preview_label.grid(row=1, column=0, pady=5, sticky="n")

        def select_pdf():
            p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
            if p:
                pdf_var.set(p)
                pdf_label.config(text=os.path.basename(p), fg="green")

                thumb = generate_pdf_thumbnail(p, width=200)
                if thumb:
                    preview_label.config(image=thumb)
                    preview_label.image = thumb

        ttk.Button(
            right_inner,
            text=self.t("select_pdf"),
            command=select_pdf
        ).grid(row=2, column=0, pady=5, padx=100, sticky="n")

        # ---------- Bottom buttons ----------
        btn_frame = tk.Frame(win)
        btn_frame.pack(fill="x", pady=5, padx=5)

        ttk.Button(
            btn_frame,
            text=self.t("cancel"),
            command=win.destroy
        ).pack(side="right", padx=(0, 10))

        ttk.Button(
            btn_frame,
            text=self.t("save_entry"),
            command=lambda: self.save_entry(win, fields, pdf_var)
        ).pack(side="right")


    def save_entry(self, win, fields, pdf_var):
        if not fields["Search No"].get() or not pdf_var.get():
            messagebox.showerror(self.t("error"), self.t("required_error"))
            return

        if not os.path.exists(PDF_DIR):
            os.makedirs(PDF_DIR)
        shutil.copy(pdf_var.get(), PDF_DIR)

        # Check duplicates
        duplicate = self.df[
            (self.df["Search No"] == fields["Search No"].get()) &
            (self.df["Reference model"] == fields["Reference model"].get())
        ]
        if not duplicate.empty:
            messagebox.showerror(self.t("error"), self.t("duplicate_error"))
            return

        # Add new entry
        self.df = pd.concat([self.df, pd.DataFrame([{c: fields[c].get() for c in COLUMNS}])], ignore_index=True)

        # -------------------------------
        # Sort by numeric Search No before saving
        # -------------------------------
        self.df["Search No"] = pd.to_numeric(self.df["Search No"], errors="coerce").fillna(0)
        self.df = self.df.sort_values("Search No").reset_index(drop=True)

        # Save Excel
        save_excel(self.df)

        # Update UI
        self.update_headers()
        self.create_filters()
        self.refresh_table(self.df)

        win.destroy()


# ===============================
# RUN
# ===============================
if __name__ == "__main__":
    DiagramApp().mainloop()