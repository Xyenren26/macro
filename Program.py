import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import win32com.client  # Requires pywin32 installed (Windows only)
import webbrowser
import pandas as pd
import os
import sys
import getpass
import subprocess
import time
import shutil
import json
import fitz  # PyMuPDF
from PIL import Image, ImageTk
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Font
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import threading
import zipfile

# ===============================
# CONFIG & CONSTANTS
# ===============================
CONFIG_FILE = "config.json"

DEFAULT_CONFIG = {
    "excel_path": "excel/diagram_list.xlsx",
    "pdf_dir": "pdf",
    "language": "English"
}
DEFAULT_COLUMNS = {
    "english": ["Search No","Reference model","Contents","Before correction","After correction",  "Verification Meeting – Implementation Period",
                "Record No.", "Model Name","Target Part Name","Motor Specification","Issue Classification","Update Info",
                "Updated By","Upload Date"],
    "japanese": ["検索No.","参考機種","内容","訂正前","訂正後", "検証会_実施期","記録No.",
                 "機種名","対象部品名","モータ仕様","指摘項目の分類","更新情報",
                 "更新者","アップロード日"]
}

username = getpass.getuser()

upload_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def load_config():
    # Try reading the config file first
    if not os.path.exists(CONFIG_FILE):
        return DEFAULT_CONFIG.copy()
    
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config = json.load(f)
    except (json.JSONDecodeError, OSError):
        show_config_warning("Configuration file is corrupted or unreadable. Using default settings.")
        return DEFAULT_CONFIG.copy()

    # Now validate network paths
    excel_path = config.get("excel_path")
    pdf_dir = config.get("pdf_dir")

    paths_ok = True
    missing_paths = []

    if excel_path and not os.path.exists(excel_path):
        paths_ok = False
        missing_paths.append(f"Excel file: {excel_path}")

    if pdf_dir and not os.path.isdir(pdf_dir):
        paths_ok = False
        missing_paths.append(f"PDF directory: {pdf_dir}")

    if not paths_ok:
        missing_str = "\n".join(missing_paths)
        show_config_warning(f"The following paths are unreachable:\n{missing_str}\nUsing default settings instead.")
        return DEFAULT_CONFIG.copy()

    return config

def show_config_warning(msg):
    try:
        with open("lang.json", "r", encoding="utf-8") as lf:
            lang_text = json.load(lf)
        warning_title = lang_text.get("config_warning", "Warning")
        warning_msg = lang_text.get("config_msg", msg)
    except Exception:
        warning_title = "Warning"
        warning_msg = msg

    messagebox.showwarning(warning_title, warning_msg)
    
def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=4, ensure_ascii=False)

def load_columns():
    if not os.path.exists(COLUMNS_FILE):
        with open(COLUMNS_FILE, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_COLUMNS, f, indent=4, ensure_ascii=False)
        return DEFAULT_COLUMNS.copy()

    try:
        with open(COLUMNS_FILE, "r", encoding="utf-8") as f:
            columns = json.load(f)

            # Merge with defaults so missing keys don’t break things
            merged = DEFAULT_COLUMNS.copy()
            merged.update(columns)
            return merged
    except (json.JSONDecodeError, OSError) as e:
        try:
            with open("lang.json", "r", encoding="utf-8") as lf:
                lang_text = json.load(lf)
            warning_title = lang_text.get("columns_warning", "Warning")
            warning_msg = lang_text.get("columns_msg", "Columns file error — using defaults")
        except Exception:
            warning_title = "Warning"
            warning_msg = "Columns file error — using defaults"

        messagebox.showwarning(warning_title, warning_msg)

        # Reset file to defaults so next run is clean
        with open(COLUMNS_FILE, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_COLUMNS, f, indent=4, ensure_ascii=False)

        return DEFAULT_COLUMNS.copy()

def save_columns(columns):
    with open(COLUMNS_FILE, "w", encoding="utf-8") as f:
        json.dump(columns, f, indent=4, ensure_ascii=False)

config = load_config()
EXCEL_PATH = config["excel_path"]
PDF_DIR = config["pdf_dir"]

# Build columns.json path one level up from the Excel folder
excel_dir = os.path.dirname(EXCEL_PATH)       # .../excel/
parent_dir = os.path.dirname(excel_dir)       # go outside the excel folder
COLUMNS_FILE = os.path.join(parent_dir, "columns.json")

columns_data = load_columns()

DEFAULT_LANG = config.get("language", "Japanese")
os.makedirs(PDF_DIR, exist_ok=True)

COLUMNS = columns_data["english"]
JAPANESE_COLUMNS = columns_data["japanese"]

# ===============================
# LANGUAGE TEXT
# ===============================
LANG_FILE = os.path.join(parent_dir, "lang.json")
with open(LANG_FILE, "r", encoding="utf-8") as f:
    LANG_TEXT = json.load(f)

# ===============================
# LOAD JSON DROPDOWN
# ===============================
DROPDOWN_FILE = os.path.join(parent_dir, "dropdowns.json")
with open(DROPDOWN_FILE, "r", encoding="utf-8") as f: 
    DROPDOWN_OPTIONS = json.load(f)

# ===============================
# EXCEL HANDLING
# ===============================
def load_excel():
    if not os.path.exists(EXCEL_PATH):
        return pd.DataFrame(columns=COLUMNS)
    df = pd.read_excel(EXCEL_PATH, dtype=str).fillna("")
    clean_df = pd.DataFrame({col: df[col] if col in df.columns else "" for col in COLUMNS})
    return clean_df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

def safe_load_excel():
    for attempt in range(5):
        try:
            df = pd.read_excel(EXCEL_PATH, dtype=str, engine="openpyxl").fillna("")
            clean_df = pd.DataFrame({col: df[col] if col in df.columns else "" for col in COLUMNS})
            return clean_df.map(lambda x: x.strip() if isinstance(x, str) else x)
        except (PermissionError, ValueError, OSError, zipfile.BadZipFile) as e:
            # print(f"Retrying load_excel (attempt {attempt+1}) due to: {e}")
            time.sleep(0.5)
    # print("Failed to load Excel after retries")
    return pd.DataFrame(columns=COLUMNS)


def save_excel(df):
    os.makedirs(os.path.dirname(EXCEL_PATH), exist_ok=True)
    df[COLUMNS].to_excel(EXCEL_PATH, index=False)

def export_excel(df, lang, master_list):
    headers = JAPANESE_COLUMNS if lang == "Japanese" else COLUMNS

    # Build default filename: YYYY-MM-DD_AppTitle.xlsx 
    today = datetime.now().strftime("%Y-%m-%d")
    default_name = f"{today}_{master_list}.xlsx"

    file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile=default_name,
        filetypes=[("Excel","*.xlsx")]
    )
    if not file:
        return

    wb = Workbook()
    ws = wb.active

    # Write headers (include PDF column)
    ws.append(headers + ["PDF"])

    # Write data rows
    for _, row in df.iterrows():
        values = [row.get(c, "") for c in COLUMNS]
        pdf_path = find_pdf(row.get("Search No", ""))
        if pdf_path:
            pdf_name = os.path.basename(pdf_path)
            ws.append(values + [pdf_name])
            cell = ws.cell(row=ws.max_row, column=len(values) + 1)
            try:
                rel_path = os.path.relpath(pdf_path, os.path.dirname(file))
                cell.hyperlink = rel_path
            except ValueError:
                cell.hyperlink = pdf_path
            cell.font = Font(color="0000FF", underline="single")
        else:
            ws.append(values + ["Missing"])

    # Define table range (from A1 to last row/column)
    end_col = get_column_letter(ws.max_column)
    end_row = ws.max_row
    table_ref = f"A1:{end_col}{end_row}"

    # Create table with Dark Teal, Table Style Medium 2
    table = Table(displayName="ExportTable", ref=table_ref)
    style = TableStyleInfo(
        name="TableStyleMedium2",  # Dark Teal style
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )
    table.tableStyleInfo = style
    ws.add_table(table)

    # Auto-fit column widths
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

    # Save file
    wb.save(file)

    # Auto-open the file
    if sys.platform.startswith("darwin"):  # macOS
        subprocess.call(["open", file])
    elif os.name == "nt":  # Windows
        os.startfile(file)
    elif os.name == "posix":  # Linux
        subprocess.call(["xdg-open", file])

    messagebox.showinfo(
        LANG_TEXT[lang]["export_done"],
        LANG_TEXT[lang]["export_msg"].format(file=file)
    )


# ===============================
# EXCEL LOCK FUNCTION (For Saving don't remove - echo)
# ===============================
LOCK_FILE = EXCEL_PATH + ".lock"

def acquire_lock(retries=10, delay=1):
    for attempt in range(retries):
        try:
            if not os.path.exists(LOCK_FILE):
                # Create lock file atomically
                with open(LOCK_FILE, "x") as f:  # "x" fails if file exists
                    f.write("locked")
                return True
        except FileExistsError:
            # Another user is saving → show popup
            messagebox.showinfo(
                LANG_TEXT[DEFAULT_LANG]["lock_title"],
                LANG_TEXT[DEFAULT_LANG]["lock_msg"]
            )
        time.sleep(delay)
    return False

def release_lock():
    try:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)
    except Exception:
        pass  # Ensure we don’t crash if lock file already gone

def save_excel_with_lock(df, path=EXCEL_PATH, retries=5, delay=1):
    if not acquire_lock(retries, delay):
        raise Exception("Could not acquire lock for saving Excel file")

    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)

        # Refresh latest Excel after acquiring lock
        if os.path.exists(path):
            latest_df = pd.read_excel(path, dtype=str).fillna("")
        else:
            latest_df = pd.DataFrame(columns=COLUMNS)

        for attempt in range(retries):
            try:
                df[COLUMNS].to_excel(path, index=False)
                return True
            except PermissionError:
                time.sleep(delay)
        raise Exception("Could not save Excel file after multiple retries")
    finally:
        release_lock()


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
# WATCHDOG IMPLEMENTATION USED TO MAKE THE TREE DATA LIVE
# ===============================
def load_columns_json(path=COLUMNS_FILE):
    import json, os
    if not os.path.exists(path):
        # return default structure if missing
        return {
            "english": [],
            "japanese": [],
            "visible": {}
        }
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    # ensure keys exist
    if "english" not in data:
        data["english"] = []
    if "japanese" not in data:
        data["japanese"] = []
    if "visible" not in data:
        data["visible"] = {}
    return data

class ExcelHandler(FileSystemEventHandler):
    def __init__(self, filepath, app):
        self.filepath = os.path.abspath(filepath)
        self.app = app
        self.last_update = 0
        self.popup_scheduled = False  # prevent multiple popups

    def on_any_event(self, event):
        if event.is_directory:
            return

        filename = os.path.basename(event.src_path)
        if filename.startswith("~$") or filename.endswith(".tmp"):
            return

        if os.path.normcase(os.path.abspath(event.src_path)) == os.path.normcase(self.filepath):
            now = time.time()
            if now - self.last_update < 2:  # debounce
                return
            self.last_update = now

            try:
                new_df = safe_load_excel()
                self.app.columns_data = load_columns_json()
                global COLUMNS, JAPANESE_COLUMNS
                COLUMNS = self.app.columns_data["english"]
                JAPANESE_COLUMNS = self.app.columns_data["japanese"]
                self.app.columns_visibility = self.app.columns_data.get(
                    "visible", {col: True for col in self.app.columns_data["english"]}
                )

                def update_ui():
                    self.app.update_headers()
                    self.app.refresh_table(new_df)
                    if not self.popup_scheduled:
                        self.popup_scheduled = True
                        messagebox.showinfo(
                            LANG_TEXT[self.app.lang].get("excel_update_title", "Info"),
                            LANG_TEXT[self.app.lang].get("excel_update_msg", "Excel file updated and headers refreshed.")
                        )
                        self.popup_scheduled = False  # reset after closing

                self.app.after(0, update_ui)

                # --- Update dropdowns.json ---
                try:
                    with open(DROPDOWN_FILE, "r", encoding="utf-8") as f:
                        dropdowns = json.load(f)
                except FileNotFoundError:
                    dropdowns = {}
                global DROPDOWN_OPTIONS
                DROPDOWN_OPTIONS = dropdowns

            except PermissionError:
                warning_msg = LANG_TEXT[self.app.lang].get(
                    "excel_lock_warning",
                    "Excel is currently open/locked — syncing will resume shortly."
                )
                if not self.popup_scheduled:
                    self.popup_scheduled = True
                    self.app.after(0, lambda msg=warning_msg: (
                        messagebox.showinfo(LANG_TEXT[self.app.lang].get("excel_warning_title", "Warning"), msg),
                        setattr(self, "popup_scheduled", False)
                    ))
                time.sleep(1)

            except Exception as e:
                print(f"[Watchdog] Failed to sync Excel: {e}")
                    
# ===============================
# MULTI-SELECT DROPDOWN (White Theme + Hover)
# ===============================
class MultiSelectDropdown(tk.Frame):
    def __init__(self, parent, values, lang_text, lang="English", width=20, callback=None):
        super().__init__(parent)

        self.values = values
        self.selected = []
        self.callback = callback
        self.width = width
        self.lang_text = lang_text
        self.lang = lang

        # Localized labels
        self.default_label = self.get_text("all_label", default="All")
        self.selected_label = self.get_text("selected_label", default="selected")

        # White button with hover
        self.button = tk.Button(
            self,
            text=self.default_label,
            width=self.width,
            fg="#333333",
            bg="white",
            activebackground="#f0f0f0",
            activeforeground="#007acc",
            font=("Segoe UI", 9, "bold"),
            bd=1,
            relief="solid",
            cursor="hand2",
            command=self.toggle_menu
        )
        self.button.pack(padx=3, pady=3)

        # Hover effect
        self.button.bind("<Enter>", lambda e: self.button.configure(bg="#f9f9f9"))
        self.button.bind("<Leave>", lambda e: self.button.configure(bg="white"))

    def get_text(self, key, default=""):
        """Fetch localized text from lang.json"""
        return self.lang_text.get(self.lang, {}).get(key, default)

    def toggle_menu(self):
        menu = tk.Toplevel(self)
        menu.wm_overrideredirect(True)

        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        menu.geometry(f"+{x}+{y}")

        menu.configure(bg="white", bd=1, relief="solid")

        self.vars = {}

        for v in self.values:
            var = tk.BooleanVar(value=v in self.selected)

            chk = tk.Checkbutton(
                menu,
                text=v,
                variable=var,
                command=self.update_selection,
                anchor="w",
                bg="white",
                fg="#333333",
                activebackground="#f0f0f0",
                activeforeground="#007acc",
                selectcolor="#f0f8ff",
                font=("Segoe UI", 9)
            )
            chk.pack(fill="x", padx=5, pady=2)

            # Hover effect for each option
            chk.bind("<Enter>", lambda e, c=chk: c.configure(bg="#f9f9f9"))
            chk.bind("<Leave>", lambda e, c=chk: c.configure(bg="white"))

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
            self.button.config(text=f"{len(self.selected)} {self.selected_label}")

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

        # --- Language and text resources ---
        self.lang = DEFAULT_LANG
        self.text = LANG_TEXT[self.lang]

        # --- Data ---
        self.df = load_excel()
        self.columns_data = columns_data

        # --- Column visibility (default all True if not in JSON) ---
        if "visible" in self.columns_data:
            self.columns_visibility = self.columns_data["visible"]
        else:
            self.columns_visibility = {col: True for col in self.columns_data["english"]}

        # --- Window setup ---
        self.title(self.text["app_title"])
        self.geometry("1800x900")

        # Make window start maximized with control buttons
        self.state("zoomed")


        # --- Styles and UI ---
        self.create_styles()
        self.create_ui()

        # --- Table setup ---
        self.update_headers()
        self.refresh_table(self.df)
        self.start_excel_watcher(EXCEL_PATH)

    def t(self, key):
        return LANG_TEXT[self.lang][key]

    # ---------- WATCHDOG Don't Touch ----------
    def start_excel_watcher(self, filepath):
        handler = ExcelHandler(filepath, self)
        observer = Observer()
        watch_dir = os.path.dirname(filepath) or "."  # ensure valid directory
        observer.schedule(handler, path=watch_dir, recursive=False)
        observer.start()
        threading.Thread(target=lambda: observer.join(), daemon=True).start()

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

        # Background striping only
        self.tree.tag_configure("even", background="#d6e6f2")   # soft light blue
        self.tree.tag_configure("odd", background="#eaf1f8")    # very soft blue/white
        self.tree.tag_configure("hover", background="#004274")  # hover effect

    def create_ui(self):
        # ===============================
        # Header
        # ===============================
        header = tk.Frame(self, height=60)
        header.pack(fill="x")
        header.pack_propagate(False)


        self.title_lbl = tk.Label(
            header,
            text=self.t("app_title") if hasattr(self, "t") else "Document Manager",
            fg="#005f99",
            font=("Segoe UI", 20, "bold")
        )
        self.title_lbl.pack(side="left", padx=20)

        self.settings_btn = tk.Button(
            header,
            text=self.t("settings") if hasattr(self, "t") else "Setting",
            fg="white",
            bg="#005f99",
            activebackground="#005f99",
            activeforeground="white",
            font=("Segoe UI", 10),
            padx=10,
            pady=5,
            bd=1,
            relief="solid",
            cursor="hand2",
            command=self.open_settings
        )

        self.export_btn = tk.Button(
            header,
            text=self.t("export_excel") if hasattr(self, "t") else "Export Excel",
            fg="white",
            bg="#005f99",
            activebackground="#005f99",
            activeforeground="white",
            font=("Segoe UI", 10),
            padx=10,
            pady=5,
            bd=1,
            relief="solid",
            cursor="hand2",
            command=lambda: export_excel(self.df, self.lang, self.t("master_list"))
        )
        self.export_btn.pack(side="right", padx=10)
        self.settings_btn.pack(side="right", padx=10)

        # ===============================
        # Hyperlink just below buttons
        # ===============================
        link_frame = tk.Frame(self)
        link_frame.pack(fill="x", padx=25, pady=(0, 5))  # directly below header

        link = tk.Label(
            link_frame,
            text="4D2図面検証会_実施記録",
            fg="blue",
            cursor="hand2",
            font=("Segoe UI", 15)
        )
        link.pack(side="right")  # align left under the buttons

        def open_link(event):
            webbrowser.open("https://mitsuba.box.com/s/gumfdt0ie1l0df5st8c3fk8gkv06rm84")

        link.bind("<Button-1>", open_link)

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
        # Table Actions
        # ===============================
        table_actions = tk.Frame(self, bg="#f0f4f8")
        table_actions.pack(fill="x", padx=20, pady=(0, 5))

        self.export_filtered_btn = tk.Button(
            table_actions,
            text=self.t("export_filtered") if hasattr(self, "t") else "Export",
            fg="white",
            bg="#005f99",
            activebackground="#3399cc",
            activeforeground="white",
            font=("Segoe UI", 10, "bold"),
            padx=10,
            pady=5,
            bd=1,
            relief="solid",
            cursor="hand2",
            command=self.export_filtered
        )
        self.export_filtered_btn.pack(side="right", padx=5)

        def add_hover_effect(widget, normal_bg="#005f99", hover_bg="#3399cc", fg="white"):
            widget.bind("<Enter>", lambda e: widget.configure(bg=hover_bg, fg=fg))
            widget.bind("<Leave>", lambda e: widget.configure(bg=normal_bg, fg=fg))

        # Apply to both buttons
        add_hover_effect(self.settings_btn)
        add_hover_effect(self.export_btn)
        add_hover_effect(self.export_filtered_btn)

        # ===============================
        # Table Section
        # ===============================
        container = tk.Frame(self, bg="#f0f4f8")
        container.pack(fill="both", expand=True, padx=20, pady=10)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", font=("Segoe UI", 11), rowheight=32,
                        background="#ffffff", fieldbackground="#ffffff")
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"),
                        background="#005f99", foreground="white")

        style.map("Treeview.Heading", 
                background=[("active", "#004274"), ("pressed", "#3399cc")], 
                foreground=[("active", "white"), ("pressed", "white")] )
        self.tree = ttk.Treeview(container, columns=COLUMNS + ["PDF"], show="headings")

        # Scrollbars
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=1, column=0, sticky="nsew")
        vsb.grid(row=1, column=1, sticky="ns")
        hsb.grid(row=2, column=0, sticky="ew")
        container.grid_rowconfigure(1, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # Double click preview
        self.tree.bind("<Double-1>", self.open_pdf_preview)

        # Right-click context menu
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(
            label=LANG_TEXT[self.lang].get("edit", "Edit"),
            command=self.edit_selected_row
        )
        self.menu.add_command(
            label=LANG_TEXT[self.lang].get("delete", "Delete"),
            command=self.delete_selected_row
        )

        self.tree.bind("<Button-3>", self.show_context_menu)

        # ===============================
        # Hover Tooltips for specific columns
        # ===============================
        def show_tooltip(text, x, y):
            if hasattr(self, "tooltip"):
                self.tooltip.destroy()
            self.tooltip = tk.Toplevel(self.tree)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x+20}+{y+20}")
            label = tk.Label(self.tooltip, text=text, bg="lightyellow",
                            relief="solid", borderwidth=1, font=("Segoe UI", 9))
            label.pack()

        def hide_tooltip(event=None):
            if hasattr(self, "tooltip"):
                self.tooltip.destroy()
                del self.tooltip

        def on_tree_hover(event):
            region = self.tree.identify("region", event.x, event.y)
            if region == "heading":
                col = self.tree.identify_column(event.x)
                col_index = int(col.replace("#", "")) - 1
                if col_index < len(COLUMNS):
                    col_name = COLUMNS[col_index]
                    info = LANG_TEXT[self.lang]["TypeInfo"].get(col_name)
                    if info:
                        text = f"{info['title']}: {info['details']}"
                        show_tooltip(text, event.x_root, event.y_root)
            else:
                hide_tooltip()

        self.tree.bind("<Motion>", on_tree_hover)
        self.tree.bind("<Leave>", hide_tooltip)

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
            fg="#005f99",
            font=("Segoe UI", 10)
        )
        self.result_label.pack(side="left", padx=15)


    # ---------- Filters ----------
    def create_filters(self):
        # Clear old filters
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

        filters_container = tk.Frame(container, bg="#f0f4f8")
        filters_container.grid(row=0, column=0, sticky="w")

        row, col = 0, 0
        self.filters = {}

        # Loop through all columns
        for idx, col_name in enumerate(labels):
            eng_col = self.columns_data["english"][idx]

            # Skip hidden columns
            if not self.columns_visibility.get(eng_col, True):
                continue

            # Auto wrap if too wide
            if col > 6:
                row += 1
                col = 0

            # Label
            tk.Label(filters_container, text=col_name,width=15, fg="#005f99", bg="#f0f4f8", font=("Segoe UI", 9, "bold"))\
                .grid(row=row, column=col, sticky="w", padx=(10, 0))
            col += 1

            # Build widget depending on column type
            if eng_col not in self.df.columns:
                continue

            if eng_col == "Search No":
                var = tk.StringVar()

                entry = tk.Entry(
                    filters_container,
                    textvariable=var,
                    width=14,
                    relief="flat",
                    bd=1,
                    highlightthickness=1,
                    highlightbackground="black",
                    font=("Segoe UI", 10),
                    bg="#ccffcc"
                )
                entry.grid(row=row, column=col, padx=5, pady=0)

                # Localized placeholder
                placeholder = LANG_TEXT[self.lang].get("search_placeholder", " Search No...")

                def show_placeholder():
                    entry.delete(0, "end")
                    entry.insert(0, placeholder)
                    entry.config(fg="gray")

                def hide_placeholder():
                    if entry.get() == placeholder:
                        entry.delete(0, "end")
                    entry.config(fg="black")

                # Initialize placeholder
                show_placeholder()

                def on_focus_in(event):
                    hide_placeholder()

                def on_focus_out(event):
                    if entry.get().strip() == "":
                        show_placeholder()

                entry.bind("<FocusIn>", on_focus_in)
                entry.bind("<FocusOut>", on_focus_out)

                # Only allow numbers
                def validate_input(new_value):
                    return new_value == "" or new_value.isdigit()

                vcmd = (self.register(validate_input), "%P")
                entry.config(validate="key", validatecommand=vcmd)

                # Trigger filter when typing
                def on_var_change(*args):
                    if entry.get() != placeholder:  # Only filter if not placeholder
                        self.apply_filters()

                var.trace_add("write", on_var_change)

                widget = var
            else:
                unique_vals = sorted(self.df[eng_col].dropna().unique())
                if 0 < len(unique_vals) <= 20:  # categorical → dropdown
                    widget = MultiSelectDropdown(filters_container, unique_vals, lang_text=LANG_TEXT, lang=self.lang,width=14, callback=self.apply_filters)
                    widget.grid(row=row, column=col, padx=5)
                else:  # free text → entry
                    var = tk.StringVar()
                    tk.Entry(filters_container, textvariable=var, width=14, relief="solid", bd=1)\
                        .grid(row=row, column=col, padx=5)
                    var.trace_add("write", lambda *_: self.apply_filters())
                    widget = var

            self.filters[eng_col] = widget
            col += 1

        # ==========================
        # ACTION BUTTONS (RIGHT)
        # ==========================
        action_container = tk.Frame(container, bg="#f0f4f8")
        action_container.grid(row=0, column=1, sticky="e", padx=10)

        # Styled Add Button
        add_btn = tk.Button(
            action_container,
            text=self.t("add_entry") if hasattr(self, "t") else "Add Entry",
            fg="white",
            bg="#005f99",
            activebackground="#3399cc",
            activeforeground="white",
            font=("Segoe UI", 10),
            bd=1,
            relief="solid",
            padx=12,
            pady=6,
            cursor="hand2",
            command=self.open_add_window
        )
        add_btn.pack(side="right")

        # Styled Clear Button
        clear_btn = tk.Button(
            action_container,
            text=self.t("clear_filters") if hasattr(self, "t") else "Clear",
            fg="white",
            bg="#005f99",
            activebackground="#3399cc",
            activeforeground="white",
            font=("Segoe UI", 10),
            bd=1,
            relief="solid",
            padx=12,
            pady=6,
            cursor="hand2",
            command=self.clear_all_filters
        )
        clear_btn.pack(side="right", padx=5)

        # --- Hover effect helper ---
        def add_hover_effect(widget, normal_bg, hover_bg, fg="white"):
            widget.bind("<Enter>", lambda e: widget.configure(bg=hover_bg, fg=fg))
            widget.bind("<Leave>", lambda e: widget.configure(bg=normal_bg, fg=fg))

        # Apply hover effects
        add_hover_effect(add_btn, normal_bg="#005f99", hover_bg="#3399cc")
        add_hover_effect(clear_btn, normal_bg="#005f99", hover_bg="#3399cc")

    def clear_all_filters(self):
        # Reset all filter widgets
        for col, widget in self.filters.items():
            if isinstance(widget, tk.StringVar):
                widget.set("")
            else:  # MultiSelectDropdown
                widget.clear_selection()

        self.apply_filters()

    def apply_filters(self):
        df = self.df.copy()

        # Collect active filters
        active_filters = {}
        for col, widget in self.filters.items():
            if isinstance(widget, tk.StringVar):
                val = widget.get().strip()
                if val:
                    if col == "Search No":
                        # Special case: substring match for Search No
                        df = df[df[col].astype(str).str.contains(val, na=False)]
                    else:
                        # Exact match for other text inputs
                        df = df[df[col].astype(str) == val]
                    active_filters[col] = val
            else:
                selected = widget.get_selected()
                if selected:
                    df = df[df[col].isin(selected)]
                    active_filters[col] = selected

        # --- Update dropdown options dynamically ---
        for col, widget in self.filters.items():
            if not isinstance(widget, tk.StringVar):
                # Build mask excluding current column
                mask = pd.Series([True] * len(self.df))
                for other_col, val in active_filters.items():
                    if other_col == col:
                        continue
                    if isinstance(val, str):
                        if other_col == "Search No":
                            mask &= self.df[other_col].astype(str).str.contains(val, na=False)
                        else:
                            mask &= self.df[other_col].astype(str) == val
                    else:
                        mask &= self.df[other_col].isin(val)

                available = sorted(self.df[mask][col].dropna().unique())
                widget.values = available
                widget.selected = [v for v in widget.selected if v in available]

        # --- Refresh table ---
        self.filtered_df = df
        self.refresh_table(df)

        # Update status bar
        self.result_label.config(text=f"{len(df)} Results")

    def export_filtered(self):
        today = datetime.now().strftime("%Y-%m-%d")
        checklist_name = LANG_TEXT[self.lang].get("checklist", "Checklist")
        default_name = f"{today}_{checklist_name}.xlsx"




        if not hasattr(self, "filtered_df") or self.filtered_df.empty:
            messagebox.showinfo(self.t("error"), self.t("no_data_to_export"))
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Files", "*.xlsx")]
        )

        if not file_path:
            return

        try:
            # Load template workbook
            template_path = os.path.join("template", "template_export_filter.xlsx")
            wb = load_workbook(template_path)
            ws = wb.active

            # Extract only required columns
            export_df = self.filtered_df[["Search No", "Contents", "Before correction", "After correction"]]

            # Start writing at row 5, columns D–G
            start_row = 5
            start_col = 4  # D = 4
            for r_idx, row in enumerate(export_df.itertuples(index=False), start=start_row):
                for c_idx, value in enumerate(row, start=start_col):
                    ws.cell(row=r_idx, column=c_idx, value=value)

            # Save to chosen file path
            wb.save(file_path)

            # === Add Form Control checkboxes via Excel COM automation ===
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False  # keep Excel hidden
                wb_excel = excel.Workbooks.Open(file_path)
                ws_excel = wb_excel.Sheets(1)

                for r in range(start_row, start_row + len(export_df)):
                    # Column B
                    cell_b = ws_excel.Cells(r, 2)
                    cb_b = ws_excel.CheckBoxes().Add(cell_b.Left, cell_b.Top, 15, 15)
                    cb_b.Text = ""  # remove label

                    # Column C
                    cell_c = ws_excel.Cells(r, 3)
                    cb_c = ws_excel.CheckBoxes().Add(cell_c.Left, cell_c.Top, 15, 15)
                    cb_c.Text = ""  # remove label

                wb_excel.Save()
                wb_excel.Close()
                excel.Quit()
            except Exception as e:
                messagebox.showwarning("Checkbox Insert",
                                    f"Data exported, but checkboxes could not be added.\n{e}")

            # Show confirmation
            messagebox.showinfo(self.t("export_done"),
                                self.t("export_msg").format(file=file_path))

            # Auto-open the file
            try:
                if os.name == "nt":  # Windows
                    os.startfile(file_path)
                elif sys.platform == "darwin":  # macOS
                    subprocess.call(["open", file_path])
                else:  # Linux and others
                    subprocess.call(["xdg-open", file_path])
            except Exception as e:
                messagebox.showwarning("Open File", f"File saved but could not be opened automatically.\n{e}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export: {e}")

    def update_headers(self):
        self.text = LANG_TEXT[self.lang]

        # Update window title
        self.title(self.text["app_title"])
        self.title_lbl.config(text=self.text["app_title"], font=("Segoe UI", 20, "bold"))
        self.export_btn.config(text=self.text["export_excel"], font=("Segoe UI", 10))
        self.settings_btn.config(text=self.text["settings"], font=("Segoe UI", 10))
        self.export_filtered_btn.config(text=self.text["export_filtered"], font=("Segoe UI", 10))

        # Update filter frame title
        self.filter_frame.config(text=self.text["filters"], font=("Segoe UI", 10, "bold"))

        # Reset Treeview columns
        self.tree["columns"] = COLUMNS + ["PDF"]

        # Language-based column headers
        headers = (
            self.columns_data["japanese"]
            if self.lang == "Japanese"
            else self.columns_data["english"]
        )

        note_columns_en = ["Model Name", "Target Part Name", "Motor Specification", "Issue Classification", "Update Info"]
        note_columns_jp = ["モデル名", "対象部品名", "モータ仕様", "課題分類", "更新情報"]

        for i, col in enumerate(COLUMNS):
            header_text = headers[i]
            if col in note_columns_en or header_text in note_columns_jp:
                self.tree.heading(col, text=header_text, anchor="center")
                self.tree.column(col, width=200, anchor="center", stretch=False)
            elif col == "Updated By":
                self.tree.heading(col, text=header_text, anchor="center")
                self.tree.column(col, width=150, anchor="center", stretch=False)
            elif col == "Upload Date":
                self.tree.heading(col, text=header_text, anchor="center")
                self.tree.column(col, width=180, anchor="center", stretch=False)
            else:
                self.tree.heading(col, text=header_text, anchor="center")
                self.tree.column(col, width=160, anchor="center", stretch=False)

        # PDF column (fixed)
        self.tree.heading("PDF", text=self.t("pdf_header"), anchor="center")
        self.tree.column("PDF", width=120, anchor="center", stretch=False)

    def refresh_table(self, df):
        # Sort by numeric Search No
        if "Search No" in df.columns and not df.empty:
            df = df.copy()
            df["Search No"] = pd.to_numeric(df["Search No"], errors="coerce").fillna(0)
            df = df.sort_values("Search No")  # ascending order

        # Clear existing rows
        self.tree.delete(*self.tree.get_children())
        self.filtered_df = df.copy()

        # Reset columns and apply headers/widths
        self.tree["columns"] = COLUMNS + ["PDF"]
        self.update_headers()

        if df.empty:
            no_record_msg = self.t("no_record_found")
            values = [no_record_msg] + ["" for _ in COLUMNS[1:]] + [""]
            self.tree.insert("", "end", values=values, tags=("missing",))
            self.tree.tag_configure(
                "missing",
                background="#f0f0f0",
                foreground="gray",
                font=("Segoe UI", 10, "italic")   # modern italic font
            )
            for col in self.tree["columns"]:
                self.tree.column(col, anchor="center")
        else:
            for _, row in df.iterrows():
                pdf = find_pdf(row.get("Search No", ""))

                if pdf:
                    # PDF exists → show plain text only
                    status = self.t("pdf_exists")
                    item_id = self.tree.insert(
                        "",
                        "end",
                        values=[row.get(c, "") for c in COLUMNS] + [status],
                    )
                    self.tree.item(item_id, tags=("pdf_exists",))
                else:
                    # PDF missing → show black ✖ symbol
                    status = f"✖ {self.t('pdf_missing')}"
                    item_id = self.tree.insert(
                        "",
                        "end",
                        values=[row.get(c, "") for c in COLUMNS] + [status],
                    )
                    self.tree.item(item_id, tags=("pdf_missing",))

            # Configure tags with modern italic font
            self.tree.tag_configure("pdf_exists", foreground="#000000", font=("Segoe UI", 10, "italic"))
            self.tree.tag_configure("pdf_missing", foreground="#000000", font=("Segoe UI", 10, "italic"))

            # Reset non-PDF columns to default foreground
            for item in self.tree.get_children():
                values = self.tree.item(item, "values")
                for i, col in enumerate(COLUMNS):
                    self.tree.set(item, col, values[i])

        # Apply row striping (alternating row colors for better readability)
        self.stripe_rows()

        # Update Result Counter with modern font
        result_count = len(df)
        total_count = len(self.df)
        result_text = f"{result_count} of {total_count} Results" if result_count != total_count else f"{result_count} Results"
        self.result_label.config(text=result_text, font=("Segoe UI", 10, "italic", "bold"))
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
            try:
                save_excel_with_lock(self.df)
            except Exception:
                messagebox.showerror(self.t("error"), self.t("save_failed"))
                return

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
        win.title(self.t("edit_title"))
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

        info_columns = ["Model Name", "Target Part Name", "Motor Specification", "Issue Classification", "Update Info"]

        for i, col in enumerate(COLUMNS):
            # Frame for label + info icon
            label_frame = tk.Frame(left_inner)
            label_frame.pack(anchor="w", padx=10, pady=(5, 0))

            # Field label
            tk.Label(label_frame, text=labels[i]).pack(side="left")

            # Add yellow ⓘ icon only for the five special fields
            if col in info_columns:
                info_icon = tk.Label(label_frame, text="ⓘ", fg="gold", font=("Segoe UI", 10, "bold"), cursor="hand2")
                info_icon.pack(side="left", padx=(5, 0))

                def show_tooltip(event, col=col):
                    info = LANG_TEXT[self.lang]["TypeInfo"].get(col)
                    if info:
                        if hasattr(self, "tooltip"):
                            self.tooltip.destroy()
                        self.tooltip = tk.Toplevel(self)
                        self.tooltip.wm_overrideredirect(True)
                        x, y = event.x_root, event.y_root
                        self.tooltip.wm_geometry(f"+{x+20}+{y+20}")
                        tk.Label(
                            self.tooltip,
                            text=f"{info['title']}: {info['details']}",
                            bg="lightyellow",
                            relief="solid",
                            borderwidth=1,
                            font=("Segoe UI", 9),
                            wraplength=300,
                            justify="left"
                        ).pack()

                def hide_tooltip(event):
                    if hasattr(self, "tooltip"):
                        self.tooltip.destroy()
                        del self.tooltip

                info_icon.bind("<Enter>", show_tooltip)
                info_icon.bind("<Leave>", hide_tooltip)

            # Input field below the label
            var = tk.StringVar(value=values[i])
            if col in DROPDOWN_OPTIONS:
                ent = ttk.Combobox(
                    left_inner,
                    textvariable=var,
                    values=DROPDOWN_OPTIONS[col],
                    width=80
                )
                ent.state(["!readonly"])
            else:
                ent = tk.Entry(left_inner, textvariable=var, width=80)

            ent.pack(fill="x", padx=10, pady=(0, 5))
            fields[col] = var

            # Prevent editing Search No
            if col == "Search No":
                ent.config(state="disabled")


        # ---------- Right: PDF preview ----------
        right_frame_outer = tk.Frame(paned)
        paned.add(right_frame_outer, stretch="never")

        right_canvas = tk.Canvas(right_frame_outer, width=300)
        right_scrollbar = tk.Scrollbar(right_frame_outer, orient="vertical", command=right_canvas.yview)
        right_inner = tk.Frame(right_canvas, padx=20, pady=20)

        right_inner.bind("<Configure>", lambda e: right_canvas.configure(scrollregion=right_canvas.bbox("all")))
        right_canvas.create_window((0, 0), window=right_inner, anchor="nw")
        right_canvas.configure(yscrollcommand=right_scrollbar.set)

        right_canvas.pack(side="left", fill="both", expand=True)
        right_scrollbar.pack(side="right", fill="y")

        right_inner.grid_columnconfigure(0, weight=1)

        pdf_var = tk.StringVar()
        existing_pdf = find_pdf(original_search_no)

        pdf_label = tk.Label(
            right_inner,
            text=os.path.basename(existing_pdf) if existing_pdf else self.t("no_pdf"),
            fg="green" if existing_pdf else "red"
        )
        pdf_label.grid(row=0, column=0, pady=(0, 5), sticky="n")

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

        ttk.Button(right_inner, text=self.t("replace_pdf"), command=select_new_pdf).grid(
            row=2, column=0, pady=5, padx=100, sticky="n"
        )

        # ---------- Bottom buttons ----------
        btn_frame = tk.Frame(win)
        btn_frame.pack(fill="x", pady=5, padx=5)

        ttk.Button(
            btn_frame, text=self.t("cancel"), command=win.destroy
        ).pack(side="right", padx=(0, 10))

        ttk.Button(
            btn_frame, text=self.t("save_changes"),
            command=lambda: self.save_edited_entry(win, fields, pdf_var, original_search_no)
        ).pack(side="right")

    def save_edited_entry(self, win, fields, pdf_var, original_search_no):
        # Reload latest Excel before editing
        if os.path.exists(EXCEL_PATH):
            latest_df = pd.read_excel(EXCEL_PATH, dtype=str).fillna("")
        else:
            latest_df = pd.DataFrame(columns=COLUMNS)

        # Find the row to edit
        idx = latest_df[latest_df["Search No"] == str(original_search_no)].index
        if idx.empty:
            messagebox.showerror(self.t("error"), self.t("not_found_error"))
            return

        # Validate Search No before saving
        search_no_val = fields["Search No"].get().strip()
        if not search_no_val.isdigit():
            messagebox.showerror(self.t("error"), self.t("search_no_numeric_error"))
            return

        # Update DataFrame values (excluding "Updated By" and "Upload Date" for now)
        for col in COLUMNS:
            if col == "Search No":
                latest_df.loc[idx, col] = search_no_val
            elif col not in ["Updated By", "Upload Date"]:
                latest_df.loc[idx, col] = fields[col].get()

        # Now update only the "Updated By" and "Upload Date" fields
        latest_df.loc[idx, "Updated By"] = getpass.getuser()  # Set the current user as "Updated By"
        latest_df.loc[idx, "Upload Date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Set the current date/time

        # Handle PDF replacement if needed
        if pdf_var.get():
            if not os.path.exists(PDF_DIR):
                os.makedirs(PDF_DIR)

            old_pdf = find_pdf(original_search_no)
            if old_pdf and os.path.exists(old_pdf):
                os.remove(old_pdf)

            search_no_norm = str(search_no_val).zfill(3)
            type1 = fields.get("Type 1", tk.StringVar(value="")).get().strip()
            type2 = fields.get("Type 2", tk.StringVar(value="")).get().strip()

            new_pdf_name = f"検索No.{search_no_norm}_{type1}_{type2}.pdf"
            new_pdf_path = os.path.join(PDF_DIR, new_pdf_name)
            shutil.copy(pdf_var.get(), new_pdf_path)

        # Sort before saving
        latest_df["Search No"] = pd.to_numeric(latest_df["Search No"], errors="coerce").fillna(0)
        latest_df = latest_df.sort_values("Search No").reset_index(drop=True)

        # Save with lock
        try:
            save_excel_with_lock(latest_df)
        except Exception:
            messagebox.showerror(self.t("error"), self.t("save_failed"))
            return

        # Refresh UI
        self.df = latest_df
        self.update_headers()
        self.create_filters()
        self.refresh_table(self.df)

        win.destroy()
        messagebox.showinfo(self.t("success"), self.t("save_entry"))

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

    def manage_columns(self):
        if hasattr(self, "columns_win") and self.columns_win is not None:
            if self.columns_win.winfo_exists():
                self.columns_win.deiconify()
                self.columns_win.lift()
                self.columns_win.focus_force()
                return

        win = tk.Toplevel(self)
        self.columns_win = win
        win.title(self.t("manage_columns"))
        win.geometry("540x460")
        win.resizable(False, False)

        # Main container with padding
        container = ttk.Frame(win, padding=20)
        container.pack(fill="both", expand=True)

        def on_close():
            self.columns_visibility = {}
            for item in self.columns_tree.get_children():
                col_name = self.columns_tree.set(item, "Column")
                if self.lang == "Japanese":
                    idx = JAPANESE_COLUMNS.index(col_name)
                    eng_col = COLUMNS[idx]
                else:
                    eng_col = col_name
                checkbox = self.columns_tree.set(item, "Visible")
                self.columns_visibility[eng_col] = (checkbox == "☑")

            save_columns({
                "english": COLUMNS,
                "japanese": JAPANESE_COLUMNS,
                "visible": self.columns_visibility
            })

            self.columns_win = None
            win.destroy()
            self.create_filters()
            self.update_headers()
            self.refresh_table(self.df)

        win.protocol("WM_DELETE_WINDOW", on_close)

        # -------- Treeview (Modernized) --------
        self.columns_tree = ttk.Treeview(
            container,
            columns=("Column", "Visible"),
            show="headings",
            selectmode="browse"
        )

        self.columns_tree.heading("Column", text=self.t("column_header"), anchor="center")
        self.columns_tree.heading("Visible", text=self.t("visible_header"), anchor="center")

        self.columns_tree.column("Column", width=360, anchor="w")
        self.columns_tree.column("Visible", width=120, anchor="center")
        self.columns_tree.pack(fill="both", expand=True, pady=(0, 15))

        # Style for row height (slightly taller rows)
        style = ttk.Style()
        style.configure("Treeview", rowheight=30)

        for eng, jpn in zip(COLUMNS, JAPANESE_COLUMNS):
            visible = "☑" if self.columns_visibility.get(eng, True) else "☐"
            self.columns_tree.insert(
                "", "end",
                values=(jpn if self.lang == "Japanese" else eng, visible)
            )

        # Toggle visibility on double-click
        def toggle_visible(event):
            selection = self.columns_tree.selection()
            if selection:
                item_id = selection[0]
                current = self.columns_tree.set(item_id, "Visible")
                new_val = "☐" if current == "☑" else "☑"

                if new_val == "☑":
                    visible_count = sum(
                        1 for i in self.columns_tree.get_children()
                        if self.columns_tree.set(i, "Visible") == "☑"
                    )
                    if visible_count >= 8:
                        tk.messagebox.showwarning(
                            self.t("limit_title"),
                            self.t("limit_message").format(limit=8)
                        )
                        return
                self.columns_tree.set(item_id, "Visible", new_val)

        self.columns_tree.bind("<Double-1>", toggle_visible)

        # -------- Add Column Button --------
        add_column_btn = tk.Button(
            container,
            text=self.t("add_column"),
            bg="#005f99", fg="white", activebackground="#3399cc", activeforeground="white",
            font=("Segoe UI", 10, "bold"), cursor="hand2", bd=0, relief="solid",
            padx=12, pady=6, command=lambda: self._open_add_column_popup()
        )
        add_column_btn.pack(pady=(0, 10))

        # -------- Remove Column Button --------
        remove_column_btn = tk.Button(
            container,
            text=self.t("remove_column"),
            bg="#dc3545", fg="white", activebackground="#e63946", activeforeground="white",
            font=("Segoe UI", 10, "bold"), cursor="hand2", bd=0, relief="solid",
            padx=12, pady=6, command=self._remove_column
        )
        remove_column_btn.pack(pady=(0, 10))

    def _open_add_column_popup(self):
        popup = tk.Toplevel(self)
        popup.title(self.t("add_column"))
        popup.geometry("440x280")
        popup.resizable(False, False)
        popup.transient(self)
        popup.grab_set()

        container = ttk.Frame(popup, padding=20)
        container.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure("Modern.TLabel", font=("Segoe UI", 10))
        style.configure("Modern.TEntry", padding=6)
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=8, background="#28a745", foreground="white")
        style.map("Primary.TButton", background=[("active", "#45c767")], foreground=[("active", "white")])
        style.configure("Secondary.TButton", font=("Segoe UI", 9), padding=6, background="#005f99", foreground="white")
        style.map("Secondary.TButton", background=[("active", "#3399cc")], foreground=[("active", "white")])

        # English column
        ttk.Label(container, text=self.t("enter_column_name_english"), style="Modern.TLabel").pack(anchor="w")
        eng_var = tk.StringVar()
        ttk.Entry(container, textvariable=eng_var, style="Modern.TEntry").pack(fill="x", pady=(5, 15))

        # Japanese column
        ttk.Label(container, text=self.t("enter_column_name_japanese"), style="Modern.TLabel").pack(anchor="w")
        jpn_var = tk.StringVar()
        ttk.Entry(container, textvariable=jpn_var, style="Modern.TEntry").pack(fill="x", pady=(5, 20))

        # Buttons
        btn_frame = ttk.Frame(container)
        btn_frame.pack(fill="x")

        def on_submit():
            eng_name = eng_var.get().strip()
            jpn_name = jpn_var.get().strip()

            if not eng_name or not jpn_name:
                messagebox.showerror(self.t("error"), self.t("invalid_column_name"))
                return
            if eng_name in COLUMNS or jpn_name in JAPANESE_COLUMNS:
                messagebox.showerror(self.t("error"), self.t("column_exists").format(col=eng_name))
                return

            try:
                if "Upload Date" in COLUMNS:
                    idx = COLUMNS.index("Upload Date")
                else:
                    idx = len(COLUMNS)

                COLUMNS.insert(idx, eng_name)
                JAPANESE_COLUMNS.insert(idx, jpn_name)
                self.df.insert(idx, eng_name, "")

                self.columns_tree.insert(
                    "", "end",
                    values=(jpn_name if self.lang == "Japanese" else eng_name, "☐")
                )
                self.columns_visibility[eng_name] = False

                save_columns({
                    "english": COLUMNS,
                    "japanese": JAPANESE_COLUMNS,
                    "visible": self.columns_visibility
                })
                save_excel_with_lock(self.df)

                self.create_filters()
                self.update_headers()
                self.refresh_table(self.df)

                messagebox.showinfo(self.t("success"), self.t("column_added").format(col=eng_name))
                popup.destroy()

            except Exception as e:
                messagebox.showerror(self.t("error"), f"{self.t('add_failed')}: {e}")

        # Buttons side by side
        ttk.Button(btn_frame, text=self.t("cancel"), style="Secondary.TButton", command=popup.destroy).pack(side="right")
        ttk.Button(btn_frame, text=self.t("save_changes"), style="Primary.TButton", command=on_submit).pack(side="right", padx=(0, 10))

    def _remove_column(self):
        item = self.columns_tree.selection()
        if not item:
            messagebox.showerror(self.t("error"), self.t("no_selection"))
            return

        col_name = self.columns_tree.set(item, "Column")
        if self.lang == "Japanese":
            idx = JAPANESE_COLUMNS.index(col_name)
            eng_col = COLUMNS[idx]
        else:
            idx = COLUMNS.index(col_name)
            eng_col = col_name

        if eng_col in ["Search No", "Reference model"]:
            messagebox.showerror(self.t("error"), self.t("cannot_remove"))
            return
        try:
            del COLUMNS[idx]
            del JAPANESE_COLUMNS[idx]
            self.df.drop(columns=[eng_col], inplace=True)
            self.columns_tree.delete(item)
            self.columns_visibility.pop(eng_col, None)

            save_columns({"english": COLUMNS, "japanese": JAPANESE_COLUMNS, "visible": self.columns_visibility})
            save_excel_with_lock(self.df)

            self.create_filters()
            self.update_headers()
            self.refresh_table(self.df)
            messagebox.showinfo(self.t("success"), self.t("column_removed").format(col=eng_col))
        except Exception as e:
            messagebox.showerror(self.t("error"), f"{self.t('remove_failed')}: {e}")
    # ---------- Settings ----------
    def open_settings(self):
        # Prevent multiple windows
        if hasattr(self, "settings_win") and self.settings_win is not None:
            if self.settings_win.winfo_exists():
                self.settings_win.deiconify()
                self.settings_win.lift()
                self.settings_win.focus_force()
                return

        win = tk.Toplevel(self)
        self.settings_win = win
        win.title(self.t("settings_title"))
        win.geometry("520x450")
        win.resizable(False, False)

        # Close callback
        def on_close():
            if hasattr(self, "columns_win") and self.columns_win is not None:
                if self.columns_win.winfo_exists():
                    self.columns_win.destroy()
                    self.columns_win = None
            self.settings_win = None
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", on_close)

        excel_var = tk.StringVar(value=EXCEL_PATH)
        pdf_var = tk.StringVar(value=PDF_DIR)
        lang_var = tk.StringVar(value=self.lang)

        def add_hover_effect(widget, normal_bg, hover_bg, fg="white"):
            widget.bind("<Enter>", lambda e: widget.configure(bg=hover_bg, fg=fg))
            widget.bind("<Leave>", lambda e: widget.configure(bg=normal_bg, fg=fg))

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

        # ---------- Main container ----------
        container = tk.Frame(win, padx=25, pady=20)
        container.pack(fill="both", expand=True)

        field_pad_y = 8  # space between label and entry
        section_pad_y = 15  # space between sections

        # ---------- Excel ----------
        excel_label = tk.Label(container,
                            text=LANG_TEXT[lang_var.get()]["excel_file"],
                            font=("Segoe UI", 10, "bold"))
        excel_label.pack(anchor="w", pady=(0, field_pad_y))

        excel_frame = tk.Frame(container)
        excel_frame.pack(fill="x", pady=(0, section_pad_y))

        excel_browse_btn = tk.Button(
            excel_frame,
            text=LANG_TEXT[lang_var.get()]["browse_excel"],
            fg="white", bg="#005f99",
            activebackground="#3399cc", activeforeground="white",
            font=("Segoe UI", 10, "bold"), bd=0,
            padx=12, pady=6, cursor="hand2",
            command=lambda: excel_var.set(filedialog.askopenfilename(
                title=LANG_TEXT[lang_var.get()]["browse_excel"],
                filetypes=[("Excel files", "*.xlsx *.xls")]
            ))
        )
        excel_browse_btn.pack(side="left", padx=(0, 8))
        add_hover_effect(excel_browse_btn, "#005f99", "#3399cc")

        excel_entry = tk.Entry(excel_frame, textvariable=excel_var,
                            font=("Segoe UI", 10),
                            state="readonly",
                            relief="solid", bd=1,
                            highlightthickness=1, highlightbackground="#cccccc",
                            justify="left")
        excel_entry.pack(side="left", fill="x", expand=True, ipady=6)

        # ---------- PDF ----------
        pdf_label = tk.Label(container,
                            text=LANG_TEXT[lang_var.get()]["pdf_folder"],
                            font=("Segoe UI", 10, "bold"))
        pdf_label.pack(anchor="w", pady=(0, field_pad_y))

        pdf_frame = tk.Frame(container)
        pdf_frame.pack(fill="x", pady=(0, section_pad_y))
        
        pdf_browse_btn = tk.Button(
            pdf_frame,
            text=LANG_TEXT[lang_var.get()]["browse_pdf"],
            fg="white", bg="#005f99",
            activebackground="#3399cc", activeforeground="white",
            font=("Segoe UI", 10, "bold"), bd=0,
            padx=12, pady=6, cursor="hand2",
            command=lambda: pdf_var.set(filedialog.askdirectory(
                title=LANG_TEXT[lang_var.get()]["browse_pdf"]
            ))
        )
        pdf_browse_btn.pack(side="left", padx=(0, 8))
        add_hover_effect(pdf_browse_btn, "#005f99", "#3399cc")

        pdf_entry = tk.Entry(pdf_frame, textvariable=pdf_var,
                            font=("Segoe UI", 10),
                            state="readonly",
                            relief="solid", bd=1,
                            highlightthickness=1, highlightbackground="#cccccc",
                            justify="left")
        pdf_entry.pack(side="left", fill="x", expand=True, ipady=6)


        # ---------- Language ----------
        lang_label = tk.Label(container,
                            text=LANG_TEXT[lang_var.get()]["default_language"],
                            font=("Segoe UI", 10, "bold"))
        lang_label.pack(anchor="w", pady=(0, field_pad_y))

        lang_combobox = ttk.Combobox(container, textvariable=lang_var,
                                    values=list(LANG_TEXT.keys()),
                                    font=("Segoe UI", 10), state="readonly")
        lang_combobox.pack(anchor="w", pady=(0, section_pad_y), fill="x")

        # ---------- Buttons ----------
        button_frame = tk.Frame(container)
        button_frame.pack(fill="x", pady=(10, 0))

        columns_btn = tk.Button(
            button_frame,
            text=LANG_TEXT[lang_var.get()]["manage_columns"],
            fg="white", bg="#005f99",
            activebackground="#3399cc", activeforeground="white",
            font=("Segoe UI", 10, "bold"), bd=0,
            padx=14, pady=6, cursor="hand2",
            command=self.manage_columns
        )
        columns_btn.pack(fill="x", pady=(0, 8))
        add_hover_effect(columns_btn, "#005f99", "#3399cc")

        save_btn = tk.Button(
            button_frame,
            text=LANG_TEXT[lang_var.get()]["save_settings"],
            fg="white", bg="#28a745",
            activebackground="#45c767", activeforeground="white",
            font=("Segoe UI", 10, "bold"), bd=0,
            padx=14, pady=6, cursor="hand2",
            command=lambda: self.save_settings(win, excel_var, pdf_var, lang_var)
        )
        save_btn.pack(fill="x")
        add_hover_effect(save_btn, "#28a745", "#45c767")

        lang_var.trace_add("write", update_text)

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
        if hasattr(self, "add_win") and self.add_win is not None:
            if self.add_win.winfo_exists():
                self.add_win.deiconify()
                self.add_win.lift()
                self.add_win.focus_force()
                return

        win = tk.Toplevel(self)
        self.add_win = win
        win.title(self.t("add_title"))
        win.geometry("900x500")
        win.minsize(700, 400)

        def on_close():
            self.add_win = None
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", on_close)

        paned = tk.PanedWindow(win, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=5, pady=5)

        # Left side
        left_frame_outer = tk.Frame(paned)
        paned.add(left_frame_outer, stretch="always")

        left_canvas = tk.Canvas(left_frame_outer)
        left_scrollbar = tk.Scrollbar(left_frame_outer, orient="vertical", command=left_canvas.yview)
        left_inner = tk.Frame(left_canvas)

        left_inner.bind("<Configure>", lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all")))
        left_canvas.create_window((0, 0), window=left_inner, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)
        left_canvas.pack(side="left", fill="both", expand=True)
        left_scrollbar.pack(side="right", fill="y")

        labels = self.columns_data["japanese"] if self.lang == "Japanese" else self.columns_data["english"]
        fields = {}

        info_columns = ["Model Name", "Target Part Name", "Motor Specification", "Issue Classification", "Update Info"]

        # Define which columns to skip
        skip_columns = {"Updated By", "Upload Date"}

        def only_numbers(char): 
            return char.isdigit() or char == "" 
        
        vcmd = (self.register(only_numbers), "%S")

        for i, col in enumerate(COLUMNS):
            if col in skip_columns:
                continue  # Skip these fields entirely

            # Frame for label + info icon
            label_frame = tk.Frame(left_inner)
            label_frame.pack(anchor="w", padx=10, pady=(5, 0))

            # Field label
            tk.Label(label_frame, text=labels[i]).pack(side="left")

            # Add yellow ⓘ icon only for the five special fields
            if col in info_columns:
                info_icon = tk.Label(label_frame, text="ⓘ", fg="gold", font=("Segoe UI", 10, "bold"), cursor="hand2")
                info_icon.pack(side="left", padx=(5, 0))

                def show_tooltip(event, col=col):
                    info = LANG_TEXT[self.lang]["TypeInfo"].get(col)
                    if info:
                        if hasattr(self, "tooltip"):
                            self.tooltip.destroy()
                        self.tooltip = tk.Toplevel(self)
                        self.tooltip.wm_overrideredirect(True)
                        x, y = event.x_root, event.y_root
                        self.tooltip.wm_geometry(f"+{x+20}+{y+20}")
                        tk.Label(
                            self.tooltip,
                            text=f"{info['title']}: {info['details']}",
                            bg="lightyellow",
                            relief="solid",
                            borderwidth=1,
                            font=("Segoe UI", 9),
                            wraplength=300,
                            justify="left"
                        ).pack()

                def hide_tooltip(event):
                    if hasattr(self, "tooltip"):
                        self.tooltip.destroy()
                        del self.tooltip

                info_icon.bind("<Enter>", show_tooltip)
                info_icon.bind("<Leave>", hide_tooltip)

            # Input field below the label
            var = tk.StringVar()
            if col in DROPDOWN_OPTIONS:
                ent = ttk.Combobox(left_inner, textvariable=var, values=DROPDOWN_OPTIONS[col], width=80)
                ent.state(["!readonly"])
            else:
                if col == "Search No":
                    ent = tk.Entry(
                        left_inner,
                        textvariable=var,
                        width=80,
                        validate="key",
                        validatecommand=vcmd
                    )
                else:
                    ent = tk.Entry(left_inner, textvariable=var, width=80)

            ent.pack(fill="x", padx=10, pady=(0, 5))
            fields[col] = {"var": var, "widget": ent}

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


    # -------------------------------
    # Save entry method
    # -------------------------------
    def save_entry(self, win, fields, pdf_var):
        if not fields["Search No"]["var"].get() or not pdf_var.get():
            messagebox.showerror(self.t("error"), self.t("required_error"))
            return

        if not os.path.exists(PDF_DIR):
            os.makedirs(PDF_DIR)

        search_no = str(fields["Search No"]["var"].get()).strip()

        # Reload latest Excel
        if os.path.exists(EXCEL_PATH):
            latest_df = pd.read_excel(EXCEL_PATH, dtype=str).fillna("")
        else:
            latest_df = pd.DataFrame(columns=COLUMNS)

        duplicate = latest_df[latest_df["Search No"].astype(str).str.strip() == search_no]
        if not duplicate.empty:
            messagebox.showerror(self.t("error"), self.t("duplicate_error"))
            return

        # Handle PDF saving
        search_no_norm = str(fields["Search No"]["var"].get()).zfill(3)
        type1 = fields.get("Type 1", {"var": tk.StringVar(value="")})["var"].get().strip()
        type2 = fields.get("Type 2", {"var": tk.StringVar(value="")})["var"].get().strip()

        new_pdf_name = f"検索No.{search_no_norm}_{type1}_{type2}.pdf"
        new_pdf_path = os.path.join(PDF_DIR, new_pdf_name)
        shutil.copy(pdf_var.get(), new_pdf_path)

        # Add new entry
        new_entry = {c: fields[c]["var"].get() for c in COLUMNS if c not in ["Updated By", "Upload Date"]}
        new_entry["Updated By"] = getpass.getuser()
        new_entry["Upload Date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        latest_df = pd.concat([latest_df, pd.DataFrame([new_entry])], ignore_index=True)

        # --- Update dropdowns.json ---
        try:
            with open(DROPDOWN_FILE, "r", encoding="utf-8") as f:
                dropdowns = json.load(f)
        except FileNotFoundError:
            dropdowns = {}

        for col in ["Model Name", "Target Part Name", "Motor Specification", "Issue Classification", "Update Info"]:
            val = fields[col]["var"].get().strip()
            if val:
                if col not in dropdowns:
                    dropdowns[col] = []
                if val not in dropdowns[col]:
                    dropdowns[col].append(val)
                    dropdowns[col].sort(key=str.lower)

                # Update combobox immediately
                widget = fields[col]["widget"]
                if isinstance(widget, ttk.Combobox):
                    widget.configure(values=dropdowns[col])

        with open(DROPDOWN_FILE, "w", encoding="utf-8") as f:
            json.dump(dropdowns, f, ensure_ascii=False, indent=2)

        # Sort before saving
        latest_df["Search No"] = pd.to_numeric(latest_df["Search No"], errors="coerce").fillna(0)
        latest_df = latest_df.sort_values("Search No").reset_index(drop=True)

        try:
            save_excel_with_lock(latest_df)
        except Exception:
            messagebox.showerror(self.t("error"), self.t("save_failed"))
            return

        self.df = latest_df
        self.update_headers()
        self.create_filters()
        self.refresh_table(self.df)

        win.destroy()
        messagebox.showinfo(self.t("success"), self.t("save_entry"))


# ===============================
# RUN
# ===============================
if __name__ == "__main__":
    DiagramApp().mainloop()