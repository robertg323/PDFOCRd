import os
import sys
import json
import shutil
import subprocess
import fitz # PyMuPDF
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
import tkinter.ttk as ttk
import csv
from datetime import datetime
import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageTk

_PREVIEW_W = 360        # width of the preview panel (pixels)
_PREVIEW_RENDER_W = 320 # target render width for PDF pages

# =========================
# TESSERACT PATH DISCOVERY
# =========================

def _find_tesseract_path():
    """
    Find Tesseract executable by checking common locations, then PATH.
    Returns tuple: (tesseract_exe_path, tessdata_prefix_path) or (None, None) if not found.
    """
    # Common Windows installation locations
    common_paths = [
        r'C:\Program Files\Tesseract-OCR',
        r'C:\Program Files (x86)\Tesseract-OCR',
        os.path.expanduser(r'~\AppData\Local\Tesseract-OCR'),
    ]
    
    # Check common locations
    for base_path in common_paths:
        tesseract_exe = os.path.join(base_path, 'tesseract.exe')
        tessdata_dir = os.path.join(base_path, 'tessdata')
        if os.path.exists(tesseract_exe):
            return (tesseract_exe, tessdata_dir)
    
    # Fall back to checking PATH
    try:
        result = subprocess.run(['where', 'tesseract.exe'], 
                              capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            tesseract_exe = result.stdout.strip().split('\n')[0]
            base_path = os.path.dirname(tesseract_exe)
            tessdata_dir = os.path.join(base_path, 'tessdata')
            return (tesseract_exe, tessdata_dir)
    except Exception:
        pass
    
    return (None, None)

# Get Tesseract paths and configure them
_TESSERACT_EXE, _TESSDATA_PREFIX = _find_tesseract_path()

if _TESSERACT_EXE and os.path.exists(_TESSERACT_EXE):
    pytesseract.pytesseract.tesseract_cmd = _TESSERACT_EXE
    if _TESSDATA_PREFIX and os.path.exists(_TESSDATA_PREFIX):
        os.environ['TESSDATA_PREFIX'] = _TESSDATA_PREFIX

# =========================
# POPPLER PATH DISCOVERY
# =========================

def _find_poppler_path():
    """
    Find Poppler executables by checking bundled path, common locations, then PATH.
    Returns the directory containing pdftoppm.exe, or None if not found.
    """
    candidate_paths = []
    
    # Check if running as a frozen PyInstaller exe (one-file or one-dir mode)
    if getattr(sys, 'frozen', False):
        # PyInstaller one-file mode: bundles are in sys._MEIPASS (temporary directory)
        if hasattr(sys, '_MEIPASS'):
            meipass_poppler = os.path.join(sys._MEIPASS, 'poppler')
            candidate_paths.append(meipass_poppler)
        
        # PyInstaller one-dir mode: bundles are next to executable
        exe_dir = os.path.dirname(sys.executable)
        exe_poppler = os.path.join(exe_dir, 'poppler')
        candidate_paths.append(exe_poppler)
    
    # Check relative to script directory (for development)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    script_poppler = os.path.join(script_dir, 'poppler')
    candidate_paths.append(script_poppler)
    
    # Search all candidate directories
    for poppler_root in candidate_paths:
        if os.path.isdir(poppler_root):
            # Check for nested poppler-<version> directories
            try:
                for item in os.listdir(poppler_root):
                    if item.startswith('poppler-'):
                        candidate = os.path.join(poppler_root, item, 'Library', 'bin')
                        if os.path.isdir(candidate) and os.path.exists(os.path.join(candidate, 'pdftoppm.exe')):
                            return candidate
            except (OSError, PermissionError):
                pass
            
            # Also check flat: poppler/Library/bin (if someone extracts it without nested dir)
            flat_poppler = os.path.join(poppler_root, 'Library', 'bin')
            if os.path.isdir(flat_poppler) and os.path.exists(os.path.join(flat_poppler, 'pdftoppm.exe')):
                return flat_poppler
    
    # Check common Windows installation paths
    common_paths = [
        r'C:\Program Files\poppler\Library\bin',
        r'C:\Program Files (x86)\poppler\Library\bin',
        os.path.expanduser(r'~\AppData\Local\poppler\Library\bin'),
    ]
    for poppler_path in common_paths:
        if os.path.isdir(poppler_path) and os.path.exists(os.path.join(poppler_path, 'pdftoppm.exe')):
            return poppler_path
    
    # Fall back to checking PATH
    try:
        result = subprocess.run(['where', 'pdftoppm.exe'], 
                              capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            pdftoppm_exe = result.stdout.strip().split('\n')[0]
            return os.path.dirname(pdftoppm_exe)
    except Exception:
        pass
    
    return None

# =========================
# APPLICATION CLASS (GUI)
# =========================

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Configure Window ---
        self.title("PDF OCR CHECK v1.01")
        self.geometry("950x850")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Configuration Defaults ---
        self.SOURCE_DIR = r"C:\PDFs\All_PDFs"
        self.DEST_DIR = r"C:\PDFs\Non_OCR_PDFs"
        self.REPORT_DIR = r"C:\PDFs\Reports"
        self.FILE_ACTION = ctk.StringVar(value="move")  # "move", "ocr", or "none"
        self.ENABLE_REPORT = ctk.BooleanVar(value=True)
        self.MIN_TEXT_CHARS = 50

        # --- New Feature Vars ---
        self.RECURSIVE_SCAN = ctk.BooleanVar(value=True)
        self.DRY_RUN = ctk.BooleanVar(value=False)
        self.tree_item_map = {}    # maps full_path -> treeview iid for Phase 2 row updates
        self.tree_iid_to_path = {} # reverse map: iid -> full_path for interactions

        # --- Preview State ---
        self.preview_pdf_path = None
        self.preview_page = 0
        self.preview_total_pages = 0
        self._preview_visible = False

        # --- Scan State ---
        self.scan_running = False
        self.cancel_requested = False
        self.last_csv_path = None

        # --- Config ---
        self.config_path = self._get_config_path()
        self._load_config()

        # --- Build UI Elements ---
        self.create_widgets()

        # --- Save on close ---
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # =========================
    # CONFIG PERSISTENCE
    # =========================

    def _get_config_path(self):
        """Return path to config JSON. Use LOCALAPPDATA on Windows, fallback to script dir."""
        if os.name == 'nt':
            appdata = os.environ.get('LOCALAPPDATA', '')
            if appdata:
                config_dir = os.path.join(appdata, 'PDFOCRd')
                os.makedirs(config_dir, exist_ok=True)
                return os.path.join(config_dir, 'config.json')
        # Fallback: same directory as the script/exe
        if getattr(sys, 'frozen', False):
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base, 'pdfocrd_config.json')

    def _load_config(self):
        """Load settings from JSON config. If file missing or corrupt, use defaults."""
        if not os.path.exists(self.config_path):
            return
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            self.SOURCE_DIR = cfg.get('source_dir', self.SOURCE_DIR)
            self.DEST_DIR = cfg.get('dest_dir', self.DEST_DIR)
            self.REPORT_DIR = cfg.get('report_dir', self.REPORT_DIR)
            self.MIN_TEXT_CHARS = cfg.get('threshold', self.MIN_TEXT_CHARS)
            # Load file_action; fall back to reading legacy move_files/ocr_in_place booleans
            if 'file_action' in cfg:
                self.FILE_ACTION.set(cfg['file_action'])
            elif cfg.get('ocr_in_place', False):
                self.FILE_ACTION.set('ocr')
            elif cfg.get('move_files', True):
                self.FILE_ACTION.set('move')
            else:
                self.FILE_ACTION.set('none')
            self.ENABLE_REPORT.set(cfg.get('enable_report', True))
            self.RECURSIVE_SCAN.set(cfg.get('recursive_scan', True))
            self.DRY_RUN.set(cfg.get('dry_run', False))
        except (json.JSONDecodeError, IOError, KeyError):
            pass  # Corrupt config -- silently use defaults

    def _save_config(self):
        """Persist current settings to JSON config file."""
        try:
            threshold = int(self.threshold_entry.get())
        except (ValueError, AttributeError):
            threshold = self.MIN_TEXT_CHARS
        cfg = {
            'source_dir': self.source_entry.get(),
            'dest_dir': self.dest_entry.get(),
            'report_dir': self.report_entry.get(),
            'threshold': threshold,
            'file_action': self.FILE_ACTION.get(),
            'enable_report': self.ENABLE_REPORT.get(),
            'recursive_scan': self.RECURSIVE_SCAN.get(),
            'dry_run': self.DRY_RUN.get(),
        }
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, indent=4)
        except IOError:
            pass  # Silently fail -- don't disrupt user flow

    def _on_close(self):
        """Called when user closes the window. Save config then destroy."""
        if self.scan_running:
            self.cancel_requested = True
        self._save_config()
        self.destroy()

    # =========================
    # REPORT VIEWER
    # =========================

    def view_report(self):
        """Opens the last generated CSV file using the default system program."""
        if self.last_csv_path and os.path.exists(self.last_csv_path):
            try:
                os.startfile(self.last_csv_path)
                self.log(f"Opened report: {os.path.basename(self.last_csv_path)}")
            except Exception as e:
                self.log(f"ERROR: Could not open report file. Check file association. Error: {e}")
        else:
            self.log("ERROR: No report file available to view.")

    # =========================
    # DIRECTORY SELECTION
    # =========================

    def select_source_dir(self):
        """Opens dialog to select the source directory and updates the entry field."""
        initial_dir = self.source_entry.get() if os.path.isdir(self.source_entry.get()) else self.SOURCE_DIR
        folder_selected = filedialog.askdirectory(initialdir=initial_dir)
        if folder_selected:
            self.source_entry.delete(0, "end")
            self.source_entry.insert(0, folder_selected)
            self.log(f"Source Directory set to: {folder_selected}")

    def select_dest_dir(self):
        """Opens dialog to select the destination directory and updates the entry field."""
        initial_dir = self.dest_entry.get() if os.path.isdir(self.dest_entry.get()) else self.DEST_DIR
        folder_selected = filedialog.askdirectory(initialdir=initial_dir)
        if folder_selected:
            self.dest_entry.delete(0, "end")
            self.dest_entry.insert(0, folder_selected)
            self.log(f"Destination Directory set to: {folder_selected}")

    def select_report_dir(self):
        """Opens dialog to select the report directory and updates the entry field."""
        initial_dir = self.report_entry.get() if os.path.isdir(self.report_entry.get()) else self.REPORT_DIR
        folder_selected = filedialog.askdirectory(initialdir=initial_dir)
        if folder_selected:
            self.report_entry.delete(0, "end")
            self.report_entry.insert(0, folder_selected)
            self.log(f"Report Directory set to: {folder_selected}")

    # =========================
    # FILE ACTION HANDLER
    # =========================

    def _on_file_action_change(self):
        """Enable dest entry/button only when 'Move files' is selected."""
        if self.FILE_ACTION.get() == "move":
            self.dest_entry.configure(state="normal")
            self.dest_button.configure(state="normal")
        else:
            self.dest_entry.configure(state="disabled")
            self.dest_button.configure(state="disabled")

    # =========================
    # GUI CONSTRUCTION
    # =========================

    def create_widgets(self):

        # --- Main PanedWindow (fills entire window, holds left controls + right preview) ---
        self.paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self.paned.grid(row=0, column=0, sticky="nsew")

        left_frame = ctk.CTkFrame(self.paned, fg_color="transparent", corner_radius=0)
        self.paned.add(left_frame, weight=1)
        left_frame.grid_columnconfigure(0, weight=1)
        left_frame.grid_rowconfigure(8, weight=2)   # Treeview gets more space
        left_frame.grid_rowconfigure(11, weight=1)  # Log gets some space

        # --- Treeview light theme styling (must be set before creating Treeview) ---
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview",
            background="white",
            foreground="black",
            fieldbackground="white",
            rowheight=25
        )
        style.configure("Treeview.Heading",
            background="#1F6AA5",
            foreground="white",
            font=("", 10, "bold")
        )
        style.map("Treeview.Heading",
            background=[("active", "#174F7A")]
        )
        style.map("Treeview", background=[("selected", "#93C5FD")])

        # ============================================
        # ROW 0: Header Frame (Source/Destination)
        # ============================================
        header_frame = ctk.CTkFrame(left_frame)
        header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        header_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(header_frame, text="Source Directory:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.source_entry = ctk.CTkEntry(header_frame, width=450)
        self.source_entry.insert(0, self.SOURCE_DIR)
        self.source_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.source_button = ctk.CTkButton(header_frame, text="...", width=30, command=self.select_source_dir)
        self.source_button.grid(row=0, column=2, padx=(0, 10), pady=5, sticky="e")

        ctk.CTkLabel(header_frame, text="Destination Directory:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.dest_entry = ctk.CTkEntry(header_frame, width=450)
        self.dest_entry.insert(0, self.DEST_DIR)
        self.dest_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        self.dest_button = ctk.CTkButton(header_frame, text="...", width=30, command=self.select_dest_dir)
        self.dest_button.grid(row=1, column=2, padx=(0, 10), pady=5, sticky="e")

        # ============================================
        # ROW 1: Threshold Input Frame
        # ============================================
        threshold_frame = ctk.CTkFrame(left_frame)
        threshold_frame.grid(row=1, column=0, padx=20, pady=5, sticky="ew")
        threshold_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(threshold_frame, text="Min. Text Chars (Threshold):").grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.threshold_entry = ctk.CTkEntry(threshold_frame, width=100)
        self.threshold_entry.insert(0, str(self.MIN_TEXT_CHARS))
        self.threshold_entry.grid(row=0, column=1, padx=10, pady=5, sticky="w")

        ctk.CTkLabel(threshold_frame, text="Set minimum characters needed to be considered 'OCR-D'").grid(row=0, column=2, padx=10, pady=5, sticky="w")

        # ============================================
        # ROW 2: Report Control Frame
        # ============================================
        report_frame = ctk.CTkFrame(left_frame)
        report_frame.grid(row=2, column=0, padx=20, pady=5, sticky="ew")
        report_frame.grid_columnconfigure(1, weight=1)

        self.report_check = ctk.CTkCheckBox(
            report_frame,
            text="Generate CSV Report",
            variable=self.ENABLE_REPORT
        )
        self.report_check.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.view_report_button = ctk.CTkButton(
            report_frame,
            text="View Report",
            command=self.view_report,
            state="disabled"
        )
        self.view_report_button.grid(row=0, column=2, padx=(0, 10), pady=5, sticky="e")

        ctk.CTkLabel(report_frame, text="Report Directory:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.report_entry = ctk.CTkEntry(report_frame, width=450)
        self.report_entry.insert(0, self.REPORT_DIR)
        self.report_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        self.report_button = ctk.CTkButton(report_frame, text="...", width=30, command=self.select_report_dir)
        self.report_button.grid(row=1, column=2, padx=(0, 10), pady=5, sticky="e")

        # ============================================
        # ROW 3: Options Frame (Recursive + Dry Run)
        # ============================================
        options_frame = ctk.CTkFrame(left_frame)
        options_frame.grid(row=3, column=0, padx=20, pady=5, sticky="ew")

        self.recursive_check = ctk.CTkCheckBox(
            options_frame,
            text="Scan subdirectories (recursive)",
            variable=self.RECURSIVE_SCAN
        )
        self.recursive_check.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.dryrun_check = ctk.CTkCheckBox(
            options_frame,
            text="Dry run (no files moved)",
            variable=self.DRY_RUN
        )
        self.dryrun_check.grid(row=0, column=1, padx=20, pady=5, sticky="w")

        # ============================================
        # ROW 4: Control Frame (Move, View Report, Start, Cancel)
        # ============================================
        control_frame = ctk.CTkFrame(left_frame)
        control_frame.grid(row=4, column=0, padx=20, pady=10, sticky="ew")
        control_frame.grid_columnconfigure(0, weight=1)

        radio_frame = ctk.CTkFrame(control_frame, fg_color="transparent")
        radio_frame.grid(row=0, column=0, rowspan=2, padx=10, pady=5, sticky="w")

        ctk.CTkLabel(radio_frame, text="Non-OCR file action:").grid(row=0, column=0, padx=(0, 10), pady=(5, 2), sticky="w")

        ctk.CTkRadioButton(
            radio_frame,
            text="Move files to destination directory",
            variable=self.FILE_ACTION,
            value="move",
            command=self._on_file_action_change
        ).grid(row=1, column=0, padx=10, pady=2, sticky="w")

        ctk.CTkRadioButton(
            radio_frame,
            text="OCR in place (creates _aocrded.pdf alongside original)",
            variable=self.FILE_ACTION,
            value="ocr",
            command=self._on_file_action_change
        ).grid(row=2, column=0, padx=10, pady=2, sticky="w")

        ctk.CTkRadioButton(
            radio_frame,
            text="No action (report only)",
            variable=self.FILE_ACTION,
            value="none",
            command=self._on_file_action_change
        ).grid(row=3, column=0, padx=10, pady=(2, 5), sticky="w")

        self.run_button = ctk.CTkButton(control_frame, text="START SCAN", command=self.start_scan)
        self.run_button.grid(row=0, column=1, padx=(5, 5), pady=10, sticky="e")

        self.cancel_button = ctk.CTkButton(
            control_frame,
            text="CANCEL",
            command=self.cancel_scan,
            state="disabled",
            fg_color="red",
            hover_color="darkred"
        )
        self.cancel_button.grid(row=0, column=2, padx=(5, 10), pady=10, sticky="e")

        # ============================================
        # ROW 5: Progress Frame
        # ============================================
        progress_frame = ctk.CTkFrame(left_frame)
        progress_frame.grid(row=5, column=0, padx=20, pady=5, sticky="ew")
        progress_frame.grid_columnconfigure(0, weight=1)

        self.progress_label = ctk.CTkLabel(progress_frame, text="0 of 0 files scanned")
        self.progress_label.grid(row=0, column=0, padx=10, pady=(5, 0), sticky="w")

        self.progress_bar = ctk.CTkProgressBar(progress_frame)
        self.progress_bar.grid(row=1, column=0, padx=10, pady=(0, 5), sticky="ew")
        self.progress_bar.set(0)

        self.ocr_page_status_label = ctk.CTkLabel(
            progress_frame, text="", text_color="gray", font=("", 11)
        )
        self.ocr_page_status_label.grid(row=2, column=0, padx=10, pady=(0, 5), sticky="w")

        # ============================================
        # ROW 6: Summary Stats Bar
        # ============================================
        stats_frame = ctk.CTkFrame(left_frame)
        stats_frame.grid(row=6, column=0, padx=20, pady=5, sticky="ew")

        self.stats_total = ctk.CTkLabel(stats_frame, text="Total: 0", font=("", 13, "bold"))
        self.stats_total.grid(row=0, column=0, padx=15, pady=5, sticky="w")

        self.stats_ocr = ctk.CTkLabel(stats_frame, text="OCR-D: 0", text_color="#2ea043")
        self.stats_ocr.grid(row=0, column=1, padx=15, pady=5, sticky="w")

        self.stats_nonocr = ctk.CTkLabel(stats_frame, text="NON-OCR: 0", text_color="#da3633")
        self.stats_nonocr.grid(row=0, column=2, padx=15, pady=5, sticky="w")

        self.stats_errors = ctk.CTkLabel(stats_frame, text="Errors: 0", text_color="#d29922")
        self.stats_errors.grid(row=0, column=3, padx=15, pady=5, sticky="w")

        # ============================================
        # ROW 7-8: Treeview Results Table
        # ============================================
        ctk.CTkLabel(left_frame, text="Scan Results:").grid(row=7, column=0, padx=20, pady=(10, 0), sticky="sw")

        tree_frame = ctk.CTkFrame(left_frame)
        tree_frame.grid(row=8, column=0, padx=20, pady=(5, 5), sticky="nsew")
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        columns = ("filename", "ocr_status", "char_count", "action")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=10)

        self.tree.heading("filename", text="Filename", command=lambda: self._sort_tree("filename", False))
        self.tree.heading("ocr_status", text="OCR Status", anchor="center", command=lambda: self._sort_tree("ocr_status", False))
        self.tree.heading("char_count", text="Char Count", anchor="center", command=lambda: self._sort_tree("char_count", False))
        self.tree.heading("action", text="Action Taken", command=lambda: self._sort_tree("action", False))

        self.tree.column("filename", width=350, minwidth=150)
        self.tree.column("ocr_status", width=100, minwidth=80, anchor="center")
        self.tree.column("char_count", width=100, minwidth=80, anchor="e")
        self.tree.column("action", width=200, minwidth=100, anchor="center")

        # Alternating row colors
        self.tree.tag_configure("even", background="white")
        self.tree.tag_configure("odd", background="#F0F0F0")

        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll.grid(row=0, column=1, sticky="ns")

        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Motion>", self._tree_tooltip_show)
        self.tree.bind("<Leave>", self._tree_tooltip_hide)
        self.tree.bind("<<TreeviewSelect>>", self._update_preview)

        # ============================================
        # ROW 9: Show/Hide Preview button (right-aligned, under treeview)
        # ============================================
        preview_btn_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        preview_btn_frame.grid(row=9, column=0, padx=20, pady=(2, 0), sticky="ew")
        preview_btn_frame.grid_columnconfigure(0, weight=1)

        self.preview_toggle_btn = ctk.CTkButton(
            preview_btn_frame,
            text="Show Preview",
            command=self._toggle_preview,
            width=120
        )
        self.preview_toggle_btn.grid(row=0, column=1, pady=2, sticky="e")

        # ============================================
        # ROW 10-11: Real-time Log
        # ============================================
        ctk.CTkLabel(left_frame, text="Real-time Log:").grid(row=10, column=0, padx=20, pady=(10, 0), sticky="sw")
        self.log_textbox = ctk.CTkTextbox(left_frame, wrap="word")
        self.log_textbox.grid(row=11, column=0, padx=20, pady=(5, 20), sticky="nsew")
        self.log_textbox.insert("end", "Ready to start. Click 'START SCAN'.\n")
        self.log_textbox.configure(state="disabled")

        # Apply initial dest entry state based on loaded config
        self._on_file_action_change()

        # ============================================
        # Preview Panel (right pane — added to PanedWindow on demand)
        # ============================================
        self.preview_frame = ctk.CTkFrame(self.paned, corner_radius=0)
        self.preview_frame.grid_rowconfigure(1, weight=1)
        self.preview_frame.grid_columnconfigure(0, weight=1)

        self.preview_title_label = ctk.CTkLabel(
            self.preview_frame, text="No file selected",
            font=("", 11, "bold"), wraplength=_PREVIEW_W - 20
        )
        self.preview_title_label.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 4), sticky="ew")

        canvas_outer = ctk.CTkFrame(self.preview_frame, fg_color="white")
        canvas_outer.grid(row=1, column=0, columnspan=2, padx=(10, 0), pady=4, sticky="nsew")
        canvas_outer.grid_rowconfigure(0, weight=1)
        canvas_outer.grid_columnconfigure(0, weight=1)

        self.preview_canvas = tk.Canvas(canvas_outer, bg="white", highlightthickness=0)
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")

        # Scrollbar repurposed as document-position indicator
        self.preview_vscroll = ttk.Scrollbar(
            self.preview_frame, orient="vertical",
            command=self._preview_scrollbar_command
        )
        self.preview_vscroll.grid(row=1, column=1, padx=(0, 10), pady=4, sticky="ns")

        self.preview_canvas.bind("<MouseWheel>", self._preview_mousewheel)
        self.preview_canvas.bind("<Configure>", self._on_preview_canvas_resize)

        self.preview_page_label = ctk.CTkLabel(self.preview_frame, text="")
        self.preview_page_label.grid(row=2, column=0, columnspan=2, padx=10, pady=(4, 10))

    # =========================
    # TREEVIEW SORTING
    # =========================

    def _sort_tree(self, col, reverse):
        """Sort treeview by column. Numeric sort for char_count, string sort otherwise."""
        data = [(self.tree.set(child, col), child) for child in self.tree.get_children("")]

        if col == "char_count":
            def sort_key(item):
                try:
                    return int(item[0].replace(",", ""))
                except ValueError:
                    return -1 if reverse else float('inf')
            data.sort(key=sort_key, reverse=reverse)
        else:
            data.sort(key=lambda t: t[0].lower(), reverse=reverse)

        for index, (_, child) in enumerate(data):
            self.tree.move(child, "", index)
            # Reapply alternating row colors after sort
            self.tree.item(child, tags=("even" if index % 2 == 0 else "odd",))

        # Toggle sort direction for next click
        self.tree.heading(col, command=lambda: self._sort_tree(col, not reverse))

    # =========================
    # TREEVIEW INTERACTIONS
    # =========================

    def _on_tree_double_click(self, event):
        """Open the PDF under the cursor with the default system program."""
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        path = self.tree_iid_to_path.get(iid, "")
        if not path:
            return
        if os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                self.log(f"ERROR: Could not open file: {e}")
        else:
            self.log(f"ERROR: File not found (may have been moved): {os.path.basename(path)}")

    def _tree_tooltip_show(self, event):
        """Show a tooltip with the full file path for the row under the cursor, only if selected."""
        iid = self.tree.identify_row(event.y)
        if not iid or iid not in self.tree.selection():
            self._tree_tooltip_hide()
            return
        path = self.tree_iid_to_path.get(iid, "")
        if not path:
            self._tree_tooltip_hide()
            return
        if not hasattr(self, '_tooltip') or self._tooltip is None:
            self._tooltip = tk.Toplevel(self)
            self._tooltip.wm_overrideredirect(True)
            self._tooltip.wm_attributes("-topmost", True)
            self._tooltip_label = tk.Label(
                self._tooltip, text=os.path.basename(path),
                bg="#ffffe0", fg="black",
                relief="solid", borderwidth=1,
                font=("", 9), padx=4, pady=2
            )
            self._tooltip_label.pack()
        else:
            self._tooltip_label.configure(text=os.path.basename(path))
            self._tooltip.deiconify()
        self._tooltip.wm_geometry(f"+{event.x_root + 14}+{event.y_root + 14}")

    def _tree_tooltip_hide(self, event=None):
        """Hide the tooltip window."""
        if hasattr(self, '_tooltip') and self._tooltip is not None:
            self._tooltip.withdraw()

    # =========================
    # PDF PREVIEW
    # =========================

    def _toggle_preview(self):
        """Show or hide the preview panel via PanedWindow add/forget."""
        if self._preview_visible:
            self.paned.forget(self.preview_frame)
            self._preview_visible = False
            self.preview_toggle_btn.configure(text="Show Preview")
        else:
            self.paned.add(self.preview_frame, weight=0)
            self._preview_visible = True
            self.preview_toggle_btn.configure(text="Hide Preview")
            # Delay render until the canvas has been sized by the layout manager
            self.after(50, self._update_preview)

    def _update_preview(self, event=None):
        """Refresh preview when treeview selection changes."""
        if not self.preview_frame.winfo_ismapped():
            return
        selection = self.tree.selection()
        if not selection:
            self._preview_show_placeholder("Select a row to preview")
            return
        path = self.tree_iid_to_path.get(selection[0], "")
        if not path:
            self._preview_show_placeholder("No path available")
            return
        if not os.path.exists(path):
            self._preview_show_placeholder("File not found\n(may have been moved)")
            return
        if path != self.preview_pdf_path:
            self.preview_pdf_path = path
            self.preview_page = 0
        self._render_preview_page()

    def _render_preview_page(self):
        """Render the current page scaled to fit the canvas (full-page view)."""
        try:
            self.preview_canvas.update_idletasks()
            cw = self.preview_canvas.winfo_width()
            ch = self.preview_canvas.winfo_height()
            if cw < 10:
                cw = _PREVIEW_RENDER_W
            if ch < 10:
                ch = 600

            doc = fitz.open(self.preview_pdf_path)
            self.preview_total_pages = len(doc)
            if self.preview_page >= self.preview_total_pages:
                self.preview_page = 0
            page = doc.load_page(self.preview_page)

            # Scale to fit canvas (maintain aspect ratio, show full page)
            scale = min(cw / page.rect.width, ch / page.rect.height)
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale))
            doc.close()

            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            photo = ImageTk.PhotoImage(img)

            # Centre the page on the canvas
            x = (cw - pix.width) // 2
            y = (ch - pix.height) // 2

            self.preview_canvas.configure(scrollregion=(0, 0, cw, ch))
            self.preview_canvas.delete("all")
            self.preview_canvas.create_rectangle(0, 0, cw, ch, fill="#404040", outline="")
            self.preview_canvas.create_image(x, y, anchor="nw", image=photo)
            self.preview_canvas._photo_ref = photo  # prevent GC

            self.preview_title_label.configure(text=os.path.basename(self.preview_pdf_path))
            self.preview_page_label.configure(text=f"Page {self.preview_page + 1} of {self.preview_total_pages}")

            # Update scrollbar thumb to reflect position within document
            if self.preview_total_pages <= 1:
                self.preview_vscroll.set(0.0, 1.0)
            else:
                start = self.preview_page / self.preview_total_pages
                end = (self.preview_page + 1) / self.preview_total_pages
                self.preview_vscroll.set(start, end)
        except Exception as e:
            self._preview_show_placeholder(f"Preview error:\n{e}")

    def _preview_show_placeholder(self, message="Select a row to preview"):
        """Show a centred message on the preview canvas."""
        self.preview_canvas.delete("all")
        cw = self.preview_canvas.winfo_width() or _PREVIEW_RENDER_W
        ch = self.preview_canvas.winfo_height() or 400
        self.preview_canvas.create_text(
            cw // 2, ch // 2,
            text=message, fill="gray", font=("", 11), justify="center"
        )
        self.preview_title_label.configure(text="No file selected")
        self.preview_page_label.configure(text="")
        self.preview_vscroll.set(0.0, 1.0)
        self.preview_pdf_path = None
        self.preview_total_pages = 0

    def _preview_scrollbar_command(self, *args):
        """Scrollbar drives page navigation, not pixel scrolling."""
        if not self.preview_total_pages:
            return
        cmd = args[0]
        if cmd == "moveto":
            fraction = float(args[1])
            page = int(fraction * self.preview_total_pages)
            page = max(0, min(page, self.preview_total_pages - 1))
        elif cmd == "scroll":
            delta = int(args[1])
            page = max(0, min(self.preview_page + delta, self.preview_total_pages - 1))
        else:
            return
        if page != self.preview_page:
            self.preview_page = page
            self._render_preview_page()

    def _preview_mousewheel(self, event):
        """Scroll wheel navigates pages."""
        if not self.preview_pdf_path:
            return
        if event.delta < 0:
            self._preview_page_next()
        else:
            self._preview_page_prev()

    def _on_preview_canvas_resize(self, event):
        """Re-render when the preview panel is resized."""
        if self.preview_pdf_path and self._preview_visible:
            self._render_preview_page()

    def _preview_page_prev(self):
        if self.preview_page > 0:
            self.preview_page -= 1
            self._render_preview_page()

    def _preview_page_next(self):
        if self.preview_page < self.preview_total_pages - 1:
            self.preview_page += 1
            self._render_preview_page()

    # =========================
    # LOGGING UTILITY
    # =========================

    def log(self, message):
        """Adds a message to the text box and forces a GUI update."""
        try:
            self.log_textbox.configure(state="normal")
            self.log_textbox.insert("end", f"{message}\n")
            self.log_textbox.see("end")
            self.log_textbox.configure(state="disabled")
            self.update()
        except Exception:
            pass  # Window may have been destroyed during scan

    # =========================
    # CORE OCR DETECTION
    # =========================

    def is_pdf_ocred(self, pdf_path):
        """
        Returns a tuple (is_ocr: bool, char_count: int).
        Counts ALL extractable text characters in the PDF.
        Returns (False, -1) on error.
        """
        try:
            doc = fitz.open(pdf_path)
            total_text = 0
            for page in doc:
                text = page.get_text("text")
                if text:
                    total_text += len(text.strip())
            doc.close()
            return (total_text >= self.MIN_TEXT_CHARS, total_text)
        except Exception as e:
            self.log(f"ERROR reading {pdf_path}: {e}")
            return (False, -1)

    # =========================
    # CANCEL / CLEANUP
    # =========================

    def cancel_scan(self):
        """Signal the scan loop to stop after the current file."""
        if self.scan_running:
            self.cancel_requested = True
            self.cancel_button.configure(state="disabled", text="Cancelling...")
            self.log("Cancel requested -- finishing current file...")

    def _scan_cleanup(self):
        """Reset UI state after scan completes or is cancelled."""
        self.scan_running = False
        self.cancel_requested = False
        self.run_button.configure(state="normal", text="START SCAN")
        self.cancel_button.configure(state="disabled", text="CANCEL")

        self._on_file_action_change()

    # =========================
    # OCR METHODS
    # =========================

    def _check_tesseract(self):
        try:
            self.log(f"DEBUG: TESSDATA_PREFIX={os.environ.get('TESSDATA_PREFIX')}")
            self.log(f"DEBUG: tesseract_cmd={pytesseract.pytesseract.tesseract_cmd}")

            # Check if the executable exists
            if _TESSERACT_EXE and os.path.exists(_TESSERACT_EXE):
                self.log(f"Tesseract found at: {_TESSERACT_EXE}")
                return True
            else:
                self.log("ERROR: Tesseract OCR engine not found.")
                self.log("  Tried locations:")
                self.log("    - C:\\Program Files\\Tesseract-OCR")
                self.log("    - C:\\Program Files (x86)\\Tesseract-OCR")
                self.log("    - System PATH")
                self.log("  Install from: https://github.com/UB-Mannheim/tesseract/wiki")
                self.log("  Then restart this application.")
                return False

        except Exception as e:
            self.log(f"ERROR: Could not verify Tesseract: {e}")
            self.log(f"  Tesseract path: {pytesseract.pytesseract.tesseract_cmd}")
            self.log("  https://github.com/UB-Mannheim/tesseract/wiki")
            return False

    def ocr_pdf(self, pdf_path, output_path, status_callback=None):
        """
        OCR a PDF using Tesseract.
        Returns tuple: (success: bool, error_message: str)
        Calls status_callback(page_num, total_pages) before each page if provided.
        """
        try:
            # Convert PDF pages to images
            poppler_path = _find_poppler_path()
            try:
                self.log(f"DEBUG: poppler_path={poppler_path}")
            except Exception:
                pass
            if poppler_path:
                images = convert_from_path(pdf_path, dpi=300, poppler_path=poppler_path)
            else:
                images = convert_from_path(pdf_path, dpi=300)
            total_pages = len(images)

            # OCR each page and collect PDF bytes
            pdf_bytes_list = []
            for page_num, image in enumerate(images, start=1):
                # Call status callback
                if status_callback:
                    try:
                        status_callback(page_num, total_pages)
                    except Exception:
                        pass  # Window may have been destroyed

                # OCR image to PDF bytes
                pdf_bytes = pytesseract.image_to_pdf_or_hocr(image, extension='pdf')
                pdf_bytes_list.append(pdf_bytes)

            # Merge all page PDFs using fitz
            merged_doc = fitz.open(stream=pdf_bytes_list[0], filetype="pdf")
            for page_bytes in pdf_bytes_list[1:]:
                temp_doc = fitz.open(stream=page_bytes, filetype="pdf")
                merged_doc.insert_pdf(temp_doc)
                temp_doc.close()

            # Save merged PDF
            merged_doc.save(output_path)
            merged_doc.close()

            return (True, "")
        except Exception as e:
            return (False, str(e))

    def _get_ocr_output_path(self, original_path):
        r"""
        Given a path like C:\dir\MyFile.pdf, return C:\dir\MyFile_aocrded.pdf
        If that exists, increment counter: _aocrded_1.pdf, _aocrded_2.pdf, etc.
        """
        base, ext = os.path.splitext(original_path)
        output_path = f"{base}_aocrded{ext}"

        if not os.path.exists(output_path):
            return output_path

        counter = 1
        while True:
            output_path = f"{base}_aocrded_{counter}{ext}"
            if not os.path.exists(output_path):
                return output_path
            counter += 1

    # =========================
    # MAIN SCAN LOGIC
    # =========================

    def start_scan(self):
        """The main scan function, run from the GUI button."""

        # Save config at scan start
        self._save_config()

        # Reset state
        self.last_csv_path = None
        self.view_report_button.configure(state="disabled")
        self.scan_running = True
        self.cancel_requested = False

        # Update button states
        self.run_button.configure(state="disabled", text="Scanning...")
        self.cancel_button.configure(state="normal")

        # Clear log and treeview
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Reset summary stats
        self.stats_total.configure(text="Total: 0")
        self.stats_ocr.configure(text="OCR-D: 0")
        self.stats_nonocr.configure(text="NON-OCR: 0")
        self.stats_errors.configure(text="Errors: 0")

        # --- Read GUI values ---
        try:
            self.MIN_TEXT_CHARS = int(self.threshold_entry.get())
        except ValueError:
            self.log("ERROR: Invalid threshold value. Using default of 50.")
            self.MIN_TEXT_CHARS = 50

        source_dir = self.source_entry.get()
        dest_dir = self.dest_entry.get()
        report_dir = self.report_entry.get()
        move_files = self.FILE_ACTION.get() == "move"
        enable_report = self.ENABLE_REPORT.get()
        recursive = self.RECURSIVE_SCAN.get()
        dry_run = self.DRY_RUN.get()
        ocr_in_place = self.FILE_ACTION.get() == "ocr"

        # Check Tesseract if OCR in place is selected
        if ocr_in_place:
            if not self._check_tesseract():
                self._scan_cleanup()
                return

        # --- Count PDF files first for progress bar ---
        self.log("Counting PDF files...")
        total_files = 0
        try:
            if recursive:
                for root, _, files in os.walk(source_dir):
                    for f in files:
                        if f.lower().endswith(".pdf"):
                            total_files += 1
            else:
                for f in os.listdir(source_dir):
                    if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(source_dir, f)):
                        total_files += 1
        except FileNotFoundError:
            self.log(f"FATAL ERROR: Source directory not found: {source_dir}")
            self._scan_cleanup()
            return

        self.log(f"Found {total_files} PDF file(s) to scan.")
        self.progress_bar.set(0)
        self.progress_label.configure(text=f"0 of {total_files} files scanned")
        self.ocr_page_status_label.configure(text="")

        # --- Log scan settings ---
        if dry_run:
            self.log("*** DRY RUN MODE -- No files will be moved or OCR'd ***")
        if ocr_in_place:
            if dry_run:
                self.log("*** OCR IN PLACE MODE (DRY RUN) -- Would OCR non-OCR files (no actual changes) ***")
            else:
                self.log("*** OCR IN PLACE MODE -- Non-OCR files will be OCR'd in their source directory ***")
        self.log(f"Starting scan of: {source_dir}")
        self.log(f"Recursive: {'Yes' if recursive else 'No'}")
        self.log(f"Move files: {move_files}")
        self.log(f"Report generation: {'ON' if enable_report else 'OFF'}")
        self.log(f"OCR Check Threshold: {self.MIN_TEXT_CHARS} characters.")

        # --- Prepare CSV ---
        log_data = []
        csv_header = [
            "Timestamp", "Filename", "Full Path", "OCR Status", "Char Count",
            "Action", "Final Destination", "Error",
            "OCR Phase Status", "OCR Output File", "OCR Verify Status"
        ]

        if move_files and not dry_run:
            os.makedirs(dest_dir, exist_ok=True)
            self.log(f"Destination folder created/checked: {dest_dir}")

        if enable_report:
            os.makedirs(report_dir, exist_ok=True)
            self.log(f"Report folder created/checked: {report_dir}")

        non_ocr_files = []
        scanned_count = 0
        ocr_count = 0
        error_count = 0
        cancelled = False
        log_data_map = {}       # full_path -> file_log dict reference
        self.tree_item_map = {}    # reset for this scan
        self.tree_iid_to_path = {} # reset reverse map
        ocr_success_count = 0
        ocr_fail_count = 0

        # --- Build file iterator based on recursive setting ---
        if recursive:
            file_iterator = (
                (root, filename)
                for root, _, files in os.walk(source_dir)
                for filename in files
                if filename.lower().endswith(".pdf")
            )
        else:
            try:
                entries = os.listdir(source_dir)
            except FileNotFoundError:
                self.log(f"FATAL ERROR: Source directory not found: {source_dir}")
                self._scan_cleanup()
                return
            file_iterator = (
                (source_dir, filename)
                for filename in entries
                if filename.lower().endswith(".pdf")
                and os.path.isfile(os.path.join(source_dir, filename))
            )

        # --- Main scan loop ---
        try:
            for root, filename in file_iterator:
                # Cancel check
                if self.cancel_requested:
                    cancelled = True
                    self.log("\n*** SCAN CANCELLED BY USER ***")
                    break

                scanned_count += 1
                full_path = os.path.join(root, filename)
                file_log = {
                    "Timestamp": datetime.now().strftime("%H:%M:%S"),
                    "Filename": filename,
                    "Full Path": full_path,
                    "Error": "",
                    "OCR Phase Status": "", "OCR Output File": "", "OCR Verify Status": ""
                }

                # OCR check (returns tuple)
                is_ocr, char_count = self.is_pdf_ocred(full_path)

                ocr_status = "OCR-D" if is_ocr else "NON-OCR"
                action = "Kept"
                final_dest = "N/A"
                char_display = f"{char_count:,}" if char_count >= 0 else "Error"

                if not is_ocr:
                    non_ocr_files.append(full_path)
                    self.log(f"--> NON-OCR FOUND: {filename} ({char_display} chars)")

                    if move_files:
                        if dry_run:
                            action = "Would Move (Dry Run)"
                            dest_path = os.path.join(dest_dir, filename)
                            final_dest = dest_path
                            self.log(f"    [DRY RUN] Would move to {dest_dir}")
                        else:
                            action = "Moved"
                            dest_path = os.path.join(dest_dir, filename)
                            # Auto-rename if duplicate exists
                            if os.path.exists(dest_path):
                                name, ext = os.path.splitext(filename)
                                counter = 1
                                while os.path.exists(dest_path):
                                    dest_path = os.path.join(dest_dir, f"{name}_{counter}{ext}")
                                    counter += 1
                            final_dest = dest_path
                            try:
                                shutil.move(full_path, dest_path)
                                self.log(f"    (Moved to {dest_dir})")
                            except Exception as e:
                                action = "Error Moving"
                                file_log["Error"] = str(e)
                                self.log(f"    (ERROR moving file: {e})")
                else:
                    ocr_count += 1
                    self.log(f"OCR-D: {filename} ({char_display} chars)")

                # Track errors
                if char_count < 0:
                    error_count += 1

                # Update summary stats bar
                self.stats_total.configure(text=f"Total: {scanned_count:,}")
                self.stats_ocr.configure(text=f"OCR-D: {ocr_count:,}")
                self.stats_nonocr.configure(text=f"NON-OCR: {len(non_ocr_files):,}")
                self.stats_errors.configure(text=f"Errors: {error_count:,}")

                # Insert into treeview with alternating row colors
                try:
                    row_tag = "even" if scanned_count % 2 == 0 else "odd"
                    iid = self.tree.insert("", "end", values=(filename, ocr_status, char_display, action), tags=(row_tag,))
                    self.tree_item_map[full_path] = iid
                    self.tree_iid_to_path[iid] = full_path
                except Exception:
                    pass  # Window may have been destroyed

                # Compile CSV row
                file_log["OCR Status"] = ocr_status
                file_log["Char Count"] = char_display
                file_log["Action"] = action
                file_log["Final Destination"] = final_dest
                log_data.append(file_log)
                log_data_map[full_path] = file_log

                # Update progress
                if total_files > 0:
                    self.progress_bar.set(scanned_count / total_files)
                self.progress_label.configure(text=f"{scanned_count} of {total_files} files scanned")
                self.update()

        except FileNotFoundError:
            self.log(f"\nFATAL ERROR: Source directory not found: {source_dir}")
        except Exception as e:
            self.log(f"\nAN UNEXPECTED ERROR OCCURRED: {e}")

        # --- Phase 2: OCR In Place ---
        if ocr_in_place and not cancelled and not dry_run and len(non_ocr_files) > 0:
            self.log("\n=========================")
            self.log("PHASE 2: OCR IN PLACE")
            self.log("=========================")
            self.progress_bar.set(0)
            self.run_button.configure(text="OCR Phase...")

            for ocr_index, orig_path in enumerate(non_ocr_files):
                # Cancel check
                if self.cancel_requested:
                    cancelled = True
                    self.log("\n*** OCR PHASE CANCELLED BY USER ***")
                    break

                filename = os.path.basename(orig_path)
                dirname = os.path.dirname(orig_path)

                # Update treeview
                if orig_path in self.tree_item_map:
                    iid = self.tree_item_map[orig_path]
                    self.tree.item(iid, values=(filename, "OCRing...", "", "Pending OCR"))

                # Update status label
                self.ocr_page_status_label.configure(text=f"OCRing: {dirname} | {filename} | Starting...")
                self.update()

                # Get output path
                output_path = self._get_ocr_output_path(orig_path)

                # Define page callback with default args to avoid late-binding
                def _page_callback(page_num, total_pages, dir_name=dirname, file_name=filename):
                    self.ocr_page_status_label.configure(text=f"OCRing: {dir_name} | {file_name} | Page {page_num} of {total_pages}")
                    self.update()

                # Run OCR
                ocr_ok, ocr_error = self.ocr_pdf(orig_path, output_path, _page_callback)

                ocr_phase_status = ""
                ocr_output_file = ""
                ocr_verify_status = ""
                action_result = "Pending OCR"

                if ocr_ok:
                    # Verify the output
                    verify_ok, verify_chars = self.is_pdf_ocred(output_path)
                    if verify_ok:
                        ocr_success_count += 1
                        ocr_phase_status = "OCR Success"
                        action_result = "OCR'd in place"
                        ocr_output_file = output_path
                        ocr_verify_status = f"Verified - {verify_chars:,} chars"
                    else:
                        ocr_fail_count += 1
                        ocr_phase_status = "OCR Failed"
                        action_result = "OCR Failed"
                        ocr_verify_status = f"Verification failed - {verify_chars:,} chars"
                        try:
                            os.remove(output_path)
                        except Exception:
                            pass
                        output_path = ""
                else:
                    ocr_fail_count += 1
                    ocr_phase_status = "OCR Failed"
                    action_result = "OCR Failed"
                    ocr_verify_status = f"Error: {ocr_error}"
                    output_path = ""

                # Update treeview row
                if orig_path in self.tree_item_map:
                    iid = self.tree_item_map[orig_path]
                    self.tree.item(iid, values=(filename, ocr_phase_status, "", action_result))

                # Update log data map
                if orig_path in log_data_map:
                    log_data_map[orig_path]["OCR Phase Status"] = ocr_phase_status
                    log_data_map[orig_path]["OCR Output File"] = ocr_output_file
                    log_data_map[orig_path]["OCR Verify Status"] = ocr_verify_status

                # Update progress
                if len(non_ocr_files) > 0:
                    self.progress_bar.set((ocr_index + 1) / len(non_ocr_files))
                self.progress_label.configure(text=f"{ocr_index + 1} of {len(non_ocr_files)} files OCR'd | Failed: {ocr_fail_count}")
                self.update()

            self.ocr_page_status_label.configure(text="")

        elif ocr_in_place and len(non_ocr_files) == 0:
            self.log("No non-OCR files found. Nothing to OCR.")

        # --- Save CSV Log (Conditional) ---
        if enable_report:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            csv_filename = os.path.join(report_dir, f"pdf_scan_log_{timestamp}.csv")

            try:
                with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=csv_header)
                    writer.writeheader()
                    writer.writerows(log_data)

                self.last_csv_path = csv_filename
                self.view_report_button.configure(state="normal")

                self.log("\n=========================")
                self.log(f"CSV log successfully saved to: {csv_filename}")
            except Exception as e:
                self.log(f"\nERROR saving CSV log to {report_dir}: {e}")
                self.last_csv_path = None

        # --- Summary ---
        self.log("\nSUMMARY")
        self.log("=========================")
        self.log(f"Total files scanned: {scanned_count}" + (" (cancelled)" if cancelled else ""))
        self.log(f"Total non-OCR PDFs found: {len(non_ocr_files)}")
        if dry_run:
            self.log("*** This was a DRY RUN -- no files were moved or OCR'd ***")
        elif move_files:
            self.log(f"Files moved to: {dest_dir}")
        if ocr_in_place:
            if dry_run:
                self.log(f"OCR In Place (DRY RUN): Would have OCR'd {len(non_ocr_files)} files")
            else:
                self.log(f"OCR In Place: {ocr_success_count} succeeded, {ocr_fail_count} failed.")
        self.log("Scan complete.")

        self._scan_cleanup()


if __name__ == "__main__":
    app = App()
    app.mainloop()
