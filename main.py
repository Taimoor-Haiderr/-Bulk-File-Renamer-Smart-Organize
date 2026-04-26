
import os
import sys
import shutil
import csv
import json
import datetime
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────
TYPE_MAP = {
    "Images":    {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".svg", ".tiff"},
    "PDFs":      {".pdf"},
    "Videos":    {".mp4", ".mkv", ".avi", ".mov", ".wmv", ".flv"},
    "Audio":     {".mp3", ".wav", ".flac", ".aac", ".ogg"},
    "Documents": {".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt", ".odt"},
    "Archives":  {".zip", ".rar", ".tar", ".gz", ".7z"},
    "Code":      {".py", ".js", ".html", ".css", ".java", ".cpp", ".c", ".ts"},
}

C = {
    "bg":         "#f0f4ff",
    "sidebar":    "#2563eb",
    "sidebar_dk": "#1d4ed8",
    "panel":      "#ffffff",
    "card":       "#f8faff",
    "accent":     "#2563eb",
    "accent_lt":  "#dbeafe",
    "text":       "#1e293b",
    "subtext":    "#64748b",
    "success":    "#16a34a",
    "success_lt": "#dcfce7",
    "warning":    "#d97706",
    "error":      "#dc2626",
    "error_lt":   "#fee2e2",
    "border":     "#e2e8f0",
    "entry":      "#f8fafc",
    "white":      "#ffffff",
    "drag_bg":    "#eff6ff",
    "drag_bd":    "#93c5fd",
}

FONT_BOLD  = ("Segoe UI", 10, "bold")
FONT_REG   = ("Segoe UI", 9)
FONT_SM    = ("Segoe UI", 8)
FONT_LG    = ("Segoe UI", 12, "bold")
FONT_TITLE = ("Segoe UI", 16, "bold")

# ─────────────────────────────────────────────
# CORE LOGIC
# ─────────────────────────────────────────────

def _now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def get_files(folder: str) -> list:
    try:
        return [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
    except PermissionError:
        raise PermissionError(f"Permission denied: {folder}")
    except FileNotFoundError:
        raise FileNotFoundError(f"Folder not found: {folder}")

def build_new_name(original, prefix, suffix, find, replace, counter, use_numbering, use_date):
    name, ext = os.path.splitext(original)
    if find:
        name = name.replace(find, replace)
    if use_date:
        name = f"{name}_{datetime.date.today().strftime('%Y%m%d')}"
    if use_numbering:
        name = f"{name}_{counter}"
    return f"{prefix}{name}{suffix}{ext}"

def safe_name(folder: str, name: str) -> str:
    base, ext = os.path.splitext(name)
    candidate, i = name, 1
    while os.path.exists(os.path.join(folder, candidate)):
        candidate = f"{base}_({i}){ext}"
        i += 1
    return candidate

def preview_renames(folder, files, options):
    results, counter = [], int(options.get("start_num", 1))
    for f in files:
        new = build_new_name(f, options.get("prefix",""), options.get("suffix",""),
                             options.get("find",""), options.get("replace",""),
                             counter, options.get("use_numbering", False), options.get("use_date", False))
        new = safe_name(folder, new) if new != f else new
        results.append((f, new))
        if options.get("use_numbering"): counter += 1
    return results

def apply_renames(folder, pairs):
    log = []
    for old, new in pairs:
        src, dst = os.path.join(folder, old), os.path.join(folder, new)
        try:
            if old == new: continue
            if not os.path.exists(src): raise FileNotFoundError(f"Missing: {src}")
            os.rename(src, dst)
            log.append({"action":"rename","old":old,"new":new,"status":"ok","time":_now()})
        except Exception as e:
            log.append({"action":"rename","old":old,"new":new,"status":f"ERROR: {e}","time":_now()})
    return log

def get_file_type(filename):
    ext = Path(filename).suffix.lower()
    for cat, exts in TYPE_MAP.items():
        if ext in exts: return cat
    return "Others"

def _move_file(folder, f, dest_dir):
    src = os.path.join(folder, f)
    os.makedirs(dest_dir, exist_ok=True)
    dst_name = safe_name(dest_dir, f)
    shutil.move(src, os.path.join(dest_dir, dst_name))
    return dst_name

def organize_by_type(folder, files):
    log = []
    for f in files:
        cat = get_file_type(f)
        dest_dir = os.path.join(folder, cat)
        try:
            dst_name = _move_file(folder, f, dest_dir)
            log.append({"action":"organize","old":f,"new":os.path.join(cat,dst_name),"status":"ok","time":_now()})
        except Exception as e:
            log.append({"action":"organize","old":f,"new":"","status":f"ERROR:{e}","time":_now()})
    return log

def organize_by_date(folder, files):
    log = []
    for f in files:
        src = os.path.join(folder, f)
        try:
            date_folder = datetime.datetime.fromtimestamp(os.path.getmtime(src)).strftime("%Y-%m")
            dest_dir = os.path.join(folder, date_folder)
            dst_name = _move_file(folder, f, dest_dir)
            log.append({"action":"organize_date","old":f,"new":os.path.join(date_folder,dst_name),"status":"ok","time":_now()})
        except Exception as e:
            log.append({"action":"organize_date","old":f,"new":"","status":f"ERROR:{e}","time":_now()})
    return log

def organize_by_size(folder, files):
    BINS = [(1_000_000,"Small_under1MB"),(10_000_000,"Medium_1to10MB"),
            (100_000_000,"Large_10to100MB"),(float("inf"),"Huge_over100MB")]
    log = []
    for f in files:
        src = os.path.join(folder, f)
        try:
            cat = next(label for limit,label in BINS if os.path.getsize(src) < limit)
            dest_dir = os.path.join(folder, cat)
            dst_name = _move_file(folder, f, dest_dir)
            log.append({"action":"organize_size","old":f,"new":os.path.join(cat,dst_name),"status":"ok","time":_now()})
        except Exception as e:
            log.append({"action":"organize_size","old":f,"new":"","status":f"ERROR:{e}","time":_now()})
    return log

def save_log(log_entries, folder):
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path  = os.path.join(folder, f"organizer_log_{ts}.csv")
    json_path = os.path.join(folder, f"organizer_log_{ts}.json")
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["action","old","new","status","time"])
        w.writeheader(); w.writerows(log_entries)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(log_entries, f, indent=2)
    return csv_path, json_path

def open_folder(path):
    """Open folder in system file explorer."""
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass

# ─────────────────────────────────────────────
# GUI HELPERS
# ─────────────────────────────────────────────

def flat_btn(parent, text, cmd, bg, fg="#ffffff", padx=14, pady=6, font=FONT_BOLD):
    btn = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                    activebackground=C["sidebar_dk"], activeforeground="#fff",
                    relief="flat", bd=0, font=font, padx=padx, pady=pady, cursor="hand2")
    return btn

def label(parent, text, fg=None, font=FONT_REG, bg=None):
    return tk.Label(parent, text=text, fg=fg or C["text"], font=font, bg=bg or C["panel"])

def section_card(parent, title, icon=""):
    frame = tk.Frame(parent, bg=C["panel"], bd=0)
    frame.pack(fill="x", padx=0, pady=(0, 10))
    hdr = tk.Frame(frame, bg=C["accent_lt"], padx=12, pady=8)
    hdr.pack(fill="x")
    tk.Label(hdr, text=f"{icon}  {title}", bg=C["accent_lt"],
             fg=C["accent"], font=FONT_BOLD).pack(side="left")
    body = tk.Frame(frame, bg=C["panel"], padx=14, pady=10)
    body.pack(fill="x")
    return body

# ─────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("File Renamer & Smart Organizer  v2.0")
        self.geometry("1060x700")
        self.minsize(900, 580)
        self.configure(bg=C["bg"])
        self.resizable(True, True)

        self.folder        = tk.StringVar()
        self.log_entries   = []
        self.undo_stack    = []
        self.preview_pairs = []

        # Rename options
        self.prefix_var   = tk.StringVar()
        self.suffix_var   = tk.StringVar()
        self.find_var     = tk.StringVar()
        self.replace_var  = tk.StringVar()
        self.numbering_var= tk.BooleanVar()
        self.date_var     = tk.BooleanVar()
        self.start_num_var= tk.IntVar(value=1)
        self.org_mode     = tk.StringVar(value="type")
        self.open_after   = tk.BooleanVar(value=True)   # ← open folder after action

        self._build_styles()
        self._build_ui()

    # ── Styles ────────────────────────────────
    def _build_styles(self):
        s = ttk.Style(self); s.theme_use("clam")
        s.configure("TNotebook",     background=C["bg"],   borderwidth=0)
        s.configure("TNotebook.Tab", background=C["border"], foreground=C["subtext"],
                    padding=[16, 8], font=FONT_BOLD)
        s.map("TNotebook.Tab",
              background=[("selected", C["accent"])],
              foreground=[("selected", "#ffffff")])
        s.configure("TFrame",    background=C["bg"])
        s.configure("Treeview",  background=C["white"], foreground=C["text"],
                    fieldbackground=C["white"], rowheight=26, font=FONT_REG)
        s.configure("Treeview.Heading", background=C["accent_lt"],
                    foreground=C["accent"], font=FONT_BOLD, relief="flat")
        s.map("Treeview", background=[("selected", C["accent_lt"])],
              foreground=[("selected", C["accent"])])
        s.configure("TScrollbar", background=C["border"],
                    troughcolor=C["bg"], arrowcolor=C["accent"])
        s.configure("TCheckbutton", background=C["panel"], foreground=C["text"],
                    font=FONT_REG)

    # ── UI Layout ─────────────────────────────
    def _build_ui(self):
        # ── Sidebar ──
        sidebar = tk.Frame(self, bg=C["sidebar"], width=220)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        self._build_sidebar(sidebar)

        # ── Main area ──
        main = tk.Frame(self, bg=C["bg"])
        main.pack(side="left", fill="both", expand=True)
        self._build_topbar(main)
        self._build_notebook(main)
        self._build_statusbar(main)

    # ── Sidebar ───────────────────────────────
    def _build_sidebar(self, sb):
        # Logo
        logo = tk.Frame(sb, bg=C["sidebar_dk"], pady=20)
        logo.pack(fill="x")
        tk.Label(logo, text="⚡", bg=C["sidebar_dk"], fg="#ffffff",
                 font=("Segoe UI", 28)).pack()
        tk.Label(logo, text="File Organizer", bg=C["sidebar_dk"], fg="#ffffff",
                 font=("Segoe UI", 11, "bold")).pack()
        tk.Label(logo, text="v2.0", bg=C["sidebar_dk"], fg="#93c5fd",
                 font=FONT_SM).pack()

        # Drag & Drop zone
        dnd_outer = tk.Frame(sb, bg=C["sidebar"], pady=12, padx=12)
        dnd_outer.pack(fill="x")
        self.dnd_zone = tk.Label(
            dnd_outer,
            text="\n\nDrag & Drop\nFolder Here\n\nor click Browse",
            bg=C["drag_bg"], fg=C["accent"],
            font=("Segoe UI", 9), relief="flat",
            bd=0, pady=16, padx=10,
            justify="center", cursor="hand2"
        )
        self.dnd_zone.pack(fill="x", ipady=8)
        # Make drag-drop zone look like a dashed border via a frame trick
        dnd_border = tk.Frame(dnd_outer, bg=C["drag_bd"], padx=2, pady=2)
        dnd_border.pack(fill="x")
        self.dnd_inner = tk.Label(
            dnd_border,
            text="Drop folder here\n     or click Browse",
            bg=C["drag_bg"], fg=C["accent"],
            font=("Segoe UI", 9), pady=14, padx=10,
            justify="center", cursor="hand2", wraplength=160
        )
        self.dnd_inner.pack(fill="x")

        # Bind click to browse on the drop zone
        for w in (self.dnd_zone, self.dnd_inner, dnd_border):
            w.bind("<Button-1>", lambda e: self._browse_folder())
            w.bind("<Enter>",    lambda e: self.dnd_inner.config(bg="#dbeafe"))
            w.bind("<Leave>",    lambda e: self.dnd_inner.config(bg=C["drag_bg"]))

        # Enable native drag-drop if OS supports it
        self._setup_dnd(self.dnd_inner)
        self.dnd_zone.pack_forget()   # hide the first one, keep the bordered one

        tk.Frame(sb, bg=C["sidebar_dk"], height=1).pack(fill="x", pady=4)

        # Folder path display
        fp_frame = tk.Frame(sb, bg=C["sidebar"], padx=10, pady=4)
        fp_frame.pack(fill="x")
        tk.Label(fp_frame, text="Selected Folder:", bg=C["sidebar"],
                 fg="#bfdbfe", font=FONT_SM).pack(anchor="w")
        self.folder_lbl = tk.Label(fp_frame, text="None selected", bg=C["sidebar"],
                                   fg="#ffffff", font=FONT_SM, wraplength=190, justify="left")
        self.folder_lbl.pack(anchor="w", pady=(2, 6))

        # Sidebar action buttons
        btn_frame = tk.Frame(sb, bg=C["sidebar"], padx=10)
        btn_frame.pack(fill="x")

        for text, cmd, bg in [
            (" Browse Folder",  self._browse_folder,  C["sidebar_dk"]),
            ("  Undo Last",       self._undo,           "#7c3aed"),
            ("  Save Log",        self._save_log,       "#16a34a"),
        ]:
            flat_btn(btn_frame, text, cmd, bg, padx=8, pady=8).pack(fill="x", pady=3)

        # Open after toggle
        oa_frame = tk.Frame(sb, bg=C["sidebar"], padx=10, pady=8)
        oa_frame.pack(fill="x")
        tk.Checkbutton(oa_frame, text=" Open folder after action",
                       variable=self.open_after,
                       bg=C["sidebar"], fg="#ffffff",
                       selectcolor=C["sidebar_dk"],
                       activebackground=C["sidebar"],
                       activeforeground="#ffffff",
                       font=FONT_SM).pack(anchor="w")

        # Stats footer
        self.stats_lbl = tk.Label(sb, text="Files: 0", bg=C["sidebar_dk"],
                                  fg="#93c5fd", font=FONT_SM, pady=6)
        self.stats_lbl.pack(side="bottom", fill="x")

    def _setup_dnd(self, widget):

        if sys.platform != "win32":
            return
        self.after(200, self._register_drop_target)

    def _register_drop_target(self):

        try:
            import ctypes, ctypes.wintypes

            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            ctypes.windll.shell32.DragAcceptFiles(hwnd, True)

            WM_DROPFILES = 0x0233
            WndProcType  = ctypes.WINFUNCTYPE(
                ctypes.c_long, ctypes.wintypes.HWND,
                ctypes.c_uint, ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM
            )
            old_proc = ctypes.windll.user32.GetWindowLongPtrW(hwnd, -4)

            def wndproc(hwnd_, msg, wparam, lparam):
                if msg == WM_DROPFILES:
                    hDrop = wparam
                    n = ctypes.windll.shell32.DragQueryFileW(hDrop, 0xFFFFFFFF, None, 0)
                    for i in range(n):
                        buf = ctypes.create_unicode_buffer(260)
                        ctypes.windll.shell32.DragQueryFileW(hDrop, i, buf, 260)
                        p = buf.value
                        if os.path.isdir(p):
                            self.after(0, lambda x=p: self._load_folder(x))
                        elif os.path.isfile(p):
                            self.after(0, lambda x=os.path.dirname(p): self._load_folder(x))
                        break
                    ctypes.windll.shell32.DragFinish(hDrop)
                    return 0
                return ctypes.windll.user32.CallWindowProcW(old_proc, hwnd_, msg, wparam, lparam)

            self._wndproc = WndProcType(wndproc)   # keep reference alive
            ctypes.windll.user32.SetWindowLongPtrW(hwnd, -4, self._wndproc)
        except Exception:
            pass

    # ── Top bar ───────────────────────────────
    def _build_topbar(self, parent):
        bar = tk.Frame(parent, bg=C["white"], pady=10, padx=16,
                       relief="flat", bd=0)
        bar.pack(fill="x")
        tk.Label(bar, text="Bulk File Renamer & Smart Organizer",
                 bg=C["white"], fg=C["text"], font=FONT_LG).pack(side="left")
        tk.Label(bar, text="Rename · Organize · Preview · Log",
                 bg=C["white"], fg=C["subtext"], font=FONT_SM).pack(side="left", padx=12)
        # Separator
        tk.Frame(parent, bg=C["border"], height=1).pack(fill="x")

    # ── Notebook ──────────────────────────────
    def _build_notebook(self, parent):
        self.nb = ttk.Notebook(parent)
        self.nb.pack(fill="both", expand=True, padx=12, pady=8)
        self._tab_rename()
        self._tab_organize()
        self._tab_preview()
        self._tab_log()

    # ── Status bar ────────────────────────────
    def _build_statusbar(self, parent):
        bar = tk.Frame(parent, bg=C["accent_lt"], pady=5, padx=14)
        bar.pack(fill="x", side="bottom")
        self.status_var = tk.StringVar(value="Ready — drag a folder or click Browse to begin.")
        tk.Label(bar, textvariable=self.status_var, bg=C["accent_lt"],
                 fg=C["accent"], font=FONT_SM, anchor="w").pack(side="left")

    # ── Tab: Rename ───────────────────────────
    def _tab_rename(self):
        outer = tk.Frame(self.nb, bg=C["bg"])
        self.nb.add(outer, text="    Rename Files  ")

        scroll_canvas = tk.Canvas(outer, bg=C["bg"], highlightthickness=0)
        scroll_canvas.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=scroll_canvas.yview)
        vsb.pack(side="right", fill="y")
        scroll_canvas.configure(yscrollcommand=vsb.set)

        content = tk.Frame(scroll_canvas, bg=C["bg"])
        win = scroll_canvas.create_window((0,0), window=content, anchor="nw")

        def on_configure(e):
            scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))
        def on_canvas_resize(e):
            scroll_canvas.itemconfig(win, width=e.width)
        content.bind("<Configure>", on_configure)
        scroll_canvas.bind("<Configure>", on_canvas_resize)

        # ── Options card ──
        opt = section_card(content, "Rename Options", )

        def row(parent):
            r = tk.Frame(parent, bg=C["panel"]); r.pack(fill="x", pady=4); return r

        def field(parent, lbl, var, w=16):
            tk.Label(parent, text=lbl, bg=C["panel"], fg=C["subtext"],
                     font=FONT_SM, width=9, anchor="e").pack(side="left", padx=(0,4))
            e = tk.Entry(parent, textvariable=var, width=w, bg=C["entry"],
                         fg=C["text"], insertbackground=C["accent"],
                         relief="flat", font=FONT_REG, bd=1,
                         highlightthickness=1, highlightbackground=C["border"],
                         highlightcolor=C["accent"])
            e.pack(side="left", padx=(0,18))

        r1 = row(opt); field(r1,"Prefix:", self.prefix_var); field(r1,"Suffix:", self.suffix_var)
        r2 = row(opt); field(r2,"Find:",   self.find_var);   field(r2,"Replace:", self.replace_var)

        r3 = row(opt)
        tk.Checkbutton(r3, text="Auto Numbering", variable=self.numbering_var,
                       bg=C["panel"], fg=C["text"], selectcolor=C["accent_lt"],
                       activebackground=C["panel"], font=FONT_REG).pack(side="left", padx=(0,6))
        tk.Label(r3, text="Start:", bg=C["panel"], fg=C["subtext"], font=FONT_SM).pack(side="left")
        tk.Spinbox(r3, textvariable=self.start_num_var, from_=1, to=9999, width=5,
                   bg=C["entry"], fg=C["text"], relief="flat",
                   buttonbackground=C["border"], font=FONT_REG).pack(side="left", padx=(2,18))
        tk.Checkbutton(r3, text="Append Today's Date (YYYYMMDD)",
                       variable=self.date_var,
                       bg=C["panel"], fg=C["text"], selectcolor=C["accent_lt"],
                       activebackground=C["panel"], font=FONT_REG).pack(side="left")

        # ── Action buttons ──
        act = tk.Frame(content, bg=C["bg"], pady=6, padx=0)
        act.pack(fill="x")
        flat_btn(act, "👁  Preview Changes", self._do_preview, C["accent"]).pack(side="left", padx=(0,8))
        flat_btn(act, "▶  Apply Rename",     self._do_rename,  C["success"]).pack(side="left")

        # ── File list card ──
        lc = section_card(content, "Files in Folder", )
        self.file_tree = self._make_tree(lc, ("Filename","Type","Size"),
                                         widths=[420,100,90])
        self.file_tree.pack(fill="both", expand=True)

        # Double-click to open file location
        self.file_tree.bind("<Double-1>", lambda e: self._open_selected_file())

    # ── Tab: Organize ─────────────────────────
    def _tab_organize(self):
        outer = tk.Frame(self.nb, bg=C["bg"])
        self.nb.add(outer, text="   Organize Files  ")

        top = section_card(outer, "Organization Mode", )
        tk.Label(top, text="Choose how files will be sorted into subfolders:",
                 bg=C["panel"], fg=C["subtext"], font=FONT_REG).pack(anchor="w", pady=(0,10))

        modes = [
            ("type",  "  By File Type",  "Images → Images/,  PDFs → PDFs/,  Videos → Videos/ …"),
            ("date",  "  By Date",        "Sorted by last-modified date  →  2025-01/,  2026-04/ …"),
            ("size",  "  By File Size",   "Small (<1MB)  ·  Medium  ·  Large  ·  Huge (>100MB)"),
        ]
        for val, lbl, desc in modes:
            f = tk.Frame(top, bg=C["card"], relief="flat", bd=1, pady=8, padx=12)
            f.pack(fill="x", pady=4)
            tk.Radiobutton(f, text=lbl, variable=self.org_mode, value=val,
                           bg=C["card"], fg=C["text"], selectcolor=C["accent_lt"],
                           activebackground=C["card"], font=FONT_BOLD).pack(side="left")
            tk.Label(f, text=desc, bg=C["card"], fg=C["subtext"],
                     font=FONT_SM).pack(side="left", padx=12)

        flat_btn(outer, "▶  Organize Now  →  Opens Folder When Done",
                 self._do_organize, C["success"], pady=10).pack(anchor="w", padx=14, pady=10)

        # Category map
        cm = section_card(outer, "Category Map", )
        for cat, exts in TYPE_MAP.items():
            row = tk.Frame(cm, bg=C["panel"]); row.pack(fill="x", pady=1)
            tk.Label(row, text=f"{cat:12s}", bg=C["accent_lt"], fg=C["accent"],
                     font=FONT_BOLD, width=12, padx=6).pack(side="left")
            tk.Label(row, text="  →  " + "  ·  ".join(sorted(exts)),
                     bg=C["panel"], fg=C["subtext"], font=FONT_SM).pack(side="left")

    # ── Tab: Preview ──────────────────────────
    def _tab_preview(self):
        outer = tk.Frame(self.nb, bg=C["bg"])
        self.nb.add(outer, text="  👁  Preview  ")

        hdr = tk.Frame(outer, bg=C["accent_lt"], padx=14, pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Before → After rename preview",
                 bg=C["accent_lt"], fg=C["accent"], font=FONT_BOLD).pack(side="left")
        self.preview_count = tk.Label(hdr, text="", bg=C["accent_lt"],
                                      fg=C["subtext"], font=FONT_SM)
        self.preview_count.pack(side="right")

        self.preview_tree = self._make_tree(outer, ("Original Name","New Name"),
                                            widths=[440,440], pack_fill="both")
        self.preview_tree.pack(fill="both", expand=True, padx=12, pady=8)

    # ── Tab: Log ──────────────────────────────
    def _tab_log(self):
        outer = tk.Frame(self.nb, bg=C["bg"])
        self.nb.add(outer, text="   Log  ")

        bar = tk.Frame(outer, bg=C["bg"], pady=6, padx=12)
        bar.pack(fill="x")
        flat_btn(bar, "  Clear Log", self._clear_log, C["error"]).pack(side="left")
        flat_btn(bar, "  Open Folder", lambda: open_folder(self.folder.get()),
                 C["accent"]).pack(side="left", padx=8)

        self.log_tree = self._make_tree(
            outer,
            ("Time","Action","Old Name","New Name","Status"),
            widths=[130, 90, 220, 220, 110],
            pack_fill="both"
        )
        self.log_tree.pack(fill="both", expand=True, padx=12, pady=(0,8))

    # ── Helpers ───────────────────────────────
    def _make_tree(self, parent, columns, widths=None, pack_fill=None):
        frame = tk.Frame(parent, bg=C["bg"])
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
        for i, col in enumerate(columns):
            w = widths[i] if widths else 160
            tree.heading(col, text=col)
            tree.column(col, width=w, minwidth=60)
        sb_y = ttk.Scrollbar(frame, orient="vertical",   command=tree.yview)
        sb_x = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.pack(side="right",  fill="y")
        sb_x.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        # Alternating row colors
        tree.tag_configure("odd",  background="#f8faff")
        tree.tag_configure("even", background=C["white"])
        return tree

    def _insert_tree(self, tree, values, tags=()):
        n = len(tree.get_children())
        row_tag = "odd" if n % 2 == 0 else "even"
        all_tags = list(tags) + [row_tag]
        tree.insert("", "end", values=values, tags=all_tags)

    def _set_status(self, msg):
        self.status_var.set(msg)

    def _assert_folder(self):
        f = self.folder.get().strip()
        if not f:
            messagebox.showwarning("No Folder", "Please select or drop a folder first.")
            return None
        if not os.path.isdir(f):
            messagebox.showerror("Not Found", f"Folder not found:\n{f}")
            return None
        return f

    def _load_folder(self, path):
        self.folder.set(path)
        short = path if len(path) <= 28 else "…" + path[-26:]
        self.folder_lbl.config(text=short)
        self._refresh_file_list(path)
        self._set_status(f"Folder loaded: {path}")
        # Highlight the drop zone green briefly
        self.dnd_inner.config(bg=C["success_lt"], fg=C["success"])
        self.after(1200, lambda: self.dnd_inner.config(bg=C["drag_bg"], fg=C["accent"]))

    def _browse_folder(self):
        path = filedialog.askdirectory(title="Select Working Folder")
        if path:
            self._load_folder(path)

    def _refresh_file_list(self, folder):
        for row in self.file_tree.get_children():
            self.file_tree.delete(row)
        try:
            files = get_files(folder)
            for i, f in enumerate(files):
                fpath = os.path.join(folder, f)
                size  = os.path.getsize(fpath)
                size_s = (f"{size/1_000_000:.2f} MB" if size>=1_000_000
                          else f"{size/1000:.1f} KB" if size>=1000
                          else f"{size} B")
                tag = "odd" if i%2==0 else "even"
                self.file_tree.insert("", "end",
                                      values=(f, get_file_type(f), size_s), tags=(tag,))
            self.stats_lbl.config(text=f"Files: {len(files)}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _get_options(self):
        return {"prefix": self.prefix_var.get(), "suffix": self.suffix_var.get(),
                "find":   self.find_var.get(),   "replace": self.replace_var.get(),
                "use_numbering": self.numbering_var.get(),
                "use_date":      self.date_var.get(),
                "start_num":     self.start_num_var.get()}

    def _open_selected_file(self):
        sel = self.file_tree.selection()
        if sel:
            fname = self.file_tree.item(sel[0])["values"][0]
            fpath = os.path.join(self.folder.get(), str(fname))
            open_folder(os.path.dirname(fpath))

    # ── Actions ───────────────────────────────
    def _do_preview(self):
        folder = self._assert_folder()
        if not folder: return
        try:
            files = get_files(folder)
            pairs = preview_renames(folder, files, self._get_options())
            self.preview_pairs = pairs
            for row in self.preview_tree.get_children():
                self.preview_tree.delete(row)
            changed = 0
            for i, (old, new) in enumerate(pairs):
                if old != new:
                    tags = ("changed", "odd" if i%2==0 else "even")
                    changed += 1
                else:
                    tags = ("same",    "odd" if i%2==0 else "even")
                self.preview_tree.insert("", "end", values=(old, new), tags=tags)
            self.preview_tree.tag_configure("changed", foreground=C["success"])
            self.preview_tree.tag_configure("same",    foreground=C["subtext"])
            self.preview_count.config(text=f"{changed} files will change")
            self.nb.select(2)
            self._set_status(f"Preview: {changed}/{len(pairs)} files will be renamed.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _do_rename(self):
        folder = self._assert_folder()
        if not folder: return
        try:
            files  = get_files(folder)
            pairs  = preview_renames(folder, files, self._get_options())
            changed= [(o,n) for o,n in pairs if o != n]
            if not changed:
                messagebox.showinfo("Nothing to do", "No files match the rename settings.")
                return
            if not messagebox.askyesno("Confirm Rename", f"Rename {len(changed)} file(s)?"):
                return
            log = apply_renames(folder, changed)
            for e in log:
                if e["status"] == "ok":
                    self.undo_stack.append((
                        os.path.join(folder, e["new"]),
                        os.path.join(folder, e["old"])
                    ))
            self._add_log(log)
            self._refresh_file_list(folder)
            ok  = sum(1 for e in log if e["status"]=="ok")
            err = len(log) - ok
            self._set_status(f"Done: {ok} renamed, {err} error(s).")
            if self.open_after.get():
                open_folder(folder)
        except Exception as ex:
            messagebox.showerror("Error", str(ex))

    def _do_organize(self):
        folder = self._assert_folder()
        if not folder: return
        mode = self.org_mode.get()
        try:
            files = get_files(folder)
            if not files:
                messagebox.showinfo("Empty", "No files found in the folder.")
                return
            if not messagebox.askyesno("Confirm Organize",
                                       f"Move {len(files)} file(s) into subfolders by {mode}?"):
                return
            if   mode == "type": log = organize_by_type(folder, files)
            elif mode == "date": log = organize_by_date(folder, files)
            else:                log = organize_by_size(folder, files)

            for e in log:
                if e["status"] == "ok":
                    self.undo_stack.append((
                        os.path.join(folder, e["new"]),
                        os.path.join(folder, e["old"])
                    ))
            self._add_log(log)
            self._refresh_file_list(folder)
            ok = sum(1 for e in log if e["status"]=="ok")
            self._set_status(f"Organized: {ok}/{len(log)} files moved.")

            # ── Open the result folder ──
            if self.open_after.get():
                open_folder(folder)

        except Exception as ex:
            messagebox.showerror("Error", str(ex))

    def _undo(self):
        if not self.undo_stack:
            messagebox.showinfo("Nothing to undo", "No operations to undo.")
            return
        new_path, old_path = self.undo_stack.pop()
        try:
            if not os.path.exists(new_path):
                messagebox.showerror("Undo Failed", f"File not found:\n{new_path}")
                return
            os.makedirs(os.path.dirname(old_path), exist_ok=True)
            os.rename(new_path, old_path)
            self._add_log([{"action":"undo","old":new_path,"new":old_path,
                            "status":"ok","time":_now()}])
            folder = self.folder.get().strip()
            if folder: self._refresh_file_list(folder)
            self._set_status(f"Undo OK: {os.path.basename(new_path)}")
        except Exception as ex:
            messagebox.showerror("Undo Error", str(ex))

    def _save_log(self):
        if not self.log_entries:
            messagebox.showinfo("Empty Log", "Nothing logged yet.")
            return
        # Save next to the script file (same folder as file_organizer_v2.py)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        try:
            csv_p, json_p = save_log(self.log_entries, script_dir)
            messagebox.showinfo("Log Saved",
                                f"Saved next to the script:\n\nCSV : {csv_p}\nJSON: {json_p}")
        except Exception as ex:
            messagebox.showerror("Save Error", str(ex))

    def _clear_log(self):
        self.log_entries.clear()
        for row in self.log_tree.get_children():
            self.log_tree.delete(row)
        self._set_status("Log cleared.")

    def _add_log(self, entries):
        self.log_entries.extend(entries)
        for e in entries:
            tag = "ok" if e["status"]=="ok" else "err"
            self._insert_tree(self.log_tree,
                              (e["time"], e["action"], e["old"], e["new"], e["status"]),
                              tags=(tag,))
        self.log_tree.tag_configure("ok",  foreground=C["success"])
        self.log_tree.tag_configure("err", foreground=C["error"])
        kids = self.log_tree.get_children()
        if kids: self.log_tree.see(kids[-1])


# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()