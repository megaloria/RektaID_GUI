"""
RektaPrint - ID Photo Processor v5
====================================
Changes from v4:
  - Removed background color feature
  - Fixed drag & drop → replaced with proper Add Photos button (file dialog)
  - Face auto-center kept and improved
  - Added output format choice: PDF / Word (.docx) / Both

Requirements:
  pip install Pillow reportlab opencv-python-headless python-docx
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json, os, tempfile, threading, shutil
from PIL import Image
import cv2
import numpy as np
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from docx import Document as DocxDocument
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─── CONFIG ────────────────────────────────────────────────────────────────────
INPUT_FOLDER   = "input_photos"
OUTPUT_FOLDER  = "output_sheets"
PACKAGES_FILE  = "packages.json"
DPI            = 300
SUPPORTED      = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}

PAPER_SIZES = {
    "A5  (5.83 x 8.27 in)": (5.83, 8.27),
    "4R  (4 x 6 in)":       (4.0,  6.0),
}
DEFAULT_PAPER = "A5  (5.83 x 8.27 in)"

SIZES = {
    "1x1 inch":         (1.0,       1.0),
    "2x2 inch":         (2.0,       2.0),
    "Passport 35x45mm": (35/25.4,   45/25.4),
    "Passport 40x50mm": (40/25.4,   50/25.4),
    "Wallet 2x3 inch":  (2.0,       3.0),
}

# Face detection settings
FACE_CASCADE    = cv2.CascadeClassifier(
    cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
)
FACE_TOP_RATIO  = 0.12
FACE_SIZE_RATIO = 0.50

# Colors
BG      = "#1a1a2e"
PANEL   = "#16213e"
ACCENT  = "#e94560"
ACCENT2 = "#0f3460"
TEXT    = "#eaeaea"
SUBTEXT = "#a0a0b0"
WHITE   = "#ffffff"

# ─── PACKAGE STORAGE ──────────────────────────────────────────────────────────

def load_packages():
    if os.path.exists(PACKAGES_FILE):
        try:
            with open(PACKAGES_FILE) as f:
                return json.load(f)
        except: pass
    return {
        "Package A": [{"size": "2x2 inch",         "qty": 6}],
        "Package B": [{"size": "2x2 inch",         "qty": 4},
                      {"size": "1x1 inch",         "qty": 8}],
        "Package C": [{"size": "Passport 35x45mm", "qty": 6}],
    }

def save_packages(packages):
    with open(PACKAGES_FILE, "w") as f:
        json.dump(packages, f, indent=2)

# ─── IMAGE PROCESSING ─────────────────────────────────────────────────────────

def face_crop(pil_img: Image.Image, w_px: int, h_px: int) -> Image.Image:
    """Detect face and crop with face centered. Falls back to center crop."""
    img_np = np.array(pil_img.convert("RGB"))
    gray   = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    faces  = FACE_CASCADE.detectMultiScale(
        gray, scaleFactor=1.1, minNeighbors=5, minSize=(60, 60)
    )
    if len(faces) == 0:
        return _center_crop(pil_img, w_px, h_px)

    fx, fy, fw, fh = max(faces, key=lambda r: r[2]*r[3])
    face_cx = fx + fw // 2
    face_ty = fy

    scale = (h_px * FACE_SIZE_RATIO) / fh
    orig_w, orig_h = pil_img.size
    new_w = int(orig_w * scale)
    new_h = int(orig_h * scale)

    if new_w < w_px or new_h < h_px:
        return _center_crop(pil_img, w_px, h_px)

    resized    = pil_img.resize((new_w, new_h), Image.LANCZOS)
    s_face_cx  = int(face_cx * scale)
    s_face_ty  = int(face_ty * scale)
    crop_x     = max(0, min(s_face_cx - w_px // 2, new_w - w_px))
    crop_y     = max(0, min(s_face_ty - int(h_px * FACE_TOP_RATIO), new_h - h_px))
    return resized.crop((crop_x, crop_y, crop_x + w_px, crop_y + h_px))


def _center_crop(pil_img: Image.Image, w_px: int, h_px: int) -> Image.Image:
    ow, oh = pil_img.size
    if ow/oh > w_px/h_px:
        nw, nh = int(ow * h_px/oh), h_px
    else:
        nw, nh = w_px, int(oh * w_px/ow)
    img = pil_img.resize((nw, nh), Image.LANCZOS)
    l = (nw - w_px) // 2
    t = (nh - h_px) // 2
    return img.crop((l, t, l+w_px, t+h_px))


def crop_for_slot(pil_img, size_key):
    w_in, h_in = SIZES.get(size_key, (2.0, 2.0))
    w_px = int(w_in * DPI)
    h_px = int(h_in * DPI)
    return _center_crop(pil_img, w_px, h_px), w_in, h_in

# ─── PDF GENERATION ───────────────────────────────────────────────────────────

def generate_pdf(photo_path, pkg_slots, out_path, paper_key=None):
    if paper_key is None:
        paper_key = DEFAULT_PAPER
    pw_in, ph_in = PAPER_SIZES.get(paper_key, PAPER_SIZES[DEFAULT_PAPER])
    P_W = pw_in * inch
    P_H = ph_in * inch

    pil_img = Image.open(photo_path).convert("RGB")
    c = canvas.Canvas(out_path, pagesize=(P_W, P_H))
    c.setLineWidth(0.5)
    c.setStrokeColorRGB(0.5, 0.5, 0.5)

    y_top = P_H

    for slot in pkg_slots:
        cropped, w_in, h_in = crop_for_slot(pil_img, slot["size"])
        qty  = slot["qty"]
        w_pt = w_in * inch
        h_pt = h_in * inch

        cols = max(1, int(P_W / w_pt))
        rows = -(-qty // cols)

        tmp = tempfile.NamedTemporaryFile(suffix=".jpg", delete=False)
        cropped.save(tmp.name, "JPEG", quality=95)
        tmp.close()

        count = 0
        for row in range(rows):
            items_in_row = min(cols, qty - row * cols)
            row_w    = items_in_row * w_pt
            x_start  = (P_W - row_w) / 2   # center each row
            y_bottom = y_top - (row + 1) * h_pt
            for col in range(items_in_row):
                if count >= qty: break
                x = x_start + col * w_pt
                c.drawImage(tmp.name, x, y_bottom, width=w_pt, height=h_pt)
                c.rect(x, y_bottom, w_pt, h_pt)
                count += 1

        y_top -= rows * h_pt
        os.unlink(tmp.name)

    c.save()

# ─── WORD GENERATION ──────────────────────────────────────────────────────────

def generate_docx(photo_path, pkg_slots, out_path, paper_key=None):
    if paper_key is None:
        paper_key = DEFAULT_PAPER
    pw_in, ph_in = PAPER_SIZES.get(paper_key, PAPER_SIZES[DEFAULT_PAPER])

    pil_img = Image.open(photo_path).convert("RGB")
    doc = DocxDocument()

    # Set page size and zero margins
    section = doc.sections[0]
    section.page_width  = Inches(pw_in)
    section.page_height = Inches(ph_in)
    section.top_margin    = Inches(0)
    section.bottom_margin = Inches(0)
    section.left_margin   = Inches(0)
    section.right_margin  = Inches(0)

    first_group = True

    for slot in pkg_slots:
        cropped, w_in, h_in = crop_for_slot(pil_img, slot["size"])
        qty  = slot["qty"]
        cols = max(1, int(pw_in / w_in))
        rows = -(-qty // cols)

        # Save cropped image to temp file
        tmp = tempfile.NamedTemporaryFile(suffix=".jpg", delete=False)
        cropped.save(tmp.name, "JPEG", quality=95)
        tmp.close()

        if not first_group:
            # Add a tiny spacing paragraph between groups
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after  = Pt(0)
        first_group = False

        count = 0
        for row in range(rows):
            items_in_row = min(cols, qty - row * cols)

            # Create a table with exact number of cells for this row
            tbl = doc.add_table(rows=1, cols=items_in_row)
            tbl.style = "Table Grid"

            # Set table properties
            tbl_pr = tbl._tbl.find(qn("w:tblPr"))
            if tbl_pr is None:
                tbl_pr = OxmlElement("w:tblPr")
                tbl._tbl.insert(0, tbl_pr)

            # Center the table on page
            tbl_jc = OxmlElement("w:jc")
            tbl_jc.set(qn("w:val"), "center")
            tbl_pr.append(tbl_jc)

            # Set exact table width
            tbl_w = OxmlElement("w:tblW")
            tbl_w.set(qn("w:w"), str(int(items_in_row * w_in * 1440)))
            tbl_w.set(qn("w:type"), "dxa")
            tbl_pr.append(tbl_w)

            # Remove cell spacing
            tbl_cell_sp = OxmlElement("w:tblCellSpacing")
            tbl_cell_sp.set(qn("w:w"), "0")
            tbl_cell_sp.set(qn("w:type"), "dxa")
            tbl_pr.append(tbl_cell_sp)

            for col_idx, cell in enumerate(tbl.rows[0].cells):
                # Set exact cell width
                tc_pr = cell._tc.get_or_add_tcPr()
                tc_w = OxmlElement("w:tcW")
                tc_w.set(qn("w:w"), str(int(w_in * 1440)))
                tc_w.set(qn("w:type"), "dxa")
                tc_pr.append(tc_w)

                # Zero cell margins
                tc_mar = OxmlElement("w:tcMar")
                for side in ["top", "left", "bottom", "right"]:
                    m = OxmlElement(f"w:{side}")
                    m.set(qn("w:w"), "0")
                    m.set(qn("w:type"), "dxa")
                    tc_mar.append(m)
                tc_pr.append(tc_mar)

                if count < qty:
                    # Add cut border
                    tc_borders = OxmlElement("w:tcBorders")
                    for side in ["top", "left", "bottom", "right"]:
                        b = OxmlElement(f"w:{side}")
                        b.set(qn("w:val"), "single")
                        b.set(qn("w:sz"), "4")
                        b.set(qn("w:color"), "999999")
                        tc_borders.append(b)
                    tc_pr.append(tc_borders)

                    # Add image
                    para = cell.paragraphs[0]
                    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    para.paragraph_format.space_before = Pt(0)
                    para.paragraph_format.space_after  = Pt(0)
                    run = para.add_run()
                    run.add_picture(tmp.name, width=Inches(w_in), height=Inches(h_in))
                    count += 1
                else:
                    # Empty cell — no border, no content
                    tc_borders = OxmlElement("w:tcBorders")
                    for side in ["top", "left", "bottom", "right"]:
                        b = OxmlElement(f"w:{side}")
                        b.set(qn("w:val"), "none")
                        b.set(qn("w:sz"), "0")
                        b.set(qn("w:color"), "auto")
                        tc_borders.append(b)
                    tc_pr.append(tc_borders)

        os.unlink(tmp.name)

    doc.save(out_path)

# ─── GUI ──────────────────────────────────────────────────────────────────────

class RektaPrintApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RektaPrint — ID Photo Processor v5")
        self.root.geometry("820x660")
        self.root.configure(bg=BG)
        self.root.resizable(False, False)

        self.packages       = load_packages()
        self.selected_pkg   = tk.StringVar()
        self.selected_paper = tk.StringVar(value=DEFAULT_PAPER)
        self.output_format  = tk.StringVar(value="PDF")
        self.auto_clear_var = tk.BooleanVar(value=True)
        self.status_var     = tk.StringVar(value="Ready.")

        self._build_ui()
        self._refresh_package_list()

    # ── UI BUILDER ────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self.root, bg=ACCENT, height=52)
        hdr.pack(fill="x")
        tk.Label(hdr, text="RektaPrint  ·  ID Photo Processor  v5",
                 font=("Segoe UI", 15, "bold"), bg=ACCENT, fg=WHITE).pack(side="left", padx=18, pady=12)

        main = tk.Frame(self.root, bg=BG)
        main.pack(fill="both", expand=True, padx=16, pady=12)

        # LEFT — packages
        left = tk.Frame(main, bg=PANEL, width=230)
        left.pack(side="left", fill="y", padx=(0,10))
        left.pack_propagate(False)

        tk.Label(left, text="PACKAGES", font=("Segoe UI", 9, "bold"),
                 bg=PANEL, fg=SUBTEXT).pack(anchor="w", padx=12, pady=(12,4))

        self.pkg_listbox = tk.Listbox(
            left, bg=PANEL, fg=TEXT, selectbackground=ACCENT,
            selectforeground=WHITE, font=("Segoe UI", 11),
            relief="flat", bd=0, highlightthickness=0,
            activestyle="none", cursor="hand2"
        )
        self.pkg_listbox.pack(fill="both", expand=True, padx=8)
        self.pkg_listbox.bind("<<ListboxSelect>>", self._on_pkg_select)

        btn_row = tk.Frame(left, bg=PANEL)
        btn_row.pack(fill="x", padx=8, pady=8)
        self._btn(btn_row, "＋ New",  self._new_package,  ACCENT).pack(side="left", expand=True, fill="x", padx=(0,3))
        self._btn(btn_row, "✎ Edit",  self._edit_package, ACCENT2).pack(side="left", expand=True, fill="x", padx=(3,0))
        self._btn(left, "🗑 Delete", self._delete_package, "#3a1a1a", fg=ACCENT).pack(fill="x", padx=8, pady=(0,10))

        # RIGHT
        right = tk.Frame(main, bg=BG)
        right.pack(side="left", fill="both", expand=True)

        # Package detail card
        detail_card = tk.Frame(right, bg=PANEL)
        detail_card.pack(fill="x", pady=(0,10))
        self.pkg_title = tk.Label(detail_card, text="← Select a package",
                                  font=("Segoe UI", 13, "bold"), bg=PANEL, fg=TEXT)
        self.pkg_title.pack(anchor="w", padx=14, pady=(10,2))
        self.pkg_detail = tk.Label(detail_card, text="",
                                   font=("Segoe UI", 10), bg=PANEL, fg=SUBTEXT, justify="left")
        self.pkg_detail.pack(anchor="w", padx=14, pady=(0,10))

        # Photos section header
        photos_hdr = tk.Frame(right, bg=BG)
        photos_hdr.pack(fill="x")
        tk.Label(photos_hdr, text="INPUT PHOTOS", font=("Segoe UI", 9, "bold"),
                 bg=BG, fg=SUBTEXT).pack(side="left")

        # Add photos + open folder buttons
        btn_row2 = tk.Frame(right, bg=BG)
        btn_row2.pack(fill="x", pady=(4,4))
        self._btn(btn_row2, "＋ Add Photos", self._add_photos, ACCENT,
                  pady=5).pack(side="left", padx=(0,6))
        self._btn(btn_row2, "Open Folder", self._open_input_folder,
                  ACCENT2, pady=5).pack(side="left", padx=(0,6))
        self._btn(btn_row2, "↺ Refresh", self._refresh_photos,
                  ACCENT2, pady=5).pack(side="left")

        # Photos list
        self.photos_box = tk.Listbox(
            right, bg=PANEL, fg=TEXT, font=("Segoe UI", 10),
            relief="flat", bd=0, highlightthickness=0,
            height=6, selectbackground=ACCENT2
        )
        self.photos_box.pack(fill="x", pady=(0,6))

        # Auto-clear checkbox
        tk.Checkbutton(
            right, text="Auto-clear input folder after processing",
            variable=self.auto_clear_var,
            bg=BG, fg=TEXT, selectcolor=ACCENT2,
            activebackground=BG, activeforeground=TEXT,
            font=("Segoe UI", 9), cursor="hand2"
        ).pack(anchor="w", pady=(0,8))

        # Paper size
        paper_row = tk.Frame(right, bg=BG)
        paper_row.pack(fill="x", pady=(0,6))
        tk.Label(paper_row, text="PAPER", font=("Segoe UI", 9, "bold"),
                 bg=BG, fg=SUBTEXT).pack(side="left", padx=(0,10))
        for key in PAPER_SIZES:
            tk.Radiobutton(
                paper_row, text=key, variable=self.selected_paper, value=key,
                bg=BG, fg=TEXT, selectcolor=ACCENT2,
                activebackground=BG, activeforeground=TEXT,
                font=("Segoe UI", 10), cursor="hand2"
            ).pack(side="left", padx=6)

        # Output format
        fmt_row = tk.Frame(right, bg=BG)
        fmt_row.pack(fill="x", pady=(0,8))
        tk.Label(fmt_row, text="OUTPUT", font=("Segoe UI", 9, "bold"),
                 bg=BG, fg=SUBTEXT).pack(side="left", padx=(0,10))
        for fmt in ["PDF", "Word (.docx)", "Both"]:
            tk.Radiobutton(
                fmt_row, text=fmt, variable=self.output_format, value=fmt,
                bg=BG, fg=TEXT, selectcolor=ACCENT2,
                activebackground=BG, activeforeground=TEXT,
                font=("Segoe UI", 10), cursor="hand2"
            ).pack(side="left", padx=6)

        # Process button
        self.process_btn = self._btn(
            right, "▶  PRINT THIS PACKAGE",
            self._process, ACCENT, pady=10,
            font=("Segoe UI", 12, "bold")
        )
        self.process_btn.pack(fill="x", pady=(0,4))

        # Status bar
        status_bar = tk.Frame(self.root, bg=ACCENT2, height=28)
        status_bar.pack(fill="x", side="bottom")
        tk.Label(status_bar, textvariable=self.status_var,
                 font=("Segoe UI", 9), bg=ACCENT2, fg=TEXT).pack(side="left", padx=12, pady=4)

        self._refresh_photos()

    def _btn(self, parent, text, cmd, color, fg=WHITE, padx=10, pady=6, font=("Segoe UI", 10)):
        return tk.Button(parent, text=text, command=cmd,
                         bg=color, fg=fg, relief="flat",
                         activebackground=color, activeforeground=fg,
                         font=font, padx=padx, pady=pady,
                         cursor="hand2", bd=0)

    # ── PHOTOS ────────────────────────────────────────────────────────────────

    def _add_photos(self):
        """Open file dialog to select photos and copy to input folder."""
        files = filedialog.askopenfilenames(
            title="Select ID Photos",
            filetypes=[
                ("Image files", "*.jpg *.jpeg *.png *.webp *.bmp"),
                ("All files", "*.*")
            ]
        )
        if not files:
            return
        os.makedirs(INPUT_FOLDER, exist_ok=True)
        copied = 0
        for fpath in files:
            ext = os.path.splitext(fpath)[1].lower()
            if ext in SUPPORTED:
                dest = os.path.join(INPUT_FOLDER, os.path.basename(fpath))
                shutil.copy2(fpath, dest)
                copied += 1
        self._refresh_photos()
        self.status_var.set(f"Added {copied} photo(s) to input folder.")

    def _refresh_photos(self):
        os.makedirs(INPUT_FOLDER, exist_ok=True)
        files = [f for f in os.listdir(INPUT_FOLDER)
                 if os.path.splitext(f)[1].lower() in SUPPORTED]
        self.photos_box.delete(0, "end")
        if files:
            for f in files:
                self.photos_box.insert("end", f"  {f}")
            self.status_var.set(f"{len(files)} photo(s) ready.")
        else:
            self.photos_box.insert("end", "  No photos — click '＋ Add Photos' to select")
            self.status_var.set("No photos found. Click '＋ Add Photos'.")

    def _open_input_folder(self):
        os.makedirs(INPUT_FOLDER, exist_ok=True)
        os.startfile(os.path.abspath(INPUT_FOLDER))

    # ── PACKAGE LIST ──────────────────────────────────────────────────────────

    def _refresh_package_list(self):
        self.pkg_listbox.delete(0, "end")
        for name in self.packages:
            self.pkg_listbox.insert("end", f"  {name}")
        if self.packages:
            self.pkg_listbox.select_set(0)
            self._on_pkg_select()

    def _on_pkg_select(self, event=None):
        sel = self.pkg_listbox.curselection()
        if not sel: return
        name  = self.pkg_listbox.get(sel[0]).strip()
        self.selected_pkg.set(name)
        slots = self.packages.get(name, [])
        self.pkg_title.config(text=name)
        self.pkg_detail.config(
            text="\n".join(f"  • {s['qty']} pcs  —  {s['size']}" for s in slots) or "No sizes defined."
        )

    # ── PACKAGE EDITOR ────────────────────────────────────────────────────────

    def _new_package(self):    self._open_editor(None)
    def _edit_package(self):
        name = self.selected_pkg.get()
        if not name:
            messagebox.showwarning("No Selection", "Select a package to edit.")
            return
        self._open_editor(name)

    def _delete_package(self):
        name = self.selected_pkg.get()
        if not name:
            messagebox.showwarning("No Selection", "Select a package to delete.")
            return
        if messagebox.askyesno("Delete", f'Delete "{name}"?'):
            del self.packages[name]
            save_packages(self.packages)
            self.selected_pkg.set("")
            self.pkg_title.config(text="← Select a package")
            self.pkg_detail.config(text="")
            self._refresh_package_list()

    def _open_editor(self, edit_name):
        win = tk.Toplevel(self.root)
        win.title("Edit Package" if edit_name else "New Package")
        win.geometry("420x420")
        win.configure(bg=BG)
        win.grab_set()

        tk.Label(win, text="Package Name", font=("Segoe UI", 10, "bold"),
                 bg=BG, fg=TEXT).pack(anchor="w", padx=16, pady=(16,2))
        name_var = tk.StringVar(value=edit_name or "")
        tk.Entry(win, textvariable=name_var, font=("Segoe UI", 11),
                 bg=PANEL, fg=TEXT, insertbackground=TEXT,
                 relief="flat", bd=6).pack(fill="x", padx=16)

        tk.Label(win, text="Sizes & Quantities", font=("Segoe UI", 10, "bold"),
                 bg=BG, fg=TEXT).pack(anchor="w", padx=16, pady=(14,4))

        slot_frame = tk.Frame(win, bg=BG)
        slot_frame.pack(fill="both", expand=True, padx=16)

        slots    = []
        existing = self.packages.get(edit_name, []) if edit_name else []

        def add_slot(size_val="2x2 inch", qty_val=1):
            row = tk.Frame(slot_frame, bg=BG)
            row.pack(fill="x", pady=3)
            size_var = tk.StringVar(value=size_val)
            ttk.Combobox(row, textvariable=size_var, values=list(SIZES.keys()),
                         state="readonly", width=20,
                         font=("Segoe UI", 10)).pack(side="left")
            tk.Label(row, text="  Qty:", bg=BG, fg=SUBTEXT,
                     font=("Segoe UI", 10)).pack(side="left")
            qty_var = tk.IntVar(value=qty_val)
            tk.Spinbox(row, from_=1, to=99, textvariable=qty_var, width=4,
                       font=("Segoe UI", 10), bg=PANEL, fg=TEXT,
                       relief="flat", buttonbackground=PANEL,
                       insertbackground=TEXT).pack(side="left", padx=4)
            def remove():
                slots.remove((size_var, qty_var))
                row.destroy()
            tk.Button(row, text="✕", command=remove, bg=BG, fg=ACCENT,
                      relief="flat", font=("Segoe UI", 10),
                      activebackground=BG, cursor="hand2").pack(side="left")
            slots.append((size_var, qty_var))

        for s in existing:
            add_slot(s["size"], s["qty"])
        if not existing:
            add_slot()

        self._btn(win, "＋ Add Size", add_slot, ACCENT2, pady=4).pack(fill="x", padx=16, pady=6)

        def save():
            n = name_var.get().strip()
            if not n:
                messagebox.showwarning("Required", "Enter a package name.", parent=win)
                return
            if not slots:
                messagebox.showwarning("Required", "Add at least one size.", parent=win)
                return
            if edit_name and edit_name != n:
                del self.packages[edit_name]
            self.packages[n] = [{"size": sv.get(), "qty": qv.get()} for sv, qv in slots]
            save_packages(self.packages)
            self._refresh_package_list()
            for i, key in enumerate(self.packages):
                if key == n:
                    self.pkg_listbox.select_clear(0, "end")
                    self.pkg_listbox.select_set(i)
                    self.pkg_listbox.see(i)
                    self.selected_pkg.set(n)
                    self._on_pkg_select()
                    break
            self.status_var.set(f'Package "{n}" saved.')
            win.destroy()

        self._btn(win, "💾  Save Package", save, ACCENT,
                  pady=8, font=("Segoe UI", 11, "bold")).pack(fill="x", padx=16, pady=10)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TCombobox", fieldbackground=PANEL, background=PANEL,
                        foreground=TEXT, selectbackground=ACCENT2)

    # ── PROCESS ───────────────────────────────────────────────────────────────

    def _process(self):
        pkg_name = self.selected_pkg.get()
        if not pkg_name or pkg_name not in self.packages:
            messagebox.showwarning("No Package", "Select a package first.")
            return

        files = [f for f in os.listdir(INPUT_FOLDER)
                 if os.path.splitext(f)[1].lower() in SUPPORTED]
        if not files:
            messagebox.showwarning("No Photos",
                'No photos found. Click "＋ Add Photos" to select photos.')
            return

        self.process_btn.config(state="disabled", text="Processing…")
        self.status_var.set("Processing… please wait.")

        paper_key  = self.selected_paper.get()
        fmt        = self.output_format.get()
        auto_clear = self.auto_clear_var.get()
        slots      = self.packages[pkg_name]

        def run():
            os.makedirs(OUTPUT_FOLDER, exist_ok=True)
            done, errors = 0, 0
            for fname in files:
                fpath      = os.path.join(INPUT_FOLDER, fname)
                basename   = os.path.splitext(fname)[0]
                cust_dir   = os.path.join(OUTPUT_FOLDER, basename)
                os.makedirs(cust_dir, exist_ok=True)
                safe_pkg   = pkg_name.replace("/", "-").replace("\\", "-")
                safe_paper = "A5" if "A5" in paper_key else "4R"
                base_out   = os.path.join(cust_dir, f"{basename}_{safe_pkg}_{safe_paper}")

                try:
                    if fmt in ("PDF", "Both"):
                        generate_pdf(fpath, slots, base_out + ".pdf", paper_key)
                    if fmt in ("Word (.docx)", "Both"):
                        generate_docx(fpath, slots, base_out + ".docx", paper_key)
                    done += 1
                except Exception as e:
                    errors += 1
                    print(f"Error {fname}: {e}")

            if auto_clear:
                for fname in files:
                    try: os.remove(os.path.join(INPUT_FOLDER, fname))
                    except: pass

            def finish():
                self.process_btn.config(state="normal", text="▶  PRINT THIS PACKAGE")
                msg = f"✅ Done! {done} file(s) saved."
                if errors: msg += f" ({errors} error(s))"
                if auto_clear: msg += " · Input cleared."
                self.status_var.set(msg)
                self._refresh_photos()
                if messagebox.askyesno("Done!", f"{msg}\n\nOpen output folder?"):
                    os.startfile(os.path.abspath(OUTPUT_FOLDER))

            self.root.after(0, finish)

        threading.Thread(target=run, daemon=True).start()


# ─── ENTRY POINT ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    os.makedirs(INPUT_FOLDER,  exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    root = tk.Tk()
    app  = RektaPrintApp(root)
    root.mainloop()