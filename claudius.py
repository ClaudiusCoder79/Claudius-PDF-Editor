"""
Claudius PDF Editor
Windows 10 x64 (also works on macOS / Linux)

Requirements:  Python 3.9+   and   pip install pymupdf pillow
               (optional, for PDF<->Word/Excel/PowerPoint:  pip install python-docx openpyxl python-pptx)
Keep claudius_extra.py in the same folder as this file.
Run:           python claudius.py   (or double-click run.bat)
"""
import copy
import io
import math
import webbrowser
import os
import sys
import time
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk, filedialog, messagebox, simpledialog, colorchooser

# PyMuPDF 2.x tries to load an optional AI "layout" add-on (onnx models, several 100 MB) when it starts. This program
# does not use it, and inside the EXE its model files are not packed, so it is switched off here.
for _m in ("pymupdf.layout", "pymupdf_layout", "pymupdf4llm"):
    sys.modules.setdefault(_m, None)

try:
    import pymupdf as fitz          # newer PyMuPDF
except ImportError:                 # older PyMuPDF
    import fitz

try:
    from PIL import Image, ImageTk, ImageOps, ImageEnhance, ImageChops, ImageGrab
except ImportError:
    sys.exit("Pillow is missing. Run:  python -m pip install pillow pymupdf")

import claudius_i18n as i18n
from claudius_i18n import tr
i18n.init()                      # interface language (File > Language); must run before any window exists
import claudius_theme as TH
TH.init()                        # light / dark look (File > Appearance); also before any window exists

try:
    from claudius_extra import (safe_zoom, safe_dpi, SCALE, S, ExtraTools, EXTRA_HINTS, EXTRA_TOOLS, RECT_TOOLS, HL_COLORS, CONVERT_KINDS)
except ImportError:
    sys.exit("claudius_extra.py is missing - keep it in the same folder as claudius.py")

APP_NAME = "Claudius PDF Editor"
APP_VERSION = "3.0"
APP_URL = "https://github.com/ClaudiusCoder79/Claudius-PDF"
APP_FREEWARE = "Free for personal and commercial use, source on GitHub"
MAX_UNDO = 25
WIN_ARIAL = r"C:\Windows\Fonts\arial.ttf"

# ---- look & feel -----------------------------------------------------------
from claudius_theme import round_rect, rr_points, PillTab
from claudius_features import FeatureTools, page_label
from claudius_view import ViewTools
from claudius_v3 import V3Tools
from claudius_v3b import V3bTools
from claudius_pro import ProTools
ACCENT = TH.ACCENT
ACCENT_HOVER = TH.ACCENT_HOVER
RIBBON_BG = TH.CARD
RIBBON_LINE = TH.LINE_SOFT
SIDEBAR_BG = TH.BG
WORK_BG = TH.WORK
BAR_BG = TH.NAV_BG
ACTIVE_BG = TH.TINT_DOWN
HOVER_BG = TH.TINT
UI_FONT = TH.UI_FONT
ICON_FONT = "Segoe UI Symbol"

SIDEBAR_W = S(216)
THUMB_W, THUMB_H, CELL_H = S(124), S(160), S(196)
PAD = S(24)


RES = getattr(Image, "Resampling", Image)
LANCZOS, BICUBIC, BILINEAR = RES.LANCZOS, RES.BICUBIC, RES.BILINEAR

# output quality for photos:  name -> (max dpi, JPEG quality)
QUALITY = {
    "Original": (None, 95),
    "High (300 dpi)": (300, 90),
    "Medium (200 dpi)": (200, 80),
    "Small file (120 dpi)": (120, 70),
}
# ---- embedded logo (PNG, base64) - no extra files needed at run time
LOGO_26_B64 = (
    "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"
)
LOGO_32_B64 = (
    "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"
)
LOGO_128_B64 = (
    "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"
)

IMG_TYPES = [("Images", "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.webp"), ("All files", "*.*")]


def load_image(path):
    """Open a photo, honour the phone's EXIF rotation, return RGBA."""
    im = Image.open(path)
    im = ImageOps.exif_transpose(im)
    return im.convert("RGBA")


class FloatImage:
    """A photo that floats above the page until it is applied (flattened)."""
    _next = 1

    def __init__(self, img, name=""):
        self.id = FloatImage._next
        FloatImage._next += 1
        self.name = name
        self.pristine = img
        self.set_base(img)
        self.cx = self.cy = 0.0
        self.w = self.h = 100.0
        self.reset_adjust()
        self.border = False

    def set_base(self, img):
        self.base = img
        m = max(img.size)
        if m > 1600:
            s = 1600 / m
            self.work = img.resize((max(1, int(img.width * s)), max(1, int(img.height * s))), LANCZOS)
        else:
            self.work = img
        self.pv_key = None
        self.pv_photo = None

    def reset_adjust(self):
        self.angle = 0.0
        self.flip_h = self.flip_v = False
        self.opacity = 100
        self.brightness = 0
        self.contrast = 0
        self.transp = False
        self.transp_thr = 235
        self.gray = False
        self.bw = False


def process_image(img, o, use_opacity=True):
    """Apply flips / effects / opacity / rotation of FloatImage `o` to an RGBA image."""
    im = img
    if o.flip_h:
        im = ImageOps.mirror(im)
    if o.flip_v:
        im = ImageOps.flip(im)
    if o.transp:
        thr = int(o.transp_thr)
        soft = 25
        lut = [0 if v >= thr else (255 if v < thr - soft else int((thr - v) * 255 / soft)) for v in range(256)]
        mask = im.convert("L").point(lut)
        a = ImageChops.multiply(im.getchannel("A"), mask)
        im = im.copy()
        im.putalpha(a)
    if o.gray or o.bw:
        a = im.getchannel("A")
        g = im.convert("L")
        if o.bw:
            g = ImageOps.autocontrast(g).point(lambda v: 255 if v >= 128 else 0)
        im = Image.merge("RGBA", (g, g, g, a))
    if o.brightness or o.contrast:
        a = im.getchannel("A")
        rgb = im.convert("RGB")
        if o.brightness:
            rgb = ImageEnhance.Brightness(rgb).enhance(1 + o.brightness / 100)
        if o.contrast:
            rgb = ImageEnhance.Contrast(rgb).enhance(1 + o.contrast / 100)
        im = rgb.convert("RGBA")
        im.putalpha(a)
    if use_opacity and o.opacity < 100:
        a = im.getchannel("A").point(lambda v: v * o.opacity // 100)
        im = im.copy()
        im.putalpha(a)
    if abs(o.angle) % 360 > 0.01:
        im = im.rotate(-o.angle, resample=BICUBIC, expand=True, fillcolor=(255, 255, 255, 0))
    return im


SHAPE_TOOLS = ("line", "arrow", "darrow", "rect", "rrect", "ellipse", "circle",
               "triangle", "diamond", "star", "check", "cross", "plus")


def shape_items(kind, a, b, lw, shift=False):
    """Geometry (in view space) of a shape dragged from point a to point b.
    Returns a list of (type, data, fillable) used for both the live preview and the PDF."""
    ax, ay = a
    bx, by = b
    if kind in ("line", "arrow", "darrow"):
        if shift:                                   # snap direction to 15 degree steps
            ang = math.atan2(by - ay, bx - ax)
            ang = round(ang / (math.pi / 12)) * (math.pi / 12)
            d = math.hypot(bx - ax, by - ay)
            bx, by = ax + d * math.cos(ang), ay + d * math.sin(ang)
        dx, dy = bx - ax, by - ay
        L = math.hypot(dx, dy)
        if L < 1:
            return []
        if kind == "line":
            return [("line", [(ax, ay), (bx, by)], False)]
        ux, uy = dx / L, dy / L
        nx, ny = -uy, ux
        hl = min(max(12.0, lw * 5.0), L * (0.4 if kind == "darrow" else 0.6))
        hw = hl * 0.45
        items = []
        tip_b = (bx, by)
        base_b = (bx - ux * hl, by - uy * hl)
        s0 = (ax, ay)
        if kind == "darrow":
            base_a = (ax + ux * hl, ay + uy * hl)
            s0 = (ax + ux * hl * 0.9, ay + uy * hl * 0.9)
            items.append(("head", [(ax, ay), (base_a[0] + nx * hw, base_a[1] + ny * hw),
                                   (base_a[0] - nx * hw, base_a[1] - ny * hw)], False))
        s1 = (bx - ux * hl * 0.9, by - uy * hl * 0.9)
        items.insert(0, ("line", [s0, s1], False))
        items.append(("head", [tip_b, (base_b[0] + nx * hw, base_b[1] + ny * hw),
                               (base_b[0] - nx * hw, base_b[1] - ny * hw)], False))
        return items

    dx, dy = bx - ax, by - ay
    if kind == "circle" or shift:                   # keep the box square
        side = max(abs(dx), abs(dy))
        dx = math.copysign(side, dx if dx else 1)
        dy = math.copysign(side, dy if dy else 1)
    l, r = sorted((ax, ax + dx))
    t, btm = sorted((ay, ay + dy))
    w, h = r - l, btm - t
    if w < 1 or h < 1:
        return []
    cx, cy = (l + r) / 2, (t + btm) / 2
    box = (l, t, r, btm)
    if kind == "rect":
        return [("rect", box, True)]
    if kind == "rrect":
        return [("rrect", box, True)]
    if kind in ("ellipse", "circle"):
        return [("oval", box, True)]
    if kind == "triangle":
        return [("poly", [(cx, t), (r, btm), (l, btm)], True)]
    if kind == "diamond":
        return [("poly", [(cx, t), (r, cy), (cx, btm), (l, cy)], True)]
    if kind == "star":
        pts = []
        for i in range(10):
            ang = -math.pi / 2 + i * math.pi / 5
            k = 1.0 if i % 2 == 0 else 0.4
            pts.append((cx + math.cos(ang) * k * w / 2, cy + math.sin(ang) * k * h / 2))
        return [("poly", pts, True)]
    if kind == "check":
        return [("open", [(l + 0.08 * w, t + 0.55 * h), (l + 0.38 * w, t + 0.85 * h), (r - 0.04 * w, t + 0.12 * h)], False)]
    if kind == "cross":
        return [("line", [(l, t), (r, btm)], False), ("line", [(r, t), (l, btm)], False)]
    if kind == "plus":
        return [("line", [(cx, t), (cx, btm)], False), ("line", [(l, cy), (r, cy)], False)]
    return []


class RibbonButton(tk.Canvas):
    """Ribbon button: a glossy 3-D icon tile over a label. It lifts when hovered and pushes in when pressed."""

    def __init__(self, parent, icon, text, command, icon_size=17, width=60, color=None):
        need = tkfont.Font(family=UI_FONT, size=8).measure(tr(text)) + S(16)      # longer translations get wider buttons
        self._bw, self._bh = max(S(width), need), S(64)
        super().__init__(parent, width=self._bw, height=self._bh, bg=RIBBON_BG, highlightthickness=0, cursor="hand2")
        self.command = command
        self.active = False
        self.hover = False
        self.pressed = False
        self._icon, self._text, self._isize = icon, text, icon_size
        self._color = color or getattr(parent, "_tile_color", None) or ACCENT
        self._tile = None
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)
        self.bind("<ButtonPress-1>", self._press)
        self.bind("<ButtonRelease-1>", self._release)
        self._paint()

    def _paint(self):
        self.delete("all")
        state = "pressed" if self.pressed else "active" if self.active else "hover" if self.hover else "normal"
        self._tile = TH.tile_photo(self, S(36), self._color, state, S)
        cx, cy = self._bw // 2, S(24)
        if self.active:
            round_rect(self, S(2), S(1), self._bw - S(2), self._bh - S(2), S(12), fill=HOVER_BG, outline="")
        lift = S(1) if state == "pressed" else 0
        self.create_image(cx, cy, image=self._tile)
        # glyph with a faint shadow so it looks embossed into the tile
        self.create_text(cx + 1, cy + S(1) + lift + 1, text=self._icon, font=(ICON_FONT, self._isize), fill="#14305c")
        self.create_text(cx, cy + lift, text=self._icon, font=(ICON_FONT, self._isize), fill="white")
        self.create_text(self._bw // 2, S(54), text=self._text, font=(UI_FONT, 8),
                         fill=ACCENT if self.active else TH.LABEL)

    def _enter(self, _e):
        self.hover = True
        self._paint()

    def _leave(self, _e):
        self.hover = False
        self.pressed = False
        self._paint()

    def _press(self, _e):
        self.pressed = True
        self._paint()

    def _release(self, e):
        was = self.pressed
        self.pressed = False
        self._paint()
        if was and 0 <= e.x < self._bw and 0 <= e.y < self._bh:
            self.command()

    def _click(self, _e=None):
        self.command()

    def set_active(self, on):
        self.active = on
        self._paint()


class QuickButton(tk.Canvas):
    """Small 3-D icon button for the quick-access toolbar (Open / Save / Save As) with a tooltip."""

    def __init__(self, parent, icon, tip, command, color):
        self._qs = S(24)
        side = self._qs + 2 * S(4)
        super().__init__(parent, width=side, height=side, bg=BAR_BG, highlightthickness=0, cursor="hand2")
        self._icon, self._tip, self.command, self._color = icon, tip, command, color
        self.hover = self.pressed = False
        self._tipwin = self._tipjob = None
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)
        self.bind("<ButtonPress-1>", self._press)
        self.bind("<ButtonRelease-1>", self._release)
        self._paint()

    def _paint(self):
        self.delete("all")
        state = "pressed" if self.pressed else "hover" if self.hover else "normal"
        self._img = TH.tile_photo(self, self._qs, self._color, state, S)
        c = (self._qs + 2 * S(4)) // 2
        lift = S(1) if self.pressed else 0
        self.create_image(c, c, image=self._img)
        self.create_text(c + 1, c + lift + 1, text=self._icon, font=(ICON_FONT, 11), fill="#14305c")
        self.create_text(c, c + lift, text=self._icon, font=(ICON_FONT, 11), fill="white")

    def _enter(self, _e):
        self.hover = True
        self._paint()
        self._tipjob = self.after(550, self._show_tip)

    def _leave(self, _e):
        self.hover = self.pressed = False
        self._paint()
        self._hide_tip()

    def _press(self, _e):
        self.pressed = True
        self._hide_tip()
        self._paint()

    def _release(self, e):
        was, self.pressed = self.pressed, False
        self._paint()
        side = self._qs + 2 * S(4)
        if was and 0 <= e.x < side and 0 <= e.y < side:
            self.command()

    def _show_tip(self):
        if self._tipwin is not None:
            return
        w = tk.Toplevel(self)
        w.wm_overrideredirect(True)
        w.wm_geometry("+%d+%d" % (self.winfo_rootx(), self.winfo_rooty() + self.winfo_height() + S(4)))
        tk.Label(w, text=self._tip, bg=TH.CARD, fg=TH.LABEL, font=(UI_FONT, 9), padx=8, pady=3,
                 highlightbackground=TH.LINE, highlightthickness=1).pack()
        self._tipwin = w

    def _hide_tip(self):
        if self._tipjob:
            self.after_cancel(self._tipjob)
            self._tipjob = None
        if self._tipwin is not None:
            self._tipwin.destroy()
            self._tipwin = None


class Snap(tuple):
    """an undo step: (document bytes, page number) with a readable name for the undo history"""
    def __new__(cls, data, page_no, label=""):
        s = super().__new__(cls, (data, page_no))
        s.label = label
        return s


class PdfEditor(V3Tools, V3bTools, ViewTools, FeatureTools, ProTools, ExtraTools, tk.Tk):
    def __init__(self, path=None):
        super().__init__()
        self._splash = None
        if not os.environ.get("CLAUDIUS_NO_SPLASH"):
            self.withdraw()
            self._show_splash()
        paths = [path] if isinstance(path, str) else list(path or [])
        if SCALE != 1:
            self.tk.call("tk", "scaling", SCALE * 96 / 72)
        self.title(APP_NAME)
        self.geometry(f"{S(1320)}x{S(820)}")
        self.minsize(S(900), S(560))
        self.configure(bg=BAR_BG)

        # state
        self.doc = None
        self.path = None
        self.page_no = 0
        self.zoom = 1.3
        self.tool = tk.StringVar(value="view")
        self.color = (0.85, 0.0, 0.0)
        self.font_size = tk.IntVar(value=14)
        self.line_width = tk.IntVar(value=2)
        self.undo_stack = []
        self.modified = False
        self._photo = None
        self._drag_start = None
        self._drag_item = None
        self._stroke = []
        self.off_x, self.off_y = PAD, PAD
        self.page_w = self.page_h = 0
        self.tool_buttons = {}
        self.swatches = []
        self.thumb_photos = {}
        self._thumb_job = 0
        self._sync_zoom = False
        self._zoom_after = None
        self.floats = []              # photos floating above the page
        self.sel = None
        self.float_page = 0
        self.fdrag = None
        self._handles = {}
        self.quality_var = tk.StringVar(value="High (300 dpi)")
        self.angle_var = tk.IntVar(value=0)
        self.keep_aspect = tk.BooleanVar(value=True)
        self.fill_color = (1.0, 0.85, 0.0)
        self.fill_on = tk.BooleanVar(value=False)
        self.shape_opacity = tk.IntVar(value=100)
        self.fill_swatches = []
        self.fit_mode = "page"           # default view: whole page visible
        self._fit_after = None

        TH.setup_styles(self, S)

        self.hl_color = HL_COLORS[0]
        self._extra_init()
        self._view_init()
        self._pro_init()
        self._v3_init()
        self._load_icons()
        self._build_topbar()
        self._build_ribbons()
        self._build_body()
        self._build_bottombar()
        self._bind_keys()
        self.show_tab("Home")
        self.set_tool("view")
        self._v3_after_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self._win11_init()
        self.restore_geometry()
        for p in paths:
            self.open_path(p)
        self.render()
        self.update_ui()
        self.after(200, self._refit)
        if not os.environ.get("CLAUDIUS_NO_RECOVERY"):
            self.after(700, self.start_recovery)
        self._close_splash()

    # ------------------------------------------------------------ splash screen
    def _show_splash(self):
        """small start-up window (the EXE shows its own splash while it unpacks; this one covers building the window)"""
        try:
            import pyi_splash                     # inside the EXE the build's own splash is already on screen - do not show a second one
            return
        except Exception:
            pass
        try:
            sp = tk.Toplevel(self)
            sp.overrideredirect(True)
            sp.configure(bg=TH.LINE if hasattr(TH, "LINE") else "#c7c7cc", padx=1, pady=1)
            card = tk.Frame(sp, bg=TH.CARD, padx=S(40), pady=S(26))
            card.pack()
            self._splash_logo = tk.PhotoImage(data=LOGO_128_B64, master=sp)
            tk.Label(card, image=self._splash_logo, bg=TH.CARD).pack()
            tk.Label(card, text=APP_NAME, font=(UI_FONT, 20, "bold"), bg=TH.CARD, fg=TH.LABEL).pack(pady=(10, 0))
            tk.Label(card, text=f"{tr('Version')} {APP_VERSION}", font=(UI_FONT, 10), bg=TH.CARD, fg=TH.SECOND).pack(pady=(4, 0))
            tk.Label(card, text=tr(APP_FREEWARE), font=(UI_FONT, 9), bg=TH.CARD, fg=TH.TERTIARY).pack(pady=(10, 0))
            tk.Label(card, text=tr("Starting..."), font=(UI_FONT, 9), bg=TH.CARD, fg=TH.SECOND).pack(pady=(14, 0))
            sp.update_idletasks()
            w, h = sp.winfo_reqwidth(), sp.winfo_reqheight()
            sp.geometry("%dx%d+%d+%d" % (w, h, (sp.winfo_screenwidth() - w) // 2, (sp.winfo_screenheight() - h) // 2))
            sp.attributes("-topmost", True)
            sp.update()
            self._splash = sp
        except Exception:
            self._splash = None

    def _close_splash(self):
        try:
            import pyi_splash                     # the splash made by the EXE build (only exists inside the EXE)
            pyi_splash.close()
        except Exception:
            pass
        sp, self._splash = getattr(self, "_splash", None), None
        if sp is not None:
            try:
                sp.destroy()
            except tk.TclError:
                pass
        if not os.environ.get("CLAUDIUS_NO_SPLASH"):
            self.deiconify()
            self.lift()

    # ------------------------------------------------------------ top bar
    def _build_topbar(self):
        top = tk.Frame(self, bg=BAR_BG)
        top.pack(side="top", fill="x")
        self.top_bar = top
        brand = tk.Frame(top, bg=BAR_BG, cursor="hand2")
        brand.pack(side="left", padx=(14, 12), pady=S(6))
        logo_lbl = tk.Label(brand, image=self.logo32 if SCALE > 1.2 else self.logo26, bg=BAR_BG)
        logo_lbl.pack(side="left")
        name_lbl = tk.Label(brand, text="Claudius PDF Editor", bg=BAR_BG, fg=TH.LABEL, font=(UI_FONT, 11, "bold"))
        name_lbl.pack(side="left", padx=(7, 0))
        for w in (brand, logo_lbl, name_lbl):
            w.bind("<Button-1>", lambda e: self.show_about())

        # quick-access toolbar (Open / Save / Save As / Print)
        qa = tk.Frame(top, bg=BAR_BG)
        qa.pack(side="left", padx=(0, 10))
        self.quick_buttons = []
        for icon, label, keys, cmd, col in (("📂", "Open", "Ctrl+O", self.open_dialog, TH.TILE_COLORS[0]),
                                            ("💾", "Save", "Ctrl+S", self.save, TH.TILE_COLORS[1]),
                                            ("📄", "Save As", "Ctrl+Shift+S", self.save_as, TH.TILE_COLORS[2]),
                                            ("🖨", "Print", "Ctrl+P", self.do_print, TH.TILE_COLORS[3])):
            qb = QuickButton(qa, icon, f"{tr(label)}  ({keys})", cmd, col)
            qb.pack(side="left", padx=1, pady=S(3))
            self.quick_buttons.append(qb)
        tk.Frame(top, bg=RIBBON_LINE, width=1).pack(side="left", fill="y", pady=S(8), padx=(0, 8))

        self.file_menu = tk.Menu(self, tearoff=0)
        self.file_menu.add_command(label="Open...", accelerator="Ctrl+O", command=self.open_dialog)
        self.recent_menu = tk.Menu(self.file_menu, tearoff=0, postcommand=self._fill_recent_menu)
        self.file_menu.add_cascade(label="Open recent", menu=self.recent_menu)
        self.file_menu.add_command(label="Create PDF from files...", command=self.dlg_create_pdf)
        self.file_menu.add_command(label="Save", accelerator="Ctrl+S", command=self.save)
        self.file_menu.add_command(label="Save As...", accelerator="Ctrl+Shift+S", command=self.save_as)
        self.file_menu.add_command(label="Print...", accelerator="Ctrl+P", command=self.do_print)
        self.file_menu.add_command(label="Close tab", accelerator="Ctrl+W", command=self.close_tab)
        self.file_menu.add_command(label="Document properties...", command=self.dlg_properties)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Settings...", accelerator="Ctrl+,", command=self.dlg_settings)
        self.file_menu.add_command(label="Undo history...", command=self.dlg_undo_history)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Export current page as PNG...", command=self.export_png)
        self.file_menu.add_command(label="Extract pages...", command=self.dlg_extract)
        self.file_menu.add_separator()
        if sys.platform == "win32":
            self.file_menu.add_command(label="Add Claudius PDF Editor to the 'Open with' list...", command=self.register_open_with)
            self.file_menu.add_command(label="Make Claudius PDF Editor the default PDF app...", command=self.choose_default_app)
            self.file_menu.add_separator()
        self.lang_menu = tk.Menu(self.file_menu, tearoff=0)
        self.lang_var = tk.StringVar(value=i18n.current)
        for code, name in i18n.LANGS:
            self.lang_menu.add_radiobutton(label=name, value=code, variable=self.lang_var,
                                           command=lambda c=code: self.set_ui_language(c))
        self.file_menu.add_cascade(label="Language", menu=self.lang_menu)
        self.theme_menu = tk.Menu(self.file_menu, tearoff=0)
        self.theme_var = tk.StringVar(value=TH.load_setting())
        for code, name in (("auto", "System (follow Windows)"), ("light", "Light"), ("dark", "Dark")):
            self.theme_menu.add_radiobutton(label=name, value=code, variable=self.theme_var,
                                            command=lambda c=code: self.set_theme(c))
        self.file_menu.add_cascade(label="Appearance", menu=self.theme_menu)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="About Claudius PDF Editor...", command=self.show_about)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.on_close)

        self.tab_labels = {}
        file_pill = PillTab(top, "File", lambda: self.file_menu.tk_popup(
            file_pill.winfo_rootx(), file_pill.winfo_rooty() + file_pill.winfo_height()), S, bg=BAR_BG)
        file_pill.pack(side="left", padx=(0, 4), pady=S(6))

        for name in ("Home", "Comment", "Shapes", "Edit", "Sign", "Photos", "Adjust", "Organize", "Document", "Convert", "Tools", "Review", "View"):
            pill = PillTab(top, name, lambda n=name: self.show_tab(n), S, bg=BAR_BG)
            pill.pack(side="left", padx=1, pady=S(6))
            self.tab_labels[name] = pill

        self.doc_lbl = tk.Label(top, text="", bg=BAR_BG, fg=TH.SECOND, font=(UI_FONT, 9))
        self.doc_lbl.pack(side="right", padx=14)
        tk.Frame(self, bg=RIBBON_LINE, height=1).pack(side="top", fill="x")

    def show_tab(self, name):
        self.current_tab = name
        for n, pill in self.tab_labels.items():
            pill.set_selected(n == name)
        for n, fr in self.ribbons.items():
            fr.pack_forget()
        self.ribbons[name].pack(fill="both", expand=True)

    # ------------------------------------------------------------- ribbon
    def group(self, parent, caption):
        outer = tk.Frame(parent, bg=RIBBON_BG)
        outer.pack(side="left", fill="y", pady=(4, 2))
        inner = tk.Frame(outer, bg=RIBBON_BG)
        self._grp_i = getattr(self, "_grp_i", -1) + 1
        inner._tile_color = TH.TILE_COLORS[self._grp_i % len(TH.TILE_COLORS)]
        inner.pack(side="top", fill="both", expand=True)
        tk.Label(outer, text=caption, font=(UI_FONT, 8), bg=RIBBON_BG, fg=TH.TERTIARY).pack(side="bottom")
        tk.Frame(parent, bg=RIBBON_LINE, width=1).pack(side="left", fill="y", pady=8, padx=4)
        return inner

    def rb(self, parent, icon, text, cmd, tool=None, size=17, width=60):
        if tool:
            b = RibbonButton(parent, icon, text, lambda t=tool: self.set_tool(t), size, width)
            self.tool_buttons.setdefault(tool, []).append(b)
        else:
            b = RibbonButton(parent, icon, text, cmd, size, width)
        b.pack(side="left", padx=1)
        return b

    def style_group(self, parent):
        g = self.group(parent, "Style")
        f = tk.Frame(g, bg=RIBBON_BG)
        f.pack(padx=6, pady=2)

        def lab(t, r):
            tk.Label(f, text=t, font=(UI_FONT, 8), bg=RIBBON_BG, fg=TH.LABEL).grid(row=r, column=0, sticky="w", pady=1)

        lab("Color", 0)
        sw = tk.Button(f, width=4, relief="solid", bd=1, bg=self._hex(self.color),
                       activebackground=self._hex(self.color), command=self.pick_color, cursor="hand2")
        sw.grid(row=0, column=1, padx=6, pady=1)
        self.swatches.append(sw)
        lab("Size", 1)
        ttk.Spinbox(f, from_=6, to=96, width=4, textvariable=self.font_size).grid(row=1, column=1, padx=6, pady=1)
        lab("Line", 2)
        ttk.Spinbox(f, from_=1, to=20, width=4, textvariable=self.line_width).grid(row=2, column=1, padx=6, pady=1)

    def _build_ribbons(self):
        holder = tk.Frame(self, bg=RIBBON_BG, height=S(96))
        holder.pack(side="top", fill="x")
        holder.pack_propagate(False)
        tk.Frame(self, bg=RIBBON_LINE, height=1).pack(side="top", fill="x")
        self.ribbons = {}
        for name in ("Home", "Comment", "Shapes", "Edit", "Sign", "Photos", "Adjust", "Organize", "Document", "Convert", "Tools", "Review", "View"):
            self.ribbons[name] = tk.Frame(holder, bg=RIBBON_BG)

        # Home
        r = self.ribbons["Home"]
        g = self.group(r, "File")
        self.rb(g, "📂", "Open", self.open_dialog)
        self.rb(g, "💾", "Save", self.save)
        self.rb(g, "📄", "Save As", self.save_as)
        self.rb(g, "🖨", "Print", self.do_print)
        g = self.group(r, "Edit")
        self.rb(g, "↶", "Undo", self.undo, size=19)
        self.rb(g, "↷", "Redo", self.redo, size=19)
        g = self.group(r, "Tools")
        self.rb(g, "✋", "Hand", None, tool="view")
        self.rb(g, "🔍", "Find", self.show_find)
        self.rb(g, "✏", "Edit text", None, tool="edittext", size=19, width=62)
        g = self.group(r, "Photos")
        self.rb(g, "🖼", "Add Image", self.add_images)
        self.rb(g, "📋", "Paste", self.paste_image)
        g = self.group(r, "Zoom")
        self.rb(g, "−", "Zoom out", lambda: self.set_zoom(self.zoom / 1.2), size=20)
        self.rb(g, "+", "Zoom in", lambda: self.set_zoom(self.zoom * 1.2), size=20)
        self.rb(g, "↔", "Fit width", self.fit_width, size=19)
        self.rb(g, "⛶", "Fit page", self.fit_page, size=19)
        g = self.group(r, "Rotate")
        self.rb(g, "⟲", "Left", lambda: self.rotate_selected(-90), size=19)
        self.rb(g, "⟳", "Right", lambda: self.rotate_selected(90), size=19)

        # Comment
        r = self.ribbons["Comment"]
        g = self.group(r, "Text markup")
        self.rb(g, "🖍", "Highlight", None, tool="highlight")
        self.rb(g, "U̲", "Underline", None, tool="underline", size=19, width=62)
        self.rb(g, "S̶", "Strikeout", None, tool="strike", size=19, width=62)
        self.rb(g, "≈", "Squiggly", None, tool="squiggly", size=20, width=62)
        g = self.group(r, "Draw")
        self.rb(g, "✎", "Pen", None, tool="pen", size=19)
        self.rb(g, "▭", "Rectangle", None, tool="rect", size=20)
        self.rb(g, "○", "Ellipse", None, tool="ellipse", size=20)
        self.rb(g, "➜", "Arrow", None, tool="arrow", size=19)
        g = self.group(r, "Comments list")
        self.rb(g, "💬", "Show list", lambda: self.set_sidebar("comments"), size=18, width=62)
        self.rb(g, "⤓", "Export", self.export_comments, size=19, width=56)
        self.rb(g, "⤒", "Import", self.import_comments, size=19, width=56)
        self.rb(g, "≣", "Summary", self.export_comment_summary, size=19, width=62)
        g = self.group(r, "Notes")
        self.rb(g, "🗒", "Sticky note", None, tool="sticky", size=18, width=70)
        self.rb(g, "▤", "Text box", None, tool="textbox", size=19)
        self.rb(g, "T", "Text", None, tool="text", size=19)
        self.rb(g, "🔖", "Stamps", self.show_stamp_menu, size=18)
        g = self.group(r, "Highlight color")
        hf = tk.Frame(g, bg=RIBBON_BG)
        hf.pack(padx=6, pady=14)
        for col in HL_COLORS:
            hb = tk.Button(hf, width=2, relief="solid", bd=1, bg=self._hex(col), activebackground=self._hex(col),
                           cursor="hand2", command=lambda c=col: setattr(self, "hl_color", c))
            hb.pack(side="left", padx=2)
            self.hl_buttons.append(hb)
        self.style_group(r)

        # Shapes
        r = self.ribbons["Shapes"]
        g = self.group(r, "Lines")
        self.rb(g, "╱", "Line", None, tool="line", size=20)
        self.rb(g, "➜", "Arrow", None, tool="arrow", size=19)
        self.rb(g, "⟷", "Double arrow", None, tool="darrow", size=19, width=82)
        g = self.group(r, "Shapes")
        self.rb(g, "▭", "Rectangle", None, tool="rect", size=20)
        self.rb(g, "▢", "Rounded", None, tool="rrect", size=20)
        self.rb(g, "⬭", "Ellipse", None, tool="ellipse", size=20)
        self.rb(g, "○", "Circle", None, tool="circle", size=20)
        self.rb(g, "△", "Triangle", None, tool="triangle", size=20)
        self.rb(g, "◇", "Diamond", None, tool="diamond", size=20)
        self.rb(g, "☆", "Star", None, tool="star", size=20)
        g = self.group(r, "Marks")
        self.rb(g, "✓", "Check", None, tool="check", size=20)
        self.rb(g, "✗", "Cross", None, tool="cross", size=20)
        self.rb(g, "＋", "Plus", None, tool="plus", size=20)
        self.shape_style_group(r)

        # Edit
        r = self.ribbons["Edit"]
        g = self.group(r, "Change")
        self.rb(g, "✏", "Edit text", None, tool="edittext", size=19, width=62)
        self.rb(g, "T", "Add text", None, tool="text", size=19, width=62)
        self.rb(g, "▤", "Text box", None, tool="textbox", size=19)
        self.rb(g, "🖼", "Add Image", self.add_images)
        self.rb(g, "🔗", "Link", None, tool="link", size=18)
        self.rb(g, "📅", "Date", None, tool="date", size=18)
        g = self.group(r, "Cover / remove")
        self.rb(g, "◻", "Whiteout", None, tool="whiteout", size=20)
        self.rb(g, "▮", "Redact", None, tool="redact", size=20)
        self.rb(g, "⌫", "Redact text", self.dlg_redact_text, size=19, width=66)
        self.rb(g, "✱", "Redact data", self.dlg_redact_patterns, size=19, width=66)
        g = self.group(r, "Copy")
        self.rb(g, "⎘", "Extract text", None, tool="extract", size=19, width=70)
        g = self.group(r, "Undo")
        self.rb(g, "↶", "Undo", self.undo, size=19)
        self.rb(g, "↷", "Redo", self.redo, size=19)
        self.style_group(r)

        # Sign
        r = self.ribbons["Sign"]
        g = self.group(r, "Sign")
        self.rb(g, "✍", "Signature", lambda: self.show_sign_menu("sign"), size=19, width=68)
        self.rb(g, "Aa", "Initials", lambda: self.show_sign_menu("init"), size=17)
        self.rb(g, "🔖", "Stamps", self.show_stamp_menu, size=18)
        self.rb(g, "📅", "Date", None, tool="date", size=18)
        g = self.group(r, "Fill & create forms")
        self.rb(g, "☐", "Text field", None, tool="formtext", size=20, width=66)
        self.rb(g, "☑", "Check box", None, tool="formcheck", size=20, width=66)
        self.rb(g, "▦", "Show fields", self.toggle_fields, size=18, width=68)
        g = self.group(r, "How to fill a form")
        tk.Label(g, text="Choose the Hand tool and click a field\n(text, check box, drop-down list).",
                 font=(UI_FONT, 8), bg=RIBBON_BG, fg=TH.SECOND, justify="left").pack(padx=8, pady=14)
        self.style_group(r)

        # Photos
        r = self.ribbons["Photos"]
        g = self.group(r, "Insert")
        self.rb(g, "🖼", "Add Image", self.add_images)
        self.rb(g, "📋", "Paste", self.paste_image)
        self.rb(g, "📑", "Photos→Pages", self.photos_to_pages, width=86)
        g = self.group(r, "Arrange")
        self.rb(g, "⛶", "Fit page", self.img_fit_page)
        self.rb(g, "⊞", "Align", self.img_align_menu)
        self.rb(g, "❐", "Duplicate", self.img_duplicate)
        self.rb(g, "⬆", "To front", lambda: self.img_order(True), size=15)
        self.rb(g, "⬇", "To back", lambda: self.img_order(False), size=15)
        self.rb(g, "🗑", "Delete", self.img_delete)
        g = self.group(r, "Finish")
        self.rb(g, "✔", "Apply", self.apply_images, size=19)
        qf = tk.Frame(g, bg=RIBBON_BG)
        qf.pack(side="left", padx=8, pady=8)
        tk.Label(qf, text="Output quality", font=(UI_FONT, 8), bg=RIBBON_BG, fg=TH.LABEL).pack(anchor="w")
        ttk.Combobox(qf, textvariable=self.quality_var, values=list(QUALITY), state="readonly",
                     width=18).pack(pady=2)

        # Adjust
        r = self.ribbons["Adjust"]
        g = self.group(r, "Rotate / flip")
        self.rb(g, "⟲", "Left 90°", lambda: self.img_rotate(-90), size=19)
        self.rb(g, "⟳", "Right 90°", lambda: self.img_rotate(90), size=19)
        self.rb(g, "⇋", "Flip H", lambda: self.img_flip(True), size=19)
        self.rb(g, "⇅", "Flip V", lambda: self.img_flip(False), size=19)
        af = tk.Frame(g, bg=RIBBON_BG)
        af.pack(side="left", padx=8, pady=14)
        tk.Label(af, text="Angle °", font=(UI_FONT, 8), bg=RIBBON_BG, fg=TH.LABEL).pack(anchor="w")
        asp = ttk.Spinbox(af, from_=-180, to=180, width=5, textvariable=self.angle_var,
                          command=self._angle_changed)
        asp.pack()
        asp.bind("<Return>", self._angle_changed)
        ttk.Checkbutton(af, text="Keep aspect", variable=self.keep_aspect).pack(anchor="w", pady=(3, 0))
        g = self.group(r, "Crop")
        self.rb(g, "✂", "Crop", self.start_crop, size=19)
        g = self.group(r, "Effects")
        self.rb(g, "☀", "Adjust…", self.adjust_dialog, size=19)
        self.rb(g, "◧", "No white", lambda: self.img_toggle("transp"), size=19)
        self.rb(g, "◐", "Grayscale", lambda: self.img_toggle("gray"), size=19, width=64)
        self.rb(g, "◑", "B & W", lambda: self.img_toggle("bw"), size=19)
        self.rb(g, "▣", "Border", lambda: self.img_toggle("border"), size=19)
        g = self.group(r, "Undo")
        self.rb(g, "↺", "Reset image", self.img_reset, size=19, width=70)

        # Organize
        r = self.ribbons["Organize"]
        g = self.group(r, "Rotate")
        self.rb(g, "⟲", "Left", lambda: self.rotate_selected(-90), size=19)
        self.rb(g, "⟳", "Right", lambda: self.rotate_selected(90), size=19)
        g = self.group(r, "Pages")
        self.rb(g, "🗑", "Delete", self.delete_selected)
        self.rb(g, "▲", "Move up", lambda: self.move_page(-1), size=15)
        self.rb(g, "▼", "Move down", lambda: self.move_page(1), size=15)
        self.rb(g, "📄", "Extract", self.dlg_extract)
        self.rb(g, "✂", "Split", self.dlg_split, size=19)
        g = self.group(r, "Insert")
        self.rb(g, "＋", "Blank page", self.insert_blank, size=19, width=66)
        self.rb(g, "⧉", "Insert PDF", self.dlg_insert_pdf, size=19, width=66)
        self.rb(g, "⛓", "Merge PDFs", self.dlg_merge, size=19, width=66)
        self.rb(g, "📑", "Photos→Pages", self.photos_to_pages, width=86)
        g = self.group(r, "Crop pages")
        self.rb(g, "⛶", "Drag area", None, tool="pagecrop", size=18, width=62)
        self.rb(g, "▭", "Margins…", self.dlg_crop, size=19, width=62)
        g = self.group(r, "Export")
        self.rb(g, "🖼", "To PNG", self.export_png)

        # Document
        r = self.ribbons["Document"]
        g = self.group(r, "Page decoration")
        self.rb(g, "💧", "Watermark", self.dlg_watermark, size=18, width=76)
        self.rb(g, "🎨", "Background", self.dlg_background, size=18, width=78)
        self.rb(g, "⇱", "Header/Footer", self.dlg_header_footer, size=18, width=92)
        self.rb(g, "#", "Page numbers", self.dlg_page_numbers, size=19, width=88)
        g = self.group(r, "Optimize")
        self.rb(g, "🗜", "Compress", self.dlg_compress, size=18, width=66)
        self.rb(g, "▩", "Flatten", self.do_flatten, size=18)
        self.rb(g, "👁", "OCR", self.dlg_ocr, size=18)
        g = self.group(r, "Security")
        self.rb(g, "🔒", "Protect", self.dlg_protect, size=18)
        self.rb(g, "🔓", "Unprotect", self.do_unprotect, size=18, width=70)
        g = self.group(r, "Numbering")
        self.rb(g, "§", "Bates", self.dlg_bates, size=19, width=60)
        self.rb(g, "ⅰ", "Page labels", self.dlg_page_labels, size=19, width=72)
        g = self.group(r, "Archive")
        self.rb(g, "🗄", "PDF/A", self.dlg_pdfa, size=18, width=56)
        g = self.group(r, "Check / attach")
        self.rb(g, "♿", "Accessibility", self.dlg_accessibility, size=18, width=84)
        self.rb(g, "📎", "Attachments", self.dlg_attachments, size=18, width=78)
        self.rb(g, "⇄", "Compare", self.dlg_compare, size=18, width=62)
        g = self.group(r, "Info")
        self.rb(g, "ℹ", "Properties", self.dlg_properties, size=18, width=66)
        self.rb(g, "🖨", "Print", self.do_print, size=18)

        # Convert
        r = self.ribbons["Convert"]
        g = self.group(r, "Convert PDF to")
        for key, icon, label in (("word", "W", "Word"), ("excel", "X", "Excel"), ("pptx", "P", "PowerPoint"),
                                 ("images", "🖼", "Images"), ("text", "T", "Text"), ("html", "◇", "HTML"),
                                 ("xml", "‹›", "XML")):
            self.rb(g, icon, label, lambda k=key: self.dlg_convert(k), size=18, width=70 if key == "pptx" else 60)
        g = self.group(r, "Create PDF")
        self.rb(g, "➕", "From files", self.dlg_create_pdf, size=18, width=66)
        self.rb(g, "🖷", "From scanner", self.dlg_scan, size=18, width=78)
        self.rb(g, "📑", "Photos→Pages", self.photos_to_pages, width=86)

        # Tools (version 3)
        r = self.ribbons["Tools"]
        g = self.group(r, "Pictures & signature")
        self.rb(g, "🖼", "Edit pictures", None, tool="editimage", size=18, width=74)
        self.rb(g, "✍", "Signature block", self.dlg_signature_block, size=18, width=84)
        g = self.group(r, "Pages")
        self.rb(g, "⛶", "Crop / resize", self.dlg_page_crop, size=18, width=76)
        self.rb(g, "⛓", "Merge + pick", self.dlg_merge_pick, size=18, width=72)
        self.rb(g, "↶", "Undo history", self.dlg_undo_history, size=18, width=78)
        g = self.group(r, "Forms")
        self.rb(g, "▾", "Drop-down", None, tool="formcombo", size=18, width=66)
        self.rb(g, "◉", "Radio button", None, tool="formradio", size=18, width=76)
        self.rb(g, "☰", "Field editor", self.dlg_field_editor, size=18, width=72)
        self.rb(g, "⇥", "Fill from table", self.dlg_fill_from_data, size=18, width=88)
        g = self.group(r, "Files")
        self.rb(g, "▤", "Batch folder", self.dlg_batch, size=18, width=72)
        self.rb(g, "🗜", "Smart compress", self.dlg_smart_compress, size=18, width=88)
        self.rb(g, "🔍", "Search folder", self.dlg_folder_search, size=18, width=82)
        r = self.ribbons["Review"]
        g = self.group(r, "Highlights & checks")
        self.rb(g, "🖍", "Export highlights", self.dlg_export_highlights, size=18, width=96)
        self.rb(g, "🌐", "Look up", self.dlg_lookup, size=18, width=56)
        self.rb(g, "✔", "Redaction check", self.dlg_redaction_check, size=18, width=92)
        self.rb(g, "🧹", "Clean data", self.dlg_clean_metadata, size=18, width=72)
        g = self.group(r, "Program")
        self.rb(g, "⚙", "Settings", self.dlg_settings, size=18, width=60)

        # View
        r = self.ribbons["View"]
        g = self.group(r, "Zoom")
        self.rb(g, "−", "Zoom out", lambda: self.set_zoom(self.zoom / 1.2), size=20)
        self.rb(g, "+", "Zoom in", lambda: self.set_zoom(self.zoom * 1.2), size=20)
        self.rb(g, "↔", "Fit width", self.fit_width, size=19)
        self.rb(g, "⛶", "Fit page", self.fit_page, size=19)
        g = self.group(r, "Navigate")
        self.rb(g, "◀", "Previous", lambda: self.goto(self.page_no - 1), size=15)
        self.rb(g, "▶", "Next", lambda: self.goto(self.page_no + 1), size=15)
        g = self.group(r, "Page view")
        self.rb(g, "☰", "Continuous", self.toggle_continuous, size=19, width=70)
        self.rb(g, "☾", "Night view", self.toggle_night, size=19, width=70)
        self.rb(g, "⛶", "Full screen", lambda: self.toggle_fullscreen(False), size=19, width=70)
        self.rb(g, "▶", "Presentation", lambda: self.toggle_fullscreen(True), size=17, width=78)
        g = self.group(r, "Listen")
        self.rb(g, "🔊", "Read aloud", self.dlg_read_aloud, size=18, width=70)
        self.rb(g, "⏹", "Stop", self.stop_speaking, size=17, width=48)
        g = self.group(r, "Side panel")
        self.rb(g, "◧", "Hide / show", self.toggle_sidebar, size=18, width=70)
        self.rb(g, "▤", "Pages", lambda: self.set_sidebar("pages"), size=18)
        self.rb(g, "🔖", "Bookmarks", lambda: self.set_sidebar("bookmarks"), size=18, width=66)
        self.rb(g, "💬", "Comments", lambda: self.set_sidebar("comments"), size=18, width=66)
        g = self.group(r, "Search / forms")
        self.rb(g, "🔍", "Find", self.show_find)
        self.rb(g, "▦", "Show fields", self.toggle_fields, size=18, width=68)

    # --------------------------------------------------------------- body
    def _build_body(self):
        body = tk.Frame(self, bg=BAR_BG)
        body.pack(side="top", fill="both", expand=True)
        self.body_frame = body
        self._build_tabbar()

        # sidebar with thumbnails
        side = tk.Frame(body, bg=SIDEBAR_BG, width=SIDEBAR_W)
        self.side_frame = side
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        self._build_sidebar_tabs(side)
        wrap = tk.Frame(side, bg=SIDEBAR_BG)
        wrap.pack(fill="both", expand=True)
        self.pages_wrap = wrap
        self.thumbs = tk.Canvas(wrap, bg=SIDEBAR_BG, highlightthickness=0, width=SIDEBAR_W - S(14))
        tsb = ttk.Scrollbar(wrap, orient="vertical", command=self.thumbs.yview)
        self.thumbs.config(yscrollcommand=tsb.set)
        tsb.pack(side="right", fill="y")
        self.thumbs.pack(side="left", fill="both", expand=True)
        self.thumbs.bind("<ButtonPress-1>", self._thumb_press)
        self.thumbs.bind("<B1-Motion>", self._thumb_drag)
        self.thumbs.bind("<ButtonRelease-1>", self._thumb_release)
        self.thumbs.bind("<Button-3>", self._thumb_menu)
        self.thumbs.bind("<Button-2>", self._thumb_menu)
        self.thumbs.config(yscrollincrement=40)
        self.side_sep = tk.Frame(body, bg=RIBBON_LINE, width=1)
        self.side_sep.pack(side="left", fill="y")

        # page canvas
        right = tk.Frame(body, bg=WORK_BG)
        self.page_area = right
        right.pack(side="left", fill="both", expand=True)
        self.canvas = tk.Canvas(right, bg=WORK_BG, highlightthickness=0)
        vs = ttk.Scrollbar(right, orient="vertical", command=self.canvas.yview)
        hs = ttk.Scrollbar(right, orient="horizontal", command=self.canvas.xview)
        self._vs = vs
        self.canvas.config(yscrollcommand=self._on_yscroll, xscrollcommand=hs.set)
        vs.pack(side="right", fill="y")
        hs.pack(side="bottom", fill="x")
        self.canvas.pack(fill="both", expand=True)

        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Motion>", self._on_motion)
        self.canvas.bind("<Double-Button-1>", lambda e: self._text_dblclick(e, 2))
        self.canvas.bind("<Triple-Button-1>", lambda e: self._text_dblclick(e, 3))
        self.canvas.config(yscrollincrement=30, xscrollincrement=30)
        self._last_scroll = 0.0
        # one global handler: works wherever the pointer is, even if the page area has no keyboard focus
        self.bind_all("<MouseWheel>", self._wheel_router)
        self.bind_all("<Button-4>", lambda e: self._wheel_router(e, 1))
        self.bind_all("<Button-5>", lambda e: self._wheel_router(e, -1))
        self.canvas.bind("<Configure>", self._on_canvas_resize)
        self.canvas.bind("<Button-3>", self._on_right_click)
        self.canvas.bind("<Button-2>", self._on_right_click)

    def _build_bottombar(self):
        tk.Frame(self, bg=RIBBON_LINE, height=1).pack(side="bottom", fill="x")
        bar = tk.Frame(self, bg=BAR_BG)
        bar.pack(side="bottom", fill="x")

        self.status = tk.Label(bar, text="Open a PDF to start (Ctrl+O)", bg=BAR_BG, fg=TH.SECOND,
                               font=(UI_FONT, 9), anchor="w")
        self.status.pack(side="left", padx=12, pady=5)

        # right: zoom controls
        zf = tk.Frame(bar, bg=BAR_BG)
        zf.pack(side="right", padx=12)
        self._flat_btn(zf, "⛶", self.fit_page).pack(side="left", padx=(0, 10))
        self._flat_btn(zf, "−", lambda: self.set_zoom(self.zoom / 1.2)).pack(side="left")
        self.zoom_var = tk.DoubleVar(value=self.zoom * 100)
        TH.IOSSlider(zf, S, from_=20, to=400, variable=self.zoom_var, length=S(150),
                     command=self._on_scale).pack(side="left", padx=4)
        self._flat_btn(zf, "+", lambda: self.set_zoom(self.zoom * 1.2)).pack(side="left")
        self.zoom_lbl = tk.Label(zf, text="130%", width=5, bg=BAR_BG, font=(UI_FONT, 9))
        self.zoom_lbl.pack(side="left", padx=(6, 0))

        # centre: page navigation
        nav = tk.Frame(bar, bg=BAR_BG)
        nav.pack(side="right", padx=30)
        self._flat_btn(nav, "◀", lambda: self.goto(self.page_no - 1)).pack(side="left")
        self.page_var = tk.StringVar(value="")
        self.page_entry = ttk.Entry(nav, width=4, textvariable=self.page_var, justify="center")
        self.page_entry.pack(side="left", padx=4)
        self.page_entry.bind("<Return>", self._on_page_entry)
        self.total_lbl = tk.Label(nav, text="/ -", bg=BAR_BG, font=(UI_FONT, 9), width=12, anchor="w")
        self.total_lbl.pack(side="left")
        self._flat_btn(nav, "▶", lambda: self.goto(self.page_no + 1)).pack(side="left")

    def _flat_btn(self, parent, text, cmd):
        b = tk.Label(parent, text=text, font=(ICON_FONT, 11), bg=BAR_BG, fg=ACCENT,
                     width=2, cursor="hand2")
        b.bind("<Button-1>", lambda e: cmd())
        b.bind("<Enter>", lambda e: b.config(bg=HOVER_BG))
        b.bind("<Leave>", lambda e: b.config(bg=BAR_BG))
        return b

    def _bind_keys(self):
        self.bind("<Control-o>", lambda e: self.open_dialog())
        self.bind("<Control-s>", lambda e: self.save())
        self.bind("<Control-S>", lambda e: self.save_as())
        self.bind("<Control-z>", lambda e: self.undo())
        self.bind("<Control-f>", lambda e: self.show_find())
        self.bind("<Control-b>", lambda e: self.toggle_sidebar())
        self.bind("<Control-y>", lambda e: self.redo())
        self.bind("<Control-Z>", lambda e: self.redo())
        self.bind("<Control-p>", lambda e: self.do_print())
        self.bind("<Control-c>", self._key_copy)
        self.bind("<Control-w>", lambda e: self.close_tab())
        self.bind("<Control-Tab>", lambda e: self.next_tab(1))
        self.bind("<Control-Shift-Tab>", lambda e: self.next_tab(-1))
        self.bind("<Control-ISO_Left_Tab>", lambda e: self.next_tab(-1))
        self.bind("<F11>", lambda e: self.toggle_fullscreen(False))
        self.bind("<F5>", lambda e: self.toggle_fullscreen(True))
        self.bind("<space>", lambda e: self._pres_key(1, 0) if self.presenting else None)
        self.bind("<BackSpace>", lambda e: self._pres_key(-1, 0) if self.presenting else None)
        self.bind("<Home>", lambda e: self.goto(0) if (self.presenting or not self._typing()) and self.doc else None)
        self.bind("<End>", lambda e: self.goto(len(self.doc) - 1) if (self.presenting or not self._typing()) and self.doc else None)
        self.bind("<Control-a>", lambda e: None if self._typing() else self.select_all_pages())
        self.bind("<Delete>", self._key_delete)
        self.bind("<Return>", self._key_enter)
        self.bind("<Control-v>", self._key_paste)
        self.bind("<Control-d>", lambda e: None if self._typing() else self.img_duplicate())
        self.bind("<Escape>", lambda e: (self.exit_fullscreen() if self._fs else None, self.select(None), self.hide_find() if self.find_frame is not None and self.find_frame.winfo_ismapped() else None))
        for key, dx, dy in (("Left", -1, 0), ("Right", 1, 0), ("Up", 0, -1), ("Down", 0, 1)):
            self.bind(f"<{key}>", lambda e, a=dx, b=dy: self._key_arrow(a, b, e))
        self.bind("<Prior>", lambda e: self._key_nav(-1))
        self.bind("<Next>", lambda e: self._key_nav(1))

    def _typing(self):
        return isinstance(self.focus_get(), (tk.Entry, ttk.Entry, ttk.Spinbox))

    def _key_delete(self, _e):
        if self._typing():
            return
        if self.sel in self.floats and self.sel is not None:
            self.img_delete()
        else:
            self.delete_page()

    def _key_nav(self, step):
        if not self._typing():
            self.goto(self.page_no + step)

    # ------------------------------------------------------------ helpers
    @staticmethod
    def _hex(rgb):
        return "#%02x%02x%02x" % tuple(int(round(c * 255)) for c in rgb)

    HINTS = {"view": "Hand tool: drag over text to select it (Ctrl+C copies), drag elsewhere to move the page. Ctrl + mouse wheel to zoom.",
             "text": "Text: click on the page where the text should go.",
             "highlight": "Highlight: drag over the text.",
             "rect": "Rectangle: drag to draw. Hold Shift for a square. Turn on Fill in the Shapes tab for solid shapes.",
             "pen": "Pen: hold the mouse button and draw.",
             "whiteout": "Whiteout: drag over content to cover it with white.",
             "image": "Image: drag a box, then choose an image.",
             "crop": "Crop: drag a rectangle over the part of the image to keep.",
             "line": "Line: drag from start to end. Hold Shift for 15° steps.",
             "arrow": "Arrow: drag from the tail to the arrow tip. Hold Shift for 15° steps.",
             "darrow": "Double arrow: drag between the two tips. Hold Shift for 15° steps.",
             "rrect": "Rounded rectangle: drag to draw. Hold Shift for a square.",
             "ellipse": "Ellipse: drag to draw. Hold Shift for a circle.",
             "circle": "Circle: drag from one corner to the other.",
             "triangle": "Triangle: drag to draw. Hold Shift for equal sides.",
             "diamond": "Diamond: drag to draw. Hold Shift for a square diamond.",
             "star": "Star: drag to draw. Hold Shift for a regular star.",
             "check": "Check mark: drag a box where the mark should go.",
             "cross": "Cross: drag a box where the mark should go.",
             "plus": "Plus: drag a box where the mark should go.",
             "pagecrop": "Crop pages: drag a rectangle - everything outside it is cut away."}

    def set_tool(self, t):
        self.tool.set(t)
        for name, btns in self.tool_buttons.items():
            for b in btns:
                b.set_active(name == t)
        self.canvas.config(cursor="fleur" if t == "view" else "crosshair")
        self.status.config(text=self.HINTS.get(t) or EXTRA_HINTS.get(t, ""))

    def pick_color(self):
        c = colorchooser.askcolor(color=self._hex(self.color), title="Choose color")
        if c and c[0]:
            self.color = tuple(v / 255 for v in c[0])
            for sw in self.swatches:
                sw.config(bg=self._hex(self.color), activebackground=self._hex(self.color))

    def need_doc(self):
        if self.doc is None:
            messagebox.showinfo(APP_NAME, "Open a PDF first.")
            return False
        return True

    def page(self):
        return self.doc[self.page_no]

    def to_pdf_point(self, cx, cy):
        x = (self.canvas.canvasx(cx) - self.off_x) / self.zoom
        y = (self.canvas.canvasy(cy) - self.off_y) / self.zoom
        return fitz.Point(x, y) * self.page().derotation_matrix

    def to_pdf_rect(self, x0, y0, x1, y1):
        """x/y are canvas (already scrolled) coordinates."""
        m = self.page().derotation_matrix
        p0 = fitz.Point((x0 - self.off_x) / self.zoom, (y0 - self.off_y) / self.zoom) * m
        p1 = fitz.Point((x1 - self.off_x) / self.zoom, (y1 - self.off_y) / self.zoom) * m
        r = fitz.Rect(p0, p1)
        r.normalize()
        return r

    # ----------------------------------------------------------- file ops
    def open_dialog(self):
        p = filedialog.askopenfilename(filetypes=[
            ("PDF and other files", "*.pdf *.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.webp *.txt *.doc *.docx *.xls *.xlsx *.ppt *.pptx *.rtf"),
            ("PDF files", "*.pdf"), ("All files", "*.*")])
        if p:
            self.open_path(p)

    def open_path(self, p):
        k = self._tab_for_path(p) if self.tabs else -1
        if k >= 0:                                    # already open: just show that tab
            self.switch_tab(k)
            return
        try:
            with open(p, "rb") as fh:
                head = fh.read(8)
        except Exception as ex:
            messagebox.showerror(APP_NAME, f"Could not open file:\n{ex}")
            return
        ext = os.path.splitext(p)[1].lower()
        if not head.startswith(b"%PDF") and ext in (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp",
                                                    ".txt", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
                                                    ".rtf", ".odt", ".ods", ".odp"):
            self.open_non_pdf(p)
            return
        try:
            with open(p, "rb") as fh:
                data = fh.read()
            doc = fitz.open(stream=data, filetype="pdf")
            prot = None
            if doc.needs_pass:
                pw = simpledialog.askstring(APP_NAME, "This PDF is password-protected.\nPassword:", show="*")
                if not pw or not doc.authenticate(pw):
                    messagebox.showerror(APP_NAME, "Wrong password.")
                    return
                prot = dict(user_pw=pw, owner_pw=pw, perms=doc.permissions)
        except Exception as ex:
            messagebox.showerror(APP_NAME, f"Could not open file:\n{ex}")
            return
        self._adopt_doc(doc, p)
        self.protect = prot
        self._after_open_check_signed(data)
        self.add_recent(p)
        self._v3_restore_position(p)
        self.status.config(text=f"Opened {os.path.basename(p)}" + ("  (password kept when you save)" if prot else ""))

    def save(self):
        if not self.need_doc():
            return
        if not self.path:
            return self.save_as()
        self._write(self.path)

    def save_as(self):
        if not self.need_doc():
            return
        p = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                         initialfile=os.path.basename(self.path or "edited.pdf"))
        if p:
            self._write(p)
            self.path = p
            self.update_ui()

    def _write(self, p):
        self.commit_floats()
        if self._sig_orig is not None and self.modified:
            if not messagebox.askyesno(APP_NAME, "This document is digitally signed.\nSaving your changes will make the signature invalid.\n\nSave anyway?"):
                return
            self._sig_orig = None
        try:
            data = self._sig_orig if (self._sig_orig is not None) else self.doc.tobytes(**self._save_kwargs())
            tmp = p + ".claudius-tmp"          # write beside the target first, then swap: a failed save never damages the old file
            try:
                with open(tmp, "wb") as fh:
                    fh.write(data)
                    fh.flush()
                    os.fsync(fh.fileno())
                os.replace(tmp, p)
            finally:
                if os.path.exists(tmp):
                    try:
                        os.remove(tmp)
                    except OSError:
                        pass
            self.modified = False
            self.add_recent(p)
            self.status.config(text=f"Saved: {p}")
            self.update_ui()
        except PermissionError:
            messagebox.showerror(APP_NAME, "Windows would not let Claudius PDF Editor save here.\n\nThe file may be open in another "
                                           "program, read-only, or in a protected folder.\nTry Save As and pick another place.")
        except Exception as ex:
            messagebox.showerror(APP_NAME, f"Could not save:\n{ex}")

    def set_ui_language(self, code):
        """save the language and restart the program so every window is built in it"""
        if code == i18n.current:
            return
        if not self.confirm_discard():
            self.lang_var.set(i18n.current)
            return
        i18n.save_setting(code)
        i18n.set_language(code)
        messagebox.showinfo(APP_NAME, "The language has been changed.\nClaudius PDF Editor will restart now.")
        self.restart_app()

    def set_theme(self, code):
        """save light / dark choice and restart so every window is built with it"""
        if code == TH.load_setting():
            return
        if not self.confirm_discard():
            self.theme_var.set(TH.load_setting())
            return
        TH.save_setting(code)
        messagebox.showinfo(APP_NAME, "The appearance has been changed.\nClaudius PDF Editor will restart now.")
        self.restart_app()

    def restart_app(self):
        import subprocess
        args = [sys.executable] if getattr(sys, "frozen", False) else [sys.executable, os.path.abspath(sys.argv[0])]
        args.extend(self.open_paths_list())
        env = dict(os.environ, PYINSTALLER_RESET_ENVIRONMENT="1")
        self.modified = False
        self._v3_save_state()
        self.save_geometry()
        self.stop_speaking()
        self.clear_recovery()
        try:
            subprocess.Popen(args, env=env)
        except Exception as ex:
            messagebox.showerror(APP_NAME, f"Could not restart:\n{ex}\nPlease start Claudius PDF Editor again yourself.")
            return
        self.destroy()

    def on_close(self):
        if self.confirm_discard():
            self._v3_save_state()
            self.save_geometry()
            self.stop_speaking()
            self.clear_recovery()
            self.destroy()

    def export_png(self):
        if not self.need_doc():
            return
        self.commit_floats()
        p = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG", "*.png")],
                                         initialfile=f"page_{self.page_no + 1}.png")
        if p:
            z = min(2.0, safe_zoom(self.page(), 2.0))
            self.page().get_pixmap(matrix=fitz.Matrix(z, z)).save(p)
            self.status.config(text=f"Exported {p}")

    def extract_page(self):
        self.dlg_extract()

    # --------------------------------------------------------------- undo
    def _tool_label(self):
        t = self.tool.get() if hasattr(self, "tool") else ""
        return tr(t.replace("_", " ").capitalize()) if t and t != "view" else tr("Change")

    def _push_snapshot(self):
        self.undo_stack.append(Snap(self.doc.tobytes(encryption=fitz.PDF_ENCRYPT_NONE), self.page_no,
                                    getattr(self, "_snap_label", None) or self._tool_label()))
        self.redo_stack.clear()
        if len(self.undo_stack) > MAX_UNDO:
            self.undo_stack.pop(0)
        self.modified = True

    def snapshot(self):
        self.commit_floats()          # pending photos become part of the page first
        self._push_snapshot()

    def undo(self):
        if self.floats:               # undo removes the most recently added pending photo
            o = self.floats.pop()
            if self.sel is o:
                self.sel = self.floats[-1] if self.floats else None
            self.draw_floats()
            self.status.config(text="Removed last added image.")
            return
        if not self.undo_stack:
            self.status.config(text="Nothing to undo.")
            return
        data, pno = self.undo_stack.pop()
        self.redo_stack.append((self.doc.tobytes(encryption=fitz.PDF_ENCRYPT_NONE), self.page_no))
        self.doc = fitz.open(stream=data, filetype="pdf")
        self.page_no = min(pno, len(self.doc) - 1)
        self.modified = True
        self.build_thumbs()
        self.render()
        self.update_ui()

    def edited(self):
        """Call after a change that affects only the current page."""
        self.render()
        self.refresh_thumb(self.page_no)
        self.update_ui()

    # -------------------------------------------------------------- pages
    def rotate(self, deg):
        self.rotate_selected(deg)

    def delete_page(self):
        self.delete_selected()

    def move_page(self, step):
        if not self.need_doc():
            return
        new = self.page_no + step
        if not 0 <= new < len(self.doc):
            return
        self.snapshot()
        if step > 0:
            self.doc.move_page(self.page_no, new + 1 if new + 1 < len(self.doc) else -1)
        else:
            self.doc.move_page(self.page_no, new)
        self.page_no = new
        self.build_thumbs()
        self.render()
        self.update_ui()

    def insert_blank(self):
        if not self.need_doc():
            return
        self.snapshot()
        r = self.page().rect
        self.doc.new_page(pno=self.page_no + 1, width=r.width, height=r.height)
        self.page_no += 1
        self.build_thumbs()
        self.render()
        self.update_ui()

    def merge_pdf(self):
        self.dlg_insert_pdf()

    def _unused_find(self):
        pass

    def _show_hits(self, rects):
        m = self.page().rotation_matrix
        for r in rects:
            rr = r * m
            self.canvas.create_rectangle(self.off_x + rr.x0 * self.zoom, self.off_y + rr.y0 * self.zoom,
                                         self.off_x + rr.x1 * self.zoom, self.off_y + rr.y1 * self.zoom,
                                         outline="#ff9900", width=2, fill="", tags="pg")

    # ---------------------------------------------------------- thumbnails
    def build_thumbs(self):
        self._cont_raw = None
        self._thumb_job += 1
        job = self._thumb_job
        self.thumbs.delete("all")
        self.thumb_photos = {}
        if not self.doc:
            return
        n = len(self.doc)
        self.thumbs.config(scrollregion=(0, 0, SIDEBAR_W - S(14), n * CELL_H + 10))
        cx = (SIDEBAR_W - S(14)) // 2
        for i in range(n):
            top = i * CELL_H + 12
            round_rect(self.thumbs, cx - THUMB_W // 2 - 6, top - 6, cx + THUMB_W // 2 + 6,
                       top + THUMB_H + 28, S(10), outline="", fill=SIDEBAR_BG, tags=("selbg", f"selbg{i}"))
            self.thumbs.create_text(cx, top + THUMB_H + 14, text=str(i + 1), font=(UI_FONT, 9), fill=TH.SECOND)
        self._select_thumb()
        if self.sidebar_mode == "bookmarks":
            self._refresh_bookmarks()
        self.after(10, lambda: self._thumb_step(0, job))

    def _thumb_step(self, i, job):
        if job != self._thumb_job or not self.doc or i >= len(self.doc):
            return
        self.refresh_thumb(i)
        self.after(1, lambda: self._thumb_step(i + 1, job))

    def refresh_thumb(self, i):
        if not self.doc or not 0 <= i < len(self.doc):
            return
        pg = self.doc[i]
        r = pg.rect
        s = min(THUMB_W / r.width, THUMB_H / r.height)
        pix = pg.get_pixmap(matrix=fitz.Matrix(s, s), alpha=False)
        ph = tk.PhotoImage(data=pix.tobytes("ppm"))
        self.thumb_photos[i] = ph
        self.thumbs.delete(f"img{i}")
        cx = (SIDEBAR_W - S(14)) // 2
        cy = i * CELL_H + 12 + THUMB_H // 2
        self.thumbs.create_image(cx, cy, image=ph, tags=f"img{i}")
        self.thumbs.create_rectangle(cx - pix.width // 2 - 1, cy - pix.height // 2 - 1,
                                     cx + pix.width // 2, cy + pix.height // 2,
                                     outline=TH.LINE, tags=f"img{i}")

    def _select_thumb(self):
        self.thumbs.itemconfig("selbg", fill=SIDEBAR_BG, outline="")
        for i in (self.sel_pages or {self.page_no}):
            self.thumbs.itemconfig(f"selbg{i}", fill=ACTIVE_BG, outline=ACCENT)
        self.thumbs.itemconfig(f"selbg{self.page_no}", fill=ACTIVE_BG, outline=ACCENT)
        total = max(len(self.doc) * CELL_H + 10, 1) if self.doc else 1
        y = self.page_no * CELL_H
        top = self.thumbs.canvasy(0)
        h = self.thumbs.winfo_height()
        if y < top or y + CELL_H > top + h:
            self.thumbs.yview_moveto(max(0, (y - 10) / total))

    def _on_thumb_click(self, e):
        if not self.doc:
            return
        i = int(self.thumbs.canvasy(e.y) // CELL_H)
        if 0 <= i < len(self.doc):
            self.goto(i)

    # ------------------------------------------------------------ viewing
    def goto(self, n):
        if self.doc is None or not 0 <= n < len(self.doc):
            return
        if self.cont:
            return self._cont_goto(n)
        self.commit_floats()
        self.page_no = n
        self.sel_pages = {n}
        if self.fit_mode:
            self._apply_fit(render=False)
        self._select_thumb()
        self.render()
        self.canvas.yview_moveto(0)
        self.update_ui()

    def _on_page_entry(self, _e):
        try:
            self.goto(int(self.page_var.get()) - 1)
        except ValueError:
            pass
        self.update_ui()
        self.canvas.focus_set()

    def set_zoom(self, z, render=True, mode=None):
        self.fit_mode = mode
        self.zoom = max(0.2, min(4.0, z))
        self._sync_zoom = True
        self.zoom_var.set(self.zoom * 100)
        self._sync_zoom = False
        self.zoom_lbl.config(text=f"{int(self.zoom * 100)}%")
        if render:
            self.render()

    def _on_scale(self, v):
        if self._sync_zoom or not self.doc:
            return
        if self._zoom_after:
            self.after_cancel(self._zoom_after)
        self._zoom_after = self.after(60, lambda: self.set_zoom(float(v) / 100))

    def fit_width(self, render=True):
        if not self.doc:
            return
        cw = max(self.canvas.winfo_width(), 400) - 2 * PAD - 20
        self.set_zoom(cw / self.page().rect.width, render, mode="width")

    def fit_page(self, render=True):
        if not self.doc:
            return
        cw = max(self.canvas.winfo_width(), 400) - 2 * PAD - 20
        ch = max(self.canvas.winfo_height(), 300) - 2 * PAD
        r = self.page().rect
        self.set_zoom(min(cw / r.width, ch / r.height), render, mode="page")

    def _offsets(self):
        cw = max(self.canvas.winfo_width(), 1)
        return max(PAD, (cw - self.page_w) // 2), PAD

    def _update_scrollregion(self):
        cw = max(self.canvas.winfo_width(), 1)
        ch = max(self.canvas.winfo_height(), 1)
        self.canvas.config(scrollregion=(0, 0, max(self.page_w + 2 * PAD, cw),
                                         max(self.page_h + 2 * PAD, ch)))

    def _on_canvas_resize(self, _e):
        if not self.doc:
            self.render()
            return
        if self.fit_mode:
            if self._fit_after:
                self.after_cancel(self._fit_after)
            self._fit_after = self.after(120, self._refit)
        if self.cont:
            if getattr(self, "_cont_rs", None):
                self.after_cancel(self._cont_rs)
            self._cont_rs = self.after(90, self.render)
            return
        nx, _ = self._offsets()
        dx = nx - self.off_x
        if dx:
            self.canvas.move("pg", dx, 0)
            self.off_x = nx
            if self.floats:
                self.draw_floats()
        self._update_scrollregion()

    def render(self):
        if self.doc is not None and self.cont:
            return self._render_cont()
        self._gen += 1
        self.canvas.delete("all")
        if self.doc is None:
            self._draw_start_screen()
            return
        z = safe_zoom(self.page(), self.zoom)
        if z < self.zoom - 1e-6:                 # a very large page: never build a gigantic bitmap
            self.zoom = max(0.02, z)
            self._sync_zoom = True
            self.zoom_var.set(self.zoom * 100)
            self._sync_zoom = False
            self.zoom_lbl.config(text=f"{int(self.zoom * 100)}%")
            self.status.config(text="Zoom limited because this page is very large.")
        pix = self.page().get_pixmap(matrix=fitz.Matrix(self.zoom, self.zoom), alpha=False)
        if self.night:
            pix.invert_irect(pix.irect)
        self._photo = tk.PhotoImage(data=pix.tobytes("ppm"))
        self.page_w, self.page_h = pix.width, pix.height
        self.off_x, self.off_y = self._offsets()
        for k, col in zip((7, 5, 3, 1), TH.PAGE_SHADOW):
            self.canvas.create_rectangle(self.off_x - k + 2, self.off_y - k + 4,
                                         self.off_x + self.page_w + k - 2, self.off_y + self.page_h + k + 1,
                                         fill=col, outline="", tags="pg")
        self.canvas.create_image(self.off_x, self.off_y, anchor="nw", image=self._photo, tags="pg")
        self._update_scrollregion()
        self._after_render()
        self.draw_floats()
        self._draw_textsel()

    def update_ui(self):
        name = os.path.basename(self.path) if self.path else ""
        star = "● " if self.modified else ""
        self.title(f"{'*' if self.modified else ''}{name + ' - ' if name else ''}{APP_NAME}")
        self.doc_lbl.config(text=f"{star}{name}")
        if self.doc:
            self.page_var.set(str(self.page_no + 1))
            lab = page_label(self.doc, self.page_no)
            self.total_lbl.config(text=f"/ {len(self.doc)}" + (f"  ({lab})" if lab and lab != str(self.page_no + 1) else ""))
        else:
            self.page_var.set("")
            self.total_lbl.config(text="/ -")
        self._tabs_redraw()
        if self.sidebar_mode == "comments":
            self._refresh_comments()

    def _wheel_router(self, e, direction=None):
        """Mouse wheel over the page view scrolls it, over the thumbnails scrolls them."""
        if direction is not None:
            n = 1
            d = direction
        else:
            steps = self.wheel_steps(e.delta)
            if steps == 0:
                return "break"
            n, d = abs(steps), (1 if steps > 0 else -1)
        w = self.winfo_containing(e.x_root, e.y_root)
        if w is None:
            return
        path = str(w)
        if path.startswith(str(self.thumbs.master)):
            self.thumbs.yview_scroll(-3 * d * n, "units")
            return "break"
        if path.startswith(str(self.canvas.master)):
            for _ in range(n):
                self._scroll_view(d, bool(e.state & 0x0004), bool(e.state & 0x0001))
            return "break"

    def _scroll_view(self, d, ctrl=False, shift=False):
        """d > 0 = scroll up / zoom in.  Reaching the end of a page continues on the next / previous one."""
        if ctrl:
            if self.doc:
                self.set_zoom(self.zoom * (1.1 if d > 0 else 1 / 1.1))
            return
        if shift:
            self.canvas.xview_scroll(-3 * d, "units")
            return
        if self.doc is None:
            return
        if self.cont:
            self.canvas.yview_scroll(-3 * d, "units")
            return
        top, bottom = self.canvas.yview()
        now = time.monotonic()
        if (d > 0 and top > 0.001) or (d < 0 and bottom < 0.999):
            self.canvas.yview_scroll(-3 * d, "units")
            self._last_scroll = now
            return
        if now - self._last_scroll < 0.35:           # a short pause at the edge before turning the page
            return
        self._last_scroll = now
        if d < 0 and self.page_no < len(self.doc) - 1:
            self.goto(self.page_no + 1)
        elif d > 0 and self.page_no > 0:
            self.goto(self.page_no - 1)
            self.canvas.update_idletasks()
            self.canvas.yview_moveto(1.0)            # arrive at the bottom of the previous page

    # ---------------------------------------------------------------- logo
    def _load_icons(self):
        self.logo26 = tk.PhotoImage(data=LOGO_26_B64)
        self.logo32 = tk.PhotoImage(data=LOGO_32_B64)
        self.logo128 = tk.PhotoImage(data=LOGO_128_B64)
        try:                                   # window + taskbar icon
            self.iconphoto(True, self.logo128, self.logo32)
        except tk.TclError:
            pass

    def show_about(self):
        dlg = tk.Toplevel(self)
        dlg.title(f"{tr('About')} {APP_NAME}")
        dlg.transient(self)
        dlg.resizable(False, False)
        dlg.configure(bg=TH.CARD, padx=48, pady=28)
        dlg.geometry("+%d+%d" % (self.winfo_rootx() + 380, self.winfo_rooty() + 130))
        tk.Label(dlg, image=self.logo128, bg=TH.CARD).pack()
        tk.Label(dlg, text=APP_NAME, font=(UI_FONT, 20, "bold"), bg=TH.CARD, fg=TH.LABEL).pack(pady=(10, 0))
        tk.Label(dlg, text=f"{tr('Version')} {APP_VERSION}", font=(UI_FONT, 10), bg=TH.CARD, fg=TH.SECOND).pack(pady=(4, 0))
        tk.Label(dlg, text=tr("Created by {}").replace("{}", "Michal Czapla & Claude AI"),
                 font=(UI_FONT, 10, "bold"), bg=TH.CARD, fg=TH.LABEL).pack(pady=(14, 0))
        tk.Label(dlg, text=tr(APP_FREEWARE), font=(UI_FONT, 10), bg=TH.CARD, fg=TH.LABEL).pack(pady=(6, 0))
        link = tk.Label(dlg, text=APP_URL.replace("https://", ""), font=(UI_FONT, 10, "underline"), bg=TH.CARD,
                        fg=ACCENT, cursor="hand2")
        link.pack(pady=(2, 0))
        link.bind("<Button-1>", lambda e: webbrowser.open(APP_URL))
        tk.Label(dlg, text=f"Powered by PyMuPDF {fitz.VersionBind}",
                 font=(UI_FONT, 9), bg=TH.CARD, fg=TH.TERTIARY).pack(pady=(14, 0))
        ttk.Button(dlg, text="Close", command=dlg.destroy).pack(pady=(18, 0))

    # ================================================================ shapes
    def pick_fill_color(self):
        c = colorchooser.askcolor(color=self._hex(self.fill_color), title="Fill color")
        if c and c[0]:
            self.fill_color = tuple(v / 255 for v in c[0])
            for sw in self.fill_swatches:
                sw.config(bg=self._hex(self.fill_color), activebackground=self._hex(self.fill_color))
            self.fill_on.set(True)

    def shape_style_group(self, parent):
        g = self.group(parent, "Style")
        f = tk.Frame(g, bg=RIBBON_BG)
        f.pack(padx=6, pady=(2, 0))

        def lab(t, r, c):
            tk.Label(f, text=t, font=(UI_FONT, 8), bg=RIBBON_BG, fg=TH.LABEL).grid(
                row=r, column=c, sticky="w", padx=(10 if c else 0, 4), pady=1)

        lab("Line", 0, 0)
        sw = tk.Button(f, width=4, relief="solid", bd=1, bg=self._hex(self.color),
                       activebackground=self._hex(self.color), command=self.pick_color, cursor="hand2")
        sw.grid(row=0, column=1, pady=1)
        self.swatches.append(sw)
        lab("Width", 1, 0)
        ttk.Spinbox(f, from_=1, to=20, width=4, textvariable=self.line_width).grid(row=1, column=1, pady=1)
        tk.Checkbutton(f, text="Fill", variable=self.fill_on, bg=RIBBON_BG, activebackground=RIBBON_BG,
                       font=(UI_FONT, 8)).grid(row=0, column=2, sticky="w", padx=(10, 0))
        fsw = tk.Button(f, width=4, relief="solid", bd=1, bg=self._hex(self.fill_color),
                        activebackground=self._hex(self.fill_color), command=self.pick_fill_color, cursor="hand2")
        fsw.grid(row=0, column=3, pady=1)
        self.fill_swatches.append(fsw)
        lab("Opacity %", 1, 2)
        ttk.Spinbox(f, from_=10, to=100, increment=10, width=4, textvariable=self.shape_opacity).grid(
            row=1, column=3, pady=1)

    def draw_shape_preview(self, kind, a_canvas, b_canvas, shift):
        cv = self.canvas
        cv.delete("shp")
        lw = self.line_width.get()
        items = shape_items(kind, self._c2v(*a_canvas), self._c2v(*b_canvas), lw, shift)
        col = self._hex(self.color)
        fillc = self._hex(self.fill_color) if self.fill_on.get() else ""
        w = max(1, lw * self.zoom)

        def flat(pts):
            out = []
            for p in pts:
                out.extend(self._v2c(*p))
            return out

        for typ, data, _ in items:
            if typ == "line":
                cv.create_line(*flat(data), fill=col, width=w, capstyle="round", tags="shp")
            elif typ == "open":
                cv.create_line(*flat(data), fill=col, width=w, capstyle="round", joinstyle="round", tags="shp")
            elif typ == "poly":
                cv.create_polygon(*flat(data), outline=col, fill=fillc, width=w, tags="shp")
            elif typ == "head":
                cv.create_polygon(*flat(data), outline=col, fill=col, tags="shp")
            else:
                x0, y0 = self._v2c(data[0], data[1])
                x1, y1 = self._v2c(data[2], data[3])
                if typ == "oval":
                    cv.create_oval(x0, y0, x1, y1, outline=col, fill=fillc, width=w, tags="shp")
                else:
                    cv.create_rectangle(x0, y0, x1, y1, outline=col, fill=fillc, width=w, tags="shp")

    def commit_shape(self, kind, a_canvas, b_canvas, shift):
        lw = self.line_width.get()
        items = shape_items(kind, self._c2v(*a_canvas), self._c2v(*b_canvas), lw, shift)
        if not items:
            return
        self.snapshot()
        pg = self.page()
        m = pg.derotation_matrix
        op = max(0.05, min(1.0, self.shape_opacity.get() / 100))
        fill = self.fill_color if self.fill_on.get() else None

        def P(p):
            return fitz.Point(p[0], p[1]) * m

        def R(box):
            r = fitz.Rect(P((box[0], box[1])), P((box[2], box[3])))
            r.normalize()
            return r

        sh = pg.new_shape()
        base = dict(width=lw, color=self.color, lineCap=1, lineJoin=1, stroke_opacity=op, fill_opacity=op)
        for typ, data, fillable in items:
            if typ == "line":
                sh.draw_line(P(data[0]), P(data[1]))
                sh.finish(closePath=False, **base)
            elif typ == "open":
                sh.draw_polyline([P(p) for p in data])
                sh.finish(closePath=False, **base)
            elif typ == "poly":
                pts = [P(p) for p in data]
                sh.draw_polyline(pts + [pts[0]])
                sh.finish(fill=fill, closePath=True, **base)
            elif typ == "head":
                pts = [P(p) for p in data]
                sh.draw_polyline(pts + [pts[0]])
                sh.finish(fill=self.color, closePath=True, **base)
            elif typ == "oval":
                sh.draw_oval(R(data))
                sh.finish(fill=fill, closePath=True, **base)
            elif typ == "rrect":
                sh.draw_rect(R(data), radius=0.2)
                sh.finish(fill=fill, closePath=True, **base)
            else:
                sh.draw_rect(R(data))
                sh.finish(fill=fill, closePath=True, **base)
        sh.commit()
        self.edited()

    def _apply_fit(self, render=True):
        if self.fit_mode == "page":
            self.fit_page(render)
        elif self.fit_mode == "width":
            self.fit_width(render)

    def _refit(self):
        self._fit_after = None
        if self.doc and self.fit_mode:
            self._apply_fit(render=True)

    # ================================================================ photos
    def _v2c(self, x, y):
        return self.off_x + x * self.zoom, self.off_y + y * self.zoom

    def _c2v(self, cx, cy):
        return (cx - self.off_x) / self.zoom, (cy - self.off_y) / self.zoom

    @staticmethod
    def f_corners(o):
        th = math.radians(o.angle)
        c, s = math.cos(th), math.sin(th)
        out = []
        for dx, dy in ((-o.w / 2, -o.h / 2), (o.w / 2, -o.h / 2), (o.w / 2, o.h / 2), (-o.w / 2, o.h / 2)):
            out.append((o.cx + dx * c - dy * s, o.cy + dx * s + dy * c))
        return out

    @classmethod
    def f_bbox(cls, o):
        pts = cls.f_corners(o)
        xs, ys = [p[0] for p in pts], [p[1] for p in pts]
        return min(xs), min(ys), max(xs), max(ys)

    @staticmethod
    def f_hit(o, x, y):
        th = math.radians(-o.angle)
        dx, dy = x - o.cx, y - o.cy
        rx = dx * math.cos(th) - dy * math.sin(th)
        ry = dx * math.sin(th) + dy * math.cos(th)
        return abs(rx) <= o.w / 2 and abs(ry) <= o.h / 2

    def float_preview(self, o):
        pw, ph = max(1, int(o.w * self.zoom)), max(1, int(o.h * self.zoom))
        key = (pw, ph, o.angle, o.flip_h, o.flip_v, o.opacity, o.brightness, o.contrast,
               o.transp, o.transp_thr, o.gray, o.bw, id(o.work))
        if o.pv_key == key and o.pv_photo is not None:
            return o.pv_photo
        small = o.work.resize((pw, ph), BILINEAR if self.fdrag else LANCZOS)
        o.pv_photo = ImageTk.PhotoImage(process_image(small, o))
        o.pv_key = key
        return o.pv_photo

    def draw_floats(self):
        self.canvas.delete("fl")
        self._handles = {}
        if not self.doc or not self.floats:
            return
        tags = ("pg", "fl")
        for o in self.floats:
            ph = self.float_preview(o)
            cx, cy = self._v2c(o.cx, o.cy)
            self.canvas.create_image(cx, cy, image=ph, tags=tags)
            if o.border:
                pts = []
                for x, y in self.f_corners(o):
                    pts.extend(self._v2c(x, y))
                self.canvas.create_polygon(*pts, outline=self._hex(self.color), fill="",
                                           width=max(1, self.line_width.get() * self.zoom), tags=tags)
        o = self.sel
        if o in self.floats:
            pts = []
            for x, y in self.f_corners(o):
                pts.extend(self._v2c(x, y))
            self.canvas.create_polygon(*pts, outline=ACCENT, fill="", width=1, tags=tags)
            bx0, by0, bx1, by1 = self.f_bbox(o)
            x0, y0 = self._v2c(bx0, by0)
            x1, y1 = self._v2c(bx1, by1)
            self.canvas.create_rectangle(x0, y0, x1, y1, outline=ACCENT, dash=(3, 3), tags=tags)
            self._handles = {"nw": (x0, y0), "ne": (x1, y0), "se": (x1, y1), "sw": (x0, y1)}
            mx = (x0 + x1) / 2
            ry = y0 - 28 if y0 - 28 >= 6 else y1 + 28
            self.canvas.create_line(mx, y0 if ry < y0 else y1, mx, ry, fill=ACCENT, tags=tags)
            self._handles["rot"] = (mx, ry)
            self.canvas.create_oval(mx - 7, ry - 7, mx + 7, ry + 7, fill=ACCENT, outline="white", width=2, tags=tags)
            for k in ("nw", "ne", "se", "sw"):
                hx, hy = self._handles[k]
                self.canvas.create_rectangle(hx - 5, hy - 5, hx + 5, hy + 5, fill="white",
                                             outline=ACCENT, width=2, tags=tags)

    def select(self, o):
        self.sel = o
        if o is not None:
            self.angle_var.set(int(round(o.angle)))
            self.status.config(text="Drag to move  •  corners to resize  •  blue dot to rotate  •  "
                                    "Enter = apply to page  •  Delete = remove")
        self.draw_floats()

    def need_sel(self):
        if self.sel is None or self.sel not in self.floats:
            messagebox.showinfo(APP_NAME, "Click an image on the page first.\n"
                                          "(Use Add Image or Paste to place one.)")
            return False
        return True

    # ---- adding photos
    def add_images(self):
        if self.doc is None:
            self.photos_to_pages()
            return
        paths = filedialog.askopenfilenames(title="Add images to this page", filetypes=IMG_TYPES)
        failed = []
        for p in paths:
            try:
                self.place_image(load_image(p), os.path.basename(p))
            except Exception as ex:
                failed.append(f"{os.path.basename(p)}: {ex}")
        if failed:
            messagebox.showerror(APP_NAME, "Could not open:\n{}".format("\n".join(failed)))

    def paste_image(self):
        if not self.need_doc():
            return
        try:
            data = ImageGrab.grabclipboard()
        except Exception as ex:
            messagebox.showerror(APP_NAME, f"Clipboard not available:\n{ex}")
            return
        if data is None:
            self.status.config(text="No image in the clipboard. Copy a picture (or an image file) first.")
            return
        if isinstance(data, (list, tuple)):
            for p in data:
                try:
                    self.place_image(load_image(p), os.path.basename(p))
                except Exception:
                    pass
        else:
            self.place_image(data.convert("RGBA"), "clipboard")

    def place_image(self, img, name="image"):
        if self.floats and self.float_page != self.page_no:
            self.commit_floats()
        self.float_page = self.page_no
        r = self.page().rect
        w = min(r.width * 0.5, img.width * 0.75)
        h = w * img.height / img.width
        if h > r.height * 0.7:
            h = r.height * 0.7
            w = h * img.width / img.height
        o = FloatImage(img, name)
        o.w, o.h = w, h
        k = len(self.floats) % 6
        o.cx, o.cy = r.width / 2 + 18 * k, r.height / 2 + 18 * k
        self.floats.append(o)
        self.modified = True
        self.set_tool("view")
        self.select(o)
        self.update_ui()

    # ---- applying to the page
    def commit_floats(self):
        if not self.floats:
            self.sel = None
            return
        if not self.doc or not 0 <= self.float_page < len(self.doc):
            self.floats, self.sel = [], None
            return
        pno = self.float_page
        self._push_snapshot()
        pg = self.doc[pno]
        dpi, q = QUALITY[self.quality_var.get()]
        errors = []
        for o in self.floats:
            try:
                self._flatten(pg, o, dpi, q)
            except Exception as ex:
                errors.append(f"{o.name}: {ex}")
        self.floats, self.sel = [], None
        self.canvas.delete("fl")
        if pno == self.page_no:
            self.render()
        self.refresh_thumb(pno)
        self.update_ui()
        if errors:
            messagebox.showerror(APP_NAME, "Could not place:\n{}".format("\n".join(errors)))
        else:
            self.status.config(text="Image applied to the page.")

    def _flatten(self, pg, o, dpi, q):
        th = math.radians(o.angle)
        W = abs(o.w * math.cos(th)) + abs(o.h * math.sin(th))
        H = abs(o.w * math.sin(th)) + abs(o.h * math.cos(th))
        full = o.base
        if dpi:
            tw = max(1, min(full.width, int(o.w / 72 * dpi)))
            thp = max(1, int(round(tw * full.height / full.width)))
            if (tw, thp) != full.size:
                full = full.resize((tw, thp), LANCZOS)
        im = process_image(full, o)
        alpha = im.getchannel("A")
        has_alpha = alpha.getextrema()[0] < 255
        buf = io.BytesIO()
        im.convert("RGB").save(buf, "JPEG", quality=q, optimize=True)
        mask = None
        if has_alpha:
            mb = io.BytesIO()
            alpha.save(mb, "PNG")
            mask = mb.getvalue()
        rect = fitz.Rect(o.cx - W / 2, o.cy - H / 2, o.cx + W / 2, o.cy + H / 2) * pg.derotation_matrix
        rect.normalize()
        pg.insert_image(rect, stream=buf.getvalue(), mask=mask, rotate=pg.rotation,
                        keep_proportion=False, overlay=True)
        if o.border:
            pts = [fitz.Point(x, y) * pg.derotation_matrix for x, y in self.f_corners(o)]
            sh = pg.new_shape()
            sh.draw_polyline(pts + [pts[0]])
            sh.finish(color=self.color, width=self.line_width.get(), closePath=False)
            sh.commit()

    def apply_images(self):
        if not self.floats:
            self.status.config(text="No pending images.")
            return
        self.commit_floats()

    # ---- editing the selected photo
    def _touch(self):
        self.modified = True
        self.draw_floats()
        self.update_ui()

    def img_rotate(self, deg):
        if self.need_sel():
            o = self.sel
            a = (o.angle + deg) % 360
            o.angle = a - 360 if a > 180 else a
            self.angle_var.set(int(round(o.angle)))
            self._touch()

    def _angle_changed(self, *_):
        if self.sel in self.floats:
            try:
                self.sel.angle = float(self.angle_var.get())
            except (tk.TclError, ValueError):
                return
            self._touch()

    def img_flip(self, horizontal):
        if self.need_sel():
            if horizontal:
                self.sel.flip_h = not self.sel.flip_h
            else:
                self.sel.flip_v = not self.sel.flip_v
            self._touch()

    def img_toggle(self, attr):
        if self.need_sel():
            o = self.sel
            setattr(o, attr, not getattr(o, attr))
            if attr == "bw" and o.bw:
                o.gray = False
            if attr == "gray" and o.gray:
                o.bw = False
            self._touch()

    def img_fit_page(self):
        if self.need_sel():
            o = self.sel
            r = self.page().rect
            m = 18
            bx0, by0, bx1, by1 = self.f_bbox(o)
            s = min((r.width - 2 * m) / (bx1 - bx0), (r.height - 2 * m) / (by1 - by0))
            o.w *= s
            o.h *= s
            o.cx, o.cy = r.width / 2, r.height / 2
            self._touch()

    def img_duplicate(self):
        if self.need_sel():
            o = copy.copy(self.sel)
            o.id = FloatImage._next
            FloatImage._next += 1
            o.cx += 14
            o.cy += 14
            o.pv_key = None
            o.pv_photo = None
            self.floats.append(o)
            self.select(o)
            self._touch()

    def img_delete(self):
        if self.need_sel():
            self.floats.remove(self.sel)
            self.sel = self.floats[-1] if self.floats else None
            self.draw_floats()
            self.status.config(text="Image removed.")

    def img_order(self, front):
        if self.need_sel():
            o = self.sel
            self.floats.remove(o)
            if front:
                self.floats.append(o)
            else:
                self.floats.insert(0, o)
            self._touch()

    def img_reset(self):
        if self.need_sel():
            o = self.sel
            aspect = o.pristine.height / o.pristine.width
            o.set_base(o.pristine)
            o.reset_adjust()
            o.h = o.w * aspect
            self.angle_var.set(0)
            self._touch()

    def start_crop(self):
        if self.need_sel():
            self.set_tool("crop")
            self.status.config(text="Crop: drag a rectangle over the part of the image you want to keep.")

    def crop_selected(self, x0, y0, x1, y1):
        o = self.sel
        if o not in self.floats:
            return
        vx0, vy0 = self._c2v(min(x0, x1), min(y0, y1))
        vx1, vy1 = self._c2v(max(x0, x1), max(y0, y1))
        bx0, by0, bx1, by1 = self.f_bbox(o)
        ix0, iy0, ix1, iy1 = max(vx0, bx0), max(vy0, by0), min(vx1, bx1), min(vy1, by1)
        if ix1 - ix0 < 4 or iy1 - iy0 < 4:
            return
        full = process_image(o.base, o, use_opacity=False)
        sx, sy = full.width / (bx1 - bx0), full.height / (by1 - by0)
        box = (int((ix0 - bx0) * sx), int((iy0 - by0) * sy), int((ix1 - bx0) * sx), int((iy1 - by0) * sy))
        new = full.crop(box)
        opacity, border = o.opacity, o.border
        o.set_base(new)
        o.reset_adjust()
        o.opacity, o.border = opacity, border
        o.w, o.h = ix1 - ix0, iy1 - iy0
        o.cx, o.cy = (ix0 + ix1) / 2, (iy0 + iy1) / 2
        self.angle_var.set(0)
        self._touch()

    def adjust_dialog(self):
        if not self.need_sel():
            return
        o = self.sel
        dlg = tk.Toplevel(self)
        dlg.title("Adjust image")
        dlg.transient(self)
        dlg.geometry("+%d+%d" % (self.winfo_rootx() + 140, self.winfo_rooty() + 120))
        dlg.resizable(False, False)
        dlg.configure(bg=RIBBON_BG, padx=14, pady=10)
        vars_ = {"opacity": tk.IntVar(value=o.opacity), "brightness": tk.IntVar(value=o.brightness),
                 "contrast": tk.IntVar(value=o.contrast), "transp_thr": tk.IntVar(value=o.transp_thr)}
        transp = tk.BooleanVar(value=o.transp)

        def apply(*_):
            o.opacity = vars_["opacity"].get()
            o.brightness = vars_["brightness"].get()
            o.contrast = vars_["contrast"].get()
            o.transp_thr = vars_["transp_thr"].get()
            o.transp = transp.get()
            self._touch()

        def row(r, text, key, lo, hi):
            tk.Label(dlg, text=text, bg=RIBBON_BG, font=(UI_FONT, 9), anchor="w", width=14).grid(row=r, column=0, sticky="w")
            tk.Scale(dlg, from_=lo, to=hi, orient="horizontal", length=240, variable=vars_[key],
                     command=apply, bg=RIBBON_BG, highlightthickness=0).grid(row=r, column=1, pady=2)

        row(0, "Opacity %", "opacity", 0, 100)
        row(1, "Brightness", "brightness", -100, 100)
        row(2, "Contrast", "contrast", -100, 100)
        tk.Checkbutton(dlg, text="Remove white background", variable=transp, command=apply, bg=RIBBON_BG,
                       font=(UI_FONT, 9)).grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))
        row(4, "Sensitivity", "transp_thr", 150, 254)

        def reset():
            for k, v in (("opacity", 100), ("brightness", 0), ("contrast", 0), ("transp_thr", 235)):
                vars_[k].set(v)
            transp.set(False)
            apply()

        bf = tk.Frame(dlg, bg=RIBBON_BG)
        bf.grid(row=5, column=0, columnspan=2, pady=(10, 0), sticky="e")
        ttk.Button(bf, text="Reset", command=reset).pack(side="left", padx=4)
        ttk.Button(bf, text="Close", command=dlg.destroy).pack(side="left")

    # ---- mouse handling for floating images
    def _float_press(self, e):
        if not self.floats:
            return False
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        o = self.sel
        if o in self.floats:
            for name, (hx, hy) in self._handles.items():
                if abs(cx - hx) <= 9 and abs(cy - hy) <= 9:
                    if name == "rot":
                        self.fdrag = {"mode": "rotate", "o": o}
                    else:
                        bx0, by0, bx1, by1 = self.f_bbox(o)
                        opp = {"nw": (bx1, by1), "ne": (bx0, by1), "se": (bx0, by0), "sw": (bx1, by0)}[name]
                        corner = {"nw": (bx0, by0), "ne": (bx1, by0), "se": (bx1, by1), "sw": (bx0, by1)}[name]
                        cs = self.f_corners(o)
                        ci = min(range(4), key=lambda k: math.dist(cs[k], self._c2v(hx, hy)))
                        self.fdrag = {"mode": "scale", "o": o, "opp": opp, "d0": max(1e-3, math.dist(corner, opp)),
                                      "g0": (o.cx, o.cy, o.w, o.h), "ci": ci, "P": cs[(ci + 2) % 4]}
                    return True
        vx, vy = self._c2v(cx, cy)
        for o in reversed(self.floats):
            if self.f_hit(o, vx, vy):
                self.select(o)
                self.fdrag = {"mode": "move", "o": o, "last": (vx, vy)}
                return True
        if self.sel is not None:
            self.select(None)
        return False

    def _float_drag(self, e):
        d = self.fdrag
        o = d["o"]
        vx, vy = self._c2v(self.canvas.canvasx(e.x), self.canvas.canvasy(e.y))
        if d["mode"] == "move":
            o.cx += vx - d["last"][0]
            o.cy += vy - d["last"][1]
            d["last"] = (vx, vy)
        elif d["mode"] == "scale":
            cx0, cy0, w0, h0 = d["g0"]
            ox, oy = d["opp"]
            if bool(self.keep_aspect.get()) == bool(e.state & 0x0001):   # Shift inverts the option
                th = math.radians(o.angle)
                ux, uy, vx2, vy2 = math.cos(th), math.sin(th), -math.sin(th), math.cos(th)
                px, py = d["P"]
                dx, dy = vx - px, vy - py
                sx = 1 if d["ci"] in (1, 2) else -1
                sy = 1 if d["ci"] in (2, 3) else -1
                nw = max(12.0, (dx * ux + dy * uy) * sx)
                nh = max(12.0, (dx * vx2 + dy * vy2) * sy)
                o.w, o.h = nw, nh
                o.cx = px + (ux * sx * nw + vx2 * sy * nh) / 2
                o.cy = py + (uy * sx * nw + vy2 * sy * nh) / 2
                self.draw_floats()
                return
            s = math.dist((vx, vy), (ox, oy)) / d["d0"]
            s = max(s, 12 / min(w0, h0))
            o.w, o.h = w0 * s, h0 * s
            o.cx, o.cy = ox + (cx0 - ox) * s, oy + (cy0 - oy) * s
        else:
            a = math.degrees(math.atan2(vx - o.cx, -(vy - o.cy)))
            if e.state & 0x0001:
                a = round(a / 15) * 15
            else:
                for snap in (-180, -90, 0, 90, 180):
                    if abs(a - snap) < 4:
                        a = snap
            o.angle = 0.0 if abs(a) < 1e-6 else (-180 if a == 180 else a)
            self.angle_var.set(int(round(o.angle)))
        self.draw_floats()

    def _float_release(self, _e):
        o = self.fdrag["o"]
        self.fdrag = None
        o.pv_key = None
        self.modified = True
        self.draw_floats()
        self.update_ui()

    def _on_right_click(self, e):
        if self.presenting:
            self._pres_click(-1)
            return
        if self._text_menu(e):
            return
        if not self.floats:
            self._cont_focus_at(e)
            self.annot_menu(e)
            return
        vx, vy = self._c2v(self.canvas.canvasx(e.x), self.canvas.canvasy(e.y))
        for o in reversed(self.floats):
            if self.f_hit(o, vx, vy):
                self.select(o)
                m = tk.Menu(self, tearoff=0)
                m.add_command(label="Apply to page", command=self.apply_images)
                m.add_command(label="Duplicate", command=self.img_duplicate)
                m.add_command(label="Fit to page", command=self.img_fit_page)
                self.align_menu(m, self.img_align)
                m.add_separator()
                m.add_command(label="Rotate right", command=lambda: self.img_rotate(90))
                m.add_command(label="Rotate left", command=lambda: self.img_rotate(-90))
                m.add_command(label="Flip horizontally", command=lambda: self.img_flip(True))
                m.add_command(label="Flip vertically", command=lambda: self.img_flip(False))
                m.add_separator()
                m.add_command(label="Bring to front", command=lambda: self.img_order(True))
                m.add_command(label="Send to back", command=lambda: self.img_order(False))
                m.add_separator()
                m.add_command(label="Delete", command=self.img_delete)
                m.tk_popup(e.x_root, e.y_root)
                return
        self.annot_menu(e)

    def _key_arrow(self, dx, dy, e):
        if self._typing():
            return
        if self.presenting:
            self._pres_key(dx, dy)
            return
        if self.sel is not None and self.sel in self.floats:      # nudge the selected photo
            step = 10 if (e.state & 0x0001) else 1
            self.sel.cx += dx * step
            self.sel.cy += dy * step
            self._touch()
        elif self.doc is not None:                                  # otherwise scroll the page
            if dy:
                self._scroll_view(-dy)
            if dx:
                self.canvas.xview_scroll(dx * 3, "units")

    def _key_enter(self, _e):
        if not self._typing() and self.floats:
            self.apply_images()

    def _key_paste(self, _e):
        if not self._typing():
            self.paste_image()

    # ---- photos -> PDF pages
    def photos_to_pages(self):
        paths = filedialog.askopenfilenames(title="Choose photos to turn into PDF pages", filetypes=IMG_TYPES)
        if not paths:
            return
        files = list(paths)
        dlg = tk.Toplevel(self)
        dlg.title("Photos to PDF pages")
        dlg.transient(self)
        dlg.geometry("+%d+%d" % (self.winfo_rootx() + 140, self.winfo_rooty() + 120))
        dlg.grab_set()
        dlg.configure(bg=RIBBON_BG, padx=14, pady=12)

        tk.Label(dlg, text="One page per photo, in this order:", bg=RIBBON_BG, font=(UI_FONT, 9, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w")
        lb = tk.Listbox(dlg, width=58, height=10, selectmode="browse", activestyle="none")
        lb.grid(row=1, column=0, rowspan=4, sticky="nsew", pady=4)

        def refill(sel=None):
            lb.delete(0, "end")
            for p in files:
                lb.insert("end", os.path.basename(p))
            if sel is not None and files:
                sel = max(0, min(sel, len(files) - 1))
                lb.selection_set(sel)
                lb.see(sel)

        def cur():
            s = lb.curselection()
            return s[0] if s else None

        def up():
            i = cur()
            if i:
                files[i - 1], files[i] = files[i], files[i - 1]
                refill(i - 1)

        def down():
            i = cur()
            if i is not None and i < len(files) - 1:
                files[i + 1], files[i] = files[i], files[i + 1]
                refill(i + 1)

        def remove():
            i = cur()
            if i is not None:
                files.pop(i)
                refill(i)

        def add_more():
            more = filedialog.askopenfilenames(parent=dlg, filetypes=IMG_TYPES)
            files.extend(more)
            refill()

        def sort_names():
            files.sort(key=lambda p: os.path.basename(p).lower())
            refill()

        for r, (t, c) in enumerate((("▲ Up", up), ("▼ Down", down), ("Remove", remove), ("Add more…", add_more))):
            ttk.Button(dlg, text=t, command=c, width=11).grid(row=1 + r, column=1, padx=(8, 0), pady=1, sticky="n")
        ttk.Button(dlg, text="Sort by name", command=sort_names, width=11).grid(row=5, column=1, padx=(8, 0), pady=1)
        refill(0)

        opts = tk.Frame(dlg, bg=RIBBON_BG)
        opts.grid(row=6, column=0, columnspan=2, sticky="w", pady=(10, 0))
        size_v = tk.StringVar(value="A4")
        ori_v = tk.StringVar(value="Auto")
        mar_v = tk.IntVar(value=18)
        qual_v = tk.StringVar(value=self.quality_var.get())
        where_v = tk.StringVar(value="Insert after current page" if self.doc else "Create a new PDF")
        where_vals = ["Insert after current page", "Add at end of document", "Create a new PDF"] if self.doc \
            else ["Create a new PDF"]

        def lab(t, r, c):
            tk.Label(opts, text=t, bg=RIBBON_BG, font=(UI_FONT, 9)).grid(row=r, column=c, sticky="w", padx=(0, 6), pady=2)

        lab("Page size", 0, 0)
        ttk.Combobox(opts, textvariable=size_v, state="readonly", width=24,
                     values=["A4", "Letter", "Match photo shape"]).grid(row=0, column=1, padx=(0, 18))
        lab("Orientation", 0, 2)
        ttk.Combobox(opts, textvariable=ori_v, state="readonly", width=10,
                     values=["Auto", "Portrait", "Landscape"]).grid(row=0, column=3)
        lab("Margin (pt)", 1, 0)
        ttk.Spinbox(opts, from_=0, to=100, width=6, textvariable=mar_v).grid(row=1, column=1, sticky="w")
        lab("Quality", 1, 2)
        ttk.Combobox(opts, textvariable=qual_v, state="readonly", width=18,
                     values=list(QUALITY)).grid(row=1, column=3)
        lab("Where", 2, 0)
        ttk.Combobox(opts, textvariable=where_v, state="readonly", width=24,
                     values=where_vals).grid(row=2, column=1, columnspan=1)

        def go():
            if not files:
                return
            args = (list(files), size_v.get(), ori_v.get(), mar_v.get(), where_v.get(), qual_v.get())
            dlg.destroy()
            self.build_pages_from_photos(*args)

        bf = tk.Frame(dlg, bg=RIBBON_BG)
        bf.grid(row=7, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(bf, text="Create pages", command=go).pack(side="left", padx=4)
        ttk.Button(bf, text="Cancel", command=dlg.destroy).pack(side="left")

    def build_pages_from_photos(self, files, size, orient, margin, where, quality):
        dpi, q = QUALITY[quality]
        if self.doc is None or where == "Create a new PDF":
            seed = fitz.open()                 # a placeholder page keeps the viewer happy; it is removed again below
            seed.new_page(width=595, height=842)
            self._next_title = "Photos"
            self._adopt_doc(seed, None, modified=True)
            placeholder = True
            start = 0
        else:
            placeholder = False
            self.snapshot()
            start = self.page_no + 1 if where.startswith("Insert") else len(self.doc)
        added, failed = 0, []
        for p in files:
            try:
                im = ImageOps.exif_transpose(Image.open(p))
                if im.mode in ("RGBA", "LA", "P"):
                    rgba = im.convert("RGBA")
                    bg = Image.new("RGB", rgba.size, (255, 255, 255))
                    bg.paste(rgba, mask=rgba.getchannel("A"))
                    im = bg
                else:
                    im = im.convert("RGB")
                iw, ih = im.size
                if size == "Match photo shape":
                    pw, ph = (841.89, 841.89 * ih / iw) if iw >= ih else (841.89 * iw / ih, 841.89)
                else:
                    pw, ph = (595.28, 841.89) if size == "A4" else (612.0, 792.0)
                    land = iw > ih if orient == "Auto" else orient == "Landscape"
                    if land:
                        pw, ph = ph, pw
                m = max(0, min(margin, min(pw, ph) / 3))
                s = min((pw - 2 * m) / iw, (ph - 2 * m) / ih)
                dw, dh = iw * s, ih * s
                if dpi:
                    tw = max(1, min(iw, int(dw / 72 * dpi)))
                    if tw < iw:
                        im = im.resize((tw, max(1, int(round(tw * ih / iw)))), LANCZOS)
                buf = io.BytesIO()
                im.save(buf, "JPEG", quality=q, optimize=True)
                page = self.doc.new_page(pno=start + added, width=pw, height=ph)
                rect = fitz.Rect((pw - dw) / 2, (ph - dh) / 2, (pw + dw) / 2, (ph + dh) / 2)
                page.insert_image(rect, stream=buf.getvalue(), keep_proportion=False)
                added += 1
                self.status.config(text=f"Adding photos… {added}/{len(files)}")
                self.update_idletasks()
            except Exception as ex:
                failed.append(f"{os.path.basename(p)}: {ex}")
        if placeholder:
            if added:
                self.doc.delete_page(len(self.doc) - 1)
            else:
                self.modified = False
                self.close_tab()
        if added:
            self.modified = True
            self.page_no = min(start, len(self.doc) - 1)
            self.update_idletasks()
            self.build_thumbs()
            self.fit_page(render=False)
            self.render()
            self.update_ui()
            self.status.config(text=f"Added {added} photo page(s).")
        if failed:
            messagebox.showerror(APP_NAME, "Could not add:\n{}".format("\n".join(failed)))


    # ------------------------------------------------------ mouse / tools
    def _on_press(self, e):
        if self.doc is None:
            return
        if self.presenting:
            self._pres_click(1)
            return
        t = self.tool.get()
        if t == "view":
            if self.cont and self.floats:
                self._cont_focus(self.float_page)
            if self._float_press(e):
                return
            self._cont_focus_at(e)
            self._view_click = (e.x, e.y)
            if self._text_press(e):
                return
            self.canvas.scan_mark(e.x, e.y)
            return
        self._cont_focus_at(e)
        self._drag_start = (self.canvas.canvasx(e.x), self.canvas.canvasy(e.y))
        if t == "text":
            self._add_text(e)
            self._drag_start = None
        elif t == "pen":
            self._stroke = [self._drag_start]

    def _on_drag(self, e):
        if self.doc is None:
            return
        if self.fdrag:
            self._float_drag(e)
            return
        t = self.tool.get()
        if t == "view":
            if self._tsel_drag:
                self._text_drag(e)
            else:
                self.canvas.scan_dragto(e.x, e.y, gain=1)
            return
        if not self._drag_start:
            return
        x, y = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        if t == "pen":
            px, py = self._stroke[-1]
            self.canvas.create_line(px, py, x, y, fill=self._hex(self.color),
                                    width=self.line_width.get() * self.zoom, tags="preview")
            self._stroke.append((x, y))
        elif t in SHAPE_TOOLS:
            self.draw_shape_preview(t, self._drag_start, (x, y), bool(e.state & 1))
        elif t in ("highlight", "whiteout", "image", "crop", "pagecrop") + tuple(RECT_TOOLS):
            if self._drag_item:
                self.canvas.delete(self._drag_item)
            self._drag_item = self.canvas.create_rectangle(
                *self._drag_start, x, y, outline=ACCENT, dash=(4, 3), width=2)

    def _on_release(self, e):
        if self.doc is None:
            return
        if self.fdrag:
            self._float_release(e)
            return
        if self.tool.get() == "view":
            if self._text_release(e):
                return
            if self._view_release(e):
                return
        if not self._drag_start:
            return
        t = self.tool.get()
        x0, y0 = self._drag_start
        x1, y1 = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        self._drag_start = None
        self.canvas.delete("preview")
        if self._drag_item:
            self.canvas.delete(self._drag_item)
            self._drag_item = None
        if t == "crop":
            self.crop_selected(x0, y0, x1, y1)
            self.set_tool("view")
            return
        if t == "pagecrop":
            if abs(x1 - x0) > 8 and abs(y1 - y0) > 8:
                self.crop_drag(x0, y0, x1, y1)
            return
        if t in EXTRA_TOOLS:
            try:
                self._extra_release(t, x0, y0, x1, y1, bool(e.state & 1))
            except Exception as ex:
                messagebox.showerror(APP_NAME, f"That did not work:\n{ex}")
            return
        self.canvas.delete("shp")
        if t in SHAPE_TOOLS:
            if math.hypot(x1 - x0, y1 - y0) > 4:
                try:
                    self.commit_shape(t, (x0, y0), (x1, y1), bool(e.state & 1))
                except Exception as ex:
                    messagebox.showerror(APP_NAME, f"Could not draw shape:\n{ex}")
            return
        try:
            if t == "pen":
                if len(self._stroke) > 1:
                    self._add_stroke(self._stroke)
            elif abs(x1 - x0) > 3 and abs(y1 - y0) > 3:
                rect = self.to_pdf_rect(x0, y0, x1, y1)
                if t == "highlight":
                    self.markup_rect("highlight", rect)
                    return
                elif t == "rect":
                    self.snapshot()
                    sh = self.page().new_shape()
                    sh.draw_rect(rect)
                    sh.finish(color=self.color, width=self.line_width.get())
                    sh.commit()
                elif t == "whiteout":
                    self.snapshot()
                    sh = self.page().new_shape()
                    sh.draw_rect(rect)
                    sh.finish(color=(1, 1, 1), fill=(1, 1, 1), width=0)
                    sh.commit()
                elif t == "image":
                    self._add_image(rect)
                    return
                else:
                    return
        except Exception as ex:
            messagebox.showerror(APP_NAME, f"Could not apply edit:\n{ex}")
            return
        self.edited()

    def _add_stroke(self, pts):
        self.snapshot()
        m = self.page().derotation_matrix
        mapped = [fitz.Point((cx - self.off_x) / self.zoom, (cy - self.off_y) / self.zoom) * m
                  for cx, cy in pts]
        sh = self.page().new_shape()
        sh.draw_polyline(mapped)
        sh.finish(color=self.color, width=self.line_width.get(), closePath=False)
        sh.commit()

    def _add_text(self, e):
        txt = simpledialog.askstring(APP_NAME, "Text to insert:")
        if not txt:
            return
        pt = self.to_pdf_point(e.x, e.y)
        pg = self.page()
        kwargs = dict(fontsize=self.font_size.get(), color=self.color, rotate=pg.rotation)
        if os.path.exists(WIN_ARIAL):
            kwargs.update(fontname="arial", fontfile=WIN_ARIAL)
        else:
            kwargs.update(fontname="helv")
        try:
            self.snapshot()
            pg.insert_text(pt, txt, **kwargs)
        except Exception as ex:
            self.undo_stack.pop()
            messagebox.showerror(APP_NAME, f"Could not insert text:\n{ex}")
            return
        self.edited()

    def _add_image(self, rect):
        p = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp *.gif"),
                                                  ("All files", "*.*")])
        if not p:
            return
        try:
            self.snapshot()
            self.page().insert_image(rect, filename=p, rotate=self.page().rotation)
        except Exception as ex:
            self.undo_stack.pop()
            messagebox.showerror(APP_NAME, f"Could not insert image:\n{ex}")
            return
        self.edited()


def main():
    paths = [a for a in sys.argv[1:] if not a.startswith("-")]
    PdfEditor(paths).mainloop()


if __name__ == "__main__":
    main()
