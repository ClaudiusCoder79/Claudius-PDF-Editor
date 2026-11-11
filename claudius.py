"""
Claudius PDF Editor
Windows 10 x64 (also works on macOS / Linux)

Requirements:  Python 3.9+   and   pip install pymupdf pillow
               (optional, for PDF<->Word/Excel/PowerPoint:  pip install python-docx openpyxl python-pptx)
Keep claudius_extra.py in the same folder as this file.
Run:           python claudius.py   (or double-click run.bat)
"""
import time as _time
_T0 = _time.time()                # start-up timing (written to ~/.claudius/startup.log)
import copy
import io
import math
import os
import sys
import time
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk, filedialog, messagebox, simpledialog, colorchooser


# ---- first sign of life: a small window that appears BEFORE the heavy modules (PyMuPDF, Pillow ...) are loaded
_PRE = None


_CENTRE_DBG = ""


def _work_area(win):
    """usable screen area (x, y, w, h) - the primary screen without the taskbar"""
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    if sys.platform == "win32":
        try:
            import ctypes
            from ctypes import wintypes
            rc = wintypes.RECT()
            if ctypes.windll.user32.SystemParametersInfoW(0x30, 0, ctypes.byref(rc), 0) and rc.right > rc.left:
                return rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top
        except Exception:
            pass
    return 0, 0, sw, sh


def _centre(win, w=None, h=None, lift=0):
    """put a window exactly in the middle of the screen; measured again after it is shown, because the real size can differ
    from the requested one on high-DPI screens"""
    global _CENTRE_DBG
    win.update_idletasks()
    w = w or win.winfo_reqwidth()
    h = h or win.winfo_reqheight()
    ax, ay, aw, ah = _work_area(win)
    win.geometry("%dx%d+%d+%d" % (w, h, ax + (aw - w) // 2, ay + max(0, (ah - h) // 2 - lift)))
    win.update()
    rw, rh = win.winfo_width(), win.winfo_height()
    if abs(rw - w) > 2 or abs(rh - h) > 2:                       # the window got another size than asked: correct the position
        win.geometry("+%d+%d" % (ax + (aw - rw) // 2, ay + max(0, (ah - rh) // 2 - lift)))
        win.update()
    _CENTRE_DBG = "  window %dx%d at %d,%d on screen %dx%d (area %dx%d)" % (win.winfo_width(), win.winfo_height(), win.winfo_x(),
                                                                          win.winfo_y(), win.winfo_screenwidth(), win.winfo_screenheight(), aw, ah)

SPLASH_COLOURS = ("#FF3B30", "#FF9500", "#FFCC00", "#34C759", "#30B0C7", "#007AFF", "#5856D6", "#AF52DE")


def _splash_ui(parent, scale, bg, fg, sub, tert, font, title, version, free, starting):
    """the start-up card: coloured app tile, name, version, a strip in the iOS system colours.  Used by the very first
    window and by the full one, so nothing jumps when one replaces the other.  Plain Tk only (no Pillow needed yet)."""
    sc = lambda v: int(round(v * scale))
    W = sc(360)
    card = tk.Frame(parent, bg=bg, padx=0, pady=0)
    cv = tk.Canvas(card, width=W, height=sc(268), bg=bg, highlightthickness=0)
    cv.pack()
    cx, t, sz, rad = W // 2, sc(26), sc(84), sc(20)
    x0, x1 = cx - sz // 2, cx + sz // 2

    def rr(a, b, c, d, r, **kw):
        pts = [a + r, b, c - r, b, c, b, c, b + r, c, d - r, c, d, c - r, d, a + r, d, a, d, a, d - r, a, b + r, a, b]
        return cv.create_polygon(pts, smooth=True, **kw)
    rr(x0 + sc(2), t + sc(5), x1 + sc(2), t + sz + sc(5), rad, fill="#d8d8de" if bg == "#ffffff" else "#1a1a1c", outline="")   # soft shadow
    rr(x0, t, x1, t + sz, rad, fill="#2C8CFF", outline="")
    pw, ph = int(sz * 0.50), int(sz * 0.72)                                                      # white page like the program icon
    rr(cx - pw // 2, t + sz // 2 - ph // 2, cx + pw // 2, t + sz // 2 + ph // 2, sc(4), fill="white", outline="")
    cv.create_text(cx, t + sz // 2 - sc(3), text="C", font=(font, 28, "bold"), fill="#0A6EF0")
    for i, c in enumerate(SPLASH_COLOURS):                                                       # its colour strip
        cv.create_rectangle(cx - pw // 2 + sc(7) + i * (pw - sc(14)) // 8, t + sz // 2 + ph // 2 - sc(11),
                            cx - pw // 2 + sc(7) + (i + 1) * (pw - sc(14)) // 8 - 1, t + sz // 2 + ph // 2 - sc(8), fill=c, outline="")
    y = t + sz + sc(22)
    cv.create_text(cx, y, text=title, font=(font, 19, "bold"), fill=fg)
    cv.create_text(cx, y + sc(30), text=version, font=(font, 10), fill=sub)
    cv.create_text(cx, y + sc(56), text=free, font=(font, 9), fill=tert)
    n, seg = len(SPLASH_COLOURS), sc(26)
    bx = cx - n * seg // 2
    for i, c in enumerate(SPLASH_COLOURS):
        cv.create_rectangle(bx + i * seg, y + sc(82), bx + (i + 1) * seg, y + sc(86), fill=c, outline="")
    cv.create_text(cx, y + sc(106), text=starting, font=(font, 9), fill=sub)
    return card



def _pre_splash():
    scale = 1.0
    if sys.platform == "win32":
        try:                                            # same high-DPI setup as claudius_extra (must happen before any window exists)
            import ctypes
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                ctypes.windll.user32.SetProcessDPIAware()
            scale = max(1.0, ctypes.windll.user32.GetDpiForSystem() / 96.0)
        except Exception:
            pass
    dark = False
    if sys.platform == "win32":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as k:
                dark = winreg.QueryValueEx(k, "AppsUseLightTheme")[0] == 0
        except Exception:
            pass
    bg, fg, sub, line = ("#2c2c2e", "#f2f2f7", "#aeaeb2", "#48484a") if dark else ("#ffffff", "#1c1c1e", "#636366", "#c7c7cc")
    r = tk.Tk()
    r.overrideredirect(True)
    r.configure(bg=line, padx=1, pady=1)
    r.tk.call("tk", "scaling", scale * 96 / 72)
    font = "Segoe UI" if sys.platform == "win32" else "TkDefaultFont"
    _splash_ui(r, scale, bg, fg, sub, "#8e8e93", font, "Claudius PDF Editor", "", "Free for personal and commercial use, source on GitHub",
               "Starting...").pack()
    r.attributes("-topmost", True)
    _centre(r)
    return r


def _drop_pre_splash():
    global _PRE
    if _PRE is not None:
        try:
            _PRE.destroy()
        except Exception:
            pass
        _PRE = None


if __name__ == "__main__" and not os.environ.get("CLAUDIUS_NO_SPLASH"):
    try:
        _PRE = _pre_splash()
    except Exception:
        _PRE = None

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
    from claudius_extra import webbrowser
    from claudius_extra import (CONFIG_DIR, safe_zoom, safe_dpi, SCALE, S, ExtraTools, EXTRA_HINTS, EXTRA_TOOLS, RECT_TOOLS, HL_COLORS, CONVERT_KINDS)
except ImportError:
    sys.exit("claudius_extra.py is missing - keep it in the same folder as claudius.py")

from claudius_cad import CadTools, CAD_TOOLS, CAD_HINTS
from claudius_palette import PaletteTools

ZOOM_MIN, ZOOM_MAX = 0.2, 32.0           # 20 % ... 3200 %
TILE_PIXELS = 12_000_000                 # a page that would need a bigger bitmap than this is drawn in pieces (only what is visible)
APP_NAME = "Claudius PDF Editor"
APP_VERSION = "4.8"
APP_URL = "https://github.com/ClaudiusCoder79/Claudius-PDF-Editor"
APP_FREEWARE = "Free for personal and commercial use, source on GitHub"
MAX_UNDO = 25
WIN_ARIAL = r"C:\Windows\Fonts\arial.ttf"

# ---- look & feel -----------------------------------------------------------
from claudius_theme import round_rect, rr_points, PillTab
from claudius_features import FeatureTools, page_label
from claudius_view import ViewTools
from claudius_v3 import V3Tools
from claudius_text import TextTools
from claudius_frame import FrameTools
from claudius_quick import QuickTools
from claudius_update import UpdateTools
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
    "iVBORw0KGgoAAAANSUhEUgAAABoAAAAaCAYAAACpSkzOAAAFD0lEQVR42pWWX4hcdxXHP+f87p29Mzv7J60So2misjYJVAumpOIf"
    "ECsKkpdIKG3xpYIt9aX4IFsRuhR86YsQEFpakEKhoVBYhfqizZtUCIUiCG1sqViizT+zm9nJ3Jl77+8cH34zs7ObDcYLhx/38jt/"
    "v99zzpU1d31OxJ560/ctzduqo6ct2mF3DeDC//WIi1jUoP8U7I0bN/X5F07Kxpq7CsDT5/xIS1jP5jhWD8EaT2p34Mb91m+aCXkB"
    "zYj3KufUmYfkgjzzpu8rg51vL+hK2Ys1QiaIIEIdHfPbOwkKuYI5zMbkuOM07cWQl1v2YTvqiawvtlp0dKW/ERtRyRkbbpqaA12n"
    "yARnt6H03q/g8iBQ5ErckZoIkPc3YlMshpV+31YzM390NMANghiIQjkyvrQw4OffXkT19vW71HNe/MuQi2WLTisQzXcHFEYD3Mwf"
    "1Rj9YF2ZmCHRHDOoGmN5zlFVoss4/iQ+FhCCKiePFhzuVGyVDThE86mYIXVlEqMfVEHUzIkGZuDjs47jIsyUqzGIls5mfCe68oOj"
    "BQfbFf1hBE/6Nr5r5giieqVvNJbinI3GZ2ruY7AzTZKPz+4c3D0Pn15QHj/R5lB3xM3KMU82xJ3GnCt9I7uyBZuls6+A7pwQJGU1"
    "8eMOqimj9Xdrzr3f4MB3jmSc/mrOYjEhh3Ko2/D36w3tLKeKzubI2RhC1UCGO8Ma/l1BHpylQrDo2ITXAhsD5+GXbvLWOzWduxWA"
    "F/4w5GtfyXnjiQ6fWVRUYFgbN0dOGY1e6VQx8U+BrBkbFGBUw9UGbgyMa59K34PAU2dL3nq35sUn53nkgRwReP2dmmfXSy71nM8t"
    "p5iuD5yP/mMsdVLpVVKPGZDFuJOyWUhAVnV6v3DZeP3titVTbZ78Vmt67yffbPH411tkmkAPmghiloKOlhxM7cY9Wj+aYWOQ3r8U"
    "0eicvC+jidsTITqo7BxB7k6MThMTGXaMpRhhL7FxOHfNC+bw8aaTheQgemKdys55aAbNbezpLKVnZaJ//FDg8MHAc78rubLlzGXJ"
    "yeWe88v1If/aNCbDw3c17KzcghGeIpiUpNMSXvpRh+//us/RtR6PnWghAq+dr7h+1fjh8ZzPLm9nNM3CdmG/J0bRpxjVEb53LOPP"
    "v+jy7O+H/PbtCgce/HzgVz+d5/ihQB0hD0wbdSI7HUW/JaNmZj0oqae+8cXAuZ/Ns1mm9lxuyzQLnahaIkOMezky9mBdmg6TRlIE"
    "IxmcOJiySWfI4El3IreWzndhZE6c3HzlVbz9V/S+IVeXu7w8up+ifICPP/kC+6+VfLnX5967Otx7eokYx9XYnZFAFhs3RHQnRqAq"
    "ia4PfZeQPYjMG9088LAtEGyBcr8yV7XpxhZzQYmWdCZE2JGRu2WNcVFCdg+xAhHBE+Ou9xtUYe6eA8ABANrAykR5gXExs6m9jX6N"
    "e5Ya1j11cGjhVl/MzPystIpVr4YNErKI024pf/pbyWNnLrDUDvh41blsr3X37U0lCP1hwx8/UDqdLlUzGf0WpVVkXldnZemZzX3a"
    "6HkpFla8vFEDGSLiFun1y+0p/j8eVWGxWyCaMV5mjbSXch9ufWiZnRCAhadvHJE8X5esfczrAVgNIgQV7vTHzkkkSgssR/IO3pTv"
    "eV2f2jqzdCGw5lo9X1wr7n/itUZz3Dng5ovuLtFMErB+B2K44+4SHf2HWf1yGFz7ce83+y+y5vpfl5lZT5UgzYMAAAAASUVORK5C"
    "YII="
)
LOGO_32_B64 = (
    "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"
)
LOGO_128_B64 = (
    "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"
    "Dso9Riw9WwAAAABJRU5ErkJggg=="
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


def constrain_shape(kind, a, b, shift):
    """the end point b of a shape dragged from a: Shift / circle keep a square box, lines snap to 15 degree steps"""
    ax, ay = a
    bx, by = b
    if kind in ("line", "arrow", "darrow"):
        if shift:
            ang = math.atan2(by - ay, bx - ax)
            ang = round(ang / (math.pi / 12)) * (math.pi / 12)
            d = math.hypot(bx - ax, by - ay)
            bx, by = ax + d * math.cos(ang), ay + d * math.sin(ang)
        return (ax, ay), (bx, by)
    if kind == "circle" or shift:
        dx, dy = bx - ax, by - ay
        side = max(abs(dx), abs(dy))
        bx = ax + math.copysign(side, dx if dx else 1)
        by = ay + math.copysign(side, dy if dy else 1)
    return (ax, ay), (bx, by)


class FloatShape:
    """A shape (line, arrow, box, star ...) that floats above the page until it is applied: it can still be moved,
    resized and rotated, like a pending photo."""
    kind_label = "shape"

    def __init__(self, kind, a, b, lw, color, fill, opacity):
        self.id = FloatImage._next
        FloatImage._next += 1
        self.kind, self.lw, self.color, self.fill = kind, lw, tuple(color), (tuple(fill) if fill else None)
        self.name = kind
        self.opacity = int(opacity)
        (ax, ay), (bx, by) = a, b
        self.cx, self.cy = (ax + bx) / 2, (ay + by) / 2
        self.w, self.h = max(abs(bx - ax), 12.0), max(abs(by - ay), 12.0)
        self.a_n = ((ax - self.cx) / self.w, (ay - self.cy) / self.h)
        self.b_n = ((bx - self.cx) / self.w, (by - self.cy) / self.h)
        self.angle = 0.0
        # the image-only settings exist so that the common picture commands do no harm
        self.flip_h = self.flip_v = self.border = False
        self.brightness = self.contrast = 0
        self.transp = self.gray = self.bw = False
        self.transp_thr = 235
        self.pv_key = self.pv_photo = None

    def local_items(self):
        a = (self.a_n[0] * self.w, self.a_n[1] * self.h)
        b = (self.b_n[0] * self.w, self.b_n[1] * self.h)
        return shape_items(self.kind, a, b, self.lw, False)

    def world_items(self):
        """the geometry in page (view) space, rotated and moved; boxes and ovals become polygons when rotated"""
        th = math.radians(self.angle)
        c, s = math.cos(th), math.sin(th)
        rotated = abs(self.angle) % 360 > 0.01

        def T(p):
            return (self.cx + p[0] * c - p[1] * s, self.cy + p[0] * s + p[1] * c)
        out = []
        for typ, data, fillable in self.local_items():
            if typ in ("line", "open", "poly", "head"):
                out.append((typ, [T(p) for p in data], fillable))
            else:                                           # rect / rrect / oval: data = (l, t, r, b)
                l, t, r, b = data
                if not rotated:
                    out.append((typ, (self.cx + l, self.cy + t, self.cx + r, self.cy + b), fillable))
                    continue
                if typ == "oval":
                    mx, my, rx, ry = (l + r) / 2, (t + b) / 2, (r - l) / 2, (b - t) / 2
                    pts = [(mx + rx * math.cos(2 * math.pi * k / 72), my + ry * math.sin(2 * math.pi * k / 72)) for k in range(72)]
                elif typ == "rrect":
                    rad = 0.2 * min(r - l, b - t)
                    pts = []
                    for (ccx, ccy, a0) in ((r - rad, t + rad, -90), (r - rad, b - rad, 0), (l + rad, b - rad, 90), (l + rad, t + rad, 180)):
                        for k in range(9):
                            ang = math.radians(a0 + k * 10)
                            pts.append((ccx + rad * math.cos(ang), ccy + rad * math.sin(ang)))
                else:
                    pts = [(l, t), (r, t), (r, b), (l, b)]
                out.append(("poly", [T(p) for p in pts], fillable))
        return out


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


TAB_NAMES = ("Home", "Comment", "Edit", "Pictures", "Organize", "Forms", "Convert", "Document", "CAD", "View")
RIBBON_FLAT = True                       # simple flat buttons (File > Settings can switch back to the coloured tiles)
ICON_STYLE = os.environ.get("CLAUDIUS_ICONSTYLE", "mono")        # mono | colour | tile  (flat toolbar icons)
SYSTEM_COLOURS = {"blue": "#007AFF", "green": "#34C759", "orange": "#FF9500", "red": "#FF3B30", "purple": "#AF52DE",
                  "teal": "#30B0C7", "indigo": "#5856D6", "pink": "#FF2D55"}
_ICON_RULES = (                                          # first matching word in the label wins
    ("red", ("delete", "redact", "whiteout", "clean", "remove", "unprotect", "clear")),
    ("green", ("save", "apply", "check", "accessib", "pdf/a", "sign", "date", "protect", "verify")),
    ("orange", ("highlight", "underline", "strike", "squiggly", "sticky", "stamp", "pen", "shapes", "text box", "note", "comment", "show list", "more")),
    ("purple", ("image", "picture", "photo", "adjust", "crop", "flip", "border", "gray", "b & w", "no white", "reset", "align", "front", "back")),
    ("teal", ("rotate", "move", "extract", "split", "blank", "insert", "merge", "page", "undo history", "folder")),
    ("indigo", ("convert", "png", "word", "excel", "export", "to ", "ocr", "optimi", "compare", "attach", "propert", "settings", "guided", "watermark", "background", "header")),
)


def icon_colour(label):
    t = (label or "").lower()
    for name, words in _ICON_RULES:
        if any(w in t for w in words):
            return SYSTEM_COLOURS[name]
    return SYSTEM_COLOURS["blue"]


def mix_hex(c, o, t):
    a = [int(c[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(o[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02x%02x%02x" % tuple(round(x + (y - x) * t) for x, y in zip(a, b))


class RibbonButton(tk.Canvas):
    """Ribbon button: a glossy 3-D icon tile over a label. It lifts when hovered and pushes in when pressed."""

    def __init__(self, parent, icon, text, command, icon_size=17, width=60, color=None, arrow=False):
        self._arrow = arrow
        need = tkfont.Font(family=UI_FONT, size=8).measure(tr(text) + (" \u25be" if arrow else "")) + S(16)      # longer translations get wider buttons
        self._bw, self._bh = max(S(width), need), S(58) if RIBBON_FLAT else S(64)
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
        label = tr(self._text) + (" \u25be" if self._arrow else "") if self._arrow else self._text
        if RIBBON_FLAT:                                   # simple look: no tile, grey line-style glyph, light highlight
            cx, cy = self._bw // 2, S(21)
            if state != "normal":
                round_rect(self, S(2), S(1), self._bw - S(2), self._bh - S(2), S(8),
                           fill=TH.TINT if state != "hover" else HOVER_BG, outline="")
            on = state in ("active", "pressed", "hover")
            col = ACCENT if on else TH.SECOND
            if ICON_STYLE != "mono":
                col = icon_colour(self._text)
                if ICON_STYLE == "tile":                  # soft tinted rounded square behind a coloured glyph
                    base = TH.CARD if TH.IS_DARK else "#FFFFFF"
                    round_rect(self, cx - S(17), cy - S(15), cx + S(17), cy + S(15), S(9),
                               fill=mix_hex(base, col, 0.30 if TH.IS_DARK else 0.14), outline="")
            self.create_text(cx, cy + (S(1) if state == "pressed" else 0), text=self._icon, font=(ICON_FONT, self._isize),
                             fill=col)
            self.create_text(cx, S(46), text=label, font=(UI_FONT, 8), fill=ACCENT if self.active else TH.LABEL)
            return
        self._tile = TH.tile_photo(self, S(36), self._color, state, S)
        cx, cy = self._bw // 2, S(24)
        if self.active:
            round_rect(self, S(2), S(1), self._bw - S(2), self._bh - S(2), S(12), fill=HOVER_BG, outline="")
        lift = S(1) if state == "pressed" else 0
        self.create_image(cx, cy, image=self._tile)
        # glyph with a faint shadow so it looks embossed into the tile
        self.create_text(cx + 1, cy + S(1) + lift + 1, text=self._icon, font=(ICON_FONT, self._isize), fill="#14305c")
        self.create_text(cx, cy + lift, text=self._icon, font=(ICON_FONT, self._isize), fill="white")
        self.create_text(self._bw // 2, S(54), text=label, font=(UI_FONT, 8),
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
        self.active = False
        self._tipwin = self._tipjob = None
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)
        self.bind("<ButtonPress-1>", self._press)
        self.bind("<ButtonRelease-1>", self._release)
        self._paint()

    def _paint(self):
        self.delete("all")
        state = "pressed" if self.pressed else "hover" if self.hover else "normal"
        c = (self._qs + 2 * S(4)) // 2
        if RIBBON_FLAT:
            if state != "normal" or self.active:
                round_rect(self, S(1), S(1), 2 * c - S(1), 2 * c - S(1), S(7), fill=TH.TINT if (self.active or state == "pressed") else HOVER_BG, outline="")
            self.create_text(c, c + (S(1) if state == "pressed" else 0), text=self._icon, font=(ICON_FONT, 12),
                             fill=ACCENT if (self.active or state != "normal") else TH.SECOND)
            return
        self._img = TH.tile_photo(self, self._qs, self._color, state, S)
        lift = S(1) if self.pressed else 0
        if self.active:                                   # switched on: a ring around the tile
            self.create_rectangle(1, 1, 2 * c - 1, 2 * c - 1, outline=ACCENT, width=2)
        self.create_image(c, c, image=self._img)
        self.create_text(c + 1, c + lift + 1, text=self._icon, font=(ICON_FONT, 11), fill="#14305c")
        self.create_text(c, c + lift, text=self._icon, font=(ICON_FONT, 11), fill="white")

    def set_active(self, on):
        self.active = bool(on)
        self._paint()

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


class PdfEditor(PaletteTools, CadTools, TextTools, FrameTools, QuickTools, UpdateTools, V3Tools, V3bTools, ViewTools, FeatureTools, ProTools, ExtraTools, tk.Tk):
    def __init__(self, path=None):
        self._t_build = _time.time()
        super().__init__()
        if _PRE is not None:
            tk._default_root = self                      # the early splash window must not become the home of images and variables
        self._splash = None
        if not os.environ.get("CLAUDIUS_NO_SPLASH"):
            self.withdraw()
            self._show_splash()
        paths = [path] if isinstance(path, str) else list(path or [])
        if SCALE != 1:
            self.tk.call("tk", "scaling", SCALE * 96 / 72)
        self.title(APP_NAME)
        _ax, _ay, _aw, _ah = _work_area(self)
        _w, _h = min(S(1320), _aw - S(40)), min(S(820), _ah - S(60))
        self.geometry("%dx%d+%d+%d" % (_w, _h, _ax + (_aw - _w) // 2, _ay + (_ah - _h) // 2))      # centred on the screen
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
        globals()["RIBBON_FLAT"] = bool(self.cfg.get("ribbon_flat", True))
        if not os.environ.get("CLAUDIUS_ICONSTYLE"):
            globals()["ICON_STYLE"] = self.cfg.get("icon_style", "colour") if self.cfg.get("icon_style", "colour") in ("mono", "colour", "tile") else "colour"
        self._load_icons()
        self._build_topbar()
        self._build_ribbons()
        self._build_body()
        self._build_bottombar()
        self._bind_keys()
        self.show_tab("Home")
        self.set_tool("view")
        self._v3_after_ui()
        self.after(400, self.qt_restore)
        self.after(8000, self.upd_schedule)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self._win11_init()
        self.restore_geometry()
        self.render()
        self.update_ui()
        self.after(200, self._refit)
        if not os.environ.get("CLAUDIUS_NO_RECOVERY"):
            self.after(700, self.start_recovery)
        self._close_splash()
        self._log_startup()
        if paths:
            self.update()                       # the window must be on screen first: password and error boxes would stay invisible otherwise
            for p in paths:
                self.open_path(p)

    def _log_startup(self):
        """one line per start in ~/.claudius/startup.log: seconds since the process began, program loading, window building"""
        try:
            now = _time.time()
            since = ""
            line = "%s  v%s  program loaded in %.1fs, window ready %.1fs after that%s\n" % (
                time.strftime("%Y-%m-%d %H:%M:%S"), APP_VERSION, self._t_build - _T0, now - self._t_build,
                _CENTRE_DBG + ("  (single file)" if getattr(sys, "_MEIPASS", None) and not os.path.isdir(os.path.join(os.path.dirname(sys.executable), "_internal")) else ""))
            os.makedirs(CONFIG_DIR, exist_ok=True)
            p = os.path.join(CONFIG_DIR, "startup.log")
            try:
                old = open(p, encoding="utf-8").read().splitlines()[-29:]
            except Exception:
                old = []
            with open(p, "w", encoding="utf-8") as f:
                f.write("\n".join(old + [line.rstrip("\n")]) + "\n")
        except Exception:
            pass

    # ------------------------------------------------------------ splash screen
    def _show_splash(self):
        """small start-up window (centred on the screen and scaled for high-DPI; the EXE build has no splash of its own any more)"""
        try:
            import pyi_splash                     # inside the EXE the build's own splash is already on screen - do not show a second one
            return
        except Exception:
            pass
        try:
            sp = tk.Toplevel(self)
            sp.overrideredirect(True)
            sp.configure(bg=TH.LINE if hasattr(TH, "LINE") else "#c7c7cc", padx=1, pady=1)
            _splash_ui(sp, SCALE, "#2c2c2e" if TH.IS_DARK else "#ffffff", TH.LABEL, TH.SECOND, TH.TERTIARY, UI_FONT, APP_NAME,
                       f"{tr('Version')} {APP_VERSION}", tr(APP_FREEWARE), tr("Starting...")).pack()
            sp.attributes("-topmost", True)
            _centre(sp)
            sp.configure(bg=TH.LINE if not TH.IS_DARK else "#48484a")
            self._splash = sp
            _drop_pre_splash()
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
        self.qt_top_btn = QuickButton(qa, "★", f"{tr('Show / hide the quick toolbar')}  (Ctrl+Shift+Q)", self.qt_toggle,
                                      TH.TILE_COLORS[4 % len(TH.TILE_COLORS)])
        self.qt_top_btn.pack(side="left", padx=(8, 1), pady=S(3))
        self.pal_btn = QuickButton(qa, "\U0001F50D", f"{tr('Find tools and commands')}  (Ctrl+K)", self.pal_open, TH.TILE_COLORS[5 % len(TH.TILE_COLORS)])
        self.pal_btn.pack(side="left", padx=1, pady=S(3))
        self._name_lbl = name_lbl
        self.bind("<Configure>", self._topbar_fit, add="+")
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
        self.file_menu.add_command(label="Check for updates...", command=lambda: self.upd_check(True))
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

        for name in TAB_NAMES:
            pill = PillTab(top, name, lambda n=name: self.show_tab(n), S, bg=BAR_BG)
            pill.pack(side="left", padx=1, pady=S(6))
            self.tab_labels[name] = pill

        self.doc_lbl = tk.Label(top, text="", bg=BAR_BG, fg=TH.SECOND, font=(UI_FONT, 9))
        self.doc_lbl.pack(side="right", padx=14)
        tk.Frame(self, bg=RIBBON_LINE, height=1).pack(side="top", fill="x")

    def _topbar_fit(self, e=None):
        """on narrower windows the program name is hidden so that all tabs stay visible"""
        if e is not None and e.widget is not self:
            return
        try:
            narrow = self.winfo_width() < 1620
            shown = self._name_lbl.winfo_manager() == "pack"
            if narrow and shown:
                self._name_lbl.pack_forget()
            elif not narrow and not shown:
                self._name_lbl.pack(side="left", padx=(7, 0))
        except Exception:
            pass

    def show_tab(self, name):
        self._ensure_ribbon(name)
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
        if not RIBBON_FLAT:
            tk.Label(outer, text=tr(caption) if False else caption, font=(UI_FONT, 8), bg=RIBBON_BG, fg=TH.TERTIARY).pack(side="bottom")
        tk.Frame(parent, bg=RIBBON_LINE, width=1).pack(side="left", fill="y", pady=S(12) if RIBBON_FLAT else 8, padx=4)
        return inner

    def rb(self, parent, icon, text, cmd, tool=None, size=17, width=60):
        if tool:
            b = RibbonButton(parent, icon, text, lambda t=tool: self.set_tool(t), size, width)
            self.tool_buttons.setdefault(tool, []).append(b)
        else:
            b = RibbonButton(parent, icon, text, cmd, size, width)
        b.pack(side="left", padx=1)
        self.qt_register(b, icon, text, cmd, tool, size, width)
        return b

    def rbm(self, parent, icon, text, items, size=17, width=60):
        """ribbon button that opens a drop-down menu.  items: (label, tool, command) tuples, None = separator"""
        def popup():
            m = tk.Menu(self, tearoff=0)
            for it in items:
                if it is None:
                    m.add_separator()
                    continue
                label, tool, cmd = it
                m.add_command(label=tr(label), command=(lambda t=tool: self.set_tool(t)) if tool else cmd)
            m.tk_popup(b.winfo_rootx(), b.winfo_rooty() + b.winfo_height())
        b = RibbonButton(parent, icon, text, popup, size, width, arrow=True)
        b._tools = {it[1] for it in items if it and it[1]}
        self.__dict__.setdefault("menu_buttons", []).append(b)
        b.pack(side="left", padx=1)
        reg = self.__dict__.setdefault("qt_reg", {})
        for it in items:                                   # every menu entry can be found by Ctrl+K and put on the quick toolbar
            if it is None:
                continue
            label, tool, cmd = it
            reg["%s|%s|%s" % (label, tool or "", icon)] = {"icon": icon, "text": label, "cmd": cmd, "tool": tool, "size": size,
                                                           "width": width, "color": None, "btn": b}
        return b

    def combined_style_group(self, parent):
        """one compact style box for markup, notes and shapes"""
        g = self.group(parent, "Style")
        f = tk.Frame(g, bg=RIBBON_BG)
        f.pack(padx=6, pady=(2, 0))

        def lab(t, r, c):
            tk.Label(f, text=tr(t), font=(UI_FONT, 8), bg=RIBBON_BG, fg=TH.LABEL).grid(
                row=r, column=c, sticky="w", padx=(10 if c else 0, 4), pady=1)

        lab("Color", 0, 0)
        sw = tk.Button(f, width=4, relief="solid", bd=1, bg=self._hex(self.color),
                       activebackground=self._hex(self.color), command=self.pick_color, cursor="hand2")
        sw.grid(row=0, column=1, pady=1)
        self.swatches.append(sw)
        lab("Width", 1, 0)
        ttk.Spinbox(f, from_=1, to=20, width=4, textvariable=self.line_width).grid(row=1, column=1, pady=1)
        tk.Checkbutton(f, text=tr("Fill"), variable=self.fill_on, bg=RIBBON_BG, activebackground=RIBBON_BG,
                       font=(UI_FONT, 8)).grid(row=0, column=2, sticky="w", padx=(10, 0))
        fsw = tk.Button(f, width=4, relief="solid", bd=1, bg=self._hex(self.fill_color),
                        activebackground=self._hex(self.fill_color), command=self.pick_fill_color, cursor="hand2")
        fsw.grid(row=0, column=3, pady=1)
        self.fill_swatches.append(fsw)
        lab("Opacity %", 1, 2)
        ttk.Spinbox(f, from_=10, to=100, increment=10, width=4, textvariable=self.shape_opacity).grid(row=1, column=3, pady=1)
        lab("Size", 0, 4)
        ttk.Spinbox(f, from_=6, to=96, width=4, textvariable=self.font_size).grid(row=0, column=5, pady=1)

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
        holder = tk.Frame(self, bg=RIBBON_BG, height=S(76) if RIBBON_FLAT else S(96))
        holder.pack(side="top", fill="x")
        holder.pack_propagate(False)
        tk.Frame(self, bg=RIBBON_LINE, height=1).pack(side="top", fill="x")
        self.ribbons = {}
        for name in TAB_NAMES:
            self.ribbons[name] = tk.Frame(holder, bg=RIBBON_BG)

        self._rb_built = set()
        self._rb_order = [n for n in self.ribbons]
        self._ensure_ribbon("Home")                 # the other tabs are built right after the window is on screen
        self.after(30, self._build_ribbons_idle)

    def _ensure_ribbon(self, name):
        if name in self._rb_built:
            return
        self._rb_built.add(name)
        getattr(self, "_rb_" + name)()
        try:
            t = self.tool.get()
            for b in self.tool_buttons.get(t, []):
                b.set_active(True)
            for mb in self.__dict__.get("menu_buttons", ()):
                mb.set_active(t in mb._tools)
        except Exception:
            pass

    def _ensure_all_ribbons(self):
        for n in self._rb_order:
            self._ensure_ribbon(n)

    def _build_ribbons_idle(self):
        """build one more tab per idle step so that the first window appears sooner"""
        try:
            for n in self._rb_order:
                if n not in self._rb_built:
                    self._ensure_ribbon(n)
                    self.after(8, self._build_ribbons_idle)
                    return
            self._ribbons_done()
        except tk.TclError:
            pass

    def _ribbons_done(self):
        pass

    # ---------------------------------------------------------------- the tabs (kept short: rarely used buttons live in drop-down menus)
    def _rb_Home(self):
        r = self.ribbons["Home"]
        g = self.group(r, "File")
        self.rb(g, "📂", "Open", self.open_dialog)
        self.rb(g, "💾", "Save", self.save)
        self.rb(g, "🖨", "Print", self.do_print)
        g = self.group(r, "Edit")
        self.rb(g, "↶", "Undo", self.undo, size=19)
        self.rb(g, "↷", "Redo", self.redo, size=19)
        g = self.group(r, "Tools")
        self.rb(g, "✋", "Hand", None, tool="view")
        self.rb(g, "🔍", "Find", self.show_find)
        self.rb(g, "✏", "Edit text", None, tool="edittext", size=19, width=62)
        self.rb(g, "🌐", "Look up", self.dlg_lookup, size=18, width=56)
        g = self.group(r, "View")
        self.rb(g, "−", "Zoom out", lambda: self.set_zoom(self.zoom / 1.2), size=20)
        self.rb(g, "+", "Zoom in", lambda: self.set_zoom(self.zoom * 1.2), size=20)
        self.rbm(g, "⛶", "Fit", [("Fit page", None, self.fit_page), ("Fit width", None, self.fit_width)], size=19)
        self.rbm(g, "⟳", "Rotate", [("Rotate left", None, lambda: self.rotate_selected(-90)),
                                    ("Rotate right", None, lambda: self.rotate_selected(90))], size=19)
        g = self.group(r, "Image")
        self.rb(g, "🖼", "Add Image", self.add_images)

    def _rb_Comment(self):
        r = self.ribbons["Comment"]
        g = self.group(r, "Text markup")
        self.rb(g, "🖍", "Highlight", None, tool="highlight")
        self.rb(g, "U̲", "Underline", None, tool="underline", size=19, width=62)
        self.rb(g, "S̶", "Strikeout", None, tool="strike", size=19, width=62)
        self.rb(g, "≈", "Squiggly", None, tool="squiggly", size=20, width=62)
        g = self.group(r, "Notes")
        self.rb(g, "🗒", "Sticky note", None, tool="sticky", size=18, width=70)
        self.rb(g, "▤", "Text box", None, tool="textbox", size=19)
        self.rb(g, "T", "Text", None, tool="text", size=19)
        self.rb(g, "🔖", "Stamps", self.show_stamp_menu, size=18)
        g = self.group(r, "Draw")
        self.rb(g, "✎", "Pen", None, tool="pen", size=19)
        self.rbm(g, "▭", "Shapes", [
            ("Line", "line", None), ("Arrow", "arrow", None), ("Double arrow", "darrow", None), None,
            ("Rectangle", "rect", None), ("Rounded", "rrect", None), ("Ellipse", "ellipse", None), ("Circle", "circle", None),
            ("Triangle", "triangle", None), ("Diamond", "diamond", None), ("Star", "star", None), None,
            ("Check", "check", None), ("Cross", "cross", None), ("Plus", "plus", None)], size=20)
        g = self.group(r, "Comments")
        self.rb(g, "💬", "Show list", lambda: self.set_sidebar("comments"), size=18, width=62)
        self.rbm(g, "⤓", "More", [("Export comments", None, self.export_comments), ("Import comments", None, self.import_comments),
                                  ("Comment summary", None, self.export_comment_summary),
                                  ("Export highlights", None, self.dlg_export_highlights)], size=19)
        g = self.group(r, "Highlight color")
        hf = tk.Frame(g, bg=RIBBON_BG)
        hf.pack(padx=6, pady=14)
        for col in HL_COLORS:
            hb = tk.Button(hf, width=2, relief="solid", bd=1, bg=self._hex(col), activebackground=self._hex(col),
                           cursor="hand2", command=lambda c=col: setattr(self, "hl_color", c))
            hb.pack(side="left", padx=2)
            self.hl_buttons.append(hb)
        self.combined_style_group(r)

    def _rb_Edit(self):
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
        self.rbm(g, "⌫", "Redact more", [("Redact text", None, self.dlg_redact_text), ("Redact data", None, self.dlg_redact_patterns)], size=19, width=72)
        g = self.group(r, "Copy")
        self.rb(g, "⎘", "Extract text", None, tool="extract", size=19, width=70)
        self.style_group(r)

    def _rb_Pictures(self):
        r = self.ribbons["Pictures"]
        g = self.group(r, "Insert")
        self.rb(g, "🖼", "Add Image", self.add_images)
        self.rb(g, "📋", "Paste", self.paste_image)
        self.rb(g, "📑", "Photos→Pages", self.photos_to_pages, width=86)
        self.rb(g, "🖼", "Edit pictures", None, tool="editimage", size=18, width=74)
        g = self.group(r, "Arrange")
        self.rb(g, "⊞", "Align", self.img_align_menu)
        self.rb(g, "❐", "Duplicate", self.img_duplicate)
        self.rbm(g, "⬆", "Order", [("Fit page", None, self.img_fit_page), ("To front", None, lambda: self.img_order(True)),
                                   ("To back", None, lambda: self.img_order(False))], size=17)
        self.rb(g, "🗑", "Delete", self.img_delete)
        g = self.group(r, "Transform")
        self.rbm(g, "⟲", "Rotate / flip", [("Left 90°", None, lambda: self.img_rotate(-90)), ("Right 90°", None, lambda: self.img_rotate(90)),
                                           ("Flip H", None, lambda: self.img_flip(True)), ("Flip V", None, lambda: self.img_flip(False))],
                 size=19, width=72)
        self.rb(g, "✂", "Crop", self.start_crop, size=19)
        af = tk.Frame(g, bg=RIBBON_BG)
        af.pack(side="left", padx=8, pady=10)
        tk.Label(af, text=tr("Angle °"), font=(UI_FONT, 8), bg=RIBBON_BG, fg=TH.LABEL).pack(anchor="w")
        asp = ttk.Spinbox(af, from_=-180, to=180, width=5, textvariable=self.angle_var, command=self._angle_changed)
        asp.pack()
        asp.bind("<Return>", self._angle_changed)
        ttk.Checkbutton(af, text=tr("Keep aspect"), variable=self.keep_aspect).pack(anchor="w", pady=(3, 0))
        g = self.group(r, "Effects")
        self.rbm(g, "☀", "Effects", [("Adjust…", None, self.adjust_dialog), ("No white", None, lambda: self.img_toggle("transp")),
                                     ("Grayscale", None, lambda: self.img_toggle("gray")), ("B & W", None, lambda: self.img_toggle("bw")),
                                     ("Border", None, lambda: self.img_toggle("border"))], size=19)
        self.rb(g, "↺", "Reset image", self.img_reset, size=19, width=70)
        g = self.group(r, "Finish")
        self.rb(g, "✔", "Apply", self.apply_images, size=19)
        qf = tk.Frame(g, bg=RIBBON_BG)
        qf.pack(side="left", padx=8, pady=8)
        tk.Label(qf, text=tr("Output quality"), font=(UI_FONT, 8), bg=RIBBON_BG, fg=TH.LABEL).pack(anchor="w")
        ttk.Combobox(qf, textvariable=self.quality_var, values=list(QUALITY), state="readonly", width=16).pack(pady=2)

    def _rb_Organize(self):
        r = self.ribbons["Organize"]
        g = self.group(r, "Pages")
        self.rbm(g, "⟳", "Rotate", [("Rotate left", None, lambda: self.rotate_selected(-90)),
                                    ("Rotate right", None, lambda: self.rotate_selected(90))], size=19)
        self.rb(g, "🗑", "Delete", self.delete_selected)
        self.rbm(g, "▲", "Move", [("Move up", None, lambda: self.move_page(-1)), ("Move down", None, lambda: self.move_page(1))], size=15)
        self.rb(g, "📄", "Extract", self.dlg_extract)
        self.rb(g, "✂", "Split", self.dlg_split, size=19)
        g = self.group(r, "Insert")
        self.rb(g, "＋", "Blank page", self.insert_blank, size=19, width=66)
        self.rbm(g, "⧉", "Insert", [("Insert PDF", None, self.dlg_insert_pdf), ("Merge PDFs", None, self.dlg_merge),
                                     ("Merge + pick", None, self.dlg_merge_pick), ("Photos→Pages", None, self.photos_to_pages)], size=19)
        g = self.group(r, "Crop pages")
        self.rbm(g, "⛶", "Crop", [("Drag area", "pagecrop", None), ("Margins…", None, self.dlg_crop),
                                  ("Crop / resize", None, self.dlg_page_crop)], size=18)
        g = self.group(r, "Export")
        self.rb(g, "🖼", "To PNG", self.export_png)
        self.rb(g, "↶", "Undo history", self.dlg_undo_history, size=18, width=78)

    def _rb_Forms(self):
        r = self.ribbons["Forms"]
        g = self.group(r, "Sign")
        self.rb(g, "✍", "Signature", lambda: self.show_sign_menu("sign"), size=19, width=68)
        self.rb(g, "Aa", "Initials", lambda: self.show_sign_menu("init"), size=17)
        self.rb(g, "🔖", "Stamps", self.show_stamp_menu, size=18)
        self.rb(g, "📅", "Date", None, tool="date", size=18)
        self.rb(g, "✍", "Signature block", self.dlg_signature_block, size=18, width=84)
        g = self.group(r, "Create form fields")
        self.rb(g, "☐", "Text field", None, tool="formtext", size=20, width=66)
        self.rb(g, "☑", "Check box", None, tool="formcheck", size=20, width=66)
        self.rb(g, "▾", "Drop-down", None, tool="formcombo", size=18, width=66)
        self.rb(g, "◉", "Radio button", None, tool="formradio", size=18, width=76)
        g = self.group(r, "Manage fields")
        self.rb(g, "☰", "Field editor", self.dlg_field_editor, size=18, width=72)
        self.rb(g, "⇥", "Fill from table", self.dlg_fill_from_data, size=18, width=88)
        self.rb(g, "▦", "Show fields", self.toggle_fields, size=18, width=68)
        g = self.group(r, "How to fill a form")
        tk.Label(g, text=tr("Choose the Hand tool and click a field\n(text, check box, drop-down list)."),
                 font=(UI_FONT, 8), bg=RIBBON_BG, fg=TH.SECOND, justify="left").pack(padx=8, pady=14)
        self.style_group(r)

    def _rb_Convert(self):
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

    def _rb_Document(self):
        r = self.ribbons["Document"]
        g = self.group(r, "Protect")
        self.rb(g, "🔒", "Protect", self.dlg_protect, size=18)
        self.rb(g, "🔓", "Unprotect", self.do_unprotect, size=18, width=70)
        self.rb(g, "✔", "Redaction check", self.dlg_redaction_check, size=18, width=92)
        self.rb(g, "🧹", "Clean data", self.dlg_clean_metadata, size=18, width=72)
        g = self.group(r, "Pages")
        self.rb(g, "💧", "Watermark", self.dlg_watermark, size=18, width=76)
        self.rb(g, "🎨", "Background", self.dlg_background, size=18, width=78)
        self.rbm(g, "⇱", "Header / numbers", [("Header/Footer", None, self.dlg_header_footer), ("Page numbers", None, self.dlg_page_numbers),
                                              ("Bates", None, self.dlg_bates), ("Page labels", None, self.dlg_page_labels)], size=18, width=96)
        g = self.group(r, "Optimize")
        self.rbm(g, "🗜", "Optimize", [("Compress", None, self.dlg_compress), ("Smart compress", None, self.dlg_smart_compress),
                                       ("Flatten", None, self.do_flatten), ("OCR", None, self.dlg_ocr)], size=18)
        self.rb(g, "🗄", "PDF/A", self.dlg_pdfa, size=18, width=56)
        g = self.group(r, "Check")
        self.rb(g, "♿", "Accessibility", self.dlg_accessibility, size=18, width=84)
        self.rb(g, "⇄", "Compare", self.dlg_compare, size=18, width=62)
        self.rb(g, "📎", "Attachments", self.dlg_attachments, size=18, width=78)
        self.rb(g, "ℹ", "Properties", self.dlg_properties, size=18, width=66)
        g = self.group(r, "Files")
        self.rbm(g, "▤", "Folder", [("Batch folder", None, self.dlg_batch), ("Search folder", None, self.dlg_folder_search)], size=18)
        g = self.group(r, "Program")
        self.rb(g, "✔", "Guided actions", self.dlg_guided, size=18, width=78)
        self.rb(g, "⚙", "Settings", self.dlg_settings, size=18, width=60)

    def _rb_View(self):
        r = self.ribbons["View"]
        g = self.group(r, "Navigate")
        self.rb(g, "◀", "Previous", lambda: self.goto(self.page_no - 1), size=15)
        self.rb(g, "▶", "Next", lambda: self.goto(self.page_no + 1), size=15)
        g = self.group(r, "Page view")
        self.rb(g, "☰", "Continuous", self.toggle_continuous, size=19, width=70)
        self.rb(g, "☾", "Night view", self.toggle_night, size=19, width=70)
        self.rb(g, "⛶", "Full screen", lambda: self.toggle_fullscreen(False), size=19, width=70)
        self.rb(g, "▶", "Presentation", lambda: self.toggle_fullscreen(True), size=17, width=78)
        g = self.group(r, "Side panel")
        self.rb(g, "◧", "Hide / show", self.toggle_sidebar, size=18, width=70)
        self.rb(g, "▤", "Pages", lambda: self.set_sidebar("pages"), size=18)
        self.rb(g, "🔖", "Bookmarks", lambda: self.set_sidebar("bookmarks"), size=18, width=66)
        self.rb(g, "💬", "Comments", lambda: self.set_sidebar("comments"), size=18, width=66)
        g = self.group(r, "Listen")
        self.rb(g, "🔊", "Read aloud", self.dlg_read_aloud, size=18, width=70)
        self.rb(g, "⏹", "Stop", self.stop_speaking, size=17, width=48)
        g = self.group(r, "Toolbar")
        self.rb(g, "★", "Quick toolbar", self.qt_toggle, size=19, width=70)

    def _rb_CAD(self):
        self.build_cad_ribbon(self.ribbons["CAD"])

    # --------------------------------------------------------------- body

    def _rb_CAD(self):
        self.build_cad_ribbon(self.ribbons["CAD"])

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
        self.zoom_var = tk.DoubleVar(value=self._zoom_pos(self.zoom))
        TH.IOSSlider(zf, S, from_=0, to=100, variable=self.zoom_var, length=S(150),
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
        self.bind("<Control-k>", lambda e: (self.pal_open(), "break")[1])
        self.bind("<Control-s>", lambda e: self.save())
        self.bind("<Control-S>", lambda e: self.save_as())
        self.bind("<Control-z>", lambda e: self.undo())
        self.bind("<Control-f>", lambda e: self.show_find())
        self.bind("<Control-b>", lambda e: self.toggle_sidebar())
        self.bind("<Control-y>", lambda e: self.redo())
        self.bind("<Control-Z>", lambda e: self.redo())
        self.bind("<Control-p>", lambda e: self.do_print())
        self.bind("<Control-Q>", lambda e: self.qt_toggle())
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
        self.bind("<Escape>", lambda e: (self.exit_fullscreen() if self._fs else None, self.select(None), self.cad_cancel() if hasattr(self, "_cad") else None, self.hide_find() if self.find_frame is not None and self.find_frame.winfo_ismapped() else None))
        for key, dx, dy in (("Left", -1, 0), ("Right", 1, 0), ("Up", 0, -1), ("Down", 0, 1)):
            self.bind(f"<{key}>", lambda e, a=dx, b=dy: self._key_arrow(a, b, e))
        self.bind("<Prior>", lambda e: self._key_nav(-1))
        self.bind("<Next>", lambda e: self._key_nav(1))

    def _typing(self):
        return isinstance(self.focus_get(), (tk.Entry, ttk.Entry, ttk.Spinbox, tk.Text, ttk.Combobox))

    def _key_delete(self, _e):
        if self._typing():
            return
        if self.tf:                                   # a text frame is open: Delete removes it
            self.tf_cancel()
            return False
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
             "text": "Text: click or drag on the page to make a text frame, then type. Drag the frame to move it.",
             "highlight": "Highlight: drag over the text.",
             "rect": "Rectangle: drag to draw. Hold Shift for a square. Turn on Fill in the Comment tab for solid shapes.",
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
        for mb in self.__dict__.get("menu_buttons", ()):
            mb.set_active(t in mb._tools)
        self.canvas.config(cursor=self._tool_cursor(t))
        if hasattr(self, "_cad") and t != self._cad.get("tool"):
            self.cad_cancel()
        self.status.config(text=self.HINTS.get(t) or EXTRA_HINTS.get(t) or CAD_HINTS.get(t, ""))

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
                pw = self.unlock_document(doc, os.path.basename(p))
                if pw is None:
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

    def single_listen(self, receiver):
        """files sent by other copies of the program (Explorer starts one copy per file) open here as tabs"""
        self._single = receiver

        def poll():
            try:
                got = receiver.take()
                if got:
                    try:
                        self.deiconify()
                        self.lift()
                        self.focus_force()
                    except tk.TclError:
                        pass
                    for p in got:
                        self.open_path(p)
            except Exception:
                pass
            try:
                self.after(250, poll)
            except tk.TclError:
                pass
        self.after(250, poll)

    def unlock_document(self, doc, name, parent=None, tries=3):
        """Ask for the password until `doc` opens. Returns the password, or None when cancelled / not unlocked."""
        msg = "This PDF is password-protected.\nPassword:"
        for k in range(tries):
            pw = self.ask_password(msg, name, parent)
            if pw is None:
                return None
            try:
                if doc.authenticate(pw):
                    return pw
            except Exception:
                pass
            msg = "That password is not correct. Try again.\nPassword:"
        messagebox.showerror(APP_NAME, "Wrong password.", parent=parent or self)
        return None

    def ask_password(self, prompt, name="", parent=None):
        """Small password box that always comes to the front. Returns the text, or None when cancelled."""
        owner = parent or self
        d = tk.Toplevel(owner)
        d.title(APP_NAME)
        d.configure(bg=TH.BG, padx=18, pady=14)
        d.resizable(False, False)
        try:
            if owner.winfo_viewable():
                d.transient(owner)
        except tk.TclError:
            pass
        out = {}
        if name:
            ttk.Label(d, text=name, font=(UI_FONT, 9, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(d, text=prompt, justify="left").grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 6))
        var = tk.StringVar()
        ent = ttk.Entry(d, textvariable=var, show="*", width=34)
        ent.grid(row=2, column=0, columnspan=2, sticky="we")
        shown = tk.BooleanVar(value=False)
        ttk.Checkbutton(d, text="Show password", variable=shown,
                        command=lambda: ent.config(show="" if shown.get() else "*")).grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))

        def ok(_e=None):
            out["v"] = var.get()
            d.destroy()
        bf = ttk.Frame(d)
        bf.grid(row=4, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(bf, text="OK", command=ok).pack(side="left", padx=4)
        ttk.Button(bf, text="Cancel", command=d.destroy).pack(side="left")
        d.bind("<Return>", ok)
        d.bind("<Escape>", lambda e: d.destroy())
        d.update_idletasks()
        w, h = d.winfo_reqwidth(), d.winfo_reqheight()
        d.geometry("+%d+%d" % ((d.winfo_screenwidth() - w) // 2, (d.winfo_screenheight() - h) // 3))
        d.attributes("-topmost", True)                      # a hidden dialog looks like a frozen program
        d.lift()
        d.focus_force()
        ent.focus_set()
        try:
            d.grab_set()
        except tk.TclError:
            pass
        d.wait_window()
        return out.get("v") or None

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
            try:
                if getattr(self, "_single", None):
                    self._single.close()
            except Exception:
                pass
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
        self._edit_serial = getattr(self, "_edit_serial", 0) + 1
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
        old = self.zoom
        keep = None
        if render and mode is None and self.doc is not None and not self.cont and abs(z - old) > 1e-9:
            keep = self._view_center()                  # zoom around what is in the middle of the window
        self.zoom = max(ZOOM_MIN, min(ZOOM_MAX, z))
        if self.cont and self.doc is not None and self._cont_too_big():
            self._cont_auto = True                      # very large zoom: one page at a time (only the visible part is drawn)
            self.cont = False
            self.status.config(text=tr("One page at a time for this zoom."))
        elif getattr(self, "_cont_auto", False) and not self.cont and self.doc is not None and not self._cont_too_big():
            self._cont_auto = False
            self.cont = True
        self._sync_zoom = True
        self.zoom_var.set(self._zoom_pos(self.zoom))
        self._sync_zoom = False
        self.zoom_lbl.config(text=f"{int(round(self.zoom * 100))}%")
        if render:
            self.render()
            if keep is not None:
                self._view_center_set(keep)

    @staticmethod
    def _zoom_pos(z):
        """slider position 0..100 for a zoom factor (logarithmic: the small zooms stay easy to set)"""
        return 100.0 * math.log(max(ZOOM_MIN, min(ZOOM_MAX, z)) / ZOOM_MIN) / math.log(ZOOM_MAX / ZOOM_MIN)

    @staticmethod
    def _zoom_from_pos(p):
        return ZOOM_MIN * (ZOOM_MAX / ZOOM_MIN) ** (max(0.0, min(100.0, p)) / 100.0)

    def _cont_too_big(self):
        """True when the pages are too large to be drawn whole at the current zoom (continuous view cannot show that)"""
        try:
            r = self.page().rect
            return r.width * r.height * self.zoom * self.zoom > TILE_PIXELS
        except Exception:
            return False

    def _view_center(self):
        """the page point (in view points) that is in the middle of the window"""
        cv = self.canvas
        cx = cv.canvasx(cv.winfo_width() / 2)
        cy = cv.canvasy(cv.winfo_height() / 2)
        return ((cx - self.off_x) / max(self.zoom, 1e-6), (cy - self.off_y) / max(self.zoom, 1e-6))

    def _view_center_set(self, pt):
        cv = self.canvas
        try:
            x0, y0, x1, y1 = [float(v) for v in str(cv.cget("scrollregion")).split()]
        except ValueError:
            return
        W, H = max(1.0, x1 - x0), max(1.0, y1 - y0)
        left = self.off_x + pt[0] * self.zoom - cv.winfo_width() / 2
        top = self.off_y + pt[1] * self.zoom - cv.winfo_height() / 2
        cv.xview_moveto(max(0.0, left / W))
        cv.yview_moveto(max(0.0, top / H))
        self.update_idletasks()
        self._tile_update()

    def _on_scale(self, v):
        if self._sync_zoom or not self.doc:
            return
        if self._zoom_after:
            self.after_cancel(self._zoom_after)
        self._zoom_after = self.after(60, lambda: self.set_zoom(self._zoom_from_pos(float(v))))

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
        self._tiled = False
        if self.doc is not None and self.cont:
            return self._render_cont()
        self._gen += 1
        self.canvas.delete("all")
        if self.doc is None:
            self._draw_start_screen()
            return
        if self._cont_too_big():
            return self._render_tiled()
        z = safe_zoom(self.page(), self.zoom)
        if z < self.zoom - 1e-6:                 # a very large page: never build a gigantic bitmap
            self.zoom = max(0.02, z)
            self._sync_zoom = True
            self.zoom_var.set(self._zoom_pos(self.zoom))
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

    # ---- very large zoom: only the part of the page that is on screen is drawn
    def _render_tiled(self):
        r = self.page().rect
        self.page_w, self.page_h = int(math.ceil(r.width * self.zoom)), int(math.ceil(r.height * self.zoom))
        self.off_x, self.off_y = self._offsets()
        self._tiled = True
        self._tile_box = None
        self._tile_photo = None
        for k, col in zip((7, 5, 3, 1), TH.PAGE_SHADOW):
            self.canvas.create_rectangle(self.off_x - k + 2, self.off_y - k + 4,
                                         self.off_x + self.page_w + k - 2, self.off_y + self.page_h + k + 1,
                                         fill=col, outline="", tags="pg")
        self.canvas.create_rectangle(self.off_x, self.off_y, self.off_x + self.page_w, self.off_y + self.page_h,
                                     fill="#1a1a1a" if self.night else "white", outline="", tags="pg")
        self._tile_item = self.canvas.create_image(self.off_x, self.off_y, anchor="nw", tags="pg")
        self._update_scrollregion()
        self._after_render()
        self.draw_floats()
        self._draw_textsel()
        self._tile_update(force=True)
        self._tile_start()

    def _tile_start(self):
        if getattr(self, "_tile_job", None):
            try:
                self.after_cancel(self._tile_job)
            except (tk.TclError, ValueError):
                pass
        self._tile_job = self.after(120, self._tile_tick)

    def _tile_tick(self):
        self._tile_job = None
        if not (self.doc is not None and getattr(self, "_tiled", False) and not self.cont):
            return
        try:
            self._tile_update()
        except Exception:
            pass
        self._tile_job = self.after(120, self._tile_tick)

    def _tile_update(self, force=False):
        if not getattr(self, "_tiled", False) or self.doc is None or self.cont:
            return
        cv = self.canvas
        cw, ch = max(cv.winfo_width(), 1), max(cv.winfo_height(), 1)
        vx0, vy0 = cv.canvasx(0), cv.canvasy(0)
        px0, py0 = self.off_x, self.off_y
        px1, py1 = px0 + self.page_w, py0 + self.page_h
        need = (max(vx0, px0), max(vy0, py0), min(vx0 + cw, px1), min(vy0 + ch, py1))
        if need[2] <= need[0] or need[3] <= need[1]:
            return
        box = self._tile_box
        if not force and box and box[0] <= need[0] and box[1] <= need[1] and box[2] >= need[2] and box[3] >= need[3]:
            return
        mx, my = cw * 0.5, ch * 0.5                         # draw a little more than what is visible, so scrolling does not redraw at once
        while (cw + 2 * mx) * (ch + 2 * my) > 14e6 and mx > 8:
            mx, my = mx * 0.7, my * 0.7
        reg = (max(vx0 - mx, px0), max(vy0 - my, py0), min(vx0 + cw + mx, px1), min(vy0 + ch + my, py1))
        z = self.zoom
        clip = fitz.Rect((reg[0] - px0) / z, (reg[1] - py0) / z, (reg[2] - px0) / z, (reg[3] - py0) / z)
        pix = self.page().get_pixmap(matrix=fitz.Matrix(z, z), clip=clip, alpha=False)
        if self.night:
            pix.invert_irect(pix.irect)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        self._tile_photo = ImageTk.PhotoImage(img)
        self.canvas.itemconfig(self._tile_item, image=self._tile_photo)
        self.canvas.coords(self._tile_item, px0 + pix.x, py0 + pix.y)
        self._tile_box = (px0 + pix.x, py0 + pix.y, px0 + pix.x + pix.width, py0 + pix.y + pix.height)

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
        ttk.Button(dlg, text="Check for updates...", command=lambda: (dlg.destroy(), self.upd_check(True))).pack(pady=(14, 0))
        ttk.Button(dlg, text="Close", command=dlg.destroy).pack(pady=(8, 0))

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
        self.canvas.delete("shp")
        lw = self.line_width.get()
        a, b = constrain_shape(kind, self._c2v(*a_canvas), self._c2v(*b_canvas), shift)
        items = shape_items(kind, a, b, lw, False)
        self._draw_shape_items(items, self._hex(self.color), self._hex(self.fill_color) if self.fill_on.get() else "", lw, ("shp",))

    def _draw_shape_items(self, items, col, fillc, lw, tags):
        """draw shape geometry (view space) on the canvas"""
        cv = self.canvas
        w = max(1, lw * self.zoom)

        def flat(pts):
            out = []
            for p in pts:
                out.extend(self._v2c(*p))
            return out

        for typ, data, _ in items:
            if typ == "line":
                cv.create_line(*flat(data), fill=col, width=w, capstyle="round", tags=tags)
            elif typ == "open":
                cv.create_line(*flat(data), fill=col, width=w, capstyle="round", joinstyle="round", tags=tags)
            elif typ == "poly":
                cv.create_polygon(*flat(data), outline=col, fill=fillc, width=w, tags=tags)
            elif typ == "head":
                cv.create_polygon(*flat(data), outline=col, fill=col, tags=tags)
            else:
                x0, y0 = self._v2c(data[0], data[1])
                x1, y1 = self._v2c(data[2], data[3])
                if typ == "oval":
                    cv.create_oval(x0, y0, x1, y1, outline=col, fill=fillc, width=w, tags=tags)
                else:
                    cv.create_rectangle(x0, y0, x1, y1, outline=col, fill=fillc, width=w, tags=tags)

    def commit_shape(self, kind, a_canvas, b_canvas, shift):
        """the dragged shape becomes a floating object: it can be moved, resized and rotated until it is applied"""
        lw = self.line_width.get()
        a, b = constrain_shape(kind, self._c2v(*a_canvas), self._c2v(*b_canvas), shift)
        if not shape_items(kind, a, b, lw, False):
            return
        if self.floats and self.float_page != self.page_no:
            self.commit_floats()
        self.float_page = self.page_no
        o = FloatShape(kind, a, b, lw, self.color, self.fill_color if self.fill_on.get() else None, self.shape_opacity.get())
        self.floats.append(o)
        self.modified = True
        self.select(o)
        self.update_ui()
        self.status.config(text=tr("Shape added: drag it to move, corners to resize, blue dot to rotate. Enter = apply to the page."))

    def _flatten_shape(self, pg, o):
        self._write_shape(pg, o.world_items(), o.lw, o.color, o.fill, max(0.05, min(1.0, o.opacity / 100)))

    def _write_shape(self, pg, items, lw, color, fill, op):
        m = pg.derotation_matrix

        def P(p):
            return fitz.Point(p[0], p[1]) * m

        def R(box):
            r = fitz.Rect(P((box[0], box[1])), P((box[2], box[3])))
            r.normalize()
            return r

        sh = pg.new_shape()
        base = dict(width=lw, color=color, lineCap=1, lineJoin=1, stroke_opacity=op, fill_opacity=op)
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
                sh.finish(fill=color, closePath=True, **base)
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
        self.tf_draw()
        if hasattr(self, "_cad"):
            self.cad_scale_refresh()
            self.cad_draw()
        self.canvas.delete("fl")
        self._handles = {}
        if not self.doc or not self.floats:
            return
        tags = ("pg", "fl")
        for o in self.floats:
            if isinstance(o, FloatShape):
                self._draw_shape_items(o.world_items(), self._hex(o.color), self._hex(o.fill) if o.fill else "", o.lw, tags)
                continue
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
        self.tf_commit()
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
        if pno == self.page_no or self.cont:        # continuous view: the page may be off screen, its picture must be redrawn too
            self.render()
        self.refresh_thumb(pno)
        self.update_ui()
        if errors:
            messagebox.showerror(APP_NAME, "Could not place:\n{}".format("\n".join(errors)))
        else:
            self.status.config(text="Image applied to the page.")

    def _flatten(self, pg, o, dpi, q):
        if isinstance(o, FloatShape):
            return self._flatten_shape(pg, o)
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
        if self.need_sel() and isinstance(self.sel, FloatShape):
            self.status.config(text=tr("Flip works on pictures. Use the rotate dot or the angle box for shapes."))
            return
        if self.need_sel():
            if horizontal:
                self.sel.flip_h = not self.sel.flip_h
            else:
                self.sel.flip_v = not self.sel.flip_v
            self._touch()

    def img_toggle(self, attr):
        if self.need_sel() and isinstance(self.sel, FloatShape):
            return
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
        if self.need_sel() and isinstance(self.sel, FloatShape):
            self.sel.angle = 0.0
            self.angle_var.set(0)
            self._touch()
            return
        if self.need_sel():
            o = self.sel
            aspect = o.pristine.height / o.pristine.width
            o.set_base(o.pristine)
            o.reset_adjust()
            o.h = o.w * aspect
            self.angle_var.set(0)
            self._touch()

    def start_crop(self):
        if self.need_sel() and isinstance(self.sel, FloatShape):
            self.status.config(text=tr("Crop works on pictures."))
            return
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
    def _float_press(self, e, only_sel=False):
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
        if only_sel:                                      # while a shape tool is active only the selected shape reacts
            if o in self.floats and self.f_hit(o, vx, vy):
                self.fdrag = {"mode": "move", "o": o, "last": (vx, vy)}
                return True
            return False
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
        if self.tool.get() in CAD_TOOLS:
            if self.cad_pending():
                self.cad_finish()
            return
        if not self.floats:
            self._cont_focus_at(e)
            if not self.annot_menu(e):
                self.text_line_menu(e)
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
        if not self.annot_menu(e):
            self.text_line_menu(e)

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
        if hasattr(self, "_cad") and self.cad_pending() and not self._typing():
            self.cad_finish()
            return
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
        if self.tf_press(e):
            return
        t = self.tool.get()
        if t in CAD_TOOLS:
            self._drag_start = None
            if not self.cad_pending():
                self._cont_focus_at(e)
            self.cad_press(e)
            return
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
        if t in SHAPE_TOOLS and isinstance(self.sel, FloatShape) and self.sel in self.floats:
            if self.cont:
                self._cont_focus(self.float_page)
            if self._float_press(e, only_sel=True):       # grab the handles / body of the shape that was just drawn
                return
        self._cont_focus_at(e)
        self._drag_start = (self.canvas.canvasx(e.x), self.canvas.canvasy(e.y))
        if t in SHAPE_TOOLS:
            self._drag_start = self.cad_snap_canvas(*self._drag_start, shift=False)
        if t == "edittext":
            self.tf_move_begin(e)
        if t == "text":
            self._add_text(e)
            self._drag_start = None
        elif t == "pen":
            self._stroke = [self._drag_start]

    def _on_drag(self, e):
        if self.doc is None:
            return
        if self._tfd:
            self.tf_drag(e)
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
        elif t == "edittext":
            self.tf_move_preview(e)
        elif t in SHAPE_TOOLS:
            x, y = self.cad_snap_canvas(x, y, bool(e.state & 1), self._drag_start)
            self.draw_shape_preview(t, self._drag_start, (x, y), bool(e.state & 1))
        elif t in ("highlight", "whiteout", "image", "crop", "pagecrop") + tuple(RECT_TOOLS):
            if self._drag_item:
                self.canvas.delete(self._drag_item)
            self._drag_item = self.canvas.create_rectangle(
                *self._drag_start, x, y, outline=ACCENT, dash=(4, 3), width=2)

    def _on_release(self, e):
        if self.doc is None:
            return
        if self._tfd:
            self.tf_release(e)
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
        if t in SHAPE_TOOLS:
            x1, y1 = self.cad_snap_canvas(x1, y1, bool(e.state & 1), (x0, y0))
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
        self.tt_add_click(e)

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
    receiver = None
    if not os.environ.get("CLAUDIUS_NO_SINGLE"):
        try:
            from claudius_v3 import load_cfg
            if load_cfg().get("single_instance", True):
                import claudius_single
                receiver, handed = claudius_single.start(paths)
                if handed:                                  # the running window took the files as new tabs
                    return
        except Exception:
            receiver = None
    app = PdfEditor(paths)
    _drop_pre_splash()
    if os.environ.get("CLAUDIUS_QUIT_MS"):                 # used by the automatic tests: close the window after this many ms
        app.after(int(os.environ["CLAUDIUS_QUIT_MS"]), app.destroy)
    if receiver is not None:
        app.single_listen(receiver)
    app.mainloop()


if __name__ == "__main__":
    main()
