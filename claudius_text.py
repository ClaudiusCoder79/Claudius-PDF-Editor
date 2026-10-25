# -*- coding: utf-8 -*-
"""Claudius PDF Editor - adding and editing text.

One text window for all three jobs (click to add text, draw a text box, edit existing text) with:
  * a format bar: font, size, bold, italic, colour, alignment, line spacing
  * a live preview that shows the real page with the new text on it (drawn by the same engine that saves it)
  * Remembers the last format that was used.
"""
import re
import tkinter as tk
from tkinter import ttk, colorchooser, messagebox
from tkinter import font as tkfont

try:
    import pymupdf as fitz
except ImportError:
    import fitz
from PIL import Image, ImageDraw, ImageTk

import claudius_theme as TH
from claudius_i18n import tr as _tr
import claudius_extra as EX
from claudius_extra import APP, UI_FONT, BG, S
from claudius_view import get_setting, set_setting

# ------------------------------------------------------------------------------------------------ fonts
BASE14 = {"he": ("Helvetica / Arial (built in)", ("helv", "hebo", "heit", "hebi"), "Arial"),
          "ti": ("Times (built in)", ("tiro", "tibo", "tiit", "tibi"), "Times New Roman"),
          "co": ("Courier (built in)", ("cour", "cobo", "coit", "cobi"), "Courier New")}

# label, (regular, bold, italic, bold italic) file names - only the ones installed on this computer are offered
SYSTEM_FONTS = [
    ("Arial", ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf")),
    ("Times New Roman", ("times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf")),
    ("Calibri", ("calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf")),
    ("Cambria", ("cambria.ttc",) * 4),
    ("Verdana", ("verdana.ttf", "verdanab.ttf", "verdanai.ttf", "verdanaz.ttf")),
    ("Tahoma", ("tahoma.ttf", "tahomabd.ttf", "tahoma.ttf", "tahomabd.ttf")),
    ("Segoe UI", ("segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "segoeuiz.ttf")),
    ("Georgia", ("georgia.ttf", "georgiab.ttf", "georgiai.ttf", "georgiaz.ttf")),
    ("Trebuchet MS", ("trebuc.ttf", "trebucbd.ttf", "trebucit.ttf", "trebucbi.ttf")),
    ("Palatino Linotype", ("pala.ttf", "palab.ttf", "palai.ttf", "palabi.ttf")),
    ("Comic Sans MS", ("comic.ttf", "comicbd.ttf", "comici.ttf", "comicz.ttf")),
    ("Consolas", ("consola.ttf", "consolab.ttf", "consolai.ttf", "consolaz.ttf")),
    ("Courier New", ("cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf")),
    ("DejaVu Sans", ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans-Oblique.ttf", "DejaVuSans-BoldOblique.ttf")),
    ("DejaVu Serif", ("DejaVuSerif.ttf", "DejaVuSerif-Bold.ttf", "DejaVuSerif-Italic.ttf", "DejaVuSerif-BoldItalic.ttf")),
    ("Liberation Sans", ("LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf", "LiberationSans-Italic.ttf", "LiberationSans-BoldItalic.ttf")),
    ("Liberation Serif", ("LiberationSerif-Regular.ttf", "LiberationSerif-Bold.ttf", "LiberationSerif-Italic.ttf", "LiberationSerif-BoldItalic.ttf")),
    ("Carlito", ("Carlito-Regular.ttf", "Carlito-Bold.ttf", "Carlito-Italic.ttf", "Carlito-BoldItalic.ttf")),
]
ORIGINAL = "original"
_CATALOG = None


def font_catalog():
    """[(label, key)] of the fonts that can be used on this computer. Key: 'base:he' / 'file:Calibri'."""
    global _CATALOG
    if _CATALOG is None:
        out = [(v[0], "base:" + k) for k, v in BASE14.items()]
        for label, files in SYSTEM_FONTS:
            if EX._find_in_dirs([files[0]]):
                out.append((label, "file:" + label))
        _CATALOG = out
    return _CATALOG


def _files_of(label):
    for lab, files in SYSTEM_FONTS:
        if lab == label:
            return files
    return None


def _unicode_fk(text, bold, ital):
    """A font with real letters for any language (Polish, Czech, Greek, ...) in the style asked for."""
    for lab in ("Arial", "Segoe UI", "Liberation Sans", "DejaVu Sans", "Carlito"):
        files = _files_of(lab)
        path = EX._find_in_dirs([files[int(bold) + 2 * int(ital)]]) or EX._find_in_dirs([files[0]])
        if path:
            try:
                f = fitz.Font(fontfile=path)
                if all(f.has_glyph(ord(c)) for c in text if c not in "\n\r\t "):
                    return {"fontname": "TF%s%d%d" % (re.sub(r"[^A-Za-z0-9]", "", lab), int(bold), int(ital)),
                            "fontfile": path}
            except Exception:
                pass
    return EX.font_kwargs(text, bold)


def resolve_font(spec, text, original=None):
    """insert_text / insert_textbox keyword arguments for the chosen font."""
    bold, ital = bool(spec.get("bold")), bool(spec.get("italic"))
    key = spec.get("font", "base:he")
    if key == ORIGINAL and original:
        import claudius_v3
        doc, pg, span = original
        return claudius_v3.match_font(doc, pg, span, text, bold, ital)
    kind, _, data = key.partition(":")
    if kind == "file":
        files = _files_of(data)
        if files:
            path = EX._find_in_dirs([files[int(bold) + 2 * int(ital)]]) or EX._find_in_dirs([files[0]])
            if path:
                try:
                    f = fitz.Font(fontfile=path)
                    if all(f.has_glyph(ord(c)) for c in text if c not in "\n\r\t "):
                        return {"fontname": "TF%s%d%d" % (re.sub(r"[^A-Za-z0-9]", "", data), int(bold), int(ital)),
                                "fontfile": path}
                except Exception:
                    pass
        return _unicode_fk(text, bold, ital)
    fam = data if data in BASE14 else "he"
    if EX.needs_unicode(text):
        return _unicode_fk(text, bold, ital)
    return {"fontname": BASE14[fam][1][int(bold) + 2 * int(ital)]}


def _prep(pg, fk):
    """fontbuffer fonts (taken from the PDF itself) must be added to the page first"""
    fk = dict(fk)
    if "fontbuffer" in fk:
        pg.insert_font(fontname=fk["fontname"], fontbuffer=fk.pop("fontbuffer"))
    return fk


_BUF_FONTS = {}


def _width(fk, text, size):
    try:
        if "fontbuffer" in fk:
            f = _BUF_FONTS.get(id(fk["fontbuffer"]))
            if f is None:
                if len(_BUF_FONTS) > 8:
                    _BUF_FONTS.clear()
                f = _BUF_FONTS[id(fk["fontbuffer"])] = fitz.Font(fontbuffer=fk["fontbuffer"])
            return f.text_length(text, fontsize=size)
        return EX.text_width(text, size, fk)
    except Exception:
        return len(text) * size * 0.5


DEFAULT_SPEC = {"font": "base:he", "size": 14.0, "bold": False, "italic": False, "color": (0, 0, 0),
                "align": "left", "spacing": 1.0, "fit": True, "fill": None, "border": None,
                "underline": False, "strike": False}
ALIGNS = (("left", "Left"), ("center", "Centre"), ("right", "Right"), ("justify", "Justify"))
QUICK_COLORS = ("#000000", "#c0392b", "#d35400", "#27ae60", "#2471a3", "#6c3483", "#7f8c8d", "#ffffff")


def _hex(rgb):
    return "#%02x%02x%02x" % tuple(int(round(max(0, min(1, c)) * 255)) for c in rgb)


def _rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


# ------------------------------------------------------------------------------------------------ writing
# All text is laid out by this small engine (not by PyMuPDF's insert_textbox) so that underline, strike-through,
# justified text and line spacing work the same everywhere - in the preview and in the saved file.
def _wrap(fk, text, size, maxw):
    """[(line, last_line_of_paragraph)] - words are wrapped to the width `maxw`"""
    out = []
    for para in text.split("\n"):
        if not para.strip():
            out.append(("", True))
            continue
        cur = ""
        for w in para.split(" "):
            trial = (cur + " " + w) if cur else w
            if cur and _width(fk, trial, size) > maxw:
                out.append((cur, False))
                cur = w
            else:
                cur = trial
        out.append((cur, True))
    return out


class _Pen:
    """Writes lines of text. Positions are given in 'layout space'; conv() turns them into PDF page coordinates."""

    def __init__(self, pg, spec, fk, conv, rotate):
        self.pg, self.spec, self.conv, self.rotate = pg, spec, conv, rotate
        self.fk = _prep(pg, fk)

    def width(self, text, size):
        return _width(self.fk, text, size)

    def line(self, x, y, text, size, spread=0.0):
        """one line, baseline starts at (x, y); `spread` = extra space for every blank (justified text)"""
        sp = self.spec
        if not text:
            return
        if spread:
            cx = x
            sw = self.width(" ", size) + spread
            for w in text.split(" "):
                self.pg.insert_text(self.conv(cx, y), w, fontsize=size, color=sp["color"], rotate=self.rotate, **self.fk)
                cx += self.width(w, size) + sw
            wline = cx - sw - x
        else:
            self.pg.insert_text(self.conv(x, y), text, fontsize=size, color=sp["color"], rotate=self.rotate, **self.fk)
            wline = self.width(text, size)
        th = max(0.4, size / 20)
        for flag, dy in (("underline", size * 0.14), ("strike", -size * 0.28)):
            if sp.get(flag):
                self.pg.draw_line(self.conv(x, y + dy), self.conv(x + wline, y + dy), color=sp["color"], width=th)


def _line_x(align, x0, x1, w):
    return x0 if align in ("left", "justify") else (x0 + x1) / 2 - w / 2 if align == "center" else x1 - w


def _paint(pen, lines, size, x0, x1, y_first, lh, align):
    for k, (txt, last) in enumerate(lines):
        w = pen.width(txt, size)
        spread = 0.0
        if align == "justify" and not last and txt.count(" ") and w < (x1 - x0):
            spread = ((x1 - x0) - w) / txt.count(" ")
        pen.line(_line_x(align, x0, x1, w), y_first + k * lh, txt, size, spread)


def _view_conv(pg):
    m = pg.derotation_matrix
    return lambda x, y: fitz.Point(x, y) * m


def write_lines(pg, spec, text, fk, pt):
    """Text that starts at the point `pt` (baseline of the first line); several lines go downwards."""
    pen = _Pen(pg, spec, fk, _view_conv(pg), pg.rotation)
    size = float(spec["size"])
    lh = size * 1.2 * float(spec.get("spacing", 1.0))
    vp = fitz.Point(pt) * pg.rotation_matrix                          # work in what the viewer sees
    for k, line in enumerate(text.split("\n")):
        w = pen.width(line, size)
        pen.line(vp.x - (w / 2 if spec["align"] == "center" else w if spec["align"] == "right" else 0),
                 vp.y + k * lh, line, size)


def write_box(pg, spec, text, fk, rect):
    """Text inside a rectangle; shrinks the text until it fits when 'fit' is on."""
    pen = _Pen(pg, spec, fk, _view_conv(pg), pg.rotation)
    vr = fitz.Rect(rect) * pg.rotation_matrix
    vr.normalize()
    if spec.get("fill"):
        pg.draw_rect(rect, color=None, fill=spec["fill"])
    if spec.get("border"):
        pg.draw_rect(rect, color=spec["border"], width=0.75)
    inner = vr + (3, 2, -3, -2) if vr.width > 10 and vr.height > 8 else vr
    s = float(spec["size"])
    while s >= 4:
        lh = s * 1.2 * float(spec.get("spacing", 1.0))
        lines = _wrap(pen.fk, text, s, inner.width)
        if s * 0.95 + (len(lines) - 1) * lh + s * 0.25 <= inner.height:
            _paint(pen, lines, s, inner.x0, inner.x1, inner.y0 + s * 0.95, lh, spec.get("align", "left"))
            return
        if not spec.get("fit", True):
            break
        s -= 0.5
    raise RuntimeError(_tr("The text does not fit in that box - draw a bigger box or turn on \"Shrink text to fit\"."))


_DOWN = {0: (0, 1), 90: (1, 0), 180: (0, -1), 270: (-1, 0)}


def write_replacement(pg, spec, text, fk, area, rot, origin, size, block):
    """The new text in place of the old one (the old text has already been removed)."""
    sp = float(spec.get("spacing", 1.0))
    s = float(spec["size"])
    ox, oy = origin
    if rot == 0:                                                       # normal horizontal text
        pen = _Pen(pg, spec, fk, lambda x, y: fitz.Point(x, y), 0)
        al = spec.get("align", "left")
        if block:
            x0, x1 = area.x0, area.x1 + 2
            while s >= 4:
                lh = s * 1.2 * sp
                lines = _wrap(pen.fk, text, s, x1 - x0)
                if oy + (len(lines) - 1) * lh <= area.y1 + 4 * size:
                    _paint(pen, lines, s, x0 if al != "left" else ox, x1, oy, lh, al)
                    return
                s -= 0.5
            raise RuntimeError(_tr("The new text is too long for this paragraph."))
        lines = [(l, True) for l in text.split("\n")]
        x0 = ox if al == "left" else area.x0
        _paint(pen, lines, s, x0, area.x1, oy, s * 1.2 * sp, al)
        return
    fk = _prep(pg, fk)                                                 # vertical text: plain version
    col = spec["color"]
    if block:
        al = {"left": 0, "center": 1, "right": 2, "justify": 3}[spec.get("align", "left")]
        while s >= 4:
            if pg.insert_textbox(area, text, fontsize=s, color=col, rotate=rot, align=al, lineheight=1.2 * sp, **fk) >= 0:
                return
            s -= 0.5
        raise RuntimeError(_tr("The new text is too long for this paragraph."))
    dx, dy = _DOWN[rot]
    for k, line in enumerate(text.split("\n")):
        pg.insert_text(fitz.Point(ox + dx * s * 1.2 * sp * k, oy + dy * s * 1.2 * sp * k), line, fontsize=s,
                       color=col, rotate=rot, **fk)


# ------------------------------------------------------------------------------------------------ the window
def _align_icon(kind, color="#303030"):
    im = Image.new("RGBA", (22, 16), (0, 0, 0, 0))
    dr = ImageDraw.Draw(im)
    widths = {"left": (20, 12, 20, 8), "center": (20, 12, 20, 8), "right": (20, 12, 20, 8), "justify": (20, 20, 20, 20)}[kind]
    for k, w in enumerate(widths):
        y = 1 + k * 4
        x0 = 1 if kind in ("left", "justify") else (1 + (20 - w) // 2 if kind == "center" else 21 - w)
        dr.rectangle([x0, y, x0 + w - 1, y + 1], fill=color)
    return im


class TextTools:
    # ---- remembered format
    def tt_spec(self, **over):
        spec = dict(DEFAULT_SPEC)
        spec["color"] = (0, 0, 0)
        spec["size"] = float(self.font_size.get()) if hasattr(self, "font_size") else 14.0
        try:
            saved = get_setting("text_fmt", {}) or {}
            for k in ("font", "bold", "italic", "underline", "strike", "align"):
                if k in saved and type(saved[k]) is type(spec[k]):
                    spec[k] = saved[k]
            if isinstance(saved.get("spacing"), (int, float)):
                spec["spacing"] = float(saved["spacing"])
            c = saved.get("color")
            if isinstance(c, (list, tuple)) and len(c) == 3 and all(isinstance(v, (int, float)) for v in c):
                spec["color"] = tuple(max(0.0, min(1.0, float(v))) for v in c)
        except Exception:
            pass
        if spec["font"] != ORIGINAL and spec["font"] not in [k for _l, k in font_catalog()]:
            spec["font"] = "base:he"
        spec.update(over)
        return spec

    def tt_remember(self, spec):
        try:
            keep = {k: spec[k] for k in ("font", "bold", "italic", "underline", "strike", "align", "spacing")
                    if not (k == "font" and spec[k] == ORIGINAL)}
            keep["color"] = list(spec["color"])
            set_setting("text_fmt", keep)
        except Exception:
            pass
        try:
            self.font_size.set(int(round(spec["size"])))
        except Exception:
            pass

    # ---- preview helpers
    def _tt_base(self, pg):
        """the page as a picture (once per window): (PIL image, zoom)"""
        longest = max(pg.rect.width, pg.rect.height)
        z = max(0.8, min(2.5, 3600 / longest))
        pix = pg.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
        return Image.frombytes("RGB", (pix.width, pix.height), pix.samples), z

    def _tt_temp_page(self, pg):
        doc = fitz.open()
        mb = pg.mediabox
        tp = doc.new_page(width=mb.width, height=mb.height)
        tp.set_rotation(pg.rotation)
        return doc, tp

    # ---- the dialog: controls on the left, big preview on the right
    def tt_dialog(self, title, text, spec, painter, region, erase=None, original=None, box=False, edit=False,
                  hint=""):
        """Returns (text, spec) when the user pressed Apply, otherwise None.
        painter(tp, text, spec, fk) draws the text on a blank copy of the page; region / erase are in view coordinates."""
        pg = self.page()
        base, z = self._tt_base(pg)
        d = tk.Toplevel(self)
        d.title(title)
        d.transient(self)
        d.configure(bg=BG, padx=14, pady=12)
        d.resizable(False, False)
        d.geometry("+%d+%d" % (self.winfo_rootx() + 60, self.winfo_rooty() + 40))
        res = {}
        v_font = tk.StringVar()
        v_size = tk.StringVar(value=("%g" % spec["size"]))
        v_sizeS = tk.DoubleVar(value=spec["size"])
        v_bold = tk.BooleanVar(value=spec["bold"])
        v_ital = tk.BooleanVar(value=spec["italic"])
        v_under = tk.BooleanVar(value=spec.get("underline", False))
        v_strike = tk.BooleanVar(value=spec.get("strike", False))
        v_align = tk.StringVar(value=spec["align"])
        v_space = tk.DoubleVar(value=spec["spacing"])
        v_zoom = tk.DoubleVar(value=1.0)
        v_fit = tk.BooleanVar(value=spec.get("fit", True))
        v_fill = tk.BooleanVar(value=bool(spec.get("fill")))
        v_brd = tk.BooleanVar(value=bool(spec.get("border")))
        st = {"color": tuple(spec["color"]), "fill": tuple(spec.get("fill") or (1, 1, 0.8)),
              "border": tuple(spec.get("border") or (0, 0, 0)), "job": None, "sync": False}

        choices = ([(_tr("Same as the original text"), ORIGINAL)] if (edit and original) else []) + font_catalog()
        labels = [l for l, _k in choices]
        by_label = dict(choices)
        v_font.set(next((l for l, k in choices if k == spec["font"]), labels[0]))

        left = tk.Frame(d, bg=BG)
        left.grid(row=0, column=0, sticky="n", padx=(0, 16))
        right = tk.Frame(d, bg=BG)
        right.grid(row=0, column=1, sticky="n")

        def section(parent, name):
            f = tk.Frame(parent, bg=BG)
            f.pack(fill="x", pady=(10, 0))
            tk.Label(f, text=_tr(name), bg=BG, fg=TH.TERTIARY, font=(UI_FONT, 8, "bold")).pack(anchor="w")
            tk.Frame(f, bg="#c7c7cc", height=1).pack(fill="x", pady=(1, 4))
            inner = tk.Frame(f, bg=BG)
            inner.pack(fill="x")
            return inner

        # ---------------- text
        tk.Label(left, text=_tr(hint or "Type the text:"), bg=BG, font=(UI_FONT, 9), anchor="w").pack(anchor="w")
        tf = tkfont.Font(family="Arial", size=11)
        t = tk.Text(left, width=44, height=6 if not edit else 5, wrap="word", font=tf, relief="solid", bd=1, undo=True)
        t.pack(fill="x", pady=(2, 0))
        t.insert("1.0", text)
        t.focus_set()
        t.mark_set("insert", "end-1c")
        if edit:
            t.tag_add("sel", "1.0", "end-1c")

        # ---------------- font
        fs = section(left, "FONT")
        ttk.Combobox(fs, textvariable=v_font, values=labels, state="readonly", width=40).pack(fill="x")
        row = tk.Frame(fs, bg=BG)
        row.pack(fill="x", pady=(6, 0))
        tk.Label(row, text=_tr("Size"), bg=BG, font=(UI_FONT, 9)).pack(side="left")
        slider = ttk.Scale(row, from_=6, to=120, orient="horizontal", variable=v_sizeS, length=S(200))
        slider.pack(side="left", padx=8)
        ttk.Spinbox(row, from_=4, to=300, increment=1, width=5, textvariable=v_size).pack(side="left")
        tk.Label(row, text="pt", bg=BG, font=(UI_FONT, 9)).pack(side="left", padx=(3, 0))
        srow = tk.Frame(fs, bg=BG)
        srow.pack(fill="x", pady=(6, 0))
        btns = {}
        for var, txt, fnt in ((v_bold, "B", (UI_FONT, 10, "bold")), (v_ital, "I", (UI_FONT, 10, "italic")),
                              (v_under, "U", (UI_FONT, 10, "underline")), (v_strike, "S", (UI_FONT, 10, "overstrike"))):
            b = tk.Checkbutton(srow, text=txt, variable=var, indicatoron=False, width=3, font=fnt, bg=BG, selectcolor="#c9dcf5",
                               activebackground=BG, relief="groove", bd=1)
            b.pack(side="left", padx=(0, 3))
            btns[txt] = b

        # ---------------- paragraph
        ps = section(left, "PARAGRAPH")
        arow = tk.Frame(ps, bg=BG)
        arow.pack(fill="x")
        icons = {}
        for kind, tip in (("left", "Left"), ("center", "Centre"), ("right", "Right"), ("justify", "Justify")):
            icons[kind] = ImageTk.PhotoImage(_align_icon(kind))
            tk.Radiobutton(arow, image=icons[kind], variable=v_align, value=kind, indicatoron=False, bg=BG, selectcolor="#c9dcf5",
                           activebackground=BG, relief="groove", bd=1, width=S(30), height=S(22)).pack(side="left", padx=(0, 3))
        tk.Label(arow, text=_tr("Line spacing"), bg=BG, font=(UI_FONT, 9)).pack(side="left", padx=(14, 6))
        ttk.Scale(arow, from_=0.8, to=3.0, orient="horizontal", variable=v_space, length=S(110)).pack(side="left")
        sp_lbl = tk.Label(arow, text="1.0", bg=BG, font=(UI_FONT, 9), width=4)
        sp_lbl.pack(side="left")

        # ---------------- colour
        cs = section(left, "COLOUR")
        crow = tk.Frame(cs, bg=BG)
        crow.pack(fill="x")
        cur_sw = tk.Frame(crow, width=34, height=20, bg=_hex(st["color"]), highlightthickness=2, highlightbackground="#4a4a4a")
        cur_sw.pack(side="left", padx=(0, 10))

        def set_color(rgb):
            st["color"] = tuple(rgb)
            cur_sw.config(bg=_hex(rgb))
            refresh()

        for h in QUICK_COLORS:
            b = tk.Frame(crow, width=18, height=18, bg=h, highlightthickness=1, highlightbackground="#9aa0a6", cursor="hand2")
            b.pack(side="left", padx=2)
            b.bind("<Button-1>", lambda _e, h=h: set_color(_rgb(h)))

        def pick(key="color", btn=None):
            c = colorchooser.askcolor(color=_hex(st[key]), parent=d, title=_tr("Choose color"))
            if c and c[0]:
                st[key] = tuple(v / 255 for v in c[0])
                if btn is not None:
                    btn.config(bg=_hex(st[key]), activebackground=_hex(st[key]))
                if key == "color":
                    cur_sw.config(bg=_hex(st[key]))
                refresh()
        ttk.Button(crow, text=_tr("More..."), command=pick, width=8).pack(side="left", padx=(8, 0))

        # ---------------- box options
        if box:
            bs = section(left, "BOX")
            ttk.Checkbutton(bs, text=_tr("Shrink text to fit"), variable=v_fit).grid(row=0, column=0, sticky="w", columnspan=4)
            ttk.Checkbutton(bs, text=_tr("Background"), variable=v_fill).grid(row=1, column=0, sticky="w", pady=(4, 0))
            fb = tk.Button(bs, width=3, relief="solid", bd=1, bg=_hex(st["fill"]), cursor="hand2")
            fb.config(command=lambda: pick("fill", fb))
            fb.grid(row=1, column=1, padx=(4, 18), pady=(4, 0))
            ttk.Checkbutton(bs, text=_tr("Border"), variable=v_brd).grid(row=1, column=2, sticky="w", pady=(4, 0))
            bb = tk.Button(bs, width=3, relief="solid", bd=1, bg=_hex(st["border"]), cursor="hand2")
            bb.config(command=lambda: pick("border", bb))
            bb.grid(row=1, column=3, padx=4, pady=(4, 0))

        # ---------------- right side: preview(s)
        PW = S(560)
        PH = S(250) if edit else S(380)
        tk.Label(right, text=_tr("PREVIEW"), bg=BG, fg=TH.TERTIARY, font=(UI_FONT, 8, "bold")).pack(anchor="w")
        tk.Frame(right, bg="#c7c7cc", height=1).pack(fill="x", pady=(1, 4))
        if edit:
            tk.Label(right, text=_tr("Before"), bg=BG, font=(UI_FONT, 9, "bold")).pack(anchor="w")
            bfr = tk.Frame(right, width=PW, height=S(100), bg="#ffffff", relief="solid", bd=1)
            bfr.pack()
            bfr.pack_propagate(False)
            bl = tk.Label(bfr, bg="#ffffff")
            bl.pack(expand=True)
            tk.Label(right, text=_tr("After"), bg=BG, font=(UI_FONT, 9, "bold")).pack(anchor="w", pady=(8, 0))
        pvf = tk.Frame(right, width=PW, height=PH, bg="#ffffff", relief="solid", bd=1)
        pvf.pack()
        pvf.pack_propagate(False)
        pv = tk.Label(pvf, bg="#ffffff")
        pv.pack(expand=True)
        zrow = tk.Frame(right, bg=BG)
        zrow.pack(fill="x", pady=(6, 0))
        tk.Label(zrow, text=_tr("Zoom"), bg=BG, font=(UI_FONT, 9)).pack(side="left")
        ttk.Scale(zrow, from_=1.0, to=4.0, orient="horizontal", variable=v_zoom, length=S(200)).pack(side="left", padx=8)
        zl = tk.Label(zrow, text="100%", bg=BG, font=(UI_FONT, 9), width=6)
        zl.pack(side="left")
        msg = tk.Label(right, text="", bg=BG, fg="#c0392b", font=(UI_FONT, 9), anchor="w", justify="left", wraplength=PW)
        msg.pack(anchor="w", pady=(4, 0))

        def collect():
            try:
                size = max(4.0, min(300.0, float(v_size.get().replace(",", "."))))
            except ValueError:
                size = spec["size"]
            return {"font": by_label.get(v_font.get(), "base:he"), "size": size, "bold": v_bold.get(),
                    "italic": v_ital.get(), "underline": v_under.get(), "strike": v_strike.get(), "color": st["color"],
                    "align": v_align.get(), "spacing": round(max(0.5, min(4.0, float(v_space.get()))), 1),
                    "fit": v_fit.get(), "fill": st["fill"] if v_fill.get() else None,
                    "border": st["border"] if v_brd.get() else None}

        def crop_box(rg, factor, centre=None):
            """the part of the page that is shown: the region, made smaller (= zoomed in) around `centre`"""
            w, h = rg.width / factor, rg.height / factor
            cx, cy = centre if centre else (rg.x0 + w / 2, rg.y0 + h / 2)
            x0 = max(rg.x0, min(rg.x1 - w, cx - w / 2))
            y0 = max(rg.y0, min(rg.y1 - h, cy - h / 2))
            return (int(x0), int(y0), max(int(x0) + 1, int(x0 + w)), max(int(y0) + 1, int(y0 + h)))

        def draw():
            st["job"] = None
            s = collect()
            txt = t.get("1.0", "end-1c")
            key = s["font"]
            fam = BASE14.get(key[5:], BASE14["he"])[2] if key.startswith("base:") else key[5:] if key.startswith("file:") else "Arial"
            tf.configure(family=fam, size=max(9, min(20, int(round(s["size"])))), weight="bold" if s["bold"] else "normal",
                         slant="italic" if s["italic"] else "roman", underline=s["underline"], overstrike=s["strike"])
            for k in ("B", "I"):
                btns[k].config(state="normal")
            sp_lbl.config(text="%.1f" % s["spacing"])
            zf = max(1.0, min(4.0, float(v_zoom.get())))
            zl.config(text="%d%%" % round(zf * 100))
            msg.config(text="")
            doc, tp = self._tt_temp_page(pg)
            try:
                try:
                    fk = resolve_font(s, txt, original)
                    if txt.strip():
                        painter(tp, txt, s, fk)
                except Exception as ex:
                    msg.config(text=str(ex))
                ov = tp.get_pixmap(matrix=fitz.Matrix(z, z), alpha=True)
                ovi = Image.frombytes("RGBA", (ov.width, ov.height), ov.samples)
            finally:
                doc.close()
            im = base.copy()
            if erase is not None:
                er = fitz.Rect(erase) * z
                px = im.getpixel((max(0, min(im.width - 1, int(er.x0) - 3)), max(0, min(im.height - 1, int((er.y0 + er.y1) / 2)))))
                ImageDraw.Draw(im).rectangle([er.x0, er.y0, er.x1, er.y1], fill=px)
            im.paste(ovi, (0, 0), ovi)
            rg = fitz.Rect(region) * z
            ink = ovi.getbbox()                                  # where the new text is: always keep it in view
            centre = None
            if ink:
                ir = fitz.Rect(ink) + (-14, -14, 14, 14)
                rg = rg | ir
                centre = (ink[0] - 14 + rg.width / zf / 2, (ink[1] + ink[3]) / 2)       # zoom starts where the text starts
            rg = rg & fitz.Rect(0, 0, im.width, im.height)
            box_ = crop_box(rg, zf, centre)
            crop = im.crop(box_)
            k = min(PW / crop.width, PH / crop.height, 6.0)
            crop = crop.resize((max(1, int(crop.width * k)), max(1, int(crop.height * k))), Image.LANCZOS)
            ph = ImageTk.PhotoImage(crop)
            pv.config(image=ph)
            pv.image = ph
            if edit:
                b0 = base.crop(box_)
                k0 = min(PW / b0.width, S(96) / b0.height, 6.0)
                b0 = b0.resize((max(1, int(b0.width * k0)), max(1, int(b0.height * k0))), Image.LANCZOS)
                st["before"] = ImageTk.PhotoImage(b0)
                bl.config(image=st["before"])

        def refresh(*_a):
            if st["job"] is not None:
                d.after_cancel(st["job"])
            st["job"] = d.after(150, draw)

        def slider_moved(*_a):                         # slider -> number box
            if st["sync"]:
                return
            st["sync"] = True
            v_size.set("%g" % round(v_sizeS.get()))
            st["sync"] = False

        def number_typed(*_a):                         # number box -> slider
            if st["sync"]:
                return
            try:
                st["sync"] = True
                v_sizeS.set(max(6.0, min(120.0, float(v_size.get().replace(",", ".")))))
            except ValueError:
                pass
            finally:
                st["sync"] = False

        v_sizeS.trace_add("write", slider_moved)
        v_size.trace_add("write", number_typed)
        for v in (v_font, v_size, v_bold, v_ital, v_under, v_strike, v_align, v_space, v_zoom, v_fit, v_fill, v_brd):
            v.trace_add("write", refresh)
        t.bind("<<Modified>>", lambda e: (t.edit_modified(False), refresh()))

        def apply(_e=None):
            txt = t.get("1.0", "end-1c")
            if not txt.strip():
                d.destroy()
                return "break"
            res["v"] = (txt, collect())
            d.destroy()
            return "break"

        def delete_it():
            res["v"] = (None, collect())
            d.destroy()

        bf = tk.Frame(d, bg=BG)
        bf.grid(row=1, column=0, columnspan=2, sticky="e", pady=(12, 0))
        if edit:
            tk.Button(bf, text=_tr("Delete text"), command=delete_it, bg="#d92d20", fg="white", relief="flat", padx=12,
                      pady=2, cursor="hand2", font=(UI_FONT, 9)).pack(side="left", padx=(0, 30))
        tk.Label(bf, text=_tr("Ctrl+Enter = Apply"), bg=BG, fg=TH.TERTIARY, font=(UI_FONT, 8)).pack(side="left", padx=(0, 16))
        ttk.Button(bf, text=_tr("Apply"), command=apply).pack(side="left", padx=4)
        ttk.Button(bf, text=_tr("Cancel"), command=d.destroy).pack(side="left")
        d.bind("<Control-Return>", apply)
        d.bind("<Escape>", lambda e: d.destroy())
        draw()
        try:
            d.grab_set()
        except tk.TclError:
            pass
        d.wait_window()
        return res.get("v")

    # ---- add text with one click
    def tt_add_click(self, e):
        pt = self.to_pdf_point(e.x, e.y)
        pg = self.page()
        vpt = fitz.Point(pt) * pg.rotation_matrix
        spec0 = self.tt_spec()
        size = spec0["size"]
        region = fitz.Rect(vpt.x - 50, vpt.y - size * 2.2 - 8, vpt.x + 420, vpt.y + 150) & pg.rect

        def painter(tp, txt, s, fk):
            write_lines(tp, s, txt, fk, pt)
        r = self.tt_dialog(_tr("Add text"), "", spec0, painter, region, hint="Type the text that should appear where you clicked:")
        if not r:
            return
        txt, spec = r
        fk = resolve_font(spec, txt)
        self.tt_remember(spec)
        self._run_edit(lambda: write_lines(pg, spec, txt, fk, pt), msg=_tr("Text added."))

    # ---- text box
    def add_textbox(self, rect):
        pg = self.page()
        spec0 = self.tt_spec()
        view = fitz.Rect(rect) * pg.rotation_matrix
        view.normalize()
        region = fitz.Rect(view.x0 - 40, view.y0 - 30, max(view.x1, view.x0 + 300) + 40, max(view.y1, view.y0 + 90) + 30) & pg.rect

        def painter(tp, txt, s, fk):
            write_box(tp, s, txt, fk, fitz.Rect(rect))
        r = self.tt_dialog(_tr("Text box"), "", spec0, painter, region, box=True, hint="Type the text for the box:")
        if not r:
            return
        txt, spec = r
        fk = resolve_font(spec, txt)
        self.tt_remember(spec)
        self._run_edit(lambda: write_box(pg, spec, txt, fk, fitz.Rect(rect)), msg=_tr("Text box added."))

    # ---- edit existing text
    def edit_text_at(self, pt, block=False):
        b, l = self.find_line(pt)
        if not l:
            self.status.config(text=_tr("There is no text at that spot (scanned pages need OCR first)."))
            return
        rot = EX._rot_from_dir(l["dir"])
        if rot is None:
            messagebox.showinfo(APP, _tr("This text is slanted - editing is only supported for horizontal or vertical text."))
            return
        if block:
            lines = [l2 for l2 in b["lines"] if l2["spans"]]
            old = "\n".join("".join(s["text"] for s in l2["spans"]) for l2 in lines)
            area = fitz.Rect(b["bbox"])
        else:
            old = "".join(s["text"] for s in l["spans"])
            area = fitz.Rect(l["bbox"])
        sp = l["spans"][0]
        size = sp["size"]
        col = EX._int_color(sp["color"])
        bold = bool(sp["flags"] & 16) or "bold" in sp["font"].lower()
        ital = bool(sp["flags"] & 2) or "italic" in sp["font"].lower() or "oblique" in sp["font"].lower()
        pg = self.page()
        spec0 = self.tt_spec(font=ORIGINAL, size=round(float(size), 1), color=tuple(col), bold=bold, italic=ital,
                             align="left", spacing=1.0, underline=False, strike=False)
        original = (self.doc, pg, sp)
        view = fitz.Rect(area) * pg.rotation_matrix
        view.normalize()
        region = fitz.Rect(view.x0 - 40, view.y0 - 24, max(view.x1, view.x0 + 360) + 40, view.y1 + 40) & pg.rect

        def red_rect():
            red = fitz.Rect(area)
            red.y0 += area.height * 0.08 if not block else 1
            red.y1 -= area.height * 0.08 if not block else 1
            red.x0 += 0.5
            red.x1 -= 0.5
            return red

        def painter(tp, txt, s, fk):
            write_replacement(tp, s, txt, fk, area, rot, sp["origin"], size, block)
        r = self.tt_dialog(_tr("Edit text"), old, spec0, painter, region, erase=view, original=original, edit=True, box=False,
                           hint="Change the text:")
        if r is not None and r[0] is None:                       # "Delete text"
            self.delete_text_area(area, block)
            return
        if r is None or (r[0].strip() == old.strip() and not self._tt_changed(r[1], spec0)):
            return
        new, spec = r[0].rstrip("\n"), r[1]
        fk = resolve_font(spec, new, original)
        if spec["font"] != ORIGINAL:
            self.tt_remember(spec)

        def work():
            self._remove_text(pg, fitz.Rect(area), block)          # also takes old underline / frame with it
            write_replacement(pg, spec, new, fk, area, rot, sp["origin"], size, block)
        self._run_edit(work, msg=_tr("Text changed."))

    @staticmethod
    def _tt_changed(a, b):
        return any(a[k] != b[k] for k in ("font", "size", "bold", "italic", "underline", "strike", "color", "align", "spacing"))
