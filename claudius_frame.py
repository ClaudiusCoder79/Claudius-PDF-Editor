# -*- coding: utf-8 -*-
"""Claudius PDF Editor - movable text frames.

Add text / Text box: click or drag on the page and a frame appears. Type straight into it; drag its edge (or the
little handle on top) to move it, drag the squares to resize it. A format bar floats on top of the page (font, size,
bold, italic, underline, strike, colour, alignment, frame, background). "Apply" writes the text onto the page.
The Edit text tool can also drag existing text to a new place.
"""
import math
import sys
import tkinter as tk
from tkinter import ttk, colorchooser, messagebox
from tkinter import font as tkfont

try:
    import pymupdf as fitz
except ImportError:
    import fitz
from PIL import ImageTk

import claudius_theme as TH
from claudius_i18n import tr as _tr
import claudius_extra as EX
from claudius_extra import APP, UI_FONT, S
from claudius_view import get_setting, set_setting
from claudius_text import (font_catalog, resolve_font, write_box, _wrap, _prep, ORIGINAL, _hex, _rgb, _align_icon,
                           QUICK_COLORS, BASE14)

ACCENT = getattr(TH, "ACCENT", "#0a84ff")
RING = 6            # grab ring around the frame (pixels)
HANDLE = 5          # half size of a resize square
MIN_W, MIN_H = 36.0, 18.0
FILL0 = (1.0, 0.96, 0.75)
_WIN = sys.platform.startswith("win")
# the normal Windows pointers on Windows, the nearest ones elsewhere
CUR = {"move": "size" if _WIN else "fleur", "text": "ibeam" if _WIN else "xterm", "hand": "hand2",
       "cross": "crosshair", "ns": "size_ns" if _WIN else "sb_v_double_arrow",
       "we": "size_we" if _WIN else "sb_h_double_arrow",
       "nwse": "size_nw_se" if _WIN else "bottom_right_corner", "nesw": "size_ne_sw" if _WIN else "bottom_left_corner",
       "arrow": "arrow"}
_HANDLE_CUR = {"nw": "nwse", "se": "nwse", "ne": "nesw", "sw": "nesw", "n": "ns", "s": "ns", "e": "we", "w": "we",
               "move": "move", "delete": "arrow"}


def needed_height(pg, spec, text, fk, width):
    """height (points) the text needs inside a frame of this width - same maths as write_box"""
    s = float(spec["size"])
    inner = width - 6 if width > 10 else width
    lines = _wrap(_prep(pg, fk), text, s, inner)
    lh = s * 1.2 * float(spec.get("spacing", 1.0))
    return s * 0.95 + (len(lines) - 1) * lh + s * 0.25 + 4.5


def _family(key):
    kind, _, data = key.partition(":")
    if kind == "file":
        return data
    return BASE14.get(data, BASE14["he"])[2]


class FrameTools:
    tf = None            # the text frame that is being edited (dict) or None
    _tfd = None          # mouse drag in progress
    _mvl = None          # line grabbed by the Edit text tool

    # ================================================================= helpers
    @staticmethod
    def _tool_cursor(t):
        """pointer for a tool: the normal Windows ones (hand, I-beam, cross-hair, move arrows)"""
        return {"view": CUR["hand"], "beam": CUR["text"], "text": CUR["text"], "textbox": CUR["cross"],
                "edittext": CUR["text"], "move": CUR["move"]}.get(t, CUR["cross"])

    def tf_rect_c(self):
        tf = self.tf
        x0, y0 = self._v2c(tf["x0"], tf["y0"])
        x1, y1 = self._v2c(tf["x1"], tf["y1"])
        return x0, y0, x1, y1

    def tf_handles(self):
        x0, y0, x1, y1 = self.tf_rect_c()
        x0, y0, x1, y1 = x0 - RING, y0 - RING, x1 + RING, y1 + RING
        mx, my = (x0 + x1) / 2, (y0 + y1) / 2
        return {"nw": (x0, y0), "n": (mx, y0), "ne": (x1, y0), "e": (x1, my),
                "se": (x1, y1), "s": (mx, y1), "sw": (x0, y1), "w": (x0, my)}

    def tf_hit(self, cx, cy):
        """'move', a handle name or None for a canvas position"""
        if not self.tf:
            return None
        x0, y0, x1, y1 = self.tf_rect_c()
        if x0 - RING + 40 <= cx <= x0 - RING + 58 and y0 - RING - 18 <= cy <= y0 - RING:
            return "delete"                                         # the little cross: remove the frame
        if x0 - RING + 14 <= cx <= x0 - RING + 38 and y0 - RING - 18 <= cy <= y0 - RING:
            return "move"                                           # the little grip on top
        for name, (hx, hy) in self.tf_handles().items():
            if abs(cx - hx) <= HANDLE + 4 and abs(cy - hy) <= HANDLE + 4:
                return name
        if x0 - RING - 4 <= cx <= x1 + RING + 4 and y0 - RING - 4 <= cy <= y1 + RING + 4:
            return "move"
        return None

    # ================================================================= creating / closing
    def tf_start(self, x0, y0, x1, y1, text="", spec=None, fmt=None):
        """open a frame (view points on the current page)"""
        self.tf_commit()
        pw, ph = self.page_w / self.zoom, self.page_h / self.zoom
        spec = dict(spec or self.tt_spec())
        spec["fit"] = True
        saved = get_setting("frame_fmt", {}) or {}
        fmt = fmt or {}
        fill = saved.get("fill")
        if not (isinstance(fill, (list, tuple)) and len(fill) == 3):
            fill = FILL0
        size = float(spec["size"])
        x0, x1 = sorted((max(0.0, x0), min(pw, x1)))
        y0, y1 = sorted((max(0.0, y0), min(ph, y1)))
        if x1 - x0 < MIN_W:
            x1 = min(pw, x0 + max(MIN_W, 220.0))
            x0 = max(0.0, x1 - max(MIN_W, 220.0))
        if y1 - y0 < MIN_H:
            y1 = min(ph, y0 + max(MIN_H, size * 1.2 + 8))
        self.tf = {"page": self.page_no, "x0": x0, "y0": y0, "x1": x1, "y1": y1, "spec": spec,
                   "fill": tuple(fill), "border": bool(fmt.get("border", saved.get("border", False))),
                   "fillon": bool(fmt.get("fillon", saved.get("fillon", False))), "win": None, "font": None}
        self._tf_bar()
        self._tf_load_vars()
        t = tk.Text(self.canvas, wrap="word", undo=True, bd=0, relief="flat", highlightthickness=0,
                    insertwidth=2, cursor=CUR["text"], exportselection=False)
        tf = self.tf
        tf["text"] = t
        tf["font"] = tkfont.Font(family=_family(spec["font"]), size=-10)
        t.config(font=tf["font"])
        t.tag_configure("a", justify="left")
        t.bindtags((str(t), "Text"))
        t.bind("<Control-Return>", lambda e: (self.tf_apply(), "break")[1])
        t.bind("<Escape>", lambda e: (self.canvas.focus_set(), "break")[1])
        t.bind("<Control-z>", lambda e: (self._tf_try(t.edit_undo), "break")[1])
        t.bind("<Control-y>", lambda e: (self._tf_try(t.edit_redo), "break")[1])
        t.bind("<Control-Z>", lambda e: (self._tf_try(t.edit_redo), "break")[1])
        t.bind("<Control-a>", lambda e: (t.tag_add("sel", "1.0", "end-1c"), "break")[1])
        t.bind("<Control-s>", lambda e: (self.save(), "break")[1])
        t.bind("<<Modified>>", self._tf_modified)
        if text:
            t.insert("1.0", text)
        t.edit_reset()
        t.edit_modified(False)
        self._tf_style()
        self.tf_draw()
        self.update_idletasks()
        t.focus_set()
        self.status.config(text=_tr("Type in the frame.  Drag its edge or the little handle on top to move it, "
                                    "drag the squares to resize.  Ctrl+Enter or Apply puts it on the page."))

    @staticmethod
    def _tf_try(fn):
        try:
            fn()
        except tk.TclError:
            pass
        return True

    def tf_close(self):
        tf, self.tf = self.tf, None
        self._tfd = None
        if tf:
            try:
                tf["text"].destroy()
            except Exception:
                pass
        try:
            self.canvas.delete("tfc")
            self.canvas.delete("tfw")
            self.tf_bar.place_forget()
        except Exception:
            pass

    def tf_cancel(self):
        if self.tf:
            self.tf_close()
            self.status.config(text=_tr("Text frame removed."))

    def tf_apply(self):
        """Apply button: put the text on the page and go back to the hand tool"""
        had = self.tf is not None
        self.tf_commit()
        if had:
            self.set_tool("view")

    def tf_commit(self):
        """write the text of the open frame onto its page (called before anything else changes the document)"""
        tf = self.tf
        if not tf:
            return False
        txt = tf["text"].get("1.0", "end-1c")
        spec = self.tf_collect()
        pno, rect = tf["page"], (tf["x0"], tf["y0"], tf["x1"], tf["y1"])
        fmt = {"border": tf["border"], "fillon": tf["fillon"]}
        self.tf_close()
        if not txt.strip() or not self.doc or not 0 <= pno < len(self.doc):
            return False
        pg = self.doc[pno]
        vr = fitz.Rect(*rect)
        fk = resolve_font(spec, txt)
        need = needed_height(pg, spec, txt, fk, vr.width)
        if need > vr.height:
            vr.y1 = vr.y0 + need
        prect = vr * pg.derotation_matrix
        prect.normalize()
        self.tt_remember(spec)
        try:
            set_setting("frame_fmt", {"border": fmt["border"], "fillon": fmt["fillon"],
                                      "fill": list(tf["fill"])})
        except Exception:
            pass
        if self.cont:
            self._cont_focus(pno)
        else:
            self.page_no = pno
        ok = self._run_edit(lambda: write_box(self.doc[pno], spec, txt, fk, prect), msg=_tr("Text added."),
                            thumbs=[pno])
        if not ok:                                              # keep the text, the user can try again
            self.tf_start(vr.x0, vr.y0, vr.x1, vr.y1, txt, spec, fmt)
        return ok

    def tf_collect(self):
        """format of the open frame as a spec for the writing engine"""
        tf = self.tf
        spec = dict(tf["spec"])
        v = self._tfv
        labels = {lab: key for lab, key in font_catalog()}
        spec["font"] = labels.get(v["font"].get(), spec["font"])
        try:
            spec["size"] = max(4.0, min(300.0, float(str(v["size"].get()).replace(",", "."))))
        except ValueError:
            pass
        spec["bold"], spec["italic"] = bool(v["bold"].get()), bool(v["italic"].get())
        spec["underline"], spec["strike"] = bool(v["under"].get()), bool(v["strike"].get())
        spec["align"] = v["align"].get()
        spec["fit"] = True
        spec["fill"] = tuple(tf["fill"]) if tf["fillon"] else None
        spec["border"] = tuple(spec["color"]) if tf["border"] else None
        return spec

    # ================================================================= drawing
    def tf_draw(self):
        tf = self.tf
        if not tf or tf["page"] != self.page_no:
            return
        c = self.canvas
        c.delete("tfc")
        x0, y0, x1, y1 = self.tf_rect_c()
        z = self.zoom
        t = tf["text"]
        spec = tf["spec"]
        tf["font"].configure(size=-max(2, int(round(float(spec["size"]) * z))))
        w, h = max(8, int(x1 - x0)), max(8, int(y1 - y0))
        if not c.find_withtag("tfw"):
            tf["win"] = c.create_window(x0, y0, anchor="nw", window=t, width=w, height=h, tags=("tfw",))
        else:
            c.coords("tfw", x0, y0)
            c.itemconfigure("tfw", width=w, height=h)
        t.config(padx=max(1, int(3 * z)), pady=max(1, int(2 * z)))
        c.create_rectangle(x0 - RING, y0 - RING, x1 + RING, y1 + RING, outline=ACCENT, dash=(4, 3), width=1, tags="tfc")
        gx, gy = x0 - RING + 14, y0 - RING - 17                  # grip on top
        c.create_rectangle(gx, gy, gx + 24, gy + 17, fill=ACCENT, outline=ACCENT, tags="tfc")
        c.create_text(gx + 12, gy + 8, text="✥", fill="white", font=(UI_FONT, 10), tags="tfc")
        c.create_rectangle(gx + 26, gy, gx + 44, gy + 17, fill="#d92d20", outline="#d92d20", tags="tfc")
        c.create_text(gx + 35, gy + 8, text="✕", fill="white", font=(UI_FONT, 9, "bold"), tags="tfc")
        for hx, hy in self.tf_handles().values():
            c.create_rectangle(hx - HANDLE, hy - HANDLE, hx + HANDLE, hy + HANDLE, fill="white", outline=ACCENT,
                               width=2, tags="tfc")
        c.tag_raise("tfc")
        self._tf_bar_place()

    def _tf_bar_place(self):
        """the format bar floats just above the frame (below it when there is no room above)"""
        try:
            bar = self.tf_bar
            bar.update_idletasks()
            x0, y0, x1, y1 = self.tf_rect_c()
            sx0, sy0 = x0 - self.canvas.canvasx(0), y0 - self.canvas.canvasy(0)
            sy1 = y1 - self.canvas.canvasy(0)
            cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
            bw, bh = bar.winfo_reqwidth(), bar.winfo_reqheight()
            x = max(4, min(sx0 - RING, cw - bw - 4))
            y = sy0 - RING - 17 - bh - 6
            if y < 4:
                y = sy1 + RING + 10
            y = max(4, min(y, ch - bh - 4))
            bar.place(in_=self.canvas, x=int(x), y=int(y), anchor="nw")
            bar.lift()
        except tk.TclError:
            pass

    def _tf_modified(self, _e=None):
        tf = self.tf
        if not tf:
            return
        t = tf["text"]
        if not t.edit_modified():
            return
        t.edit_modified(False)
        self.modified = True
        t.tag_add("a", "1.0", "end")
        # the frame grows while the text gets longer
        try:
            t.update_idletasks()
            n = t.count("1.0", "end-1c", "displaylines")
            n = n[0] if n else 1
            ls = tf["font"].metrics("linespace")
            need = (n * ls + 2 * max(1, int(2 * self.zoom)) + 3) / self.zoom
            if need > tf["y1"] - tf["y0"] + 0.5:
                tf["y1"] = min(self.page_h / self.zoom, tf["y0"] + need)
                self.tf_draw()
        except tk.TclError:
            pass

    # ================================================================= format bar
    def _tf_bar(self):
        if getattr(self, "tf_bar", None) is not None:
            return
        BGc = "#f2f2f7"
        bar = tk.Frame(self.canvas, bg=BGc, highlightthickness=1, highlightbackground="#b8b8c0", padx=6, pady=4)
        self.tf_bar = bar
        v = self._tfv = {"font": tk.StringVar(), "size": tk.StringVar(), "bold": tk.BooleanVar(), "italic": tk.BooleanVar(),
                         "under": tk.BooleanVar(), "strike": tk.BooleanVar(), "align": tk.StringVar(value="left"),
                         "border": tk.BooleanVar(), "fillon": tk.BooleanVar()}
        labels = [lab for lab, _k in font_catalog()]
        row1 = tk.Frame(bar, bg=BGc)
        row1.pack(fill="x")
        row2 = tk.Frame(bar, bg=BGc)
        row2.pack(fill="x", pady=(5, 0))
        cb = ttk.Combobox(row1, textvariable=v["font"], values=labels, state="readonly", width=26)
        cb.pack(side="left")
        sp = ttk.Spinbox(row1, from_=4, to=300, increment=1, width=4, textvariable=v["size"])
        sp.pack(side="left", padx=(6, 8))
        for var, txt, fnt in ((v["bold"], "B", (UI_FONT, 10, "bold")), (v["italic"], "I", (UI_FONT, 10, "italic")),
                              (v["under"], "U", (UI_FONT, 10, "underline")), (v["strike"], "S", (UI_FONT, 10, "overstrike"))):
            tk.Checkbutton(row1, text=txt, variable=var, indicatoron=False, width=3, font=fnt, bg=BGc, selectcolor="#c9dcf5",
                           activebackground=BGc, relief="groove", bd=1).pack(side="left", padx=(0, 2))
        self._tf_sw = tk.Label(row1, text=" A ", font=(UI_FONT, 10, "bold"), width=3, relief="solid", bd=1, cursor="hand2")
        self._tf_sw.pack(side="left", padx=(8, 4))
        self._tf_sw.bind("<Button-1>", lambda e: self._tf_pick("color"))
        for h in QUICK_COLORS[:7]:
            q = tk.Frame(row1, width=14, height=14, bg=h, highlightthickness=1, highlightbackground="#9aa0a6", cursor="hand2")
            q.pack(side="left", padx=1)
            q.bind("<Button-1>", lambda _e, h=h: self._tf_set_color(_rgb(h)))
        self._tf_icons = {}
        for kind in ("left", "center", "right", "justify"):
            self._tf_icons[kind] = ImageTk.PhotoImage(_align_icon(kind))
            tk.Radiobutton(row2, image=self._tf_icons[kind], variable=v["align"], value=kind, indicatoron=False, bg=BGc,
                           selectcolor="#c9dcf5", activebackground=BGc, relief="groove", bd=1,
                           width=S(28), height=S(20)).pack(side="left", padx=(0, 2))
        tk.Frame(row2, width=1, bg="#c7c7cc").pack(side="left", fill="y", padx=8)
        tk.Checkbutton(row2, text=_tr("Frame"), variable=v["border"], bg=BGc, activebackground=BGc,
                       font=(UI_FONT, 9)).pack(side="left")
        tk.Checkbutton(row2, text=_tr("Background"), variable=v["fillon"], bg=BGc, activebackground=BGc,
                       font=(UI_FONT, 9)).pack(side="left")
        self._tf_fillsw = tk.Label(row2, text="   ", relief="solid", bd=1, cursor="hand2")
        self._tf_fillsw.pack(side="left", padx=(0, 6))
        self._tf_fillsw.bind("<Button-1>", lambda e: self._tf_pick("fill"))
        ttk.Button(row2, text=_tr("Remove"), width=8, command=self.tf_cancel).pack(side="right")
        ttk.Button(row2, text=_tr("Apply"), width=7, command=self.tf_apply).pack(side="right", padx=(6, 4))
        self._tf_loading = False
        for k, var in v.items():
            var.trace_add("write", lambda *_a: self._tf_style())

    def _tf_load_vars(self):
        tf = self.tf
        spec = tf["spec"]
        v = self._tfv
        self._tf_loading = True
        try:
            labels = {key: lab for lab, key in font_catalog()}
            v["font"].set(labels.get(spec["font"], font_catalog()[0][0]))
            v["size"].set(str(int(round(float(spec["size"])))))
            v["bold"].set(bool(spec["bold"]))
            v["italic"].set(bool(spec["italic"]))
            v["under"].set(bool(spec["underline"]))
            v["strike"].set(bool(spec["strike"]))
            v["align"].set(spec["align"] if spec["align"] in ("left", "center", "right", "justify") else "left")
            v["border"].set(tf["border"])
            v["fillon"].set(tf["fillon"])
        finally:
            self._tf_loading = False

    def _tf_set_color(self, rgb):
        if self.tf:
            self.tf["spec"]["color"] = tuple(rgb)
            self._tf_style()

    def _tf_pick(self, key):
        tf = self.tf
        if not tf:
            return
        cur = tf["spec"]["color"] if key == "color" else tf["fill"]
        c = colorchooser.askcolor(color=_hex(cur), parent=self, title=_tr("Choose color"))
        if c and c[0]:
            rgb = tuple(x / 255 for x in c[0])
            if key == "color":
                tf["spec"]["color"] = rgb
            else:
                tf["fill"] = rgb
                tf["fillon"] = True
                self._tfv["fillon"].set(True)
            self._tf_style()
        try:
            tf["text"].focus_set()
        except tk.TclError:
            pass

    def _tf_style(self):
        """bring the editing widget in line with the format bar"""
        tf = self.tf
        if not tf or getattr(self, "_tf_loading", False):
            return
        v = self._tfv
        spec = tf["spec"]
        labels = {lab: key for lab, key in font_catalog()}
        spec["font"] = labels.get(v["font"].get(), spec["font"])
        try:
            spec["size"] = max(4.0, min(300.0, float(str(v["size"].get()).replace(",", "."))))
        except ValueError:
            pass
        spec["bold"], spec["italic"] = bool(v["bold"].get()), bool(v["italic"].get())
        spec["underline"], spec["strike"] = bool(v["under"].get()), bool(v["strike"].get())
        spec["align"] = v["align"].get()
        tf["border"], tf["fillon"] = bool(v["border"].get()), bool(v["fillon"].get())
        f = tf["font"]
        f.configure(family=_family(spec["font"]), weight="bold" if spec["bold"] else "normal",
                    slant="italic" if spec["italic"] else "roman", underline=int(spec["underline"]),
                    overstrike=int(spec["strike"]))
        t = tf["text"]
        col = _hex(spec["color"])
        bg = _hex(tf["fill"]) if tf["fillon"] else "#ffffff"
        t.config(fg=col, insertbackground=col, bg=bg, highlightthickness=2 if tf["border"] else 0,
                 highlightbackground=col, highlightcolor=col)
        t.tag_configure("a", justify={"left": "left", "center": "center", "right": "right", "justify": "left"}[spec["align"]])
        self._tf_sw.config(fg=col, bg="#ffffff")
        self._tf_fillsw.config(bg=_hex(tf["fill"]))
        self.tf_draw()

    # ================================================================= mouse
    def tf_press(self, e):
        """called first for every mouse press; True when the frame code used it"""
        if self.doc is None:
            return False
        t = self.tool.get()
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        if self.tf:
            if self.cont:
                self._cont_focus(self.tf["page"])
            hit = self.tf_hit(cx, cy)
            if hit == "delete":
                self.tf_cancel()
                return True
            if hit:
                self.canvas.focus_set()
                self._tfd = {"mode": hit, "last": (cx, cy)}
                return True
            if t not in ("text", "textbox"):
                return False
            self.tf_commit()                                   # click away = put it on the page
        if t in ("text", "textbox"):
            self._cont_focus_at(e)
            self._tfd = {"mode": "new", "start": (cx, cy), "item": None}
            return True
        return False

    def tf_motion(self, e):
        """pointer shape over the frame (resize arrows on the squares, move arrows on the edge); True when it is on the frame"""
        if not self.tf:
            return False
        keep = self.page_no
        try:
            if self.cont:
                self._cont_focus(self.tf["page"])
            hit = self.tf_hit(self.canvas.canvasx(e.x), self.canvas.canvasy(e.y))
        finally:
            if self.cont:
                self._cont_focus(keep)
        if not hit:
            return False
        want = CUR[_HANDLE_CUR[hit]]
        if self.canvas.cget("cursor") != want:
            self.canvas.config(cursor=want)
        return True

    def tf_drag(self, e):
        d = self._tfd
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        if d["mode"] == "new":
            if d["item"]:
                self.canvas.delete(d["item"])
            d["item"] = self.canvas.create_rectangle(*d["start"], cx, cy, outline=ACCENT, dash=(4, 3), width=2,
                                                     tags="preview")
            return
        tf = self.tf
        if not tf:
            return
        dx, dy = (cx - d["last"][0]) / self.zoom, (cy - d["last"][1]) / self.zoom
        d["last"] = (cx, cy)
        pw, ph = self.page_w / self.zoom, self.page_h / self.zoom
        m = d["mode"]
        if m == "move":
            w, h = tf["x1"] - tf["x0"], tf["y1"] - tf["y0"]
            nx = min(max(0.0, tf["x0"] + dx), max(0.0, pw - w))
            ny = min(max(0.0, tf["y0"] + dy), max(0.0, ph - h))
            tf["x0"], tf["y0"], tf["x1"], tf["y1"] = nx, ny, nx + w, ny + h
        else:
            if "w" in m:
                tf["x0"] = min(max(0.0, tf["x0"] + dx), tf["x1"] - MIN_W)
            if "e" in m:
                tf["x1"] = max(min(pw, tf["x1"] + dx), tf["x0"] + MIN_W)
            if "n" in m:
                tf["y0"] = min(max(0.0, tf["y0"] + dy), tf["y1"] - MIN_H)
            if "s" in m:
                tf["y1"] = max(min(ph, tf["y1"] + dy), tf["y0"] + MIN_H)
        self.tf_draw()

    def tf_release(self, e):
        d, self._tfd = self._tfd, None
        if d["mode"] != "new":
            self.modified = True
            return
        self.canvas.delete("preview")
        x0, y0 = d["start"]
        x1, y1 = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        a = self._c2v(x0, y0)
        if abs(x1 - x0) > 8 and abs(y1 - y0) > 8:
            b = self._c2v(x1, y1)
            self.tf_start(a[0], a[1], b[0], b[1])
        else:                                                   # a click: frame hangs from the clicked point
            self.tf_start(a[0], a[1], a[0] + 220, a[1])

    # ================================================================= moving existing text
    def tf_move_preview(self, e):
        if self._mvl is None:
            return
        b, l = self._mvl
        if l is None or not self._drag_start:
            return
        pg = self.page()
        r = fitz.Rect(l["bbox"]) * pg.rotation_matrix
        x0, y0 = self._v2c(r.x0, r.y0)
        x1, y1 = self._v2c(r.x1, r.y1)
        dx = self.canvas.canvasx(e.x) - self._drag_start[0]
        dy = self.canvas.canvasy(e.y) - self._drag_start[1]
        self.canvas.delete("preview")
        self.canvas.create_rectangle(x0, y0, x1, y1, outline="#8e8e93", dash=(2, 3), tags="preview")
        self.canvas.create_rectangle(x0 + dx, y0 + dy, x1 + dx, y1 + dy, outline=ACCENT, width=2, tags="preview")

    def tf_move_begin(self, e):
        try:
            self._mvl = self.find_line(self._canvas_pt(self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)))
        except Exception:
            self._mvl = None

    # ================================================================= removing text
    def text_line_menu(self, e):
        """right-click on text: edit / delete it. True when a menu was shown"""
        if self.doc is None:
            return False
        pt = self._canvas_pt(self.canvas.canvasx(e.x), self.canvas.canvasy(e.y))
        b, l = self.find_line(pt)
        if not l:
            return False
        m = tk.Menu(self, tearoff=0)
        m.add_command(label=_tr("Edit this text..."), command=lambda: self.edit_text_at(pt))
        m.add_command(label=_tr("Delete this line of text"), command=lambda: self.delete_text_at(pt, False))
        if len([x for x in b["lines"] if x["spans"]]) > 1:
            m.add_command(label=_tr("Delete this paragraph"), command=lambda: self.delete_text_at(pt, True))
        m.tk_popup(e.x_root, e.y_root)
        return True

    def delete_text_at(self, pt, block=False):
        b, l = self.find_line(pt)
        if not l:
            self.status.config(text=_tr("There is no text at that spot (scanned pages need OCR first)."))
            return
        lines = [x for x in b["lines"] if x["spans"]] if block else [l]
        area = fitz.Rect(lines[0]["bbox"])
        for x in lines[1:]:
            area |= fitz.Rect(x["bbox"])
        self.delete_text_area(area, block)

    def delete_text_area(self, area, block=False):
        pg = self.page()
        red = fitz.Rect(area)
        if block:
            red.y0 += 1
            red.y1 -= 1
        else:
            red.y0 += area.height * 0.08
            red.y1 -= area.height * 0.08
        red.x0 += 0.5
        red.x1 -= 0.5

        def work():
            pg.add_redact_annot(red, fill=False)
            pg.apply_redactions(images=getattr(fitz, "PDF_REDACT_IMAGE_NONE", 0),
                                graphics=getattr(fitz, "PDF_REDACT_LINE_ART_NONE", 0))
        self._run_edit(work, msg=_tr("Text deleted."))

    def move_text_line(self, p0, p1, block=False):
        """drag a line (or with Shift the whole paragraph) of existing text to another place"""
        b, l = self.find_line(p0)
        if not l:
            self.status.config(text=_tr("There is no text at that spot (scanned pages need OCR first)."))
            return
        rot = EX._rot_from_dir(l["dir"])
        if rot is None:
            messagebox.showinfo(APP, _tr("This text is slanted - editing is only supported for horizontal or vertical text."))
            return
        lines = [x for x in b["lines"] if x["spans"]] if block else [l]
        area = fitz.Rect(lines[0]["bbox"])
        for x in lines[1:]:
            area |= fitz.Rect(x["bbox"])
        d = fitz.Point(p1) - fitz.Point(p0)
        pg = self.page()
        red = fitz.Rect(area)
        if block:
            red.y0 += 1
            red.y1 -= 1
        else:
            red.y0 += area.height * 0.08
            red.y1 -= area.height * 0.08
        red.x0 += 0.5
        red.x1 -= 0.5

        def work():
            pg.add_redact_annot(red, fill=False)
            pg.apply_redactions(images=getattr(fitz, "PDF_REDACT_IMAGE_NONE", 0),
                                graphics=getattr(fitz, "PDF_REDACT_LINE_ART_NONE", 0))
            for ln in lines:
                for sp in ln["spans"]:
                    txt = sp["text"]
                    if not txt.strip():
                        continue
                    fk = _prep(pg, resolve_font({"font": ORIGINAL, "bold": False, "italic": False}, txt,
                                                (self.doc, pg, sp)))
                    pg.insert_text(fitz.Point(sp["origin"]) + d, txt, fontsize=sp["size"],
                                   color=tuple(EX._int_color(sp["color"])), rotate=rot, **fk)
        self._run_edit(work, msg=_tr("Text moved."))
