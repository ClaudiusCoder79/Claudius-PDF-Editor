"""
Claudius PDF Editor - viewing features: select & copy text, continuous scrolling, night view, full screen / presentation,
and several documents open at once (tabs).

Mixed into PdfEditor by claudius.py (see the hooks marked "view:" there).
"""
import bisect
import json
import math
import os
import time
import tkinter as tk
from PIL import Image, ImageTk
from collections import OrderedDict
from tkinter import ttk, messagebox

try:
    import pymupdf as fitz
except ImportError:
    import fitz

import claudius_theme as TH
from claudius_i18n import tr as _tr
from claudius_extra import APP, UI_FONT, S, MAX_PIXELS

PAD = S(24)
GAP = S(14)
PHOTO_CACHE = 10

# per-document state that belongs to a tab
TAB_KEYS = ("doc", "path", "page_no", "zoom", "fit_mode", "undo_stack", "redo_stack", "modified", "protect",
            "sel_pages", "_anchor", "floats", "sel", "float_page", "find_hits", "find_idx", "_sig_orig")


# ------------------------------------------------------------------------------------------ small settings store
def _cfg_path():
    return TH._config_file()


def get_setting(key, default=None):
    try:
        with open(_cfg_path(), "r", encoding="utf-8") as fh:
            return json.load(fh).get(key, default)
    except Exception:
        return default


def set_setting(key, value):
    try:
        os.makedirs(os.path.dirname(_cfg_path()), exist_ok=True)
        data = {}
        try:
            with open(_cfg_path(), "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:
            pass
        data[key] = value
        with open(_cfg_path(), "w", encoding="utf-8") as fh:
            json.dump(data, fh)
    except Exception:
        pass


# ------------------------------------------------------------------------------------------ text selection helpers
def nearest_word(words, x, y, max_dist=None):
    """words: list of (r0, rv, text, linekey) -> index of the word nearest to (x, y) in view points, or None"""
    best, bd = None, 1e18
    for i, (_r0, rv, _t, _k) in enumerate(words):
        dx = max(rv.x0 - x, 0, x - rv.x1)
        dy = max(rv.y0 - y, 0, y - rv.y1)
        d = dx * dx + dy * dy * 3.0
        if d < bd:
            best, bd = i, d
            if d == 0:
                break
    if best is None or (max_dist is not None and bd > max_dist * max_dist):
        return None
    return best


def selection_lines(words, i0, i1):
    """-> list of (linekey, r0 union, rv union) for words i0..i1 (inclusive)"""
    a, b = (i0, i1) if i0 <= i1 else (i1, i0)
    out = []
    for r0, rv, _t, k in words[a:b + 1]:
        if out and out[-1][0] == k:
            out[-1] = (k, out[-1][1] | r0, out[-1][2] | rv)
        else:
            out.append((k, fitz.Rect(r0), fitz.Rect(rv)))
    return out


def selection_text(words, i0, i1):
    a, b = (i0, i1) if i0 <= i1 else (i1, i0)
    parts, last = [], None
    for _r0, _rv, t, k in words[a:b + 1]:
        if last is not None:
            if k == last:
                parts.append(" ")
            else:
                parts.append("\n" if k[0] != last[0] else " ")
        parts.append(t)
        last = k
    return "".join(parts)


class ViewTools:
    # ===================================================================== init
    def _view_init(self):
        self.cont = bool(get_setting("continuous", True))
        if os.environ.get("CLAUDIUS_CONT") in ("0", "1"):
            self.cont = os.environ["CLAUDIUS_CONT"] == "1"
        self.night = bool(get_setting("night", False))
        self.presenting = False
        self._fs = None
        self._tsel = None                 # dict(page, i0, i1)
        self._tsel_drag = None
        self._tsel_moved = False
        self._gen = 0
        self._wcache = {}
        self._cont_raw = None             # [(w, h)] in points
        self._cont_key = None
        self._cont_tops, self._cont_x, self._cont_sizes = [], [], []
        self._cont_total = 0
        self._cont_photos = OrderedDict()
        self._cont_drawn = set()
        self._cont_upd_after = None
        self._cont_rel = 0.0
        self._cont_doc = None
        self._motion_t = 0.0
        self._cursor_after = None
        self._next_title = None
        self.tabs = []
        self.cur_tab = -1
        self._tab_sig = None
        self.tabbar = None

    # ===================================================================== geometry helpers
    def _page_origin(self, i):
        if self.cont and self._cont_tops:
            return self._cont_x[i], self._cont_tops[i]
        return self.off_x, self.off_y

    def _cont_page_at_y(self, y):
        if not self._cont_tops:
            return 0
        i = bisect.bisect_right(self._cont_tops, y) - 1
        return max(0, min(len(self._cont_tops) - 1, i))

    def _cont_focus(self, i):
        """Make page i the 'current' one for the editing tools (they work in page_no / off_x / off_y)."""
        if not self._cont_tops or not 0 <= i < len(self._cont_tops):
            return
        self.page_no = i
        self.off_x, self.off_y = self._cont_x[i], self._cont_tops[i]
        self.page_w, self.page_h = self._cont_sizes[i]

    def _cont_focus_at(self, e):
        if self.cont and self.doc is not None and self._cont_tops:
            self._cont_focus(self._cont_page_at_y(self.canvas.canvasy(e.y)))

    def _current_rel(self):
        """scroll position as a fraction of the current page (continuous) or of the canvas (single page)"""
        try:
            if self.cont and self._cont_tops and self._cont_doc is self.doc:
                y = self.canvas.canvasy(0)
                i = max(0, min(len(self._cont_tops) - 1, self.page_no))
                return (i, (y - self._cont_tops[i]) / max(1, self._cont_sizes[i][1]))
            return (self.page_no, self.canvas.yview()[0])
        except Exception:
            return (self.page_no, 0.0)

    # ===================================================================== night view / full screen
    def toggle_night(self):
        self.night = not self.night
        set_setting("night", self.night)
        self._cont_photos.clear()
        self.render()
        self.status.config(text=_tr("Night view on.") if self.night else _tr("Night view off."))

    def toggle_continuous(self):
        if self.doc is not None:
            cur = self._current_rel()
        self.cont = not self.cont
        set_setting("continuous", self.cont)
        if self.doc is not None:
            self.commit_floats()
            if not self.cont:
                self.page_no = max(0, min(len(self.doc) - 1, cur[0]))
                if self.fit_mode:
                    self._apply_fit(render=False)
            self._cont_rel = 0.0
            self._cont_doc = None
            self.render()
            if not self.cont:
                self.canvas.yview_moveto(0)
            self.update_ui()
        self.status.config(text=_tr("Continuous scrolling on.") if self.cont else _tr("One page at a time."))

    def toggle_fullscreen(self, presentation=False):
        if self.presenting or self._fs:
            self.exit_fullscreen()
            return
        if self.doc is None:
            self.status.config(text=_tr("Open a PDF first."))
            return
        self.commit_floats()
        body_slaves = [w for w in self.pack_slaves() if w is not self.body_frame]
        order = list(self.pack_slaves())
        before_body = {w: (order.index(w) < order.index(self.body_frame)) for w in body_slaves}
        rec = []
        for w in body_slaves:
            rec.append((w, w.pack_info()))
            w.pack_forget()
        inner = [w for w in self.body_frame.pack_slaves() if w is not self.canvas.master]
        rec_in = []
        for w in inner:
            rec_in.append((w, w.pack_info()))
            w.pack_forget()
        bars = []
        if presentation:
            for w in self.canvas.master.pack_slaves():
                if w is not self.canvas:
                    bars.append((w, w.pack_info()))
                    w.pack_forget()
        self._fs = dict(rec=rec, rec_in=rec_in, bars=bars, cont=self.cont, fit=self.fit_mode, zoom=self.zoom,
                        tool=self.tool.get(), bg=self.canvas.cget("bg"), before=before_body, geo=self.geometry(),
                        state=self.state())
        self.presenting = bool(presentation)
        self.attributes("-fullscreen", True)
        if presentation:
            self.cont = False
            self.set_tool("view")
            self.canvas.config(bg="black")
            self.fit_mode = "page"
            self._pres_hide_cursor_later()
        self.update_idletasks()
        self.after(150, self._fs_refit)
        self.status.config(text="")

    def _fs_refit(self):
        if self._fs and self.doc is not None:
            if self.fit_mode:
                self._apply_fit(render=True)
            else:
                self.render()

    def exit_fullscreen(self):
        fs = self._fs
        if not fs:
            return
        self._fs = None
        was_pres = self.presenting
        self.presenting = False
        self.attributes("-fullscreen", False)
        body = self.body_frame
        for w, info in fs["rec"]:
            info = {k: v for k, v in info.items() if k != "in"}
            if fs["before"].get(w, True):
                w.pack(before=body, **info)
            else:
                w.pack(**info)
        # the sidebar parts go back in front of the page area
        for w, info in fs["rec_in"]:
            info = {k: v for k, v in info.items() if k != "in"}
            w.pack(before=self.canvas.master, **info)
        # restore stacking of the bars next to the canvas (scrollbars must be packed before the canvas)
        if fs["bars"]:
            self.canvas.pack_forget()
            for w, _ in fs["bars"]:
                w.pack_forget()
            for w, info in fs["bars"]:
                w.pack(**{k: v for k, v in info.items() if k != "in"})
            self.canvas.pack(fill="both", expand=True)
        self.canvas.config(bg=fs["bg"], cursor="")
        if self._cursor_after:
            self.after_cancel(self._cursor_after)
            self._cursor_after = None
        self.cont = fs["cont"]
        self.fit_mode = fs["fit"]
        if not was_pres:
            self.set_tool(fs["tool"])
        self.after(100, lambda: (self._apply_fit(render=True) if (self.fit_mode and self.doc) else self.render()))

    def _pres_hide_cursor_later(self):
        if self._cursor_after:
            self.after_cancel(self._cursor_after)
        self.canvas.config(cursor="")
        if self.presenting:
            self._cursor_after = self.after(2500, lambda: self.canvas.config(cursor="none") if self.presenting else None)

    def _pres_click(self, step):
        self.goto(self.page_no + step)

    def _pres_key(self, dx, dy):
        step = 1 if (dx > 0 or dy > 0) else -1
        self.goto(self.page_no + step)

    # ===================================================================== text selection
    def _page_words(self, i):
        key = (id(self.doc), i, self._gen)
        w = self._wcache.get(key)
        if w is None:
            if len(self._wcache) > 8:
                self._wcache.clear()
            pg = self.doc[i]
            m = pg.rotation_matrix
            w = []
            for t in pg.get_text("words"):
                r0 = fitz.Rect(t[:4])
                w.append((r0, r0 * m, t[4], (t[5], t[6])))
            self._wcache[key] = w
        return w

    def _event_page_pt(self, e):
        """-> (page index, x, y) with x / y in view points of that page"""
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        if self.cont and self._cont_tops:
            i = self._cont_page_at_y(cy)
            ox, oy = self._cont_x[i], self._cont_tops[i]
        else:
            i, ox, oy = self.page_no, self.off_x, self.off_y
        return i, (cx - ox) / self.zoom, (cy - oy) / self.zoom

    def _text_press(self, e):
        """Start a text selection if the pointer is on text.  Returns True when it did."""
        self._tsel_drag = None
        self._tsel_moved = False
        i, x, y = self._event_page_pt(e)
        if not 0 <= i < len(self.doc):
            return False
        words = self._page_words(i)
        k = nearest_word(words, x, y, max_dist=4)
        if self._tsel:
            self._tsel = None
            self.canvas.delete("tsel")
        if k is None:
            return False
        self._tsel_drag = dict(page=i, anchor=k, start=(e.x, e.y))
        return True

    def _text_drag(self, e):
        d = self._tsel_drag
        if not d:
            return
        if not self._tsel_moved and math.hypot(e.x - d["start"][0], e.y - d["start"][1]) < 4:
            return
        self._tsel_moved = True
        i = d["page"]
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        ox, oy = self._page_origin(i)
        k = nearest_word(self._page_words(i), (cx - ox) / self.zoom, (cy - oy) / self.zoom)
        if k is None:
            return
        self._tsel = dict(page=i, i0=d["anchor"], i1=k)
        self._draw_textsel(i)

    def _text_release(self, e):
        """True when the release ended a text selection (so it is not also a click)."""
        d, moved = self._tsel_drag, self._tsel_moved
        self._tsel_drag = None
        self._tsel_moved = False
        self._view_click = self._view_click if not moved else None
        return bool(d and moved)

    def _text_dblclick(self, e, level):
        if self.doc is None or self.tool.get() != "view" or self.presenting:
            return
        if self.cont:
            self._cont_focus_at(e)
        i, x, y = self._event_page_pt(e)
        words = self._page_words(i)
        k = nearest_word(words, x, y, max_dist=4)
        if k is None:
            return
        i0 = i1 = k
        if level == 3:
            key = words[k][3]
            while i0 > 0 and words[i0 - 1][3] == key:
                i0 -= 1
            while i1 < len(words) - 1 and words[i1 + 1][3] == key:
                i1 += 1
        self._tsel = dict(page=i, i0=i0, i1=i1)
        self._draw_textsel(i)
        self._tsel_drag = None
        self._view_click = None

    def _draw_textsel(self, i=None):
        if i is None:
            self.canvas.delete("tsel")
        else:
            self.canvas.delete(f"tselp{i}")
        s = self._tsel
        if not s or self.doc is None:
            return
        if i is not None and s["page"] != i:
            return
        if not (self.cont or s["page"] == self.page_no):
            return
        words = self._page_words(s["page"])
        ox, oy = self._page_origin(s["page"])
        z = self.zoom
        if i is None or not hasattr(self, "_tsel_imgs"):
            self._tsel_imgs = {}
        imgs = []
        for _k, _r0, rv in selection_lines(words, s["i0"], s["i1"]):
            w, h = max(1, int(rv.width * z)), max(1, int(rv.height * z))
            if w * h > 4_000_000:
                continue
            ph = ImageTk.PhotoImage(Image.new("RGBA", (w, h), self._sel_rgba()))
            imgs.append(ph)
            self.canvas.create_image(ox + rv.x0 * z, oy + rv.y0 * z, anchor="nw", image=ph,
                                     tags=("tsel", f"tselp{s['page']}", "pg"))
        self._tsel_imgs[s["page"]] = imgs

    @staticmethod
    def _sel_rgba():
        h = TH.ACCENT.lstrip("#")
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), 85)

    def _sel_text(self):
        s = self._tsel
        if not s or self.doc is None:
            return ""
        return selection_text(self._page_words(s["page"]), s["i0"], s["i1"])

    def copy_selection(self):
        t = self._sel_text()
        if not t:
            self.status.config(text=_tr("Nothing selected. Drag over text with the Hand tool to select it."))
            return False
        self.clipboard_clear()
        self.clipboard_append(t)
        self.update_idletasks()
        self.status.config(text=_tr("Copied {} characters.").replace("{}", str(len(t))))
        return True

    def _key_copy(self, _e=None):
        if self._typing():
            return None
        if self.doc is not None:
            self.copy_selection()
            return "break"

    def select_all_text(self):
        if self.doc is None:
            return
        i = self.page_no
        words = self._page_words(i)
        if not words:
            self.status.config(text=_tr("This page has no selectable text (scanned pages need OCR first)."))
            return
        self._tsel = dict(page=i, i0=0, i1=len(words) - 1)
        self._draw_textsel(i)

    def _sel_markup(self, kind):
        s = self._tsel
        if not s:
            return
        i = s["page"]
        lines = selection_lines(self._page_words(i), s["i0"], s["i1"])
        quads = [r0.quad for _k, r0, _rv in lines]
        col = getattr(self, "hl_color", (1, 0.9, 0.2)) if kind == "highlight" else self.color
        add = {"highlight": "add_highlight_annot", "underline": "add_underline_annot", "strike": "add_strikeout_annot"}[kind]
        pg = self.doc[i]

        def work():
            a = getattr(pg, add)(quads=quads)
            if a is not None:
                a.set_colors(stroke=col)
                a.update()
        self._tsel = None
        self._run_edit(work, thumbs=[i])

    def _sel_redact(self):
        s = self._tsel
        if not s:
            return
        i = s["page"]
        words = self._page_words(i)
        a, b = sorted((s["i0"], s["i1"]))
        hits = [(i, w[0]) for w in words[a:b + 1]]
        if messagebox.askyesno(APP, _tr("Permanently remove the selected text? (This can be undone until you save.)")):
            self._tsel = None
            self.do_redact_hits(hits, [w[2] for w in words[a:b + 1]])

    def _text_menu(self, e):
        s = self._tsel
        if not s or self.doc is None or self.tool.get() != "view":
            return False
        i, x, y = self._event_page_pt(e)
        if i != s["page"]:
            return False
        words = self._page_words(i)
        k = nearest_word(words, x, y, max_dist=4)
        a, b = sorted((s["i0"], s["i1"]))
        if k is None or not a <= k <= b:
            return False
        m = tk.Menu(self, tearoff=0)
        m.add_command(label="Copy", accelerator="Ctrl+C", command=self.copy_selection)
        m.add_separator()
        m.add_command(label="Highlight", command=lambda: self._sel_markup("highlight"))
        m.add_command(label="Underline", command=lambda: self._sel_markup("underline"))
        m.add_command(label="Strikethrough", command=lambda: self._sel_markup("strike"))
        m.add_separator()
        m.add_command(label="Redact selected text...", command=self._sel_redact)
        m.add_separator()
        m.add_command(label="Look up / translate...", command=lambda: self.dlg_lookup(self._sel_text()))
        m.tk_popup(e.x_root, e.y_root)
        return True

    def _on_motion(self, e):
        if self.doc is None or self.tool.get() != "view":
            return
        if self.presenting:
            self._pres_hide_cursor_later()
            return
        now = time.monotonic()
        if now - self._motion_t < 0.05:
            return
        self._motion_t = now
        try:
            i, x, y = self._event_page_pt(e)
            over = 0 <= i < len(self.doc) and nearest_word(self._page_words(i), x, y, max_dist=1) is not None
        except Exception:
            over = False
        want = "xterm" if over else "fleur"
        if self.canvas.cget("cursor") != want:
            self.canvas.config(cursor=want)

    # ===================================================================== rendering
    def _photo_for(self, pg, zi):
        pix = pg.get_pixmap(matrix=fitz.Matrix(zi, zi), alpha=False)
        if self.night:
            pix.invert_irect(pix.irect)
        return tk.PhotoImage(data=pix.tobytes("ppm")), pix.width, pix.height

    def _eff_zoom(self, w, h):
        return min(self.zoom, (MAX_PIXELS / max(1.0, w * h)) ** 0.5)

    def _cont_layout(self):
        n = len(self.doc)
        key = (id(self.doc), n)
        if self._cont_raw is None or self._cont_key != key:
            self._cont_raw = [(self.doc[i].rect.width, self.doc[i].rect.height) for i in range(n)]
            self._cont_key = key
        zmax = min(((MAX_PIXELS / max(1.0, w * h)) ** 0.5 for w, h in self._cont_raw), default=self.zoom)
        if self.zoom > zmax + 1e-6:                     # a very large page: never build a gigantic bitmap
            self.zoom = max(0.02, zmax)
            self._sync_zoom = True
            self.zoom_var.set(self.zoom * 100)
            self._sync_zoom = False
            self.zoom_lbl.config(text=f"{int(self.zoom * 100)}%")
            self.status.config(text=_tr("Zoom limited because this page is very large."))
        sizes, tops, y = [], [], PAD
        for w, h in self._cont_raw:
            zi = self._eff_zoom(w, h)
            sw, sh = int(math.ceil(w * zi)), int(math.ceil(h * zi))
            sizes.append((sw, sh))
            tops.append(y)
            y += sh + GAP
        total = y - GAP + PAD
        cw = max(self.canvas.winfo_width(), 1)
        ch = max(self.canvas.winfo_height(), 1)
        maxw = max((s[0] for s in sizes), default=0)
        width = max(maxw + 2 * PAD, cw)
        self._cont_x = [max(PAD, (width - s[0]) // 2) for s in sizes]
        self._cont_sizes, self._cont_tops, self._cont_total = sizes, tops, max(total, ch)
        self.canvas.config(scrollregion=(0, 0, width, self._cont_total))

    def _render_cont(self):
        c = self.canvas
        if self._cont_tops and self._cont_doc is self.doc and len(self._cont_tops) == len(self.doc):
            pn, rel = self._current_rel()
        else:
            pn, rel = self.page_no, self._cont_rel
        self._gen += 1
        self._cont_photos.clear()
        self._cont_drawn = set()
        c.delete("all")
        self._cont_layout()
        self._cont_doc = self.doc
        pn = max(0, min(len(self.doc) - 1, pn))
        fill = "#1c1c1e" if self.night else "white"
        for i, (sw, sh) in enumerate(self._cont_sizes):
            x, y = self._cont_x[i], self._cont_tops[i]
            c.create_rectangle(x - 1, y - 1, x + sw + 1, y + sh + 1, fill=fill, outline=TH.LINE, tags=("pg", f"cp{i}"))
        y = self._cont_tops[pn] + rel * self._cont_sizes[pn][1]
        c.update_idletasks()
        c.yview_moveto(max(0.0, min(1.0, y / max(1, self._cont_total))))
        pn = max(0, min(len(self.doc) - 1, pn))
        self._cont_focus(pn)
        self._cont_update(force_current=pn)

    def _cont_schedule(self):
        if self._cont_upd_after is None:
            self._cont_upd_after = self.after(15, self._cont_update)

    def _cont_update(self, force_current=None):
        self._cont_upd_after = None
        if not (self.cont and self.doc is not None and self._cont_tops and self._cont_doc is self.doc):
            return
        c = self.canvas
        y0 = c.canvasy(0)
        h = max(c.winfo_height(), 1)
        y1 = y0 + h
        lo = self._cont_page_at_y(y0)
        hi = self._cont_page_at_y(y1)
        vis = [i for i in range(lo, hi + 1) if self._cont_tops[i] + self._cont_sizes[i][1] >= y0 and self._cont_tops[i] <= y1]
        if not vis:
            vis = [lo]
        # page with the biggest visible part is "current"
        best, best_a = vis[0], -1
        for i in vis:
            a = min(y1, self._cont_tops[i] + self._cont_sizes[i][1]) - max(y0, self._cont_tops[i])
            if a > best_a:
                best, best_a = i, a
        if force_current is not None and force_current in vis:
            best = force_current
        keep = set(vis)
        for i in sorted(self._cont_drawn - keep):
            c.delete(f"ci{i}")
            c.delete(f"ov{i}")
        self._cont_drawn &= keep
        redraw_needed = False
        for i in vis:
            if i in self._cont_drawn:
                continue
            pg = self.doc[i]
            r = pg.rect
            raw = self._cont_raw[i]
            if abs(r.width - raw[0]) > 0.5 or abs(r.height - raw[1]) > 0.5:
                self._cont_raw = None                         # page size changed (rotation, crop): lay out again
                redraw_needed = True
                break
            zi = self._eff_zoom(raw[0], raw[1])
            ph = self._cont_photos.get(i)
            if ph is None:
                ph, _w, _h = self._photo_for(pg, zi)
                self._cont_photos[i] = ph
                while len(self._cont_photos) > PHOTO_CACHE:
                    self._cont_photos.popitem(last=False)
            else:
                self._cont_photos.move_to_end(i)
            x, y = self._cont_x[i], self._cont_tops[i]
            c.create_image(x, y, anchor="nw", image=ph, tags=("pg", f"ci{i}"))
            self._cont_drawn.add(i)
            self._cont_overlays(i)
        if redraw_needed:
            self.render()
            return
        if best != self.page_no or (force_current is not None):
            old = self.page_no
            self._cont_focus(best)
            if best != old:
                if self.sel_pages <= {old}:
                    self.sel_pages = {best}
                self._select_thumb()
                self.update_ui()

    def _cont_overlays(self, i):
        """find hits, form field boxes, photos and the text selection of page i"""
        keep_no = self.page_no
        before = set(self.canvas.find_all())
        self._cont_focus(i)
        try:
            self._after_render()
            if self.floats and self.float_page == i:
                self.draw_floats()
            self.tf_draw()
            self._draw_textsel(i)
        finally:
            self._cont_focus(keep_no)
        for it in set(self.canvas.find_all()) - before:
            self.canvas.addtag_withtag(f"ov{i}", it)

    def _on_yscroll(self, a, b):
        self._vs.set(a, b)
        if self.cont and self.doc is not None:
            self._cont_schedule()

    def _cont_goto(self, n):
        self.commit_floats()
        if not self._cont_tops or self._cont_doc is not self.doc:
            self.page_no = n
            self.sel_pages = {n}
            self._cont_rel = 0.0
            self.render()
        else:
            self.sel_pages = {n}
            y = max(0, self._cont_tops[n] - PAD // 2)
            self.canvas.yview_moveto(y / max(1, self._cont_total))
            self._cont_focus(n)
            self._cont_update(force_current=n)
        self._select_thumb()
        self.update_ui()

    # ===================================================================== tabs
    def _build_tabbar(self):
        self.tabbar = tk.Frame(self, bg=TH.NAV_BG)
        self.tabbar_line = tk.Frame(self, bg=TH.LINE, height=1)

    def _tabs_show(self):
        if self.tabbar is None:
            return
        if self.tabs and not self.tabbar.winfo_ismapped() and not self._fs:
            self.tabbar.pack(side="top", fill="x", before=self.body_frame)
            self.tabbar_line.pack(side="top", fill="x", before=self.body_frame)
        elif not self.tabs and self.tabbar.winfo_ismapped():
            self.tabbar.pack_forget()
            self.tabbar_line.pack_forget()

    def _tab_title(self, i):
        st = self.tabs[i]
        if i == self.cur_tab:
            path, title = self.path, st.get("title")
        else:
            path, title = st.get("path"), st.get("title")
        return os.path.basename(path) if path else _tr(title or "Untitled")

    def _tab_modified(self, i):
        return self.modified if i == self.cur_tab else bool(self.tabs[i].get("modified"))

    def _tabs_redraw(self, force=False):
        if self.tabbar is None:
            return
        sig = tuple((self._tab_title(i), self._tab_modified(i)) for i in range(len(self.tabs))) + (self.cur_tab,)
        if sig == self._tab_sig and not force:
            return
        self._tab_sig = sig
        for w in self.tabbar.winfo_children():
            w.destroy()
        self._tabs_show()
        for i in range(len(self.tabs)):
            cur = i == self.cur_tab
            bg = TH.CARD if cur else TH.NAV_BG
            fr = tk.Frame(self.tabbar, bg=bg, cursor="hand2")
            fr.pack(side="left", padx=(6 if i == 0 else 1, 0), pady=(S(4), 0))
            name = self._tab_title(i)
            if len(name) > 26:
                name = name[:24] + "…"
            lab = tk.Label(fr, text=("● " if self._tab_modified(i) else "") + name, bg=bg,
                           fg=TH.LABEL if cur else TH.SECOND, font=(UI_FONT, 9, "bold" if cur else "normal"), padx=10, pady=4)
            lab.pack(side="left")
            x = tk.Label(fr, text="✕", bg=bg, fg=TH.TERTIARY, font=(UI_FONT, 8), padx=5, cursor="hand2")
            x.pack(side="left", padx=(0, 4))
            x.bind("<Enter>", lambda e, w=x: w.config(fg="#d93025"))
            x.bind("<Leave>", lambda e, w=x: w.config(fg=TH.TERTIARY))
            x.bind("<Button-1>", lambda e, k=i: self.close_tab(k))
            for w in (fr, lab):
                w.bind("<Button-1>", lambda e, k=i: self.switch_tab(k))
                w.bind("<Button-2>", lambda e, k=i: self.close_tab(k))
            tk.Frame(fr, bg=TH.ACCENT if cur else bg, height=2).pack(side="bottom", fill="x")
        plus = tk.Label(self.tabbar, text="＋", bg=TH.NAV_BG, fg=TH.ACCENT, font=(UI_FONT, 12), padx=10, cursor="hand2")
        plus.pack(side="left", pady=(S(4), 0))
        plus.bind("<Button-1>", lambda e: self.open_dialog())

    def _tab_stash(self):
        if 0 <= self.cur_tab < len(self.tabs):
            st = self.tabs[self.cur_tab]
            for k in TAB_KEYS:
                st[k] = getattr(self, k, None)
            st["rel"] = self._current_rel()

    def _tab_fresh_state(self):
        self.doc = None
        self.path = None
        self.page_no = 0
        self.undo_stack = []
        self.redo_stack = []
        self.floats = []
        self.sel = None
        self.float_page = 0
        self.find_hits = []
        self.find_idx = -1
        self.sel_pages = {0}
        self._anchor = 0
        self.protect = None
        self.modified = False
        self._sig_orig = None
        self._tsel = None
        self.fdrag = None

    def _tab_load(self, i):
        st = self.tabs[i]
        self.cur_tab = i
        for k in TAB_KEYS:
            if k in st:
                setattr(self, k, st[k])
        self.fdrag = None
        self._tsel = None
        self._cont_doc = None
        rel = st.get("rel") or (self.page_no, 0.0)
        self._cont_rel = rel[1] if self.cont else 0.0
        self.set_zoom(self.zoom, render=False, mode=self.fit_mode)
        self.build_thumbs()
        self.render()
        if not self.cont:
            try:
                self.canvas.yview_moveto(rel[1])
            except Exception:
                pass
        self.update_ui()
        if self.sidebar_mode == "bookmarks":
            self._refresh_bookmarks()
        elif self.sidebar_mode == "comments":
            self._refresh_comments()

    def switch_tab(self, i):
        if i == self.cur_tab or not 0 <= i < len(self.tabs):
            return
        self.commit_floats()
        self._tab_stash()
        self._tab_load(i)
        self._tabs_redraw()

    def next_tab(self, step=1):
        if len(self.tabs) > 1:
            self.switch_tab((self.cur_tab + step) % len(self.tabs))
        return "break"

    def _adopt_doc(self, doc, path=None, modified=False):
        """Open `doc` in a new tab (or in the empty window)."""
        from claudius_extra import ExtraTools
        if self.doc is not None:
            self.commit_floats()
            self._tab_stash()
            self.tabs.append({})
            self.cur_tab = len(self.tabs) - 1
            self._tab_fresh_state()
        elif not self.tabs:
            self.tabs.append({})
            self.cur_tab = 0
        self.tabs[self.cur_tab]["title"] = self._next_title
        self._next_title = None
        self._cont_doc = None
        self._cont_rel = 0.0
        ExtraTools._adopt_doc(self, doc, path, modified)
        self._tabs_redraw(force=True)

    def _open_result(self, doc, title):
        """A document made by Claudius PDF Editor (comparison, scan, ...) opens in its own tab."""
        self._next_title = title
        self._adopt_doc(doc, None, modified=True)

    def _tab_for_path(self, p):
        key = os.path.normcase(os.path.abspath(p))
        for i in range(len(self.tabs)):
            tp = self.path if i == self.cur_tab else self.tabs[i].get("path")
            if tp and os.path.normcase(os.path.abspath(tp)) == key:
                return i
        return -1

    def _ask_save_current(self):
        if self.doc is not None and self.modified:
            a = messagebox.askyesnocancel(APP, _tr("Save changes to {} before closing?").replace("{}", self._tab_title(self.cur_tab)))
            if a is None:
                return False
            if a:
                self.save()
                return not self.modified
        return True

    def close_tab(self, i=None):
        if not self.tabs:
            return
        i = self.cur_tab if i is None else i
        if not 0 <= i < len(self.tabs):
            return
        if i != self.cur_tab:
            self.switch_tab(i)
        self.commit_floats()
        if not self._ask_save_current():
            return
        del self.tabs[i]
        if not self.tabs:
            self.cur_tab = -1
            self._tab_fresh_state()
            self._cont_doc = None
            self._cont_tops = []
            self.build_thumbs()
            self.render()
            self.update_ui()
            self.status.config(text=_tr("Open a PDF to start (Ctrl+O)"))
        else:
            j = min(i, len(self.tabs) - 1)
            self._tab_load(j)
        self._tabs_redraw(force=True)
        self._tabs_show()

    def confirm_discard(self):
        """Ask about unsaved changes in every open tab.  False = the user cancelled."""
        if not self.tabs:
            return True
        self.commit_floats()
        start = self.cur_tab
        for i in range(len(self.tabs)):
            mod = self.modified if i == self.cur_tab else bool(self.tabs[i].get("modified"))
            if not mod:
                continue
            self.switch_tab(i)
            if not self._ask_save_current():
                return False
        return True

    def open_paths_list(self):
        """paths of all open files, current first (used when the program restarts)"""
        self._tab_stash()
        out = []
        for i, st in enumerate(self.tabs):
            p = st.get("path")
            if p and os.path.exists(p):
                out.insert(0, p) if i == self.cur_tab else out.append(p)
        return out
