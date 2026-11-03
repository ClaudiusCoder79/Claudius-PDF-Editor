"""
Claudius PDF Editor - CAD tools for drawings and plans.

  * Scale          set the scale of a drawing (1:100 ... or by clicking two points of a known length)
  * Measure        distance, length of a path, angle, area and perimeter - in real units
  * Draw           polyline, polygon, 3-point arc, exact line (type length and angle), dimension line
  * Snap           to line ends / middles / corners of the drawing, to a grid and to 15 degree steps

All points are kept in "view points" (the page as you see it, 1 unit = 1 PDF point), like the other drawing tools.
"""
import datetime
import math
import tkinter as tk
from tkinter import ttk, messagebox

import pymupdf as fitz

import claudius_theme as TH
from claudius_i18n import tr as _tr
from claudius_extra import APP, UI_FONT, BG, ACCENT, S
from claudius_view import get_setting, set_setting

CAD_TOOLS = ("c_dist", "c_len", "c_angle", "c_area", "c_poly", "c_polygon", "c_arc", "c_dim", "c_xline", "c_cal")

CAD_HINTS = {
    "c_dist": "Distance: click two points. Snap helps you hit line ends and corners. Esc = cancel.",
    "c_len": "Length of a path: click the points one after another; double-click, Enter or right-click finishes.",
    "c_angle": "Angle: click a point on the first arm, then the corner, then a point on the second arm.",
    "c_area": "Area: click the corners of the area; double-click, Enter or right-click finishes (area and perimeter are shown).",
    "c_poly": "Polyline: click the points; double-click, Enter or right-click finishes and draws it on the page.",
    "c_polygon": "Polygon: click the corners; double-click, Enter or right-click closes and draws it (Fill in this tab for a solid one).",
    "c_arc": "Arc: click the start, a point on the arc, then the end.",
    "c_dim": "Dimension: click the two points to measure, then click where the dimension line should run.",
    "c_xline": "Exact line: click the start point, then type the length and angle.",
    "c_cal": "Calibrate: click two points whose real distance you know, then type that distance.",
}

# one unit in millimetres
UNITS = {"mm": 1.0, "cm": 10.0, "m": 1000.0, "in": 25.4, "ft": 304.8}
PAPER_MM_PER_PT = 25.4 / 72.0
RATIOS = ["1:1", "1:2", "1:5", "1:10", "1:20", "1:25", "1:50", "1:75", "1:100", "1:125", "1:150", "1:200", "1:250",
          "1:500", "1:1000", "2:1", "5:1", "10:1"]
MARK_COLOR = "#d6336c"
SNAP_PIXELS = 12
CELL = 24.0


def fmt_num(v):
    a = abs(v)
    if a >= 1000:
        return "%.1f" % v
    if a >= 100:
        return "%.1f" % v
    if a >= 1:
        return "%.2f" % v
    return "%.3f" % v


def poly_area(pts):
    s = 0.0
    for (x0, y0), (x1, y1) in zip(pts, pts[1:] + pts[:1]):
        s += x0 * y1 - x1 * y0
    return abs(s) / 2.0


def path_length(pts, closed=False):
    seq = pts + (pts[:1] if closed else [])
    return sum(math.dist(a, b) for a, b in zip(seq, seq[1:]))


def circle3(p1, p2, p3):
    """centre and radius of the circle through three points, or None when they lie on a line"""
    ax, ay = p1
    bx, by = p2
    cx, cy = p3
    d = 2 * (ax * (by - cy) + bx * (cy - ay) + cx * (ay - by))
    if abs(d) < 1e-9:
        return None
    ux = ((ax * ax + ay * ay) * (by - cy) + (bx * bx + by * by) * (cy - ay) + (cx * cx + cy * cy) * (ay - by)) / d
    uy = ((ax * ax + ay * ay) * (cx - bx) + (bx * bx + by * by) * (ax - cx) + (cx * cx + cy * cy) * (bx - ax)) / d
    return (ux, uy), math.hypot(ax - ux, ay - uy)


def arc_points(p1, p2, p3, step_deg=3.0):
    """points of the arc that starts at p1, passes p2 and ends at p3 (a straight line when the three points are aligned)"""
    c = circle3(p1, p2, p3)
    if c is None:
        return [p1, p3]
    (cx, cy), r = c
    a1 = math.atan2(p1[1] - cy, p1[0] - cx)
    a2 = math.atan2(p2[1] - cy, p2[0] - cx)
    a3 = math.atan2(p3[1] - cy, p3[0] - cx)

    def ccw(a, b):
        return (b - a) % (2 * math.pi)
    if ccw(a1, a2) <= ccw(a1, a3):                  # p2 lies on the way a1 -> a3 going counter-clockwise (in angle terms)
        sweep = ccw(a1, a3)
    else:
        sweep = -((a1 - a3) % (2 * math.pi))
    n = max(2, int(abs(math.degrees(sweep)) / step_deg))
    return [(cx + r * math.cos(a1 + sweep * k / n), cy + r * math.sin(a1 + sweep * k / n)) for k in range(n + 1)]


def dimension_geometry(a, b, off_pt, lw):
    """lines and arrow heads of an aligned dimension.  Returns (items, text_anchor, direction) in view space;
    items use the same types as the shape tools ('line', 'head')."""
    ax, ay = a
    bx, by = b
    dx, dy = bx - ax, by - ay
    L = math.hypot(dx, dy)
    if L < 1e-6:
        return [], a, (1.0, 0.0)
    ux, uy = dx / L, dy / L
    nx, ny = -uy, ux
    d = (off_pt[0] - ax) * nx + (off_pt[1] - ay) * ny              # signed distance of the dimension line from the measured line
    a2 = (ax + nx * d, ay + ny * d)
    b2 = (bx + nx * d, by + ny * d)
    gap, over = 2.0, 3.0
    sgn = 1 if d >= 0 else -1
    items = []
    for p, q in ((a, a2), (b, b2)):                                  # extension lines
        items.append(("line", [(p[0] + nx * sgn * gap, p[1] + ny * sgn * gap), (q[0] + nx * sgn * over, q[1] + ny * sgn * over)], False))
    hl = min(max(7.0, lw * 4.0), L * 0.3)
    hw = hl * 0.3
    items.append(("line", [(a2[0] + ux * hl * 0.9, a2[1] + uy * hl * 0.9), (b2[0] - ux * hl * 0.9, b2[1] - uy * hl * 0.9)], False))
    for tip, s in ((a2, 1), (b2, -1)):
        base = (tip[0] + ux * hl * s, tip[1] + uy * hl * s)
        items.append(("head", [tip, (base[0] + nx * hw, base[1] + ny * hw), (base[0] - nx * hw, base[1] - ny * hw)], False))
    mid = ((a2[0] + b2[0]) / 2, (a2[1] + b2[1]) / 2)
    return items, mid, (ux, uy)


class CadTools:
    """mixin of PdfEditor"""

    # ------------------------------------------------------------------ settings and scale
    def _cad_vars(self):
        if not hasattr(self, "_cadv"):
            self._cadv = {"geo": tk.BooleanVar(value=True), "grid": tk.BooleanVar(value=False),
                          "ang": tk.BooleanVar(value=False), "gridsize": tk.StringVar(value="1")}
            self._cad = {"tool": None, "pts": [], "page": None}
            self._cad_marks = []
            self._cad_geo_key = None
            self._cad_geo_cells = {}
        return self._cadv

    def cad_scale(self):
        """-> dict(per_pt, unit, ratio): real units per PDF point of this document"""
        try:
            allv = get_setting("cad_scales", {}) or {}
            d = allv.get(getattr(self, "path", None) or "")
            if isinstance(d, dict) and float(d.get("per_pt", 0)) > 0 and d.get("unit") in UNITS:
                return {"per_pt": float(d["per_pt"]), "unit": d["unit"], "ratio": str(d.get("ratio", ""))}
        except Exception:
            pass
        return {"per_pt": PAPER_MM_PER_PT, "unit": "mm", "ratio": "1:1"}

    def cad_set_scale(self, per_pt, unit, ratio=""):
        try:
            allv = dict(get_setting("cad_scales", {}) or {})
            allv[getattr(self, "path", None) or ""] = {"per_pt": float(per_pt), "unit": unit, "ratio": ratio}
            if len(allv) > 200:                                    # keep the file small
                allv = dict(list(allv.items())[-200:])
            set_setting("cad_scales", allv)
        except Exception:
            pass
        self.cad_scale_refresh()
        self.cad_draw()

    def cad_scale_refresh(self):
        lbl = getattr(self, "cad_scale_lbl", None)
        if lbl is None:
            return
        sc = self.cad_scale()
        txt = _tr("Scale") + " " + (sc["ratio"] or ("1 pt = " + fmt_num(sc["per_pt"]) + " " + sc["unit"])) + "  (" + sc["unit"] + ")"
        try:
            lbl.config(text=txt)
        except tk.TclError:
            pass

    def _cad_len(self, d_pt):
        sc = self.cad_scale()
        return "%s %s" % (fmt_num(d_pt * sc["per_pt"]), sc["unit"])

    def _cad_area(self, a_pt):
        sc = self.cad_scale()
        return "%s %s²" % (fmt_num(a_pt * sc["per_pt"] ** 2), sc["unit"])

    def cad_scale_dialog(self):
        if not self.need_doc():
            return
        self._cad_vars()
        sc = self.cad_scale()
        d = self._dlg(_tr("Drawing scale"))
        self._label(d, _tr("Paper scale"), bold=True).grid(row=0, column=0, sticky="w", pady=3)
        rv = tk.StringVar(value=sc["ratio"] if sc["ratio"] in RATIOS else "1:100")
        ttk.Combobox(d, textvariable=rv, values=RATIOS, width=10).grid(row=0, column=1, sticky="w")
        self._label(d, _tr("Show results in"), bold=True).grid(row=1, column=0, sticky="w", pady=3)
        uv = tk.StringVar(value=sc["unit"])
        ttk.Combobox(d, textvariable=uv, values=list(UNITS), width=6, state="readonly").grid(row=1, column=1, sticky="w")
        self._label(d, _tr("1:100 means 1 mm on the printed page is 100 mm in reality. If the drawing has no scale, use "
                           "\"Calibrate\" and click two points whose real distance you know."), fg=TH.TERTIARY).grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(8, 0))

        def ok():
            txt = rv.get().strip().replace(" ", "")
            try:
                a, b = txt.split(":")
                a, b = float(a.replace(",", ".")), float(b.replace(",", "."))
                if a <= 0 or b <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror(APP, _tr("Type the scale like 1:100."), parent=d)
                return
            unit = uv.get() if uv.get() in UNITS else "mm"
            per_pt = PAPER_MM_PER_PT * (b / a) / UNITS[unit]
            d.destroy()
            self.cad_set_scale(per_pt, unit, "%g:%g" % (a, b))
            self.status.config(text=_tr("Scale set."))

        def cal():
            d.destroy()
            self.set_tool("c_cal")
        f = self._buttons(d, _tr("OK"), ok, 3, 3)
        ttk.Button(f, text=_tr("Calibrate with two points..."), command=cal).pack(side="left", padx=4)

    def _cad_calibrate(self, p1, p2):
        dist = math.dist(p1, p2)
        if dist < 1:
            return
        d = self._dlg(_tr("Calibrate"))
        sc = self.cad_scale()
        self._label(d, _tr("Real distance between the two points"), bold=True).grid(row=0, column=0, columnspan=3, sticky="w")
        vv = tk.StringVar(value="")
        e = ttk.Entry(d, textvariable=vv, width=12)
        e.grid(row=1, column=0, sticky="w", pady=6)
        e.focus_set()
        uv = tk.StringVar(value=sc["unit"])
        ttk.Combobox(d, textvariable=uv, values=list(UNITS), width=6, state="readonly").grid(row=1, column=1, sticky="w", padx=6)

        def ok(_e=None):
            try:
                val = float(vv.get().replace(",", "."))
                if val <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror(APP, _tr("Type a number greater than zero."), parent=d)
                return
            unit = uv.get() if uv.get() in UNITS else "mm"
            d.destroy()
            self.cad_set_scale(val / dist, unit, "")
            self.set_tool("view")
            self.status.config(text=_tr("Scale set from the two points."))
        d.bind("<Return>", ok)
        self._buttons(d, _tr("OK"), ok, 2, 3)

    # ------------------------------------------------------------------ ribbon
    def build_cad_ribbon(self, r):
        v = self._cad_vars()
        bg = r.cget("bg")
        g = self.group(r, "Scale")
        self.rb(g, "⚙", "Set scale", self.cad_scale_dialog, size=19, width=62)
        self.rb(g, "↔", "Calibrate", None, tool="c_cal", size=19, width=62)
        self.cad_scale_lbl = tk.Label(g, text="", bg=bg, fg=TH.SECOND, font=(UI_FONT, 8), wraplength=S(110), justify="left")
        self.cad_scale_lbl.pack(side="left", padx=4)
        g = self.group(r, "Measure")
        self.rb(g, "↔", "Distance", None, tool="c_dist", size=19, width=62)
        self.rb(g, "〰", "Length", None, tool="c_len", size=19, width=62)
        self.rb(g, "∠", "Angle", None, tool="c_angle", size=19, width=62)
        self.rb(g, "▧", "Area", None, tool="c_area", size=19, width=62)
        self.rb(g, "⌦", "Clear", self.cad_clear_marks, size=19, width=62)
        self.rb(g, "\U0001F4CC", "Put on page", self.cad_marks_to_page, size=18, width=70)
        g = self.group(r, "Draw")
        self.rb(g, "〰", "Polyline", None, tool="c_poly", size=19, width=62)
        self.rb(g, "⬠", "Polygon", None, tool="c_polygon", size=19, width=62)
        self.rb(g, "◜", "Arc", None, tool="c_arc", size=19, width=62)
        self.rb(g, "╱", "Exact line", None, tool="c_xline", size=19, width=70)
        self.rb(g, "⇹", "Dimension", None, tool="c_dim", size=19, width=70)
        g = self.group(r, "Snap")
        f = tk.Frame(g, bg=bg)
        f.pack(padx=6, pady=(2, 0))

        def chk(text, var, row, col=0):
            tk.Checkbutton(f, text=_tr(text), variable=var, bg=bg, activebackground=bg, font=(UI_FONT, 8),
                           command=self._cad_snap_changed).grid(row=row, column=col, sticky="w")
        chk("Line ends, middles, corners", v["geo"], 0)
        chk("Grid", v["grid"], 1)
        chk("Angles of 15° (or hold Shift)", v["ang"], 2)
        tk.Label(f, text=_tr("Grid size"), bg=bg, fg=TH.LABEL, font=(UI_FONT, 8)).grid(row=1, column=1, sticky="e", padx=(10, 2))
        ttk.Spinbox(f, from_=0.01, to=1000, increment=0.5, width=6, textvariable=v["gridsize"]).grid(row=1, column=2, sticky="w")
        self.shape_style_group(r)
        self.cad_scale_refresh()

    def _cad_snap_changed(self):
        self.canvas.delete("cadsnap")

    # ------------------------------------------------------------------ snapping
    def _cad_geo(self):
        """the points of the drawing on this page that can be snapped to (view space), bucketed for a fast search"""
        key = (id(self.doc), self.page_no, getattr(self, "_edit_serial", 0))
        if key == self._cad_geo_key:
            return self._cad_geo_cells
        cells = {}
        n = 0
        pg = self.page()
        m = pg.rotation_matrix

        def add(p, kind):
            nonlocal n
            if n > 40000:
                return
            q = fitz.Point(p) * m
            cells.setdefault((int(q.x // CELL), int(q.y // CELL)), []).append((q.x, q.y, kind))
            n += 1
        try:
            for d in pg.get_drawings()[:8000]:
                for it in d["items"]:
                    k = it[0]
                    if k == "l":
                        add(it[1], 0)
                        add(it[2], 0)
                        add(((it[1].x + it[2].x) / 2, (it[1].y + it[2].y) / 2), 1)
                    elif k == "c":
                        add(it[1], 0)
                        add(it[4], 0)
                    elif k in ("re", "qu"):
                        q = it[1]
                        cs = [fitz.Point(q.x0, q.y0), fitz.Point(q.x1, q.y0), fitz.Point(q.x1, q.y1), fitz.Point(q.x0, q.y1)] if k == "re" \
                            else [q.ul, q.ur, q.lr, q.ll]
                        for i, p in enumerate(cs):
                            add(p, 0)
                            nx = cs[(i + 1) % 4]
                            add(((p.x + nx.x) / 2, (p.y + nx.y) / 2), 1)
        except Exception:
            pass
        self._cad_geo_key, self._cad_geo_cells = key, cells
        return cells

    def cad_snap(self, vx, vy, shift=False, ref=None):
        """-> (x, y, kind) kind: 'end', 'mid', 'grid', 'angle' or None.  Coordinates are view points."""
        v = self._cad_vars()
        zoom = max(self.zoom, 1e-6)
        R = SNAP_PIXELS / zoom
        best = None
        if v["geo"].get():
            cells = self._cad_geo()
            rng = int(R // CELL) + 1
            cx0, cy0 = int(vx // CELL), int(vy // CELL)
            bd = R
            for ix in range(cx0 - rng, cx0 + rng + 1):
                for iy in range(cy0 - rng, cy0 + rng + 1):
                    for x, y, kind in cells.get((ix, iy), ()):
                        dd = math.hypot(x - vx, y - vy) + (0.0 if kind == 0 else R * 0.15)
                        if dd < bd:
                            bd, best = dd, (x, y, "end" if kind == 0 else "mid")
        if best:
            return best
        if v["grid"].get():
            try:
                g = float(str(v["gridsize"].get()).replace(",", "."))
                step = g / self.cad_scale()["per_pt"]
            except (ValueError, ZeroDivisionError):
                step = 0
            if step * zoom >= 5:
                gx, gy = round(vx / step) * step, round(vy / step) * step
                if math.hypot(gx - vx, gy - vy) <= max(R, step * 0.5):
                    return gx, gy, "grid"
        if ref is not None and (v["ang"].get() or shift):
            dx, dy = vx - ref[0], vy - ref[1]
            L = math.hypot(dx, dy)
            if L > 1e-6:
                ang = round(math.atan2(dy, dx) / (math.pi / 12)) * (math.pi / 12)
                return ref[0] + L * math.cos(ang), ref[1] + L * math.sin(ang), "angle"
        return vx, vy, None

    def cad_snap_canvas(self, cx, cy, shift=False, ref_canvas=None):
        """snap canvas coordinates (used by the shape tools too); returns canvas coordinates"""
        if not hasattr(self, "_cadv"):
            return cx, cy
        v = self._cadv
        if not (v["geo"].get() or v["grid"].get() or v["ang"].get()):
            return cx, cy
        ref = self._c2v(*ref_canvas) if ref_canvas else None
        x, y, kind = self.cad_snap(*self._c2v(cx, cy), shift=shift, ref=ref)
        return self._v2c(x, y) if kind else (cx, cy)

    def _cad_marker(self, x, y, kind):
        cv = self.canvas
        cv.delete("cadsnap")
        if not kind:
            return
        cx, cy = self._v2c(x, y)
        col = "#00a86b"
        if kind == "end":
            cv.create_rectangle(cx - 6, cy - 6, cx + 6, cy + 6, outline=col, width=2, tags="cadsnap")
        elif kind == "mid":
            cv.create_polygon(cx, cy - 7, cx + 7, cy + 6, cx - 7, cy + 6, outline=col, fill="", width=2, tags="cadsnap")
        elif kind == "grid":
            cv.create_line(cx - 6, cy, cx + 6, cy, fill=col, width=2, tags="cadsnap")
            cv.create_line(cx, cy - 6, cx, cy + 6, fill=col, width=2, tags="cadsnap")
        else:
            cv.create_oval(cx - 4, cy - 4, cx + 4, cy + 4, outline=col, width=2, tags="cadsnap")

    # ------------------------------------------------------------------ input
    def cad_pending(self):
        return hasattr(self, "_cad") and bool(self._cad["pts"])

    def cad_cancel(self):
        if hasattr(self, "_cad"):
            self._cad["pts"] = []
            self.canvas.delete("cad")
            self.canvas.delete("cadsnap")

    def _cad_cursor_pt(self, e):
        shift = bool(e.state & 0x0001)
        cx, cy = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        vx, vy = self._c2v(cx, cy)
        ref = self._cad["pts"][-1] if self._cad["pts"] else None
        x, y, kind = self.cad_snap(vx, vy, shift, ref)
        return (x, y), kind

    def cad_motion(self, e):
        now = getattr(self, "_cad_mt", 0.0)
        import time
        t = time.monotonic()
        if t - now < 0.025:
            return
        self._cad_mt = t
        self._cad_vars()
        pt, kind = self._cad_cursor_pt(e)
        self._cad["cur"] = pt
        self._cad_marker(pt[0], pt[1], kind)
        if self._cad["pts"]:
            self.cad_draw()

    def cad_press(self, e):
        self._cad_vars()
        st = self._cad
        tool = self.tool.get()
        if st["page"] is not None and st["page"] != self.page_no:           # another page: start again
            st["pts"] = []
        st["page"] = self.page_no
        st["tool"] = tool
        pt, kind = self._cad_cursor_pt(e)
        pts = st["pts"]
        if pts and math.dist(pts[-1], pt) * self.zoom < 4:                   # second click of a double-click
            if tool in ("c_len", "c_area", "c_poly", "c_polygon"):
                self.cad_finish()
            return
        pts.append(pt)
        need = {"c_dist": 2, "c_cal": 2, "c_angle": 3, "c_arc": 3, "c_dim": 3, "c_xline": 1}.get(tool)
        if need and len(pts) >= need:
            self.cad_finish()
        else:
            self.cad_draw()

    def cad_finish(self):
        st = self._cad
        tool, pts = st["tool"], list(st["pts"])
        st["pts"] = []
        self.canvas.delete("cadsnap")
        if not pts:
            return
        try:
            if tool == "c_dist" and len(pts) >= 2:
                self._cad_add_mark("dist", pts[:2], False)
            elif tool == "c_len" and len(pts) >= 2:
                self._cad_add_mark("len", pts, False)
            elif tool == "c_angle" and len(pts) >= 3:
                self._cad_add_mark("angle", pts[:3], False)
            elif tool == "c_area" and len(pts) >= 3:
                self._cad_add_mark("area", pts, True)
            elif tool == "c_cal" and len(pts) >= 2:
                self.cad_draw()
                self._cad_calibrate(pts[0], pts[1])
            elif tool == "c_poly" and len(pts) >= 2:
                self._cad_commit([("open", pts, False)], closed=False)
            elif tool == "c_polygon" and len(pts) >= 3:
                self._cad_commit([("poly", pts, True)], closed=True)
            elif tool == "c_arc" and len(pts) >= 3:
                self._cad_commit([("open", arc_points(*pts[:3]), False)], closed=False)
            elif tool == "c_dim" and len(pts) >= 3:
                self._cad_dimension(pts[0], pts[1], pts[2])
            elif tool == "c_xline" and len(pts) >= 1:
                self._cad_exact_line(pts[0])
        except Exception as ex:
            messagebox.showerror(APP, _tr("That did not work:") + "\n" + str(ex))
        self.cad_draw()

    # ------------------------------------------------------------------ measuring
    def _cad_add_mark(self, kind, pts, closed):
        if kind == "dist":
            d = math.dist(pts[0], pts[1])
            dx, dy = pts[1][0] - pts[0][0], pts[1][1] - pts[0][1]
            ang = math.degrees(math.atan2(-dy, dx)) % 360
            sc = self.cad_scale()["per_pt"]
            text = self._cad_len(d)
            info = "%s   Δx %s   Δy %s   %.1f°" % (text, fmt_num(abs(dx) * sc), fmt_num(abs(dy) * sc), ang)
            at = ((pts[0][0] + pts[1][0]) / 2, (pts[0][1] + pts[1][1]) / 2)
        elif kind == "len":
            text = self._cad_len(path_length(pts))
            info = _tr("Length") + ": " + text
            at = pts[-1]
        elif kind == "angle":
            v1 = (pts[0][0] - pts[1][0], pts[0][1] - pts[1][1])
            v2 = (pts[2][0] - pts[1][0], pts[2][1] - pts[1][1])
            a = abs(math.degrees(math.atan2(v1[0] * v2[1] - v1[1] * v2[0], v1[0] * v2[0] + v1[1] * v2[1])))
            text = "%.2f°" % a
            info = _tr("Angle") + ": " + text
            at = pts[1]
        else:
            ar = poly_area(pts)
            pr = path_length(pts, True)
            text = self._cad_area(ar)
            info = "%s: %s   %s: %s" % (_tr("Area"), text, _tr("Perimeter"), self._cad_len(pr))
            at = (sum(p[0] for p in pts) / len(pts), sum(p[1] for p in pts) / len(pts))
        self._cad_marks.append({"doc": id(self.doc), "page": self.page_no, "kind": kind, "pts": pts, "closed": closed, "text": text, "at": at})
        self.status.config(text=info)

    def cad_clear_marks(self):
        if hasattr(self, "_cad_marks"):
            self._cad_marks = []
        self.cad_cancel()
        self.cad_draw()
        self.status.config(text=_tr("Measurements cleared."))

    def cad_marks_to_page(self):
        """draw the measurements of this page onto the page itself (lines and text)"""
        if not self.need_doc():
            return
        self._cad_vars()
        marks = [m for m in self._cad_marks if m["page"] == self.page_no and m.get("doc") == id(self.doc)]
        if not marks:
            self.status.config(text=_tr("There are no measurements on this page."))
            return
        self.commit_floats()
        self.snapshot()
        pg = self.page()
        col = (0.84, 0.2, 0.42)
        for m in marks:
            items = [("poly" if m["closed"] else "open", m["pts"], False)]
            self._write_shape(pg, items, 1.0, col, None, 1.0)
            self._cad_text(pg, m["at"], m["text"], 9.0, (1.0, 0.0), col)
        self._cad_marks = [m for m in self._cad_marks if m not in marks]
        self.edited()
        self.cad_draw()
        self.status.config(text=_tr("Measurements put on the page."))

    # ------------------------------------------------------------------ drawing on the page
    def _cad_commit(self, items, closed):
        self.commit_floats()
        self.snapshot()
        pg = self.page()
        op = max(0.05, min(1.0, self.shape_opacity.get() / 100))
        fill = self.fill_color if (self.fill_on.get() and closed) else None
        self._write_shape(pg, items, self.line_width.get(), self.color, fill, op)
        self.edited()
        self.status.config(text=_tr("Drawn. Right-click it > Rotate this shape if you want to turn it."))

    def _cad_text(self, pg, at, text, size, direction, color):
        """write `text` centred on view point `at`, running along `direction` (view space), readable from the left"""
        dx, dy = direction
        if dx < -1e-9 or (abs(dx) <= 1e-9 and dy > 0):
            dx, dy = -dx, -dy
        L = math.hypot(dx, dy) or 1.0
        dx, dy = dx / L, dy / L
        nx, ny = dy, -dx                                           # "above" the text line
        w = fitz.get_text_length(text, fontname="helv", fontsize=size)
        gap = 2.5
        start = (at[0] - dx * w / 2 + nx * gap, at[1] - dy * w / 2 + ny * gap)
        m = pg.derotation_matrix
        p0 = fitz.Point(*start) * m
        v = fitz.Point(dx, dy) * m - fitz.Point(0, 0) * m
        ang = math.degrees(math.atan2(v.y, v.x))
        pg.insert_text(p0, text, fontsize=size, fontname="helv", color=color, morph=(p0, fitz.Matrix(ang)))

    def _cad_dimension(self, a, b, off):
        lw = max(0.5, min(2.0, float(self.line_width.get())))
        items, mid, direction = dimension_geometry(a, b, off, lw)
        if not items:
            return
        self.commit_floats()
        self.snapshot()
        pg = self.page()
        op = max(0.05, min(1.0, self.shape_opacity.get() / 100))
        self._write_shape(pg, items, lw, self.color, None, op)
        self._cad_text(pg, mid, self._cad_len(math.dist(a, b)), 10.0, direction, self.color)
        self.edited()
        self.status.config(text=_tr("Dimension added."))

    def _cad_exact_line(self, start):
        sc = self.cad_scale()
        d = self._dlg(_tr("Exact line"))
        self._label(d, _tr("Length"), bold=True).grid(row=0, column=0, sticky="w", pady=3)
        lv = tk.StringVar(value="")
        e = ttk.Entry(d, textvariable=lv, width=12)
        e.grid(row=0, column=1, sticky="w")
        e.focus_set()
        self._label(d, sc["unit"]).grid(row=0, column=2, sticky="w", padx=4)
        self._label(d, _tr("Angle"), bold=True).grid(row=1, column=0, sticky="w", pady=3)
        av = tk.StringVar(value="0")
        ttk.Entry(d, textvariable=av, width=12).grid(row=1, column=1, sticky="w")
        self._label(d, "°  (0 = right, 90 = up)").grid(row=1, column=2, sticky="w", padx=4)

        def ok(_e=None):
            try:
                L = float(lv.get().replace(",", "."))
                a = float(av.get().replace(",", "."))
                if L <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror(APP, _tr("Type a length and an angle as numbers."), parent=d)
                return
            d.destroy()
            Lp = L / sc["per_pt"]
            end = (start[0] + Lp * math.cos(math.radians(a)), start[1] - Lp * math.sin(math.radians(a)))
            self._cad_commit([("line", [start, end], False)], closed=False)
        d.bind("<Return>", ok)
        self._buttons(d, _tr("Draw"), ok, 2, 3)

    # ------------------------------------------------------------------ canvas drawing (live preview and measurements)
    def cad_draw(self):
        if not hasattr(self, "_cad"):
            return
        cv = self.canvas
        cv.delete("cad")
        if self.doc is None:
            return
        z = self.zoom

        def C(p):
            return self._v2c(*p)

        def label(p, text, dy=-14):
            x, y = C(p)
            t = cv.create_text(x, y + dy, text=text, fill=MARK_COLOR, font=(UI_FONT, 10, "bold"), tags="cad")
            bb = cv.bbox(t)
            if bb:
                r = cv.create_rectangle(bb[0] - 3, bb[1] - 1, bb[2] + 3, bb[3] + 1, fill="white", outline=MARK_COLOR, tags="cad")
                cv.tag_lower(r, t)

        for m in self._cad_marks:
            if m["page"] != self.page_no or m.get("doc") != id(self.doc):
                continue
            flat = [c for p in m["pts"] for c in C(p)]
            if m["closed"]:
                cv.create_polygon(*flat, outline=MARK_COLOR, fill="", width=2, tags="cad")
            else:
                cv.create_line(*flat, fill=MARK_COLOR, width=2, tags="cad")
            for p in m["pts"]:
                x, y = C(p)
                cv.create_oval(x - 3, y - 3, x + 3, y + 3, fill=MARK_COLOR, outline="", tags="cad")
            label(m["at"], m["text"])
        st = self._cad
        pts = st["pts"]
        if not pts or st.get("page") != self.page_no:
            return
        tool = st.get("tool")
        cur = st.get("cur") or pts[-1]
        seq = pts + [cur]
        col = self._hex(self.color) if tool in ("c_poly", "c_polygon", "c_arc", "c_dim", "c_xline") else MARK_COLOR
        w = max(1, (self.line_width.get() if tool in ("c_poly", "c_polygon", "c_arc", "c_dim") else 1.5) * z)
        if tool == "c_arc" and len(pts) == 2:
            ap = arc_points(pts[0], pts[1], cur)
            cv.create_line(*[c for p in ap for c in C(p)], fill=col, width=w, tags="cad")
        elif tool == "c_dim" and len(pts) == 2:
            items, mid, direction = dimension_geometry(pts[0], pts[1], cur, 1.0)
            for typ, data, _ in items:
                flat = [c for p in data for c in C(p)]
                if typ == "line":
                    cv.create_line(*flat, fill=col, width=1.5, tags="cad")
                else:
                    cv.create_polygon(*flat, fill=col, outline=col, tags="cad")
            label(mid, self._cad_len(math.dist(pts[0], pts[1])), -12)
        else:
            flat = [c for p in seq for c in C(p)]
            if len(flat) >= 4:
                cv.create_line(*flat, fill=col, width=w, dash=(5, 3), tags="cad")
            if tool in ("c_area", "c_polygon") and len(seq) >= 3:
                cv.create_line(*C(seq[-1]), *C(seq[0]), fill=col, width=1, dash=(2, 4), tags="cad")
        for p in pts:
            x, y = C(p)
            cv.create_rectangle(x - 3, y - 3, x + 3, y + 3, fill="white", outline=MARK_COLOR, tags="cad")
        # live numbers next to the cursor
        if tool in ("c_dist", "c_len", "c_cal", "c_xline") and len(pts) >= 1:
            tot = path_length(pts + [cur]) if tool == "c_len" else math.dist(pts[-1], cur)
            label(cur, self._cad_len(tot), -16)
        elif tool == "c_area" and len(pts) >= 2:
            label(cur, self._cad_area(poly_area(seq)), -16)
        elif tool == "c_angle" and len(pts) == 2:
            v1 = (pts[0][0] - pts[1][0], pts[0][1] - pts[1][1])
            v2 = (cur[0] - pts[1][0], cur[1] - pts[1][1])
            a = abs(math.degrees(math.atan2(v1[0] * v2[1] - v1[1] * v2[0], v1[0] * v2[0] + v1[1] * v2[1])))
            label(pts[1], "%.1f°" % a, -16)
