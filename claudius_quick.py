# -*- coding: utf-8 -*-
"""Claudius PDF Editor - the quick toolbar.

A small free-floating window with the buttons you use most. Move it anywhere (even next to the page on a second
screen), resize it and the buttons re-flow. Right-click any ribbon button -> "Add to quick toolbar", or press "＋"
on the quick toolbar to pick buttons from a list. Right-click a quick button to remove it or move it.
What is on the toolbar, where it is and whether it is shown is remembered.
"""
import tkinter as tk
from tkinter import ttk

from claudius_i18n import tr as _tr
from claudius_view import get_setting, set_setting

DEFAULTS = ("Open", "Save", "Undo", "Redo", "Print", "Hand", "Find", "Edit text")
CELL_W = 74


class QuickTools:
    _qt_win = None

    # ------------------------------------------------------------ registry (every ribbon button reports here)
    def qt_register(self, btn, icon, text, cmd, tool, size, width):
        reg = self.__dict__.setdefault("qt_reg", {})
        key = "%s|%s|%s" % (text, tool or "", icon)
        reg[key] = {"icon": icon, "text": text, "cmd": cmd, "tool": tool, "size": size, "width": width,
                    "color": getattr(btn, "_color", None)}
        btn.bind("<Button-3>", lambda e, k=key: self.qt_ribbon_menu(e, k), add="+")
        btn._qt_key = key

    def qt_items(self):
        items = self.__dict__.setdefault("_qt_items", None)
        if items is None:
            saved = (get_setting("quick_tools", {}) or {}).get("items")
            if isinstance(saved, list):
                items = [k for k in saved if k in self.qt_reg]
            else:
                items = []
                for d in DEFAULTS:
                    k = next((k for k, v in self.qt_reg.items() if v["text"] == d), None)
                    if k:
                        items.append(k)
            self._qt_items = items
        return items

    def _qt_save(self, **extra):
        try:
            cfg = dict(get_setting("quick_tools", {}) or {})
            cfg["items"] = list(self.qt_items())
            cfg.update(extra)
            set_setting("quick_tools", cfg)
        except Exception:
            pass

    # ------------------------------------------------------------ menus
    def qt_ribbon_menu(self, e, key):
        m = tk.Menu(self, tearoff=0)
        if key in self.qt_items():
            m.add_command(label=_tr("Remove from quick toolbar"), command=lambda: self.qt_remove(key))
        else:
            m.add_command(label=_tr("Add to quick toolbar"), command=lambda: self.qt_add(key))
        m.add_command(label=_tr("Show quick toolbar"), command=self.qt_show)
        m.tk_popup(e.x_root, e.y_root)

    def _qt_button_menu(self, e, key):
        m = tk.Menu(self, tearoff=0)
        m.add_command(label=_tr("Remove from quick toolbar"), command=lambda: self.qt_remove(key))
        m.add_command(label=_tr("Move left"), command=lambda: self.qt_move(key, -1))
        m.add_command(label=_tr("Move right"), command=lambda: self.qt_move(key, 1))
        m.add_separator()
        m.add_command(label=_tr("Choose buttons..."), command=self.qt_picker)
        m.tk_popup(e.x_root, e.y_root)

    # ------------------------------------------------------------ edit the list
    def qt_add(self, key):
        if key in self.qt_reg and key not in self.qt_items():
            self.qt_items().append(key)
            self._qt_save()
            self.qt_show()
            self._qt_rebuild()
            self.status.config(text=_tr("Added to the quick toolbar."))

    def qt_remove(self, key):
        if key in self.qt_items():
            self.qt_items().remove(key)
            self._qt_save()
            self._qt_rebuild()

    def qt_move(self, key, d):
        it = self.qt_items()
        i = it.index(key)
        j = i + d
        if 0 <= j < len(it):
            it[i], it[j] = it[j], it[i]
            self._qt_save()
            self._qt_rebuild()

    # ------------------------------------------------------------ the window
    def qt_toggle(self):
        if self._qt_win is not None and self._qt_win.winfo_exists():
            self.qt_hide()
        else:
            self.qt_show()

    def qt_show(self):
        if self._qt_win is not None and self._qt_win.winfo_exists():
            self._qt_win.deiconify()
            self._qt_win.lift()
            return
        w = self._qt_win = tk.Toplevel(self)
        w.title(_tr("Quick tools"))
        w.transient(self)
        try:
            w.attributes("-toolwindow", True)
        except tk.TclError:
            pass
        cfg = get_setting("quick_tools", {}) or {}
        geo = cfg.get("geometry")
        w.geometry(geo if isinstance(geo, str) and geo.count("x") == 1 and "+" in geo else "%dx%d+%d+%d" % (
            CELL_W * 4 + 16, 190, self.winfo_rootx() + 260, self.winfo_rooty() + 120))
        w.minsize(CELL_W + 16, 90)
        from claudius import RIBBON_BG
        w.configure(bg=RIBBON_BG)
        self._qt_flow = tk.Frame(w, bg=RIBBON_BG)
        self._qt_flow.pack(fill="both", expand=True, padx=4, pady=4)
        w.bind("<Configure>", self._qt_configure)
        w.protocol("WM_DELETE_WINDOW", self.qt_hide)
        self._qt_rebuild()
        self._qt_save(visible=True)

    def qt_hide(self):
        w = self._qt_win
        if w is not None and w.winfo_exists():
            self._qt_save(visible=False, geometry=self._qt_geometry())
            self._qt_forget_buttons()
            w.destroy()
        self._qt_win = None

    def _qt_geometry(self):
        w = self._qt_win
        return "%dx%d+%d+%d" % (w.winfo_width(), w.winfo_height(), w.winfo_x(), w.winfo_y())

    def _qt_configure(self, e):
        if e.widget is not self._qt_win:
            return
        self._qt_layout()
        if self.__dict__.get("_qt_after"):
            self.after_cancel(self._qt_after)
        self._qt_after = self.after(500, lambda: self._qt_win is not None and self._qt_win.winfo_exists()
                                    and self._qt_save(geometry=self._qt_geometry(), visible=True))

    def _qt_forget_buttons(self):
        for b in self.__dict__.get("_qt_btns", []):
            lst = self.tool_buttons.get(getattr(b, "_qt_tool", None))
            if lst and b in lst:
                lst.remove(b)
        self._qt_btns = []

    def _qt_rebuild(self):
        w = self._qt_win
        if w is None or not w.winfo_exists():
            return
        self._qt_forget_buttons()
        for c in self._qt_flow.winfo_children():
            c.destroy()
        from claudius import RibbonButton
        self._qt_btns = []
        for key in self.qt_items():
            d = self.qt_reg[key]
            if d["tool"]:
                b = RibbonButton(self._qt_flow, d["icon"], d["text"], lambda t=d["tool"]: self.set_tool(t),
                                 d["size"], d["width"], color=d["color"])
                b._qt_tool = d["tool"]
                self.tool_buttons.setdefault(d["tool"], []).append(b)
                b.set_active(self.tool.get() == d["tool"])
            else:
                b = RibbonButton(self._qt_flow, d["icon"], d["text"], d["cmd"], d["size"], d["width"], color=d["color"])
            b.bind("<Button-3>", lambda e, k=key: self._qt_button_menu(e, k))
            b._qt_key = key
            self._qt_btns.append(b)
        add = RibbonButton(self._qt_flow, "＋", "Add...", self.qt_picker, 20, 60)
        add._qt_key = None
        self._qt_btns.append(add)
        self._qt_layout()

    def _qt_layout(self):
        fl = getattr(self, "_qt_flow", None)
        if fl is None or not self._qt_btns:
            return
        width = max(self._qt_win.winfo_width() - 8, CELL_W)
        cols = max(1, width // CELL_W)
        for b in self._qt_btns:
            b.grid_forget()
        for i, b in enumerate(self._qt_btns):
            b.grid(row=i // cols, column=i % cols, padx=1, pady=1, sticky="n")

    # ------------------------------------------------------------ pick buttons from a list
    def qt_picker(self):
        old = self.__dict__.get("_qt_pick")
        if old is not None and old.winfo_exists():
            old.lift()
            return
        d = self._qt_pick = tk.Toplevel(self)
        d.title(_tr("Choose quick toolbar buttons"))
        d.transient(self)
        d.geometry("340x460+%d+%d" % (self.winfo_rootx() + 300, self.winfo_rooty() + 140))
        top = ttk.Frame(d, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text=_tr("Search")).pack(side="left")
        q = tk.StringVar()
        ttk.Entry(top, textvariable=q).pack(side="left", fill="x", expand=True, padx=6)
        body = ttk.Frame(d)
        body.pack(fill="both", expand=True, padx=8)
        cv = tk.Canvas(body, highlightthickness=0)
        sb = ttk.Scrollbar(body, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        cv.pack(side="left", fill="both", expand=True)
        inner = ttk.Frame(cv)
        cv.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<MouseWheel>", lambda e: cv.yview_scroll(-1 if e.delta > 0 else 1, "units"))
        vars_ = {}

        def toggle(key):
            if vars_[key].get():
                if key not in self.qt_items():
                    self.qt_items().append(key)
            elif key in self.qt_items():
                self.qt_items().remove(key)
            self._qt_save()
            self._qt_rebuild()

        def fill(*_a):
            for c in inner.winfo_children():
                c.destroy()
            needle = q.get().strip().lower()
            seen = set()
            for key, v in sorted(self.qt_reg.items(), key=lambda kv: kv[1]["text"].lower()):
                label = _tr(v["text"])
                if needle and needle not in label.lower() and needle not in v["text"].lower():
                    continue
                if (v["text"], v["icon"]) in seen:
                    continue
                seen.add((v["text"], v["icon"]))
                var = vars_.setdefault(key, tk.BooleanVar())
                var.set(key in self.qt_items())
                ttk.Checkbutton(inner, text="%s   %s" % (v["icon"], label), variable=var,
                                command=lambda k=key: toggle(k)).pack(anchor="w", pady=1)
        q.trace_add("write", fill)
        fill()
        ttk.Button(d, text=_tr("Close"), command=d.destroy).pack(pady=8)
        d.bind("<Escape>", lambda e: d.destroy())

    def qt_restore(self):
        """called once after start-up: re-open the toolbar when it was open last time"""
        try:
            if (get_setting("quick_tools", {}) or {}).get("visible"):
                self.qt_show()
        except Exception:
            pass
