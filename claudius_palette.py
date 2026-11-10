"""
Claudius PDF Editor - command palette (Ctrl+K) and guided actions.

  * Find tools and commands   one search box for every ribbon button and the main dialogs; shows where each lives
  * Guided actions            ready-made step lists (share safely, archive, legal documents ...) that walk you through the
                              program's own tools one after another
  * Recommended tools         cards on the start screen
"""
import os
import tkinter as tk
from tkinter import ttk

import claudius_theme as TH
from claudius_i18n import tr as _tr
from claudius_extra import APP, UI_FONT, BG, ACCENT, S
from claudius_view import get_setting, set_setting

RECENT_MAX = 8

# (text, keywords, method, sidebar-or-tab hint)
EXTRA_COMMANDS = (
    ("Compare two PDFs", "difference changes versions", "dlg_compare", "Review"),
    ("Accessibility check", "pdf/ua screen reader tags title language", "dlg_accessibility", "Review"),
    ("Apply PDF standards (PDF/A)", "archive pdfa standard", "dlg_pdfa", "Document"),
    ("Remove metadata", "clean hidden information author privacy", "dlg_clean_metadata", "Document"),
    ("Check redactions", "redaction verify blackout hidden text", "dlg_redaction_check", "Review"),
    ("Redact patterns", "emails phone numbers iban regex blackout", "dlg_redact_patterns", "Review"),
    ("Compress PDF", "reduce size optimise optimize smaller", "dlg_compress", "Document"),
    ("Protect with password", "encrypt security permissions", "dlg_protect", "Document"),
    ("Watermark", "stamp background text draft confidential", "dlg_watermark", "Document"),
    ("Header and footer", "page numbers date", "dlg_header_footer", "Document"),
    ("Bates numbering", "legal numbering stamp", "dlg_bates", "Document"),
    ("Guided actions", "wizard steps workflow checklist", "dlg_guided", "Tools"),
)

# preset name -> (description, [(label, method)])
GUIDED = (
    ("Share safely",
     "Clean a file before you send it to someone outside.",
     [("Remove hidden metadata", "dlg_clean_metadata"), ("Check that nothing is hidden under blackouts", "dlg_redaction_check"),
      ("Make the file smaller", "dlg_compress"), ("Protect with a password", "dlg_protect")]),
    ("Legal documents",
     "Black out personal data, number the pages and lock the result.",
     [("Redact emails, numbers, IBANs...", "dlg_redact_patterns"), ("Check the redactions", "dlg_redaction_check"),
      ("Bates numbering", "dlg_bates"), ("Protect with a password", "dlg_protect")]),
    ("Archive (long-term)",
     "Make the file accessible and convert it to PDF/A.",
     [("Remove hidden metadata", "dlg_clean_metadata"), ("Accessibility check and fix", "dlg_accessibility"),
      ("Convert to PDF/A", "dlg_pdfa")]),
    ("Prepare for printing or publishing",
     "Add page numbers and a watermark, then make the file smaller.",
     [("Header and footer / page numbers", "dlg_header_footer"), ("Watermark", "dlg_watermark"), ("Make the file smaller", "dlg_compress")]),
)

# start-screen cards: (icon, title, subtitle, command text)
START_CARDS = (
    ("▦", "Organize pages", "Rotate, move, delete", "Organize pages"),
    ("\U0001F4AC", "Add comments", "Notes, highlights", "Comments"),
    ("✎", "Edit text", "Change words on the page", "Edit text"),
    ("\U0001F512", "Protect", "Password and redaction", "Protect with password"),
    ("↔", "Compare files", "Find what changed", "Compare two PDFs"),
    ("⚙", "CAD and measure", "Scale, distances, areas", "Distance"),
)


def match(query, cand):
    q = query.lower().split()
    hay = cand["hay"]
    return all(t in hay for t in q)


class PaletteTools:
    _pal_win = None

    # ------------------------------------------------------------------ the list of commands
    def pal_commands(self):
        self._ensure_all_ribbons()
        cmds, seen = [], set()
        ribbon_of = {fr: n for n, fr in self.ribbons.items()}
        for key, v in getattr(self, "qt_reg", {}).items():
            w, tab = v.get("btn"), ""
            while w is not None:
                if w in ribbon_of:
                    tab = ribbon_of[w]
                    break
                w = getattr(w, "master", None)
            ident = (v["text"], tab)
            if ident in seen or not v["text"].strip():
                continue
            seen.add(ident)
            cmds.append({"id": "r:%s|%s" % ident, "text": v["text"], "kw": "", "tab": tab, "run": self._pal_ribbon_runner(v, tab)})
        for text, kw, meth, tab in EXTRA_COMMANDS:
            if hasattr(self, meth):
                cmds.append({"id": "x:" + meth, "text": text, "kw": kw, "tab": tab,
                             "run": lambda m=meth: getattr(self, m)()})
        for text, kw, tab, fn in (
                ("Measure objects", "measure distance length angle area perimeter scale cad", "CAD", lambda: (self.show_tab("CAD"), self.set_tool("c_dist"))),
                ("Add comments", "comment note sticky highlight underline annotate", "Comment", lambda: self.show_tab("Comment")),
                ("Organize pages", "pages rotate move delete reorder extract insert", "Organize", lambda: self.show_tab("Organize")),
                ("Fill and sign", "form signature initials checkmark", "Sign", lambda: self.show_tab("Sign")),
                ("Prepare a form", "form fields", "Sign", lambda: self.show_tab("Sign"))):
            if tab in self.ribbons:
                cmds.append({"id": "a:" + text, "text": text, "kw": kw, "tab": tab, "run": fn})
        for name in self.ribbons:
            cmds.append({"id": "t:" + name, "text": "Go to the %s tab" % name, "kw": "ribbon tab", "tab": "",
                         "run": lambda n=name: self.show_tab(n)})
        for label, mode in (("Show the page thumbnails", "pages"), ("Show the bookmarks", "bookmarks"), ("Show the comments list", "comments")):
            cmds.append({"id": "s:" + mode, "text": label, "kw": "side panel sidebar", "tab": "Side panel",
                         "run": lambda m=mode: self.set_sidebar(m)})
        try:
            for p in (self.load_recents() or [])[:6]:
                cmds.append({"id": "o:" + p, "text": "Open recent: " + os.path.basename(p), "kw": p, "tab": "File",
                             "run": lambda p=p: self._open_recent(p)})
        except Exception:
            pass
        for c in cmds:
            c["shown"] = _tr(c["text"]) if not c["text"].startswith(("Go to the ", "Open recent: ")) else self._pal_compound(c["text"])
            c["where"] = _tr(c["tab"]) if c["tab"] else ""
            c["hay"] = " ".join((c["shown"], c["text"], c["kw"], c["tab"], c["where"])).lower()
        return cmds

    @staticmethod
    def _pal_compound(text):
        if text.startswith("Go to the "):
            tab = text[len("Go to the "):-len(" tab")]
            return _tr("Go to the {} tab").format(_tr(tab)) if "{}" in _tr("Go to the {} tab") else "%s: %s" % (_tr("Tab"), _tr(tab))
        return text.replace("Open recent: ", _tr("Open recent") + ": ")

    def _pal_ribbon_runner(self, v, tab):
        def run():
            if tab and tab in self.ribbons:
                self.show_tab(tab)
            if v.get("tool"):
                self.set_tool(v["tool"])
            elif v.get("cmd"):
                v["cmd"]()
        return run

    # ------------------------------------------------------------------ the palette window
    def pal_open(self, initial=""):
        w = self._pal_win
        if w is not None:
            try:
                if w.winfo_exists():
                    w.lift()
                    w._entry.focus_set()
                    return
            except tk.TclError:
                pass
        d = tk.Toplevel(self)
        self._pal_win = d
        d.title(_tr("Find tools and commands"))
        d.transient(self)
        d.configure(bg=BG, padx=10, pady=10)
        width, height = S(560), S(380)
        d.geometry("%dx%d+%d+%d" % (width, height, self.winfo_rootx() + max(20, (self.winfo_width() - width) // 2), self.winfo_rooty() + S(70)))
        d.resizable(False, False)
        var = tk.StringVar(value=initial)
        e = ttk.Entry(d, textvariable=var, font=(UI_FONT, 12))
        e.pack(fill="x", pady=(0, 8))
        d._entry = e
        tree = ttk.Treeview(d, columns=("name", "where"), show="", selectmode="browse", height=12)
        tree.column("name", width=S(380), anchor="w")
        tree.column("where", width=S(150), anchor="e")
        tree.pack(fill="both", expand=True)
        tk.Label(d, text=_tr("Enter = run   Up / Down = choose   Esc = close"), bg=BG, fg=TH.TERTIARY, font=(UI_FONT, 8)).pack(anchor="w", pady=(6, 0))
        cmds = self.pal_commands()
        by_id = {c["id"]: c for c in cmds}
        rows = {}

        def refill(*_a):
            tree.delete(*tree.get_children())
            rows.clear()
            q = var.get().strip()
            recent = [by_id[i] for i in (get_setting("palette_recent", []) or []) if i in by_id]
            if q:
                hits = [c for c in cmds if match(q, c)]
                ql = q.lower()
                hits.sort(key=lambda c: (not c["shown"].lower().startswith(ql), c not in recent, c["shown"].lower()))
            else:
                hits = recent + sorted((c for c in cmds if c not in recent and c["id"][0] in "rx"), key=lambda c: c["shown"].lower())
            for c in hits[:80]:
                iid = tree.insert("", "end", values=(c["shown"], c["where"]))
                rows[iid] = c
            kids = tree.get_children()
            if kids:
                tree.selection_set(kids[0])
                tree.focus(kids[0])

        def run(_e=None):
            sel = tree.selection()
            if not sel:
                return "break"
            c = rows[sel[0]]
            try:
                rec = [i for i in (get_setting("palette_recent", []) or []) if i != c["id"]]
                set_setting("palette_recent", ([c["id"]] + rec)[:RECENT_MAX])
            except Exception:
                pass
            d.destroy()
            self.after(30, c["run"])
            return "break"

        def move(delta):
            kids = tree.get_children()
            if not kids:
                return "break"
            sel = tree.selection()
            i = kids.index(sel[0]) if sel else -1
            j = max(0, min(len(kids) - 1, i + delta))
            tree.selection_set(kids[j])
            tree.see(kids[j])
            return "break"
        var.trace_add("write", refill)
        e.bind("<Return>", run)
        e.bind("<Up>", lambda ev: move(-1))
        e.bind("<Down>", lambda ev: move(1))
        tree.bind("<Double-1>", run)
        d.bind("<Escape>", lambda ev: d.destroy())
        refill()
        d.after(30, lambda: (d.lift(), e.focus_force(), e.icursor("end")))

    def pal_run_by_text(self, text):
        """used by the start-screen cards"""
        for c in self.pal_commands():
            if c["text"] == text or c["id"] == text:
                c["run"]()
                return True
        return False

    def pal_start_card(self, name):
        if name == "Comments":
            if self.doc is None:
                self.open_dialog()
                return
            self.show_tab("Comment")
            self.set_sidebar("comments")
            return
        if not self.pal_run_by_text(name):
            self.pal_open(_tr(name))

    # ------------------------------------------------------------------ guided actions
    def dlg_guided(self):
        if not self.need_doc():
            return
        w = getattr(self, "_guided_win", None)
        try:
            if w is not None and w.winfo_exists():
                w.lift()
                return
        except tk.TclError:
            pass
        d = tk.Toplevel(self)
        self._guided_win = d
        d.title(_tr("Guided actions"))
        d.transient(self)
        d.configure(bg=BG, padx=16, pady=12)
        d.geometry("+%d+%d" % (self.winfo_rootx() + 120, self.winfo_rooty() + 90))
        tk.Label(d, text=_tr("Choose a job"), bg=BG, fg=TH.LABEL, font=(UI_FONT, 10, "bold")).grid(row=0, column=0, sticky="w")
        names = [g[0] for g in GUIDED]
        pv = tk.StringVar(value=_tr(names[0]))
        cb = ttk.Combobox(d, textvariable=pv, values=[_tr(n) for n in names], state="readonly", width=34)
        cb.grid(row=1, column=0, sticky="w", pady=(2, 4))
        desc = tk.Label(d, text="", bg=BG, fg=TH.SECOND, font=(UI_FONT, 9), wraplength=S(420), justify="left")
        desc.grid(row=2, column=0, sticky="w", pady=(0, 8))
        lb = tk.Listbox(d, height=6, width=52, font=(UI_FONT, 10), activestyle="none", bg=TH.CARD, fg=TH.LABEL,
                        selectbackground=ACCENT, selectforeground="white", relief="flat", highlightthickness=1,
                        highlightbackground=TH.LINE)
        lb.grid(row=3, column=0, sticky="ew")
        info = tk.Label(d, text="", bg=BG, fg=TH.TERTIARY, font=(UI_FONT, 8), wraplength=S(420), justify="left")
        info.grid(row=4, column=0, sticky="w", pady=(6, 0))
        st = {"preset": 0, "state": [], "i": 0}

        def label_of(i):
            mark = {"": "○", "done": "✓", "skip": "↷"}[st["state"][i]]
            return "%s  %d. %s" % (mark, i + 1, _tr(GUIDED[st["preset"]][2][i][0]))

        def fill():
            p = GUIDED[st["preset"]]
            desc.config(text=_tr(p[1]))
            st["state"] = [""] * len(p[2])
            st["i"] = 0
            lb.delete(0, "end")
            for i in range(len(p[2])):
                lb.insert("end", label_of(i))
            lb.selection_set(0)
            btn_do.config(state="normal")
            btn_skip.config(state="normal")
            info.config(text=_tr("Each step opens the normal tool. Finish it (or close it), and the next step is ready."))

        def pick(_e=None):
            st["preset"] = cb.current()
            fill()

        def advance(kind):
            i = st["i"]
            if i >= len(st["state"]):
                return
            st["state"][i] = kind
            lb.delete(i)
            lb.insert(i, label_of(i))
            st["i"] = i + 1
            if st["i"] >= len(st["state"]):
                lb.selection_clear(0, "end")
                btn_do.config(state="disabled")
                btn_skip.config(state="disabled")
                info.config(text=_tr("All steps are done. Save the file (Ctrl+S) to keep the result."))
            else:
                lb.selection_clear(0, "end")
                lb.selection_set(st["i"])

        def do_step():
            i = st["i"]
            if i >= len(st["state"]) or self.doc is None:
                return
            meth = GUIDED[st["preset"]][2][i][1]
            before = {c for c in self.winfo_children() if isinstance(c, tk.Toplevel)}
            try:
                getattr(self, meth)()
            except Exception as ex:
                info.config(text=_tr("That step did not work:") + " " + str(ex))
                return
            self.update_idletasks()
            for c in [c for c in self.winfo_children() if isinstance(c, tk.Toplevel) and c not in before and c is not d]:
                try:
                    self.wait_window(c)
                except tk.TclError:
                    pass
            if d.winfo_exists():
                advance("done")

        bf = tk.Frame(d, bg=BG)
        bf.grid(row=5, column=0, sticky="e", pady=(12, 0))
        btn_do = ttk.Button(bf, text=_tr("Do this step"), command=do_step)
        btn_do.pack(side="left", padx=3)
        btn_skip = ttk.Button(bf, text=_tr("Skip"), command=lambda: advance("skip"))
        btn_skip.pack(side="left", padx=3)
        ttk.Button(bf, text=_tr("Close"), command=d.destroy).pack(side="left", padx=3)
        cb.bind("<<ComboboxSelected>>", pick)
        d.bind("<Escape>", lambda e: d.destroy())
        fill()
        d._st, d._do, d._skip, d._cb = st, do_step, advance, cb
