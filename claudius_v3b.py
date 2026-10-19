"""
Claudius PDF Editor - version 3 features, part 2:
forms (drop-down / radio group tools, field editor, fill from CSV or Excel), batch processing of a folder,
smart compress with before / after preview, merge with a page picker, search in all PDFs of a folder.
Mixed into PdfEditor by claudius.py.
"""
import csv
import io
import os
import re
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pymupdf as fitz
except ImportError:
    import fitz
from PIL import Image, ImageTk

import claudius_theme as TH
from claudius_i18n import tr as _tr
import claudius_extra as EX
from claudius_extra import (APP, UI_FONT, BG, ACCENT, S, parse_pages, pages_to_text, runs, safe_name, human_size,
                            COMPRESS_LEVELS, compress_bytes, make_watermark_image)

MAX_ROWS = 2000                       # rows read from a data file
MAX_FILES = 5000                      # files handled by one batch / search
MAX_HITS = 500
MAX_FILE_MB = 300
TRUE_WORDS = {"1", "true", "yes", "y", "x", "on", "checked", "tak", "ja", "si", "sí", "sì", "oui", "wahr"}
FIELD_KINDS = {fitz.PDF_WIDGET_TYPE_TEXT: "Text", fitz.PDF_WIDGET_TYPE_CHECKBOX: "Check box",
               fitz.PDF_WIDGET_TYPE_RADIOBUTTON: "Radio button", fitz.PDF_WIDGET_TYPE_COMBOBOX: "Drop-down",
               fitz.PDF_WIDGET_TYPE_LISTBOX: "List", fitz.PDF_WIDGET_TYPE_BUTTON: "Button",
               fitz.PDF_WIDGET_TYPE_SIGNATURE: "Signature"}


# ====================================================================================== low-level PDF helpers
def _array_refs(text):
    return [int(m) for m in re.findall(r"(\d+)\s+0\s+R", text or "")]


def _get_array(doc, owner, key):
    """-> (refs, setter) for an array that may be direct or indirect"""
    kind, val = doc.xref_get_key(owner, key)
    if kind == "xref":
        x = int(val.split()[0])
        refs = _array_refs(doc.xref_object(x, compact=True))
        return refs, lambda new: doc.update_object(x, "[" + " ".join("%d 0 R" % r for r in new) + "]")
    refs = _array_refs(val) if kind == "array" else []
    return refs, lambda new: doc.xref_set_key(owner, key, "[" + " ".join("%d 0 R" % r for r in new) + "]")


def field_top(doc, xref):
    """the xref that carries the field name (the highest ancestor)"""
    seen = 0
    while seen < 20:
        kind, val = doc.xref_get_key(xref, "Parent")
        if kind != "xref":
            break
        xref = int(val.split()[0])
        seen += 1
    return xref


def widget_by_xref(page, xref):
    for w in page.widgets() or []:
        if w.xref == xref:
            return w
    return None


def radio_add(doc, page, group, export, rect):
    """add one radio button to the group `group` (the group is created when it does not exist yet)"""
    group = re.sub(r"[()\\]", "", group).strip()[:60]
    export = re.sub(r"[^A-Za-z0-9_]", "_", export.strip())[:40] or "Option"
    parent = None
    for pg in doc:
        for w in pg.widgets() or []:
            if w.field_name == group:
                if w.field_type != fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
                    raise ValueError(_tr("A field named \"{}\" already exists and is not a radio group.").replace("{}", group))
                kind, val = doc.xref_get_key(w.xref, "Parent")
                if kind != "xref":
                    raise ValueError(_tr("A field named \"{}\" already exists.").replace("{}", group))
                parent = int(val.split()[0])
    tmp = "__claudius_tmp_%d" % int(time.time() * 1000)
    w = fitz.Widget()
    w.field_type = fitz.PDF_WIDGET_TYPE_RADIOBUTTON
    w.field_name = tmp
    w.rect = fitz.Rect(rect)
    w.border_color = (0, 0, 0)
    w.border_width = 1
    w.field_value = False
    page.add_widget(w)
    kid = next(x.xref for x in page.widgets() if x.field_name == tmp)
    ap = doc.xref_get_key(kid, "AP/N")
    if ap[0] != "xref" and ap[1]:
        doc.xref_set_key(kid, "AP/N", re.sub(r"/(?!Off\b)[A-Za-z0-9_]+(?=\s*\d+ 0 R)", "/" + export, ap[1]))   # the "on" state is called Yes (PyMuPDF 1.x) or False (2.x)
    for key in ("T", "FT", "Ff", "V"):
        doc.xref_set_key(kid, key, "null")
    root = doc.pdf_catalog()
    fields, set_fields = _get_array(doc, root, "AcroForm/Fields")
    fields = [f for f in fields if f != kid]
    if parent is None:
        parent = doc.get_new_xref()
        doc.update_object(parent, "<< /FT /Btn /Ff 49152 /T %s /V /Off /Kids [%d 0 R] >>" % (fitz.get_pdf_str(group), kid))
        fields.append(parent)
    else:
        kids, set_kids = _get_array(doc, parent, "Kids")
        set_kids(kids + [kid])
    set_fields(fields)
    doc.xref_set_key(kid, "Parent", "%d 0 R" % parent)
    doc.xref_set_key(kid, "AS", "/Off")
    return parent


def read_table(path):
    """CSV or Excel file -> (header, rows) with all cells as strings"""
    ext = os.path.splitext(path)[1].lower()
    rows = []
    if os.path.getsize(path) > 50 * 1024 * 1024:
        raise ValueError(_tr("The file is too large."))
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb.worksheets[0]
            for r in ws.iter_rows(values_only=True):
                rows.append(["" if c is None else str(c) for c in r])
                if len(rows) > MAX_ROWS + 1:
                    break
        finally:
            wb.close()
    else:
        raw = open(path, "rb").read(30 * 1024 * 1024)
        for enc in ("utf-8-sig", "cp1250", "latin-1"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        try:
            dialect = csv.Sniffer().sniff(text[:4000], delimiters=",;\t|")
        except csv.Error:
            dialect = csv.excel
        for r in csv.reader(io.StringIO(text), dialect):
            rows.append(r)
            if len(rows) > MAX_ROWS + 1:
                break
    rows = [r for r in rows if any(c.strip() for c in r)]
    if len(rows) < 2:
        raise ValueError(_tr("The file needs a header row and at least one data row."))
    header = [c.strip() or "Column %d" % (i + 1) for i, c in enumerate(rows[0])]
    body = [(r + [""] * len(header))[:len(header)] for r in rows[1:MAX_ROWS + 1]]
    return header, body


def form_fields(doc):
    """unique form fields of the document: name -> (type, sample widget, [choices], [radio exports])"""
    out = {}
    for pg in doc:
        for w in pg.widgets() or []:
            e = out.setdefault(w.field_name, [w.field_type, w.field_type_string, list(w.choice_values or []), []])
            if w.field_type == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
                try:
                    e[3].append(w.on_state())
                except Exception:
                    pass
    return out


def fill_document(doc, values, flatten=False):
    """values: field name -> text.  Returns the number of fields that were filled."""
    n = 0
    for pg in doc:
        for w in list(pg.widgets() or []):
            if w.field_name not in values:
                continue
            v = values[w.field_name]
            t = w.field_type
            try:
                if t == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                    w.field_value = w.on_state() if str(v).strip().lower() in TRUE_WORDS else "Off"
                elif t == fitz.PDF_WIDGET_TYPE_RADIOBUTTON:
                    on = w.on_state()
                    same = re.sub(r"[^A-Za-z0-9_]", "_", str(v).strip()) == str(on).lstrip("/")
                    w.field_value = on if same else "Off"
                elif t in (fitz.PDF_WIDGET_TYPE_BUTTON, fitz.PDF_WIDGET_TYPE_SIGNATURE):
                    continue
                else:
                    w.field_value = str(v)
                w.update()
                n += 1
            except Exception:
                pass
    if flatten and n:
        try:
            doc.bake(annots=False, widgets=True)
        except Exception:
            pass
    return n


def unique_path(folder, name, ext):
    base = re.sub(r"\.{2,}", "_", safe_name(name)).strip(" .")[:80] or "file"
    p = os.path.join(folder, base + ext)
    k = 2
    while os.path.exists(p):
        p = os.path.join(folder, "%s_%d%s" % (base, k, ext))
        k += 1
    return p


def list_pdfs(folder, sub=True):
    out = []
    for root, dirs, files in os.walk(folder):
        dirs[:] = [d for d in dirs if not d.startswith(".")]
        for f in sorted(files):
            if f.lower().endswith(".pdf"):
                out.append(os.path.join(root, f))
                if len(out) >= MAX_FILES:
                    return out
        if not sub:
            break
    return out


def search_pdfs(folder, query, sub=True, case=False, whole=False, progress=None, stats=None):
    """yield (path, page_no, snippet) for every page that contains `query`"""
    stats = stats if stats is not None else {}
    files = list_pdfs(folder, sub)
    stats.update(files=len(files), skipped=0)
    needle = query if case else query.lower()
    rx = re.compile(r"\b" + re.escape(query) + r"\b", 0 if case else re.I) if whole else None
    found = 0
    for k, p in enumerate(files):
        if progress and progress(k, len(files), os.path.basename(p)):
            break
        try:
            if os.path.getsize(p) > MAX_FILE_MB * 1024 * 1024:
                stats["skipped"] += 1
                continue
            doc = fitz.open(p)
            if doc.needs_pass:
                stats["skipped"] += 1
                doc.close()
                continue
            try:
                for pno in range(len(doc)):
                    text = doc[pno].get_text()
                    if rx is not None:
                        m = rx.search(text)
                        pos = m.start() if m else -1
                    else:
                        pos = (text if case else text.lower()).find(needle)
                    if pos >= 0:
                        a = max(0, pos - 45)
                        yield p, pno, re.sub(r"\s+", " ", text[a:pos + len(query) + 70]).strip()
                        found += 1
                        if found >= MAX_HITS:
                            return
            finally:
                doc.close()
        except Exception:
            stats["skipped"] += 1


# ====================================================================================== batch actions
def _wm_doc(doc, text, opacity):
    img = make_watermark_image(text, (0.8, 0.1, 0.1), opacity, 45)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    data = buf.getvalue()
    ratio = img.height / img.width
    for pg in doc:
        v = pg.rect
        w = v.width * 0.6
        h = w * ratio
        c = fitz.Point(v.x0 + v.width / 2, v.y0 + v.height / 2)
        pg.insert_image(fitz.Rect(c.x - w / 2, c.y - h / 2, c.x + w / 2, c.y + h / 2), stream=data, overlay=True)


def batch_pdf(path, action, opts):
    """process one PDF; returns (bytes_or_None, extension, suffix)"""
    data = open(path, "rb").read()
    doc = fitz.open(stream=data, filetype="pdf")
    if doc.needs_pass:
        raise ValueError(_tr("The file is password protected."))
    if action == "compress":
        out = compress_bytes(data, opts["level"], opts.get("strip", False))
        return out, ".pdf", "_compressed"
    if action == "watermark":
        _wm_doc(doc, opts["text"], opts["opacity"])
        return doc.tobytes(garbage=3, deflate=True), ".pdf", "_watermark"
    if action == "rotate":
        for pg in doc:
            pg.set_rotation((pg.rotation + opts["angle"]) % 360)
        return doc.tobytes(garbage=3, deflate=True), ".pdf", "_rotated"
    if action == "clean":
        doc.scrub(metadata=True, xml_metadata=True, javascript=True, thumbnails=True, hidden_text=True, redactions=False,
                  clean_pages=False, reset_fields=False, reset_responses=False, attached_files=False, embedded_files=False,
                  remove_links=False, redact_images=0)
        return doc.tobytes(garbage=3, deflate=True), ".pdf", "_clean"
    if action == "numbers":
        n = len(doc)
        for i, pg in enumerate(doc):
            v = pg.rect
            t = "%d / %d" % (i + 1, n)
            EX.put_text(pg, v.x0 + v.width / 2, v.y1 - 28, t, 10, (0.2, 0.2, 0.2), EX.font_kwargs(t), "center")
        return doc.tobytes(garbage=3, deflate=True), ".pdf", "_numbered"
    if action == "ocr":
        td = EX.tessdata_dir()
        if not td:
            raise ValueError(_tr("OCR needs the free Tesseract program."))
        new = EX.ocr_document(doc, list(range(len(doc))), opts["lang"], td, 250, True, None)
        return new.tobytes(garbage=3, deflate=True), ".pdf", "_ocr"
    raise ValueError("unknown action")


def batch_convert(path, kind, outdir):
    doc = fitz.open(path)
    if doc.needs_pass:
        raise ValueError(_tr("The file is password protected."))
    base = re.sub(r"\.{2,}", "_", safe_name(os.path.splitext(os.path.basename(path))[0])).strip(" .")[:80] or "file"
    pages = list(range(len(doc)))
    if kind == "word":
        EX.pdf_to_docx(doc, pages, unique_path(outdir, base, ".docx"))
    elif kind == "text":
        EX.pdf_to_text(doc, pages, unique_path(outdir, base, ".txt"))
    elif kind == "png":
        sub = os.path.join(outdir, base)
        os.makedirs(sub, exist_ok=True)
        EX.pdf_to_images(doc, pages, "PNG", 150, sub, base)
    doc.close()


# ====================================================================================== the mixin
class V3bTools:
    # ------------------------------------------------------------------ form tools
    def v3_form_combo(self, rect):
        r = self.ask_text("Drop-down list", "One choice per line:", initial="Option 1\nOption 2\nOption 3", height=6)
        if r is None:
            return
        choices = [x.strip() for x in r[0].splitlines() if x.strip()][:200]
        if not choices:
            return
        pg = self.page()
        n = sum(1 for _ in (pg.widgets() or [])) + 1
        w = fitz.Widget()
        w.field_type = fitz.PDF_WIDGET_TYPE_COMBOBOX
        w.field_name = "choice_%d_%d" % (self.page_no + 1, n)
        w.choice_values = choices
        w.field_value = choices[0]
        w.rect = rect
        w.text_fontsize = 11
        w.fill_color = (0.93, 0.96, 1)
        w.border_color = (0.2, 0.2, 0.2)
        w.border_width = 1
        w.text_color = (0, 0, 0)
        self._run_edit(lambda: pg.add_widget(w), msg=_tr("Drop-down list added."))

    def v3_form_radio(self, pt):
        groups = sorted({n for n, e in form_fields(self.doc).items() if e[0] == fitz.PDF_WIDGET_TYPE_RADIOBUTTON})
        d = self._dlg("Radio button")
        grp = tk.StringVar(value=groups[-1] if groups else "group1")
        opt = tk.StringVar(value="Option")
        self._label(d, "Group (buttons of one group exclude each other)").grid(row=0, column=0, sticky="w")
        ttk.Combobox(d, textvariable=grp, values=groups, width=28).grid(row=1, column=0, sticky="w", pady=3)
        self._label(d, "Value of this button").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(d, textvariable=opt, width=30).grid(row=3, column=0, sticky="w", pady=3)

        def go():
            g, o = grp.get().strip(), opt.get().strip()
            if not g or not o:
                return
            d.destroy()
            pg = self.page()
            rect = fitz.Rect(pt.x - 7, pt.y - 7, pt.x + 7, pt.y + 7)
            self._run_edit(lambda: radio_add(self.doc, pg, g, o, rect), msg=_tr("Radio button added."))
        self._buttons(d, "Add", go, 4, colspan=1)

    # ------------------------------------------------------------------ form field editor
    def dlg_field_editor(self):
        if not self.need_doc():
            return
        self.commit_floats()
        d = self._dlg("Form fields")
        tree = ttk.Treeview(d, columns=("n", "name", "type", "page", "value"), show="headings", height=14, selectmode="browse")
        for c, t, w in (("n", "#", 36), ("name", "Name", 190), ("type", "Field type", 120), ("page", "Page", 64), ("value", "Value", 170)):
            tree.heading(c, text=_tr(t))
            tree.column(c, width=S(w), anchor="w" if c in ("name", "value", "type") else "center")
        tree.grid(row=0, column=0, rowspan=9, padx=(0, 10))
        rows = {}

        def refill(select=None):
            tree.delete(*tree.get_children())
            rows.clear()
            k = 0
            for pno, pg in enumerate(self.doc):
                for w in pg.widgets() or []:
                    k += 1
                    iid = "w%d" % w.xref
                    rows[iid] = (pno, w.xref)
                    tree.insert("", "end", iid=iid, values=(k, w.field_name, _tr(FIELD_KINDS.get(w.field_type, "Field")), pno + 1,
                                                          str(w.field_value if w.field_value not in (None, False) else "")[:40]))
            if select and select in rows:
                tree.selection_set(select)
                tree.see(select)

        def current():
            s = tree.selection()
            return (s[0], rows[s[0]]) if s else (None, None)

        def rename():
            iid, r = current()
            if not r:
                return
            pno, xref = r
            w = widget_by_xref(self.doc[pno], xref)
            old = w.field_name
            a = self.ask_text("Rename field", "New name:", initial=old, height=1)
            if a is None:
                return
            new = a[0].strip().replace("\n", " ")[:80]
            if not new or new == old:
                return
            others = {x for x in form_fields(self.doc)} - {old}
            if new in others:
                messagebox.showinfo(APP, "There is already a field with that name.", parent=d)
                return
            top = field_top(self.doc, xref)
            self._run_edit(lambda: self.doc.xref_set_key(top, "T", fitz.get_pdf_str(new)), thumbs=[], msg=_tr("Field renamed."))
            refill(iid)

        def delete():
            iid, r = current()
            if not r:
                return
            pno, xref = r
            pg = self.doc[pno]
            w = widget_by_xref(pg, xref)
            if w is not None and messagebox.askyesno(APP, "Delete this form field?", parent=d):
                self._run_edit(lambda: pg.delete_widget(w), msg=_tr("Field deleted."))
                refill()

        def options():
            iid, r = current()
            if not r:
                return
            pno, xref = r
            pg = self.doc[pno]
            w = widget_by_xref(pg, xref)
            if w.field_type not in (fitz.PDF_WIDGET_TYPE_COMBOBOX, fitz.PDF_WIDGET_TYPE_LISTBOX):
                messagebox.showinfo(APP, "Only drop-down lists and lists have choices.", parent=d)
                return
            a = self.ask_text("Choices", "One choice per line:", initial="\n".join(str(x) for x in (w.choice_values or [])), height=8)
            if a is None:
                return
            ch = [x.strip() for x in a[0].splitlines() if x.strip()][:200]

            def work():
                ww = widget_by_xref(pg, xref)
                ww.choice_values = ch
                if ww.field_value not in ch:
                    ww.field_value = ch[0] if ch else ""
                ww.update()
            self._run_edit(work, msg=_tr("Choices changed."))
            refill(iid)

        def move(step):
            iid, r = current()
            if not r:
                return
            pno, xref = r
            pg = self.doc[pno]

            def work():
                refs, setter = _get_array(self.doc, pg.xref, "Annots")
                if xref not in refs:
                    return
                i = refs.index(xref)
                j = i + step
                # swap with the next form field on the same page in the requested direction
                wx = {w.xref for w in pg.widgets() or []}
                while 0 <= j < len(refs) and refs[j] not in wx:
                    j += step
                if 0 <= j < len(refs):
                    refs[i], refs[j] = refs[j], refs[i]
                    setter(refs)
            self._run_edit(work, thumbs=[], msg=_tr("Tab order changed."))
            refill(iid)

        def auto_order():
            def work():
                for pg in self.doc:
                    ws = list(pg.widgets() or [])
                    if len(ws) < 2:
                        continue
                    refs, setter = _get_array(self.doc, pg.xref, "Annots")
                    pos = {w.xref: (round(w.rect.y0 / 6), w.rect.x0) for w in ws}
                    slots = [i for i, x in enumerate(refs) if x in pos]
                    ordered = sorted((refs[i] for i in slots), key=lambda x: pos[x])
                    for i, x in zip(slots, ordered):
                        refs[i] = x
                    setter(refs)
            self._run_edit(work, thumbs=[], msg=_tr("Tab order set: top to bottom, left to right."))
            refill()

        def show():
            iid, r = current()
            if r:
                d.destroy()
                self.goto(r[0])
        for i, (t, c) in enumerate((("Rename...", rename), ("Choices...", options), ("Delete", delete), ("Move up", lambda: move(-1)),
                                    ("Move down", lambda: move(1)), ("Order top to bottom", auto_order), ("Go to page", show))):
            ttk.Button(d, text=t, command=c, width=20).grid(row=i, column=1, pady=2, sticky="n")
        self._label(d, "The order of the list is the Tab-key order in other PDF programs.", fg=TH.TERTIARY).grid(row=9, column=0, columnspan=2, sticky="w", pady=(6, 0))
        refill()
        if not rows:
            self._label(d, "This document has no form fields. Use Sign > Fill & create forms to add some.", fg=TH.SECOND).grid(row=10, column=0, columnspan=2, sticky="w")
        ttk.Button(d, text="Close", command=d.destroy).grid(row=11, column=1, pady=(10, 0), sticky="e")

    # ------------------------------------------------------------------ fill from CSV / Excel
    def dlg_fill_from_data(self):
        if not self.need_doc():
            return
        self.commit_floats()
        fields = form_fields(self.doc)
        if not fields:
            messagebox.showinfo(APP, "This document has no form fields to fill. Add fields first (Sign > Fill & create forms).")
            return
        p = filedialog.askopenfilename(title=_tr("Choose the data file (CSV or Excel)"),
                                       filetypes=[("CSV / Excel", "*.csv *.txt *.xlsx *.xlsm"), ("All", "*.*")])
        if not p:
            return
        try:
            header, rows = read_table(p)
        except Exception as ex:
            messagebox.showerror(APP, _tr("Could not read the data file:") + f"\n{ex}")
            return
        d = self._dlg("Fill forms from a table")
        self._label(d, "{} rows found. One PDF will be made for every row.".replace("{}", str(len(rows))), bold=True).grid(row=0, column=0, columnspan=3, sticky="w")
        self._label(d, "Form field", bold=True).grid(row=1, column=0, sticky="w", pady=(8, 2))
        self._label(d, "Column", bold=True).grid(row=1, column=1, sticky="w", pady=(8, 2))
        none = _tr("(leave empty)")
        maps = {}
        low = {h.strip().lower(): h for h in header}
        names = sorted(fields)
        box = tk.Frame(d, bg=BG)
        box.grid(row=2, column=0, columnspan=3, sticky="w")
        cv = tk.Canvas(box, width=S(430), height=S(min(260, 30 * len(names) + 8)), bg=BG, highlightthickness=0)
        sb = ttk.Scrollbar(box, orient="vertical", command=cv.yview)
        inner = tk.Frame(cv, bg=BG)
        cv.create_window(0, 0, window=inner, anchor="nw")
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side="left")
        if len(names) * 30 > 260:
            sb.pack(side="left", fill="y")
        for i, n in enumerate(names):
            tk.Label(inner, text=f"{n}  ({_tr(FIELD_KINDS.get(fields[n][0], 'Field'))})"[:48], bg=BG, fg=TH.LABEL, font=(UI_FONT, 9), anchor="w",
                     width=34).grid(row=i, column=0, sticky="w", pady=2)
            v = tk.StringVar(value=low.get(n.strip().lower(), none))
            ttk.Combobox(inner, textvariable=v, values=[none] + header, state="readonly", width=24).grid(row=i, column=1, pady=2)
            maps[n] = v
        inner.update_idletasks()
        cv.configure(scrollregion=cv.bbox("all"))
        self._label(d, "File names").grid(row=3, column=0, sticky="w", pady=(10, 2))
        namecol = tk.StringVar(value=header[0])
        ttk.Combobox(d, textvariable=namecol, values=header + [_tr("(numbers only)")], state="readonly", width=24).grid(row=3, column=1, sticky="w", pady=(10, 2))
        flat = tk.BooleanVar(value=False)
        tk.Checkbutton(d, text="Flatten (the filled values can no longer be edited)", variable=flat, bg=BG, activebackground=BG,
                       fg=TH.LABEL, selectcolor=TH.CARD, font=(UI_FONT, 9)).grid(row=4, column=0, columnspan=3, sticky="w")

        def go():
            mapping = {n: v.get() for n, v in maps.items() if v.get() != none}
            if not mapping:
                messagebox.showinfo(APP, "Match at least one form field with a column.", parent=d)
                return
            outdir = filedialog.askdirectory(parent=d, title=_tr("Choose the folder for the filled PDFs"))
            if not outdir:
                return
            nc = namecol.get()
            d.destroy()
            self.do_fill_from_data(header, rows, mapping, nc if nc in header else None, outdir, flat.get())
        self._buttons(d, "Create PDFs", go, 5, colspan=3)

    def do_fill_from_data(self, header, rows, mapping, namecol, outdir, flatten):
        self.commit_floats()
        template = self.doc.tobytes(encryption=fitz.PDF_ENCRYPT_NONE)
        idx = {h: i for i, h in enumerate(header)}
        upd, close = self._progress("Creating PDFs...", len(rows))
        made = 0
        try:
            for k, row in enumerate(rows):
                upd(k, len(rows), _tr("Creating PDF {} of {}...").replace("{}", str(k + 1), 1).replace("{}", str(len(rows)), 1))
                doc = fitz.open(stream=template, filetype="pdf")
                fill_document(doc, {n: row[idx[c]] for n, c in mapping.items()}, flatten)
                base = ("%03d_%s" % (k + 1, row[idx[namecol]])) if namecol else "%03d" % (k + 1)
                path = unique_path(outdir, base, ".pdf")
                doc.save(path, garbage=3, deflate=True)
                doc.close()
                made += 1
        except Exception as ex:
            close()
            messagebox.showerror(APP, f"{ex}")
            return
        close()
        self.status.config(text=_tr("{} PDF files were created.").replace("{}", str(made)))
        messagebox.showinfo(APP, _tr("{} PDF files were created.").replace("{}", str(made)) + "\n" + outdir)

    # ------------------------------------------------------------------ batch processing of a folder
    def dlg_batch(self):
        d = self._dlg("Batch processing")
        src, dst = tk.StringVar(), tk.StringVar()
        sub = tk.BooleanVar(value=False)
        for r, (t, var) in enumerate((("Folder with PDF files", src), ("Folder for the results", dst))):
            self._label(d, t).grid(row=r, column=0, sticky="w", pady=3)
            ttk.Entry(d, textvariable=var, width=44).grid(row=r, column=1, pady=3, padx=4)
            ttk.Button(d, text="Browse...", command=lambda v=var: v.set(filedialog.askdirectory(parent=d) or v.get())).grid(row=r, column=2)
        tk.Checkbutton(d, text="Include sub-folders", variable=sub, bg=BG, activebackground=BG, fg=TH.LABEL, selectcolor=TH.CARD,
                       font=(UI_FONT, 9)).grid(row=2, column=1, sticky="w")
        actions = (("compress", "Compress"), ("watermark", "Add a text watermark"), ("rotate", "Rotate all pages"),
                   ("ocr", "Make searchable (OCR)"), ("clean", "Remove hidden data (metadata)"), ("numbers", "Add page numbers"),
                   ("word", "Convert to Word"), ("text", "Convert to text"), ("png", "Convert to PNG pictures"))
        act = tk.StringVar(value=_tr(actions[0][1]))
        self._label(d, "What to do", bold=True).grid(row=3, column=0, sticky="w", pady=(10, 3))
        ttk.Combobox(d, textvariable=act, values=[_tr(n) for _k, n in actions], state="readonly", width=34).grid(row=3, column=1, sticky="w", pady=(10, 3), padx=4)
        opt = tk.Frame(d, bg=BG)
        opt.grid(row=4, column=0, columnspan=3, sticky="w")
        level = tk.StringVar(value="Medium  (recommended)")
        wtext, wop = tk.StringVar(value="CONFIDENTIAL"), tk.IntVar(value=30)
        angle = tk.StringVar(value="90")
        langs = EX.ocr_languages(EX.tessdata_dir()) if EX.tessdata_dir() else []
        lang = tk.StringVar(value="eng" if "eng" in langs else (langs[0] if langs else ""))

        def show_opts(*_a):
            for w in opt.winfo_children():
                w.destroy()
            k = {_tr(n): key for key, n in actions}.get(act.get(), "compress")
            if k == "compress":
                self._label(opt, "Compression").grid(row=0, column=0, sticky="w", pady=3)
                ttk.Combobox(opt, textvariable=level, values=list(COMPRESS_LEVELS), state="readonly", width=24).grid(row=0, column=1, padx=4)
            elif k == "watermark":
                self._label(opt, "Text").grid(row=0, column=0, sticky="w", pady=3)
                ttk.Entry(opt, textvariable=wtext, width=24).grid(row=0, column=1, padx=4)
                self._label(opt, "Opacity %").grid(row=0, column=2, sticky="w")
                ttk.Spinbox(opt, from_=5, to=100, width=5, textvariable=wop).grid(row=0, column=3, padx=4)
            elif k == "rotate":
                self._label(opt, "Turn by (degrees)").grid(row=0, column=0, sticky="w", pady=3)
                ttk.Combobox(opt, textvariable=angle, values=["90", "180", "270"], state="readonly", width=6).grid(row=0, column=1, padx=4)
            elif k == "ocr":
                if langs:
                    self._label(opt, "Language").grid(row=0, column=0, sticky="w", pady=3)
                    ttk.Combobox(opt, textvariable=lang, values=langs, state="readonly", width=10).grid(row=0, column=1, padx=4)
                else:
                    self._label(opt, "OCR needs the free Tesseract program (see Document > OCR).", fg=TH.TERTIARY).grid(row=0, column=0)
        act.trace_add("write", show_opts)
        show_opts()
        self._label(d, "Your original files are never changed. Results are saved as new files.", fg=TH.TERTIARY).grid(row=5, column=0, columnspan=3, sticky="w", pady=(8, 0))

        def go():
            s, o = src.get().strip(), dst.get().strip()
            if not os.path.isdir(s):
                messagebox.showinfo(APP, "Choose the folder with the PDF files.", parent=d)
                return
            if not o:
                o = os.path.join(s, "Claudius output")
            k = {_tr(n): key for key, n in actions}.get(act.get(), "compress")
            opts = {"level": level.get(), "text": wtext.get().strip() or "CONFIDENTIAL", "opacity": max(5, min(100, wop.get())),
                    "angle": int(angle.get() or 90), "lang": lang.get(), "strip": False}
            if k == "ocr" and not langs:
                messagebox.showinfo(APP, "OCR needs the free Tesseract program (see Document > OCR).", parent=d)
                return
            d.destroy()
            self.do_batch(s, o, sub.get(), k, opts)
        self._buttons(d, "Start", go, 6, colspan=3)

    def do_batch(self, src, dst, sub, action, opts):
        files = [f for f in list_pdfs(src, sub) if os.path.abspath(os.path.dirname(f)) != os.path.abspath(dst) or not sub]
        files = [f for f in files if not os.path.abspath(f).startswith(os.path.abspath(dst) + os.sep)]
        if not files:
            messagebox.showinfo(APP, "There are no PDF files in that folder.")
            return
        os.makedirs(dst, exist_ok=True)
        upd, close = self._progress("Batch processing...", len(files))
        report, ok = [], 0
        try:
            for k, f in enumerate(files):
                name = os.path.basename(f)
                upd(k, len(files), f"{name}  ({k + 1}/{len(files)})")
                try:
                    if os.path.getsize(f) > MAX_FILE_MB * 1024 * 1024:
                        raise ValueError(_tr("The file is too large."))
                    if action in ("word", "text", "png"):
                        batch_convert(f, action, dst)
                    else:
                        data, ext, suffix = batch_pdf(f, action, opts)
                        with open(unique_path(dst, os.path.splitext(name)[0] + suffix, ext), "wb") as fh:
                            fh.write(data)
                    ok += 1
                except Exception as ex:
                    report.append(f"{name}: {ex}")
                upd(k + 1, len(files))
        finally:
            close()
        msg = _tr("{} of {} files done.").replace("{}", str(ok), 1).replace("{}", str(len(files)), 1) + "\n" + dst
        if report:
            msg += "\n\n" + _tr("Problems:") + "\n" + "\n".join(report[:12])
        messagebox.showinfo(APP, msg)
        self.status.config(text=_tr("Batch finished."))

    # ------------------------------------------------------------------ smart compress
    def dlg_smart_compress(self):
        if not self.need_doc():
            return
        self.commit_floats()
        data = self.doc.tobytes(encryption=fitz.PDF_ENCRYPT_NONE)
        if len(data) > 150 * 1024 * 1024:
            messagebox.showinfo(APP, "This document is very large. Use Document > Compress instead.")
            return
        d = self._dlg("Smart compress")
        self._label(d, "Current size: {}".replace("{}", human_size(len(data))), bold=True).grid(row=0, column=0, columnspan=3, sticky="w")
        variants = [("clean", "Lossless (clean up only)", None)] + [(k, k, v) for k, v in COMPRESS_LEVELS.items()] + [("custom", "Custom...", None)]
        results = {}
        tree = ttk.Treeview(d, columns=("opt", "size", "save"), show="headings", height=5, selectmode="browse")
        for c, t, w in (("opt", "Quality", 210), ("size", "New size", 110), ("save", "Saved", 100)):
            tree.heading(c, text=_tr(t))
            tree.column(c, width=S(w), anchor="w" if c == "opt" else "e")
        tree.grid(row=1, column=0, columnspan=3, pady=6)
        cust = {"dpi": tk.IntVar(value=130), "q": tk.IntVar(value=65)}
        cf = tk.Frame(d, bg=BG)
        cf.grid(row=2, column=0, columnspan=3, sticky="w")
        self._label(cf, "Picture resolution (dpi)").pack(side="left")
        ttk.Spinbox(cf, from_=50, to=300, width=5, textvariable=cust["dpi"]).pack(side="left", padx=4)
        self._label(cf, "JPEG quality").pack(side="left", padx=(8, 0))
        ttk.Spinbox(cf, from_=20, to=95, width=5, textvariable=cust["q"]).pack(side="left", padx=4)
        ttk.Button(cf, text="Calculate", command=lambda: calc("custom")).pack(side="left", padx=6)

        def make(key):
            doc = fitz.open(stream=data, filetype="pdf")
            if key == "clean":
                pass
            elif key == "custom":
                dpi = max(50, min(300, int(cust["dpi"].get())))
                doc.rewrite_images(dpi_threshold=int(dpi * 1.3), dpi_target=dpi, quality=max(20, min(95, int(cust["q"].get()))))
            else:
                thr, target, q = COMPRESS_LEVELS[key]
                doc.rewrite_images(dpi_threshold=thr, dpi_target=target, quality=q)
            out = doc.tobytes(garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, clean=True)
            doc.close()
            return out

        def calc(key):
            try:
                results[key] = make(key)
            except Exception as ex:
                results[key] = None
                messagebox.showerror(APP, f"{ex}", parent=d)
            show_row(key)

        def show_row(key):
            r = results.get(key)
            label = next(_tr(t) for k, t, _v in variants if k == key)
            vals = (label, human_size(len(r)) if r else "-", ("%d %%" % round((1 - len(r) / len(data)) * 100)) if r else "")
            if tree.exists(key):
                tree.item(key, values=vals)
            else:
                tree.insert("", "end", iid=key, values=vals)

        upd, close = self._progress("Calculating...", len(variants) - 1)
        try:
            for i, (k, _t, _v) in enumerate(variants[:-1]):
                upd(i, len(variants) - 1)
                calc(k)
        finally:
            close()
        tree.insert("", "end", iid="custom", values=(_tr("Custom..."), "-", ""))
        tree.selection_set("clean" if not results.get("Medium  (recommended)") else "Medium  (recommended)")
        prev = tk.Frame(d, bg=BG)
        prev.grid(row=3, column=0, columnspan=3, pady=6)
        lb_a = tk.Label(prev, bg="white", bd=1, relief="solid")
        lb_b = tk.Label(prev, bg="white", bd=1, relief="solid")
        lb_a.grid(row=1, column=0, padx=6)
        lb_b.grid(row=1, column=1, padx=6)
        self._label(prev, "Before", bold=True).grid(row=0, column=0)
        self._label(prev, "After", bold=True).grid(row=0, column=1)
        keep = {}

        def render_one(src, label):
            doc = fitz.open(stream=src, filetype="pdf")
            pg = doc[min(self.page_no, len(doc) - 1)]
            z = min(S(250) / pg.rect.width, S(250) / pg.rect.height) * 1.6
            pix = pg.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
            im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            im.thumbnail((S(250), S(250)))
            keep[label] = ImageTk.PhotoImage(im, master=d)
            doc.close()
            return keep[label]

        def refresh(_e=None):
            s = tree.selection()
            if not s:
                return
            if s[0] == "custom" and "custom" not in results:
                lb_b.config(image="")
                return
            try:
                lb_a.config(image=render_one(data, "a"))
                if results.get(s[0]):
                    lb_b.config(image=render_one(results[s[0]], "b"))
            except Exception:
                pass
        tree.bind("<<TreeviewSelect>>", refresh)
        refresh()
        self._label(d, "The preview shows the current page at screen size.", fg=TH.TERTIARY).grid(row=4, column=0, columnspan=3, sticky="w")

        def apply(mode):
            s = tree.selection()
            if not s or not results.get(s[0]):
                messagebox.showinfo(APP, "Choose an option with a calculated size.", parent=d)
                return
            new = results[s[0]]
            if mode == "save":
                base = os.path.splitext(os.path.basename(self.path))[0] if self.path else "document"
                out = filedialog.asksaveasfilename(parent=d, defaultextension=".pdf", initialfile=f"{base}_compressed.pdf",
                                                   filetypes=[("PDF files", "*.pdf")])
                if not out:
                    return
                with open(out, "wb") as fh:
                    fh.write(new)
                d.destroy()
                self.status.config(text=_tr("Saved {}.").replace("{}", human_size(len(new))))
            else:
                d.destroy()

                def work():
                    self.doc = fitz.open(stream=new, filetype="pdf")
                self._run_edit(work, rebuild=True, msg=_tr("Compressed: {} -> {}").replace("{}", human_size(len(data)), 1).replace("{}", human_size(len(new)), 1))
        f = tk.Frame(d, bg=BG)
        f.grid(row=5, column=0, columnspan=3, sticky="e", pady=(10, 0))
        ttk.Button(f, text="Apply to this document", command=lambda: apply("replace")).pack(side="left", padx=4)
        ttk.Button(f, text="Save as new file...", command=lambda: apply("save")).pack(side="left", padx=4)
        ttk.Button(f, text="Cancel", command=d.destroy).pack(side="left")

    # ------------------------------------------------------------------ merge with page picker
    def dlg_merge_pick(self):
        d = self._dlg("Merge with page picker")
        items = []                                     # [path, pages(list)|None for all, total]
        self._label(d, "Add files, then choose the pages you want from each one.", fg=TH.TERTIARY).grid(row=0, column=0, columnspan=2, sticky="w")
        tree = ttk.Treeview(d, columns=("file", "pages", "n"), show="headings", height=9, selectmode="browse")
        for c, t, w in (("file", "File", 260), ("pages", "Pages used", 110), ("n", "Count", 60)):
            tree.heading(c, text=_tr(t))
            tree.column(c, width=S(w), anchor="w" if c == "file" else "center")
        tree.grid(row=1, column=0, rowspan=8, pady=4)
        use_cur = tk.BooleanVar(value=False)
        if self.doc is not None:
            tk.Checkbutton(d, text="Put the open document first", variable=use_cur, bg=BG, activebackground=BG, fg=TH.LABEL,
                           selectcolor=TH.CARD, font=(UI_FONT, 9)).grid(row=9, column=0, sticky="w")
            use_cur.set(True)

        def refill(sel=None):
            tree.delete(*tree.get_children())
            for i, (p, pages, total) in enumerate(items):
                shown = _tr("all") if pages is None else pages_to_text(pages)
                tree.insert("", "end", iid=str(i), values=(os.path.basename(p), shown[:30], total if pages is None else len(pages)))
            if sel is not None and items:
                tree.selection_set(str(max(0, min(sel, len(items) - 1))))

        def cur():
            s = tree.selection()
            return int(s[0]) if s else None

        def add():
            for p in filedialog.askopenfilenames(parent=d, filetypes=[("PDF files", "*.pdf")]):
                try:
                    x = fitz.open(p)
                    if x.needs_pass:
                        messagebox.showinfo(APP, _tr("Skipped (password protected):") + " " + os.path.basename(p), parent=d)
                        continue
                    items.append([p, None, len(x)])
                    x.close()
                except Exception as ex:
                    messagebox.showerror(APP, f"{os.path.basename(p)}: {ex}", parent=d)
            refill()

        def pick():
            i = cur()
            if i is None:
                return
            res = self.pick_pages_dialog(items[i][0], items[i][1])
            if res is not None:
                items[i][1] = None if len(res) == items[i][2] else res
                refill(i)

        def typed():
            i = cur()
            if i is None:
                return
            a = self.ask_text("Pages", _tr("Pages of {} (e.g. all, 1-3, 5, 8-):").replace("{}", os.path.basename(items[i][0])), initial="all", height=1)
            if a is None:
                return
            try:
                pg = parse_pages(a[0], items[i][2])
            except ValueError as ex:
                messagebox.showwarning(APP, str(ex), parent=d)
                return
            items[i][1] = None if len(pg) == items[i][2] else pg
            refill(i)

        def move(step):
            i = cur()
            if i is None or not 0 <= i + step < len(items):
                return
            items[i], items[i + step] = items[i + step], items[i]
            refill(i + step)

        def remove():
            i = cur()
            if i is not None:
                items.pop(i)
                refill(i)
        for r, (t, c) in enumerate((("Add files...", add), ("Pick pages (pictures)...", pick), ("Type page numbers...", typed),
                                    ("Up", lambda: move(-1)), ("Down", lambda: move(1)), ("Remove", remove))):
            ttk.Button(d, text=t, command=c, width=22).grid(row=1 + r, column=1, padx=(8, 0), pady=1, sticky="n")
        mode = tk.StringVar(value="open")
        tk.Radiobutton(d, text="Open the result here", variable=mode, value="open", bg=BG, activebackground=BG, fg=TH.LABEL,
                       selectcolor=TH.CARD).grid(row=10, column=0, sticky="w")
        tk.Radiobutton(d, text="Save the result as a new file...", variable=mode, value="save", bg=BG, activebackground=BG, fg=TH.LABEL,
                       selectcolor=TH.CARD).grid(row=11, column=0, sticky="w")

        def go():
            n_src = len(items) + (1 if self.doc is not None and use_cur.get() else 0)
            if n_src < 1 or (n_src < 2 and not any(it[1] is not None for it in items)):
                messagebox.showinfo(APP, "Add at least two files, or pick pages from one file.", parent=d)
                return
            out = None
            if mode.get() == "save":
                out = filedialog.asksaveasfilename(parent=d, defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], initialfile="merged.pdf")
                if not out:
                    return
            d.destroy()
            self.do_merge_pick(items, out, self.doc is not None and use_cur.get())
        self._buttons(d, "Merge", go, 12, colspan=2)

    def do_merge_pick(self, items, out, with_current):
        new = fitz.open()
        if with_current:
            self.commit_floats()
            new.insert_pdf(self.doc)
        for p, pages, total in items:
            src = fitz.open(p)
            for a, b in (runs(pages) if pages is not None else [(0, len(src) - 1)]):
                new.insert_pdf(src, from_page=a, to_page=b)
            src.close()
        if len(new) == 0:
            return
        if out:
            new.save(out, garbage=3, deflate=True)
            self.status.config(text=_tr("Merged {} pages into {}").replace("{}", str(len(new)), 1).replace("{}", out, 1))
        else:
            self._adopt_doc(new, None, modified=True)
            self.status.config(text=_tr("Merged document with {} pages.").replace("{}", str(len(new))))

    def pick_pages_dialog(self, path, current):
        """thumbnail grid: click pages to select.  Returns the chosen page list (0-based) or None."""
        try:
            src = fitz.open(path)
        except Exception as ex:
            messagebox.showerror(APP, f"{ex}")
            return None
        n = min(len(src), 400)
        chosen = set(range(len(src))) if current is None else set(current)
        d = self._dlg(os.path.basename(path))
        d.resizable(True, True)
        self._label(d, "Click a page to add or remove it. Selected pages have a blue frame.", fg=TH.TERTIARY).grid(row=0, column=0, columnspan=3, sticky="w")
        W, H = S(110), S(150)
        cols = 6
        cv = tk.Canvas(d, width=cols * (W + 14) + 8, height=min(S(430), (H + 34) * 3), bg=BG, highlightthickness=0)
        sb = ttk.Scrollbar(d, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=sb.set)
        cv.grid(row=1, column=0, columnspan=2)
        sb.grid(row=1, column=2, sticky="ns")
        photos, frames = [], {}
        rows = (n + cols - 1) // cols
        cv.configure(scrollregion=(0, 0, cols * (W + 14), rows * (H + 34) + 6))

        def paint(i):
            x0, y0, x1, y1 = frames[i][0]
            cv.itemconfigure(frames[i][1], outline=ACCENT if i in chosen else TH.LINE, width=4 if i in chosen else 1)

        upd, close = self._progress("Loading pages...", n)
        try:
            for i in range(n):
                pg = src[i]
                z = min(W / pg.rect.width, H / pg.rect.height)
                pix = pg.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
                ph = ImageTk.PhotoImage(Image.frombytes("RGB", (pix.width, pix.height), pix.samples), master=d)
                photos.append(ph)
                cx, cy = (i % cols) * (W + 14) + 8, (i // cols) * (H + 34) + 4
                cv.create_image(cx + W // 2, cy + H // 2, image=ph)
                rect = cv.create_rectangle(cx - 3, cy - 3, cx + W + 3, cy + H + 3, outline=TH.LINE, width=1)
                cv.create_text(cx + W // 2, cy + H + 14, text=str(i + 1), fill=TH.LABEL, font=(UI_FONT, 8))
                frames[i] = ((cx, cy, cx + W, cy + H), rect)
                if i % 12 == 0:
                    upd(i)
        finally:
            close()
        cv._photos = photos
        for i in frames:
            paint(i)

        def click(e):
            x, y = cv.canvasx(e.x), cv.canvasy(e.y)
            for i, ((a, b, c, dd), _r) in frames.items():
                if a - 3 <= x <= c + 3 and b - 3 <= y <= dd + 3:
                    chosen.symmetric_difference_update({i})
                    paint(i)
                    return
        cv.bind("<Button-1>", click)
        cv.bind("<MouseWheel>", lambda e: cv.yview_scroll(-1 if e.delta > 0 else 1, "units"))
        result = {"v": None}

        def setall(on):
            chosen.clear()
            if on:
                chosen.update(range(len(src)))
            for i in frames:
                paint(i)

        def ok():
            if not chosen:
                messagebox.showinfo(APP, "Select at least one page.", parent=d)
                return
            result["v"] = sorted(chosen)
            d.destroy()
        f = tk.Frame(d, bg=BG)
        f.grid(row=2, column=0, columnspan=3, sticky="e", pady=(10, 0))
        ttk.Button(f, text="Select all", command=lambda: setall(True)).pack(side="left", padx=2)
        ttk.Button(f, text="Select none", command=lambda: setall(False)).pack(side="left", padx=2)
        ttk.Button(f, text="OK", command=ok).pack(side="left", padx=(12, 2))
        ttk.Button(f, text="Cancel", command=d.destroy).pack(side="left")
        d.wait_window()
        src.close()
        return result["v"]

    # ------------------------------------------------------------------ search in all PDFs of a folder
    def dlg_folder_search(self):
        d = self._dlg("Search in a folder")
        d.resizable(True, True)
        folder, query = tk.StringVar(value=os.path.dirname(self.path) if self.path else ""), tk.StringVar()
        sub, case, whole = tk.BooleanVar(value=True), tk.BooleanVar(value=False), tk.BooleanVar(value=False)
        self._label(d, "Folder").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(d, textvariable=folder, width=52).grid(row=0, column=1, padx=4)
        ttk.Button(d, text="Browse...", command=lambda: folder.set(filedialog.askdirectory(parent=d) or folder.get())).grid(row=0, column=2)
        self._label(d, "Search for").grid(row=1, column=0, sticky="w", pady=3)
        e = ttk.Entry(d, textvariable=query, width=52)
        e.grid(row=1, column=1, padx=4)
        e.focus_set()
        opts = tk.Frame(d, bg=BG)
        opts.grid(row=2, column=1, sticky="w")
        for t, v in (("Include sub-folders", sub), ("Match case", case), ("Whole words", whole)):
            tk.Checkbutton(opts, text=t, variable=v, bg=BG, activebackground=BG, fg=TH.LABEL, selectcolor=TH.CARD,
                           font=(UI_FONT, 9)).pack(side="left", padx=(0, 10))
        tree = ttk.Treeview(d, columns=("file", "page", "text"), show="headings", height=14, selectmode="browse")
        for c, t, w in (("file", "File", 190), ("page", "Page", 50), ("text", "Text", 430)):
            tree.heading(c, text=_tr(t))
            tree.column(c, width=S(w), anchor="w" if c != "page" else "center")
        tree.grid(row=3, column=0, columnspan=3, pady=8, sticky="nsew")
        d.rowconfigure(3, weight=1)
        d.columnconfigure(1, weight=1)
        info = self._label(d, "", fg=TH.SECOND)
        info.grid(row=4, column=0, columnspan=3, sticky="w")
        hits, state = {}, {"stop": False, "running": False}

        def search():
            q = query.get().strip()
            f = folder.get().strip()
            if not q or not os.path.isdir(f):
                messagebox.showinfo(APP, "Choose a folder and type what to search for.", parent=d)
                return
            tree.delete(*tree.get_children())
            hits.clear()
            state.update(stop=False, running=True)
            stats = {}

            def prog(k, total, name):
                info.config(text=_tr("Searching {} of {}: {}").replace("{}", str(k + 1), 1).replace("{}", str(total), 1).replace("{}", name, 1))
                d.update()
                return state["stop"] or not d.winfo_exists()
            found = 0
            for p, pno, snippet in search_pdfs(f, q, sub.get(), case.get(), whole.get(), prog, stats):
                if not d.winfo_exists():
                    return
                iid = tree.insert("", "end", values=(os.path.relpath(p, f), pno + 1, snippet))
                hits[iid] = (p, pno)
                found += 1
            state["running"] = False
            if d.winfo_exists():
                msg = _tr("{} result(s) in {} file(s).").replace("{}", str(found), 1).replace("{}", str(stats.get("files", 0)), 1)
                if found >= MAX_HITS:
                    msg += " " + _tr("(showing the first {})").replace("{}", str(MAX_HITS))
                if stats.get("skipped"):
                    msg += " " + _tr("{} file(s) skipped (protected, damaged or too large).").replace("{}", str(stats["skipped"]))
                info.config(text=msg)

        def open_hit(_e=None):
            s = tree.selection()
            if not s:
                return
            p, pno = hits[s[0]]
            q = query.get().strip()
            self.open_path(p)
            if self.doc is not None:
                self.goto(min(pno, len(self.doc) - 1))
                try:
                    self.find_var.set(q)
                    self.show_find(); self.find_run()
                except Exception:
                    pass
        tree.bind("<Double-1>", open_hit)
        e.bind("<Return>", lambda ev: search())
        f = tk.Frame(d, bg=BG)
        f.grid(row=5, column=0, columnspan=3, sticky="e", pady=(6, 0))
        ttk.Button(f, text="Search", command=search).pack(side="left", padx=3)
        ttk.Button(f, text="Stop", command=lambda: state.update(stop=True)).pack(side="left", padx=3)
        ttk.Button(f, text="Open selected", command=open_hit).pack(side="left", padx=3)
        ttk.Button(f, text="Close", command=lambda: (state.update(stop=True), d.destroy())).pack(side="left", padx=3)
        d.protocol("WM_DELETE_WINDOW", lambda: (state.update(stop=True), d.destroy()))
