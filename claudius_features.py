"""
Claudius PDF Editor - document features (redaction patterns, Bates numbers, page labels, attachments,
comments list / export / import, compare, accessibility check, print preview).

Mixed into PdfEditor by claudius.py.  Pure helper functions are at the top so they can be tested without a window.
"""
import difflib
import io
import json
import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pymupdf as fitz
except ImportError:                      # older PyMuPDF
    import fitz
from PIL import Image, ImageChops, ImageFilter, ImageTk

import claudius_theme as TH
from claudius_i18n import tr as _tr
from claudius_extra import (APP, UI_FONT, BG, SIDEBAR_BG, S, parse_pages, pages_to_text, scrub_phrases, put_text, font_kwargs,
                            MAX_PIXELS)

# ===================================================================================== sensitive-data patterns
_EMAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}")
_PHONE = re.compile(r"(?<![\w/])(?:\+|00)?\d[\d \-().]{7,17}\d(?![\w/])")
_IBAN = re.compile(r"\b[A-Z]{2}\d{2}(?: ?[A-Z0-9]{4}){2,7}(?: ?[A-Z0-9]{1,4})?\b")
_CARD = re.compile(r"(?<![\d\-])\d(?:[ \-]?\d){12,18}(?![\d\-])")
_SSN = re.compile(r"(?<![\d\-])\d{3}-\d{2}-\d{4}(?![\d\-])")
_NIP = re.compile(r"(?<![\d\-])\d{3}-\d{3}-\d{2}-\d{2}(?![\d\-])|(?<![\d\-])\d{2,3}-\d{2}-\d{2}-\d{3}(?![\d\-])")
_PESEL = re.compile(r"(?<!\d)\d{11}(?!\d)")
_DATE_OF_BIRTH = re.compile(r"\b(?:0?[1-9]|[12]\d|3[01])[./-](?:0?[1-9]|1[0-2])[./-](?:19|20)\d{2}\b")

PATTERN_KINDS = (
    ("email", "E-mail addresses"),
    ("phone", "Phone numbers"),
    ("iban", "Bank account numbers (IBAN)"),
    ("card", "Payment card numbers"),
    ("ssn", "US Social Security numbers"),
    ("pesel", "Polish PESEL / NIP numbers"),
    ("date", "Dates (dd.mm.yyyy)"),
)


def luhn_ok(digits):
    s = [int(c) for c in digits if c.isdigit()]
    if not 13 <= len(s) <= 19:
        return False
    tot = 0
    for i, d in enumerate(reversed(s)):
        if i % 2:
            d *= 2
            if d > 9:
                d -= 9
        tot += d
    return tot % 10 == 0


def iban_ok(text):
    t = re.sub(r"\s", "", text).upper()
    if not 15 <= len(t) <= 34:
        return False
    t = t[4:] + t[:4]
    num = "".join(str(int(c, 36)) for c in t)
    return int(num) % 97 == 1


def pesel_ok(s):
    if len(s) != 11 or not s.isdigit():
        return False
    w = (1, 3, 7, 9, 1, 3, 7, 9, 1, 3)
    return (10 - sum(int(a) * b for a, b in zip(s, w)) % 10) % 10 == int(s[10])


def nip_ok(s):
    d = re.sub(r"\D", "", s)
    if len(d) != 10:
        return False
    w = (6, 5, 7, 2, 3, 4, 5, 6, 7)
    c = sum(int(a) * b for a, b in zip(d, w)) % 11
    return c != 10 and c == int(d[9])


def find_sensitive(text, kinds, custom=None):
    """-> list of (kind, matched string) found in text.  `kinds` is an iterable of keys from PATTERN_KINDS;
    `custom` an optional regular expression.  Numbers are validated (checksums) so ordinary figures are left alone."""
    found = []
    kinds = set(kinds)
    if "email" in kinds:
        found += [("email", m.group()) for m in _EMAIL.finditer(text)]
    if "iban" in kinds:
        found += [("iban", m.group()) for m in _IBAN.finditer(text) if iban_ok(m.group())]
    if "card" in kinds:
        found += [("card", m.group()) for m in _CARD.finditer(text) if luhn_ok(m.group())]
    if "ssn" in kinds:
        found += [("ssn", m.group()) for m in _SSN.finditer(text)]
    if "pesel" in kinds:
        found += [("pesel", m.group()) for m in _PESEL.finditer(text) if pesel_ok(m.group())]
        found += [("pesel", m.group()) for m in _NIP.finditer(text) if nip_ok(m.group())]
    if "phone" in kinds:
        taken = [m.span() for rx in (_EMAIL, _IBAN, _CARD, _SSN, _NIP, _PESEL) for m in rx.finditer(text)
                 if any(k == kk and s == m.group() for kk, s in found for k in ("email", "iban", "card", "ssn", "pesel"))]
        for m in _PHONE.finditer(text):
            s = m.group().strip(" -.")
            n = len(re.sub(r"\D", "", s))
            if not 9 <= n <= 15 or luhn_ok(s) or _DATE_OF_BIRTH.fullmatch(s):
                continue
            if re.fullmatch(r"\d+(?:[.,]\d+)?", s) or re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
                continue                                            # a plain number or ISO date, not a phone number
            if any(m.start() < b and a < m.end() for a, b in taken):
                continue
            found.append(("phone", s))
    if "date" in kinds:
        found += [("date", m.group()) for m in _DATE_OF_BIRTH.finditer(text)]
    if custom:
        try:
            rx = re.compile(custom)
        except re.error as ex:
            raise ValueError(f"Invalid pattern: {ex}")
        for m in rx.finditer(text):
            if m.group().strip():
                found.append(("custom", m.group()))
    seen, out = set(), []
    for k, s in found:
        if s not in seen:
            seen.add(s)
            out.append((k, s))
    return out


# ===================================================================================== page labels
def roman(n, upper=True):
    vals = ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
            (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"))
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out if upper else out.lower()


def letters(n, upper=True):
    s = ""
    n = max(1, n)
    k = (n - 1) // 26 + 1
    s = chr(ord("A") + (n - 1) % 26) * k
    return s if upper else s.lower()


LABEL_STYLES = (("D", "1, 2, 3"), ("r", "i, ii, iii"), ("R", "I, II, III"), ("a", "a, b, c"), ("A", "A, B, C"), ("", "(no number)"))


def label_for(style, first, idx, prefix=""):
    n = first + idx
    body = {"D": str(n), "r": roman(n, False), "R": roman(n, True), "a": letters(n, False), "A": letters(n, True), "": ""}[style]
    return prefix + body


def page_label_ranges(doc):
    """-> list of dict(startpage, prefix, style, firstpagenum) or []"""
    try:
        return list(doc.get_page_labels() or [])
    except Exception:
        return []


def page_label(doc, i):
    try:
        return doc[i].get_label() or ""
    except Exception:
        return ""


# ===================================================================================== annotations <-> dicts
_MARKUP = {"Highlight": "add_highlight_annot", "Underline": "add_underline_annot",
           "StrikeOut": "add_strikeout_annot", "Squiggly": "add_squiggly_annot"}
_SKIP = ("Link", "Widget", "Popup")
ANNOT_LABEL = {"Highlight": "Highlight", "Underline": "Underline", "StrikeOut": "Strikeout", "Squiggly": "Squiggly",
               "Text": "Sticky note", "FreeText": "Text box", "Square": "Rectangle", "Circle": "Ellipse", "Line": "Line",
               "Ink": "Pen", "Polygon": "Polygon", "PolyLine": "Polyline", "Stamp": "Stamp"}


def _pt(p):
    return [round(float(p[0]), 3), round(float(p[1]), 3)]


def annot_to_dict(a, pno):
    name = a.type[1]
    d = {"page": pno, "type": name, "rect": [round(float(v), 3) for v in a.rect],
         "content": a.info.get("content", ""), "author": a.info.get("title", ""), "subject": a.info.get("subject", ""),
         "created": a.info.get("creationDate", ""), "opacity": a.opacity}
    col = a.colors or {}
    if col.get("stroke"):
        d["stroke"] = list(col["stroke"])
    if col.get("fill"):
        d["fill"] = list(col["fill"])
    try:
        d["width"] = a.border.get("width", 1)
    except Exception:
        pass
    try:
        v = a.vertices
        if v:
            if isinstance(v[0], (list, tuple)) and v and len(v[0]) and isinstance(v[0][0], (list, tuple, fitz.Point)):
                d["vertices"] = [[_pt(p) for p in stroke] for stroke in v]
                d["nested"] = True
            else:
                d["vertices"] = [_pt(p) for p in v]
    except Exception:
        pass
    return d


def collect_annots(doc, pages=None):
    out = []
    for i in (pages if pages is not None else range(len(doc))):
        pg = doc[i]
        for a in pg.annots() or []:
            if a.type[1] in _SKIP:
                continue
            try:
                out.append(annot_to_dict(a, i))
            except Exception:
                pass
    return out


def apply_annots(doc, items, page_offset=0):
    """Create the described annotations in doc.  -> (created, skipped)"""
    created = skipped = 0
    budget = 400000                                 # total points allowed, so a hostile file cannot exhaust memory
    for d in items:
        try:
            v = d.get("vertices")
            if v:
                budget -= sum(len(x) if isinstance(x, list) and x and isinstance(x[0], list) else 1 for x in v)
                if budget < 0:
                    skipped += 1
                    continue
            pno = int(d["page"]) + page_offset
            if not 0 <= pno < len(doc):
                skipped += 1
                continue
            pg = doc[pno]
            t = d.get("type")
            r = fitz.Rect(d.get("rect", (0, 0, 10, 10)))
            v = d.get("vertices")
            a = None
            if t in _MARKUP and v:
                pts = [fitz.Point(p) for p in v]
                quads = [fitz.Quad(*pts[k:k + 4]) for k in range(0, len(pts) - 3, 4)]
                a = getattr(pg, _MARKUP[t])(quads=quads)
            elif t == "Text":
                a = pg.add_text_annot(r.tl, d.get("content", "") or " ")
            elif t == "FreeText":
                a = pg.add_freetext_annot(r, d.get("content", ""), fontsize=12, text_color=d.get("stroke") or (0, 0, 0))
            elif t == "Square":
                a = pg.add_rect_annot(r)
            elif t == "Circle":
                a = pg.add_circle_annot(r)
            elif t == "Line" and v and len(v) >= 2:
                a = pg.add_line_annot(fitz.Point(v[0]), fitz.Point(v[1]))
            elif t == "Ink" and v:
                strokes = v if d.get("nested") else [v]
                a = pg.add_ink_annot([[tuple(p) for p in s] for s in strokes])
            elif t == "Polygon" and v:
                a = pg.add_polygon_annot([fitz.Point(p) for p in v])
            elif t == "PolyLine" and v:
                a = pg.add_polyline_annot([fitz.Point(p) for p in v])
            if a is None:
                skipped += 1
                continue
            col = {}
            if d.get("stroke"):
                col["stroke"] = tuple(d["stroke"])
            if d.get("fill") and t in ("Square", "Circle", "Polygon", "FreeText"):
                col["fill"] = tuple(d["fill"])
            if col and t != "FreeText":
                a.set_colors(**col)
            if d.get("width") and t in ("Square", "Circle", "Line", "Ink", "Polygon", "PolyLine"):
                a.set_border(width=float(d["width"]))
            if d.get("opacity") is not None and 0 <= float(d["opacity"]) <= 1:
                a.set_opacity(float(d["opacity"]))
            a.set_info(content=d.get("content", ""), title=d.get("author", ""), subject=d.get("subject", ""))
            a.update()
            created += 1
        except Exception:
            skipped += 1
    return created, skipped


# ===================================================================================== compare
def _gray_page(page, size=None, zoom=1.0):
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    if size and img.size != size:
        img = img.resize(size, Image.LANCZOS)
    return img


def diff_images(img_a, img_b):
    """-> (overlay RGB image, removed_px, added_px).  Red = only in A, green = only in B, grey = same."""
    if img_b.size != img_a.size:
        img_b = img_b.resize(img_a.size, Image.LANCZOS)
    ga, gb = img_a.convert("L"), img_b.convert("L")
    dark_a = ga.point(lambda v: 255 if v < 150 else 0)
    dark_b = gb.point(lambda v: 255 if v < 150 else 0)
    removed = ImageChops.subtract(dark_a, dark_b.filter(ImageFilter.MaxFilter(5)))
    added = ImageChops.subtract(dark_b, dark_a.filter(ImageFilter.MaxFilter(5)))
    out = Image.new("RGB", img_a.size, (255, 255, 255))
    out.paste((175, 175, 175), mask=ImageChops.lighter(dark_a, dark_b))
    out.paste((215, 30, 30), mask=removed)
    out.paste((10, 150, 50), mask=added)
    rem = sum(1 for v in removed.getdata() if v)
    add = sum(1 for v in added.getdata() if v)
    return out, rem, add


def text_diff(words_a, words_b):
    """-> (removed_words, added_words) lists using a word level diff"""
    sm = difflib.SequenceMatcher(None, words_a, words_b, autojunk=False)
    rem, add = [], []
    for op, a0, a1, b0, b1 in sm.get_opcodes():
        if op in ("replace", "delete"):
            rem += words_a[a0:a1]
        if op in ("replace", "insert"):
            add += words_b[b0:b1]
    return rem, add


def compare_docs(a, b, progress=None, zoom=1.0):
    """Compare two documents page by page.  -> list of dict(page, words_removed, words_added, rem_px, add_px, overlay,
    removed_text, added_text, only_in)"""
    res = []
    n = max(len(a), len(b))
    for i in range(n):
        if progress:
            progress(i, n)
        row = {"page": i, "words_removed": 0, "words_added": 0, "rem_px": 0, "add_px": 0, "overlay": None,
               "removed_text": [], "added_text": [], "only_in": ""}
        if i >= len(a) or i >= len(b):
            row["only_in"] = "B" if i >= len(a) else "A"
            res.append(row)
            continue
        wa = [w[4] for w in a[i].get_text("words")]
        wb = [w[4] for w in b[i].get_text("words")]
        rem, add = text_diff(wa, wb)
        row.update(words_removed=len(rem), words_added=len(add), removed_text=rem, added_text=add)
        z = min(zoom, (MAX_PIXELS / max(1.0, a[i].rect.width * a[i].rect.height)) ** 0.5)
        ia = _gray_page(a[i], zoom=z)
        ib = _gray_page(b[i], size=ia.size, zoom=z)
        ov, rpx, apx = diff_images(ia, ib)
        row.update(rem_px=rpx, add_px=apx)
        if rpx + apx > 12 or rem or add:
            row["overlay"] = ov
        res.append(row)
    if progress:
        progress(n, n)
    return res


def build_compare_report(a_name, b_name, rows, a_doc):
    """-> new fitz.Document with a summary page and one page per changed page"""
    out = fitz.open()
    changed = [r for r in rows if r["overlay"] is not None or r["only_in"]]
    lines = [f"Comparison of two PDFs", "",
             f"A: {a_name}", f"B: {b_name}", "",
             f"Pages compared: {len(rows)}",
             f"Pages with differences: {len(changed)}", "",
             "Red = only in A (removed), green = only in B (added).", ""]
    for r in changed:
        if r["only_in"]:
            lines.append(f"Page {r['page'] + 1}: exists only in {r['only_in']}")
        else:
            lines.append(f"Page {r['page'] + 1}: {r['words_removed']} word(s) removed, {r['words_added']} added")
            if r["removed_text"]:
                lines.append("    removed: " + " ".join(r["removed_text"][:25]) + (" ..." if len(r["removed_text"]) > 25 else ""))
            if r["added_text"]:
                lines.append("    added:   " + " ".join(r["added_text"][:25]) + (" ..." if len(r["added_text"]) > 25 else ""))
    if not changed:
        lines.append("No differences found.")
    pg = None
    y = 0
    for ln in lines:
        wrapped = [ln[i:i + 95] for i in range(0, len(ln), 95)] or [""]
        for w in wrapped:
            if pg is None or y > 800:
                pg = out.new_page(width=595, height=842)
                y = 56
            pg.insert_text((48, y), w, fontsize=11 if ln != lines[0] else 16, fontname="helv", color=(0.1, 0.1, 0.1))
            y += 15 if ln != lines[0] else 24
    for r in changed:
        if r["overlay"] is None:
            continue
        ov = r["overlay"]
        w, h = ov.size
        pg = out.new_page(width=w, height=h + 30)
        buf = io.BytesIO()
        ov.save(buf, "PNG")
        pg.insert_image(fitz.Rect(0, 30, w, h + 30), stream=buf.getvalue())
        pg.insert_text((10, 20), f"Page {r['page'] + 1}   -{r['words_removed']} words / +{r['words_added']} words",
                       fontsize=12, fontname="helv", color=(0.1, 0.1, 0.1))
    return out


# ===================================================================================== accessibility
def accessibility_report(doc):
    """-> list of (status, key, message); status is 'ok', 'warn' or 'fail'."""
    rep = []
    meta = doc.metadata or {}
    cat = doc.pdf_catalog()

    def key(k):
        try:
            return doc.xref_get_key(cat, k)[1]
        except Exception:
            return ""
    if (meta.get("title") or "").strip():
        rep.append(("ok", "title", "The document has a title."))
    else:
        rep.append(("fail", "title", "The document has no title."))
    if "true" in doc.xref_get_key(cat, "ViewerPreferences")[1].replace("/DisplayDocTitle", "DisplayDocTitle").lower() \
            and "DisplayDocTitle" in doc.xref_get_key(cat, "ViewerPreferences")[1]:
        rep.append(("ok", "disptitle", "The title (not the file name) is shown in the window title."))
    else:
        rep.append(("warn", "disptitle", "The window will show the file name instead of the title."))
    if key("Lang") not in ("", "null"):
        rep.append(("ok", "lang", "The document language is set."))
    else:
        rep.append(("fail", "lang", "The document language is not set (screen readers need it)."))
    text_pages = sum(1 for p in doc if len(p.get_text("text").strip()) > 20)
    if text_pages == len(doc):
        rep.append(("ok", "text", "All pages contain real text."))
    elif text_pages == 0:
        rep.append(("fail", "text", "No page contains real text - this looks like a scan. Run OCR first."))
    else:
        rep.append(("warn", "text", f"{len(doc) - text_pages} page(s) contain no real text (scanned?). Run OCR."))
    if len(doc) > 5 and not doc.get_toc():
        rep.append(("warn", "toc", "A document of this length should have bookmarks."))
    elif len(doc) > 5:
        rep.append(("ok", "toc", "The document has bookmarks."))
    marked = key("MarkInfo")
    if "true" in marked.lower() and key("StructTreeRoot") not in ("", "null"):
        rep.append(("ok", "tags", "The document is tagged (has a structure tree)."))
    else:
        rep.append(("warn", "tags", "The document is not tagged. Claudius PDF Editor cannot create tags; export the source "
                                    "(Word, etc.) as a 'tagged PDF' for full accessibility."))
    nofield = 0
    for p in doc:
        for w in p.widgets() or []:
            if not (w.field_label or "").strip():
                nofield += 1
    if nofield:
        rep.append(("warn", "fields", f"{nofield} form field(s) have no description."))
    imgs = sum(len(p.get_images()) for p in doc)
    if imgs and "true" not in marked.lower():
        rep.append(("warn", "alt", f"{imgs} image(s) found; without tags they have no alternative text."))
    return rep


def fix_accessibility(doc, title=None, lang=None, disp_title=True):
    cat = doc.pdf_catalog()
    if title is not None:
        meta = dict(doc.metadata or {})
        meta["title"] = title
        doc.set_metadata(meta)
    if lang:
        doc.xref_set_key(cat, "Lang", fitz.get_pdf_str(lang))
    if disp_title:
        doc.xref_set_key(cat, "ViewerPreferences", "<</DisplayDocTitle true>>")


# ===================================================================================== the mixin
class FeatureTools:
    # ------------------------------------------------------------------ redaction patterns
    def dlg_redact_patterns(self):
        if not self.need_doc():
            return
        d = self._dlg("Redact sensitive data")
        self._label(d, "Find and black out in the whole document:", bold=True).grid(row=0, column=0, columnspan=3, sticky="w")
        vars_ = {}
        for r, (k, label) in enumerate(PATTERN_KINDS, start=1):
            vars_[k] = tk.BooleanVar(value=k in ("email", "phone", "iban", "card"))
            tk.Checkbutton(d, text=label, variable=vars_[k], bg=BG, activebackground=BG, fg=TH.LABEL, selectcolor=TH.FILL,
                           font=(UI_FONT, 9)).grid(row=r, column=0, columnspan=3, sticky="w")
        n = len(PATTERN_KINDS) + 1
        self._label(d, "Own pattern (regular expression, optional)").grid(row=n, column=0, columnspan=3, sticky="w", pady=(8, 0))
        cv = tk.StringVar()
        ttk.Entry(d, textvariable=cv, width=40).grid(row=n + 1, column=0, columnspan=3, sticky="w")
        self._label(d, "Numbers are checked (card, IBAN, PESEL, NIP checksums) so ordinary figures are left alone.",
                    fg=TH.TERTIARY).grid(row=n + 2, column=0, columnspan=3, sticky="w", pady=(6, 0))

        def go():
            kinds = [k for k, v in vars_.items() if v.get()]
            custom = cv.get().strip()
            if not kinds and not custom:
                return
            try:
                find_sensitive("", kinds, custom)
            except ValueError as ex:
                messagebox.showwarning(APP, str(ex), parent=d)
                return
            d.destroy()
            self.redact_patterns(kinds, custom)
        self._buttons(d, "Find", go, n + 3, colspan=3)

    def redact_patterns(self, kinds, custom=""):
        """Scan, show what was found, ask, then permanently remove (with metadata scrub)."""
        hits, words, per_kind = [], set(), {}
        for i, pg in enumerate(self.doc):
            for k, s in find_sensitive(pg.get_text("text"), kinds, custom):
                rects = pg.search_for(s)
                if not rects and " " in s:                     # line breaks inside the number
                    rects = pg.search_for(s.replace(" ", ""))
                for rc in rects:
                    hits.append((i, rc))
                    per_kind[k] = per_kind.get(k, 0) + 1
                if rects:
                    words.add(s)
        if not hits:
            messagebox.showinfo(APP, "Nothing found.")
            return 0
        names = dict(PATTERN_KINDS)
        names["custom"] = "Own pattern"
        summary = "\n".join(f"  {_tr(names[k])}: {c}" for k, c in per_kind.items())
        if not messagebox.askyesno(APP, f"{len(hits)} match(es) found on {len({p for p, _ in hits})} page(s):\n\n{summary}\n\n"
                                        "Permanently remove them? (This can be undone until you save.)"):
            return 0
        self.do_redact_hits(hits, sorted(words))
        return len(hits)

    # ------------------------------------------------------------------ Bates numbering
    def dlg_bates(self):
        if not self.need_doc():
            return
        d = self._dlg("Bates numbering")
        pre, suf = tk.StringVar(value="CASE-"), tk.StringVar(value="")
        start, digits, size, margin = tk.IntVar(value=1), tk.IntVar(value=6), tk.IntVar(value=10), tk.IntVar(value=24)
        pos = tk.StringVar(value="Bottom right")
        rows = (("Prefix", ttk.Entry(d, textvariable=pre, width=20)), ("Suffix", ttk.Entry(d, textvariable=suf, width=20)),
                ("First number", ttk.Spinbox(d, from_=0, to=99999999, width=10, textvariable=start)),
                ("Digits", ttk.Spinbox(d, from_=1, to=12, width=6, textvariable=digits)),
                ("Position", ttk.Combobox(d, textvariable=pos, state="readonly", width=18,
                                          values=["Bottom right", "Bottom center", "Bottom left", "Top right", "Top center", "Top left"])),
                ("Font size", ttk.Spinbox(d, from_=5, to=36, width=6, textvariable=size)),
                ("Margin (pt)", ttk.Spinbox(d, from_=5, to=150, width=6, textvariable=margin)))
        for r, (lab, w) in enumerate(rows):
            self._label(d, lab).grid(row=r, column=0, sticky="w", pady=3)
            w.grid(row=r, column=1, sticky="w", padx=4)
        prev = self._label(d, "", fg=TH.TERTIARY)
        prev.grid(row=8, column=0, columnspan=3, sticky="w", pady=(4, 0))
        rng = self._range_row(d, 9, "all")

        def label():
            try:
                return bates_text(pre.get(), start.get(), digits.get(), suf.get(), 0)
            except tk.TclError:
                return ""

        def upd(*_):
            prev.config(text=_tr("Example") + ": " + label())
        for v in (pre, suf, start, digits):
            v.trace_add("write", upd)
        upd()

        def params():
            return dict(prefix=pre.get(), suffix=suf.get(), start=int(start.get()), digits=int(digits.get()),
                        position=pos.get(), size=int(size.get()), margin=int(margin.get()))

        def go():
            pages = self._pages_or_warn(rng.get(), d)
            if pages is None:
                return
            p = params()
            d.destroy()
            self.do_bates(pages, **p)

        def many():
            p = params()
            d.destroy()
            self.bates_several_files(**p)
        f = self._buttons(d, "Apply", go, 11, colspan=3)
        ttk.Button(f, text="Several files...", command=many).pack(side="left", padx=4)

    def do_bates(self, pages, prefix="", suffix="", start=1, digits=6, position="Bottom right", size=10, margin=24):
        def work():
            apply_bates(self.doc, pages, prefix, suffix, start, digits, position, size, margin)
        return self._run_edit(work, thumbs=pages, msg=f"Bates numbers added to {len(pages)} page(s).")

    def bates_several_files(self, **p):
        files = filedialog.askopenfilenames(title=_tr("Choose PDFs in order"), filetypes=[("PDF files", "*.pdf")])
        if not files:
            return
        out = filedialog.askdirectory(title=_tr("Folder for the numbered copies"))
        if not out:
            return
        n = start = p["start"]
        upd, close = self._progress("Bates numbering", len(files))
        done = 0
        try:
            for k, f in enumerate(files):
                upd(k, len(files), os.path.basename(f))
                doc = fitz.open(f)
                if doc.needs_pass:
                    doc.close()
                    continue
                q = dict(p, start=n)
                apply_bates(doc, list(range(len(doc))), **q)
                n += len(doc)
                base, ext = os.path.splitext(os.path.basename(f))
                dst = os.path.join(out, base + "_bates" + ext)
                k = 1
                while os.path.exists(dst):                  # never overwrite an existing file
                    k += 1
                    dst = os.path.join(out, f"{base}_bates{k}{ext}")
                doc.save(dst, garbage=3, deflate=True)
                doc.close()
                done += 1
        except Exception as ex:
            close()
            messagebox.showerror(APP, f"That did not work:\n{ex}")
            return
        close()
        messagebox.showinfo(APP, f"{done} file(s) numbered from {bates_text(p['prefix'], start, p['digits'], p['suffix'], 0)} "
                                 f"to {bates_text(p['prefix'], n - 1, p['digits'], p['suffix'], 0)}.")
        self.reveal(out)

    # ------------------------------------------------------------------ page labels
    def dlg_page_labels(self):
        if not self.need_doc():
            return
        d = self._dlg("Page labels")
        self._label(d, "Name pages like i, ii, iii for the front matter and 1, 2, 3 for the body.", fg=TH.TERTIARY).grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 6))
        cols = ("start", "style", "prefix", "first")
        tv = ttk.Treeview(d, columns=cols, show="headings", height=6, selectmode="browse")
        for c, t, w in (("start", "From page", 80), ("style", "Style", 110), ("prefix", "Prefix", 90), ("first", "Start at", 70)):
            tv.heading(c, text=_tr(t))
            tv.column(c, width=S(w), anchor="w")
        tv.grid(row=1, column=0, columnspan=4, sticky="we")
        ranges = [dict(r) for r in page_label_ranges(self.doc)] or [dict(startpage=0, prefix="", style="D", firstpagenum=1)]
        styles = dict(LABEL_STYLES)

        def refill():
            tv.delete(*tv.get_children())
            for r in sorted(ranges, key=lambda r: r["startpage"]):
                tv.insert("", "end", values=(r["startpage"] + 1, styles.get(r.get("style", ""), "1, 2, 3"),
                                             r.get("prefix", ""), r.get("firstpagenum", 1)))
        refill()
        sv, stv, pv, fv = tk.IntVar(value=1), tk.StringVar(value="i, ii, iii"), tk.StringVar(), tk.IntVar(value=1)
        self._label(d, "From page").grid(row=2, column=0, sticky="w", pady=(8, 2))
        ttk.Spinbox(d, from_=1, to=len(self.doc), width=6, textvariable=sv).grid(row=2, column=1, sticky="w")
        self._label(d, "Style").grid(row=2, column=2, sticky="w", padx=(10, 0))
        ttk.Combobox(d, textvariable=stv, state="readonly", width=12, values=[t for _, t in LABEL_STYLES]).grid(row=2, column=3, sticky="w")
        self._label(d, "Prefix").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Entry(d, textvariable=pv, width=10).grid(row=3, column=1, sticky="w")
        self._label(d, "Start at").grid(row=3, column=2, sticky="w", padx=(10, 0))
        ttk.Spinbox(d, from_=1, to=99999, width=6, textvariable=fv).grid(row=3, column=3, sticky="w")

        def add():
            code = {t: c for c, t in LABEL_STYLES}[stv.get()]
            sp = max(1, min(int(sv.get()), len(self.doc))) - 1
            ranges[:] = [r for r in ranges if r["startpage"] != sp] + [dict(startpage=sp, prefix=pv.get(), style=code, firstpagenum=max(1, int(fv.get())))]
            if not any(r["startpage"] == 0 for r in ranges):
                ranges.append(dict(startpage=0, prefix="", style="D", firstpagenum=1))
            refill()

        def remove():
            s = tv.selection()
            if not s:
                return
            sp = int(tv.item(s[0], "values")[0]) - 1
            ranges[:] = [r for r in ranges if r["startpage"] != sp]
            refill()

        def clear():
            ranges.clear()
            refill()
        bf = tk.Frame(d, bg=BG)
        bf.grid(row=4, column=0, columnspan=4, sticky="w", pady=6)
        ttk.Button(bf, text="Add / replace", command=add).pack(side="left", padx=2)
        ttk.Button(bf, text="Remove", command=remove).pack(side="left", padx=2)
        ttk.Button(bf, text="Remove all labels", command=clear).pack(side="left", padx=2)

        def ok():
            d.destroy()
            self.set_page_labels(ranges)
        self._buttons(d, "OK", ok, 5)

    def set_page_labels(self, ranges):
        def work():
            self.doc.set_page_labels(sorted(ranges, key=lambda r: r["startpage"]) if ranges else [])
        return self._run_edit(work, msg="Page labels updated.")

    # ------------------------------------------------------------------ attachments
    def dlg_attachments(self):
        if not self.need_doc():
            return
        d = self._dlg("Attachments")
        tv = ttk.Treeview(d, columns=("name", "size", "desc"), show="headings", height=8, selectmode="browse")
        for c, t, w in (("name", "File", 200), ("size", "Size", 80), ("desc", "Description", 180)):
            tv.heading(c, text=_tr(t))
            tv.column(c, width=S(w), anchor="w")
        tv.grid(row=0, column=0, columnspan=4, sticky="we")

        def refill():
            tv.delete(*tv.get_children())
            for nm in self.doc.embfile_names():
                inf = self.doc.embfile_info(nm)
                tv.insert("", "end", iid=nm, values=(inf.get("filename") or nm, _fmt_size(inf.get("size", 0)), inf.get("desc", "")))
        refill()

        def add():
            p = filedialog.askopenfilename(parent=d)
            if not p:
                return
            if os.path.getsize(p) > 100 * 1024 * 1024:
                messagebox.showwarning(APP, "That file is larger than 100 MB.", parent=d)
                return
            with open(p, "rb") as fh:
                data = fh.read()
            nm = os.path.basename(p)
            base, k = nm, 1
            while nm in self.doc.embfile_names():
                k += 1
                nm = f"{k}_{base}"
            self._run_edit(lambda: self.doc.embfile_add(nm, data, filename=base), msg=f"Attached {base}.")
            refill()

        def save():
            s = tv.selection()
            if not s:
                return
            inf = self.doc.embfile_info(s[0])
            fn = os.path.basename(inf.get("filename") or s[0]) or "attachment"
            p = filedialog.asksaveasfilename(parent=d, initialfile=fn)
            if p:
                with open(p, "wb") as fh:
                    fh.write(self.doc.embfile_get(s[0]))
                self.status.config(text=f"Saved {p}")

        def delete():
            s = tv.selection()
            if s:
                self._run_edit(lambda: self.doc.embfile_del(s[0]), msg="Attachment removed.")
                refill()
        bf = tk.Frame(d, bg=BG)
        bf.grid(row=1, column=0, columnspan=4, sticky="w", pady=8)
        ttk.Button(bf, text="Attach file...", command=add).pack(side="left", padx=2)
        ttk.Button(bf, text="Save as...", command=save).pack(side="left", padx=2)
        ttk.Button(bf, text="Remove", command=delete).pack(side="left", padx=2)
        ttk.Button(bf, text="Close", command=d.destroy).pack(side="left", padx=12)

    # ------------------------------------------------------------------ comments panel
    def _build_comments_panel(self, side):
        self.cm_wrap = tk.Frame(side, bg=SIDEBAR_BG)
        bf = tk.Frame(self.cm_wrap, bg=SIDEBAR_BG)
        bf.pack(side="bottom", fill="x", pady=4)
        for t, c in (("Export", self.export_comments), ("Import", self.import_comments), ("Summary", self.export_comment_summary)):
            ttk.Button(bf, text=t, command=c, width=7).pack(side="left", padx=1)
        self.cm_tree = ttk.Treeview(self.cm_wrap, columns=("pg", "type", "text"), show="headings", selectmode="browse")
        for c, t, w in (("pg", "Pg", 34), ("type", "Kind", 70), ("text", "Comment", 120)):
            self.cm_tree.heading(c, text=_tr(t))
            self.cm_tree.column(c, width=S(w), anchor="w", stretch=(c == "text"))
        sb = ttk.Scrollbar(self.cm_wrap, orient="vertical", command=self.cm_tree.yview)
        self.cm_tree.config(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.cm_tree.pack(side="left", fill="both", expand=True)
        self.cm_tree.bind("<<TreeviewSelect>>", self._cm_select)
        self.cm_tree.bind("<Delete>", lambda e: self.delete_comment())
        self.cm_items = {}

    def _refresh_comments(self):
        if not hasattr(self, "cm_tree"):
            return
        t = self.cm_tree
        t.delete(*t.get_children())
        self.cm_items = {}
        if not self.doc:
            return
        try:
            for i, pg in enumerate(self.doc):
                for a in pg.annots() or []:
                    nm = a.type[1]
                    if nm in _SKIP:
                        continue
                    txt = (a.info.get("content") or "").replace("\n", " ").strip()
                    if not txt and nm in _MARKUP:
                        try:
                            txt = pg.get_textbox(a.rect).strip().replace("\n", " ")[:80]
                        except Exception:
                            txt = ""
                    iid = t.insert("", "end", values=(i + 1, _tr(ANNOT_LABEL.get(nm, nm)), txt))
                    self.cm_items[iid] = (i, a.xref, fitz.Rect(a.rect))
        except Exception:
            pass

    def _cm_select(self, _e=None):
        s = self.cm_tree.selection()
        if not s or s[0] not in self.cm_items:
            return
        pno, _, rc = self.cm_items[s[0]]
        if pno != self.page_no:
            self.goto(pno)
        self.canvas.update_idletasks()
        try:
            y = self.off_y + rc.y0 * self.zoom
            total = max(1, int(float(self.canvas.cget("scrollregion").split()[3])))
            self.canvas.yview_moveto(max(0.0, (y - 60) / total))
        except Exception:
            pass

    def delete_comment(self):
        s = self.cm_tree.selection()
        if not s or s[0] not in self.cm_items:
            return
        pno, xref, _ = self.cm_items[s[0]]

        def work():
            pg = self.doc[pno]
            for a in pg.annots() or []:
                if a.xref == xref:
                    pg.delete_annot(a)
                    break
        self._run_edit(work, thumbs=[pno], msg="Comment deleted.")
        self._refresh_comments()

    def export_comments(self):
        if not self.need_doc():
            return
        items = collect_annots(self.doc)
        if not items:
            messagebox.showinfo(APP, "There are no comments or annotations in this document.")
            return
        base = os.path.splitext(os.path.basename(self.path or "document"))[0]
        p = filedialog.asksaveasfilename(defaultextension=".json", initialfile=base + "_comments.json",
                                         filetypes=[("Claudius PDF Editor comments", "*.json")])
        if not p:
            return
        with open(p, "w", encoding="utf-8") as fh:
            json.dump({"format": "claudius-annotations-1", "annotations": items}, fh, ensure_ascii=False, indent=1)
        self.status.config(text=f"Exported {len(items)} annotation(s) to {p}")

    def import_comments(self):
        if not self.need_doc():
            return
        p = filedialog.askopenfilename(filetypes=[("Comments file or PDF", "*.json *.pdf"), ("All files", "*.*")])
        if not p:
            return
        try:
            if p.lower().endswith(".pdf"):
                other = fitz.open(p)
                if other.needs_pass:
                    messagebox.showwarning(APP, "That PDF is password-protected.")
                    return
                items = collect_annots(other)
                other.close()
            else:
                with open(p, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
                items = data.get("annotations") if isinstance(data, dict) else None
                if not isinstance(items, list):
                    raise ValueError("This is not a Claudius PDF Editor comments file.")
                if len(items) > 20000:
                    raise ValueError("The file contains too many annotations.")
        except Exception as ex:
            messagebox.showerror(APP, f"Could not read the file:\n{ex}")
            return
        res = {}

        def work():
            res["r"] = apply_annots(self.doc, [i for i in items if isinstance(i, dict)])
        if self._run_edit(work, thumbs=list(range(len(self.doc))), msg=None):
            c, s = res["r"]
            self.status.config(text=f"Imported {c} annotation(s), skipped {s}." if s else f"Imported {c} annotation(s).")
            self._refresh_comments()

    def export_comment_summary(self):
        """A readable list of all comments, saved as a text file or as a new PDF."""
        if not self.need_doc():
            return
        items = collect_annots(self.doc)
        if not items:
            messagebox.showinfo(APP, "There are no comments or annotations in this document.")
            return
        base = os.path.splitext(os.path.basename(self.path or "document"))[0]
        p = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=base + "_comments.pdf",
                                         filetypes=[("PDF summary", "*.pdf"), ("Text file", "*.txt")])
        if not p:
            return
        lines = [f"Comments in {os.path.basename(self.path or 'document')}", ""]
        for it in items:
            who = f" ({it['author']})" if it.get("author") else ""
            txt = it.get("content") or ""
            if not txt and it["type"] in _MARKUP:
                try:
                    txt = self.doc[it["page"]].get_textbox(fitz.Rect(it["rect"])).strip().replace("\n", " ")
                except Exception:
                    txt = ""
            lines.append(f"Page {it['page'] + 1} - {it['type']}{who}: {txt}")
        if p.lower().endswith(".txt"):
            with open(p, "w", encoding="utf-8") as fh:
                fh.write("\n".join(lines))
        else:
            doc = fitz.open()
            pg, y = None, 0
            for ln in lines:
                for w in ([ln[i:i + 100] for i in range(0, len(ln), 100)] or [""]):
                    if pg is None or y > 800:
                        pg = doc.new_page(width=595, height=842)
                        y = 56
                    pg.insert_text((48, y), w, fontsize=10, **font_kwargs(w))
                    y += 14
            doc.save(p)
            doc.close()
        self.status.config(text=f"Summary saved: {p}")

    # ------------------------------------------------------------------ compare
    def dlg_compare(self):
        if not self.need_doc():
            return
        p = filedialog.askopenfilename(title=_tr("PDF to compare with the open document"), filetypes=[("PDF files", "*.pdf")])
        if not p:
            return
        try:
            other = fitz.open(p)
            if other.needs_pass:
                messagebox.showwarning(APP, "That PDF is password-protected.")
                return
        except Exception as ex:
            messagebox.showerror(APP, f"Could not open file:\n{ex}")
            return
        self.commit_floats()
        a = fitz.open(stream=self.doc.tobytes(encryption=fitz.PDF_ENCRYPT_NONE), filetype="pdf")
        upd, close = self._progress("Comparing...", len(a))
        try:
            rows = compare_docs(a, other, progress=lambda i, n: upd(i, n))
            rep = build_compare_report(os.path.basename(self.path or "open document"), os.path.basename(p), rows, a)
        except Exception as ex:
            close()
            messagebox.showerror(APP, f"That did not work:\n{ex}")
            return
        close()
        changed = sum(1 for r in rows if r["overlay"] is not None or r["only_in"])
        self._open_result(rep, "Comparison")
        self.status.config(text=f"{changed} of {len(rows)} page(s) differ.")

    # ------------------------------------------------------------------ accessibility
    def dlg_accessibility(self):
        if not self.need_doc():
            return
        d = self._dlg("Accessibility check")
        icons = {"ok": ("✔", "#1f9d55"), "warn": ("▲", "#d98a00"), "fail": ("✖", "#d93025")}
        body = tk.Frame(d, bg=BG)
        body.grid(row=0, column=0, sticky="w")

        def fill():
            for w in body.winfo_children():
                w.destroy()
            for st, key, msg in accessibility_report(self.doc):
                msg = _tr(msg)
                row = tk.Frame(body, bg=BG)
                row.pack(anchor="w", pady=2)
                tk.Label(row, text=icons[st][0], fg=icons[st][1], bg=BG, font=(UI_FONT, 11, "bold"), width=2).pack(side="left")
                tk.Label(row, text=msg, bg=BG, fg=TH.LABEL, font=(UI_FONT, 9), wraplength=S(430), justify="left").pack(side="left")
        fill()
        tv, lv = tk.StringVar(value=(self.doc.metadata or {}).get("title") or os.path.splitext(os.path.basename(self.path or ""))[0]), tk.StringVar(value="en-US")
        fr = tk.Frame(d, bg=BG)
        fr.grid(row=1, column=0, sticky="w", pady=(10, 0))
        self._label(fr, "Title").grid(row=0, column=0, sticky="w")
        ttk.Entry(fr, textvariable=tv, width=36).grid(row=0, column=1, padx=4, pady=2)
        self._label(fr, "Language").grid(row=1, column=0, sticky="w")
        ttk.Combobox(fr, textvariable=lv, width=10, values=["en-US", "en-GB", "pl-PL", "de-DE", "es-ES", "it-IT", "fr-FR"]).grid(row=1, column=1, sticky="w", padx=4)

        def fix():
            title, lang = tv.get().strip(), lv.get().strip()
            if not re.fullmatch(r"[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*", lang or "x"):
                messagebox.showwarning(APP, "Enter a language code such as en-US.", parent=d)
                return
            self._run_edit(lambda: fix_accessibility(self.doc, title or None, lang or None), msg="Document title and language set.")
            fill()

        def hint():
            messagebox.showinfo(APP, "To read a scanned document aloud or make it searchable, use Document > OCR.\n\n"
                                     "For a fully tagged PDF, export from the original program (Word, Writer, ...) with "
                                     "'Create tagged PDF' ticked.", parent=d)
        bf = tk.Frame(d, bg=BG)
        bf.grid(row=2, column=0, sticky="e", pady=(14, 0))
        ttk.Button(bf, text="Fix title and language", command=fix).pack(side="left", padx=4)
        ttk.Button(bf, text="Tips", command=hint).pack(side="left", padx=4)
        ttk.Button(bf, text="Close", command=d.destroy).pack(side="left", padx=4)

    # ------------------------------------------------------------------ print preview
    def print_preview(self, pages, fit="fit", on_print=None):
        if not self.doc or not pages:
            return
        self.commit_floats()
        d = tk.Toplevel(self)
        d.title(_tr("Print preview"))
        d.configure(bg=BG)
        d.transient(self)
        d.geometry(f"{S(560)}x{S(720)}+{self.winfo_rootx() + 120}+{self.winfo_rooty() + 40}")
        d.bind("<Escape>", lambda e: d.destroy())
        papers = {"A4": (595, 842), "Letter": (612, 792), "Legal": (612, 1008), "A3": (842, 1191), "A5": (420, 595)}
        top = tk.Frame(d, bg=BG)
        top.pack(fill="x", padx=10, pady=8)
        pv = tk.StringVar(value="A4")
        ttk.Combobox(top, textvariable=pv, values=list(papers), state="readonly", width=8).pack(side="left")
        idx = {"i": 0}
        info = tk.Label(top, bg=BG, fg=TH.LABEL, font=(UI_FONT, 9))
        cv = tk.Canvas(d, bg=TH.WORK, highlightthickness=0)
        photo = {}

        def draw(*_):
            cv.delete("all")
            cw, ch = max(cv.winfo_width(), 200), max(cv.winfo_height(), 200)
            pw, ph = papers[pv.get()]
            sc = min((cw - 40) / pw, (ch - 40) / ph)
            px0, py0 = (cw - pw * sc) / 2, (ch - ph * sc) / 2
            cv.create_rectangle(px0 + 3, py0 + 4, px0 + pw * sc + 3, py0 + ph * sc + 4, fill=TH.PAGE_SHADOW[0], outline="")
            cv.create_rectangle(px0, py0, px0 + pw * sc, py0 + ph * sc, fill="white", outline=TH.LINE)
            import claudius_print as cp
            page = self.doc[pages[idx["i"]]]
            mg = 20                                           # typical unprintable margin in points
            aw, ah = pw - 2 * mg, ph - 2 * mg
            rot, scale, box = cp.layout(page.rect.width, page.rect.height, aw, ah, fit, 72, 72)
            z = scale * sc
            pix = page.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            if rot:
                img = img.transpose(Image.ROTATE_90)
            photo["p"] = ImageTk.PhotoImage(img)
            cv.create_image(px0 + (mg + box[0]) * sc, py0 + (mg + box[1]) * sc, anchor="nw", image=photo["p"])
            info.config(text=f"{_tr('Sheet')} {idx['i'] + 1} / {len(pages)}   ({_tr('page')} {pages[idx['i']] + 1})")

        def step(k):
            idx["i"] = max(0, min(len(pages) - 1, idx["i"] + k))
            draw()
        ttk.Button(top, text="◀", width=3, command=lambda: step(-1)).pack(side="left", padx=(12, 2))
        ttk.Button(top, text="▶", width=3, command=lambda: step(1)).pack(side="left")
        info.pack(side="left", padx=10)
        if on_print:
            ttk.Button(top, text="Print...", command=lambda: (d.destroy(), on_print())).pack(side="right")
        ttk.Button(top, text="Close", command=d.destroy).pack(side="right", padx=6)
        cv.pack(fill="both", expand=True)
        pv.trace_add("write", draw)
        cv.bind("<Configure>", draw)
        d.after(50, draw)
        return d


def bates_text(prefix, number, digits, suffix, k):
    return f"{prefix}{int(number) + k:0{max(1, int(digits))}d}{suffix}"


def apply_bates(doc, pages, prefix="", suffix="", start=1, digits=6, position="Bottom right", size=10, margin=24,
                color=(0.1, 0.1, 0.1)):
    vert, _, horiz = position.lower().partition(" ")
    for k, i in enumerate(pages):
        pg = doc[i]
        v = pg.rect
        t = bates_text(prefix, start, digits, suffix, k)
        x = {"left": v.x0 + margin, "center": v.x0 + v.width / 2, "right": v.x1 - margin}[horiz]
        y = v.y1 - margin if vert == "bottom" else v.y0 + margin + size * 0.85
        put_text(pg, x, y, t, size, color, font_kwargs(t, True), horiz)


def _fmt_size(n):
    n = float(n or 0)
    for u in ("B", "KB", "MB", "GB"):
        if n < 1024 or u == "GB":
            return f"{n:.0f} {u}" if u == "B" else f"{n:.1f} {u}"
        n /= 1024
