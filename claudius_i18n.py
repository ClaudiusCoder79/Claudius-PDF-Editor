"""
Claudius PDF Editor - interface languages.

The program text is written in English.  This module translates it on the way to the screen:
every Tk widget text / menu label / dialog title / message box goes through tr().
Translations live in claudius_lang.py   (English text -> (pl, de, es, it, fr)).

Texts with variable parts are written with {} in the table, e.g.  "Moved {} page(s)."
"""
import json
import locale
import os
import re
import sys
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog

LANGS = [("en", "English"), ("pl", "Polski"), ("de", "Deutsch"), ("es", "Español"), ("it", "Italiano"), ("fr", "Français")]
_ORDER = {"pl": 0, "de": 1, "es": 2, "it": 3, "fr": 4}
CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".claudius")
DEBUG = bool(os.environ.get("CLAUDIUS_I18N_DEBUG"))

current = "en"
MISSING = set()
_exact = {}
_templates = []            # (regex, translated template, literal length)
_cb_forward = {}           # English combo-box value -> shown text
_cb_reverse = {}           # shown text -> English value
_cb_vars = set()

# ---------------------------------------------------------------------------------- settings
def _settings_file():
    return os.path.join(CONFIG_DIR, "settings.json")


def load_setting():
    try:
        with open(_settings_file(), "r", encoding="utf-8") as fh:
            return json.load(fh).get("language", "auto")
    except Exception:
        return "auto"


def save_setting(code):
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        data = {}
        try:
            with open(_settings_file(), "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:
            pass
        data["language"] = code
        with open(_settings_file(), "w", encoding="utf-8") as fh:
            json.dump(data, fh)
    except Exception:
        pass


def system_language():
    """language of the Windows display / the user's locale, if we have a translation for it"""
    code = None
    if sys.platform == "win32":
        try:
            import ctypes
            lid = ctypes.windll.kernel32.GetUserDefaultUILanguage() & 0xFF
            code = {0x15: "pl", 0x07: "de", 0x0A: "es", 0x10: "it", 0x0C: "fr"}.get(lid)
        except Exception:
            code = None
    if code is None:
        try:
            loc = (locale.getlocale()[0] or os.environ.get("LANG", "") or "").lower()
            code = loc[:2]
        except Exception:
            code = None
    return code if code in _ORDER else "en"


# ---------------------------------------------------------------------------------- tables
def _compile(lang):
    global _exact, _templates
    _exact, _templates = {}, []
    if lang == "en":
        return
    try:
        from claudius_lang import T
    except ImportError:
        return
    idx = _ORDER[lang]
    for key, vals in T.items():
        key = key.strip()                       # tr() works on the text without surrounding blanks
        vals = tuple(v.strip() if isinstance(v, str) else v for v in vals)
        if idx >= len(vals) or not vals[idx]:
            continue
        val = vals[idx]
        if "{}" in key:
            parts = key.split("{}")
            rx = re.compile("(.+?)".join(re.escape(p) for p in parts), re.S)
            _templates.append((rx, val, sum(len(p) for p in parts)))
        else:
            _exact[key] = val
    _templates.sort(key=lambda t: -t[2])          # most specific pattern first


def set_language(code):
    global current
    current = code if code in _ORDER else "en"
    _compile(current)


def init(code=None):
    """call once, before any window exists"""
    if code is None:
        code = load_setting()
    if code == "auto":
        code = system_language()
    set_language(code)
    _install()
    return current


# ---------------------------------------------------------------------------------- translate
def _fill(tpl, groups):
    it = iter(groups)

    def sub(m):
        if m.group(1):
            i = int(m.group(1))
            return groups[i] if i < len(groups) else ""
        return next(it, "")
    return re.sub(r"\{(\d*)\}", sub, tpl)


def _tr_core(core):
    v = _exact.get(core)
    if v is not None:
        return v
    for rx, tpl, _n in _templates:
        m = rx.fullmatch(core)
        if m:
            return _fill(tpl, [_exact.get(g, g) for g in m.groups()])
    return None


def tr(s):
    if current == "en" or not isinstance(s, str) or not s:
        return s
    core = s.strip()
    if not core:
        return s
    lead, trail = s[:len(s) - len(s.lstrip())], s[len(s.rstrip()):]
    v = _tr_core(core)
    if v is None and "\n" in core:                      # multi-line messages: line by line
        out, hit = [], False
        for line in core.split("\n"):
            t = _tr_core(line.strip()) if line.strip() else line
            if t is not None and line.strip():
                hit = True
                out.append(t)
            else:
                out.append(line)
        if hit:
            v = "\n".join(out)
    if v is None:
        if DEBUG and re.search(r"[A-Za-z]{3}", core):
            MISSING.add(core)
        return s
    return lead + v + trail


# ---------------------------------------------------------------------------------- patches
_TEXT_KEYS = ("text", "label")
_installed = False


def _tr_opts(d):
    if not d:
        return d
    out = None
    for k in _TEXT_KEYS:
        v = d.get(k) if hasattr(d, "get") else None
        if isinstance(v, str):
            if out is None:
                out = dict(d)
            out[k] = tr(v)
    return out if out is not None else d


def _install():
    global _installed
    if _installed:
        return
    _installed = True

    orig_options = tk.Misc._options

    def _options(self, cnf, kw=None):
        if current != "en":
            cnf = _tr_opts(cnf) if isinstance(cnf, dict) else cnf
            kw = _tr_opts(kw)
        return orig_options(self, cnf, kw)
    tk.Misc._options = _options

    # window titles
    orig_title = tk.Wm.wm_title

    def wm_title(self, string=None):
        return orig_title(self, tr(string) if isinstance(string, str) else string)
    tk.Wm.wm_title = wm_title
    tk.Wm.title = wm_title

    # message boxes
    orig_show = messagebox._show

    def _show(title=None, message=None, _icon=None, _type=None, **options):
        return orig_show(tr(title), tr(message), _icon, _type, **options)
    messagebox._show = _show

    # simple dialogs
    orig_ask = simpledialog.askstring

    def askstring(title, prompt, **kw):
        return orig_ask(tr(title), tr(prompt), **kw)
    simpledialog.askstring = askstring

    # file dialogs
    def wrap_file(name):
        orig = getattr(filedialog, name)

        def f(**options):
            if "title" in options:
                options["title"] = tr(options["title"])
            if options.get("filetypes"):
                options["filetypes"] = [(tr(a), b) for a, b in options["filetypes"]]
            return orig(**options)
        setattr(filedialog, name, f)
    for n in ("askopenfilename", "askopenfilenames", "asksaveasfilename", "askdirectory"):
        wrap_file(n)

    # notebook tabs
    orig_nb_add = ttk.Notebook.add

    def nb_add(self, child, **kw):
        if "text" in kw:
            kw["text"] = tr(kw["text"])
        return orig_nb_add(self, child, **kw)
    ttk.Notebook.add = nb_add

    # read-only combo boxes: show translated entries, but the program keeps seeing the English value
    orig_cb_init = ttk.Combobox.__init__

    def cb_init(self, master=None, **kw):
        translate = current != "en" and kw.get("state") == "readonly" and kw.get("values")
        if translate:
            en_vals = list(kw["values"])
            shown = []
            for v in en_vals:
                t = tr(v) if isinstance(v, str) else v
                if isinstance(v, str):
                    _cb_forward[v] = t
                    _cb_reverse[t] = v
                shown.append(t)
            kw["values"] = shown
        orig_cb_init(self, master, **kw)
        var = kw.get("textvariable")
        if translate and var is not None:
            _cb_vars.add(str(var))
            cur = var.get()
            if cur in _cb_forward:
                orig_set(var, _cb_forward[cur])
    ttk.Combobox.__init__ = cb_init

    orig_get, orig_set = tk.StringVar.get, tk.StringVar.set

    def sv_get(self):
        v = orig_get(self)
        return _cb_reverse.get(v, v) if str(self) in _cb_vars else v

    def sv_set(self, value):
        if str(self) in _cb_vars and isinstance(value, str):
            value = _cb_forward.get(value, value)
        return orig_set(self, value)
    tk.StringVar.get = sv_get
    tk.StringVar.set = sv_set
