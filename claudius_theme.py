"""
Claudius PDF Editor - iOS-style look: colours, rounded shapes, ttk styles.
Everything about the appearance lives here so it can be tuned in one place.
"""
import tkinter as tk
from tkinter import ttk
from tkinter import font as tkfont

UI_FONT = "Segoe UI"

# ---- iOS system colours: light and dark sets.  init() copies the chosen set into this module's globals,
#      so every other file just reads TH.BG, TH.ACCENT ... (a change of theme takes effect after a restart).
LIGHT = dict(
    ACCENT="#007AFF", ACCENT_HOVER="#0A84FF", ACCENT_DOWN="#0062CC",
    BG="#F2F2F7",            # grouped background (dialogs, sidebar)
    CARD="#FFFFFF",          # cards / ribbon / fields
    NAV_BG="#F9F9FB",        # navigation + status bars
    LINE="#D1D1D6", LINE_SOFT="#E5E5EA",
    FILL="#E9E9EE",          # quiet fills
    TINT="#E5F1FF", TINT_DOWN="#D3E7FF", HOVER="#EDEDF2",
    LABEL="#1C1C1E", SECOND="#6C6C70", TERTIARY="#8E8E93",
    WORK="#E5E5EA",          # the area behind the page
    SCROLL="#C1C1C6", SCROLL_HOT="#9A9AA0",
    PAGE_SHADOW=("#E0E0E6", "#D6D6DD", "#CBCBD3", "#C0C0C8"),
    KNOB_SHADOW="#D4D4DA", IS_DARK=False,
)
DARK = dict(
    ACCENT="#0A84FF", ACCENT_HOVER="#409CFF", ACCENT_DOWN="#0064D1",
    BG="#1C1C1E", CARD="#2C2C2E", NAV_BG="#242426",
    LINE="#48484A", LINE_SOFT="#38383A",
    FILL="#3A3A3C",
    TINT="#13324F", TINT_DOWN="#1B4673", HOVER="#3A3A3C",
    LABEL="#F2F2F7", SECOND="#AEAEB2", TERTIARY="#8E8E93",
    WORK="#121214",
    SCROLL="#5A5A5E", SCROLL_HOT="#7C7C80",
    PAGE_SHADOW=("#1A1A1C", "#161618", "#111113", "#0A0A0B"),
    KNOB_SHADOW="#0E0E10", IS_DARK=True,
)
NAMES = list(LIGHT)
mode = "light"


def _config_file():
    import os
    return os.path.join(os.path.expanduser("~"), ".claudius", "settings.json")


def load_setting():
    import json
    try:
        with open(_config_file(), "r", encoding="utf-8") as fh:
            return json.load(fh).get("theme", "auto")
    except Exception:
        return "auto"


def save_setting(code):
    import json, os
    try:
        os.makedirs(os.path.dirname(_config_file()), exist_ok=True)
        data = {}
        try:
            with open(_config_file(), "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:
            pass
        data["theme"] = code
        with open(_config_file(), "w", encoding="utf-8") as fh:
            json.dump(data, fh)
    except Exception:
        pass


def system_is_dark():
    """Windows 'Choose your default app mode' = Dark"""
    import sys
    if sys.platform != "win32":
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as k:
            return winreg.QueryValueEx(k, "AppsUseLightTheme")[0] == 0
    except Exception:
        return False


def apply(name):
    global mode
    import os
    name = os.environ.get("CLAUDIUS_THEME", name)         # (also handy for testing)
    dark = name == "dark" or (name == "auto" and system_is_dark())
    pal = DARK if dark else LIGHT
    globals().update(pal)
    mode = "dark" if dark else "light"
    return mode


def init(name=None):
    return apply(name if name else load_setting())


apply("light")           # module is usable even before init()


def rr_points(x0, y0, x1, y1, r):
    """polygon points of a rounded rectangle (use with smooth=True)"""
    r = max(0, min(r, (x1 - x0) / 2, (y1 - y0) / 2))
    return [x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r, x1, y1 - r, x1, y1, x1 - r, y1,
            x0 + r, y1, x0, y1, x0, y1 - r, x0, y0 + r, x0, y0]


def round_rect(canvas, x0, y0, x1, y1, r, **kw):
    return canvas.create_polygon(rr_points(x0, y0, x1, y1, r), smooth=True, **kw)


class PillTab(tk.Canvas):
    """A navigation tab: plain text, with a rounded capsule behind it when selected or hovered."""

    def __init__(self, parent, text, command, S, bg=None, font_size=10, bold_selected=True):
        bg = bg or NAV_BG
        self._S = S
        self._text = text
        self._font = tkfont.Font(family=UI_FONT, size=font_size)
        self._bold = tkfont.Font(family=UI_FONT, size=font_size, weight="bold")
        self._bold_selected = bold_selected
        from claudius_i18n import tr
        w = max(self._font.measure(tr(text)), self._bold.measure(tr(text))) + S(20)
        h = S(30)
        super().__init__(parent, width=w, height=h, bg=bg, highlightthickness=0, cursor="hand2")
        self._pw, self._ph = w, h
        self.selected = False
        self.hover = False
        self.command = command
        self.bind("<Button-1>", lambda e: self.command())
        self.bind("<Enter>", lambda e: self._set(hover=True))
        self.bind("<Leave>", lambda e: self._set(hover=False))
        self._draw()

    def _set(self, selected=None, hover=None):
        if selected is not None:
            self.selected = selected
        if hover is not None:
            self.hover = hover
        self._draw()

    def set_selected(self, on):
        self._set(selected=on)

    def _draw(self):
        S = self._S
        self.delete("all")
        if self.selected:
            round_rect(self, S(2), S(2), self._pw - S(2), self._ph - S(2), S(14), fill=TINT, outline="")
        elif self.hover:
            round_rect(self, S(2), S(2), self._pw - S(2), self._ph - S(2), S(14), fill=HOVER, outline="")
        font = self._bold if (self.selected and self._bold_selected) else self._font
        self.create_text(self._pw // 2, self._ph // 2, text=self._text, font=font,
                         fill=ACCENT if self.selected else LABEL)


def setup_styles(root, S):
    """ttk widgets in iOS clothes"""
    st = ttk.Style(root)
    try:
        st.theme_use("clam")
    except tk.TclError:
        pass
    root.option_add("*Font", (UI_FONT, 9))
    # classic tk widgets that were created without explicit colours
    for pat, val in (("*Foreground", LABEL), ("*Background", BG), ("*Entry.Background", CARD), ("*Text.Background", CARD),
                     ("*Listbox.Background", CARD), ("*Entry.insertBackground", LABEL), ("*Entry.selectBackground", TINT_DOWN),
                     ("*Entry.selectForeground", LABEL), ("*Checkbutton.selectColor", FILL), ("*Radiobutton.selectColor", FILL),
                     ("*Checkbutton.activeBackground", BG), ("*Radiobutton.activeBackground", BG),
                     ("*Checkbutton.highlightBackground", BG), ("*Radiobutton.highlightBackground", BG),
                     ("*Checkbutton.activeForeground", LABEL), ("*Radiobutton.activeForeground", LABEL)):
        root.option_add(pat, val)
    st.configure(".", background=BG, foreground=LABEL, font=(UI_FONT, 9), bordercolor=LINE,
                 lightcolor=BG, darkcolor=BG, troughcolor=LINE_SOFT, focuscolor=BG)
    st.configure("TFrame", background=BG)
    st.configure("TLabel", background=BG, foreground=LABEL)
    st.configure("TLabelframe", background=BG, bordercolor=LINE_SOFT)
    st.configure("TLabelframe.Label", background=BG, foreground=SECOND)
    st.configure("TCheckbutton", background=BG)
    st.configure("TRadiobutton", background=BG)

    # buttons: quiet tinted capsules with blue text
    st.configure("TButton", background=TINT, foreground=ACCENT, bordercolor=TINT, lightcolor=TINT,
                 darkcolor=TINT, relief="flat", borderwidth=0, padding=(S(14), S(5)), focuscolor=TINT,
                 font=(UI_FONT, 9, "bold"))
    st.map("TButton", background=[("pressed", TINT_DOWN), ("active", TINT_DOWN), ("disabled", FILL)],
           foreground=[("disabled", TERTIARY)], bordercolor=[("pressed", TINT_DOWN), ("active", TINT_DOWN)],
           lightcolor=[("active", TINT_DOWN)], darkcolor=[("active", TINT_DOWN)])
    st.configure("Toolbutton", background=BG, foreground=ACCENT)

    # text fields
    for name in ("TEntry", "TSpinbox", "TCombobox"):
        st.configure(name, fieldbackground=CARD, background=CARD, foreground=LABEL, bordercolor=LINE,
                     lightcolor=CARD, darkcolor=CARD, padding=S(4), insertcolor=ACCENT,
                     selectbackground=TINT_DOWN, selectforeground=LABEL, arrowcolor=ACCENT)
        st.map(name, bordercolor=[("focus", ACCENT)], lightcolor=[("focus", ACCENT)], darkcolor=[("focus", ACCENT)])
    st.map("TCombobox", fieldbackground=[("readonly", CARD)], background=[("readonly", CARD), ("active", CARD)],
           foreground=[("readonly", LABEL)], selectbackground=[("readonly", CARD)],
           selectforeground=[("readonly", LABEL)])
    root.option_add("*TCombobox*Listbox.background", CARD)
    root.option_add("*TCombobox*Listbox.foreground", LABEL)
    root.option_add("*TCombobox*Listbox.selectBackground", ACCENT)
    root.option_add("*TCombobox*Listbox.selectForeground", "white")
    root.option_add("*TCombobox*Listbox.font", (UI_FONT, 9))

    # slim scroll bars without arrows
    w = S(10)
    for orient, name in (("vertical", "Vertical.TScrollbar"), ("horizontal", "Horizontal.TScrollbar")):
        trough = "Vertical.Scrollbar.trough" if orient == "vertical" else "Horizontal.Scrollbar.trough"
        thumb = "Vertical.Scrollbar.thumb" if orient == "vertical" else "Horizontal.Scrollbar.thumb"
        try:
            st.layout(name, [(trough, {"sticky": "nswe", "children": [(thumb, {"expand": "1", "sticky": "nswe"})]})])
        except tk.TclError:
            pass
        st.configure(name, background=SCROLL, troughcolor=BG, bordercolor=BG, lightcolor=SCROLL,
                     darkcolor=SCROLL, gripcount=0, width=w, arrowsize=w, relief="flat", borderwidth=0)
        st.map(name, background=[("active", SCROLL_HOT), ("pressed", TERTIARY)],
               lightcolor=[("active", SCROLL_HOT)], darkcolor=[("active", SCROLL_HOT)])

    # zoom slider
    st.configure("Horizontal.TScale", background=BG, troughcolor=LINE, lightcolor=ACCENT, darkcolor=ACCENT,
                 bordercolor=LINE, sliderrelief="flat")
    st.map("Horizontal.TScale", background=[("active", ACCENT_HOVER)])
    st.configure("TScale", background=NAV_BG, troughcolor=LINE)

    # tabs inside dialogs (segmented-control feel)
    st.configure("TNotebook", background=BG, bordercolor=LINE_SOFT, lightcolor=BG, darkcolor=BG, tabmargins=(0, S(2), 0, 0))
    st.configure("TNotebook.Tab", background=FILL, foreground=SECOND, padding=(S(14), S(5)), bordercolor=LINE_SOFT,
                 lightcolor=FILL, darkcolor=FILL, font=(UI_FONT, 9))
    st.map("TNotebook.Tab", background=[("selected", CARD)], foreground=[("selected", ACCENT)],
           lightcolor=[("selected", CARD)], darkcolor=[("selected", CARD)], bordercolor=[("selected", LINE)])

    # lists
    st.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=LABEL, rowheight=S(24),
                 bordercolor=LINE_SOFT, lightcolor=CARD, darkcolor=CARD, borderwidth=0)
    st.map("Treeview", background=[("selected", TINT_DOWN)], foreground=[("selected", ACCENT)])
    st.configure("Treeview.Heading", background=NAV_BG, foreground=SECOND, relief="flat")
    st.configure("TProgressbar", background=ACCENT, troughcolor=LINE_SOFT, bordercolor=LINE_SOFT,
                 lightcolor=ACCENT, darkcolor=ACCENT)
    st.configure("TSeparator", background=LINE_SOFT)
    return st


class IOSSlider(tk.Canvas):
    """Thin blue slider with a round white knob (replaces ttk.Scale)."""

    def __init__(self, parent, S, from_=0, to=100, variable=None, length=150, command=None, bg=None):
        bg = bg or NAV_BG
        self._S = S
        self._len = length
        self._from, self._to = from_, to
        self._var = variable or tk.DoubleVar(value=from_)
        self._cmd = command
        super().__init__(parent, width=length, height=S(22), bg=bg, highlightthickness=0, cursor="hand2")
        self._pad = S(10)
        self._trace = self._var.trace_add("write", lambda *a: self._draw())
        self.bind("<Button-1>", self._press)
        self.bind("<B1-Motion>", self._press)
        self.bind("<Configure>", lambda e: self._draw())
        self._draw()

    def _value_to_x(self, v):
        span = max(1, self._to - self._from)
        w = max(self.winfo_width(), self._len)
        return self._pad + (min(max(v, self._from), self._to) - self._from) / span * (w - 2 * self._pad)

    def _press(self, e):
        w = max(self.winfo_width(), self._len)
        f = min(1, max(0, (e.x - self._pad) / max(1, (w - 2 * self._pad))))
        v = self._from + f * (self._to - self._from)
        self._var.set(v)
        if self._cmd:
            self._cmd(str(v))

    def _draw(self):
        S = self._S
        try:
            v = float(self._var.get())
        except (tk.TclError, ValueError):
            v = self._from
        w = max(self.winfo_width(), self._len)
        h = S(22)
        y = h // 2
        x = self._value_to_x(v)
        self.delete("all")
        round_rect(self, self._pad, y - S(2), w - self._pad, y + S(2), S(2), fill=LINE, outline="")
        if x > self._pad:
            round_rect(self, self._pad, y - S(2), x, y + S(2), S(2), fill=ACCENT, outline="")
        r = S(8)
        self.create_oval(x - r + 1, y - r + 2, x + r + 1, y + r + 2, fill=KNOB_SHADOW, outline="")   # soft shadow
        self.create_oval(x - r, y - r, x + r, y + r, fill="white", outline=LINE)


# ======================================================================================
#  3-D "glass" icon tiles for the ribbon buttons
# ======================================================================================
TILE_COLORS = ["#0A84FF", "#30B454", "#FF9F0A", "#BF5AF2", "#32ADE6", "#5E5CE6", "#FF6482", "#FF453A"]


def _hex_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _mix(c, other, t):
    return tuple(int(round(a + (b - a) * t)) for a, b in zip(c, other))


def render_tile(size, color, state="normal", pad=4, radius=9, scale=4):
    """PIL RGBA image of a glossy, raised (or pressed) rounded tile with a soft drop shadow.
    size / pad / radius are in screen pixels; drawn at `scale`x and shrunk for smooth edges."""
    from PIL import Image, ImageDraw, ImageFilter, ImageOps
    base = _hex_rgb(color)
    K = scale
    W = (size + 2 * pad) * K
    H = (size + 2 * pad) * K
    r = radius * K
    box = (pad * K, pad * K, (pad + size) * K, (pad + size) * K)
    out = Image.new("RGBA", (W, H), (0, 0, 0, 0))

    pressed = state in ("pressed", "active")
    hover = state == "hover"

    # drop shadow (smaller & tighter when pressed)
    sh = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    off = (1 if pressed else 3) * K
    blur = (2 if pressed else 3) * K
    ImageDraw.Draw(sh).rounded_rectangle((box[0] + K, box[1] + off, box[2] - K, box[3] + off - K), r, fill=(20, 30, 60, 90 if not pressed else 60))
    out = Image.alpha_composite(out, sh.filter(ImageFilter.GaussianBlur(blur)))

    # body: vertical gradient
    top = _mix(base, (255, 255, 255), 0.28 if not hover else 0.40)
    bot = _mix(base, (0, 0, 0), 0.22 if not pressed else 0.34)
    if pressed:
        top, bot = _mix(base, (0, 0, 0), 0.18), _mix(base, (255, 255, 255), 0.12)      # lit from below = pushed in
    gw, gh = box[2] - box[0], box[3] - box[1]
    grad = ImageOps.colorize(Image.linear_gradient("L").resize((gw, gh)), black=top, white=bot).convert("RGBA")
    mask = Image.new("L", (gw, gh), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, gw - 1, gh - 1), r, fill=255)
    body = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    body.paste(grad, (box[0], box[1]), mask)

    # glossy highlight on the upper half
    if not pressed:
        gloss = Image.new("RGBA", (gw, gh), (0, 0, 0, 0))
        gl = ImageDraw.Draw(gloss)
        gl.rounded_rectangle((K * 2, K * 2, gw - K * 2, gh // 2 + K), max(r - K * 2, 2), fill=(255, 255, 255, 70 if not hover else 95))
        fade = Image.linear_gradient("L").transpose(Image.FLIP_TOP_BOTTOM).resize((gw, gh))
        a = ImageChopsMultiply(gloss.split()[3], fade)
        gloss.putalpha(a)
        body.paste(gloss, (box[0], box[1]), gloss)
    # crisp edges: dark rim + bright inner rim at the top
    d = ImageDraw.Draw(body)
    rim = _mix(base, (0, 0, 0), 0.45)
    d.rounded_rectangle(box, r, outline=rim + (200,), width=max(1, K // 2 + 1))
    if not pressed:
        d.arc((box[0] + K, box[1] + K, box[0] + 2 * r, box[1] + 2 * r), 180, 270, fill=(255, 255, 255, 150), width=K)
        d.line((box[0] + r, box[1] + K, box[2] - r, box[1] + K), fill=(255, 255, 255, 130), width=K)
        d.arc((box[2] - 2 * r, box[1] + K, box[2] - K, box[1] + 2 * r), 270, 360, fill=(255, 255, 255, 110), width=K)
    else:                                     # inner shadow along the top edge
        d.line((box[0] + r, box[1] + K * 2, box[2] - r, box[1] + K * 2), fill=(0, 0, 0, 70), width=K * 2)
    out = Image.alpha_composite(out, body)
    return out.resize((size + 2 * pad, size + 2 * pad), Image.LANCZOS)


def ImageChopsMultiply(a, b):
    from PIL import ImageChops
    return ImageChops.multiply(a, b)


def tile_photo(widget, size, color, state, S):
    """cached ImageTk.PhotoImage for a tile (cache lives on the main window, so a new window gets new images)"""
    from PIL import ImageTk
    root = widget.winfo_toplevel()
    cache = getattr(root, "_tile_cache", None)
    if cache is None:
        cache = root._tile_cache = {}
    key = (size, color, state)
    ph = cache.get(key)
    if ph is None:
        ph = cache[key] = ImageTk.PhotoImage(render_tile(size, color, state, pad=S(4), radius=S(9)))
    return ph
