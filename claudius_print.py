"""
Claudius PDF Editor - printing.

Windows : pages are drawn with the Windows printing system (GDI) straight to the chosen printer - no other PDF
          program is needed.  If that is not possible the PDF is handed to Windows' own 'print' command.
macOS / Linux : the chosen pages are saved to a temporary PDF and sent with  lp.
"""
import os
import shutil
import subprocess
import sys
import tempfile

try:
    import pymupdf as fitz
except ImportError:
    import fitz
from PIL import Image


# ------------------------------------------------------------------------------------ printer list
def list_printers():
    """-> (names, default_name_or_None)"""
    if sys.platform == "win32":
        try:
            return _win_printers()
        except Exception:
            return [], None
    if shutil.which("lpstat"):
        try:
            names = [ln.split()[0] for ln in subprocess.run(["lpstat", "-e"], capture_output=True, text=True,
                                                            timeout=8).stdout.splitlines() if ln.strip()]
            out = subprocess.run(["lpstat", "-d"], capture_output=True, text=True, timeout=8).stdout
            default = out.split(":")[-1].strip() if ":" in out else None
            return names, (default or None)
        except Exception:
            pass
    return [], None


def _win_printers():
    import ctypes
    from ctypes import wintypes

    class PRINTER_INFO_4W(ctypes.Structure):
        _fields_ = [("pPrinterName", wintypes.LPWSTR), ("pServerName", wintypes.LPWSTR), ("Attributes", wintypes.DWORD)]

    sp = ctypes.WinDLL("winspool.drv", use_last_error=True)
    flags = 2 | 4                                            # PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS
    needed, returned = wintypes.DWORD(), wintypes.DWORD()
    sp.EnumPrintersW(flags, None, 4, None, 0, ctypes.byref(needed), ctypes.byref(returned))
    names = []
    if needed.value:
        buf = ctypes.create_string_buffer(needed.value)
        if sp.EnumPrintersW(flags, None, 4, buf, needed, ctypes.byref(needed), ctypes.byref(returned)):
            arr = ctypes.cast(buf, ctypes.POINTER(PRINTER_INFO_4W))
            names = [arr[i].pPrinterName for i in range(returned.value)]
    size = wintypes.DWORD(0)
    sp.GetDefaultPrinterW(None, ctypes.byref(size))
    default = None
    if size.value:
        b = ctypes.create_unicode_buffer(size.value)
        if sp.GetDefaultPrinterW(b, ctypes.byref(size)):
            default = b.value
    return names, default


# ------------------------------------------------------------------------------------ Windows GDI backend
class WinGdiPrinter:
    HORZRES, VERTRES, LOGPIXELSX, LOGPIXELSY = 8, 10, 88, 90

    def __init__(self, printer_name, title):
        import ctypes
        from ctypes import wintypes
        self.ct = ctypes
        self.gdi = ctypes.WinDLL("gdi32", use_last_error=True)
        g = self.gdi
        g.CreateDCW.restype = wintypes.HDC
        g.CreateDCW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.c_void_p]
        for fn in ("StartPage", "EndPage", "EndDoc", "AbortDoc", "DeleteDC"):
            getattr(g, fn).argtypes = [wintypes.HDC]
        g.StartDocW.argtypes = [wintypes.HDC, ctypes.c_void_p]
        g.GetDeviceCaps.argtypes = [wintypes.HDC, ctypes.c_int]
        g.SetStretchBltMode.argtypes = [wintypes.HDC, ctypes.c_int]

        class DOCINFOW(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_int), ("lpszDocName", wintypes.LPCWSTR), ("lpszOutput", wintypes.LPCWSTR),
                        ("lpszDatatype", wintypes.LPCWSTR), ("fwType", wintypes.DWORD)]
        self.hdc = g.CreateDCW("WINSPOOL", printer_name, None, None)
        if not self.hdc:
            raise OSError(f"Could not open printer '{printer_name}'")
        g.SetStretchBltMode(self.hdc, 4)                      # HALFTONE
        di = DOCINFOW(ctypes.sizeof(DOCINFOW), title[:120], None, None, 0)
        if g.StartDocW(self.hdc, ctypes.byref(di)) <= 0:
            g.DeleteDC(self.hdc)
            raise OSError("The printer did not accept the document")
        self.area = (g.GetDeviceCaps(self.hdc, self.HORZRES), g.GetDeviceCaps(self.hdc, self.VERTRES))
        self.dpi = (g.GetDeviceCaps(self.hdc, self.LOGPIXELSX), g.GetDeviceCaps(self.hdc, self.LOGPIXELSY))
        self._in_page = False

    def page_area(self):
        return self.area

    def draw_page(self, image, box):
        from PIL import ImageWin
        self.gdi.StartPage(self.hdc)
        self._in_page = True
        ImageWin.Dib(image).draw(int(self.hdc), box)
        self.gdi.EndPage(self.hdc)
        self._in_page = False

    def close(self, ok=True):
        if not self.hdc:
            return
        if ok:
            self.gdi.EndDoc(self.hdc)
        else:
            self.gdi.AbortDoc(self.hdc)
        self.gdi.DeleteDC(self.hdc)
        self.hdc = None


# ------------------------------------------------------------------------------------ layout + loop
def layout(page_w, page_h, area_w, area_h, fit, dpi_x, dpi_y):
    """-> (rotate90, zoom_px_per_pt, (x0, y0, x1, y1)) for one page on a printable area given in device pixels"""
    rotate = (page_w > page_h) != (area_w > area_h)
    pw, ph = (page_h, page_w) if rotate else (page_w, page_h)
    fit_scale = min(area_w / pw, area_h / ph)                 # device pixels per point
    if fit == "actual":
        actual = dpi_x / 72.0
        scale = min(actual, fit_scale)                        # never spill over the edge
    else:
        scale = fit_scale
    w, h = int(pw * scale), int(ph * scale)
    x0, y0 = (area_w - w) // 2, (area_h - h) // 2
    return rotate, scale, (x0, y0, x0 + w, y0 + h)


def render_for_printer(page, area, dpis, fit="fit", max_dpi=300):
    r = page.rect
    rotate, scale, box = layout(r.width, r.height, area[0], area[1], fit, dpis[0], dpis[1])
    zoom = min(scale, max_dpi / 72.0)                         # no need for more than 300 dpi of detail
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False, annots=True)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    if rotate:
        img = img.transpose(Image.ROTATE_90)
    return img, box


def print_pages_gdi(doc, pages, backend, copies=1, fit="fit", progress=None):
    """draw the pages with a backend that has page_area() / dpi / draw_page(image, box)"""
    total = len(pages) * copies
    done = 0
    for _c in range(copies):
        for pno in pages:
            img, box = render_for_printer(doc[pno], backend.page_area(), backend.dpi, fit)
            backend.draw_page(img, box)
            done += 1
            if progress:
                progress(done, total)
    return done


def print_document(doc, pages, printer, title, copies=1, fit="fit", progress=None):
    """send the pages to the printer.  Returns a short description of the method used."""
    copies = max(1, int(copies))
    if sys.platform == "win32":
        backend = None
        try:
            backend = WinGdiPrinter(printer, title)
            ok = False
            try:
                print_pages_gdi(doc, pages, backend, copies, fit, progress)
                ok = True
            finally:
                backend.close(ok)
            return "gdi"
        except Exception as ex:
            if backend is not None and backend.hdc is None and not isinstance(ex, OSError):
                raise
            # fall back: let Windows hand the file to whichever program handles 'print' for PDFs
            tmp = _pages_to_temp_pdf(doc, pages)
            try:
                os.startfile(tmp, "print")
            except Exception as ex2:
                raise RuntimeError(f"{ex}\n{ex2}")
            return "startfile"
    tmp = _pages_to_temp_pdf(doc, pages)
    lp = shutil.which("lp")
    if not lp:
        raise RuntimeError("No printing system found (the 'lp' command is missing).")
    cmd = [lp]
    if printer:
        cmd += ["-d", printer]
    cmd += ["-n", str(copies), "-o", "fit-to-page" if fit == "fit" else "scaling=100", tmp]
    try:
        out = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    finally:
        try:
            os.remove(tmp)               # the spooler has its own copy; do not leave the document in the temp folder
        except OSError:
            pass
    if out.returncode != 0:
        raise RuntimeError(out.stderr.strip() or "lp failed")
    return "lp"


def _pages_to_temp_pdf(doc, pages):
    new = fitz.open()
    try:
        for p in pages:
            new.insert_pdf(doc, from_page=p, to_page=p)
        fd, tmp = tempfile.mkstemp(prefix="claudius_print_", suffix=".pdf")
        os.close(fd)
        new.save(tmp, garbage=1)
        return tmp
    finally:
        new.close()
