"""
Claudius PDF Editor - safety net and professional features: crash recovery, digital (certificate) signatures,
PDF/A conversion through Ghostscript, scanner import (Windows WIA) and read aloud.

Mixed into PdfEditor by claudius.py.  Heavy / optional libraries are imported only when used.
"""
import datetime
import glob
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import tkinter as tk
import webbrowser
from tkinter import ttk, filedialog, messagebox

try:
    import pymupdf as fitz
except ImportError:
    import fitz
from PIL import Image

import claudius_theme as TH
from claudius_i18n import tr as _tr
from claudius_extra import APP, UI_FONT, BG, S, parse_pages, pages_to_text

AUTOSAVE_SECONDS = 60
RECOVERY_KEEP_DAYS = 14
NO_WINDOW = 0x08000000 if sys.platform == "win32" else 0          # CREATE_NO_WINDOW


def home_dir():
    return os.path.join(os.path.expanduser("~"), ".claudius")


def recovery_dir():
    return os.path.join(home_dir(), "recovery")


# ====================================================================================== recovery helpers (pure)
def pid_alive(pid):
    if pid <= 0:
        return False
    if sys.platform == "win32":
        try:
            import ctypes
            k = ctypes.windll.kernel32
            h = k.OpenProcess(0x1000, False, int(pid))          # PROCESS_QUERY_LIMITED_INFORMATION
            if not h:
                return False
            code = ctypes.c_ulong()
            ok = k.GetExitCodeProcess(h, ctypes.byref(code))
            k.CloseHandle(h)
            return bool(ok) and code.value == 259               # STILL_ACTIVE
        except Exception:
            return True
    try:
        os.kill(int(pid), 0)
        return True
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False


def write_private(path, data):
    """atomic write, readable only by the current user where the system supports it"""
    tmp = path + ".tmp"
    fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "wb") as fh:
        fh.write(data)
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(tmp, path)


def save_recovery(rid, data, meta, folder=None):
    folder = folder or recovery_dir()
    os.makedirs(folder, exist_ok=True)
    try:
        os.chmod(folder, 0o700)
    except OSError:
        pass
    write_private(os.path.join(folder, rid + ".pdf"), data)
    meta = dict(meta, pid=os.getpid(), saved=time.time())
    write_private(os.path.join(folder, rid + ".json"), json.dumps(meta).encode("utf-8"))


def remove_recovery(rid, folder=None):
    folder = folder or recovery_dir()
    for ext in (".pdf", ".json", ".pdf.tmp", ".json.tmp"):
        try:
            os.remove(os.path.join(folder, rid + ext))
        except OSError:
            pass


def find_orphans(folder=None, own_pid=None):
    """-> list of (rid, meta) left behind by a program that is no longer running"""
    folder = folder or recovery_dir()
    own_pid = own_pid or os.getpid()
    out = []
    for jf in glob.glob(os.path.join(folder, "*.json")):
        rid = os.path.basename(jf)[:-5]
        try:
            with open(jf, "r", encoding="utf-8") as fh:
                meta = json.load(fh)
            pid = int(meta.get("pid", 0))
            if pid == own_pid or pid_alive(pid):
                continue
            if not os.path.exists(os.path.join(folder, rid + ".pdf")):
                remove_recovery(rid, folder)
                continue
            if time.time() - float(meta.get("saved", 0)) > RECOVERY_KEEP_DAYS * 86400:
                remove_recovery(rid, folder)
                continue
            out.append((rid, meta))
        except Exception:
            continue
    out.sort(key=lambda t: t[1].get("saved", 0))
    return out


# ====================================================================================== signatures (pure)
class SignError(Exception):
    pass


def have_pyhanko():
    try:
        import pyhanko  # noqa: F401
        import pyhanko.sign  # noqa: F401
        return True
    except Exception:
        return False


def _quiet_pyhanko():
    import logging
    for n in ("pyhanko", "pyhanko_certvalidator"):
        logging.getLogger(n).setLevel(logging.CRITICAL)


def make_self_signed_pfx(path, common_name, email, password, years=3):
    """Create a personal signing ID (.pfx).  Others will see it as 'self-signed' until they trust it."""
    from cryptography import x509
    from cryptography.x509.oid import NameOID
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.hazmat.primitives.serialization import pkcs12
    key = rsa.generate_private_key(public_exponent=65537, key_size=3072)
    attrs = [x509.NameAttribute(NameOID.COMMON_NAME, common_name)]
    if email:
        attrs.append(x509.NameAttribute(NameOID.EMAIL_ADDRESS, email))
    name = x509.Name(attrs)
    now = datetime.datetime.now(datetime.timezone.utc)
    cert = (x509.CertificateBuilder().subject_name(name).issuer_name(name).public_key(key.public_key())
            .serial_number(x509.random_serial_number())
            .not_valid_before(now - datetime.timedelta(minutes=5))
            .not_valid_after(now + datetime.timedelta(days=365 * years))
            .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
            .add_extension(x509.KeyUsage(digital_signature=True, content_commitment=True, key_encipherment=False,
                                         data_encipherment=False, key_agreement=False, key_cert_sign=False,
                                         crl_sign=False, encipher_only=False, decipher_only=False), critical=True)
            .sign(key, hashes.SHA256()))
    data = pkcs12.serialize_key_and_certificates(common_name.encode("utf-8"), key, cert, None,
                                                 serialization.BestAvailableEncryption(password.encode("utf-8")))
    write_private(path, data)
    return path


def sign_pdf_bytes(pdf_bytes, pfx_path, password, reason="", location="", contact="", page=0, box=None, field="Signature1"):
    """Sign with a PKCS#12 file.  box = (x0, y0, x1, y1) in PyMuPDF page coordinates (top-left origin) or None for an
    invisible signature.  -> signed PDF bytes."""
    _quiet_pyhanko()
    from pyhanko.sign import signers
    from pyhanko.sign.fields import SigFieldSpec, append_signature_field
    from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
    from pyhanko.stamp import TextStampStyle
    try:
        signer = signers.SimpleSigner.load_pkcs12(pfx_path, passphrase=(password or "").encode("utf-8"))
    except Exception as ex:
        raise SignError(_tr("Could not open the certificate file. Is the password right?") + f"\n{ex}")
    if signer is None:
        raise SignError(_tr("Could not open the certificate file. Is the password right?"))
    w = IncrementalPdfFileWriter(io.BytesIO(pdf_bytes), strict=False)
    spec = dict(sig_field_name=field)
    style = None
    if box is not None:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        H = doc[page].rect.height
        doc.close()
        x0, y0, x1, y1 = box
        spec.update(on_page=page, box=(x0, H - y1, x1, H - y0))
        who = signer.signing_cert.subject.human_friendly
        lines = [_tr("Digitally signed by") + " %(signer)s", "%(ts)s"]
        style = TextStampStyle(stamp_text="\n".join(lines))
    append_signature_field(w, SigFieldSpec(**spec))
    meta = signers.PdfSignatureMetadata(field_name=field, reason=reason or None, location=location or None,
                                        contact_info=contact or None)
    ps = signers.PdfSigner(meta, signer=signer, stamp_style=style)
    out = io.BytesIO()
    ps.sign_pdf(w, output=out)
    return out.getvalue()


def verify_signatures(pdf_bytes):
    """-> list of dict(name, signer, intact, valid, trusted, coverage, time, problem)"""
    _quiet_pyhanko()
    from pyhanko.pdf_utils.reader import PdfFileReader
    from pyhanko.sign.validation import validate_pdf_signature
    out = []
    r = PdfFileReader(io.BytesIO(pdf_bytes), strict=False)
    sigs = list(r.embedded_signatures)
    for k, sig in enumerate(sigs):
        row = dict(name=sig.field_name, signer="?", intact=False, valid=False, trusted=False, coverage="", time="", problem="",
                   last=(k == len(sigs) - 1), self_signed=False)
        try:
            cert = sig.signer_cert
            row["signer"] = cert.subject.human_friendly
            row["self_signed"] = cert.issuer == cert.subject
            st = validate_pdf_signature(sig)
            row.update(intact=bool(st.intact), valid=bool(st.valid), trusted=bool(st.trusted),
                       coverage=str(st.coverage).split(".")[-1])
            if st.signer_reported_dt:
                row["time"] = st.signer_reported_dt.strftime("%Y-%m-%d %H:%M")
        except Exception as ex:
            row["problem"] = str(ex)[:200]
        out.append(row)
    return out


def signature_fields_count(doc):
    n = 0
    try:
        for p in doc:
            for w in p.widgets() or []:
                if w.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE:
                    n += 1
    except Exception:
        pass
    return n


def doc_is_signed(pdf_bytes):
    return b"/ByteRange" in pdf_bytes and (b"/Type/Sig" in pdf_bytes or b"/Type /Sig" in pdf_bytes or b"/SubFilter" in pdf_bytes)


# ====================================================================================== PDF/A through Ghostscript
def find_ghostscript():
    for exe in ("gswin64c", "gswin32c", "gs"):
        p = shutil.which(exe)
        if p:
            return p
    if sys.platform == "win32":
        for base in (os.environ.get("ProgramFiles", r"C:\Program Files"), os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
            hits = sorted(glob.glob(os.path.join(base, "gs", "gs*", "bin", "gswin64c.exe")) +
                          glob.glob(os.path.join(base, "gs", "gs*", "bin", "gswin32c.exe")), reverse=True)
            if hits:
                return hits[0]
    return None


def _ps_string(s):
    return "(" + s.replace("\\", "/").replace("(", r"\(").replace(")", r"\)") + ")"


def find_icc(gs):
    root = os.path.dirname(os.path.dirname(gs))
    for pat in ("iccprofiles/default_rgb.icc", "iccprofiles/srgb.icc", "lib/default_rgb.icc", "../iccprofiles/default_rgb.icc",
                "../../iccprofiles/default_rgb.icc", "../share/ghostscript/*/iccprofiles/default_rgb.icc"):
        hits = glob.glob(os.path.join(root, pat))
        if hits:
            return hits[0]
    for pat in ("/usr/share/color/icc/ghostscript/default_rgb.icc", "/usr/share/ghostscript/*/iccprofiles/default_rgb.icc"):
        hits = glob.glob(pat)
        if hits:
            return hits[0]
    return None


def pdfa_definition(icc_path, title="Document"):
    return f"""%!
[ /Title {_ps_string(title)} /DOCINFO pdfmark
[/_objdef {{icc_PDFA}} /type /stream /OBJ pdfmark
[{{icc_PDFA}} <</N 3>> /PUT pdfmark
[{{icc_PDFA}} {_ps_string(icc_path)} (r) file /PUT pdfmark
[/_objdef {{OutputIntent_PDFA}} /type /dict /OBJ pdfmark
[{{OutputIntent_PDFA}} <<
  /Type /OutputIntent /S /GTS_PDFA1 /DestOutputProfile {{icc_PDFA}}
  /OutputConditionIdentifier (sRGB) >> /PUT pdfmark
[{{Catalog}} <</OutputIntents [ {{OutputIntent_PDFA}} ]>> /PUT pdfmark
"""


def convert_pdfa(src_pdf, dst_pdf, level=2, gs=None, title="Document"):
    """-> (ok, message).  Needs Ghostscript.  The result is checked for the PDF/A marker afterwards."""
    gs = gs or find_ghostscript()
    if not gs:
        return False, "no-gs"
    icc = find_icc(gs)
    if not icc:
        return False, "Ghostscript was found, but its colour profile (default_rgb.icc) was not."
    work = tempfile.mkdtemp(prefix="claudius_pdfa_")
    try:
        defs = os.path.join(work, "PDFA_def.ps")
        with open(defs, "w", encoding="latin-1", errors="replace") as fh:
            fh.write(pdfa_definition(icc, title))
        out = os.path.join(work, "out.pdf")
        cmd = [gs, f"-dPDFA={level}", "-dBATCH", "-dNOPAUSE", "-dQUIET", "-dNOOUTERSAVE", "-dSAFER",
               "-sColorConversionStrategy=RGB", "-sProcessColorModel=DeviceRGB", "-dPDFACompatibilityPolicy=1",
               "-sDEVICE=pdfwrite", f"-sOutputFile={out}", defs, src_pdf]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=600, creationflags=NO_WINDOW)
        if r.returncode != 0 or not os.path.exists(out):
            return False, (r.stderr or r.stdout or "Ghostscript failed.")[:400]
        d = fitz.open(out)
        xmp = d.get_xml_metadata() or ""
        d.close()
        if "pdfaid" not in xmp:
            return False, "Ghostscript finished, but the result does not declare PDF/A."
        shutil.copyfile(out, dst_pdf)
        return True, "OK"
    except subprocess.TimeoutExpired:
        return False, "Ghostscript took too long."
    finally:
        shutil.rmtree(work, ignore_errors=True)


# ====================================================================================== scanner (WIA) / images -> PDF
WIA_SCRIPT = r'''
$ErrorActionPreference = "Stop"
$dlg = New-Object -ComObject WIA.CommonDialog
$img = $dlg.ShowAcquireImage(1, 1, 0, "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}", $false, $true, $false)
if ($img -ne $null) { $img.SaveFile($env:CLAUDIUS_OUT); exit 0 } else { exit 3 }
'''


def scan_one(out_png):
    """Show the Windows scanner dialog and save the page as PNG.  -> True when a page was scanned."""
    if sys.platform != "win32":
        raise RuntimeError(_tr("Scanning is available on Windows only."))
    if os.path.exists(out_png):
        os.remove(out_png)
    env = dict(os.environ, CLAUDIUS_OUT=out_png)
    r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", WIA_SCRIPT],
                       capture_output=True, text=True, env=env, timeout=900, creationflags=NO_WINDOW)
    if r.returncode == 3:
        return False
    if r.returncode != 0 or not os.path.exists(out_png):
        raise RuntimeError((r.stderr or r.stdout or "The scanner did not respond.").strip()[:300])
    return True


def images_to_pdf(files, default_dpi=200):
    """Make a document with one page per image, sized from the image resolution."""
    doc = fitz.open()
    for f in files:
        with Image.open(f) as im:
            im = im.convert("RGB")
            dpi = im.info.get("dpi", (default_dpi, default_dpi))
            try:
                dx = float(dpi[0]) or default_dpi
            except Exception:
                dx = default_dpi
            if not 50 <= dx <= 1200:
                dx = default_dpi
            w, h = im.size[0] / dx * 72, im.size[1] / dx * 72
            buf = io.BytesIO()
            im.save(buf, "JPEG", quality=88)
        pg = doc.new_page(width=w, height=h)
        pg.insert_image(pg.rect, stream=buf.getvalue())
    return doc


# ====================================================================================== read aloud
def speak_command(text_file, rate=0, voice=None):
    """-> (argv, env) for the platform's speech engine, or None.  The text goes through a file / environment
    variable, never through the command line, so nothing in the document can be run as a command."""
    env = dict(os.environ, CLAUDIUS_TXT=text_file, CLAUDIUS_RATE=str(int(rate)), CLAUDIUS_VOICE=voice or "")
    if sys.platform == "win32":
        ps = ("Add-Type -AssemblyName System.Speech; $s = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
              "$s.Rate = [int]$env:CLAUDIUS_RATE; if ($env:CLAUDIUS_VOICE) { try { $s.SelectVoice($env:CLAUDIUS_VOICE) } catch {} }; "
              "$t = Get-Content -Raw -Encoding UTF8 -LiteralPath $env:CLAUDIUS_TXT; $s.Speak($t)")
        return ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", ps], env
    if sys.platform == "darwin" and shutil.which("say"):
        return ["say", "-r", str(175 + 15 * int(rate)), "-f", text_file], env
    for exe in ("espeak-ng", "espeak"):
        p = shutil.which(exe)
        if p:
            return [p, "-s", str(160 + 12 * int(rate)), "-f", text_file], env
    return None


def list_voices():
    if sys.platform != "win32":
        return []
    try:
        ps = ("Add-Type -AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer)"
              ".GetInstalledVoices() | ForEach-Object { $_.VoiceInfo.Name }")
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps], capture_output=True, text=True,
                           timeout=20, creationflags=NO_WINDOW)
        return [v.strip() for v in r.stdout.splitlines() if v.strip()]
    except Exception:
        return []


# ====================================================================================== the mixin
class ProTools:
    # ---------------------------------------------------------------- recovery
    def _pro_init(self):
        self._rec_after = None
        self._rec_sig = {}
        self._speech = None
        self._speech_file = None
        self._pdf_signed_bytes = None
        self._sig_orig = None

    def start_recovery(self):
        """called once the window is up: offer lost work, then start the autosave timer"""
        try:
            orphans = find_orphans()
        except Exception:
            orphans = []
        if orphans:
            names = "\n".join("  " + (m.get("title") or "Untitled") for _r, m in orphans[:8])
            if messagebox.askyesno(APP, _tr("Claudius PDF Editor did not close normally last time.\nRecover these documents?") + "\n\n" + names):
                for rid, meta in orphans:
                    try:
                        data = open(os.path.join(recovery_dir(), rid + ".pdf"), "rb").read()
                        doc = fitz.open(stream=data, filetype="pdf")
                        self._next_title = (meta.get("title") or "Untitled") + " (" + _tr("recovered") + ")"
                        self._adopt_doc(doc, None, modified=True)
                    except Exception as ex:
                        messagebox.showerror(APP, f"Could not recover a document:\n{ex}")
            for rid, _m in orphans:
                remove_recovery(rid)
        self._rec_after = self.after(self._autosave_ms(), self._autosave_tick)

    def _autosave_tick(self):
        try:
            if self.cfg.get("autosave", True):
                self.autosave_now()
        except Exception:
            pass
        self._rec_after = self.after(self._autosave_ms(), self._autosave_tick)

    def autosave_now(self):
        """Write every changed, unsaved document to the recovery folder.  Password-protected documents are never
        written (the password is not kept anywhere)."""
        if not self.tabs:
            return 0
        self._tab_stash()
        n = 0
        for i, st in enumerate(self.tabs):
            doc = st.get("doc")
            rid = st.setdefault("rid", f"{os.getpid()}_{int(time.time())}_{i}_{id(st) & 0xFFFF:x}")
            if doc is None:
                continue
            if not st.get("modified"):
                if self._rec_sig.pop(rid, None) is not None:
                    remove_recovery(rid)
                continue
            if st.get("protect"):
                continue
            sig = (id(doc), len(st.get("undo_stack") or []), len(st.get("redo_stack") or []), len(doc), bool(st.get("floats")))
            if self._rec_sig.get(rid) == sig:
                continue
            try:
                data = doc.tobytes(garbage=0, deflate=True, encryption=fitz.PDF_ENCRYPT_NONE)
                title = os.path.basename(st["path"]) if st.get("path") else (st.get("title") or "Untitled")
                save_recovery(rid, data, dict(title=title, path=st.get("path") or ""))
                self._rec_sig[rid] = sig
                n += 1
            except Exception:
                pass
        return n

    def clear_recovery(self):
        if self._rec_after:
            try:
                self.after_cancel(self._rec_after)
            except Exception:
                pass
        for st in self.tabs:
            if st.get("rid"):
                remove_recovery(st["rid"])
        for rid in list(self._rec_sig):
            remove_recovery(rid)

    # ---------------------------------------------------------------- signatures
    def _need_pyhanko(self):
        if have_pyhanko():
            return True
        if getattr(sys, "frozen", False):
            messagebox.showinfo(APP, _tr("This program build does not contain the signature component.\n\n"
                                         "Install the full version, or build it again with build_exe.bat "
                                         "after running:  pip install pyhanko"))
            return False
        if not messagebox.askyesno(APP, _tr("Digital signatures need an extra component (pyHanko).\n\n"
                                            "Download and install it now? (needs internet, about 30 seconds)")):
            return False
        return self._install_pyhanko()

    def _install_pyhanko(self):
        top = tk.Toplevel(self)
        top.title(APP)
        top.transient(self)
        top.resizable(False, False)
        ttk.Label(top, text=_tr("Installing the signature component..."), padding=24).pack()
        bar = ttk.Progressbar(top, mode="indeterminate", length=S(280))
        bar.pack(padx=24, pady=(0, 20))
        bar.start(12)
        top.protocol("WM_DELETE_WINDOW", lambda: None)
        top.update_idletasks()
        top.geometry("+%d+%d" % (self.winfo_rootx() + (self.winfo_width() - top.winfo_reqwidth()) // 2,
                                 self.winfo_rooty() + (self.winfo_height() - top.winfo_reqheight()) // 3))
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        exe = sys.executable
        if os.name == "nt" and exe.lower().endswith("pythonw.exe"):
            exe = exe[:-11] + "python.exe"
        out = {}
        try:
            proc = subprocess.Popen([exe, "-m", "pip", "install", "--user", "--disable-pip-version-check", "pyhanko"],
                                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                                    creationflags=flags)
        except Exception as ex:
            top.destroy()
            messagebox.showerror(APP, str(ex))
            return False

        def reader():
            try:
                out["log"] = proc.communicate()[0] or ""
            except Exception as ex:
                out["log"] = str(ex)
        import threading
        threading.Thread(target=reader, daemon=True).start()
        while proc.poll() is None:
            try:
                top.update()
            except tk.TclError:
                break
            time.sleep(0.03)
        time.sleep(0.1)
        try:
            top.destroy()
        except tk.TclError:
            pass
        import importlib
        importlib.invalidate_caches()
        try:
            import site
            us = site.getusersitepackages()
            if us and us not in sys.path and os.path.isdir(us):
                sys.path.append(us)
        except Exception:
            pass
        if have_pyhanko():
            messagebox.showinfo(APP, _tr("The signature component is installed."))
            return True
        messagebox.showerror(APP, _tr("Installation failed. Please run:  pip install pyhanko") +
                             "\n\n" + (out.get("log", "")[-600:]))
        return False

    def dlg_create_id(self):
        if not self._need_pyhanko():
            return
        d = self._dlg("Create my signing ID")
        name, mail, pw, pw2 = tk.StringVar(), tk.StringVar(), tk.StringVar(), tk.StringVar()
        self._label(d, "Your name").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(d, textvariable=name, width=30).grid(row=0, column=1, pady=3)
        self._label(d, "E-mail (optional)").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(d, textvariable=mail, width=30).grid(row=1, column=1, pady=3)
        self._label(d, "Password for the ID").grid(row=2, column=0, sticky="w", pady=3)
        ttk.Entry(d, textvariable=pw, show="*", width=30).grid(row=2, column=1, pady=3)
        self._label(d, "Repeat password").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Entry(d, textvariable=pw2, show="*", width=30).grid(row=3, column=1, pady=3)
        self._label(d, "This creates a personal ID file (.pfx). Others will see it as 'self-signed' until they choose to trust it.\n"
                       "Keep the file and its password safe.", fg=TH.TERTIARY).grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

        def go():
            if not name.get().strip():
                messagebox.showwarning(APP, "Please enter your name.", parent=d)
                return
            if len(pw.get()) < 6 or pw.get() != pw2.get():
                messagebox.showwarning(APP, "The passwords must match and have at least 6 characters.", parent=d)
                return
            p = filedialog.asksaveasfilename(parent=d, defaultextension=".pfx", initialfile="my-signing-id.pfx",
                                             filetypes=[("Signing ID", "*.pfx *.p12")])
            if not p:
                return
            try:
                make_self_signed_pfx(p, name.get().strip(), mail.get().strip(), pw.get())
            except Exception as ex:
                messagebox.showerror(APP, f"That did not work:\n{ex}", parent=d)
                return
            d.destroy()
            messagebox.showinfo(APP, _tr("Your signing ID was saved.\nUse Sign > Sign with certificate to sign documents."))
        self._buttons(d, "Create", go, 5, colspan=2)

    def dlg_sign_cert(self):
        if not self.need_doc() or not self._need_pyhanko():
            return
        if self.protect:
            messagebox.showinfo(APP, "Remove the password protection first (Document > Unprotect), then sign.")
            return
        self.commit_floats()
        d = self._dlg("Sign with certificate")
        pfx, pw = tk.StringVar(value=getattr(self, "_last_pfx", "")), tk.StringVar()
        reason, loc, contact = tk.StringVar(), tk.StringVar(), tk.StringVar()
        vis = tk.BooleanVar(value=True)
        pos = tk.StringVar(value="Bottom right")
        pgv = tk.IntVar(value=self.page_no + 1)
        self._label(d, "Signing ID file").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(d, textvariable=pfx, width=34).grid(row=0, column=1, pady=3)
        ttk.Button(d, text="Browse...", command=lambda: pfx.set(filedialog.askopenfilename(
            parent=d, filetypes=[("Signing ID", "*.pfx *.p12"), ("All files", "*.*")]) or pfx.get())).grid(row=0, column=2, padx=4)
        self._label(d, "Password").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(d, textvariable=pw, show="*", width=34).grid(row=1, column=1, pady=3)
        for r, (lab, v) in enumerate((("Reason", reason), ("Location", loc), ("Contact", contact)), start=2):
            self._label(d, lab).grid(row=r, column=0, sticky="w", pady=3)
            ttk.Entry(d, textvariable=v, width=34).grid(row=r, column=1, pady=3)
        tk.Checkbutton(d, text="Show the signature on the page", variable=vis, bg=BG, activebackground=BG, fg=TH.LABEL,
                       selectcolor=TH.FILL, font=(UI_FONT, 9)).grid(row=5, column=0, columnspan=2, sticky="w", pady=(6, 0))
        self._label(d, "Page").grid(row=6, column=0, sticky="w")
        ttk.Spinbox(d, from_=1, to=len(self.doc), width=6, textvariable=pgv).grid(row=6, column=1, sticky="w")
        self._label(d, "Position").grid(row=7, column=0, sticky="w", pady=3)
        ttk.Combobox(d, textvariable=pos, state="readonly", width=16,
                     values=["Bottom right", "Bottom left", "Bottom center", "Top right", "Top left", "Top center"]).grid(row=7, column=1, sticky="w")
        self._label(d, "The signed copy is saved as a new file. Changing it afterwards breaks the signature.",
                    fg=TH.TERTIARY).grid(row=8, column=0, columnspan=3, sticky="w", pady=(8, 0))

        def go():
            if not pfx.get() or not os.path.exists(pfx.get()):
                messagebox.showwarning(APP, "Choose your signing ID file first.", parent=d)
                return
            page = max(0, min(len(self.doc) - 1, int(pgv.get()) - 1))
            box = None
            if vis.get():
                if self.doc[page].rotation:
                    messagebox.showinfo(APP, "This page is rotated, so the signature will be invisible.", parent=d)
                else:
                    r = self.doc[page].rect
                    w, h, m = 190, 54, 28
                    v, _, hz = pos.get().lower().partition(" ")
                    x0 = {"left": r.x0 + m, "center": r.x0 + (r.width - w) / 2, "right": r.x1 - m - w}[hz]
                    y0 = r.y1 - m - h if v == "bottom" else r.y0 + m
                    box = (x0, y0, x0 + w, y0 + h)
            out = filedialog.asksaveasfilename(parent=d, defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                               initialfile=os.path.splitext(os.path.basename(self.path or "document"))[0] + "_signed.pdf")
            if not out:
                return
            self._last_pfx = pfx.get()
            try:
                data = self.doc.tobytes(garbage=0, deflate=True, encryption=fitz.PDF_ENCRYPT_NONE)
                signed = sign_pdf_bytes(data, pfx.get(), pw.get(), reason.get(), loc.get(), contact.get(), page, box)
                with open(out, "wb") as fh:
                    fh.write(signed)
            except SignError as ex:
                messagebox.showerror(APP, str(ex), parent=d)
                return
            except Exception as ex:
                messagebox.showerror(APP, f"That did not work:\n{ex}", parent=d)
                return
            d.destroy()
            self.status.config(text=f"Signed copy saved: {out}")
            self.open_path(out)
        self._buttons(d, "Sign", go, 9, colspan=3)

    def dlg_verify_signatures(self):
        if not self.need_doc() or not self._need_pyhanko():
            return
        data = self._sig_orig or (self.doc.tobytes(garbage=0, encryption=fitz.PDF_ENCRYPT_NONE) if not self.protect else None)
        if data is None:
            messagebox.showinfo(APP, "Remove the password protection first to check signatures.")
            return
        try:
            rows = verify_signatures(data)
        except Exception as ex:
            messagebox.showerror(APP, f"That did not work:\n{ex}")
            return
        d = self._dlg("Digital signatures")
        if not rows:
            self._label(d, "This document has no digital signatures.").grid(row=0, column=0, pady=8)
        for r, row in enumerate(rows):
            ok = row["intact"] and row["valid"]
            col = "#1f9d55" if ok else "#d93025"
            fr = tk.Frame(d, bg=BG)
            fr.grid(row=r, column=0, sticky="w", pady=6)
            tk.Label(fr, text="✔" if ok else "✖", fg=col, bg=BG, font=(UI_FONT, 16, "bold")).pack(side="left", padx=(0, 10))
            txt = tk.Frame(fr, bg=BG)
            txt.pack(side="left")
            self._label(txt, row["signer"], bold=True).pack(anchor="w")
            if row["problem"]:
                self._label(txt, row["problem"], fg="#d93025").pack(anchor="w")
            else:
                self._label(txt, _tr("The signature is valid and the document has not been changed since.") if ok and row["last"] and row["coverage"] == "ENTIRE_FILE"
                            else (_tr("The signature is valid, but the document was changed after signing.") if ok
                                  else _tr("The signature is NOT valid: the document was changed, or the signature is damaged."))).pack(anchor="w")
                trust = _tr("Identity: self-signed certificate - not verified.") if row["self_signed"] else (
                    _tr("Identity: the certificate chain is trusted.") if row["trusted"] else _tr("Identity: the certificate authority is not in your trusted list."))
                self._label(txt, trust + (f"   {row['time']}" if row["time"] else ""), fg=TH.TERTIARY).pack(anchor="w")
        self._label(d, "Claudius PDF Editor checks that the document is unchanged. It does not use the Windows certificate store, "
                       "so confirm the signer's identity with them.", fg=TH.TERTIARY).grid(row=len(rows) + 1, column=0, sticky="w", pady=(8, 0))
        f = tk.Frame(d, bg=BG)
        f.grid(row=len(rows) + 2, column=0, sticky="e", pady=(10, 0))
        ttk.Button(f, text="Close", command=d.destroy).pack()

    def _after_open_check_signed(self, data):
        """remember the exact bytes of a signed file so it can be saved again without breaking the signature"""
        try:
            if doc_is_signed(data) and signature_fields_count(self.doc) > 0:
                self._sig_orig = data
                self.status.config(text=_tr("This document is digitally signed. Editing and saving will invalidate the signature."))
                return True
        except Exception:
            pass
        self._sig_orig = None
        return False

    # ---------------------------------------------------------------- PDF/A
    def dlg_pdfa(self):
        if not self.need_doc():
            return
        gs = find_ghostscript()
        if not gs:
            if messagebox.askyesno(APP, _tr("PDF/A conversion uses the free program Ghostscript, which is not installed.\n\n"
                                            "Open the download page now?")):
                webbrowser.open("https://ghostscript.com/releases/gsdnld.html")
            return
        d = self._dlg("Save as PDF/A")
        lv = tk.StringVar(value="PDF/A-2b (recommended)")
        self._label(d, "Standard").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Combobox(d, textvariable=lv, state="readonly", width=24, values=["PDF/A-1b", "PDF/A-2b (recommended)", "PDF/A-3b"]).grid(row=0, column=1)
        self._label(d, "Fonts are embedded and colours are converted to sRGB. Forms and some effects may be flattened.",
                    fg=TH.TERTIARY).grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

        def go():
            level = int(lv.get()[5])
            out = filedialog.asksaveasfilename(parent=d, defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                               initialfile=os.path.splitext(os.path.basename(self.path or "document"))[0] + "_pdfa.pdf")
            if not out:
                return
            d.destroy()
            self.commit_floats()
            tmp = os.path.join(tempfile.gettempdir(), f"claudius_cv_pdfa_{os.getpid()}.pdf")
            try:
                self.doc.save(tmp, garbage=3, deflate=True, encryption=fitz.PDF_ENCRYPT_NONE)
                self.status.config(text=_tr("Converting to PDF/A..."))
                self.update_idletasks()
                ok, msg = convert_pdfa(tmp, out, level, gs, os.path.basename(self.path or "Document"))
            finally:
                try:
                    os.remove(tmp)
                except OSError:
                    pass
            if ok:
                self.status.config(text=f"Saved: {out}")
                messagebox.showinfo(APP, _tr("Saved as PDF/A.\n\nFor official archiving, check the file with a validator such as veraPDF."))
            else:
                messagebox.showerror(APP, f"PDF/A conversion failed:\n{msg}")
        self._buttons(d, "Convert", go, 2, colspan=2)

    # ---------------------------------------------------------------- scanner
    def dlg_scan(self):
        if sys.platform != "win32":
            messagebox.showinfo(APP, _tr("Scanning is available on Windows only."))
            return
        work = tempfile.mkdtemp(prefix="claudius_scan_")
        files = []
        try:
            while True:
                out = os.path.join(work, f"scan{len(files) + 1}.png")
                self.status.config(text=_tr("Waiting for the scanner..."))
                self.update_idletasks()
                try:
                    got = scan_one(out)
                except Exception as ex:
                    messagebox.showerror(APP, f"Scanner problem:\n{ex}")
                    break
                if not got:
                    break
                files.append(out)
                if not messagebox.askyesno(APP, _tr("Page {} scanned.\n\nScan another page?").replace("{}", str(len(files)))):
                    break
            if files:
                doc = images_to_pdf(files)
                self._open_result(doc, "Scan")
                self.status.config(text=_tr("Scanned {} page(s). Use Document > OCR to make the text searchable.").replace("{}", str(len(files))))
        finally:
            shutil.rmtree(work, ignore_errors=True)

    # ---------------------------------------------------------------- read aloud
    def dlg_read_aloud(self):
        if not self.need_doc():
            return
        if speak_command(os.devnull) is None:
            messagebox.showinfo(APP, _tr("No speech engine was found on this computer."))
            return
        d = self._dlg("Read aloud")
        scope = tk.StringVar(value="selection" if self._tsel else "page")
        rate = tk.IntVar(value=0)
        voice = tk.StringVar()
        for r, (v, t) in enumerate((("selection", "The selected text"), ("page", "This page"), ("rest", "From this page to the end"))):
            tk.Radiobutton(d, text=t, variable=scope, value=v, bg=BG, activebackground=BG, fg=TH.LABEL, selectcolor=TH.FILL,
                           font=(UI_FONT, 9)).grid(row=r, column=0, columnspan=2, sticky="w")
        self._label(d, "Speed").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Scale(d, from_=-5, to=5, variable=rate, orient="horizontal", length=160).grid(row=3, column=1, sticky="w", pady=(8, 0))
        voices = list_voices()
        if voices:
            self._label(d, "Voice").grid(row=4, column=0, sticky="w", pady=4)
            ttk.Combobox(d, textvariable=voice, values=voices, state="readonly", width=28).grid(row=4, column=1, sticky="w")
        self._label(d, "The document's language must match the voice. Scanned pages need OCR first.", fg=TH.TERTIARY).grid(
            row=5, column=0, columnspan=2, sticky="w", pady=(6, 0))

        def go():
            text = self._text_for_speech(scope.get())
            if not text.strip():
                messagebox.showinfo(APP, _tr("There is no text to read here."), parent=d)
                return
            d.destroy()
            self.speak(text, int(round(rate.get())), voice.get() or None)
        f = self._buttons(d, "Read", go, 6, colspan=2)
        ttk.Button(f, text="Stop", command=lambda: (self.stop_speaking(), None)).pack(side="left", padx=4)

    def _text_for_speech(self, scope):
        if scope == "selection" and self._tsel:
            return self._sel_text().replace("\n", " ")
        pages = [self.page_no] if scope in ("page", "selection") else range(self.page_no, len(self.doc))
        out = []
        for i in pages:
            out.append(self.doc[i].get_text("text", sort=True))
        return "\n".join(out)[:400000]

    def speak(self, text, rate=0, voice=None):
        self.stop_speaking()
        fd, path = tempfile.mkstemp(prefix="claudius_tts_", suffix=".txt")
        with os.fdopen(fd, "w", encoding="utf-8-sig" if sys.platform == "win32" else "utf-8") as fh:
            fh.write(text)
        cmd = speak_command(path, rate, voice)
        if cmd is None:
            os.remove(path)
            return
        argv, env = cmd
        try:
            self._speech = subprocess.Popen(argv, env=env, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                                            stderr=subprocess.DEVNULL, creationflags=NO_WINDOW)
        except Exception as ex:
            os.remove(path)
            messagebox.showerror(APP, f"Could not start the speech engine:\n{ex}")
            return
        self._speech_file = path
        self.status.config(text=_tr("Reading aloud... (View > Read aloud > Stop to end)"))
        self.after(800, self._speech_poll)

    def _speech_poll(self):
        p = self._speech
        if p is None:
            return
        if p.poll() is None:
            self.after(800, self._speech_poll)
        else:
            self.stop_speaking()
            self.status.config(text=_tr("Finished reading."))

    def stop_speaking(self):
        p, self._speech = self._speech, None
        if p is not None and p.poll() is None:
            try:
                p.terminate()
            except Exception:
                pass
        if self._speech_file:
            try:
                os.remove(self._speech_file)
            except OSError:
                pass
            self._speech_file = None
