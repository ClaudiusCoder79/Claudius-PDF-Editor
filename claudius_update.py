# -*- coding: utf-8 -*-
"""Claudius PDF Editor - update check through the GitHub releases page.

* Off by default. Settings > Updates: "Look for a new version when the program starts" (at most once a day), or
  File > Check for updates... at any time.
* One request to  api.github.com/repos/ClaudiusCoder79/Claudius-PDF-Editor/releases/latest  - nothing about you or your files is sent.
* A new version is only downloaded when you press "Download and install". The installer is taken from the same GitHub
  release, checked against SHA256SUMS.txt of that release (a wrong checksum deletes the file), and started only after
  you confirm. Claudius closes first (it asks to save open work).
* Only https, only github.com / githubusercontent.com hosts, size limits, file names checked. Nothing is ever started
  without your click.
"""
import datetime
import hashlib
import json
import os
import re
import shutil
import sys
import tempfile
import threading
import tkinter as tk
import urllib.error
import urllib.request
import webbrowser
from tkinter import ttk, messagebox
from urllib.parse import urlparse

import claudius_theme as TH
from claudius_i18n import tr as _tr
from claudius_extra import APP, UI_FONT, BG

REPO = "ClaudiusCoder79/Claudius-PDF-Editor"
API_URL = "https://api.github.com/repos/%s/releases/latest" % REPO
RELEASE_PAGE = "https://github.com/%s/releases" % REPO
ALLOWED_HOSTS = ("api.github.com", "github.com", "objects.githubusercontent.com", "release-assets.githubusercontent.com")
MAX_JSON = 2_000_000
MAX_SUMS = 64_000
MAX_INSTALLER = 400 * 1024 * 1024
TIMEOUT = 15
INSTALLER_RE = re.compile(r"^Claudius-PDF-Editor-Setup-\d+(\.\d+){1,3}\.exe$")
TAG_RE = re.compile(r"^v?(\d+(?:\.\d+){1,3})$")


# ------------------------------------------------------------------------------------------------ network (pure functions)
def _host_ok(url):
    try:
        u = urlparse(url)
        return u.scheme == "https" and (u.hostname or "").lower() in ALLOWED_HOSTS and not u.username
    except Exception:
        return False


class _SafeRedirect(urllib.request.HTTPRedirectHandler):
    """follow redirects only to https pages on GitHub's own hosts"""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        if not _host_ok(newurl):
            raise urllib.error.URLError("redirect to a host that is not allowed: %s" % urlparse(newurl).hostname)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _open(url, version="0"):
    """urlopen with the safe redirect rule (tests replace this function)"""
    if not _host_ok(url):
        raise urllib.error.URLError("address not allowed")
    opener = urllib.request.build_opener(_SafeRedirect)
    req = urllib.request.Request(url, headers={"User-Agent": "Claudius-PDF-Editor/%s" % version,
                                               "Accept": "application/vnd.github+json, application/octet-stream"})
    return opener.open(req, timeout=TIMEOUT)


def parse_version(text):
    m = TAG_RE.match(str(text or "").strip())
    return tuple(int(x) for x in m.group(1).split(".")) if m else None


def newer(latest, current):
    a, b = parse_version(latest), parse_version(current)
    if a is None or b is None:
        return False
    n = max(len(a), len(b))
    return a + (0,) * (n - len(a)) > b + (0,) * (n - len(b))


def fetch_latest(current, opener=None):
    """-> dict(tag, notes, page, installer=(name, url, size) | None, sums_url | None) or None when nothing was published.
    Raises on network / format errors."""
    opener = opener or _open
    try:
        with opener(API_URL, current) as r:
            raw = r.read(MAX_JSON + 1)
    except urllib.error.HTTPError as ex:
        if ex.code == 404:
            return None
        raise
    if len(raw) > MAX_JSON:
        raise ValueError("answer too large")
    data = json.loads(raw.decode("utf-8", "replace"))
    if not isinstance(data, dict) or parse_version(data.get("tag_name")) is None:
        raise ValueError("unexpected answer")
    info = {"tag": str(data["tag_name"]).lstrip("v"), "notes": str(data.get("body") or "")[:3000],
            "page": RELEASE_PAGE, "installer": None, "sums_url": None}
    page = str(data.get("html_url") or "")
    if page.startswith("https://github.com/%s/" % REPO):
        info["page"] = page
    for a in data.get("assets") or []:
        if not isinstance(a, dict):
            continue
        name, url = str(a.get("name") or ""), str(a.get("browser_download_url") or "")
        size = a.get("size") if isinstance(a.get("size"), int) else 0
        if not _host_ok(url) or not url.startswith("https://github.com/%s/releases/download/" % REPO):
            continue
        if INSTALLER_RE.match(name) and 0 < size <= MAX_INSTALLER and info["installer"] is None:
            info["installer"] = (name, url, size)
        elif name == "SHA256SUMS.txt":
            info["sums_url"] = url
    return info


def download_verified(info, folder, current, progress=None, opener=None, cancel=lambda: False):
    """download the installer of `info` into `folder` and check its SHA-256 against the release's SHA256SUMS.txt.
    Returns the file path; raises (and deletes the file) when anything is wrong."""
    opener = opener or _open
    name, url, size = info["installer"]
    if not info.get("sums_url"):
        raise ValueError(_tr("This release has no checksum file (SHA256SUMS.txt), so the download is not trusted."))
    with opener(info["sums_url"], current) as r:
        sums = r.read(MAX_SUMS + 1)
    if len(sums) > MAX_SUMS:
        raise ValueError("checksum file too large")
    want = None
    for line in sums.decode("utf-8", "replace").splitlines():
        m = re.match(r"^([0-9a-fA-F]{64})\s+\*?(.+?)\s*$", line)
        if m and m.group(2) == name:
            want = m.group(1).lower()
    if not want:
        raise ValueError(_tr("The checksum of the installer is not in SHA256SUMS.txt, so the download is not trusted."))
    path = os.path.join(folder, name)
    h = hashlib.sha256()
    got = 0
    try:
        with opener(url, current) as r, open(path, "wb") as out:
            while True:
                chunk = r.read(256 * 1024)
                if not chunk:
                    break
                got += len(chunk)
                if got > size or got > MAX_INSTALLER:
                    raise ValueError("download larger than announced")
                h.update(chunk)
                out.write(chunk)
                if progress:
                    progress(got, size)
                if cancel():
                    raise RuntimeError("cancelled")
        if got != size:
            raise ValueError("download incomplete")
        if h.hexdigest() != want:
            raise ValueError(_tr("The downloaded file does not match its checksum - it was deleted."))
    except Exception:
        try:
            os.remove(path)
        except OSError:
            pass
        raise
    return path


# ------------------------------------------------------------------------------------------------ the window part
class UpdateTools:
    _upd_busy = False

    def upd_schedule(self):
        """start-up: look for a new version once a day when the user switched that on"""
        try:
            if not self.cfg.get("check_updates"):
                return
            today = datetime.date.today().isoformat()
            if self.cfg.get("update_last") == today:
                return
            self.cfg["update_last"] = today
            from claudius_v3 import save_cfg
            save_cfg(self.cfg)
            self.upd_check(manual=False)
        except Exception:
            pass

    def upd_check(self, manual=True):
        if self._upd_busy:
            return
        self._upd_busy = True
        box = {}

        def work():
            try:
                box["info"] = fetch_latest(self._upd_version())
            except Exception as ex:
                box["err"] = ex
            box["done"] = True

        threading.Thread(target=work, daemon=True).start()
        if manual:
            self.status.config(text=_tr("Looking for a new version..."))

        def poll():
            if not box.get("done"):
                self.after(150, poll)
                return
            self._upd_busy = False
            if "err" in box:
                if manual:
                    messagebox.showinfo(APP, _tr("Could not look for updates (is the computer online?).") + "\n\n%s" % box["err"])
                return
            info = box.get("info")
            if info is None or not newer(info["tag"], self._upd_version()):
                if manual:
                    self.status.config(text="")
                    messagebox.showinfo(APP, _tr("You have the latest version ({}).").replace("{}", self._upd_version()))
                return
            if not manual and self.cfg.get("update_skip") == info["tag"]:
                return
            self.upd_dialog(info)
        self.after(150, poll)

    def _upd_version(self):
        return sys.modules[type(self).__module__].APP_VERSION

    def upd_dialog(self, info):
        d = tk.Toplevel(self)
        d.title(_tr("Update available"))
        d.transient(self)
        d.configure(bg=BG, padx=18, pady=14)
        d.resizable(False, False)
        d.geometry("+%d+%d" % (self.winfo_rootx() + 220, self.winfo_rooty() + 120))
        tk.Label(d, text=_tr("A new version of Claudius PDF Editor is available"), bg=BG, fg=TH.LABEL,
                 font=(UI_FONT, 12, "bold")).pack(anchor="w")
        tk.Label(d, text=_tr("Your version: {}   -   New version: {}").replace("{}", self._upd_version(), 1).replace("{}", info["tag"], 1),
                 bg=BG, fg=TH.SECOND, font=(UI_FONT, 10)).pack(anchor="w", pady=(4, 8))
        if info["notes"].strip():
            t = tk.Text(d, width=64, height=9, wrap="word", font=(UI_FONT, 9), relief="solid", bd=1)
            t.insert("1.0", info["notes"].strip())
            t.config(state="disabled")
            t.pack(fill="x")
        msg = tk.Label(d, text="", bg=BG, fg=TH.SECOND, font=(UI_FONT, 9), anchor="w", justify="left", wraplength=470)
        msg.pack(anchor="w", pady=(8, 0), fill="x")
        bar = ttk.Progressbar(d, length=470, mode="determinate")
        row = tk.Frame(d, bg=BG)
        row.pack(fill="x", pady=(12, 0))
        can_install = sys.platform == "win32" and info["installer"] is not None
        state = {"cancel": False, "working": False}

        def skip():
            self.cfg["update_skip"] = info["tag"]
            from claudius_v3 import save_cfg
            save_cfg(self.cfg)
            d.destroy()

        def close():
            state["cancel"] = True
            d.destroy()

        def install():
            if state["working"]:
                return
            state["working"] = True
            go.config(state="disabled")
            bar.pack(pady=(8, 0))
            msg.config(text=_tr("Downloading..."))
            folder = tempfile.mkdtemp(prefix="claudius_update_")
            res = {}

            def work():
                try:
                    res["path"] = download_verified(
                        info, folder, self._upd_version(),
                        progress=lambda a, b: res.__setitem__("p", (a, b)), cancel=lambda: state["cancel"])
                except Exception as ex:
                    res["err"] = ex
                res["done"] = True
            threading.Thread(target=work, daemon=True).start()

            def poll():
                if not d.winfo_exists():
                    return
                if "p" in res:
                    a, b = res["p"]
                    bar.config(maximum=max(1, b), value=a)
                if not res.get("done"):
                    d.after(120, poll)
                    return
                state["working"] = False
                if "err" in res:
                    shutil.rmtree(folder, ignore_errors=True)
                    msg.config(text=_tr("The update was not installed:") + " %s" % res["err"], fg="#c0392b")
                    go.config(state="normal")
                    bar.pack_forget()
                    return
                msg.config(text=_tr("Downloaded and checked (SHA-256 matches)."))
                if messagebox.askyesno(APP, _tr("Claudius PDF Editor will close now and the installer will start.\n\nContinue?"), parent=d):
                    d.destroy()
                    self.upd_run_installer(res["path"])
                else:
                    bar.pack_forget()
                    go.config(state="normal")
            poll()

        go = ttk.Button(row, text=_tr("Download and install"), command=install)
        if can_install:
            go.pack(side="left")
        ttk.Button(row, text=_tr("Open release page"),
                   command=lambda: webbrowser.open(info["page"] if str(info["page"]).startswith("https://github.com/") else RELEASE_PAGE)
                   ).pack(side="left", padx=6)
        ttk.Button(row, text=_tr("Skip this version"), command=skip).pack(side="right")
        ttk.Button(row, text=_tr("Later"), command=close).pack(side="right", padx=6)
        if not can_install:
            msg.config(text=_tr("Automatic installing works on Windows with the installer from the release; use the release page to download."))
        d.protocol("WM_DELETE_WINDOW", close)
        d.bind("<Escape>", lambda e: close())

    def upd_run_installer(self, path):
        """close the program (it asks to save open work) and start the installer"""
        self.on_close()
        try:
            alive = self.winfo_exists()
        except tk.TclError:
            alive = False
        if alive:                                    # the user cancelled closing
            return
        try:
            os.startfile(path)                       # noqa - Windows only, the user pressed "Download and install"
        except Exception:
            pass
