"""
Claudius PDF Editor - one program window for many files.

When several PDFs are opened together (select 4 files in Explorer > Open, or double-click them one after another),
Windows starts the program once per file.  This module lets the first copy keep running and take the files of the
others as new tabs; the other copies hand their file names over and close at once.

How it works (all on this computer, nothing goes to the internet):
  * the running copy listens on 127.0.0.1 only (a random port) and writes the port and a random secret into
    ~/.claudius/instance.json
  * a second copy reads that file, connects, sends the secret and the file names, and exits
  * only existing files are accepted, at most 50 per message, and a message needs the secret
Switch it off in Settings (Several files: "Open files as tabs in the running window").
"""
import atexit
import hmac
import json
import os
import queue
import secrets
import socket
import threading
import time

from claudius_extra import CONFIG_DIR

INFO = os.path.join(CONFIG_DIR, "instance.json")
LOCK = os.path.join(CONFIG_DIR, "instance.start")
MAX_PATHS = 50
MAX_BYTES = 200000


def _read_info():
    try:
        with open(INFO, "r", encoding="utf-8") as fh:
            d = json.load(fh)
        if isinstance(d, dict) and isinstance(d.get("port"), int) and isinstance(d.get("token"), str):
            return d
    except Exception:
        pass
    return None


def send_to_running(paths):
    """hand `paths` to a copy that is already running.  True when it took them."""
    d = _read_info()
    if not d:
        return False
    try:
        with socket.create_connection(("127.0.0.1", d["port"]), timeout=2) as s:
            s.settimeout(3)
            msg = json.dumps({"token": d["token"], "paths": [os.path.abspath(p) for p in paths][:MAX_PATHS]}) + "\n"
            s.sendall(msg.encode("utf-8"))
            return s.recv(8).startswith(b"ok")
    except Exception:
        return False


class Receiver:
    """the listening side: collects the file names that other copies send"""

    def __init__(self):
        self.q = queue.Queue()
        self.token = secrets.token_hex(24)
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.bind(("127.0.0.1", 0))                      # this computer only
        self.sock.listen(8)
        self.port = self.sock.getsockname()[1]
        os.makedirs(CONFIG_DIR, exist_ok=True)
        tmp = INFO + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump({"port": self.port, "token": self.token, "pid": os.getpid()}, fh)
        os.replace(tmp, INFO)
        threading.Thread(target=self._serve, daemon=True).start()
        atexit.register(self.close)                           # also removes instance.json when the program ends without the normal close

    def _serve(self):
        while True:
            try:
                conn, _addr = self.sock.accept()
            except OSError:
                return
            try:
                conn.settimeout(2)
                buf = b""
                while b"\n" not in buf and len(buf) < MAX_BYTES:
                    chunk = conn.recv(8192)
                    if not chunk:
                        break
                    buf += chunk
                d = json.loads(buf.split(b"\n", 1)[0].decode("utf-8"))
                if isinstance(d, dict) and isinstance(d.get("token"), str) and hmac.compare_digest(d["token"], self.token):
                    paths = [p for p in (d.get("paths") or [])[:MAX_PATHS]
                             if isinstance(p, str) and 0 < len(p) < 4096 and os.path.isfile(p)]
                    for p in paths:
                        self.q.put(p)
                    conn.sendall(b"ok")
            except Exception:
                pass
            finally:
                try:
                    conn.close()
                except OSError:
                    pass

    def take(self):
        out = []
        while True:
            try:
                out.append(self.q.get_nowait())
            except queue.Empty:
                return out

    def close(self):
        try:
            self.sock.close()
        except OSError:
            pass
        try:
            d = _read_info()
            if d and d.get("pid") == os.getpid():
                os.remove(INFO)
        except OSError:
            pass


def start(paths):
    """Call first thing in main().  Returns (receiver, handed_over):
         handed_over True  -> another copy took the files: just exit
         receiver not None -> this copy is the main one and listens for files from later copies
         both False/None   -> work alone (the server could not start)"""
    os.makedirs(CONFIG_DIR, exist_ok=True)
    deadline = time.time() + 8
    while True:
        if paths and send_to_running(paths):
            return None, True
        # nobody is running (or still starting): one copy wins the right to start, the others wait for it
        try:
            fd = os.open(LOCK, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
        except FileExistsError:
            try:
                if time.time() - os.path.getmtime(LOCK) > 10:       # left behind by a copy that crashed
                    os.remove(LOCK)
                    continue
            except OSError:
                pass
            if paths and time.time() < deadline:
                time.sleep(0.15)
                continue
            return None, False
        except OSError:
            return None, False
        try:
            if not paths and _read_info() and _alive():                # no files given: a normal extra window
                return None, False
            return Receiver(), False
        except Exception:
            return None, False
        finally:
            try:
                os.remove(LOCK)
            except OSError:
                pass


def _alive():
    d = _read_info()
    if not d:
        return False
    try:
        with socket.create_connection(("127.0.0.1", d["port"]), timeout=1):
            return True
    except OSError:
        return False
