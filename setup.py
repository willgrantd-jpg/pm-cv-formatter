"""
setup.py  —  PM CV Formatter launcher.
Called by start.bat via the bundled runtime\python.exe.
No external dependencies needed here — stdlib only.
"""

import os
import sys
import time
import subprocess
import urllib.request
import webbrowser
from pathlib import Path

APP_DIR = Path(__file__).parent
RUNTIME = APP_DIR / "runtime"
PYTHONW = RUNTIME / "pythonw.exe"
PYTHON  = RUNTIME / "python.exe"
LOG     = Path(r"C:\Users\Public\pm_cv_formatter_setup.log")

# cscript is always at this path on any modern Windows
CSCRIPT = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "cscript.exe"


def _log(msg: str):
    """Append a timestamped line to setup_log.txt in the app folder."""
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            import datetime
            f.write(f"[{datetime.datetime.now():%H:%M:%S}] {msg}\n")
    except Exception:
        pass


def say(msg):
    print(f"  {msg}", flush=True)
    _log(msg)


def fail(msg):
    _log(f"FAIL: {msg}")
    print(flush=True)
    print(f"  ERROR: {msg}", flush=True)
    print(flush=True)
    input("  Press Enter to close this window...")
    sys.exit(1)


REPO         = "willgrantd-jpg/pm-cv-formatter"
VERSION_URL  = f"https://raw.githubusercontent.com/{REPO}/main/version.txt"
ARCHIVE_URL  = f"https://github.com/{REPO}/archive/refs/heads/main.zip"

# These are never touched by an update
UPDATE_SKIP  = {"config.json", "outputs", "runtime", "setup_log.txt"}


def _kill_port_5000():
    """Kill any process currently listening on port 5000."""
    try:
        result = subprocess.run(
            ['netstat', '-ano'],
            capture_output=True, text=True
        )
        for line in result.stdout.splitlines():
            if ':5000' in line and 'LISTENING' in line:
                parts = line.strip().split()
                pid = parts[-1]
                if pid and pid.isdigit() and int(pid) > 0:
                    subprocess.run(['taskkill', '/F', '/PID', pid], capture_output=True)
                    _log(f"killed PID {pid} on port 5000 after update")
                    time.sleep(1)   # give the process time to die
                    break
    except Exception as exc:
        _log(f"_kill_port_5000 error: {exc}")


def _check_for_update():
    """Silently check GitHub for a newer version and apply it if found.
    Returns True if an update was downloaded and applied, False otherwise."""
    try:
        remote_ver = (
            urllib.request.urlopen(VERSION_URL, timeout=4)
            .read().decode().strip()
        )
        local_ver_file = APP_DIR / "version.txt"
        local_ver = local_ver_file.read_text().strip() if local_ver_file.exists() else "0.0.0"

        _log(f"version check: local={local_ver}  remote={remote_ver}")
        if remote_ver == local_ver:
            return False  # already up to date

        say(f"Update available ({local_ver} → {remote_ver}) — applying...")
        import io, zipfile

        data = urllib.request.urlopen(ARCHIVE_URL, timeout=60).read()
        prefix = f"pm-cv-formatter-main/"   # GitHub archive root folder name

        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            for member in zf.namelist():
                # strip the "pm-cv-formatter-main/" prefix
                if not member.startswith(prefix):
                    continue
                rel = member[len(prefix):]          # e.g. "app.py" or "templates/index.html"
                if not rel:
                    continue
                top = rel.split("/")[0]
                if top in UPDATE_SKIP:
                    continue

                dest = APP_DIR / rel
                if member.endswith("/"):
                    dest.mkdir(parents=True, exist_ok=True)
                else:
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    dest.write_bytes(zf.read(member))

        say(f"Updated to v{remote_ver}.")
        _log(f"update applied: {remote_ver}")
        return True

    except Exception as exc:
        _log(f"update check skipped: {exc}")
        # Offline or GitHub unreachable — just continue normally
        return False


_log("=== setup.py started ===")
_log(f"APP_DIR: {APP_DIR}")
_log(f"sys.executable: {sys.executable}")

# ─── Font installer ───────────────────────────────────────────────────────────
def _install_fonts():
    """Install bundled fonts to the user font folder — no admin rights needed."""
    import shutil, ctypes
    fonts_src = APP_DIR / "assets" / "fonts"
    fonts_dst = Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Windows" / "Fonts"
    fonts_dst.mkdir(parents=True, exist_ok=True)

    registry = {
        "Montserrat.ttf":        "Montserrat (TrueType)",
        "Montserrat-Italic.ttf": "Montserrat Italic (TrueType)",
    }
    installed_any = False
    for filename, reg_name in registry.items():
        src = fonts_src / filename
        dst = fonts_dst / filename
        if src.exists() and not dst.exists():
            shutil.copy2(str(src), str(dst))
            subprocess.run(
                ["reg", "add",
                 r"HKCU\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts",
                 "/v", reg_name, "/d", str(dst), "/f"],
                capture_output=True
            )
            installed_any = True
            _log(f"Font installed: {filename}")

    if installed_any:
        # Broadcast WM_FONTCHANGE so running apps pick up the new fonts
        try:
            result = ctypes.c_long(0)
            ctypes.windll.user32.SendMessageTimeoutW(
                0xFFFF, 0x001D, 0, 0, 2, 1000, ctypes.byref(result))
        except Exception:
            pass


# ─── Install fonts ────────────────────────────────────────────────────────────
_install_fonts()

# ─── Auto-update from GitHub ──────────────────────────────────────────────────
say("Checking for updates...")
_updated = _check_for_update()
if _updated:
    # New files are on disk — kill the running server so it restarts with the update
    say("Restarting server with new version...")
    _kill_port_5000()

# ─── Desktop shortcut (always check, so it works on re-runs too) ──────────────
def _get_desktop() -> Path:
    """Return the real Desktop folder (handles OneDrive redirect)."""
    try:
        import ctypes, ctypes.wintypes
        buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, 0x0010, None, 0, buf)
        p = Path(buf.value)
        _log(f"SHGetFolderPathW -> {p} (exists={p.exists()})")
        if p.exists():
            return p
    except Exception as e:
        _log(f"SHGetFolderPathW error: {e}")
    for candidate in [
        Path(os.path.expandvars("%USERPROFILE%")) / "OneDrive" / "Desktop",
        Path(os.path.expandvars("%USERPROFILE%")) / "Desktop",
    ]:
        _log(f"fallback candidate: {candidate} (exists={candidate.exists()})")
        if candidate.exists():
            return candidate
    fallback = Path(os.path.expandvars("%USERPROFILE%")) / "Desktop"
    _log(f"using last-resort fallback: {fallback}")
    return fallback


def _create_shortcut(lnk_path: Path, target: Path, working_dir: Path,
                     description: str) -> bool:
    """Create a .lnk file via cscript/VBScript — no execution-policy issues."""
    import tempfile
    icon = working_dir / "static" / "favicon.ico"
    icon_line = (f'oLink.IconLocation = "{icon}, 0"\n') if icon.exists() else ""
    vbs = (
        f'Set oWS = WScript.CreateObject("WScript.Shell")\n'
        f'Set oLink = oWS.CreateShortcut("{lnk_path}")\n'
        f'oLink.TargetPath = "{target}"\n'
        f'oLink.WorkingDirectory = "{working_dir}"\n'
        f'oLink.Description = "{description}"\n'
        + icon_line +
        f'oLink.Save\n'
    )
    _log(f"VBScript:\n{vbs}")
    _log(f"cscript path: {CSCRIPT} (exists={CSCRIPT.exists()})")
    tmp = None
    try:
        fd, tmp = tempfile.mkstemp(suffix=".vbs")
        os.write(fd, vbs.encode("utf-8"))
        os.close(fd)
        _log(f"VBS temp file: {tmp}")
        result = subprocess.run(
            [str(CSCRIPT), "//NoLogo", tmp],
            capture_output=True, text=True, timeout=15
        )
        _log(f"cscript returncode: {result.returncode}")
        if result.stdout:
            _log(f"cscript stdout: {result.stdout.strip()}")
        if result.stderr:
            _log(f"cscript stderr: {result.stderr.strip()}")
        created = lnk_path.exists()
        _log(f"lnk exists after run: {created}")
        return created
    except Exception as exc:
        _log(f"_create_shortcut exception: {exc}")
        return False
    finally:
        if tmp:
            try:
                os.unlink(tmp)
            except Exception:
                pass


desktop = _get_desktop()
lnk     = desktop / "PM CV Formatter.lnk"
_log(f"desktop={desktop}  lnk={lnk}  lnk.exists={lnk.exists()}")

if not lnk.exists():
    say("Creating desktop shortcut...")
    ok = _create_shortcut(
        lnk_path    = lnk,
        target      = APP_DIR / "start.bat",
        working_dir = APP_DIR,
        description = "Patrick Morgan CV Formatter",
    )
    _log(f"_create_shortcut returned: {ok}")
    if not ok:
        say("(Shortcut could not be created — check setup_log.txt)")
else:
    _log("lnk already exists, skipping shortcut creation")

# ─── Already running? ─────────────────────────────────────────────────────────
say("Checking server status...")
try:
    urllib.request.urlopen("http://localhost:5000", timeout=1)
    say("Server already running — opening browser...")
    webbrowser.open("http://localhost:5000")
    time.sleep(0.5)
    sys.exit(0)
except Exception:
    pass

# ─── Start Flask server ───────────────────────────────────────────────────────
say("Starting server...")
launcher = str(PYTHONW) if PYTHONW.exists() else str(PYTHON)
_log(f"launcher: {launcher}")
try:
    subprocess.Popen(
        [launcher, str(APP_DIR / "app.py")],
        cwd=str(APP_DIR),
        creationflags=0x08000000,   # CREATE_NO_WINDOW
        close_fds=True
    )
except Exception as exc:
    fail(f"Could not start server: {exc}")

# ─── Wait for server to respond ────────────────────────────────────────────────
say("Waiting for server to start...")
ready = False
for i in range(25):
    time.sleep(0.8)
    try:
        urllib.request.urlopen("http://localhost:5000", timeout=1)
        ready = True
        break
    except Exception:
        pass

if not ready:
    fail(
        "Server did not start in time.\n"
        "  Try running start.bat again.\n"
        "  If the problem continues, contact Will."
    )

# ─── Open browser ─────────────────────────────────────────────────────────────
say("Opening PM CV Formatter in your browser...")
_log("Opening browser at http://localhost:5000")
webbrowser.open("http://localhost:5000")
time.sleep(1)
_log("=== setup.py done ===")
