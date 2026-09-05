"""
TrayMinimizer — Automatically minimize applications to the system tray.

Usage:
  Watch mode (no args):
    python tray_minimizer.py
    Right-click the tray icon to add/remove watched apps.

  Launch mode (with args):
    python tray_minimizer.py <program> [args...]
    Launches the program and immediately hides its window to the tray.
    Works with GUI apps and console/terminal apps alike.

Dependencies: pip install pystray Pillow pywin32 psutil
"""

import ctypes
import ctypes.wintypes
import json
import os
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import simpledialog

import psutil
import pystray
from PIL import Image, ImageDraw
import win32gui
import win32con
import win32process
import win32ui

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tray_minimizer.json")

# WinEvent constants
EVENT_OBJECT_SHOW = 0x8002
EVENT_SYSTEM_FOREGROUND = 0x0003
EVENT_SYSTEM_MINIMIZESTART = 0x0016
WINEVENT_OUTOFCONTEXT = 0x0000
WINEVENT_SKIPOWNPROCESS = 0x0002
OBJID_WINDOW = 0

# Process creation flags
CREATE_NEW_CONSOLE = 0x00000010

# Windows whose title starts with this marker are NEVER auto-hidden to the
# tray.  It lets a launched program keep a transient window visible while the
# rest of it sits hidden in the tray.  orchestrator2's --resume/--copy session
# picker sets this prefix on its own console window (via SetConsoleTitleW and
# its Textual app title) so the picker stays on screen while the orchestrator2
# server itself remains hidden.
NO_HIDE_TITLE_PREFIX = "[nohide]"

# Callback type for SetWinEventHook
WinEventProcType = ctypes.WINFUNCTYPE(
    None,
    ctypes.wintypes.HANDLE,   # hWinEventHook
    ctypes.wintypes.DWORD,    # event
    ctypes.wintypes.HWND,     # hwnd
    ctypes.wintypes.LONG,     # idObject
    ctypes.wintypes.LONG,     # idChild
    ctypes.wintypes.DWORD,    # idEventThread
    ctypes.wintypes.DWORD,    # dwmsEventTime
)

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# Fix ctypes return/argument types for console and hook functions.
# Without these, 64-bit HWND/HANDLE values can be truncated to 32-bit c_int,
# corrupting handles and causing silent failures.
kernel32.GetConsoleWindow.restype = ctypes.wintypes.HWND
kernel32.GetConsoleWindow.argtypes = []
kernel32.AttachConsole.restype = ctypes.wintypes.BOOL
kernel32.AttachConsole.argtypes = [ctypes.wintypes.DWORD]
kernel32.FreeConsole.restype = ctypes.wintypes.BOOL
kernel32.FreeConsole.argtypes = []
user32.SetWinEventHook.restype = ctypes.wintypes.HANDLE

# AttachConsole/FreeConsole are per-process (not per-thread), so concurrent
# callers would corrupt each other's console state.  Serialize all access.
_console_lock = threading.Lock()


def _extract_exe_icon(exe_path, size=64):
    """Extract the icon from an executable and return a PIL Image, or None."""
    try:
        resolved = shutil.which(exe_path) or exe_path
        if not os.path.isfile(resolved):
            return None

        large, small = win32gui.ExtractIconEx(resolved, 0, 1)
        if not large and not small:
            return None
        hicon = large[0] if large else small[0]

        try:
            dc_screen = win32gui.GetDC(0)
            dc = win32ui.CreateDCFromHandle(dc_screen)
            dc_mem = dc.CreateCompatibleDC()

            bmp = win32ui.CreateBitmap()
            bmp.CreateCompatibleBitmap(dc, size, size)
            old = dc_mem.SelectObject(bmp)

            # Clear to black-transparent, then draw the icon over it
            dc_mem.FillSolidRect((0, 0, size, size), 0)
            win32gui.DrawIconEx(
                dc_mem.GetHandleOutput(), 0, 0, hicon,
                size, size, 0, None, win32con.DI_NORMAL,
            )

            bmpinfo = bmp.GetInfo()
            bmpbits = bmp.GetBitmapBits(True)
            img = Image.frombuffer(
                "RGBA", (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
                bmpbits, "raw", "BGRA", 0, 1,
            )

            dc_mem.SelectObject(old)
            dc_mem.DeleteDC()
            win32gui.ReleaseDC(0, dc_screen)
        finally:
            for h in list(large or ()) + list(small or ()):
                win32gui.DestroyIcon(h)

        return img
    except Exception:
        return None


def _snapshot_windows():
    """Return set of all current top-level window handles."""
    hwnds = set()
    def cb(hwnd, _):
        hwnds.add(hwnd)
    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass
    return hwnds


# Window classes that are never a program's real window.  EnumWindows returns
# *owned* top-level windows too, and every process that talks to a console gets
# a 0x0, disabled "Default IME" helper window owned by conhost.exe.  Hiding one
# of those produces a tray icon that restores nothing — the bug that used to
# leave dozens of dead TrayMinimizer icons in the notification area.
_PHANTOM_WINDOW_CLASSES = frozenset({
    "ime",
    "msctfime ui",
    "default ime",
    "tooltips_class32",
    "sysshadow",
    "workerw",
    "progman",
    "shell_traywnd",
})

WS_EX_NOACTIVATE = 0x08000000

# A console window is the thing we want in the overwhelmingly common case
# (launch mode wraps `cmd`), so it outranks any other candidate.
_CONSOLE_WINDOW_CLASSES = frozenset({"consolewindowclass", "pseudoconsolewindow"})


def _is_console_window(hwnd):
    """True when *hwnd* is a console window (whoever hosts it)."""
    try:
        return win32gui.GetClassName(hwnd).lower() in _CONSOLE_WINDOW_CLASSES
    except Exception:
        return False


def _window_is_restorable(hwnd, *, allow_hidden=False):
    """True when *hwnd* is a genuine top-level window a user could restore.

    *allow_hidden* exists because the two kinds of window we hide behave
    differently under STARTUPINFO's SW_HIDE:

    * A **console** window really is created hidden and stays that way, so
      visibility can never be required for it.  That is safe because consoles
      are identified by AttachConsole/GetConsoleWindow, which names the right
      window outright rather than guessing.
    * A **GUI** window has to be guessed at by PID, and toolkits leave
      invisible helper windows lying around that are indistinguishable from a
      real one by style alone — Tk's 1920x1015 "TtkMonitorWindow" has
      WS_CAPTION, WS_SYSMENU, WS_THICKFRAME, no owner and no WS_EX_TOOLWINDOW.
      What separates it from the real window is that it is never shown.  GUI
      toolkits call ShowWindow themselves, so a real main window does become
      visible even when the process was started with SW_HIDE (measured: Tk's
      main window is visible, its monitor window is not).

    So: consoles are exempt from the visibility test, GUI windows are not.
    """
    try:
        if not win32gui.IsWindow(hwnd):
            return False
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        if style & win32con.WS_CHILD:
            return False
        if style & win32con.WS_DISABLED:
            return False
        if ex_style & win32con.WS_EX_TOOLWINDOW:
            return False
        if ex_style & WS_EX_NOACTIVATE:
            return False
        # An owned window (GetParent returns the owner for non-child windows)
        # is a dialog/helper, not the program's main window.
        if win32gui.GetParent(hwnd):
            return False
        if win32gui.GetClassName(hwnd).lower() in _PHANTOM_WINDOW_CLASSES:
            return False
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        if right - left <= 0 or bottom - top <= 0:
            return False
        if (not allow_hidden
                and not _is_console_window(hwnd)
                and not win32gui.IsWindowVisible(hwnd)):
            return False
        return True
    except Exception:
        return False


def _window_rank(hwnd):
    """Sort key for candidate windows — lower is better.

    Console windows win, because launch mode nearly always wraps `cmd`.
    """
    return 0 if _is_console_window(hwnd) else 1


def _describe_window(hwnd):
    """Human-readable one-liner about a window, for the log."""
    try:
        _, wpid = win32process.GetWindowThreadProcessId(hwnd)
        return (f"hwnd={hwnd} cls={win32gui.GetClassName(hwnd)!r} "
                f"title={win32gui.GetWindowText(hwnd)!r} pid={wpid}")
    except Exception:
        return f"hwnd={hwnd} <unreadable>"


def _find_window_by_pid(pid):
    """Find the best real top-level window owned by the given PID, or None.

    Hidden windows are accepted (launch mode's child is born SW_HIDE); the
    _window_is_restorable() gate is what keeps phantoms out.
    """
    result = []
    def cb(hwnd, _):
        try:
            if not win32gui.GetWindowText(hwnd):
                return
            _, wpid = win32process.GetWindowThreadProcessId(hwnd)
            if wpid != pid:
                return
            if not _window_is_restorable(hwnd):
                return
            result.append(hwnd)
        except Exception:
            pass
    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass
    if not result:
        return None
    result.sort(key=_window_rank)
    return result[0]


def _get_process_tree_pids(root_pid):
    """Return a set of PIDs: the root process and all descendants."""
    pids = set()
    try:
        root = psutil.Process(root_pid)
        pids.add(root_pid)
        for child in root.children(recursive=True):
            pids.add(child.pid)
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        pids.add(root_pid)
    return pids


def _find_console_hwnd(pid):
    """Find the console window for a process via AttachConsole.

    Thread-safe: serialized by _console_lock.  Expects the calling process
    to have *no* console of its own (see TrayMinimizer.run() which detaches
    early in launch mode).  If the caller does still own a console, we
    detach/reattach around the probe.
    """
    with _console_lock:
        had_console = kernel32.GetConsoleWindow()
        if had_console:
            kernel32.FreeConsole()
        hwnd = None
        if kernel32.AttachConsole(pid):
            hwnd = kernel32.GetConsoleWindow()
            kernel32.FreeConsole()
        if had_console:
            kernel32.AttachConsole(0xFFFFFFFF)  # ATTACH_PARENT_PROCESS
        return hwnd


def _title_is_no_hide(hwnd):
    """True when the window's title marks it as never-hide.

    See NO_HIDE_TITLE_PREFIX.  A window can set this after it's created
    (SetConsoleTitleW / a TUI setting its own title), so this is re-checked
    at every hide decision rather than cached.
    """
    try:
        title = win32gui.GetWindowText(hwnd)
    except Exception:
        return False
    return bool(title) and title.startswith(NO_HIDE_TITLE_PREFIX)


class TrayMinimizer:
    # How long the authoritative AttachConsole probe gets to itself before the
    # guessing strategies are allowed to run.  A console window takes a few
    # tens of milliseconds to exist after CreateProcess; without this head
    # start the fallbacks fire into that gap and can only find phantoms.
    _FALLBACK_GRACE = 2.0
    _WATCHDOG_INTERVAL = 10       # seconds between watchdog passes
    _READOPT_TIMEOUT = 2.0        # per-pass budget when re-adopting a window
    # Empty passes before the icon gives up.  Deliberately generous: giving up
    # is right for a program that will never have a window, but wrong for one
    # that is merely slow to show its first one, and only the passage of time
    # tells those apart.
    _MAX_EMPTY_CHECKS = 6

    def __init__(self):
        self.config = self._load_config()
        # hwnd -> {"title", "exe", "pid", "launch_pid"}
        self.hidden_windows = {}
        self.lock = threading.Lock()
        self.running = True
        self.icon = None
        self.known_hwnds = set()
        self._hook_proc = WinEventProcType(self._win_event_callback)
        self._hooks = []
        self._launched_procs = []  # Popen objects from launch mode
        self._proc_exe_names = {}  # pid -> exe name, for re-adoption
        self._launch_mode = False
        # Console window inherited from our launcher, hidden in launch
        # mode and handed back in _restore_own_console().
        self._inherited_console_hwnd = None
        # Set once launch-mode's first window search has finished, so the
        # watchdog doesn't judge an icon before detection has had its say.
        self._initial_detect_done = threading.Event()
        # pystray rebuilds the native HMENU on every update; doing that from
        # several threads at once corrupts it.
        self._menu_lock = threading.Lock()
        # Windows that were restored but should re-hide when minimized.
        # hwnd -> {"title", "exe", "pid", "launch_pid"}
        self._watched_hwnds = {}

    # ── Config ────────────────────────────────────────────────────────

    def _load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        default = {"apps": []}
        self._save_config(default)
        return default

    def _save_config(self, config=None):
        with open(CONFIG_FILE, "w") as f:
            json.dump(config or self.config, f, indent=2)

    # ── Tray icon image ──────────────────────────────────────────────

    def _create_icon_image(self):
        img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle([4, 4, 60, 60], radius=8,
                               fill=(50, 120, 200), outline=(30, 80, 160), width=2)
        draw.polygon([(20, 22), (44, 22), (32, 42)], fill="white")
        draw.rectangle([18, 47, 46, 51], fill="white")
        return img

    # ── Window helpers ────────────────────────────────────────────────

    def _get_exe_for_hwnd(self, hwnd):
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            proc = psutil.Process(pid)
            return proc.name().lower(), pid
        except Exception:
            return None, None

    def _is_app_window(self, hwnd):
        """Watch-mode test: a *visible* window worth auto-hiding.

        Same phantom filtering as launch mode (_window_is_restorable), plus a
        visibility requirement — in watch mode we react to windows appearing
        on screen, so an invisible one is never a candidate.
        """
        if not win32gui.IsWindowVisible(hwnd):
            return False
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        if not (style & win32con.WS_VISIBLE):
            return False
        if not win32gui.GetWindowText(hwnd):
            return False
        return _window_is_restorable(hwnd)

    def _hide_window(self, hwnd, exe_override=None, launch_pid=None):
        """Hide *hwnd* to the tray and record it.

        *launch_pid* is the process we started in launch mode; it is tracked
        separately from the window's real owner pid because the window often
        belongs to a descendant (or to conhost.exe) rather than to the process
        whose exit we are waiting on.
        """
        if not win32gui.IsWindow(hwnd):
            return
        # Never hide a window that has marked itself no-hide (e.g. the
        # orchestrator2 session picker).  This is the single chokepoint every
        # hide path funnels through, so the guard here covers launch-mode
        # detection, the event hook, and the minimize re-hide alike.
        if _title_is_no_hide(hwnd):
            _log(f"skip hide: hwnd={hwnd} is marked no-hide")
            return
        # Last line of defence: never put a phantom (conhost's owned 0x0
        # "Default IME" helper and friends) in the tray.  A tray entry that
        # restores an invisible window is indistinguishable, to the user, from
        # a tray icon that is simply broken.
        if not _window_is_restorable(hwnd, allow_hidden=True):
            _log(f"skip hide: not a restorable window — {_describe_window(hwnd)}")
            return
        title = win32gui.GetWindowText(hwnd)
        owner_exe, owner_pid = self._get_exe_for_hwnd(hwnd)
        exe = exe_override or owner_exe
        if not exe:
            return
        _log(f"hiding {_describe_window(hwnd)} exe={exe} launch_pid={launch_pid}")
        # Window may already be hidden (started with SW_HIDE); calling
        # ShowWindow(SW_HIDE) on an already-hidden window is a harmless no-op.
        win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
        with self.lock:
            self.hidden_windows[hwnd] = {
                "title": title or exe, "exe": exe,
                "pid": owner_pid, "launch_pid": launch_pid,
            }
        self._update_menu()

    def _force_foreground(self, hwnd):
        """Show a hidden/minimized window and forcibly bring it to the front.

        A plain SetForegroundWindow() silently fails when our process does not
        own the foreground (Windows' foreground-lock).  The window then gets
        shown *behind* whatever is currently focused (e.g. a fullscreen
        browser) and the user thinks "nothing happened".  Attaching our input
        queue to the current foreground thread lifts the lock so the restore
        actually surfaces the window.
        """
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        except Exception:
            pass
        try:
            cur_tid = kernel32.GetCurrentThreadId()
            fg = win32gui.GetForegroundWindow()
            fg_tid = win32process.GetWindowThreadProcessId(fg)[0] if fg else 0
            tgt_tid = win32process.GetWindowThreadProcessId(hwnd)[0]
            attached = []
            for tid in {fg_tid, tgt_tid}:
                if tid and tid != cur_tid and user32.AttachThreadInput(cur_tid, tid, True):
                    attached.append(tid)
            try:
                win32gui.BringWindowToTop(hwnd)
                win32gui.SetForegroundWindow(hwnd)
                win32gui.SetActiveWindow(hwnd)
            finally:
                for tid in attached:
                    user32.AttachThreadInput(cur_tid, tid, False)
        except Exception:
            pass

    def _restore_window(self, hwnd):
        # Self-heal: if the backing process has exited its window handle is
        # dead, so "restoring" would silently do nothing and leave a zombie
        # entry in the tray.  Prune it instead.
        if not win32gui.IsWindow(hwnd):
            with self.lock:
                self.hidden_windows.pop(hwnd, None)
                self._watched_hwnds.pop(hwnd, None)
            self.known_hwnds.discard(hwnd)
            _log(f"restore: hwnd={hwnd} is no longer a valid window; pruned")
            self._update_menu()
            return

        self._force_foreground(hwnd)
        with self.lock:
            info = self.hidden_windows.pop(hwnd, None)
            # Remember this window so we re-hide it if the user minimizes it.
            if info:
                self._watched_hwnds[hwnd] = info
        try:
            shown = win32gui.IsWindowVisible(hwnd)
        except Exception:
            shown = False
        _log(f"restore: {_describe_window(hwnd)} -> "
             f"{'visible' if shown else 'STILL NOT VISIBLE'}")
        self._update_menu()

    def _restore_all(self):
        with self.lock:
            hwnds = list(self.hidden_windows.keys())
        for hwnd in hwnds:
            self._restore_window(hwnd)

    # ── Launch mode ───────────────────────────────────────────────────

    def launch_and_hide(self, cmd_args):
        """Launch a program and hide its window to the tray."""
        exe_name = os.path.basename(cmd_args[0]).lower()

        # Snapshot all windows before launching
        before = _snapshot_windows()

        # Start with the console window born hidden (SW_HIDE) so the user
        # never sees a flash.  The window still exists and can be restored
        # from the tray later.
        si = subprocess.STARTUPINFO()
        si.dwFlags = subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = win32con.SW_HIDE
        proc = subprocess.Popen(
            cmd_args,
            creationflags=CREATE_NEW_CONSOLE,
            startupinfo=si,
        )
        self._launched_procs.append(proc)
        self._proc_exe_names[proc.pid] = exe_name

        _log(f"launched pid={proc.pid} exe={exe_name}")

        # Start a thread to find and hide the window
        def detect():
            try:
                self._find_and_hide_launched(proc, exe_name, before)
            finally:
                self._initial_detect_done.set()

        threading.Thread(target=detect, daemon=True).start()

        # Monitor the process — clean up tray entry and optionally exit when done
        threading.Thread(
            target=self._monitor_process,
            args=(proc,),
            daemon=True,
        ).start()

    def _find_and_hide_launched(self, proc, exe_name, windows_before,
                                timeout=20, quiet=False, grace=None):
        """Find the window created by a launched process and hide it.

        Returns True when a window was found and hidden.

        The child is started with SW_HIDE, so the window we are looking for is
        legitimately invisible and visibility cannot be used as a filter.
        Everything therefore hinges on picking the *right* window:

        1. AttachConsole — authoritative.  GetConsoleWindow() names the exact
           console window of the process, whoever owns it (conhost.exe,
           OpenConsole.exe, ...).  This is the only strategy that cannot pick
           the wrong window, so it is given a head start (_FALLBACK_GRACE)
           before the guessing strategies are allowed to run at all.
        2. Window search by PID — catches GUI children that have no console.
        3. Window-list diff — catches windows owned by a console host process
           that is not in our process tree.

        Strategies 2 and 3 used to fire on the very first pass, microseconds
        after CreateProcess, when the console window does not exist yet.  The
        only window around at that instant is conhost's owned 0x0 "Default
        IME" helper, which passed the old filters and got hidden in place of
        the real console — producing a tray icon that restores nothing.  Both
        now go through _window_is_restorable() and only after strategy 1 has
        had its chance.
        """
        deadline = time.time() + timeout
        # The head start only matters right after CreateProcess.  Re-adoption
        # runs against a process that has been up for a while, so it passes
        # grace=0 and may use every strategy immediately.
        if grace is None:
            grace = self._FALLBACK_GRACE
        fallback_at = time.time() + grace

        while time.time() < deadline:
            if proc.poll() is not None:
                if not quiet:
                    _log(f"detect: pid={proc.pid} exited before a window appeared")
                return False

            # Refresh the process tree (children may spawn after launch).
            # Sort, and probe the process we actually launched first, so the
            # outcome does not depend on set iteration order.
            tree_pids = _get_process_tree_pids(proc.pid)
            ordered_pids = [proc.pid] + sorted(tree_pids - {proc.pid})

            # ── Strategy 1: AttachConsole (authoritative) ──────────────
            for pid in ordered_pids:
                hwnd = _find_console_hwnd(pid)
                if not hwnd or hwnd in self.known_hwnds:
                    continue
                if not win32gui.IsWindow(hwnd):
                    continue
                if _title_is_no_hide(hwnd):
                    continue  # e.g. the orchestrator2 picker console
                _log(f"detect: console probe matched {_describe_window(hwnd)} "
                     f"(via pid={pid})")
                self.known_hwnds.add(hwnd)
                self._hide_window(hwnd, exe_override=exe_name,
                                  launch_pid=proc.pid)
                return True

            if time.time() >= fallback_at:
                # ── Strategy 2: a real window owned by the process tree ──
                for pid in ordered_pids:
                    hwnd = _find_window_by_pid(pid)
                    if not hwnd or hwnd in self.known_hwnds:
                        continue
                    if _title_is_no_hide(hwnd):
                        continue
                    _log(f"detect: pid search matched {_describe_window(hwnd)}")
                    self.known_hwnds.add(hwnd)
                    self._hide_window(hwnd, exe_override=exe_name,
                                      launch_pid=proc.pid)
                    return True

                # ── Strategy 3: diff the window list ─────────────────────
                new_hwnds = _snapshot_windows() - windows_before
                candidates = []
                for hwnd in new_hwnds:
                    if hwnd in self.known_hwnds:
                        continue
                    if _title_is_no_hide(hwnd):
                        continue  # leave marked windows (e.g. the picker) alone
                    if not _window_is_restorable(hwnd):
                        continue
                    try:
                        _, wpid = win32process.GetWindowThreadProcessId(hwnd)
                    except Exception:
                        continue
                    if wpid in tree_pids:
                        candidates.append(hwnd)
                        continue
                    # The window may belong to the conhost.exe /
                    # OpenConsole.exe hosting our child's console.  Verify by
                    # checking that AttachConsole on our child yields it.
                    try:
                        if psutil.Process(wpid).name().lower() in (
                                "conhost.exe", "openconsole.exe"):
                            if _find_console_hwnd(proc.pid) == hwnd:
                                candidates.append(hwnd)
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
                if candidates:
                    # Deterministic: console windows first, then by handle.
                    candidates.sort(key=lambda h: (_window_rank(h), h))
                    hwnd = candidates[0]
                    _log(f"detect: window diff matched {_describe_window(hwnd)}")
                    self.known_hwnds.add(hwnd)
                    self._hide_window(hwnd, exe_override=exe_name,
                                      launch_pid=proc.pid)
                    return True

            time.sleep(0.2)

        if not quiet:
            _log(f"WARNING: could not find a window for {exe_name} "
                 f"(pid={proc.pid}) within {timeout}s")
        return False

    def _monitor_process(self, proc):
        """Wait for a launched process to exit, then clean up its tray entry."""
        proc.wait()
        _log(f"process {proc.pid} exited, code={proc.returncode}")
        # Give a moment for any final window cleanup
        time.sleep(0.5)

        with self.lock:
            stale = [h for h, info in self.hidden_windows.items()
                     if info.get("launch_pid") == proc.pid
                     or not win32gui.IsWindow(h)]
            for h in stale:
                del self.hidden_windows[h]
                self.known_hwnds.discard(h)
            for h in [h for h, info in self._watched_hwnds.items()
                      if info.get("launch_pid") == proc.pid
                      or not win32gui.IsWindow(h)]:
                del self._watched_hwnds[h]
            # Decide whether to auto-exit while still holding the lock,
            # so another thread can't sneak a new entry in between.
            should_exit = (
                self._launch_mode
                and all(p.poll() is not None for p in self._launched_procs)
                and not self.hidden_windows
            )

        self._update_menu()
        if should_exit:
            _log("all launched processes exited, auto-exiting")
            self._exit()

    # ── Windows event hook (watch mode) ───────────────────────────────

    def _win_event_callback(self, hWinEventHook, event, hwnd, idObject,
                            idChild, idEventThread, dwmsEventTime):
        if idObject != OBJID_WINDOW:
            return
        if not hwnd:
            return

        # Re-hide a restored window when the user minimizes it.
        if event == EVENT_SYSTEM_MINIMIZESTART:
            with self.lock:
                info = self._watched_hwnds.pop(hwnd, None)
            if info:
                threading.Timer(
                    0.15, self._hide_window,
                    args=(hwnd,),
                    kwargs={"exe_override": info["exe"],
                            "launch_pid": info.get("launch_pid")},
                ).start()
            return

        if hwnd in self.known_hwnds:
            return
        if not self._is_app_window(hwnd):
            return
        if _title_is_no_hide(hwnd):
            return

        apps = [a.lower() for a in self.config.get("apps", [])]
        if not apps:
            return

        exe, pid = self._get_exe_for_hwnd(hwnd)
        if exe and exe in apps:
            self.known_hwnds.add(hwnd)
            threading.Timer(0.15, self._hide_window, args=(hwnd,)).start()

    def _hook_thread(self):
        hook1 = user32.SetWinEventHook(
            EVENT_OBJECT_SHOW, EVENT_OBJECT_SHOW,
            0, self._hook_proc, 0, 0,
            WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS,
        )
        hook2 = user32.SetWinEventHook(
            EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND,
            0, self._hook_proc, 0, 0,
            WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS,
        )
        hook3 = user32.SetWinEventHook(
            EVENT_SYSTEM_MINIMIZESTART, EVENT_SYSTEM_MINIMIZESTART,
            0, self._hook_proc, 0, 0,
            WINEVENT_OUTOFCONTEXT,  # no SKIPOWNPROCESS — we need our own windows
        )
        self._hooks = [hook1, hook2, hook3]

        msg = ctypes.wintypes.MSG()
        while self.running:
            result = user32.GetMessageW(ctypes.byref(msg), 0, 0, 0)
            if result <= 0:
                break
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

        for hook in self._hooks:
            if hook:
                user32.UnhookWinEvent(hook)

    # ── Watchdog: no dead tray icons ──────────────────────────────────

    def _watchdog_thread(self):
        """Keep the tray icon honest.

        Prunes windows that no longer exist, and — in launch mode — makes sure
        the icon never outlives its usefulness.  An icon whose window has been
        destroyed while the launched process keeps running (a `cmd /K` whose
        inner program exited, a GUI child that was closed) is exactly the
        "tray icon that does nothing when I click it" the user sees.  When
        that happens we re-run detection to adopt whatever window the process
        does have now, and if it has none at all we stop, rather than sitting
        in the notification area forever with nothing to show.
        """
        misses = 0
        while self.running:
            time.sleep(self._WATCHDOG_INTERVAL)
            if not self.running:
                return

            with self.lock:
                stale = [h for h in self.hidden_windows
                         if not win32gui.IsWindow(h)]
                for h in stale:
                    del self.hidden_windows[h]
                    self.known_hwnds.discard(h)
                stale_watched = [h for h in self._watched_hwnds
                                 if not win32gui.IsWindow(h)]
                for h in stale_watched:
                    del self._watched_hwnds[h]
                    self.known_hwnds.discard(h)
                tracked = len(self.hidden_windows) + len(self._watched_hwnds)
            if stale or stale_watched:
                _log(f"watchdog: pruned {len(stale) + len(stale_watched)} "
                     f"dead window(s)")
                self._update_menu()

            if not self._launch_mode or tracked:
                misses = 0
                continue

            # Initial detection may still be searching (a program can take a
            # while to put up its first window); don't pull the rug out.
            if not self._initial_detect_done.is_set():
                continue

            # Launch mode with nothing tracked.  Try to (re-)adopt a window
            # from any still-running launched process.
            live = [p for p in self._launched_procs if p.poll() is None]
            if not live:
                # _monitor_process handles the all-exited case; if it somehow
                # did not, do not linger.
                continue
            adopted = False
            for proc in live:
                if self._find_and_hide_launched(
                        proc, self._exe_name_for(proc), set(),
                        timeout=self._READOPT_TIMEOUT, quiet=True, grace=0):
                    _log(f"watchdog: re-adopted a window for pid={proc.pid}")
                    adopted = True
                    break
            if adopted:
                misses = 0
                continue

            misses += 1
            _log(f"watchdog: nothing to show for {misses} check(s) "
                 f"({self._WATCHDOG_INTERVAL}s apart)")
            if misses >= self._MAX_EMPTY_CHECKS:
                _log("watchdog: launched process has no window to restore; "
                     "exiting so a dead icon is not left in the tray")
                self._exit()
                return

    def _exe_name_for(self, proc):
        return self._proc_exe_names.get(proc.pid, "?")

    # ── Add / remove app dialogs ──────────────────────────────────────

    def _add_app_dialog(self):
        def dialog():
            root = tk.Tk()
            root.withdraw()
            exe = simpledialog.askstring(
                "Add Application",
                "Enter the executable name (e.g. notepad.exe):",
                parent=root,
            )
            root.destroy()
            if exe and exe.strip():
                exe = exe.strip().lower()
                if exe not in [a.lower() for a in self.config["apps"]]:
                    self.config["apps"].append(exe)
                    self._save_config()
                    self._update_menu()

        threading.Thread(target=dialog, daemon=True).start()

    def _pick_running_app_dialog(self):
        def dialog():
            apps_found = {}

            def enum_cb(hwnd, _):
                if not self._is_app_window(hwnd):
                    return
                exe, _ = self._get_exe_for_hwnd(hwnd)
                if exe and exe not in apps_found:
                    title = win32gui.GetWindowText(hwnd)
                    apps_found[exe] = title

            try:
                win32gui.EnumWindows(enum_cb, None)
            except Exception:
                pass

            if not apps_found:
                return

            root = tk.Tk()
            root.title("Pick a running application")
            root.geometry("420x350")
            root.resizable(False, False)

            frame = tk.Frame(root)
            frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

            scrollbar = tk.Scrollbar(frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set, font=("Consolas", 10))
            listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=listbox.yview)

            sorted_apps = sorted(apps_found.items())
            for exe, title in sorted_apps:
                display = f"{exe}  —  {title[:50]}" if title else exe
                listbox.insert(tk.END, display)

            def on_add():
                sel = listbox.curselection()
                if sel:
                    exe = sorted_apps[sel[0]][0]
                    if exe not in [a.lower() for a in self.config["apps"]]:
                        self.config["apps"].append(exe)
                        self._save_config()
                        self._update_menu()
                root.destroy()

            btn = tk.Button(root, text="Add Selected", command=on_add)
            btn.pack(pady=(0, 8))

            root.mainloop()

        threading.Thread(target=dialog, daemon=True).start()

    # ── Menu building ─────────────────────────────────────────────────

    def _build_menu(self):
        hidden_items = []
        with self.lock:
            for hwnd, info in self.hidden_windows.items():
                label = f"{info['title'][:40]}  ({info['exe']})"

                def make_restore(h):
                    return lambda icon, item: self._restore_window(h)

                hidden_items.append(pystray.MenuItem(label, make_restore(hwnd)))

        if not hidden_items:
            hidden_items.append(pystray.MenuItem("(no hidden windows)", None, enabled=False))

        remove_items = []
        for app in self.config.get("apps", []):
            def make_remove(a):
                def remove(icon, item):
                    self.config["apps"] = [x for x in self.config["apps"] if x.lower() != a.lower()]
                    self._save_config()
                    self._update_menu()
                return remove
            remove_items.append(pystray.MenuItem(app, make_remove(app)))
        if not remove_items:
            remove_items.append(pystray.MenuItem("(none)", None, enabled=False))

        watching = [pystray.MenuItem(a, None, enabled=False)
                    for a in self.config.get("apps", [])]
        if not watching:
            watching.append(pystray.MenuItem("(none)", None, enabled=False))

        items = [
            pystray.MenuItem("Hidden Windows", pystray.Menu(*hidden_items)),
            pystray.MenuItem("Restore All", lambda icon, item: self._restore_all(), default=True),
        ]
        if self._launch_mode:
            # `cmd /K` outlives the program it ran, so without this the only
            # way to clear the icon is Task Manager.
            items.append(pystray.MenuItem(
                "Quit Program", lambda icon, item: self._kill_launched()))
        items += [
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Add App (type name)...", lambda icon, item: self._add_app_dialog()),
            pystray.MenuItem("Add App (pick running)...", lambda icon, item: self._pick_running_app_dialog()),
            pystray.MenuItem("Remove App", pystray.Menu(*remove_items)),
            pystray.MenuItem("Watching", pystray.Menu(*watching)),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit", lambda icon, item: self._exit()),
        ]
        return pystray.Menu(*items)

    def _update_menu(self):
        # Serialised: pystray's win32 backend destroys and recreates the
        # native HMENU here, and concurrent rebuilds from the detection,
        # monitor and watchdog threads would leave a dangling handle.
        with self._menu_lock:
            if self.icon:
                try:
                    self.icon.menu = self._build_menu()
                    self.icon.update_menu()
                except Exception as e:
                    _log(f"menu update failed: {e}")

    # ── Lifecycle ─────────────────────────────────────────────────────

    def _kill_launched(self):
        """Terminate the launched process tree, then exit.

        The tray icon stands for "my program, running hidden".  When the user
        says quit, the program should actually stop — leaving a stray `cmd /K`
        behind is what stacks up unkillable icons in the notification area.
        """
        for proc in self._launched_procs:
            if proc.poll() is not None:
                continue
            try:
                parent = psutil.Process(proc.pid)
            except psutil.NoSuchProcess:
                continue
            children = []
            try:
                children = parent.children(recursive=True)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
            for victim in children + [parent]:
                try:
                    _log(f"quit program: terminating pid={victim.pid} "
                         f"({victim.name()})")
                    victim.terminate()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            _, alive = psutil.wait_procs(children + [parent], timeout=3)
            for victim in alive:
                try:
                    victim.kill()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
        self._exit()

    def _restore_own_console(self):
        """Un-hide the console we inherited from our launcher.

        Only matters when the launcher outlives us — an interactive
        `cmd /K ... tray_minimizer.py ...`.  Then that shell is still sitting
        at a prompt behind a window we hid, and if we exit without showing it
        again it becomes invisible and unreachable forever.  When the launcher
        already exited (the documented `start "" /MIN` form) the hwnd is dead
        and ShowWindow is a harmless no-op.
        """
        hwnd = self._inherited_console_hwnd
        if not hwnd:
            return
        self._inherited_console_hwnd = None
        try:
            if win32gui.IsWindow(hwnd) and not win32gui.IsWindowVisible(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                _log(f"restored our launcher's console window hwnd={hwnd}")
        except Exception as e:
            _log(f"could not restore launcher console: {e}")

    def _exit(self):
        # Clear watched set first so _restore_all doesn't re-add them.
        with self.lock:
            self._watched_hwnds.clear()
        self._restore_all()
        self._restore_own_console()
        self.running = False
        for hook in self._hooks:
            if hook:
                user32.UnhookWinEvent(hook)
        if self.icon:
            self.icon.stop()

    def run(self, launch_cmd=None, tray_name=None):
        self._launch_mode = launch_cmd is not None

        if self._launch_mode:
            # In launch mode TrayMinimizer inherits the batch file's console
            # window.  That window is useless (the main thread blocks in
            # icon.run()), but it sits there confusingly — the user may try
            # to type 'exit' in it and wonder why nothing happens.
            #
            # Hide it and then fully detach.  This also eliminates the
            # FreeConsole/AttachConsole tug-of-war with the detection threads
            # (fixing the "random characters" bug).
            #
            # The hwnd is remembered so _restore_own_console() can give it back
            # on the way out.  It is not ours to keep: when the launcher is an
            # interactive shell (`cmd /K ... tray_minimizer.py ...` rather than
            # the documented `start "" /MIN ...`), that shell survives us, and
            # leaving its window hidden strands it as an invisible, immortal
            # prompt with no way to reach it.  Twenty-nine of those had piled
            # up before this was fixed.
            own_hwnd = kernel32.GetConsoleWindow()
            if own_hwnd and win32gui.IsWindowVisible(int(own_hwnd)):
                # Only remember a window we actually changed.  Under a ConPTY
                # (Windows Terminal, or a parent that is itself a pseudo
                # console) GetConsoleWindow() returns a hidden 0x0
                # PseudoConsoleWindow that was never on screen; "restoring"
                # that on exit would pop a bogus empty window into existence.
                self._inherited_console_hwnd = int(own_hwnd)
                win32gui.ShowWindow(int(own_hwnd), win32con.SW_HIDE)
            kernel32.FreeConsole()

        # Start the event hook thread (used by both modes)
        hook_thread = threading.Thread(target=self._hook_thread, daemon=True)
        hook_thread.start()

        # Start cleanup thread
        cleanup_thread = threading.Thread(target=self._watchdog_thread, daemon=True)
        cleanup_thread.start()

        # If launch mode, launch the program and hide it
        if launch_cmd:
            self.launch_and_hide(launch_cmd)

        # Pick the tray icon image: use the launched program's own icon
        # in launch mode, otherwise the default TrayMinimizer icon.
        if launch_cmd:
            exe_icon = _extract_exe_icon(launch_cmd[0])
            icon_image = exe_icon or self._create_icon_image()
            icon_title = tray_name or f"Tray Minimizer — {os.path.basename(launch_cmd[0])}"
        else:
            icon_image = self._create_icon_image()
            icon_title = tray_name or "Tray Minimizer"

        # Run tray icon on main thread (blocks)
        self.icon = pystray.Icon(
            "TrayMinimizer",
            icon_image,
            icon_title,
            self._build_menu(),
        )
        _log("starting icon.run()")
        try:
            self.icon.run()
        finally:
            # Covers every way out of the message loop, including ones that
            # never go through _exit().
            self._restore_own_console()
        _log("icon.run() returned")


LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tray_minimizer.log")

def _log(msg):
    try:
        with open(LOG_FILE, "a") as f:
            f.write(f"{time.strftime('%H:%M:%S')} {msg}\n")
    except Exception:
        pass

USAGE = """\
TrayMinimizer - hide application windows to the Windows system tray.

Usage:
  tray_minimizer.py                       Watch mode: auto-hide configured apps.
  tray_minimizer.py [--name NAME] PROG [ARGS...]
                                          Launch mode: run PROG hidden in the
                                          tray.  NAME sets the tray tooltip.

Options:
  -h, --help, /?   Show this help and exit.
  --name NAME      Tray tooltip / title for the launched program.

Examples:
  tray_minimizer.py                       (watch mode; configure via tray menu)
  tray_minimizer.py notepad.exe
  tray_minimizer.py --name "my bot" cmd /c python bot.py
"""

if __name__ == "__main__":
    _log(f"started, argv={sys.argv}")

    # Help must be handled BEFORE anything enters launch mode.  Otherwise
    # "--help" is treated as a program to launch, and launch mode hides and
    # detaches our own console — making the window vanish instead of printing.
    if any(a in ("-h", "--help", "/?", "/h") for a in sys.argv[1:2]):
        print(USAGE)
        sys.exit(0)

    try:
        # Parse optional --name flag
        args = sys.argv[1:]
        tray_name = None
        if len(args) >= 2 and args[0] == "--name":
            tray_name = args[1]
            args = args[2:]

        minimizer = TrayMinimizer()
        if args:
            _log(f"launch mode: {args}, name={tray_name}")
            minimizer.run(launch_cmd=args, tray_name=tray_name)
        else:
            _log("watch mode")
            print(
                "TrayMinimizer is running in watch mode (a tray icon is now "
                "active).\nRight-click the tray icon to add/remove watched "
                "apps, or choose Exit.\nThis console will stay open while it "
                "runs; press Ctrl+C here to quit.",
                flush=True,
            )
            minimizer.run()
    except KeyboardInterrupt:
        _log("interrupted by user (Ctrl+C)")
        try:
            minimizer._exit()
        except Exception:
            pass
    except Exception as e:
        _log(f"CRASHED: {e}")
        import traceback
        _log(traceback.format_exc())
        raise
    _log("exited normally")
