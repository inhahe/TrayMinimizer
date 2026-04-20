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
WINEVENT_OUTOFCONTEXT = 0x0000
WINEVENT_SKIPOWNPROCESS = 0x0002
OBJID_WINDOW = 0

# Process creation flags
CREATE_NEW_CONSOLE = 0x00000010

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


def _find_window_by_pid(pid, *, require_visible=True):
    """Find a titled, top-level window owned by the given PID.

    When *require_visible* is False the window may be hidden (SW_HIDE),
    which is needed when we launch child processes born-hidden.
    """
    result = []
    def cb(hwnd, _):
        if require_visible and not win32gui.IsWindowVisible(hwnd):
            return
        if not win32gui.GetWindowText(hwnd):
            return
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        if ex_style & win32con.WS_EX_TOOLWINDOW:
            return
        _, wpid = win32process.GetWindowThreadProcessId(hwnd)
        if wpid == pid:
            result.append(hwnd)
    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass
    return result[0] if result else None


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


class TrayMinimizer:
    def __init__(self):
        self.config = self._load_config()
        self.hidden_windows = {}  # hwnd -> {"title", "exe", "pid", "process"}
        self.lock = threading.Lock()
        self.running = True
        self.icon = None
        self.known_hwnds = set()
        self._hook_proc = WinEventProcType(self._win_event_callback)
        self._hooks = []
        self._launched_procs = []  # Popen objects from launch mode

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
        if not win32gui.IsWindowVisible(hwnd):
            return False
        if not win32gui.GetWindowText(hwnd):
            return False
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        if not (style & win32con.WS_VISIBLE):
            return False
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        if ex_style & win32con.WS_EX_TOOLWINDOW:
            return False
        return True

    def _hide_window(self, hwnd, exe_override=None, pid_override=None):
        if not win32gui.IsWindow(hwnd):
            return
        title = win32gui.GetWindowText(hwnd)
        if exe_override:
            exe, pid = exe_override, pid_override
        else:
            exe, pid = self._get_exe_for_hwnd(hwnd)
        if not exe:
            return
        # Window may already be hidden (started with SW_HIDE); calling
        # ShowWindow(SW_HIDE) on an already-hidden window is a harmless no-op.
        win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
        with self.lock:
            self.hidden_windows[hwnd] = {
                "title": title or exe, "exe": exe, "pid": pid,
            }
        self._update_menu()

    def _restore_window(self, hwnd):
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass
        with self.lock:
            self.hidden_windows.pop(hwnd, None)
            self.known_hwnds.discard(hwnd)
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

        _log(f"launched pid={proc.pid} exe={exe_name}")

        # Start a thread to find and hide the window
        threading.Thread(
            target=self._find_and_hide_launched,
            args=(proc, exe_name, before),
            daemon=True,
        ).start()

        # Monitor the process — clean up tray entry and optionally exit when done
        threading.Thread(
            target=self._monitor_process,
            args=(proc,),
            daemon=True,
        ).start()

    def _find_and_hide_launched(self, proc, exe_name, windows_before, timeout=15):
        """Find the window created by a launched process and hide it.

        The child was started with SW_HIDE, so visibility checks are skipped.
        Strategies are ordered by reliability for console apps on Win 10/11:

        1. AttachConsole — works regardless of which process owns the actual
           console window (conhost.exe, OpenConsole.exe, etc.).
        2. GUI window search by PID — catches non-console (GUI) children.
        3. Window-list diff — fallback that also handles conhost-owned windows
           by verifying ownership through AttachConsole.
        """
        deadline = time.time() + timeout

        while time.time() < deadline:
            if proc.poll() is not None:
                return  # process already exited

            # Refresh the process tree (children may spawn after launch)
            tree_pids = _get_process_tree_pids(proc.pid)

            # Strategy 1: AttachConsole — most reliable for console apps.
            # GetConsoleWindow() returns the HWND regardless of visibility,
            # and works even when the window is owned by conhost.exe.
            for pid in tree_pids:
                hwnd = _find_console_hwnd(pid)
                if hwnd and hwnd not in self.known_hwnds and win32gui.IsWindow(hwnd):
                    _log(f"found via AttachConsole: hwnd={hwnd} pid={pid}")
                    self.known_hwnds.add(hwnd)
                    self._hide_window(hwnd, exe_override=exe_name, pid_override=proc.pid)
                    return

            # Strategy 2: find a GUI window owned by any PID in the tree.
            # Skip the IsWindowVisible filter — the window may be born hidden.
            for pid in tree_pids:
                hwnd = _find_window_by_pid(pid, require_visible=False)
                if hwnd and hwnd not in self.known_hwnds:
                    self.known_hwnds.add(hwnd)
                    self._hide_window(hwnd, exe_override=exe_name, pid_override=proc.pid)
                    return

            # Strategy 3: diff the window list.  New windows may be owned by
            # conhost.exe (not in tree_pids), so we verify via AttachConsole.
            after = _snapshot_windows()
            new_hwnds = after - windows_before
            for hwnd in new_hwnds:
                if hwnd in self.known_hwnds:
                    continue
                if not win32gui.IsWindow(hwnd):
                    continue
                # Don't require visibility — child was started hidden.
                _, wpid = win32process.GetWindowThreadProcessId(hwnd)

                # Direct PID match
                if wpid in tree_pids:
                    self.known_hwnds.add(hwnd)
                    self._hide_window(hwnd, exe_override=exe_name, pid_override=proc.pid)
                    return

                # The window might belong to conhost.exe / OpenConsole.exe
                # hosting our child's console.  Verify by checking whether
                # AttachConsole on our child PID yields this same HWND.
                try:
                    owner = psutil.Process(wpid)
                    if owner.name().lower() in ("conhost.exe", "openconsole.exe"):
                        verify_hwnd = _find_console_hwnd(proc.pid)
                        if verify_hwnd == hwnd:
                            self.known_hwnds.add(hwnd)
                            self._hide_window(hwnd, exe_override=exe_name, pid_override=proc.pid)
                            return
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass

            time.sleep(0.2)

        _log(f"WARNING: could not find window for {exe_name} within {timeout}s")

    def _monitor_process(self, proc):
        """Wait for a launched process to exit, then clean up its tray entry."""
        proc.wait()
        _log(f"process {proc.pid} exited, code={proc.returncode}")
        # Give a moment for any final window cleanup
        time.sleep(0.5)

        with self.lock:
            stale = [h for h, info in self.hidden_windows.items()
                     if info.get("pid") == proc.pid or not win32gui.IsWindow(h)]
            for h in stale:
                del self.hidden_windows[h]
                self.known_hwnds.discard(h)
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
        if hwnd in self.known_hwnds:
            return
        if not self._is_app_window(hwnd):
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
        self._hooks = [hook1, hook2]

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

    # ── Stale window cleanup ──────────────────────────────────────────

    def _cleanup_thread(self):
        while self.running:
            time.sleep(10)
            with self.lock:
                stale = [h for h in self.hidden_windows if not win32gui.IsWindow(h)]
                for h in stale:
                    del self.hidden_windows[h]
                    self.known_hwnds.discard(h)
            if stale:
                self._update_menu()

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

        return pystray.Menu(
            pystray.MenuItem("Hidden Windows", pystray.Menu(*hidden_items)),
            pystray.MenuItem("Restore All", lambda icon, item: self._restore_all(), default=True),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Add App (type name)...", lambda icon, item: self._add_app_dialog()),
            pystray.MenuItem("Add App (pick running)...", lambda icon, item: self._pick_running_app_dialog()),
            pystray.MenuItem("Remove App", pystray.Menu(*remove_items)),
            pystray.MenuItem("Watching", pystray.Menu(*watching)),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Exit", lambda icon, item: self._exit()),
        )

    def _update_menu(self):
        if self.icon:
            self.icon.menu = self._build_menu()
            self.icon.update_menu()

    # ── Lifecycle ─────────────────────────────────────────────────────

    def _exit(self):
        self._restore_all()
        self.running = False
        for hook in self._hooks:
            if hook:
                user32.UnhookWinEvent(hook)
        if self.icon:
            self.icon.stop()

    def run(self, launch_cmd=None):
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
            own_hwnd = kernel32.GetConsoleWindow()
            if own_hwnd:
                win32gui.ShowWindow(int(own_hwnd), win32con.SW_HIDE)
            kernel32.FreeConsole()

        # Start the event hook thread (used by both modes)
        hook_thread = threading.Thread(target=self._hook_thread, daemon=True)
        hook_thread.start()

        # Start cleanup thread
        cleanup_thread = threading.Thread(target=self._cleanup_thread, daemon=True)
        cleanup_thread.start()

        # If launch mode, launch the program and hide it
        if launch_cmd:
            self.launch_and_hide(launch_cmd)

        # Pick the tray icon image: use the launched program's own icon
        # in launch mode, otherwise the default TrayMinimizer icon.
        if launch_cmd:
            exe_icon = _extract_exe_icon(launch_cmd[0])
            icon_image = exe_icon or self._create_icon_image()
            icon_title = f"Tray Minimizer — {os.path.basename(launch_cmd[0])}"
        else:
            icon_image = self._create_icon_image()
            icon_title = "Tray Minimizer"

        # Run tray icon on main thread (blocks)
        self.icon = pystray.Icon(
            "TrayMinimizer",
            icon_image,
            icon_title,
            self._build_menu(),
        )
        _log("starting icon.run()")
        self.icon.run()
        _log("icon.run() returned")


LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tray_minimizer.log")

def _log(msg):
    try:
        with open(LOG_FILE, "a") as f:
            f.write(f"{time.strftime('%H:%M:%S')} {msg}\n")
    except Exception:
        pass

if __name__ == "__main__":
    _log(f"started, argv={sys.argv}")
    try:
        minimizer = TrayMinimizer()
        if len(sys.argv) > 1:
            _log(f"launch mode: {sys.argv[1:]}")
            minimizer.run(launch_cmd=sys.argv[1:])
        else:
            _log("watch mode")
            minimizer.run()
    except Exception as e:
        _log(f"CRASHED: {e}")
        import traceback
        _log(traceback.format_exc())
        raise
    _log("exited normally")
