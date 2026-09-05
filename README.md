# TrayMinimizer

Hide windows to the system tray. Two modes: **launch mode** (hide a specific program) and **watch mode** (automatically hide programs by exe name).

## Requirements

- Python 3.10+
- `pip install pystray Pillow pywin32 psutil`

## Launch Mode

Run TrayMinimizer with a command and it launches that command in a hidden console, showing the program's icon in the system tray.

```
python tray_minimizer.py <program> [args...]
```

**Examples:**

```batch
python tray_minimizer.py cmd /K "cd \myproject & python server.py"
python tray_minimizer.py cmd /K "supybot d:\bots\bot.conf"
python tray_minimizer.py notepad.exe
```

**Using from a batch file** (e.g. in the Windows Startup folder):

```batch
@start "" /MIN python "d:\path\to\tray_minimizer.py" cmd /K "your commands here"
```

`start /MIN` lets the original console exit immediately. TrayMinimizer hides its own console and runs silently in the background.

**Prefer this `start` form over `cmd /K ... tray_minimizer.py ...`.** TrayMinimizer hides the console it inherits from its launcher; with `start` that launcher exits straight away and there is nothing to hide. With `cmd /K` the launching shell survives, so TrayMinimizer restores its window when it exits — but until then you have a shell you cannot see. (Before this was handled, those shells were never given their window back and piled up invisibly, one per launch.)

**Combining multiple commands** with `&`:

```batch
@start "" /MIN python "d:\path\to\tray_minimizer.py" cmd /K "d: & cd \myproject & call run.bat"
```

### Tray icon behavior

- The tray icon shows the launched program's own icon (e.g. cmd.exe's icon for batch files).
- **Left-click** (or double-click) the icon to restore the window.
- **Right-click** for a menu with Restore All, Quit Program and Exit.
- When the launched process exits, TrayMinimizer auto-exits.
- Choosing **Exit** restores all hidden windows before quitting, leaving the launched program running.
- Choosing **Quit Program** terminates the launched program (and its children) first, then quits. Use this for a `cmd /K` session that has outlived the program it ran — `/K` keeps the console alive at a prompt forever, so without this the icon would have nothing to auto-exit on.
- Restoring a window forcibly brings it to the foreground (it won't get stranded behind a fullscreen browser). If the window's process has already exited, the dead tray entry is pruned instead.

### No dead tray icons

The icon exists to restore a window, so TrayMinimizer works to guarantee it always has a real one:

- **Only genuine top-level windows are ever hidden.** A window is rejected unless it is unowned, enabled, non-child, not a tool window, has a non-zero size, and is not one of the known helper classes. This matters more than it sounds: every process attached to a console gets a 0x0, disabled `Default IME` helper window owned by `conhost.exe`, and it is the only window in existence for the first few milliseconds after launch. Hiding *that* by mistake produces an icon that restores nothing.
- **The console probe gets a head start.** `AttachConsole`/`GetConsoleWindow` names a process's console window exactly and cannot pick the wrong one, so it runs alone for the first couple of seconds before the by-PID and window-diff fallbacks are allowed to guess.
- **A watchdog re-checks every 10 seconds.** Windows that no longer exist are pruned. If the tracked window is destroyed while the launched program keeps running (a `cmd /K` whose inner program exited, a GUI child that was closed), TrayMinimizer re-runs detection and adopts whatever window the program has now. If it has none at all, TrayMinimizer exits rather than leave an icon that does nothing when clicked.

### Keeping a window visible (`[nohide]` marker)

A launched program can keep a specific window on screen while the rest of it stays hidden in the tray. Any top-level window whose **title starts with `[nohide]`** is never auto-hidden — TrayMinimizer skips it in launch-mode detection and in the watch-mode hook alike.

Set it from the program, e.g. on Windows before/while showing the window:

```python
import ctypes
ctypes.windll.kernel32.SetConsoleTitleW("[nohide] my transient picker")
```

(orchestrator2 uses this so its `--resume`/`--copy` session picker pops a visible console while the server itself remains hidden in the tray.)

## Help

```
python tray_minimizer.py --help
```

Prints usage and exits. (`-h`, `/?`, and `/h` also work.) Help is handled before launch mode, so it never hides or detaches the console.

## Watch Mode

Run TrayMinimizer with no arguments to start it as a background watcher.

```
python tray_minimizer.py
```

Watch mode keeps its console open and prints a status line so it's clear it's running (not hung). Press **Ctrl+C** in that console to quit, or use the tray menu's Exit.

Right-click the tray icon to:
- **Add App (type name)** — enter an exe name like `notepad.exe`
- **Add App (pick running)** — choose from currently running programs
- **Remove App** — stop watching an app

When a watched app's window appears, TrayMinimizer automatically hides it to the tray. The watched app list is saved in `tray_minimizer.json` next to the script.

## Config File

`tray_minimizer.json` stores the watch list for watch mode:

```json
{
  "apps": ["notepad.exe", "some_app.exe"]
}
```

This file is not used in launch mode.

## Tests

`test_trayminimizer.py` is a live regression suite — it launches real TrayMinimizer processes, clicks their tray icons the way the shell does, and inspects the real window tree. It runs a scratch copy of the script from a temp directory, so it never touches your real `tray_minimizer.log` or `tray_minimizer.json`.

```
python test_trayminimizer.py                    # everything (~3 min)
python test_trayminimizer.py console gui help   # named tests only
```

Tests: `stress` (20 rapid launches must all hide the console, never a helper window), `console`, `gui`, `readopt`, `giveup`, `watch`, `launcher`, `help`.

## Log File

`tray_minimizer.log` (next to the script) records startup, window detection, and process exit events for debugging.

Detection lines name the exact window that was chosen, so a misbehaving tray icon can be diagnosed from the log alone:

```
detect: console probe matched hwnd=4073044 cls='ConsoleWindowClass' title='C:\WINDOWS\SYSTEM32\cmd.exe' pid=9380 (via pid=9380)
hiding hwnd=4073044 cls='ConsoleWindowClass' title='C:\WINDOWS\SYSTEM32\cmd.exe' pid=9380 exe=cmd launch_pid=9380
```

`skip hide:` lines record a window that was rejected and why; `watchdog:` lines record pruning, re-adoption and give-up decisions.
