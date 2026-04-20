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

**Combining multiple commands** with `&`:

```batch
@start "" /MIN python "d:\path\to\tray_minimizer.py" cmd /K "d: & cd \myproject & call run.bat"
```

### Tray icon behavior

- The tray icon shows the launched program's own icon (e.g. cmd.exe's icon for batch files).
- **Double-click** the icon to restore the window.
- **Right-click** for a menu with Restore All and Exit.
- When the launched process exits, TrayMinimizer auto-exits.
- Choosing Exit restores all hidden windows before quitting.

## Watch Mode

Run TrayMinimizer with no arguments to start it as a background watcher.

```
python tray_minimizer.py
```

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

## Log File

`tray_minimizer.log` (next to the script) records startup, window detection, and process exit events for debugging.
