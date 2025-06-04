# Mute Volume on RDP Machine upon Connect (Windows Only)

This script, when run on the remote machine (the one you connect to via RDP), mutes the system volume upon connection and keeps it muted until you manually unmute it.

---

## Prerequisites

Install the required Python packages using the following commands:

```bash
pip install pywin32 pystray Pillow pycaw comtypes
pip install pywin32
pip install pystray pycaw pywin32 pillow winshell

```

---

## How to Compile the Script into an EXE

You can compile this Python script into a standalone executable (.exe) using [PyInstaller](https://pyinstaller.org/).

### Steps to Compile

1. **Open Command Prompt (CMD)** in the directory where PyInstaller is installed, or ensure PyInstaller is available in your system PATH.

2. **Run one of the following commands:**

   - To compile **with a custom icon**:
     
     ```bash
     pyinstaller --onefile --noconsole --icon=mute_icon.ico "C:\path\to\MuteDesktopOnRemote.py"
     ```

   - To compile **without a custom icon**:
     
     ```bash
     pyinstaller --onefile --noconsole "C:\path\to\MuteDesktopOnRemote.py"
     ```

### Notes

- `--onefile`: Bundles everything into a single executable.
- `--noconsole`: Hides the terminal window on launch.
- `--icon`: Optional, sets a custom .ico file for the app icon.

### Icon File Requirements

The `mute_icon.ico` file must be:
- In the same directory as the script when running the PyInstaller command, **or**
- Specified with a full file path, e.g.:

  ```bash
  pyinstaller --onefile --noconsole --icon="C:\full\path\to\mute_icon.ico" "C:\path\to\MuteDesktopOnRemote.py"
  ```
### Startup Location

The start location where the shortcut will be located:

  ```bash
  %APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
  ```
