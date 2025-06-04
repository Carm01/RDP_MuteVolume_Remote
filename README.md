# Mute Volume on RDP machine upon Connect ( Windows Only )

This script, when run on the remote machine (the one you connect to via RDP), mutes the system volume upon connection and keeps it muted until you manually unmute it.

---

## How to Compile the Script into an EXE

You can compile this Python script into a standalone executable (`.exe`) using [PyInstaller](https://pyinstaller.org/).

### Steps to Compile

1. **Open Command Prompt (CMD)** in the directory where PyInstaller is installed, or make sure `pyinstaller` is available in your system PATH.

2. **Run one of the following commands:**

   - To compile **with a custom icon**:
     ```bash
     pyinstaller --onefile --noconsole --icon=icon.ico "C:\path\to\MuteDesktopOnRemote.py"
     ```

   - To compile **without a custom icon**:
     ```bash
     pyinstaller --onefile --noconsole "C:\path\to\MuteDesktopOnRemote.py"
     ```

### Notes

- `--onefile`: Bundles everything into a single executable  
- `--noconsole`: Hides the terminal window on launch  
- `--icon`: Optional, sets a custom `.ico` file for the app icon  

### Icon File Requirements

The `icon.ico` file must either be:
- In the same directory where you run the `pyinstaller` command, **or**
- Provided with a full file path:
  ```bash
  pyinstaller --onefile --noconsole --icon="C:\full\path\to\icon.ico" "C:\path\to\MuteDesktopOnRemote.py"
  ```

### Installing PyInstaller

If PyInstaller isn't installed yet, run:
```bash
pip install pyinstaller
```

To verify it's available:
```bash
pyinstaller --version
```

---

## System Tray Icon Options

The script uses a system tray icon with `pystray` and `Pillow`. You can choose a **custom icon** or a **built-in default icon**.

---

### ✅ Option 1: Use a Custom Icon (Default)

This is the default function included in the script:

```python
def create_icon():
    """Load your custom icon file."""
    return Image.open(r"icon.ico")
```

Update the path in `Image.open()` to reflect your actual icon location, if needed.

**Example:**
```python
return Image.open(r"C:\Users\YourName\Desktop\my_icon.ico")
```

---

### ✅ Option 2: Use a Built-in Default Icon

To use a simple default icon instead:

1. Comment out the custom icon loader.
2. Use this function instead:

```python
def create_icon():
    """Create a simple system tray icon."""
    image = Image.new('RGB', (16, 16), color='blue')
    pixels = image.load()
    for x in range(16):
        for y in range(16):
            pixels[x, y] = (0, 0, 255)
    return image
```
