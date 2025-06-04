import win32ts
import time
import pystray
from PIL import Image
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
import sys
from threading import Thread, Event
import pythoncom
import os
import winshell
from win32com.client import Dispatch

# Define threading events
stop_thread = Event()
startup_enabled = Event()


def is_rdp_active():
    """Check if an RDP session is active."""
    try:
        sessions = win32ts.WTSEnumerateSessions(win32ts.WTS_CURRENT_SERVER_HANDLE)
        for session in sessions:
            if session['State'] == win32ts.WTSActive:
                client_name = win32ts.WTSQuerySessionInformation(
                    win32ts.WTS_CURRENT_SERVER_HANDLE,
                    session['SessionId'],
                    win32ts.WTSClientName
                )
                if client_name:
                    return True
        return False
    except Exception as e:
        print(f"Error checking RDP: {e}")
        return False


def mute_volume():
    """Mute the master volume."""
    pythoncom.CoInitialize()
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = interface.QueryInterface(IAudioEndpointVolume)
        volume.SetMute(1, None)
        print("Master volume muted")
    except Exception as e:
        print(f"Error muting volume: {e}")
    finally:
        pythoncom.CoUninitialize()


def unmute_volume():
    """Unmute the master volume."""
    pythoncom.CoInitialize()
    try:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = interface.QueryInterface(IAudioEndpointVolume)
        volume.SetMute(0, None)
        print("Master volume unmuted")
    except Exception as e:
        print(f"Error unmuting volume: {e}")
    finally:
        pythoncom.CoUninitialize()


def monitor_rdp(icon):
    """Monitor RDP connections and control volume."""
    was_rdp_active = False
    while not stop_thread.is_set():
        try:
            rdp_active = is_rdp_active()
            if rdp_active and not was_rdp_active:
                print("RDP detected, muting volume.")
                mute_volume()
            elif not rdp_active and was_rdp_active:
                print("RDP disconnected, unmuting volume.")
                unmute_volume()
            was_rdp_active = rdp_active
        except Exception as e:
            print(f"Monitor error: {e}")
        time.sleep(5)


def on_exit(icon, item):
    """Handle exit menu item."""
    stop_thread.set()
    unmute_volume()
    icon.stop()


def on_unmute(icon, item):
    """Unmute immediately from tray."""
    unmute_volume()


def create_icon():
    """Create system tray icon - either custom or fallback."""
    custom_icon_path = r"P:\Apps\Python\Icons\Google-Noto-Emoji-Objects-62790-speaker-high-volume.ico"
    if os.path.exists(custom_icon_path):
        return Image.open(custom_icon_path)
    else:
        return create_generic_icon()


def create_generic_icon():
    """Fallback generic icon."""
    image = Image.new('RGB', (16, 16), color='blue')
    pixels = image.load()
    for x in range(16):
        pixels[x, 0] = (0, 0, 0)
        pixels[x, 15] = (0, 0, 0)
        pixels[0, x] = (0, 0, 0)
        pixels[15, x] = (0, 0, 0)
    return image


def is_in_startup():
    """Check if script is in Startup folder."""
    startup_path = winshell.startup()
    shortcut_path = os.path.join(startup_path, "RDPVolumeControl.lnk")
    return os.path.exists(shortcut_path)


def toggle_startup(icon, item):
    """Toggle script autostart."""
    startup_path = winshell.startup()
    shortcut_path = os.path.join(startup_path, "RDPVolumeControl.lnk")

    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)
        print("Removed from startup.")
        startup_enabled.clear()
        stop_thread.set()
    else:
        try:
            target = sys.executable
            script_path = os.path.abspath(__file__)
            arguments = f'"{script_path}"'
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.TargetPath = target
            shortcut.Arguments = arguments
            shortcut.WorkingDirectory = os.path.dirname(script_path)
            shortcut.IconLocation = script_path
            shortcut.save()
            print("Added to startup.")
            startup_enabled.set()
            start_monitor_thread(icon)
        except Exception as e:
            print(f"Error adding to startup: {e}")


def start_monitor_thread(icon):
    """Start the RDP monitoring thread."""
    if stop_thread.is_set():
        stop_thread.clear()
    monitor_thread = Thread(target=monitor_rdp, args=(icon,))
    monitor_thread.daemon = True
    monitor_thread.start()


def main():
    global stop_thread
    stop_thread.clear()
    icon = pystray.Icon("RDP Volume Control")
    icon.icon = create_icon()
    icon.menu = pystray.Menu(
        pystray.MenuItem("Unmute Now", on_unmute),
        pystray.MenuItem(
            "Start with Windows",
            toggle_startup,
            checked=lambda item: is_in_startup()
        ),
        pystray.MenuItem("Exit", on_exit)
    )

    # Only monitor if already in startup
    if is_in_startup():
        startup_enabled.set()
        start_monitor_thread(icon)

    icon.run()


if __name__ == "__main__":
    main()
