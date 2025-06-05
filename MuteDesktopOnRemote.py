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
import io
import base64

# Define threading events
stop_thread = Event()
startup_enabled = Event()

# === EMBEDDED ICON (base64) ===
embedded_icon_base64 = """
<YOUR_BASE64_ICON_STRING_HERE>
"""  # Replace this with your actual base64-encoded .ico content


def is_rdp_active():
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
    stop_thread.set()
    unmute_volume()
    icon.stop()


def on_unmute(icon, item):
    unmute_volume()


def create_icon():
    try:
        icon_bytes = base64.b64decode(embedded_icon_base64)
        return Image.open(io.BytesIO(icon_bytes))
    except Exception as e:
        print(f"Failed to load embedded icon: {e}")
        return create_generic_icon()


def create_generic_icon():
    image = Image.new('RGB', (16, 16), color='blue')
    pixels = image.load()
    for x in range(16):
        pixels[x, 0] = (0, 0, 0)
        pixels[x, 15] = (0, 0, 0)
        pixels[0, x] = (0, 0, 0)
        pixels[15, x] = (0, 0, 0)
    return image


def is_in_startup():
    startup_path = winshell.startup()
    shortcut_path = os.path.join(startup_path, "RDPVolumeControl.lnk")
    return os.path.exists(shortcut_path)


def toggle_startup(icon, item):
    startup_path = winshell.startup()
    shortcut_path = os.path.join(startup_path, "RDPVolumeControl.lnk")

    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)
        print("Removed from startup.")
        startup_enabled.clear()
        stop_thread.set()
    else:
        try:
            if getattr(sys, 'frozen', False):
                target = sys.executable
                arguments = ''
            else:
                target = sys.executable
                script_path = os.path.abspath(__file__)
                arguments = f'"{script_path}"'

            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.TargetPath = target
            shortcut.Arguments = arguments
            shortcut.WorkingDirectory = os.path.dirname(target)
            shortcut.IconLocation = target
            shortcut.save()
            print("Added to startup.")
            startup_enabled.set()
            start_monitor_thread(icon)
        except Exception as e:
            print(f"Error adding to startup: {e}")


def start_monitor_thread(icon):
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

    if is_in_startup():
        startup_enabled.set()
        start_monitor_thread(icon)

    icon.run()


if __name__ == "__main__":
    main()
