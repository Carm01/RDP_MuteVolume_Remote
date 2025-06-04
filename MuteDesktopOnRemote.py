import win32ts
import time
import pystray
from PIL import Image
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
import sys
from threading import Thread
import pythoncom


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
                print(f"Session ID: {session['SessionId']}, Client Name: {client_name}")
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
    while not stop_thread:
        try:
            rdp_active = is_rdp_active()
            print(f"RDP Active: {rdp_active}, Was Active: {was_rdp_active}")
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
    """Handle exit menu item and ensure volume is unmuted."""
    global stop_thread
    stop_thread = True
    unmute_volume()  # Unmute before exiting
    icon.stop()


##def create_icon():
##    """Create a simple system tray icon."""
##    image = Image.new('RGB', (16, 16), color='blue')
##    pixels = image.load()
##    for x in range(16):
##        pixels[x, 0] = (0, 0, 0)
##        pixels[x, 15] = (0, 0, 0)
##        pixels[0, x] = (0, 0, 0)
##        pixels[15, x] = (0, 0, 0)
##    return image

def create_icon():
    """Load your custom icon file."""
    return Image.open(r"icon.ico")  # <-- Change this path to your icon file path, or un comment the def create_icon():  section above and comment this section for generic one


def main():
    global stop_thread
    stop_thread = False
    icon = pystray.Icon("RDP Volume Control")
    icon.icon = create_icon()
    icon.menu = pystray.Menu(
        pystray.MenuItem("Unmute Now", on_unmute),
        pystray.MenuItem("Exit", on_exit)
    )
    monitor_thread = Thread(target=monitor_rdp, args=(icon,))
    monitor_thread.daemon = True
    monitor_thread.start()
    icon.run()
    
def on_unmute(icon, item):
    """Handle unmute menu item."""
    unmute_volume()

if __name__ == "__main__":
    main()
