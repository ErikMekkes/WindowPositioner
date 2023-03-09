import win32gui
import win32con
import ctypes
import ctypes.wintypes
import sqlite3
from sqlite3 import Error
import asyncio
from pynput import keyboard

"""
    Class to define a window and its relevant attributes.
        title: Printable string
        position: Array of [x1,y1,x2,y2]
        handle: Current active window handle if available, might be None.
"""
class WindowPosition:
    def __init__(self, title, handle, window_placement, window_rectangle):
        self.title = title
        self.handle = handle
        self.window_placement = window_placement
        self.window_rectangle = window_rectangle

"""
    Class used to allocate space for reading window attributes from wingui.
"""
class TitleBarInfo(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.wintypes.DWORD),
        ("rcTitleBar", ctypes.wintypes.RECT),
        ("rgstate", ctypes.wintypes.DWORD * 6)
    ]

"""
    Basic sqlite database connection to store window positions.
"""
def create_connection(path):
    connection = None
    try:
        connection = sqlite3.connect(path)
        print("Connection to SQLite DB successful")
    except Error as e:
        print(f"The error '{e}' occurred")

    return connection

"""
    Basic keyboard event handler.
"""
def on_press(key):
    if key == keyboard.Key.esc:
        global continue_running
        continue_running = False
        return False  # stop listener thread
    try:
        k = key.char  # single-char keys
    except:
        k = key.name  # other keys
    if k == "ß":  # Ctrl+Alt+S
        snapshot()
    if k == "®": # Ctrl+Alt+R
        restore()

"""
    Takes a snapshot of all currently open window positions and stores it.
"""
def snapshot():
    global window_positions
    window_positions = []
    win32gui.EnumWindows(FilterWindows, window_positions)
    print("\nSnapshot made: \n")
    for wPos in window_positions:
        print(f'Window Recorded: {wPos.title}\t{wPos.window_rectangle}\t{wPos.window_placement}')

"""
    Restores open windows to their previous positions from the last snapshot.
"""
def restore():
    global window_positions
    openWindows = []
    print("\nRestoring Snapshot: \n")
    win32gui.EnumWindows(FilterWindows, openWindows)
    for wPos in openWindows:
        # find if a previously known position exists
        for knownPos in window_positions:
            if knownPos.title == wPos.title:
                restoreWindow(wPos.handle, knownPos.window_placement, knownPos.window_rectangle)
                #restoreWindow(wPos.handle, knownPos.position, knownPos.flags)
                print(f'Window Moved: {wPos.title}\t{wPos.position}')
                break
def restoreWindow(handle, window_placement, window_rectangle):
    new_placement = (
        window_placement[0],
        window_placement[1],
        window_placement[2],
        window_placement[3],
        window_rectangle
    )
    win32gui.SetWindowPlacement(handle, new_placement)

"""
    Filters out background windows that are not relevant.
"""
def FilterWindows(handle, window_list):
    # Title Info Initialization
    title_info = TitleBarInfo()
    title_info.cbSize = ctypes.sizeof(title_info)
    ctypes.windll.user32.GetTitleBarInfo(handle, ctypes.byref(title_info))
    title = win32gui.GetWindowText(handle)
    # some titles might contain non-standard encoding
    titlestring = title.encode('utf8')
    window_placement = win32gui.GetWindowPlacement(handle)
    window_rectangle = win32gui.GetWindowRect(handle)
    
    # DWM Cloaked Check
    is_cloaked = ctypes.c_int(0)
    ctypes.WinDLL("dwmapi").DwmGetWindowAttribute(handle, 14, ctypes.byref(is_cloaked), ctypes.sizeof(is_cloaked))
    if is_cloaked == 1 or title == "" or not win32gui.IsWindowVisible(handle):
        return
    if (title_info.rgstate[0] & win32con.STATE_SYSTEM_INVISIBLE):
        return
    window_list.append(
        WindowPosition(titlestring, handle, window_placement, window_rectangle)
    )

db_conn          = None      # global sqlite connection
continue_running = True      # global main loop condition
window_positions = []        # global last snapshot of window positions
"""
    Main program loop
    Using a non-blocking asynchronous main loop to minimize resources used
"""
async def main():
    # connect to local database for window positions
    global db_conn
    db_conn = create_connection("WindowPositions.sqlite")

    # start keyboard listener on separate thread
    listener = keyboard.Listener(on_press=on_press)
    listener.start()
    print(
        f"\nWindowPositioner:\n"
        f"\n - Press [Esc] to exit program\n"
        f"\n - Press [Ctrl+Alt+S] to record current window positions\n"
        f" - Press [Ctrl+Alt+R] to restore last known window positions\n"
    )

    # main loop to keep program running, non blocking sleep.
    while(continue_running):
        await asyncio.sleep(1)

asyncio.run(main())