import os
import subprocess
import time
import win32gui
import win32con
import win32api
from pathlib import Path

from src.logic.file_operations import validate_pdf_path


def position_window(window_handle, x, y, width, height):
    """Brings window to front and positions it."""
    try:
        # If window is minimized, restore it
        if win32gui.IsIconic(window_handle):
            win32gui.ShowWindow(window_handle, win32con.SW_RESTORE)
        else:
            win32gui.ShowWindow(window_handle, win32con.SW_NORMAL)

        win32gui.SetForegroundWindow(window_handle)
        win32gui.MoveWindow(window_handle, x, y, width, height, True)
    except Exception as e:
        print(f"Error positioning window: {e}")


def open_pdfs_side_by_side_acrobat(con_pdf_path, fin_pdf_path):
    # 1. Get and validate paths using your existing project function
    valid1, err1 = validate_pdf_path(con_pdf_path)
    valid2, err2 = validate_pdf_path(fin_pdf_path)

    if not valid1 or not valid2:
        print(f"Error: {err1 if not valid1 else err2}")
        return

    # Create Path objects locally to access .name without changing validate_pdf_path
    pdf1_obj = Path(con_pdf_path.strip('"'))
    pdf2_obj = Path(fin_pdf_path.strip('"'))

    # 2. Locate Adobe Acrobat
    acrobat_paths = [
        r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
    ]

    acrobat_exe = None
    for p in acrobat_paths:
        if os.path.exists(p):
            acrobat_exe = p
            break

    if not acrobat_exe:
        print("Error: Adobe Acrobat executable not found.")
        return

    print(f"Opening {pdf1_obj.name} and {pdf2_obj.name} in Adobe Acrobat...")

    # 3. Launch instances
    # The '/n' flag forces a new instance of Acrobat
    subprocess.Popen([acrobat_exe, "/n", str(pdf1_obj)])
    subprocess.Popen([acrobat_exe, "/n", str(pdf2_obj)])

    # 4. Wait for windows to initialize
    time.sleep(5)

    # 5. Calculate screen dimensions
    monitor_info = win32api.GetMonitorInfo(win32api.MonitorFromPoint((0, 0)))
    work_area = monitor_info['Work']
    screen_h = work_area[3] - work_area[1]
    half_w = (work_area[2] - work_area[0]) // 2

    # 6. Find windows by filename in title
    def find_window_by_filename(filename):
        hwnds = []

        def callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if filename.lower() in title.lower():
                    hwnds.append(hwnd)

        win32gui.EnumWindows(callback, None)
        return hwnds

    hwnds_left = find_window_by_filename(pdf1_obj.name)
    hwnds_right = find_window_by_filename(pdf2_obj.name)

    # 7. Apply positioning
    if hwnds_left:
        position_window(hwnds_left[0], work_area[0], work_area[1], half_w, screen_h)
    else:
        print(f"Could not find window for {pdf1_obj.name}")

    if hwnds_right:
        # If both files have the same name, use the second handle
        target_hwnd = hwnds_right[0]
        if pdf1_obj.name == pdf2_obj.name and len(hwnds_right) > 1:
            target_hwnd = hwnds_right[1]

        position_window(target_hwnd, work_area[0] + half_w, work_area[1], half_w, screen_h)
    else:
        print(f"Could not find window for {pdf2_obj.name}")


if __name__ == "__main__":
    # Prompting for paths since the function now requires them as arguments
    p1 = input("Enter path for the LEFT PDF: ").strip()
    p2 = input("Enter path for the RIGHT PDF: ").strip()
    open_pdfs_side_by_side_acrobat(p1, p2)