import psutil
import subprocess
from openpyxl import load_workbook

from src.constants import BOOK_TRACKER_EXCEL_FILE_PATH

def check_if_excel_is_open():
    """Check if Excel is open and prompt the user to close it."""
    for process in psutil.process_iter(['name', 'pid']):
        if process.info['name'] and 'EXCEL' in process.info['name'].upper():
            # Optionally, print out the process ID for debugging
            print(f"Excel process detected: {process.info['name']} (PID: {process.info['pid']})")
            print_red("Excel is currently open. Please save and close Excel to proceed.")
            return True
    return False

def updateExcelCellAndOpenExcel(danaCode, folderName):
    if danaCode:
        # Check if Excel is open, wait until it's closed
        while check_if_excel_is_open():
            input("Press Enter after closing Excel...")  # Wait for user to close Excel

        # Load the workbook and select the active worksheet
        workbook = load_workbook(BOOK_TRACKER_EXCEL_FILE_PATH)
        sheet = workbook.active

        # Iterate through column B to find the row with the value danaCode
        for row in sheet.iter_rows(min_col=2, max_col=2):  # Only column B
            cell = row[0]  # Get the first (and only) cell in the column
            if cell.value == danaCode:
                # Write folderName in the same row, column F (column 6)
                sheet.cell(row=cell.row, column=6).value = folderName
                print_green(f"Successfully wrote {folderName} in the row of {danaCode}\n")
                break

        workbook.save(BOOK_TRACKER_EXCEL_FILE_PATH)  # Save the changes to the Excel file

        subprocess.run(['start', 'excel', BOOK_TRACKER_EXCEL_FILE_PATH], shell=True)  # Open the Excel file