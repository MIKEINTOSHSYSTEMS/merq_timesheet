# setup_cxfreeze.py
import sys
from cx_Freeze import setup, Executable
import os

# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {
    "packages": ["os", "tkinter", "pandas", "openpyxl", "PIL", "requests", "json", "threading", "datetime"],
    "include_files": [
        ("src/merq.png", "merq.png"),
        ("src/MERQ_TIMESHEET_ETH-CAL_TEMPLATE.xlsx", "MERQ_TIMESHEET_ETH-CAL_TEMPLATE.xlsx")
    ],
    "excludes": ["tkinter.test"],
    "include_msvcr": True,
}

# GUI applications require a different base on Windows
base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="MERQ Timesheet",
    version="1.0",
    description="MERQ Consultancy Timesheet Application",
    options={"build_exe": build_exe_options},
    executables=[Executable("src/timesheet.py", 
                          base=base,
                          target_name="MERQ_Timesheet.exe",
                          icon="src/merq.ico")]
)