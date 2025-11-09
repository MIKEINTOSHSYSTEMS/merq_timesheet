# build_exe.py
import PyInstaller.__main__
import os
import shutil

def create_version_file():
    """Create a version file for the executable"""
    version_content = '''# UTF-8
#
# For more details about fixed file info 'ffi' see:
# http://msdn.microsoft.com/en-us/library/ms646997.aspx
VSVersionInfo(
  ffi=FixedFileInfo(
    # filevers and prodvers should be always a tuple with four items: (1, 2, 3, 4)
    # Set not needed items to zero 0.
    filevers=(1, 0, 0, 0),
    prodvers=(1, 0, 0, 0),
    # Contains a bitmask that specifies the valid bits 'flags'r
    mask=0x3f,
    # Contains a bitmask that specifies the Boolean attributes of the file.
    flags=0x0,
    # The operating system for which this file was designed.
    # 0x4 - NT and there is no need to change it.
    OS=0x40004,
    # The general type of file.
    # 0x1 - the file is an application.
    fileType=0x1,
    # The function of the file.
    # 0x0 - the function is not defined for this fileType
    subtype=0x0,
    # Creation date and time stamp.
    date=(0, 0)
  ),
  kids=[
    StringFileInfo(
      [
        StringTable(
          u'040904B0',
          [StringStruct(u'CompanyName', u'MERQ Consultancy PLC'),
          StringStruct(u'FileDescription', u'MERQ Timesheet - Ethiopian Calendar Work Hour Tracker'),
          StringStruct(u'FileVersion', u'1.0.0.0'),
          StringStruct(u'InternalName', u'MERQ_Timesheet'),
          StringStruct(u'LegalCopyright', u'Copyright © 2025 MERQ Consultancy PLC. All rights reserved.'),
          StringStruct(u'OriginalFilename', u'MERQ_Timesheet.exe'),
          StringStruct(u'ProductName', u'MERQ Timesheet Application'),
          StringStruct(u'ProductVersion', u'1.0.0.0')])
      ]),
    VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)
'''
    
    version_file = 'version_info.txt'
    with open(version_file, 'w', encoding='utf-8') as f:
        f.write(version_content)
    
    return version_file

def build_executable():
    # Clean previous builds
    if os.path.exists('build'):
        shutil.rmtree('build')
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    
    # Create version file
    version_file = create_version_file()
    
    # PyInstaller configuration
    args = [
        'src/timesheet.py',
        '--name=MERQ_Timesheet',
        '--onefile',
        '--windowed',
        '--icon=src/merq.ico',  # Changed to .ico for better Windows support
        '--version-file=' + version_file,
        '--add-data=src/merq.png;.',
        '--add-data=src/merq.ico;.',
        '--add-data=src/MERQ_TIMESHEET_ETH-CAL_TEMPLATE.xlsx;.',
        '--hidden-import=openpyxl',
        '--hidden-import=PIL',
        '--hidden-import=PIL._tkinter_finder',
        '--hidden-import=tkinter',
        '--hidden-import=requests',
        '--hidden-import=json',
        '--hidden-import=threading',
        '--hidden-import=datetime',
        '--hidden-import=os',
        '--hidden-import=math',
        '--hidden-import=webbrowser',
        '--noconfirm',
        '--clean'
    ]
    
    try:
        PyInstaller.__main__.run(args)
        print("Build completed! Executable is in the 'dist' folder.")
        
        # Clean up version file
        if os.path.exists(version_file):
            os.remove(version_file)
            
    except Exception as e:
        print(f"Build failed: {e}")
        # Clean up version file even if build fails
        if os.path.exists(version_file):
            os.remove(version_file)

if __name__ == '__main__':
    build_executable()