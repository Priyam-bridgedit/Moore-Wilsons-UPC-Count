import sys
from cx_Freeze import setup, Executable

# Build options
build_options = {
    'packages': ['os', 'tkinter', 'pandas', 'pyodbc', 'configparser', 'apscheduler'],
    'excludes': [],
    'include_files': ['upccount_config.ini'],  # Include the config.ini file
}

# Determine the base for the current platform
base = None
if sys.platform == "win32":
    base = "Win32GUI"

# Executable options
executables = [
    Executable(
        script='UPC.py',  # This is the target script for the first executable
        base=base,  # Use the platform-specific base
        targetName='UPCMain.exe',
    ),
    Executable(
        script='upccount.py',  # This is the target script for the second executable
        base=base,  # Use the platform-specific base
        targetName='UPCBackground.exe',
    )
]

# Create the setup
setup(
    name='Moore Wilsons Report',  # Name of the application
    version='1.0',
    description='Moore Wilsons Report',
    options={'build_exe': build_options},
    executables=executables,
)
