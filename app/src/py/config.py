import os
import sys

def get_root_path():
    """ Returns the root path of the application.
    This function checks if the application is running as a bundled executable
    (e.g., created with PyInstaller) or as a regular Python script.
    If bundled, it returns the temporary folder where the bundled files are located.
    If running as a script, it returns the absolute path to the project root directory.
    """
    if getattr(sys, 'frozen', False):
        # temp folder with bundled files
        return getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    else:
        # else get to the project root
        return os.path.abspath(os.path.join(os.path.dirname(__file__), '../../../'))

ROOT_DIR = get_root_path()
BIN_DIR = os.path.join(ROOT_DIR, 'app', 'bin')
SRC_DIR = os.path.join(ROOT_DIR, 'app', 'src')
WEB_DIR = os.path.join(ROOT_DIR, 'app', 'web')
INDEX_HTML = os.path.join(WEB_DIR, 'index.html')