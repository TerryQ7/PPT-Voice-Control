"""English entry point for PyInstaller packaging.

Behaves exactly like ``main.py --lang en``. Having a dedicated file lets
PyInstaller (and Windows shortcuts) target the English build without relying
on command-line arguments.
"""

import os
import sys

os.environ["PPT_LANG"] = "en"
sys.argv = [sys.argv[0], "--lang", "en", *sys.argv[1:]]

from main import main

if __name__ == "__main__":
    main()
