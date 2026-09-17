# -*- coding: utf-8 -*-
"""Double-click or: python launch_ste_lab.py"""

from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from ste_lab.gui_app import main

if __name__ == "__main__":
    main()
