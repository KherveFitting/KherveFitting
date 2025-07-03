#!/usr/bin/env python3
import sys
import os

# Add current directory to path for imports
if hasattr(sys, '_MEIPASS'):
    # Running as PyInstaller bundle
    current_dir = sys._MEIPASS
else:
    current_dir = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, current_dir)

# Import the main class
from LibraryID import PeriodicTableXPS

if __name__ == "__main__":
    app = PeriodicTableXPS()
    app.mainloop()