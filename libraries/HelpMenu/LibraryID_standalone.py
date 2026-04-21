#!/usr/bin/env python3
# KherveFitting - XPS Data Analysis Software
# Copyright (C) 2024-2026 Gwilherm Kerherve <g.kerherve@ic.ac.uk>
#
# KherveFitting is dual-licensed:
#   - GNU GPL v3.0 (see LICENSE-GPL.txt) for open-source use
#   - Commercial Licence (see LICENSE-COMMERCIAL.txt) for proprietary use
# SPDX-License-Identifier: GPL-3.0-only OR LicenseRef-KherveFitting-Commercial

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
from libraries.HelpMenu.LibraryID import PeriodicTableXPS

if __name__ == "__main__":
    app = PeriodicTableXPS()
    app.mainloop()