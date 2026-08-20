"""Locations KherveFitting reads from and writes to at runtime.

On Windows the application keeps its user data (Backup folder, Data-Examples,
config.json) next to the executable, and on macOS next to the .app bundle. A
Linux package installs the executable into a read-only prefix such as
/opt/khervefitting, so that data has to live in the user's home directory
instead. The launcher shipped with the Debian package exports
KHERVEFITTING_DATA_DIR and starts the application from that folder.
"""

import os
import platform
import sys


def is_linux_package():
    """Return True when running as the frozen Linux build."""
    return getattr(sys, 'frozen', False) and platform.system() == 'Linux'


def user_data_dir():
    """Return the writable folder holding Backup, Data-Examples and config.json."""
    data_dir = os.environ.get('KHERVEFITTING_DATA_DIR')
    if not data_dir:
        xdg_data_home = os.environ.get('XDG_DATA_HOME') or os.path.join(os.path.expanduser('~'), '.local', 'share')
        data_dir = os.path.join(xdg_data_home, 'KherveFitting')
    return data_dir
