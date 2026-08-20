#!/bin/sh
# Launcher installed as /usr/bin/khervefitting by the Debian package.
#
# KherveFitting keeps its user data (config.json, Backup, Data-Examples,
# Peaks Library) in the folder it runs from. The package installs the program
# under /opt, which is read only for normal users, so create a per-user data
# folder and start the application from there.
set -e

INSTALL_DIR=/opt/khervefitting
DATA_DIR="${KHERVEFITTING_DATA_DIR:-${XDG_DATA_HOME:-$HOME/.local/share}/KherveFitting}"

mkdir -p "$DATA_DIR/Backup"

# Copy the read only folders shipped with the package on first run. Existing
# files are never overwritten, so edits and new files survive an upgrade.
for folder in Data-Examples 'Peaks Library'; do
    if [ -d "$INSTALL_DIR/$folder" ]; then
        mkdir -p "$DATA_DIR/$folder"
        cp -Rn "$INSTALL_DIR/$folder/." "$DATA_DIR/$folder/" 2>/dev/null || true
    fi
done

# The application only initialises its settings when a configuration file is
# present, so seed the default one shipped with the package on first run.
if [ ! -f "$DATA_DIR/config.json" ] && [ -f "$INSTALL_DIR/config.json" ]; then
    cp "$INSTALL_DIR/config.json" "$DATA_DIR/config.json"
fi

KHERVEFITTING_DATA_DIR="$DATA_DIR"
export KHERVEFITTING_DATA_DIR

cd "$DATA_DIR"
exec "$INSTALL_DIR/KherveFitting" "$@"
