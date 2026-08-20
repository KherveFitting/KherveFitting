# Building the Linux (Debian) package

`build_deb.sh` freezes KherveFitting with PyInstaller and wraps the result in a
`.deb`. The package installs a normal desktop application: it appears in the
applications menu, registers the XPS file types, and runs on any desktop that
follows the freedesktop standards — Xfce (Xubuntu, Debian Xfce, MX Linux),
GNOME, KDE, Cinnamon, MATE and LXQt alike. There is nothing Xfce specific to
build; the `.desktop` entry Xfce reads is the same one every other desktop uses.

## Prerequisites

Build on the **oldest** distribution you want to support. The frozen binary
links against the build machine's glibc and needs one at least that new at run
time: Ubuntu 22.04 ships glibc 2.35, Debian 12 ships 2.36 and Ubuntu 24.04 ships
2.39, so a package built on Ubuntu 22.04 covers all three, while one built on
Ubuntu 24.04 runs on neither of the older two. The architecture is fixed at
build time as well — build on arm64 to get an arm64 package.

```bash
sudo apt install python3-venv python3-wxgtk4.0 dpkg-dev binutils
```

`python3-wxgtk4.0` is the quickest way to get wxPython: the PyPI package has no
Linux wheels and `pip install wxPython` compiles it from source, which takes
about half an hour and needs the GTK development headers.

Create the build environment, reusing the system wxPython:

```bash
cd /path/to/KherveFitting
python3 -m venv --system-site-packages .venv-build
.venv-build/bin/pip install -r <(grep -v '^wxPython' requirements.txt) pyinstaller
```

Some pins in `requirements.txt` (for example `lmfitxps==0.2.3`) only resolve on
Python 3.11 and older, which is what Debian 12 ships. On a distribution with
Python 3.12 or newer, install those packages unpinned:

```bash
.venv-build/bin/pip install -r <(grep -v '^wxPython' requirements.txt | sed 's/==.*//') pyinstaller
```

## Building

```bash
.venv-build/bin/python -c 'import wx'          # sanity check
packaging/linux/build_deb.sh --version 1.80 --python .venv-build/bin/python
```

The package is written to `dist/khervefitting_1.80-1_amd64.deb`. Useful options:

| Option | Effect |
| --- | --- |
| `--version X.YY` / `--revision N` | version recorded in the package |
| `--python PATH` | interpreter that provides wxPython and PyInstaller |
| `--skip-pyinstaller` | reuse an existing `dist/KherveFitting` folder |
| `--no-examples` | drop the 32 MB `Data-Examples` folder |
| `--output-dir DIR` | where to write the `.deb` |

The package is around 195 MB and installs about 610 MB, most of it PyArrow.
KherveFitting only uses it to read Parquet files and `fastparquet` handles those
too, so adding `'pyarrow'` to the `excludes` list in `KherveFittingLinux.spec`
cuts roughly 150 MB if size matters more than read speed.

## Installing

```bash
sudo apt install ./dist/khervefitting_1.80-1_amd64.deb   # pulls in the GTK libraries
sudo apt remove khervefitting                            # uninstall
```

`apt` is preferred over `dpkg -i` because it resolves the dependencies listed in
the control file.

## What the package contains

| Path | Contents |
| --- | --- |
| `/opt/khervefitting/` | frozen application, bundled Python, data files, manual |
| `/usr/bin/khervefitting` | launcher script |
| `/usr/share/applications/khervefitting.desktop` | menu entry and file associations |
| `/usr/share/icons/hicolor/*/apps/khervefitting.png` | icons |
| `/usr/share/mime/packages/khervefitting.xml` | `.vms`, `.kal`, `.spe`, `.avg`, `.mrs` file types |
| `/usr/share/doc/khervefitting/` | licence and third party notices |

## Where user data goes

`/opt` is read only for normal users, so the launcher creates
`~/.local/share/KherveFitting` (or `$XDG_DATA_HOME/KherveFitting`), starts the
application from there, and exports `KHERVEFITTING_DATA_DIR`. That folder holds
`config.json`, `Backup/`, and writable copies of `Data-Examples/` and
`Peaks Library/`, which are seeded from `/opt` on first run and never
overwritten afterwards, so preferences and edits survive an upgrade.

Point `KHERVEFITTING_DATA_DIR` somewhere else to override it:

```bash
KHERVEFITTING_DATA_DIR=~/xps-work khervefitting
```

Removing the package leaves that folder in place. Delete it by hand to start
from a clean configuration.

## Differences from the Windows and macOS builds

Three things are handled differently here, all in `KherveFittingLinux.spec`,
`build_deb.sh` or `launcher.sh` if you want to change them:

- **System GTK.** PyInstaller collects the GTK libraries wxWidgets links
  against but not their gdk-pixbuf loaders or icon themes, which leaves the
  bundled copy unable to draw some assets. The spec drops those libraries so
  the application uses the system GTK and follows the desktop's theme; the
  package depends on it. Delete the `SYSTEM_GTK_PREFIXES` block in the spec for
  a fully self contained bundle like the Windows one.
- **A default `config.json` is shipped.** The application only initialises many
  of its settings when a configuration file is present — without one it raises
  `'MyFrame' object has no attribute 'usage_tracking'` on the first file it
  opens. The Windows and macOS bundles always carry one, so this never shows up
  there. The launcher seeds the packaged default into the user's data folder on
  first run.
- **Liberation Sans instead of Arial.** The default plot font does not exist on
  Linux. The shipped `config.json` selects Liberation Sans, which has Arial's
  metrics, so plots keep the same proportions instead of falling back to
  DejaVu Sans with a few hundred matplotlib warnings.

## Troubleshooting

**The application does not start and prints nothing.** Run `khervefitting` from
a terminal — the frozen binary is built without a console, but errors still go
to stderr.

**`Gtk-WARNING: cannot open display`.** Under Wayland, install
`libgtk-3-0`'s X11 support or run with `GDK_BACKEND=x11`.

**The KherveDB window is blank.** That view uses WebKit; install the
recommended `libwebkit2gtk-4.1-0`.

**`Could not load a pixbuf from ... check-symbolic.svg`.** The GTK SVG loader is
missing; install `librsvg2-common`, which the package depends on.

**Building on a distribution without `python3-wxgtk4.0`.** Either install
wxPython from the wxPython wheel index for your distribution, or build it from
source with `pip install wxPython` after installing `libgtk-3-dev` and
`libwebkit2gtk-4.1-dev`.
