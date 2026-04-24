# KAL duplicate regions, Casa Info parsing, and Sample Manager metadata columns

This covers three related changes:

1. Fix the KAL reader so duplicate core-level regions (multiple Sample Positions) aren't dropped.
2. Parse the Kratos "Casa Info Follows" block by keyword in both KAL and VAMAS readers, adding the fields to each sheet's Experimental Description.
3. In the Sample Manager (`FileManager`), add a right-click "Add column" submenu that appends a metadata-filled column, persists it across sessions, lets the user remove it, and grows the window so the new column is visible.

All three changes live in two files:

- `libraries/FileMenu/Open.py`
- `libraries/ViewMenu/FileManager.py`
- plus 2 lines of config wiring in `KherveFitting.py`

---

## 1. KAL duplicate regions

### Why

Kratos `.kal` files can contain multiple `Sample Position` blocks, each followed by the same set of acquisition regions (e.g. `Ce 3d`, `O 1s`, `wide` …). The previous reader keyed `spectra` by region name alone, so the second sample position silently overwrote the first. The ReadMe in `DATA/Conversion/Kratos Data files/` even notes: *"Sample 3 contains 2 sets of data"*, but the conversion only produced the first.

### What to change in `convert_kal_to_excel` (in `libraries/FileMenu/Open.py`)

**a.** Replace the sample-ID extraction + main block loop header:

```python
blocks = content.split('Dataset filename')
spectra = []  # list of dicts, preserves duplicates across sample positions
PHOTON_ENERGY = 1486.67

# Track the current sample position as blocks are processed.
# Kratos only records these once per sample position, not per acquisition region.
current_sample_id = "Unknown"
current_sample_tilt = ''

for block in blocks:
    # Update current sample position / tilt whenever a Sample Position block is seen.
    for line in block.split('\n'):
        if 'Stage Position Name' in line:
            current_sample_id = line.split('=')[1].strip()
        elif 'Stage X Rotation' in line:
            current_sample_tilt = line.split('=')[1].strip()

    if 'Ordinate values' not in block or 'Object name' not in block:
        continue
    # ... rest of per-block parsing unchanged ...
```

**b.** In the `metadata` dict, change:

```python
'Sample ID': sample_id,
```

to:

```python
'Sample ID': current_sample_id,
```

and add two new entries (between `Sputtering Energy` and `Take-off Polar Angle`):

```python
'Sputter Time': '',
'Sample Tilt': current_sample_tilt,
```

**c.** Change the spectra-storage line at the bottom of the per-block loop:

```python
# OLD:
# spectra[name] = {'data': df, 'metadata': metadata, 'id': spectrum_id}

# NEW:
spectra.append({
    'sample_id': current_sample_id,
    'name': name,
    'data': df,
    'metadata': metadata,
    'id': spectrum_id,
})
```

**d.** Update the sheet-creation loop to iterate the list and disambiguate duplicate sheet names (same style as VAMAS):

```python
for info in spectra:
    name = info['name']
    base_sheet_name = name.replace(' ', '')
    sheet_name = base_sheet_name
    if sheet_name in wb.sheetnames:
        count = 1
        while f"{base_sheet_name}{count}" in wb.sheetnames:
            count += 1
        sheet_name = f"{base_sheet_name}{count}"
    df = info['data']
    metadata = info['metadata']
    ws = wb.create_sheet(sheet_name)
    # ... rest unchanged ...
```

**e.** In the same loop, the `fields` list that drives Experimental Description rows becomes:

```python
fields = [
    'Sample ID', 'Date', 'Time', 'Technique', 'Species & Transition',
    'Number of scans', 'Source Label', 'Source Energy', 'Source width X',
    'Source width Y', 'Pass Energy', 'Work Function', 'Analyzer Mode',
    'Sputtering Energy', 'Sputter Time', 'Sample Tilt',
    'Take-off Polar Angle', 'Take-off Azimuth',
    'Target Bias', 'Analysis Width X', 'Analysis Width Y', 'X Label',
    'X Units', 'X Start', 'X Step', 'Num Y Values', 'Num Scans',
    'Collection Time', 'Time Correction', 'Y Unit', '# Comment Lines',
    'Block Comment'
] + list(CASA_INFO_FIELDS)  # CASA_INFO_FIELDS is defined below
```

**f.** The "Experimental description" (global sheet) loop also needs to iterate the list:

```python
row = 1
for info in spectra:
    name = info['name']
    metadata = info['metadata']
    spectrum_id = info['id']
    header = exp_sheet.cell(row=row, column=1,
                            value=f"Spectrum: {name}/{spectrum_id} [{info['sample_id']}]")
    # ... rest unchanged ...
```

### Result

A Kratos file with 2 sample positions × 6 regions now produces 12 sheets (`Ce3d`, …, `wide`, `Ce3d1`, …, `wide1`) instead of silently dropping the second set.

---

## 2. Casa Info parsing (KAL + VAMAS)

### Why

Both `.vms` and `.kal` files (when converted by Kratos/Casa tools) contain a block starting with `Casa Info Follows` whose first ~12 lines hold useful metadata:

```
Casa Info Follows
0
0
0
0
CeO2 160C
(0.04701 m, -2.9e-05 m, -0.00094 m)  Angle: 0 degrees
FoV Position [0,0] um
Lens Mode:HSA_LENS_HYBRID
Neutraliser : On
Charge Balance = 3.8 : Filament Current = 0.38 : Filament Bias = 1.4 : Magnet Lens Trim Coil = 0.175
Aperture Description: Slot
Iris Position Description: slot
```

Each of these values should appear as its own row in the sheet's Experimental Description.

### Parser — **anchor by keyword, not by line number**

Add this to `libraries/FileMenu/Open.py` **just before** `def parse_casa_peak_fitting(...)`:

```python
CASA_INFO_FIELDS = [
    'Casa Sample Name', 'Stage Position', 'Angle', 'FoV Position', 'Lens Mode',
    'Neutraliser', 'Charge Balance', 'Filament Current', 'Filament Bias',
    'Magnet Lens Trim Coil', 'Aperture', 'Iris Position'
]


def parse_casa_info_lines(comment_text):
    """Extract Kratos/Casa informational fields from a 'Casa Info Follows' block.

    Returns a dict keyed by the labels in CASA_INFO_FIELDS. Missing fields map to
    empty strings. Lines are matched by keyword anchors (not by position) so the
    parser survives reordering or missing entries.
    """
    result = {k: '' for k in CASA_INFO_FIELDS}
    if not comment_text or 'Casa Info Follows' not in comment_text:
        return result

    tail = comment_text.split('Casa Info Follows', 1)[1]
    candidate_lines = [ln.strip() for ln in tail.split('\n') if ln.strip()][:15]

    # Sample Name: first non-numeric, non-structured line.
    for ln in candidate_lines:
        try:
            float(ln); continue
        except ValueError:
            pass
        if ln.lower().startswith(('xps', 'aes', 'uhv')):
            break
        if ':' in ln or '(' in ln or '=' in ln or ln.startswith('FoV') \
                or ln.startswith('Aperture') or ln.startswith('Iris'):
            break
        result['Casa Sample Name'] = ln
        break

    for ln in candidate_lines:
        if ln.startswith('(') and 'Angle' in ln:
            paren_end = ln.find(')')
            if paren_end != -1:
                result['Stage Position'] = ln[:paren_end + 1]
            if 'Angle:' in ln:
                result['Angle'] = ln.split('Angle:', 1)[1].strip()
        elif ln.startswith('FoV Position'):
            result['FoV Position'] = ln.split('FoV Position', 1)[1].strip()
        elif ln.startswith('Lens Mode'):
            result['Lens Mode'] = ln.split(':', 1)[1].strip() if ':' in ln else ''
        elif ln.startswith('Neutraliser'):
            result['Neutraliser'] = ln.split(':', 1)[1].strip() if ':' in ln else ''
        elif ln.startswith('Charge Balance'):
            for part in ln.split(':'):
                if '=' not in part:
                    continue
                key, val = part.split('=', 1)
                key, val = key.strip(), val.strip()
                if key in result:
                    result[key] = val
        elif ln.startswith('Aperture Description'):
            result['Aperture'] = ln.split(':', 1)[1].strip() if ':' in ln else ''
        elif ln.startswith('Iris Position Description'):
            result['Iris Position'] = ln.split(':', 1)[1].strip().rstrip('.') if ':' in ln else ''

    return result
```

### Wire it into the KAL reader

In `convert_kal_to_excel`, **after** the existing `metadata['Block Comment'] = ...` line, add:

```python
# Parse Casa Info keyword-anchored fields.
casa_info = parse_casa_info_lines('\n'.join(comment_lines))
for k, v in casa_info.items():
    metadata[k] = v
```

The KAL `fields` list already gets `+ list(CASA_INFO_FIELDS)` appended (see change 1e above).

### Wire it into the VAMAS reader

In `open_vamas_file`, after `block_exp_data = [...]` is built and **before** `exp_data.append(block_exp_data)`, add:

```python
# Parse Casa Info keyword-anchored fields from the block comment.
casa_info = parse_casa_info_lines(block.block_comment)
block_exp_data.extend(casa_info.get(f, '') for f in CASA_INFO_FIELDS)
```

Then both `exp_labels` (the per-sheet list, ~line 3518) and `block_info_order` (the global sheet list, ~line 3561) need this at the end:

```python
] + list(CASA_INFO_FIELDS)
```

(The same `exp_labels` also needs `"Sputter Time", "Sample Tilt",` inserted between `"Sputtering Energy"` and `"Take-off Polar Angle"`, matching the KAL change above.)

### Sample Tilt and Sputter Time in the VAMAS block_exp_data

Before the `block_exp_data = [...]` list, compute:

```python
sputter_source = getattr(block, 'sputtering_source', None)
sputter_time_val = ''
if sputter_source is not None:
    # VAMAS has no dedicated sputter duration; expose the mode as a proxy.
    sputter_time_val = getattr(sputter_source, 'mode', '') or ''
sample_tilt_val = getattr(block, 'sample_normal_polar_angle_tilt', '')
```

Then insert between `block.sputtering_source_energy ...` and `block.analyzer_axis_take_off_polar_angle,`:

```python
sputter_time_val,
sample_tilt_val,
```

---

## 3. Sample Manager: "Add column" right-click + config persistence

### Why

With Casa Info fields now living in each sheet's Experimental Description, the user wants a way to surface any of them (Angle, Sample Tilt, Lens Mode, etc.) as a column in the FileManager grid. The column:

- Must persist across restarts (stored in `config.json`).
- Must be removable via right-click on the column (but only the user-added ones — built-ins are protected).
- Must trigger a window resize so the new column is actually visible.

### Config wiring — `KherveFitting.py`

In `load_config`, add inside the `if os.path.exists('config.json'):` branch (near `heatmap_colormap`):

```python
self.file_manager_custom_columns = config.get('file_manager_custom_columns', [])
```

In `save_config`, add in the config dict:

```python
'file_manager_custom_columns': getattr(self, 'file_manager_custom_columns', []),
```

### FileManager changes — `libraries/ViewMenu/FileManager.py`

#### a. Placeholder in `__init__`

After `self.load_be_corrections()` (before the window positioning block):

```python
# Placeholder for user-added metadata columns; actual rebuild happens
# after populate_grid() runs (otherwise rows have no sheet names yet).
self.custom_columns = []
```

#### b. Restore after populate_grid

After `self.populate_grid()` (which is the existing call in `__init__`):

```python
# Restore any user-added metadata columns from the persisted config.
for label in list(getattr(self.parent, 'file_manager_custom_columns', []) or []):
    self._append_metadata_column(label, persist=False, grow_window=False)
if self.custom_columns:
    self._grow_window_for_columns()
```

> **Important**: restore must run after `populate_grid()`, otherwise `_get_row_sheet_names()` returns empty lists and every new column is blank.

#### c. Right-click menu additions — `on_grid_right_click`

Just before the final `self.grid.PopupMenu(menu)`:

```python
# "Add column" submenu — one entry per unique experimental_description label.
metadata_keys = self._get_available_metadata_fields()
if metadata_keys:
    menu.AppendSeparator()
    meta_submenu = wx.Menu()
    for key in metadata_keys:
        item = meta_submenu.Append(wx.ID_ANY, key)
        self.Bind(wx.EVT_MENU, lambda evt, k=key: self.fill_column_from_metadata(k), item)
    menu.AppendSubMenu(meta_submenu, "Add column")

# Allow removing a user-added column (not the built-in ones).
if col >= self._first_custom_col_index() and col < self.grid.GetNumberCols():
    label = self.grid.GetColLabelValue(col)
    if label in self.custom_columns:
        remove_item = menu.Append(wx.ID_ANY, f"Remove column '{label}'")
        self.Bind(wx.EVT_MENU, lambda evt, c=col: self.remove_custom_column(c), remove_item)
```

#### d. New methods (place right after `on_grid_right_click`)

```python
def _get_row_sheet_names(self, row):
    """Sheet names present in the core-level columns of a grid row."""
    names = []
    for c in range(1, len(self.core_levels) + 1):
        v = self.grid.GetCellValue(row, c)
        if v and v in self.parent.Data.get('Core levels', {}):
            names.append(v)
    return names

def _get_available_metadata_fields(self):
    """Union of experimental_description labels across all loaded sheets."""
    seen = []
    for sheet_data in self.parent.Data.get('Core levels', {}).values():
        exp = sheet_data.get('experimental_description', [])
        for item in exp:
            if not item:
                continue
            key = str(item[0]).strip() if len(item) >= 1 else ''
            if key and key not in seen:
                seen.append(key)
    return seen

def fill_column_from_metadata(self, field_name):
    """Append a user-added column filled with the given metadata field."""
    if field_name in self.custom_columns:
        wx.MessageBox(f"Column '{field_name}' already exists.",
                      "Duplicate column", wx.OK | wx.ICON_INFORMATION)
        return
    self._append_metadata_column(field_name, persist=True, grow_window=True)

def _append_metadata_column(self, field_name, persist=True, grow_window=True):
    new_col = self.grid.GetNumberCols()
    self.grid.AppendCols(1)
    self.grid.SetColLabelValue(new_col, field_name)

    for row in range(self.grid.GetNumberRows()):
        sheets = self._get_row_sheet_names(row)
        value = ''
        for sheet in sheets:
            exp = self.parent.Data['Core levels'][sheet].get('experimental_description', [])
            for item in exp:
                if len(item) >= 2 and str(item[0]).strip() == field_name:
                    v = str(item[1]).strip() if item[1] is not None else ''
                    if v:
                        value = v
                        break
            if value:
                break
        self.grid.SetCellValue(row, new_col, value)

    self.grid.AutoSizeColumn(new_col)
    col_size = self.grid.GetColSize(new_col)
    self.grid.SetColSize(new_col, max(80, min(col_size, 200)))
    self.custom_columns.append(field_name)

    if persist:
        saved = getattr(self.parent, 'file_manager_custom_columns', []) or []
        if field_name not in saved:
            saved.append(field_name)
        self.parent.file_manager_custom_columns = saved
        if hasattr(self.parent, 'save_config'):
            self.parent.save_config()

    if grow_window:
        self._grow_window_for_columns()

    self.grid.ForceRefresh()

def remove_custom_column(self, col):
    first_custom = self._first_custom_col_index()
    if col < first_custom or col >= self.grid.GetNumberCols():
        return
    label = self.grid.GetColLabelValue(col)
    self.grid.DeleteCols(col, 1)
    if label in self.custom_columns:
        self.custom_columns.remove(label)

    saved = getattr(self.parent, 'file_manager_custom_columns', []) or []
    if label in saved:
        saved.remove(label)
        self.parent.file_manager_custom_columns = saved
        if hasattr(self.parent, 'save_config'):
            self.parent.save_config()

    self._grow_window_for_columns()
    self.grid.ForceRefresh()

def _first_custom_col_index(self):
    """Index of the first user-added column (after Experiment + core levels + Xshift + 2 Norm)."""
    return len(self.core_levels) + 4

def _grow_window_for_columns(self):
    """Resize the window so all grid columns are visible."""
    total = self.grid.GetRowLabelSize() + 40
    for c in range(self.grid.GetNumberCols()):
        total += self.grid.GetColSize(c)
    total = max(total, 630)
    current = self.GetSize()
    if total != current.width:
        self.SetSize(total, current.height)
```

---

## Testing checklist

1. **KAL duplicate regions**: open `DATA/Conversion/Kratos Data files/20191030_03.kal` — should produce 12 sheets, with the second 6 suffixed `1` (e.g. `Ce3d`, `Ce3d1`). Each sheet's Experimental Description should show the correct `Sample ID` (e.g. `CeO2 160C` vs `CeO2 180C`).
2. **Casa Info on VAMAS**: open `20191030_03.vms`. Open the Info window for any sheet — it should now list `Casa Sample Name`, `Angle`, `Lens Mode`, `Aperture`, etc. with real values.
3. **Add column**: right-click any cell in the Sample Manager → "Add column" → choose `Angle`. New column appears on the right, filled with per-row angle values, window grows to show it.
4. **Persistence**: close and reopen the Sample Manager. Custom columns come back filled (not empty).
5. **Remove**: right-click inside a custom column → "Remove column '…'" → column disappears and the window shrinks.
6. **Built-in protection**: right-clicking inside `Experiment`, any core-level column, `Xshift`, or `Norm. @ BE` / `Norm. to A` shows no "Remove column" option.

---

## Notes for the other build

- If the paid AI build has a different version of `Open.py`, the parsers are additive — `CASA_INFO_FIELDS` and `parse_casa_info_lines` are standalone and can be dropped in verbatim.
- The FileManager methods are all new and don't touch any of the AI-specific code paths.
- `experimental_description` on `window.Data['Core levels'][sheet]` is populated at xlsx load time by `libraries/FileMenu/Save.py` (the refresh path). As long as that path runs, the "Add column" submenu will see the keys. If the paid build has a different loader, verify that `experimental_description` is populated as a list of `[label, value]` pairs.
