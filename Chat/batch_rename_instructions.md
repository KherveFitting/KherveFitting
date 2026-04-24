# Batch Rename Feature for Sample Manager (FileManager)

## Summary

When multiple sheets from the same column are selected in the Sample Manager grid, the rename facility should propose renaming the base name while preserving each sheet's row number suffix.

**Example:** Select `VB1`, `VB2`, `VB3` → Rename → user types `Fermi` → sheets become `Fermi1`, `Fermi2`, `Fermi3`.

Sheets without a numeric suffix (e.g. bare `VB`) are also supported — they map to the bare new name (e.g. `Fermi`), while `VB0` maps to `Fermi0`.

## What to change

All changes are in `libraries/ViewMenu/FileManager.py`, inside the `FileManagerWindow` class. Three things to add/modify:

### 1. Add two new methods: `_extract_base_and_index` and `_batch_rename_sheets`

Add these two methods to `FileManagerWindow`, just before `on_rename`:

```python
def _extract_base_and_index(self, sheet_name):
    """Extract base name and numeric index from a sheet name (e.g. 'VB3' -> ('VB', 3, True))
    Returns (base_name, index, has_suffix).
    has_suffix is False when the name has no trailing digits (e.g. 'VB' -> ('VB', 0, False)).
    """
    if "Raman_" in sheet_name or "Ra_" in sheet_name:
        base_parts = sheet_name.split('_')
        base_name = base_parts[0] + "_" + base_parts[1]
        if len(base_parts) > 2 and base_parts[2].isdigit():
            return base_name, int(base_parts[2]), True
        return base_name, 0, False
    match = re.match(r'([\w\-~.]+?)(\d+)$', sheet_name)
    if match:
        return match.group(1), int(match.group(2)), True
    return sheet_name, 0, False

def _batch_rename_sheets(self, sheet_names):
    """Rename multiple sheets that share a common base name, preserving row numbers."""
    # Extract base names and indices
    # has_suffix: whether the original name had an explicit numeric suffix
    parsed = []
    for name in sheet_names:
        base, idx, has_suffix = self._extract_base_and_index(name)
        parsed.append((name, base, idx, has_suffix))

    # Check all share the same base name
    bases = set(b for _, b, _, _ in parsed)
    if len(bases) != 1:
        self.parent.show_popup_message2("Invalid Selection",
            "All selected sheets must belong to the same core level column.")
        return
    old_base = bases.pop()

    dlg = wx.TextEntryDialog(self,
        f"Enter new base name for {', '.join(sheet_names)}:\n"
        f"(Row numbers will be preserved)",
        "Rename Core Levels", old_base)
    if dlg.ShowModal() == wx.ID_OK:
        new_base = dlg.GetValue().strip()
        if not new_base or new_base == old_base:
            dlg.Destroy()
            return
        if len(new_base.split()) > 1:
            self.parent.show_popup_message2("Invalid Name",
                "Only single words are allowed for sheet names.")
            dlg.Destroy()
            return

        # Check for conflicts
        existing = set(self.parent.Data['Core levels'].keys())
        old_names_set = set(n for n, _, _, _ in parsed)
        rename_map = {}
        for old_name, _, idx, has_suffix in parsed:
            # Preserve original suffix style: "VB" (no suffix) -> "Fermi", "VB0" -> "Fermi0"
            new_name = f"{new_base}{idx}" if has_suffix else new_base
            if new_name in existing and new_name not in old_names_set:
                self.parent.show_popup_message2("Name Conflict",
                    f"'{new_name}' already exists. Rename aborted.")
                dlg.Destroy()
                return
            rename_map[old_name] = new_name

        # Sort by index descending to avoid conflicts during rename
        sorted_items = sorted(rename_map.items(),
            key=lambda x: next(idx for n, _, idx, _ in parsed if n == x[0]), reverse=True)

        # Save position/size before closing
        pos = self.GetPosition()
        size = self.GetSize()

        # Temporarily hide file_manager from rename_sheet so it doesn't
        # close/reopen this window on every iteration
        self.parent.file_manager = None
        try:
            for old_name, new_name in sorted_items:
                self.parent.sheet_combobox.SetValue(old_name)
                from libraries.Sheet_Operations import on_sheet_selected
                on_sheet_selected(self.parent, old_name)
                from libraries.Utilities import rename_sheet
                rename_sheet(self.parent, new_name)
        finally:
            pass

        parent = self.parent
        self.Close()
        self.Destroy()

        def reopen():
            from libraries.ViewMenu.FileManager import FileManagerWindow
            parent.file_manager = FileManagerWindow(parent)
            parent.file_manager.SetPosition(pos)
            parent.file_manager.SetSize(size)
            parent.file_manager.Show()

        wx.CallLater(50, reopen)
        return
    dlg.Destroy()
```

### 2. Modify `on_rename` to route to batch rename when multiple sheets are selected

The existing `on_rename` only uses `sheet_names[0]`. Change it so that when `len(sheet_names) > 1`, it calls `self._batch_rename_sheets(sheet_names)` and returns early. The single-sheet path stays unchanged.

**Before:**
```python
def on_rename(self, event):
    """Rename the selected core level"""
    save_state(self.parent)
    sheet_names = self.get_selected_sheet_names()

    if sheet_names:
        sheet_name = sheet_names[0]  # Use the first selected sheet
        # ... single rename dialog ...
```

**After:**
```python
def on_rename(self, event):
    """Rename the selected core level(s)"""
    save_state(self.parent)
    sheet_names = self.get_selected_sheet_names()

    if not sheet_names:
        return

    if len(sheet_names) > 1:
        self._batch_rename_sheets(sheet_names)
        return

    sheet_name = sheet_names[0]
    # ... rest of single rename dialog unchanged ...
```

### 3. Modify the right-click context menu in `on_grid_right_click`

Find the rename section in `on_grid_right_click`. When multiple sheets are selected, show "Rename N Core Levels..." and route to `on_rename` (which will dispatch to batch). Otherwise keep the existing single-sheet rename.

**Before:**
```python
# Add rename option for core level cells only
if col > 0 and col <= len(self.core_levels):
    cell_value = self.grid.GetCellValue(row, col)
    if cell_value and cell_value in self.parent.Data['Core levels']:
        menu.AppendSeparator()
        rename_item = menu.Append(wx.ID_ANY, f"Rename '{cell_value}'")
        self.Bind(wx.EVT_MENU, lambda evt, sheet=cell_value: self.rename_from_context_menu(sheet), rename_item)
```

**After:**
```python
# Add rename option for core level cells only
if col > 0 and col <= len(self.core_levels):
    cell_value = self.grid.GetCellValue(row, col)
    if cell_value and cell_value in self.parent.Data['Core levels']:
        selected = self.get_selected_sheet_names()
        menu.AppendSeparator()
        if len(selected) > 1:
            rename_item = menu.Append(wx.ID_ANY, f"Rename {len(selected)} Core Levels...")
            self.Bind(wx.EVT_MENU, lambda evt: self.on_rename(evt), rename_item)
        else:
            rename_item = menu.Append(wx.ID_ANY, f"Rename '{cell_value}'")
            self.Bind(wx.EVT_MENU, lambda evt, sheet=cell_value: self.rename_from_context_menu(sheet), rename_item)
```

## Key design decisions

- **`has_suffix` flag**: Distinguishes `VB` (no digits, index=0, has_suffix=False) from `VB0` (explicit zero, index=0, has_suffix=True). A bare name like `VB` becomes bare `Fermi`; `VB0` becomes `Fermi0`.
- **Descending sort**: Sheets are renamed from highest index down to avoid name collisions mid-rename (e.g. renaming `VB1→Fermi1` before `VB2` still exists as `VB2`).
- **file_manager suppression**: `rename_sheet()` in `Utilities.py` closes and reopens the file manager window on each call. During batch rename, we set `self.parent.file_manager = None` so it skips that. After all renames complete, we close/destroy the window and use `wx.CallLater(50, reopen)` to recreate it at the same position and size.
- **Conflict check**: Before renaming anything, all target names are checked against existing sheet names. If any conflict is found, the entire batch is aborted.
- **Same base name requirement**: All selected sheets must share the same base name (i.e. belong to the same column). If they don't, an error is shown.

## Dependencies

- `re` (already imported in FileManager.py)
- `rename_sheet` from `libraries/Utilities.py` (existing function, no changes needed)
- `save_state` from `libraries/State_Management.py` (already used in `on_rename`)
- `get_selected_sheet_names` (existing method on `FileManagerWindow`)
- `show_popup_message2` (existing method on the parent window)
