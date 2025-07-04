# libraries/grid_operations.py

import wx

def populate_results_grid(window):
    # Get current sheet name and extract row number
    sheet_name = window.sheet_combobox.GetValue()
    row_number = 0

    import re
    match = re.search(r'(\d+)$', sheet_name)
    if match:
        row_number = int(match.group(1))

    results_table_key = f'Results Table{row_number}'

    # FULL RESET OF THE GRID
    # First, remove all rows from the grid
    if window.results_grid.GetNumberRows() > 0:
        window.results_grid.DeleteRows(0, window.results_grid.GetNumberRows())

    # Check if the results table exists and has data
    if results_table_key in window.Data and 'Peak' in window.Data[results_table_key] and window.Data[results_table_key][
        'Peak']:
        results = window.Data[results_table_key]['Peak']

        # Add the exact number of rows needed
        num_rows = len(results)
        if num_rows > 0:
            window.results_grid.AppendRows(num_rows)

        # Ensure enough columns exist
        num_cols = 31
        if window.results_grid.GetNumberCols() < num_cols:
            window.results_grid.AppendCols(num_cols - window.results_grid.GetNumberCols())

        # Populate the grid
        for row, (peak_label, peak_data) in enumerate(results.items()):
            window.results_grid.SetCellValue(row, 0, peak_data.get('Name', ''))
            window.results_grid.SetCellValue(row, 1, str(peak_data.get('Position', '')))
            window.results_grid.SetCellValue(row, 2, str(peak_data.get('Height', '')))
            window.results_grid.SetCellValue(row, 3, str(peak_data.get('FWHM', '')))
            window.results_grid.SetCellValue(row, 4, str(peak_data.get('L/G', '')))
            window.results_grid.SetCellValue(row, 5, f"{peak_data.get('Area', 0):.2f}")
            window.results_grid.SetCellValue(row, 6, f"{peak_data.get('at. %', 0):.2f}")



            window.results_grid.SetCellValue(row, 8, f"{peak_data.get('RSF', 0):.2f}")
            window.results_grid.SetCellValue(row, 9, f"{peak_data.get('TXFN', 1.0):.2f}")
            window.results_grid.SetCellValue(row, 10, str(peak_data.get('ECF', '')))
            window.results_grid.SetCellValue(row, 11, str(peak_data.get('Instrument', 'Al1486')))
            window.results_grid.SetCellValue(row, 12, peak_data.get('Fitting Model', ''))
            window.results_grid.SetCellValue(row, 13, f"{peak_data.get('Rel. Area', 0):.2f}")
            window.results_grid.SetCellValue(row, 14, str(peak_data.get('Sigma', '')))
            window.results_grid.SetCellValue(row, 15, str(peak_data.get('Gamma', '')))
            window.results_grid.SetCellValue(row, 16, peak_data.get('Bkg Type', ''))
            window.results_grid.SetCellValue(row, 17, str(peak_data.get('Bkg Low', '')))
            window.results_grid.SetCellValue(row, 18, str(peak_data.get('Bkg High', '')))
            window.results_grid.SetCellValue(row, 19, str(peak_data.get('Bkg Offset Low', '')))
            window.results_grid.SetCellValue(row, 20, str(peak_data.get('Bkg Offset High', '')))
            window.results_grid.SetCellValue(row, 21, peak_data.get('Sheetname', ''))
            window.results_grid.SetCellValue(row, 22, peak_data.get('Pos. Constraint', ''))
            window.results_grid.SetCellValue(row, 23, peak_data.get('Height Constraint', ''))
            window.results_grid.SetCellValue(row, 24, peak_data.get('FWHM Constraint', ''))
            window.results_grid.SetCellValue(row, 25, peak_data.get('L/G Constraint', ''))

            # if window.results_grid.GetNumberCols() > 27:
            window.results_grid.SetCellValue(row, 26, peak_data.get('Area Constraint', ''))
            window.results_grid.SetCellValue(row, 27, peak_data.get('Sigma Constraint', ''))
            window.results_grid.SetCellValue(row, 28, peak_data.get('Gamma Constraint', ''))
            window.results_grid.SetCellValue(row, 29, f"{peak_data.get('wt. %', 0):.2f}")

            # Calculate and display mass for reference
            from libraries.Area_Calculation import extract_element_symbol, ATOMIC_MASSES
            peak_name = peak_data.get('Name', '')
            element_symbol = extract_element_symbol(peak_name)
            atomic_mass = ATOMIC_MASSES.get(element_symbol, 12.01)
            window.results_grid.SetCellValue(row, 30, f"{atomic_mass:.2f}")

            # # Use custom renderer and editor for checkboxes
            checkbox_state = peak_data.get('Checkbox', '0')
            window.results_grid.SetCellValue(row, 7, checkbox_state)

            # Use custom checkbox renderer
            from libraries.Grid_Operations import CheckboxRenderer
            window.results_grid.SetCellRenderer(row, 7, CheckboxRenderer())
            window.results_grid.SetReadOnly(row, 7)

        for row in range(window.results_grid.GetNumberRows()):
            # Set atomic % column (column 6) to green
            window.results_grid.SetCellBackgroundColour(row, 6, wx.Colour(200, 245, 228))
            bold_font = window.results_grid.GetDefaultCellFont()
            bold_font.SetWeight(wx.FONTWEIGHT_BOLD)
            window.results_grid.SetCellFont(row, 6, bold_font)

            # Set calculated area column (column 3) to green
            window.results_grid.SetCellBackgroundColour(row, 13, wx.Colour(200, 245, 228))
            window.results_grid.SetCellFont(row, 13, bold_font)

            # Set weight % column (column 29) to green
            window.results_grid.SetCellBackgroundColour(row, 29, wx.Colour(200, 245, 228))
            window.results_grid.SetCellFont(row, 29, bold_font)

        # # Force a refresh to ensure renderer is applied
        # window.results_grid.ForceRefresh()

        # Force a refresh to ensure renderer is applied
        window.results_grid.ForceRefresh()

        # Set up column properties (read-only states)
        setup_results_grid_column_properties(window)

        # Add this line at the end, after all the SetCellValue calls
        setup_grid_editability(window)

        # Bind the cell changed event if not already bound
        if not hasattr(window, '_results_grid_event_bound'):
            window.results_grid.Bind(wx.grid.EVT_GRID_CELL_CHANGED,
                                     lambda event: on_results_grid_cell_changed(window, event))
            window._results_grid_event_bound = True




def setup_results_grid_column_properties(window):
    """Set up column read-only properties for results grid"""
    if not hasattr(window, 'results_grid') or not window.results_grid:
        return

    num_rows = window.results_grid.GetNumberRows()

    # Define read-only columns: Position, FWHM, Area, Atomic%, ECF, Instrument, Corr. Area, Weight %
    read_only_columns = [1, 2, 3, 5, 6, 10, 11, 13, 29,30]

    # Define editable columns: Label, RSF, TXFN
    editable_columns = [0, 8, 9]

    for row in range(num_rows):
        for col in read_only_columns:
            if col < window.results_grid.GetNumberCols():
                window.results_grid.SetReadOnly(row, col, True)

        for col in editable_columns:
            if col < window.results_grid.GetNumberCols():
                window.results_grid.SetReadOnly(row, col, False)


def setup_grid_editability(window):
    """Set which columns are editable vs read-only"""
    num_rows = window.results_grid.GetNumberRows()
    if num_rows == 0:
        return

    # All columns except 0 (Label), 8 (RSF), 9 (TXFN) are read-only
    all_columns = range(window.results_grid.GetNumberCols())
    editable_columns = [0, 8, 9]  # Label, RSF, TXFN

    for row in range(num_rows):
        for col in all_columns:
            if col in editable_columns:
                window.results_grid.SetReadOnly(row, col, False)
            else:
                window.results_grid.SetReadOnly(row, col, True)

def on_results_grid_cell_changed(window, event):
    """Handle cell changes in results grid"""
    row = event.GetRow()
    col = event.GetCol()

    # Only process changes for editable columns
    if col == 0:  # Label column
        handle_label_change(window, row)
    elif col == 8:  # RSF column
        handle_rsf_change(window, row)
    elif col == 9:  # TXFN column
        handle_txfn_change(window, row)

    event.Skip()



def handle_rsf_change(window, row):
    """Handle RSF value changes and recalculate dependent values"""
    try:
        new_rsf = float(window.results_grid.GetCellValue(row, 8))
        window.results_grid.SetCellValue(row, 8, f"{new_rsf:.2f}")

        # Update data structure
        sheet_name = window.sheet_combobox.GetValue()
        row_number = 0
        import re
        match = re.search(r'(\d+)$', sheet_name)
        if match:
            row_number = int(match.group(1))

        results_table_key = f'Results Table{row_number}'
        peak_key = f"Peak_{row}"

        if (results_table_key in window.Data and 'Peak' in window.Data[results_table_key] and
                peak_key in window.Data[results_table_key]['Peak']):
            window.Data[results_table_key]['Peak'][peak_key]['RSF'] = new_rsf

        # Recalculate all percentages
        window.update_atomic_percentages()

        # Save state
        from libraries.Save import save_state
        save_state(window)

    except ValueError:
        import wx
        wx.MessageBox("Invalid RSF value. Please enter a valid number.", "Error", wx.OK | wx.ICON_ERROR)


def handle_txfn_change(window, row):
    """Handle TXFN value changes and recalculate dependent values"""
    try:
        new_txfn = float(window.results_grid.GetCellValue(row, 9))

        # Format TXFN to 2 decimal places in the grid
        window.results_grid.SetCellValue(row, 9, f"{new_txfn:.2f}")

        # Update data structure
        sheet_name = window.sheet_combobox.GetValue()
        row_number = 0
        import re
        match = re.search(r'(\d+)$', sheet_name)
        if match:
            row_number = int(match.group(1))

        results_table_key = f'Results Table{row_number}'
        peak_key = f"Peak_{row}"

        if (results_table_key in window.Data and 'Peak' in window.Data[results_table_key] and
                peak_key in window.Data[results_table_key]['Peak']):
            window.Data[results_table_key]['Peak'][peak_key]['TXFN'] = new_txfn

        # Recalculate all percentages
        window.update_atomic_percentages()

        # Save state
        from libraries.Save import save_state
        save_state(window)

    except ValueError:
        import wx
        wx.MessageBox("Invalid TXFN value. Please enter a valid number.", "Error", wx.OK | wx.ICON_ERROR)


def handle_label_change_OLD(window, row):
    """Handle label/name changes"""
    new_label = window.results_grid.GetCellValue(row, 0)

    # Update data structure
    sheet_name = window.sheet_combobox.GetValue()
    row_number = 0
    import re
    match = re.search(r'(\d+)$', sheet_name)
    if match:
        row_number = int(match.group(1))

    results_table_key = f'Results Table{row_number}'
    peak_key = f"Peak_{row}"

    if (results_table_key in window.Data and 'Peak' in window.Data[results_table_key] and
            peak_key in window.Data[results_table_key]['Peak']):
        window.Data[results_table_key]['Peak'][peak_key]['Name'] = new_label

    # Save state
    from libraries.Save import save_state
    save_state(window)


def handle_label_change(window, row):
    """Handle label/name changes"""
    new_label = window.results_grid.GetCellValue(row, 0)

    # Calculate and display new mass for reference
    from libraries.Area_Calculation import extract_element_symbol, ATOMIC_MASSES
    element_symbol = extract_element_symbol(new_label)
    atomic_mass = ATOMIC_MASSES.get(element_symbol, 12.01)

    # Update mass display in grid
    window.results_grid.SetCellValue(row, 30, f"{atomic_mass:.2f}")

    # Update data structure
    sheet_name = window.sheet_combobox.GetValue()
    row_number = 0
    import re
    match = re.search(r'(\d+)$', sheet_name)
    if match:
        row_number = int(match.group(1))

    results_table_key = f'Results Table{row_number}'
    peak_key = f"Peak_{row}"

    if (results_table_key in window.Data and 'Peak' in window.Data[results_table_key] and
            peak_key in window.Data[results_table_key]['Peak']):
        window.Data[results_table_key]['Peak'][peak_key]['Name'] = new_label

    # Recalculate weight percentages (this will use ATOMIC_MASSES.get() internally)
    window.update_atomic_percentages()

    # Save state
    from libraries.Save import save_state
    save_state(window)

class CustomCheckboxRenderer(wx.grid.GridCellRenderer):
    def __init__(self):
        wx.grid.GridCellRenderer.__init__(self)

    def Draw(self, grid, attr, dc, rect, row, col, isSelected):
        # Clean the background
        dc.SetBackgroundMode(wx.SOLID)
        dc.SetBrush(wx.Brush(wx.WHITE, wx.SOLID))
        dc.SetPen(wx.TRANSPARENT_PEN)
        dc.DrawRectangle(rect)

        # Get and sanitize the value - ensure it's exactly '0' or '1'
        value = grid.GetCellValue(row, col)
        value = '1' if value == '1' else '0'

        # If the value wasn't valid, fix it in the grid
        if grid.GetCellValue(row, col) != value:
            grid.SetCellValue(row, col, value)

        checkboxState = wx.CONTROL_CHECKED if value == '1' else 0

        # Draw the checkbox
        checkSize = 16
        x = rect.x + (rect.width - checkSize) // 2
        y = rect.y + (rect.height - checkSize) // 2
        checkRect = wx.Rect(x, y, checkSize, checkSize)
        wx.RendererNative.Get().DrawCheckBox(grid, dc, checkRect, checkboxState)

    def GetBestSize(self, grid, attr, dc, row, col):
        return wx.Size(24, 24)

    def Clone(self):
        return CustomCheckboxRenderer()


# Add to Grid_Operations.py
class CustomCheckboxEditor(wx.grid.GridCellEditor):
    def __init__(self):
        wx.grid.GridCellEditor.__init__(self)
        self.checkbox = None

    def Create(self, parent, id, evtHandler):
        self.checkbox = wx.CheckBox(parent, id)
        self.SetControl(self.checkbox)
        if evtHandler:
            self.checkbox.Bind(wx.EVT_CHECKBOX, evtHandler)

    def SetSize(self, rect):
        # Center the checkbox in the cell
        size = self.checkbox.GetBestSize()
        x = rect.x + (rect.width - size.width) // 2
        y = rect.y + (rect.height - size.height) // 2
        self.checkbox.SetSize(wx.Rect(x, y, size.width, size.height))

    def BeginEdit(self, row, col, grid):
        self.startValue = grid.GetCellValue(row, col)
        self.checkbox.SetValue(self.startValue == '1')
        self.checkbox.SetFocus()

    def EndEdit(self, row, col, grid, oldVal=None):
        val = '1' if self.checkbox.GetValue() else '0'
        if val != self.startValue:
            return val
        return None

    def ApplyEdit(self, row, col, grid):
        val = '1' if self.checkbox.GetValue() else '0'
        grid.SetCellValue(row, col, val)

    def Reset(self):
        self.checkbox.SetValue(self.startValue == '1')

    def Clone(self):
        return CustomCheckboxEditor()


class CheckboxRenderer(wx.grid.GridCellRenderer):
    def __init__(self):
        wx.grid.GridCellRenderer.__init__(self)

    def Draw(self, grid, attr, dc, rect, row, col, isSelected):
        # Clear background
        dc.SetBackgroundMode(wx.SOLID)
        dc.SetBrush(wx.Brush(wx.WHITE, wx.SOLID))
        dc.SetPen(wx.TRANSPARENT_PEN)
        dc.DrawRectangle(rect)

        # Get the value (should be '0' or '1')
        value = grid.GetCellValue(row, col)
        # Force valid value
        checked = (value == '1')

        # Calculate checkbox size and position
        checkSize = min(rect.width, rect.height) - 4
        x = rect.x + (rect.width - checkSize) // 2
        y = rect.y + (rect.height - checkSize) // 2

        # Draw checkbox using native renderer if possible
        if hasattr(wx, 'RendererNative'):
            checkBoxFlags = 0
            if checked:
                checkBoxFlags |= wx.CONTROL_CHECKED
            wx.RendererNative.Get().DrawCheckBox(
                grid, dc, wx.Rect(x, y, checkSize, checkSize), checkBoxFlags)
        else:
            # Fallback to manual drawing
            dc.SetPen(wx.BLACK_PEN)
            dc.SetBrush(wx.WHITE_BRUSH)
            dc.DrawRectangle(x, y, checkSize, checkSize)

            # Draw check if checked
            if checked:
                dc.SetPen(wx.Pen(wx.BLACK, 2))
                dc.DrawLine(x + 3, y + checkSize // 2, x + checkSize // 2, y + checkSize - 3)
                dc.DrawLine(x + checkSize // 2, y + checkSize - 3, x + checkSize - 3, y + 3)

    def GetBestSize(self, grid, attr, dc, row, col):
        return wx.Size(20, 20)

    def Clone(self):
        return CheckboxRenderer()