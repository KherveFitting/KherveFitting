# libraries/grid_operations.py

import wx

def populate_results_grid_OLD(window):
    if 'Results' in window.Data and 'Peak' in window.Data['Results']:
        results = window.Data['Results']['Peak']

        # Clear existing data in the grid
        window.results_grid.ClearGrid()

        # Resize the grid if necessary
        num_rows = len(results)
        num_cols = 26  # Based on your Export.py structure
        if window.results_grid.GetNumberRows() < num_rows:
            window.results_grid.AppendRows(num_rows - window.results_grid.GetNumberRows())
        if window.results_grid.GetNumberCols() < num_cols:
            window.results_grid.AppendCols(num_cols - window.results_grid.GetNumberCols())

        # Populate the grid
        for row, (peak_label, peak_data) in enumerate(results.items()):
            window.results_grid.SetCellValue(row, 0, peak_data['Name'])
            window.results_grid.SetCellValue(row, 1, str(peak_data['Position']))
            window.results_grid.SetCellValue(row, 2, str(peak_data['Height']))
            window.results_grid.SetCellValue(row, 3, str(peak_data['FWHM']))
            window.results_grid.SetCellValue(row, 4, str(peak_data['L/G']))
            window.results_grid.SetCellValue(row, 5, f"{peak_data['Area']:.2f}")
            window.results_grid.SetCellValue(row, 6, f"{peak_data['at. %']:.2f}")

            # Set up checkbox and its state
            checkbox_state = peak_data.get('Checkbox', '0')
            window.results_grid.SetCellRenderer(row, 7, wx.grid.GridCellBoolRenderer())
            window.results_grid.SetCellValue(row, 7, checkbox_state)


            window.results_grid.SetCellValue(row, 8, f"{peak_data['RSF']:.2f}")
            window.results_grid.SetCellValue(row, 9, f"{peak_data['TXFN']:.2f}")
            window.results_grid.SetCellValue(row, 10, str(peak_data['ECF']))
            window.results_grid.SetCellValue(row, 11, str(peak_data.get('Instrument', 'Al1486')))
            window.results_grid.SetCellValue(row, 12, peak_data['Fitting Model'])
            window.results_grid.SetCellValue(row, 13, f"{peak_data['Rel. Area']:.2f}")
            window.results_grid.SetCellValue(row, 14, str(peak_data['Sigma']))
            window.results_grid.SetCellValue(row, 15, str(peak_data['Gamma']))
            window.results_grid.SetCellValue(row, 16, peak_data.get('Bkg Type', ''))  # Bkg Type
            window.results_grid.SetCellValue(row, 17, str(peak_data['Bkg Low']))
            window.results_grid.SetCellValue(row, 18, str(peak_data['Bkg High']))
            window.results_grid.SetCellValue(row, 19, str(peak_data.get('Bkg Offset Low', '')))
            window.results_grid.SetCellValue(row, 20, str(peak_data.get('Bkg Offset High', '')))
            window.results_grid.SetCellValue(row, 21, peak_data['Sheetname'])
            window.results_grid.SetCellValue(row, 22, peak_data['Pos. Constraint'])
            window.results_grid.SetCellValue(row, 23, peak_data['Height Constraint'])
            window.results_grid.SetCellValue(row, 24, peak_data['FWHM Constraint'])
            window.results_grid.SetCellValue(row, 25, peak_data['L/G Constraint'])
            window.results_grid.SetCellValue(row, 26, peak_data['Area Constraint'])
            window.results_grid.SetCellValue(row, 27, peak_data['Sigma Constraint'])
            window.results_grid.SetCellValue(row, 28, peak_data['Gamma Constraint'])

        # Bind events
        # window.results_grid.Bind(wx.grid.EVT_GRID_CELL_LEFT_CLICK, window.on_checkbox_update)
        window.results_grid.Bind(wx.grid.EVT_GRID_CELL_CHANGED, window.on_cell_changed)

        # Refresh the grid
        window.results_grid.ForceRefresh()
        window.results_grid.Refresh()

    else:
        print("No results data found in window.Data")

    # Calculate atomic percentages for checked elements
    window.update_atomic_percentages()


# In Grid_Operations.py
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
        num_cols = 26  # Based on your Export.py structure
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

            if window.results_grid.GetNumberCols() > 26:
                window.results_grid.SetCellValue(row, 26, peak_data.get('Area Constraint', ''))
                window.results_grid.SetCellValue(row, 27, peak_data.get('Sigma Constraint', ''))
                window.results_grid.SetCellValue(row, 28, peak_data.get('Gamma Constraint', ''))

            # # Use custom renderer and editor for checkboxes
            checkbox_state = peak_data.get('Checkbox', '0')
            window.results_grid.SetCellValue(row, 7, checkbox_state)

            # Use custom checkbox renderer
            from libraries.Grid_Operations import CheckboxRenderer
            window.results_grid.SetCellRenderer(row, 7, CheckboxRenderer())
            window.results_grid.SetReadOnly(row, 7)



        # Force a refresh to ensure renderer is applied
        window.results_grid.ForceRefresh()

        # # Calculate atomic percentages for checked elements
        # window.update_atomic_percentages()

        # # Bind events
        # window.results_grid.Bind(wx.grid.EVT_GRID_CELL_CHANGED, window.on_cell_changed)
        # window.results_grid.Bind(wx.grid.EVT_GRID_CELL_LEFT_CLICK, window.on_checkbox_update)

    else:
        # Just leave the grid empty if no results for this row
        print(f"No results data found for {results_table_key}")

    # # Calculate atomic percentages for checked elements
    # window.update_atomic_percentages()


# Add to Grid_Operations.py
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