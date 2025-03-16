import wx
import wx.grid
import os
import re
import numpy as np
from copy import deepcopy


class FileManagerWindow(wx.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, title="File Manager", size=(500, 300),
                         style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP, *args, **kwargs)

        self.parent = parent

        # Create main panel
        self.panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Create toolbar
        self.toolbar = wx.ToolBar(self.panel, style=wx.TB_HORIZONTAL | wx.TB_FLAT | wx.TB_NODIVIDER)
        self.create_toolbar()
        self.toolbar.Realize()
        main_sizer.Add(self.toolbar, 0, wx.EXPAND)

        # Create grid
        self.grid = wx.grid.Grid(self.panel)
        self.init_grid()
        main_sizer.Add(self.grid, 1, wx.EXPAND | wx.ALL, 5)

        self.panel.SetSizer(main_sizer)

        # Position window relative to main window
        main_pos = parent.GetPosition()
        main_size = parent.GetSize()
        file_manager_size = self.GetSize()
        pos_x = main_pos.x + (main_size.width - file_manager_size.width) // 2
        pos_y = main_pos.y + (main_size.height - file_manager_size.height) // 2
        self.SetPosition((pos_x, pos_y))

        # Populate the grid with core levels
        self.populate_grid()

        # Bind events
        self.grid.Bind(wx.grid.EVT_GRID_CELL_CHANGING, self.on_cell_changing)
        self.grid.Bind(wx.EVT_KEY_DOWN, self.on_key_down)  # Bind to grid, not frame
        self.Bind(wx.EVT_KEY_DOWN, self.on_key_down)  # Keep frame binding too

        # Apply consistent fonts from parent
        from libraries.ConfigFile import set_consistent_fonts
        set_consistent_fonts(self)

    def create_toolbar(self):
        """Create toolbar with buttons for core level management"""
        # Get icon path
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Icons")

        # Plot button
        plot_icon = os.path.join(icon_path, "Plot-25.png")
        if os.path.exists(plot_icon):
            plot_bmp = wx.Bitmap(plot_icon)
        else:
            plot_bmp = wx.ArtProvider.GetBitmap(wx.ART_FIND, wx.ART_TOOLBAR)
        plot_tool = self.toolbar.AddTool(wx.ID_ANY, "Plot Selected", plot_bmp, "Plot selected core level(s)")
        self.Bind(wx.EVT_TOOL, self.on_plot_selected, plot_tool)

        # Sum button
        sum_icon = os.path.join(icon_path, "SUM-25.jpg")
        sum_bmp = wx.Bitmap(sum_icon)
        sum_tool = self.toolbar.AddTool(wx.ID_ANY, "Sum Selected", sum_bmp, "Sum selected core levels")
        self.Bind(wx.EVT_TOOL, self.on_sum_selected, sum_tool)

        self.toolbar.AddSeparator()

        # Copy button with custom icon
        copy_icon = os.path.join(icon_path, "copy-25.png")
        if os.path.exists(copy_icon):
            copy_bmp = wx.Bitmap(copy_icon)
        else:
            copy_bmp = wx.ArtProvider.GetBitmap(wx.ART_COPY, wx.ART_TOOLBAR)
        copy_tool = self.toolbar.AddTool(wx.ID_ANY, "Copy Core Level", copy_bmp, "Copy selected core level")
        self.Bind(wx.EVT_TOOL, self.on_copy, copy_tool)

        # Paste button (using default icon since not specified)
        paste_bmp = wx.ArtProvider.GetBitmap(wx.ART_PASTE, wx.ART_TOOLBAR)
        paste_tool = self.toolbar.AddTool(wx.ID_ANY, "Paste Core Level", paste_bmp, "Paste core level")
        self.Bind(wx.EVT_TOOL, self.on_paste, paste_tool)

        self.toolbar.AddSeparator()

        # Rename button with custom icon
        rename_icon = os.path.join(icon_path, "rename-25.png")
        if os.path.exists(rename_icon):
            rename_bmp = wx.Bitmap(rename_icon)
        else:
            rename_bmp = wx.ArtProvider.GetBitmap(wx.ART_INFORMATION, wx.ART_TOOLBAR)
        rename_tool = self.toolbar.AddTool(wx.ID_ANY, "Rename Core Level", rename_bmp, "Rename selected core level")
        self.Bind(wx.EVT_TOOL, self.on_rename, rename_tool)

        # Delete button with custom icon
        delete_icon = os.path.join(icon_path, "delete-25.png")
        if os.path.exists(delete_icon):
            delete_bmp = wx.Bitmap(delete_icon)
        else:
            delete_bmp = wx.ArtProvider.GetBitmap(wx.ART_DELETE, wx.ART_TOOLBAR)
        delete_tool = self.toolbar.AddTool(wx.ID_ANY, "Delete Core Level", delete_bmp, "Delete selected core level")
        self.Bind(wx.EVT_TOOL, self.on_delete, delete_tool)

        self.toolbar.AddSeparator()

        # Toggle size button - using a different art ID

        size_icon = os.path.join(icon_path, "Minimize-25.png")
        size_bmp = wx.Bitmap(size_icon)
        self.size_tool = self.toolbar.AddTool(wx.ID_ANY, "Toggle Size", size_bmp, "Toggle window height")
        self.Bind(wx.EVT_TOOL, self.on_toggle_size, self.size_tool)

        # Store the original/expanded size
        self.is_collapsed = False
        self.expanded_height = 300  # Default expanded height
        self.collapsed_height = 70  # Collapsed height

    def on_toggle_size(self, event):
        """Toggle between normal and small window size"""
        current_size = self.GetSize()

        if self.is_collapsed:
            # Expand
            self.SetSize((current_size.width, self.expanded_height))
            self.is_collapsed = False
        else:
            # Collapse
            self.expanded_height = current_size.height  # Save current height
            self.SetSize((current_size.width, self.collapsed_height))
            self.is_collapsed = True

    def init_grid(self):
        """Initialize the grid with rows and columns"""
        # Get unique core level names from parent
        self.core_levels = self.get_unique_core_levels()
        num_levels = len(self.core_levels)

        # Determine number of rows based on existing data
        max_row_index = self.get_max_core_level_row_index()
        num_rows = max(10, max_row_index + 1)  # Ensure at least 10 rows

        # Create grid
        self.grid.CreateGrid(num_rows, num_levels)

        # Enable cell editing for renaming
        self.grid.EnableEditing(True)

        # Set column labels (core level names)
        self.grid.SetRowLabelSize(30)
        for i, level in enumerate(self.core_levels):
            self.grid.SetColLabelValue(i, level)

        # Set row labels (0, 1, 2, etc.)
        for i in range(num_rows):
            self.grid.SetRowLabelValue(i, str(i))

        # Set column width and row height
        default_col_width = 50
        default_row_height = 20

        # Set column sizes and row heights
        for i in range(num_levels):
            self.grid.SetColSize(i, default_col_width)
        for i in range(num_rows):
            self.grid.SetRowSize(i, default_row_height)

        # Set cell alignment
        for row in range(num_rows):
            for col in range(num_levels):
                self.grid.SetCellAlignment(row, col, wx.ALIGN_CENTER, wx.ALIGN_CENTER)

    def get_unique_core_levels(self):
        """Get list of unique core level names from parent data"""
        unique_levels = set()

        if hasattr(self.parent, 'Data') and 'Core levels' in self.parent.Data:
            for sheet_name in self.parent.Data['Core levels'].keys():
                base_name = self.extract_base_name(sheet_name)
                if base_name:
                    unique_levels.add(base_name)

        return sorted(list(unique_levels))

    def extract_base_name(self, sheet_name):
        """Extract base core level name without any trailing numbers"""
        match = re.match(r'([A-Za-z0-9]+?)(\d*)$', sheet_name)
        if match:
            return match.group(1)
        return sheet_name

    def get_max_core_level_row_index(self):
        """Get the maximum row index based on core level naming"""
        max_index = 0

        if hasattr(self.parent, 'Data') and 'Core levels' in self.parent.Data:
            for sheet_name in self.parent.Data['Core levels'].keys():
                match = re.match(r'[A-Za-z0-9]+?(\d*)$', sheet_name)
                if match:
                    index_str = match.group(1)
                    index = int(index_str) if index_str else 0
                    max_index = max(max_index, index)

        return max_index

    def populate_grid(self):
        """Populate the grid with core levels from the data"""
        if not hasattr(self.parent, 'Data') or 'Core levels' not in self.parent.Data:
            return

        # Clear grid first
        for row in range(self.grid.GetNumberRows()):
            for col in range(self.grid.GetNumberCols()):
                self.grid.SetCellValue(row, col, "")
                self.grid.SetCellBackgroundColour(row, col, wx.WHITE)

        # Map of all core levels by type and index
        core_level_map = {}

        # First, categorize all core levels
        for sheet_name in self.parent.Data['Core levels'].keys():
            match = re.match(r'([A-Za-z0-9]+?)(\d*)$', sheet_name)
            if match:
                base_name = match.group(1)
                index_str = match.group(2)
                index = int(index_str) if index_str else 0

                if base_name not in core_level_map:
                    core_level_map[base_name] = {}
                core_level_map[base_name][index] = sheet_name

        # Now populate the grid
        for col, base_name in enumerate(self.core_levels):
            if base_name in core_level_map:
                for index, sheet_name in core_level_map[base_name].items():
                    # Ensure we have enough rows
                    if index >= self.grid.GetNumberRows():
                        self.grid.AppendRows(index - self.grid.GetNumberRows() + 1)
                        # Set row label for new rows
                        for row in range(self.grid.GetNumberRows()):
                            if not self.grid.GetRowLabelValue(row):
                                self.grid.SetRowLabelValue(row, str(row))

                    # Set the cell value and color
                    self.grid.SetCellValue(index, col, sheet_name)
                    self.grid.SetCellBackgroundColour(index, col, wx.Colour(200, 245, 228))

        # Force refresh
        self.grid.ForceRefresh()

    def on_key_down(self, event):
        """Handle key press events"""
        key_code = event.GetKeyCode()

        if key_code == wx.WXK_F2:
            # Call the plot function directly and don't skip the event
            self.on_plot_selected(None)
            return  # Don't skip the event
        else:
            event.Skip()

    def on_plot_selected(self, event):
        """Plot the currently selected core level(s)"""
        sheet_names = self.get_selected_sheet_names()

        if sheet_names:
            if len(sheet_names) == 1:
                # Single sheet - normal plot
                self.parent.sheet_combobox.SetValue(sheet_names[0])
                from libraries.Sheet_Operations import on_sheet_selected
                on_sheet_selected(self.parent, sheet_names[0])
            else:
                # Multiple sheets - overlay plot
                self.plot_multiple_sheets(sheet_names)

            self.Raise()  # Bring the file manager window to the front

    def on_sum_selected(self, event):
        """Sum/average the Y values of selected cells in the same column"""
        # Get selected cells grouped by column
        selected_by_column = {}

        # Check for selected blocks first
        blocks = self.grid.GetSelectedBlocks()
        for block in blocks:  # Directly iterate over blocks instead of using GetCount()
            top = block.GetTopRow()
            bottom = block.GetBottomRow()
            left = block.GetLeftCol()
            right = block.GetRightCol()

            for col in range(left, right + 1):
                if col not in selected_by_column:
                    selected_by_column[col] = []

                for row in range(top, bottom + 1):
                    cell_value = self.grid.GetCellValue(row, col)
                    if cell_value and cell_value in self.parent.Data['Core levels']:
                        selected_by_column[col].append(cell_value)

        # Add individually selected cells
        selected_cells = self.grid.GetSelectedCells()
        for cell in selected_cells:
            row, col = cell.GetRow(), cell.GetCol()

            if col not in selected_by_column:
                selected_by_column[col] = []

            cell_value = self.grid.GetCellValue(row, col)
            if cell_value and cell_value in self.parent.Data['Core levels']:
                selected_by_column[col].append(cell_value)

        # Process each column separately
        for col, sheet_names in selected_by_column.items():
            if len(sheet_names) > 1:
                self.create_summed_spectrum(sheet_names, self.grid.GetColLabelValue(col))

    def create_summed_spectrum(self, sheet_names, base_name):
        """Create a new spectrum that is the average of the selected spectra"""
        if not sheet_names:
            return

        # Create a new sheet name
        new_sheet_name = f"{base_name}1000"
        counter = 1
        while new_sheet_name in self.parent.Data['Core levels']:
            new_sheet_name = f"{base_name}{1000 + counter}"
            counter += 1

        # Get X values from the first sheet (assuming they're similar)
        first_sheet = sheet_names[0]
        x_values = self.parent.Data['Core levels'][first_sheet]['B.E.']

        # Sum Y values from all sheets
        summed_y = np.zeros_like(x_values, dtype=float)
        for sheet_name in sheet_names:
            if sheet_name in self.parent.Data['Core levels']:
                current_y = self.parent.Data['Core levels'][sheet_name]['Raw Data']
                summed_y += np.array(current_y)

        # Average the Y values
        avg_y = summed_y / len(sheet_names)

        # Create a new entry in the parent data
        from libraries.ConfigFile import Init_Measurement_Data
        if 'Core levels' not in self.parent.Data:
            self.parent.Data['Core levels'] = {}
            self.parent.Data['Number of Core levels'] = 0

        self.parent.Data['Core levels'][new_sheet_name] = {
            'Name': new_sheet_name,
            'B.E.': x_values,
            'Raw Data': avg_y.tolist(),
            'Background': {
                'Bkg Type': 'Linear',
                'Bkg Low': min(x_values),
                'Bkg High': max(x_values),
                'Bkg Offset Low': 0,
                'Bkg Offset High': 0,
                'Bkg Y': avg_y.tolist()  # Initially use raw data as background
            }
        }
        self.parent.Data['Number of Core levels'] += 1

        # Update the Excel file
        import pandas as pd
        df = pd.DataFrame({
            'BE': x_values,
            'Raw Data': avg_y.tolist(),
            'Background': avg_y.tolist(),
            'Transmission': [1.0] * len(x_values)
        })

        with pd.ExcelWriter(self.parent.Data['FilePath'], engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=new_sheet_name, index=False)

        # Update combobox in parent
        self.parent.sheet_combobox.Append(new_sheet_name)

        # Plot the new sheet
        self.parent.sheet_combobox.SetValue(new_sheet_name)
        from libraries.Sheet_Operations import on_sheet_selected
        on_sheet_selected(self.parent, new_sheet_name)

        # Update the grid
        self.populate_grid()

        # Notify the user
        wx.MessageBox(f"Created summed spectrum: {new_sheet_name}", "Sum Complete", wx.OK | wx.ICON_INFORMATION)

    def get_selected_sheet_names(self):
        """Get names of all currently selected sheets in the grid"""
        sheet_names = []

        # Check for selected blocks first
        blocks = self.grid.GetSelectedBlocks()
        # For newer wxPython versions, blocks is an iterable object
        for block in blocks:
            top = block.GetTopRow()
            bottom = block.GetBottomRow()
            left = block.GetLeftCol()
            right = block.GetRightCol()

            for row in range(top, bottom + 1):
                for col in range(left, right + 1):
                    cell_value = self.grid.GetCellValue(row, col)
                    if cell_value and cell_value in self.parent.Data['Core levels']:
                        sheet_names.append(cell_value)

        # Add individually selected cells
        selected_cells = self.grid.GetSelectedCells()
        for cell in selected_cells:
            row, col = cell.GetRow(), cell.GetCol()
            cell_value = self.grid.GetCellValue(row, col)
            if cell_value and cell_value in self.parent.Data['Core levels']:
                sheet_names.append(cell_value)

        # If no selection, use the current cell
        if not sheet_names:
            row = self.grid.GetGridCursorRow()
            col = self.grid.GetGridCursorCol()
            if row >= 0 and col >= 0:  # Make sure we have valid coordinates
                cell_value = self.grid.GetCellValue(row, col)
                if cell_value and cell_value in self.parent.Data['Core levels']:
                    sheet_names.append(cell_value)

        return sheet_names

    def plot_multiple_sheets(self, sheet_names):
        """Plot multiple core levels together on the same graph"""
        if not sheet_names:
            return

        # Set the first sheet as the active one in the parent window
        self.parent.sheet_combobox.SetValue(sheet_names[0])
        from libraries.Sheet_Operations import on_sheet_selected
        on_sheet_selected(self.parent, sheet_names[0])

        # Clear the plot
        self.parent.ax.clear()

        # Track min/max x values
        x_min = float('inf')
        x_max = float('-inf')

        # Plot each selected sheet
        for i, sheet_name in enumerate(sheet_names):
            if sheet_name in self.parent.Data['Core levels']:
                core_level = self.parent.Data['Core levels'][sheet_name]
                x_values = core_level['B.E.']
                y_values = core_level['Raw Data']

                # Update min/max x values
                x_min = min(x_min, min(x_values))
                x_max = max(x_max, max(x_values))

                # Use a different color for each plot
                color = self.parent.peak_colors[i % len(self.parent.peak_colors)]

                # Plot the data
                self.parent.ax.plot(x_values, y_values, label=sheet_name, color=color)

        # Set labels and formatting
        self.parent.ax.set_xlabel("Binding Energy (eV)")
        self.parent.ax.set_ylabel("Intensity (CPS)")
        self.parent.ax.legend()

        # Set x-axis limits to min/max values from all datasets
        self.parent.ax.set_xlim(x_max, x_min)  # Reversed for XPS

        # Update the plot
        self.parent.canvas.draw_idle()

    def on_cell_changing(self, event):
        """Handle cell edit event for renaming and repositioning core levels"""
        row = event.GetRow()
        col = event.GetCol()
        old_value = self.grid.GetCellValue(row, col)
        new_value = event.GetString()

        # Only process if there's a real change and the cell isn't empty
        if old_value and old_value != new_value and new_value.strip():
            # Check if old name exists in parent data
            if old_value in self.parent.Data['Core levels']:
                # Check if new name already exists
                if new_value in self.parent.Data['Core levels']:
                    wx.MessageBox(f"A core level named '{new_value}' already exists.",
                                  "Duplicate Name", wx.OK | wx.ICON_ERROR)
                    event.Veto()
                    return

                # Set the sheet in parent before renaming
                self.parent.sheet_combobox.SetValue(old_value)
                from libraries.Sheet_Operations import on_sheet_selected
                on_sheet_selected(self.parent, old_value)

                # Perform the rename
                from libraries.Utilities import rename_sheet
                rename_sheet(self.parent, new_value)

                # Always refresh the grid to properly position items
                wx.CallAfter(self.populate_grid)

    def on_copy(self, event):
        """Copy the selected core level"""
        sheet_names = self.get_selected_sheet_names()

        if sheet_names:
            sheet_name = sheet_names[0]  # Use the first selected sheet

            # Set the sheet in parent before copying
            self.parent.sheet_combobox.SetValue(sheet_name)
            from libraries.Sheet_Operations import on_sheet_selected
            on_sheet_selected(self.parent, sheet_name)

            from libraries.Save import copy_core_level
            copy_core_level(self.parent)

            wx.MessageBox(f"Core level '{sheet_name}' copied", "Copy Successful", wx.OK | wx.ICON_INFORMATION)

    def on_paste(self, event):
        """Paste the core level to the selected cell"""
        from libraries.Save import paste_core_level
        paste_core_level(self.parent)

        # Reload the grid after pasting
        self.populate_grid()

    def on_rename(self, event):
        """Rename the selected core level"""
        sheet_names = self.get_selected_sheet_names()

        if sheet_names:
            sheet_name = sheet_names[0]  # Use the first selected sheet

            # Set the sheet in parent before renaming
            self.parent.sheet_combobox.SetValue(sheet_name)
            from libraries.Sheet_Operations import on_sheet_selected
            on_sheet_selected(self.parent, sheet_name)

            dlg = wx.TextEntryDialog(self, f"Enter new name for {sheet_name}:", "Rename Core Level")
            if dlg.ShowModal() == wx.ID_OK:
                new_name = dlg.GetValue()
                if new_name and new_name != sheet_name:
                    from libraries.Utilities import rename_sheet
                    rename_sheet(self.parent, new_name)

                    # Reload the grid after renaming
                    self.populate_grid()
            dlg.Destroy()

    def on_delete(self, event):
        """Delete the selected core level"""
        sheet_names = self.get_selected_sheet_names()

        if sheet_names:
            sheet_name = sheet_names[0]  # Use the first selected sheet

            # Set the sheet in parent before deleting
            self.parent.sheet_combobox.SetValue(sheet_name)
            from libraries.Sheet_Operations import on_sheet_selected
            on_sheet_selected(self.parent, sheet_name)

            if wx.MessageBox(f"Are you sure you want to delete {sheet_name}?", "Confirm Delete",
                             wx.YES_NO | wx.ICON_QUESTION) == wx.YES:
                from libraries.Utilities import on_delete_sheet
                on_delete_sheet(self.parent, None)

                # Reload the grid after deleting
                self.populate_grid()