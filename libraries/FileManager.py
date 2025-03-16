import wx
import wx.grid
import os
import re
import numpy as np
from matplotlib.ticker import ScalarFormatter
from copy import deepcopy


class FileManagerWindow(wx.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, title="File Manager", size=(500, 300),
                         style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP, *args, **kwargs)

        self.parent = parent

        # Create main panel - use only ONE main panel
        self.panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)  # Changed to HORIZONTAL for the side-by-side layout

        # Create vertical toolbar on left
        self.v_toolbar_panel = wx.Panel(self.panel, size=(40, -1))
        self.v_toolbar_panel.SetBackgroundColour(wx.Colour(240, 240, 240))
        v_toolbar_sizer = wx.BoxSizer(wx.VERTICAL)

        # Add normalization checkboxes
        self.norm_check = wx.CheckBox(self.v_toolbar_panel, label="Norm")
        self.auto_check = wx.CheckBox(self.v_toolbar_panel, label="Auto")
        self.auto_check.SetValue(True)  # Auto is checked by default

        v_toolbar_sizer.Add(self.norm_check, 0, wx.ALL, 5)
        v_toolbar_sizer.Add(self.auto_check, 0, wx.ALL, 5)
        self.v_toolbar_panel.SetSizer(v_toolbar_sizer)

        # Right side panel with content
        self.right_panel = wx.Panel(self.panel)
        right_sizer = wx.BoxSizer(wx.VERTICAL)

        # Create toolbar
        self.toolbar = wx.ToolBar(self.right_panel, style=wx.TB_HORIZONTAL | wx.TB_FLAT | wx.TB_NODIVIDER)
        self.toolbar.SetToolBitmapSize(wx.Size(25, 25))
        self.create_toolbar()
        self.toolbar.Realize()
        right_sizer.Add(self.toolbar, 0, wx.EXPAND)

        # Create grid
        self.grid = wx.grid.Grid(self.right_panel)
        self.init_grid()
        right_sizer.Add(self.grid, 1, wx.EXPAND | wx.ALL, 5)

        self.right_panel.SetSizer(right_sizer)

        # Add both panels to main sizer
        main_sizer.Add(self.v_toolbar_panel, 0, wx.EXPAND)
        main_sizer.Add(self.right_panel, 1, wx.EXPAND)

        self.panel.SetSizer(main_sizer)

        # Position window relative to main window
        main_pos = parent.GetPosition()
        main_size = parent.GetSize()
        file_manager_size = self.GetSize()
        pos_x = main_pos.x + (main_size.width - file_manager_size.width) // 2
        pos_y = main_pos.y + (main_size.height - file_manager_size.height) // 2
        self.SetPosition((pos_x, pos_y))

        # Initialize normalization values
        self.norm_min = 0
        self.norm_max = 1
        self.norm_vlines = [None, None]  # For the normalization cursors
        self.is_dragging_cursor = False

        # Bind events
        self.norm_check.Bind(wx.EVT_CHECKBOX, self.on_norm_changed)
        self.auto_check.Bind(wx.EVT_CHECKBOX, self.on_auto_changed)
        self.parent.canvas.mpl_connect('key_press_event', self.on_key_press)
        self.parent.canvas.mpl_connect('key_release_event', self.on_key_release)

        # Populate the grid with core levels
        self.populate_grid()

        # Bind grid events
        self.grid.Bind(wx.grid.EVT_GRID_CELL_CHANGING, self.on_cell_changing)
        self.grid.Bind(wx.EVT_KEY_DOWN, self.on_key_down)
        self.Bind(wx.EVT_KEY_DOWN, self.on_key_down)

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
        sum_icon = os.path.join(icon_path, "SUM-25.png")
        sum_bmp = wx.Bitmap(sum_icon)
        sum_tool = self.toolbar.AddTool(wx.ID_ANY, "Sum Selected", sum_bmp, "Sum selected core levels")
        self.Bind(wx.EVT_TOOL, self.on_sum_selected, sum_tool)

        # self.toolbar.AddSeparator()

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

        # self.toolbar.AddSeparator()

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

        # self.toolbar.AddSeparator()

        # Toggle size button - using a different art ID
        self.toolbar.AddStretchableSpace()

        pref_icon = os.path.join(icon_path, "settings-25.png")
        pref_bmp = wx.Bitmap(pref_icon)
        pref_tool = self.toolbar.AddTool(wx.ID_ANY, "Preferences", pref_bmp, "Normalization Settings")
        self.Bind(wx.EVT_TOOL, self.on_preferences, pref_tool)

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

            # Apply consistent fonts to grid
            if 'wxMac' in wx.PlatformInfo:
                default_font = 'Helvetica'
                font_size = 11
            elif 'wxGTK' in wx.PlatformInfo:
                default_font = 'DejaVu Sans'
                font_size = 9
            else:
                default_font = 'Calibri'
                font_size = 9

            self.grid.SetDefaultCellFont(
                wx.Font(font_size, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                        wx.FONTWEIGHT_NORMAL, faceName=default_font))
            self.grid.SetLabelFont(
                wx.Font(font_size, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                        wx.FONTWEIGHT_NORMAL, faceName=default_font))

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
            # Call the plot function directly
            self.on_plot_selected(None)
            return  # Don't skip the event
        elif event.ControlDown() and key_code == wx.WXK_F2:
            # CTRL+F2: Add new functionality here
            # Example: Show multiple plots in a new window
            self.on_plot_selected(None)
            return
        elif event.ControlDown() and key_code == ord('2'):
            # CTRL+2: Add new functionality here
            # Example: Toggle between single/multiple plot view
            self.on_plot_selected(None)
            return
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

        # Determine if normalization is needed
        normalize = self.norm_check.GetValue()
        auto_norm = self.auto_check.GetValue()

        # For auto normalization, we need to calculate global min/max
        global_min = float('inf')
        global_max = float('-inf')

        # Check if all sheets are from the same column (core level)
        base_names = set(self.extract_base_name(name) for name in sheet_names)
        same_column = len(base_names) == 1
        column_name = list(base_names)[0] if same_column else None

        if normalize and auto_norm:
            # Get global min/max across all selected datasets
            for sheet_name in sheet_names:
                if sheet_name in self.parent.Data['Core levels']:
                    y_values = self.parent.Data['Core levels'][sheet_name]['Raw Data']
                    global_min = min(global_min, min(y_values))
                    global_max = max(global_max, max(y_values))

        # Plot each selected sheet
        for i, sheet_name in enumerate(sheet_names):
            if sheet_name in self.parent.Data['Core levels']:
                core_level = self.parent.Data['Core levels'][sheet_name]
                x_values = core_level['B.E.']
                y_values = np.array(core_level['Raw Data'])

                # Update min/max x values
                x_min = min(x_min, min(x_values))
                x_max = max(x_max, max(x_values))

                # Apply normalization if enabled
                if normalize:
                    if auto_norm:
                        # Use global min/max from all datasets
                        norm_min = global_min
                        norm_max = global_max
                    else:
                        # Use manually set min/max
                        norm_min = self.norm_min
                        norm_max = self.norm_max

                    # Avoid division by zero
                    if norm_max != norm_min:
                        y_values = (y_values - norm_min) / (norm_max - norm_min)

                # Use a different color for each plot
                color = self.parent.peak_colors[i % len(self.parent.peak_colors)]

                # Plot the data
                if self.parent.energy_scale == 'KE':
                    self.parent.ax.plot(self.parent.photons - x_values, y_values, label=sheet_name, color=color,
                                        linewidth=self.parent.line_width)
                else:
                    self.parent.ax.plot(x_values, y_values, label=sheet_name, color=color,
                                        linewidth=self.parent.line_width)

        # Set labels and formatting
        self.parent.ax.set_xlabel("Binding Energy (eV)")
        self.parent.ax.set_ylabel("Intensity (CPS)")

        # Apply scientific format to Y-axis
        self.parent.ax.yaxis.set_major_formatter(ScalarFormatter(useMathText=True))
        self.parent.ax.ticklabel_format(style='sci', axis='y', scilimits=(0, 0))

        # Set legend on the left
        self.parent.ax.legend(loc='upper left')

        # Set labels and formatting
        self.parent.ax.set_xlabel("Binding Energy (eV)")
        if normalize:
            self.parent.ax.set_ylabel("Normalized Intensity")
        else:
            self.parent.ax.set_ylabel("Intensity (CPS)")
        self.parent.ax.legend()

        # Set x-axis limits to min/max values from all datasets
        self.parent.ax.set_xlim(x_max, x_min)  # Reversed for XPS

        # If all sheets are from the same column, add core level text in top right
        if same_column:
            formatted_name = self.parent.plot_manager.format_sheet_name(column_name)
            sheet_name_text = self.parent.ax.text(
                0.98, 0.98,  # Position (top-right corner)
                formatted_name,
                transform=self.parent.ax.transAxes,
                fontsize=self.parent.core_level_text_size,
                fontfamily=[self.parent.plot_font],
                fontweight='bold',
                verticalalignment='top',
                horizontalalignment='right',
                bbox=dict(facecolor='none', edgecolor='none', alpha=1),
            )
            sheet_name_text.sheet_name_text = True  # Mark this text object

        # Apply text settings from preferences
        self.parent.ax.tick_params(axis='both', labelsize=self.parent.axis_number_size)
        self.parent.ax.xaxis.label.set_size(self.parent.axis_title_size)
        self.parent.ax.yaxis.label.set_size(self.parent.axis_title_size)

        # Update the plot
        self.parent.canvas.draw_idle()

    def replot_with_normalization(self):
        """Replot the current selection with updated normalization settings"""
        sheet_names = self.get_selected_sheet_names()
        if sheet_names:
            self.plot_multiple_sheets(sheet_names)

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

    def on_preferences(self, event):
        dlg = wx.Dialog(self, title="Normalization Settings", size=(300, 200))
        panel = wx.Panel(dlg)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Min value input
        min_sizer = wx.BoxSizer(wx.HORIZONTAL)
        min_sizer.Add(wx.StaticText(panel, label="Min Value:"), 0, wx.ALL, 5)
        min_ctrl = wx.SpinCtrlDouble(panel, value=str(self.norm_min), min=0, max=1000000, inc=0.1)
        min_sizer.Add(min_ctrl, 1, wx.ALL, 5)

        # Max value input
        max_sizer = wx.BoxSizer(wx.HORIZONTAL)
        max_sizer.Add(wx.StaticText(panel, label="Max Value:"), 0, wx.ALL, 5)
        max_ctrl = wx.SpinCtrlDouble(panel, value=str(self.norm_max), min=0, max=1000000, inc=0.1)
        max_sizer.Add(max_ctrl, 1, wx.ALL, 5)

        # Buttons
        btn_sizer = wx.StdDialogButtonSizer()
        ok_btn = wx.Button(panel, wx.ID_OK)
        cancel_btn = wx.Button(panel, wx.ID_CANCEL)
        btn_sizer.AddButton(ok_btn)
        btn_sizer.AddButton(cancel_btn)
        btn_sizer.Realize()

        # Add to main sizer
        sizer.Add(min_sizer, 0, wx.EXPAND | wx.ALL, 5)
        sizer.Add(max_sizer, 0, wx.EXPAND | wx.ALL, 5)
        sizer.Add(btn_sizer, 0, wx.EXPAND | wx.ALL, 10)

        panel.SetSizer(sizer)

        if dlg.ShowModal() == wx.ID_OK:
            self.norm_min = min_ctrl.GetValue()
            self.norm_max = max_ctrl.GetValue()

        dlg.Destroy()

    def on_norm_changed(self, event):
        if self.norm_check.GetValue():
            # Update the plot with normalization enabled
            self.replot_with_normalization()
        else:
            # Hide normalization cursors if visible
            self.hide_norm_cursors()
            # Replot without normalization
            self.replot_with_normalization()

    def on_auto_changed(self, event):
        if self.norm_check.GetValue():
            # Re-apply normalization with new Auto setting
            self.replot_with_normalization()

    def on_key_press(self, event):
        if event.key == 'shift' and self.norm_check.GetValue() and not self.auto_check.GetValue():
            # Show normalization cursors
            self.show_norm_cursors()

    def on_key_release(self, event):
        if event.key == 'shift':
            # Hide normalization cursors unless they're being dragged
            if not self.is_dragging_cursor:
                self.hide_norm_cursors()

    def show_norm_cursors(self):
        """Display draggable cursors for manual normalization"""
        if not self.parent.ax:
            return

        sheet_names = self.get_selected_sheet_names()
        if not sheet_names:
            return

        # Get data from first sheet
        first_sheet = sheet_names[0]
        if first_sheet not in self.parent.Data['Core levels']:
            return

        x_values = self.parent.Data['Core levels'][first_sheet]['B.E.']
        y_values = self.parent.Data['Core levels'][first_sheet]['Raw Data']

        x_min = min(x_values) + 0.5  # 0.5V above min BE
        y_max = max(y_values)

        # Create or update cursors
        if self.norm_vlines[0] is None:
            self.norm_vlines[0] = self.parent.ax.axvline(x_min, color='red', linestyle='--', alpha=0.7,
                                                         label='Min Norm')
        else:
            self.norm_vlines[0].set_xdata([x_min, x_min])

        if self.norm_vlines[1] is None:
            self.norm_vlines[1] = self.parent.ax.axhline(y_max, color='green', linestyle='--', alpha=0.7,
                                                         label='Max Norm')
        else:
            self.norm_vlines[1].set_ydata([y_max, y_max])

        # Update normalization values
        self.norm_min = min(y_values)
        self.norm_max = y_max

        # Make cursors draggable
        self.make_cursors_draggable()

        self.parent.canvas.draw_idle()

    def hide_norm_cursors(self):
        """Hide normalization cursors"""
        for i, vline in enumerate(self.norm_vlines):
            if vline:
                vline.remove()
                self.norm_vlines[i] = None

        self.parent.canvas.draw_idle()

    def make_cursors_draggable(self):
        """Make the normalization cursors draggable"""
        self.is_dragging_cursor = False
        self.active_cursor = None

        def on_press(event):
            if event.inaxes != self.parent.ax:
                return

            # Check if click is near either cursor
            for i, cursor in enumerate(self.norm_vlines):
                if cursor is None:
                    continue

                if i == 0:  # Vertical line for min BE
                    xdata = cursor.get_xdata()[0]
                    if abs(event.xdata - xdata) < 5:  # 5 is threshold in data units
                        self.is_dragging_cursor = True
                        self.active_cursor = (cursor, i)
                        return
                else:  # Horizontal line for max intensity
                    ydata = cursor.get_ydata()[0]
                    if abs(event.ydata - ydata) < 5:  # 5 is threshold in data units
                        self.is_dragging_cursor = True
                        self.active_cursor = (cursor, i)
                        return

        def on_motion(event):
            if not self.is_dragging_cursor or self.active_cursor is None:
                return

            cursor, i = self.active_cursor

            if i == 0:  # Vertical cursor
                cursor.set_xdata([event.xdata, event.xdata])
                # Update min normalization value
                self.norm_min = event.xdata
            else:  # Horizontal cursor
                cursor.set_ydata([event.ydata, event.ydata])
                # Update max normalization value
                self.norm_max = event.ydata

            self.parent.canvas.draw_idle()

        def on_release(event):
            if self.is_dragging_cursor:
                self.is_dragging_cursor = False
                self.active_cursor = None
                # Reapply normalization with new bounds
                self.replot_with_normalization()

        self.press_cid = self.parent.canvas.mpl_connect('button_press_event', on_press)
        self.motion_cid = self.parent.canvas.mpl_connect('motion_notify_event', on_motion)
        self.release_cid = self.parent.canvas.mpl_connect('button_release_event', on_release)
