import wx
import numpy as np
import os
from scipy.ndimage import gaussian_filter
from scipy.signal import savgol_filter
from scipy.integrate import cumtrapz
import json


class PlotModWindow(wx.Frame):
    def __init__(self, parent, *args, **kw):
        super().__init__(parent, *args, **kw, style=wx.DEFAULT_FRAME_STYLE & ~(
                wx.RESIZE_BORDER | wx.MAXIMIZE_BOX | wx.MINIMIZE_BOX | wx.SYSTEM_MENU) | wx.STAY_ON_TOP)

        self.SetTitle("Plot Modifications")
        self.SetSize(340, 450)

        self.parent = parent
        panel = wx.Panel(self)

        main_pos = parent.GetPosition()
        main_size = parent.GetSize()
        mod_size = self.GetSize()

        x = main_pos.x + (main_size.width - mod_size.width) // 2
        y = main_pos.y + (main_size.height - mod_size.height) // 2

        self.SetPosition((x, y))

        grid_sizer = wx.GridBagSizer(5, 5)

        # First row - Smoothing
        smooth_box = wx.StaticBox(panel, label="Smoothing")
        smooth_sizer = wx.StaticBoxSizer(smooth_box, wx.VERTICAL)

        self.smooth_method = wx.ComboBox(panel, choices=["Gaussian", "Savitzky-Golay", "Moving Average"],
                                         style=wx.CB_READONLY)
        self.smooth_method.SetValue("Gaussian")
        self.smooth_width = wx.SpinCtrl(panel, min=1, max=100, initial=5)

        smooth_sizer.Add(wx.StaticText(panel, label="Method:"), 0, wx.ALL, 5)
        smooth_sizer.Add(self.smooth_method, 0, wx.EXPAND | wx.ALL, 5)
        smooth_sizer.Add(wx.StaticText(panel, label="Width:"), 0, wx.ALL, 5)
        smooth_sizer.Add(self.smooth_width, 0, wx.EXPAND | wx.ALL, 5)

        smooth_btn = wx.Button(panel, label="Apply Smoothing")
        smooth_btn.SetMinSize((125, 40))
        smooth_btn.Bind(wx.EVT_BUTTON, self.on_smooth)
        smooth_sizer.Add(smooth_btn, 0, wx.EXPAND | wx.ALL, 5)

        # First row - Differentiation
        diff_box = wx.StaticBox(panel, label="Differentiation")
        diff_sizer = wx.StaticBoxSizer(diff_box, wx.VERTICAL)

        self.diff_width = wx.SpinCtrl(panel, min=1, max=100, initial=5)
        diff_sizer.Add(wx.StaticText(panel, label="Width:"), 0, wx.ALL, 5)
        diff_sizer.Add(self.diff_width, 0, wx.EXPAND | wx.ALL, 5)

        diff_btn = wx.Button(panel, label="Apply Differentiation")
        diff_btn.SetMinSize((125, 40))
        diff_btn.Bind(wx.EVT_BUTTON, self.on_differentiate)
        diff_sizer.Add(diff_btn, 0, wx.EXPAND | wx.ALL, 5)

        # Second row - Integration
        int_box = wx.StaticBox(panel, label="Integration")
        int_sizer = wx.StaticBoxSizer(int_box, wx.VERTICAL)

        self.int_width = wx.SpinCtrl(panel, min=1, max=100, initial=5)
        int_sizer.Add(wx.StaticText(panel, label="Width:"), 0, wx.ALL, 5)
        int_sizer.Add(self.int_width, 0, wx.EXPAND | wx.ALL, 5)

        int_btn = wx.Button(panel, label="Apply Integration")
        int_btn.SetMinSize((125, 40))
        int_btn.Bind(wx.EVT_BUTTON, self.on_integrate)
        int_sizer.Add(int_btn, 0, wx.EXPAND | wx.ALL, 5)

        # Constant Operation
        const_box = wx.StaticBox(panel, label="Constant Operation")
        const_sizer = wx.StaticBoxSizer(const_box, wx.VERTICAL)

        self.const_op = wx.ComboBox(panel, choices=["Multiply", "Divide", "Add", "Subtract"], style=wx.CB_READONLY)
        self.const_op.SetValue("Multiply")
        const_sizer.Add(self.const_op, 0, wx.EXPAND | wx.ALL, 5)

        self.const_value = wx.SpinCtrlDouble(panel, value="1.0", min=0.001, max=10000000.0, inc=0.1)
        const_sizer.Add(wx.StaticText(panel, label="Value:"), 0, wx.ALL, 5)
        const_sizer.Add(self.const_value, 0, wx.EXPAND | wx.ALL, 5)

        const_btn = wx.Button(panel, label="Apply Operation")
        const_btn.SetMinSize((125, 40))
        const_btn.Bind(wx.EVT_BUTTON, self.on_apply_constant)
        const_sizer.Add(const_btn, 0, wx.EXPAND | wx.ALL, 5)

        # Add to grid
        grid_sizer.Add(smooth_sizer, pos=(0, 0), flag=wx.EXPAND | wx.ALL, border=5)
        grid_sizer.Add(diff_sizer, pos=(0, 1), flag=wx.EXPAND | wx.ALL, border=5)
        grid_sizer.Add(int_sizer, pos=(1, 0), flag=wx.EXPAND | wx.ALL, border=5)
        grid_sizer.Add(const_sizer, pos=(1, 1), flag=wx.EXPAND | wx.ALL, border=5)

        from libraries.ConfigFile import set_consistent_fonts
        set_consistent_fonts(self)

        panel.SetSizer(grid_sizer)
        self.Centre()

    # Method for constant operation
    def on_apply_constant(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        constant = self.const_value.GetValue()
        operation = self.const_op.GetValue()

        x = self.parent.Data['Core levels'][sheet_name]['B.E.']
        y = self.parent.Data['Core levels'][sheet_name]['Raw Data']

        # Apply the operation
        if operation == "Multiply":
            modified_y = [val * constant for val in y]
        elif operation == "Divide":
            modified_y = [val / constant for val in y]
        elif operation == "Add":
            modified_y = [val + constant for val in y]
        elif operation == "Subtract":
            modified_y = [val - constant for val in y]

        # Get the base name
        import re
        match = re.match(r'([A-Za-z]+\d*[spdfg]*)', sheet_name)
        if match:
            base_name = match.group(1)
        else:
            base_name = sheet_name

        # Find the earliest available row
        new_sheet_name = self.get_earliest_row_name(base_name)

        # Save the data
        self.save_modified_data(x, modified_y, new_sheet_name, f"{operation}d")

    def get_earliest_row_name(self, base_name):
        """
        Find the earliest available row for a core level.
        Returns the sheet name with appropriate row number.
        """
        import re

        # Check all sheets in parent data
        all_sheets = list(self.parent.Data['Core levels'].keys())

        # Check if base name without number exists (row 0)
        if base_name in all_sheets:
            # Base name exists, need to find next available
            rows_used = []
        else:
            # Base name doesn't exist, it's available
            return base_name

        # Pattern to match base name followed by optional number
        pattern = re.compile(f"^{re.escape(base_name)}(\\d+)$")

        # Collect all used row numbers
        for sheet in all_sheets:
            if sheet == base_name:  # This is row 0
                rows_used.append(0)
            else:
                match = pattern.match(sheet)
                if match:
                    rows_used.append(int(match.group(1)))

        # Find first unused row number
        for i in range(1000):  # Reasonable upper limit
            if i not in rows_used:
                if i == 0:
                    return base_name  # No suffix for row 0
                else:
                    return f"{base_name}{i}"

        # Fallback (unlikely to reach)
        return f"{base_name}{len(all_sheets)}"

    def save_modified_data(self, x, y, sheet_name, operation_type):
        """Save modified data to a new sheet and refresh file manager."""
        import pandas as pd
        import openpyxl

        # Create DataFrame with required columns
        df = pd.DataFrame({
            'BE': x,
            'Raw Data': y,
            'Background': y,
            'Transmission': [1.0] * len(x)
        })

        # Load workbook
        wb = openpyxl.load_workbook(self.parent.Data['FilePath'])

        # Save to Excel
        with pd.ExcelWriter(self.parent.Data['FilePath'], engine='openpyxl', mode='a',
                            if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Update window.Data
        self.parent.Data['Core levels'][sheet_name] = {
            'B.E.': x if isinstance(x, list) else x.tolist(),
            'Raw Data': y if isinstance(y, list) else y.tolist(),
            'Background': {'Bkg Y': y if isinstance(y, list) else y.tolist()},
            'Name': sheet_name
        }
        self.parent.Data['Number of Core levels'] += 1

        # Update JSON file
        json_file_path = os.path.splitext(self.parent.Data['FilePath'])[0] + '.json'
        if os.path.exists(json_file_path):
            from libraries.FileMenu.Save import convert_to_serializable_and_round
            json_data = convert_to_serializable_and_round(self.parent.Data)
            with open(json_file_path, 'w') as json_file:
                json.dump(json_data, json_file, indent=2)

        # Close and reopen file manager if it exists
        if hasattr(self.parent, 'file_manager') and self.parent.file_manager is not None:
            try:
                # Close existing file manager
                self.parent.file_manager.Close()
                self.parent.file_manager.Destroy()
                self.parent.file_manager = None

                # Reopen file manager
                import wx
                wx.CallAfter(self.parent.on_open_file_manager, None)
            except Exception as e:
                print(f"Error refreshing file manager: {e}")
                pass

        # Update sheet list
        self.parent.sheet_combobox.Append(sheet_name)
        self.parent.sheet_combobox.SetValue(sheet_name)
        from libraries.Sheet_Operations import on_sheet_selected
        on_sheet_selected(self.parent, sheet_name)

    def on_smooth(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        method = self.smooth_method.GetValue()
        width = self.smooth_width.GetValue()

        x = self.parent.Data['Core levels'][sheet_name]['B.E.']
        y = self.parent.Data['Core levels'][sheet_name]['Raw Data']

        if method == "Gaussian":
            smoothed = gaussian_filter(y, width)
        elif method == "Savitzky-Golay":
            if width % 2 == 0:  # Ensure odd window length for savgol_filter
                width += 1
            smoothed = savgol_filter(y, width, 3)
        else:  # Moving Average
            kernel = np.ones(width) / width
            smoothed = np.convolve(y, kernel, mode='same')

        # Get the base name
        import re
        match = re.match(r'([A-Za-z]+\d*[spdfg]*)', sheet_name)
        if match:
            base_name = match.group(1)
        else:
            base_name = sheet_name

        # Find the earliest available row
        new_sheet_name = self.get_earliest_row_name(base_name)

        # Save the data
        self.save_modified_data(x, smoothed, new_sheet_name, "Smoothed")

    def on_differentiate(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        width = self.diff_width.GetValue()

        x = self.parent.Data['Core levels'][sheet_name]['B.E.']
        y = self.parent.Data['Core levels'][sheet_name]['Raw Data']

        derivative = np.gradient(y, x)

        # Ensure odd window length for savgol_filter
        if width % 2 == 0:
            width += 1

        smoothed_deriv = savgol_filter(derivative, width, 3)

        # Get the base name
        import re
        match = re.match(r'([A-Za-z]+\d*[spdfg]*)', sheet_name)
        if match:
            base_name = match.group(1)
        else:
            base_name = sheet_name

        # Find the earliest available row
        new_sheet_name = self.get_earliest_row_name(base_name)

        # Save the data
        self.save_modified_data(x, smoothed_deriv, new_sheet_name, "Differentiated")

    def on_integrate(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        width = self.int_width.GetValue()

        x = self.parent.Data['Core levels'][sheet_name]['B.E.']
        y = self.parent.Data['Core levels'][sheet_name]['Raw Data']

        integrated = cumtrapz(y, x, initial=0)

        # Ensure odd window length for savgol_filter
        if width % 2 == 0:
            width += 1

        smoothed_int = savgol_filter(integrated, width, 3)

        # Get the base name
        import re
        match = re.match(r'([A-Za-z]+\d*[spdfg]*)', sheet_name)
        if match:
            base_name = match.group(1)
        else:
            base_name = sheet_name

        # Find the earliest available row
        new_sheet_name = self.get_earliest_row_name(base_name)

        # Save the data
        self.save_modified_data(x, smoothed_int, new_sheet_name, "Integrated")