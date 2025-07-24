import wx
import numpy as np
import os
from scipy.ndimage import gaussian_filter
from scipy.signal import savgol_filter
from scipy.integrate import cumtrapz
import json
import lmfit.models  # Add this if not already present
import pandas as pd


class PlotModWindow(wx.Frame):
    def __init__(self, parent, *args, **kw):
        super().__init__(parent, *args, **kw, style=wx.DEFAULT_FRAME_STYLE & ~(
                wx.RESIZE_BORDER | wx.MAXIMIZE_BOX | wx.MINIMIZE_BOX | wx.SYSTEM_MENU) | wx.STAY_ON_TOP)

        self.SetTitle("Plot Modifications")

        if 'wxMac' in wx.PlatformInfo:
            self.SetSize(430, 345)
        elif 'wxGTK' in wx.PlatformInfo:
            self.SetSize(420, 465)
        else:
            self.SetSize(450, 390)
        self.parent = parent

        self.controls_hidden = True  # Start with controls hidden

        panel = wx.Panel(self)

        main_pos = parent.GetPosition()
        main_size = parent.GetSize()
        mod_size = self.GetSize()

        x = main_pos.x + (main_size.width - mod_size.width) // 2
        y = main_pos.y + (main_size.height - mod_size.height) // 2

        self.SetPosition((x, y))

        grid_sizer = wx.GridBagSizer(5, 5)

        # Define the two green colors
        light_green = wx.Colour(220, 240, 220)  # Light green
        darker_green = wx.Colour(240, 255, 240)  # Darker green

        # Create a panel instead of StaticBox
        smooth_panel = wx.Panel(panel)
        smooth_panel.SetBackgroundColour(light_green)

        # Create sizer for the panel contents
        smooth_sizer = wx.BoxSizer(wx.VERTICAL)

        # Add title text inside the panel
        title_text = wx.StaticText(smooth_panel, label="Smoothing Mod.")
        title_font = title_text.GetFont()
        title_font.SetWeight(wx.FONTWEIGHT_BOLD)
        title_text.SetFont(title_font)
        smooth_sizer.Add(title_text, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        # Add a separator line (optional)
        line = wx.StaticLine(smooth_panel, style=wx.LI_HORIZONTAL)
        smooth_sizer.Add(line, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 5)

        # Add your controls
        self.smooth_method = wx.ComboBox(smooth_panel, choices=["Gaussian", "Savitzky-Golay", "Moving Average"],
                                         style=wx.CB_READONLY)
        self.smooth_method.SetValue("Gaussian")
        self.smooth_width = wx.SpinCtrl(smooth_panel, min=1, max=100, initial=5)

        smooth_sizer.Add(wx.StaticText(smooth_panel, label="Method:"), 0, wx.ALL, 5)
        smooth_sizer.Add(self.smooth_method, 0, wx.EXPAND | wx.ALL, 5)
        smooth_sizer.Add(wx.StaticText(smooth_panel, label="Width:"), 0, wx.ALL, 5)
        smooth_sizer.Add(self.smooth_width, 0, wx.EXPAND | wx.ALL, 5)

        smooth_btn = wx.Button(smooth_panel, label="Apply Smoothing")
        smooth_btn.SetMinSize((125, 40))
        smooth_btn.Bind(wx.EVT_BUTTON, self.on_smooth)
        smooth_sizer.Add(smooth_btn, 0, wx.EXPAND | wx.ALL, 5)

        # Set the sizer to the panel
        smooth_panel.SetSizer(smooth_sizer)

        # Noise section - Create container panel that will hold both actual and placeholder panels
        noise_container = wx.Panel(panel)
        noise_container_sizer = wx.BoxSizer(wx.VERTICAL)

        # Create the actual noise panel
        self.noise_panel = wx.Panel(noise_container)
        self.noise_panel.SetBackgroundColour(light_green)
        noise_sizer = wx.BoxSizer(wx.VERTICAL)

        title_text = wx.StaticText(self.noise_panel, label="Noise Mod.")
        title_font = title_text.GetFont()
        title_font.SetWeight(wx.FONTWEIGHT_BOLD)
        title_text.SetFont(title_font)
        noise_sizer.Add(title_text, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        line = wx.StaticLine(self.noise_panel, style=wx.LI_HORIZONTAL)
        noise_sizer.Add(line, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 5)

        self.noise_type = wx.ComboBox(self.noise_panel, choices=["Gaussian", "Uniform", "Poisson"],
                                      style=wx.CB_READONLY)
        self.noise_type.SetValue("Gaussian")
        self.noise_amplitude = wx.SpinCtrlDouble(self.noise_panel, min=0.0, max=10000.0, initial=1.0, inc=0.1)

        noise_sizer.Add(wx.StaticText(self.noise_panel, label="Type:"), 0, wx.ALL, 5)
        noise_sizer.Add(self.noise_type, 0, wx.EXPAND | wx.ALL, 5)
        noise_sizer.Add(wx.StaticText(self.noise_panel, label="Amplitude:"), 0, wx.ALL, 5)
        noise_sizer.Add(self.noise_amplitude, 0, wx.EXPAND | wx.ALL, 5)

        noise_btn = wx.Button(self.noise_panel, label="Add Noise")
        noise_btn.SetMinSize((125, 40))
        noise_btn.Bind(wx.EVT_BUTTON, self.on_add_noise)
        noise_sizer.Add(noise_btn, 0, wx.EXPAND | wx.ALL, 5)
        self.noise_panel.SetSizer(noise_sizer)

        # Create placeholder panel for noise (same appearance)
        self.noise_placeholder = wx.Panel(noise_container)
        self.noise_placeholder.SetBackgroundColour(light_green)
        placeholder_sizer = wx.BoxSizer(wx.VERTICAL)

        title_text = wx.StaticText(self.noise_placeholder, label="Other Mods.")
        title_font = title_text.GetFont()
        title_font.SetWeight(wx.FONTWEIGHT_BOLD)
        title_text.SetFont(title_font)
        placeholder_sizer.Add(title_text, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        line = wx.StaticLine(self.noise_placeholder, style=wx.LI_HORIZONTAL)
        placeholder_sizer.Add(line, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 5)

        hidden_text = wx.StaticText(self.noise_placeholder, label="Coming soon")
        hidden_text.SetForegroundColour(wx.Colour(128, 128, 128))
        placeholder_sizer.Add(hidden_text, 1, wx.ALL | wx.ALIGN_CENTER, 30)

        self.noise_placeholder.SetSizer(placeholder_sizer)

        # Add both panels to container (they'll stack, but only one will be visible)
        noise_container_sizer.Add(self.noise_panel, 1, wx.EXPAND)
        noise_container_sizer.Add(self.noise_placeholder, 1, wx.EXPAND)
        noise_container.SetSizer(noise_container_sizer)

        # Differentiation section
        diff_panel = wx.Panel(panel)
        diff_panel.SetBackgroundColour(light_green)
        diff_sizer = wx.BoxSizer(wx.VERTICAL)

        title_text = wx.StaticText(diff_panel, label="Differentiation Mod.")
        title_font = title_text.GetFont()
        title_font.SetWeight(wx.FONTWEIGHT_BOLD)
        title_text.SetFont(title_font)
        diff_sizer.Add(title_text, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        line = wx.StaticLine(diff_panel, style=wx.LI_HORIZONTAL)
        diff_sizer.Add(line, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 5)

        self.diff_width = wx.SpinCtrl(diff_panel, min=1, max=100, initial=5)
        diff_sizer.Add(wx.StaticText(diff_panel, label="Width:"), 0, wx.ALL, 5)
        diff_sizer.Add(self.diff_width, 0, wx.EXPAND | wx.ALL, 5)

        diff_btn = wx.Button(diff_panel, label="Apply Differentiation")
        diff_btn.SetMinSize((125, 40))
        diff_btn.Bind(wx.EVT_BUTTON, self.on_differentiate)
        diff_sizer.Add(diff_btn, 0, wx.EXPAND | wx.ALL, 5)
        diff_panel.SetSizer(diff_sizer)

        # Integration section
        int_panel = wx.Panel(panel)
        int_panel.SetBackgroundColour(darker_green)
        int_sizer = wx.BoxSizer(wx.VERTICAL)

        title_text = wx.StaticText(int_panel, label="Integration Mod.")
        title_font = title_text.GetFont()
        title_font.SetWeight(wx.FONTWEIGHT_BOLD)
        title_text.SetFont(title_font)
        int_sizer.Add(title_text, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        line = wx.StaticLine(int_panel, style=wx.LI_HORIZONTAL)
        int_sizer.Add(line, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 5)

        self.int_width = wx.SpinCtrl(int_panel, min=1, max=100, initial=5)
        int_sizer.Add(wx.StaticText(int_panel, label="Width:"), 0, wx.ALL, 5)
        int_sizer.Add(self.int_width, 0, wx.EXPAND | wx.ALL, 5)

        int_btn = wx.Button(int_panel, label="Apply Integration")
        int_btn.SetMinSize((125, 40))
        int_btn.Bind(wx.EVT_BUTTON, self.on_integrate)
        int_sizer.Add(int_btn, 0, wx.EXPAND | wx.ALL, 5)
        int_panel.SetSizer(int_sizer)

        # Constant Operation section
        const_panel = wx.Panel(panel)
        const_panel.SetBackgroundColour(darker_green)
        const_sizer = wx.BoxSizer(wx.VERTICAL)

        title_text = wx.StaticText(const_panel, label="Constant Mod.")
        title_font = title_text.GetFont()
        title_font.SetWeight(wx.FONTWEIGHT_BOLD)
        title_text.SetFont(title_font)
        const_sizer.Add(title_text, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        line = wx.StaticLine(const_panel, style=wx.LI_HORIZONTAL)
        const_sizer.Add(line, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 5)

        const_sizer.Add(wx.StaticText(const_panel, label="Type:"), 0, wx.ALL, 5)
        self.const_op = wx.ComboBox(const_panel, choices=["Multiply", "Divide", "Add", "Subtract", "Special"],
                                    style=wx.CB_READONLY)
        self.const_op.SetValue("Multiply")
        const_sizer.Add(self.const_op, 0, wx.EXPAND | wx.ALL, 5)

        self.const_value = wx.SpinCtrlDouble(const_panel, value="1.0", min=0.001, max=10000000.0, inc=0.1)
        const_sizer.Add(wx.StaticText(const_panel, label="Value:"), 0, wx.ALL, 5)
        const_sizer.Add(self.const_value, 0, wx.EXPAND | wx.ALL, 5)

        const_btn = wx.Button(const_panel, label="Apply Operation")
        const_btn.SetMinSize((125, 40))
        const_btn.Bind(wx.EVT_BUTTON, self.on_apply_constant)
        const_sizer.Add(const_btn, 0, wx.EXPAND | wx.ALL, 5)
        const_panel.SetSizer(const_sizer)

        # BE Shift section
        be_shift_panel = wx.Panel(panel)
        be_shift_panel.SetBackgroundColour(darker_green)
        be_shift_sizer = wx.BoxSizer(wx.VERTICAL)

        title_text = wx.StaticText(be_shift_panel, label="BE Shift Mod.")
        title_font = title_text.GetFont()
        title_font.SetWeight(wx.FONTWEIGHT_BOLD)
        title_text.SetFont(title_font)
        be_shift_sizer.Add(title_text, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        line = wx.StaticLine(be_shift_panel, style=wx.LI_HORIZONTAL)
        be_shift_sizer.Add(line, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 5)

        be_shift_sizer.Add(wx.StaticText(be_shift_panel, label="Shift (eV):"), 0, wx.ALL, 5)
        self.be_shift_value = wx.SpinCtrlDouble(be_shift_panel, value="0.0", min=-100.0, max=100.0, inc=0.1,
                                                size=(80, -1))
        be_shift_sizer.Add(self.be_shift_value, 0, wx.EXPAND | wx.ALL, 5)

        be_shift_btn = wx.Button(be_shift_panel, label="Apply BE Shift")
        be_shift_btn.SetMinSize((125, 40))
        be_shift_btn.Bind(wx.EVT_BUTTON, self.on_apply_be_shift)
        be_shift_sizer.Add(be_shift_btn, 0, wx.EXPAND | wx.ALL, 5)
        be_shift_panel.SetSizer(be_shift_sizer)

        light_blue = wx.Colour(220, 240, 255)  # Light green

        # Voigt Model Generator section
        voigt_panel = wx.Panel(panel)
        self.voigt_panel = voigt_panel
        voigt_panel.SetBackgroundColour(light_blue)
        voigt_sizer = wx.BoxSizer(wx.VERTICAL)

        title_text = wx.StaticText(voigt_panel, label="Create Voigt Model")
        title_font = title_text.GetFont()
        title_font.SetWeight(wx.FONTWEIGHT_BOLD)
        title_text.SetFont(title_font)
        voigt_sizer.Add(title_text, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        line = wx.StaticLine(voigt_panel, style=wx.LI_HORIZONTAL)
        voigt_sizer.Add(line, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 2)

        # Position control
        voigt_sizer.Add(wx.StaticText(voigt_panel, label="Position (eV):"), 0, wx.ALL, 2)
        self.voigt_position = wx.SpinCtrlDouble(voigt_panel, value="285.0", min=0.0, max=2000.0, inc=0.1)
        voigt_sizer.Add(self.voigt_position, 0, wx.EXPAND | wx.ALL, 2)

        # FWHM control
        voigt_sizer.Add(wx.StaticText(voigt_panel, label="FWHM (eV):"), 0, wx.ALL, 2)
        self.voigt_fwhm = wx.SpinCtrlDouble(voigt_panel, value="1.5", min=0.1, max=10.0, inc=0.1)
        voigt_sizer.Add(self.voigt_fwhm, 0, wx.EXPAND | wx.ALL, 2)

        # L/G Ratio control
        voigt_sizer.Add(wx.StaticText(voigt_panel, label="L/G Ratio (%):"), 0, wx.ALL, 2)
        self.voigt_lg_ratio = wx.SpinCtrlDouble(voigt_panel, value="30.0", min=0.0, max=100.0, inc=1.0)
        voigt_sizer.Add(self.voigt_lg_ratio, 0, wx.EXPAND | wx.ALL, 2)

        # Height control
        voigt_sizer.Add(wx.StaticText(voigt_panel, label="Height (CPS):"), 0, wx.ALL, 2)
        self.voigt_height = wx.SpinCtrlDouble(voigt_panel, value="1000.0", min=1.0, max=1e7, inc=100.0)
        voigt_sizer.Add(self.voigt_height, 0, wx.EXPAND | wx.ALL, 2)

        # Range Min control
        voigt_sizer.Add(wx.StaticText(voigt_panel, label="Range Min (eV):"), 0, wx.ALL, 2)
        self.voigt_min = wx.SpinCtrlDouble(voigt_panel, value="280.0", min=0.0, max=2000.0, inc=0.1)
        voigt_sizer.Add(self.voigt_min, 0, wx.EXPAND | wx.ALL, 2)

        # Range Max control
        voigt_sizer.Add(wx.StaticText(voigt_panel, label="Range Max (eV):"), 0, wx.ALL, 2)
        self.voigt_max = wx.SpinCtrlDouble(voigt_panel, value="290.0", min=0.0, max=2000.0, inc=0.1)
        voigt_sizer.Add(self.voigt_max, 0, wx.EXPAND | wx.ALL, 2)

        # Create button
        voigt_btn = wx.Button(voigt_panel, label="Create Voigt Model")
        voigt_btn.SetMinSize((125, 40))
        voigt_btn.Bind(wx.EVT_BUTTON, self.on_create_voigt)
        voigt_sizer.Add(voigt_btn, 0, wx.EXPAND | wx.ALL, 2)
        voigt_panel.SetSizer(voigt_sizer)

        # Add to grid - note all are now panels instead of sizers
        grid_sizer.Add(smooth_panel, pos=(0, 0), flag=wx.EXPAND | wx.ALL, border=1)
        grid_sizer.Add(noise_container, pos=(0, 2), flag=wx.EXPAND | wx.ALL, border=1)
        grid_sizer.Add(diff_panel, pos=(1, 1), flag=wx.EXPAND | wx.ALL, border=1)
        grid_sizer.Add(int_panel, pos=(1, 0), flag=wx.EXPAND | wx.ALL, border=1)
        grid_sizer.Add(const_panel, pos=(0, 1), flag=wx.EXPAND | wx.ALL, border=1)
        grid_sizer.Add(voigt_panel, pos=(0, 3), span=(2, 1), flag=wx.EXPAND | wx.ALL, border=1)
        grid_sizer.Add(be_shift_panel, pos=(1, 2), flag=wx.EXPAND | wx.ALL, border=1)

        # Auto-populate Voigt parameters from current plot
        self.auto_populate_voigt_from_plot()

        from libraries.ConfigFile import set_consistent_fonts
        set_consistent_fonts(self)

        # Hide noise and voigt panels initially
        if self.controls_hidden:
            self.noise_panel.Hide()
            self.voigt_panel.Hide()
            self.noise_placeholder.Show()  # Make sure placeholder is visible
            # # Set smaller window size
            # if 'wxMac' in wx.PlatformInfo:
            #     self.SetSize(380, 345)  # Reduced from 560, 345
            # elif 'wxGTK' in wx.PlatformInfo:
            #     self.SetSize(420, 465)  # Reduced from 620, 465
            # else:
            #     self.SetSize(400, 390)  # Reduced from 580, 390

        panel.SetSizer(grid_sizer)
        self.Centre()

    # Method for constant operation
    def on_apply_constant(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        constant = self.const_value.GetValue()
        operation = self.const_op.GetValue()

        # Check for special toggle operation
        if operation == "Special" and constant == 1976:
            self.toggle_hidden_controls()
            return

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
        match = re.match(r'([A-Za-z]+(?:\d+[spdfg]+)?)', sheet_name)
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
            'Corrected Data': y,  # Modified data goes here
            'Raw Data': y,  # Keep original reference
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
                wx.CallAfter(self.parent.on_open_file_manager, None)
            except Exception as e:
                print(f"Error refreshing file manager: {e}")
                pass

        # Update sheet list
        self.parent.sheet_combobox.Append(sheet_name)
        self.parent.sheet_combobox.SetValue(sheet_name)
        from libraries.Sheet_Operations import on_sheet_selected
        on_sheet_selected(self.parent, sheet_name)

    def on_add_noise(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        noise_type = self.noise_type.GetValue()
        amplitude = self.noise_amplitude.GetValue()

        x = self.parent.Data['Core levels'][sheet_name]['B.E.']
        y = np.array(self.parent.Data['Core levels'][sheet_name]['Raw Data'])

        if noise_type == "Gaussian":
            noise = np.random.normal(0, amplitude, len(y))
        elif noise_type == "Uniform":
            noise = np.random.uniform(-amplitude, amplitude, len(y))
        else:  # Poisson
            # Scale data to avoid negative values for Poisson
            scaled_y = np.abs(y) + 1
            noise = np.random.poisson(scaled_y * amplitude / 100) - scaled_y * amplitude / 100

        noisy_y = y + noise

        # Get base name and create new sheet
        import re
        # match = re.match(r'([A-Za-z]+\d*[spdfg]*)', sheet_name)
        match = re.match(r'([A-Za-z]+(?:\d+[spdfg]+)?)', sheet_name)
        base_name = match.group(1) if match else sheet_name
        new_sheet_name = self.get_earliest_row_name(base_name)

        # Save the noisy data
        self.save_modified_data(x, noisy_y, new_sheet_name, "Noisy")

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
        match = re.match(r'([A-Za-z]+(?:\d+[spdfg]+)?)', sheet_name)
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
        # match = re.match(r'([A-Za-z]+\d*[spdfg]*)', sheet_name)
        match = re.match(r'([A-Za-z]+(?:\d+[spdfg]+)?)', sheet_name)
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
        # match = re.match(r'([A-Za-z]+\d*[spdfg]*)', sheet_name)
        match = re.match(r'([A-Za-z]+(?:\d+[spdfg]+)?)', sheet_name)
        if match:
            base_name = match.group(1)
        else:
            base_name = sheet_name

        # Find the earliest available row
        new_sheet_name = self.get_earliest_row_name(base_name)

        # Save the data
        self.save_modified_data(x, smoothed_int, new_sheet_name, "Integrated")

    def on_create_voigt(self, event):
        """Create a new sheet with Voigt model data and save to Excel"""
        import openpyxl

        # Get current sheet for base name and x-axis
        current_sheet = self.parent.sheet_combobox.GetValue()

        # Get current x-axis data (use exact same x-values as original)
        current_x = np.array(self.parent.Data['Core levels'][current_sheet]['B.E.'])

        # Get parameters
        position = self.voigt_position.GetValue()
        fwhm = self.voigt_fwhm.GetValue()
        lg_ratio = self.voigt_lg_ratio.GetValue()
        height = self.voigt_height.GetValue()
        x_min = self.voigt_min.GetValue()
        x_max = self.voigt_max.GetValue()

        # Validate range
        if x_min >= x_max:
            wx.MessageBox("Min value must be less than Max value", "Invalid Range", wx.OK | wx.ICON_ERROR)
            return

        # Filter x_values to the specified range
        mask = (current_x >= x_min) & (current_x <= x_max)
        if not np.any(mask):
            wx.MessageBox("No data points found in the specified range", "Invalid Range", wx.OK | wx.ICON_ERROR)
            return

        # Use filtered x-values from original data (this ensures identical x-axis)
        x_values = current_x[mask]

        # Calculate Voigt parameters
        sigma = fwhm / (2 * np.sqrt(2 * np.log(2)))  # Convert FWHM to sigma
        gamma = (lg_ratio / 100) * sigma  # Calculate gamma from L/G ratio

        # Create Voigt model
        voigt_model = lmfit.models.VoigtModel()

        # Calculate amplitude to achieve desired height
        amplitude = height / voigt_model.eval(center=0, amplitude=1, sigma=sigma, gamma=gamma, x=0)

        # Generate Voigt profile using the exact same x-values
        y_values = voigt_model.eval(x=x_values, center=position, amplitude=amplitude, sigma=sigma, gamma=gamma)

        # Create full arrays (same length as original data) with zeros outside range
        full_x = current_x.copy()
        full_y = np.zeros_like(current_x)
        full_y[mask] = y_values  # Only fill the range where Voigt is defined

        # Get base name for new sheet
        import re
        # match = re.match(r'([A-Za-z]+\d*[spdfg]*)', current_sheet)
        match = re.match(r'([A-Za-z]+(?:\d+[spdfg]+)?)', current_sheet)
        base_name = match.group(1) if match else "Voigt"

        # Find earliest available row name
        new_sheet_name = self.get_earliest_row_name(base_name)

        # Create DataFrame with correct column names
        df = pd.DataFrame({
            'BE': full_x,
            'Corrected Data': full_y,  # Voigt model data
            'Raw Data': full_y,  # Same as corrected for Voigt
            'Transmission': np.ones_like(full_x)
        })

        # Save to Excel file
        try:
            # Save DataFrame to Excel
            with pd.ExcelWriter(self.parent.Data['FilePath'], engine='openpyxl', mode='a',
                                if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=new_sheet_name, index=False)

            # Create new core level data structure for memory
            new_core_level_data = {
                'B.E.': full_x.tolist(),
                'Raw Data': full_y.tolist(),
                'Background': {'Bkg Y': np.zeros_like(full_y).tolist()},
                'Name': new_sheet_name
            }

            # Add to parent data
            self.parent.Data['Core levels'][new_sheet_name] = new_core_level_data
            self.parent.Data['Number of Core levels'] += 1

            # Update sheet combobox
            self.parent.sheet_combobox.Append(new_sheet_name)

            # Close and reopen the file manager if it exists
            if hasattr(self.parent, 'file_manager') and self.parent.file_manager is not None:
                try:
                    # Close existing file manager
                    self.parent.file_manager.Close()
                    self.parent.file_manager.Destroy()
                    self.parent.file_manager = None

                    # Reopen file manager
                    wx.CallAfter(self.parent.on_open_file_manager, None)
                except Exception as e:
                    print(f"Error refreshing file manager: {e}")
                    pass

            # Select the new sheet
            self.parent.sheet_combobox.SetValue(new_sheet_name)

            # Refresh the plot
            from libraries.Sheet_Operations import on_sheet_selected
            on_sheet_selected(self.parent, new_sheet_name)

            # wx.MessageBox(f"Voigt model created: {new_sheet_name}\nPoints: {len(full_x)}",
            #               "Success", wx.OK | wx.ICON_INFORMATION)

        except Exception as e:
            wx.MessageBox(f"Error saving to Excel: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def on_apply_be_shift(self, event):
        """Create a new sheet with BE-shifted data and save to Excel"""
        import pandas as pd
        import openpyxl

        # Get current sheet and shift value
        current_sheet = self.parent.sheet_combobox.GetValue()
        shift_value = self.be_shift_value.GetValue()

        # Get current data
        x = self.parent.Data['Core levels'][current_sheet]['B.E.']
        y = self.parent.Data['Core levels'][current_sheet]['Raw Data']

        # Apply BE shift (add shift to all BE values)
        shifted_x = [val + shift_value for val in x]

        # Get base name for new sheet
        import re
        match = re.match(r'([A-Za-z]+\d*[spdfg]*)', current_sheet)
        base_name = match.group(1) if match else current_sheet

        # Find earliest available row name
        new_sheet_name = self.get_earliest_row_name(base_name)

        # Create DataFrame with correct column names
        df = pd.DataFrame({
            'BE': shifted_x,
            'Corrected Data': y,  # Use Corrected Data instead of Raw Data
            'Raw Data': y,  # Keep original as Raw Data
            'Transmission': np.ones(len(y))  # Unit transmission
        })

        # Save to Excel file
        try:
            # Save DataFrame to Excel
            with pd.ExcelWriter(self.parent.Data['FilePath'], engine='openpyxl', mode='a',
                                if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=new_sheet_name, index=False)

            # Create new core level data structure for memory
            new_core_level_data = {
                'B.E.': shifted_x,
                'Raw Data': y if isinstance(y, list) else y.tolist(),
                'Background': {'Bkg Y': np.zeros_like(y).tolist()},
                'Name': new_sheet_name
            }

            # Add to parent data
            self.parent.Data['Core levels'][new_sheet_name] = new_core_level_data
            self.parent.Data['Number of Core levels'] += 1

            # Update sheet combobox
            self.parent.sheet_combobox.Append(new_sheet_name)

            # Close and reopen the file manager if it exists
            if hasattr(self.parent, 'file_manager') and self.parent.file_manager is not None:
                try:
                    # Close existing file manager
                    self.parent.file_manager.Close()
                    self.parent.file_manager.Destroy()
                    self.parent.file_manager = None

                    # Reopen file manager
                    wx.CallAfter(self.parent.on_open_file_manager, None)
                except Exception as e:
                    print(f"Error refreshing file manager: {e}")
                    pass

            # Select the new sheet
            self.parent.sheet_combobox.SetValue(new_sheet_name)

            # Refresh the plot
            from libraries.Sheet_Operations import on_sheet_selected
            on_sheet_selected(self.parent, new_sheet_name)

            # wx.MessageBox(f"BE shifted by {shift_value} eV: {new_sheet_name}", "Success", wx.OK | wx.ICON_INFORMATION)

        except Exception as e:
            wx.MessageBox(f"Error saving to Excel: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def auto_populate_voigt_from_plot(self):
        """Auto-populate Voigt parameters from current plot limits"""
        try:
            sheet_name = self.parent.sheet_combobox.GetValue()

            # Get plot limits from plot_config
            if hasattr(self.parent, 'plot_config') and sheet_name in self.parent.plot_config.plot_limits:
                limits = self.parent.plot_config.plot_limits[sheet_name]

                # Get X range (BE range)
                x_min = limits['Xmin']
                x_max = limits['Xmax']
                y_max = limits['Ymax']

                # Set position to middle of plot range
                middle_position = (x_min + x_max) / 2.0
                self.voigt_position.SetValue(middle_position)

                # Set range to plot limits
                self.voigt_min.SetValue(x_min)
                self.voigt_max.SetValue(x_max)

                # Set reasonable height based on plot Y max (50% of max intensity)
                reasonable_height = y_max * 0.5
                self.voigt_height.SetValue(max(reasonable_height, 100.0))  # Minimum 100 CPS

            else:
                # Fallback: try to get from data directly
                if sheet_name in self.parent.Data['Core levels']:
                    x_data = self.parent.Data['Core levels'][sheet_name]['B.E.']
                    y_data = self.parent.Data['Core levels'][sheet_name]['Raw Data']

                    if len(x_data) > 0:
                        x_min = min(x_data)
                        x_max = max(x_data)
                        y_max = max(y_data)

                        middle_position = (x_min + x_max) / 2.0
                        self.voigt_position.SetValue(middle_position)
                        self.voigt_min.SetValue(x_min)
                        self.voigt_max.SetValue(x_max)
                        self.voigt_height.SetValue(max(y_max * 0.5, 100.0))

        except Exception as e:
            print(f"Could not auto-populate Voigt parameters: {e}")
            # Keep default values if there's an error

    def toggle_hidden_controls(self):
        """Toggle visibility of noise and voigt panels and resize window accordingly"""
        self.controls_hidden = not self.controls_hidden

        if self.controls_hidden:
            # Hide actual panels and show placeholder for noise
            self.noise_panel.Hide()
            self.voigt_panel.Hide()
            self.noise_placeholder.Show()

            # Set smaller window size (your specified sizes)
            if 'wxMac' in wx.PlatformInfo:
                self.SetSize(430, 345)
            elif 'wxGTK' in wx.PlatformInfo:
                self.SetSize(420, 465)
            else:
                self.SetSize(400, 390)
        else:
            # Show actual panels and hide placeholder
            self.noise_placeholder.Hide()
            self.noise_panel.Show()
            self.voigt_panel.Show()

            # Set larger window size (your specified original sizes)
            if 'wxMac' in wx.PlatformInfo:
                self.SetSize(560, 345)
            elif 'wxGTK' in wx.PlatformInfo:
                self.SetSize(620, 465)
            else:
                self.SetSize(580, 390)

        # Refresh the layout
        self.Layout()
        self.Refresh()

        # Re-center the window on the parent
        main_pos = self.parent.GetPosition()
        main_size = self.parent.GetSize()
        mod_size = self.GetSize()

        x = main_pos.x + (main_size.width - mod_size.width) // 2
        y = main_pos.y + (main_size.height - mod_size.height) // 2

        self.SetPosition((x, y))