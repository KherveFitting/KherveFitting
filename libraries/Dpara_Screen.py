import wx
import numpy as np
from libraries.Peak_Functions import OtherCalc


class DParameterWindow(wx.Frame):
    def __init__(self, parent,*args, **kw):
        super().__init__(parent, title="D-Parameter Measurement", *args, **kw,
                         style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX | wx.MINIMIZE_BOX | wx.SYSTEM_MENU) | wx.STAY_ON_TOP)
        self.parent = parent
        self.SetSize((300, 440))

        panel = wx.Panel(self)
        # panel.SetBackgroundColour(wx.WHITE)

        # Parameters Box
        param_box = wx.StaticBox(panel, label="Parameters")
        param_sizer = wx.StaticBoxSizer(param_box, wx.VERTICAL)

        grid_sizer = wx.GridSizer(rows=5, cols=2, hgap=5, vgap=5)

        # Smooth Width
        grid_sizer.Add(wx.StaticText(panel, label="Smooth Width"), 0, wx.ALL | wx.CENTER, 5)
        self.smooth_spin = wx.SpinCtrlDouble(panel, value="7.0", min=0.1, max=10.0, inc=0.1)
        grid_sizer.Add(self.smooth_spin, 0, wx.ALL | wx.EXPAND, 5)

        # Pre-Smooth Passes
        grid_sizer.Add(wx.StaticText(panel, label="Pre-Smooth Passes"), 0, wx.ALL | wx.CENTER, 5)
        self.pre_spin = wx.SpinCtrl(panel, value="2", min=0, max=10)
        grid_sizer.Add(self.pre_spin, 0, wx.ALL | wx.EXPAND, 5)

        # Differentiation Width
        grid_sizer.Add(wx.StaticText(panel, label="Differentiation \nWidth (eV)"), 0, wx.ALL | wx.CENTER, 5)
        self.diff_spin = wx.SpinCtrlDouble(panel, value="1.0", min=0.1, max=10.0, inc=0.1)
        grid_sizer.Add(self.diff_spin, 0, wx.ALL | wx.EXPAND, 5)

        # Post-Smooth Passes
        grid_sizer.Add(wx.StaticText(panel, label="Post-Smooth Passes"), 0, wx.ALL | wx.CENTER, 5)
        self.post_spin = wx.SpinCtrl(panel, value="1", min=0, max=10)
        grid_sizer.Add(self.post_spin, 0, wx.ALL | wx.EXPAND, 5)

        # Smooth Algorithm
        grid_sizer.Add(wx.StaticText(panel, label="Smooth Algorithm"), 0, wx.ALL | wx.CENTER, 5)
        algorithms = ["Gaussian", "Savitsky-Golay", "Moving Average", "Wiener", "None"]
        self.algo_combo = wx.ComboBox(panel, choices=algorithms, style=wx.CB_READONLY)
        self.algo_combo.SetSelection(0)
        grid_sizer.Add(self.algo_combo, 0, wx.ALL | wx.EXPAND, 5)

        param_sizer.Add(grid_sizer, 0, wx.ALL | wx.EXPAND, 5)

        # D-Parameter Value Box
        d_box = wx.StaticBox(panel, label="D-Parameter Value (eV)")
        d_sizer = wx.StaticBoxSizer(d_box, wx.HORIZONTAL)
        self.d_value = wx.TextCtrl(panel, value="0")
        self.d_unit = wx.StaticText(panel, label="eV")
        d_sizer.Add(self.d_value, 1, wx.ALL | wx.EXPAND, 5)
        d_sizer.Add(self.d_unit, 0, wx.ALL | wx.CENTER, 5)

        # Buttons
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        clear_btn = wx.Button(panel, label="Clear D-para")
        calc_btn = wx.Button(panel, label="Calculate")
        clear_btn.SetMinSize((125, 40))
        calc_btn.SetMinSize((125, 40))

        btn_sizer.Add(clear_btn, 0, wx.ALL, 5)
        btn_sizer.Add(calc_btn, 0, wx.ALL, 5)

        # Bind events
        calc_btn.Bind(wx.EVT_BUTTON, self.on_calculate)
        clear_btn.Bind(wx.EVT_BUTTON, self.on_clear)

        # Main sizer
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        main_sizer.Add(param_sizer, 0, wx.ALL | wx.EXPAND, 5)
        main_sizer.Add(d_sizer, 0, wx.ALL | wx.EXPAND, 5)
        main_sizer.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        from libraries.ConfigFile import set_consistent_fonts
        set_consistent_fonts(self)

        panel.SetSizer(main_sizer)

    def on_clear(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()

        # Clear the grid
        num_rows = self.parent.peak_params_grid.GetNumberRows()
        if num_rows > 0:
            self.parent.peak_params_grid.DeleteRows(0, num_rows)

        # Clear the data
        if sheet_name in self.parent.Data['Core levels']:
            if 'Fitting' in self.parent.Data['Core levels'][sheet_name]:
                self.parent.Data['Core levels'][sheet_name]['Fitting'] = {}

        # Reset peak count
        self.parent.peak_count = 0

        # Clear the plot
        for line in self.parent.ax.lines:
            if line.get_label() == 'Derivative':
                line.remove()
        self.parent.canvas.draw_idle()

    def on_calculate(self, event):
        x_values = self.parent.x_values
        y_values = self.parent.y_values
        sheet_name = self.parent.sheet_combobox.GetValue()

        # Clear previous derivative plots
        for line in self.parent.ax.lines:
            if line.get_label() == 'Derivative':
                line.remove()

        # Remove previous D-parameter rows if they exist
        for row in range(self.parent.peak_params_grid.GetNumberRows() - 1, -1, -1):
            if self.parent.peak_params_grid.GetCellValue(row, 13) == "D-parameter":
                self.parent.peak_params_grid.DeleteRows(row, 2)

        algorithm = self.algo_combo.GetString(self.algo_combo.GetSelection())
        normalized_deriv = OtherCalc.smooth_and_differentiate(
            x_values,
            y_values,
            self.smooth_spin.GetValue(),
            self.pre_spin.GetValue(),
            self.diff_spin.GetValue(),
            self.post_spin.GetValue(),
            algorithm
        )

        min_idx = np.argmin(normalized_deriv)
        max_idx = np.argmax(normalized_deriv)
        x_min = x_values[min_idx]
        x_max = x_values[max_idx]
        center = (x_max + x_min) / 2
        d_param = round(abs(x_max - x_min), 2)

        # Update D-value display
        self.d_value.SetValue(f"{d_param:.2f}")

        # Add new rows
        row = self.parent.peak_params_grid.GetNumberRows()
        self.parent.peak_params_grid.AppendRows(2)

        # Update grid values
        self.parent.peak_params_grid.SetCellValue(row, 0, chr(65 + row // 2))
        self.parent.peak_params_grid.SetCellValue(row, 1, "D-parameter")
        self.parent.peak_params_grid.SetCellValue(row, 2, f"{center:.2f}")
        self.parent.peak_params_grid.SetCellValue(row, 4, f"{d_param:.2f}")
        self.parent.peak_params_grid.SetCellValue(row, 7, f"{self.pre_spin.GetValue():.2f}")
        self.parent.peak_params_grid.SetCellValue(row, 8, f"{self.post_spin.GetValue():.2f}")
        self.parent.peak_params_grid.SetCellValue(row, 9, f"{self.smooth_spin.GetValue():.2f}")
        self.parent.peak_params_grid.SetCellValue(row, 5, f"{self.diff_spin.GetValue():.2f}")
        self.parent.peak_params_grid.SetCellValue(row, 13, "D-parameter")

        # Update Data dictionary
        if 'Fitting' not in self.parent.Data['Core levels'][sheet_name]:
            self.parent.Data['Core levels'][sheet_name]['Fitting'] = {}
        if 'Peaks' not in self.parent.Data['Core levels'][sheet_name]['Fitting']:
            self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = {}

        peak_data = {
            'Position': center,
            'FWHM': d_param,
            'Sigma': self.pre_spin.GetValue(),
            'Gamma': self.post_spin.GetValue(),
            'Skew': self.smooth_spin.GetValue(),
            'L/G': self.diff_spin.GetValue(),
            'Fitting Model': 'D-parameter',
            'Derivative': normalized_deriv.tolist()  # Store derivative for replotting
        }

        self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']['D-parameter'] = peak_data

        # Update plot
        self.parent.ax.plot(x_values, normalized_deriv, '-', color='red', label='Derivative')
        self.parent.canvas.draw_idle()

    def on_defaults(self, event):
        self.smooth_spin.SetValue(7.0)
        self.pre_spin.SetValue(2)
        self.diff_spin.SetValue(1.0)
        self.post_spin.SetValue(1)
        self.algo_radio.SetSelection(0)

    def on_ok(self, event):
        self.Close()

    def on_cancel(self, event):
        self.Close()