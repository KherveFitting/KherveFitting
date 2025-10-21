import wx
import wx.grid
import json
import os
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
import openpyxl


class NPLTransmissionWindow(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, title="NPL Transmission Function", size=(1000, 600))
        self.parent = parent

        # Initialize transmission parameters
        self.a0 = None
        self.a1 = None
        self.a2 = None
        self.a3 = None
        self.a4 = None
        self.b1 = None
        self.b2 = None
        self.b3 = None
        self.b4 = None
        self.min_ke = None
        self.max_ke = None
        self.photon_energy = 1486.69  # Default Al Ka

        self.init_ui()
        self.load_parameters_from_config()
        self.Centre()

    def init_ui(self):
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Left panel for controls
        left_panel = wx.Panel(panel)
        left_sizer = wx.BoxSizer(wx.VERTICAL)

        # VMS file drop zone
        drop_label = wx.StaticText(left_panel, label="Drop VMS File Here:")
        left_sizer.Add(drop_label, 0, wx.ALL, 5)

        self.vms_path_text = wx.TextCtrl(left_panel, style=wx.TE_READONLY, size=(300, 25))
        self.vms_path_text.SetDropTarget(VMSFileDropTarget(self))
        left_sizer.Add(self.vms_path_text, 0, wx.ALL | wx.EXPAND, 5)

        browse_btn = wx.Button(left_panel, label="Browse VMS File")
        browse_btn.Bind(wx.EVT_BUTTON, self.on_browse_vms)
        left_sizer.Add(browse_btn, 0, wx.ALL | wx.EXPAND, 5)

        # Parameters display
        param_label = wx.StaticText(left_panel, label="Transmission Parameters:")
        left_sizer.Add(param_label, 0, wx.ALL | wx.TOP, 10)

        self.param_grid = wx.grid.Grid(left_panel)
        self.param_grid.CreateGrid(11, 2)
        self.param_grid.SetColLabelValue(0, "Parameter")
        self.param_grid.SetColLabelValue(1, "Value")
        self.param_grid.SetColSize(0, 100)
        self.param_grid.SetColSize(1, 180)
        self.param_grid.EnableEditing(False)

        # Set parameter names
        param_names = ["a0", "a1", "a2", "a3", "a4", "b1", "b2", "b3", "b4", "Min KE (eV)", "Max KE (eV)"]
        for i, name in enumerate(param_names):
            self.param_grid.SetCellValue(i, 0, name)
            self.param_grid.SetCellValue(i, 1, "")

        left_sizer.Add(self.param_grid, 1, wx.ALL | wx.EXPAND, 5)

        # Sheet selection
        sheet_label = wx.StaticText(left_panel, label="Select Sheet:")
        left_sizer.Add(sheet_label, 0, wx.ALL | wx.TOP, 10)

        self.sheet_combo = wx.ComboBox(left_panel, style=wx.CB_READONLY)
        self.sheet_combo.Bind(wx.EVT_COMBOBOX, self.on_sheet_select)
        left_sizer.Add(self.sheet_combo, 0, wx.ALL | wx.EXPAND, 5)

        # Apply button
        apply_btn = wx.Button(left_panel, label="Write Transmission to Excel")
        apply_btn.Bind(wx.EVT_BUTTON, self.on_apply_transmission)
        left_sizer.Add(apply_btn, 0, wx.ALL | wx.EXPAND, 5)

        left_panel.SetSizer(left_sizer)
        main_sizer.Add(left_panel, 0, wx.ALL | wx.EXPAND, 5)

        # Right panel for plot
        right_panel = wx.Panel(panel)
        right_sizer = wx.BoxSizer(wx.VERTICAL)

        self.figure, self.ax = plt.subplots(figsize=(6, 5))
        self.canvas = FigureCanvas(right_panel, -1, self.figure)
        right_sizer.Add(self.canvas, 1, wx.ALL | wx.EXPAND, 5)

        right_panel.SetSizer(right_sizer)
        main_sizer.Add(right_panel, 1, wx.ALL | wx.EXPAND, 5)

        panel.SetSizer(main_sizer)

        # Update sheet combobox
        self.update_sheet_list()

    def on_browse_vms(self, event):
        wildcard = "VMS files (*.vms)|*.vms"
        dlg = wx.FileDialog(self, "Choose VMS file", wildcard=wildcard, style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)

        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            self.load_vms_file(path)

        dlg.Destroy()

    def load_vms_file(self, filepath):
        """Parse VMS file and extract transmission function parameters"""
        try:
            with open(filepath, 'r') as f:
                content = f.read()

            # Extract parameters from ASCII section
            params = {}
            for line in content.split('\n'):
                if 'Response Function Parameter a0' in line:
                    params['a0'] = float(line.split()[-1])
                elif 'Response Function Parameter a1' in line:
                    params['a1'] = float(line.split()[-1])
                elif 'Response Function Parameter a2' in line:
                    params['a2'] = float(line.split()[-1])
                elif 'Response Function Parameter a3' in line:
                    params['a3'] = float(line.split()[-1])
                elif 'Response Function Parameter a4' in line:
                    params['a4'] = float(line.split()[-1])
                elif 'Response Function Parameter b1' in line:
                    params['b1'] = float(line.split()[-1])
                elif 'Response Function Parameter b2' in line:
                    params['b2'] = float(line.split()[-1])
                elif 'Response Function Parameter b3' in line:
                    params['b3'] = float(line.split()[-1])
                elif 'Response Function Parameter b4' in line:
                    params['b4'] = float(line.split()[-1])
                elif 'Minimum Calibration Energy/eV' in line:
                    params['min_ke'] = float(line.split()[-1])
                elif 'Maximum Calibration Energy/eV' in line:
                    params['max_ke'] = float(line.split()[-1])

            if len(params) != 11:
                wx.MessageBox("Could not find all required parameters in VMS file", "Error", wx.OK | wx.ICON_ERROR)
                return

            # Update class attributes
            self.a0 = params['a0']
            self.a1 = params['a1']
            self.a2 = params['a2']
            self.a3 = params['a3']
            self.a4 = params['a4']
            self.b1 = params['b1']
            self.b2 = params['b2']
            self.b3 = params['b3']
            self.b4 = params['b4']
            self.min_ke = params['min_ke']
            self.max_ke = params['max_ke']

            # Update display
            self.vms_path_text.SetValue(filepath)
            self.update_parameter_display()
            self.save_parameters_to_config()
            self.plot_transmission_function()

            wx.MessageBox("VMS file loaded successfully!", "Success", wx.OK | wx.ICON_INFORMATION)

        except Exception as e:
            wx.MessageBox(f"Error loading VMS file: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def update_parameter_display(self):
        """Update parameter grid with current values"""
        values = [self.a0, self.a1, self.a2, self.a3, self.a4,
                  self.b1, self.b2, self.b3, self.b4, self.min_ke, self.max_ke]

        for i, val in enumerate(values):
            if val is not None:
                self.param_grid.SetCellValue(i, 1, f"{val:.6f}")

    def calculate_transmission(self, kinetic_energy):
        """
        Calculate Q(E) transmission function at given kinetic energy
        Uses normalized energy: ε = (E - 1000 eV) / 1000 eV
        Q(E) = (a0 + a1*ε + a2*ε^2 + a3*ε^3 + a4*ε^4) / (1 + b1*ε + b2*ε^2 + b3*ε^3 + b4*ε^4)
        """
        if None in [self.a0, self.a1, self.a2, self.a3, self.a4, self.b1, self.b2, self.b3, self.b4]:
            return 1.0

        # Normalize the kinetic energy
        epsilon = (kinetic_energy - 1000.0) / 1000.0

        numerator = self.a0 + self.a1 * epsilon + self.a2 * epsilon ** 2 + self.a3 * epsilon ** 3 + self.a4 * epsilon ** 4
        denominator = 1.0 + self.b1 * epsilon + self.b2 * epsilon ** 2 + self.b3 * epsilon ** 3 + self.b4 * epsilon ** 4

        return numerator / denominator

    def plot_transmission_function(self):
        """Plot the transmission function"""
        if None in [self.a0, self.min_ke, self.max_ke]:
            return

        # Create kinetic energy range
        ke_range = np.linspace(self.min_ke, self.max_ke, 500)
        transmission = [self.calculate_transmission(ke) for ke in ke_range]

        self.ax.clear()
        self.ax.plot(ke_range, transmission, 'b-', linewidth=2)
        self.ax.set_xlabel('Kinetic Energy (eV)', fontsize=12)
        self.ax.set_ylabel('Q(E) Transmission', fontsize=12)
        self.ax.set_title('NPL Transmission Function', fontsize=14)
        self.ax.grid(True, alpha=0.3)
        self.canvas.draw()

    def update_sheet_list(self):
        """Update sheet combobox with available sheets"""
        if 'FilePath' in self.parent.Data and self.parent.Data['FilePath']:
            try:
                wb = openpyxl.load_workbook(self.parent.Data['FilePath'])
                sheets = [s for s in wb.sheetnames if s not in ['Experimental description', 'Sheet']]
                self.sheet_combo.Clear()
                self.sheet_combo.AppendItems(sheets)
                if sheets:
                    self.sheet_combo.SetSelection(0)
                wb.close()
            except:
                pass

    def on_sheet_select(self, event):
        """Handle sheet selection"""
        pass

    def on_apply_transmission(self, event):
        """Apply transmission correction to selected sheet in Excel"""
        if None in [self.a0, self.a1, self.a2, self.a3, self.a4, self.b1, self.b2, self.b3, self.b4]:
            wx.MessageBox("Please load a VMS file first", "Error", wx.OK | wx.ICON_ERROR)
            return

        sheet_name = self.sheet_combo.GetValue()
        if not sheet_name:
            wx.MessageBox("Please select a sheet", "Error", wx.OK | wx.ICON_ERROR)
            return

        if 'FilePath' not in self.parent.Data or not self.parent.Data['FilePath']:
            wx.MessageBox("No Excel file is open", "Error", wx.OK | wx.ICON_ERROR)
            return

        try:
            # Get photon energy from parent
            photon_energy = self.parent.photons

            # Load workbook
            wb = openpyxl.load_workbook(self.parent.Data['FilePath'])

            if sheet_name not in wb.sheetnames:
                wx.MessageBox(f"Sheet '{sheet_name}' not found", "Error", wx.OK | wx.ICON_ERROR)
                wb.close()
                return

            ws = wb[sheet_name]

            # Read binding energies from column A and raw data from column C
            row = 2
            while ws.cell(row=row, column=1).value is not None:
                be_value = ws.cell(row=row, column=1).value
                raw_data = ws.cell(row=row, column=3).value

                if be_value is not None and raw_data is not None:
                    # Convert BE to KE
                    ke = photon_energy - float(be_value)

                    # Calculate transmission
                    transmission = self.calculate_transmission(ke)

                    # Write transmission to column D
                    ws.cell(row=row, column=4, value=float(f"{transmission:.2f}"))

                    # Calculate corrected data: raw_data / transmission
                    corrected = float(raw_data) / transmission
                    ws.cell(row=row, column=2, value=float(f"{corrected:.2f}"))

                row += 1

            # Save workbook
            wb.save(self.parent.Data['FilePath'])
            wb.close()

            wx.MessageBox(f"Transmission applied to {sheet_name}\nColumn B: Corrected data\nColumn D: Transmission values",
                          "Success", wx.OK | wx.ICON_INFORMATION)

            # Reload the data in parent
            if hasattr(self.parent, 'clear_and_replot'):
                self.parent.clear_and_replot()

        except Exception as e:
            wx.MessageBox(f"Error applying transmission: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def save_parameters_to_config(self):
        """Save transmission parameters to config.json"""
        config_path = 'config.json'

        try:
            # Load existing config
            if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    config = json.load(f)
            else:
                config = {}

            # Add NPL parameters
            config['npl_transmission'] = {
                'a0': self.a0,
                'a1': self.a1,
                'a2': self.a2,
                'a3': self.a3,
                'a4': self.a4,
                'b1': self.b1,
                'b2': self.b2,
                'b3': self.b3,
                'b4': self.b4,
                'min_ke': self.min_ke,
                'max_ke': self.max_ke
            }

            # Save config
            with open(config_path, 'w') as f:
                json.dump(config, f, indent=2)

        except Exception as e:
            print(f"Error saving config: {e}")

    def load_parameters_from_config(self):
        """Load transmission parameters from config.json"""
        config_path = 'config.json'

        try:
            if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    config = json.load(f)

                if 'npl_transmission' in config:
                    params = config['npl_transmission']
                    self.a0 = params.get('a0')
                    self.a1 = params.get('a1')
                    self.a2 = params.get('a2')
                    self.a3 = params.get('a3')
                    self.a4 = params.get('a4')
                    self.b1 = params.get('b1')
                    self.b2 = params.get('b2')
                    self.b3 = params.get('b3')
                    self.b4 = params.get('b4')
                    self.min_ke = params.get('min_ke')
                    self.max_ke = params.get('max_ke')

                    self.update_parameter_display()
                    self.plot_transmission_function()

        except Exception as e:
            print(f"Error loading config: {e}")


class VMSFileDropTarget(wx.FileDropTarget):
    def __init__(self, window):
        super().__init__()
        self.window = window

    def OnDropFiles(self, x, y, filenames):
        if filenames and filenames[0].endswith('.vms'):
            self.window.load_vms_file(filenames[0])
            return True
        return False