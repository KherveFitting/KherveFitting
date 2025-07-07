import wx
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.patches as patches


class ThickogramWindow(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, title="Thickogram XPS Thickness Calculator", size=(1200, 800))
        self.parent = parent
        self.InitUI()
        self.Centre()

    def InitUI(self):
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Left panel for controls
        left_panel = wx.Panel(panel)
        left_sizer = wx.BoxSizer(wx.VERTICAL)

        # Title
        title = wx.StaticText(left_panel, label="Thickogram Calculator")
        title_font = wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        title.SetFont(title_font)
        left_sizer.Add(title, 0, wx.ALL | wx.CENTER, 10)

        # Create input fields
        self.create_input_fields(left_panel, left_sizer)

        # Calculate button
        calc_btn = wx.Button(left_panel, label="Calculate & Plot")
        calc_btn.Bind(wx.EVT_BUTTON, self.on_calculate)
        left_sizer.Add(calc_btn, 0, wx.ALL | wx.EXPAND, 10)

        # Result label
        self.result_label = wx.StaticText(left_panel, label="")
        result_font = wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        self.result_label.SetFont(result_font)
        self.result_label.SetForegroundColour(wx.Colour(255, 0, 0))
        left_sizer.Add(self.result_label, 0, wx.ALL | wx.CENTER, 10)

        left_panel.SetSizer(left_sizer)

        # Right panel for plot
        right_panel = wx.Panel(panel)
        right_sizer = wx.BoxSizer(wx.VERTICAL)

        # Create matplotlib figure
        self.figure = Figure(figsize=(8, 6), dpi=100)
        self.canvas = FigureCanvas(right_panel, -1, self.figure)
        right_sizer.Add(self.canvas, 1, wx.EXPAND | wx.ALL, 5)

        right_panel.SetSizer(right_sizer)

        # Add panels to main sizer
        main_sizer.Add(left_panel, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(right_panel, 1, wx.EXPAND | wx.ALL, 5)

        panel.SetSizer(main_sizer)

        # Initialize with default values
        self.set_default_values()

    def create_input_fields(self, parent, sizer):
        self.inputs = {}

        # Get peak names from results grid for combo boxes
        peak_names = self.get_peak_names()

        fields = [
            ("overlayer_combo", "Overlayer Peak:", "ComboBox", peak_names),
            ("substrate_combo", "Substrate Peak:", "ComboBox", peak_names),
            ("Eo", "BE of Overlayer (eV):", "TextCtrl", "707"),
            ("Es", "BE of Substrate (eV):", "TextCtrl", "100"),
            ("Io", "Intensity of Overlayer (Io):", "TextCtrl", "1000"),
            ("Is", "Intensity of Substrate (Is):", "TextCtrl", "2000"),
            ("So", "RSF of Overlayer (So):", "TextCtrl", "2.8"),
            ("Ss", "RSF of Substrate (Ss):", "TextCtrl", "0.8"),
            ("Lambda", "Lambda (Lo) in nm:", "TextCtrl", "2.0"),
            ("Theta", "Theta (degrees):", "TextCtrl", "0"),
            ("Esource", "X-ray Source Energy (eV):", "TextCtrl", str(self.parent.photons))
        ]

        for field_name, label, ctrl_type, default_value in fields:
            # Create label
            label_ctrl = wx.StaticText(parent, label=label)
            sizer.Add(label_ctrl, 0, wx.ALL | wx.ALIGN_LEFT, 5)

            # Create control
            if ctrl_type == "ComboBox":
                ctrl = wx.ComboBox(parent, choices=default_value, style=wx.CB_READONLY)
                if field_name == "overlayer_combo":
                    ctrl.Bind(wx.EVT_COMBOBOX, self.on_overlayer_selected)
                elif field_name == "substrate_combo":
                    ctrl.Bind(wx.EVT_COMBOBOX, self.on_substrate_selected)
            else:
                ctrl = wx.TextCtrl(parent, value=str(default_value))

            self.inputs[field_name] = ctrl
            sizer.Add(ctrl, 0, wx.ALL | wx.EXPAND, 5)

    def get_peak_names(self):
        """Get peak names from the results grid"""
        peak_names = []
        if hasattr(self.parent, 'results_grid'):
            for row in range(self.parent.results_grid.GetNumberRows()):
                peak_name = self.parent.results_grid.GetCellValue(row, 0)  # Column 0 is Peak Label
                if peak_name:
                    peak_names.append(peak_name)
        return peak_names

    def set_default_values(self):
        """Set default values from KherveFitting"""
        # Set X-ray source energy
        self.inputs["Esource"].SetValue(str(self.parent.photons))
        # Set theta to 0
        self.inputs["Theta"].SetValue("0")

    def on_overlayer_selected(self, event):
        """Handle overlayer peak selection"""
        selected_peak = event.GetString()
        self.populate_peak_data(selected_peak, "overlayer")

    def on_substrate_selected(self, event):
        """Handle substrate peak selection"""
        selected_peak = event.GetString()
        self.populate_peak_data(selected_peak, "substrate")

    def populate_peak_data(self, peak_name, peak_type):
        """Populate BE, Intensity, and RSF from results grid"""
        if not hasattr(self.parent, 'results_grid'):
            return

        for row in range(self.parent.results_grid.GetNumberRows()):
            grid_peak_name = self.parent.results_grid.GetCellValue(row, 0)
            if grid_peak_name == peak_name:
                # Get values from results grid
                position = self.parent.results_grid.GetCellValue(row, 1)  # Position (BE)
                area = self.parent.results_grid.GetCellValue(row, 5)  # Area
                rsf = self.parent.results_grid.GetCellValue(row, 8)  # RSF

                # Update appropriate fields
                if peak_type == "overlayer":
                    self.inputs["Eo"].SetValue(position)
                    self.inputs["Io"].SetValue(area)
                    self.inputs["So"].SetValue(rsf)
                elif peak_type == "substrate":
                    self.inputs["Es"].SetValue(position)
                    self.inputs["Is"].SetValue(area)
                    self.inputs["Ss"].SetValue(rsf)
                break

    def on_calculate(self, event):
        """Calculate thickness and plot thickogram"""
        try:
            # Get input values
            Eo = float(self.inputs["Eo"].GetValue())
            Es = float(self.inputs["Es"].GetValue())
            Io = float(self.inputs["Io"].GetValue())
            Is = float(self.inputs["Is"].GetValue())
            So = float(self.inputs["So"].GetValue())
            Ss = float(self.inputs["Ss"].GetValue())
            La = float(self.inputs["Lambda"].GetValue())
            theta = float(self.inputs["Theta"].GetValue())
            Esource = float(self.inputs["Esource"].GetValue())

            # Calculate thickness
            Eo = Esource - Eo
            Es = Esource - Es
            Eoes = Eo / Es
            Ro = So / Ss
            Ioisro = Io / (Is * Ro)
            rad = np.radians(theta)
            xthicko = self.solve_fAB_0(Ioisro, Eoes)
            t = xthicko * La * np.cos(rad)

            # Update result
            self.result_label.SetLabel(f"THICKNESS: {t:.2f} nm")

            # Plot thickogram
            self.plot_thickogram(xthicko, Ioisro, Eoes)

        except Exception as e:
            wx.MessageBox(f"Error in calculation: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def sinh(self, x):
        """Hyperbolic sine function"""
        return (np.exp(x) - np.exp(-x)) / 2

    def func(self, x):
        """Function used in thickogram calculation"""
        return x * ((4.2 - 0.6 * x) ** 0.75 - 0.5)

    def fAB(self, A, B, x):
        """Function fAB for thickness calculation"""
        return np.log(A / 2) - (B ** 0.75 - 0.5) * x - np.log(self.sinh(x / 2))

    def solve_fAB_0(self, A, B):
        """Solve fAB = 0 for thickness calculation"""
        x = 0.001
        deltax = 0.1
        eps = 1.0e-12
        while deltax > eps:
            fx = self.fAB(A, B, x)
            fxd = self.fAB(A, B, x + deltax)
            if fx * fxd >= 0.0:
                x += deltax
            else:
                if deltax <= eps:
                    x += deltax
                else:
                    deltax /= 10
        return x

    def plot_thickogram(self, xthicko, Ioisro, Eoes):
        """Plot the thickogram"""
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.set_facecolor('white')

        # Plot main curve
        x_vals = np.linspace(0.01, 6, 600)
        y_vals = np.log(self.sinh(x_vals / 2))
        ax.plot(2 * x_vals, y_vals, color='green', linewidth=1)

        # Plot grid points
        axp = np.arange(0.1, 6.1, 0.1)
        axd = [0.1, 0.2, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0]

        for x in axp:
            y = np.log(self.sinh(x / 2))
            ax.plot(2 * x, y, 'go', markersize=4)
        for x in axd:
            y = np.log(self.sinh(x / 2))
            ax.plot(2 * x, y, 'go', markersize=7)

        # Plot ratio lines
        ayp = [0.01, 0.02, 0.03, 0.04, 0.05, 0.06, 0.07, 0.08, 0.09, 0.1,
               0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1, 2, 3, 4, 5, 6,
               7, 8, 9, 10, 20, 30]

        for k in range(4, len(ayp) - 1):
            b_vals = np.arange(0.4, 3.1, 0.1)
            x_line = (4.2 - b_vals) / 0.6
            y_line = np.log(ayp[k] / 2) - self.func(x_line)
            ax.plot(2 * x_line, y_line, color='blue', linewidth=1)

        # Plot vertical lines
        bx = np.arange(0.4, 3.2, 0.2)
        for b in bx:
            x = (4.2 - b) / 0.6
            y0 = np.log(ayp[4] / 2) - self.func(x)
            y1 = np.log(ayp[-2] / 2) - self.func(x)
            ax.plot([2 * x, 2 * x], [y0, y1], color='blue', linewidth=1)

        # Plot horizontal reference lines
        for y in ayp:
            ax.plot([0, 0.18], [np.log(y / 2)] * 2, color='black', linewidth=1)

        # Add labels
        ayd = [0.01, 0.1, 1.0, 10.0]
        for y in ayd:
            ax.text(-0.4, np.log(y / 2), f"{y}", verticalalignment='center', color='black')

        for x in axd:
            y = np.log(self.sinh(x / 2))
            ax.text(2 * x + 0.15, y - 0.25, f"{x}", color='green')

        # Plot result lines
        x = xthicko
        y0 = np.log(self.sinh(x / 2)) - 0.3
        y1 = np.log(self.sinh(x / 2)) + 0.3
        ax.plot([2 * x, 2 * x], [y0, y1], color='red', linewidth=2)
        ax.plot([0, 2 * x], [np.log(Ioisro / 2), np.log(Ioisro / 2) - self.func((4.2 - Eoes) / 0.6)],
                color='red', linewidth=2)

        # Add result text
        ax.text(2 * x + 0.6, 2.4, f"{x:.6f}", color='red')

        # Set labels and formatting
        ax.set_title("Thickogram XPS Thickness Calculation")
        ax.set_xlabel("t/L*cos(theta) (2x)")
        ax.set_ylabel("log(Intensity Ratio)")
        ax.grid(True, linestyle='--', alpha=0.5)

        # Remove outer axes for clean look
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.tick_params(left=False, bottom=False, labelleft=False, labelbottom=False)

        self.canvas.draw()