import wx
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.patches as patches
import wx.lib.dialogs


class ThickogramWindow(wx.Frame):
    """
    Thickogram XPS Thickness Calculator

    This is a modification of the webvpython code as published at https://web.hallym.ac.kr/~jwlee/thickogram/
    which was written by Prof. Jong Wan Lee (Hallym University, School of Nano Convergence Technology, Korea)
    - Original paper (https://www.npsm-kps.org/journal/view.html?uid=7443)
    The original webpage no longer is accessible but the online webvpython app can be found at:
    https://www.glowscript.org/#/user/jwsslee/folder/Private/program/thickogram

    This implementation adapts the thickogram method for XPS thickness calculations in wxPython,
    integrating with KherveFitting's peak analysis results.
    """

    def __init__(self, parent):
        # Platform-specific window sizing
        if 'wxMac' in wx.PlatformInfo:
            window_size = (800, 540)  # Smaller for macOS
        elif 'wxGTK' in wx.PlatformInfo:  # Linux
            window_size = (920, 660)
        else:  # Windows
            window_size = (920, 660)

        super().__init__(parent, title="Cumpson-Lee Thickogram / Updated from DaveXPS v0.1",
                         size=window_size,
                         style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX))
        self.parent = parent
        self.InitUI()
        self.Centre()

        # Add consistent font handling like other screens
        from libraries.ConfigFile import set_consistent_fonts
        set_consistent_fonts(self)

    def InitUI(self):
        # Create menu bar
        menubar = wx.MenuBar()

        # Create File menu
        file_menu = wx.Menu()
        export_png_item = file_menu.Append(wx.ID_ANY, "Export as PNG\tCtrl+P", "Export plot as PNG image")
        export_svg_item = file_menu.Append(wx.ID_ANY, "Export as SVG\tCtrl+S", "Export plot as SVG image")

        # Create Help menu
        help_menu = wx.Menu()
        about_item = help_menu.Append(wx.ID_ABOUT, "&About", "About Thickogram Calculator")

        # Add menus to menu bar
        menubar.Append(file_menu, "&File")
        menubar.Append(help_menu, "&Help")
        self.SetMenuBar(menubar)

        # Bind menu events
        self.Bind(wx.EVT_MENU, self.on_export_png, export_png_item)
        self.Bind(wx.EVT_MENU, self.on_export_svg, export_svg_item)
        self.Bind(wx.EVT_MENU, self.on_about, about_item)


        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Left panel for controls
        left_panel = wx.Panel(panel)
        left_sizer = wx.BoxSizer(wx.VERTICAL)

        # # Title
        # title = wx.StaticText(left_panel, label="Thickogram Calculator")
        # title_font = wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        # title.SetFont(title_font)
        # left_sizer.Add(title, 0, wx.ALL | wx.CENTER, 5)

        # Create input fields
        self.create_input_fields(left_panel, left_sizer)

        # Calculate button
        calc_btn = wx.Button(left_panel, label="Calculate & Plot")
        if 'wxMac' in wx.PlatformInfo:
            calc_btn.SetMinSize((140, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            calc_btn.SetMinSize((160, 30))
        else:
            calc_btn.SetMinSize((160, 35))
        calc_btn.Bind(wx.EVT_BUTTON, self.on_calculate)
        left_sizer.Add(calc_btn, 0, wx.ALL, 3)  # Removed wx.EXPAND

        # # Result label
        # self.result_label = wx.StaticText(left_panel, label="", size=(160, 30))
        # result_font = wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        # self.result_label.SetFont(result_font)
        # self.result_label.SetForegroundColour(wx.Colour(255, 0, 0))
        # left_sizer.Add(self.result_label, 0, wx.ALL | wx.EXPAND, 3)

        left_panel.SetSizer(left_sizer)

        # Right panel for plot
        right_panel = wx.Panel(panel)
        right_sizer = wx.BoxSizer(wx.VERTICAL)

        # Create matplotlib figure (5x5 inches)
        self.figure = Figure(figsize=(5, 5), dpi=100)
        self.canvas = FigureCanvas(right_panel, -1, self.figure)
        right_sizer.Add(self.canvas, 1, wx.EXPAND | wx.ALL, 2)

        right_panel.SetSizer(right_sizer)

        # Add panels to main sizer
        main_sizer.Add(left_panel, 0, wx.EXPAND | wx.ALL, 2)
        main_sizer.Add(right_panel, 1, wx.EXPAND | wx.ALL, 2)

        panel.SetSizer(main_sizer)

        # Initialize with default values
        self.set_default_values()

        # Show empty thickogram initially
        self.plot_empty_thickogram()

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
            ("Z", "Atomic Number (Z):", "TextCtrl", "6"),
            ("Lambda", "Lambda (Lo) in nm:", "TextCtrl", "2.0"),
            ("Theta", "Theta (degrees):", "TextCtrl", "0"),
            ("Esource", "X-ray Source Energy (eV):", "TextCtrl", str(self.parent.photons))
        ]

        for field_name, label, ctrl_type, default_value in fields:
            # Create label with platform-specific font
            label_ctrl = wx.StaticText(parent, label=label)
            # Platform-specific font sizing like other screens
            if 'wxMac' in wx.PlatformInfo:
                label_font = wx.Font(11, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL,
                                     faceName='Helvetica')
            elif 'wxGTK' in wx.PlatformInfo:
                label_font = wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL,
                                     faceName='DejaVu Sans')
            else:
                label_font = wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL,
                                     faceName='Calibri')

            label_ctrl.SetFont(label_font)
            sizer.Add(label_ctrl, 0, wx.ALL | wx.ALIGN_LEFT, 4)

            # Create control with platform-specific sizing
            if 'wxMac' in wx.PlatformInfo:
                control_size = (140, -1)  # Smaller for macOS
            elif 'wxGTK' in wx.PlatformInfo:  # Linux
                control_size = (160, -1)
            else:  # Windows
                control_size = (160, -1)

            if ctrl_type == "ComboBox":
                ctrl = wx.ComboBox(parent, choices=default_value, style=wx.CB_READONLY, size=control_size)
                if field_name == "overlayer_combo":
                    ctrl.Bind(wx.EVT_COMBOBOX, self.on_overlayer_selected)
                elif field_name == "substrate_combo":
                    ctrl.Bind(wx.EVT_COMBOBOX, self.on_substrate_selected)
            else:
                ctrl = wx.TextCtrl(parent, value=str(default_value), size=control_size)
                if field_name == "Z":
                    ctrl.Bind(wx.EVT_TEXT, self.on_z_changed)

            # Apply same platform-specific font to controls
            ctrl.SetFont(label_font)
            self.inputs[field_name] = ctrl
            sizer.Add(ctrl, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 1)

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
        """Populate BE, Intensity, RSF and calculate Lambda from results grid"""
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

                    # Set Z based on peak name
                    z_value = self.determine_z_from_peak(peak_name)
                    self.inputs["Z"].SetValue(str(z_value))

                    # Calculate EAL using Z
                    try:
                        be_overlayer = float(position)
                        esource = float(self.inputs["Esource"].GetValue())
                        kinetic_energy = esource - be_overlayer
                        calculated_lambda = self.calculate_lambda_with_z(kinetic_energy, z_value)
                        self.inputs["Lambda"].SetValue(f"{calculated_lambda:.2f}")
                    except:
                        self.inputs["Lambda"].SetValue("3.4")  # fallback

                elif peak_type == "substrate":
                    self.inputs["Es"].SetValue(position)
                    self.inputs["Is"].SetValue(area)
                    self.inputs["Ss"].SetValue(rsf)
                break

    def calculate_lambda(self, peak_name, kinetic_energy):
        """Calculate EAL using Seah universal equation: L = (0.65 + 0.007E^0.93) / Z^0.38"""

        # Determine atomic number Z based on peak name/material
        Z = 6  # Default to carbon

        # Check for different elements based on peak name
        if any(element in peak_name.upper() for element in ['C', 'CARBON']):
            Z = 6  # Pure carbon
            # Check if it's organic/polymer (lower Z due to hydrogen)
            if any(keyword in peak_name.lower() for keyword in ['polymer', 'organic', 'hydrocarbon', 'ch']):
                Z = 4  # Average for hydrocarbon polymer
        elif 'SI' in peak_name.upper():
            Z = 14  # Silicon
        elif 'AL' in peak_name.upper():
            Z = 13  # Aluminum
        elif 'O' in peak_name.upper():
            Z = 8  # Oxygen
        elif 'N' in peak_name.upper():
            Z = 7  # Nitrogen

        # Seah universal equation for EAL (nm)
        # L = (0.65 + 0.007E^0.93) / Z^0.38
        E = kinetic_energy
        eal = (0.65 + 0.007 * (E ** 0.93)) / (Z ** 0.38)

        return eal

    def calculate_lambda_with_z(self, kinetic_energy, z_value):
        """Calculate EAL using Seah universal equation with specified Z"""
        # Seah universal equation for EAL (nm)
        # L = (0.65 + 0.007E^0.93) / Z^0.38
        E = kinetic_energy
        eal = (0.65 + 0.007 * (E ** 0.93)) / (z_value ** 0.38)
        return eal

    def calculate_lambda(self, peak_name, kinetic_energy):
        """Calculate EAL using Z from input field"""
        try:
            z_value = float(self.inputs["Z"].GetValue())
        except (ValueError, KeyError):
            z_value = self.determine_z_from_peak(peak_name)

        return self.calculate_lambda_with_z(kinetic_energy, z_value)



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

            # # Update result
            # self.result_label.SetLabel(f"THICKNESS: {t:.2f} nm")

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

    def plot_thickogram_OLD(self, xthicko, Ioisro, Eoes):
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
            ax.plot(2 * x, y, 'go', markersize=2)
        for x in axd:
            y = np.log(self.sinh(x / 2))
            ax.plot(2 * x, y, 'go', markersize=4)

        # Plot ratio lines
        ayp = [0.01, 0.02, 0.03, 0.04, 0.05, 0.06, 0.07, 0.08, 0.09, 0.1,
               0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1, 2, 3, 4, 5, 6,
               7, 8, 9, 10, 20, 30]

        for k in range(4, len(ayp) - 1):
            b_vals = np.arange(0.4, 3.1, 0.1)
            x_line = (4.2 - b_vals) / 0.6
            y_line = np.log(ayp[k] / 2) - self.func(x_line)
            ax.plot(2 * x_line, y_line, color='blue', linewidth=0.5)

        # Plot vertical lines
        bx = np.arange(0.4, 3.2, 0.2)
        for b in bx:
            x = (4.2 - b) / 0.6
            y0 = np.log(ayp[4] / 2) - self.func(x)
            y1 = np.log(ayp[-2] / 2) - self.func(x)
            ax.plot([2 * x, 2 * x], [y0, y1], color='blue', linewidth=0.5)

        # Plot horizontal reference lines
        for y in ayp:
            ax.plot([0, 0.18], [np.log(y / 2)] * 2, color='black', linewidth=0.5)

        # Add labels with smaller font
        ayd = [0.01, 0.1, 1.0, 10.0]
        for y in ayd:
            ax.text(-0.4, np.log(y / 2), f"{y}", verticalalignment='center', color='black', fontsize=6)

        for x in axd:
            y = np.log(self.sinh(x / 2))
            ax.text(2 * x + 0.15, y - 0.25, f"{x}", color='green', fontsize=6)

        # **ADD MISSING BLUE LABELS**
        # Blue ratio labels on left and right
        for i in range(1, len(ayd)):
            xl = (4.2 - 3) / 0.6
            xr = (4.2 - 0.4) / 0.6
            yl = np.log(ayd[i] / 2) - self.func(xl)
            yr = np.log(ayd[i] / 2) - self.func(xr)
            ax.text(2 * xl - 0.4, yl, f"{ayd[i]:.1f}", color='blue', fontsize=6, ha='center')
            ax.text(2 * xr + 0.2, yr, f"{ayd[i]:.1f}", color='blue', fontsize=6, ha='center')

        # Blue energy ratio labels at bottom
        for i in range(len(bx)):
            if i % 2 == 0:
                x = (4.2 - bx[i]) / 0.6
                y = np.log(ayp[4] / 2) - self.func(x)
                ax.text(2 * x + 0.05, y - 0.3, f"{bx[i]:.1f}", color='blue', fontsize=6, ha='center')

        # Plot result lines
        x = xthicko
        y0 = np.log(self.sinh(x / 2)) - 0.3
        y1 = np.log(self.sinh(x / 2)) + 0.3
        ax.plot([2 * x, 2 * x], [y0, y1], color='red', linewidth=2)
        ax.plot([0, 2 * x], [np.log(Ioisro / 2), np.log(Ioisro / 2) - self.func((4.2 - Eoes) / 0.6)],
                color='red', linewidth=2)

        # **FIX RED TEXT ALIGNMENT AND SIZE**
        ax.text(2 * x + 0.3, 2.4, f"{x:.4f}", color='red', fontsize=7, ha='left', va='center')

        # Set labels and formatting
        ax.set_title("Thickogram XPS Thickness Calculation", fontsize=10)
        ax.set_xlabel("t/L*cos(theta) (2x)", fontsize=8)
        ax.set_ylabel("log(Intensity Ratio)", fontsize=8)
        ax.grid(True, linestyle='--', alpha=0.3)
        ax.tick_params(labelsize=7)

        # Remove outer axes for clean look
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.tick_params(left=False, bottom=False, labelleft=False, labelbottom=False)

        # Adjust layout to fit smaller figure
        self.figure.tight_layout(pad=0.5)
        self.canvas.draw()

    def plot_thickogram(self, xthicko, Ioisro, Eoes):
        """Plot the thickogram"""
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.set_facecolor('white')

        # Plot main curve
        x_vals = np.linspace(0.01, 6, 600)
        y_vals = np.log(self.sinh(x_vals / 2))
        ax.plot(2 * x_vals, y_vals, color='green', linewidth=1)

        # Plot grid points with perpendicular lines
        axp = np.arange(0.1, 6.1, 0.1)
        axd = [0.1, 0.2, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0]

        # Sub-units (every 0.1) as short perpendicular lines
        for x in axp:
            if x not in axd:  # Only for sub-units, not major divisions
                y = np.log(self.sinh(x / 2))

                # Calculate slope of curve at this point: dy/dx = (1/2) * coth(x/2)
                slope = 0.5 * (1.0 / np.tanh(x / 2))

                # Perpendicular direction vector (normalized)
                perp_dx = -slope
                perp_dy = 1
                norm = np.sqrt(perp_dx ** 2 + perp_dy ** 2)
                perp_dx /= norm
                perp_dy /= norm

                # Create short perpendicular line
                length = 0.05
                x1 = 2 * x - length * perp_dx
                y1 = y - length * perp_dy
                x2 = 2 * x + length * perp_dx
                y2 = y + length * perp_dy

                ax.plot([x1, x2], [y1, y2], color='green', linewidth=0.5)

        # Major divisions as longer perpendicular lines
        for x in axd:
            y = np.log(self.sinh(x / 2))

            # Calculate slope of curve at this point
            slope = 0.5 * (1.0 / np.tanh(x / 2))

            # Perpendicular direction vector (normalized)
            perp_dx = -slope
            perp_dy = 1
            norm = np.sqrt(perp_dx ** 2 + perp_dy ** 2)
            perp_dx /= norm
            perp_dy /= norm

            # Create longer perpendicular line
            length = 0.1
            x1 = 2 * x - length * perp_dx
            y1 = y - length * perp_dy
            x2 = 2 * x + length * perp_dx
            y2 = y + length * perp_dy

            ax.plot([x1, x2], [y1, y2], color='green', linewidth=1)

        # for x in axp:
        #     y = np.log(self.sinh(x / 2))
        #     ax.plot(2 * x, y, 'go', markersize=2)
        # for x in axd:
        #     y = np.log(self.sinh(x / 2))
        #     ax.plot(2 * x, y, 'go', markersize=4)

        # Plot ratio lines
        ayp = [0.01, 0.02, 0.03, 0.04, 0.05, 0.06, 0.07, 0.08, 0.09, 0.1,
               0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1, 2, 3, 4, 5, 6,
               7, 8, 9, 10]

        for k in range(4, len(ayp) - 1):
            b_vals = np.arange(0.4, 3.1, 0.1)
            x_line = (4.2 - b_vals) / 0.6
            y_line = np.log(ayp[k] / 2) - self.func(x_line)
            ax.plot(2 * x_line, y_line, color='blue', linewidth=0.5)

        # Plot vertical lines
        bx = np.arange(0.4, 3.2, 0.2)
        for b in bx:
            x = (4.2 - b) / 0.6
            y0 = np.log(ayp[4] / 2) - self.func(x)
            y1 = np.log(ayp[-2] / 2) - self.func(x)
            ax.plot([2 * x, 2 * x], [y0, y1], color='blue', linewidth=0.5)

        # Plot horizontal reference lines
        for y in ayp:
            ax.plot([0, 0.18], [np.log(y / 2)] * 2, color='black', linewidth=0.5)

        # Add black line between 0.01 and 10 on Y-axis
        ax.plot([0, 0], [np.log(0.01 / 2), np.log(10 / 2)], color='black', linewidth=2)

        # Add labels with smaller font
        ayd = [0.01, 0.1, 1.0, 10.0]
        for y in ayd:
            ax.text(-0.6, np.log(y / 2), f"{y}", verticalalignment='center', color='black', fontsize=8)

        for x in axd:
            y = np.log(self.sinh(x / 2))
            ax.text(2 * x + 0.15, y - 0.25, f"{x}", color='green', fontsize=8)


        # Blue ratio labels on left and right
        for i in range(1, len(ayd)):
            xl = (4.2 - 3) / 0.6
            xr = (4.2 - 0.4) / 0.6
            yl = np.log(ayd[i] / 2) - self.func(xl)
            yr = np.log(ayd[i] / 2) - self.func(xr)
            ax.text(2 * xl - 0.4, yl, f"{ayd[i]:.1f}", color='blue', fontsize=8, ha='center')
            ax.text(2 * xr + 0.2, yr, f"{ayd[i]:.1f}", color='blue', fontsize=8, ha='center')

        # Blue energy ratio labels at bottom
        for i in range(len(bx)):
            if i % 2 == 0:
                x = (4.2 - bx[i]) / 0.6
                y = np.log(ayp[4] / 2) - self.func(x)
                ax.text(2 * x + 0.05, y - 0.3, f"{bx[i]:.1f}", color='blue', fontsize=8, ha='center')

        # CORRECTED: Plot result lines A-B-C
        # Point A: Intensity ratio on left Y-axis
        A_x = 0
        A_y = np.log(Ioisro / 2)

        # Point B: At energy ratio position, with Y calculated from blue curve equation
        B_x = 2 * ((4.2 - Eoes) / 0.6)
        B_y = np.log(Ioisro / 2) - self.func((4.2 - Eoes) / 0.6)  # Blue curve Y-coordinate

        # Point C: Intersection with green curve at thickness
        C_x = 2 * xthicko
        C_y = np.log(self.sinh(xthicko / 2))

        # Draw the line from A to B to C
        ax.plot([A_x, B_x, C_x], [A_y, B_y, C_y], color='red', linewidth=1)

        # Mark the points
        ax.plot(A_x, A_y, 'ro', markersize=3)  # Point A
        ax.plot(B_x, B_y, 'ro', markersize=3)  # Point B
        ax.plot(C_x, C_y, 'ro', markersize=3)  # Point C

        # Add labels
        ax.text(A_x - 0.4, A_y + 0.1, 'A', color='red', fontsize=10, fontweight='bold')
        ax.text(B_x + 0.2, B_y + 0.2, 'B', color='red', fontsize=10, fontweight='bold')
        ax.text(C_x + 0.0, C_y + 0.2, 'C', color='red', fontsize=10, fontweight='bold')

        # Calculate meaningful physical values for display
        # A: Intensity ratio at point A
        A_intensity = Ioisro  # This is the intensity ratio (Io/So)/(Is/Ss)

        # B: Energy ratio and same intensity ratio as A
        B_energy = Eoes  # Energy ratio Eo/Es
        B_intensity = A_intensity  # Should be same as A

        # C: Thickness parameter (t/λ*cos(θ))
        C_thickness_param = xthicko

        # Calculate actual thickness using the formula: thickness = C * λ * cos(θ)
        lambda_value = float(self.inputs['Lambda'].GetValue())  # Get lambda from input
        theta_degrees = float(self.inputs['Theta'].GetValue())  # Get theta from input
        theta_radians = np.radians(theta_degrees)
        actual_thickness = C_thickness_param * lambda_value * np.cos(theta_radians)

        # Display all results with correct values
        result_y_start = -6
        line_spacing = 0.3
        X_start = -0

        ax.text(X_start, result_y_start, f"A = [{A_intensity:.2f}]",
                color='red', fontsize=7, ha='left', va='center')

        ax.text(X_start, result_y_start - line_spacing, f"B = [{B_energy:.2f}, {B_intensity:.2f}]",
                color='red', fontsize=7, ha='left', va='center')

        ax.text(X_start, result_y_start - 2 * line_spacing, f"C = [{C_thickness_param:.2f}]",
                color='red', fontsize=7, ha='left', va='center')

        ax.text(X_start, result_y_start - 3 * line_spacing, f"Thickness = {actual_thickness:.2f} nm",
                color='red', fontsize=8, ha='left', va='center', weight='bold')

        # Add axis labels in specific positions
        # t/L*cos(theta) next to green plot
        ax.text(8, 1.5, r't/λ*cos(θ)', color='green', fontsize=10, fontweight='bold', rotation=10)

        # Eo/Es next to x-axis of blue plot
        ax.text(8, -8.5, r'E$_o$/E$_s$', color='blue', fontsize=10, fontweight='bold', ha='center')

        # (Io/So)/(Is/Ss) next to blue plot
        ax.text(2.7, -4.5, r'(I$_o$/S$_o$)/(I$_s$/S$_s$)', color='blue', fontsize=10, fontweight='bold', rotation=90,
                va='center')

        # Set labels and formatting
        # ax.set_title("Thickogram XPS Thickness Calculation", fontsize=10)
        # ax.set_xlabel("t/L*cos(theta) (2x)", fontsize=10)
        ax.set_ylabel(r'(I$_o$/S$_o$)/(I$_s$/S$_s$)', fontsize=10)
        ax.grid(True, linestyle='--', alpha=0.3)
        ax.tick_params(labelsize=7)

        # Remove outer axes for clean look
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.tick_params(left=False, bottom=False, labelleft=False, labelbottom=False)

        # Adjust layout to fit smaller figure
        self.figure.tight_layout(pad=0.5)
        self.canvas.draw()

    def plot_empty_thickogram(self):
        """Plot the empty thickogram without results"""
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.set_facecolor('white')

        # Plot main curve
        x_vals = np.linspace(0.01, 6, 600)
        y_vals = np.log(self.sinh(x_vals / 2))
        ax.plot(2 * x_vals, y_vals, color='green', linewidth=1)

        # Plot grid points with perpendicular lines
        axp = np.arange(0.1, 6.1, 0.1)
        axd = [0.1, 0.2, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0]

        # Sub-units (every 0.1) as short perpendicular lines
        for x in axp:
            if x not in axd:
                y = np.log(self.sinh(x / 2))
                slope = 0.5 * (1.0 / np.tanh(x / 2))
                perp_dx = -slope
                perp_dy = 1
                norm = np.sqrt(perp_dx ** 2 + perp_dy ** 2)
                perp_dx /= norm
                perp_dy /= norm
                length = 0.05
                x1 = 2 * x - length * perp_dx
                y1 = y - length * perp_dy
                x2 = 2 * x + length * perp_dx
                y2 = y + length * perp_dy
                ax.plot([x1, x2], [y1, y2], color='green', linewidth=0.5)

        # Major divisions as longer perpendicular lines
        for x in axd:
            y = np.log(self.sinh(x / 2))
            slope = 0.5 * (1.0 / np.tanh(x / 2))
            perp_dx = -slope
            perp_dy = 1
            norm = np.sqrt(perp_dx ** 2 + perp_dy ** 2)
            perp_dx /= norm
            perp_dy /= norm
            length = 0.1
            x1 = 2 * x - length * perp_dx
            y1 = y - length * perp_dy
            x2 = 2 * x + length * perp_dx
            y2 = y + length * perp_dy
            ax.plot([x1, x2], [y1, y2], color='green', linewidth=1)

        # Plot ratio lines
        ayp = [0.01, 0.02, 0.03, 0.04, 0.05, 0.06, 0.07, 0.08, 0.09, 0.1,
               0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1, 2, 3, 4, 5, 6,
               7, 8, 9, 10]

        for k in range(4, len(ayp) - 1):
            b_vals = np.arange(0.4, 3.1, 0.1)
            x_line = (4.2 - b_vals) / 0.6
            y_line = np.log(ayp[k] / 2) - self.func(x_line)
            ax.plot(2 * x_line, y_line, color='blue', linewidth=0.5)

        # Plot vertical lines
        bx = np.arange(0.4, 3.2, 0.2)
        for b in bx:
            x = (4.2 - b) / 0.6
            y0 = np.log(ayp[4] / 2) - self.func(x)
            y1 = np.log(ayp[-2] / 2) - self.func(x)
            ax.plot([2 * x, 2 * x], [y0, y1], color='blue', linewidth=0.5)

        # Plot horizontal reference lines
        for y in ayp:
            ax.plot([0, 0.18], [np.log(y / 2)] * 2, color='black', linewidth=0.5)

        # Add black line between 0.01 and 10 on Y-axis
        ax.plot([0, 0], [np.log(0.01 / 2), np.log(10 / 2)], color='black', linewidth=2)

        # Add labels
        ayd = [0.01, 0.1, 1.0, 10.0]
        for y in ayd:
            ax.text(-0.6, np.log(y / 2), f"{y}", verticalalignment='center', color='black', fontsize=8)

        for x in axd:
            y = np.log(self.sinh(x / 2))
            ax.text(2 * x + 0.15, y - 0.25, f"{x}", color='green', fontsize=8)

        # Blue ratio labels on left and right
        for i in range(1, len(ayd)):
            xl = (4.2 - 3) / 0.6
            xr = (4.2 - 0.4) / 0.6
            yl = np.log(ayd[i] / 2) - self.func(xl)
            yr = np.log(ayd[i] / 2) - self.func(xr)
            ax.text(2 * xl - 0.4, yl, f"{ayd[i]:.1f}", color='blue', fontsize=8, ha='center')
            ax.text(2 * xr + 0.2, yr, f"{ayd[i]:.1f}", color='blue', fontsize=8, ha='center')

        # Blue energy ratio labels at bottom
        for i in range(len(bx)):
            if i % 2 == 0:
                x = (4.2 - bx[i]) / 0.6
                y = np.log(ayp[4] / 2) - self.func(x)
                ax.text(2 * x + 0.05, y - 0.3, f"{bx[i]:.1f}", color='blue', fontsize=8, ha='center')

        # Add axis labels in specific positions
        ax.text(8, 1.5, r't/λ*cos(θ)', color='green', fontsize=10, fontweight='bold', rotation=18)
        ax.text(8, -8.5, r'E$_o$/E$_s$', color='blue', fontsize=10, fontweight='bold', ha='center')
        ax.text(2.7, -4.5, r'(I$_o$/S$_o$)/(I$_s$/S$_s$)', color='blue', fontsize=10, fontweight='bold', rotation=90,
                va='center')

        # Set labels and formatting
        ax.set_ylabel(r'(I$_o$/S$_o$)/(I$_s$/S$_s$)', fontsize=10)
        ax.grid(True, linestyle='--', alpha=0.3)
        ax.tick_params(labelsize=7)

        # Remove outer axes for clean look
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.tick_params(left=False, bottom=False, labelleft=False, labelbottom=False)

        # Adjust layout to fit smaller figure
        self.figure.tight_layout(pad=0.5)
        self.canvas.draw()

    def on_about(self, event):
        """Show about dialog"""
        about_text = """THICKOGRAM XPS THICKNESS CALCULATOR
        
        Updated from DaveXPS v0.1

    ═══════════════════════════════════════════════════════════════

    THEORETICAL FOUNDATION:
    This calculator implements a powerful XPS analysis technique pioneered by Peter J. Cumpson 
    at Britain's National Physical Laboratory. The method enables precise measurement of thin 
    film thickness when overlayers possess distinct chemical compositions from their underlying 
    substrates.

    METHODOLOGY OVERVIEW:
    The approach utilizes intensity ratios between overlayer and substrate photoelectron peaks, 
    combined with thickogram visualization developed by Prof. Jong Wan Lee at Hallym University, 
    Korea. This dual approach provides both numerical precision and intuitive graphical analysis.

    TECHNICAL ADVANTAGES:
    ✓ Surface contaminants like adventitious carbon do not compromise measurements
    ✓ Instrumental response factors automatically cancel during ratio calculations  
    ✓ Mathematical simplicity without sacrificing analytical rigor
    ✓ Consistent performance across varied thickness ranges

    EXPERIMENTAL REQUIREMENTS:
    Analysis Geometry: Photoelectron collection angles from 30° to 90° (optimal ~45°)
    Energy Range: Photoelectron kinetic energies exceeding 500 eV recommended
    Measurement Precision: Typically ±10% uncertainty from attenuation length estimations

    INPUT SPECIFICATIONS:
    Peak Intensities: Io (overlayer), Is (substrate) - areas or heights
    Sensitivity Factors: So (overlayer RSF), Ss (substrate RSF)  
    Photoelectron Energies: Eo, Es (kinetic energies of respective peaks)
    Geometric Parameters: θ (emission angle), λo (electron attenuation length)

    SCIENTIFIC RECOGNITION:
    This work builds upon foundational research by Peter J. Cumpson (NPL) and visualization 
    innovations by Jong Wan Lee (Hallym University).

    PRIMARY REFERENCE:
    P.J. Cumpson, "Angle-resolved XPS and AES: depth-resolution limits and a general 
    comparison of properties of depth-profile reconstruction methods"
    Surface and Interface Analysis, Volume 29, Pages 403-406 (2000)

    Secondary Reference:
    J.W. Lee, "Thickogram: A web-based tool for XPS overlayer thickness calculation"
    New Physics: Sae Mulli, Volume 65, Pages 1165-1174 (2015)"""

        dlg = wx.lib.dialogs.ScrolledMessageDialog(self, about_text, "About Thickogram Calculator",
                                                   size=(650, 550))
        dlg.ShowModal()
        dlg.Destroy()

    def on_export_png(self, event):
        """Export plot as PNG image"""
        with wx.FileDialog(self, "Save PNG file",
                           wildcard="PNG files (*.png)|*.png",
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:

            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return

            pathname = fileDialog.GetPath()
            try:
                self.figure.savefig(pathname, format='png', dpi=300, bbox_inches='tight',
                                    facecolor='white', edgecolor='none')
                wx.MessageBox(f"Plot saved as PNG: {pathname}", "Export Successful",
                              wx.OK | wx.ICON_INFORMATION)
            except Exception as e:
                wx.MessageBox(f"Error saving PNG: {str(e)}", "Export Error",
                              wx.OK | wx.ICON_ERROR)

    def on_export_svg(self, event):
        """Export plot as SVG image"""
        with wx.FileDialog(self, "Save SVG file",
                           wildcard="SVG files (*.svg)|*.svg",
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:

            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return

            pathname = fileDialog.GetPath()
            try:
                self.figure.savefig(pathname, format='svg', bbox_inches='tight',
                                    facecolor='white', edgecolor='none')
                wx.MessageBox(f"Plot saved as SVG: {pathname}", "Export Successful",
                              wx.OK | wx.ICON_INFORMATION)
            except Exception as e:
                wx.MessageBox(f"Error saving SVG: {str(e)}", "Export Error",
                              wx.OK | wx.ICON_ERROR)

    def determine_z_from_peak(self, peak_name):
        """Determine atomic number Z from peak name - comprehensive periodic table"""
        peak_upper = peak_name.upper()

        # Comprehensive atomic number mapping
        element_mapping = {
            # Light elements (Z=1-10)
            1: ['H1S', 'H 1S', 'HYDROGEN'],
            2: ['HE1S', 'HE 1S', 'HELIUM'],
            3: ['LI1S', 'LI 1S', 'LITHIUM'],
            4: ['BE1S', 'BE 1S', 'BERYLLIUM'],
            5: ['B1S', 'B 1S', 'BORON'],
            6: ['C1S', 'C 1S', 'CARBON'],
            7: ['N1S', 'N 1S', 'NITROGEN'],
            8: ['O1S', 'O 1S', 'OXYGEN'],
            9: ['F1S', 'F 1S', 'FLUORINE'],
            10: ['NE1S', 'NE 1S', 'NEON'],

            # Period 3 (Z=11-18)
            11: ['NA1S', 'NA 1S', 'SODIUM'],
            12: ['MG1S', 'MG 1S', 'MG2P', 'MG 2P', 'MAGNESIUM'],
            13: ['AL2P', 'AL 2P', 'ALUMINUM', 'ALUMINIUM'],
            14: ['SI2P', 'SI 2P', 'SILICON'],
            15: ['P2P', 'P 2P', 'PHOSPHORUS'],
            16: ['S2P', 'S 2P', 'SULFUR', 'SULPHUR'],
            17: ['CL2P', 'CL 2P', 'CHLORINE'],
            18: ['AR2P', 'AR 2P', 'ARGON'],

            # Period 4 (Z=19-36)
            19: ['K2P', 'K 2P', 'POTASSIUM'],
            20: ['CA2P', 'CA 2P', 'CALCIUM'],
            21: ['SC2P', 'SC 2P', 'SCANDIUM'],
            22: ['TI2P', 'TI 2P', 'TITANIUM'],
            23: ['V2P', 'V 2P', 'VANADIUM'],
            24: ['CR2P', 'CR 2P', 'CHROMIUM'],
            25: ['MN2P', 'MN 2P', 'MANGANESE'],
            26: ['FE2P', 'FE 2P', 'IRON'],
            27: ['CO2P', 'CO 2P', 'COBALT'],
            28: ['NI2P', 'NI 2P', 'NICKEL'],
            29: ['CU2P', 'CU 2P', 'COPPER'],
            30: ['ZN2P', 'ZN 2P', 'ZINC'],
            31: ['GA2P', 'GA 2P', 'GALLIUM'],
            32: ['GE3D', 'GE 3D', 'GERMANIUM'],
            33: ['AS3D', 'AS 3D', 'ARSENIC'],
            34: ['SE3D', 'SE 3D', 'SELENIUM'],
            35: ['BR3D', 'BR 3D', 'BROMINE'],
            36: ['KR3D', 'KR 3D', 'KRYPTON'],

            # Period 5 (Z=37-54)
            37: ['RB3D', 'RB 3D', 'RUBIDIUM'],
            38: ['SR3D', 'SR 3D', 'STRONTIUM'],
            39: ['Y3D', 'Y 3D', 'YTTRIUM'],
            40: ['ZR3D', 'ZR 3D', 'ZIRCONIUM'],
            41: ['NB3D', 'NB 3D', 'NIOBIUM'],
            42: ['MO3D', 'MO 3D', 'MOLYBDENUM'],
            43: ['TC3D', 'TC 3D', 'TECHNETIUM'],
            44: ['RU3D', 'RU 3D', 'RUTHENIUM'],
            45: ['RH3D', 'RH 3D', 'RHODIUM'],
            46: ['PD3D', 'PD 3D', 'PALLADIUM'],
            47: ['AG3D', 'AG 3D', 'SILVER'],
            48: ['CD3D', 'CD 3D', 'CADMIUM'],
            49: ['IN3D', 'IN 3D', 'INDIUM'],
            50: ['SN3D', 'SN 3D', 'TIN'],
            51: ['SB3D', 'SB 3D', 'ANTIMONY'],
            52: ['TE3D', 'TE 3D', 'TELLURIUM'],
            53: ['I3D', 'I 3D', 'IODINE'],
            54: ['XE3D', 'XE 3D', 'XENON'],

            # Period 6 (Z=55-86)
            55: ['CS3D', 'CS 3D', 'CESIUM', 'CAESIUM'],
            56: ['BA3D', 'BA 3D', 'BARIUM'],
            57: ['LA3D', 'LA 3D', 'LANTHANUM'],
            58: ['CE3D', 'CE 3D', 'CERIUM'],
            59: ['PR3D', 'PR 3D', 'PRASEODYMIUM'],
            60: ['ND3D', 'ND 3D', 'NEODYMIUM'],
            61: ['PM3D', 'PM 3D', 'PROMETHIUM'],
            62: ['SM3D', 'SM 3D', 'SAMARIUM'],
            63: ['EU3D', 'EU 3D', 'EUROPIUM'],
            64: ['GD3D', 'GD 3D', 'GADOLINIUM'],
            65: ['TB3D', 'TB 3D', 'TERBIUM'],
            66: ['DY3D', 'DY 3D', 'DYSPROSIUM'],
            67: ['HO3D', 'HO 3D', 'HOLMIUM'],
            68: ['ER3D', 'ER 3D', 'ERBIUM'],
            69: ['TM3D', 'TM 3D', 'THULIUM'],
            70: ['YB3D', 'YB 3D', 'YTTERBIUM'],
            71: ['LU3D', 'LU 3D', 'LUTETIUM'],
            72: ['HF4F', 'HF 4F', 'HAFNIUM'],
            73: ['TA4F', 'TA 4F', 'TANTALUM'],
            74: ['W4F', 'W 4F', 'TUNGSTEN'],
            75: ['RE4F', 'RE 4F', 'RHENIUM'],
            76: ['OS4F', 'OS 4F', 'OSMIUM'],
            77: ['IR4F', 'IR 4F', 'IRIDIUM'],
            78: ['PT4F', 'PT 4F', 'PLATINUM'],
            79: ['AU4F', 'AU 4F', 'GOLD'],
            80: ['HG4F', 'HG 4F', 'MERCURY'],
            81: ['TL4F', 'TL 4F', 'THALLIUM'],
            82: ['PB4F', 'PB 4F', 'LEAD'],
            83: ['BI4F', 'BI 4F', 'BISMUTH'],
            84: ['PO4F', 'PO 4F', 'POLONIUM'],
            85: ['AT4F', 'AT 4F', 'ASTATINE'],
            86: ['RN4F', 'RN 4F', 'RADON'],

            # Period 7 (Z=87-118)
            87: ['FR4F', 'FR 4F', 'FRANCIUM'],
            88: ['RA4F', 'RA 4F', 'RADIUM'],
            89: ['AC4F', 'AC 4F', 'ACTINIUM'],
            90: ['TH4F', 'TH 4F', 'THORIUM'],
            91: ['PA4F', 'PA 4F', 'PROTACTINIUM'],
            92: ['U4F', 'U 4F', 'URANIUM'],
            93: ['NP4F', 'NP 4F', 'NEPTUNIUM'],
            94: ['PU4F', 'PU 4F', 'PLUTONIUM'],
            95: ['AM4F', 'AM 4F', 'AMERICIUM'],
            96: ['CM4F', 'CM 4F', 'CURIUM'],
            97: ['BK4F', 'BK 4F', 'BERKELIUM'],
            98: ['CF4F', 'CF 4F', 'CALIFORNIUM'],
            99: ['ES4F', 'ES 4F', 'EINSTEINIUM'],
            100: ['FM4F', 'FM 4F', 'FERMIUM'],
            101: ['MD4F', 'MD 4F', 'MENDELEVIUM'],
            102: ['NO4F', 'NO 4F', 'NOBELIUM'],
            103: ['LR4F', 'LR 4F', 'LAWRENCIUM'],
            104: ['RF4F', 'RF 4F', 'RUTHERFORDIUM'],
            105: ['DB4F', 'DB 4F', 'DUBNIUM'],
            106: ['SG4F', 'SG 4F', 'SEABORGIUM'],
            107: ['BH4F', 'BH 4F', 'BOHRIUM'],
            108: ['HS4F', 'HS 4F', 'HASSIUM'],
            109: ['MT4F', 'MT 4F', 'MEITNERIUM'],
            110: ['DS4F', 'DS 4F', 'DARMSTADTIUM'],
            111: ['RG4F', 'RG 4F', 'ROENTGENIUM'],
            112: ['CN4F', 'CN 4F', 'COPERNICIUM'],
            113: ['NH4F', 'NH 4F', 'NIHONIUM'],
            114: ['FL4F', 'FL 4F', 'FLEROVIUM'],
            115: ['MC4F', 'MC 4F', 'MOSCOVIUM'],
            116: ['LV4F', 'LV 4F', 'LIVERMORIUM'],
            117: ['TS4F', 'TS 4F', 'TENNESSINE'],
            118: ['OG4F', 'OG 4F', 'OGANESSON']
        }

        # Search through all elements
        for z, identifiers in element_mapping.items():
            for identifier in identifiers:
                if identifier in peak_upper:
                    # Special case for carbon - check for organic/polymer context
                    if z == 6:
                        if any(keyword in peak_name.lower() for keyword in
                               ['polymer', 'organic', 'hydrocarbon', 'ch', 'pet', 'pmma', 'ps', 'pe']):
                            return 4  # Effective Z for hydrocarbon polymer
                    return z

        # Default to carbon if no match found
        return 6

    def on_z_changed(self, event):
        """Update Lambda when Z value changes"""
        try:
            # Get current values
            z_value = float(self.inputs["Z"].GetValue())
            be_overlayer = float(self.inputs["Eo"].GetValue())
            esource = float(self.inputs["Esource"].GetValue())
            kinetic_energy = esource - be_overlayer

            # Recalculate Lambda with new Z
            new_lambda = self.calculate_lambda_with_z(kinetic_energy, z_value)
            self.inputs["Lambda"].SetValue(f"{new_lambda:.2f}")
        except ValueError:
            pass  # Ignore invalid inputs during typing
