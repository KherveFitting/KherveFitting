import wx
import wx.lib.agw.floatspin as FS
import numpy as np
from scipy.special import erf
import lmfit

# Install and import lmfitxps for better Fermi edge fitting
try:
    from lmfitxps import models as xps_models

    HAS_LMFITXPS = True
except ImportError:
    HAS_LMFITXPS = False
    print("lmfitxps not found. Install with: pip install lmfitxps")


class VB_measurements(wx.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, title="VB Measurements",
                         size=(600, 430),
                         style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX))
        self.parent = parent  # Main frame

        # Store reference in parent
        self.parent.vb_measurements_window = self

        self.controller = controller

        # Fit results storage
        self.fermi_fit_result = None
        self.thermal_fit_line = None
        self.vbm_lines = []

        self.InitUI()

        # Center window on main frame and keep on top
        self.center_on_parent()
        self.SetWindowStyle(self.GetWindowStyle() | wx.STAY_ON_TOP)

        # Force vLines to be visible for VB measurements
        self.force_vlines_visible()

        # Bind close event to cleanup
        self.Bind(wx.EVT_CLOSE, self.on_close)

        self.Centre()
        self.Show()

    def center_on_parent(self):
        """Center this window on the parent frame"""
        if self.parent:
            parent_pos = self.parent.GetPosition()
            parent_size = self.parent.GetSize()
            my_size = self.GetSize()

            # Calculate center position
            x = parent_pos.x + (parent_size.width - my_size.width) // 2
            y = parent_pos.y + (parent_size.height - my_size.height) // 2

            # Ensure window is on screen
            x = max(0, x)
            y = max(0, y)

            self.SetPosition((x, y))


    def force_vlines_visible_OLD(self):
        """Force vLines to be visible specifically for VB measurements at 10% and 90% of plot range"""
        # Initialize vLines if they don't exist
        if self.parent.vline1 is None or self.parent.vline2 is None:
            self.parent.initialize_or_restore_background_vlines()

        # Position vLines at 10% and 90% of plot range for VB measurements
        plot_range = self.get_vb_plot_range_positions()
        if plot_range:
            low_pos, high_pos = plot_range

            # Set vLines to 10% and 90% positions
            if self.parent.vline1 is not None:
                self.parent.vline1.set_xdata([low_pos, low_pos])
                self.parent.vline1.set_visible(True)
            if self.parent.vline2 is not None:
                self.parent.vline2.set_xdata([high_pos, high_pos])
                self.parent.vline2.set_visible(True)

            # Update VBM controls to match vLine positions (with .2f format)
            self.vbm_edge_ctrl.SetValue(float(f"{low_pos:.2f}"))
            self.vbm_bg_center_ctrl.SetValue(float(f"{high_pos:.2f}"))

        # Override visibility logic by setting them directly visible
        if hasattr(self.parent, 'vline1_text') and self.parent.vline1_text is not None:
            self.parent.vline1_text.set_visible(True)
        if hasattr(self.parent, 'vline2_text') and self.parent.vline2_text is not None:
            self.parent.vline2_text.set_visible(True)

        # Enable background interaction for dragging
        self.parent.background_tab_selected = True

        # Add text labels for vlines
        self.add_vline_text_labels()

        # Force canvas redraw
        self.parent.canvas.draw_idle()

    def force_vlines_visible(self):
        """Force vLines to be visible specifically for VB measurements at 10% and 90% of plot range"""

        # Check if vLines are valid and attached to current axis
        vlines_are_valid = (
                self.parent.vline1 is not None and
                self.parent.vline2 is not None and
                hasattr(self.parent, 'ax') and
                self.parent.vline1 in self.parent.ax.get_children() and
                self.parent.vline2 in self.parent.ax.get_children()
        )

        # If vLines don't exist or are invalid, clean up and recreate them
        if not vlines_are_valid:
            # Clean up any invalid vLine references
            if self.parent.vline1 is not None:
                try:
                    self.parent.vline1.remove()
                except:
                    pass
                self.parent.vline1 = None

            if self.parent.vline2 is not None:
                try:
                    self.parent.vline2.remove()
                except:
                    pass
                self.parent.vline2 = None

            # Clean up any invalid text labels
            if hasattr(self.parent, 'vline1_text') and self.parent.vline1_text is not None:
                try:
                    self.parent.vline1_text.remove()
                except:
                    pass
                self.parent.vline1_text = None

            if hasattr(self.parent, 'vline2_text') and self.parent.vline2_text is not None:
                try:
                    self.parent.vline2_text.remove()
                except:
                    pass
                self.parent.vline2_text = None

            # Create new vLines at 10% and 90% of plot range
            plot_range = self.get_vb_plot_range_positions()
            if plot_range:
                low_pos, high_pos = plot_range

                # Create new vLines
                self.parent.vline1 = self.parent.ax.axvline(low_pos, color='r', linestyle='--', alpha=0.7)
                self.parent.vline2 = self.parent.ax.axvline(high_pos, color='r', linestyle='--', alpha=0.7)

                # Update VBM controls to match vLine positions (with .2f format)
                self.vbm_edge_ctrl.SetValue(float(f"{low_pos:.2f}"))
                self.vbm_bg_center_ctrl.SetValue(float(f"{high_pos:.2f}"))
        else:
            # VLines are valid, just reposition them to 10% and 90%
            plot_range = self.get_vb_plot_range_positions()
            if plot_range:
                low_pos, high_pos = plot_range

                # Set vLines to 10% and 90% positions
                self.parent.vline1.set_xdata([low_pos, low_pos])
                self.parent.vline2.set_xdata([high_pos, high_pos])

                # Update VBM controls to match vLine positions (with .2f format)
                self.vbm_edge_ctrl.SetValue(float(f"{low_pos:.2f}"))
                self.vbm_bg_center_ctrl.SetValue(float(f"{high_pos:.2f}"))

        # Ensure vLines are visible
        if self.parent.vline1 is not None:
            self.parent.vline1.set_visible(True)
        if self.parent.vline2 is not None:
            self.parent.vline2.set_visible(True)

        # Enable background interaction for dragging
        self.parent.background_tab_selected = True

        # Add text labels for vlines
        self.add_vline_text_labels()

        # Force canvas redraw
        self.parent.canvas.draw_idle()

    def get_vb_plot_range_positions(self):
        """Get 10% and 90% positions of current plot X-axis range for VB measurements."""
        try:
            # Get current X-axis limits from the plot
            if hasattr(self.parent, 'ax') and self.parent.ax:
                xlim = self.parent.ax.get_xlim()
                x_min, x_max = xlim

                # Calculate the range
                x_range = x_max - x_min

                # Calculate 10% and 90% positions
                low_pos = x_min + (0.1 * x_range)  # 10% from left
                high_pos = x_min + (0.9 * x_range)  # 90% from left

                # Ensure proper order (BE scale is usually decreasing)
                if low_pos > high_pos:
                    low_pos, high_pos = high_pos, low_pos

                return (low_pos, high_pos)

        except (AttributeError, Exception):
            # Fallback to data range if plot limits not available
            sheet_name = self.parent.sheet_combobox.GetValue()
            if sheet_name in self.parent.Data['Core levels']:
                x_values = self.parent.Data['Core levels'][sheet_name]['B.E.']
                if x_values:
                    x_min, x_max = min(x_values), max(x_values)
                    x_range = x_max - x_min
                    low_pos = x_min + (0.1 * x_range)
                    high_pos = x_min + (0.9 * x_range)

                    if low_pos > high_pos:
                        low_pos, high_pos = high_pos, low_pos

                    return (low_pos, high_pos)

        return None

    def InitUI(self):
        panel = wx.Panel(self)
        panel.SetBackgroundColour(wx.Colour(240, 225, 225))  # Orange background
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Left side - Controls
        left_panel = wx.Panel(panel)
        left_sizer = wx.BoxSizer(wx.VERTICAL)

        # Fermi Edge Section
        fermi_box = wx.StaticBox(left_panel, label="Fermi Edge Measurement")
        fermi_sizer = wx.StaticBoxSizer(fermi_box, wx.VERTICAL)

        fermi_btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        fit_btn = wx.Button(left_panel, label="Fit Fermi Edge")
        fit_btn.Bind(wx.EVT_BUTTON, self.OnFitFermiEdge)
        fermi_btn_sizer.Add(fit_btn, 1, wx.ALL, 5)

        remove_btn = wx.Button(left_panel, label="Remove Fermi")
        remove_btn.Bind(wx.EVT_BUTTON, self.OnRemoveFermi)
        fermi_btn_sizer.Add(remove_btn, 1, wx.ALL, 5)

        fermi_sizer.Add(fermi_btn_sizer, 0, wx.EXPAND)

        left_sizer.Add(fermi_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # VBM Section
        vbm_box = wx.StaticBox(left_panel, label="Valence Band Minimum / Cut-Off")
        vbm_sizer = wx.StaticBoxSizer(vbm_box, wx.VERTICAL)

        # Method selection (fixed to Linear Extrapolation only)
        self.vbm_method = wx.Choice(left_panel, choices=["Linear Extrapolation"])
        self.vbm_method.SetSelection(0)
        vbm_sizer.Add(wx.StaticText(left_panel, label="Method:"), 0, wx.ALL, 5)
        vbm_sizer.Add(self.vbm_method, 0, wx.EXPAND | wx.ALL, 5)

        # Center Edge (renamed from Energy Range)
        edge_sizer = wx.BoxSizer(wx.HORIZONTAL)
        edge_sizer.Add(wx.StaticText(left_panel, label="Center Edge (eV):"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        self.vbm_edge_ctrl = FS.FloatSpin(left_panel, value=2.0, min_val=-35.0, max_val=5000.0, increment=0.1, digits=2)
        edge_sizer.Add(self.vbm_edge_ctrl, 1, wx.ALL, 5)
        vbm_sizer.Add(edge_sizer, 0, wx.EXPAND)

        # Number of Points for linear fit
        points_sizer = wx.BoxSizer(wx.HORIZONTAL)
        points_sizer.Add(wx.StaticText(left_panel, label="Fit Points:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        self.vbm_points_ctrl = wx.SpinCtrl(left_panel)
        self.vbm_points_ctrl.SetRange(3, 100)
        self.vbm_points_ctrl.SetValue(10)
        points_sizer.Add(self.vbm_points_ctrl, 1, wx.ALL, 5)
        vbm_sizer.Add(points_sizer, 0, wx.EXPAND)

        # Background extrapolation checkbox
        self.vbm_bg_checkbox = wx.CheckBox(left_panel, label="Background Extrapolation")
        vbm_sizer.Add(self.vbm_bg_checkbox, 0, wx.ALL, 5)

        # Background center (initially disabled)
        bg_center_sizer = wx.BoxSizer(wx.HORIZONTAL)
        bg_center_sizer.Add(wx.StaticText(left_panel, label="BG Center (eV):"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        self.vbm_bg_center_ctrl = FS.FloatSpin(left_panel, value=-1.5, min_val=-35.0, max_val=5000.0, increment=0.1,
                                               digits=2)
        self.vbm_bg_center_ctrl.Enable(False)
        bg_center_sizer.Add(self.vbm_bg_center_ctrl, 1, wx.ALL, 5)
        vbm_sizer.Add(bg_center_sizer, 0, wx.EXPAND)

        # Background points (initially disabled)
        bg_points_sizer = wx.BoxSizer(wx.HORIZONTAL)
        bg_points_sizer.Add(wx.StaticText(left_panel, label="BG Points:"), 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        self.vbm_bg_points_ctrl = wx.SpinCtrl(left_panel)
        self.vbm_bg_points_ctrl.SetRange(3, 100)
        self.vbm_bg_points_ctrl.SetValue(10)
        self.vbm_bg_points_ctrl.Enable(False)
        bg_points_sizer.Add(self.vbm_bg_points_ctrl, 1, wx.ALL, 5)
        vbm_sizer.Add(bg_points_sizer, 0, wx.EXPAND)

        # Bind checkbox event
        self.vbm_bg_checkbox.Bind(wx.EVT_CHECKBOX, self.OnBgCheckbox)

        # Bind control changes to update vLines
        self.vbm_edge_ctrl.Bind(FS.EVT_FLOATSPIN, self.OnEdgeCenterChange)
        self.vbm_bg_center_ctrl.Bind(FS.EVT_FLOATSPIN, self.OnBgCenterChange)

        # VBM buttons
        vbm_btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        vbm_calc_btn = wx.Button(left_panel, label="Calculate VBM")
        vbm_calc_btn.Bind(wx.EVT_BUTTON, self.OnCalculateVBM)
        vbm_btn_sizer.Add(vbm_calc_btn, 1, wx.ALL, 5)

        cutoff_calc_btn = wx.Button(left_panel, label="Calculate Cut-Off")
        cutoff_calc_btn.Bind(wx.EVT_BUTTON, self.OnCalculateCutOff)
        vbm_btn_sizer.Add(cutoff_calc_btn, 1, wx.ALL, 5)

        vbm_sizer.Add(vbm_btn_sizer, 0, wx.EXPAND)

        # Remove button (handles both VBM and Cut-Off)
        remove_btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        remove_vbm_btn = wx.Button(left_panel, label="Remove VBM/Cut-Off")
        remove_vbm_btn.Bind(wx.EVT_BUTTON, self.OnRemoveVBM)
        remove_btn_sizer.Add(remove_vbm_btn, 1, wx.ALL, 5)

        vbm_sizer.Add(remove_btn_sizer, 0, wx.EXPAND)

        left_sizer.Add(vbm_sizer, 0, wx.EXPAND | wx.ALL, 5)

        left_panel.SetSizer(left_sizer)

        # Right side - Results
        right_panel = wx.Panel(panel)
        right_sizer = wx.BoxSizer(wx.VERTICAL)

        # Right side - Results (replace entire right panel section)
        right_panel = wx.Panel(panel)
        right_sizer = wx.BoxSizer(wx.VERTICAL)

        # Single Results Box
        result_box = wx.StaticBox(right_panel, label="Analysis Results")
        result_sizer = wx.StaticBoxSizer(result_box, wx.VERTICAL)
        self.results = wx.TextCtrl(right_panel, style=wx.TE_MULTILINE | wx.TE_READONLY, size=(300, 350))
        result_sizer.Add(self.results, 1, wx.EXPAND | wx.ALL, 5)
        right_sizer.Add(result_sizer, 1, wx.EXPAND | wx.ALL, 5)

        right_panel.SetSizer(right_sizer)

        # Add panels to main sizer
        main_sizer.Add(left_panel, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(right_panel, 1, wx.EXPAND | wx.ALL, 5)
        panel.SetSizer(main_sizer)


    def get_current_data(self):
        """Get current active sheet data"""
        sheet_name = self.parent.sheet_combobox.GetValue()
        if not sheet_name or 'Core levels' not in self.parent.Data or sheet_name not in self.parent.Data['Core levels']:
            return None, None

        return self.parent.x_values, self.parent.y_values

    def get_vline_range(self):
        """Get range from main frame vLines"""
        try:
            if self.parent.vline1 is not None and self.parent.vline2 is not None:
                vline1_pos = self.parent.vline1.get_xdata()[0]
                vline2_pos = self.parent.vline2.get_xdata()[0]
                return min(vline1_pos, vline2_pos), max(vline1_pos, vline2_pos)
        except:
            pass
        return None, None

    def OnFitFermiEdge(self, event):
        """Fit Fermi edge using lmfitxps FermiEdgeModel"""
        if not HAS_LMFITXPS:
            wx.MessageBox("lmfitxps required. Install with: pip install lmfitxps", "Error")
            return

        x_data, y_data = self.get_current_data()
        if x_data is None:
            wx.MessageBox("No active data available", "Error")
            return

        # Get range from vLines
        x_min, x_max = self.get_vline_range()
        if x_min is None:
            wx.MessageBox("Drag the red vLines to define fitting range", "Error")
            return

        # Select fitting range
        mask = (x_data >= x_min) & (x_data <= x_max)
        x_fit = x_data[mask]
        y_fit = y_data[mask]

        if len(x_fit) < 10:
            wx.MessageBox("Insufficient data points. Adjust vLines range.", "Error")
            return

        try:
            # Define sheet_name first
            sheet_name = self.parent.sheet_combobox.GetValue()

            # Clear previous fit results from plot
            self.clear_previous_fits()

            # Redraw base data and vLines
            self.parent.plot_manager.plot_data(self.parent)
            self.force_vlines_visible()

            # Use lmfitxps FermiEdgeModel
            fermi_model = xps_models.FermiEdgeModel(prefix='fermi_')
            const_model = lmfit.models.ConstantModel(prefix='const_')
            model = fermi_model + const_model

            # Let the model determine initial parameters automatically
            params = model.make_params()

            # Set reasonable bounds
            params['fermi_amplitude'].set(value=np.max(y_fit) - np.min(y_fit), min=0)
            params['fermi_center'].set(value=np.mean(x_fit), min=x_min, max=x_max)
            params['fermi_sigma'].set(value=0.05, min=0.001, max=0.5)  # ~50 meV instrumental
            params['fermi_kt'].set(value=0.026, min=0.001, max=0.2)  # ~300K thermal
            params['const_c'].set(value=np.min(y_fit), min=0)

            # Fit the model
            result = model.fit(y_fit, x=x_fit, params=params)
            self.fermi_fit_result = result

            # Extract fitted parameters
            center = result.params['fermi_center'].value
            center_err = result.params['fermi_center'].stderr if result.params['fermi_center'].stderr else 0
            sigma = result.params['fermi_sigma'].value * 1000  # Convert to meV
            kt_fitted = result.params['fermi_kt'].value * 1000  # Convert to meV
            amplitude = result.params['fermi_amplitude'].value
            background = result.params['const_c'].value

            # Calculate effective temperature from kT
            eff_temp = (kt_fitted / 1000) / 8.617333e-5

            # Calculate edge widths (DEFAULT: 16%-84%)
            width_16_84 = (kt_fitted / 1000) * np.log(25.5)  # 16%-84% width (DEFAULT)
            width_10_90 = (kt_fitted / 1000) * np.log(81)  # 10%-90% width
            width_20_80 = (kt_fitted / 1000) * np.log(16)  # 20%-80% width
            width_1_99 = (kt_fitted / 1000) * np.log(9801)  # 1%-99% width (practical full width)

            # Calculate fit quality
            ss_res = np.sum(result.residual ** 2)
            ss_tot = np.sum((y_fit - np.mean(y_fit)) ** 2)
            r_squared = 1 - (ss_res / ss_tot)

            # Display results (16%-84% as main result)
            results_text = f"""Fermi Edge Fit Results:
Model: lmfitxps FermiEdgeModel

Center Position: {center:.3f} ± {center_err:.3f} eV
Thermal Broadening (kT): {kt_fitted:.1f} meV
Instrumental Resolution (σ): {sigma:.1f} meV


Edge Widths:
10%-90%: {width_10_90 * 1000:.1f} meV
16%-84%: {width_16_84 * 1000:.1f} meV ** (default)                            
20%-80%: {width_20_80 * 1000:.1f} meV

Fit Parameters:
Step Height: {amplitude:.1f}
Background: {background:.1f}

Fit Quality:
R²: {r_squared:.4f}
χ²: {result.chisqr:.2f}
Reduced χ²: {result.redchi:.2f}

Note: 16%-84% is the standard thermal width measure."""

            self.results.SetValue(results_text)

            # Remove previous Fermi rows if they exist (same as D-parameter does)
            for row in range(self.parent.peak_params_grid.GetNumberRows() - 1, -1, -1):
                if self.parent.peak_params_grid.GetCellValue(row, 13) == "Fermi":
                    self.parent.peak_params_grid.DeleteRows(row, 2)

            # Add new rows (same as D-parameter)
            row = self.parent.peak_params_grid.GetNumberRows()
            self.parent.peak_params_grid.AppendRows(2)

            # Update grid values directly (same pattern as D-parameter)
            self.parent.peak_params_grid.SetCellValue(row, 0, chr(65 + row // 2))  # Peak letter
            self.parent.peak_params_grid.SetCellValue(row, 1, "Fermi")
                                                              # "\n\n\n"
                                                              # f" -Center Pos.: {center:.3f} eV\n"
                                                              # f" -Edge Width: {width_16_84:.3f} eV")  # Peak name
            self.parent.peak_params_grid.SetCellValue(row, 2, f"{center:.2f}")  # Position
            self.parent.peak_params_grid.SetCellValue(row, 3, f"{amplitude:.2f}")  # Height
            self.parent.peak_params_grid.SetCellValue(row, 4, f"{width_16_84:.3f}")  # FWHM (16%-84%)
            self.parent.peak_params_grid.SetCellValue(row, 5, "0.00")  # L/G
            self.parent.peak_params_grid.SetCellValue(row, 6, f"{amplitude * width_16_84:.2f}")  # Area
            self.parent.peak_params_grid.SetCellValue(row, 7, f"{sigma / 1000:.3f}")  # Sigma (in eV)
            self.parent.peak_params_grid.SetCellValue(row, 8, f"{kt_fitted / 1000:.3f}")  # Gamma (kT in eV)
            self.parent.peak_params_grid.SetCellValue(row, 9, "0.00")  # Skew
            self.parent.peak_params_grid.SetCellValue(row, 13, "Fermi")  # Fitting Model

            # Add this section to color the constraint row green:
            constraint_row = row + 1  # Constraint row is next row
            num_cols = self.parent.peak_params_grid.GetNumberCols()

            # Set constraint row background to green (same as other peaks)
            green_color = wx.Colour(200, 245, 228)  # Light green
            for col in range(num_cols):
                self.parent.peak_params_grid.SetCellBackgroundColour(constraint_row, col, green_color)

            # # Set constraint row labels
            # self.parent.peak_params_grid.SetCellValue(constraint_row, 0, "")  # No letter for constraint
            # self.parent.peak_params_grid.SetCellValue(constraint_row, 2, "Fixed")  # Position constraint
            # self.parent.peak_params_grid.SetCellValue(constraint_row, 3, "Fixed")  # Height constraint
            # self.parent.peak_params_grid.SetCellValue(constraint_row, 4, "Fixed")  # FWHM constraint
            # self.parent.peak_params_grid.SetCellValue(constraint_row, 5, "Fixed")  # L/G constraint
            # self.parent.peak_params_grid.SetCellValue(constraint_row, 6, "Fixed")  # Area constraint
            # self.parent.peak_params_grid.SetCellValue(constraint_row, 7, "Fixed")  # Sigma constraint
            # self.parent.peak_params_grid.SetCellValue(constraint_row, 8, "Fixed")  # Gamma constraint
            # self.parent.peak_params_grid.SetCellValue(constraint_row, 9, "Fixed")  # Skew constraint

            # Refresh grid to show colors
            self.parent.peak_params_grid.ForceRefresh()

            # Update Data dictionary (same structure as D-parameter)
            if 'Fitting' not in self.parent.Data['Core levels'][sheet_name]:
                self.parent.Data['Core levels'][sheet_name]['Fitting'] = {}
            if 'Peaks' not in self.parent.Data['Core levels'][sheet_name]['Fitting']:
                self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = {}

            peak_data = {
                'Position': float(round(center,3)),
                'Height': float(round(amplitude,3)),
                'FWHM': float(round(width_16_84,3)),  # Use 16%-84% as the width
                'L/G': 0.00,
                'Area': float(round(amplitude * width_16_84,3)),  # Approximate area
                'Sigma': float(round(sigma / 1000,2)),  # Store in eV
                'Gamma': float(kt_fitted / 1000),  # Store kT in eV
                'Skew': 0.00,
                'Fitting Model': 'Fermi',
                'Bkg Type': "None",
                'Bkg Low': float(x_min),
                'Bkg High': float(x_max),
                'Constraints': {
                    'Position': "Fixed",
                    'Height': "Fixed",
                    'FWHM': "Fixed",
                    'L/G': "Fixed",
                    'Area': "Fixed",
                    'Sigma': "Fixed",
                    'Gamma': "Fixed",
                    'Skew': "Fixed"
                },
                # Store fitted data for plotting (same as D-parameter stores derivative)
                'Fitted_X': x_fit.tolist(),
                'Fitted_Y': result.best_fit.tolist(),
                'Fermi_Center': float(center),
                'Fermi_16_84_Width': float(width_16_84),
                'Fermi_kT': float(kt_fitted / 1000),  # Store in eV
                'Fermi_Sigma': float(sigma / 1000)  # Store in eV
            }

            # Create the same descriptive name for the data key
            peak_name = (f"Fermi")
                         # f"\n\n\n"
                         # f" -Center Pos.: {center:.3f} eV\n"
                         # f" -Edge Width: {width_16_84:.3f} eV")

            self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'][peak_name] = peak_data

            # Update plot (same as D-parameter - just draw on canvas)
            self.parent.canvas.draw_idle()

            # Store vline positions BEFORE clear_and_replot (they get destroyed)
            vline1_x = self.parent.vline1.get_xdata()[0] if self.parent.vline1 else None
            vline2_x = self.parent.vline2.get_xdata()[0] if self.parent.vline2 else None

            # Trigger replot to show Fermi fit
            self.parent.clear_and_replot()

            # Restore vlines at their original positions (they were destroyed by clear_and_replot)
            if vline1_x is not None and vline2_x is not None:
                self.parent.vline1 = self.parent.ax.axvline(vline1_x, color='r', linestyle='--', alpha=0.7)
                self.parent.vline2 = self.parent.ax.axvline(vline2_x, color='r', linestyle='--', alpha=0.7)

                # Recreate text labels
                self.add_vline_text_labels()

                # Force vlines to be visible
                self.force_vlines_visible()

            # Force canvas redraw
            self.parent.canvas.draw_idle()

            # # Plot fit results on main plot
            # self.plot_thermal_fit(x_fit, result.best_fit, center, width_16_84)

        except Exception as e:
            wx.MessageBox(f"Fitting failed: {str(e)}\n\nEnsure lmfitxps is installed and vLines define a good range.",
                          "Fit Error")
            import traceback
            traceback.print_exc()

    def clear_previous_fits(self):
        """Clear previous thermal fit and VBM lines from plot"""
        # Remove thermal fit line
        if self.thermal_fit_line:
            try:
                self.thermal_fit_line.remove()
            except:
                pass
            self.thermal_fit_line = None

        # Remove VBM lines
        for line in self.vbm_lines:
            try:
                line.remove()
            except:
                pass
        self.vbm_lines.clear()

        # Remove any existing Fermi fit lines from legend
        handles, labels = self.parent.ax.get_legend_handles_labels()
        new_handles, new_labels = [], []
        for handle, label in zip(handles, labels):
            if 'Thermal' not in label and 'Center:' not in label and 'VBM' not in label:
                new_handles.append(handle)
                new_labels.append(label)

        if new_handles:
            self.parent.ax.legend(new_handles, new_labels)


    def plot_thermal_fit(self, x_fit, y_fit, center, width_16_84):
        """Plot directly on main plot (same as D-parameter does)"""
        # Plot Fermi fit directly (same pattern as D-parameter plots derivative)
        self.parent.ax.plot(x_fit, y_fit, '-', color='red', linewidth=1, label='Fermi')
        self.parent.canvas.draw_idle()

        # Ensure vLines stay visible
        self.force_vlines_visible()

    def OnCalculateVBM(self, event):
        """Calculate valence band minimum"""
        x_data, y_data = self.get_current_data()
        if x_data is None:
            wx.MessageBox("No active data available", "Error")
            return

        method = self.vbm_method.GetSelection()
        center_edge = self.vbm_edge_ctrl.GetValue()

        # Find VBM region (near zero binding energy)
        mask = (x_data >= -center_edge) & (x_data <= center_edge)
        x_vbm = x_data[mask]
        y_vbm = y_data[mask]

        if len(x_vbm) < 5:
            wx.MessageBox("Insufficient data in VBM range", "Error")
            return

        # Clear previous VBM lines
        for line in self.vbm_lines:
            try:
                line.remove()
            except:
                pass
        self.vbm_lines.clear()

        if method == 0:  # Linear Extrapolation
            center_edge = self.vbm_edge_ctrl.GetValue()
            n_points = self.vbm_points_ctrl.GetValue()
            use_bg = self.vbm_bg_checkbox.GetValue()

            # Initialize bg variables (will be updated if use_bg is True)
            bg_center = self.vbm_bg_center_ctrl.GetValue()
            bg_points = self.vbm_bg_points_ctrl.GetValue()

            # Initialize background fit variables
            bg_coef = None
            x_bg_fit = None
            y_bg_fit = None

            # Find signal edge points centered around center_edge
            # Find closest index to center_edge
            center_idx = np.argmin(np.abs(x_data - center_edge))

            # Calculate half-width for centering
            half_points = n_points // 2

            # Define signal fitting range (centered on center_edge)
            start_idx = max(0, center_idx - half_points)
            end_idx = min(len(x_data), center_idx + half_points)

            # Adjust if we hit boundaries
            actual_points = end_idx - start_idx
            if actual_points < n_points and start_idx > 0:
                start_idx = max(0, end_idx - n_points)
            elif actual_points < n_points and end_idx < len(x_data):
                end_idx = min(len(x_data), start_idx + n_points)

            # Extract signal fitting data
            x_signal_fit = x_data[start_idx:end_idx]
            y_signal_fit = y_data[start_idx:end_idx]

            if len(x_signal_fit) < 3:
                wx.MessageBox(f"Insufficient signal data points around {center_edge:.2f} eV", "Error")
                return

            # Fit linear to signal points
            signal_coef = np.polyfit(x_signal_fit, y_signal_fit, 1)

            # Handle background extrapolation or baseline intersection
            if use_bg:
                # Background extrapolation
                bg_center = self.vbm_bg_center_ctrl.GetValue()
                bg_points = self.vbm_bg_points_ctrl.GetValue()

                # Find background points centered around bg_center
                bg_center_idx = np.argmin(np.abs(x_data - bg_center))
                bg_half_points = bg_points // 2

                bg_start_idx = max(0, bg_center_idx - bg_half_points)
                bg_end_idx = min(len(x_data), bg_center_idx + bg_half_points)

                # Adjust if we hit boundaries
                bg_actual_points = bg_end_idx - bg_start_idx
                if bg_actual_points < bg_points and bg_start_idx > 0:
                    bg_start_idx = max(0, bg_end_idx - bg_points)
                elif bg_actual_points < bg_points and bg_end_idx < len(x_data):
                    bg_end_idx = min(len(x_data), bg_start_idx + bg_points)

                # Extract background fitting data
                x_bg_fit = x_data[bg_start_idx:bg_end_idx]
                y_bg_fit = y_data[bg_start_idx:bg_end_idx]

                if len(x_bg_fit) >= 3:
                    # Fit background line
                    bg_coef = np.polyfit(x_bg_fit, y_bg_fit, 1)

                    # Find intersection of signal and background lines
                    if abs(signal_coef[0] - bg_coef[0]) > 1e-10:  # Avoid division by zero
                        vbm_position = (bg_coef[1] - signal_coef[1]) / (signal_coef[0] - bg_coef[0])
                    else:
                        vbm_position = 0.0  # Lines are parallel

                    # Plot background extrapolation line
                    x_bg_extrap = np.linspace(min(x_data), max(x_data), 100)
                    y_bg_extrap = bg_coef[0] * x_bg_extrap + bg_coef[1]
                    bg_line = self.parent.ax.plot(x_bg_extrap, y_bg_extrap, 'b--', linewidth=1,
                                                  label=f'Background Fit')[0]
                    self.vbm_lines.append(bg_line)

                    # Mark background fit points
                    bg_fit_points_line = self.parent.ax.plot(x_bg_fit, y_bg_fit, 'bo', markersize=2)[0]
                                                             # label=f'BG Points ({len(x_bg_fit)})')[0]
                    self.vbm_lines.append(bg_fit_points_line)
                else:
                    wx.MessageBox("Insufficient background data points", "Warning")
                    return
            else:
                # Simple baseline intersection (y = 0)
                vbm_position = -signal_coef[1] / signal_coef[0]  # x where y = 0

            # Mark the signal fit points used (centered around center_edge)
            signal_fit_points_line = self.parent.ax.plot(x_signal_fit, y_signal_fit, 'bo', markersize=2)[0]
                                                         # label=f'Signal Points ({len(x_signal_fit)})')[0]
            self.vbm_lines.append(signal_fit_points_line)

            # Plot signal extrapolation line - FROM FULL RANGE
            x_signal_extrap = np.linspace(min(x_data), max(x_data), 100)
            y_signal_extrap = signal_coef[0] * x_signal_extrap + signal_coef[1]
            signal_line = self.parent.ax.plot(x_signal_extrap, y_signal_extrap, 'r--', linewidth=1, alpha=0.6,
                                              label='Linear Extrapolation')[0]
            self.vbm_lines.append(signal_line)


        # elif method == 1:  # Leading Edge Midpoint
        #     # Find midpoint of leading edge
        #     y_normalized = (y_vbm - np.min(y_vbm)) / (np.max(y_vbm) - np.min(y_vbm))
        #     idx_midpoint = np.argmin(np.abs(y_normalized - 0.5))
        #     vbm_position = x_vbm[idx_midpoint]
        #
        # elif method == 2:  # Derivative Method
        #     # Calculate derivative
        #     dy_dx = np.gradient(y_vbm, x_vbm)
        #     # Find maximum derivative (steepest point)
        #     idx_max = np.argmax(np.abs(dy_dx))
        #     vbm_position = x_vbm[idx_max]
        elif method == 1:  # Leading Edge Midpoint
            # Initialize variables for methods that don't use background
            use_bg = False
            center_edge = 0.0
            bg_center = 0.0
            n_points = 0
            bg_points = 0
            signal_coef = [0, 0]
            bg_coef = None
            x_signal_fit = np.array([])
            y_signal_fit = np.array([])
            x_bg_fit = np.array([])
            y_bg_fit = np.array([])

            # Find midpoint of leading edge
            y_normalized = (y_vbm - np.min(y_vbm)) / (np.max(y_vbm) - np.min(y_vbm))
            idx_midpoint = np.argmin(np.abs(y_normalized - 0.5))
            vbm_position = x_vbm[idx_midpoint]

        elif method == 2:  # Derivative Method
            # Initialize variables for methods that don't use background
            use_bg = False
            center_edge = 0.0
            bg_center = 0.0
            n_points = 0
            bg_points = 0
            signal_coef = [0, 0]
            bg_coef = None
            x_signal_fit = np.array([])
            y_signal_fit = np.array([])
            x_bg_fit = np.array([])
            y_bg_fit = np.array([])

            # Calculate derivative
            dy_dx = np.gradient(y_vbm, x_vbm)
            # Find maximum derivative (steepest point)
            idx_max = np.argmax(np.abs(dy_dx))
            vbm_position = x_vbm[idx_max]

        # Add VBM position line
        vbm_line = self.parent.ax.axvline(vbm_position, color='purple',
                                          linestyle='--', linewidth=1,
                                          label=f'VBM = {vbm_position:.3f} eV')
        self.vbm_lines.append(vbm_line)

        # Update results
        method_names = ["Linear Extrapolation", "Leading Edge Midpoint", "Derivative Method"]
        method_name = method_names[method]

        if method == 0:  # Linear Extrapolation
            bg_info = ""
            if use_bg:
                bg_info = f"\nBackground Center: {bg_center:.2f} eV\nBackground Points: {len(x_bg_fit)} (requested: {bg_points})"

            results_text = f"""VBM Analysis Results:

Method: {method_name}
VBM Position: {vbm_position:.2f} eV

Signal Fitting:
Center: {center_edge:.2f} eV
Points Used: {len(x_signal_fit)} (requested: {n_points})
Range: {x_signal_fit[0]:.2f} to {x_signal_fit[-1]:.2f} eV

Background Extrapolation: {'Yes' if use_bg else 'No'}{bg_info}

Additional Information:
Max Intensity: {np.max(y_data):.2f}
Min Intensity: {np.min(y_data):.2f}"""

        else:  # Leading Edge Midpoint or Derivative Method
            results_text = f"""VBM Analysis Results:

Method: {method_name}
VBM Position: {vbm_position:.2f} eV

Data Range: {x_vbm[0]:.2f} to {x_vbm[-1]:.2f} eV
Points Used: {len(x_vbm)}

Additional Information:
Max Intensity: {np.max(y_data):.2f}
Min Intensity: {np.min(y_data):.2f}"""

        self.results.SetValue(results_text)

        # Update legend and redraw
        self.parent.ax.legend()
        self.parent.canvas.draw_idle()

        # Store vline positions BEFORE add_peak_to_grid (they get destroyed by clear_and_replot)
        vline1_x = self.parent.vline1.get_xdata()[0] if self.parent.vline1 else None
        vline2_x = self.parent.vline2.get_xdata()[0] if self.parent.vline2 else None

        # Add VBM peak to grid (this calls clear_and_replot which destroys vlines)
        self.add_peak_to_grid(vbm_position, center_edge, bg_center, n_points, bg_points, use_bg,
                              signal_coef, bg_coef, x_signal_fit, y_signal_fit, x_bg_fit, y_bg_fit, "VBM")

        # Trigger replot to show VBM
        self.parent.clear_and_replot()

        # Restore vlines at their original positions (they were destroyed by clear_and_replot)
        if vline1_x is not None and vline2_x is not None:
            self.parent.vline1 = self.parent.ax.axvline(vline1_x, color='r', linestyle='--', alpha=0.7)
            self.parent.vline2 = self.parent.ax.axvline(vline2_x, color='r', linestyle='--', alpha=0.7)

            # Recreate text labels
            self.add_vline_text_labels()

            # Force vlines to be visible
            self.force_vlines_visible()

        # Force canvas redraw
        self.parent.canvas.draw_idle()

    def OnRemoveVBM(self, event):
        """Remove VBM or CutOff peak from grid and data"""
        sheet_name = self.parent.sheet_combobox.GetValue()

        # Remove from grid - find and delete VBM or Cut-Off rows
        removed_something = False
        for row in range(self.parent.peak_params_grid.GetNumberRows() - 1, -1, -1):
            fitting_model = self.parent.peak_params_grid.GetCellValue(row, 13)
            if fitting_model in ["VBM", "Cut-Off"]:
                # Delete both data row and constraint row (2 rows total)
                self.parent.peak_params_grid.DeleteRows(row, 2)
                removed_something = True
                break

        if not removed_something:
            wx.MessageBox("No VBM or Cut-Off data found to remove", "Info", wx.OK | wx.ICON_INFORMATION)
            return

        # Remove from Data structure
        if sheet_name in self.parent.Data['Core levels']:
            fitting_data = self.parent.Data['Core levels'][sheet_name].get('Fitting', {})
            peaks_data = fitting_data.get('Peaks', {})

            # Find and remove VBM or Cut-Off peak by fitting model
            keys_to_remove = []
            for peak_name, peak_info in peaks_data.items():
                if peak_info.get('Fitting Model') in ['VBM', 'Cut-Off']:
                    keys_to_remove.append(peak_name)

            for key in keys_to_remove:
                del peaks_data[key]

        # Clear results display
        self.results.SetValue("")

        # Clear VBM lines from plot
        for line in self.vbm_lines:
            try:
                line.remove()
            except:
                pass
        self.vbm_lines.clear()

        # Store vline positions BEFORE clear_and_replot (they get destroyed)
        vline1_x = self.parent.vline1.get_xdata()[0] if self.parent.vline1 else None
        vline2_x = self.parent.vline2.get_xdata()[0] if self.parent.vline2 else None

        # Trigger replot to remove VBM lines
        self.parent.clear_and_replot()

        # Restore vlines at their original positions (they were destroyed by clear_and_replot)
        if vline1_x is not None and vline2_x is not None:
            self.parent.vline1 = self.parent.ax.axvline(vline1_x, color='r', linestyle='--', alpha=0.7)
            self.parent.vline2 = self.parent.ax.axvline(vline2_x, color='r', linestyle='--', alpha=0.7)

            # Recreate text labels
            self.add_vline_text_labels()

            # Force vlines to be visible
            self.force_vlines_visible()

        # Force canvas redraw
        self.parent.canvas.draw_idle()

        wx.MessageBox("VBM or Cut-Off data removed", "Removed", wx.OK | wx.ICON_INFORMATION)

    def add_vline_text_labels(self):
        """Add text labels to vlines showing their positions"""
        if self.parent.vline1 is not None and self.parent.vline2 is not None:
            # Get current vline positions
            vline1_x = self.parent.vline1.get_xdata()[0]
            vline2_x = self.parent.vline2.get_xdata()[0]

            # Get y-axis limits for positioning text
            ylim = self.parent.ax.get_ylim()
            text_y = ylim[1] * 0.95  # Position at 95% of y-axis height

            # Remove existing text labels if they exist
            if hasattr(self.parent, 'vline1_text') and self.parent.vline1_text is not None:
                try:
                    self.parent.vline1_text.remove()
                except:
                    pass

            if hasattr(self.parent, 'vline2_text') and self.parent.vline2_text is not None:
                try:
                    self.parent.vline2_text.remove()
                except:
                    pass

            # Create new text labels
            self.parent.vline1_text = self.parent.ax.text(vline1_x, text_y, f'{vline1_x:.2f}',
                                                          ha='center', va='top', fontsize=8,
                                                          bbox=dict(boxstyle='round,pad=0.2',
                                                                    facecolor='white', alpha=0.8))

            self.parent.vline2_text = self.parent.ax.text(vline2_x, text_y, f'{vline2_x:.2f}',
                                                          ha='center', va='top', fontsize=8,
                                                          bbox=dict(boxstyle='round,pad=0.2',
                                                                    facecolor='white', alpha=0.8))

    def on_close(self, event):
        """Cleanup when closing VB measurements"""
        # Clear all fits from plot
        self.clear_previous_fits()

        # Reset background interaction
        self.parent.background_tab_selected = False

        # Hide vLines (they will be restored when needed by other tools)
        if self.parent.vline1 is not None:
            self.parent.vline1.set_visible(False)
        if self.parent.vline2 is not None:
            self.parent.vline2.set_visible(False)
        if hasattr(self.parent, 'vline1_text') and self.parent.vline1_text is not None:
            self.parent.vline1_text.set_visible(False)
        if hasattr(self.parent, 'vline2_text') and self.parent.vline2_text is not None:
            self.parent.vline2_text.set_visible(False)

        # Clear window reference
        if hasattr(self.parent, 'vb_measurements_window'):
            self.parent.vb_measurements_window = None

        # Replot data and update legend when closing window
        self.parent.clear_and_replot()

        # Ensure canvas is redrawn
        self.parent.canvas.draw_idle()

        self.Destroy()

    def OnRemoveFermi(self, event):
        """Remove Fermi peak from grid and data"""
        sheet_name = self.parent.sheet_combobox.GetValue()

        # Remove from grid - find and delete Fermi rows
        for row in range(self.parent.peak_params_grid.GetNumberRows() - 1, -1, -1):
            if self.parent.peak_params_grid.GetCellValue(row, 13) == "Fermi":
                # Delete both data row and constraint row (2 rows total)
                self.parent.peak_params_grid.DeleteRows(row, 2)
                break

        # Remove from Data structure
        if sheet_name in self.parent.Data['Core levels']:
            fitting_data = self.parent.Data['Core levels'][sheet_name].get('Fitting', {})
            peaks_data = fitting_data.get('Peaks', {})

            # Find and remove Fermi peak by fitting model
            fermi_key = None
            for peak_name, peak_info in peaks_data.items():
                if peak_info.get('Fitting Model') == 'Fermi':
                    fermi_key = peak_name
                    break

            if fermi_key:
                del peaks_data[fermi_key]

        # Clear results display
        self.results.SetValue("")

        # Store vline positions BEFORE clear_and_replot (they get destroyed)
        vline1_x = self.parent.vline1.get_xdata()[0] if self.parent.vline1 else None
        vline2_x = self.parent.vline2.get_xdata()[0] if self.parent.vline2 else None

        # Trigger replot to remove Fermi lines
        self.parent.clear_and_replot()

        # Restore vlines at their original positions (they were destroyed by clear_and_replot)
        if vline1_x is not None and vline2_x is not None:
            self.parent.vline1 = self.parent.ax.axvline(vline1_x, color='r', linestyle='--', alpha=0.7)
            self.parent.vline2 = self.parent.ax.axvline(vline2_x, color='r', linestyle='--', alpha=0.7)

            # Recreate text labels
            self.add_vline_text_labels()

            # Force vlines to be visible
            self.force_vlines_visible()

        # Force canvas redraw
        self.parent.canvas.draw_idle()

        wx.MessageBox("Fermi edge data removed", "Removed", wx.OK | wx.ICON_INFORMATION)

    def OnBgCheckbox(self, event):
        """Enable/disable background controls"""
        enabled = self.vbm_bg_checkbox.GetValue()
        self.vbm_bg_center_ctrl.Enable(enabled)
        self.vbm_bg_points_ctrl.Enable(enabled)

    def OnEdgeCenterChange(self, event):
        """Update vline1 position when edge center control changes"""
        if self.parent.vline1 is not None:
            new_value = self.vbm_edge_ctrl.GetValue()
            self.parent.vline1.set_xdata([new_value, new_value])

            if hasattr(self.parent, 'update_vline_text_labels'):
                self.parent.update_vline_text_labels()

            # Update averaging indicator lines
            if hasattr(self.parent, 'add_averaging_indicator_lines'):
                self.parent.add_averaging_indicator_lines()

            self.parent.canvas.draw_idle()

    def OnBgCenterChange(self, event):
        """Update vline2 position when background center control changes"""
        if self.parent.vline2 is not None:
            new_value = self.vbm_bg_center_ctrl.GetValue()
            self.parent.vline2.set_xdata([new_value, new_value])

            if hasattr(self.parent, 'update_vline_text_labels'):
                self.parent.update_vline_text_labels()

            # Update averaging indicator lines
            if hasattr(self.parent, 'add_averaging_indicator_lines'):
                self.parent.add_averaging_indicator_lines()

            self.parent.canvas.draw_idle()

    def update_controls_from_vlines(self):
        """Update VBM controls when vLines are dragged"""
        if self.parent.vline1 is not None and self.parent.vline2 is not None:
            vline1_pos = self.parent.vline1.get_xdata()[0]
            vline2_pos = self.parent.vline2.get_xdata()[0]

            # vline1 controls Center Edge
            self.vbm_edge_ctrl.SetValue(round(vline1_pos, 2))

            # vline2 controls BG Center
            self.vbm_bg_center_ctrl.SetValue(round(vline2_pos, 2))

    def setup_vbm_vlines(self):
        """Position vLines at current control values"""
        if self.parent.vline1 is not None and self.parent.vline2 is not None:
            edge_center = self.vbm_edge_ctrl.GetValue()
            bg_center = self.vbm_bg_center_ctrl.GetValue()

            self.parent.vline1.set_xdata([edge_center, edge_center])
            self.parent.vline2.set_xdata([bg_center, bg_center])

            if hasattr(self.parent, 'update_vline_text_labels'):
                self.parent.update_vline_text_labels()

            # Update averaging indicator lines
            if hasattr(self.parent, 'add_averaging_indicator_lines'):
                self.parent.add_averaging_indicator_lines()

            self.parent.canvas.draw_idle()

    def add_peak_to_grid(self, peak_position, center_edge, bg_center, n_points, bg_points, use_bg, signal_coef=None,
                         bg_coef=None, x_signal_fit=None, y_signal_fit=None, x_bg_fit=None, y_bg_fit=None,
                         peak_type="VBM"):
        """Add VBM or Cut-Off peak to the peak fitting grid"""
        sheet_name = self.parent.sheet_combobox.GetValue()

        # Initialize data structure if needed
        if 'Fitting' not in self.parent.Data['Core levels'][sheet_name]:
            self.parent.Data['Core levels'][sheet_name]['Fitting'] = {}
        if 'Peaks' not in self.parent.Data['Core levels'][sheet_name]['Fitting']:
            self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = {}

        # Remove existing peak of same type if it exists
        peaks_data = self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']
        existing_key = None
        for peak_name, peak_info in peaks_data.items():
            if peak_info.get('Fitting Model') == peak_type:
                existing_key = peak_name
                break

        if existing_key:
            del peaks_data[existing_key]
            # Remove from grid
            for row in range(self.parent.peak_params_grid.GetNumberRows() - 1, -1, -1):
                if self.parent.peak_params_grid.GetCellValue(row, 13) == peak_type:
                    self.parent.peak_params_grid.DeleteRows(row, 2)
                    break

        # Add new rows to grid
        self.parent.peak_params_grid.AppendRows(2)
        row = self.parent.peak_params_grid.GetNumberRows() - 2

        # Get next letter ID
        letter_id = chr(65 + len(peaks_data))

        # Set grid values (all in .2f format)
        self.parent.peak_params_grid.SetCellValue(row, 0, letter_id)  # Letter ID
        self.parent.peak_params_grid.SetReadOnly(row, 0)
        self.parent.peak_params_grid.SetCellValue(row, 1, peak_type)  # Label
        self.parent.peak_params_grid.SetCellValue(row, 2, f"{peak_position:.2f}")  # Position
        self.parent.peak_params_grid.SetCellValue(row, 3, f"{center_edge:.2f}")  # Height (Edge Center)
        self.parent.peak_params_grid.SetCellValue(row, 4, f"{n_points:.2f}")  # FWHM (Edge Points)
        self.parent.peak_params_grid.SetCellValue(row, 5, f"{bg_center:.2f}")  # L/G (BG Center)
        self.parent.peak_params_grid.SetCellValue(row, 6, f"{bg_points:.2f}")  # Area (BG Points)
        self.parent.peak_params_grid.SetCellValue(row, 7, f"{1 if use_bg else 0:.2f}")  # Sigma (Use BG flag)
        self.parent.peak_params_grid.SetCellValue(row, 13, peak_type)  # Fitting Model

        # Set constraint row background color
        for col in range(self.parent.peak_params_grid.GetNumberCols()):
            self.parent.peak_params_grid.SetCellBackgroundColour(row + 1, col, wx.Colour(200, 245, 228))
            self.parent.peak_params_grid.SetCellBackgroundColour(row, col, wx.WHITE)

        # Store in data structure with extrapolation data
        peak_data = {
            'Position': peak_position,
            'Height': center_edge,  # Store edge center as height
            'FWHM': n_points,  # Store edge points as FWHM
            'L/G': bg_center,  # Store BG center as L/G
            'Area': bg_points,  # Store BG points as Area
            'Sigma': 1 if use_bg else 0,  # Store use_bg flag as Sigma
            'Gamma': 0.0,
            'Skew': 0.0,
            'Fitting Model': peak_type,
            f'{peak_type}_Edge_Center': center_edge,
            f'{peak_type}_BG_Center': bg_center,
            f'{peak_type}_Edge_Points': n_points,
            f'{peak_type}_BG_Points': bg_points,
            f'{peak_type}_Use_BG': use_bg,
            'Signal_Coef': signal_coef.tolist() if signal_coef is not None else None,
            'BG_Coef': bg_coef.tolist() if bg_coef is not None else None,
            'X_Signal_Fit': x_signal_fit.tolist() if x_signal_fit is not None else None,
            'Y_Signal_Fit': y_signal_fit.tolist() if y_signal_fit is not None else None,
            'X_BG_Fit': x_bg_fit.tolist() if x_bg_fit is not None else None,
            'Y_BG_Fit': y_bg_fit.tolist() if y_bg_fit is not None else None,
            'Constraints': {}
        }

        peaks_data[peak_type] = peak_data

        return True

    def restore_vlines_after_plot(self, vline1_pos, vline2_pos):
        """Restore vlines at specified positions after plotting"""
        try:
            # FIRST: Remove/destroy any existing vlines
            if self.parent.vline1 is not None:
                try:
                    self.parent.vline1.remove()
                except:
                    pass
                self.parent.vline1 = None

            if self.parent.vline2 is not None:
                try:
                    self.parent.vline2.remove()
                except:
                    pass
                self.parent.vline2 = None

            # Remove any existing text labels
            if hasattr(self.parent, 'vline1_text') and self.parent.vline1_text is not None:
                try:
                    self.parent.vline1_text.remove()
                except:
                    pass
                self.parent.vline1_text = None

            if hasattr(self.parent, 'vline2_text') and self.parent.vline2_text is not None:
                try:
                    self.parent.vline2_text.remove()
                except:
                    pass
                self.parent.vline2_text = None

            # THEN: Create new vlines at the specified positions
            self.parent.vline1 = self.parent.ax.axvline(x=vline1_pos, color='red', linestyle='--', alpha=0.7)
            self.parent.vline2 = self.parent.ax.axvline(x=vline2_pos, color='red', linestyle='--', alpha=0.7)

            # Create new text labels
            if hasattr(self.parent, 'add_vline_text_labels'):
                self.parent.add_vline_text_labels()

            elif hasattr(self.parent, 'update_vline_text_labels'):
                self.parent.update_vline_text_labels()

            # Update averaging indicator lines
            if hasattr(self.parent, 'add_averaging_indicator_lines'):
                self.parent.add_averaging_indicator_lines()

            # Make sure they're visible and background tab is selected
            self.parent.show_hide_vlines()
            self.parent.background_tab_selected = True

            # Force canvas redraw
            self.parent.canvas.draw_idle()

        except Exception as e:
            print(f"Error restoring vlines: {e}")

    def OnCalculateCutOff(self, event):
        """Calculate cut-off"""
        x_data, y_data = self.get_current_data()
        if x_data is None:
            wx.MessageBox("No active data available", "Error")
            return

        method = self.vbm_method.GetSelection()
        center_edge = self.vbm_edge_ctrl.GetValue()

        # Find Cut-Off region (near zero binding energy)
        mask = (x_data >= -center_edge) & (x_data <= center_edge)
        x_cutoff = x_data[mask]
        y_cutoff = y_data[mask]

        if len(x_cutoff) < 5:
            wx.MessageBox("Insufficient data in Cut-Off range", "Error")
            return

        # Clear previous Cut-Off lines
        for line in self.vbm_lines:
            try:
                line.remove()
            except:
                pass
        self.vbm_lines.clear()

        if method == 0:  # Linear Extrapolation
            center_edge = self.vbm_edge_ctrl.GetValue()
            n_points = self.vbm_points_ctrl.GetValue()
            use_bg = self.vbm_bg_checkbox.GetValue()

            # Initialize bg variables (will be updated if use_bg is True)
            bg_center = self.vbm_bg_center_ctrl.GetValue()
            bg_points = self.vbm_bg_points_ctrl.GetValue()

            # Initialize background fit variables
            bg_coef = None
            x_bg_fit = None
            y_bg_fit = None

            # Find signal edge points centered around center_edge
            # Find closest index to center_edge
            center_idx = np.argmin(np.abs(x_data - center_edge))

            # Calculate half-width for centering
            half_points = n_points // 2

            # Define signal fitting range (centered on center_edge)
            start_idx = max(0, center_idx - half_points)
            end_idx = min(len(x_data), center_idx + half_points)

            # Adjust if we hit boundaries
            actual_points = end_idx - start_idx
            if actual_points < n_points and start_idx > 0:
                start_idx = max(0, end_idx - n_points)
            elif actual_points < n_points and end_idx < len(x_data):
                end_idx = min(len(x_data), start_idx + n_points)

            # Extract signal fitting data
            x_signal_fit = x_data[start_idx:end_idx]
            y_signal_fit = y_data[start_idx:end_idx]

            if len(x_signal_fit) < 3:
                wx.MessageBox(f"Insufficient signal data points around {center_edge:.2f} eV", "Error")
                return

            # Fit linear to signal points
            signal_coef = np.polyfit(x_signal_fit, y_signal_fit, 1)

            # Handle background extrapolation or baseline intersection
            if use_bg:
                # Background extrapolation
                bg_center = self.vbm_bg_center_ctrl.GetValue()
                bg_points = self.vbm_bg_points_ctrl.GetValue()

                # Find background points centered around bg_center
                bg_center_idx = np.argmin(np.abs(x_data - bg_center))
                bg_half_points = bg_points // 2

                bg_start_idx = max(0, bg_center_idx - bg_half_points)
                bg_end_idx = min(len(x_data), bg_center_idx + bg_half_points)

                # Adjust if we hit boundaries
                bg_actual_points = bg_end_idx - bg_start_idx
                if bg_actual_points < bg_points and bg_start_idx > 0:
                    bg_start_idx = max(0, bg_end_idx - bg_points)
                elif bg_actual_points < bg_points and bg_end_idx < len(x_data):
                    bg_end_idx = min(len(x_data), bg_start_idx + bg_points)

                # Extract background fitting data
                x_bg_fit = x_data[bg_start_idx:bg_end_idx]
                y_bg_fit = y_data[bg_start_idx:bg_end_idx]

                if len(x_bg_fit) >= 3:
                    # Fit background line
                    bg_coef = np.polyfit(x_bg_fit, y_bg_fit, 1)

                    # Find intersection of signal and background lines
                    if abs(signal_coef[0] - bg_coef[0]) > 1e-10:  # Avoid division by zero
                        cutoff_position = (bg_coef[1] - signal_coef[1]) / (signal_coef[0] - bg_coef[0])
                    else:
                        cutoff_position = 0.0  # Lines are parallel

                    # Plot background extrapolation line
                    x_bg_extrap = np.linspace(min(x_data), max(x_data), 100)
                    y_bg_extrap = bg_coef[0] * x_bg_extrap + bg_coef[1]
                    bg_line = self.parent.ax.plot(x_bg_extrap, y_bg_extrap, 'b--', linewidth=1,
                                                  label=f'Background Fit')[0]
                    self.vbm_lines.append(bg_line)

                    # Mark background fit points
                    bg_fit_points_line = self.parent.ax.plot(x_bg_fit, y_bg_fit, 'bo', markersize=2)[0]
                    # label=f'BG Points ({len(x_bg_fit)})')[0]
                    self.vbm_lines.append(bg_fit_points_line)
                else:
                    wx.MessageBox("Insufficient background data points", "Warning")
                    return
            else:
                # Simple baseline intersection (y = 0)
                cutoff_position = -signal_coef[1] / signal_coef[0]  # x where y = 0

            # Mark the signal fit points used (centered around center_edge)
            signal_fit_points_line = self.parent.ax.plot(x_signal_fit, y_signal_fit, 'bo', markersize=2)[0]
            # label=f'Signal Points ({len(x_signal_fit)})')[0]
            self.vbm_lines.append(signal_fit_points_line)

            # Plot signal extrapolation line - FROM FULL RANGE
            x_signal_extrap = np.linspace(min(x_data), max(x_data), 100)
            y_signal_extrap = signal_coef[0] * x_signal_extrap + signal_coef[1]
            signal_line = self.parent.ax.plot(x_signal_extrap, y_signal_extrap, 'r--', linewidth=1, alpha=0.6,
                                              label='Linear Extrapolation')[0]
            self.vbm_lines.append(signal_line)

        # elif method == 1:  # Leading Edge Midpoint
        #     # Find midpoint of leading edge
        #     y_normalized = (y_cutoff - np.min(y_cutoff)) / (np.max(y_cutoff) - np.min(y_cutoff))
        #     idx_midpoint = np.argmin(np.abs(y_normalized - 0.5))
        #     cutoff_position = x_cutoff[idx_midpoint]
        #
        # elif method == 2:  # Derivative Method
        #     # Calculate derivative
        #     dy_dx = np.gradient(y_cutoff, x_cutoff)
        #     # Find maximum derivative (steepest point)
        #     idx_max = np.argmax(np.abs(dy_dx))
        #     cutoff_position = x_cutoff[idx_max]
        elif method == 1:  # Leading Edge Midpoint
            # Initialize variables for methods that don't use background
            use_bg = False
            center_edge = 0.0
            bg_center = 0.0
            n_points = 0
            bg_points = 0
            signal_coef = [0, 0]
            bg_coef = None
            x_signal_fit = np.array([])
            y_signal_fit = np.array([])
            x_bg_fit = np.array([])
            y_bg_fit = np.array([])

            # Find midpoint of leading edge
            y_normalized = (y_vbm - np.min(y_vbm)) / (np.max(y_vbm) - np.min(y_vbm))
            idx_midpoint = np.argmin(np.abs(y_normalized - 0.5))
            vbm_position = x_vbm[idx_midpoint]

        elif method == 2:  # Derivative Method
            # Initialize variables for methods that don't use background
            use_bg = False
            center_edge = 0.0
            bg_center = 0.0
            n_points = 0
            bg_points = 0
            signal_coef = [0, 0]
            bg_coef = None
            x_signal_fit = np.array([])
            y_signal_fit = np.array([])
            x_bg_fit = np.array([])
            y_bg_fit = np.array([])

            # Calculate derivative
            dy_dx = np.gradient(y_vbm, x_vbm)
            # Find maximum derivative (steepest point)
            idx_max = np.argmax(np.abs(dy_dx))
            vbm_position = x_vbm[idx_max]

        # Add Cut-Off position line
        cutoff_line = self.parent.ax.axvline(cutoff_position, color='orange',
                                             linestyle='--', linewidth=1,
                                             label=f'Cut-Off = {cutoff_position:.3f} eV')
        self.vbm_lines.append(cutoff_line)

        # Update results
        bg_info = ""
        if use_bg:
            bg_info = f"\nBackground Center: {bg_center:.2f} eV\nBackground Points: {len(x_bg_fit)} (requested: {bg_points})"

        results_text = f"""Cut-Off Analysis Results:

    Method: Linear Extrapolation
    Cut-Off Position: {cutoff_position:.2f} eV

    Signal Fitting:
    Center: {center_edge:.2f} eV
    Points Used: {len(x_signal_fit)} (requested: {n_points})
    Range: {x_signal_fit[0]:.2f} to {x_signal_fit[-1]:.2f} eV

    Background Extrapolation: {'Yes' if use_bg else 'No'}{bg_info}

    Additional Information:
    Max Intensity: {np.max(y_data):.2f}
    Min Intensity: {np.min(y_data):.2f}"""

        self.results.SetValue(results_text)

        # Update legend and redraw
        self.parent.ax.legend()
        self.parent.canvas.draw_idle()

        # Store vline positions BEFORE add_peak_to_grid (they get destroyed by clear_and_replot)
        vline1_x = self.parent.vline1.get_xdata()[0] if self.parent.vline1 else None
        vline2_x = self.parent.vline2.get_xdata()[0] if self.parent.vline2 else None

        # Add Cut-Off peak to grid
        self.add_peak_to_grid(cutoff_position, center_edge, bg_center, n_points, bg_points, use_bg,
                              signal_coef, bg_coef, x_signal_fit, y_signal_fit, x_bg_fit, y_bg_fit, "Cut-Off")

        # Trigger replot to show Cut-Off
        self.parent.clear_and_replot()

        # Restore vlines at their original positions (they were destroyed by clear_and_replot)
        if vline1_x is not None and vline2_x is not None:
            self.parent.vline1 = self.parent.ax.axvline(vline1_x, color='r', linestyle='--', alpha=0.7)
            self.parent.vline2 = self.parent.ax.axvline(vline2_x, color='r', linestyle='--', alpha=0.7)

            # Recreate text labels
            self.add_vline_text_labels()

            # Force vlines to be visible
            self.force_vlines_visible()

        # Force canvas redraw
        self.parent.canvas.draw_idle()
