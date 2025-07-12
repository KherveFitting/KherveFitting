import wx
from matplotlib.figure import Figure
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
import matplotlib.pyplot as plt
import numpy as np
import lmfit
from scipy import sparse
from scipy.sparse.linalg import spsolve
from scipy.ndimage import gaussian_filter


class TougaardRamanFitWindow(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, title="ALS-Raman Background Model", size=(1050, 600))

        self.parent = parent
        self.panel = wx.Panel(self)

        # Get data from parent window
        sheet_name = self.parent.parent.sheet_combobox.GetValue()
        self.x_values = np.array(self.parent.parent.Data['Core levels'][sheet_name]['B.E.'])
        self.y_values = np.array(self.parent.parent.Data['Core levels'][sheet_name]['Raw Data'])

        # Initialize variables
        self.vlines = []
        self.zones = []
        self.dragging = False
        self.current_vline = None
        self.y_max = max(self.y_values)
        self.y_min = 0.99 * min(self.y_values)
        self.background = None

        # Create main layout
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Create left control panel
        control_panel = wx.Panel(self.panel)
        control_sizer = wx.BoxSizer(wx.VERTICAL)

        # ALS Raman Controls Section
        als_box = wx.StaticBox(control_panel, label="ALS Raman Settings")
        als_sizer = wx.StaticBoxSizer(als_box, wx.VERTICAL)

        # Lambda parameter
        lambda_sizer = wx.BoxSizer(wx.HORIZONTAL)
        lambda_sizer.Add(wx.StaticText(als_box, label="Lambda:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.lambda_value = wx.SpinCtrlDouble(als_box, min=1, max=1e7, inc=10, value="5")
        lambda_sizer.Add(self.lambda_value, 1, wx.ALL, 5)
        self.lambda_fixed = wx.CheckBox(als_box, label="Fix")
        lambda_sizer.Add(self.lambda_fixed, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        als_sizer.Add(lambda_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # Lambda min/max
        lambda_range_sizer = wx.BoxSizer(wx.HORIZONTAL)
        lambda_range_sizer.Add(wx.StaticText(als_box, label="Min:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.lambda_min = wx.SpinCtrlDouble(als_box, min=1, max=1e7, inc=10, value="1")
        lambda_range_sizer.Add(self.lambda_min, 1, wx.RIGHT, 5)
        lambda_range_sizer.Add(wx.StaticText(als_box, label="Max:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT | wx.LEFT,
                               5)
        self.lambda_max = wx.SpinCtrlDouble(als_box, min=1, max=1e7, inc=10, value="1000")
        lambda_range_sizer.Add(self.lambda_max, 1)
        als_sizer.Add(lambda_range_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # P parameter
        p_sizer = wx.BoxSizer(wx.HORIZONTAL)
        p_sizer.Add(wx.StaticText(als_box, label="p value:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.p_value = wx.SpinCtrlDouble(als_box, min=0.0001, max=0.6, inc=0.001, value="0.001")
        p_sizer.Add(self.p_value, 1, wx.ALL, 5)
        self.p_fixed = wx.CheckBox(als_box, label="Fix")
        p_sizer.Add(self.p_fixed, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        als_sizer.Add(p_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # P min/max
        p_range_sizer = wx.BoxSizer(wx.HORIZONTAL)
        p_range_sizer.Add(wx.StaticText(als_box, label="Min:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.p_min = wx.SpinCtrlDouble(als_box, min=0.0001, max=0.5, inc=0.001, value="0.0001")
        p_range_sizer.Add(self.p_min, 1, wx.RIGHT, 5)
        p_range_sizer.Add(wx.StaticText(als_box, label="Max:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT | wx.LEFT, 5)
        self.p_max = wx.SpinCtrlDouble(als_box, min=0.0001, max=0.6, inc=0.001, value="2")
        p_range_sizer.Add(self.p_max, 1)
        als_sizer.Add(p_range_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # Iterations
        iter_sizer = wx.BoxSizer(wx.HORIZONTAL)
        iter_sizer.Add(wx.StaticText(als_box, label="Iterations:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.iter_value = wx.SpinCtrl(als_box, min=5, max=100, initial=30)
        iter_sizer.Add(self.iter_value, 1, wx.ALL, 5)
        als_sizer.Add(iter_sizer, 0, wx.EXPAND | wx.ALL, 5)

        control_sizer.Add(als_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # 4. Smoothing control
        smooth_box = wx.StaticBox(control_panel, label="Data Smoothing")
        smooth_sizer = wx.StaticBoxSizer(smooth_box, wx.VERTICAL)

        smooth_ctrl_sizer = wx.BoxSizer(wx.HORIZONTAL)
        smooth_ctrl_sizer.Add(wx.StaticText(smooth_box, label="Gaussian Width:"), 0,
                              wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.smooth_value = wx.SpinCtrlDouble(smooth_box, min=0, max=10, inc=0.1, value="0")
        self.smooth_value.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_smooth_change)
        smooth_ctrl_sizer.Add(self.smooth_value, 1, wx.ALL, 5)
        smooth_sizer.Add(smooth_ctrl_sizer, 0, wx.EXPAND | wx.ALL, 5)

        control_sizer.Add(smooth_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # 5. Fitting Zones Section
        zones_box = wx.StaticBox(control_panel, label="Exclusion Zones")
        zones_sizer = wx.StaticBoxSizer(zones_box, wx.VERTICAL)

        zones_ctrl_sizer = wx.BoxSizer(wx.HORIZONTAL)
        zones_ctrl_sizer.Add(wx.StaticText(zones_box, label="Number of Zones:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT,
                             5)
        self.num_zones = wx.SpinCtrl(zones_box, min=1, max=10, initial=1)
        self.num_zones.Bind(wx.EVT_SPINCTRL, self.on_num_zones_change)
        zones_ctrl_sizer.Add(self.num_zones, 1, wx.ALL, 5)
        zones_sizer.Add(zones_ctrl_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # Zones scroll area
        self.zones_panel = wx.ScrolledWindow(zones_box)
        self.zones_panel.SetScrollRate(0, 20)
        self.zones_panel_sizer = wx.BoxSizer(wx.VERTICAL)
        self.zones_panel.SetSizer(self.zones_panel_sizer)
        zones_sizer.Add(self.zones_panel, 1, wx.EXPAND | wx.ALL, 5)

        control_sizer.Add(zones_sizer, 1, wx.EXPAND | wx.ALL, 5)

        # 6. Buttons Section
        buttons_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.fit_button = wx.Button(control_panel, label="Fit")
        self.fit_button.SetMinSize((110, 40))
        self.fit_button.Bind(wx.EVT_BUTTON, self.on_fit)
        buttons_sizer.Add(self.fit_button, 1, wx.ALL, 5)

        self.export_button = wx.Button(control_panel, label="Export Background")
        self.export_button.SetMinSize((110, 40))
        self.export_button.Bind(wx.EVT_BUTTON, self.on_export)
        buttons_sizer.Add(self.export_button, 1, wx.ALL, 5)

        control_sizer.Add(buttons_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # Finalize control panel
        control_panel.SetSizer(control_sizer)

        # Create plot area
        self.figure = Figure(figsize=(6, 4))
        self.canvas = FigureCanvas(self.panel, -1, self.figure)
        self.ax = self.figure.add_subplot(111)

        # Add both panels to main sizer
        main_sizer.Add(control_panel, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(self.canvas, 1, wx.EXPAND | wx.ALL, 5)

        self.panel.SetSizer(main_sizer)

        # Set initial state
        self.create_zones(1)

        # Center on screen
        self.Center()

        # Bind events for plot interaction
        self.canvas.mpl_connect('button_press_event', self.on_press)
        self.canvas.mpl_connect('button_release_event', self.on_release)
        self.canvas.mpl_connect('motion_notify_event', self.on_motion)
        self.panel.Bind(wx.EVT_CHAR_HOOK, self.on_key_press)

        # Initial plot
        self.plot_initial_data()

        # Apply consistent fonts
        from libraries.ConfigFile import set_consistent_fonts
        set_consistent_fonts(self)

    def on_key_press(self, event):
        if event.ControlDown():
            if event.GetKeyCode() in [wx.WXK_UP, wx.WXK_DOWN]:
                # Adjust y_max to zoom in/out
                intensity_factor = 0.05

                if event.GetKeyCode() == wx.WXK_DOWN:
                    self.y_max = max(self.y_max - intensity_factor * max(self.y_values), self.y_min)
                else:
                    self.y_max += intensity_factor * max(self.y_values)

                self.ax.set_ylim(self.y_min, self.y_max)
                self.canvas.draw_idle()
                return
        event.Skip()

    def create_zones(self, num_zones):
        # Store current values if they exist
        old_values = []
        for i, zone in enumerate(self.zones):
            if i < len(self.zones):
                try:
                    old_values.append({
                        'min': zone['min'].GetValue(),
                        'max': zone['max'].GetValue(),
                        'use_zone': zone['use_zone'].GetValue()
                    })
                except:
                    # If controls are already deleted, use default values
                    old_values.append({
                        'min': max(self.x_values) - 15 - (i * 30) if self.x_values is not None else 0,
                        'max': max(self.x_values) - 1 - (i * 30) if self.x_values is not None else 100,
                        'use_zone': (i == 0)
                    })

        # Clear existing zones - first remove all items from the list
        self.zones = []
        # Then clear the sizer, which will delete all the windows
        self.zones_panel_sizer.Clear(True)


        # Create new zones
        for i in range(num_zones):
            if i < len(old_values):
                min_val = old_values[i]['min']
                max_val = old_values[i]['max']
                use_zone = old_values[i]['use_zone']
            else:
                if self.x_values is not None and len(self.x_values) > 0:
                    min_val = max(self.x_values) - 15 - (i * 30)
                    max_val = max(self.x_values) - 1 - (i * 30)
                else:
                    min_val = 0
                    max_val = 100
                use_zone = (i == 0)

            zone_box = wx.StaticBox(self.zones_panel, label=f"Zone {i + 1}")
            zone_sizer = wx.StaticBoxSizer(zone_box, wx.VERTICAL)

            min_max_sizer = wx.BoxSizer(wx.HORIZONTAL)
            min_max_sizer.Add(wx.StaticText(self.zones_panel, label="Min:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
            min_ctrl = wx.SpinCtrlDouble(self.zones_panel, min=0, max=20000, inc=0.1, value=str(min_val))
            min_ctrl.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_zone_change)
            min_max_sizer.Add(min_ctrl, 1, wx.RIGHT, 5)

            min_max_sizer.Add(wx.StaticText(self.zones_panel, label="Max:"), 0,
                              wx.ALIGN_CENTER_VERTICAL | wx.RIGHT | wx.LEFT, 5)
            max_ctrl = wx.SpinCtrlDouble(self.zones_panel, min=0, max=20000, inc=0.1, value=str(max_val))
            max_ctrl.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_zone_change)
            min_max_sizer.Add(max_ctrl, 1)

            use_zone_ctrl = wx.CheckBox(self.zones_panel, label="Use Zone")
            use_zone_ctrl.SetValue(use_zone)
            use_zone_ctrl.Bind(wx.EVT_CHECKBOX, self.on_zone_change)

            zone_sizer.Add(min_max_sizer, 0, wx.EXPAND | wx.ALL, 5)
            zone_sizer.Add(use_zone_ctrl, 0, wx.EXPAND | wx.ALL, 5)

            self.zones_panel_sizer.Add(zone_sizer, 0, wx.EXPAND | wx.ALL, 5)

            self.zones.append({
                'min': min_ctrl,
                'max': max_ctrl,
                'use_zone': use_zone_ctrl
            })

        self.zones_panel.Layout()
        self.plot_vlines()

    def on_zone_change(self, event):
        self.plot_vlines()

    def on_smooth_change(self, event):
        self.plot_initial_data()

    def plot_vlines(self):
        # Clear existing vlines
        for vline in self.vlines:
            vline.remove()
        self.vlines = []

        # Add vlines for active zones
        for i, zone in enumerate(self.zones):
            if zone['use_zone'].GetValue():
                min_line = self.ax.axvline(zone['min'].GetValue(), color=f'C{i}', linestyle='--', alpha=0.7)
                max_line = self.ax.axvline(zone['max'].GetValue(), color=f'C{i}', linestyle='--', alpha=0.7)
                self.vlines.extend([min_line, max_line])

        self.canvas.draw_idle()

    def on_press(self, event):
        if event.inaxes != self.ax:
            return

        # Check if click is near a vline
        for i, vline in enumerate(self.vlines):
            if abs(event.xdata - vline.get_xdata()[0]) < 5:
                self.dragging = True
                self.current_vline = (i, vline)
                break

    def on_release(self, event):
        if self.dragging:
            self.dragging = False
            self.current_vline = None

    def on_motion(self, event):
        if self.dragging and event.inaxes == self.ax and self.current_vline is not None:
            idx, vline = self.current_vline

            # Update vline position
            vline.set_xdata([event.xdata, event.xdata])

            # Update corresponding zone control
            zone_idx = idx // 2  # Each zone has 2 vlines
            is_min = (idx % 2 == 0)  # Even indices are min vlines

            if is_min:
                self.zones[zone_idx]['min'].SetValue(event.xdata)
            else:
                self.zones[zone_idx]['max'].SetValue(event.xdata)

            self.canvas.draw_idle()

    def plot_initial_data(self):
        self.ax.clear()

        # Get smoothed data if needed
        smooth_value = self.smooth_value.GetValue()
        if smooth_value > 0:
            # Apply Gaussian smoothing
            from scipy.ndimage import gaussian_filter
            y_smooth = gaussian_filter(self.y_values, smooth_value)
        else:
            y_smooth = self.y_values.copy()

        # Plot data
        self.ax.plot(self.x_values, self.y_values, 'k-', label='Raw Data')
        if smooth_value > 0:
            self.ax.plot(self.x_values, y_smooth, 'r-', label='Smoothed Data')

        self.ax.set_xlabel('Wavenumber (cm⁻¹)')
        self.ax.set_ylabel('Intensity')
        self.ax.legend()
        self.ax.set_xlim(min(self.x_values), max(self.x_values))
        self.ax.set_ylim(self.y_min, self.y_max)

        # Plot vertical lines
        self.plot_vlines()

        self.canvas.draw()

    def get_exclusion_zones(self):
        zones = []
        for zone in self.zones:
            if zone['use_zone'].GetValue():
                min_val = min(zone['min'].GetValue(), zone['max'].GetValue())
                max_val = max(zone['min'].GetValue(), zone['max'].GetValue())
                zones.append((min_val, max_val))
        return zones

    def on_fit(self, event):
        smooth_value = self.smooth_value.GetValue()
        exclusion_zones = self.get_exclusion_zones()

        # Get smoothed data if needed
        if smooth_value > 0:
            from scipy.ndimage import gaussian_filter
            y_smooth = gaussian_filter(self.y_values, smooth_value)
        else:
            y_smooth = self.y_values.copy()

        self.fit_als_raman(y_smooth, exclusion_zones)

    def fit_als_raman(self, y_data, exclusion_zones):
        from scipy import sparse
        from scipy.sparse.linalg import spsolve
        import lmfit

        if not exclusion_zones:
            wx.MessageBox("No active exclusion zones defined", "Error", wx.OK | wx.ICON_ERROR)
            return

        # Create mask for data in exclusion zones
        mask = np.zeros_like(self.x_values, dtype=bool)
        for zone in exclusion_zones:
            zone_mask = (self.x_values >= zone[0]) & (self.x_values <= zone[1])
            mask = mask | zone_mask

        # Define the baseline model function for optimization
        def als_baseline(params, x, y, mask):
            lam = params['lam'].value
            p_val = params['p_val'].value
            niter = self.iter_value.GetValue()

            m = len(y)
            D = sparse.diags([1, -2, 1], [-1, 0, 1], shape=(m - 2, m))

            # Initialize with better baseline that doesn't start or end at zero
            # Use linear interpolation between endpoints
            left_idx = np.argmin(np.abs(x - min(x) + 5))  # 5 units from left edge
            right_idx = np.argmin(np.abs(x - max(x) + 5))  # 5 units from right edge
            left_val = np.median(y[:left_idx])
            right_val = np.median(y[right_idx:])
            z = np.linspace(left_val, right_val, len(y))

            for i in range(niter):
                w = p_val * (y > z) + (1 - p_val) * (y <= z)

                # Apply mask - set weights to 0 in excluded regions
                w[mask] = 0

                # Add higher weights to endpoints to prevent dropping to zero
                edge_size = max(3, int(len(w) * 0.02))  # Use 2% of points or minimum 3 points
                w[:edge_size] = 10  # Higher weight at left edge
                w[-edge_size:] = 10  # Higher weight at right edge

                W = sparse.spdiags(w, 0, m, m)
                DtD = D.transpose() @ D
                A = W + lam * DtD
                B = W @ y
                z = spsolve(A, B)

            # Return residuals for fitting
            return z - y

        # Create parameter set
        params = lmfit.Parameters()
        lambda_fixed = self.lambda_fixed.GetValue()
        p_fixed = self.p_fixed.GetValue()

        lambda_val = self.lambda_value.GetValue()
        p_val = self.p_value.GetValue()
        lambda_min = self.lambda_min.GetValue()
        lambda_max = self.lambda_max.GetValue()
        p_min = self.p_min.GetValue()
        p_max = self.p_max.GetValue()

        params.add('lam', value=lambda_val, min=lambda_min, max=lambda_max, vary=not lambda_fixed)
        params.add('p_val', value=p_val, min=p_min, max=p_max, vary=not p_fixed)

        # Perform the fit
        result = lmfit.minimize(als_baseline, params, args=(self.x_values, y_data, mask), method='least_squares')

        # Get final parameters
        final_lambda = result.params['lam'].value
        final_p = result.params['p_val'].value

        # Update the UI controls with the optimized values
        self.lambda_value.SetValue(final_lambda)
        self.p_value.SetValue(final_p)

        print(f"Final lambda: {final_lambda}, Final p: {final_p}")

        # Calculate final background with optimized parameters
        m = len(y_data)
        D = sparse.diags([1, -2, 1], [-1, 0, 1], shape=(m - 2, m))

        # Better initialization for final calculation
        left_idx = np.argmin(np.abs(self.x_values - min(self.x_values) + 5))
        right_idx = np.argmin(np.abs(self.x_values - max(self.x_values) + 5))
        left_val = np.median(y_data[:left_idx])
        right_val = np.median(y_data[right_idx:])
        z = np.linspace(left_val, right_val, len(y_data))

        niter = self.iter_value.GetValue()

        for i in range(niter):
            w = final_p * (y_data > z) + (1 - final_p) * (y_data <= z)
            w[mask] = 0  # Zero weight in excluded zones

            # Add higher weights to endpoints
            edge_size = max(3, int(len(w) * 0.02))
            w[:edge_size] = 10  # Higher weight at left edge
            w[-edge_size:] = 10  # Higher weight at right edge

            W = sparse.spdiags(w, 0, m, m)
            DtD = D.transpose() @ D
            A = W + final_lambda * DtD
            B = W @ y_data
            z = spsolve(A, B)

        # Ensure endpoints don't drop
        z[0] = y_data[0]
        z[-1] = y_data[-1]

        self.plot_als_results(self.x_values, y_data, z, mask)

    def plot_als_results(self, x, y, baseline, mask):
        self.ax.clear()
        smooth_value = self.smooth_value.GetValue()

        # Get raw and smoothed data
        if smooth_value > 0:
            from scipy.ndimage import gaussian_filter
            y_smooth = gaussian_filter(self.y_values, smooth_value)
        else:
            y_smooth = self.y_values.copy()

        # Plot data
        self.ax.plot(x, self.y_values, 'k-', label='Raw Data')
        if smooth_value > 0:
            self.ax.plot(x, y_smooth, 'r-', label='Smoothed Data', alpha=0.5)

        # Plot baseline
        self.ax.plot(x, baseline, 'b-', label='ALS Baseline')

        # Highlight excluded zones
        for zone in self.get_exclusion_zones():
            self.ax.axvspan(zone[0], zone[1], alpha=0.2, color='red',
                            label='Excluded Zone' if zone == self.get_exclusion_zones()[0] else "")

        # Save background for export
        self.background = baseline

        # Plot fitting zones
        self.plot_vlines()

        self.ax.set_xlabel('Wavenumber (cm⁻¹)')
        self.ax.set_ylabel('Intensity')
        self.ax.legend()
        self.ax.set_ylim(self.y_min, self.y_max)
        self.ax.set_xlim(min(x), max(x))

        self.canvas.draw()

    def on_num_zones_change(self, event):
        num = self.num_zones.GetValue()
        self.create_zones(num)
        self.zones_panel.FitInside()

    def on_export(self, event):
        if not hasattr(self, 'background'):
            wx.MessageBox("No background calculated yet", "Error", wx.OK | wx.ICON_ERROR)
            return

        sheet_name = self.parent.parent.sheet_combobox.GetValue()
        if sheet_name in self.parent.parent.Data['Core levels']:
            # First create a background in the parent's context
            self.parent.background_method = "ALS-Raman"
            self.parent.parent.plot_manager.plot_background(self.parent.parent)

            # Then update the background with our calculated values
            core_level_data = self.parent.parent.Data['Core levels'][sheet_name]
            if 'Background' in core_level_data:
                core_level_data['Background']['Bkg Y'] = self.background.tolist()
                core_level_data['Background']['Bkg Type'] = "ALS-Raman"

                # Store ALS parameters in the data
                core_level_data['Background']['ALS_Lambda'] = self.lambda_value.GetValue()
                core_level_data['Background']['ALS_P'] = self.p_value.GetValue()
                core_level_data['Background']['ALS_Iterations'] = self.iter_value.GetValue()

                # Update main window background
                self.parent.parent.background = self.background

                # Set the background min/max energy variables so peak fitting knows a background exists
                self.parent.parent.bg_min_energy = min(self.x_values)
                self.parent.parent.bg_max_energy = max(self.x_values)

                # Make sure the background is set in the parent fitting window
                self.parent.background = self.background

                # Update the main window plot
                self.parent.parent.clear_and_replot()

                wx.MessageBox("Background exported successfully", "Success", wx.OK | wx.ICON_INFORMATION)