# KherveFitting - XPS Data Analysis Software
# Copyright (C) 2024-2026 Gwilherm Kerherve <g.kerherve@ic.ac.uk>
#
# KherveFitting is dual-licensed:
#   - GNU GPL v3.0 (see LICENSE-GPL.txt) for open-source use
#   - Commercial Licence (see LICENSE-COMMERCIAL.txt) for proprietary use
# SPDX-License-Identifier: GPL-3.0-only OR LicenseRef-KherveFitting-Commercial

"""
Tougaard Quantitative XPS Analysis Window

Based on: "Practical guide to the use of backgrounds in quantitative XPS"
by Sven Tougaard, J. Vac. Sci. Technol. A 39, 011201 (2021)

Features:
- Linear baseline removal before analysis (draggable)
- U2 Tougaard background calculation
- Support for up to 5 analysis ranges
"""

import wx
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from scipy.optimize import minimize_scalar
from scipy.integrate import trapezoid


class TougaardAnalysisWindow(wx.Frame):
    """Window for Tougaard-based quantitative XPS depth analysis."""

    D0_HOMOGENEOUS = 23.0
    B0_HOMOGENEOUS = 2866.0
    C_UNIVERSAL = 1643.0
    BKG_OFFSET = 30.0  # Fixed offset from peak centroid to bkg point

    CROSS_SECTIONS = {
        'Universal (metals/oxides)': {'C': 1643.0, 'D': 1.0},
        'SiO2': {'C': 542.0, 'D': 275.0},
        'Polymers': {'C': 551.0, 'D': 436.0},
        'Si': {'C': 542.0, 'D': 275.0},
        'Al': {'C': 542.0, 'D': 275.0},
    }

    MAX_RANGES = 5
    COLORS = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']

    def __init__(self, parent):
        if 'wxMac' in wx.PlatformInfo:
            window_size = (900, 700)
        elif 'wxGTK' in wx.PlatformInfo:
            window_size = (950, 750)
        else:
            window_size = (950, 750)

        super().__init__(parent, title="Tougaard Quantitative XPS Analysis",
                         size=window_size, style=wx.DEFAULT_FRAME_STYLE | wx.RESIZE_BORDER)

        self.parent = parent
        self.current_sheet = None
        self.x_data = None
        self.y_data = None
        self.dragging_line = None
        self.dragging_range_idx = None

        # Storage for multiple analysis ranges results
        self.analysis_results = []

        # Active range index (which tab is selected)
        self.active_range_idx = 0

        # Zoom selection state
        self.zoom_selecting = False
        self.zoom_start = None
        self.zoom_rect = None

        # Find peak mode state
        self.find_peak_mode = False
        self.find_peak_range_idx = None

        # Plot state
        self._plot_initialized = False

        self.InitUI()
        self.Centre()
        self.populate_core_levels()

        if self.core_level_combo.GetCount() > 0:
            self.on_core_level_change(None)

        try:
            from libraries.ConfigFile import set_consistent_fonts
            set_consistent_fonts(self)
        except ImportError:
            pass

    def InitUI(self):
        self.panel = wx.Panel(self, style=wx.BORDER_RAISED)
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Left panel with fixed width
        left_panel = wx.Panel(self.panel)
        left_panel.SetMinSize((280, -1))
        left_panel.SetMaxSize((280, -1))
        left_sizer = wx.BoxSizer(wx.VERTICAL)

        # Core Level Selection
        cl_box = wx.StaticBox(left_panel, label="Core Level")
        cl_sizer = wx.StaticBoxSizer(cl_box, wx.VERTICAL)
        self.core_level_combo = wx.ComboBox(cl_box, style=wx.CB_READONLY)
        self.core_level_combo.Bind(wx.EVT_COMBOBOX, self.on_core_level_change)
        cl_sizer.Add(self.core_level_combo, 0, wx.EXPAND | wx.ALL, 2)
        left_sizer.Add(cl_sizer, 0, wx.EXPAND | wx.ALL, 2)

        # Material Parameters
        mat_box = wx.StaticBox(left_panel, label="Material Parameters")
        mat_sizer = wx.StaticBoxSizer(mat_box, wx.VERTICAL)

        cs_row = wx.BoxSizer(wx.HORIZONTAL)
        cs_row.Add(wx.StaticText(mat_box, label="Material:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.cross_section_combo = wx.ComboBox(mat_box, style=wx.CB_READONLY,
                                                choices=list(self.CROSS_SECTIONS.keys()))
        self.cross_section_combo.SetValue('Universal (metals/oxides)')
        self.cross_section_combo.Bind(wx.EVT_COMBOBOX, self.on_cross_section_change)
        cs_row.Add(self.cross_section_combo, 1, wx.EXPAND)
        mat_sizer.Add(cs_row, 0, wx.EXPAND | wx.ALL, 2)

        imfp_row = wx.BoxSizer(wx.HORIZONTAL)
        imfp_row.Add(wx.StaticText(mat_box, label="IMFP λ (nm):"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.imfp_ctrl = wx.SpinCtrlDouble(mat_box, min=0.1, max=10.0, initial=1.5, inc=0.1)
        self.imfp_ctrl.SetDigits(2)
        imfp_row.Add(self.imfp_ctrl, 1, wx.EXPAND)
        mat_sizer.Add(imfp_row, 0, wx.EXPAND | wx.ALL, 2)

        theta_row = wx.BoxSizer(wx.HORIZONTAL)
        theta_row.Add(wx.StaticText(mat_box, label="Angle θ (°):"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.theta_ctrl = wx.SpinCtrlDouble(mat_box, min=0, max=80, initial=0, inc=5)
        self.theta_ctrl.SetDigits(1)
        theta_row.Add(self.theta_ctrl, 1, wx.EXPAND)
        mat_sizer.Add(theta_row, 0, wx.EXPAND | wx.ALL, 2)

        c_row = wx.BoxSizer(wx.HORIZONTAL)
        c_row.Add(wx.StaticText(mat_box, label="C (eV²):"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.c_ctrl = wx.SpinCtrlDouble(mat_box, min=100, max=5000, initial=1643, inc=10)
        self.c_ctrl.SetDigits(2)
        c_row.Add(self.c_ctrl, 1, wx.EXPAND)
        mat_sizer.Add(c_row, 0, wx.EXPAND | wx.ALL, 2)

        left_sizer.Add(mat_sizer, 0, wx.EXPAND | wx.ALL, 2)

        # Analysis Ranges Section with Notebook (tabs)
        # Each range has its own baseline markers
        ranges_box = wx.StaticBox(left_panel, label="Analysis Ranges")
        ranges_sizer = wx.StaticBoxSizer(ranges_box, wx.VERTICAL)

        # Number of ranges selector
        num_row = wx.BoxSizer(wx.HORIZONTAL)
        num_row.Add(wx.StaticText(ranges_box, label="Ranges:"), 0,
                    wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.num_ranges_spin = wx.SpinCtrl(ranges_box, min=1, max=self.MAX_RANGES, initial=1)
        self.num_ranges_spin.Bind(wx.EVT_SPINCTRL, self.on_num_ranges_change)
        num_row.Add(self.num_ranges_spin, 0)
        ranges_sizer.Add(num_row, 0, wx.EXPAND | wx.ALL, 2)

        # Notebook for range tabs
        self.ranges_notebook = wx.Notebook(ranges_box)
        self.ranges_notebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.on_range_tab_changed)
        self.range_controls = []

        for i in range(self.MAX_RANGES):
            range_panel = self.create_range_panel(self.ranges_notebook, i)
            self.ranges_notebook.AddPage(range_panel, f"R{i + 1}")

        ranges_sizer.Add(self.ranges_notebook, 0, wx.EXPAND | wx.ALL, 2)
        left_sizer.Add(ranges_sizer, 0, wx.EXPAND | wx.ALL, 2)

        # Results
        results_box = wx.StaticBox(left_panel, label="Results")
        results_sizer = wx.StaticBoxSizer(results_box, wx.VERTICAL)
        self.results_text = wx.TextCtrl(results_box, style=wx.TE_MULTILINE | wx.TE_READONLY,
                                         size=(-1, 150))
        self.results_text.SetFont(wx.Font(9, wx.FONTFAMILY_TELETYPE, wx.FONTSTYLE_NORMAL,
                                           wx.FONTWEIGHT_NORMAL))
        results_sizer.Add(self.results_text, 1, wx.EXPAND | wx.ALL, 2)
        copy_btn = wx.Button(results_box, label="Copy Results")
        copy_btn.Bind(wx.EVT_BUTTON, self.on_copy_results)
        results_sizer.Add(copy_btn, 0, wx.EXPAND | wx.ALL, 2)
        left_sizer.Add(results_sizer, 1, wx.EXPAND | wx.ALL, 2)

        # Analyse Button
        analyse_btn = wx.Button(left_panel, label="Analyse All Ranges")
        analyse_btn.SetMinSize((-1, 35))
        analyse_btn.Bind(wx.EVT_BUTTON, self.on_analyse)
        analyse_btn.SetToolTip("Fit Tougaard background for each range\n"
                               "and calculate Ap/B ratio and decay length")
        left_sizer.Add(analyse_btn, 0, wx.EXPAND | wx.ALL, 2)

        left_panel.SetSizer(left_sizer)

        # Right panel - Plot (expandable)
        right_panel = wx.Panel(self.panel)
        right_sizer = wx.BoxSizer(wx.VERTICAL)
        self.figure = Figure(figsize=(5, 4), dpi=100)
        self.canvas = FigureCanvas(right_panel, -1, self.figure)
        right_sizer.Add(self.canvas, 1, wx.EXPAND | wx.ALL, 0)
        right_panel.SetSizer(right_sizer)

        main_sizer.Add(left_panel, 0, wx.EXPAND | wx.ALL, 0)
        main_sizer.Add(right_panel, 1, wx.EXPAND | wx.ALL, 0)
        self.panel.SetSizer(main_sizer)

        self.init_plot()
        self.canvas.mpl_connect('button_press_event', self.on_canvas_press)
        self.canvas.mpl_connect('button_release_event', self.on_canvas_release)
        self.canvas.mpl_connect('motion_notify_event', self.on_canvas_motion)

        # Right-click context menu for zoom
        self.canvas.mpl_connect('button_press_event', self.on_right_click)

        # Keyboard events for Ctrl+Up/Down - bind to panel like TougaardRaman does
        self.panel.Bind(wx.EVT_CHAR_HOOK, self.on_key_press)

    def create_range_panel(self, parent, idx):
        """Create a panel with controls for one analysis range (for notebook tab).
        Each range has its own baseline markers, peak range, and bkg point."""
        color = self.COLORS[idx % len(self.COLORS)]
        panel = wx.Panel(parent)

        sizer = wx.BoxSizer(wx.VERTICAL)

        # Header row with color indicator and Find Peak button
        header = wx.BoxSizer(wx.HORIZONTAL)
        color_indicator = wx.Panel(panel, size=(20, 20))
        color_indicator.SetBackgroundColour(color)
        header.Add(color_indicator, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        header.Add(wx.StaticText(panel, label=f"Range {idx + 1}"), 0, wx.ALIGN_CENTER_VERTICAL)
        header.AddStretchSpacer()

        # Find Peak button
        find_peak_btn = wx.Button(panel, label="Find Peak", size=(70, 22))
        find_peak_btn.SetToolTip("Click on a peak maximum to auto-set range and baseline")
        find_peak_btn.Bind(wx.EVT_BUTTON, lambda evt, i=idx: self.on_find_peak_click(evt, i))
        header.Add(find_peak_btn, 0, wx.ALIGN_CENTER_VERTICAL)
        sizer.Add(header, 0, wx.EXPAND | wx.ALL, 3)

        # Baseline markers for this range
        bl_label = wx.StaticText(panel, label="Baseline (drag markers):")
        bl_label.SetFont(wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_ITALIC, wx.FONTWEIGHT_NORMAL))
        sizer.Add(bl_label, 0, wx.LEFT | wx.TOP, 3)

        bl_start_row = wx.BoxSizer(wx.HORIZONTAL)
        bl_start_row.Add(wx.StaticText(panel, label="Start BE:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        bl_start_ctrl = wx.SpinCtrlDouble(panel, min=0, max=2000, initial=0, inc=0.1)
        bl_start_ctrl.SetDigits(2)
        bl_start_ctrl.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_range_change)
        bl_start_row.Add(bl_start_ctrl, 1, wx.EXPAND)
        sizer.Add(bl_start_row, 0, wx.EXPAND | wx.ALL, 3)

        bl_end_row = wx.BoxSizer(wx.HORIZONTAL)
        bl_end_row.Add(wx.StaticText(panel, label="End BE:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        bl_end_ctrl = wx.SpinCtrlDouble(panel, min=0, max=2000, initial=0, inc=0.1)
        bl_end_ctrl.SetDigits(2)
        bl_end_ctrl.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_range_change)
        bl_end_row.Add(bl_end_ctrl, 1, wx.EXPAND)
        sizer.Add(bl_end_row, 0, wx.EXPAND | wx.ALL, 3)

        # Peak range
        peak_label = wx.StaticText(panel, label="Peak range (drag lines):")
        peak_label.SetFont(wx.Font(8, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_ITALIC, wx.FONTWEIGHT_NORMAL))
        sizer.Add(peak_label, 0, wx.LEFT | wx.TOP, 3)

        ps_row = wx.BoxSizer(wx.HORIZONTAL)
        ps_row.Add(wx.StaticText(panel, label="Peak Start:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        peak_start_ctrl = wx.SpinCtrlDouble(panel, min=0, max=2000, initial=0, inc=0.1)
        peak_start_ctrl.SetDigits(2)
        peak_start_ctrl.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_range_change)
        ps_row.Add(peak_start_ctrl, 1, wx.EXPAND)
        sizer.Add(ps_row, 0, wx.EXPAND | wx.ALL, 3)

        pe_row = wx.BoxSizer(wx.HORIZONTAL)
        pe_row.Add(wx.StaticText(panel, label="Peak End:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        peak_end_ctrl = wx.SpinCtrlDouble(panel, min=0, max=2000, initial=0, inc=0.1)
        peak_end_ctrl.SetDigits(2)
        peak_end_ctrl.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_range_change)
        pe_row.Add(peak_end_ctrl, 1, wx.EXPAND)
        sizer.Add(pe_row, 0, wx.EXPAND | wx.ALL, 3)

        # Bkg Point (read-only, calculated from peak centroid + 30 eV)
        bp_row = wx.BoxSizer(wx.HORIZONTAL)
        bp_row.Add(wx.StaticText(panel, label="Bkg Point:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        bg_point_label = wx.StaticText(panel, label="--")
        bg_point_label.SetToolTip("Auto: Peak Max + 30 eV")
        bp_row.Add(bg_point_label, 1, wx.EXPAND)
        sizer.Add(bp_row, 0, wx.EXPAND | wx.ALL, 3)

        panel.SetSizer(sizer)

        # Store controls reference - now includes baseline markers per range
        self.range_controls.append({
            'panel': panel,
            'bl_start': bl_start_ctrl,
            'bl_end': bl_end_ctrl,
            'bl_start_y': None,  # Y position of start marker
            'bl_end_y': None,    # Y position of end marker
            'peak_start': peak_start_ctrl,
            'peak_end': peak_end_ctrl,
            'bg_point_label': bg_point_label,
            'find_peak_btn': find_peak_btn,
            'color': color
        })

        return panel

    def on_range_tab_changed(self, event):
        """Called when the range tab is changed."""
        self.active_range_idx = self.ranges_notebook.GetSelection()
        self.update_plot()
        event.Skip()

    def on_find_peak_click(self, event, range_idx):
        """Enter find peak mode - wait for user to click on a peak."""
        if self.x_data is None:
            wx.MessageBox("Please select a core level first.", "No Data", wx.OK | wx.ICON_WARNING)
            return

        self.find_peak_mode = True
        self.find_peak_range_idx = range_idx

        # Change cursor to crosshair
        self.canvas.SetCursor(wx.Cursor(wx.CURSOR_CROSS))

        # Update button to show active state
        ctrl = self.range_controls[range_idx]
        ctrl['find_peak_btn'].SetLabel("Click peak...")
        ctrl['find_peak_btn'].SetBackgroundColour(wx.Colour(255, 200, 100))

        # Switch to this tab if not already
        self.ranges_notebook.SetSelection(range_idx)
        self.active_range_idx = range_idx
        self.update_plot()

    def on_find_peak_complete(self, click_be):
        """Complete find peak mode - set up range and baseline based on click position."""
        if self.x_data is None or self.find_peak_range_idx is None:
            return

        range_idx = self.find_peak_range_idx
        ctrl = self.range_controls[range_idx]

        # Reset button appearance
        ctrl['find_peak_btn'].SetLabel("Find Peak")
        ctrl['find_peak_btn'].SetBackgroundColour(wx.NullColour)

        # Reset cursor
        self.canvas.SetCursor(wx.Cursor(wx.CURSOR_ARROW))

        # Find the actual peak maximum near the click position
        # Search within ±3 eV of click
        search_range = 3.0
        search_mask = (self.x_data >= click_be - search_range) & (self.x_data <= click_be + search_range)

        if not np.any(search_mask):
            self.find_peak_mode = False
            self.find_peak_range_idx = None
            return

        x_search = self.x_data[search_mask]
        y_search = self.y_data[search_mask]

        # Find maximum
        max_idx = np.argmax(y_search)
        peak_max_be = x_search[max_idx]

        # Auto-detect peak boundaries by finding where intensity drops significantly
        # Start from peak maximum and go both directions
        peak_max_intensity = y_search[max_idx]

        # Find baseline level (average of edges of search region)
        baseline_estimate = (np.mean(y_search[:3]) + np.mean(y_search[-3:])) / 2
        threshold = baseline_estimate + (peak_max_intensity - baseline_estimate) * 0.1

        # Search for peak start (lower BE side, which is higher index in XPS)
        # Go from peak max towards lower BE
        peak_start_be = peak_max_be - 5.0  # Default
        peak_end_be = peak_max_be + 5.0    # Default

        # Find where intensity crosses threshold on each side
        full_max_idx = np.argmin(np.abs(self.x_data - peak_max_be))

        # Search towards lower BE (higher indices for reversed XPS data)
        for i in range(full_max_idx, min(full_max_idx + 50, len(self.x_data))):
            if self.y_data[i] < threshold:
                peak_end_be = self.x_data[i]
                break

        # Search towards higher BE (lower indices)
        for i in range(full_max_idx, max(full_max_idx - 50, 0), -1):
            if self.y_data[i] < threshold:
                peak_start_be = self.x_data[i]
                break

        # Ensure peak_start < peak_end (in BE terms, start is lower value)
        if peak_start_be > peak_end_be:
            peak_start_be, peak_end_be = peak_end_be, peak_start_be

        # Set peak range controls
        ctrl['peak_start'].SetValue(peak_start_be)
        ctrl['peak_end'].SetValue(peak_end_be)

        # Set baseline markers a few eV before peak_start
        bl_end_be = peak_start_be - 2.0
        bl_start_be = bl_end_be - 5.0

        # Clamp to data range
        bl_start_be = np.clip(bl_start_be, self.x_data.min(), self.x_data.max())
        bl_end_be = np.clip(bl_end_be, self.x_data.min(), self.x_data.max())

        ctrl['bl_start'].SetValue(bl_start_be)
        ctrl['bl_end'].SetValue(bl_end_be)

        # Set Y positions on the data
        ctrl['bl_start_y'] = self.get_baseline_intensity(bl_start_be)
        ctrl['bl_end_y'] = self.get_baseline_intensity(bl_end_be)

        # Reset state
        self.find_peak_mode = False
        self.find_peak_range_idx = None

        # Update display
        self.update_bg_point_labels()
        self.update_plot()

    def on_num_ranges_change(self, event):
        """Show/hide range tabs based on selected number."""
        num_ranges = self.num_ranges_spin.GetValue()

        # Enable/disable pages in notebook
        # We can't really hide notebook pages, so we'll just update visibility
        # The notebook will show all pages but we track which are "active"
        self.update_baseline_endpoints()
        self.update_bg_point_labels()
        self.update_plot()

    def find_peak_centroid(self, peak_start, peak_end):
        """Find the centroid (maximum intensity position) within a peak range."""
        if self.x_data is None or self.y_data is None:
            return (peak_start + peak_end) / 2

        peak_min, peak_max = min(peak_start, peak_end), max(peak_start, peak_end)
        mask = (self.x_data >= peak_min) & (self.x_data <= peak_max)

        if not np.any(mask):
            return (peak_start + peak_end) / 2

        x_peak = self.x_data[mask]
        y_peak = self.y_data[mask]
        peak_idx = np.argmax(y_peak)
        return x_peak[peak_idx]

    def get_active_ranges(self):
        """Get list of active range configurations."""
        num_ranges = self.num_ranges_spin.GetValue()
        ranges = []
        for i in range(num_ranges):
            ctrl = self.range_controls[i]
            peak_start = ctrl['peak_start'].GetValue()
            peak_end = ctrl['peak_end'].GetValue()

            # Find peak centroid (maximum) within the range
            peak_centroid = self.find_peak_centroid(peak_start, peak_end)

            # Bkg point is 30 eV after peak centroid (higher BE)
            bg_point = peak_centroid + self.BKG_OFFSET

            # Ensure within data range
            if self.x_data is not None:
                bg_point = np.clip(bg_point, self.x_data.min(), self.x_data.max())

            # Get baseline markers for this range
            bl_start = ctrl['bl_start'].GetValue()
            bl_end = ctrl['bl_end'].GetValue()
            bl_start_y = ctrl.get('bl_start_y')
            bl_end_y = ctrl.get('bl_end_y')

            # If Y values not set, get from data
            if bl_start_y is None:
                bl_start_y = self.get_baseline_intensity(bl_start)
                ctrl['bl_start_y'] = bl_start_y
            if bl_end_y is None:
                bl_end_y = self.get_baseline_intensity(bl_end)
                ctrl['bl_end_y'] = bl_end_y

            ranges.append({
                'idx': i,
                'peak_start': peak_start,
                'peak_end': peak_end,
                'peak_centroid': peak_centroid,
                'bg_point': bg_point,
                'bl_start': bl_start,
                'bl_end': bl_end,
                'bl_start_y': bl_start_y,
                'bl_end_y': bl_end_y,
                'color': ctrl['color']
            })
        return ranges

    def update_bg_point_labels(self):
        """Update the bkg point labels for all ranges."""
        ranges = self.get_active_ranges()
        for r in ranges:
            self.range_controls[r['idx']]['bg_point_label'].SetLabel(
                f"{r['bg_point']:.2f} eV (from {r['peak_centroid']:.2f})")

    def update_baseline_endpoints(self):
        """Initialize baseline endpoints for each range based on its peak position."""
        if self.x_data is None:
            return

        data_min = self.x_data.min()
        data_max = self.x_data.max()

        for i, ctrl in enumerate(self.range_controls):
            peak_start = ctrl['peak_start'].GetValue()

            # Both markers are BEFORE the peak (lower BE = before peak in XPS)
            # Second marker: 2 eV before peak start
            baseline_end = peak_start - 2.0
            # First marker: 5 eV before second marker (7 eV before peak start)
            baseline_start = baseline_end - 5.0

            # Ensure markers stay within data range
            if baseline_start < data_min:
                baseline_start = data_min
                baseline_end = baseline_start + 5.0
            if baseline_end > data_max:
                baseline_end = data_max
                baseline_start = baseline_end - 5.0
                if baseline_start < data_min:
                    baseline_start = data_min

            ctrl['bl_start'].SetValue(baseline_start)
            ctrl['bl_end'].SetValue(baseline_end)

            # Initialize Y values on the data
            ctrl['bl_start_y'] = self.get_baseline_intensity(baseline_start)
            ctrl['bl_end_y'] = self.get_baseline_intensity(baseline_end)

    def get_baseline_intensity(self, be_value):
        """Get the intensity at a given BE value (average over 3 points)."""
        if self.x_data is None or self.y_data is None:
            return 0
        idx = np.argmin(np.abs(self.x_data - be_value))
        return np.mean(self.y_data[max(0, idx - 1):idx + 2])

    def calculate_linear_baseline_for_range(self, range_idx):
        """
        Calculate linear baseline for a specific range.
        Returns baseline parameters including slope.
        """
        if self.x_data is None or range_idx >= len(self.range_controls):
            return None

        ctrl = self.range_controls[range_idx]
        bl_start_be = ctrl['bl_start'].GetValue()
        bl_end_be = ctrl['bl_end'].GetValue()

        # Use stored Y positions for the markers (can be off-data)
        bl_start_y = ctrl.get('bl_start_y')
        bl_end_y = ctrl.get('bl_end_y')

        if bl_start_y is None:
            bl_start_y = self.get_baseline_intensity(bl_start_be)
            ctrl['bl_start_y'] = bl_start_y
        if bl_end_y is None:
            bl_end_y = self.get_baseline_intensity(bl_end_be)
            ctrl['bl_end_y'] = bl_end_y

        # Calculate gradient (slope) from the two marker positions
        if abs(bl_end_be - bl_start_be) > 0.01:
            slope = (bl_end_y - bl_start_y) / (bl_end_be - bl_start_be)
        else:
            slope = 0

        # Line extends 15 eV past the second marker (towards higher BE)
        line_extend_be = bl_end_be + 15.0
        if self.x_data is not None:
            line_extend_be = np.clip(line_extend_be, self.x_data.min(), self.x_data.max())
        line_extend_y = bl_end_y + slope * (line_extend_be - bl_end_be)

        return {
            'bl_start_be': bl_start_be,
            'bl_end_be': bl_end_be,
            'bl_start_y': bl_start_y,
            'bl_end_y': bl_end_y,
            'line_extend_be': line_extend_be,
            'line_extend_y': line_extend_y,
            'slope': slope
        }

    def init_plot(self):
        self.ax = self.figure.add_subplot(111)
        self.ax.set_xlabel('Binding Energy (eV)')
        self.ax.set_ylabel('Intensity (a.u.)')
        self.ax.ticklabel_format(axis='y', style='scientific', scilimits=(0, 0))
        self.figure.tight_layout(pad=0.5)
        self.figure.subplots_adjust(left=0.07, right=0.98, top=0.97, bottom=0.07)
        self.canvas.draw()

    def populate_core_levels(self):
        self.core_level_combo.Clear()
        if hasattr(self.parent, 'Data') and 'Core levels' in self.parent.Data:
            sheets = list(self.parent.Data['Core levels'].keys())
            self.core_level_combo.AppendItems(sheets)
            if sheets:
                if hasattr(self.parent, 'sheet_combobox'):
                    current = self.parent.sheet_combobox.GetValue()
                    if current in sheets:
                        self.core_level_combo.SetValue(current)
                    else:
                        self.core_level_combo.SetValue(sheets[0])
                else:
                    self.core_level_combo.SetValue(sheets[0])

    def calculate_imfp_tpp2m(self, kinetic_energy):
        """Calculate IMFP using TPP-2M formula with average matrix parameters."""
        N_v = 4.684
        rho = 6.767
        M = 137.51
        E_g = 0

        E_p = 28.8 * np.sqrt((N_v * rho) / M)
        U = N_v * rho / M

        beta = -0.10 + 0.944 / (E_p ** 2 + E_g ** 2) ** 0.5 + 0.069 * rho ** 0.1
        gamma = 0.191 * rho ** (-0.5)
        C = 1.97 - 0.91 * U
        D = 53.4 - 20.8 * U

        imfp = kinetic_energy / (E_p ** 2 * (
                beta * np.log(gamma * kinetic_energy) -
                (C / kinetic_energy) +
                (D / kinetic_energy ** 2))) / 10

        return imfp

    def on_core_level_change(self, event):
        sheet_name = self.core_level_combo.GetValue()
        if not sheet_name:
            return
        try:
            core_data = self.parent.Data['Core levels'][sheet_name]
            self.x_data = np.array(core_data['B.E.'])
            self.y_data = np.array(core_data['Raw Data'])
            self.current_sheet = sheet_name
            self.y_data_corrected = None
            self.linear_baseline = None
            self.analysis_results = []
            self._plot_initialized = False  # Reset so new data gets proper limits

            self.auto_detect_range()

            # Calculate IMFP
            peak_idx = np.argmax(self.y_data)
            peak_be = self.x_data[peak_idx]
            photon_energy = getattr(self.parent, 'photons', 1486.68)
            kinetic_energy = photon_energy - peak_be
            imfp = 1.5
            if kinetic_energy > 0:
                imfp = self.calculate_imfp_tpp2m(kinetic_energy)
                self.imfp_ctrl.SetValue(imfp)

            self.update_bg_point_labels()
            self.update_baseline_endpoints()
            self.update_plot(preserve_limits=False)  # Don't preserve limits for new data
            self.results_text.SetValue(f"Loaded: {sheet_name}\n"
                                       f"Peak BE: {peak_be:.2f} eV\n"
                                       f"KE: {kinetic_energy:.2f} eV\n"
                                       f"λ (TPP-2M): {imfp:.2f} nm\n\n"
                                       f"Click 'Analyse All Ranges' to run analysis")
        except Exception as e:
            wx.MessageBox(f"Error loading data: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def on_cross_section_change(self, event):
        cs_name = self.cross_section_combo.GetValue()
        if cs_name in self.CROSS_SECTIONS:
            self.c_ctrl.SetValue(self.CROSS_SECTIONS[cs_name]['C'])

    def auto_detect_range(self):
        """Auto-detect peak range for first range."""
        if self.x_data is None:
            return
        peak_idx = np.argmax(self.y_data)
        peak_be = self.x_data[peak_idx]
        peak_start = np.clip(peak_be - 5, self.x_data.min(), self.x_data.max())
        peak_end = np.clip(peak_be + 5, self.x_data.min(), self.x_data.max())

        # Set first range
        self.range_controls[0]['peak_start'].SetValue(peak_start)
        self.range_controls[0]['peak_end'].SetValue(peak_end)

    def calculate_u2_tougaard_background(self, x, y, C_value, target_be):
        """
        Calculate U2-Tougaard background.
        Exactly as in original file.
        """
        baseline = np.mean(y[-5:])
        y_shifted = y - baseline
        target_idx = np.argmin(np.abs(x - target_be))
        target_intensity = y[target_idx]
        dx = np.mean(np.diff(x))

        def calculate_background(B_val):
            bg = np.zeros_like(y)
            for i in range(len(x)):
                E_prime_minus_E = x[i:] - x[i]
                K = B_val * E_prime_minus_E / ((C_value + E_prime_minus_E ** 2) ** 2 + 1e-10)
                bg[i] = np.trapz(K * y_shifted[i:], dx=dx)
            return bg + baseline

        def objective(B_val):
            try:
                bg = calculate_background(B_val)
                return (bg[target_idx] - target_intensity) ** 2
            except:
                return 1e10

        result = minimize_scalar(objective, bounds=(100, 1000000), method='bounded')
        B_fitted = result.x
        background = calculate_background(B_fitted)
        return background, B_fitted

    def on_canvas_press(self, event):
        if event.inaxes != self.ax or event.button != 1:
            return

        # Handle find peak mode
        if self.find_peak_mode:
            self.on_find_peak_complete(event.xdata)
            return

        # Handle zoom selection mode
        if self.zoom_selecting:
            self.zoom_start = (event.xdata, event.ydata)
            return

        x_click = event.xdata
        xlim = self.ax.get_xlim()
        tol_x = abs(xlim[1] - xlim[0]) * 0.015  # Smaller tolerance

        # Only interact with the active range's markers
        active_idx = self.active_range_idx
        if active_idx < len(self.range_controls):
            ctrl = self.range_controls[active_idx]

            # Get all draggable positions
            bl_start = ctrl['bl_start'].GetValue()
            bl_end = ctrl['bl_end'].GetValue()
            peak_start = ctrl['peak_start'].GetValue()
            peak_end = ctrl['peak_end'].GetValue()

            # Calculate distances to all elements
            distances = [
                ('peak_start', abs(x_click - peak_start)),
                ('peak_end', abs(x_click - peak_end)),
                ('bl_start', abs(x_click - bl_start)),
                ('bl_end', abs(x_click - bl_end)),
            ]

            # Sort by distance and pick the closest one within tolerance
            distances.sort(key=lambda x: x[1])
            closest_name, closest_dist = distances[0]

            if closest_dist < tol_x:
                self.dragging_line = closest_name
                self.dragging_range_idx = active_idx
                return

    def on_canvas_release(self, event):
        # Handle zoom selection completion
        if self.zoom_selecting and self.zoom_start is not None:
            if event.xdata is not None and event.ydata is not None:
                x0, y0 = self.zoom_start
                x1, y1 = event.xdata, event.ydata

                # Only zoom if the selection is large enough
                if abs(x1 - x0) > 0.5 and abs(y1 - y0) > 0:
                    self.ax.set_xlim(max(x0, x1), min(x0, x1))  # Reversed for XPS
                    self.ax.set_ylim(min(y0, y1), max(y0, y1))
                    self.canvas.draw()

            self.zoom_selecting = False
            self.zoom_start = None
            self.canvas.SetCursor(wx.Cursor(wx.CURSOR_ARROW))
            if self.zoom_rect:
                self.zoom_rect.remove()
                self.zoom_rect = None
            return

        if self.dragging_line:
            self.dragging_line = None
            self.dragging_range_idx = None
            self.update_bg_point_labels()
            self.update_plot()

    def on_canvas_motion(self, event):
        if event.inaxes != self.ax or event.xdata is None:
            return

        # Handle zoom selection rectangle drawing
        if self.zoom_selecting and self.zoom_start is not None:
            x0, y0 = self.zoom_start
            x1, y1 = event.xdata, event.ydata

            # Remove old rectangle
            if self.zoom_rect:
                self.zoom_rect.remove()

            # Draw new rectangle
            from matplotlib.patches import Rectangle
            width = x1 - x0
            height = y1 - y0
            self.zoom_rect = Rectangle((x0, y0), width, height,
                                        fill=False, edgecolor='red', linestyle='--', linewidth=1)
            self.ax.add_patch(self.zoom_rect)
            self.canvas.draw()
            return

        if not self.dragging_line:
            return

        idx = self.dragging_range_idx
        if idx is None or idx >= len(self.range_controls):
            return

        ctrl = self.range_controls[idx]

        # Start marker stays on data (only X moves, Y follows data)
        if self.dragging_line == 'bl_start':
            ctrl['bl_start'].SetValue(event.xdata)
            ctrl['bl_start_y'] = self.get_baseline_intensity(event.xdata)  # Stay on data
            self.update_plot()
            return
        # End marker can move freely in both X and Y
        elif self.dragging_line == 'bl_end':
            ctrl['bl_end'].SetValue(event.xdata)
            ctrl['bl_end_y'] = event.ydata  # Free Y movement
            self.update_plot()
            return
        elif self.dragging_line == 'peak_start':
            ctrl['peak_start'].SetValue(event.xdata)
        elif self.dragging_line == 'peak_end':
            ctrl['peak_end'].SetValue(event.xdata)

        self.update_bg_point_labels()
        self.update_plot()

    def on_range_change(self, event):
        self.update_bg_point_labels()
        self.update_plot()

    def update_plot(self, preserve_limits=True):
        # Store current limits before clearing
        if preserve_limits and hasattr(self, '_plot_initialized') and self._plot_initialized:
            old_xlim = self.ax.get_xlim()
            old_ylim = self.ax.get_ylim()
        else:
            old_xlim = None
            old_ylim = None

        self.ax.clear()
        if self.x_data is None:
            self.canvas.draw()
            return

        # Plot raw data
        self.ax.plot(self.x_data, self.y_data, 'k-', label='Raw Data', linewidth=1)

        num_ranges = self.num_ranges_spin.GetValue()
        active_idx = self.active_range_idx

        # Plot all active ranges, but only show markers/lines for the currently selected tab
        for i in range(num_ranges):
            ctrl = self.range_controls[i]
            color = ctrl['color']
            peak_start = ctrl['peak_start'].GetValue()
            peak_end = ctrl['peak_end'].GetValue()

            # Calculate bg_point
            peak_centroid = self.find_peak_centroid(peak_start, peak_end)
            bg_point = peak_centroid + self.BKG_OFFSET
            if self.x_data is not None:
                bg_point = np.clip(bg_point, self.x_data.min(), self.x_data.max())

            is_active = (i == active_idx)
            alpha = 0.9 if is_active else 0.3
            linewidth = 1.5 if is_active else 0.8

            # Peak range markers (vertical lines)
            self.ax.axvline(peak_start, color=color, linestyle='--', alpha=alpha, linewidth=linewidth)
            self.ax.axvline(peak_end, color=color, linestyle='--', alpha=alpha, linewidth=linewidth)
            self.ax.axvline(bg_point, color=color, linestyle=':', alpha=alpha, linewidth=linewidth + 0.5)

            # Shaded peak region
            self.ax.axvspan(min(peak_start, peak_end), max(peak_start, peak_end),
                            alpha=0.1 if is_active else 0.03, color=color)

            # Only show baseline markers for the active range
            if is_active:
                bl_result = self.calculate_linear_baseline_for_range(i)
                if bl_result is not None:
                    # Plot baseline line
                    self.ax.plot([bl_result['bl_start_be'], bl_result['bl_end_be'], bl_result['line_extend_be']],
                                 [bl_result['bl_start_y'], bl_result['bl_end_y'], bl_result['line_extend_y']],
                                 'g-', label='Baseline', linewidth=1.5, alpha=0.7)

                    # Draw draggable markers (smaller size)
                    self.ax.plot(bl_result['bl_start_be'], bl_result['bl_start_y'], 'ms',
                                 markersize=7, markeredgecolor='black')#, label='Baseline Points')
                    self.ax.plot(bl_result['bl_end_be'], bl_result['bl_end_y'], 'ms',
                                 markersize=7, markeredgecolor='black')

        # Plot analysis results (Tougaard backgrounds) if available
        for result in self.analysis_results:
            if result.get('background_with_slope') is not None:
                bg = result['background_with_slope']
                color = result['color']
                bl_start_be = result['bl_start_be']
                bg_point = result['bg_point']

                # Plot background from bl_start_be to bg_point
                min_be = min(bl_start_be, bg_point)
                max_be = max(bl_start_be, bg_point)
                mask = (self.x_data >= min_be) & (self.x_data <= max_be)

                self.ax.plot(self.x_data[mask], bg[mask], '-',
                             color=color, label=f"Tougaard {result['idx'] + 1}",
                             linewidth=1.5)

        self.ax.set_xlabel('Binding Energy (eV)')
        self.ax.set_ylabel('Intensity (a.u.)')
        # self.ax.set_title(self.current_sheet if self.current_sheet else '')
        self.ax.legend(loc='upper right', fontsize=8)

        # Restore limits if preserving, otherwise set default
        if old_xlim is not None and old_ylim is not None:
            self.ax.set_xlim(old_xlim)
            self.ax.set_ylim(old_ylim)
        else:
            self.ax.set_xlim(self.x_data.max(), self.x_data.min())
            y_min = np.min(self.y_data)
            y_max = np.max(self.y_data)
            y_margin = (y_max - y_min) * 0.05
            self.ax.set_ylim(y_min - y_margin, y_max + y_margin)
            self._plot_initialized = True

        self.ax.ticklabel_format(axis='y', style='scientific', scilimits=(0, 0))
        self.figure.subplots_adjust(left=0.07, right=0.98, top=0.97, bottom=0.07)
        self.canvas.draw()

    def on_analyse(self, event):
        """Run full Tougaard analysis for all active ranges independently."""
        if self.x_data is None:
            wx.MessageBox("Please select a core level first.", "No Data", wx.OK | wx.ICON_WARNING)
            return

        num_ranges = self.num_ranges_spin.GetValue()
        if num_ranges == 0:
            wx.MessageBox("No analysis ranges defined.", "Error", wx.OK | wx.ICON_WARNING)
            return

        C = self.c_ctrl.GetValue()
        lambda_nm = self.imfp_ctrl.GetValue()
        theta_deg = self.theta_ctrl.GetValue()

        try:
            # Clear previous results
            self.analysis_results = []

            # Analyse each range independently with its own baseline
            for i in range(num_ranges):
                ctrl = self.range_controls[i]

                # Get baseline parameters for this range
                bl_result = self.calculate_linear_baseline_for_range(i)
                if bl_result is None:
                    continue

                bl_start_be = bl_result['bl_start_be']
                slope = bl_result['slope']

                # Get raw data value at bl_start_be
                raw_data_at_start = self.get_baseline_intensity(bl_start_be)

                # Get range parameters
                peak_start = ctrl['peak_start'].GetValue()
                peak_end = ctrl['peak_end'].GetValue()
                peak_centroid = self.find_peak_centroid(peak_start, peak_end)
                bg_point = peak_centroid + self.BKG_OFFSET
                if self.x_data is not None:
                    bg_point = np.clip(bg_point, self.x_data.min(), self.x_data.max())

                range_config = {
                    'idx': i,
                    'peak_start': peak_start,
                    'peak_end': peak_end,
                    'bg_point': bg_point,
                    'bl_start_be': bl_start_be,
                    'color': ctrl['color']
                }

                result = self.analyse_single_range(range_config, C, lambda_nm, theta_deg,
                                                   slope, bl_start_be, raw_data_at_start)
                self.analysis_results.append(result)

            # Generate combined report
            report = self.generate_combined_report(lambda_nm)
            self.results_text.SetValue(report)

            # Update plot with all results
            self.update_plot()

            # Add peak area visualizations (on raw data scale)
            for result in self.analysis_results:
                if result.get('x_peak') is not None and result.get('y_peak_display') is not None:
                    self.ax.fill_between(result['x_peak'], result['lin_bg_display'],
                                         result['y_peak_display'],
                                         alpha=0.3, color=result['color'])

            self.canvas.draw()

        except Exception as e:
            import traceback
            wx.MessageBox(f"Error: {str(e)}\n\n{traceback.format_exc()}",
                          "Error", wx.OK | wx.ICON_ERROR)

    def analyse_single_range(self, range_config, C, lambda_nm, theta_deg,
                              slope, bl_start_be, raw_data_at_start):
        """Analyse a single range independently with its own baseline."""
        idx = range_config['idx']
        peak_start = range_config['peak_start']
        peak_end = range_config['peak_end']
        bg_point = range_config['bg_point']
        color = range_config['color']

        result = {
            'idx': idx,
            'peak_start': peak_start,
            'peak_end': peak_end,
            'bg_point': bg_point,
            'bl_start_be': bl_start_be,
            'color': color,
            'background': None,
            'background_with_slope': None,
            'fitted_B': None,
            'peak_area': None,
            'B_increase': None,
            'ap_b_ratio': None,
            'L_lambda': None,
            'L_nm': None,
            'x_peak': None,
            'y_peak': None,
            'y_peak_display': None,
            'lin_bg': None,
            'lin_bg_display': None
        }

        # Extract data between Start BE marker and Bkg Point for Tougaard calculation
        min_be = min(bl_start_be, bg_point)
        max_be = max(bl_start_be, bg_point)
        range_mask = (self.x_data >= min_be) & (self.x_data <= max_be)

        x_range = self.x_data[range_mask]
        y_raw_range = self.y_data[range_mask]

        if len(x_range) < 5:
            return result

        # Create subtraction line for this range based on the baseline slope
        # The line passes through (bl_start_be, raw_data_at_start) with given slope
        subtraction_line_range = raw_data_at_start + slope * (x_range - bl_start_be)

        # Subtract the baseline slope from the raw data in this range
        y_corrected_range = y_raw_range - subtraction_line_range

        # Calculate Tougaard background on the corrected data
        background_range, fitted_B = self.calculate_u2_tougaard_background(
            x_range, y_corrected_range, C, bg_point)

        if background_range is None:
            return result

        # Add the subtraction line back to get background in raw data coordinates
        background_with_slope_range = background_range + subtraction_line_range

        # Create full arrays
        background = np.zeros_like(self.x_data)
        background[range_mask] = background_range

        background_with_slope = np.zeros_like(self.x_data)
        background_with_slope[range_mask] = background_with_slope_range

        result['background'] = background
        result['background_with_slope'] = background_with_slope
        result['fitted_B'] = fitted_B

        # Calculate peak area using corrected data in the peak region
        peak_min, peak_max = min(peak_start, peak_end), max(peak_start, peak_end)
        peak_mask_in_range = (x_range >= peak_min) & (x_range <= peak_max)
        x_peak = x_range[peak_mask_in_range]
        y_peak = y_corrected_range[peak_mask_in_range]

        if len(x_peak) < 3:
            return result

        sort_idx = np.argsort(x_peak)
        x_peak, y_peak = x_peak[sort_idx], y_peak[sort_idx]
        lin_bg = np.interp(x_peak, [x_peak[0], x_peak[-1]], [y_peak[0], y_peak[-1]])
        peak_area = np.abs(trapezoid(y_peak - lin_bg, x_peak))

        result['x_peak'] = x_peak
        result['y_peak'] = y_peak
        result['lin_bg'] = lin_bg
        result['peak_area'] = peak_area

        # Calculate display versions (transform back to raw data scale)
        sub_line_peak = raw_data_at_start + slope * (x_peak - bl_start_be)
        result['y_peak_display'] = y_peak + sub_line_peak
        result['lin_bg_display'] = lin_bg + sub_line_peak

        # Calculate B increase at bg_point using corrected data
        baseline = np.mean(y_corrected_range[-5:])
        bg_idx = np.argmin(np.abs(x_range - bg_point))
        B_increase = y_corrected_range[bg_idx] - baseline
        result['B_increase'] = B_increase

        # Ap/B ratio
        ap_b_ratio = peak_area / B_increase if B_increase > 0 else float('inf')
        result['ap_b_ratio'] = ap_b_ratio

        # Decay length
        cos_theta = np.cos(np.radians(theta_deg))
        B0 = self.B0_HOMOGENEOUS
        if abs(B0 - fitted_B) > 1e-10:
            L_lambda = (fitted_B / (B0 - fitted_B)) * cos_theta
            L_nm = L_lambda * lambda_nm
        else:
            L_lambda, L_nm = float('inf'), float('inf')

        result['L_lambda'] = L_lambda
        result['L_nm'] = L_nm

        return result

    def generate_combined_report(self, lambda_nm):
        """Generate analysis report for all ranges."""
        B0 = self.B0_HOMOGENEOUS

        report = f"""TOUGAARD DEPTH ANALYSIS
{'=' * 50}
Core Level: {self.current_sheet}
IMFP (λ): {lambda_nm:.2f} nm | C: {self.c_ctrl.GetValue():.2f} eV²
Angle θ: {self.theta_ctrl.GetValue():.2f}°

INTERPRETATION GUIDE
--------------------
Ap/B (eV)  | Depth Distribution
-----------+----------------------
  > 30     | Surface localized
  20-30    | Uniform distribution
  < 20     | Subsurface/buried

L (λ)      | Concentration Profile
-----------+----------------------
  L > 0    | Surface enriched
  |L| > 6  | Nearly uniform
  L < 0    | Subsurface/buried
"""

        for result in self.analysis_results:
            idx = result['idx']
            fitted_B = result.get('fitted_B')
            ap_b_ratio = result.get('ap_b_ratio')
            L_lambda = result.get('L_lambda')
            L_nm = result.get('L_nm')
            peak_area = result.get('peak_area')
            bl_start_be = result.get('bl_start_be', 0)

            if fitted_B is None:
                report += f"\n--- Range {idx + 1}: Analysis failed ---\n"
                continue

            # Interpretation
            if ap_b_ratio is not None and ap_b_ratio != float('inf'):
                if ap_b_ratio > 30:
                    apb_interp = "Surface localized"
                elif ap_b_ratio < 20:
                    apb_interp = "Subsurface/buried"
                else:
                    apb_interp = "Uniform distribution"
            else:
                apb_interp = "N/A"

            if L_lambda is not None and L_lambda != float('inf'):
                if abs(L_lambda) > 6:
                    L_interp = "Nearly uniform"
                elif L_lambda > 0:
                    L_interp = "Surface enriched"
                else:
                    L_interp = "Subsurface/buried"
            else:
                L_interp = "N/A"

            # Format values safely (handle None)
            peak_area_str = f"{peak_area:.2f}" if peak_area is not None else "N/A"
            ap_b_str = f"{ap_b_ratio:.2f}" if ap_b_ratio is not None else "N/A"
            L_lambda_str = f"{L_lambda:.2f}" if L_lambda is not None else "N/A"
            L_nm_str = f"{L_nm:.2f}" if L_nm is not None else "N/A"

            report += f"""
--- RANGE {idx + 1} ---
Baseline Start: {bl_start_be:.2f} eV
Peak: {result['peak_start']:.2f} to {result['peak_end']:.2f} eV
Bkg Point: {result['bg_point']:.2f} eV

B (fitted): {fitted_B:.2f} eV² (B₀={B0:.2f})
Peak Area: {peak_area_str}
Ap/B: {ap_b_str} eV → {apb_interp}
L: {L_lambda_str}λ = {L_nm_str} nm → {L_interp}
"""

        return report

    def on_copy_results(self, event):
        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(wx.TextDataObject(self.results_text.GetValue()))
            wx.TheClipboard.Close()

    def on_right_click(self, event):
        """Handle right-click for zoom context menu."""
        if event.button != 3:  # Not right click
            return
        if event.inaxes != self.ax:
            return

        # Create context menu
        menu = wx.Menu()
        zoom_in_item = menu.Append(wx.ID_ANY, "Zoom In (select area)")
        zoom_out_item = menu.Append(wx.ID_ANY, "Zoom Out")
        menu.AppendSeparator()
        reset_zoom_item = menu.Append(wx.ID_ANY, "Reset Zoom")

        self.Bind(wx.EVT_MENU, self.on_zoom_in_select, zoom_in_item)
        self.Bind(wx.EVT_MENU, self.on_zoom_out, zoom_out_item)
        self.Bind(wx.EVT_MENU, self.on_reset_zoom, reset_zoom_item)

        self.PopupMenu(menu)
        menu.Destroy()

    def on_zoom_in_select(self, event):
        """Enable zoom selection mode."""
        self.zoom_selecting = True
        self.canvas.SetCursor(wx.Cursor(wx.CURSOR_CROSS))

    def on_zoom_out(self, event):
        """Zoom out by 20%."""
        if self.x_data is None:
            return

        xlim = self.ax.get_xlim()
        ylim = self.ax.get_ylim()

        # Expand by 20%
        x_range = abs(xlim[1] - xlim[0])
        y_range = abs(ylim[1] - ylim[0])

        x_center = (xlim[0] + xlim[1]) / 2
        y_center = (ylim[0] + ylim[1]) / 2

        new_x_range = x_range * 1.2
        new_y_range = y_range * 1.2

        self.ax.set_xlim(x_center + new_x_range / 2, x_center - new_x_range / 2)
        self.ax.set_ylim(y_center - new_y_range / 2, y_center + new_y_range / 2)
        self.canvas.draw()

    def on_reset_zoom(self, event):
        """Reset zoom to show full data range."""
        if self.x_data is None:
            return

        self.ax.set_xlim(self.x_data.max(), self.x_data.min())
        y_min = np.min(self.y_data)
        y_max = np.max(self.y_data)
        y_margin = (y_max - y_min) * 0.05
        self.ax.set_ylim(y_min - y_margin, y_max + y_margin)
        self.canvas.draw()

    def on_key_press(self, event):
        """Handle key press events for Ctrl+Up/Down intensity adjustment."""
        if event.ControlDown():
            if event.GetKeyCode() in [wx.WXK_UP, wx.WXK_DOWN]:
                # Adjust y limits to zoom in/out
                if self.x_data is None:
                    event.Skip()
                    return

                ylim = self.ax.get_ylim()
                y_range = abs(ylim[1] - ylim[0])
                intensity_factor = 0.1

                if event.GetKeyCode() == wx.WXK_DOWN:
                    # Decrease Y max (zoom in on Y - see more noise detail)
                    new_ymax = ylim[1] - intensity_factor * y_range
                else:  # WXK_UP
                    # Increase Y max (zoom out on Y)
                    new_ymax = ylim[1] + intensity_factor * y_range

                if new_ymax > ylim[0]:
                    self.ax.set_ylim(ylim[0], new_ymax)
                    self.canvas.draw_idle()
                return
        event.Skip()