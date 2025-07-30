import wx
from libraries.FileMenu.Save import save_state
from Functions import remove_peak
import numpy as np

class BackgroundWindow(wx.Frame):
    def __init__(self, parent, *args, **kw):
        super(BackgroundWindow, self).__init__(parent, *args, **kw, style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX | wx.MINIMIZE_BOX | wx.SYSTEM_MENU) | wx.STAY_ON_TOP)
        self.parent = parent
        self.SetTitle("Measure Area")
        if 'wxMac' in wx.PlatformInfo:
            self.SetSize((260, 385))  # Increased height to accommodate new elements
            self.SetMinSize((260, 385))
            self.SetMaxSize((260, 385))
        # elif 'wxGTK' in wx.PlatformInfo:
        #     self.SetSize((260, 470))  # Increased height to accommodate new elements
        #     self.SetMinSize((260, 470))
        #     self.SetMaxSize((260, 470))
        elif 'wxGTK' in wx.PlatformInfo:  # This is for Linux
            desktop = self.get_linux_desktop()
            if desktop == 'gnome':
                self.SetSize((310, 560))
            elif desktop == 'kde':
                self.SetSize((260, 480))
            elif desktop == 'xfce':
                print('linux xfce')
                self.SetSize((260, 480))
            else:  # Unknown or other
                self.SetSize((280, 520))
            print(f'GTK environment: {desktop}')
        else:
            self.SetSize((267, 400))
            self.SetMinSize((267, 400))
            self.SetMaxSize((267, 400))

        panel = wx.Panel(self)

        def detect_dark_mode():
            if 'wxMac' in wx.PlatformInfo:  # Mac
                return wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW).GetLuminance() < 0.5
            elif 'wxMSW' in wx.PlatformInfo:  # Windows
                return wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW).GetLuminance() < 0.5
            elif 'wxGTK' in wx.PlatformInfo:  # Linux
                return wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW).GetLuminance() < 0.5
            return False

        if not detect_dark_mode():
            panel.SetBackgroundColour(wx.Colour(250, 250, 230))

        # Create controls
        method_label = wx.StaticText(panel, label="Method:")
        self.method_combobox = wx.ComboBox(panel, choices=["Multi-Regions Smart",
                                                           # "Smart", "Shirley", "Linear",
                                                           '1x U4-Tougaard', '2x U4-Tougaard', '3x U4-Tougaard'],
                                           style=wx.CB_READONLY)
        self.method_combobox.SetSelection(0)  # Default to Shirley
        self.method_combobox.SetMaxSize((125,25))

        offset_h_label = wx.StaticText(panel, label="Offset (H):")
        self.offset_h_text = wx.TextCtrl(panel, value="0")

        offset_l_label = wx.StaticText(panel, label="Offset (L):")
        self.offset_l_text = wx.TextCtrl(panel, value="0")

        self.offset_h_text.Bind(wx.EVT_TEXT, self.on_offset_changed)
        self.offset_l_text.Bind(wx.EVT_TEXT, self.on_offset_changed)

        self.min_range_label = wx.StaticText(panel, label='Min Range:')
        self.min_range_text = wx.TextCtrl(panel, value="0.00")
        self.min_range_text.Bind(wx.EVT_TEXT, self.on_min_range_change)

        self.max_range_label = wx.StaticText(panel, label='Max Range:')
        self.max_range_text = wx.TextCtrl(panel, value="0.00")
        self.max_range_text.Bind(wx.EVT_TEXT, self.on_max_range_change)

        # Initialize range control updating flag
        self.updating_range_controls = False

        # Initialize range controls from data
        self.update_range_controls_from_data()

        # Add averaging points control
        averaging_points_label = wx.StaticText(panel, label="Averaging Points:")
        self.averaging_points_text = wx.TextCtrl(panel, value="5")
        self.averaging_points_text.Bind(wx.EVT_TEXT, self.on_averaging_points_change)

        # Add Tougaard controls
        self.cross_section_label = wx.StaticText(panel, label='Tougaard1: B,C,D,T0')
        self.cross_section = wx.TextCtrl(panel, value="2866,1643,1,0")
        self.cross_section.Bind(wx.EVT_TEXT, self.on_cross_section_change)


        # Add Tougaard model button
        self.tougaard_fit_btn = wx.Button(panel, label="Create Tougaard\nModel")
        if 'wxMac' in wx.PlatformInfo:
            self.tougaard_fit_btn.SetMinSize((125, 30))
        else:
            self.tougaard_fit_btn.SetMinSize((125, 35))
        self.tougaard_fit_btn.Bind(wx.EVT_BUTTON, self.on_tougaard_model)

        # Add remove last peak button
        remove_peak_button = wx.Button(panel, label="Remove\nLast Area")
        if 'wxMac' in wx.PlatformInfo:
            remove_peak_button.SetMinSize((90, 30))
        else:
            remove_peak_button.SetMinSize((90, 35))
        remove_peak_button.Bind(wx.EVT_BUTTON, self.on_remove_peak)

        clear_background_button = wx.Button(panel, label="Clear\nAll")
        if 'wxMac' in wx.PlatformInfo:
            clear_background_button.SetMinSize((125, 30))
        else:
            clear_background_button.SetMinSize((110, 35))
        clear_background_button.Bind(wx.EVT_BUTTON, self.on_clear_background)

        # export_button = wx.Button(panel, label="Export to\nResults Grid")
        # export_button.SetMinSize((60, 40))
        # export_button.Bind(wx.EVT_BUTTON, self.on_export_results)


        reset_vlines_button = wx.Button(panel, label="Reset \nVertical Lines")
        if 'wxMac' in wx.PlatformInfo:
            reset_vlines_button.SetMinSize((125, 30))
        else:
            reset_vlines_button.SetMinSize((125, 35))
        reset_vlines_button.Bind(wx.EVT_BUTTON, self.on_reset_vlines)


        background_only_button = wx.Button(panel, label="Create\nBackground / Area")
        if 'wxMac' in wx.PlatformInfo:
            background_only_button.SetMinSize((125, 30))
        else:
            background_only_button.SetMinSize((125, 35))
        background_only_button.Bind(wx.EVT_BUTTON, self.on_background_only)

        area_button = wx.Button(panel, label="Calculate\nArea")
        if 'wxMac' in wx.PlatformInfo:
            area_button.SetMinSize((90, 30))
        else:
            area_button.SetMinSize((90, 35))
        area_button.Bind(wx.EVT_BUTTON, self.on_area)

        peak_label_text_label = wx.StaticText(panel, label="Area Name      ")
        self.peak_label_text = wx.TextCtrl(panel, value="")



        # Layout with a GridBagSizer
        if 'wxMac' in wx.PlatformInfo or 'wxGTK' in wx.PlatformInfo:
            sizer = wx.GridBagSizer(hgap=1, vgap=1)
        else:
            sizer = wx.GridBagSizer(hgap=0, vgap=0)

        if 'wxMac' in wx.PlatformInfo or 'wxGTK' in wx.PlatformInfo:
            # First row: Method
            sizer.Add(method_label, pos=(0, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.method_combobox, pos=(0, 1), flag=wx.ALL | wx.EXPAND, border=1)

            # second Third row: Offset (H) and Offset (L)
            sizer.Add(offset_h_label, pos=(1, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.offset_h_text, pos=(1, 1), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(offset_l_label, pos=(2, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.offset_l_text, pos=(2, 1), flag=wx.ALL | wx.EXPAND, border=1)

            # Add new range controls
            sizer.Add(self.min_range_label, pos=(3, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.min_range_text, pos=(3, 1), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.max_range_label, pos=(4, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.max_range_text, pos=(4, 1), flag=wx.ALL | wx.EXPAND, border=1)

            # Fourth row: Averaging Points
            sizer.Add(averaging_points_label, pos=(5, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.averaging_points_text, pos=(5, 1), flag=wx.ALL | wx.EXPAND, border=1)

            # Tougaard parameters (only Tougaard1)
            sizer.Add(self.cross_section_label, pos=(7, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.cross_section, pos=(7, 1), flag=wx.ALL | wx.EXPAND, border=1)

            # Area row
            area_box = wx.StaticBox(panel, label="Area Calculation")
            area_sizer = wx.StaticBoxSizer(area_box, wx.VERTICAL)


            text_sizer = wx.BoxSizer(wx.HORIZONTAL)
            text_sizer.Add(peak_label_text_label, 0, flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=5)
            # 5) # | wx.ALIGN_CENTER_VERTICAL, 5)
            text_sizer.Add(self.peak_label_text, 1, flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=5)
            # 5) #| wx.EXPAND, 5)

            button_sizer = wx.BoxSizer(wx.HORIZONTAL)
            button_sizer.Add(area_button, 1, flag=wx.ALL | wx.EXPAND, border=1)
            button_sizer.Add(remove_peak_button, 1, flag=wx.ALL | wx.EXPAND, border=1)

            area_sizer.Add(text_sizer, 0, wx.EXPAND)
            area_sizer.Add(button_sizer, 1, wx.EXPAND)

            sizer.Add(area_sizer, pos=(9, 0), span=(2, 2), flag=wx.ALL | wx.EXPAND, border=1)

            # sizer.Add(area_button, pos=(8, 0), flag=wx.ALL | wx.EXPAND, border=5)
            # sizer.Add(remove_peak_button, pos=(8, 1), flag=wx.ALL | wx.EXPAND, border=5)

            # Seventh row: Remove peak and Export buttons
            sizer.Add(self.tougaard_fit_btn, pos=(11, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(reset_vlines_button, pos=(11, 1), flag=wx.ALL | wx.EXPAND, border=1)
            # sizer.Add(export_button, pos=(10, 1), flag=wx.ALL | wx.EXPAND, border=5)

            # Sixth row: Background and Clear Background buttons
            # sizer.Add(background_button, pos=(12, 0), flag=wx.ALL | wx.EXPAND, border=5)
            sizer.Add(background_only_button, pos=(12, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(clear_background_button, pos=(12, 1), flag=wx.ALL | wx.EXPAND, border=1)
        else:  # For Windows
            # First row: Method
            sizer.Add(method_label, pos=(0, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.method_combobox, pos=(0, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

            # second Third row: Offset (H) and Offset (L)
            sizer.Add(offset_h_label, pos=(1, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.offset_h_text, pos=(1, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(offset_l_label, pos=(2, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.offset_l_text, pos=(2, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

            # Add new range controls
            sizer.Add(self.min_range_label, pos=(3, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.min_range_text, pos=(3, 1), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.max_range_label, pos=(4, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.max_range_text, pos=(4, 1), flag=wx.ALL | wx.EXPAND, border=1)

            # Fourth row: Averaging Points
            sizer.Add(averaging_points_label, pos=(5, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.averaging_points_text, pos=(5, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

            # Fourth row: Tougaard parameters
            sizer.Add(self.cross_section_label, pos=(7, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.cross_section, pos=(7, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)


            # Area row
            area_box = wx.StaticBox(panel, label="Area Calculation")
            area_sizer = wx.StaticBoxSizer(area_box, wx.VERTICAL)


            text_sizer = wx.BoxSizer(wx.HORIZONTAL)
            text_sizer.Add(peak_label_text_label, 0, flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=5)
            text_sizer.Add(self.peak_label_text, 1, flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=5)


            button_sizer = wx.BoxSizer(wx.HORIZONTAL)
            button_sizer.Add(area_button, 1, flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            button_sizer.Add(remove_peak_button, 1, flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

            area_sizer.Add(text_sizer, 0, wx.EXPAND)
            area_sizer.Add(button_sizer, 1, wx.EXPAND)

            sizer.Add(area_sizer, pos=(9, 0), span=(2, 2), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)


            # Seventh row: Remove peak and Export buttons
            sizer.Add(self.tougaard_fit_btn, pos=(11, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(reset_vlines_button, pos=(11, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            # sizer.Add(export_button, pos=(10, 1), flag=wx.ALL | wx.EXPAND, border=5)

            # Sixth row: Background and Clear Background buttons
            # sizer.Add(background_button, pos=(12, 0), flag=wx.ALL | wx.EXPAND, border=5)
            sizer.Add(background_only_button, pos=(12, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(clear_background_button, pos=(12, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)


        # Initially disable all Tougaard controls
        self.cross_section.Enable(False)
        self.cross_section_label.Enable(False)
        self.tougaard_fit_btn.Enable(False)

        self.Bind(wx.EVT_CLOSE, self.on_close)

        self.method_combobox.Bind(wx.EVT_COMBOBOX, self.on_bkg_method_change)

        from libraries.ConfigFile import set_consistent_fonts
        set_consistent_fonts(self)

        panel.SetSizer(sizer)

    def update_tougaard_controls_visibility(self, new_method):
        if "Tougaard" in new_method:
            self.cross_section.Enable(True)
            self.cross_section_label.Enable(True)
            self.tougaard_fit_btn.Enable(True)
        else:
            self.cross_section.Enable(False)
            self.cross_section_label.Enable(False)
            self.tougaard_fit_btn.Enable(False)

    def on_area(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        if self.parent.vline1 is None or self.parent.vline2 is None:
            return

        x_values = np.array(self.parent.Data['Core levels'][sheet_name]['B.E.'])
        y_values = np.array(self.parent.Data['Core levels'][sheet_name]['Raw Data'])
        background = np.array(self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Y'])

        vline1_x = self.parent.vline1.get_xdata()[0]
        vline2_x = self.parent.vline2.get_xdata()[0]
        range_min = round(min(vline1_x, vline2_x),2)
        range_max = round(max(vline1_x, vline2_x),2)
        mask = (x_values >= range_min) & (x_values <= range_max)

        x_range = x_values[mask]
        y_range = y_values[mask]
        bg_range = background[mask]
        y_minus_bg = y_range - bg_range

        sorted_indices = np.argsort(x_range)
        x_sorted = x_range[sorted_indices]
        y_minus_bg_sorted = y_minus_bg[sorted_indices]

        area = np.trapz(y_minus_bg_sorted, x_sorted)
        
        # area = np.trapz(y_minus_bg, x_range)
        print(f"Area positive: {area}")
        # print(f"Area negative: {np.trapz(y_minus_bg, x_range)}")
        peak_index = np.argmax(y_minus_bg)
        peak_position = x_range[peak_index]
        peak_height = y_minus_bg[peak_index]
        fwhm = 2 * np.sqrt(2 * np.log(2)) * area / (peak_height * np.sqrt(2 * np.pi))

        area = abs(round(area, 2))
        peak_position = round(peak_position, 2)
        peak_height = round(peak_height, 2)
        fwhm = round(fwhm, 2)

        grid = self.parent.peak_params_grid
        num_peaks = grid.GetNumberRows() // 2
        peak_letter = chr(65 + num_peaks)
        # peak_name = self.peak_label_text.GetValue() f"p{num_peaks + 1}" if self.peak_label_text.GetValue() else f"{sheet_name} p{num_peaks + 1}"
        peak_name = f"{self.peak_label_text.GetValue()} p{num_peaks + 1}" if self.peak_label_text.GetValue() else f"{sheet_name} p{num_peaks + 1}"

        grid.AppendRows(2)
        row = num_peaks * 2

        grid.SetCellValue(row, 0, peak_letter)
        grid.SetCellValue(row, 1, peak_name)
        grid.SetCellValue(row, 2, f"{peak_position}")
        grid.SetCellValue(row, 3, f"{peak_height}")
        grid.SetCellValue(row, 4, f"{fwhm}")
        grid.SetCellValue(row, 5, "0.00")
        grid.SetCellValue(row, 6, f"{area}")
        grid.SetCellValue(row, 7, "0.00")
        grid.SetCellValue(row, 8, "0.00")
        grid.SetCellValue(row, 9, "0.00")
        grid.SetCellValue(row, 13, "Unfitted")
        grid.SetCellValue(row, 15, f"{range_min}")
        grid.SetCellValue(row, 16, f"{range_max}")

        for col in range(grid.GetNumberCols()):
            grid.SetCellBackgroundColour(row + 1, col, wx.Colour(200,245,228))

        grid.SetCellValue(row + 1, 2, "0,1e3")
        grid.SetCellValue(row + 1, 3, "1,1e7")
        grid.SetCellValue(row + 1, 4, "0.3,3.5")
        grid.SetCellValue(row + 1, 5, "0,0.5")
        grid.SetCellValue(row + 1, 7, "0.1,1")
        grid.SetCellValue(row + 1, 8, "0.1,1")
        grid.SetCellValue(row + 1, 9, "0.01,2")

        if 'Fitting' not in self.parent.Data['Core levels'][sheet_name]:
            self.parent.Data['Core levels'][sheet_name]['Fitting'] = {}
        if 'Peaks' not in self.parent.Data['Core levels'][sheet_name]['Fitting']:
            self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = {}

        self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'][peak_name] = {
            'Position': peak_position,
            'Height': peak_height,
            'FWHM': fwhm,
            'L/G': 0.00,
            'Area': area,
            'Sigma': 0,
            'Gamma': 0,
            'Skew': 0,
            'Fitting Model': "Unfitted",
            'Bkg Low': range_min,
            'Bkg High': range_max,
            'Constraints': {
                'Position': "0,1e3",
                'Height': "1,1e7",
                'FWHM': "0.3,3.5",
                'L/G': "0,0.5",
                'Sigma': "0.1,1",
                'Gamma': "0.1,1",
                'Skew': "0.01,2"
            }
        }

        self.parent.ax.fill_between(x_range, bg_range, y_range,
                                    facecolor='lightgreen', alpha=0.5,
                                    label=peak_name)

        self.parent.ax.legend()
        self.parent.peak_params_grid.ForceRefresh()
        self.on_reset_vlines(self)
        save_state(self.parent)

    def on_reset_vlines(self, event):
        # Calculate 1/15 and 14/15 positions
        if hasattr(self.parent, 'x_values') and len(self.parent.x_values) > 0:
            x_min = min(self.parent.x_values)
            x_max = max(self.parent.x_values)
            x_range = x_max - x_min

            new_low = x_min + x_range / 15
            new_high = x_min + 14 * x_range / 15

            # Update data structure
            sheet_name = self.parent.sheet_combobox.GetValue()
            if sheet_name in self.parent.Data['Core levels']:
                if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
                    self.parent.Data['Core levels'][sheet_name]['Background'] = {}
                self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Low'] = float(new_low)
                self.parent.Data['Core levels'][sheet_name]['Background']['Bkg High'] = float(new_high)

            self.parent.bg_min_energy = new_low
            self.parent.bg_max_energy = new_high

        # Reset vlines and recreate them
        self.parent.vline1 = None
        self.parent.vline2 = None
        if hasattr(self.parent, 'vline1_text'):
            self.parent.vline1_text = None
        if hasattr(self.parent, 'vline2_text'):
            self.parent.vline2_text = None

        # Recreate vlines
        self.initialize_or_restore_area_vlines()
        self.parent.canvas.draw_idle()

        # Update range controls
        self.update_range_controls_from_data()

    def on_toggle_vlines(self, event):
        if self.parent.vline1:
            self.parent.vline1.set_visible(not self.parent.vline1.get_visible())
        if self.parent.vline2:
            self.parent.vline2.set_visible(not self.parent.vline2.get_visible())
        self.parent.canvas.draw_idle()


    def on_background_only(self, event):
        """Create background and calculate area - combines both functionalities."""
        sheet_name = self.parent.sheet_combobox.GetValue()
        if self.parent.vline1 is None or self.parent.vline2 is None:
            return

        save_state(self.parent)

        # Set the background method from combobox
        selected_method = self.method_combobox.GetValue()
        self.parent.background_method = selected_method

        # Get offsets
        self.parent.offset_h = float(self.offset_h_text.GetValue())
        self.parent.offset_l = float(self.offset_l_text.GetValue())

        # Store vline positions BEFORE plotting background (they will get destroyed)
        vline1_x = self.parent.vline1.get_xdata()[0]
        vline2_x = self.parent.vline2.get_xdata()[0]

        # Calculate background first using plot_manager
        self.parent.plot_manager.plot_background(self.parent)

        # Get data after background calculation
        x_values = np.array(self.parent.Data['Core levels'][sheet_name]['B.E.'])
        y_values = np.array(self.parent.Data['Core levels'][sheet_name]['Raw Data'])
        background = np.array(self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Y'])

        # Use stored vline positions (since the vlines were destroyed by plot_background)
        range_min = round(min(vline1_x, vline2_x), 2)
        range_max = round(max(vline1_x, vline2_x), 2)
        mask = (x_values >= range_min) & (x_values <= range_max)

        # Get data in range
        x_range = x_values[mask]
        y_range = y_values[mask]
        bg_range = background[mask]
        y_minus_bg = y_range - bg_range

        # Sort data for proper integration
        sorted_indices = np.argsort(x_range)
        x_sorted = x_range[sorted_indices]
        y_minus_bg_sorted = y_minus_bg[sorted_indices]

        # Calculate area and peak parameters
        area = np.trapz(y_minus_bg_sorted, x_sorted)
        print(f"Area: {area}")

        peak_index = np.argmax(y_minus_bg)
        peak_position = x_range[peak_index]
        peak_height = y_minus_bg[peak_index]

        if peak_height > 0:
            fwhm = 2 * np.sqrt(2 * np.log(2)) * abs(area) / (peak_height * np.sqrt(2 * np.pi))
        else:
            fwhm = 0

        area = abs(round(area, 2))
        peak_position = round(peak_position, 2)
        peak_height = round(peak_height, 2)
        fwhm = round(fwhm, 2)

        # Find next available peak letter
        grid = self.parent.peak_params_grid
        num_peaks = grid.GetNumberRows() // 2
        # peak_letter = chr(65 + num_peaks)

        # Check if this is Survey data
        sheet_name_lower = sheet_name.lower()
        is_survey = any(x in sheet_name_lower for x in ['survey', 'wide'])

        # Generate area name from text field or auto-detect
        area_name = self.peak_label_text.GetValue().strip()
        if not area_name:
            # Auto-detect from sheet name
            if is_survey:
                # For Survey: use format "C1s ." instead of "C1s p1"
                base_name = sheet_name.replace('Survey', '').replace('Wide', '').strip()
                if not base_name:
                    # If no base name, try to detect from vline positions
                    center_position = (vline1_x + vline2_x) / 2
                    # You could add element detection logic here based on BE position
                    area_name = f"Unknown ."
                else:
                    area_name = f"{base_name} ."
            else:
                # For regular core levels: use existing format
                num_peaks = grid.GetNumberRows() // 2
                area_name = f"{sheet_name} p{num_peaks + 1}"
        else:
            # Use the area name from text field
            if is_survey and not area_name.endswith(' .'):
                area_name = f"{area_name} ."

        # Check if peak with this name already exists and overwrite
        grid = self.parent.peak_params_grid
        existing_row = -1
        for row in range(0, grid.GetNumberRows(), 2):
            if grid.GetCellValue(row, 1) == area_name:
                existing_row = row
                break

        if existing_row >= 0:
            # Overwrite existing peak
            row = existing_row
            peak_letter = grid.GetCellValue(row, 0)
        else:
            # Create new peak
            num_peaks = grid.GetNumberRows() // 2
            peak_letter = chr(65 + num_peaks)

            # Add new rows for the peak
            grid.AppendRows(2)
            row = num_peaks * 2

        # Set all the peak values
        grid.SetCellValue(row, 0, peak_letter)
        grid.SetCellValue(row, 1, area_name)  # Use area_name instead of peak_name
        grid.SetCellValue(row, 2, f"{peak_position:.2f}")
        grid.SetCellValue(row, 3, f"{peak_height:.2f}")
        grid.SetCellValue(row, 4, f"{fwhm:.2f}")
        grid.SetCellValue(row, 5, "0")
        grid.SetCellValue(row, 6, f"{area:.2f}")
        grid.SetCellValue(row, 7, "")
        grid.SetCellValue(row, 8, "")
        grid.SetCellValue(row, 9, "")
        grid.SetCellValue(row, 13, "Unfitted")
        grid.SetCellValue(row, 14, selected_method)  # Background Type
        grid.SetCellValue(row, 15, f"{range_min}")  # Bkg Low
        grid.SetCellValue(row, 16, f"{range_max}")  # Bkg High

        # Set constraints row
        grid.SetCellValue(row + 1, 2, "Fixed")
        grid.SetCellValue(row + 1, 3, "Fixed")
        grid.SetCellValue(row + 1, 4, "Fixed")

        # Set constraints
        for col in range(grid.GetNumberCols()):
            grid.SetCellBackgroundColour(row + 1, col, wx.Colour(200, 245, 228))

        grid.SetCellValue(row + 1, 6, "Fixed")
        grid.SetCellValue(row + 1, 7, "")
        grid.SetCellValue(row + 1, 8, "")
        grid.SetCellValue(row + 1, 9, "")

        # Update Data structure
        if 'Fitting' not in self.parent.Data['Core levels'][sheet_name]:
            self.parent.Data['Core levels'][sheet_name]['Fitting'] = {}
        if 'Peaks' not in self.parent.Data['Core levels'][sheet_name]['Fitting']:
            self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = {}

        self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'][area_name] = {
            'Position': peak_position,
            'Height': peak_height,
            'FWHM': fwhm,
            'L/G': 0.00,
            'Area': area,
            'Sigma': 0,
            'Gamma': 0,
            'Skew': 0,
            'Fitting Model': "Unfitted",
            'Bkg Type': selected_method,
            'Bkg Low': range_min,
            'Bkg High': range_max,
            'Constraints': {
                'Position': "Fixed",
                'Height': "Fixed",
                'FWHM': "Fixed",
                'L/G': "Fixed",
                'Area': "Fixed",
                'Sigma': "Fixed",
                'Gamma': "Fixed",
                'Skew': "Fixed"
            }
        }


        # Fill area between curves
        self.parent.ax.fill_between(x_range, bg_range, y_range,
                                    facecolor='lightgreen', alpha=0.5,
                                    label=area_name)



        # RESTORE VLINES AND MOUSE INTERACTION - THIS IS THE KEY FIX!
        # The plot_background method calls clear_and_replot which removes vlines AND disconnects mouse handlers
        # So we need to recreate them AND restore mouse interaction
        if hasattr(self, 'initialize_or_restore_area_vlines'):
            self.initialize_or_restore_area_vlines()

        # CRITICAL: Reset mouse interaction system to make vlines draggable again
        if hasattr(self.parent, 'mouse_handler'):
            # Reset any existing mouse handlers
            self.parent.mouse_handler.cleanup_vline_handlers()
            # Reset moving vline state
            self.parent.moving_vline = None

        # Update range controls to reflect current vline positions
        if hasattr(self, 'update_range_controls_from_data'):
            self.update_range_controls_from_data()

        # Use the same plotting sequence as sheet change for proper display
        # Apply choice editors to the fitting model column
        self.parent.set_model_choice_editors(self.parent)

        self.parent.plot_manager.plot_data(self.parent)  # Always plot raw data first
        if self.parent.show_fit and self.parent.peak_params_grid.GetNumberRows() > 0:
            self.parent.clear_and_replot()  # Add fit and residuals if show_fit is True

        # self.parent.disable_area_interaction()
        # self.parent.enable_area_interaction()


        self.parent.plot_config.update_plot_limits(self.parent, sheet_name)
        self.parent.plot_manager.update_legend(self.parent)

        # Restore vlines at their original positions (they were destroyed by clear_and_replot)
        if hasattr(self.parent, 'vline1') and hasattr(self.parent, 'vline2'):
            # Recreate vlines at stored positions (vline1_x and vline2_x were stored earlier)
            self.parent.vline1 = self.parent.ax.axvline(vline1_x, color='r', linestyle='--', alpha=0.7)
            self.parent.vline2 = self.parent.ax.axvline(vline2_x, color='r', linestyle='--', alpha=0.7)

            # Add text labels back
            self.add_vline_text_labels()

            # Update area tab selection state so vlines stay visible
            self.parent.area_tab_selected = True

        # Print results to console
        print(f"Results for {sheet_name}:")
        print(f"Peak Name: {area_name}")
        print(f"Peak Position: {peak_position:.2f} eV")
        print(f"Peak Height: {peak_height:.2f} counts")
        print(f"FWHM: {fwhm:.2f} eV")
        print(f"Area: {area:.2f}")
        print(f"Background Type: {selected_method}")
        print(f"Range: {range_min:.2f} - {range_max:.2f} eV")


    def on_bkg_method_change(self, event):
        new_method = self.method_combobox.GetValue()
        self.parent.background_method = new_method
        self.update_tougaard_controls_visibility(new_method)

    def on_cross_section_change(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        values = self.cross_section.GetValue().split(',')
        try:
            self.parent.Data['Core levels'][sheet_name]['Background'].update({
                'Tougaard_B': float(values[0]),
                'Tougaard_C': float(values[1]),
                'Tougaard_D': float(values[2]),
                'Tougaard_T0': float(values[3])
            })
        except (ValueError, IndexError):
            pass


    def on_remove_peak(self, event):
        remove_peak(self.parent)
        save_state(self.parent)


    def on_averaging_points_change(self, event):
        try:
            self.parent.averaging_points = int(self.averaging_points_text.GetValue())
        except ValueError:
            pass

    def on_tougaard_model(self, event):
        from libraries.ToolsMenu.Fitting_Screen import TougaardFitWindow
        TougaardFitWindow(self).Show()

    def on_close(self, event):
        """Close area screen and ensure complete cleanup."""
        # Use the proper disable method
        self.parent.disable_area_interaction()

        # Save state and destroy window
        save_state(self.parent)
        self.Destroy()



    def on_background(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        if self.parent.vline1 is None or self.parent.vline2 is None:
            return

        # Set the background method from combobox
        selected_method = self.method_combobox.GetValue()
        self.parent.background_method = selected_method

        # Get offsets
        self.parent.offset_h = float(self.offset_h_text.GetValue())
        self.parent.offset_l = float(self.offset_l_text.GetValue())

        # Calculate background first using plot_manager
        self.parent.plot_manager.plot_background(self.parent)

        # Get data after background calculation
        x_values = np.array(self.parent.Data['Core levels'][sheet_name]['B.E.'])
        y_values = np.array(self.parent.Data['Core levels'][sheet_name]['Raw Data'])
        background = np.array(self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Y'])

        # Get vline positions and mask data
        vline1_x = self.parent.vline1.get_xdata()[0]
        vline2_x = self.parent.vline2.get_xdata()[0]
        range_min = min(vline1_x, vline2_x)
        range_max = max(vline1_x, vline2_x)
        mask = (x_values >= range_min) & (x_values <= range_max)

        # Get data in range
        x_range = x_values[mask]
        y_range = y_values[mask]
        bg_range = background[mask]

        # Calculate area and peak parameters
        y_minus_bg = y_range - bg_range
        area = abs(np.trapz(y_minus_bg, x_range))
        print(f"Area positive: {area}")
        print(f"Area negative: {np.trapz(y_minus_bg, x_range)}")
        peak_index = np.argmax(y_minus_bg)
        peak_position = x_range[peak_index]
        peak_height = y_minus_bg[peak_index]

        if peak_height > 0:
            fwhm = 2 * np.sqrt(2 * np.log(2)) * area / (peak_height * np.sqrt(2 * np.pi))
        else:
            fwhm = 0

        area = abs(round(area, 2))
        peak_position = round(peak_position, 2)
        peak_height = round(peak_height, 2)
        fwhm = round(fwhm, 2)

        # Find next available peak letter
        grid = self.parent.peak_params_grid
        num_peaks = grid.GetNumberRows() // 2
        peak_letter = chr(65 + num_peaks)
        peak_name = f"{sheet_name} p{num_peaks + 1}"

        # Add new rows for the peak
        grid.AppendRows(2)
        row = num_peaks * 2

        grid.SetCellValue(row, 0, peak_letter)
        grid.SetCellValue(row, 1, peak_name)
        grid.SetCellValue(row, 2, f"{peak_position}")
        grid.SetCellValue(row, 3, f"{peak_height}")
        grid.SetCellValue(row, 4, f"{fwhm}")
        grid.SetCellValue(row, 5, "0.00")
        grid.SetCellValue(row, 6, f"{area}")
        grid.SetCellValue(row, 7, "0.00")
        grid.SetCellValue(row, 8, "0.00")
        grid.SetCellValue(row, 9, "0.00")
        grid.SetCellValue(row, 13, "Unfitted")

        # Set constraints
        for col in range(grid.GetNumberCols()):
            grid.SetCellBackgroundColour(row + 1, col, wx.Colour(230, 230, 230))


        grid.SetCellValue(row + 1, 2, "0,1e3")
        grid.SetCellValue(row + 1, 3, "1,1e7")
        grid.SetCellValue(row + 1, 4, "0.3,3.5")
        grid.SetCellValue(row + 1, 5, "0,0.5")
        grid.SetCellValue(row + 1, 7, "0.1,1")
        grid.SetCellValue(row + 1, 8, "0.1,1")
        grid.SetCellValue(row + 1, 9, "0.01,2")

        # Update Data structure
        if 'Fitting' not in self.parent.Data['Core levels'][sheet_name]:
            self.parent.Data['Core levels'][sheet_name]['Fitting'] = {}
        if 'Peaks' not in self.parent.Data['Core levels'][sheet_name]['Fitting']:
            self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = {}

        self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'][peak_name] = {
            'Position': peak_position,
            'Height': peak_height,
            'FWHM': fwhm,
            'L/G': 0.00,
            'Area': area,
            'Sigma': 0,
            'Gamma': 0,
            'Skew': 0,
            'Fitting Model': "Unfitted",
            'Constraints': {
                'Position': "0,1e3",
                'Height': "1,1e7",
                'FWHM': "0.3,3.5",
                'L/G': "0,0.5",
                'Sigma': "0.1,1",
                'Gamma': "0.1,1",
                'Skew': "0.01,2"
            }
        }

        # Fill area between curves
        self.parent.ax.fill_between(x_range, bg_range, y_range,
                                    facecolor='lightgreen', alpha=0.5,
                                    label=peak_name)

        self.parent.ax.legend()
        self.parent.peak_params_grid.ForceRefresh()
        self.parent.canvas.draw_idle()

        print(f"Results for {sheet_name}:")
        print(f"Peak Name: {peak_name}")
        print(f"Peak Position: {peak_position:.2f} eV")
        print(f"Peak Height: {peak_height:.2f} counts")
        print(f"FWHM: {fwhm:.2f} eV")
        print(f"Area: {area:.2f}")

        save_state(self.parent)


    def on_clear_background(self, event):
        if hasattr(self, 'offset_h_text') and hasattr(self, 'offset_l_text'):
            self.offset_h_text.SetValue('0')
            self.offset_l_text.SetValue('0')
        self.parent.plot_manager.clear_background(self.parent)
        self.parent.plot_data()
        save_state(self.parent)

    def on_export_results(self, event):
        self.parent.export_results()
        save_state(self.parent)

    def on_offset_changed(self, event):
        try:
            offset_h = float(self.offset_h_text.GetValue())
            offset_l = float(self.offset_l_text.GetValue())
            self.parent.set_offset_h(offset_h)
            self.parent.set_offset_l(offset_l)
            save_state(self.parent)
        except ValueError:
            print("Invalid offset value")

    # Add these methods to the BackgroundWindow class:

    def initialize_or_restore_area_vlines(self):
        """Initialize vertical lines - robust version that always works."""
        if not hasattr(self.parent, 'x_values') or len(self.parent.x_values) == 0:
            return

        sheet_name = self.parent.sheet_combobox.GetValue()

        # Only remove existing vlines if they exist (don't force cleanup)
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

        # Get positions from background data or use defaults
        saved_low = None
        saved_high = None
        if (sheet_name in self.parent.Data['Core levels'] and
                'Background' in self.parent.Data['Core levels'][sheet_name]):
            bg_data = self.parent.Data['Core levels'][sheet_name]['Background']
            saved_low = bg_data.get('Bkg Low')
            saved_high = bg_data.get('Bkg High')

        # Use saved positions if valid, otherwise default to 1/15 and 14/15
        if (saved_low is not None and saved_high is not None and
                saved_low != saved_high and saved_low != '' and saved_high != '' and
                str(saved_low).strip() != '' and str(saved_high).strip() != ''):
            try:
                vline1_pos = float(saved_low)
                vline2_pos = float(saved_high)
            except (ValueError, TypeError):
                x_min = min(self.parent.x_values)
                x_max = max(self.parent.x_values)
                x_range = x_max - x_min
                vline1_pos = x_min + x_range / 15
                vline2_pos = x_min + 14 * x_range / 15
        else:
            x_min = min(self.parent.x_values)
            x_max = max(self.parent.x_values)
            x_range = x_max - x_min
            vline1_pos = x_min + x_range / 15
            vline2_pos = x_min + 14 * x_range / 15

        # Create new vlines - these will be draggable
        self.parent.vline1 = self.parent.ax.axvline(vline1_pos, color='r', linestyle='--', alpha=0.7)
        self.parent.vline2 = self.parent.ax.axvline(vline2_pos, color='r', linestyle='--', alpha=0.7)

        # Add text labels
        self.add_vline_text_labels()

        # Update data structure
        self.parent.bg_min_energy = min(vline1_pos, vline2_pos)
        self.parent.bg_max_energy = max(vline1_pos, vline2_pos)

        if sheet_name in self.parent.Data['Core levels']:
            if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
                self.parent.Data['Core levels'][sheet_name]['Background'] = {}
            self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Low'] = float(min(vline1_pos, vline2_pos))
            self.parent.Data['Core levels'][sheet_name]['Background']['Bkg High'] = float(max(vline1_pos, vline2_pos))

        # Auto-detect area name from vline positions
        self.auto_detect_area_name(vline1_pos, vline2_pos)

    def add_vline_text_labels(self):
        """Add text labels to vertical lines showing their BE values."""
        if self.parent.vline1 is not None and self.parent.vline2 is not None:
            # Get y-axis range for positioning
            y_min, y_max = self.parent.ax.get_ylim()
            text_y = y_min + 0.9 * (y_max - y_min)  # 9/10 of Y axis

            # Get vline positions and round to 2 digits
            vline1_x = round(self.parent.vline1.get_xdata()[0], 2)
            vline2_x = round(self.parent.vline2.get_xdata()[0], 2)

            # Remove existing text if any
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

            # Create new text labels with 2 decimal places
            self.parent.vline1_text = self.parent.ax.text(vline1_x, text_y, f'{vline1_x:.2f}',
                                                          ha='center', va='center',
                                                          color='grey', fontsize=10,
                                                          bbox=dict(boxstyle='round,pad=0.2', facecolor='white',
                                                                    alpha=0.8))

            self.parent.vline2_text = self.parent.ax.text(vline2_x, text_y, f'{vline2_x:.2f}',
                                                          ha='center', va='center',
                                                          color='grey', fontsize=10,
                                                          bbox=dict(boxstyle='round,pad=0.2', facecolor='white',
                                                                    alpha=0.8))

    def update_vline_text_labels(self):
        """Update the text labels when vlines are moved."""
        if (self.parent.vline1 is not None and self.parent.vline2 is not None and
                hasattr(self.parent, 'vline1_text') and self.parent.vline1_text is not None and
                hasattr(self.parent, 'vline2_text') and self.parent.vline2_text is not None):
            # Get y-axis range for positioning
            y_min, y_max = self.parent.ax.get_ylim()
            text_y = y_min + 0.9 * (y_max - y_min)  # 9/10 of Y axis

            # Get current vline positions and round to 2 digits
            vline1_x = round(self.parent.vline1.get_xdata()[0], 2)
            vline2_x = round(self.parent.vline2.get_xdata()[0], 2)

            # Update text positions and values
            self.parent.vline1_text.set_position((vline1_x, text_y))
            self.parent.vline1_text.set_text(f'{vline1_x:.2f}')

            self.parent.vline2_text.set_position((vline2_x, text_y))
            self.parent.vline2_text.set_text(f'{vline2_x:.2f}')

            # Auto-detect area name when vlines move
            self.auto_detect_area_name(vline1_x, vline2_x)

    def show_hide_vlines(self):
        """Show/hide vlines based on current screen state and zoom/drag modes."""
        # Hide vlines if zooming or dragging
        if self.zoom_mode or self.drag_mode:
            if self.vline1 is not None:
                self.vline1.set_visible(False)
            if self.vline2 is not None:
                self.vline2.set_visible(False)
            if hasattr(self, 'vline1_text') and self.vline1_text is not None:
                self.vline1_text.set_visible(False)
            if hasattr(self, 'vline2_text') and self.vline2_text is not None:
                self.vline2_text.set_visible(False)
            if self.vline3 is not None:
                self.vline3.set_visible(False)
            if self.vline4 is not None:
                self.vline4.set_visible(False)
            return

        # Check if either fitting screen OR area screen is open and active
        fitting_screen_active = (hasattr(self, 'fitting_window') and
                                 self.fitting_window is not None and
                                 hasattr(self, 'background_tab_selected') and
                                 self.background_tab_selected)

        area_screen_active = (hasattr(self, 'background_window') and
                              self.background_window is not None and
                              hasattr(self, 'area_tab_selected') and
                              self.area_tab_selected)

        # Background lines (vline1, vline2) should be visible if EITHER screen is active
        background_lines_visible = fitting_screen_active or area_screen_active

        # Set visibility for background vlines (vline1 and vline2)
        if self.vline1 is not None:
            self.vline1.set_visible(background_lines_visible)
        if self.vline2 is not None:
            self.vline2.set_visible(background_lines_visible)
        if hasattr(self, 'vline1_text') and self.vline1_text is not None:
            self.vline1_text.set_visible(background_lines_visible)
        if hasattr(self, 'vline2_text') and self.vline2_text is not None:
            self.vline2_text.set_visible(background_lines_visible)

        # Noise lines visibility (only for fitting screen)
        noise_lines_visible = (self.noise_analysis_window is not None and
                               hasattr(self, 'noise_tab_selected') and
                               self.noise_tab_selected)
        if self.vline3 is not None:
            self.vline3.set_visible(noise_lines_visible)
        if self.vline4 is not None:
            self.vline4.set_visible(noise_lines_visible)

        self.canvas.draw_idle()

    def create_elements_database(self):
        """Create a comprehensive elements database with binding energy ranges."""
        return {
            'C': {'1s': (284, 291), '2s': (20, 25)},
            'N': {'1s': (397, 407), '2s': (20, 25)},
            'O': {'1s': (528, 538), '2s': (23, 28), '2p': (5, 8)},
            'F': {'1s': (684, 688), '2s': (31, 35)},
            'Na': {'1s': (1071, 1072), '2s': (63, 64), '2p': (31, 31.5)},
            'Mg': {'1s': (1303, 1304), '2s': (88, 90), '2p': (50, 51)},
            'Al': {'1s': (1559, 1560), '2s': (117, 119), '2p': (72, 75)},
            'Si': {'1s': (1838, 1840), '2s': (149, 151), '2p': (99, 103)},
            'P': {'1s': (2145, 2147), '2s': (189, 191), '2p': (129, 136)},
            'S': {'1s': (2470, 2473), '2s': (228, 230), '2p': (162, 170)},
            'Cl': {'1s': (2822, 2824), '2s': (270, 272), '2p': (198, 202)},
            'K': {'1s': (3607, 3609), '2s': (378, 380), '2p': (293, 297), '3s': (34, 35), '3p': (18, 19)},
            'Ca': {'1s': (4038, 4040), '2s': (438, 440), '2p': (346, 350), '3s': (44, 45), '3p': (25, 26)},
            'Ti': {'1s': (4964, 4966), '2s': (563, 565), '2p': (455, 465), '3s': (60, 61), '3p': (37, 38)},
            'Cr': {'1s': (5987, 5989), '2s': (694, 696), '2p': (574, 584), '3s': (84, 85), '3p': (43, 45)},
            'Mn': {'1s': (6539, 6541), '2s': (769, 771), '2p': (639, 651), '3s': (82, 84), '3p': (47, 49)},
            'Fe': {'1s': (7112, 7114), '2s': (844, 846), '2p': (706, 720), '3s': (91, 93), '3p': (52, 54)},
            'Co': {'1s': (7709, 7711), '2s': (925, 927), '2p': (778, 793), '3s': (101, 103), '3p': (58, 60)},
            'Ni': {'1s': (8333, 8335), '2s': (1008, 1010), '2p': (852, 870), '3s': (110, 112), '3p': (66, 68)},
            'Cu': {'1s': (8979, 8981), '2s': (1096, 1098), '2p': (932, 953), '3s': (122, 124), '3p': (75, 77)},
            'Zn': {'1s': (9659, 9661), '2s': (1193, 1195), '2p': (1021, 1045), '3s': (139, 141), '3p': (88, 90)},
            'Ga': {'1s': (10367, 10369), '2s': (1298, 1300), '2p': (1116, 1120), '3s': (159, 161), '3p': (103, 105)},
            'Ge': {'1s': (11103, 11105), '2s': (1414, 1416), '2p': (1217, 1221), '3s': (180, 182), '3p': (120, 122)},
            'As': {'1s': (11867, 11869), '2s': (1527, 1529), '2p': (1323, 1327), '3s': (204, 206), '3p': (141, 143)},
            'Se': {'1s': (12658, 12660), '2s': (1652, 1654), '2p': (1433, 1437), '3s': (229, 231), '3p': (160, 162)},
            'Br': {'1s': (13474, 13476), '2s': (1782, 1784), '2p': (1550, 1554), '3s': (257, 259), '3p': (182, 184)},
            'Sr': {'1s': (16105, 16107), '2s': (2216, 2218), '2p': (1940, 1944), '3s': (358, 360), '3p': (269, 271)},
            'Zr': {'1s': (17998, 18000), '2s': (2532, 2534), '2p': (2223, 2227), '3s': (430, 432), '3p': (343, 345)},
            'Mo': {'1s': (20000, 20002), '2s': (2866, 2868), '2p': (2520, 2524), '3s': (506, 508), '3p': (412, 414)},
            'Ag': {'1s': (25514, 25516), '2s': (3806, 3808), '2p': (3351, 3355), '3s': (719, 721), '3p': (603, 605)},
            'Cd': {'1s': (26711, 26713), '2s': (4018, 4020), '2p': (3538, 3542), '3s': (772, 774), '3p': (652, 654)},
            'In': {'1s': (27940, 27942), '2s': (4238, 4240), '2p': (3730, 3734), '3s': (827, 829), '3p': (703, 705)},
            'Sn': {'1s': (29200, 29202), '2s': (4465, 4467), '2p': (3929, 3933), '3s': (884, 886), '3p': (756, 758)},
            'Sb': {'1s': (30491, 30493), '2s': (4698, 4700), '2p': (4132, 4136), '3s': (946, 948), '3p': (812, 814)},
            'Te': {'1s': (31814, 31816), '2s': (4939, 4941), '2p': (4341, 4345), '3s': (1006, 1008), '3p': (870, 872)},
            'I': {'1s': (33169, 33171), '2s': (5188, 5190), '2p': (4557, 4561), '3s': (1072, 1074), '3p': (931, 933)},
            'Ba': {'1s': (37441, 37443), '2s': (5989, 5991), '2p': (5247, 5251), '3s': (1293, 1295),
                   '3p': (1137, 1139)},
            'Au': {'1s': (80725, 80727), '2s': (14353, 14355), '2p': (11919, 11923), '3s': (3148, 3150),
                   '3p': (2743, 2745)},
            'Pb': {'1s': (88005, 88007), '2s': (15861, 15863), '2p': (13035, 13039), '3s': (3851, 3853),
                   '3p': (3554, 3556)},
            'Bi': {'1s': (90526, 90528), '2s': (16388, 16390), '2p': (13419, 13423), '3s': (3999, 4001),
                   '3p': (3696, 3698)}
        }

    def update_range_controls_from_data(self):
        """Update min/max range controls from vline positions."""
        if not hasattr(self, 'min_range_text') or not hasattr(self, 'max_range_text'):
            return
        try:
            self.min_range_text.GetValue()
            self.max_range_text.GetValue()
        except RuntimeError:
            return

        self.updating_range_controls = True

        try:
            if (self.parent.vline1 is not None and self.parent.vline2 is not None):
                vline1_pos = self.parent.vline1.get_xdata()[0]
                vline2_pos = self.parent.vline2.get_xdata()[0]
                actual_min = min(vline1_pos, vline2_pos)
                actual_max = max(vline1_pos, vline2_pos)
            else:
                actual_min = 0
                actual_max = 0

            self.min_range_text.SetValue(f"{float(actual_min):.2f}")
            self.max_range_text.SetValue(f"{float(actual_max):.2f}")

        except (ValueError, TypeError, RuntimeError):
            try:
                self.min_range_text.SetValue("0.00")
                self.max_range_text.SetValue("0.00")
            except RuntimeError:
                return

        finally:
            self.updating_range_controls = False

    def on_min_range_change(self, event):
        """Handle min range change."""
        if self.updating_range_controls:
            return

        try:
            new_min = float(self.min_range_text.GetValue())
            max_val = float(self.max_range_text.GetValue())

            new_min = round(new_min, 2)
            max_val = round(max_val, 2)

            if new_min > max_val:
                self.updating_range_controls = True
                self.min_range_text.SetValue(f"{max_val:.2f}")
                self.max_range_text.SetValue(f"{new_min:.2f}")
                self.updating_range_controls = False
                new_min = max_val

            if self.parent.vline1 is not None:
                self.parent.vline1.set_xdata([new_min, new_min])
                self.update_vline_text_labels()
                self.parent.canvas.draw_idle()

        except ValueError:
            pass

    def on_max_range_change(self, event):
        """Handle max range change."""
        if self.updating_range_controls:
            return

        try:
            new_max = float(self.max_range_text.GetValue())
            min_val = float(self.min_range_text.GetValue())

            new_max = round(new_max, 2)
            min_val = round(min_val, 2)

            if new_max < min_val:
                self.updating_range_controls = True
                self.max_range_text.SetValue(f"{min_val:.2f}")
                self.min_range_text.SetValue(f"{new_max:.2f}")
                self.updating_range_controls = False
                new_max = min_val

            if self.parent.vline2 is not None:
                self.parent.vline2.set_xdata([new_max, new_max])
                self.update_vline_text_labels()
                self.parent.canvas.draw_idle()

        except ValueError:
            pass

    def auto_detect_area_name(self, vline1_pos, vline2_pos):
        """Auto-detect and update area name based on vline positions."""
        range_min = min(vline1_pos, vline2_pos)
        range_max = max(vline1_pos, vline2_pos)
        range_center = (range_min + range_max) / 2

        elements_db = self.create_elements_database()

        best_match = None
        min_distance = float('inf')

        for element, orbitals in elements_db.items():
            for orbital, (be_min, be_max) in orbitals.items():
                be_center = (be_min + be_max) / 2

                # Check if range overlaps with binding energy range
                if not (range_max < be_min or range_min > be_max):
                    distance = abs(range_center - be_center)
                    if distance < min_distance:
                        min_distance = distance
                        best_match = f"{element}{orbital}"

        if best_match:
            self.peak_label_text.SetValue(best_match)



    def get_linux_desktop(self):
        """Detect Linux desktop environment"""
        import os

        # Check environment variables
        desktop = os.environ.get('XDG_CURRENT_DESKTOP', '').lower()
        if desktop:
            if 'gnome' in desktop:
                return 'gnome'
            elif 'kde' in desktop or 'plasma' in desktop:
                return 'kde'
            elif 'xfce' in desktop:
                return 'xfce'
            elif 'mate' in desktop:
                return 'mate'
            elif 'cinnamon' in desktop:
                return 'cinnamon'

        # Fallback checks
        session = os.environ.get('DESKTOP_SESSION', '').lower()
        if 'gnome' in session:
            return 'gnome'
        elif 'kde' in session:
            return 'kde'
        elif 'xfce' in session:
            return 'xfce'

        return 'unknown'
