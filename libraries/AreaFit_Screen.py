import wx
from libraries.Plot_Operations import PlotManager
from libraries.Save import save_state
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

        # Add averaging points control
        averaging_points_label = wx.StaticText(panel, label="Averaging Points:")
        self.averaging_points_text = wx.TextCtrl(panel, value="5")
        self.averaging_points_text.Bind(wx.EVT_TEXT, self.on_averaging_points_change)

        # Add Tougaard controls
        self.cross_section_label = wx.StaticText(panel, label='Tougaard1: B,C,D,T0')
        self.cross_section = wx.TextCtrl(panel, value="2866,1643,1,0")
        self.cross_section.Bind(wx.EVT_TEXT, self.on_cross_section_change)

        self.cross_section2_label = wx.StaticText(panel, label='Tougaard2: B,C,D,T0')
        self.cross_section2 = wx.TextCtrl(panel, value="2866,1643,1,0")
        self.cross_section2.Bind(wx.EVT_TEXT, self.on_cross_section2_change)

        self.cross_section3_label = wx.StaticText(panel, label='Tougaard3: B,C,D,T0')
        self.cross_section3 = wx.TextCtrl(panel, value="2866,1643,1,0")
        self.cross_section3.Bind(wx.EVT_TEXT, self.on_cross_section3_change)

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


        background_only_button = wx.Button(panel, label="Create\nBackground")
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
        if 'wxMac' in wx.PlatformInfo:
            sizer = wx.GridBagSizer(hgap=1, vgap=1)
        else:
            sizer = wx.GridBagSizer(hgap=0, vgap=0)

        if 'wxMac' in wx.PlatformInfo:
            # First row: Method
            sizer.Add(method_label, pos=(0, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.method_combobox, pos=(0, 1), flag=wx.ALL | wx.EXPAND, border=1)

            # second Third row: Offset (H) and Offset (L)
            sizer.Add(offset_h_label, pos=(1, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.offset_h_text, pos=(1, 1), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(offset_l_label, pos=(2, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.offset_l_text, pos=(2, 1), flag=wx.ALL | wx.EXPAND, border=1)

            # Fourth row: Averaging Points
            sizer.Add(averaging_points_label, pos=(3, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.averaging_points_text, pos=(3, 1), flag=wx.ALL | wx.EXPAND, border=1)

            # Fourth row: Tougaard parameters
            sizer.Add(self.cross_section_label, pos=(5, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.cross_section, pos=(5, 1), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.cross_section2_label, pos=(6, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.cross_section2, pos=(6, 1), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.cross_section3_label, pos=(7, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.cross_section3, pos=(7, 1), flag=wx.ALL | wx.EXPAND, border=1)

            # Area row
            area_box = wx.StaticBox(panel, label="Area Calculation")
            area_sizer = wx.StaticBoxSizer(area_box, wx.VERTICAL)

            # sizer.Add(peak_label_text_label, pos=(7, 0), flag=wx.ALL | wx.EXPAND, border=5)
            # sizer.Add(self.peak_label_text, pos=(7, 1), flag=wx.ALL | wx.EXPAND, border=5)

            text_sizer = wx.BoxSizer(wx.HORIZONTAL)
            text_sizer.Add(peak_label_text_label, 0, flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=5) #wx.ALL | wx.EXPAND,
            # 5) # | wx.ALIGN_CENTER_VERTICAL, 5)
            text_sizer.Add(self.peak_label_text, 1, flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=5) #wx.ALL | wx.EXPAND,
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

            # Fourth row: Averaging Points
            sizer.Add(averaging_points_label, pos=(3, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.averaging_points_text, pos=(3, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

            # Fourth row: Tougaard parameters
            sizer.Add(self.cross_section_label, pos=(5, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.cross_section, pos=(5, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.cross_section2_label, pos=(6, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.cross_section2, pos=(6, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.cross_section3_label, pos=(7, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.cross_section3, pos=(7, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

            # Area row
            area_box = wx.StaticBox(panel, label="Area Calculation")
            area_sizer = wx.StaticBoxSizer(area_box, wx.VERTICAL)

            # sizer.Add(peak_label_text_label, pos=(7, 0), flag=wx.ALL | wx.EXPAND, border=5)
            # sizer.Add(self.peak_label_text, pos=(7, 1), flag=wx.ALL | wx.EXPAND, border=5)

            text_sizer = wx.BoxSizer(wx.HORIZONTAL)
            text_sizer.Add(peak_label_text_label, 0, flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=5) #wx.ALL | wx.EXPAND,
            # 5) # | wx.ALIGN_CENTER_VERTICAL, 5)
            text_sizer.Add(self.peak_label_text, 1, flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=5) #wx.ALL | wx.EXPAND,
            # 5) #| wx.EXPAND, 5)

            button_sizer = wx.BoxSizer(wx.HORIZONTAL)
            button_sizer.Add(area_button, 1, flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            button_sizer.Add(remove_peak_button, 1, flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

            area_sizer.Add(text_sizer, 0, wx.EXPAND)
            area_sizer.Add(button_sizer, 1, wx.EXPAND)

            sizer.Add(area_sizer, pos=(9, 0), span=(2, 2), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

            # sizer.Add(area_button, pos=(8, 0), flag=wx.ALL | wx.EXPAND, border=5)
            # sizer.Add(remove_peak_button, pos=(8, 1), flag=wx.ALL | wx.EXPAND, border=5)

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
        self.cross_section2.Enable(False)
        self.cross_section2_label.Enable(False)
        self.cross_section3.Enable(False)
        self.cross_section3_label.Enable(False)
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
            if "2x" in new_method:
                self.cross_section2.Enable(True)
                self.cross_section2_label.Enable(True)
                self.cross_section3.Enable(False)
                self.cross_section3_label.Enable(False)
            elif "3x" in new_method:
                self.cross_section2.Enable(True)
                self.cross_section2_label.Enable(True)
                self.cross_section3.Enable(True)
                self.cross_section3_label.Enable(True)
            else:
                self.cross_section2.Enable(False)
                self.cross_section2_label.Enable(False)
                self.cross_section3.Enable(False)
                self.cross_section3_label.Enable(False)
        else:
            self.cross_section.Enable(False)
            self.cross_section_label.Enable(False)
            self.tougaard_fit_btn.Enable(False)
            self.cross_section2.Enable(False)
            self.cross_section2_label.Enable(False)
            self.cross_section3.Enable(False)
            self.cross_section3_label.Enable(False)

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
        self.parent.vline1 = None
        self.parent.vline2 = None
        self.parent.show_hide_vlines()
        self.parent.plot_manager.clear_and_replot(self.parent)
        save_state(self.parent)

    def on_toggle_vlines(self, event):
        if self.parent.vline1:
            self.parent.vline1.set_visible(not self.parent.vline1.get_visible())
        if self.parent.vline2:
            self.parent.vline2.set_visible(not self.parent.vline2.get_visible())
        self.parent.canvas.draw_idle()

    def on_background_only(self, event):
        selected_method = self.method_combobox.GetValue()
        self.parent.background_method = selected_method
        self.parent.offset_h = float(self.offset_h_text.GetValue())
        self.parent.offset_l = float(self.offset_l_text.GetValue())

        sheet_name = self.parent.sheet_combobox.GetValue()
        x_values = np.array(self.parent.Data['Core levels'][sheet_name]['B.E.'])

        if self.parent.vline1 is None or self.parent.vline2 is None:
            min_x = min(x_values) + 0.2
            max_x = max(x_values) - 0.2
            self.parent.vline1 = self.parent.ax.axvline(min_x, color='r', linestyle='--')
            self.parent.vline2 = self.parent.ax.axvline(max_x, color='r', linestyle='--')
            self.parent.bg_min_energy = min_x
            self.parent.bg_max_energy = max_x

            if 'Background' in self.parent.Data['Core levels'][sheet_name]:
                self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Low'] = min_x
                self.parent.Data['Core levels'][sheet_name]['Background']['Bkg High'] = max_x

        self.parent.plot_manager.plot_background(self.parent)
        self.on_reset_vlines(self)
        save_state(self.parent)

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

    def on_cross_section2_change(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        values = self.cross_section2.GetValue().split(',')
        try:
            self.parent.Data['Core levels'][sheet_name]['Background'].update({
                'Tougaard_B2': float(values[0]),
                'Tougaard_C2': float(values[1]),
                'Tougaard_D2': float(values[2]),
                'Tougaard_T02': float(values[3])
            })
        except (ValueError, IndexError):
            pass

    def on_cross_section3_change(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        values = self.cross_section3.GetValue().split(',')
        try:
            self.parent.Data['Core levels'][sheet_name]['Background'].update({
                'Tougaard_B3': float(values[0]),
                'Tougaard_C3': float(values[1]),
                'Tougaard_D3': float(values[2]),
                'Tougaard_T03': float(values[3])
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
        from libraries.Fitting_Screen import TougaardFitWindow
        TougaardFitWindow(self).Show()

    def on_close(self, event):
        self.parent.background_tab_selected = False
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
