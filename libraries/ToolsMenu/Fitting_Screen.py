
import re
import wx
from Functions import fit_peaks, remove_peak
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
import matplotlib.pyplot as plt
import lmfit
from libraries.FileMenu.Save import save_state
from libraries.ToolsMenu.TougaardRaman_Screen import TougaardRamanFitWindow

class FittingWindow(wx.Frame):
    def __init__(self, parent, normal=True, *args, **kw):
        super().__init__(parent, *args, **kw, style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX | wx.MINIMIZE_BOX | wx.SYSTEM_MENU) | wx.STAY_ON_TOP)
        self.parent = parent  # Store reference to MainFrame
        self.normal = normal  # Track if this is normal or mini mode

        self._updating_offsets = False

        self.SetTitle("Peak Fitting" if normal else "Mini Peak Fitting")


        if normal:
            if 'wxMac' in wx.PlatformInfo:
                self.SetSize((262, 380))  # Increased height to accommodate new elements
                self.SetMinSize((262, 380))
                self.SetMaxSize((300, 580))
            elif 'wxGTK' in wx.PlatformInfo:  # This is for Linux
                desktop = self.get_linux_desktop()
                if desktop == 'gnome':
                    self.SetSize((320, 570))
                elif desktop == 'kde':
                    self.SetSize((290, 430))
                elif desktop == 'xfce':
                    print('linux xfce')
                    self.SetSize((260, 490))
                else:  # Unknown or other
                    self.SetSize((280, 520))
                print(f'GTK environment: {desktop}')
                print('GTK environment')
            else:
                self.SetSize((275, 400))  # Increased height to accommodate new elements
                self.SetMinSize((275, 400))
                self.SetMaxSize((275, 400))
        else:
            # Mini mode - smaller window
            if 'wxMac' in wx.PlatformInfo:
                self.SetSize((262, 280))
                self.SetMinSize((262, 280))
                self.SetMaxSize((262, 350))
            elif 'wxGTK' in wx.PlatformInfo:
                self.SetSize((260, 320))
            else:
                self.SetSize((275, 220))



        #305 480

        self.init_ui()
        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.library_data = self.parent.library_data
        self.doublet_splittings = self.load_doublet_splittings(self.parent.library_data)

        from libraries.ConfigFile import set_consistent_fonts
        set_consistent_fonts(self)




    def init_ui(self):
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
            panel.SetBackgroundColour(wx.Colour(230, 250, 250))

        main_sizer = wx.BoxSizer(wx.VERTICAL)


        notebook = wx.Notebook(panel)

        # Bind notebook page change event
        notebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.on_notebook_page_changed)

        if not detect_dark_mode():
            notebook.SetBackgroundColour(wx.Colour(240, 250, 250))
        self.init_background_tab(notebook)
        self.init_fitting_tab(notebook)

        main_sizer.Add(notebook, 1, wx.EXPAND,border=5)
        panel.SetSizer(main_sizer)

        notebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.on_tab_change)

        self.parent.background_tab_selected = True
        self.parent.peak_fitting_tab_selected = False
        self.parent.peak_manipulation.deselect_all_peaks()

        self.vline1_text = None
        self.vline2_text = None

        # Trick: Quickly switch to fitting tab and back to background tab to trigger proper vline initialization
        wx.CallAfter(self._switch_tabs_trick, notebook)

    def _switch_tabs_trick(self, notebook):
        """Quick tab switching trick to ensure proper vline initialization"""
        # Safety check: make sure the window and notebook still exist
        try:
            if not self or not notebook:
                return

            # Check if the window is being destroyed
            if not self.IsShown() or self.IsBeingDeleted():
                return

            # Check if notebook is still valid
            notebook.GetSelection()  # This will raise an exception if notebook is deleted

            # Switch to peak fitting tab (index 1)
            notebook.SetSelection(1)
            # Process any pending events
            wx.GetApp().Yield()
            # Switch back to background tab (index 0)
            notebook.SetSelection(0)

        except (RuntimeError, AttributeError, TypeError):
            # Window or notebook has been destroyed, silently return
            return
        except:
            # Catch any other unexpected exceptions
            return

    def init_background_tab(self, notebook):
        """Initialize the background tab in the notebook."""
        # self.parent.background_method = "Multi-Regions Smart"
        self.parent.background_method = "Smart"
        self.parent.offset_h = 0
        self.parent.offset_l = 0

        # self.background_ranges = []  # List to store (offset_h, offset_l, min_range, max_range) tuples


        self.background_panel = wx.Panel(notebook)
        background_sizer = wx.GridBagSizer(hgap=0, vgap=0)

        method_label = wx.StaticText(self.background_panel, label="Method:")
        if self.normal:
            self.method_combobox = wx.ComboBox(self.background_panel, choices=["Smart", "Shirley",
                                                "Linear", 'U4-Tougaard', 'U2-Tougaard',"ALS-Raman"],
                                               style=wx.CB_READONLY)
        else:
            self.method_combobox = wx.ComboBox(self.background_panel, choices=["Smart",'U2-Tougaard'],
                                               style=wx.CB_READONLY)

        self.method_combobox.SetMaxSize((125,25))

        method_index = self.method_combobox.FindString(self.parent.background_method)
        self.method_combobox.SetSelection(method_index)
        self.method_combobox.Bind(wx.EVT_COMBOBOX, self.on_bkg_method_change)

        # info_button = self.create_info_button(self.background_panel,
        #                                       self.get_background_description(self.parent.background_method))

        offset_h_label = wx.StaticText(self.background_panel, label="Offset (Left):")
        self.offset_h_text = wx.TextCtrl(self.background_panel, value=f"{self.parent.offset_h:.2f}",
                                         style=wx.TE_PROCESS_ENTER)
        self.offset_h_text.Bind(wx.EVT_TEXT_ENTER, self.on_offset_h_change)
        self.offset_h_text.Bind(wx.EVT_CHAR, self.validate_numeric_input)
        self.offset_h_text.Bind(wx.EVT_KEY_DOWN, self.on_key_down)

        offset_l_label = wx.StaticText(self.background_panel, label="Offset (Right):")
        self.offset_l_text = wx.TextCtrl(self.background_panel, value=f"{self.parent.offset_l:.2f}",
                                         style=wx.TE_PROCESS_ENTER)
        self.offset_l_text.Bind(wx.EVT_TEXT_ENTER, self.on_offset_l_change)
        self.offset_l_text.Bind(wx.EVT_CHAR, self.validate_numeric_input)
        self.offset_l_text.Bind(wx.EVT_KEY_DOWN, self.on_key_down)

        self.min_range_label = wx.StaticText(self.background_panel, label='Region (Right):')
        self.min_range_text = wx.TextCtrl(self.background_panel, value="0.00", style=wx.TE_PROCESS_ENTER)
        self.min_range_text.Bind(wx.EVT_TEXT_ENTER, self.on_min_range_change)
        self.min_range_text.Bind(wx.EVT_CHAR, self.validate_numeric_input)
        self.min_range_text.Bind(wx.EVT_KEY_DOWN, self.on_key_down)

        self.max_range_label = wx.StaticText(self.background_panel, label='Region (Left):')
        self.max_range_text = wx.TextCtrl(self.background_panel, value="0.00", style=wx.TE_PROCESS_ENTER)
        self.max_range_text.Bind(wx.EVT_TEXT_ENTER, self.on_max_range_change)
        self.max_range_text.Bind(wx.EVT_CHAR, self.validate_numeric_input)
        self.max_range_text.Bind(wx.EVT_KEY_DOWN, self.on_key_down)

        if not self.normal:
            # Hide these controls
            self.offset_h_text.Hide()
            offset_h_label.Hide()
            self.offset_l_text.Hide()
            offset_l_label.Hide()
            self.min_range_text.Hide()
            self.min_range_label.Hide()
            self.max_range_text.Hide()
            self.max_range_label.Hide()



        # Add this flag to prevent recursive updates
        self.updating_range_controls = False

        # Add active range tracking (after creating range boxes)
        self.active_range_index = -1  # -1 means no active range, 0+ means active range index

        # Initialize with current vline positions from data
        self.update_range_controls_from_data()

        averaging_points_label = wx.StaticText(self.background_panel, label="Averaging Points:")
        self.averaging_points_text = wx.TextCtrl(self.background_panel, value="5")
        self.averaging_points_text.Bind(wx.EVT_TEXT, self.on_averaging_points_change)
        if not self.normal:
            averaging_points_label.Hide()
            self.averaging_points_text.Hide()

        # Add after averaging_points_text line:
        smooth_data_label = wx.StaticText(self.background_panel, label="Smooth noisy data:")
        self.smooth_data_checkbox = wx.CheckBox(self.background_panel, label="")
        self.smooth_data_checkbox.SetToolTip("Apply Gaussian smoothing (width=2) to data before calculating shirley background")
        self.smooth_data_checkbox.Bind(wx.EVT_CHECKBOX, self.on_smooth_data_change)
        if not self.normal:
            smooth_data_label.Hide()
            self.smooth_data_checkbox.Hide()


        self.cross_section_label = wx.StaticText(self.background_panel, label = 'Tougaard1: B,C,D,T0')
        self.cross_section = wx.TextCtrl(self.background_panel, value="2866,1643,1,0")
        self.cross_section.Bind(wx.EVT_TEXT, self.on_cross_section_change)
        if not self.normal:
            self.cross_section_label.Hide()
            self.cross_section.Hide()

        # Range selection boxes (after self.cross_section creation)
        self.range_boxes_label = wx.StaticText(self.background_panel, label='Regions:')
        self.range_boxes_panel = wx.Panel(self.background_panel)
        self.range_boxes_panel.SetMinSize((-1, 50))  # Allow for 2 rows of buttons with spacing

        # Create main vertical sizer for the panel
        self.range_boxes_main_sizer = wx.BoxSizer(wx.VERTICAL)
        # Create row horizontal sizer
        self.range_boxes_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.range_boxes_sizer2 = wx.BoxSizer(wx.HORIZONTAL)

        # Add both rows to main sizer with spacing
        self.range_boxes_main_sizer.Add(self.range_boxes_sizer, 0, wx.EXPAND)
        self.range_boxes_main_sizer.Add(self.range_boxes_sizer2, 0, wx.EXPAND | wx.TOP, 0)  # 5px spacing

        self.range_boxes = []  # List to store range selection buttons
        self.range_boxes_panel.SetSizer(self.range_boxes_main_sizer)

        sheet_name = self.parent.sheet_combobox.GetValue()
        if sheet_name != '' :
            # if 'Background' in self.parent.Data['Core levels'][sheet_name]:
            if sheet_name in self.parent.Data['Core levels'] and 'Background' in self.parent.Data['Core levels'][
                sheet_name]:
                bg_data = self.parent.Data['Core levels'][sheet_name]['Background']
                saved_values = [
                    bg_data.get('Tougaard_B', 2866),
                    bg_data.get('Tougaard_C', 1643),
                    bg_data.get('Tougaard_D', 1),
                    bg_data.get('Tougaard_T0', 0)
                ]
                saved_values2 = [
                    bg_data.get('Tougaard_B2', 2866),
                    bg_data.get('Tougaard_C2', 1643),
                    bg_data.get('Tougaard_D2', 1),
                    bg_data.get('Tougaard_T02', 0)
                ]
                saved_values3 = [
                    bg_data.get('Tougaard_B2', 2866),
                    bg_data.get('Tougaard_C2', 1643),
                    bg_data.get('Tougaard_D2', 1),
                    bg_data.get('Tougaard_T02', 0)
                ]
                self.cross_section.SetValue(','.join(map(str, saved_values)))
                # self.cross_section2.SetValue(','.join(map(str, saved_values2)))
                # self.cross_section2.SetValue(','.join(map(str, saved_values3)))

        background_button = wx.Button(self.background_panel, label="Create\nRegion")
        if 'wxMac' in wx.PlatformInfo:
            background_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            background_button.SetMinSize((125, 35))
        else:
            background_button.SetMinSize((125, 35))
        background_button.Bind(wx.EVT_BUTTON, self.on_background)

        clear_background_button = wx.Button(self.background_panel, label="Remove\nRegions and Peaks")
        if 'wxMac' in wx.PlatformInfo:
            clear_background_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            clear_background_button.SetMinSize((125, 35))
        else:
            clear_background_button.SetMinSize((125, 35))
        clear_background_button.Bind(wx.EVT_BUTTON, self.on_clear_background)

        reset_vlines_button = wx.Button(self.background_panel, label="Switch Region\nTAB key")
        if 'wxMac' in wx.PlatformInfo:
            reset_vlines_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            reset_vlines_button.SetMinSize((125, 35))
        else:
            reset_vlines_button.SetMinSize((125, 35))
        reset_vlines_button.Bind(wx.EVT_BUTTON, self.on_reset_vlines2)

        remove_active_region_button = wx.Button(self.background_panel, label="Remove\nCurrent Region")
        if 'wxMac' in wx.PlatformInfo:
            remove_active_region_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            remove_active_region_button.SetMinSize((125, 35))
        else:
            remove_active_region_button.SetMinSize((125, 35))
        remove_active_region_button.Bind(wx.EVT_BUTTON, self.on_remove_active_region)

        clear_background_only_button = wx.Button(self.background_panel, label="Remove\nAll Regions")
        if 'wxMac' in wx.PlatformInfo:
            clear_background_only_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            clear_background_only_button.SetMinSize((125, 35))
        else:
            clear_background_only_button.SetMinSize((125, 35))
        clear_background_only_button.Bind(wx.EVT_BUTTON, self.on_clear_background_only)

        self.tougaard_fit_btn = wx.Button(self.background_panel, label="Tougaard / Raman\n Model")
        if 'wxMac' in wx.PlatformInfo:
            self.tougaard_fit_btn.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            self.tougaard_fit_btn.SetMinSize((125, 35))
        else:
            self.tougaard_fit_btn.SetMinSize((125, 35))
        self.tougaard_fit_btn.Bind(wx.EVT_BUTTON, self.on_tougaard_raman_model)

        if not self.normal:
            self.tougaard_fit_btn.Hide()
            clear_background_button.Hide()

        if 'wxMac' in wx.PlatformInfo or 'wxGTK' in wx.PlatformInfo:
            # Layout Background Tab
            background_sizer.Add(method_label, pos=(0, 0), flag=wx.ALL | wx.EXPAND, border=1)
            background_sizer.Add(self.method_combobox, pos=(0, 1), flag=wx.ALL | wx.EXPAND, border=1)

            if self.normal:
                # background_sizer.Add(info_button, pos=(1, 1), flag=wx.ALL | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(offset_h_label, pos=(1, 0), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(self.offset_h_text, pos=(1, 1), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(offset_l_label, pos=(2, 0), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(self.offset_l_text, pos=(2, 1), flag=wx.ALL | wx.EXPAND, border=1)

                background_sizer.Add(self.min_range_label, pos=(4, 0), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(self.min_range_text, pos=(4, 1), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(self.max_range_label, pos=(3, 0), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(self.max_range_text, pos=(3, 1), flag=wx.ALL | wx.EXPAND, border=1)

                background_sizer.Add(averaging_points_label, pos=(5, 0), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(self.averaging_points_text, pos=(5, 1), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(smooth_data_label, pos=(6, 0), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(self.smooth_data_checkbox, pos=(6, 1), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(self.cross_section_label,  pos=(7, 0), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(self.cross_section, pos=(7, 1), flag=wx.ALL | wx.EXPAND, border=1)

                # Add range boxes
                background_sizer.Add(self.range_boxes_label, pos=(8, 0), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(self.range_boxes_panel, pos=(8, 1), flag=wx.ALL | wx.EXPAND, border=1)

                background_sizer.Add(reset_vlines_button, pos=(9, 0), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(clear_background_button, pos=(9, 1), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(self.tougaard_fit_btn, pos=(10, 0), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(clear_background_only_button, pos=(10, 1), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(background_button, pos=(11, 0), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(remove_active_region_button, pos=(11, 1), flag=wx.ALL | wx.EXPAND, border=1)
            else:
                # Add range boxes
                background_sizer.Add(self.range_boxes_label, pos=(1, 0), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(self.range_boxes_panel, pos=(1, 1), flag=wx.ALL | wx.EXPAND, border=1)

                background_sizer.Add(reset_vlines_button, pos=(2, 0), flag=wx.ALL | wx.EXPAND, border=1)
                # background_sizer.Add(clear_background_button, pos=(2, 1), flag=wx.ALL | wx.EXPAND, border=1)
                # background_sizer.Add(self.tougaard_fit_btn, pos=(3, 0), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(clear_background_only_button, pos=(3, 1), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(background_button, pos=(4, 0), flag=wx.ALL | wx.EXPAND, border=1)
                background_sizer.Add(remove_active_region_button, pos=(4, 1), flag=wx.ALL | wx.EXPAND, border=1)

        else:

            background_sizer.Add(method_label, pos=(0, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            background_sizer.Add(self.method_combobox, pos=(0, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            if self.normal:
                # background_sizer.Add(info_button, pos=(1, 1), flag=wx.ALL | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(offset_h_label, pos=(1, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(self.offset_h_text, pos=(1, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(offset_l_label, pos=(2, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(self.offset_l_text, pos=(2, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

                background_sizer.Add(self.min_range_label, pos=(4, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(self.min_range_text, pos=(4, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(self.max_range_label, pos=(3, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(self.max_range_text, pos=(3, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

                background_sizer.Add(averaging_points_label, pos=(5, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(self.averaging_points_text, pos=(5, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(smooth_data_label, pos=(6, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(self.smooth_data_checkbox, pos=(6, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(self.cross_section_label,  pos=(7, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(self.cross_section, pos=(7, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

                # Add range boxes
                background_sizer.Add(self.range_boxes_label, pos=(8, 0), flag=wx.ALL | wx.EXPAND, border=0)
                background_sizer.Add(self.range_boxes_panel, pos=(8, 1), flag=wx.ALL | wx.EXPAND, border=0)

                background_sizer.Add(reset_vlines_button, pos=(9, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(clear_background_button, pos=(9, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(self.tougaard_fit_btn, pos=(10, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(clear_background_only_button, pos=(10, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP,
                                     border=0)
                background_sizer.Add(background_button, pos=(11, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(remove_active_region_button, pos=(11, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            else:
                # Add range boxes
                background_sizer.Add(self.range_boxes_label, pos=(1, 0), flag=wx.ALL | wx.EXPAND, border=0)
                background_sizer.Add(self.range_boxes_panel, pos=(1, 1), flag=wx.ALL | wx.EXPAND, border=0)

                background_sizer.Add(reset_vlines_button, pos=(2, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                # background_sizer.Add(clear_background_button, pos=(2, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                # background_sizer.Add(self.tougaard_fit_btn, pos=(3, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(clear_background_only_button, pos=(2, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP,
                                     border=0)
                background_sizer.Add(background_button, pos=(3, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                background_sizer.Add(remove_active_region_button, pos=(3, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)


        self.background_panel.SetSizer(background_sizer)
        notebook.AddPage(self.background_panel, "Background")


        self.update_tougaard_controls_visibility(self.parent.background_method)

    def init_fitting_tab(self, notebook):
        """Initialize the peak fitting tab in the notebook."""
        self.fitting_panel = wx.Panel(notebook)
        fitting_sizer = wx.GridBagSizer(hgap=0, vgap=0)


        self.model_combobox = CustomComboBox(self.fitting_panel, style=wx.CB_READONLY)
        self.model_combobox.SetMaxSize((125, 25))

        # Set items and green items
        items = ["Best Models-----------------------",
                 "SGL (Area)",
                 "GL (Area)",
                 "LA (Area, \u03c3/\u03b3, \u03b3)",
                 "Voigt (Area, L/G, \u03c3)",
                 "Asymmetric------------------------",
                 "LA (Area, \u03c3, \u03b3)",
                 "DS*G (A, \u03c3, \u03b3, S)",
                 "DS (A, \u03c3, \u03b3)",
                 "ExpGauss.(Area, \u03c3, \u03b3)",
                 "Voigt-----------------------------",
                 "Voigt (Area, L/G, \u03c3)",
                 "Voigt (Area, \u03c3, \u03b3)",
                 "Voigt (Area, L/G, \u03c3, S)",
                 "LA--------------------------------",
                 "LA (Area, \u03c3/\u03b3, \u03b3)",
                 "LA (Area, \u03c3, \u03b3)",
                 "LA*G (Area, \u03c3/\u03b3, \u03b3)",
                 "Others----------------------------",
                 "Pseudo-Voigt (Area)",
                 "GL (Height)",
                 "SGL (Height)",
                 "Beta------------------------------",

                 "Voigt (Area, L/G, \u03c3, S)"
                 ]

        if not self.normal:
            items = ["SGL (Area)",
                     "Voigt (Area, L/G, \u03c3)",
                     "LA (Area, \u03c3, \u03b3)",
                     ]


        green_items = ["Others---------------","Area Based----------", "Height Based---------", "Under Test "
                                                                                               "-----------", "Preferred Models-----"]
        self.model_combobox.SetItems(items)
        self.model_combobox.SetGreenItems(green_items)

        # Set default selection to "GL (Area)"
        default_index = items.index("SGL (Area)")
        self.model_combobox.SetSelection(default_index)
        self.model_combobox.SetValue("SGL (Area)")
        self.parent.set_fitting_method("SGL (Area)")



        # model_index = self.model_combobox.FindString(self.parent.selected_fitting_method)
        # self.model_combobox.SetSelection(model_index)
        self.model_combobox.Bind(wx.EVT_COMBOBOX, self.on_method_change)

        info_button = self.create_info_button(self.fitting_panel,
                                              self.get_fitting_description(self.parent.selected_fitting_method))
        info_button.SetMaxSize((20,20))
        # if not self.normal:
        #     info_button.Hide()

        self.optimization_method = wx.ComboBox(self.fitting_panel, choices=[
            "leastsq",
            "least_squares",
            # "differential_evolution",
            "nelder",
            # "lbfgsb",
            "powell",
            # "cg",
            # "newton",
            "cobyla",
            # "basinhopping",
            # "slsqp",
            # "tnc",
            # "trust-krylov",
            "trust-constr"
            # "dogleg",
            # "shgo",
            # "dual_annealing"
        ], style=wx.CB_READONLY)
        self.optimization_method.SetSelection(1)  # Default value
        if not self.normal:
            self.optimization_method.Hide()

        self.max_iter_spin = wx.SpinCtrl(self.fitting_panel, value=str(self.parent.max_iterations), min=20, max=500)
        self.max_iter_spin.Bind(wx.EVT_SPINCTRL, self.on_max_iter_change)
        if not self.normal:
            self.max_iter_spin.Hide()

        self.fit_iterations_spin = wx.SpinCtrl(self.fitting_panel, value="20", min=2, max=100)
        if not self.normal:
            self.fit_iterations_spin.Hide()

        self.weights_combo = wx.ComboBox(self.fitting_panel, choices=["uniform", "intensity-based", "statistical-XPS",
                                                                      "hybrid-XPS"],
                                         style=wx.CB_READONLY)
        if not self.normal:
            self.weights_combo.Hide()

        self.weights_combo.SetSelection(0)  # Default to "uniform"

        self.r_squared_label = wx.StaticText(self.fitting_panel, label="R²:")
        self.r_squared_text = wx.TextCtrl(self.fitting_panel, style=wx.TE_READONLY)
        if not self.normal:
            self.r_squared_label.Hide()
            self.r_squared_text.Hide()


        # self.chi_squared_label = wx.StaticText(self.fitting_panel, label="Chi²:")
        # self.chi_squared_text = wx.TextCtrl(self.fitting_panel, style=wx.TE_READONLY)

        # self.rsd_label = wx.StaticText(self.fitting_panel, label="RSD:")
        # self.rsd_text = wx.TextCtrl(self.fitting_panel, style=wx.TE_READONLY)


        self.red_chi_squared_label = wx.StaticText(self.fitting_panel, label="Red. Chi²:")
        self.red_chi_squared_text = wx.TextCtrl(self.fitting_panel, style=wx.TE_READONLY)
        if not self.normal:
            self.red_chi_squared_text.Hide()
            self.red_chi_squared_text.Hide()

        # self.actual_iter_label = wx.StaticText(self.fitting_panel, label="Actual Iterations:")
        # self.actual_iter_text = wx.TextCtrl(self.fitting_panel, style=wx.TE_READONLY)

        self.current_fit_label = wx.StaticText(self.fitting_panel, label="Current Fit:")
        self.current_fit_text = wx.TextCtrl(self.fitting_panel, style=wx.TE_READONLY)

        fit_report_button = wx.Button(self.fitting_panel, label="Report")
        fit_report_button.SetMinSize((125, 30))
        fit_report_button.Bind(wx.EVT_BUTTON, self.on_fit_report)
        if not self.normal:
            fit_report_button.Hide()

        # This button size controls all the size of the button on the left side
        add_peak_button = wx.Button(self.fitting_panel, label="Add 1 Peak\nSinglet")
        if 'wxMac' in wx.PlatformInfo:
            add_peak_button.SetMinSize((125, 30))
            pass
        elif 'wxGTK' in wx.PlatformInfo:
            add_peak_button.SetMinSize((125, 35))
        else:
            add_peak_button.SetMinSize((125, 35))
        add_peak_button.Bind(wx.EVT_BUTTON, self.on_add_peak)

        add_doublet_button = wx.Button(self.fitting_panel, label="Add 2 Peaks\nDoublet")
        if 'wxMac' in wx.PlatformInfo:
            add_doublet_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            add_doublet_button.SetMinSize((125, 35))
        else:
            add_doublet_button.SetMinSize((125, 35))
        add_doublet_button.Bind(wx.EVT_BUTTON, self.on_add_doublet)

        remove_peak_button = wx.Button(self.fitting_panel, label="Remove\nLast Peak")
        if 'wxMac' in wx.PlatformInfo:
            remove_peak_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            remove_peak_button.SetMinSize((125, 35))
        else:
            remove_peak_button.SetMinSize((125, 35))
        remove_peak_button.Bind(wx.EVT_BUTTON, self.on_remove_peak)

        # export_button = wx.Button(self.fitting_panel, label="Export to\nResults Grid")
        # if 'wxMac' in wx.PlatformInfo:
        #     export_button.SetMinSize((125, 30))
        # elif 'wxGTK' in wx.PlatformInfo:
        #     export_button.SetMinSize((125, 35))
        # else:
        #    export_button.SetMinSize((125, 35))
        # export_button.Bind(wx.EVT_BUTTON, self.on_export_results)

        add_named_doublet_button = wx.Button(self.fitting_panel, label="Add Doublet\nwith Name...")
        if 'wxMac' in wx.PlatformInfo:
            add_named_doublet_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            add_named_doublet_button.SetMinSize((125, 35))
        else:
            add_named_doublet_button.SetMinSize((125, 35))
        add_named_doublet_button.Bind(wx.EVT_BUTTON, self.on_add_named_doublet)
        if not self.normal:
            add_named_doublet_button.Hide()

        fit_button = wx.Button(self.fitting_panel, label="Fit \nOne Time")
        if 'wxMac' in wx.PlatformInfo:
            fit_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            fit_button.SetMinSize((125, 35))
        else:
            fit_button.SetMinSize((125, 35))
        fit_button.Bind(wx.EVT_BUTTON, self.on_fit_peaks)
        if not self.normal:
            fit_button.Hide()

        fit_multi_button = wx.Button(self.fitting_panel, label="Fit \nN# Times")
        if 'wxMac' in wx.PlatformInfo:
            fit_multi_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            fit_multi_button.SetMinSize((125, 35))
        else:
            fit_multi_button.SetMinSize((125, 35))
        fit_multi_button.Bind(wx.EVT_BUTTON, self.on_fit_multi)

        if 'wxMac' in wx.PlatformInfo or 'wxGTK' in wx.PlatformInfo:
            fitting_sizer.Add(wx.StaticText(self.fitting_panel, label="Fitting Model:"), pos=(0, 0),
                              flag=wx.ALL | wx.EXPAND, border=1)
            fitting_sizer.Add(self.model_combobox, pos=(0, 1), flag=wx.ALL, border=0)
            fitting_sizer.Add(info_button, pos=(1, 1), flag=wx.ALL, border=0)
            if self.normal:
                fitting_sizer.Add(wx.StaticText(self.fitting_panel, label="Method:"), pos=(2, 0),
                                  flag=wx.ALL | wx.EXPAND, border=1)
                fitting_sizer.Add(self.optimization_method, pos=(2, 1), flag=wx.ALL | wx.EXPAND, border=1)

                fitting_sizer.Add(wx.StaticText(self.fitting_panel, label="Convergence: "), pos=(3, 0),
                                  flag=wx.ALL | wx.EXPAND, border=1)
                fitting_sizer.Add(self.max_iter_spin, pos=(3, 1), flag=wx.ALL | wx.EXPAND, border=1)
                fitting_sizer.Add(wx.StaticText(self.fitting_panel, label="N# of Iterations:"), pos=(4, 0),
                                  flag=wx.ALL | wx.EXPAND, border=1)
                fitting_sizer.Add(self.fit_iterations_spin, pos=(4, 1), flag=wx.ALL | wx.EXPAND, border=1)

                fitting_sizer.Add(wx.StaticText(self.fitting_panel, label="Weights:"), pos=(5, 0),
                                  flag=wx.ALL | wx.EXPAND, border=1)
                fitting_sizer.Add(self.weights_combo, pos=(5, 1), flag=wx.ALL | wx.EXPAND, border=1)

                fitting_sizer.Add(self.r_squared_label, pos=(6, 0), flag=wx.ALL | wx.EXPAND, border=1)
                fitting_sizer.Add(self.r_squared_text, pos=(6, 1), flag=wx.ALL | wx.EXPAND, border=1)
                # fitting_sizer.Add(self.rsd_label, pos=(6, 0), flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
                # fitting_sizer.Add(self.rsd_text, pos=(6, 1), flag=wx.ALL | wx.EXPAND, border=5)
                fitting_sizer.Add(self.red_chi_squared_label, pos=(7, 0), flag=wx.ALL | wx.EXPAND, border=1)
                fitting_sizer.Add(self.red_chi_squared_text, pos=(7, 1), flag=wx.ALL | wx.EXPAND, border=1)

                # fitting_sizer.Add(self.actual_iter_label, pos=(8, 0), flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
                # fitting_sizer.Add(self.actual_iter_text, pos=(8, 1), flag=wx.ALL | wx.EXPAND, border=5)
                fitting_sizer.Add(self.current_fit_label, pos=(8, 0), flag=wx.ALL | wx.EXPAND, border=1)
                fitting_sizer.Add(self.current_fit_text, pos=(8, 1), flag=wx.ALL | wx.EXPAND, border=1)

                fitting_sizer.Add(fit_report_button, pos=(9, 1), flag=wx.ALL | wx.EXPAND, border=1)

                fitting_sizer.Add(add_peak_button, pos=(10, 0), flag=wx.ALL | wx.EXPAND, border=1)
                fitting_sizer.Add(add_doublet_button, pos=(10, 1), flag=wx.ALL | wx.EXPAND, border=1)

                fitting_sizer.Add(remove_peak_button, pos=(11, 0), flag=wx.ALL | wx.EXPAND, border=1)
                # fitting_sizer.Add(export_button, pos=(11, 1), flag=wx.ALL | wx.EXPAND, border=1)
                fitting_sizer.Add(add_named_doublet_button, pos=(11, 1), flag=wx.ALL | wx.EXPAND, border=1)

                # fitting_sizer.Add(fit_button, pos=(13, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(fit_button, pos=(12, 0), flag=wx.ALL | wx.EXPAND, border=1)
                # fitting_sizer.Add(fit_multi_button, pos=(13, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(fit_multi_button, pos=(12, 1), flag=wx.ALL | wx.EXPAND, border=1)
            else:
                fitting_sizer.Add(self.current_fit_label, pos=(2, 0), flag=wx.ALL | wx.EXPAND, border=1)
                fitting_sizer.Add(self.current_fit_text, pos=(2, 1), flag=wx.ALL | wx.EXPAND, border=1)
                fitting_sizer.Add(add_peak_button, pos=(4, 0), flag=wx.ALL | wx.EXPAND, border=1)
                fitting_sizer.Add(add_doublet_button, pos=(4, 1), flag=wx.ALL | wx.EXPAND, border=1)

                fitting_sizer.Add(remove_peak_button, pos=(5, 0), flag=wx.ALL | wx.EXPAND, border=1)
                fitting_sizer.Add(fit_multi_button, pos=(5, 1), flag=wx.ALL | wx.EXPAND, border=1)
        else:

            fitting_sizer.Add(wx.StaticText(self.fitting_panel, label="Fitting Model:"), pos=(0, 0),
                              flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            fitting_sizer.Add(self.model_combobox, pos=(0, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            fitting_sizer.Add(info_button, pos=(1, 1), flag= wx.ALL | wx.BOTTOM | wx.TOP, border=0)

            if self.normal:
                fitting_sizer.Add(wx.StaticText(self.fitting_panel, label="Method:"), pos=(2, 0),
                                  flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(self.optimization_method, pos=(2, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

                fitting_sizer.Add(wx.StaticText(self.fitting_panel, label="Convergence: "), pos=(3, 0),
                                  flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(self.max_iter_spin, pos=(3, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(wx.StaticText(self.fitting_panel, label="N# of Iterations:"), pos=(4, 0),
                                  flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(self.fit_iterations_spin, pos=(4, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

                fitting_sizer.Add(wx.StaticText(self.fitting_panel, label="Weights:"), pos=(5, 0),
                                  flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(self.weights_combo, pos=(5, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

                fitting_sizer.Add(self.r_squared_label, pos=(6, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(self.r_squared_text, pos=(6, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                # fitting_sizer.Add(self.rsd_label, pos=(6, 0), flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
                # fitting_sizer.Add(self.rsd_text, pos=(6, 1), flag=wx.ALL | wx.EXPAND, border=5)
                fitting_sizer.Add(self.red_chi_squared_label, pos=(7, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(self.red_chi_squared_text, pos=(7, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

                # fitting_sizer.Add(self.actual_iter_label, pos=(8, 0), flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)
                # fitting_sizer.Add(self.actual_iter_text, pos=(8, 1), flag=wx.ALL | wx.EXPAND, border=5)
                fitting_sizer.Add(self.current_fit_label, pos=(8, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(self.current_fit_text, pos=(8, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

                fitting_sizer.Add(fit_report_button, pos=(9, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

                fitting_sizer.Add(add_peak_button, pos=(10, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(add_doublet_button, pos=(10, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

                fitting_sizer.Add(remove_peak_button, pos=(11, 0), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                # fitting_sizer.Add(export_button, pos=(11, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(add_named_doublet_button, pos=(11, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

                fitting_sizer.Add(fit_button, pos=(12, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(fit_multi_button, pos=(12, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            else:
                fitting_sizer.Add(self.current_fit_label, pos=(2, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(self.current_fit_text, pos=(2, 1), flag= wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(add_peak_button, pos=(4, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(add_doublet_button, pos=(4, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

                fitting_sizer.Add(remove_peak_button, pos=(5, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
                fitting_sizer.Add(fit_multi_button, pos=(5, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

        self.fitting_panel.SetSizer(fitting_sizer)
        notebook.AddPage(self.fitting_panel, "Peak Fitting")

    def on_fit_report(self, event):
        if hasattr(self.parent, 'fit_results') and 'result' in self.parent.fit_results:
            dlg = wx.Frame(self, title="Fit Report", size=(600, 400))
            text = wx.TextCtrl(dlg, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.HSCROLL)

            result = self.parent.fit_results['result']
            report = result.fit_report()
            report += "\n\nAdditional Information:"
            report += f"\nNumber of function evaluations: {result.nfev}"
            report += f"\nNumber of variables: {result.nvarys}"
            report += f"\nNumber of data points: {result.ndata}"
            report += f"\nDegrees of freedom: {result.nfree}"
            report += f"\nAborted: {result.aborted}"
            report += f"\nError flags: {result.errorbars}"
            report += f"\nCovariance matrix: \n{result.covar}"

            text.SetValue(report)

            sizer = wx.BoxSizer(wx.VERTICAL)
            sizer.Add(text, 1, wx.EXPAND | wx.ALL, 5)

            dlg.SetSizer(sizer)
            dlg.Show()
        else:
            # wx.MessageBox("No fit report available. Please perform a fit first.", "No Report")
            self.parent.show_popup_message2("No Report", "No fit report available. Please perform a fit first.")

    def get_optimization_method(self):
        selection = self.optimization_method.GetValue()
        return selection.split()[0]  # Get just the method name without description

    def get_weights_method(self):
        return self.weights_combo.GetStringSelection()

    def on_smooth_data_change(self, event):
        save_state(self.parent)
        # Pass - we'll access the checkbox state directly when calculating background

    def on_method_change_OLD(self, event):
        new_method = self.model_combobox.GetValue()
        old_method = self.parent.selected_fitting_method
        self.parent.set_fitting_method(new_method)
        self.update_fitting_info_button()

    def on_method_change(self, event):
        new_method = self.model_combobox.GetValue()
        if "skew" in new_method:
            self.optimization_method.SetValue("leastsq")
            self.optimization_method.Enable(True)
        else:
            self.optimization_method.SetValue("least_squares")
            self.optimization_method.Enable(True)
        self.parent.set_fitting_method(new_method)
        self.update_fitting_info_button()

    def on_max_iter_change(self, event):
        self.parent.set_max_iterations(self.max_iter_spin.GetValue())

    # def on_method_change(self, event):
    #     new_method = self.model_combobox.GetValue()
    #     self.parent.set_fitting_method(new_method)

    def on_bkg_method_change(self, event):
        new_method = self.method_combobox.GetValue()
        self.parent.set_background_method(new_method)
        self.update_background_info_button()
        self.update_tougaard_controls_visibility(new_method)

    def on_tougaard_raman_model(self, event):
        bg_method = self.method_combobox.GetValue()
        if bg_method.startswith("U4-Tougaard") or bg_method.startswith("U2-Tougaard") or bg_method.startswith("2x U4-Tougaard") or \
                bg_method.startswith("3x U4-Tougaard"):
            tougaard_window = TougaardFitWindow(self)
            tougaard_window.Show()
        elif bg_method == "ALS-Raman":
            raman_window = TougaardRamanFitWindow(self)
            raman_window.Show()
        else:
            self.parent.show_popup_message2("Error",
                                            "Tougaard/Raman Model is only available for Tougaard and ALS-Raman background methods.")

    def update_tougaard_controls_visibility(self, new_method):
        if new_method.startswith("U4-Tougaard"):
            self.cross_section.Enable(True)
            self.cross_section_label.Enable(True)
            self.tougaard_fit_btn.Enable(True)
        elif new_method == "U2-Tougaard":
            self.cross_section.Enable(False)
            self.cross_section_label.Enable(False)
            self.tougaard_fit_btn.Enable(False)
        elif new_method == "ALS-Raman":
            self.cross_section.Enable(False)
            self.cross_section_label.Enable(False)
            self.tougaard_fit_btn.Enable(True)
        else:
            self.cross_section.Enable(False)
            self.cross_section_label.Enable(False)
            self.tougaard_fit_btn.Enable(False)
        self.background_panel.Layout()


    def on_notebook_page_changed(self, event):
        """Handle notebook page changes to initialize ranges when background tab is selected"""
        if event.GetSelection() == 0:  # Background tab index
            self.initialize_recorded_ranges()
        event.Skip()

    def on_averaging_points_change(self, event):
        try:
            new_value = int(self.averaging_points_text.GetValue())

            # Ensure averaging points is always at least 1
            if new_value <= 0:
                new_value = 1
                self.averaging_points_text.SetValue("1")

            self.parent.averaging_points = new_value

            # Update averaging indicator lines when points change
            if hasattr(self.parent, 'add_averaging_indicator_lines'):
                self.parent.add_averaging_indicator_lines()

        except ValueError:
            # If invalid input, reset to 1
            self.averaging_points_text.SetValue("1")
            self.parent.averaging_points = 1

    def on_reset_vlines(self, event):
        # if hasattr(self.parent, 'x_values') and len(self.parent.x_values) > 0:
        #     x_min = min(self.parent.x_values)
        #     x_max = max(self.parent.x_values)
        #     x_range = x_max - x_min
        #
        #     new_low = x_min + x_range / 15
        #     new_high = x_min + 14 * x_range / 15
        #
        #     # Update bg_min_energy and bg_max_energy
        #     self.parent.bg_min_energy = new_low
        #     self.parent.bg_max_energy = new_high

        # # Reset vlines to None first
        # self.parent.vline1 = None
        # self.parent.vline2 = None
        # if hasattr(self.parent, 'vline1_text'):
        #     self.parent.vline1_text = None
        # if hasattr(self.parent, 'vline2_text'):
        #     self.parent.vline2_text = None

        # Use tab switching trick to properly recreate vlines at new positions
        if self.parent.background_tab_selected:
            # Find the notebook
            notebook = None
            for child in self.GetChildren():
                for grandchild in child.GetChildren():
                    if isinstance(grandchild, wx.Notebook):
                        notebook = grandchild
                        break
                if notebook:
                    break

            if notebook:
                wx.CallAfter(self._switch_tabs_trick, notebook)
        else:
            # If not on background tab, just clear and replot
            self.parent.plot_manager.clear_and_replot(self.parent)

    def on_reset_vlines2(self, event):
        # NEW: Check for region switching first
        ranges = self.get_recorded_ranges_from_data()
        if ranges:
            # Get current active region
            current_active = getattr(self, 'active_range_index', -1)

            # Calculate next region index (cycle back to 0 after last region)
            if current_active < 0:
                # No active region, start with region 1 (index 0)
                next_region = 0
            else:
                # Move to next region, cycling back to 0 after the last one
                next_region = (current_active + 1) % len(ranges)

            # Set the next region as active
            self.set_active_range(next_region)

            # Get the next region's data
            offset_h, offset_l, min_range, max_range = ranges[next_region]

            # Update controls with next region values
            self.updating_range_controls = True
            self._updating_offsets = True
            self.offset_h_text.SetValue(f"{offset_h:.2f}")
            self.offset_l_text.SetValue(f"{offset_l:.2f}")
            self.min_range_text.SetValue(f"{min_range:.2f}")
            self.max_range_text.SetValue(f"{max_range:.2f}")
            self._updating_offsets = False
            self.updating_range_controls = False

            # Move vLines to the next region's positions
            if self.parent.vline1 is not None:
                min_display = self.parent.convert_energy_for_display(min_range)
                self.parent.vline1.set_xdata([min_display, min_display])
            if self.parent.vline2 is not None:
                max_display = self.parent.convert_energy_for_display(max_range)
                self.parent.vline2.set_xdata([max_display, max_display])

            # Update text labels
            if hasattr(self.parent, 'update_vline_text_labels'):
                self.parent.update_vline_text_labels()

            # Update averaging indicator lines
            if hasattr(self.parent, 'add_averaging_indicator_lines'):
                self.parent.add_averaging_indicator_lines()

            # Force canvas redraw
            self.parent.canvas.draw_idle()

            # Update window.Data background range
            self.update_window_data_background_range()

            # Update visual highlighting
            wx.CallAfter(self.update_range_box_colors)
            wx.CallAfter(self.force_range_box_refresh)

            # print(f"Switched to region {next_region + 1}: {min_range:.2f} - {max_range:.2f}")
            return

        # # ORIGINAL CODE: Calculate 1/15 and 14/15 positions
        # if hasattr(self.parent, 'x_values') and len(self.parent.x_values) > 0:
        #     x_min = min(self.parent.x_values)
        #     x_max = max(self.parent.x_values)
        #     x_range = x_max - x_min
        #
        #     new_low = x_min + x_range / 15
        #     new_high = x_min + 14 * x_range / 15
        #
        #     # Update bg_min_energy and bg_max_energy
        #     self.parent.bg_min_energy = new_low
        #     self.parent.bg_max_energy = new_high

        # Reset vlines to None first
        self.parent.vline1 = None
        self.parent.vline2 = None
        if hasattr(self.parent, 'vline1_text'):
            self.parent.vline1_text = None
        if hasattr(self.parent, 'vline2_text'):
            self.parent.vline2_text = None

        # # Use tab switching trick to properly recreate vlines at new positions
        # if self.parent.background_tab_selected:
        #     # Find the notebook
        #     notebook = None
        #     for child in self.GetChildren():
        #         for grandchild in child.GetChildren():
        #             if isinstance(grandchild, wx.Notebook):
        #                 notebook = grandchild
        #                 break
        #         if notebook:
        #             break
        #
        #     if notebook:
        #         wx.CallAfter(self._switch_tabs_trick, notebook)
        # else:
        #     # If not on background tab, just clear and replot
        #     self.parent.plot_manager.clear_and_replot(self.parent)

    def on_add_peak(self, event):
        current_model = self.model_combobox.GetValue()
        if "----" in current_model:
            self.parent.show_popup_message("Please select a valid peak model.")
            return
        self.parent.peak_fitting_grid.add_peak_params()

    def on_remove_peak(self, event):
        save_state(self.parent)
        remove_peak(self.parent)

    def on_fit_multi_OLD(self, event):
        save_state(self.parent)
        if self.parent.peak_params_grid.GetNumberRows() == 0:
            # wx.MessageBox("No peaks to fit. Add at least one peak first.", "Error", wx.OK | wx.ICON_ERROR)
            self.parent.show_popup_message2("Error", "No peaks to fit. Add at least one peak first.")
            return
        save_state(self.parent)
        iterations = self.fit_iterations_spin.GetValue()
        for i in range(1, iterations + 1):
            self.current_fit_text.SetValue(f"{i}/{iterations}")
            result = fit_peaks(self.parent, self.parent.peak_params_grid)
            if result:
                r_squared, rsd, red_chi_square = result
                self.update_fit_indicators(r_squared, rsd, red_chi_square)
                # self.actual_iter_text.SetValue(str(self.parent.fit_results['nfev']))
            self.parent.clear_and_replot()
            wx.Yield()
        self.current_fit_text.SetValue("Complete")

    def on_fit_multi(self, event):
        save_state(self.parent)
        if self.parent.peak_params_grid.GetNumberRows() == 0:
            self.parent.show_popup_message2("Error", "No peaks to fit. Add at least one peak first.")
            return

        # Disable UI controls to prevent interference
        self.disable_fitting_ui()

        try:
            save_state(self.parent)
            iterations = self.fit_iterations_spin.GetValue()
            for i in range(1, iterations + 1):
                self.current_fit_text.SetValue(f"{i}/{iterations}")
                result = fit_peaks(self.parent, self.parent.peak_params_grid)
                if result:
                    r_squared, rsd, red_chi_square = result
                    self.update_fit_indicators(r_squared, rsd, red_chi_square)
                self.parent.clear_and_replot()
                wx.Yield()
            self.current_fit_text.SetValue("Complete")
        finally:
            # Always re-enable UI controls, even if fitting fails
            self.enable_fitting_ui()


    def on_fit_peaks(self, event):
        save_state(self.parent)
        result = fit_peaks(self.parent, self.parent.peak_params_grid)
        if result:
            r_squared, rsd, red_chi_squared = result
            self.update_fit_indicators(r_squared, rsd, red_chi_squared)
        else:
            print("Fitting failed or was cancelled.")

    def update_fit_indicators(self, r_squared, rsd, red_chi_squared):
        self.r_squared_text.SetValue(f"{r_squared:.5f}")
        # self.rsd_text.SetValue(f"{rsd:.5f}")
        self.red_chi_squared_text.SetValue(f"{red_chi_squared:.2f}")
        self.Layout()

    def on_export_results(self, event):
        save_state(self.parent)
        self.parent.export_results()

    def load_doublet_splittings(self, library_data):
        splittings = {}

        return


    def get_doublet_splitting(self, element, orbital, instrument):
        # Extract number and letter from orbital (e.g., "3d" -> "3" and "d")
        orbital_match = re.match(r'(\d)([spdf])', orbital)
        if orbital_match:
            num, letter = orbital_match.groups()
            key = (element, f"{num}{letter}")

            if key in self.library_data and 'C-Al1486' in self.library_data[key]:
                print(f'Found splitting for {element}{num}{letter}: {self.library_data[key]["C-Al1486"]["ds"]}')
                return self.library_data[key]['C-Al1486']['ds']
            else:
                print(f'No splitting found for key: {key}')

        print(f"Warning: No doublet splitting found for {element} {orbital}")
        return 0.0


    def on_add_doublet(self, event):
        current_model = self.model_combobox.GetValue()
        if "----" in current_model:
            self.parent.show_popup_message("Please select a valid peak model.")
            return
        save_state(self.parent)
        if self.parent.bg_min_energy is None or self.parent.bg_max_energy is None:
            # wx.MessageBox("Please create a background first.", "No Background", wx.OK | wx.ICON_WARNING)
            self.parent.show_popup_message2("No Background", "Please create a background first.")
            return

        # Get overall background range for doublet creation
        overall_bg_low, overall_bg_high = self.get_overall_background_range()

        sheet_name = self.parent.sheet_combobox.GetValue()

        # Extract element and orbital, ignoring sample numbers
        element_orbital_match = re.match(r'([A-Z][a-z]*\d+[spdf])', sheet_name)
        if element_orbital_match:
            element_orbital = element_orbital_match.group(1)
        else:
            element_orbital = sheet_name.split()[0]

        # Check if it's a valid orbital for doublet
        orbital = re.search(r'(\d+[spdf])', element_orbital)
        if not orbital:
            # wx.MessageBox("Cannot fit doublet peak on a S orbital core level.",
            #               "Error", wx.OK | wx.ICON_ERROR)
            self.parent.show_popup_message2("Error", "Cannot fit doublet peak on a S orbital core level.")
            return

        orbital = orbital.group()
        element = re.match(r'([A-Z][a-z]*)', element_orbital).group(1)

        if orbital[-1] == 's':
            self.parent.peak_fitting_grid.add_peak_params()
        else:
            # Create the doublet peaks - first peak uses highest intensity
            first_peak = self.parent.peak_fitting_grid.add_peak_params()

            # Get the first peak's position and intensity from grid for second peak calculation
            first_peak_row = first_peak * 2
            first_peak_position = float(self.parent.peak_params_grid.GetCellValue(first_peak_row, 2))
            first_peak_height = float(self.parent.peak_params_grid.GetCellValue(first_peak_row, 3))

            # Calculate second peak intensity based on orbital constraint
            orbital_type = orbital[-1]  # Extract p, d, or f
            intensity_factors = {'p': 0.50, 'd': 0.667, 'f': 0.75}
            intensity_factor = intensity_factors.get(orbital_type, 0.50)

            # Calculate second peak position using doublet splitting
            splitting = self.get_doublet_splitting(element, orbital, self.parent.current_instrument)
            second_peak_position = first_peak_position + splitting
            second_peak_height = first_peak_height * intensity_factor

            # Create second peak with calculated intensity
            second_peak = self.parent.peak_fitting_grid.add_peak_params(
                custom_peak_x=second_peak_position,
                custom_peak_y=second_peak_height
            )

            self.parent.peak_fill_types[second_peak] = self.parent.peak_fill_types[first_peak]
            self.parent.peak_hatch_patterns[second_peak] = self.parent.peak_hatch_patterns[first_peak]
            self.hatch_density = 2

            # Set constraints for the second peak
            row1 = first_peak * 2
            row2 = second_peak * 2

            # L/G ratio constraint
            if any(element in element_orbital for element in ['Ti2p', 'V2p']) and any(
                    x in self.parent.selected_fitting_method for x in ["Voigt (Area, L/G"]):
                lg_constraint = "2:80"  # Independent L/G for Ti2p and V2p
            else:
                lg_constraint = f"{chr(65 + first_peak)}*1"
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 5, lg_constraint)

            # FWHM constraint
            if any(element in element_orbital for element in ['Ti2p', 'V2p']) and any(
                    x in self.parent.selected_fitting_method for x in ["LA", "GL", "SGL"]):
                fwhm_constraint = "0.3:3.5"  # Independent FWHM for Ti2p and V2p
            else:
                if "Voigt" in self.parent.selected_fitting_method:
                    fwhm_constraint = "0.3:3.5"
                else:
                    fwhm_constraint = f"{chr(65 + first_peak)}*1"
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 4, fwhm_constraint)

            # Height and Area constraints based on fitting model
            height_factor = {'p': 0.5, 'd': 0.667, 'f': 0.75}
            area_factor = {'p': 0.5, 'd': 0.667, 'f': 0.75}

            current_model = self.parent.selected_fitting_method

            # Check if model uses area or height as primary parameter
            area_based_models = ["GL (Area)", "SGL (Area)", "Pseudo-Voigt (Area)",
                                 "Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)",
                                 "Voigt (Area, L/G, \u03c3, S)", "ExpGauss.(Area, \u03c3, \u03b3)",
                                 "LA (Area, \u03c3, \u03b3)", "LA (Area, \u03c3/\u03b3, \u03b3)",
                                 "LA*G (Area, \u03c3/\u03b3, \u03b3)", "DS (A, \u03c3, \u03b3)",
                                 "DS*G (A, \u03c3, \u03b3, S)"]

            # Set default constraints
            height_constraint = "1:1e7"
            area_constraint = "1:1e7"

            if current_model in area_based_models:
                # For area-based models, constrain area and use default for height
                area_constraint = f"{chr(65 + first_peak)}*{area_factor[orbital[-1]]}#0.01"
            else:
                # For height-based models, constrain height and use default for area
                height_constraint = f"{chr(65 + first_peak)}*{height_factor[orbital[-1]]}#0.01"

            # Apply constraints
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 3, height_constraint)
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 6, area_constraint)

            # Position constraint
            element = re.match(r'([A-Z][a-z]*)', element_orbital).group(1)
            splitting = self.get_doublet_splitting(element, orbital, self.parent.current_instrument)

            position_constraint = f"{chr(65 + first_peak)}+{splitting}#0.2"
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 2, position_constraint)

            # Sigma constraint
            # if any(element in element_orbital for element in ['Ti2p', 'V2p']) and any(
            #         x in self.parent.selected_fitting_method for x in ["Voigt (Area, L/G"]):
            #     sigma_constraint = "0.3:3"  # Independent FWHM for Ti2p and V2p
            # else:
            sigma_constraint = f"{chr(65 + first_peak)}*1"
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 7, sigma_constraint)

            # Gamma constraint
            if any(element in element_orbital for element in ['Ti2p', 'V2p']) and any(
                    x in self.parent.selected_fitting_method for x in ["Voigt (Area, \u03c3", "DS*G"]):
                gamma_constraint = "0.3:3"  # Independent FWHM for Ti2p and V2p. gaussian is fixed but lorentzian vary
            else:
                gamma_constraint = f"{chr(65 + first_peak)}*1"
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 8, gamma_constraint)

            # Skew constraint
            skew_constraint = f"{chr(65 + first_peak)}*1"
            self.parent.peak_params_grid.SetCellValue(row2 + 1, 9, skew_constraint)

            # Calculate peak numbers
            peak_number1 = first_peak + 1
            peak_number2 = second_peak + 1

            # Set peak names with correct suborbital designations
            if orbital[-1] == 'p':
                peak1_name = f"{element_orbital}3/2 p{peak_number1}"
                peak2_name = f"{element_orbital}1/2_p{peak_number2}"
            elif orbital[-1] == 'd':
                peak1_name = f"{element_orbital}5/2 p{peak_number1}"
                peak2_name = f"{element_orbital}3/2_p{peak_number2}"
            elif orbital[-1] == 'f':
                peak1_name = f"{element_orbital}7/2 p{peak_number1}"
                peak2_name = f"{element_orbital}5/2_p{peak_number2}"

            self.parent.peak_params_grid.SetCellValue(row1, 1, peak1_name)
            self.parent.peak_params_grid.SetCellValue(row2, 1, peak2_name)

            # Position the second peak
            first_peak_position = float(self.parent.peak_params_grid.GetCellValue(row1, 2))
            second_peak_position = first_peak_position + splitting
            self.parent.peak_params_grid.SetCellValue(row2, 2, f"{second_peak_position:.2f}")

            # Update window.Data with new constraints and names
            if 'Fitting' in self.parent.Data['Core levels'][sheet_name] and 'Peaks' in \
                    self.parent.Data['Core levels'][sheet_name]['Fitting']:
                peaks = self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                new_peaks = {}
                for i, (key, value) in enumerate(peaks.items()):
                    if i == first_peak:
                        new_peaks[peak1_name] = value
                        new_peaks[peak1_name]['Name'] = peak1_name
                    elif i == second_peak:
                        new_peaks[peak2_name] = value
                        new_peaks[peak2_name]['Name'] = peak2_name
                        new_peaks[peak2_name]['Position'] = second_peak_position
                        new_peaks[peak2_name]['Constraints'] = {
                            'Position': position_constraint,
                            'Height': height_constraint,
                            'FWHM': fwhm_constraint,
                            'L/G': lg_constraint,
                            'Area': area_constraint,
                            'Sigma': sigma_constraint,
                            'Gamma': gamma_constraint,
                            'Skew': skew_constraint
                        }
                    else:
                        new_peaks[key] = value
                self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = new_peaks

        self.parent.clear_and_replot()

    def on_add_named_doublet(self, event):
        current_model = self.model_combobox.GetValue()
        if "----" in current_model:
            self.parent.show_popup_message("Please select a valid peak model.")
            return

        save_state(self.parent)

        if self.parent.bg_min_energy is None or self.parent.bg_max_energy is None:
            self.parent.show_popup_message2("No Background", "Please create a background first.")
            return

        # Get custom core level name from user
        dlg = wx.TextEntryDialog(self,
                                 'Enter a core level name for the doublet\n(e.g., Fe2p, Si2p, Ag3d, Au4f.. & NOT C1s or 2s or 3s):',
                                 'Doublet Core Level Name',
                                 '')

        # Set the dialog background color to match the fitting screen
        fitting_bg_color = wx.Colour(240, 250, 250)
        dlg.SetBackgroundColour(fitting_bg_color)

        # Also set the text control background if needed
        for child in dlg.GetChildren():
            if isinstance(child, wx.TextCtrl):
                child.SetBackgroundColour(fitting_bg_color)
            elif hasattr(child, 'SetBackgroundColour'):
                child.SetBackgroundColour(fitting_bg_color)

        # Position the dialog on top of the fitting window
        fitting_pos = self.GetPosition()
        fitting_size = self.GetSize()
        dlg_size = dlg.GetSize()

        # Center the dialog on the fitting window
        x = fitting_pos.x + (fitting_size.width - dlg_size.width) // 2
        y = fitting_pos.y + (fitting_size.height - dlg_size.height) // 2
        dlg.SetPosition((x, y))

        # Keep dialog on top
        dlg.SetWindowStyle(dlg.GetWindowStyle() | wx.STAY_ON_TOP)

        # Refresh the dialog to apply color changes
        dlg.Refresh()

        if dlg.ShowModal() != wx.ID_OK:
            dlg.Destroy()
            return

        custom_name = dlg.GetValue().strip()
        dlg.Destroy()

        if not custom_name:
            self.parent.show_popup_message2("Error", "Please enter a valid core level name.")
            return

        # Extract orbital information from the custom name
        orbital_match = re.search(r'(\d+[spdf])', custom_name)
        if not orbital_match:
            self.parent.show_popup_message2("Error", "Invalid core level name. Must contain orbital (e.g., 1s, 2p, 3d, 4f).")
            return

        orbital = orbital_match.group()
        element_match = re.match(r'([A-Z][a-z]*)', custom_name)
        if not element_match:
            self.parent.show_popup_message2("Error", "Invalid core level name. Must start with element symbol.")
            return

        element = element_match.group(1)

        # Check if it's a valid orbital for doublet
        if orbital[-1] == 's':
            self.parent.show_popup_message2("Error", "Cannot fit doublet peak on a S orbital core level.")
            return

        # Get overall background range for doublet creation
        overall_bg_low, overall_bg_high = self.get_overall_background_range()
        sheet_name = self.parent.sheet_combobox.GetValue()

        # Create the doublet peaks - first peak uses highest intensity
        first_peak = self.parent.peak_fitting_grid.add_peak_params()

        # Get the first peak's position and intensity from grid for second peak calculation
        first_peak_row = first_peak * 2
        first_peak_position = float(self.parent.peak_params_grid.GetCellValue(first_peak_row, 2))
        first_peak_height = float(self.parent.peak_params_grid.GetCellValue(first_peak_row, 3))

        # Calculate second peak intensity based on orbital constraint
        orbital_type = orbital[-1]  # Extract p, d, or f
        intensity_factors = {'p': 0.50, 'd': 0.667, 'f': 0.75}
        intensity_factor = intensity_factors.get(orbital_type, 0.50)

        # Calculate second peak position using doublet splitting
        splitting = self.get_doublet_splitting(element, orbital, self.parent.current_instrument)
        second_peak_position = first_peak_position + splitting
        second_peak_height = first_peak_height * intensity_factor

        # Create second peak with calculated intensity
        second_peak = self.parent.peak_fitting_grid.add_peak_params(
            custom_peak_x=second_peak_position,
            custom_peak_y=second_peak_height
        )

        self.parent.peak_fill_types[second_peak] = self.parent.peak_fill_types[first_peak]
        self.parent.peak_hatch_patterns[second_peak] = self.parent.peak_hatch_patterns[first_peak]
        self.hatch_density = 2

        # Set constraints for the second peak
        row1 = first_peak * 2
        row2 = second_peak * 2

        # L/G ratio constraint
        if any(element in custom_name for element in ['Ti2p', 'V2p']) and any(
                x in self.parent.selected_fitting_method for x in ["Voigt (Area, L/G"]):
            lg_constraint = "2:80"
        else:
            lg_constraint = f"{chr(65 + first_peak)}*1"
        self.parent.peak_params_grid.SetCellValue(row2 + 1, 5, lg_constraint)

        # FWHM constraint
        if any(element in custom_name for element in ['Ti2p', 'V2p']) and any(
                x in self.parent.selected_fitting_method for x in ["LA", "GL", "SGL"]):
            fwhm_constraint = "0.3:3.5"
        else:
            if "Voigt" in self.parent.selected_fitting_method:
                fwhm_constraint = "0.3:3.5"
            else:
                fwhm_constraint = f"{chr(65 + first_peak)}*1"
        self.parent.peak_params_grid.SetCellValue(row2 + 1, 4, fwhm_constraint)

        # Height and Area constraints based on orbital type from custom name
        if orbital[-1] == 'p':
            height_factor = 0.5
            area_factor = 0.5
        elif orbital[-1] == 'd':
            height_factor = 0.667
            area_factor = 0.667
        elif orbital[-1] == 'f':
            height_factor = 0.75
            area_factor = 0.75
        else:
            height_factor = 0.5
            area_factor = 0.5

        current_model = self.parent.selected_fitting_method

        # Check if model uses area or height as primary parameter
        area_based_models = ["GL (Area)", "SGL (Area)", "Pseudo-Voigt (Area)",
                             "Voigt (Area, L/G, \u03c3)", "Voigt (Area, \u03c3, \u03b3)",
                             "Voigt (Area, L/G, \u03c3, S)", "ExpGauss (Area, \u03c3, \u03b3)",
                             "LA (Area, \u03c3, \u03b3)", "LA (Area, \u03c3/\u03b3, \u03b3)",
                             "LA*G (Area, \u03c3/\u03b3, \u03b3)", "DS (A, \u03c3, \u03b3)",
                             "DS*G (A, \u03c3, \u03b3, S)"]

        # Set default constraints
        height_constraint = "1:1e7"
        area_constraint = "1:1e7"

        if current_model in area_based_models:
            area_constraint = f"{chr(65 + first_peak)}*{area_factor:.3f}#0.01"
        else:
            height_constraint = f"{chr(65 + first_peak)}*{height_factor:.3f}#0.01"

        # Apply constraints
        self.parent.peak_params_grid.SetCellValue(row2 + 1, 3, height_constraint)
        self.parent.peak_params_grid.SetCellValue(row2 + 1, 6, area_constraint)

        # Position constraint - get splitting from library using custom name
        splitting = self.get_doublet_splitting(element, orbital, self.parent.current_instrument)
        position_constraint = f"{chr(65 + first_peak)}+{splitting:.2f}#0.2"
        self.parent.peak_params_grid.SetCellValue(row2 + 1, 2, position_constraint)

        # Sigma constraint
        sigma_constraint = f"{chr(65 + first_peak)}*1"
        self.parent.peak_params_grid.SetCellValue(row2 + 1, 7, sigma_constraint)

        # Gamma constraint
        if any(element in custom_name for element in ['Ti2p', 'V2p']) and any(
                x in self.parent.selected_fitting_method for x in ["Voigt (Area, \u03c3", "DS*G"]):
            gamma_constraint = "0.3:3"
        else:
            gamma_constraint = f"{chr(65 + first_peak)}*1"
        self.parent.peak_params_grid.SetCellValue(row2 + 1, 8, gamma_constraint)

        # Skew constraint
        skew_constraint = f"{chr(65 + first_peak)}*1"
        self.parent.peak_params_grid.SetCellValue(row2 + 1, 9, skew_constraint)

        # Calculate peak numbers
        peak_number1 = first_peak + 1
        peak_number2 = second_peak + 1

        # Set peak names based on orbital type from custom name
        if orbital[-1] == 'p':
            peak1_name = f"{custom_name}3/2 p{peak_number1}"
            peak2_name = f"{custom_name}1/2_p{peak_number2}"
        elif orbital[-1] == 'd':
            peak1_name = f"{custom_name}5/2 p{peak_number1}"
            peak2_name = f"{custom_name}3/2_p{peak_number2}"
        elif orbital[-1] == 'f':
            peak1_name = f"{custom_name}7/2 p{peak_number1}"
            peak2_name = f"{custom_name}5/2_p{peak_number2}"

        self.parent.peak_params_grid.SetCellValue(row1, 1, peak1_name)
        self.parent.peak_params_grid.SetCellValue(row2, 1, peak2_name)

        # Position the second peak
        first_peak_position = float(self.parent.peak_params_grid.GetCellValue(row1, 2))
        second_peak_position = first_peak_position + splitting
        self.parent.peak_params_grid.SetCellValue(row2, 2, f"{second_peak_position:.2f}")

        # Update window.Data with new constraints and names
        if 'Fitting' in self.parent.Data['Core levels'][sheet_name] and 'Peaks' in \
                self.parent.Data['Core levels'][sheet_name]['Fitting']:
            peaks = self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']
            new_peaks = {}
            for i, (key, value) in enumerate(peaks.items()):
                if i == first_peak:
                    new_peaks[peak1_name] = value
                    new_peaks[peak1_name]['Name'] = peak1_name
                elif i == second_peak:
                    new_peaks[peak2_name] = value
                    new_peaks[peak2_name]['Name'] = peak2_name
                    new_peaks[peak2_name]['Position'] = second_peak_position
                    new_peaks[peak2_name]['Constraints'] = {
                        'Position': position_constraint,
                        'Height': height_constraint,
                        'FWHM': fwhm_constraint,
                        'L/G': lg_constraint,
                        'Area': area_constraint,
                        'Sigma': sigma_constraint,
                        'Gamma': gamma_constraint,
                        'Skew': skew_constraint
                    }
                else:
                    new_peaks[key] = value
            self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = new_peaks

        self.parent.clear_and_replot()

    def on_background(self, event):
        save_state(self.parent)

        # Validate and set background method from combobox
        selected_method = self.method_combobox.GetValue()
        if not selected_method or selected_method == "":
            selected_method = "Smart"  # Default fallback
            self.method_combobox.SetSelection(0)  # Select first item (Smart)
        self.parent.background_method = selected_method

        # Get vline positions if they exist, otherwise use full range
        if self.parent.vline1 is not None and self.parent.vline2 is not None:
            # Use actual vline positions
            vline1_pos = self.parent.vline1.get_xdata()[0]
            vline2_pos = self.parent.vline2.get_xdata()[0]
            self.parent.bg_min_energy = min(vline1_pos, vline2_pos)
            self.parent.bg_max_energy = max(vline1_pos, vline2_pos)

            # ADD THIS: Update the range controls with current vline positions
            self.updating_range_controls = True
            self.min_range_text.SetValue(f"{self.parent.bg_min_energy:.2f}")
            self.max_range_text.SetValue(f"{self.parent.bg_max_energy:.2f}")
            self.updating_range_controls = False
        elif self.parent.bg_min_energy is None or self.parent.bg_max_energy is None:
            # Fallback to full range only if no vlines and no stored values
            sheet_name = self.parent.sheet_combobox.GetValue()
            x_values = self.parent.Data['Core levels'][sheet_name]['B.E.']
            self.parent.bg_min_energy = min(x_values)
            self.parent.bg_max_energy = max(x_values)

            # ADD THIS: Update the range controls with full range values
            self.updating_range_controls = True
            self.min_range_text.SetValue(f"{self.parent.bg_min_energy:.2f}")
            self.max_range_text.SetValue(f"{self.parent.bg_max_energy:.2f}")
            self.updating_range_controls = False

        # Update the data structure with the correct values
        sheet_name = self.parent.sheet_combobox.GetValue()
        if 'Background' in self.parent.Data['Core levels'][sheet_name]:
            self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Low'] = self.parent.bg_min_energy
            self.parent.Data['Core levels'][sheet_name]['Background']['Bkg High'] = self.parent.bg_max_energy

        # Parse and save cross-section values when creating region
        self.parse_cross_section(self.cross_section.GetValue())

        # Get smooth checkbox state
        use_smoothing = self.smooth_data_checkbox.GetValue()

        # Call plot_background with the smoothing flag
        self.parent.plot_manager.plot_background(self.parent, use_smoothing=use_smoothing)

        # Update all peaks with new background information
        self.update_all_peaks_background_info()

        self.record_current_range()

        # Use tab switching trick to ensure vlines are properly displayed
        if self.parent.background_tab_selected:
            # Find the notebook
            notebook = None
            for child in self.GetChildren():
                for grandchild in child.GetChildren():
                    if isinstance(grandchild, wx.Notebook):
                        notebook = grandchild
                        break
                if notebook:
                    break

            if notebook:
                wx.CallAfter(self._switch_tabs_trick, notebook)

    def update_all_peaks_background_info(self):
        """Update background information for all peaks in the current core level"""
        sheet_name = self.parent.sheet_combobox.GetValue()

        # Get current background information
        current_bg_type = self.parent.background_method

        # Use overall background range instead of current vline positions
        overall_bg_low, overall_bg_high = self.get_overall_background_range()

        # Update peak fitting grid
        if hasattr(self.parent, 'peak_params_grid') and self.parent.peak_params_grid.GetNumberRows() > 0:
            num_peaks = self.parent.peak_params_grid.GetNumberRows() // 2

            for i in range(num_peaks):
                row = i * 2  # Each peak takes 2 rows (data + constraints)

                # Update grid columns: 14=Bkg Type, 15=Bkg Low, 16=Bkg High
                self.parent.peak_params_grid.SetCellValue(row, 14, current_bg_type)
                self.parent.peak_params_grid.SetCellValue(row, 15, f"{overall_bg_low:.2f}")
                self.parent.peak_params_grid.SetCellValue(row, 16, f"{overall_bg_high:.2f}")

        # Update window.Data structure
        if (sheet_name in self.parent.Data['Core levels'] and
                'Fitting' in self.parent.Data['Core levels'][sheet_name] and
                'Peaks' in self.parent.Data['Core levels'][sheet_name]['Fitting']):

            peaks = self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']

            for peak_label, peak_data in peaks.items():
                peak_data['Bkg Type'] = current_bg_type
                peak_data['Bkg Low'] = overall_bg_low
                peak_data['Bkg High'] = overall_bg_high

    def on_clear_background(self, event):
        # Show confirmation dialog
        dlg = wx.MessageDialog(self,
                               "Are you sure you want to clear all background data?",
                               "Confirm Clear All",
                               wx.YES_NO | wx.ICON_EXCLAMATION)

        result = dlg.ShowModal()
        dlg.Destroy()

        if result == wx.ID_YES:
            save_state(self.parent)
            self.clear_recorded_ranges_from_data()
            self.clear_active_range()  # Clear active range highlighting
            self.update_range_boxes()
            self.parent.plot_manager.clear_background(self.parent)
            self.parent.bg_min_energy = None
            self.parent.bg_max_energy = None
            self.parent.plot_data()
            save_state(self.parent)

    def on_clear_background_only(self, event):
        # Clear recorded ranges and active state
        self.clear_recorded_ranges_from_data()
        self.clear_active_range()  # Clear active range highlighting
        self.update_range_boxes()

        # Existing clear background only code
        self.parent.plot_manager.clear_background_only(self.parent)

    def set_active_range(self, index):
        """Set the active range and update visual highlighting"""
        self.active_range_index = index
        self.update_range_box_colors()

    def update_range_box_colors(self):
        """Update range box colors to show active state"""
        ranges = self.get_recorded_ranges_from_data()

        for i, box in enumerate(self.range_boxes):
            try:
                if i == self.active_range_index:
                    # Active range - yellow background
                    box.SetBackgroundColour(wx.Colour(255, 0, 0))  # Yellow
                    box.SetForegroundColour(wx.Colour(0, 0, 0))  # Black text
                else:
                    # Inactive range - default colors
                    box.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_BTNFACE))
                    box.SetForegroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_BTNTEXT))
                box.Refresh()
            except RuntimeError:
                # Button was destroyed, skip
                pass

    def clear_active_range(self):
        """Clear active range selection"""
        self.active_range_index = -1
        self.update_range_box_colors()

    def on_offset_h_change(self, event):
        if self._updating_offsets:  # Prevent recursion
            return

        self._updating_offsets = True
        try:
            save_state(self.parent)
            try:
                display_value = float(self.offset_h_text.GetValue())
                # Convert positive display to negative actual value
                offset_h_value = -abs(display_value)

                self.parent.set_offset_h(offset_h_value)

                # Update active range if one is selected
                if hasattr(self, 'active_range_index') and self.active_range_index >= 0:
                    offset_l_value = self.get_offset_actual_value(self.offset_l_text)
                    self.update_active_range_offsets(offset_h_value, offset_l_value)

                # Redraw background from all regions
                if hasattr(self.parent, 'mouse_handler') and hasattr(self.parent.mouse_handler,
                                                                     'redraw_all_regions_background'):
                    self.parent.mouse_handler.redraw_all_regions_background()

            except ValueError:
                self.parent.set_offset_h(0)
                self.offset_h_text.SetValue('0.0')
        finally:
            self._updating_offsets = False

    def on_offset_l_change(self, event):
        if self._updating_offsets:  # Prevent recursion
            return

        self._updating_offsets = True
        try:
            save_state(self.parent)
            try:
                display_value = float(self.offset_l_text.GetValue())
                # Convert positive display to negative actual value
                offset_l_value = -abs(display_value)

                self.parent.set_offset_l(offset_l_value)

                # Update active range if one is selected
                if hasattr(self, 'active_range_index') and self.active_range_index >= 0:
                    offset_h_value = self.get_offset_actual_value(self.offset_h_text)
                    self.update_active_range_offsets(offset_h_value, offset_l_value)

                # Redraw background from all regions
                if hasattr(self.parent, 'mouse_handler') and hasattr(self.parent.mouse_handler,
                                                                     'redraw_all_regions_background'):
                    self.parent.mouse_handler.redraw_all_regions_background()

            except ValueError:
                self.parent.set_offset_l(0)
                self.offset_l_text.SetValue('0.0')
        finally:
            self._updating_offsets = False

    def get_background_description(self, method):
        descriptions = {
            "Multi-Regions Smart": "Same as the Smart background but not restricted to a single region",
            "Smart": "Single region Smart background estimation. It applies a Shirley background "
                     "when the intensity is going up and a linear background when the intensity is "
                     "going down. If the calculated background is above the data then the "
                     "background is set equal to the data.",
            "Shirley": "Iterative background calculation. Reliable for increasing background when "
                       "the data contains symmetrical peak. the number of iteration is set to 100",
            "Linear": "Simple linear background. Usually used on negative background",
            "U4-Tougaard": "U4 Tougaard background for Advanced users. B, C, D and T0 can be varied.",
            "U2-Tougaard": "2-parameter Tougaard background (auto B, user C, D=0, T0=0)",
            "ALS-Raman": "Asymmetric Least Squares background estimation. Good for uneven backgrounds "
                         "with broad features. Lambda controls smoothness and p controls asymmetry."
        }
        return descriptions.get(method, "No description available")

    def get_fitting_description(self, model):
        descriptions = {
            "GL (Height)": "Gaussian-Lorentzian product function (height-based).  Constrained by Peak position, Height, FWHM and L/G. Based on GL from Thermo Avantage",
            "SGL (Height)": "Sum of Gaussian and Lorentzian functions (height-based).  Constrained by Peak position, Height, FWHM and L/G.   Based on GL from Thermo Avantage",
            "GL (Area)": "Product of Gaussian and Lorentzian model (area-based).  Constrained by Peak position, Area, FWHM and L/G.   Based on GL from CasaXPS",
            "SGL (Area)": "Sum of Gaussian and Lorentzian functions (area-based). Constrained by Peak position, Area, FWHM and L/G.   Based on SGL from CasaXPS",
            "Pseudo-Voigt (Area)": "Linear combination of Gaussian and Lorentzian provided by LMFIT.  Constrained by Peak position, Area, FWHM and L/G",
            "Voigt (Area, L/G, \u03c3, S)": "Asymmetric Voigt Function controlled by A, L/G, Wg and S.  Wl is "
                                            "calculated from Wg and L/G. In this case S controls the Asymmetry."
                                            " The model is provided by LMFIT",
            "Voigt (Area,\u03c3, \u03b3)": "Voigt function with separate Gaussian and Lorentzian widths.  "
                                           "It is controlled by A, Wg and Wl. Wl is independent of Wg and L/G  "
                                           "is calculated from Wl and Wg.  The model is provided by LMFIT",
            "Voigt (Area, L/G, \u03c3)": "Voigt Function controlled by A, L/G, Wg.   Wl is "
                                            "calculated from Wg and L/G. "
                                            "The Model is provided by LMFIT",
            "ExpGauss.(Area, \u03c3, \u03b3)": "Gaussian shape model with asymmetric side. The asymmetry is modelled "
                                               "using an exponential decay. Model provided by LMFIT",
            "LA (Area, \u03c3, \u03b3)": "Asymmetrical lorentzian similar to the model used in casa XPS.  \u03c3 and \u03b3 can be varied independently."
                                         "Constrained by Peak position, Area, FWHM, L/G, \u03c3 and \u03b3.  \u03c3 controlled the "
                                         "high BE side of the peak while \u03b3 controlled the low BE side of the peak." ,
            "LA (Area, \u03c3/\u03b3, \u03b3)": "Asymmetrical lorentzian with asymmetry controlled by the ratio "
                                                "between \u03c3 and \u03b3. The model is similar to the one used in casa XPS."
                                                " \u03c3 controlled the high BE side of the peak while \u03b3 "
                                                "controlled the low BE side of the peak.",
            "LA*G (Area, \u03c3/\u03b3, \u03b3)": "Asymmetrical lorentzian convoluted with a gaussian peak of a width Wg (see skew colum). "
                                           "The model is similar to the one used in casa XPS. \u03c3 controlled the "
                                         "high BE side of the peak while \u03b3 controlled the low BE side of the peak."

        }
        return descriptions.get(model, "No description available")


    def create_info_button(self, parent, tooltip):
        info_button = wx.Button(parent, label="?", size=(20, 20))
        info_button.SetToolTip(tooltip)
        # info_button.Bind(wx.EVT_ENTER_WINDOW, self.on_info_hover)
        info_button.Bind(wx.EVT_LEAVE_WINDOW, self.on_info_leave)
        return info_button

    def on_info_hover(self, event):
        button = event.GetEventObject()
        pos = button.ClientToScreen(button.GetPosition())
        self.show_info_popup(pos, button.GetToolTip().GetTip())

    def on_info_leave(self, event):
        if hasattr(self, 'info_popup'):
            self.info_popup.Destroy()

    def show_info_popup(self, pos, text):
        self.info_popup = wx.PopupWindow(self, wx.SIMPLE_BORDER)
        panel = wx.Panel(self.info_popup)
        st = wx.StaticText(panel, label=text, pos=(10, 10))
        sz = st.GetBestSize()
        self.info_popup.SetSize((sz.width + 20, sz.height + 20))
        panel.SetSize(self.info_popup.GetSize())
        self.info_popup.Position(pos, (0, 0))
        self.info_popup.Show()

    def update_background_info_button(self):
        for child in self.background_panel.GetChildren():
            if isinstance(child, wx.Button) and child.GetLabel() == "?":
                child.SetToolTip(self.get_background_description(self.parent.background_method))
                break

    def update_fitting_info_button(self):
        for child in self.fitting_panel.GetChildren():
            if isinstance(child, wx.Button) and child.GetLabel() == "?":
                child.SetToolTip(self.get_fitting_description(self.parent.selected_fitting_method))
                break

    def on_tab_change(self, event):
        selected_page = event.GetSelection()

        # Store previous state
        was_background_tab = self.parent.background_tab_selected

        # Update tab states
        self.parent.background_tab_selected = (selected_page == 0)
        self.parent.peak_fitting_tab_selected = (selected_page == 1)

        # Handle background interaction
        if self.parent.background_tab_selected:
            self.parent.enable_background_interaction()
            # Show vlines when entering background tab
            self.parent.show_hide_vlines()

            # AUTO-activate first region if regions exist
            ranges = self.get_recorded_ranges_from_data()
            if ranges:
                # Set first region as active
                self.set_active_range(0)

                # Get first region data
                offset_h, offset_l, min_range, max_range = ranges[0]

                # Update all controls with first region values
                self.updating_range_controls = True
                self._updating_offsets = True  # Prevent save_state calls from offset handlers
                self.offset_h_text.SetValue(f"{offset_h:.2f}")
                self.offset_l_text.SetValue(f"{offset_l:.2f}")
                self.min_range_text.SetValue(f"{min_range:.2f}")
                self.max_range_text.SetValue(f"{max_range:.2f}")
                self._updating_offsets = False
                self.updating_range_controls = False

                # Position vLines at first region's range
                if self.parent.vline1 is not None and self.parent.vline2 is not None:
                    min_display = self.parent.convert_energy_for_display(min_range)
                    max_display = self.parent.convert_energy_for_display(max_range)
                    self.parent.vline1.set_xdata([min_display, min_display])
                    self.parent.vline2.set_xdata([max_display, max_display])

                    # Update text labels
                    if hasattr(self.parent, 'update_vline_text_labels'):
                        self.parent.update_vline_text_labels()

                    # Update averaging indicator lines
                    if hasattr(self.parent, 'add_averaging_indicator_lines'):
                        self.parent.add_averaging_indicator_lines()

                    # Force canvas redraw
                    self.parent.canvas.draw_idle()

                # DELAY the color update to ensure UI is ready
                wx.CallAfter(self.update_range_box_colors)
                wx.CallAfter(self.force_range_box_refresh)

                # print(f"Auto-activated region 1: {min_range:.2f} - {max_range:.2f}")

        else:
            self.parent.disable_background_interaction()
            # Reset vlines when leaving background tab (same as Reset Vertical Lines button)
            if was_background_tab:
                self.on_reset_vlines(None)

        # Handle peak selection
        if not self.parent.peak_fitting_tab_selected:
            self.parent.peak_manipulation.deselect_all_peaks()

        event.Skip()


    def on_key_down_OLD(self, event):
        """Handle KEY_DOWN events for UP/DOWN arrow increment before TextCtrl processes them"""
        key_code = event.GetKeyCode()

        # Only handle UP/DOWN arrows - let everything else pass through
        if key_code not in [wx.WXK_UP, wx.WXK_DOWN]:
            event.Skip()
            return

        text_ctrl = event.GetEventObject()
        current_value = text_ctrl.GetValue()
        cursor_pos = text_ctrl.GetInsertionPoint()

        # Skip if empty or cursor at invalid position
        if not current_value or cursor_pos >= len(current_value):
            event.Skip()
            return

        try:
            # Parse the number
            original_number = float(current_value)
            is_negative = original_number < 0

            # Work with absolute value for easier processing
            abs_value = abs(original_number)

            # Split into integer and decimal parts
            str_abs = f"{abs_value:.10f}".rstrip('0').rstrip('.')
            if '.' in str_abs:
                integer_part, decimal_part = str_abs.split('.')
            else:
                integer_part = str_abs
                decimal_part = ""

            # Find which digit position cursor is on
            full_str = current_value.replace('-', '')  # Remove minus for position calculation
            adjusted_cursor = cursor_pos
            if is_negative:
                adjusted_cursor -= 1  # Account for minus sign

            # Determine digit position (power of 10)
            if '.' in full_str:
                decimal_pos = full_str.index('.')
                if adjusted_cursor < decimal_pos:
                    # Integer part - modify character to the RIGHT of cursor
                    # Increment adjusted_cursor to get the character we want to modify
                    target_char_pos = adjusted_cursor
                    digit_power = (decimal_pos - 1) - target_char_pos
                else:
                    # Decimal part - modify character to the RIGHT of cursor
                    target_char_pos = adjusted_cursor
                    digit_power = decimal_pos - target_char_pos
            else:
                # No decimal point - modify character to the RIGHT of cursor
                target_char_pos = adjusted_cursor
                digit_power = (len(full_str) - 1) - target_char_pos

            # Skip if we're at the end or on decimal point
            if adjusted_cursor >= len(full_str) or (
                    adjusted_cursor < len(full_str) and full_str[adjusted_cursor] == '.'):
                event.Skip()
                return

            # Calculate increment value
            increment_value = 10 ** digit_power
            if key_code == wx.WXK_DOWN:
                increment_value = -increment_value

            # Apply increment
            new_number = original_number + increment_value

            # Format the result
            if digit_power >= 0:
                # Integer digit changed
                if abs(new_number - round(new_number)) < 1e-10:
                    formatted_result = f"{int(round(new_number))}"
                else:
                    decimal_places = len(decimal_part) if decimal_part else 0
                    formatted_result = f"{new_number:.{decimal_places}f}"
            else:
                # Decimal digit changed
                required_decimals = abs(digit_power)
                existing_decimals = len(decimal_part) if decimal_part else 0
                decimal_places = max(required_decimals, existing_decimals)
                formatted_result = f"{new_number:.{decimal_places}f}".rstrip('0').rstrip('.')

            # Update the text control
            text_ctrl.SetValue(formatted_result)

            # Calculate desired cursor position BEFORE calling change event
            new_length = len(formatted_result)
            old_length = len(current_value)
            desired_cursor_pos = min(cursor_pos + (new_length - old_length), new_length)

            # Trigger the appropriate change event manually
            if text_ctrl == self.offset_h_text:
                self.on_offset_h_change(event)
            elif text_ctrl == self.offset_l_text:
                self.on_offset_l_change(event)
            elif text_ctrl == self.min_range_text:
                self.on_min_range_change(event)
                # RESTORE cursor position after change event (which may call SetValue again)
                wx.CallAfter(lambda: text_ctrl.SetInsertionPoint(desired_cursor_pos))
            elif text_ctrl == self.max_range_text:
                self.on_max_range_change(event)
                # RESTORE cursor position after change event (which may call SetValue again)
                wx.CallAfter(lambda: text_ctrl.SetInsertionPoint(desired_cursor_pos))

            # For offset controls, set cursor position immediately since they don't interfere
            if text_ctrl in [self.offset_h_text, self.offset_l_text]:
                text_ctrl.SetInsertionPoint(desired_cursor_pos)

            # DO NOT call event.Skip() - we handled this completely

        except (ValueError, IndexError):
            # If parsing fails, allow default behavior
            event.Skip()
            return

    def on_key_down(self, event):
        """Handle KEY_DOWN events for UP/DOWN arrow increment before TextCtrl processes them"""
        key_code = event.GetKeyCode()

        # Only handle UP/DOWN arrows - let everything else pass through
        if key_code not in [wx.WXK_UP, wx.WXK_DOWN]:
            event.Skip()
            return

        text_ctrl = event.GetEventObject()
        current_value = text_ctrl.GetValue()
        cursor_pos = text_ctrl.GetInsertionPoint()

        # Skip if empty or cursor at invalid position
        if not current_value or cursor_pos >= len(current_value):
            event.Skip()
            return

        try:
            # Parse the number
            original_number = float(current_value)
            is_negative = original_number < 0

            # Work with absolute value for easier processing
            abs_value = abs(original_number)

            # Split into integer and decimal parts
            str_abs = f"{abs_value:.10f}".rstrip('0').rstrip('.')
            if '.' in str_abs:
                integer_part, decimal_part = str_abs.split('.')
            else:
                integer_part = str_abs
                decimal_part = ""

            # Find which digit position cursor is on
            full_str = current_value.replace('-', '')  # Remove minus for position calculation
            adjusted_cursor = cursor_pos
            if is_negative:
                adjusted_cursor -= 1  # Account for minus sign

            # Determine digit position (power of 10)
            if '.' in full_str:
                decimal_pos = full_str.index('.')
                if adjusted_cursor < decimal_pos:
                    # Integer part - modify character to the RIGHT of cursor
                    target_char_pos = adjusted_cursor
                    digit_power = (decimal_pos - 1) - target_char_pos
                else:
                    # Decimal part - modify character to the RIGHT of cursor
                    target_char_pos = adjusted_cursor
                    digit_power = decimal_pos - target_char_pos
            else:
                # No decimal point - modify character to the RIGHT of cursor
                target_char_pos = adjusted_cursor
                digit_power = (len(full_str) - 1) - target_char_pos

            # Skip if we're at the end or on decimal point
            if adjusted_cursor >= len(full_str) or (
                    adjusted_cursor < len(full_str) and full_str[adjusted_cursor] == '.'):
                event.Skip()
                return

            # Calculate increment value
            increment_value = 10 ** digit_power
            if key_code == wx.WXK_DOWN:
                increment_value = -increment_value

            # Apply increment
            new_number = original_number + increment_value

            # CONSISTENT FORMATTING: Maintain same decimal places as original
            if '.' in current_value:
                # Has decimal - maintain same number of decimal places
                original_decimal_places = len(current_value.split('.')[1])
                formatted_result = f"{new_number:.{original_decimal_places}f}"
            else:
                # No decimal - format as integer
                formatted_result = f"{int(round(new_number))}"

            # Update the text control
            text_ctrl.SetValue(formatted_result)

            # Since string length stays the same, cursor position stays the same
            desired_cursor_pos = cursor_pos

            # Trigger the appropriate change event manually
            if text_ctrl == self.offset_h_text:
                self.on_offset_h_change(event)
            elif text_ctrl == self.offset_l_text:
                self.on_offset_l_change(event)
            elif text_ctrl == self.min_range_text:
                self.on_min_range_change(event)
                # RESTORE cursor position after change event (which may call SetValue again)
                wx.CallAfter(lambda: text_ctrl.SetInsertionPoint(desired_cursor_pos))
            elif text_ctrl == self.max_range_text:
                self.on_max_range_change(event)
                # RESTORE cursor position after change event (which may call SetValue again)
                wx.CallAfter(lambda: text_ctrl.SetInsertionPoint(desired_cursor_pos))

            # For offset controls, set cursor position immediately since they don't interfere
            if text_ctrl in [self.offset_h_text, self.offset_l_text]:
                text_ctrl.SetInsertionPoint(desired_cursor_pos)

            # DO NOT call event.Skip() - we handled this completely

        except (ValueError, IndexError):
            # If parsing fails, allow default behavior
            event.Skip()
            return

    def validate_numeric_input(self, event):
        """Validate input to allow only numbers, single dot, and minus sign"""
        key_code = event.GetKeyCode()

        # Always allow these control keys
        if key_code in [wx.WXK_BACK, wx.WXK_DELETE, wx.WXK_TAB, wx.WXK_RETURN,
                        wx.WXK_LEFT, wx.WXK_RIGHT, wx.WXK_HOME, wx.WXK_END]:
            event.Skip()
            return

        # Get the current text control and its current value
        text_ctrl = event.GetEventObject()
        current_value = text_ctrl.GetValue()
        insertion_point = text_ctrl.GetInsertionPoint()

        # Allow digits 0-9
        if 48 <= key_code <= 57:  # ASCII codes for '0' to '9'
            event.Skip()
            return

        # Allow minus sign only at the beginning for offset controls
        if key_code == 45 and insertion_point == 0:  # ASCII code for '-'
            if text_ctrl in [self.offset_h_text, self.offset_l_text]:
                if '-' not in current_value:
                    event.Skip()
                    return

        # Allow dot only if there isn't one already
        if key_code == 46:  # ASCII code for '.'
            if '.' not in current_value:
                event.Skip()
                return

        # Block all other characters
        return

    def handle_digit_increment(self, event):
        """Handle UP/DOWN arrow keys to increment/decrement digit at cursor position"""
        key_code = event.GetKeyCode()

        # Only handle UP and DOWN arrow keys
        if key_code not in [wx.WXK_UP, wx.WXK_DOWN]:
            event.Skip()
            return

        text_ctrl = event.GetEventObject()
        current_value = text_ctrl.GetValue()
        cursor_pos = text_ctrl.GetInsertionPoint()

        # Skip if empty or cursor at invalid position
        if not current_value or cursor_pos >= len(current_value):
            event.Skip()
            return

        try:
            # Parse the number
            original_number = float(current_value)
            is_negative = original_number < 0

            # Work with absolute value for easier processing
            abs_value = abs(original_number)

            # Split into integer and decimal parts
            str_abs = f"{abs_value:.10f}".rstrip('0').rstrip('.')
            if '.' in str_abs:
                integer_part, decimal_part = str_abs.split('.')
            else:
                integer_part = str_abs
                decimal_part = ""

            # Find which digit position cursor is on
            full_str = current_value.replace('-', '')  # Remove minus for position calculation
            adjusted_cursor = cursor_pos
            if is_negative:
                adjusted_cursor -= 1  # Account for minus sign

            # Skip if cursor is on decimal point
            if adjusted_cursor < len(full_str) and full_str[adjusted_cursor] == '.':
                event.Skip()
                return

            # Determine digit position (power of 10)
            if '.' in full_str:
                decimal_pos = full_str.index('.')
                if adjusted_cursor < decimal_pos:
                    # Integer part - power is positive
                    digit_power = decimal_pos - adjusted_cursor - 1
                else:
                    # Decimal part - power is negative
                    digit_power = decimal_pos - adjusted_cursor
            else:
                # No decimal point
                digit_power = len(full_str) - adjusted_cursor - 1

            # Calculate increment value
            increment_value = 10 ** digit_power
            if key_code == wx.WXK_DOWN:
                increment_value = -increment_value

            # Apply increment
            new_number = original_number + increment_value

            # Format the result to preserve appropriate decimal places
            if digit_power >= 0:
                # Integer digit changed - format as integer if no decimals needed
                if abs(new_number - round(new_number)) < 1e-10:
                    formatted_result = f"{int(round(new_number))}"
                else:
                    # Keep existing decimal places
                    decimal_places = len(decimal_part) if decimal_part else 0
                    formatted_result = f"{new_number:.{decimal_places}f}"
            else:
                # Decimal digit changed - ensure we have enough decimal places
                required_decimals = abs(digit_power)
                existing_decimals = len(decimal_part) if decimal_part else 0
                decimal_places = max(required_decimals, existing_decimals)
                formatted_result = f"{new_number:.{decimal_places}f}".rstrip('0').rstrip('.')

            # Update the text control
            text_ctrl.SetValue(formatted_result)

            # Restore cursor position (adjust for possible length change)
            new_length = len(formatted_result)
            old_length = len(current_value)
            new_cursor_pos = min(cursor_pos + (new_length - old_length), new_length)
            text_ctrl.SetInsertionPoint(new_cursor_pos)

            # Trigger the appropriate change event
            if text_ctrl == self.offset_h_text:
                self.on_offset_h_change(event)
            elif text_ctrl == self.offset_l_text:
                self.on_offset_l_change(event)
            elif text_ctrl == self.min_range_text:
                self.on_min_range_change(event)
            elif text_ctrl == self.max_range_text:
                self.on_max_range_change(event)

        except (ValueError, IndexError):
            # If parsing fails, just skip
            event.Skip()
            return

    def force_range_box_refresh(self):
        """Force refresh of range box colors and layout"""
        try:
            # Double-check active range is set
            # print(f"Active range index: {getattr(self, 'active_range_index', 'NOT SET')}")

            # Force refresh of each range box
            for i, box in enumerate(self.range_boxes):
                if i == self.active_range_index:
                    box.SetBackgroundColour(wx.Colour(255, 0, 0))  # Red
                    box.SetForegroundColour(wx.Colour(255, 255, 255))  # White text
                    # print(f"Setting box {i + 1} to RED (active)")
                else:
                    box.SetBackgroundColour(wx.NullColour)
                    box.SetForegroundColour(wx.NullColour)
                    # print(f"Setting box {i + 1} to DEFAULT")

                box.Refresh()
                box.Update()

            # Force layout refresh
            self.range_boxes_panel.Layout()
            self.range_boxes_panel.Refresh()

        except Exception as e:
            print(f"Error in force_range_box_refresh: {e}")

    def load_active_range_to_controls(self):
        """Load active range settings into the UI controls"""
        if not hasattr(self, 'active_range_index') or self.active_range_index < 0:
            return

        ranges = self.get_recorded_ranges_from_data()
        if self.active_range_index < len(ranges):
            offset_h, offset_l, min_range, max_range = ranges[self.active_range_index]

            # Update controls with active range values
            self.updating_range_controls = True
            self.offset_h_text.SetValue(f"{offset_h:.2f}")
            self.offset_l_text.SetValue(f"{offset_l:.2f}")
            self.min_range_text.SetValue(f"{min_range:.2f}")
            self.max_range_text.SetValue(f"{max_range:.2f}")
            self.updating_range_controls = False

            # Update vlines to match active region positions
            if self.parent.vline1 is not None:
                min_display = self.parent.convert_energy_for_display(min_range)
                self.parent.vline1.set_xdata([min_display, min_display])
            if self.parent.vline2 is not None:
                max_display = self.parent.convert_energy_for_display(max_range)
                self.parent.vline2.set_xdata([max_display, max_display])

            # Update text labels and redraw
            if hasattr(self.parent, 'update_vline_text_labels'):
                self.parent.update_vline_text_labels()

            # Update averaging indicator lines
            if hasattr(self.parent, 'add_averaging_indicator_lines'):
                self.parent.add_averaging_indicator_lines()

            self.parent.canvas.draw_idle()

            print(f"Loaded region {self.active_range_index + 1} as active")

    def on_close(self, event):
        # Store previous state
        was_background_tab = self.parent.background_tab_selected

        # Cancel any pending wx.CallAfter calls by setting a flag
        self._is_closing = True

        # Reset vlines when closing (same as Reset Vertical Lines button)
        self.on_reset_vlines(None)

        self.parent.background_tab_selected = False
        self.parent.peak_fitting_tab_selected = False

        # self.parent.show_hide_vlines()
        # self.parent.peak_manipulation.deselect_all_peaks()
        self.parent.disable_background_interaction()
        # Reset vlines when leaving background tab (same as Reset Vertical Lines button)
        if was_background_tab:
            self.on_reset_vlines(None)

        # Small delay to ensure all pending operations complete
        wx.CallLater(10, self.Destroy)

    def parse_cross_section(self, text):
        try:
            values = [float(x.strip()) for x in text.split(',')]
            if len(values) == 4:
                sheet_name = self.parent.sheet_combobox.GetValue()
                self.parent.Data['Core levels'][sheet_name]['Background'].update({
                    'Tougaard_B': values[0],
                    'Tougaard_C': values[1],
                    'Tougaard_D': values[2],
                    'Tougaard_T0': values[3]
                })
                return values
            raise ValueError
        except ValueError:
            return [2866, 1643, 1, 0]

    def on_cross_section_change(self, event):
        self.parse_cross_section(self.cross_section.GetValue())
        # self.parent.plot_manager.plot_background(self.parent)

    def initialize_recorded_ranges(self):
        """Initialize recorded ranges from window.Data when tab becomes active"""
        self.update_range_boxes()
        # Also update window.Data background range
        if self.get_recorded_ranges_from_data():  # Only if we have recorded ranges
            self.update_window_data_background_range()

    def record_current_range(self):
        """Record current background range settings"""
        try:
            offset_h = float(self.offset_h_text.GetValue())
            offset_l = float(self.offset_l_text.GetValue())
            min_range = float(self.min_range_text.GetValue())
            max_range = float(self.max_range_text.GetValue())

            range_tuple = (offset_h, offset_l, min_range, max_range)

            # Get current ranges from data
            current_ranges = self.get_recorded_ranges_from_data()

            # REMOVE THIS DUPLICATE CHECK - Always add the range
            # OLD CODE: if not current_ranges or current_ranges[-1] != range_tuple:
            current_ranges.append(range_tuple)

            # Save back to data
            sheet_name = self.parent.sheet_combobox.GetValue()
            if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
                self.parent.Data['Core levels'][sheet_name]['Background'] = {}
            self.parent.Data['Core levels'][sheet_name]['Background']['Recorded_Ranges'] = current_ranges

            # Set the new range as active
            self.set_active_range(len(current_ranges) - 1)

            # Update window.Data background range with overall range
            self.update_window_data_background_range()

            self.update_range_boxes()

        except ValueError:
            pass  # Invalid values, don't record

    def update_active_range_offsets(self, new_offset_h, new_offset_l):
        """Update the offsets for the currently active range"""
        if self.active_range_index >= 0:
            ranges = self.get_recorded_ranges_from_data()
            if self.active_range_index < len(ranges):
                # Get current range data
                old_offset_h, old_offset_l, min_range, max_range = ranges[self.active_range_index]

                # Update with new offsets
                ranges[self.active_range_index] = (new_offset_h, new_offset_l, min_range, max_range)

                # Save back to data
                sheet_name = self.parent.sheet_combobox.GetValue()
                if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
                    self.parent.Data['Core levels'][sheet_name]['Background'] = {}
                self.parent.Data['Core levels'][sheet_name]['Background']['Recorded_Ranges'] = ranges

                # print(
                #     f"Updated active range {self.active_range_index + 1}: Offset H={new_offset_h:.1f}, L={new_offset_l:.1f}")


    def update_range_boxes(self):
        """Update the numbered range selection boxes with wrapping after 4"""
        # Clear existing boxes
        for box in self.range_boxes:
            box.Destroy()
        self.range_boxes.clear()

        # Clear both sizers
        self.range_boxes_sizer.Clear()
        self.range_boxes_sizer2.Clear()

        # Get ranges from data
        ranges = self.get_recorded_ranges_from_data()

        # Create new boxes with wrapping after 4
        for i, range_data in enumerate(ranges):
            box = wx.Button(self.range_boxes_panel, label=str(i + 1), size=(25, 25))
            box.Bind(wx.EVT_BUTTON, lambda evt, idx=i: self.on_range_box_click(idx))
            self.range_boxes.append(box)

            # Add to appropriate row
            if i < 4:
                # First row (boxes 1-4)
                self.range_boxes_sizer.Add(box, 0, wx.ALL, 2)
            else:
                # Second row (boxes 5+)
                self.range_boxes_sizer2.Add(box, 0, wx.ALL, 2)

        # Layout the panel
        self.range_boxes_panel.Layout()

    def on_range_box_click(self, index):
        """Handle clicking on a numbered range box"""
        ranges = self.get_recorded_ranges_from_data()
        if index < len(ranges):
            offset_h, offset_l, min_range, max_range = ranges[index]

            # Set this as the active range
            self.set_active_range(index)

            # Update the controls with flags to prevent unwanted events
            self.updating_range_controls = True
            self._updating_offsets = True
            self.offset_h_text.SetValue(str(offset_h))
            self.offset_l_text.SetValue(str(offset_l))
            self.min_range_text.SetValue(str(min_range))
            self.max_range_text.SetValue(str(max_range))
            self._updating_offsets = False
            self.updating_range_controls = False

            # Move vLines to the selected region's positions
            if self.parent.vline1 is not None:
                min_display = self.parent.convert_energy_for_display(min_range)
                self.parent.vline1.set_xdata([min_display, min_display])
            if self.parent.vline2 is not None:
                max_display = self.parent.convert_energy_for_display(max_range)
                self.parent.vline2.set_xdata([max_display, max_display])

            # Update text labels
            if hasattr(self.parent, 'update_vline_text_labels'):
                self.parent.update_vline_text_labels()

            # Update averaging indicator lines
            if hasattr(self.parent, 'add_averaging_indicator_lines'):
                self.parent.add_averaging_indicator_lines()

            # Force canvas redraw
            self.parent.canvas.draw_idle()

            # Update window.Data background range with overall range
            self.update_window_data_background_range()

            # Update visual highlighting
            wx.CallAfter(self.update_range_box_colors)
            wx.CallAfter(self.force_range_box_refresh)

            print(f"Selected region {index + 1}: {min_range:.2f} - {max_range:.2f}")

    def save_recorded_ranges_to_data(self):
        """Save recorded ranges to window.Data"""
        sheet_name = self.parent.sheet_combobox.GetValue()
        if sheet_name in self.parent.Data['Core levels']:
            if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
                self.parent.Data['Core levels'][sheet_name]['Background'] = {}

            ranges = self.get_recorded_ranges_from_data()
            self.parent.Data['Core levels'][sheet_name]['Background']['Recorded_Ranges'] = ranges

    def get_recorded_ranges_from_data(self):
        """Get recorded ranges from window.Data"""
        sheet_name = self.parent.sheet_combobox.GetValue()
        if (sheet_name in self.parent.Data['Core levels'] and
                'Background' in self.parent.Data['Core levels'][sheet_name] and
                'Recorded_Ranges' in self.parent.Data['Core levels'][sheet_name]['Background']):
            return self.parent.Data['Core levels'][sheet_name]['Background']['Recorded_Ranges']
        return []

    def clear_recorded_ranges_from_data(self):
        """Clear recorded ranges from window.Data"""
        sheet_name = self.parent.sheet_combobox.GetValue()
        if (sheet_name in self.parent.Data['Core levels'] and
                'Background' in self.parent.Data['Core levels'][sheet_name]):
            self.parent.Data['Core levels'][sheet_name]['Background']['Recorded_Ranges'] = []

    def get_overall_background_range(self):
        """Get the overall minimum and maximum values from all recorded background ranges"""
        ranges = self.get_recorded_ranges_from_data()

        if not ranges:
            # No recorded ranges, fall back to current vline positions or default
            if (hasattr(self.parent, 'vline1') and self.parent.vline1 is not None and
                    hasattr(self.parent, 'vline2') and self.parent.vline2 is not None):
                vline1_pos = self.parent.vline1.get_xdata()[0]
                vline2_pos = self.parent.vline2.get_xdata()[0]
                return min(vline1_pos, vline2_pos), max(vline1_pos, vline2_pos)
            else:
                return self.parent.bg_min_energy, self.parent.bg_max_energy

        # Extract min and max ranges from all recorded ranges
        all_mins = []
        all_maxs = []

        for offset_h, offset_l, min_range, max_range in ranges:
            all_mins.append(min_range)
            all_maxs.append(max_range)

        overall_min = min(all_mins)
        overall_max = max(all_maxs)

        # print(f"Overall background range from {len(ranges)} recorded ranges: {overall_min:.2f} - {overall_max:.2f}")

        return overall_min, overall_max

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

    def on_clear_between_vlines(self, event):
        sheet_name = self.parent.sheet_combobox.GetValue()
        if sheet_name in self.parent.Data['Core levels']:
            core_level_data = self.parent.Data['Core levels'][sheet_name]
            if 'Background' in core_level_data:
                bg_low = core_level_data['Background']['Bkg Low']
                bg_high = core_level_data['Background']['Bkg High']
                if bg_low is not None and bg_high is not None:
                    x_values = np.array(self.parent.x_values)
                    raw_data = np.array(core_level_data['Raw Data'])
                    background = np.array(core_level_data['Background']['Bkg Y'])

                    mask = (x_values >= min(bg_low, bg_high)) & (x_values <= max(bg_low, bg_high))
                    background[mask] = raw_data[mask]

                    core_level_data['Background']['Bkg Y'] = background.tolist()
                    self.parent.background = background
                    self.parent.clear_and_replot()

            # Use tab switching trick to properly recreate vlines at new positions
            if self.parent.background_tab_selected:
                # Find the notebook
                notebook = None
                for child in self.GetChildren():
                    for grandchild in child.GetChildren():
                        if isinstance(grandchild, wx.Notebook):
                            notebook = grandchild
                            break
                    if notebook:
                        break

                if notebook:
                    wx.CallAfter(self._switch_tabs_trick, notebook)

    def on_min_range_change(self, event):
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
                # Convert BE value to display position for vLine placement
                display_pos = self.parent.convert_energy_for_display(new_min)
                self.parent.vline1.set_xdata([display_pos, display_pos])

                # Update data structure
                sheet_name = self.parent.sheet_combobox.GetValue()
                if sheet_name in self.parent.Data['Core levels']:
                    if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
                        self.parent.Data['Core levels'][sheet_name]['Background'] = {}
                    self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Low'] = float(new_min)

                # Update text labels
                if hasattr(self.parent, 'update_vline_text_labels'):
                    self.parent.update_vline_text_labels()

                # Update averaging indicator lines
                if hasattr(self.parent, 'add_averaging_indicator_lines'):
                    self.parent.add_averaging_indicator_lines()

                self.parent.canvas.draw_idle()

            # Update active range positions if one is selected
            if hasattr(self, 'active_range_index') and self.active_range_index >= 0:
                if hasattr(self.parent, 'mouse_handler') and hasattr(self.parent.mouse_handler,
                                                                     'update_active_region_positions'):
                    self.parent.mouse_handler.update_active_region_positions()

            # Redraw background from all regions
            if hasattr(self.parent, 'mouse_handler') and hasattr(self.parent.mouse_handler,
                                                                 'redraw_all_regions_background'):
                self.parent.mouse_handler.redraw_all_regions_background()

        except ValueError:
            pass

    def on_max_range_change_OLD_BE(self, event):
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

                # Update data structure
                sheet_name = self.parent.sheet_combobox.GetValue()
                if sheet_name in self.parent.Data['Core levels']:
                    if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
                        self.parent.Data['Core levels'][sheet_name]['Background'] = {}
                    self.parent.Data['Core levels'][sheet_name]['Background']['Bkg High'] = float(new_max)

                # Update text labels
                if hasattr(self.parent, 'update_vline_text_labels'):
                    self.parent.update_vline_text_labels()

                # Update averaging indicator lines
                if hasattr(self.parent, 'add_averaging_indicator_lines'):
                    self.parent.add_averaging_indicator_lines()

                self.parent.canvas.draw_idle()

            # Update active range positions if one is selected
            if hasattr(self, 'active_range_index') and self.active_range_index >= 0:
                if hasattr(self.parent, 'mouse_handler') and hasattr(self.parent.mouse_handler,
                                                                     'update_active_region_positions'):
                    self.parent.mouse_handler.update_active_region_positions()

            # Redraw background from all regions
            if hasattr(self.parent, 'mouse_handler') and hasattr(self.parent.mouse_handler,
                                                                 'redraw_all_regions_background'):
                self.parent.mouse_handler.redraw_all_regions_background()

        except ValueError:
            pass

    def on_max_range_change(self, event):
        if self.updating_range_controls:
            return

        try:
            # Get display values from text controls
            new_max_display = float(self.max_range_text.GetValue())
            min_display = float(self.min_range_text.GetValue())

            # Convert display values to BE for storage and comparison
            new_max_be = self.parent.convert_energy_from_display(new_max_display)
            min_be = self.parent.convert_energy_from_display(min_display)

            # Handle order swapping in BE coordinates
            if new_max_be < min_be:
                self.updating_range_controls = True
                self.max_range_text.SetValue(f"{min_display:.2f}")
                self.min_range_text.SetValue(f"{new_max_display:.2f}")
                self.updating_range_controls = False
                new_max_be = min_be

            # Update vLine position using display coordinates
            if self.parent.vline2 is not None:
                self.parent.vline2.set_xdata([new_max_display, new_max_display])

                # Store BE value in data structure
                sheet_name = self.parent.sheet_combobox.GetValue()
                if sheet_name in self.parent.Data['Core levels']:
                    if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
                        self.parent.Data['Core levels'][sheet_name]['Background'] = {}
                    self.parent.Data['Core levels'][sheet_name]['Background']['Bkg High'] = float(new_max_be)

                # Update text labels and redraw
                if hasattr(self.parent, 'update_vline_text_labels'):
                    self.parent.update_vline_text_labels()

                # Update averaging indicator lines
                if hasattr(self.parent, 'add_averaging_indicator_lines'):
                    self.parent.add_averaging_indicator_lines()
                self.parent.canvas.draw_idle()

            # Update active range positions and redraw background
            if hasattr(self, 'active_range_index') and self.active_range_index >= 0:
                if hasattr(self.parent, 'mouse_handler'):
                    self.parent.mouse_handler.update_active_region_positions()
                    self.parent.mouse_handler.redraw_all_regions_background()

        except ValueError:
            pass

    def update_window_data_background_range(self):
        """Update window.Data background range with overall range from all recorded ranges"""
        sheet_name = self.parent.sheet_combobox.GetValue()

        if sheet_name in self.parent.Data['Core levels']:
            # Get overall range
            overall_min, overall_max = self.get_overall_background_range()

            # Update window.Data
            if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
                self.parent.Data['Core levels'][sheet_name]['Background'] = {}

            # Update with overall range
            self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Low'] = overall_min
            self.parent.Data['Core levels'][sheet_name]['Background']['Bkg High'] = overall_max

            # Also update the window attributes
            self.parent.bg_min_energy = overall_min
            self.parent.bg_max_energy = overall_max

            # print(f"Updated window.Data background range: {overall_min:.2f} - {overall_max:.2f}")

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
                # Get vline display positions
                vline1_display = self.parent.vline1.get_xdata()[0]
                vline2_display = self.parent.vline2.get_xdata()[0]

                # The vlines are already positioned in display coordinates,
                # so these values are correct for the text controls
                actual_min = min(vline1_display, vline2_display)
                actual_max = max(vline1_display, vline2_display)
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

    def on_remove_active_region(self, event):
        """Remove the currently active region"""
        if not hasattr(self, 'active_range_index') or self.active_range_index < 0:
            wx.MessageBox("No active region selected to remove.", "No Active Region",
                          wx.OK | wx.ICON_INFORMATION)
            return

        # Get current ranges
        ranges = self.get_recorded_ranges_from_data()
        if self.active_range_index >= len(ranges):
            return

        # Store which region we're removing for the print message
        removed_region_number = self.active_range_index + 1

        # STEP 1: Remove the region data from window.data
        ranges.pop(self.active_range_index)

        # Save back to data
        sheet_name = self.parent.sheet_combobox.GetValue()
        if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
            self.parent.Data['Core levels'][sheet_name]['Background'] = {}
        self.parent.Data['Core levels'][sheet_name]['Background']['Recorded_Ranges'] = ranges

        # STEP 2: Update range boxes (removes the button)
        self.update_range_boxes()

        # STEP 3: Set region 1 as active if it exists
        if ranges:
            # Set region 1 (index 0) as active
            self.set_active_range(0)

            # Load region 1 values to controls
            offset_h, offset_l, min_range, max_range = ranges[0]
            self.updating_range_controls = True
            self.offset_h_text.SetValue(f"{offset_h:.2f}")
            self.offset_l_text.SetValue(f"{offset_l:.2f}")
            self.min_range_text.SetValue(f"{min_range:.2f}")
            self.max_range_text.SetValue(f"{max_range:.2f}")
            self.updating_range_controls = False

            # Update vlines to region 1 position
            if self.parent.vline1 is not None and self.parent.vline2 is not None:
                min_display = self.parent.convert_energy_for_display(min_range)
                max_display = self.parent.convert_energy_for_display(max_range)
                self.parent.vline1.set_xdata([min_display, min_display])
                self.parent.vline2.set_xdata([max_display, max_display])

                if hasattr(self.parent, 'update_vline_text_labels'):
                    self.parent.update_vline_text_labels()

                # Update averaging indicator lines
                if hasattr(self.parent, 'add_averaging_indicator_lines'):
                    self.parent.add_averaging_indicator_lines()

            # Update window data background range
            self.update_window_data_background_range()

            # Force visual refresh to show region 1 as active
            wx.CallAfter(self.update_range_box_colors)
            wx.CallAfter(self.force_range_box_refresh)

        else:
            # No regions left
            self.clear_active_range()

            # Clear window data background
            self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Low'] = None
            self.parent.Data['Core levels'][sheet_name]['Background']['Bkg High'] = None
            self.parent.bg_min_energy = None
            self.parent.bg_max_energy = None

            # Clear the UI controls
            self.updating_range_controls = True
            self.offset_h_text.SetValue("0.0")
            self.offset_l_text.SetValue("0.0")
            self.min_range_text.SetValue("0.00")
            self.max_range_text.SetValue("0.00")
            self.updating_range_controls = False

        # STEP 4: Remove whole background and redraw from region 1 to xxx
        if ranges:
            # Redraw background from all remaining regions in sequence
            if hasattr(self.parent, 'mouse_handler') and hasattr(self.parent.mouse_handler,
                                                                 'redraw_all_regions_background'):
                self.parent.mouse_handler.redraw_all_regions_background()
            else:
                # Fallback: plot background normally
                self.parent.plot_manager.plot_background(self.parent)
        else:
            # STEP 5: No regions - don't create any background, just plot raw data
            self.parent.plot_manager.clear_background(self.parent)
            self.parent.plot_data()

        save_state(self.parent)
        print(f"Removed region {removed_region_number}")

    def set_offset_display_value(self, text_ctrl, actual_value):
        """Set display value as positive while keeping actual value negative"""
        display_value = abs(float(actual_value))
        text_ctrl.SetValue(f"{display_value:.1f}")

    def get_offset_actual_value(self, text_ctrl):
        """Get actual negative value from positive display value"""
        display_value = float(text_ctrl.GetValue())
        return -abs(display_value)  # Always return negative

    def disable_fitting_ui(self):
        """Disable all UI controls during multiple fitting"""
        # Fitting panel controls
        self.model_combobox.Enable(False)
        self.optimization_method.Enable(False)
        self.max_iter_spin.Enable(False)
        self.fit_iterations_spin.Enable(False)
        self.weights_combo.Enable(False)

        # Buttons - find them by iterating through fitting_panel children
        for child in self.fitting_panel.GetChildren():
            if isinstance(child, wx.Button):
                if child.GetLabel() != "Fit \nN# Times":  # Keep multi-fit button enabled to show it's running
                    child.Enable(False)

        # Background panel controls
        self.method_combobox.Enable(False)
        self.offset_h_text.Enable(False)
        self.offset_l_text.Enable(False)
        self.averaging_points_text.Enable(False)
        self.smooth_data_checkbox.Enable(False)

        # Background buttons
        for child in self.background_panel.GetChildren():
            if isinstance(child, wx.Button):
                child.Enable(False)

        # Find and update the multi-fit button
        for child in self.fitting_panel.GetChildren():
            if isinstance(child, wx.Button) and "N# Times" in child.GetLabel():
                child.SetLabel("Fitting...")
                break

    def enable_fitting_ui(self):
        """Re-enable all UI controls after multiple fitting"""
        # Fitting panel controls
        self.model_combobox.Enable(True)
        self.optimization_method.Enable(True)
        self.max_iter_spin.Enable(True)
        self.fit_iterations_spin.Enable(True)
        self.weights_combo.Enable(True)

        # Buttons
        for child in self.fitting_panel.GetChildren():
            if isinstance(child, wx.Button):
                child.Enable(True)

        # Background panel controls
        self.method_combobox.Enable(True)
        self.offset_h_text.Enable(True)
        self.offset_l_text.Enable(True)
        self.averaging_points_text.Enable(True)
        self.smooth_data_checkbox.Enable(True)

        # Background buttons - restore based on method
        for child in self.background_panel.GetChildren():
            if isinstance(child, wx.Button):
                child.Enable(True)

        # Restore Tougaard controls visibility based on current method
        current_method = self.method_combobox.GetValue()
        self.update_tougaard_controls_visibility(current_method)

        # Restore multi-fit button label
        for child in self.fitting_panel.GetChildren():
            if isinstance(child, wx.Button) and "Fitting..." in child.GetLabel():
                child.SetLabel("Fit \nN# Times")
                break

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


class TougaardFitWindow(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, title="Tougaard Fit", size=(1050, 600))

        self.parent = parent
        self.panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)

        main_pos = parent.parent.GetPosition()
        main_size = parent.parent.GetSize()
        tougaard_size = self.GetSize()

        x = main_pos.x + (main_size.width - tougaard_size.width) // 2
        y = main_pos.y + (main_size.height - tougaard_size.height) // 2

        self.SetPosition((x, y))

        # Get data from parent window
        sheet_name = self.parent.parent.sheet_combobox.GetValue()
        self.x_values = np.array(self.parent.parent.Data['Core levels'][sheet_name]['B.E.'])
        self.y_values = np.array(self.parent.parent.Data['Core levels'][sheet_name]['Raw Data'])

        min_x = min(self.x_values)
        max_x = max(self.x_values)

        # Left panel for controls
        control_panel = wx.Panel(self.panel)
        control_sizer = wx.BoxSizer(wx.VERTICAL)

        # # Number of Tougaard backgrounds control
        # num_box = wx.StaticBox(control_panel, label="Number of Tougaard Backgrounds")
        # num_sizer = wx.StaticBoxSizer(num_box, wx.HORIZONTAL)
        # self.num_tougaard = wx.SpinCtrl(control_panel, min=1, max=10, initial=1)
        # num_sizer.Add(self.num_tougaard, 1, wx.ALL, 5)

        # Background start control
        bg_box = wx.StaticBox(control_panel, label="Background Start")
        bg_sizer = wx.StaticBoxSizer(bg_box, wx.HORIZONTAL)
        self.bg_start = wx.SpinCtrlDouble(control_panel, min=0, max=2000, inc=0.1, value=str(min_x + 1))
        bg_sizer.Add(self.bg_start, 1, wx.ALL, 5)

        # Fitting ranges controls
        range_box = wx.StaticBox(control_panel, label="Fitting Ranges")
        range_sizer = wx.StaticBoxSizer(range_box, wx.VERTICAL)

        # Number of ranges control
        num_ranges_sizer = wx.BoxSizer(wx.HORIZONTAL)
        num_ranges_sizer.Add(wx.StaticText(range_box, label="Number of Ranges:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.num_ranges = wx.SpinCtrl(range_box, min=1, max=10, initial=1)
        self.num_ranges.Bind(wx.EVT_SPINCTRL, self.on_num_ranges_change)
        num_ranges_sizer.Add(self.num_ranges, 1, wx.ALL, 5)
        range_sizer.Add(num_ranges_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # Ranges scroll area
        self.ranges_panel = wx.ScrolledWindow(range_box)
        self.ranges_panel.SetScrollRate(0, 20)
        self.ranges_panel.SetMinSize((-1, 180))  # Increased from 120 to 180
        self.ranges_panel_sizer = wx.BoxSizer(wx.VERTICAL)
        self.ranges_panel.SetSizer(self.ranges_panel_sizer)
        range_sizer.Add(self.ranges_panel, 1, wx.EXPAND | wx.ALL, 5)

        # Initialize ranges array
        self.ranges = []

        # Plot setup - MUST be before create_initial_ranges
        self.figure = Figure(figsize=(6, 4))
        self.canvas = FigureCanvas(self.panel, -1, self.figure)
        self.ax = self.figure.add_subplot(111)

        self.y_max = max(self.y_values)
        self.y_min = 0.99 * min(self.y_values)

        # Now create the ranges (which calls plot_initial_data)
        self.create_initial_ranges(1)

        # Scrolled window for Tougaard parameters
        self.param_scroll = wx.ScrolledWindow(control_panel)
        self.param_scroll.SetScrollRate(0, 20)
        self.param_sizer = wx.BoxSizer(wx.VERTICAL)

        # Initialize empty list for parameter controls
        self.tougaard_params = []

        # Create initial parameter set
        self.create_tougaard_params(1)

        # Fit button
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.fit_button = wx.Button(control_panel, label="Fit")
        self.fit_button.SetMinSize((110, 40))
        self.fit_button.Bind(wx.EVT_BUTTON, self.on_fit)
        self.copy_button = wx.Button(control_panel, label="Copy to Peak Fitting Screen")
        self.copy_button.SetMinSize((110, 40))
        self.copy_button.Bind(wx.EVT_BUTTON, self.on_copy_values)
        button_sizer.Add(self.fit_button, 1, wx.ALL, 5)
        button_sizer.Add(self.copy_button, 1, wx.ALL, 5)


        # Initialize vertical lines as None
        self.vline_min = None
        self.vline_max = None

        # Mouse interaction
        self.dragging_range = None
        self.mouse_press_x = None
        self.range_vlines = []  # Store vlines for each range

        # self.y_max = max(self.y_values)
        # self.y_min = 0.99*min(self.y_values)


        # Initial plot setup
        self.ax.plot(self.x_values, self.y_values, 'k-', label='Data')
        self.ax.set_xlabel('Binding Energy (eV)')
        self.ax.set_ylabel('Intensity (CPS)')
        self.ax.legend()
        self.ax.set_xlim(max(self.x_values), min(self.x_values))

        # Plot initial data
        self.plot_initial_data()

        # Layout
        control_sizer.Add(bg_sizer, 0, wx.EXPAND | wx.ALL, 5)
        control_sizer.Add(range_sizer, 2, wx.EXPAND | wx.ALL, 5)  # Changed from 1 to 2 for more space
        control_sizer.Add(self.param_scroll, 1, wx.EXPAND | wx.ALL, 5)
        control_sizer.Add(button_sizer, 0, wx.EXPAND|wx.ALL, 5)
        control_panel.SetSizer(control_sizer)

        main_sizer.Add(control_panel, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(self.canvas, 1, wx.EXPAND | wx.ALL, 5)

        self.panel.SetSizer(main_sizer)



        # Bind events
        # self.min_range.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_range_change)
        # self.max_range.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_range_change)
        # self.num_tougaard.Bind(wx.EVT_SPINCTRL, self.on_num_tougaard_change)
        self.panel.Bind(wx.EVT_CHAR_HOOK, self.on_key_press)

        from libraries.ConfigFile import set_consistent_fonts
        set_consistent_fonts(self)

    def on_key_press(self, event):
        if event.ControlDown():
            if event.GetKeyCode() in [wx.WXK_UP, wx.WXK_DOWN]:
                # y_min =
                intensity_factor = 0.05

                if event.GetKeyCode() == wx.WXK_DOWN:
                    self.y_max = max(self.y_max - intensity_factor * max(self.y_values), self.y_min)
                else:
                    self.y_max += intensity_factor * max(self.y_values)

                self.ax.set_ylim(self.y_min, self.y_max)
                self.canvas.draw_idle()
                return

        event.Skip()

    def create_tougaard_params_OLD(self, num_tougaard):
        # Store current values if they exist
        old_values = []
        old_fixed = []
        for params in self.tougaard_params:
            old_values.append({
                'B': params['B']['value'].GetValue(),
                'C': params['C']['value'].GetValue(),
                'D': params['D']['value'].GetValue()
            })
            old_fixed.append({
                'B': params['B']['fixed'].GetValue(),
                'C': params['C']['fixed'].GetValue(),
                'D': params['D']['fixed'].GetValue()
            })

        self.param_sizer.Clear(True)
        self.tougaard_params = []

        for i in range(num_tougaard):
            param_box = wx.StaticBox(self.param_scroll, label=f"Tougaard {i + 1} B [Scaling], C [Position], "
                                                              f"D [Width], T0 is 0")
            box_sizer = wx.StaticBoxSizer(param_box, wx.VERTICAL)

            params = {}
            param_grid = wx.GridBagSizer(5, 5)

            for row, param in enumerate(['B', 'C', 'D']):
                if i < len(old_values):
                    value = old_values[i][param]
                    is_fixed = old_fixed[i][param]
                else:
                    value = 2866 if param == 'B' else 1643 if param == 'C' else 1
                    is_fixed = False

                # Set default min values for C and D
                min_default = '0'
                if param == 'C':
                    min_default = '600'
                elif param == 'D':
                    min_default = '1200'

                params[param] = {
                    'value': wx.SpinCtrlDouble(param_box, min=0, max=20000, inc=0.1, value=str(value)),
                    'min': wx.SpinCtrlDouble(param_box, min=0, max=6000, inc=0.1, value=min_default),
                    'max': wx.SpinCtrlDouble(param_box, min=0, max=20000, inc=0.1, value='6000'),
                    'fixed': wx.CheckBox(param_box, label="Fix")
                }
                params[param]['fixed'].SetValue(is_fixed)

                param_grid.Add(wx.StaticText(param_box, label=f"{param}:"), pos=(row, 0), flag=wx.ALIGN_CENTER_VERTICAL)
                param_grid.Add(params[param]['value'], pos=(row, 1), flag=wx.EXPAND)
                param_grid.Add(params[param]['fixed'], pos=(row, 2), flag=wx.ALIGN_CENTER_VERTICAL)
                param_grid.Add(wx.StaticText(param_box, label="Min:"), pos=(row, 3), flag=wx.ALIGN_CENTER_VERTICAL)
                param_grid.Add(params[param]['min'], pos=(row, 4), flag=wx.EXPAND)
                param_grid.Add(wx.StaticText(param_box, label="Max:"), pos=(row, 5), flag=wx.ALIGN_CENTER_VERTICAL)
                param_grid.Add(params[param]['max'], pos=(row, 6), flag=wx.EXPAND)

            box_sizer.Add(param_grid, 0, wx.EXPAND | wx.ALL, 5)
            self.tougaard_params.append(params)
            self.param_sizer.Add(box_sizer, 0, wx.EXPAND | wx.ALL, 5)

        self.param_scroll.SetSizer(self.param_sizer)
        self.param_scroll.Layout()

    def create_tougaard_params(self, num=1):  # Always 1 now
        # Clear existing parameters
        for child in self.param_scroll.GetChildren():
            child.Destroy()

        self.tougaard_params = []

        # Create single Tougaard parameter set
        box_sizer = wx.StaticBoxSizer(wx.StaticBox(self.param_scroll, label="Tougaard Parameters"), wx.VERTICAL)
        param_grid = wx.GridBagSizer(hgap=5, vgap=2)

        params = {}
        for row, (param, default, is_fixed, min_default) in enumerate([
            ('B', '2866.00', False, '0'),
            ('C', '1643.00', False, '600'),
            ('D', '1.00', False, '1200')
        ]):
            params[param] = {
                'value': wx.SpinCtrlDouble(self.param_scroll, min=0, max=20000, inc=0.1, value=default),
                'min': wx.SpinCtrlDouble(self.param_scroll, min=0, max=6000, inc=0.1, value=min_default),
                'max': wx.SpinCtrlDouble(self.param_scroll, min=0, max=20000, inc=0.1, value='6000'),
                'fixed': wx.CheckBox(self.param_scroll, label="Fix")
            }
            params[param]['value'].SetDigits(2)
            params[param]['min'].SetDigits(2)
            params[param]['max'].SetDigits(2)
            params[param]['fixed'].SetValue(is_fixed)

            param_grid.Add(wx.StaticText(self.param_scroll, label=f"{param}:"), pos=(row, 0), flag=wx.ALIGN_CENTER_VERTICAL)
            param_grid.Add(params[param]['value'], pos=(row, 1), flag=wx.EXPAND)
            param_grid.Add(params[param]['fixed'], pos=(row, 2), flag=wx.ALIGN_CENTER_VERTICAL)
            param_grid.Add(wx.StaticText(self.param_scroll, label="Min:"), pos=(row, 3), flag=wx.ALIGN_CENTER_VERTICAL)
            param_grid.Add(params[param]['min'], pos=(row, 4), flag=wx.EXPAND)
            param_grid.Add(wx.StaticText(self.param_scroll, label="Max:"), pos=(row, 5), flag=wx.ALIGN_CENTER_VERTICAL)
            param_grid.Add(params[param]['max'], pos=(row, 6), flag=wx.EXPAND)

        box_sizer.Add(param_grid, 0, wx.EXPAND | wx.ALL, 5)
        self.tougaard_params.append(params)

        self.param_sizer = wx.BoxSizer(wx.VERTICAL)
        self.param_sizer.Add(box_sizer, 0, wx.EXPAND | wx.ALL, 5)
        self.param_scroll.SetSizer(self.param_sizer)
        self.param_scroll.Layout()

    def on_num_tougaard_change(self, event):
        num = self.num_tougaard.GetValue()
        self.create_tougaard_params(num)
        self.param_scroll.FitInside()

    def on_fit(self, event):
        fitting_ranges = self.get_fitting_ranges()
        if not fitting_ranges:
            print("Error: No active fitting ranges defined")
            return
        bg_start = self.bg_start.GetValue()

        if self.x_values is None or self.y_values is None:
            print("Error: No data available for fitting")
            return

        params = lmfit.Parameters()
        for i, tougaard in enumerate(self.tougaard_params):
            for param_name in ['B', 'C', 'D']:
                name = f'{param_name}{i + 1}'
                value = float(tougaard[param_name]['value'].GetValue())
                is_fixed = tougaard[param_name]['fixed'].GetValue()

                if is_fixed:
                    params.add(name, value=value, vary=False)
                else:
                    min_val = float(tougaard[param_name]['min'].GetValue())
                    max_val = float(tougaard[param_name]['max'].GetValue())
                    params.add(name, value=value, min=min_val, max=max_val, vary=True)

        bg_mask = self.x_values >= bg_start
        x_full = self.x_values[bg_mask]
        y_full = self.y_values[bg_mask]

        if len(x_full) == 0 or len(y_full) == 0:
            print("Error: No data points in selected range")
            return

        # Combine all fitting ranges
        combined_mask = np.zeros_like(x_full, dtype=bool)
        for min_e, max_e in fitting_ranges:
            range_mask = (x_full >= min_e) & (x_full <= max_e)
            combined_mask = combined_mask | range_mask

        x_fit = x_full[combined_mask]
        y_fit = y_full[combined_mask]

        if len(x_fit) == 0 or len(y_fit) == 0:
            print("Error: No data points in fitting ranges")
            return

        def tougaard_model(params, x_full, y_full, x_fit, y_fit):
            baseline = y_full[-1]
            y_shifted = y_full - baseline
            dx = np.mean(np.diff(x_full))
            background_full = np.zeros_like(y_full)

            for i in range(len(x_full)):
                E = x_full[i:] - x_full[i]
                K_total = 0
                for j in range(len(self.tougaard_params)):
                    B = params[f'B{j + 1}'].value
                    C = params[f'C{j + 1}'].value
                    D = params[f'D{j + 1}'].value
                    K_total += B * E / ((C - E ** 2) ** 2 + D * E ** 2)
                background_full[i] = np.trapz(K_total * y_shifted[i:], dx=dx)

            background_full += baseline
            fit_indices = np.where(combined_mask)[0]
            background_fit = background_full[fit_indices]
            return y_fit - background_fit

        result = lmfit.minimize(tougaard_model,
                                params,
                                args=(x_full, y_full, x_fit, y_fit),
                                method='least_squares',
                                ftol=1e-10,
                                xtol=1e-10,
                                max_nfev=100,
                                scale_covar=True,
                                verbose=True)

        print("\nFit Results:")
        print(result.message)
        print(f"Success: {result.success}")
        print(f"Number of function evaluations: {result.nfev}")
        print("\nFitted Parameters:")
        for i in range(len(self.tougaard_params)):
            for param in ['B', 'C', 'D']:
                name = f"{param}{i + 1}"
                value = result.params[name].value
                stderr = result.params[name].stderr if result.params[name].stderr is not None else 0
                print(f"{name}: {value:.2f} ± {stderr:.2f}")

        if result.success:
            self.plot_results(x_fit, y_fit, result.params, result)
        else:
            self.plot_results(x_fit, y_fit, result.params, result)
            print("Warning: Fit did not converge")

    def plot_results(self, x_fit, y_fit, fitted_params, result):
        self.ax.clear()
        bg_start = self.bg_start.GetValue()

        # Update control values with fitted parameters
        for i, tougaard in enumerate(self.tougaard_params):
            tougaard['B']['value'].SetValue(f"{fitted_params[f'B{i + 1}'].value:.2f}")
            tougaard['C']['value'].SetValue(f"{fitted_params[f'C{i + 1}'].value:.2f}")
            tougaard['D']['value'].SetValue(f"{fitted_params[f'D{i + 1}'].value:.2f}")

        # Calculate each background separately
        mask = self.x_values >= bg_start
        x_bg = self.x_values[mask]
        y_bg = self.y_values[mask]
        baseline = y_bg[-1]

        total_background = np.zeros_like(y_bg) + baseline

        # Plot data
        self.ax.plot(self.x_values, self.y_values, 'k-', label='Data')

        # Plot each Tougaard background
        colors = plt.cm.tab10(np.linspace(0, 1, len(self.tougaard_params)))
        for j, color in enumerate(colors):
            background = np.zeros_like(y_bg) + baseline
            y_shifted = y_bg - baseline
            dx = np.mean(np.diff(x_bg))

            for i in range(len(x_bg)):
                E = x_bg[i:] - x_bg[i]
                B = fitted_params[f'B{j + 1}'].value
                C = fitted_params[f'C{j + 1}'].value
                D = fitted_params[f'D{j + 1}'].value
                K = B * E / ((C - E ** 2) ** 2 + D * E ** 2)
                background[i] += np.trapz(K * y_shifted[i:], dx=dx)

            total_background += (background - baseline)
            self.ax.plot(x_bg, background, '--', color=color, alpha=0.5,
                         label=f'Tougaard {j + 1}')

        # Plot total background
        self.ax.plot(x_bg, total_background, 'b-', label='Total Background')

        # Plot vertical lines for all active fitting ranges
        fitting_ranges = self.get_fitting_ranges()
        range_colors = ['red', 'blue', 'orange', 'purple', 'brown', 'pink', 'gray', 'olive', 'cyan', 'magenta']

        for i, (min_val, max_val) in enumerate(fitting_ranges):
            color = range_colors[i % len(range_colors)]
            self.ax.axvline(x=min_val, color=color, linestyle='--', alpha=0.7)
            self.ax.axvline(x=max_val, color=color, linestyle=':', alpha=0.7)

        self.ax.set_xlabel('Binding Energy (eV)')
        self.ax.set_ylabel('Intensity (CPS)')

        # Only show legend if there are items to show
        handles, labels = self.ax.get_legend_handles_labels()
        if handles:
            self.ax.legend()

        self.ax.set_xlim(max(self.x_values), min(self.x_values))
        self.ax.set_ylim(self.y_min, self.y_max)
        self.canvas.draw()

    def plot_initial_data(self):
        self.ax.clear()
        self.ax.plot(self.x_values, self.y_values, 'k-', label='Data')
        self.ax.set_xlabel('Binding Energy (eV)')
        self.ax.set_ylabel('Intensity (CPS)')
        self.ax.set_xlim(max(self.x_values), min(self.x_values))

        # Plot vertical lines for all active fitting ranges
        fitting_ranges = self.get_fitting_ranges()
        colors = ['red', 'blue', 'orange', 'purple', 'brown', 'pink', 'gray', 'olive', 'cyan', 'magenta']

        for i, (min_val, max_val) in enumerate(fitting_ranges):
            color = colors[i % len(colors)]
            self.ax.axvline(x=min_val, color=color, linestyle='--', alpha=0.7,
                            label=f'Range {i + 1} Min')
            self.ax.axvline(x=max_val, color=color, linestyle=':', alpha=0.7,
                            label=f'Range {i + 1} Max')

        # Only show legend if there are items to show
        handles, labels = self.ax.get_legend_handles_labels()
        if handles:
            self.ax.legend()

        self.canvas.draw()

    def update_range_lines(self):
        """Update range lines when controls change"""
        self.plot_initial_data()

    def plot_range_lines(self):
        """Plot vertical lines for all active fitting ranges"""
        fitting_ranges = self.get_fitting_ranges()
        colors = ['red', 'blue', 'orange', 'purple', 'brown', 'pink', 'gray', 'olive', 'cyan', 'magenta']

        for i, (min_val, max_val) in enumerate(fitting_ranges):
            color = colors[i % len(colors)]
            self.ax.axvline(x=min_val, color=color, linestyle='--', alpha=0.7,
                            label=f'Range {i + 1} Min' if i == 0 else "")
            self.ax.axvline(x=max_val, color=color, linestyle=':', alpha=0.7,
                            label=f'Range {i + 1} Max' if i == 0 else "")

    def on_num_ranges_change(self, event):
        num_ranges = self.num_ranges.GetValue()
        self.create_initial_ranges(num_ranges)

    def create_initial_ranges(self, num_ranges):
        # Store previous values before clearing
        previous_values = []
        for range_data in self.ranges:
            previous_values.append({
                'min': range_data['min'].GetValue(),
                'max': range_data['max'].GetValue(),
                'use_range': range_data['use_range'].GetValue(),
                'moveable': range_data.get('moveable', wx.CheckBox()).GetValue() if 'moveable' in range_data else False
            })

        # Clear existing ranges
        for range_data in self.ranges:
            range_data['panel'].Destroy()
        self.ranges = []

        min_x = min(self.x_values)
        max_x = max(self.x_values)

        for i in range(num_ranges):
            range_panel = wx.Panel(self.ranges_panel)
            range_panel.SetMinSize((-1, 55))  # Increased height
            range_panel_sizer = wx.BoxSizer(wx.VERTICAL)

            # Top row: Use range and moveable checkboxes
            checkbox_sizer = wx.BoxSizer(wx.HORIZONTAL)

            use_range = wx.CheckBox(range_panel, label=f"Use Range {i + 1}")
            if i < len(previous_values):
                use_range.SetValue(previous_values[i]['use_range'])
            else:
                use_range.SetValue(True)
            use_range.Bind(wx.EVT_CHECKBOX, self.on_range_change)

            moveable = wx.CheckBox(range_panel, label="Moveable")
            if i < len(previous_values):
                moveable.SetValue(previous_values[i]['moveable'])
            moveable.Bind(wx.EVT_CHECKBOX, self.on_moveable_change)

            checkbox_sizer.Add(use_range, 0, wx.ALL, 2)
            checkbox_sizer.Add(moveable, 0, wx.ALL, 2)
            range_panel_sizer.Add(checkbox_sizer, 0, wx.EXPAND | wx.ALL, 2)

            # Min/Max controls in a horizontal layout
            range_controls_sizer = wx.BoxSizer(wx.HORIZONTAL)

            # Min control
            range_controls_sizer.Add(wx.StaticText(range_panel, label="Min:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)

            # Use previous values if available
            if i < len(previous_values):
                min_val = previous_values[i]['min']
            else:
                min_val = max_x - 15

            min_ctrl = wx.SpinCtrlDouble(range_panel, min=0.00, max=2000.00, inc=0.10, value=f"{min_val:.2f}")
            min_ctrl.SetDigits(2)
            min_ctrl.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_range_change)
            range_controls_sizer.Add(min_ctrl, 1, wx.EXPAND | wx.RIGHT, 10)

            # Max control
            range_controls_sizer.Add(wx.StaticText(range_panel, label="Max:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)

            if i < len(previous_values):
                max_val = previous_values[i]['max']
            else:
                max_val = max_x - 1

            max_ctrl = wx.SpinCtrlDouble(range_panel, min=0.00, max=2000.00, inc=0.10, value=f"{max_val:.2f}")
            max_ctrl.SetDigits(2)
            max_ctrl.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_range_change)
            range_controls_sizer.Add(max_ctrl, 1, wx.EXPAND)

            range_panel_sizer.Add(range_controls_sizer, 0, wx.EXPAND | wx.ALL, 5)
            range_panel.SetSizer(range_panel_sizer)

            range_data = {
                'panel': range_panel,
                'use_range': use_range,
                'min': min_ctrl,
                'max': max_ctrl,
                'moveable': moveable
            }

            self.ranges.append(range_data)
            self.ranges_panel_sizer.Add(range_panel, 0, wx.EXPAND | wx.ALL, 5)

        self.ranges_panel.FitInside()
        self.ranges_panel.Layout()

        # Call plot_initial_data only if matplotlib is set up
        if hasattr(self, 'ax'):
            self.setup_mouse_events()
            self.plot_initial_data()

    def get_fitting_ranges(self):
        ranges = []
        for range_data in self.ranges:
            if range_data['use_range'].GetValue():
                min_val = min(range_data['min'].GetValue(), range_data['max'].GetValue())
                max_val = max(range_data['min'].GetValue(), range_data['max'].GetValue())
                ranges.append((min_val, max_val))
        return ranges

    def on_range_change(self, event):
        """Handle changes to fitting range controls"""
        self.plot_initial_data()

    def on_copy_values(self, event):
        for i, tougaard in enumerate(self.tougaard_params):
            B = tougaard['B']['value'].GetValue()
            C = tougaard['C']['value'].GetValue()
            D = tougaard['D']['value'].GetValue()

            if i == 0:
                self.parent.cross_section.SetValue(f"{B},{C},{D},0")
            elif i == 1:
                self.parent.cross_section2.SetValue(f"{B},{C},{D},0")
            elif i == 2:
                self.parent.cross_section3.SetValue(f"{B},{C},{D},0")

    def setup_mouse_events(self):
        """Setup mouse event handlers for dragging vlines"""
        self.canvas.mpl_connect('button_press_event', self.on_mouse_press)
        self.canvas.mpl_connect('motion_notify_event', self.on_mouse_motion)
        self.canvas.mpl_connect('button_release_event', self.on_mouse_release)

    def on_moveable_change(self, event):
        """Handle moveable checkbox changes - ensure only one can be selected"""
        changed_checkbox = event.GetEventObject()

        if changed_checkbox.GetValue():
            # Uncheck all other moveable checkboxes
            for range_data in self.ranges:
                if range_data['moveable'] != changed_checkbox:
                    range_data['moveable'].SetValue(False)

        self.plot_initial_data()

    def on_mouse_press(self, event):
        """Handle mouse press for starting vline drag"""
        if event.inaxes != self.ax:
            return

        # Find which moveable range is selected
        moveable_range_idx = None
        for i, range_data in enumerate(self.ranges):
            if range_data['moveable'].GetValue() and range_data['use_range'].GetValue():
                moveable_range_idx = i
                break

        if moveable_range_idx is None:
            return

        # Check if click is near a vline for the moveable range
        range_data = self.ranges[moveable_range_idx]
        min_val = range_data['min'].GetValue()
        max_val = range_data['max'].GetValue()

        click_x = event.xdata
        tolerance = (max(self.x_values) - min(self.x_values)) * 0.02  # 2% of x range

        if abs(click_x - min_val) < tolerance:
            self.dragging_range = (moveable_range_idx, 'min')
            self.mouse_press_x = click_x
        elif abs(click_x - max_val) < tolerance:
            self.dragging_range = (moveable_range_idx, 'max')
            self.mouse_press_x = click_x

    def on_mouse_motion(self, event):
        """Handle mouse motion for dragging vlines"""
        if self.dragging_range is None or event.inaxes != self.ax:
            return

        range_idx, line_type = self.dragging_range
        range_data = self.ranges[range_idx]

        new_x = event.xdata
        if new_x is None:
            return

        # Update the control value
        if line_type == 'min':
            range_data['min'].SetValue(f"{new_x:.2f}")
        else:
            range_data['max'].SetValue(f"{new_x:.2f}")

        # Redraw the plot
        self.plot_initial_data()

    def on_mouse_release(self, event):
        """Handle mouse release to stop dragging"""
        self.dragging_range = None
        self.mouse_press_x = None


class CustomComboBox2(wx.ComboBox):
    def __init__(self, parent, style=wx.CB_READONLY):
        style |= wx.CB_OWNERDRAW
        super().__init__(parent, style=style)
        self.green_items = []
        self.Bind(wx.EVT_DRAW_ITEM, self.OnDrawItem)

    def SetGreenItems(self, items):
        self.green_items = items
        self.Refresh()

    def OnDrawItem(self, evt):
        dc = wx.PaintDC(self)
        dc.SetFont(self.GetFont())
        text = self.GetString(evt.GetIndex())
        rect = evt.GetRect()

        if evt.IsSelected():
            bg_color = wx.SystemSettings.GetColour(wx.SYS_COLOUR_HIGHLIGHT)
            text_color = wx.SystemSettings.GetColour(wx.SYS_COLOUR_HIGHLIGHTTEXT)
        else:
            bg_color = wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW)
            text_color = wx.GREEN if text in self.green_items else wx.BLACK

        dc.SetBrush(wx.Brush(bg_color))
        dc.SetTextForeground(text_color)
        dc.DrawRectangle(rect)
        dc.DrawText(text, rect.x + 3, rect.y + 2)

class CustomComboBox(wx.ComboBox):
    def __init__(self, parent, style=wx.CB_READONLY):
        super().__init__(parent, style=style)
        self.green_items = []

    def SetGreenItems(self, items):
        self.green_items = items
        for i, item in enumerate(self.GetItems()):
            self.SetForegroundColour(wx.GREEN if item in self.green_items else wx.BLACK)
            self.Refresh()

    def GetGreenItems(self):
        return self.green_items

    def OnMeasureItem(self, item):
        return 24

    def OnSelect(self, event):
        selection = self.GetStringSelection()
        if selection in self.green_items:
            wx.CallAfter(self.SetSelection, self.GetSelection())
        else:
            event.Skip()