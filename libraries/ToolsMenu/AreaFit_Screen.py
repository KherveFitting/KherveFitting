import wx
from libraries.FileMenu.Save import save_state
from Functions import remove_peak
import numpy as np
import time
from libraries.ToolsMenu.AutoID import AutoSurveyID

class BackgroundWindow(wx.Frame):
    def __init__(self, parent, *args, **kw):
        super(BackgroundWindow, self).__init__(parent, *args, **kw, style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX | wx.MINIMIZE_BOX | wx.SYSTEM_MENU) | wx.STAY_ON_TOP)
        self.parent = parent
        self.SetTitle("Measure Area")

        if 'wxMac' in wx.PlatformInfo:
            self.SetSize((260, 345))
            # self.SetMinSize((260, 345))
            # self.SetMaxSize((260, 345))
        elif 'wxGTK' in wx.PlatformInfo:
            desktop = self.get_linux_desktop()
            if desktop == 'gnome':
                self.SetSize((310, 560))
            elif desktop == 'kde':
                self.SetSize((260, 480))
            elif desktop == 'xfce':
                print('linux xfce')
                self.SetSize((260, 460))
            else:
                self.SetSize((280, 520))
            print(f'GTK environment: {desktop}')
        else:
            self.SetSize((272, 370))
            # self.SetMinSize((270, 370))
            # self.SetMaxSize((270, 370))

        # Initialize UI with tabs
        self.init_ui()

        from libraries.ConfigFile import set_consistent_fonts
        set_consistent_fonts(self)


    def init_ui(self):
        """Initialize UI with notebook tabs"""
        panel = wx.Panel(self)

        def detect_dark_mode():
            if 'wxMac' in wx.PlatformInfo:
                return wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW).GetLuminance() < 0.5
            elif 'wxMSW' in wx.PlatformInfo:
                return wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW).GetLuminance() < 0.5
            elif 'wxGTK' in wx.PlatformInfo:
                return wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW).GetLuminance() < 0.5
            return False

        if not detect_dark_mode():
            panel.SetBackgroundColour(wx.Colour(250, 250, 230))

        main_sizer = wx.BoxSizer(wx.VERTICAL)

        notebook = wx.Notebook(panel)

        if not detect_dark_mode():
            notebook.SetBackgroundColour(wx.Colour(250, 250, 230))

        # Create tabs
        self.init_measure_area_tab(notebook)
        self.init_batching_tab(notebook)

        main_sizer.Add(notebook, 1, wx.EXPAND, border=5)
        panel.SetSizer(main_sizer)

        # Track which tab is selected
        self.parent.area_tab_selected = True
        self.parent.area_batch_tab_selected = False

        notebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.on_tab_change)

        # Initially disable all Tougaard controls
        self.cross_section.Enable(False)
        self.cross_section_label.Enable(False)
        self.tougaard_fit_btn.Enable(False)

        self.Bind(wx.EVT_CLOSE, self.on_close)
        self.method_combobox.Bind(wx.EVT_COMBOBOX, self.on_bkg_method_change)

        # Add key event handling for TAB functionality
        self.Bind(wx.EVT_CHAR_HOOK, self.on_key_press)
        self.current_peak_index = 0

    def init_measure_area_tab(self, notebook):
        """Initialize the Measure Area tab - contains all existing controls"""
        self.measure_area_panel = wx.Panel(notebook)

        def detect_dark_mode():
            if 'wxMac' in wx.PlatformInfo:
                return wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW).GetLuminance() < 0.5
            elif 'wxMSW' in wx.PlatformInfo:
                return wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW).GetLuminance() < 0.5
            elif 'wxGTK' in wx.PlatformInfo:
                return wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW).GetLuminance() < 0.5
            return False

        if not detect_dark_mode():
            self.measure_area_panel.SetBackgroundColour(wx.Colour(250, 250, 230))

        # Create controls - USE PANEL NOT SELF.MEASURE_AREA_PANEL FOR NOW
        panel = self.measure_area_panel

        method_label = wx.StaticText(panel, label="Method:")
        self.method_combobox = wx.ComboBox(panel, choices=["Multi-Regions Smart", "Shirley", "Linear",
                                                           'U4-Tougaard', 'U2-Tougaard'],
                                           style=wx.CB_READONLY)
        self.method_combobox.SetSelection(4)
        self.method_combobox.SetMaxSize((125, 25))

        offset_h_label = wx.StaticText(panel, label="Offset (Left):")
        self.offset_h_text = wx.TextCtrl(panel, value="0.00", style=wx.TE_PROCESS_ENTER)
        self.offset_h_text.Bind(wx.EVT_TEXT_ENTER, self.on_text_control_enter)
        self.offset_h_text.Bind(wx.EVT_KILL_FOCUS, self.on_text_control_focus_lost)

        offset_l_label = wx.StaticText(panel, label="Offset (Right):")
        self.offset_l_text = wx.TextCtrl(panel, value="0.00", style=wx.TE_PROCESS_ENTER)
        self.offset_l_text.Bind(wx.EVT_TEXT_ENTER, self.on_text_control_enter)
        self.offset_l_text.Bind(wx.EVT_KILL_FOCUS, self.on_text_control_focus_lost)

        self.min_range_label = wx.StaticText(panel, label='Range (Right):')
        self.min_range_text = wx.TextCtrl(panel, value="0.00", style=wx.TE_PROCESS_ENTER)
        self.min_range_text.Bind(wx.EVT_TEXT_ENTER, self.on_min_range_enter)
        self.min_range_text.Bind(wx.EVT_KILL_FOCUS, self.on_min_range_enter)

        self.max_range_label = wx.StaticText(panel, label='Range (Left):')
        self.max_range_text = wx.TextCtrl(panel, value="0.00", style=wx.TE_PROCESS_ENTER)
        self.max_range_text.Bind(wx.EVT_TEXT_ENTER, self.on_max_range_enter)
        self.max_range_text.Bind(wx.EVT_KILL_FOCUS, self.on_max_range_enter)

        # Initialize range control updating flag
        self.updating_range_controls = False

        # Initialize range controls from data
        self.update_range_controls_from_data()

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
        remove_peak_button = wx.Button(panel, label="Remove Last\nBackground / Area")
        if 'wxMac' in wx.PlatformInfo:
            remove_peak_button.SetMinSize((90, 30))
        else:
            remove_peak_button.SetMinSize((90, 35))
        remove_peak_button.Bind(wx.EVT_BUTTON, self.on_remove_peak)

        auto_id_button = wx.Button(panel, label="Automatic\nIdentification")
        if 'wxMac' in wx.PlatformInfo:
            auto_id_button.SetMinSize((125, 30))
        else:
            auto_id_button.SetMinSize((110, 35))
        auto_id_button.Bind(wx.EVT_BUTTON, self.on_auto_id)

        # Change the reset button to switch positions button
        switch_vlines_button = wx.Button(panel, label="Switch Regions\n TAB key")
        if 'wxMac' in wx.PlatformInfo:
            switch_vlines_button.SetMinSize((125, 30))
        else:
            switch_vlines_button.SetMinSize((125, 35))
        switch_vlines_button.Bind(wx.EVT_BUTTON, self.on_switch_vlines)
        switch_vlines_button.SetToolTip("Switch between plot range (10%-90%) and peak background ranges\nSame as pressing TAB key")

        background_only_button = wx.Button(panel, label="Create\nBackground / Area")
        if 'wxMac' in wx.PlatformInfo:
            background_only_button.SetMinSize((125, 30))
        else:
            background_only_button.SetMinSize((125, 35))
        background_only_button.Bind(wx.EVT_BUTTON, self.on_background_only)

        peak_label_text_label = wx.StaticText(panel, label="Area Name      ")
        self.peak_label_text = wx.TextCtrl(panel, value="")

        # Core Level List button and Remove Peak button
        core_level_button = wx.Button(panel, label="Core Level\nList")
        if 'wxMac' in wx.PlatformInfo:
            core_level_button.SetMinSize((90, 30))
        else:
            core_level_button.SetMinSize((90, 35))
        core_level_button.Bind(wx.EVT_BUTTON, self.on_show_core_level_list)

        # Layout with GridBagSizer
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
            sizer.Add(self.min_range_label, pos=(4, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.min_range_text, pos=(4, 1), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.max_range_label, pos=(3, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.max_range_text, pos=(3, 1), flag=wx.ALL | wx.EXPAND, border=1)

            # Fourth row: Averaging Points
            sizer.Add(averaging_points_label, pos=(5, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.averaging_points_text, pos=(5, 1), flag=wx.ALL | wx.EXPAND, border=1)

            # Tougaard parameters (only Tougaard1)
            sizer.Add(self.cross_section_label, pos=(6, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.cross_section, pos=(6, 1), flag=wx.ALL | wx.EXPAND, border=1)

            # Help button at row 7
            help_button = wx.Button(panel, label="?")
            help_button.SetMinSize((20, 20))
            help_button.SetToolTip("Click to watch tutorial video on how to use the Area Fit function")
            help_button.Bind(wx.EVT_BUTTON, self.on_help_button)
            sizer.Add(help_button, pos=(7, 1), flag=wx.ALL, border=0)

            sizer.Add(peak_label_text_label, pos=(8, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.peak_label_text, pos=(8, 1), flag=wx.ALL | wx.EXPAND, border=1)

            sizer.Add(core_level_button, pos=(11, 1), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(auto_id_button, pos=(10, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

            # Seventh row: Remove peak and Export buttons
            sizer.Add(self.tougaard_fit_btn, pos=(11, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(switch_vlines_button, pos=(10, 0), flag=wx.ALL | wx.EXPAND, border=1)

            sizer.Add(background_only_button, pos=(12, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(remove_peak_button, pos=(12, 1), flag=wx.ALL | wx.EXPAND, border=1)
        else:  # For Windows
            # First row: Method
            sizer.Add(method_label, pos=(0, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.method_combobox, pos=(0, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

            # second Third row: Offset (H) and Offset (L)
            sizer.Add(offset_h_label, pos=(1, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.offset_h_text, pos=(1, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(offset_l_label, pos=(2, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.offset_l_text, pos=(2, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

            # Add new range controls
            sizer.Add(self.min_range_label, pos=(4, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.min_range_text, pos=(4, 1), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.max_range_label, pos=(3, 0), flag=wx.ALL | wx.EXPAND, border=1)
            sizer.Add(self.max_range_text, pos=(3, 1), flag=wx.ALL | wx.EXPAND, border=1)

            # Fourth row: Averaging Points
            sizer.Add(averaging_points_label, pos=(5, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.averaging_points_text, pos=(5, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

            # Fourth row: Tougaard parameters
            sizer.Add(self.cross_section_label, pos=(6, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(self.cross_section, pos=(6, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

            # Help button at row 7
            help_button = wx.Button(panel, label="?")
            help_button.SetMinSize((20, 20))
            help_button.SetToolTip("Click to watch tutorial video on how to use the Area Fit function")
            help_button.Bind(wx.EVT_BUTTON, self.on_help_button)
            sizer.Add(help_button, pos=(7, 1), flag=wx.ALL, border=0)

            sizer.Add(peak_label_text_label, pos=(8, 0), flag=wx.ALL | wx.EXPAND, border=0)
            sizer.Add(self.peak_label_text, pos=(8, 1), flag=wx.ALL | wx.EXPAND, border=0)

            sizer.Add(switch_vlines_button, pos=(9, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(auto_id_button, pos=(9, 1), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)

            # Seventh row: Remove peak and Export buttons
            sizer.Add(self.tougaard_fit_btn, pos=(10, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(core_level_button, pos=(10, 1), flag=wx.ALL | wx.EXPAND, border=0)

            sizer.Add(background_only_button, pos=(11, 0), flag=wx.EXPAND | wx.BOTTOM | wx.TOP, border=0)
            sizer.Add(remove_peak_button, pos=(11, 1), flag=wx.ALL | wx.EXPAND, border=0)

        self.measure_area_panel.SetSizer(sizer)
        notebook.AddPage(self.measure_area_panel, "Measure Area")

    def init_batching_tab(self, notebook):
        """Initialize the batching tab for area measurements"""
        self.batch_panel = wx.Panel(notebook)

        # Create main sizer
        batch_sizer = wx.GridBagSizer(0, 0)

        # Progress indicator
        self.batch_progress_label = wx.StaticText(self.batch_panel, label="Progress:")
        batch_sizer.Add(self.batch_progress_label, pos=(0, 0), flag=wx.ALL | wx.EXPAND, border=1)

        self.batch_progress_text = wx.TextCtrl(self.batch_panel, style=wx.TE_READONLY)
        batch_sizer.Add(self.batch_progress_text, pos=(0, 1), flag=wx.ALL | wx.EXPAND, border=1)

        # Select All and Unselect All buttons
        select_all_button = wx.Button(self.batch_panel, label="Select All")
        if 'wxMac' in wx.PlatformInfo:
            select_all_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            select_all_button.SetMinSize((125, 35))
        else:
            select_all_button.SetMinSize((125, 35))
        select_all_button.Bind(wx.EVT_BUTTON, self.on_select_all)
        if 'wxMac' or 'wxGTK' in wx.PlatformInfo:
            batch_sizer.Add(select_all_button, pos=(1, 0), flag=wx.ALL | wx.EXPAND, border=1)
        else:
            batch_sizer.Add(select_all_button, pos=(1, 0), flag=wx.ALL | wx.EXPAND, border=0)

        unselect_all_button = wx.Button(self.batch_panel, label="Unselect All")
        if 'wxMac' in wx.PlatformInfo:
            unselect_all_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            unselect_all_button.SetMinSize((125, 35))
        else:
            unselect_all_button.SetMinSize((125, 35))
        unselect_all_button.Bind(wx.EVT_BUTTON, self.on_unselect_all)
        if 'wxMac' in wx.PlatformInfo:
            batch_sizer.Add(unselect_all_button, pos=(1, 1), flag=wx.ALL | wx.EXPAND, border=1)
        else:
            batch_sizer.Add(unselect_all_button, pos=(1, 1), flag=wx.ALL | wx.EXPAND, border=0)

        # Core levels checklist box
        self.core_levels_checklist = wx.CheckListBox(self.batch_panel)
        self.core_levels_checklist.SetMinSize((250, 160))
        self.populate_core_levels_list()

        self.core_levels_checklist.Bind(wx.EVT_CONTEXT_MENU, self.on_core_levels_context_menu)
        batch_sizer.Add(self.core_levels_checklist, pos=(2, 0), span=(1, 2), flag=wx.ALL | wx.EXPAND, border=1)

        # Propagate Area to Column button (instead of Prop. Fit to Column)
        propagate_area_button = wx.Button(self.batch_panel, label="Prop. Area\nto Column")
        if 'wxMac' in wx.PlatformInfo:
            propagate_area_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            propagate_area_button.SetMinSize((125, 35))
        else:
            propagate_area_button.SetMinSize((125, 35))
        propagate_area_button.Bind(wx.EVT_BUTTON, self.on_propagate_area)
        if 'wxMac' or 'wxGTK' in wx.PlatformInfo:
            batch_sizer.Add(propagate_area_button, pos=(3, 0), flag=wx.ALL | wx.EXPAND, border=1)
        else:
            batch_sizer.Add(propagate_area_button, pos=(3, 0), flag=wx.ALL | wx.EXPAND, border=0)

        # Propagate Row to Selected Core Levels button
        propagate_row_button = wx.Button(self.batch_panel, label="Prop. Row to\nSelected Core Levels")
        if 'wxMac' in wx.PlatformInfo:
            propagate_row_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            propagate_row_button.SetMinSize((125, 35))
        else:
            propagate_row_button.SetMinSize((125, 35))
        propagate_row_button.Bind(wx.EVT_BUTTON, self.on_propagate_row)
        if 'wxMac' or 'wxGTK' in wx.PlatformInfo:
            batch_sizer.Add(propagate_row_button, pos=(3, 1), flag=wx.ALL | wx.EXPAND, border=1)
        else:
            batch_sizer.Add(propagate_row_button, pos=(3, 1), flag=wx.ALL | wx.EXPAND, border=0)

        # Measure Selected Areas button (instead of Fit Selected Core Levels)
        measure_all_button = wx.Button(self.batch_panel, label="Measure Selected\nCore Levels")
        if 'wxMac' in wx.PlatformInfo:
            measure_all_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            measure_all_button.SetMinSize((125, 35))
        else:
            measure_all_button.SetMinSize((125, 35))
        measure_all_button.Bind(wx.EVT_BUTTON, self.on_measure_all)
        if 'wxMac' or 'wxGTK' in wx.PlatformInfo:
            batch_sizer.Add(measure_all_button, pos=(4, 0), span=(1, 2), flag=wx.ALL | wx.EXPAND, border=1)
        else:
            batch_sizer.Add(measure_all_button, pos=(4, 0), span=(1, 2), flag=wx.ALL | wx.EXPAND, border=0)

        self.batch_panel.SetSizer(batch_sizer)
        notebook.AddPage(self.batch_panel, "Batching")

    def populate_core_levels_list(self):
        """Populate the core levels checklist with current core levels"""
        self.core_levels_checklist.Clear()

        if hasattr(self.parent, 'Data') and 'Core levels' in self.parent.Data:
            core_levels = list(self.parent.Data['Core levels'].keys())
            for core_level in core_levels:
                self.core_levels_checklist.Append(core_level)

            # Select all by default
            for i in range(self.core_levels_checklist.GetCount()):
                self.core_levels_checklist.Check(i, True)

    def on_select_all(self, event):
        """Select all core levels in the checklist"""
        for i in range(self.core_levels_checklist.GetCount()):
            self.core_levels_checklist.Check(i, True)

    def on_unselect_all(self, event):
        """Unselect all core levels in the checklist"""
        for i in range(self.core_levels_checklist.GetCount()):
            self.core_levels_checklist.Check(i, False)

    def on_core_levels_context_menu(self, event):
        """Handle right-click on core levels checklist"""
        # Get all core level names from the checklist
        all_core_levels = []
        for i in range(self.core_levels_checklist.GetCount()):
            all_core_levels.append(self.core_levels_checklist.GetString(i))

        if not all_core_levels:
            return

        # Group core levels by type
        core_level_groups = {}
        for core_level in all_core_levels:
            core_type = self.extract_column_type(core_level)
            if core_type not in core_level_groups:
                core_level_groups[core_type] = []
            core_level_groups[core_type].append(core_level)

        # Create context menu
        menu = wx.Menu()

        # Sort core level types for consistent ordering
        sorted_core_types = sorted(core_level_groups.keys())

        for core_type in sorted_core_types:
            core_levels_of_type = core_level_groups[core_type]

            # Create select and unselect items for this core level type
            select_item = menu.Append(wx.ID_ANY, f"Select All {core_type} ({len(core_levels_of_type)})")
            unselect_item = menu.Append(wx.ID_ANY, f"Unselect All {core_type} ({len(core_levels_of_type)})")

            # Bind events
            self.Bind(wx.EVT_MENU, lambda evt, ct=core_type: self.select_all_core_type(ct), select_item)
            self.Bind(wx.EVT_MENU, lambda evt, ct=core_type: self.unselect_all_core_type(ct), unselect_item)

            # Add separator after each core level type (except the last one)
            if core_type != sorted_core_types[-1]:
                menu.AppendSeparator()

        # Show the menu
        self.core_levels_checklist.PopupMenu(menu)
        menu.Destroy()

    def select_all_core_type(self, core_type):
        """Select all core levels of a specific type"""
        for i in range(self.core_levels_checklist.GetCount()):
            core_level = self.core_levels_checklist.GetString(i)
            if self.extract_column_type(core_level) == core_type:
                self.core_levels_checklist.Check(i, True)

    def unselect_all_core_type(self, core_type):
        """Unselect all core levels of a specific type"""
        for i in range(self.core_levels_checklist.GetCount()):
            core_level = self.core_levels_checklist.GetString(i)
            if self.extract_column_type(core_level) == core_type:
                self.core_levels_checklist.Check(i, False)

    def extract_column_type(self, sheet_name):
        """Extract the column type (element + orbital) from sheet name"""
        if not sheet_name:
            return ""

        # Handle common XPS naming patterns
        import re

        # Remove common suffixes/prefixes that might indicate sample info
        cleaned_name = sheet_name.strip()

        # Look for pattern: Element + orbital (e.g., C1s, O1s, Au4f, Ti2p, etc.)
        # Pattern matches: Letters followed by numbers and optional letters, followed by optional digits
        pattern = r'^([A-Z][a-z]?\d+[a-z]*)(?:\d+)?.*'
        match = re.match(pattern, cleaned_name)

        if match:
            return match.group(1)

        # Fallback: look for first part before underscore, space, or dash
        separators = ['_', ' ', '-', '.']
        for sep in separators:
            if sep in cleaned_name:
                first_part = cleaned_name.split(sep)[0]
                # Check if first part looks like a core level (has both letters and numbers)
                if re.match(r'^[A-Z][a-z]?\d+[a-z]*\d*$', first_part):
                    # Remove trailing digits to get just the core level type
                    core_level_match = re.match(r'^([A-Z][a-z]?\d+[a-z]*)', first_part)
                    if core_level_match:
                        return core_level_match.group(1)
                break

        # Final fallback: extract core level pattern from anywhere in name
        core_level_match = re.search(r'([A-Z][a-z]?\d+[a-z]*)', cleaned_name)
        if core_level_match:
            return core_level_match.group(1)

        # Last resort: return first 4 characters or whole name if shorter
        return cleaned_name[:4] if len(cleaned_name) > 4 else cleaned_name

    def get_selected_core_levels_for_area(self):
        """Get list of selected core levels from the checklist"""
        selected = []
        for i in range(self.core_levels_checklist.GetCount()):
            if self.core_levels_checklist.IsChecked(i):
                selected.append(self.core_levels_checklist.GetString(i))
        return selected

    def on_tab_change(self, event):
        """Handle tab change events"""
        selection = event.GetSelection()

        if selection == 0:  # Measure Area tab
            self.parent.area_tab_selected = True
            self.parent.area_batch_tab_selected = False
        else:  # Batching tab
            self.parent.area_tab_selected = False
            self.parent.area_batch_tab_selected = True

        event.Skip()

    def on_propagate_area(self, event):
        """Propagate all areas from current core level to selected core levels in same column"""
        from libraries.FileMenu.Save import save_state
        import numpy as np

        save_state(self.parent)

        # Get current sheet
        current_sheet = self.parent.sheet_combobox.GetValue()
        if not current_sheet or current_sheet not in self.parent.Data['Core levels']:
            wx.MessageBox("No core level selected", "Propagate Failed", wx.OK | wx.ICON_WARNING)
            return

        # Check if current sheet has peaks
        if not hasattr(self.parent, 'peak_params_grid') or self.parent.peak_params_grid.GetNumberRows() == 0:
            wx.MessageBox("No areas/peaks to propagate from current core level", "Propagate Failed", wx.OK | wx.ICON_WARNING)
            return

        grid = self.parent.peak_params_grid

        # Extract ALL peaks and their background info from current sheet
        peaks_to_propagate = []
        num_peaks = grid.GetNumberRows() // 2

        for peak_idx in range(num_peaks):
            row = peak_idx * 2  # Data row (every other row)

            try:
                bkg_type = grid.GetCellValue(row, 14)  # Column 14: Bkg Type
                bkg_low_str = grid.GetCellValue(row, 15)  # Column 15: Bkg Low
                bkg_high_str = grid.GetCellValue(row, 16)  # Column 16: Bkg High
                peak_name = grid.GetCellValue(row, 1)  # Column 1: Peak Name

                # Skip if no background type defined
                if not bkg_type or not bkg_low_str or not bkg_high_str:
                    print(f"Skipping peak {peak_name} - missing background info")
                    continue

                bkg_low = float(bkg_low_str)
                bkg_high = float(bkg_high_str)

                peaks_to_propagate.append({
                    'name': peak_name,
                    'bkg_type': bkg_type,
                    'bkg_low': bkg_low,
                    'bkg_high': bkg_high
                })

            except (ValueError, IndexError) as e:
                print(f"Error extracting background info from row {row}: {e}")
                continue

        if not peaks_to_propagate:
            wx.MessageBox("No valid peaks with background information found in current core level",
                          "Propagate Failed", wx.OK | wx.ICON_WARNING)
            return

        print(f"Found {len(peaks_to_propagate)} peaks to propagate from {current_sheet}")

        # Get selected core levels from checklist
        selected_sheets = []
        for i in range(self.core_levels_checklist.GetCount()):
            if self.core_levels_checklist.IsChecked(i):
                selected_sheets.append(self.core_levels_checklist.GetString(i))

        if not selected_sheets:
            wx.MessageBox("No core levels selected for propagation", "Error", wx.OK | wx.ICON_ERROR)
            return

        # Filter to same column only
        current_column = self.extract_column_type(current_sheet)
        same_column_sheets = []
        omitted_sheets = []

        for sheet in selected_sheets:
            if sheet == current_sheet:
                continue  # Skip self

            sheet_column = self.extract_column_type(sheet)
            if sheet_column == current_column:
                same_column_sheets.append(sheet)
            else:
                omitted_sheets.append(sheet)

        if not same_column_sheets:
            wx.MessageBox(f"No core levels of the same column type ({current_column}) are selected for propagation.",
                          "Nothing to Propagate", wx.OK | wx.ICON_WARNING)
            return

        # Store original sheet
        original_sheet = self.parent.sheet_combobox.GetValue()

        # Propagate to selected core levels
        success_count = 0
        failed_sheets = []
        total_areas_created = 0

        try:
            total_sheets = len(same_column_sheets)

            for i, target_sheet in enumerate(same_column_sheets):
                try:
                    # Update progress
                    self.batch_progress_text.SetValue(f"Processing {target_sheet} ({i + 1}/{total_sheets})")
                    wx.Yield()

                    # Switch to target sheet
                    self.parent.sheet_combobox.SetValue(target_sheet)
                    from libraries.Sheet_Operations import on_sheet_selected
                    on_sheet_selected(self.parent, target_sheet)

                    # Create ALL areas for this sheet
                    sheet_success = True
                    areas_created_for_sheet = 0

                    for peak_info in peaks_to_propagate:
                        if self.create_area_for_sheet(target_sheet,
                                                      peak_info['bkg_type'],
                                                      peak_info['bkg_low'],
                                                      peak_info['bkg_high'],
                                                      peak_info['name']):
                            areas_created_for_sheet += 1
                            total_areas_created += 1
                        else:
                            sheet_success = False

                    if areas_created_for_sheet > 0:
                        success_count += 1
                        print(f"Created {areas_created_for_sheet} areas in {target_sheet}")
                    else:
                        failed_sheets.append(target_sheet)

                except Exception as e:
                    print(f"Failed to propagate areas to {target_sheet}: {e}")
                    import traceback
                    traceback.print_exc()
                    failed_sheets.append(target_sheet)

            # Restore original sheet
            self.parent.sheet_combobox.SetValue(original_sheet)
            on_sheet_selected(self.parent, original_sheet)

            # Show completion message
            completion_msg = f"Successfully propagated {len(peaks_to_propagate)} areas to {success_count}/{total_sheets} {current_column} core levels.\n"
            completion_msg += f"Total areas created: {total_areas_created}"

            if failed_sheets:
                completion_msg += f"\n\nFailed to propagate to {len(failed_sheets)} core levels:\n{', '.join(failed_sheets[:5])}"
                if len(failed_sheets) > 5:
                    completion_msg += f" and {len(failed_sheets) - 5} more..."

            if omitted_sheets:
                completion_msg += f"\n\n{len(omitted_sheets)} core levels were omitted (different column type):\n"
                completion_msg += ", ".join(omitted_sheets[:5])
                if len(omitted_sheets) > 5:
                    completion_msg += f" and {len(omitted_sheets) - 5} more..."

            if success_count > 0:
                self.batch_progress_text.SetValue(f"{total_areas_created} areas created in {success_count}/{total_sheets} core levels")
                wx.MessageBox(completion_msg, "Success", wx.OK | wx.ICON_INFORMATION)
            else:
                wx.MessageBox("Failed to create areas in any core levels", "Propagate Failed", wx.OK | wx.ICON_ERROR)

        except Exception as e:
            # Restore original sheet on error
            self.parent.sheet_combobox.SetValue(original_sheet)
            on_sheet_selected(self.parent, original_sheet)
            wx.MessageBox(f"Error propagating areas: {str(e)}", "Propagate Failed", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

    def create_area_for_sheet(self, sheet_name, bkg_type, bkg_low, bkg_high, original_peak_name=None):
        """Create area measurement for a specific sheet with given background parameters"""
        import numpy as np

        try:
            # Check if sheet has data
            if sheet_name not in self.parent.Data['Core levels']:
                print(f"Sheet {sheet_name} not found in data")
                return False

            core_level_data = self.parent.Data['Core levels'][sheet_name]

            if 'B.E.' not in core_level_data or 'Raw Data' not in core_level_data:
                print(f"Sheet {sheet_name} missing B.E. or Raw Data")
                return False

            # Set background parameters
            self.parent.background_method = bkg_type

            # Set the offset values to 0
            try:
                offset_h = float(self.offset_h_text.GetValue())
                offset_l = float(self.offset_l_text.GetValue())
            except:
                offset_h = 0.0
                offset_l = 0.0

            # Create background using plot_manager
            # First, temporarily set the range in parent
            self.parent.bg_min_energy = bkg_low
            self.parent.bg_max_energy = bkg_high

            # Initialize background structure if needed
            if 'Background' not in core_level_data:
                core_level_data['Background'] = {}

            core_level_data['Background']['Bkg Type'] = bkg_type
            core_level_data['Background']['Bkg Low'] = float(bkg_low)
            core_level_data['Background']['Bkg High'] = float(bkg_high)
            core_level_data['Background']['Bkg Offset Low'] = float(offset_l)
            core_level_data['Background']['Bkg Offset High'] = float(offset_h)

            # Create the background
            x_values = np.array(core_level_data['B.E.'])
            y_values = np.array(core_level_data['Raw Data'])

            # Use the parent's background creation method
            self.parent.plot_manager.plot_background(self.parent, use_smoothing=False)

            # Now calculate the area
            background = np.array(core_level_data['Background']['Bkg Y'])

            # Mask for the range
            mask = (x_values >= bkg_low) & (x_values <= bkg_high)
            x_range = x_values[mask]
            y_range = y_values[mask]
            bg_range = background[mask]

            if len(x_range) < 3:
                print(f"Insufficient data points in range for {sheet_name}")
                return False

            # Calculate area
            y_minus_bg = y_range - bg_range

            # Sort for proper integration
            sorted_indices = np.argsort(x_range)
            x_sorted = x_range[sorted_indices]
            y_minus_bg_sorted = y_minus_bg[sorted_indices]

            area = np.trapz(y_minus_bg_sorted, x_sorted)

            # Calculate peak parameters
            peak_index = np.argmax(y_minus_bg)
            peak_position = x_range[peak_index]
            peak_height = y_minus_bg[peak_index]

            if peak_height > 0:
                fwhm = 2 * np.sqrt(2 * np.log(2)) * abs(area) / (peak_height * np.sqrt(2 * np.pi))
            else:
                fwhm = 0

            # Round values to .2f
            area = abs(round(area, 2))
            peak_position = round(peak_position, 2)
            peak_height = round(peak_height, 2)
            fwhm = round(fwhm, 2)

            # Add to peak params grid
            grid = self.parent.peak_params_grid
            num_peaks = grid.GetNumberRows() // 2
            peak_letter = chr(65 + num_peaks)

            # Use original peak name if provided, otherwise generate new one
            if original_peak_name:
                area_name = original_peak_name
            else:
                area_name = f"{sheet_name} p{num_peaks + 1}"

            # Add new rows
            grid.AppendRows(2)
            row = num_peaks * 2

            # Set values
            grid.SetCellValue(row, 0, peak_letter)
            grid.SetCellValue(row, 1, area_name)  # Use propagated name
            grid.SetCellValue(row, 2, f"{peak_position:.2f}")
            grid.SetCellValue(row, 3, f"{peak_height:.2f}")
            grid.SetCellValue(row, 4, f"{fwhm:.2f}")
            grid.SetCellValue(row, 5, "0.00")
            grid.SetCellValue(row, 6, f"{area:.2f}")
            grid.SetCellValue(row, 7, "0.00")
            grid.SetCellValue(row, 8, "0.00")
            grid.SetCellValue(row, 9, "0.00")
            grid.SetCellValue(row, 13, "Unfitted")
            grid.SetCellValue(row, 14, bkg_type)
            grid.SetCellValue(row, 15, f"{bkg_low:.2f}")
            grid.SetCellValue(row, 16, f"{bkg_high:.2f}")

            # Set constraints row background color
            for col in range(grid.GetNumberCols()):
                grid.SetCellBackgroundColour(row + 1, col, wx.Colour(200, 245, 228))

            # Set constraint values
            grid.SetCellValue(row + 1, 2, "Fixed")
            grid.SetCellValue(row + 1, 3, "Fixed")
            grid.SetCellValue(row + 1, 4, "Fixed")
            grid.SetCellValue(row + 1, 5, "Fixed")
            grid.SetCellValue(row + 1, 6, "Fixed")
            grid.SetCellValue(row + 1, 7, "Fixed")
            grid.SetCellValue(row + 1, 8, "Fixed")
            grid.SetCellValue(row + 1, 9, "Fixed")

            # Update Data structure
            if 'Fitting' not in core_level_data:
                core_level_data['Fitting'] = {}
            if 'Peaks' not in core_level_data['Fitting']:
                core_level_data['Fitting']['Peaks'] = {}

            core_level_data['Fitting']['Peaks'][area_name] = {
                'Position': peak_position,
                'Height': peak_height,
                'FWHM': fwhm,
                'L/G': 0.00,
                'Area': area,
                'Sigma': 0.00,
                'Gamma': 0.00,
                'Skew': 0.00,
                'Fitting Model': "Unfitted",
                'Bkg Type': bkg_type,
                'Bkg Low': bkg_low,
                'Bkg High': bkg_high,
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

            # Update ratios
            self.parent.update_ratios()

            print(f"Successfully created area '{area_name}' for {sheet_name}: {area:.2f}")
            return True

        except Exception as e:
            print(f"Error creating area for {sheet_name}: {e}")
            import traceback
            traceback.print_exc()
            return False

    def on_propagate_row(self, event):
        """Propagate current row to selected core levels"""
        # TODO: Implement row propagation logic
        wx.MessageBox("Propagate Row to Selected Core Levels - To be implemented", "Info", wx.OK | wx.ICON_INFORMATION)

    def on_measure_all(self, event):
        """Measure areas for all selected core levels"""
        # TODO: Implement batch area measurement logic
        wx.MessageBox("Measure Selected Core Levels - To be implemented", "Info", wx.OK | wx.ICON_INFORMATION)

    def on_help_button(self, event):
        """Open YouTube tutorial video for AreaFit functionality"""
        import webbrowser
        youtube_url = "https://youtu.be/WycrqquAHQU"  # Replace with actual YouTube URL
        try:
            webbrowser.open(youtube_url)
        except Exception as e:
            wx.MessageBox(f"Could not open browser: {e}", "Error", wx.OK | wx.ICON_ERROR)

    def update_tougaard_controls_visibility(self, new_method):
        if "U4-Tougaard" in new_method:
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

        grid.SetCellValue(row + 1, 2, "0:1e3")
        grid.SetCellValue(row + 1, 3, "1:1e7")
        grid.SetCellValue(row + 1, 4, "0.3:3.5")
        grid.SetCellValue(row + 1, 5, "0:0.5")
        grid.SetCellValue(row + 1, 7, "0.1:1")
        grid.SetCellValue(row + 1, 8, "0.1:1")
        grid.SetCellValue(row + 1, 9, "0.01:2")

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
                'Position': "0:1e3",
                'Height': "1:1e7",
                'FWHM': "0.3:3.5",
                'L/G': "0:0.5",
                'Sigma': "0.1:1",
                'Gamma': "0.1:1",
                'Skew': "0.01:2"
            }
        }

        self.parent.ax.fill_between(x_range, bg_range, y_range,
                                    facecolor='lightgreen', alpha=0.5,
                                    label=peak_name)

        self.parent.ax.legend()
        self.parent.peak_params_grid.ForceRefresh()
        # self.on_reset_vlines(self)
        save_state(self.parent)

    def on_switch_vlines(self, event):
        """Switch to next vertical line position - same as TAB key functionality."""
        self.switch_to_next_peak()

    # def on_reset_vlines(self, event):
    #     # Calculate 1/15 and 14/15 positions
    #     if hasattr(self.parent, 'x_values') and len(self.parent.x_values) > 0:
    #         x_min = min(self.parent.x_values)
    #         x_max = max(self.parent.x_values)
    #         x_range = x_max - x_min
    #
    #         new_low = x_min + x_range / 15
    #         new_high = x_min + 14 * x_range / 15
    #
    #         # Update data structure
    #         sheet_name = self.parent.sheet_combobox.GetValue()
    #         if sheet_name in self.parent.Data['Core levels']:
    #             if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
    #                 self.parent.Data['Core levels'][sheet_name]['Background'] = {}
    #             self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Low'] = float(new_low)
    #             self.parent.Data['Core levels'][sheet_name]['Background']['Bkg High'] = float(new_high)
    #
    #         self.parent.bg_min_energy = new_low
    #         self.parent.bg_max_energy = new_high
    #
    #     # Reset vlines and recreate them
    #     self.parent.vline1 = None
    #     self.parent.vline2 = None
    #     if hasattr(self.parent, 'vline1_text'):
    #         self.parent.vline1_text = None
    #     if hasattr(self.parent, 'vline2_text'):
    #         self.parent.vline2_text = None
    #
    #     # Recreate vlines
    #     self.initialize_or_restore_area_vlines()
    #     self.parent.canvas.draw_idle()

        # Update range controls
        self.update_range_controls_from_data()

    def on_toggle_vlines(self, event):
        if self.parent.vline1:
            self.parent.vline1.set_visible(not self.parent.vline1.get_visible())
        if self.parent.vline2:
            self.parent.vline2.set_visible(not self.parent.vline2.get_visible())
        self.parent.canvas.draw_idle()

    def clear_background_between_range(self, bg_low, bg_high):
        """Clear background between specified range - same logic as Fitting_Screen."""
        sheet_name = self.parent.sheet_combobox.GetValue()
        if sheet_name in self.parent.Data['Core levels']:
            core_level_data = self.parent.Data['Core levels'][sheet_name]
            if 'Background' in core_level_data and 'Bkg Y' in core_level_data['Background']:
                # Use the exact same logic as on_clear_between_vlines from Fitting_Screen
                x_values = np.array(self.parent.x_values)
                raw_data = np.array(core_level_data['Raw Data'])
                background = np.array(core_level_data['Background']['Bkg Y'])

                # Create mask for the range to clear
                mask = (x_values >= min(bg_low, bg_high)) & (x_values <= max(bg_low, bg_high))

                # Set background equal to raw data in the specified range (clears the background)
                background[mask] = raw_data[mask]

                # Update the data structure
                core_level_data['Background']['Bkg Y'] = background.tolist()
                self.parent.background = background

                # Replot to show the cleared background
                self.parent.clear_and_replot()

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

        # Generate area name from text field or auto-detect (moved up to check early)
        # Check if this is Survey data
        sheet_name_lower = sheet_name.lower()
        is_survey = any(x in sheet_name_lower for x in ['survey', 'wide'])

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
                grid = self.parent.peak_params_grid
                num_peaks = grid.GetNumberRows() // 2
                area_name = f"{sheet_name} p{num_peaks + 1}"
        else:
            # Use the area name from text field
            if is_survey and not area_name.endswith(' .'):
                area_name = f"{area_name} ."

        # CHECK IF AREA ALREADY EXISTS AND CLEAR PREVIOUS RANGE
        grid = self.parent.peak_params_grid
        existing_row = -1
        for row in range(0, grid.GetNumberRows(), 2):
            if grid.GetCellValue(row, 1) == area_name:
                existing_row = row
                break

        if existing_row >= 0:
            # Area exists - get previous range and clear background there
            prev_bg_low = float(grid.GetCellValue(existing_row, 15))  # Bkg Low
            prev_bg_high = float(grid.GetCellValue(existing_row, 16))  # Bkg High

            print(
                f"Area '{area_name}' already exists. Clearing previous background range {prev_bg_low:.2f} - {prev_bg_high:.2f}")

            # Clear background between previous range using the same logic as Fitting_Screen
            self.clear_background_between_range(prev_bg_low, prev_bg_high)

        # **SPECIAL HANDLING FOR TOUGAARD BACKGROUNDS**
        if "Tougaard" in selected_method:
            # For Tougaard: calculate background for ENTIRE data range, apply only between vLines
            self.handle_tougaard_background_special(sheet_name, selected_method, vline1_x, vline2_x)
        else:
            # For other methods: use normal background creation
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

        # Use existing logic for peak creation/overwriting
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

        # Update ratios after area calculation
        self.parent.update_ratios()

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

        self.parent.plot_config.update_plot_limits(self.parent, sheet_name)
        self.parent.plot_manager.update_legend(self.parent)

        # Restore vlines at their original positions (they were destroyed by clear_and_replot)
        if hasattr(self.parent, 'vline1') and hasattr(self.parent, 'vline2'):
            # Recreate vlines at stored positions (vline1_x and vline2_x were stored earlier)
            self.parent.vline1 = self.parent.ax.axvline(vline1_x, color='r', linestyle='--', alpha=0.7)
            self.parent.vline2 = self.parent.ax.axvline(vline2_x, color='r', linestyle='--', alpha=0.7)

            # Recreate center vline
            center_pos = (vline1_x + vline2_x) / 2
            self.parent.vline_center = self.parent.ax.axvline(center_pos, color='blue', linestyle=':', alpha=0.4, linewidth=1.0)

            # Add text labels back
            self.add_vline_text_labels()

            # Add averaging indicator lines
            if hasattr(self.parent, 'add_averaging_indicator_lines'):
                self.parent.add_averaging_indicator_lines()

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

        save_state(self.parent)

        # Enable survey table if this is a survey plot
        if is_survey and hasattr(self.parent, 'plot_manager'):
            # Enable survey table state
            self.parent.plot_manager.survey_table_state = 1
            # Draw the survey table (this will also hide the legend)
            self.parent.plot_manager.draw_survey_table()
            print(f"Survey table enabled for {sheet_name}")


    def handle_tougaard_background_special(self, sheet_name, selected_method, vline1_x, vline2_x):
        import numpy as np
        from libraries.Peak_Functions import BackgroundCalculations

        # Get FULL data range
        full_x_values = np.array(self.parent.Data['Core levels'][sheet_name]['B.E.'])
        full_raw_data = np.array(self.parent.Data['Core levels'][sheet_name]['Raw Data'])

        # Get current background or create from raw data
        if 'Background' in self.parent.Data['Core levels'][sheet_name] and 'Bkg Y' in \
                self.parent.Data['Core levels'][sheet_name]['Background']:
            current_background = np.array(self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Y'])
        else:
            # Initialize background as raw data
            current_background = full_raw_data.copy()
            if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
                self.parent.Data['Core levels'][sheet_name]['Background'] = {}
            self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = current_background.tolist()

        # Define vLine range
        range_min = min(vline1_x, vline2_x)
        range_max = max(vline1_x, vline2_x)

        # Create mask for vLine range
        vline_mask = (full_x_values >= range_min) & (full_x_values <= range_max)

        # STEP 1: Remove any existing background in the vLine range
        current_background[vline_mask] = full_raw_data[vline_mask]

        # STEP 2: Calculate Tougaard background
        print(f"Calculating {selected_method} background...")

        try:
            if selected_method == "U4-Tougaard":
                # U4-Tougaard: Use ENTIRE data range
                print("U4-Tougaard: Using entire data range")
                full_tougaard_bg = BackgroundCalculations.calculate_tougaard_background(
                    full_x_values, full_raw_data, sheet_name, self.parent)
            elif selected_method == "U2-Tougaard":
                # U2-Tougaard: Use ONLY vLine range data
                print(f"U2-Tougaard: Using only vLine range {range_min:.2f} - {range_max:.2f}")

                # Extract only the data between vLines
                vline_x_values = full_x_values[vline_mask]
                vline_raw_data = full_raw_data[vline_mask]

                print(f"DEBUG: vLine data points: {len(vline_x_values)} (from {len(full_x_values)} total)")

                # Calculate U2-Tougaard background using only vLine range data
                vline_range = (vline1_x, vline2_x)
                vline_tougaard_bg = BackgroundCalculations.calculate_u2_tougaard_background(
                    vline_x_values, vline_raw_data, sheet_name, self.parent, vline_range)

                # Create full background array (start with current background)
                full_tougaard_bg = current_background.copy()

                # Insert the calculated vLine background into the full array
                full_tougaard_bg[vline_mask] = vline_tougaard_bg

            elif selected_method == "2x U4-Tougaard":
                full_tougaard_bg = BackgroundCalculations.calculate_double_tougaard_background(
                    full_x_values, full_raw_data, sheet_name, self.parent)
            elif selected_method == "3x U4-Tougaard":
                full_tougaard_bg = BackgroundCalculations.calculate_triple_tougaard_background(
                    full_x_values, full_raw_data, sheet_name, self.parent)
            else:
                print(f"Unknown Tougaard method: {selected_method}")
                # Fallback to normal background
                self.parent.plot_manager.plot_background(self.parent)
                return

            # STEP 3: Apply Tougaard background ONLY in vLine range
            current_background[vline_mask] = full_tougaard_bg[vline_mask]

            # Update data structure
            self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = current_background.tolist()
            self.parent.background = current_background

            print(f"Tougaard background applied between {range_min:.2f} - {range_max:.2f} eV")

        except Exception as e:
            print(f"Error calculating Tougaard background: {e}")
            # Fallback to normal background
            self.parent.plot_manager.plot_background(self.parent)

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
        """Remove last background/area and properly reinitialize vLines"""
        # Store vLine positions before removal
        vline1_x = None
        vline2_x = None
        if self.parent.vline1 is not None:
            vline1_x = self.parent.vline1.get_xdata()[0]
        if self.parent.vline2 is not None:
            vline2_x = self.parent.vline2.get_xdata()[0]

        # Call the remove peak function
        # from libraries.Peak_Functions import remove_peak
        remove_peak(self.parent)

        # Update atomic concentrations after removing peak
        self.parent.update_ratios()

        # Reinitialize vLines after remove_peak calls clear_and_replot
        if hasattr(self, 'initialize_or_restore_area_vlines'):
            self.initialize_or_restore_area_vlines()

        # Restore original positions if they existed
        if vline1_x is not None and vline2_x is not None:
            if self.parent.vline1 is not None:
                self.parent.vline1.set_xdata([float(f"{vline1_x:.2f}"), float(f"{vline1_x:.2f}")])
            if self.parent.vline2 is not None:
                self.parent.vline2.set_xdata([float(f"{vline2_x:.2f}"), float(f"{vline2_x:.2f}")])

            # Update text labels and range controls
            self.update_vline_text_labels()
            self.update_range_controls_from_data()

        # Reset mouse interaction system
        if hasattr(self.parent, 'mouse_handler'):
            self.parent.mouse_handler.cleanup_vline_handlers()
            self.parent.moving_vline = None

        save_state(self.parent)

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

    def on_tougaard_model(self, event):
        from libraries.ToolsMenu.Fitting_Screen import TougaardFitWindow
        TougaardFitWindow(self).Show()

    def on_close(self, event):
        """Close area screen and ensure complete cleanup."""
        # Clear peak highlighting before closing
        self.clear_peak_grid_highlights()

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


        grid.SetCellValue(row + 1, 2, "0:1e3")
        grid.SetCellValue(row + 1, 3, "1:1e7")
        grid.SetCellValue(row + 1, 4, "0.3:3.5")
        grid.SetCellValue(row + 1, 5, "0:0.5")
        grid.SetCellValue(row + 1, 7, "0:1:1")
        grid.SetCellValue(row + 1, 8, "0.1:1")
        grid.SetCellValue(row + 1, 9, "0.01:2")

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
                'Position': "0:1e3",
                'Height': "1:1e7",
                'FWHM': "0.3:3.5",
                'L/G': "0:0.5",
                'Sigma': "0.1:1",
                'Gamma': "0.1:1",
                'Skew': "0.01:2"
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
        """Export results to main results grid"""
        save_state(self.parent)
        self.parent.export_results()

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

        # Clean up center vline
        if hasattr(self.parent, 'vline_center') and self.parent.vline_center is not None:
            try:
                self.parent.vline_center.remove()
            except:
                pass
            self.parent.vline_center = None
        if hasattr(self.parent, 'vline_center_text') and self.parent.vline_center_text is not None:
            try:
                self.parent.vline_center_text.remove()
            except:
                pass
            self.parent.vline_center_text = None

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

        # Convert BE positions to display positions
        vline1_display = self.parent.convert_energy_for_display(vline1_pos)
        vline2_display = self.parent.convert_energy_for_display(vline2_pos)

        # Create new vlines - these will be draggable
        self.parent.vline1 = self.parent.ax.axvline(vline1_display, color='r', linestyle='--', alpha=0.7)
        self.parent.vline2 = self.parent.ax.axvline(vline2_display, color='r', linestyle='--', alpha=0.7)

        # Create center vline
        center_pos = (vline1_display + vline2_display) / 2
        self.parent.vline_center = self.parent.ax.axvline(center_pos, color='blue', linestyle=':', alpha=0.4, linewidth=1)

        # Add text labels
        self.add_vline_text_labels()

        # Add averaging indicator lines
        if hasattr(self.parent, 'add_averaging_indicator_lines'):
            self.parent.add_averaging_indicator_lines()

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
            y_range = y_max - y_min

            # Get vline positions and round to 2 digits
            vline1_x = round(self.parent.vline1.get_xdata()[0], 2)
            vline2_x = round(self.parent.vline2.get_xdata()[0], 2)

            # Determine which is high BE and which is low BE
            if vline1_x > vline2_x:
                high_be_x, low_be_x = vline1_x, vline2_x
                high_be_vline, low_be_vline = 'vline1', 'vline2'
            else:
                high_be_x, low_be_x = vline2_x, vline1_x
                high_be_vline, low_be_vline = 'vline2', 'vline1'

            # Position high BE text at 90% of Y axis
            high_be_text_y = y_min + 0.9 * y_range
            # Position low BE text 10% lower (80% of Y axis)
            low_be_text_y = y_min + 0.8 * y_range

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

            # Create text labels with appropriate heights
            if high_be_vline == 'vline1':
                self.parent.vline1_text = self.parent.ax.text(vline1_x, high_be_text_y, f'{vline1_x:.2f}',
                                                              ha='center', va='center',
                                                              color='grey', fontsize=10,
                                                              bbox=dict(boxstyle='round,pad=0.2', facecolor='white',
                                                                        alpha=0.8))
                self.parent.vline2_text = self.parent.ax.text(vline2_x, low_be_text_y, f'{vline2_x:.2f}',
                                                              ha='center', va='center',
                                                              color='grey', fontsize=10,
                                                              bbox=dict(boxstyle='round,pad=0.2', facecolor='white',
                                                                        alpha=0.8))

                # Add center vline text label
                if hasattr(self.parent, 'vline_center') and self.parent.vline_center is not None:
                    center_x = (vline1_x + vline2_x) / 2
                    center_text_y = y_min + 0.7 * y_range  # Position at 70% of Y axis
                    self.parent.vline_center_text = self.parent.ax.text(center_x, center_text_y,
                                                                        f'{center_x:.2f}', ha='center', va='bottom',
                                                                        bbox=dict(boxstyle='round,pad=0.2', facecolor='lightgreen', alpha=0.8))
            else:
                self.parent.vline1_text = self.parent.ax.text(vline1_x, low_be_text_y, f'{vline1_x:.2f}',
                                                              ha='center', va='center',
                                                              color='grey', fontsize=10,
                                                              bbox=dict(boxstyle='round,pad=0.2', facecolor='white',
                                                                        alpha=0.8))
                self.parent.vline2_text = self.parent.ax.text(vline2_x, high_be_text_y, f'{vline2_x:.2f}',
                                                              ha='center', va='center',
                                                              color='grey', fontsize=10,
                                                              bbox=dict(boxstyle='round,pad=0.2', facecolor='white',
                                                                        alpha=0.8))
                # Add center vline text label
                if hasattr(self.parent, 'vline_center') and self.parent.vline_center is not None:
                    center_x = (vline1_x + vline2_x) / 2
                    center_text_y = y_min + 0.7 * y_range  # Position at 70% of Y axis
                    self.parent.vline_center_text = self.parent.ax.text(center_x, center_text_y,
                                                                        f'{center_x:.2f}', ha='center', va='bottom',
                                                                        bbox=dict(boxstyle='round,pad=0.2', facecolor='lightgreen', alpha=0.8))

    def update_vline_text_labels(self):
        """Update the text labels when vlines are moved."""
        if (self.parent.vline1 is not None and self.parent.vline2 is not None and
                hasattr(self.parent, 'vline1_text') and self.parent.vline1_text is not None and
                hasattr(self.parent, 'vline2_text') and self.parent.vline2_text is not None):
            # Get y-axis range for positioning
            y_min, y_max = self.parent.ax.get_ylim()
            y_range = y_max - y_min

            # Get current vline positions and round to 2 digits
            vline1_x = round(self.parent.vline1.get_xdata()[0], 2)
            vline2_x = round(self.parent.vline2.get_xdata()[0], 2)

            # Determine which is high BE and which is low BE
            if vline1_x > vline2_x:
                high_be_text_y = y_min + 0.9 * y_range  # High BE at 90%
                low_be_text_y = y_min + 0.8 * y_range  # Low BE at 80%

                # vline1 is high BE, vline2 is low BE
                self.parent.vline1_text.set_position((vline1_x, high_be_text_y))
                self.parent.vline1_text.set_text(f'{vline1_x:.2f}')
                self.parent.vline2_text.set_position((vline2_x, low_be_text_y))
                self.parent.vline2_text.set_text(f'{vline2_x:.2f}')
            else:
                high_be_text_y = y_min + 0.9 * y_range  # High BE at 90%
                low_be_text_y = y_min + 0.8 * y_range  # Low BE at 80%

                # vline2 is high BE, vline1 is low BE
                self.parent.vline1_text.set_position((vline1_x, low_be_text_y))
                self.parent.vline1_text.set_text(f'{vline1_x:.2f}')
                self.parent.vline2_text.set_position((vline2_x, high_be_text_y))
                self.parent.vline2_text.set_text(f'{vline2_x:.2f}')

            # Auto-detect area name when vlines move
            self.auto_detect_area_name(vline1_x, vline2_x)

            # Update core level list if open
            self.update_core_level_list_if_open()

            # Update center vline position when other vlines move
            self.update_center_vline_position()

    def update_center_vline_position(self):
        """Update center vline position based on vline1 and vline2 positions."""
        if (hasattr(self.parent, 'vline_center') and self.parent.vline_center is not None and
                self.parent.vline1 is not None and self.parent.vline2 is not None):

            vline1_x = self.parent.vline1.get_xdata()[0]
            vline2_x = self.parent.vline2.get_xdata()[0]
            center_x = (vline1_x + vline2_x) / 2

            self.parent.vline_center.set_xdata([center_x, center_x])

            # Update center text label if it exists
            if hasattr(self.parent, 'vline_center_text') and self.parent.vline_center_text is not None:
                y_pos = self.parent.vline_center_text.get_position()[1]
                self.parent.vline_center_text.set_position((center_x, y_pos))
                self.parent.vline_center_text.set_text(f'{center_x:.2f}')

    def show_hide_vlines(self):
        """Show/hide vlines based on current screen state and zoom/drag modes."""
        # Hide vlines ONLY if dragging (NOT during zoom)
        if self.drag_mode:  # REMOVED: or self.zoom_mode
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

        # During zoom_mode: keep vLines visible but make them non-draggable
        if self.zoom_mode:
            # Make vLines non-draggable by removing mouse event connections
            if hasattr(self, 'mouse_handler'):
                self.mouse_handler.cleanup_vline_handlers()
            # But keep them visible - continue with normal visibility logic below

        # Rest of existing visibility logic...
        fitting_screen_active = (hasattr(self, 'fitting_window') and
                                 self.fitting_window is not None and
                                 hasattr(self, 'background_tab_selected') and
                                 self.background_tab_selected)

        area_screen_active = (hasattr(self, 'background_window') and
                              self.background_window is not None and
                              hasattr(self, 'area_tab_selected') and
                              self.area_tab_selected)

        background_lines_visible = fitting_screen_active or area_screen_active

        # Set visibility for background vlines
        if self.vline1 is not None:
            self.vline1.set_visible(background_lines_visible)
        if self.vline2 is not None:
            self.vline2.set_visible(background_lines_visible)
        if hasattr(self, 'vline1_text') and self.vline1_text is not None:
            self.vline1_text.set_visible(background_lines_visible)
        if hasattr(self, 'vline2_text') and self.vline2_text is not None:
            self.vline2_text.set_visible(background_lines_visible)
        if hasattr(self, 'vline_center') and self.vline_center is not None:
            self.vline_center.set_visible(background_lines_visible)
        if hasattr(self, 'vline_center_text') and self.vline_center_text is not None:
            self.vline_center_text.set_visible(background_lines_visible)

        # Noise lines visibility
        noise_lines_visible = (self.noise_analysis_window is not None and
                               hasattr(self, 'noise_tab_selected') and
                               self.noise_tab_selected)
        if self.vline3 is not None:
            self.vline3.set_visible(noise_lines_visible)
        if self.vline4 is not None:
            self.vline4.set_visible(noise_lines_visible)

        self.canvas.draw_idle()

    def format_text_control_to_2f(self, text_ctrl):
        """Format text control value to always show 2 decimal places."""
        try:
            current_value = text_ctrl.GetValue()
            if current_value.strip():  # Only format if not empty
                float_value = float(current_value)
                formatted_value = f"{float_value:.2f}"
                text_ctrl.SetValue(formatted_value)
        except ValueError:
            # If invalid number, reset to 0.00
            text_ctrl.SetValue("0.00")

    def on_text_control_focus_lost(self, event):
        """Handle focus lost event for any text control that needs 2f formatting."""
        text_ctrl = event.GetEventObject()
        self.format_text_control_to_2f(text_ctrl)
        event.Skip()  # Allow normal processing

    def on_text_control_enter(self, event):
        """Handle Enter key for any text control that needs 2f formatting."""
        text_ctrl = event.GetEventObject()
        self.format_text_control_to_2f(text_ctrl)

        # Call the original handler based on which control triggered this
        if text_ctrl == self.min_range_text:
            self.on_min_range_change(event)
        elif text_ctrl == self.max_range_text:
            self.on_max_range_change(event)
        elif text_ctrl == self.offset_h_text:
            self.on_offset_changed(event)
        elif text_ctrl == self.offset_l_text:
            self.on_offset_changed(event)

    def create_elements_database_OLD(self):
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

    def create_elements_database(self):
        """Create a comprehensive elements database with binding energy ranges."""
        return {
            # Light elements
            'C': {'1s': (284.00, 291.00), '2s': (20.00, 25.00)},
            'N': {'1s': (397.00, 407.00)},
            'O': {'1s': (528.00, 538.00)},
            'F': {'1s': (684.00, 688.00)},
            'Na': {'1s': (1071.00, 1072.00), 'kll': (491.00, 501.00)},
            'Mg': {'1s': (1303.00, 1304.00), '2s': (88.00, 90.00), '2p': (50.00, 51.00)},
            'Al': {'1s': (1559.00, 1560.00), '2s': (117.00, 119.00), '2p': (72.00, 75.00)},
            'Si': {'1s': (1838.00, 1840.00), '2s': (149.00, 151.00), '2p': (99.00, 103.00)},
            'P': {'1s': (2145.00, 2147.00), '2s': (189.00, 191.00), '2p': (129.00, 136.00)},
            'S': {'1s': (2470.00, 2473.00), '2s': (228.00, 230.00), '2p': (162.00, 170.00)},
            'Cl': {'1s': (2822.00, 2824.00), '2s': (270.00, 272.00), '2p': (198.00, 202.00)},
            'K': {'1s': (3607.00, 3609.00), '2s': (378.00, 380.00), '2p': (293.00, 297.00), '3s': (34.00, 35.00),
                  '3p': (18.00, 19.00)},
            'Ca': {'1s': (4038.00, 4040.00), '2s': (438.00, 440.00), '2p': (346.00, 350.00), '3s': (44.00, 45.00),
                   '3p': (25.00, 26.00)},

            # Transition metals
            'Ti': {'1s': (4964.00, 4966.00), '2s': (563.00, 565.00), '2p': (455.00, 465.00), '3s': (60.00, 61.00),
                   '3p': (37.00, 38.00)},
            'Cr': {'1s': (5987.00, 5989.00), '2s': (694.00, 696.00), '2p': (574.00, 584.00), '3s': (84.00, 85.00),
                   '3p': (43.00, 45.00)},
            'Mn': {'1s': (6539.00, 6541.00), '2s': (769.00, 771.00), '2p': (639.00, 651.00), '3s': (82.00, 84.00),
                   '3p': (47.00, 49.00)},
            'Fe': {'1s': (7112.00, 7114.00), '2s': (844.00, 846.00), '2p': (706.00, 720.00), '3s': (91.00, 93.00),
                   '3p': (52.00, 54.00)},
            'Co': {'1s': (7709.00, 7711.00), '2s': (925.00, 927.00), '2p': (778.00, 793.00), '3s': (101.00, 103.00),
                   '3p': (58.00, 60.00)},
            'Ni': {'1s': (8333.00, 8335.00), '2s': (1008.00, 1010.00), '2p': (852.00, 870.00), '3s': (110.00, 112.00),
                   '3p': (66.00, 68.00)},
            'Cu': {'1s': (8979.00, 8981.00), '2s': (1096.00, 1098.00), '2p': (932.00, 953.00), '3s': (122.00, 124.00),
                   '3p': (75.00, 77.00)},
            'Zn': {'2s': (1193.00, 1195.00), '2p': (1021.00, 1045.00), '3s': (139.00, 141.00), '3p': (88.00, 90.00)},
            'Ga': {'2s': (1298.00, 1300.00), '2p': (1116.00, 1120.00), '3s': (159.00, 161.00), '3p': (103.00, 105.00), '3d': (18.00, 20.00)},
            'Ge': {'2s': (1414.00, 1416.00), '2p': (1217.00, 1221.00), '3s': (180.00, 182.00), '3p': (120.00, 122.00), '3d': (29.00, 31.00)},
            'As': {'2s': (1527.00, 1529.00), '2p': (1323.00, 1327.00), '3s': (204.00, 206.00), '3p': (141.00, 143.00), '3d': (41.00, 43.00)},
            'Se': {'2s': (1652.00, 1654.00), '2p': (1433.00, 1437.00), '3s': (229.00, 231.00), '3p': (160.00, 162.00),
                   '3d': (55.00, 59.00)},
            'Br': {'2s': (1782.00, 1784.00), '2p': (1550.00, 1554.00), '3s': (257.00, 259.00), '3p': (182.00, 184.00),
                   '3d': (67.00, 71.00)},
            'Sr': {'2s': (2216.00, 2218.00), '2p': (1940.00, 1944.00), '3s': (358.00, 360.00), '3p': (269.00, 271.00),
                   '3d': (132.00, 136.00)},
            'Zr': {'2s': (2532.00, 2534.00), '2p': (2223.00, 2227.00), '3s': (430.00, 432.00), '3p': (343.00, 345.00),
                   '3d': (177.00, 181.00)},
            'Mo': {'2s': (2866.00, 2868.00), '2p': (2520.00, 2524.00), '3s': (506.00, 508.00), '3p': (412.00, 414.00),
                   '3d': (226.00, 230.00)},
            'Ag': {'2s': (3806.00, 3808.00), '2p': (3351.00, 3355.00), '3s': (719.00, 721.00), '3p': (603.00, 605.00),
                   '3d': (366.00, 370.00)},

            # Heavy elements (removed high energy orbitals > 9000eV)
            'Cs': {'2s': (5714.00, 5716.00), '2p': (5359.00, 5361.00), '3s': (1217.00, 1219.00),
                   '3p': (1065.00, 1067.00), '3d': (726.00, 740.00)},
            'Ba': {'2s': (5989.00, 5991.00), '2p': (5624.00, 5626.00), '3s': (1293.00, 1295.00),
                   '3p': (1137.00, 1139.00), '3d': (780.00, 795.00)},

            # Lanthanides (removed high energy orbitals > 9000eV)
            'La': {'2s': (6266.00, 6268.00), '2p': (5891.00, 5893.00), '3s': (1362.00, 1364.00),
                   '3p': (1209.00, 1211.00), '3d': (836.00, 851.00), '4d': (98.00, 103.00)},
            'Ce': {'2s': (6549.00, 6551.00), '2p': (6164.00, 6166.00), '3s': (1436.00, 1438.00),
                   '3p': (1274.00, 1276.00), '3d': (884.00, 901.00), '4d': (109.00, 114.00)},
            'Pr': {'2s': (6835.00, 6837.00), '2p': (6440.00, 6442.00), '3s': (1511.00, 1513.00),
                   '3p': (1337.00, 1339.00), '3d': (932.00, 950.00), '4d': (115.00, 120.00)},
            'Nd': {'2s': (7126.00, 7128.00), '2p': (6722.00, 6724.00), '3s': (1575.00, 1577.00),
                   '3p': (1403.00, 1405.00), '3d': (981.00, 1000.00), '4d': (121.00, 126.00)},
            'Pm': {'2s': (7428.00, 7430.00), '2p': (7013.00, 7015.00), '3s': (1471.00, 1473.00),
                   '3p': (1357.00, 1359.00), '3d': (1027.00, 1048.00)},
            'Sm': {'2s': (7737.00, 7739.00), '2p': (7312.00, 7314.00), '3s': (1723.00, 1725.00),
                   '3p': (1541.00, 1543.00), '3d': (1081.00, 1101.00)},
            'Eu': {'2s': (8052.00, 8054.00), '2p': (7617.00, 7619.00), '3s': (1800.00, 1802.00),
                   '3p': (1614.00, 1616.00), '3d': (1131.00, 1151.00)},
            'Gd': {'2s': (8376.00, 8378.00), '2p': (7930.00, 7932.00), '3s': (1881.00, 1883.00),
                   '3p': (1688.00, 1690.00), '3d': (1185.00, 1205.00)},
            'Tb': {'2s': (8708.00, 8710.00), '2p': (8252.00, 8254.00), '3s': (1968.00, 1970.00),
                   '3p': (1768.00, 1770.00), '3d': (1241.00, 1261.00)},

            # Heavy transition metals (< 9000eV only)
            'Hf': {'3s': (2601.00, 2603.00), '3p': (2365.00, 2367.00), '3d': (1716.00, 1740.00), '4s': (660.00, 665.00),
                   '4p': (380.00, 385.00), '4d': (220.00, 225.00), '4f': (14.00, 17.00)},
            'Ta': {'3s': (2708.00, 2710.00), '3p': (2469.00, 2471.00), '3d': (1793.00, 1817.00), '4s': (708.00, 713.00),
                   '4p': (405.00, 410.00), '4d': (230.00, 235.00), '4f': (22.00, 25.00)},
            'W': {'3s': (2820.00, 2822.00), '3p': (2575.00, 2577.00), '3d': (1872.00, 1896.00), '4s': (756.00, 761.00),
                  '4p': (423.00, 428.00), '4d': (243.00, 248.00), '4f': (31.00, 34.00)},
            'Re': {'3s': (2932.00, 2934.00), '3p': (2682.00, 2684.00), '3d': (1949.00, 1973.00), '4s': (625.00, 630.00),
                   '4p': (445.00, 450.00), '4d': (255.00, 260.00), '4f': (40.00, 43.00)},
            'Os': {'3s': (3049.00, 3051.00), '3p': (2792.00, 2794.00), '3d': (2031.00, 2055.00), '4s': (658.00, 663.00),
                   '4p': (470.00, 475.00), '4d': (278.00, 283.00), '4f': (51.00, 54.00)},
            'Ir': {'3s': (3174.00, 3176.00), '3p': (2909.00, 2911.00), '3d': (2116.00, 2140.00), '4s': (691.00, 696.00),
                   '4p': (495.00, 500.00), '4d': (311.00, 316.00), '4f': (61.00, 64.00)},
            'Pt': {'3s': (3296.00, 3298.00), '3p': (3027.00, 3029.00), '3d': (2202.00, 2226.00), '4s': (725.00, 730.00),
                   '4p': (519.00, 524.00), '4d': (314.00, 319.00), '4f': (71.00, 74.00)},
            'Au': {'3s': (3425.00, 3427.00), '3p': (3148.00, 3150.00), '3d': (2291.00, 2315.00), '4s': (760.00, 765.00),
                   '4p': (546.00, 551.00), '4d': (335.00, 340.00), '4f': (84.00, 87.00)},
            'Hg': {'3s': (3562.00, 3564.00), '3p': (3279.00, 3281.00), '3d': (2385.00, 2409.00), '4s': (800.00, 805.00),
                   '4p': (572.00, 577.00), '4d': (358.00, 363.00), '4f': (101.00, 104.00)},
            'Tl': {'3s': (3704.00, 3706.00), '3p': (3416.00, 3418.00), '3d': (2485.00, 2509.00), '4s': (846.00, 851.00),
                   '4p': (608.00, 613.00), '4d': (386.00, 391.00), '4f': (118.00, 121.00)},
            'Pb': {'3s': (3851.00, 3853.00), '3p': (3554.00, 3556.00), '3d': (2586.00, 2610.00), '4s': (894.00, 899.00),
                   '4p': (644.00, 649.00), '4d': (412.00, 417.00), '4f': (137.00, 140.00)},
            'Bi': {'3s': (4002.00, 4004.00), '3p': (3696.00, 3698.00), '3d': (2688.00, 2712.00), '4s': (938.00, 943.00),
                   '4p': (680.00, 685.00), '4d': (440.00, 445.00), '4f': (157.00, 160.00)},

            # Additional common elements (< 9000eV only)
            'Rb': {'2s': (2692.00, 2694.00), '2p': (2472.00, 2474.00), '3s': (484.00, 486.00), '3p': (394.00, 396.00),
                   '3d': (111.00, 113.00)},
            'Y': {'2s': (2373.00, 2375.00), '2p': (2156.00, 2158.00), '3s': (392.00, 394.00), '3p': (298.00, 300.00),
                  '3d': (158.00, 160.00)},
            'Nb': {'2s': (2698.00, 2700.00), '2p': (2465.00, 2467.00), '3s': (468.00, 470.00), '3p': (360.00, 362.00),
                   '3d': (202.00, 204.00)},
            'Ru': {'2s': (2967.00, 2969.00), '2p': (2838.00, 2840.00), '3s': (586.00, 588.00), '3p': (461.00, 463.00),
                   '3d': (280.00, 284.00)},
            'Rh': {'2s': (3146.00, 3148.00), '2p': (3004.00, 3006.00), '3s': (628.00, 630.00), '3p': (496.00, 498.00),
                   '3d': (307.00, 311.00)},
            'Pd': {'2s': (3330.00, 3332.00), '2p': (3173.00, 3175.00), '3s': (671.00, 673.00), '3p': (532.00, 534.00),
                   '3d': (335.00, 339.00)},
            'Cd': {'2s': (4018.00, 4020.00), '2p': (3727.00, 3729.00), '3s': (772.00, 774.00), '3p': (618.00, 620.00),
                   '3d': (405.00, 409.00)},
            'In': {'2s': (4238.00, 4240.00), '2p': (3938.00, 3940.00), '3s': (827.00, 829.00), '3p': (665.00, 667.00),
                   '3d': (444.00, 448.00)},
            'Sn': {'2s': (4465.00, 4467.00), '2p': (4156.00, 4158.00), '3s': (884.00, 886.00), '3p': (714.00, 716.00),
                   '3d': (485.00, 489.00)},
            'Sb': {'2s': (4698.00, 4700.00), '2p': (4380.00, 4382.00), '3s': (946.00, 948.00), '3p': (766.00, 768.00),
                   '3d': (528.00, 532.00)},
            'Te': {'2s': (4939.00, 4941.00), '2p': (4612.00, 4614.00), '3s': (1006.00, 1008.00), '3p': (820.00, 822.00),
                   '3d': (573.00, 577.00)},
            'I': {'2s': (5188.00, 5190.00), '2p': (4852.00, 4854.00), '3s': (1072.00, 1074.00), '3p': (875.00, 877.00),
                  '3d': (619.00, 623.00)},
            'Xe': {'2s': (5453.00, 5455.00), '2p': (5107.00, 5109.00), '3s': (1148.00, 1150.00), '3p': (940.00, 942.00),
                   '3d': (670.00, 674.00)}
        }

    def get_detected_elements(self):
        """Get list of elements already detected in the current sheet."""
        detected_elements = set()

        # Check existing peaks in the grid
        grid = self.parent.peak_params_grid
        for row in range(0, grid.GetNumberRows(), 2):
            peak_name = grid.GetCellValue(row, 1)
            if peak_name:
                # Extract element from peak name (e.g., "C1s" -> "C", "Ti2p" -> "Ti")
                element = ''.join([c for c in peak_name if
                                   c.isupper() or (c.islower() and peak_name[peak_name.index(c) - 1].isupper())])
                if element:
                    detected_elements.add(element)

        return detected_elements

    def get_priority_orbital(self, element, overlapping_orbitals):
        """Determine the priority orbital for element detection with dynamic priority for detected elements."""

        # Get already detected elements - these get priority 1 for ALL their peaks
        detected_elements = self.get_detected_elements()

        # Priority Level 1: Primary core levels + ALL peaks from detected elements
        priority_1 = {
            'C': '1s', 'N': '1s', 'O': '1s', 'F': '1s', 'Na': ['1s', 'kll'],'Mg': '1s',
            'K': '2p', 'Ca': '2p', 'Si': '2p', 'P': '2p', 'S': '2p', 'Cl': '2p', 'I': '3d'
        }

        # Priority Level 2: Transition metals and common elements
        priority_2 = {
            'Ti': '2p', 'V': '2p', 'Cr': '2p', 'Mn': '2p', 'Fe': '2p', 'Co': '2p',
            'Ni': '2p', 'Cu': '2p', 'Zn': '2p', 'Ag': '3d', 'Au': '4f', 'In': '3d', 'Sn': '3d',
            'Na': 'kll'
        }

        # Priority Level 3: Less common elements
        priority_3 = {
            'Sb': '3d', 'Ba': '3d', 'Ta': '4f', 'W': '4f', 'Re': '4f', 'Ir': '4f', 'Pt': '4f',
            'La': '3d', 'Ce': '3d', 'Pr': '3d', 'Sm': '3d', 'Nd': '3d', 'Eu': '3d',
            'Gd': '3d', 'Dy': '3d', 'Cs': '3d'
        }

        # DYNAMIC PRIORITY: If element is already detected, ALL its orbitals get priority 1
        if element in detected_elements:
            return overlapping_orbitals[0]  # Return any orbital since all get priority 1

        # Check static priorities in order
        if element in priority_1 and priority_1[element] in overlapping_orbitals:
            return priority_1[element]
        elif element in priority_2 and priority_2[element] in overlapping_orbitals:
            return priority_2[element]
        elif element in priority_3 and priority_3[element] in overlapping_orbitals:
            return priority_3[element]
        else:
            # Default selection based on general rules
            if '1s' in overlapping_orbitals:
                return '1s'
            elif '2p' in overlapping_orbitals:
                return '2p'
            elif '3d' in overlapping_orbitals:
                return '3d'
            elif '4f' in overlapping_orbitals:
                return '4f'
            else:
                return overlapping_orbitals[0]  # Return first available

    def get_element_priority_level(self, element, orbital):
        """Get the priority level (1=highest, 4=lowest) for an element-orbital combination."""

        # Get already detected elements - these get priority 1 for ALL their peaks
        detected_elements = self.get_detected_elements()

        # DYNAMIC PRIORITY: If element is already detected, ALL its orbitals get priority 1
        if element in detected_elements:
            return 1

        # Priority Level 1: Primary core levels (highest priority)
        priority_1 = {
            'C': '1s', 'N': '1s', 'O': '1s', 'F': '1s', 'Na': ['1s', 'kll'], 'Mg': '1s',
            'K': '2p', 'Ca': '2p', 'Si': '2p', 'P': '2p', 'S': '2p', 'Cl': '2p', 'I': '3d'
        }

        # Priority Level 2: Transition metals and common elements
        priority_2 = {
            'Ti': '2p', 'V': '2p', 'Cr': '2p', 'Mn': '2p', 'Fe': '2p', 'Co': '2p',
            'Ni': '2p', 'Cu': '2p', 'Zn': '2p', 'Ag': '3d', 'Au': '4f', 'In': '3d', 'Sn': '3d',
            'Na': 'kll'
        }

        # Priority Level 3: Less common elements
        priority_3 = {
            'Sb': '3d', 'Ba': '3d', 'Ta': '4f', 'W': '4f', 'Re': '4f', 'Ir': '4f', 'Pt': '4f',
            'La': '3d', 'Ce': '3d', 'Pr': '3d', 'Sm': '3d', 'Nd': '3d', 'Eu': '3d',
            'Gd': '3d', 'Dy': '3d', 'Cs': '3d'
        }

        if element in priority_1 and priority_1[element] == orbital:
            return 1
        elif element in priority_2 and priority_2[element] == orbital:
            return 2
        elif element in priority_3 and priority_3[element] == orbital:
            return 3
        else:
            return 4  # Lowest priority for non-specified combinations

    def update_range_controls_from_data_without_events(self, bkg_low, bkg_high):
        """Update range controls without triggering change events."""
        if not (hasattr(self, 'min_range_text') and hasattr(self, 'max_range_text')):
            return
        if not (self.min_range_text and self.max_range_text):
            return

        try:
            # Temporarily unbind Enter and focus events
            self.min_range_text.Unbind(wx.EVT_TEXT_ENTER)
            self.max_range_text.Unbind(wx.EVT_TEXT_ENTER)
            self.min_range_text.Unbind(wx.EVT_KILL_FOCUS)
            self.max_range_text.Unbind(wx.EVT_KILL_FOCUS)

            # Update values with 2f formatting
            min_val = min(bkg_low, bkg_high)
            max_val = max(bkg_low, bkg_high)

            self.min_range_text.ChangeValue(f"{min_val:.2f}")
            self.max_range_text.ChangeValue(f"{max_val:.2f}")

            # Rebind events
            self.min_range_text.Bind(wx.EVT_TEXT_ENTER, self.on_min_range_enter)
            self.max_range_text.Bind(wx.EVT_TEXT_ENTER, self.on_max_range_enter)
            self.min_range_text.Bind(wx.EVT_KILL_FOCUS, self.on_min_range_enter)
            self.max_range_text.Bind(wx.EVT_KILL_FOCUS, self.on_max_range_enter)

        except (RuntimeError, AttributeError):
            pass

    def on_min_range_enter(self, event):
        """Handle min range change only on Enter or focus lost."""
        if self.updating_range_controls:
            return

        try:
            new_min = float(self.min_range_text.GetValue())
            max_val = float(self.max_range_text.GetValue())

            # Format to 2 decimal places
            new_min = round(new_min, 2)
            max_val = round(max_val, 2)

            # Auto-swap if min > max
            if new_min > max_val:
                self.updating_range_controls = True
                self.min_range_text.SetValue(f"{max_val:.2f}")
                self.max_range_text.SetValue(f"{new_min:.2f}")
                self.updating_range_controls = False
                new_min = max_val
            else:
                # Format the current control to 2f
                self.min_range_text.SetValue(f"{new_min:.2f}")

            # Update vLine position
            if self.parent.vline1 is not None:
                self.parent.vline1.set_xdata([new_min, new_min])
                self.update_vline_text_labels()
                self.parent.canvas.draw_idle()

        except ValueError:
            # Reset to 0.00 if invalid
            self.min_range_text.SetValue("0.00")

    def on_max_range_enter(self, event):
        """Handle max range change only on Enter or focus lost."""
        if self.updating_range_controls:
            return

        try:
            new_max = float(self.max_range_text.GetValue())
            min_val = float(self.min_range_text.GetValue())

            # Format to 2 decimal places
            new_max = round(new_max, 2)
            min_val = round(min_val, 2)

            # Auto-swap if max < min
            if new_max < min_val:
                self.updating_range_controls = True
                self.max_range_text.SetValue(f"{min_val:.2f}")
                self.min_range_text.SetValue(f"{new_max:.2f}")
                self.updating_range_controls = False
                new_max = min_val
            else:
                # Format the current control to 2f
                self.max_range_text.SetValue(f"{new_max:.2f}")

            # Update vLine position
            if self.parent.vline2 is not None:
                self.parent.vline2.set_xdata([new_max, new_max])
                self.update_vline_text_labels()
                self.parent.canvas.draw_idle()

        except ValueError:
            # Reset to 0.00 if invalid
            self.max_range_text.SetValue("0.00")

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

    def auto_detect_area_name(self, vline1_pos, vline2_pos):
        """Auto-detect and update area name based on vline positions using decision metric."""
        range_min = min(vline1_pos, vline2_pos)
        range_max = max(vline1_pos, vline2_pos)
        range_center = (range_min + range_max) / 2

        elements_db = self.create_elements_database()
        detected_elements = self.get_detected_elements()

        # Find all overlapping orbitals
        overlapping_matches = {}

        for element, orbitals in elements_db.items():
            overlapping_orbitals = []
            element_distances = {}

            for orbital, (be_min, be_max) in orbitals.items():
                be_center = (be_min + be_max) / 2

                # Check if range overlaps with binding energy range
                if not (range_max < be_min or range_min > be_max):
                    distance = abs(range_center - be_center)
                    overlapping_orbitals.append(orbital)
                    element_distances[orbital] = distance

            if overlapping_orbitals:
                # Use priority system to choose the best orbital
                priority_orbital = self.get_priority_orbital(element, overlapping_orbitals)
                distance = element_distances[priority_orbital]

                # Calculate decision value using your formula
                decision = self.calculate_decision_value(element, priority_orbital, distance)

                overlapping_matches[element] = {
                    'orbital': priority_orbital,
                    'distance': distance,
                    'decision': decision,  # ← NEW DECISION VALUE
                    'priority_level': self.get_element_priority_level(element, priority_orbital)
                }

        # Find the best match using decision value (lower is better)
        if overlapping_matches:
            # Calculate decision values for all matches
            for element in overlapping_matches:
                match_data = overlapping_matches[element]
                decision = self.calculate_decision_value(element, match_data['orbital'], match_data['distance'])
                overlapping_matches[element]['decision'] = decision

            # Sort by decision value - lowest decision wins
            best_element = min(overlapping_matches.keys(),
                               key=lambda e: overlapping_matches[e]['decision'])
            best_orbital = overlapping_matches[best_element]['orbital']
            best_match = f"{best_element}{best_orbital}"

            self.peak_label_text.SetValue(best_match)

    def on_key_press(self, event):
        """Handle key press events, TAB for next peak, Q for previous peak."""
        keycode = event.GetKeyCode()

        if keycode == wx.WXK_TAB:
            self.switch_to_next_peak()
            return  # Don't skip - we handled it
        elif keycode == ord('Q'):
            self.switch_to_previous_peak()
            return  # Don't skip - we handled it

        event.Skip()  # Let other keys pass through

    def on_background_only_pure(self, event):
        """Create background only without calculating area."""
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

        # Store vline positions BEFORE plotting background
        vline1_x = self.parent.vline1.get_xdata()[0]
        vline2_x = self.parent.vline2.get_xdata()[0]

        # Create background only
        self.parent.plot_manager.plot_background(self.parent)

        # Restore vlines at their positions
        self.parent.vline1 = self.parent.ax.axvline(vline1_x, color='r', linestyle='--', alpha=0.7)
        self.parent.vline2 = self.parent.ax.axvline(vline2_x, color='r', linestyle='--', alpha=0.7)

        # Recreate center vline
        center_pos = (vline1_x + vline2_x) / 2
        self.parent.vline_center = self.parent.ax.axvline(center_pos, color='blue', linestyle=':', alpha=0.4, linewidth=1.0)

        # Add text labels back
        self.add_vline_text_labels()

        # Add averaging indicator lines
        if hasattr(self.parent, 'add_averaging_indicator_lines'):
            self.parent.add_averaging_indicator_lines()

        # Update area tab selection state so vlines stay visible
        self.parent.area_tab_selected = True

        print(f"Background created using {selected_method} method")
        save_state(self.parent)

    def switch_to_next_peak(self):
        """Switch to next position: 0=plot range 10%-90%, then peak background ranges."""
        sheet_name = self.parent.sheet_combobox.GetValue()

        # Check if we have peaks in the fitting data
        if (sheet_name not in self.parent.Data['Core levels'] or
                'Fitting' not in self.parent.Data['Core levels'][sheet_name] or
                'Peaks' not in self.parent.Data['Core levels'][sheet_name]['Fitting']):
            return

        peaks = self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']
        peak_keys = list(peaks.keys())

        # Initialize timing variables if not exists
        if not hasattr(self, 'last_key_press_time'):
            self.last_key_press_time = 0

        current_time = time.time()
        time_since_last_press = current_time - self.last_key_press_time

        # Initialize position index if not exists or reset if more than 2 seconds
        if not hasattr(self, 'current_position_index') or time_since_last_press > 2.0:
            self.current_position_index = -1  # Start at -1 so first press goes to position 0

        # Update last key press time
        self.last_key_press_time = current_time

        # Total positions = 1 (plot range) + number of peaks
        total_positions = 1 + len(peak_keys)

        # Move to next position
        self.current_position_index = (self.current_position_index + 1) % total_positions

        if self.current_position_index == 0:
            # Position 0: Use 10%-90% of current plot range
            plot_range = self.get_plot_range_positions()
            if plot_range:
                low_pos, high_pos = plot_range
                self.move_vlines_to_range(low_pos, high_pos)
                self.update_range_controls_silent(low_pos, high_pos)
                self.show_position_feedback("Plot Range (10%-90%)", low_pos, high_pos)
                # Auto-detect and update area name
                self.auto_detect_area_name(low_pos, high_pos)
                # Update core level list if open
                self.update_core_level_list_if_open()
        else:
            # Position 1+: Use peak background ranges
            peak_index = self.current_position_index - 1
            if peak_index < len(peak_keys):
                current_peak_key = peak_keys[peak_index]
                current_peak = peaks[current_peak_key]

                if 'Bkg Low' in current_peak and 'Bkg High' in current_peak:
                    bkg_low = float(current_peak['Bkg Low'])
                    bkg_high = float(current_peak['Bkg High'])

                    self.move_vlines_to_range(bkg_low, bkg_high)
                    self.update_range_controls_silent(bkg_low, bkg_high)
                    self.show_position_feedback(current_peak_key, bkg_low, bkg_high)
                    self.highlight_peak_in_grid(current_peak_key)
                    # Auto-detect and update area name
                    self.auto_detect_area_name(bkg_low, bkg_high)
                    # Update core level list if open
                    self.update_core_level_list_if_open()

    def switch_to_previous_peak(self):
        """Switch to previous position: 0=plot range 10%-90%, then peak background ranges."""
        sheet_name = self.parent.sheet_combobox.GetValue()

        # Check if we have peaks in the fitting data
        if (sheet_name not in self.parent.Data['Core levels'] or
                'Fitting' not in self.parent.Data['Core levels'][sheet_name] or
                'Peaks' not in self.parent.Data['Core levels'][sheet_name]['Fitting']):
            return

        peaks = self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']
        peak_keys = list(peaks.keys())

        # Initialize timing variables if not exists
        if not hasattr(self, 'last_key_press_time'):
            self.last_key_press_time = 0

        current_time = time.time()
        time_since_last_press = current_time - self.last_key_press_time

        # Initialize position index if not exists or reset if more than 2 seconds
        if not hasattr(self, 'current_position_index') or time_since_last_press > 2.0:
            self.current_position_index = 1  # Start at 1 so first press goes to position 0

        # Update last key press time
        self.last_key_press_time = current_time

        # Total positions = 1 (plot range) + number of peaks
        total_positions = 1 + len(peak_keys)

        # Move to previous position
        self.current_position_index = (self.current_position_index - 1) % total_positions

        if self.current_position_index == 0:
            # Position 0: Use 10%-90% of current plot range
            plot_range = self.get_plot_range_positions()
            if plot_range:
                low_pos, high_pos = plot_range
                self.move_vlines_to_range(low_pos, high_pos)
                self.update_range_controls_silent(low_pos, high_pos)
                self.show_position_feedback("Plot Range (10%-90%)", low_pos, high_pos)
                # Auto-detect and update area name
                self.auto_detect_area_name(low_pos, high_pos)
                # Update core level list if open
                self.update_core_level_list_if_open()
        else:
            # Position 1+: Use peak background ranges
            peak_index = self.current_position_index - 1
            if peak_index < len(peak_keys):
                current_peak_key = peak_keys[peak_index]
                current_peak = peaks[current_peak_key]

                if 'Bkg Low' in current_peak and 'Bkg High' in current_peak:
                    bkg_low = float(current_peak['Bkg Low'])
                    bkg_high = float(current_peak['Bkg High'])

                    self.move_vlines_to_range(bkg_low, bkg_high)
                    self.update_range_controls_silent(bkg_low, bkg_high)
                    self.show_position_feedback(current_peak_key, bkg_low, bkg_high)
                    self.highlight_peak_in_grid(current_peak_key)
                    # Auto-detect and update area name
                    self.auto_detect_area_name(bkg_low, bkg_high)
                    # Update core level list if open
                    self.update_core_level_list_if_open()

    def get_plot_range_positions(self):
        """Get 10% and 90% positions of current plot X-axis range."""
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

    def move_vlines_to_range(self, bkg_low, bkg_high):
        """Move vlines to specified background range."""
        # Ensure proper min/max order
        min_val = min(bkg_low, bkg_high)
        max_val = max(bkg_low, bkg_high)

        # Update vlines positions
        if self.parent.vline1 is not None:
            self.parent.vline1.set_xdata([min_val, min_val])
        if self.parent.vline2 is not None:
            self.parent.vline2.set_xdata([max_val, max_val])

        # Update parent's background range variables
        self.parent.bg_min_energy = min_val
        self.parent.bg_max_energy = max_val

        # Update data structure
        sheet_name = self.parent.sheet_combobox.GetValue()
        if sheet_name in self.parent.Data['Core levels']:
            if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
                self.parent.Data['Core levels'][sheet_name]['Background'] = {}
            self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Low'] = min_val
            self.parent.Data['Core levels'][sheet_name]['Background']['Bkg High'] = max_val

        # # Update text labels
        # if hasattr(self.parent, 'update_vline_text_labels'):
        #     self.parent.update_vline_text_labels()

        # Update text labels
        self.update_vline_text_labels()

        # Update averaging indicator lines
        if hasattr(self.parent, 'add_averaging_indicator_lines'):
            self.parent.add_averaging_indicator_lines()

        # Redraw canvas
        self.parent.canvas.draw_idle()

    def update_range_controls_silent(self, bkg_low, bkg_high):
        """Update min/max range text controls WITHOUT triggering events."""
        if not (hasattr(self, 'min_range_text') and hasattr(self, 'max_range_text')):
            return
        if not (self.min_range_text and self.max_range_text):
            return

        try:
            # Temporarily unbind events to prevent recursive updates
            self.min_range_text.Unbind(wx.EVT_TEXT)
            self.max_range_text.Unbind(wx.EVT_TEXT)

            # Update values using ChangeValue (doesn't trigger events)
            min_val = min(bkg_low, bkg_high)
            max_val = max(bkg_low, bkg_high)

            self.min_range_text.ChangeValue(f"{min_val:.2f}")
            self.max_range_text.ChangeValue(f"{max_val:.2f}")

            # Rebind events
            self.min_range_text.Bind(wx.EVT_TEXT, self.on_min_range_change)
            self.max_range_text.Bind(wx.EVT_TEXT, self.on_max_range_change)

        except (RuntimeError, AttributeError):
            # Controls may have been destroyed, ignore silently
            pass

    def update_range_controls(self, bkg_low, bkg_high):
        """Update min/max range text controls."""
        if hasattr(self, 'min_range_text') and hasattr(self, 'max_range_text'):
            self.min_range_text.SetValue(f"{min(bkg_low, bkg_high):.2f}")
            self.max_range_text.SetValue(f"{max(bkg_low, bkg_high):.2f}")

    def show_position_feedback(self, identifier, low_pos, high_pos):
        """Show visual feedback for current position."""
        min_val = min(low_pos, high_pos)
        max_val = max(low_pos, high_pos)
        if identifier == "Plot Range (10%-90%)":
            self.SetTitle(
                f"Measure Area - 10%-90%")
            # Clear peak highlighting for plot range position
            self.clear_peak_grid_highlights()
        else:
            self.SetTitle(f"Measure Area - {identifier}")

        # Highlight the corresponding peak in the grid
        self.highlight_peak_in_grid(identifier)

    def Show(self, show=True):
        """Override Show to reset peak selection when window opens."""
        if show:
            self.current_peak_index = -1  # Reset to first peak when opening
            self.clear_peak_grid_highlights()  # Clear any existing highlights
        super().Show(show)

    def highlight_peak_in_grid(self, peak_key):
        """Highlight the corresponding peak rows in the peak fitting grid."""
        if not hasattr(self.parent, 'peak_params_grid'):
            return

        # Clear existing highlights first
        self.clear_peak_grid_highlights()

        # Find the peak index from the peak key
        sheet_name = self.parent.sheet_combobox.GetValue()
        if (sheet_name not in self.parent.Data['Core levels'] or
                'Fitting' not in self.parent.Data['Core levels'][sheet_name] or
                'Peaks' not in self.parent.Data['Core levels'][sheet_name]['Fitting']):
            return

        peaks = self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']
        peak_keys = list(peaks.keys())

        if peak_key not in peak_keys:
            return

        peak_index = peak_keys.index(peak_key)

        # Each peak uses 2 rows (data row and constraint row)
        data_row = peak_index * 2
        constraint_row = data_row + 1

        # Set grey background for both rows of the selected peak
        grey_color = wx.Colour(220, 220, 220)  # Light grey

        try:
            num_cols = self.parent.peak_params_grid.GetNumberCols()

            # Highlight data row
            for col in range(num_cols):
                self.parent.peak_params_grid.SetCellBackgroundColour(data_row, col, grey_color)

            # Highlight constraint row
            # for col in range(num_cols):
            #     self.parent.peak_params_grid.SetCellBackgroundColour(constraint_row, col, grey_color)

            # Make the highlighted rows visible
            self.parent.peak_params_grid.MakeCellVisible(data_row, 0)

            # Force refresh to show changes
            self.parent.peak_params_grid.ForceRefresh()

        except (RuntimeError, AttributeError):
            # Grid may not be accessible, ignore silently
            pass

    def clear_peak_grid_highlights(self):
        """Clear all peak highlighting in the peak fitting grid."""
        if not hasattr(self.parent, 'peak_params_grid'):
            return

        try:
            num_rows = self.parent.peak_params_grid.GetNumberRows()
            num_cols = self.parent.peak_params_grid.GetNumberCols()

            # Reset all backgrounds to default colors
            for row in range(num_rows):
                for col in range(num_cols):
                    if row % 2 == 0:  # Data rows (even rows)
                        self.parent.peak_params_grid.SetCellBackgroundColour(row, col, wx.WHITE)
                    else:  # Constraint rows (odd rows)
                        self.parent.peak_params_grid.SetCellBackgroundColour(row, col, wx.Colour(200, 245, 228))

            self.parent.peak_params_grid.ForceRefresh()

        except (RuntimeError, AttributeError):
            # Grid may not be accessible, ignore silently
            pass

    def update_core_level_list_if_open(self):
        """Update core level list window if it's open and vlines have changed."""
        if (hasattr(self, 'core_level_list_window') and
                self.core_level_list_window and
                self.parent.vline1 is not None and
                self.parent.vline2 is not None):

            try:
                # Get current vline positions
                vline1_pos = self.parent.vline1.get_xdata()[0]
                vline2_pos = self.parent.vline2.get_xdata()[0]

                # Get updated possible matches
                possible_matches = self.get_all_possible_core_levels(vline1_pos, vline2_pos)

                # Update the window
                self.core_level_list_window.update_list(possible_matches)
            except:
                # Window was destroyed or error occurred
                self.core_level_list_window = None

    def on_show_core_level_list(self, event):
        """Show core level list window with possible matches."""
        if self.parent.vline1 is None or self.parent.vline2 is None:
            wx.MessageBox("Please set vertical lines first.", "No Range Set", wx.OK | wx.ICON_WARNING)
            return

        # Check if window already exists
        if hasattr(self, 'core_level_list_window') and self.core_level_list_window:
            try:
                self.core_level_list_window.Raise()  # Bring to front
                return
            except:
                # Window was destroyed, continue to create new one
                pass

        # Get current vline positions
        vline1_pos = self.parent.vline1.get_xdata()[0]
        vline2_pos = self.parent.vline2.get_xdata()[0]

        # Get all possible core levels in this range
        possible_matches = self.get_all_possible_core_levels(vline1_pos, vline2_pos)

        if not possible_matches:
            wx.MessageBox("No core levels found in this range.", "No Matches", wx.OK | wx.ICON_INFORMATION)
            return

        # Show core level list window
        self.core_level_list_window = CoreLevelListWindow(self, possible_matches)
        self.core_level_list_window.Show()

    def get_all_possible_core_levels(self, vline1_pos, vline2_pos):
        """Get all possible core levels in the range, sorted by decision value."""
        range_min = min(vline1_pos, vline2_pos)
        range_max = max(vline1_pos, vline2_pos)
        range_center = (range_min + range_max) / 2

        elements_db = self.create_elements_database()
        all_matches = []

        for element, orbitals in elements_db.items():
            for orbital, (be_min, be_max) in orbitals.items():
                be_center = (be_min + be_max) / 2

                # Check if range overlaps with binding energy range
                if not (range_max < be_min or range_min > be_max):
                    distance = abs(range_center - be_center)
                    priority_level = self.get_element_priority_level(element, orbital)
                    decision = self.calculate_decision_value(element, orbital, distance)

                    all_matches.append({
                        'name': f"{element}{orbital}",
                        'element': element,
                        'orbital': orbital,
                        'center': be_center,
                        'range': (be_min, be_max),
                        'distance': distance,
                        'priority_level': priority_level,
                        'decision': decision
                    })

        # Sort by decision value (lower is better)
        all_matches.sort(key=lambda x: x['decision'])

        return all_matches

    def get_element_priority_level(self, element, orbital):
        """Get priority level for element-orbital combination."""
        # Common/important elements get higher priority (lower number)
        common_elements = {'C': 1, 'O': 1, 'N': 1, 'Si': 2, 'P': 2, 'Ca': 2, 'Na': 2, 'Fe': 3, 'Cr': 3, 'Ti': 3,
                           'Cu': 3, 'Ag': 3, 'Au': 3, 'Zn': 3, 'Co': 3, 'Ni': 3, 'Mn': 3, 'Al': 3, 'Mg': 3, }
        element_priority = common_elements.get(element, 4)

        # Common orbitals get higher priority
        orbital_priority = {'1s': 1, '2p': 1, '3d': 1, '4f': 1, 'kll': 2, '2s': 2, '3p': 2, '4d': 2, '3s': 3,
                            '4p': 3}.get(orbital, 4)

        return element_priority + orbital_priority

    def get_priority_orbital(self, element, orbitals):
        """Choose the best orbital for an element based on priority."""
        priority_order = ['1s', '2p', '3d', '4f', 'kll', '2s', '4d', '3p', '3s', '4p', '4s']

        for priority_orbital in priority_order:
            if priority_orbital in orbitals:
                return priority_orbital

        return orbitals[0] if orbitals else None

    def calculate_decision_value(self, element, orbital, distance):
        """Calculate decision value based on distance and priorities."""
        # Get individual priorities
        common_elements = {'C': 1, 'O': 1, 'N': 1, 'Si': 2, 'P': 2, 'Ca': 2, 'Na': 2, 'Fe': 3, 'Cr': 3, 'Ti': 3,
                           'Cu': 3, 'Ag': 3, 'Au': 3, 'Zn': 3, 'Co': 3, 'Ni': 3, 'Mn': 3, 'Al': 3, 'Mg': 3, }
        priority_element = common_elements.get(element, 4)

        orbital_priorities = {'1s': 1, '2p': 1, '3d': 1, '4f': 1, 'kll': 2, '2s': 2, '3p': 2, '4d': 2, '3s': 3,
                              '4p': 3}
        priority_orbital = orbital_priorities.get(orbital, 5)

        # Handle division by zero and negative denominators
        orbital_denominator = max(0.1, 5 - priority_orbital)  # Prevent division by zero/negative
        element_denominator = max(0.1, 5 - priority_element)  # Prevent division by zero

        # Calculate decision value (lower is better)
        decision = (distance / orbital_denominator) + (distance / element_denominator)

        return decision

    def on_core_level_selected(self, core_level_name):
        """Called when a core level is selected from the list."""
        self.peak_label_text.SetValue(core_level_name)
        # if hasattr(self, 'core_level_list_window') and self.core_level_list_window:
        #     self.core_level_list_window.Close()
        #     self.core_level_list_window = None

    def on_auto_id(self, event):
        """Open Auto ID window for survey/wide scan identification"""
        try:
            sheet_name = self.parent.sheet_combobox.GetValue()

            # Check if this is a survey or wide scan
            if not any(x in sheet_name.lower() for x in ['survey', 'wide']):
                wx.MessageBox("Auto ID is only available for Survey or Wide scan sheets", "Info",
                              wx.OK | wx.ICON_INFORMATION)
                return

            # Import and open AutoID window
            from libraries.ToolsMenu.AutoID import AutoIDWindow
            auto_id_window = AutoIDWindow(self.parent)
            auto_id_window.Show()

        except ImportError:
            wx.MessageBox("AutoID module not found", "Error", wx.OK | wx.ICON_ERROR)
        except Exception as e:
            wx.MessageBox(f"Auto ID failed: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)


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


class CoreLevelListWindow(wx.Frame):
    def __init__(self, parent, possible_matches):
        self.parent = parent
        self.possible_matches = possible_matches
        self.sort_column = 3  # Default sort by Distance column
        self.sort_ascending = True

        super().__init__(None, title="Core Level List",
                         style=(wx.DEFAULT_FRAME_STYLE & ~wx.RESIZE_BORDER) | wx.STAY_ON_TOP)

        self.init_ui()
        self.position_window()

    def init_ui(self):
        """Initialize the user interface."""
        panel = wx.Panel(self)

        # Set yellow background to match AreaFit Screen
        panel.SetBackgroundColour(wx.Colour(250, 250, 230))

        # Create list control with sortable columns
        self.list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.list_ctrl.InsertColumn(0, "Core Level", width=70)
        self.list_ctrl.InsertColumn(1, "BE Center", width=70)
        self.list_ctrl.InsertColumn(2, "Range", width=70)
        self.list_ctrl.InsertColumn(3, "Distance", width=60)
        self.list_ctrl.InsertColumn(4, "Decision", width=60)

        # Bind column click event for sorting
        self.list_ctrl.Bind(wx.EVT_LIST_COL_CLICK, self.on_column_click)

        # Sort by distance initially and populate list
        self.sort_data()
        self.populate_list()

        # Buttons
        select_btn = wx.Button(panel, label="Apply")
        cancel_btn = wx.Button(panel, label="Close")

        select_btn.Bind(wx.EVT_BUTTON, self.on_select)
        cancel_btn.Bind(wx.EVT_BUTTON, self.on_cancel)
        self.list_ctrl.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.on_item_activated)

        # Layout
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.list_ctrl, 1, wx.EXPAND | wx.ALL, 5)

        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_sizer.Add(select_btn, 0, wx.ALL, 5)
        btn_sizer.Add(cancel_btn, 0, wx.ALL, 5)

        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER)
        panel.SetSizer(sizer)

        self.Bind(wx.EVT_CLOSE, self.on_close)

        self.SetSize((390, 300))

    def on_column_click(self, event):
        """Handle column click for sorting."""
        clicked_col = event.GetColumn()

        # Toggle sort order if same column, otherwise sort ascending
        if self.sort_column == clicked_col:
            self.sort_ascending = not self.sort_ascending
        else:
            self.sort_column = clicked_col
            self.sort_ascending = True

        self.sort_data()
        self.populate_list()

    def sort_data(self):
        """Sort possible_matches based on current sort column and order."""
        if self.sort_column == 0:  # Core Level (string)
            self.possible_matches.sort(key=lambda x: x['name'], reverse=not self.sort_ascending)
        elif self.sort_column == 1:  # BE Center (float)
            self.possible_matches.sort(key=lambda x: x['center'], reverse=not self.sort_ascending)
        elif self.sort_column == 2:  # Range (sort by range center)
            self.possible_matches.sort(key=lambda x: (x['range'][0] + x['range'][1]) / 2,
                                       reverse=not self.sort_ascending)
        elif self.sort_column == 3:  # Distance (float)
            self.possible_matches.sort(key=lambda x: x['distance'], reverse=not self.sort_ascending)
        elif self.sort_column == 4:  # Decision (float) ← NEW SORTING OPTION
            self.possible_matches.sort(key=lambda x: x.get('decision', 999), reverse=not self.sort_ascending)

    def populate_list(self):
        """Populate the list control with sorted data."""
        # Clear existing items
        self.list_ctrl.DeleteAllItems()

        # Populate with matches using .2f format
        for i, match in enumerate(self.possible_matches):
            index = self.list_ctrl.InsertItem(i, match['name'])
            self.list_ctrl.SetItem(index, 1, f"{match['center']:.2f}")
            self.list_ctrl.SetItem(index, 2, f"{match['range'][0]:.0f}-{match['range'][1]:.0f}")
            self.list_ctrl.SetItem(index, 3, f"{match['distance']:.2f}")

            # Add decision value if available
            if 'decision' in match:
                self.list_ctrl.SetItem(index, 4, f"{match['decision']:.2f}")  # ← NEW DECISION COLUMN
            else:
                self.list_ctrl.SetItem(index, 4, "N/A")

            # Highlight the top choice
            if i == 0:
                self.list_ctrl.SetItemBackgroundColour(index, wx.Colour(200, 255, 200))

        # Select first item
        if self.list_ctrl.GetItemCount() > 0:
            self.list_ctrl.Select(0)

    def position_window(self):
        """Position window on top-right of parent AreaFit window."""
        parent_pos = self.parent.GetPosition()
        parent_size = self.parent.GetSize()

        # Position to the right and slightly above the parent window
        new_x = parent_pos.x + parent_size.width
        new_y = parent_pos.y - 50  # Slightly above

        self.SetPosition((new_x, new_y))

    def on_select(self, event):
        """Handle select button click."""
        selected = self.list_ctrl.GetFirstSelected()
        if selected >= 0:
            core_level_name = self.possible_matches[selected]['name']
            self.parent.on_core_level_selected(core_level_name)
            # Window remains open for additional selections

    def on_cancel(self, event):
        """Handle cancel button click."""
        self.Close()

    def on_item_activated(self, event):
        """Handle double-click on list item."""
        selected = event.GetIndex()
        if selected >= 0:
            core_level_name = self.possible_matches[selected]['name']
            self.parent.on_core_level_selected(core_level_name)
            # Window remains open for additional selections

    def update_list(self, new_possible_matches):
        """Update the list with new possible matches."""
        self.possible_matches = new_possible_matches
        self.sort_data()
        self.populate_list()

    def on_close(self, event):
        """Handle window close event."""
        self.parent.core_level_list_window = None
        self.Destroy()
