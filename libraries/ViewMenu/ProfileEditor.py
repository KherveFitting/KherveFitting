import wx
import wx.grid
import pandas as pd
import json
import os
import re


class ProfileEditorWindow(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, title="Profile Data Editor", size=(700, 500))
        self.parent = parent
        self.sheet_name = parent.sheet_combobox.GetValue() if hasattr(parent, 'sheet_combobox') else None

        self.init_ui()

        if self.sheet_name and self.sheet_name.startswith('zzProfile'):
            self.load_data()

        self.Centre()

    def init_ui(self):
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Left panel - Core levels selection
        left_panel = wx.Panel(panel)
        left_sizer = wx.BoxSizer(wx.VERTICAL)

        # Title for left panel
        left_title = wx.StaticText(left_panel, label="Core Levels Selection")
        font = left_title.GetFont()
        font.PointSize += 1
        font = font.Bold()
        left_title.SetFont(font)
        left_sizer.Add(left_title, 0, wx.ALL | wx.CENTER, 5)

        # Select/Unselect buttons
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        select_all_btn = wx.Button(left_panel, label="Select All")
        select_all_btn.Bind(wx.EVT_BUTTON, self.on_select_all)
        button_sizer.Add(select_all_btn, 1, wx.ALL | wx.EXPAND, 2)

        unselect_all_btn = wx.Button(left_panel, label="Unselect All")
        unselect_all_btn.Bind(wx.EVT_BUTTON, self.on_unselect_all)
        button_sizer.Add(unselect_all_btn, 1, wx.ALL | wx.EXPAND, 2)
        left_sizer.Add(button_sizer, 0, wx.ALL | wx.EXPAND, 5)

        # Core levels checklist
        self.core_levels_checklist = wx.CheckListBox(left_panel)
        self.core_levels_checklist.SetMinSize((150, 250))
        self.populate_core_levels_list()
        self.core_levels_checklist.Bind(wx.EVT_CONTEXT_MENU, self.on_core_levels_context_menu)
        left_sizer.Add(self.core_levels_checklist, 1, wx.ALL | wx.EXPAND, 5)

        # Create profile buttons
        create_label = wx.StaticText(left_panel, label="Create New Profile:")
        left_sizer.Add(create_label, 0, wx.ALL, 5)

        # Dropdown for profile type
        type_sizer = wx.BoxSizer(wx.HORIZONTAL)
        type_label = wx.StaticText(left_panel, label="Type:")
        type_sizer.Add(type_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 2)

        self.profile_type_combo = wx.ComboBox(left_panel, choices=["Concentration", "Area", "Position"],
                                              style=wx.CB_READONLY, value="Concentration")
        type_sizer.Add(self.profile_type_combo, 1, wx.ALL | wx.EXPAND, 2)
        left_sizer.Add(type_sizer, 0, wx.ALL | wx.EXPAND, 5)

        # Create from Peak Fitting Grid
        create_peak_btn = wx.Button(left_panel, label="Create Profile from\nPeak Fitting Grid")
        create_peak_btn.SetMinSize((125, 35))
        create_peak_btn.Bind(wx.EVT_BUTTON, self.on_create_from_peak_grid)
        left_sizer.Add(create_peak_btn, 0, wx.ALL | wx.EXPAND, 5)

        # Create from Results Grid
        create_results_btn = wx.Button(left_panel, label="Create Profile from\nResults Grid")
        create_results_btn.SetMinSize((125, 35))
        create_results_btn.Bind(wx.EVT_BUTTON, self.on_create_from_results_grid)
        left_sizer.Add(create_results_btn, 0, wx.ALL | wx.EXPAND, 5)

        left_panel.SetSizer(left_sizer)
        main_sizer.Add(left_panel, 0, wx.ALL | wx.EXPAND, 5)

        # Right panel - Data editor
        right_panel = wx.Panel(panel)
        right_sizer = wx.BoxSizer(wx.VERTICAL)

        # Title
        if self.sheet_name:
            title = wx.StaticText(right_panel, label=f"Editing: {self.sheet_name}")
        else:
            title = wx.StaticText(right_panel, label="Create New Profile")
        title_font = title.GetFont()
        title_font.PointSize += 2
        title_font = title_font.Bold()
        title.SetFont(title_font)
        right_sizer.Add(title, 0, wx.ALL | wx.CENTER, 5)

        # Axis labels section
        labels_box = wx.StaticBoxSizer(wx.VERTICAL, right_panel, "Axis Labels")

        # X-axis label
        x_label_sizer = wx.BoxSizer(wx.HORIZONTAL)
        x_label_text = wx.StaticText(right_panel, label="X-Axis Label:")
        x_label_sizer.Add(x_label_text, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        self.x_axis_ctrl = wx.TextCtrl(right_panel, size=(300, -1))
        x_label_sizer.Add(self.x_axis_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        labels_box.Add(x_label_sizer, 0, wx.EXPAND)

        # Y-axis label
        y_label_sizer = wx.BoxSizer(wx.HORIZONTAL)
        y_label_text = wx.StaticText(right_panel, label="Y-Axis Label:")
        y_label_sizer.Add(y_label_text, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        self.y_axis_ctrl = wx.TextCtrl(right_panel, size=(300, -1))
        y_label_sizer.Add(self.y_axis_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        labels_box.Add(y_label_sizer, 0, wx.EXPAND)

        right_sizer.Add(labels_box, 0, wx.ALL | wx.EXPAND, 5)

        # Data grid section with controls
        grid_box = wx.StaticBoxSizer(wx.VERTICAL, right_panel, "Profile Data")

        # Row/Column controls
        controls_sizer = wx.BoxSizer(wx.HORIZONTAL)

        add_row_btn = wx.Button(right_panel, label="Add\nRow")
        add_row_btn.SetMinSize((70, 35))
        add_row_btn.Bind(wx.EVT_BUTTON, self.on_add_row)
        controls_sizer.Add(add_row_btn, 0, wx.ALL, 2)

        remove_row_btn = wx.Button(right_panel, label="Remove\nRow")
        remove_row_btn.SetMinSize((70, 35))
        remove_row_btn.Bind(wx.EVT_BUTTON, self.on_remove_row)
        controls_sizer.Add(remove_row_btn, 0, wx.ALL, 2)

        controls_sizer.AddSpacer(10)

        add_col_btn = wx.Button(right_panel, label="Add\nColumn")
        add_col_btn.SetMinSize((70, 35))
        add_col_btn.Bind(wx.EVT_BUTTON, self.on_add_column)
        controls_sizer.Add(add_col_btn, 0, wx.ALL, 2)

        remove_col_btn = wx.Button(right_panel, label="Remove\nColumn")
        remove_col_btn.SetMinSize((70, 35))
        remove_col_btn.Bind(wx.EVT_BUTTON, self.on_remove_column)
        controls_sizer.Add(remove_col_btn, 0, wx.ALL, 2)

        controls_sizer.AddSpacer(10)

        edit_header_btn = wx.Button(right_panel, label="Edit\nHeader")
        edit_header_btn.SetMinSize((70, 35))
        edit_header_btn.Bind(wx.EVT_BUTTON, self.on_edit_header)
        controls_sizer.Add(edit_header_btn, 0, wx.ALL, 2)

        grid_box.Add(controls_sizer, 0, wx.ALL | wx.EXPAND, 5)

        # Grid
        self.data_grid = wx.grid.Grid(right_panel)
        self.data_grid.CreateGrid(0, 0)
        self.data_grid.EnableEditing(True)
        self.data_grid.SetDefaultCellAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)
        self.data_grid.SetRowLabelSize(35)  # Narrower row labels like other grids

        # Bind grid events
        self.data_grid.Bind(wx.grid.EVT_GRID_LABEL_LEFT_DCLICK, self.on_header_double_click)

        grid_box.Add(self.data_grid, 1, wx.ALL | wx.EXPAND, 5)
        right_sizer.Add(grid_box, 1, wx.ALL | wx.EXPAND, 5)

        # Buttons
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)

        save_btn = wx.Button(right_panel, label="Save Changes")
        save_btn.SetMinSize((125, 35))
        save_btn.Bind(wx.EVT_BUTTON, self.on_save)
        button_sizer.Add(save_btn, 0, wx.ALL, 5)

        cancel_btn = wx.Button(right_panel, label="Cancel")
        cancel_btn.SetMinSize((125, 35))
        cancel_btn.Bind(wx.EVT_BUTTON, self.on_cancel)
        button_sizer.Add(cancel_btn, 0, wx.ALL, 5)

        right_sizer.Add(button_sizer, 0, wx.ALL | wx.CENTER, 10)

        right_panel.SetSizer(right_sizer)
        main_sizer.Add(right_panel, 1, wx.ALL | wx.EXPAND, 5)

        panel.SetSizer(main_sizer)

    def populate_core_levels_list(self):
        """Populate the core levels checklist with current core levels"""
        self.core_levels_checklist.Clear()

        if hasattr(self.parent, 'Data') and 'Core levels' in self.parent.Data:
            core_levels = list(self.parent.Data['Core levels'].keys())
            # Filter out zzProfile sheets
            core_levels = [cl for cl in core_levels if not cl.startswith('zzProfile')]
            for core_level in core_levels:
                self.core_levels_checklist.Append(core_level)

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

        cleaned_name = sheet_name.strip()

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
                # Check if first part looks like a core level
                if re.match(r'^[A-Z][a-z]?\d+[a-z]*\d*$', first_part):
                    core_level_match = re.match(r'^([A-Z][a-z]?\d+[a-z]*)', first_part)
                    if core_level_match:
                        return core_level_match.group(1)
                break

        return cleaned_name[:4] if len(cleaned_name) > 4 else cleaned_name

    def get_selected_core_levels(self):
        """Get list of selected core levels from the checklist"""
        selected = []
        for i in range(self.core_levels_checklist.GetCount()):
            if self.core_levels_checklist.IsChecked(i):
                selected.append(self.core_levels_checklist.GetString(i))
        return selected

    def on_create_from_peak_grid(self, event):
        """Create profile from peak fitting grid"""
        selected_sheets = self.get_selected_core_levels()
        if not selected_sheets:
            wx.MessageBox("No core levels selected", "Create Profile Failed", wx.OK | wx.ICON_WARNING)
            return

        profile_type = self.profile_type_combo.GetValue()

        # Determine profile name
        lowest_number = self._get_lowest_number(selected_sheets)
        profile_sheet_name = "zzProfile" if lowest_number == 0 else f"zzProfile{lowest_number}"

        # Check if exists
        if profile_sheet_name in self.parent.Data['Core levels']:
            response = wx.MessageBox(
                f"Profile sheet '{profile_sheet_name}' already exists. Overwrite?",
                "Confirm Overwrite",
                wx.YES_NO | wx.ICON_QUESTION
            )
            if response != wx.YES:
                # Find next available name
                profile_sheet_name = self._get_next_available_profile_name(profile_sheet_name)

        # Create profile based on type
        if profile_type == "Concentration":
            self._create_concentration_profile(selected_sheets, profile_sheet_name)
        elif profile_type == "Area":
            self._create_area_profile(selected_sheets, profile_sheet_name)
        elif profile_type == "Position":
            self._create_position_profile(selected_sheets, profile_sheet_name)

    def on_create_from_results_grid(self, event):
        """Create profile from results grid"""
        selected_sheets = self.get_selected_core_levels()
        if not selected_sheets:
            wx.MessageBox("No core levels selected", "Create Profile Failed", wx.OK | wx.ICON_WARNING)
            return

        profile_type = self.profile_type_combo.GetValue()

        # Determine profile name
        lowest_number = self._get_lowest_number(selected_sheets)
        profile_sheet_name = "zzProfile" if lowest_number == 0 else f"zzProfile{lowest_number}"

        # Check if exists
        if profile_sheet_name in self.parent.Data['Core levels']:
            response = wx.MessageBox(
                f"Profile sheet '{profile_sheet_name}' already exists. Overwrite?",
                "Confirm Overwrite",
                wx.YES_NO | wx.ICON_QUESTION
            )
            if response != wx.YES:
                return

        # Create profile from results grid
        if profile_type == "Concentration":
            self._create_concentration_profile_from_results(selected_sheets, profile_sheet_name)
        elif profile_type == "Area":
            self._create_area_profile_from_results(selected_sheets, profile_sheet_name)
        elif profile_type == "Position":
            self._create_position_profile_from_results(selected_sheets, profile_sheet_name)

    def _get_lowest_number(self, selected_sheets):
        """Extract lowest number from selected sheets"""
        lowest_number = float('inf')
        for sheet in selected_sheets:
            num_match = re.search(r'(\d+)$', sheet)
            if num_match:
                number = int(num_match.group(1))
            else:
                number = 0
            lowest_number = min(lowest_number, number)
        return lowest_number

    def _create_concentration_profile(self, selected_sheets, profile_sheet_name):
        """Create atomic concentration profile from peak fitting grid"""
        from libraries.Sheet_Operations import on_sheet_selected

        profile_data = {}
        profile_data['Number'] = []
        peak_names = set()

        # First pass: collect peak names
        for idx, sheet_name in enumerate(selected_sheets):
            num_match = re.search(r'(\d+)$', sheet_name)
            number = int(num_match.group(1)) if num_match else 0
            profile_data['Number'].append(number)

            if (sheet_name in self.parent.Data['Core levels'] and
                    'Fitting' in self.parent.Data['Core levels'][sheet_name] and
                    'Peaks' in self.parent.Data['Core levels'][sheet_name]['Fitting']):
                peaks = self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                for peak_name in peaks.keys():
                    peak_names.add(peak_name)

        peak_names = sorted(list(peak_names))
        for peak_name in peak_names:
            profile_data[f"{peak_name} At(%)"] = []

        # Second pass: get atomic concentrations
        for sheet_name in selected_sheets:
            sheet_peak_data = {}

            self.parent.sheet_combobox.SetValue(sheet_name)
            on_sheet_selected(self.parent, sheet_name)
            self.parent.update_ratios()

            grid = self.parent.peak_params_grid
            num_peaks = grid.GetNumberRows() // 2

            for i in range(num_peaks):
                row = i * 2
                try:
                    peak_name = grid.GetCellValue(row, 1)
                    atomic_conc_str = grid.GetCellValue(row, 10)
                    if atomic_conc_str and atomic_conc_str.strip():
                        sheet_peak_data[peak_name] = float(atomic_conc_str)
                except (ValueError, IndexError):
                    continue

            for peak_name in peak_names:
                column_name = f"{peak_name} At(%)"
                if peak_name in sheet_peak_data:
                    profile_data[column_name].append(sheet_peak_data[peak_name])
                else:
                    profile_data[column_name].append(0.0)

        df = pd.DataFrame(profile_data)
        cols = ['Number'] + [col for col in df.columns if col != 'Number']
        df = df[cols]
        df = df.sort_values(by='Number').reset_index(drop=True)

        self._save_profile(df, profile_sheet_name, "Atomic Concentration (%)", "Number", "Atomic_Concentration")

    def _create_area_profile(self, selected_sheets, profile_sheet_name):
        """Create area profile from peak fitting grid"""
        from libraries.Sheet_Operations import on_sheet_selected

        profile_data = {}
        profile_data['Number'] = []
        peak_names = set()

        for idx, sheet_name in enumerate(selected_sheets):
            num_match = re.search(r'(\d+)$', sheet_name)
            number = int(num_match.group(1)) if num_match else 0
            profile_data['Number'].append(number)

            if (sheet_name in self.parent.Data['Core levels'] and
                    'Fitting' in self.parent.Data['Core levels'][sheet_name] and
                    'Peaks' in self.parent.Data['Core levels'][sheet_name]['Fitting']):
                peaks = self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                for peak_name in peaks.keys():
                    peak_names.add(peak_name)

        peak_names = sorted(list(peak_names))
        for peak_name in peak_names:
            profile_data[f"{peak_name} Area"] = []

        for sheet_name in selected_sheets:
            sheet_peak_data = {}

            self.parent.sheet_combobox.SetValue(sheet_name)
            on_sheet_selected(self.parent, sheet_name)

            grid = self.parent.peak_params_grid
            num_peaks = grid.GetNumberRows() // 2

            for i in range(num_peaks):
                row = i * 2
                try:
                    peak_name = grid.GetCellValue(row, 1)
                    area_str = grid.GetCellValue(row, 6)
                    if area_str and area_str.strip():
                        sheet_peak_data[peak_name] = float(area_str)
                except (ValueError, IndexError):
                    continue

            for peak_name in peak_names:
                column_name = f"{peak_name} Area"
                if peak_name in sheet_peak_data:
                    profile_data[column_name].append(sheet_peak_data[peak_name])
                else:
                    profile_data[column_name].append(0.0)

        df = pd.DataFrame(profile_data)
        cols = ['Number'] + [col for col in df.columns if col != 'Number']
        df = df[cols]
        df = df.sort_values(by='Number').reset_index(drop=True)

        self._save_profile(df, profile_sheet_name, "Area (CPS)", "Number", "Area")

    def _create_position_profile(self, selected_sheets, profile_sheet_name):
        """Create position profile from peak fitting grid"""
        from libraries.Sheet_Operations import on_sheet_selected

        profile_data = {}
        profile_data['Number'] = []
        peak_names = set()

        for idx, sheet_name in enumerate(selected_sheets):
            num_match = re.search(r'(\d+)$', sheet_name)
            number = int(num_match.group(1)) if num_match else 0
            profile_data['Number'].append(number)

            if (sheet_name in self.parent.Data['Core levels'] and
                    'Fitting' in self.parent.Data['Core levels'][sheet_name] and
                    'Peaks' in self.parent.Data['Core levels'][sheet_name]['Fitting']):
                peaks = self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                for peak_name in peaks.keys():
                    peak_names.add(peak_name)

        peak_names = sorted(list(peak_names))
        for peak_name in peak_names:
            profile_data[f"{peak_name} Pos"] = []

        for sheet_name in selected_sheets:
            sheet_peak_data = {}

            self.parent.sheet_combobox.SetValue(sheet_name)
            on_sheet_selected(self.parent, sheet_name)

            grid = self.parent.peak_params_grid
            num_peaks = grid.GetNumberRows() // 2

            for i in range(num_peaks):
                row = i * 2
                try:
                    peak_name = grid.GetCellValue(row, 1)
                    pos_str = grid.GetCellValue(row, 2)
                    if pos_str and pos_str.strip():
                        sheet_peak_data[peak_name] = float(pos_str)
                except (ValueError, IndexError):
                    continue

            for peak_name in peak_names:
                column_name = f"{peak_name} Pos"
                if peak_name in sheet_peak_data:
                    profile_data[column_name].append(sheet_peak_data[peak_name])
                else:
                    profile_data[column_name].append(0.0)

        df = pd.DataFrame(profile_data)
        cols = ['Number'] + [col for col in df.columns if col != 'Number']
        df = df[cols]
        df = df.sort_values(by='Number').reset_index(drop=True)

        self._save_profile(df, profile_sheet_name, "Position (eV)", "Number", "Position")

    def _create_concentration_profile_from_results(self, selected_sheets, profile_sheet_name):
        """Create concentration profile from results grid"""
        profile_data = {}
        profile_data['Number'] = []
        peak_names = set()

        # Collect data from results grid
        for sheet_name in selected_sheets:
            num_match = re.search(r'(\d+)$', sheet_name)
            number = int(num_match.group(1)) if num_match else 0
            profile_data['Number'].append(number)

            # Get data from results grid for this sheet
            grid = self.parent.results_grid
            num_rows = grid.GetNumberRows()

            for row in range(num_rows):
                try:
                    result_sheet = grid.GetCellValue(row, 18)  # Column 18 has sheet name
                    if result_sheet == sheet_name:
                        peak_name = grid.GetCellValue(row, 0)  # Column 0 has peak label
                        peak_names.add(peak_name)
                except (ValueError, IndexError):
                    continue

        peak_names = sorted(list(peak_names))
        for peak_name in peak_names:
            profile_data[f"{peak_name} At(%)"] = []

        for sheet_name in selected_sheets:
            sheet_peak_data = {}
            grid = self.parent.results_grid
            num_rows = grid.GetNumberRows()

            for row in range(num_rows):
                try:
                    result_sheet = grid.GetCellValue(row, 18)
                    if result_sheet == sheet_name:
                        peak_name = grid.GetCellValue(row, 0)
                        atomic_conc_str = grid.GetCellValue(row, 10)  # Atomic %
                        if atomic_conc_str and atomic_conc_str.strip():
                            sheet_peak_data[peak_name] = float(atomic_conc_str)
                except (ValueError, IndexError):
                    continue

            for peak_name in peak_names:
                column_name = f"{peak_name} At(%)"
                if peak_name in sheet_peak_data:
                    profile_data[column_name].append(sheet_peak_data[peak_name])
                else:
                    profile_data[column_name].append(0.0)

        df = pd.DataFrame(profile_data)
        cols = ['Number'] + [col for col in df.columns if col != 'Number']
        df = df[cols]
        df = df.sort_values(by='Number').reset_index(drop=True)

        self._save_profile(df, profile_sheet_name, "Atomic Concentration (%)", "Number", "Atomic_Concentration")

    def _create_area_profile_from_results(self, selected_sheets, profile_sheet_name):
        """Create area profile from results grid"""
        profile_data = {}
        profile_data['Number'] = []
        peak_names = set()

        for sheet_name in selected_sheets:
            num_match = re.search(r'(\d+)$', sheet_name)
            number = int(num_match.group(1)) if num_match else 0
            profile_data['Number'].append(number)

            grid = self.parent.results_grid
            num_rows = grid.GetNumberRows()

            for row in range(num_rows):
                try:
                    result_sheet = grid.GetCellValue(row, 18)
                    if result_sheet == sheet_name:
                        peak_name = grid.GetCellValue(row, 0)
                        peak_names.add(peak_name)
                except (ValueError, IndexError):
                    continue

        peak_names = sorted(list(peak_names))
        for peak_name in peak_names:
            profile_data[f"{peak_name} Area"] = []

        for sheet_name in selected_sheets:
            sheet_peak_data = {}
            grid = self.parent.results_grid
            num_rows = grid.GetNumberRows()

            for row in range(num_rows):
                try:
                    result_sheet = grid.GetCellValue(row, 18)
                    if result_sheet == sheet_name:
                        peak_name = grid.GetCellValue(row, 0)
                        area_str = grid.GetCellValue(row, 5)  # Corrected Area
                        if area_str and area_str.strip():
                            sheet_peak_data[peak_name] = float(area_str)
                except (ValueError, IndexError):
                    continue

            for peak_name in peak_names:
                column_name = f"{peak_name} Area"
                if peak_name in sheet_peak_data:
                    profile_data[column_name].append(sheet_peak_data[peak_name])
                else:
                    profile_data[column_name].append(0.0)

        df = pd.DataFrame(profile_data)
        cols = ['Number'] + [col for col in df.columns if col != 'Number']
        df = df[cols]
        df = df.sort_values(by='Number').reset_index(drop=True)

        self._save_profile(df, profile_sheet_name, "Area (CPS)", "Number", "Area")

    def _create_position_profile_from_results(self, selected_sheets, profile_sheet_name):
        """Create position profile from results grid"""
        profile_data = {}
        profile_data['Number'] = []
        peak_names = set()

        for sheet_name in selected_sheets:
            num_match = re.search(r'(\d+)$', sheet_name)
            number = int(num_match.group(1)) if num_match else 0
            profile_data['Number'].append(number)

            grid = self.parent.results_grid
            num_rows = grid.GetNumberRows()

            for row in range(num_rows):
                try:
                    result_sheet = grid.GetCellValue(row, 18)
                    if result_sheet == sheet_name:
                        peak_name = grid.GetCellValue(row, 0)
                        peak_names.add(peak_name)
                except (ValueError, IndexError):
                    continue

        peak_names = sorted(list(peak_names))
        for peak_name in peak_names:
            profile_data[f"{peak_name} Pos"] = []

        for sheet_name in selected_sheets:
            sheet_peak_data = {}
            grid = self.parent.results_grid
            num_rows = grid.GetNumberRows()

            for row in range(num_rows):
                try:
                    result_sheet = grid.GetCellValue(row, 18)
                    if result_sheet == sheet_name:
                        peak_name = grid.GetCellValue(row, 0)
                        pos_str = grid.GetCellValue(row, 1)  # Position
                        if pos_str and pos_str.strip():
                            sheet_peak_data[peak_name] = float(pos_str)
                except (ValueError, IndexError):
                    continue

            for peak_name in peak_names:
                column_name = f"{peak_name} Pos"
                if peak_name in sheet_peak_data:
                    profile_data[column_name].append(sheet_peak_data[peak_name])
                else:
                    profile_data[column_name].append(0.0)

        df = pd.DataFrame(profile_data)
        cols = ['Number'] + [col for col in df.columns if col != 'Number']
        df = df[cols]
        df = df.sort_values(by='Number').reset_index(drop=True)

        self._save_profile(df, profile_sheet_name, "Position (eV)", "Number", "Position")

    def _save_profile(self, df, profile_sheet_name, y_label, x_label, profile_type):
        """Save profile to file and switch to it"""
        from libraries.FileMenu.Save import convert_to_serializable_and_round
        from libraries.Sheet_Operations import on_sheet_selected

        file_path = self.parent.Data['FilePath']

        try:
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=profile_sheet_name, index=False)

            first_column = df.columns[1] if len(df.columns) > 1 else df.columns[0]
            self.parent.Data['Core levels'][profile_sheet_name] = {
                'Name': profile_sheet_name,
                'B.E.': df[df.columns[0]].tolist(),
                'Raw Data': df[first_column].tolist(),
                'Profile Data': df.to_dict('list'),
                'Y_Axis_Label': y_label,
                'X_Axis_Label': x_label,
                'Profile_Type': profile_type,
                'Background': {
                    'Bkg Type': '',
                    'Bkg Low': '',
                    'Bkg High': '',
                    'Bkg Offset Low': '',
                    'Bkg Offset High': '',
                    'Bkg Y': df[first_column].tolist()
                }
            }

            json_file_path = os.path.splitext(file_path)[0] + '.json'
            json_data = convert_to_serializable_and_round(self.parent.Data)
            with open(json_file_path, 'w') as json_file:
                json.dump(json_data, json_file, indent=2)

            if profile_sheet_name not in self.parent.sheet_combobox.GetStrings():
                self.parent.sheet_combobox.Append(profile_sheet_name)

            self.parent.sheet_combobox.SetValue(profile_sheet_name)
            on_sheet_selected(self.parent, profile_sheet_name)

            # Update current editor
            self.sheet_name = profile_sheet_name
            self.load_data()

            wx.MessageBox(
                f"Profile '{profile_sheet_name}' created successfully!",
                "Success",
                wx.OK | wx.ICON_INFORMATION
            )

        except Exception as e:
            wx.MessageBox(f"Error creating profile: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

    def load_data(self):
        """Load profile data from window.Data"""
        if not self.sheet_name or self.sheet_name not in self.parent.Data['Core levels']:
            return

        profile_info = self.parent.Data['Core levels'][self.sheet_name]

        # Load axis labels
        self.x_axis_ctrl.SetValue(profile_info.get('X_Axis_Label', 'Number'))
        self.y_axis_ctrl.SetValue(profile_info.get('Y_Axis_Label', 'Atomic Concentration (%)'))

        # Load profile data
        if 'Profile Data' not in profile_info:
            return

        profile_data = profile_info['Profile Data']
        df = pd.DataFrame(profile_data)

        # Set up grid
        num_rows = len(df)
        num_cols = len(df.columns)

        if self.data_grid.GetNumberRows() > 0:
            self.data_grid.DeleteRows(0, self.data_grid.GetNumberRows())
        if self.data_grid.GetNumberCols() > 0:
            self.data_grid.DeleteCols(0, self.data_grid.GetNumberCols())

        self.data_grid.AppendRows(num_rows)
        self.data_grid.AppendCols(num_cols)

        # Set column headers
        for col_idx, col_name in enumerate(df.columns):
            self.data_grid.SetColLabelValue(col_idx, col_name)
            self.data_grid.SetColSize(col_idx, 120)

        # Fill data
        for row_idx in range(num_rows):
            for col_idx, col_name in enumerate(df.columns):
                value = df.iloc[row_idx, col_idx]
                self.data_grid.SetCellValue(row_idx, col_idx, f"{float(value):.2f}")

        self.data_grid.AutoSizeColumns()

    def on_add_row(self, event):
        """Add a new row at the end"""
        num_cols = self.data_grid.GetNumberCols()
        if num_cols == 0:
            wx.MessageBox("Please add columns first", "Info", wx.OK | wx.ICON_INFORMATION)
            return

        self.data_grid.AppendRows(1)
        new_row = self.data_grid.GetNumberRows() - 1

        # Fill with default values (0.00)
        for col in range(num_cols):
            self.data_grid.SetCellValue(new_row, col, "0.00")

    def on_remove_row(self, event):
        """Remove the last row"""
        num_rows = self.data_grid.GetNumberRows()
        if num_rows == 0:
            wx.MessageBox("No rows to remove", "Info", wx.OK | wx.ICON_INFORMATION)
            return

        response = wx.MessageBox(f"Remove the last row?", "Confirm", wx.YES_NO | wx.ICON_QUESTION)
        if response == wx.YES:
            self.data_grid.DeleteRows(num_rows - 1, 1)

    def on_add_column(self, event):
        """Add a new column"""
        dlg = wx.TextEntryDialog(self, "Enter column name:", "Add Column")
        if dlg.ShowModal() == wx.ID_OK:
            col_name = dlg.GetValue().strip()
            if not col_name:
                wx.MessageBox("Column name cannot be empty", "Error", wx.OK | wx.ICON_ERROR)
                dlg.Destroy()
                return

            # Check if column name already exists
            existing_cols = [self.data_grid.GetColLabelValue(i) for i in range(self.data_grid.GetNumberCols())]
            if col_name in existing_cols:
                wx.MessageBox("Column name already exists", "Error", wx.OK | wx.ICON_ERROR)
                dlg.Destroy()
                return

            self.data_grid.AppendCols(1)
            new_col = self.data_grid.GetNumberCols() - 1
            self.data_grid.SetColLabelValue(new_col, col_name)
            self.data_grid.SetColSize(new_col, 120)

            # Fill with default values
            for row in range(self.data_grid.GetNumberRows()):
                self.data_grid.SetCellValue(row, new_col, "0.00")

        dlg.Destroy()

    def on_remove_column(self, event):
        """Remove a column"""
        num_cols = self.data_grid.GetNumberCols()
        if num_cols == 0:
            wx.MessageBox("No columns to remove", "Info", wx.OK | wx.ICON_INFORMATION)
            return

        # Show list of columns
        col_names = [self.data_grid.GetColLabelValue(i) for i in range(num_cols)]
        dlg = wx.SingleChoiceDialog(self, "Select column to remove:", "Remove Column", col_names)

        if dlg.ShowModal() == wx.ID_OK:
            col_idx = dlg.GetSelection()
            col_name = col_names[col_idx]

            response = wx.MessageBox(f"Remove column '{col_name}'?", "Confirm", wx.YES_NO | wx.ICON_QUESTION)
            if response == wx.YES:
                self.data_grid.DeleteCols(col_idx, 1)

        dlg.Destroy()

    def on_edit_header(self, event):
        """Edit a column header"""
        num_cols = self.data_grid.GetNumberCols()
        if num_cols == 0:
            wx.MessageBox("No columns to edit", "Info", wx.OK | wx.ICON_INFORMATION)
            return

        # Show list of columns
        col_names = [self.data_grid.GetColLabelValue(i) for i in range(num_cols)]
        dlg = wx.SingleChoiceDialog(self, "Select column to rename:", "Edit Column Header", col_names)

        if dlg.ShowModal() == wx.ID_OK:
            col_idx = dlg.GetSelection()
            old_name = col_names[col_idx]

            # Get new name
            new_dlg = wx.TextEntryDialog(self, f"Enter new name for '{old_name}':", "Rename Column", old_name)
            if new_dlg.ShowModal() == wx.ID_OK:
                new_name = new_dlg.GetValue().strip()
                if not new_name:
                    wx.MessageBox("Column name cannot be empty", "Error", wx.OK | wx.ICON_ERROR)
                else:
                    # Check if new name already exists (excluding current column)
                    existing_cols = [self.data_grid.GetColLabelValue(i) for i in range(num_cols) if i != col_idx]
                    if new_name in existing_cols:
                        wx.MessageBox("Column name already exists", "Error", wx.OK | wx.ICON_ERROR)
                    else:
                        self.data_grid.SetColLabelValue(col_idx, new_name)
            new_dlg.Destroy()

        dlg.Destroy()

    def on_header_double_click(self, event):
        """Handle double-click on column header"""
        col_idx = event.GetCol()
        if col_idx < 0:
            return

        old_name = self.data_grid.GetColLabelValue(col_idx)
        dlg = wx.TextEntryDialog(self, f"Enter new name for '{old_name}':", "Rename Column", old_name)

        if dlg.ShowModal() == wx.ID_OK:
            new_name = dlg.GetValue().strip()
            if not new_name:
                wx.MessageBox("Column name cannot be empty", "Error", wx.OK | wx.ICON_ERROR)
            else:
                # Check if new name already exists (excluding current column)
                existing_cols = [self.data_grid.GetColLabelValue(i)
                                 for i in range(self.data_grid.GetNumberCols()) if i != col_idx]
                if new_name in existing_cols:
                    wx.MessageBox("Column name already exists", "Error", wx.OK | wx.ICON_ERROR)
                else:
                    self.data_grid.SetColLabelValue(col_idx, new_name)

        dlg.Destroy()

    def on_save(self, event):
        """Save changes back to window.Data and Excel"""
        from libraries.FileMenu.Save import convert_to_serializable_and_round

        try:
            # Update axis labels
            x_label = self.x_axis_ctrl.GetValue()
            y_label = self.y_axis_ctrl.GetValue()

            # Get data from grid
            num_rows = self.data_grid.GetNumberRows()
            num_cols = self.data_grid.GetNumberCols()

            if num_rows == 0 or num_cols == 0:
                wx.MessageBox("Grid is empty. Cannot save.", "Error", wx.OK | wx.ICON_ERROR)
                return

            # Build new dataframe
            data_dict = {}
            for col_idx in range(num_cols):
                col_name = self.data_grid.GetColLabelValue(col_idx)
                col_data = []
                for row_idx in range(num_rows):
                    value = self.data_grid.GetCellValue(row_idx, col_idx)
                    try:
                        col_data.append(float(value))
                    except ValueError:
                        col_data.append(0.0)
                data_dict[col_name] = col_data

            df = pd.DataFrame(data_dict)

            # Update window.Data
            first_data_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]

            self.parent.Data['Core levels'][self.sheet_name]['X_Axis_Label'] = x_label
            self.parent.Data['Core levels'][self.sheet_name]['Y_Axis_Label'] = y_label
            self.parent.Data['Core levels'][self.sheet_name]['Profile Data'] = df.to_dict('list')
            self.parent.Data['Core levels'][self.sheet_name]['B.E.'] = df[df.columns[0]].tolist()
            self.parent.Data['Core levels'][self.sheet_name]['Raw Data'] = df[first_data_col].tolist()
            self.parent.Data['Core levels'][self.sheet_name]['Background']['Bkg Y'] = df[first_data_col].tolist()

            # Save to Excel
            file_path = self.parent.Data['FilePath']
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=self.sheet_name, index=False)

            # Save to JSON
            json_file_path = os.path.splitext(file_path)[0] + '.json'
            json_data = convert_to_serializable_and_round(self.parent.Data)
            with open(json_file_path, 'w') as json_file:
                json.dump(json_data, json_file, indent=2)

            # Replot
            self.parent.plot_manager.plot_data(self.parent)

            wx.MessageBox("Profile data saved successfully!", "Success", wx.OK | wx.ICON_INFORMATION)
            self.Close()

        except Exception as e:
            wx.MessageBox(f"Error saving profile data: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

    def _get_next_available_profile_name(self, base_name="zzProfile"):
        """Find the next available profile name"""
        existing_profiles = [name for name in self.parent.Data['Core levels'].keys()
                             if name.startswith('zzProfile')]

        if base_name not in existing_profiles:
            return base_name

        # Find next available number
        counter = 1
        while True:
            new_name = f"{base_name}{counter}"
            if new_name not in existing_profiles:
                return new_name
            counter += 1

    def on_cancel(self, event):
        """Close without saving"""
        self.Close()