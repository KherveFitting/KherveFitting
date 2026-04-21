# KherveFitting - XPS Data Analysis Software
# Copyright (C) 2024-2026 Gwilherm Kerherve <g.kerherve@ic.ac.uk>
#
# KherveFitting is dual-licensed:
#   - GNU GPL v3.0 (see LICENSE-GPL.txt) for open-source use
#   - Commercial Licence (see LICENSE-COMMERCIAL.txt) for proprietary use
# SPDX-License-Identifier: GPL-3.0-only OR LicenseRef-KherveFitting-Commercial

import wx
import wx.grid
import pandas as pd
import json
import os
import re
import wx.lib.agw.flatnotebook as fnb


class ProfileCreatorWindow(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, title="Profile Data Creator", size=(280, 300), style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX | wx.MINIMIZE_BOX | wx.SYSTEM_MENU) | wx.STAY_ON_TOP)
        self.parent = parent
        self.Centre()
        self.init_ui()

    def init_ui(self):
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
            panel.SetBackgroundColour(wx.Colour(250, 220, 240))

        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Select/Unselect buttons
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        select_all_btn = wx.Button(panel, label="Select All")
        select_all_btn.SetMinSize((125, 35))
        select_all_btn.Bind(wx.EVT_BUTTON, self.on_select_all)
        button_sizer.Add(select_all_btn, 1, wx.ALL | wx.EXPAND, 1)

        unselect_all_btn = wx.Button(panel, label="Unselect All")
        unselect_all_btn.SetMinSize((125, 35))
        unselect_all_btn.Bind(wx.EVT_BUTTON, self.on_unselect_all)
        button_sizer.Add(unselect_all_btn, 1, wx.ALL | wx.EXPAND, 0)
        main_sizer.Add(button_sizer, 0, wx.ALL | wx.EXPAND, 1)

        # Core levels checklist
        self.core_levels_checklist = wx.CheckListBox(panel)
        self.core_levels_checklist.SetMinSize((250, 200))
        self.populate_core_levels_list()
        self.core_levels_checklist.Bind(wx.EVT_CONTEXT_MENU, self.on_core_levels_context_menu)
        main_sizer.Add(self.core_levels_checklist, 1, wx.ALL | wx.EXPAND, 3)

        # Profile type dropdown
        type_sizer = wx.BoxSizer(wx.HORIZONTAL)
        type_label = wx.StaticText(panel, label="Type:")
        type_sizer.Add(type_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 1)
        self.profile_type_combo = wx.ComboBox(panel, choices=["Concentration", "Area", "Position", "FWHM", "Height", "L/G", "Corrected Area", "Weight (%)"],
                                              style=wx.CB_READONLY, value="Concentration")
        type_sizer.Add(self.profile_type_combo, 1, wx.ALL | wx.EXPAND, 1)
        main_sizer.Add(type_sizer, 0, wx.ALL | wx.EXPAND, 1)

        # 2x2 arrangement for Create Profile buttons
        create_buttons_sizer = wx.GridSizer(2, 2, 1, 1)

        # Create from Peak Fitting Grid
        create_peak_btn = wx.Button(panel, label="Create Profile from\nPeak Fitting Grid")
        create_peak_btn.SetMinSize((120, 40))
        create_peak_btn.Bind(wx.EVT_BUTTON, self.on_create_from_peak_grid)
        create_buttons_sizer.Add(create_peak_btn, 0, wx.EXPAND)

        # Create from Results Grid
        create_results_btn = wx.Button(panel, label="Create Profile from\nResults Grid")
        create_results_btn.SetMinSize((120, 40))
        create_results_btn.Bind(wx.EVT_BUTTON, self.on_create_from_results_grid)
        create_buttons_sizer.Add(create_results_btn, 0, wx.EXPAND)

        # Edit Profiles button (spans both columns in second row)
        edit_profiles_btn = wx.Button(panel, label="Edit Profiles")
        edit_profiles_btn.SetMinSize((120, 40))
        edit_profiles_btn.Bind(wx.EVT_BUTTON, self.on_edit_profiles)
        create_buttons_sizer.Add(edit_profiles_btn, 0, wx.EXPAND)

        # # Add empty space for symmetry
        # create_buttons_sizer.Add((0, 0))

        main_sizer.Add(create_buttons_sizer, 0, wx.ALL | wx.EXPAND, 1)

        panel.SetSizer(main_sizer)
        self.Bind(wx.EVT_CLOSE, self.on_close)

    def populate_core_levels_list(self):
        """Populate checklist with core levels"""
        self.core_levels_checklist.Clear()
        if hasattr(self.parent, 'Data') and 'Core levels' in self.parent.Data:
            for core_level in self.parent.Data['Core levels'].keys():
                if not core_level.startswith('zzProfile'):
                    self.core_levels_checklist.Append(core_level)

    def on_select_all(self, event):
        """Select all core levels"""
        for i in range(self.core_levels_checklist.GetCount()):
            self.core_levels_checklist.Check(i, True)

    def on_unselect_all(self, event):
        """Unselect all core levels"""
        for i in range(self.core_levels_checklist.GetCount()):
            self.core_levels_checklist.Check(i, False)

    def on_core_levels_context_menu(self, event):
        """Handle right-click context menu"""
        menu = wx.Menu()

        # Get all core levels
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

        # Add Select All / Unselect All at the top
        select_all_item = menu.Append(wx.ID_ANY, "Select All")
        unselect_all_item = menu.Append(wx.ID_ANY, "Unselect All")

        self.Bind(wx.EVT_MENU, self.on_select_all, select_all_item)
        self.Bind(wx.EVT_MENU, self.on_unselect_all, unselect_all_item)

        menu.AppendSeparator()

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

        self.PopupMenu(menu)
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

        # Special handling for Survey sheets - group all surveys together
        if cleaned_name.startswith('Survey'):
            return 'Survey'

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


    def on_edit_profiles(self, event):
        """Open the profile editor window"""
        if hasattr(self, 'editor_window') and self.editor_window:
            self.editor_window.Raise()
        else:
            self.editor_window = ProfileEditWindow(self.parent)
            self.editor_window.Show()

    def on_create_from_peak_grid_OLD(self, event):
        """Create profile from peak fitting grid"""
        import pandas as pd
        import os
        import json
        import re

        # Get selected core levels
        checked_levels = []
        for i in range(self.core_levels_checklist.GetCount()):
            if self.core_levels_checklist.IsChecked(i):
                checked_levels.append(self.core_levels_checklist.GetString(i))

        if not checked_levels:
            wx.MessageBox("No core levels selected", "Create Profile Failed", wx.OK | wx.ICON_WARNING)
            return

        profile_type = self.profile_type_combo.GetValue()

        # Extract the lowest number from selected sheets
        lowest_number = float('inf')
        for sheet in checked_levels:
            num_match = re.search(r'(\d+)$', sheet)
            if num_match:
                number = int(num_match.group(1))
            else:
                number = 0  # No number suffix means it's the first one (0)
            lowest_number = min(lowest_number, number)

        # Find next available profile name starting from lowest_number
        counter = lowest_number
        while True:
            if counter == 0:
                profile_sheet_name = "zzProfile"
            else:
                profile_sheet_name = f"zzProfile{counter}"

            # If this name doesn't exist, use it
            if profile_sheet_name not in self.parent.Data['Core levels']:
                break

            # Otherwise, try next number
            counter += 1

        try:
            # Collect data from all selected sheets
            profile_data = {}
            profile_data['Number'] = []
            peak_names = set()

            # First pass: collect all unique peak names and assign numbers
            for idx, sheet_name in enumerate(checked_levels):
                # Extract number from sheet name (Survey=0, Survey1=1, etc.)
                num_match = re.search(r'(\d+)$', sheet_name)
                if num_match:
                    number = int(num_match.group(1))
                else:
                    number = 0  # No number suffix means it's the first one

                profile_data['Number'].append(number)

                # Get peaks from this sheet
                if (sheet_name in self.parent.Data['Core levels'] and
                        'Fitting' in self.parent.Data['Core levels'][sheet_name] and
                        'Peaks' in self.parent.Data['Core levels'][sheet_name]['Fitting']):

                    peaks = self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                    for peak_name in peaks.keys():
                        peak_names.add(peak_name)

            # Initialize columns for each peak based on profile type
            peak_names = sorted(list(peak_names))
            for peak_name in peak_names:
                if profile_type == "Concentration":
                    column_name = f"{peak_name}\n At.(%)"
                elif profile_type == "Area":
                    column_name = f"{peak_name}/ Area (CPS) "
                elif profile_type == "Position":
                    column_name = f"{peak_name} / Position (eV)"
                profile_data[column_name] = []

            # Second pass: populate data based on profile type
            for sheet_name in checked_levels:
                sheet_peak_data = {}

                # Switch to sheet and get data from peak_params_grid
                self.parent.sheet_combobox.SetValue(sheet_name)
                from libraries.Sheet_Operations import on_sheet_selected
                on_sheet_selected(self.parent, sheet_name)

                # For Concentration, update ratios to ensure atomic concentrations are calculated
                if profile_type == "Concentration":
                    self.parent.update_ratios()

                # Extract data from grid based on profile type
                grid = self.parent.peak_params_grid
                num_peaks = grid.GetNumberRows() // 2

                for i in range(num_peaks):
                    row = i * 2
                    try:
                        peak_name = grid.GetCellValue(row, 1)  # Peak name from column 1

                        if profile_type == "Concentration":
                            # Atomic % from column 10
                            value_str = grid.GetCellValue(row, 10)
                        elif profile_type == "Area":
                            # Area from column 6
                            value_str = grid.GetCellValue(row, 6)
                        elif profile_type == "Position":
                            # Position from column 2
                            value_str = grid.GetCellValue(row, 2)

                        if value_str and value_str.strip():
                            value = float(value_str)
                            sheet_peak_data[peak_name] = value
                    except (ValueError, IndexError):
                        continue

                # Fill in data for this sheet
                for peak_name in peak_names:
                    if profile_type == "Concentration":
                        column_name = f"{peak_name}\n At.(%)"
                    elif profile_type == "Area":
                        column_name = f"{peak_name} Area"
                    elif profile_type == "Position":
                        column_name = f"{peak_name} Position"

                    if peak_name in sheet_peak_data:
                        profile_data[column_name].append(sheet_peak_data[peak_name])
                    else:
                        profile_data[column_name].append(0.0)  # No data for this peak

            # Create DataFrame
            df = pd.DataFrame(profile_data)

            # Ensure Number column is first
            cols = ['Number'] + [col for col in df.columns if col != 'Number']
            df = df[cols]

            # SORT by Number column to ensure proper line plotting
            df = df.sort_values(by='Number').reset_index(drop=True)

            # Save to Excel
            file_path = self.parent.Data['FilePath']

            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=profile_sheet_name, index=False)

            # Update Data structure with profile metadata
            first_column = f"{peak_names[0]}\n At.(%)" if profile_type == "Concentration" and peak_names else f"{peak_names[0]} {profile_type}" if peak_names else 'Number'

            if profile_type == "Concentration":
                y_axis_label = 'Atomic Concentration (%)'
            elif profile_type == "Area":
                y_axis_label = 'Area (CPS.eV)'
            elif profile_type == "Position":
                y_axis_label = 'Position (eV)'

            self.parent.Data['Core levels'][profile_sheet_name] = {
                'Name': profile_sheet_name,
                'B.E.': df['Number'].tolist(),
                'Raw Data': df[first_column].tolist() if peak_names else [],
                'Profile Data': df.to_dict('list'),
                'Y_Axis_Label': y_axis_label,
                'X_Axis_Label': 'Number',
                'Profile_Type': profile_type,
                'Background': {
                    'Bkg Type': '',
                    'Bkg Low': '',
                    'Bkg High': '',
                    'Bkg Offset Low': '',
                    'Bkg Offset High': '',
                    'Bkg Y': df[first_column].tolist() if peak_names else []
                }
            }

            # Save to JSON
            json_file_path = os.path.splitext(file_path)[0] + '.json'
            from libraries.FileMenu.Save import convert_to_serializable_and_round
            json_data = convert_to_serializable_and_round(self.parent.Data)
            with open(json_file_path, 'w') as json_file:
                json.dump(json_data, json_file, indent=2)

            # Update sheet list
            if profile_sheet_name not in self.parent.sheet_combobox.GetStrings():
                self.parent.sheet_combobox.Append(profile_sheet_name)

            # Switch to the new profile sheet and plot it
            self.parent.sheet_combobox.SetValue(profile_sheet_name)
            from libraries.Sheet_Operations import on_sheet_selected
            on_sheet_selected(self.parent, profile_sheet_name)

            wx.MessageBox(
                f"Profile '{profile_sheet_name}' created successfully!\n"
                f"Data from {len(checked_levels)} sheets with {len(peak_names)} peaks.",
                "Success",
                wx.OK | wx.ICON_INFORMATION
            )

        except Exception as e:
            wx.MessageBox(f"Error creating profile: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

    def on_create_from_peak_grid(self, event):
        """Create profile from peak fitting grid"""
        import pandas as pd
        import os
        import json
        import re

        # Get selected core levels
        checked_levels = []
        for i in range(self.core_levels_checklist.GetCount()):
            if self.core_levels_checklist.IsChecked(i):
                checked_levels.append(self.core_levels_checklist.GetString(i))

        if not checked_levels:
            wx.MessageBox("No core levels selected", "Create Profile Failed", wx.OK | wx.ICON_WARNING)
            return

        profile_type = self.profile_type_combo.GetValue()

        # Extract the lowest number from selected sheets
        lowest_number = float('inf')
        for sheet in checked_levels:
            num_match = re.search(r'(\d+)$', sheet)
            if num_match:
                number = int(num_match.group(1))
            else:
                number = 0  # No number suffix means it's the first one (0)
            lowest_number = min(lowest_number, number)

        # Find next available profile name starting from lowest_number
        counter = lowest_number
        while True:
            if counter == 0:
                profile_sheet_name = "zzProfile"
            else:
                profile_sheet_name = f"zzProfile{counter}"

            # If this name doesn't exist, use it
            if profile_sheet_name not in self.parent.Data['Core levels']:
                break

            # Otherwise, try next number
            counter += 1

        try:
            # Collect data from all selected sheets
            profile_data = {}
            profile_data['Number'] = []
            peak_names = set()

            # First pass: collect all unique peak names and assign numbers
            for idx, sheet_name in enumerate(checked_levels):
                # Extract number from sheet name (Survey=0, Survey1=1, etc.)
                num_match = re.search(r'(\d+)$', sheet_name)
                if num_match:
                    number = int(num_match.group(1))
                else:
                    number = 0  # No number suffix means it's the first one

                profile_data['Number'].append(number)

                # Get peaks from this sheet
                if (sheet_name in self.parent.Data['Core levels'] and
                        'Fitting' in self.parent.Data['Core levels'][sheet_name] and
                        'Peaks' in self.parent.Data['Core levels'][sheet_name]['Fitting']):

                    peaks = self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']
                    for peak_name in peaks.keys():
                        peak_names.add(peak_name)

            # Initialize columns for each peak based on profile type
            peak_names = sorted(list(peak_names))
            for peak_name in peak_names:
                if profile_type == "Concentration":
                    column_name = f"{peak_name}\n At. (%)"
                elif profile_type == "Area":
                    column_name = f"{peak_name}\n Area (CPS) "
                elif profile_type == "Position":
                    column_name = f"{peak_name}\n Position (eV)"
                elif profile_type == "FWHM":
                    column_name = f"{peak_name}\n FWHM (eV)"
                elif profile_type == "Height":
                    column_name = f"{peak_name}\n Height (CPS)"
                elif profile_type == "L/G":
                    column_name = f"{peak_name}\n L/G (%)"
                profile_data[column_name] = []

            # Second pass: populate data based on profile type
            for sheet_name in checked_levels:
                sheet_peak_data = {}

                # Switch to sheet and get data from peak_params_grid
                self.parent.sheet_combobox.SetValue(sheet_name)
                from libraries.Sheet_Operations import on_sheet_selected
                on_sheet_selected(self.parent, sheet_name)

                # For Concentration, update ratios to ensure atomic concentrations are calculated
                if profile_type == "Concentration":
                    self.parent.update_ratios()

                # Extract data from grid based on profile type
                grid = self.parent.peak_params_grid
                num_peaks = grid.GetNumberRows() // 2

                for i in range(num_peaks):
                    row = i * 2
                    try:
                        peak_name = grid.GetCellValue(row, 1)  # Peak name from column 1

                        if profile_type == "Concentration":
                            # Atomic % from column 10
                            value_str = grid.GetCellValue(row, 10)
                        elif profile_type == "Area":
                            # Area from column 6
                            value_str = grid.GetCellValue(row, 6)
                        elif profile_type == "Position":
                            # Position from column 2
                            value_str = grid.GetCellValue(row, 2)
                        elif profile_type == "FWHM":
                            # FWHM from column 4
                            value_str = grid.GetCellValue(row, 4)
                        elif profile_type == "Height":
                            # Height from column 3
                            value_str = grid.GetCellValue(row, 3)
                        elif profile_type == "L/G":
                            # L/G from column 5
                            value_str = grid.GetCellValue(row, 5)

                        if value_str and value_str.strip():
                            value = float(value_str)
                            sheet_peak_data[peak_name] = value
                    except (ValueError, IndexError):
                        continue

                # Fill in data for this sheet
                for peak_name in peak_names:
                    if profile_type == "Concentration":
                        column_name = f"{peak_name}\n At. (%)"
                    elif profile_type == "Area":
                        column_name = f"{peak_name}\n Area (CPS) "
                    elif profile_type == "Position":
                        column_name = f"{peak_name}\n Position (eV)"
                    elif profile_type == "FWHM":
                        column_name = f"{peak_name}\n FWHM (eV)"
                    elif profile_type == "Height":
                        column_name = f"{peak_name}\n Height (CPS)"
                    elif profile_type == "L/G":
                        column_name = f"{peak_name}\n L/G (%)"

                    if peak_name in sheet_peak_data:
                        profile_data[column_name].append(sheet_peak_data[peak_name])
                    else:
                        profile_data[column_name].append(0.0)  # No data for this peak

            # Create DataFrame
            df = pd.DataFrame(profile_data)

            # Format all numeric columns to .2f
            for col in df.columns:
                if col != 'Number':
                    df[col] = df[col].apply(lambda x: float(f"{x:.2f}"))

            # Ensure Number column is first
            cols = ['Number'] + [col for col in df.columns if col != 'Number']
            df = df[cols]

            # SORT by Number column to ensure proper line plotting
            df = df.sort_values(by='Number').reset_index(drop=True)

            # Save to Excel
            file_path = self.parent.Data['FilePath']

            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=profile_sheet_name, index=False)

            # Update Data structure with profile metadata
            if peak_names:
                if profile_type == "Concentration":
                    first_column = f"{peak_names[0]}\n At. (%)"
                elif profile_type == "Area":
                    first_column = f"{peak_names[0]}\n Area (CPS) "
                elif profile_type == "Position":
                    first_column = f"{peak_names[0]}\n Position (eV)"
                elif profile_type == "FWHM":
                    first_column = f"{peak_names[0]}\n FWHM (eV)"
                elif profile_type == "Height":
                    first_column = f"{peak_names[0]}\n Height (CPS)"
                elif profile_type == "L/G":
                    first_column = f"{peak_names[0]}\n L/G (%)"
            else:
                first_column = 'Number'

            if profile_type == "Concentration":
                y_axis_label = 'Atomic Concentration (%)'
            elif profile_type == "Area":
                y_axis_label = 'Area (CPS.eV)'
            elif profile_type == "Position":
                y_axis_label = 'Position (eV)'
            elif profile_type == "FWHM":
                y_axis_label = 'FWHM (eV)'
            elif profile_type == "Height":
                y_axis_label = 'Height (CPS)'
            elif profile_type == "L/G":
                y_axis_label = 'L/G (%)'

            self.parent.Data['Core levels'][profile_sheet_name] = {
                'Name': profile_sheet_name,
                'B.E.': df['Number'].tolist(),
                'Raw Data': df[first_column].tolist() if peak_names else [],
                'Profile Data': df.to_dict('list'),
                'Y_Axis_Label': y_axis_label,
                'X_Axis_Label': 'Number',
                'Profile_Type': profile_type,
                'Background': {
                    'Bkg Type': '',
                    'Bkg Low': '',
                    'Bkg High': '',
                    'Bkg Offset Low': '',
                    'Bkg Offset High': '',
                    'Bkg Y': df[first_column].tolist() if peak_names else []
                }
            }

            # Save to JSON
            json_file_path = os.path.splitext(file_path)[0] + '.json'
            from libraries.FileMenu.Save import convert_to_serializable_and_round
            json_data = convert_to_serializable_and_round(self.parent.Data)
            with open(json_file_path, 'w') as json_file:
                json.dump(json_data, json_file, indent=2)

            # Update sheet list
            if profile_sheet_name not in self.parent.sheet_combobox.GetStrings():
                self.parent.sheet_combobox.Append(profile_sheet_name)

            # Switch to the new profile sheet and plot it
            self.parent.sheet_combobox.SetValue(profile_sheet_name)
            from libraries.Sheet_Operations import on_sheet_selected
            on_sheet_selected(self.parent, profile_sheet_name)

            wx.MessageBox(
                f"Profile '{profile_sheet_name}' created successfully!\n"
                f"Data from {len(checked_levels)} sheets with {len(peak_names)} peaks.",
                "Success",
                wx.OK | wx.ICON_INFORMATION
            )

        except Exception as e:
            wx.MessageBox(f"Error creating profile: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

    def on_create_from_results_grid_OLD(self, event):
        """Create profile from results grid - uses selected core levels to determine rows"""
        import pandas as pd
        import os
        import json
        import re

        # Get selected core levels
        checked_levels = []
        for i in range(self.core_levels_checklist.GetCount()):
            if self.core_levels_checklist.IsChecked(i):
                checked_levels.append(self.core_levels_checklist.GetString(i))

        if not checked_levels:
            wx.MessageBox("No core levels selected", "Create Profile Failed", wx.OK | wx.ICON_WARNING)
            return

        profile_type = self.profile_type_combo.GetValue()

        # Extract unique row numbers from selected core levels
        row_numbers = set()
        for sheet in checked_levels:
            num_match = re.search(r'(\d+)$', sheet)
            if num_match:
                row_numbers.add(int(num_match.group(1)))
            else:
                row_numbers.add(0)  # No number suffix means row 0

        row_numbers = sorted(list(row_numbers))

        if not row_numbers:
            wx.MessageBox("Could not determine row numbers from selected core levels", "Create Profile Failed", wx.OK | wx.ICON_WARNING)
            return

        # Find next available profile name
        counter = 0
        while True:
            if counter == 0:
                profile_sheet_name = "zzProfile"
            else:
                profile_sheet_name = f"zzProfile{counter}"

            if profile_sheet_name not in self.parent.Data['Core levels']:
                break
            counter += 1

        try:
            # Collect data from each row
            profile_data = {}
            profile_data['Number'] = row_numbers
            peak_names = set()

            # Determine which column to extract based on profile type
            if profile_type == "Concentration":
                data_column = 6  # Atomic (%)
                column_suffix = "At(%)"
            elif profile_type == "Area":
                data_column = 5  # Area (CPS.eV)
                column_suffix = "Area"
            elif profile_type == "Position":
                data_column = 1  # Position (eV)
                column_suffix = "Position"

            # First pass: collect all unique peak names across all rows
            for row_num in row_numbers:
                # Switch to the sheet for this row
                for sheet_name in self.parent.Data['Core levels'].keys():
                    num_match = re.search(r'(\d+)$', sheet_name)
                    if num_match:
                        sheet_row = int(num_match.group(1))
                    else:
                        sheet_row = 0

                    if sheet_row == row_num:
                        # Switch to this sheet
                        self.parent.sheet_combobox.SetValue(sheet_name)
                        from libraries.Sheet_Operations import on_sheet_selected
                        on_sheet_selected(self.parent, sheet_name)

                        # Get all peak names from results grid
                        grid = self.parent.results_grid
                        for row in range(grid.GetNumberRows()):
                            try:
                                peak_name = grid.GetCellValue(row, 0).strip()  # Peak Label
                                if peak_name:
                                    peak_names.add(peak_name)
                            except (ValueError, IndexError):
                                continue
                        break

            # Initialize columns for each peak
            peak_names = sorted(list(peak_names))
            for peak_name in peak_names:
                column_name = f"{peak_name} {column_suffix}"
                profile_data[column_name] = []

            # Second pass: populate data for each row
            for row_num in row_numbers:
                row_peak_data = {}

                # Find and switch to the sheet for this row
                for sheet_name in self.parent.Data['Core levels'].keys():
                    num_match = re.search(r'(\d+)$', sheet_name)
                    if num_match:
                        sheet_row = int(num_match.group(1))
                    else:
                        sheet_row = 0

                    if sheet_row == row_num:
                        # Switch to this sheet
                        self.parent.sheet_combobox.SetValue(sheet_name)
                        from libraries.Sheet_Operations import on_sheet_selected
                        on_sheet_selected(self.parent, sheet_name)

                        # Extract data from results grid
                        grid = self.parent.results_grid
                        for row in range(grid.GetNumberRows()):
                            try:
                                peak_name = grid.GetCellValue(row, 0).strip()
                                value_str = grid.GetCellValue(row, data_column).strip()

                                if peak_name and value_str:
                                    value = float(value_str)
                                    row_peak_data[peak_name] = value
                            except (ValueError, IndexError):
                                continue
                        break

                # Fill in data for this row
                for peak_name in peak_names:
                    column_name = f"{peak_name} {column_suffix}"
                    if peak_name in row_peak_data:
                        profile_data[column_name].append(row_peak_data[peak_name])
                    else:
                        profile_data[column_name].append(0.0)

            # Create DataFrame with .2f formatting
            df = pd.DataFrame(profile_data)

            # Format all numeric columns to .2f
            for col in df.columns:
                if col != 'Number':
                    df[col] = df[col].apply(lambda x: float(f"{x:.2f}"))

            # Ensure Number column is first
            cols = ['Number'] + [col for col in df.columns if col != 'Number']
            df = df[cols]

            # SORT by Number column
            df = df.sort_values(by='Number').reset_index(drop=True)

            # Save to Excel
            file_path = self.parent.Data['FilePath']

            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=profile_sheet_name, index=False)

            # Update Data structure with profile metadata
            first_column = df.columns[1] if len(df.columns) > 1 else 'Number'

            if profile_type == "Concentration":
                y_axis_label = 'Atomic Concentration (%)'
            elif profile_type == "Area":
                y_axis_label = 'Area (CPS.eV)'
            elif profile_type == "Position":
                y_axis_label = 'Position (eV)'

            self.parent.Data['Core levels'][profile_sheet_name] = {
                'Name': profile_sheet_name,
                'B.E.': df['Number'].tolist(),
                'Raw Data': df[first_column].tolist(),
                'Profile Data': df.to_dict('list'),
                'Y_Axis_Label': y_axis_label,
                'X_Axis_Label': 'Number',
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

            # Save to JSON
            json_file_path = os.path.splitext(file_path)[0] + '.json'
            from libraries.FileMenu.Save import convert_to_serializable_and_round
            json_data = convert_to_serializable_and_round(self.parent.Data)
            with open(json_file_path, 'w') as json_file:
                json.dump(json_data, json_file, indent=2)

            # Update sheet list
            if profile_sheet_name not in self.parent.sheet_combobox.GetStrings():
                self.parent.sheet_combobox.Append(profile_sheet_name)

            # Switch to the new profile sheet and plot it
            self.parent.sheet_combobox.SetValue(profile_sheet_name)
            from libraries.Sheet_Operations import on_sheet_selected
            on_sheet_selected(self.parent, profile_sheet_name)

            wx.MessageBox(
                f"Profile '{profile_sheet_name}' created successfully!\n"
                f"Data from {len(row_numbers)} rows with {len(peak_names)} peaks.",
                "Success",
                wx.OK | wx.ICON_INFORMATION
            )

        except Exception as e:
            wx.MessageBox(f"Error creating profile: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

    def on_create_from_results_grid(self, event):
        """Create profile from results grid - uses selected core levels to determine rows"""
        import pandas as pd
        import os
        import json
        import re

        # Get selected core levels
        checked_levels = []
        for i in range(self.core_levels_checklist.GetCount()):
            if self.core_levels_checklist.IsChecked(i):
                checked_levels.append(self.core_levels_checklist.GetString(i))

        if not checked_levels:
            wx.MessageBox("No core levels selected", "Create Profile Failed", wx.OK | wx.ICON_WARNING)
            return

        profile_type = self.profile_type_combo.GetValue()

        # Extract unique row numbers from selected core levels
        row_numbers = set()
        for sheet in checked_levels:
            num_match = re.search(r'(\d+)$', sheet)
            if num_match:
                row_numbers.add(int(num_match.group(1)))
            else:
                row_numbers.add(0)  # No number suffix means row 0

        row_numbers = sorted(list(row_numbers))

        if not row_numbers:
            wx.MessageBox("Could not determine row numbers from selected core levels", "Create Profile Failed", wx.OK | wx.ICON_WARNING)
            return

        # Find next available profile name
        counter = 0
        while True:
            if counter == 0:
                profile_sheet_name = "zzProfile"
            else:
                profile_sheet_name = f"zzProfile{counter}"

            if profile_sheet_name not in self.parent.Data['Core levels']:
                break
            counter += 1

        try:
            # Collect data from each row
            profile_data = {}
            profile_data['Number'] = row_numbers
            peak_names = set()

            # Determine which column to extract based on profile type
            if profile_type == "Concentration":
                data_column = 6  # Atomic (%)
                column_suffix = "At(%)"
            elif profile_type == "Area":
                data_column = 5  # Area (CPS.eV)
                column_suffix = "Area"
            elif profile_type == "Position":
                data_column = 1  # Position (eV)
                column_suffix = "Position"
            elif profile_type == "FWHM":
                data_column = 3  # FWHM (eV)
                column_suffix = "FWHM"
            elif profile_type == "Height":
                data_column = 2  # Height (CPS)
                column_suffix = "Height"
            elif profile_type == "L/G":
                data_column = 4  # L/G σ/γ (%)
                column_suffix = "L/G"
            elif profile_type == "Corrected Area":
                data_column = 13  # Corr. Area (a.u.)
                column_suffix = "Corrected Area"
            elif profile_type == "Weight (%)":
                data_column = 29  # Weight (%)
                column_suffix = "Weight (%)"

            # First pass: collect all unique peak names across all rows
            for row_num in row_numbers:
                # Switch to the sheet for this row
                for sheet_name in self.parent.Data['Core levels'].keys():
                    num_match = re.search(r'(\d+)$', sheet_name)
                    if num_match:
                        sheet_row = int(num_match.group(1))
                    else:
                        sheet_row = 0

                    if sheet_row == row_num:
                        # Switch to this sheet
                        self.parent.sheet_combobox.SetValue(sheet_name)
                        from libraries.Sheet_Operations import on_sheet_selected
                        on_sheet_selected(self.parent, sheet_name)

                        # Get all peak names from results grid
                        grid = self.parent.results_grid
                        for row in range(grid.GetNumberRows()):
                            try:
                                peak_name = grid.GetCellValue(row, 0).strip()  # Peak Label
                                if peak_name:
                                    peak_names.add(peak_name)
                            except (ValueError, IndexError):
                                continue
                        break

            # Initialize columns for each peak
            peak_names = sorted(list(peak_names))
            for peak_name in peak_names:
                column_name = f"{peak_name} {column_suffix}"
                profile_data[column_name] = []

            # Second pass: populate data for each row
            for row_num in row_numbers:
                row_peak_data = {}

                # Find and switch to the sheet for this row
                for sheet_name in self.parent.Data['Core levels'].keys():
                    num_match = re.search(r'(\d+)$', sheet_name)
                    if num_match:
                        sheet_row = int(num_match.group(1))
                    else:
                        sheet_row = 0

                    if sheet_row == row_num:
                        # Switch to this sheet
                        self.parent.sheet_combobox.SetValue(sheet_name)
                        from libraries.Sheet_Operations import on_sheet_selected
                        on_sheet_selected(self.parent, sheet_name)

                        # Extract data from results grid
                        grid = self.parent.results_grid
                        for row in range(grid.GetNumberRows()):
                            try:
                                peak_name = grid.GetCellValue(row, 0).strip()
                                value_str = grid.GetCellValue(row, data_column).strip()

                                if peak_name and value_str:
                                    value = float(value_str)
                                    row_peak_data[peak_name] = value
                            except (ValueError, IndexError):
                                continue
                        break

                # Fill in data for this row
                for peak_name in peak_names:
                    column_name = f"{peak_name} {column_suffix}"
                    if peak_name in row_peak_data:
                        profile_data[column_name].append(row_peak_data[peak_name])
                    else:
                        profile_data[column_name].append(0.0)

            # Create DataFrame with .2f formatting
            df = pd.DataFrame(profile_data)

            # Format all numeric columns to .2f
            for col in df.columns:
                if col != 'Number':
                    df[col] = df[col].apply(lambda x: float(f"{x:.2f}"))

            # Ensure Number column is first
            cols = ['Number'] + [col for col in df.columns if col != 'Number']
            df = df[cols]

            # SORT by Number column
            df = df.sort_values(by='Number').reset_index(drop=True)

            # Save to Excel
            file_path = self.parent.Data['FilePath']

            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=profile_sheet_name, index=False)

            # Update Data structure with profile metadata
            first_column = df.columns[1] if len(df.columns) > 1 else 'Number'

            if profile_type == "Concentration":
                y_axis_label = 'Atomic Concentration (%)'
            elif profile_type == "Area":
                y_axis_label = 'Area (CPS.eV)'
            elif profile_type == "Position":
                y_axis_label = 'Position (eV)'
            elif profile_type == "FWHM":
                y_axis_label = 'FWHM (eV)'
            elif profile_type == "Height":
                y_axis_label = 'Height (CPS)'
            elif profile_type == "L/G":
                y_axis_label = 'L/G (%)'
            elif profile_type == "Corrected Area":
                y_axis_label = 'Corrected Area (a.u.)'
            elif profile_type == "Weight (%)":
                y_axis_label = 'Weight (%)'

            self.parent.Data['Core levels'][profile_sheet_name] = {
                'Name': profile_sheet_name,
                'B.E.': df['Number'].tolist(),
                'Raw Data': df[first_column].tolist(),
                'Profile Data': df.to_dict('list'),
                'Y_Axis_Label': y_axis_label,
                'X_Axis_Label': 'Number',
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

            # Save to JSON
            json_file_path = os.path.splitext(file_path)[0] + '.json'
            from libraries.FileMenu.Save import convert_to_serializable_and_round
            json_data = convert_to_serializable_and_round(self.parent.Data)
            with open(json_file_path, 'w') as json_file:
                json.dump(json_data, json_file, indent=2)

            # Update sheet list
            if profile_sheet_name not in self.parent.sheet_combobox.GetStrings():
                self.parent.sheet_combobox.Append(profile_sheet_name)

            # Switch to the new profile sheet and plot it
            self.parent.sheet_combobox.SetValue(profile_sheet_name)
            from libraries.Sheet_Operations import on_sheet_selected
            on_sheet_selected(self.parent, profile_sheet_name)

            wx.MessageBox(
                f"Profile '{profile_sheet_name}' created successfully!\n"
                f"Data from {len(row_numbers)} rows with {len(peak_names)} peaks.",
                "Success",
                wx.OK | wx.ICON_INFORMATION
            )

        except Exception as e:
            wx.MessageBox(f"Error creating profile: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

    def on_close(self, event):
        """Handle window close"""
        if hasattr(self, 'editor_window') and self.editor_window:
            self.editor_window.Destroy()
        self.Destroy()


class ProfileEditWindow(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent,
                         title="Profile Data Editor",
                         size=(500, 400),
                         style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP)
        self.parent = parent
        self.current_profile = None
        self.Centre()
        self.init_ui()

    def init_ui(self):
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
            panel.SetBackgroundColour(wx.Colour(250, 220, 240))

        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Create toolbar
        self.toolbar = wx.ToolBar(panel, style=wx.TB_HORIZONTAL | wx.TB_FLAT | wx.TB_NODIVIDER)
        self.toolbar.SetToolBitmapSize(wx.Size(25, 25))

        # darker pink:
        if not detect_dark_mode():
            self.toolbar.SetBackgroundColour(wx.Colour(230, 200, 220))  # Darker pink

        # Get icon path
        current_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(current_dir, "Icons")

        # NEW EMPTY PROFILE TOOL - ADD THIS BEFORE SAVE TOOL
        new_profile_icon = os.path.join(icon_path, "NewSheet-3.png")
        if os.path.exists(new_profile_icon):
            new_profile_bmp = wx.Bitmap(new_profile_icon)
        else:
            new_profile_bmp = wx.ArtProvider.GetBitmap(wx.ART_NEW, wx.ART_TOOLBAR)
        new_profile_tool = self.toolbar.AddTool(wx.ID_ANY, "New Profile", new_profile_bmp, "Create Empty Profile")
        self.Bind(wx.EVT_TOOL, self.on_create_empty_profile, new_profile_tool)

        # Save tool
        save_icon = os.path.join(icon_path, "Save_Json-3.png")
        save_bmp = wx.Bitmap(save_icon)
        save_tool = self.toolbar.AddTool(wx.ID_ANY, "Save", save_bmp, "Save Changes")
        self.Bind(wx.EVT_TOOL, self.on_save, save_tool)

        self.toolbar.AddSeparator()

        # Add Row tool
        add_row_icon = os.path.join(icon_path, "Add_Row-3.png")
        add_row_bmp = wx.Bitmap(add_row_icon)
        add_row_tool = self.toolbar.AddTool(wx.ID_ANY, "", add_row_bmp, "Add Row")
        self.Bind(wx.EVT_TOOL, self.on_add_row, add_row_tool)

        # Remove Row tool
        rem_row_icon = os.path.join(icon_path, "Rem_Row-3.png")
        rem_row_bmp = wx.Bitmap(rem_row_icon)
        remove_row_tool = self.toolbar.AddTool(wx.ID_ANY, "Remove Row", rem_row_bmp, "Remove Row")
        self.Bind(wx.EVT_TOOL, self.on_remove_row, remove_row_tool)

        self.toolbar.AddSeparator()

        # Add Column tool
        add_col_icon = os.path.join(icon_path, "Add_column-3.png")
        add_col_bmp = wx.Bitmap(add_col_icon)
        add_col_tool = self.toolbar.AddTool(wx.ID_ANY, "Add Column", add_col_bmp, "Add Column")
        self.Bind(wx.EVT_TOOL, self.on_add_column, add_col_tool)

        # Remove Column tool
        rem_col_icon = os.path.join(icon_path, "Rem_column-3.png")
        rem_col_bmp = wx.Bitmap(rem_col_icon)
        remove_col_tool = self.toolbar.AddTool(wx.ID_ANY, "Remove Column", rem_col_bmp, "Remove Column")
        self.Bind(wx.EVT_TOOL, self.on_remove_column, remove_col_tool)

        self.toolbar.AddSeparator()

        # Edit Header tool
        edit_header_icon = os.path.join(icon_path, "rename-3.png")
        edit_header_bmp = wx.Bitmap(edit_header_icon)
        edit_header_tool = self.toolbar.AddTool(wx.ID_ANY, "Edit Header", edit_header_bmp, "Edit Header")
        self.Bind(wx.EVT_TOOL, self.on_edit_header, edit_header_tool)

        # X Label tool
        x_label_icon = os.path.join(icon_path, "XYplot_X-3.png")
        x_label_bmp = wx.Bitmap(x_label_icon)
        x_label_tool = self.toolbar.AddTool(wx.ID_ANY, "X Label", x_label_bmp, "Edit X Axis Label")
        self.Bind(wx.EVT_TOOL, self.on_edit_x_label, x_label_tool)

        # Y Label tool
        y_label_icon = os.path.join(icon_path, "XYplot_Y-3.png")
        y_label_bmp = wx.Bitmap(y_label_icon)
        y_label_tool = self.toolbar.AddTool(wx.ID_ANY, "Y Label", y_label_bmp, "Edit Y Axis Label")
        self.Bind(wx.EVT_TOOL, self.on_edit_y_label, y_label_tool)

        self.toolbar.AddStretchableSpace()



        self.toolbar.Realize()
        main_sizer.Add(self.toolbar, 0, wx.EXPAND)

        # Notebook with sheets for each profile
        self.notebook = fnb.FlatNotebook(panel, agwStyle=fnb.FNB_NO_X_BUTTON | fnb.FNB_NO_NAV_BUTTONS | fnb.FNB_BOTTOM)
        if not detect_dark_mode():
            self.notebook.SetBackgroundColour(wx.Colour(250, 220, 240))
            self.notebook.SetTabAreaColour(wx.Colour(250, 220, 240))
            self.notebook.SetActiveTabColour(wx.Colour(230, 250, 250))  # KherveFitting green
            self.notebook.SetNonActiveTabTextColour(wx.Colour(0, 0, 0))
            self.notebook.SetActiveTabTextColour(wx.Colour(0, 0, 0))

        self.grids = {}  # Dictionary to store grids for each profile
        self.x_labels = {}  # Store X labels for each profile
        self.y_labels = {}  # Store Y labels for each profile

        # Add undo/redo support for profile editor
        self.grid_history = []  # Stack for undo operations
        self.grid_redo_stack = []  # Stack for redo operations
        self.max_grid_history = 50  # Maximum number of undo states

        self.populate_notebook()

        main_sizer.Add(self.notebook, 1, wx.ALL | wx.EXPAND, 5)

        panel.SetSizer(main_sizer)

        # Bind notebook page change to update current profile
        self.notebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.on_page_changed)

        # self.Bind(wx.EVT_CLOSE, self.on_close)

        # Bind auto-save to all grids in the window
        for child in self.GetChildren():
            self._bind_grids_recursive(child)

    def _bind_grids_recursive(self, widget):
        """Recursively find and bind all wx.grid.Grid widgets"""
        if isinstance(widget, wx.grid.Grid):
            widget.Bind(wx.grid.EVT_GRID_CELL_CHANGED, self.on_cell_changed_auto_save)
            widget.Bind(wx.EVT_KEY_DOWN, self.on_grid_key_down)

        # Recursively check children
        if hasattr(widget, 'GetChildren'):
            for child in widget.GetChildren():
                self._bind_grids_recursive(child)

    def on_cell_changed_auto_save(self, event):
        """Auto-save using existing on_save method and update plot if active sheet"""
        try:
            # Save grid state for undo before the change
            self.save_grid_state()

            # Get the grid and cell that was changed
            grid = event.GetEventObject()
            row = event.GetRow()
            col = event.GetCol()

            # Format the cell value to .2f if it's a number
            value = grid.GetCellValue(row, col).strip()
            if value:
                try:
                    num_value = float(value)
                    formatted_value = f"{num_value:.2f}"
                    if value != formatted_value:
                        grid.SetCellValue(row, col, formatted_value)
                except ValueError:
                    pass  # Not a number, leave as-is

            # Call the existing save method (without showing message box)
            self.on_save(event)

            # Find the profile name
            profile_name = None
            if hasattr(self, 'current_profile') and self.current_profile:
                profile_name = self.current_profile
            elif hasattr(self, 'notebook'):
                page_index = self.notebook.GetSelection()
                profile_name = self.notebook.GetPageText(page_index)

            # If this is the active sheet, update the plot
            if profile_name:
                current_sheet = self.parent.sheet_combobox.GetValue()
                if current_sheet == profile_name:
                    self.parent.plot_manager.plot_data(self.parent)
                    self.parent.canvas.draw_idle()

        except Exception as e:
            print(f"Error in auto-save: {str(e)}")
            import traceback
            traceback.print_exc()
        finally:
            event.Skip()

    def populate_notebook(self):
        """Create a tab for each zzProfile in the data"""
        if not hasattr(self.parent, 'Data') or 'Core levels' not in self.parent.Data:
            return

        for profile_name in self.parent.Data['Core levels'].keys():
            if profile_name.startswith('zzProfile'):
                self.add_profile_tab(profile_name)

        # Select first tab if available
        if self.notebook.GetPageCount() > 0:
            self.notebook.SetSelection(0)
            self.current_profile = self.notebook.GetPageText(0)

    def add_profile_tab(self, profile_name):
        """Add a tab for a specific profile"""

        def detect_dark_mode():
            if 'wxMac' in wx.PlatformInfo:
                return wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW).GetLuminance() < 0.5
            elif 'wxMSW' in wx.PlatformInfo:
                return wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW).GetLuminance() < 0.5
            elif 'wxGTK' in wx.PlatformInfo:
                return wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW).GetLuminance() < 0.5
            return False


        panel = wx.Panel(self.notebook)
        panel.SetBackgroundColour(wx.Colour(240, 210, 230))

        sizer = wx.BoxSizer(wx.VERTICAL)

        # Create grid with 50 rows and 10 columns
        grid = wx.grid.Grid(panel)
        grid.CreateGrid(50, 10)
        grid.EnableEditing(True)
        grid.SetDefaultCellAlignment(wx.ALIGN_CENTER, wx.ALIGN_CENTER)
        grid.SetRowLabelSize(25)

        if not detect_dark_mode():
            grid.SetLabelBackgroundColour(wx.Colour(250, 220, 240))  # Pink for row/column headers

        # Set default column labels
        for col in range(10):
            grid.SetColLabelValue(col, f"Col{col}")
            grid.SetColSize(col, 80)

        # Bind grid events
        grid.Bind(wx.grid.EVT_GRID_LABEL_LEFT_DCLICK, self.on_header_double_click)

        # Bind auto-save event to this grid
        grid.Bind(wx.grid.EVT_GRID_CELL_CHANGED, self.on_cell_changed_auto_save)

        # Bind grid events
        grid.Bind(wx.grid.EVT_GRID_LABEL_LEFT_DCLICK, self.on_header_double_click)

        # Add this new binding for right-click menu
        grid.Bind(wx.grid.EVT_GRID_CELL_RIGHT_CLICK, self.on_grid_right_click)
        grid.Bind(wx.grid.EVT_GRID_LABEL_RIGHT_CLICK, self.on_header_right_click)

        # Bind auto-save event to this grid
        grid.Bind(wx.grid.EVT_GRID_CELL_CHANGED, self.on_cell_changed_auto_save)

        sizer.Add(grid, 1, wx.ALL | wx.EXPAND, 5)
        panel.SetSizer(sizer)

        self.notebook.AddPage(panel, profile_name)
        self.grids[profile_name] = grid

        # Load data for this profile
        self.load_profile_data(profile_name, grid)

    def save_grid_state(self):
        """Save current grid state for undo/redo"""
        grid = self.get_current_grid()
        if not grid:
            return

        # Get all grid data
        grid_data = []
        for row in range(grid.GetNumberRows()):
            row_data = []
            for col in range(grid.GetNumberCols()):
                row_data.append(grid.GetCellValue(row, col))
            grid_data.append(row_data)

        # Get column labels
        col_labels = []
        for col in range(grid.GetNumberCols()):
            col_labels.append(grid.GetColLabelValue(col))

        state = {
            'data': grid_data,
            'col_labels': col_labels,
            'num_rows': grid.GetNumberRows(),
            'num_cols': grid.GetNumberCols()
        }

        self.grid_history.append(state)
        if len(self.grid_history) > self.max_grid_history:
            self.grid_history.pop(0)
        self.grid_redo_stack.clear()

    def restore_grid_state(self, state):
        """Restore grid from saved state"""
        grid = self.get_current_grid()
        if not grid:
            return

        # Clear current grid
        if grid.GetNumberRows() > 0:
            grid.DeleteRows(0, grid.GetNumberRows())
        if grid.GetNumberCols() > 0:
            grid.DeleteCols(0, grid.GetNumberCols())

        # Restore structure
        grid.AppendRows(state['num_rows'])
        grid.AppendCols(state['num_cols'])

        # Restore column labels
        for col, label in enumerate(state['col_labels']):
            grid.SetColLabelValue(col, label)
            grid.SetColSize(col, 80)

        # Restore data
        for row, row_data in enumerate(state['data']):
            for col, value in enumerate(row_data):
                grid.SetCellValue(row, col, str(value))

        grid.ForceRefresh()

    def on_grid_undo(self, event):
        """Undo last grid operation"""
        if len(self.grid_history) > 1:
            current_state = self.grid_history.pop()
            self.grid_redo_stack.append(current_state)
            previous_state = self.grid_history[-1]
            self.restore_grid_state(previous_state)

    def on_grid_redo(self, event):
        """Redo last undone grid operation"""
        if self.grid_redo_stack:
            next_state = self.grid_redo_stack.pop()
            self.grid_history.append(next_state)
            self.restore_grid_state(next_state)

    def on_header_right_click(self, event):
        """Handle right-click on column header"""
        col = event.GetCol()
        if col < 0:
            return

        grid = event.GetEventObject()

        # Create context menu
        menu = wx.Menu()

        # Column operations
        insert_col_item = menu.Append(wx.ID_ANY, "Insert Column")
        remove_col_item = menu.Append(wx.ID_ANY, "Remove Column")
        menu.AppendSeparator()

        # Math operations
        multiply_item = menu.Append(wx.ID_ANY, "Multiply Column by...")
        divide_item = menu.Append(wx.ID_ANY, f"Divide Column by Another Column...")
        menu.AppendSeparator()

        # Edit operations
        undo_item = menu.Append(wx.ID_ANY, "Undo\tCtrl+Z")
        redo_item = menu.Append(wx.ID_ANY, "Redo\tCtrl+Y")

        # Enable/disable based on state
        undo_item.Enable(len(self.grid_history) > 1)
        redo_item.Enable(len(self.grid_redo_stack) > 0)

        # Bind events
        self.Bind(wx.EVT_MENU, lambda evt: self.on_insert_column_header(evt, col), insert_col_item)
        self.Bind(wx.EVT_MENU, lambda evt: self.on_remove_column_header(evt, col), remove_col_item)
        self.Bind(wx.EVT_MENU, lambda evt: self.on_multiply_column(evt, col), multiply_item)
        self.Bind(wx.EVT_MENU, lambda evt: self.on_divide_column(evt, col), divide_item)
        self.Bind(wx.EVT_MENU, self.on_grid_undo, undo_item)
        self.Bind(wx.EVT_MENU, self.on_grid_redo, redo_item)

        # Show menu
        grid.PopupMenu(menu)
        menu.Destroy()

    def on_insert_column_header(self, event, col_idx):
        """Insert column at specified position from header right-click"""
        self.save_grid_state()  # Save state for undo

        grid = self.get_current_grid()
        if not grid:
            return

        # Insert column
        grid.InsertCols(col_idx, 1)
        grid.SetColLabelValue(col_idx, f"NewCol{col_idx}")
        grid.SetColSize(col_idx, 80)

        # Fill with empty values
        for row in range(grid.GetNumberRows()):
            grid.SetCellValue(row, col_idx, "")

        grid.ForceRefresh()

    def on_remove_column_header(self, event, col_idx):
        """Remove column from header right-click"""
        grid = self.get_current_grid()
        if not grid:
            return

        if grid.GetNumberCols() <= 1:
            wx.MessageBox("Cannot remove last column", "Info", wx.OK | wx.ICON_INFORMATION)
            return

        col_name = grid.GetColLabelValue(col_idx)
        response = wx.MessageBox(f"Remove column '{col_name}'?", "Confirm", wx.YES_NO | wx.ICON_QUESTION)
        if response == wx.YES:
            self.save_grid_state()  # Save state for undo
            grid.DeleteCols(col_idx, 1)
            grid.ForceRefresh()

    def on_multiply_column(self, event, col_idx):
        """Multiply all values in a column by a factor"""
        grid = self.get_current_grid()
        if not grid:
            return

        col_name = grid.GetColLabelValue(col_idx)

        # Get multiplier from user
        dlg = wx.TextEntryDialog(self, f"Enter multiplier for column '{col_name}':", "Multiply Column", "1.0")
        if dlg.ShowModal() == wx.ID_OK:
            try:
                multiplier = float(dlg.GetValue())

                self.save_grid_state()  # Save state for undo

                # Multiply all numeric values in the column
                modified_count = 0
                for row in range(grid.GetNumberRows()):
                    cell_value = grid.GetCellValue(row, col_idx)
                    if cell_value and cell_value.strip():
                        try:
                            old_value = float(cell_value)
                            new_value = old_value * multiplier
                            grid.SetCellValue(row, col_idx, f"{new_value:.2f}")
                            modified_count += 1
                        except ValueError:
                            continue  # Skip non-numeric values

                grid.ForceRefresh()

                # Auto-save using the existing on_save method
                self.on_save(None)

                # If this is the active sheet, update the plot
                profile_name = None
                if hasattr(self, 'current_profile') and self.current_profile:
                    profile_name = self.current_profile
                elif hasattr(self, 'notebook'):
                    page_index = self.notebook.GetSelection()
                    profile_name = self.notebook.GetPageText(page_index)

                if profile_name:
                    current_sheet = self.parent.sheet_combobox.GetValue()
                    if current_sheet == profile_name:
                        self.parent.plot_manager.plot_data(self.parent)
                        self.parent.canvas.draw_idle()

                wx.MessageBox(f"Multiplied {modified_count} values by {multiplier}", "Success", wx.OK | wx.ICON_INFORMATION)

            except ValueError:
                wx.MessageBox("Invalid multiplier value", "Error", wx.OK | wx.ICON_ERROR)

        dlg.Destroy()

    def on_multiply_selected(self, event, col_idx):
        """Multiply only selected values in a column by a factor"""
        grid = self.get_current_grid()
        if not grid:
            return

        col_name = grid.GetColLabelValue(col_idx)

        # Get selected rows
        selected_rows = []
        if grid.GetSelectedRows():
            selected_rows = list(grid.GetSelectedRows())
        else:
            # Check if individual cells are selected in this column
            for row in range(grid.GetNumberRows()):
                if grid.IsInSelection(row, col_idx):
                    selected_rows.append(row)

        if not selected_rows:
            wx.MessageBox("No cells selected in this column", "Info", wx.OK | wx.ICON_INFORMATION)
            return

        # Get multiplier from user
        dlg = wx.TextEntryDialog(self, f"Enter multiplier for selected cells in column '{col_name}':", "Multiply Selected", "1.0")
        if dlg.ShowModal() == wx.ID_OK:
            try:
                multiplier = float(dlg.GetValue())

                self.save_grid_state()  # Save state for undo

                # Multiply selected numeric values in the column
                modified_count = 0
                for row in selected_rows:
                    cell_value = grid.GetCellValue(row, col_idx)
                    if cell_value and cell_value.strip():
                        try:
                            old_value = float(cell_value)
                            new_value = old_value * multiplier
                            grid.SetCellValue(row, col_idx, f"{new_value:.2f}")
                            modified_count += 1
                        except ValueError:
                            continue  # Skip non-numeric values

                grid.ForceRefresh()

                # Auto-save
                self.on_save(None)

                # Update plot if active sheet
                profile_name = self.current_profile if hasattr(self, 'current_profile') else None
                if profile_name and profile_name == self.parent.sheet_combobox.GetValue():
                    self.parent.plot_manager.plot_data(self.parent)
                    self.parent.canvas.draw_idle()

                wx.MessageBox(f"Multiplied {modified_count} selected values by {multiplier}", "Success", wx.OK | wx.ICON_INFORMATION)

            except ValueError:
                wx.MessageBox("Invalid multiplier value", "Error", wx.OK | wx.ICON_ERROR)

        dlg.Destroy()

    def on_make_array_selected(self, event, col_idx):
        """Create an evenly spaced array in selected cells"""
        grid = self.get_current_grid()
        if not grid:
            return

        col_name = grid.GetColLabelValue(col_idx)

        # Get selected rows
        selected_rows = []
        if grid.GetSelectedRows():
            selected_rows = sorted(list(grid.GetSelectedRows()))
        else:
            # Check if individual cells are selected in this column
            for row in range(grid.GetNumberRows()):
                if grid.IsInSelection(row, col_idx):
                    selected_rows.append(row)
            selected_rows.sort()

        if len(selected_rows) < 2:
            wx.MessageBox("Please select at least 2 cells to create an array", "Info", wx.OK | wx.ICON_INFORMATION)
            return

        # Create dialog to get min and max values
        dlg = wx.Dialog(self, title=f"Create Array in Column '{col_name}'", size=(300, 180))

        panel = wx.Panel(dlg)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Min value
        min_sizer = wx.BoxSizer(wx.HORIZONTAL)
        min_label = wx.StaticText(panel, label="Minimum value:")
        min_sizer.Add(min_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        min_ctrl = wx.TextCtrl(panel, value="0.00")
        min_sizer.Add(min_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        sizer.Add(min_sizer, 0, wx.ALL | wx.EXPAND, 5)

        # Max value
        max_sizer = wx.BoxSizer(wx.HORIZONTAL)
        max_label = wx.StaticText(panel, label="Maximum value:")
        max_sizer.Add(max_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        max_ctrl = wx.TextCtrl(panel, value="100.00")
        max_sizer.Add(max_ctrl, 1, wx.ALL | wx.EXPAND, 5)
        sizer.Add(max_sizer, 0, wx.ALL | wx.EXPAND, 5)

        # Info text
        info_text = wx.StaticText(panel, label=f"Creating array for {len(selected_rows)} selected cells")
        sizer.Add(info_text, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        # Buttons
        btn_sizer = wx.StdDialogButtonSizer()
        ok_btn = wx.Button(panel, wx.ID_OK)
        cancel_btn = wx.Button(panel, wx.ID_CANCEL)
        btn_sizer.AddButton(ok_btn)
        btn_sizer.AddButton(cancel_btn)
        btn_sizer.Realize()
        sizer.Add(btn_sizer, 0, wx.ALL | wx.EXPAND, 5)

        panel.SetSizer(sizer)

        if dlg.ShowModal() == wx.ID_OK:
            try:
                min_val = float(min_ctrl.GetValue())
                max_val = float(max_ctrl.GetValue())

                if min_val >= max_val:
                    wx.MessageBox("Maximum value must be greater than minimum value", "Error", wx.OK | wx.ICON_ERROR)
                    dlg.Destroy()
                    return

                self.save_grid_state()  # Save state for undo

                # Create array
                import numpy as np
                array_values = np.linspace(min_val, max_val, len(selected_rows))

                # Fill selected cells with array values
                for i, row in enumerate(selected_rows):
                    grid.SetCellValue(row, col_idx, f"{array_values[i]:.2f}")

                grid.ForceRefresh()

                # Auto-save
                self.on_save(None)

                # Update plot if active sheet
                profile_name = self.current_profile if hasattr(self, 'current_profile') else None
                if profile_name and profile_name == self.parent.sheet_combobox.GetValue():
                    self.parent.plot_manager.plot_data(self.parent)
                    self.parent.canvas.draw_idle()

                wx.MessageBox(f"Created array from {min_val:.2f} to {max_val:.2f} in {len(selected_rows)} cells",
                              "Success", wx.OK | wx.ICON_INFORMATION)

            except ValueError:
                wx.MessageBox("Invalid numeric values", "Error", wx.OK | wx.ICON_ERROR)

        dlg.Destroy()

    def on_divide_column(self, event, col_idx):
        """Divide this column by another column"""
        grid = self.get_current_grid()
        if not grid:
            return

        col_name = grid.GetColLabelValue(col_idx)

        # Get list of other columns (excluding the current one)
        other_columns = []
        other_col_indices = []
        for i in range(grid.GetNumberCols()):
            if i != col_idx:
                other_columns.append(grid.GetColLabelValue(i))
                other_col_indices.append(i)

        if not other_columns:
            wx.MessageBox("No other columns available for division", "Info", wx.OK | wx.ICON_INFORMATION)
            return

        # Create selection dialog
        dlg = wx.SingleChoiceDialog(
            self,
            f"Select column to divide '{col_name}' by:",
            "Divide Column",
            other_columns
        )

        if dlg.ShowModal() == wx.ID_OK:
            divisor_idx = other_col_indices[dlg.GetSelection()]
            divisor_name = grid.GetColLabelValue(divisor_idx)

            self.save_grid_state()  # Save state for undo

            # Perform division
            modified_count = 0
            zero_division_count = 0

            for row in range(grid.GetNumberRows()):
                dividend_value = grid.GetCellValue(row, col_idx).strip()
                divisor_value = grid.GetCellValue(row, divisor_idx).strip()

                if dividend_value and divisor_value:
                    try:
                        dividend = float(dividend_value)
                        divisor = float(divisor_value)

                        if divisor == 0:
                            zero_division_count += 1
                            # Leave cell unchanged or you could set to empty/NaN
                            continue

                        result = dividend / divisor
                        grid.SetCellValue(row, col_idx, f"{result:.2f}")
                        modified_count += 1

                    except ValueError:
                        continue  # Skip non-numeric values

            grid.ForceRefresh()

            # Auto-save
            self.on_save(None)

            # Update plot if active sheet
            profile_name = self.current_profile if hasattr(self, 'current_profile') else None
            if profile_name and profile_name == self.parent.sheet_combobox.GetValue():
                self.parent.plot_manager.plot_data(self.parent)
                self.parent.canvas.draw_idle()

            # Show result message
            msg = f"Divided {modified_count} cells by '{divisor_name}'"
            if zero_division_count > 0:
                msg += f"\n(Skipped {zero_division_count} cells due to division by zero)"

            wx.MessageBox(msg, "Success", wx.OK | wx.ICON_INFORMATION)

        dlg.Destroy()

    def load_profile_data(self, profile_name, grid):
        """Load profile data into grid"""
        if profile_name not in self.parent.Data['Core levels']:
            return

        profile_info = self.parent.Data['Core levels'][profile_name]

        # Store axis labels
        self.x_labels[profile_name] = profile_info.get('X_Axis_Label', 'Number')
        self.y_labels[profile_name] = profile_info.get('Y_Axis_Label', 'Value')

        # Load profile data
        if 'Profile Data' not in profile_info:
            return

        profile_data = profile_info['Profile Data']

        # Check if data is in new format (with x_values/y_values) or old format (DataFrame-style)
        if profile_data and isinstance(list(profile_data.values())[0], dict):
            # New format: each column has x_values and y_values
            all_columns = list(profile_data.keys())
            max_rows = 50

            # Determine how many columns we need
            num_cols = max(len(all_columns), 10)

            # Adjust grid size
            current_rows = grid.GetNumberRows()
            current_cols = grid.GetNumberCols()

            if max_rows > current_rows:
                grid.AppendRows(max_rows - current_rows)
            elif max_rows < current_rows:
                grid.DeleteRows(max_rows, current_rows - max_rows)

            if num_cols > current_cols:
                grid.AppendCols(num_cols - current_cols)
            elif num_cols < current_cols:
                grid.DeleteCols(num_cols, current_cols - num_cols)

            # Set column headers and clear all cells
            for col_idx in range(num_cols):
                if col_idx < len(all_columns):
                    col_name = all_columns[col_idx]
                    grid.SetColLabelValue(col_idx, col_name)

                    # Fill data for this column
                    col_info = profile_data[col_name]
                    x_values = col_info.get('x_values', [])
                    y_values = col_info.get('y_values', [])

                    # Clear all cells in this column first
                    for row_idx in range(max_rows):
                        grid.SetCellValue(row_idx, col_idx, "")

                    # Fill only the rows that have data
                    for i, (x_pos, y_val) in enumerate(zip(x_values, y_values)):
                        if x_pos < max_rows:
                            grid.SetCellValue(x_pos, col_idx, f"{float(y_val):.2f}")
                else:
                    grid.SetColLabelValue(col_idx, f"Col{col_idx}")
                    # Clear this column
                    for row_idx in range(max_rows):
                        grid.SetCellValue(row_idx, col_idx, "")

                grid.SetColSize(col_idx, 80)

        else:
            # Old format: try to convert to DataFrame (backwards compatibility)
            try:
                df = pd.DataFrame(profile_data)

                # Handle old format - convert to new format and save
                new_profile_data = {}
                for col_name in df.columns:
                    col_data = []
                    x_data = []
                    for row_idx, value in enumerate(df[col_name]):
                        if pd.notna(value) and str(value).strip():
                            try:
                                float_val = float(value)
                                col_data.append(float(f"{float_val:.2f}"))
                                x_data.append(row_idx)
                            except:
                                pass

                    if col_data:
                        new_profile_data[col_name] = {
                            'x_values': x_data,
                            'y_values': col_data
                        }

                # Update to new format
                profile_info['Profile Data'] = new_profile_data

                # Recursively call this method with new format
                self.load_profile_data(profile_name, grid)
                return

            except Exception as e:
                print(f"Error loading profile data: {e}")
                return

        grid.AutoSizeColumns()

    def on_page_changed(self, event):
        """Update current profile when notebook page changes"""
        selection = event.GetSelection()
        if selection != -1:
            self.current_profile = self.notebook.GetPageText(selection)
        event.Skip()

    def get_current_grid(self):
        """Get the grid for the currently selected profile"""
        if not self.current_profile or self.current_profile not in self.grids:
            return None
        return self.grids[self.current_profile]

    def on_add_row(self, event):
        """Add a new row"""
        grid = self.get_current_grid()
        if not grid:
            wx.MessageBox("No profile selected", "Error", wx.OK | wx.ICON_ERROR)
            return

        grid.AppendRows(1)
        new_row = grid.GetNumberRows() - 1

        # Fill with default values (0.00)
        for col in range(grid.GetNumberCols()):
            grid.SetCellValue(new_row, col, "")

    def on_remove_row(self, event):
        """Remove the last row"""
        grid = self.get_current_grid()
        if not grid:
            wx.MessageBox("No profile selected", "Error", wx.OK | wx.ICON_ERROR)
            return

        num_rows = grid.GetNumberRows()
        if num_rows <= 1:
            wx.MessageBox("Cannot remove last row", "Info", wx.OK | wx.ICON_INFORMATION)
            return

        response = wx.MessageBox(f"Remove the last row?", "Confirm", wx.YES_NO | wx.ICON_QUESTION)
        if response == wx.YES:
            grid.DeleteRows(num_rows - 1, 1)

    def on_add_column(self, event):
        """Add a new column"""
        grid = self.get_current_grid()
        if not grid:
            wx.MessageBox("No profile selected", "Error", wx.OK | wx.ICON_ERROR)
            return

        dlg = wx.TextEntryDialog(self, "Enter column name:", "Add Column")
        if dlg.ShowModal() == wx.ID_OK:
            col_name = dlg.GetValue().strip()
            if not col_name:
                wx.MessageBox("Column name cannot be empty", "Error", wx.OK | wx.ICON_ERROR)
                dlg.Destroy()
                return

            # Check if column name already exists
            existing_cols = [grid.GetColLabelValue(i) for i in range(grid.GetNumberCols())]
            if col_name in existing_cols:
                wx.MessageBox("Column name already exists", "Error", wx.OK | wx.ICON_ERROR)
                dlg.Destroy()
                return

            grid.AppendCols(1)
            new_col = grid.GetNumberCols() - 1
            grid.SetColLabelValue(new_col, col_name)
            grid.SetColSize(new_col, 80)

            # Fill with default values
            for row in range(grid.GetNumberRows()):
                grid.SetCellValue(row, new_col, "")

        dlg.Destroy()

    def on_remove_column(self, event):
        """Remove a column"""
        grid = self.get_current_grid()
        if not grid:
            wx.MessageBox("No profile selected", "Error", wx.OK | wx.ICON_ERROR)
            return

        num_cols = grid.GetNumberCols()
        if num_cols <= 1:
            wx.MessageBox("Cannot remove last column", "Info", wx.OK | wx.ICON_INFORMATION)
            return

        # Show list of columns
        col_names = [grid.GetColLabelValue(i) for i in range(num_cols)]
        dlg = wx.SingleChoiceDialog(self, "Select column to remove:", "Remove Column", col_names)

        if dlg.ShowModal() == wx.ID_OK:
            col_idx = dlg.GetSelection()
            col_name = col_names[col_idx]

            response = wx.MessageBox(f"Remove column '{col_name}'?", "Confirm", wx.YES_NO | wx.ICON_QUESTION)
            if response == wx.YES:
                grid.DeleteCols(col_idx, 1)

        dlg.Destroy()

    def on_edit_header(self, event):
        """Edit a column header"""
        grid = self.get_current_grid()
        if not grid:
            wx.MessageBox("No profile selected", "Error", wx.OK | wx.ICON_ERROR)
            return

        num_cols = grid.GetNumberCols()
        if num_cols == 0:
            wx.MessageBox("No columns available", "Info", wx.OK | wx.ICON_INFORMATION)
            return

        # Show list of columns
        col_names = [grid.GetColLabelValue(i) for i in range(num_cols)]
        dlg = wx.SingleChoiceDialog(self, "Select column to edit:", "Edit Header", col_names)

        if dlg.ShowModal() == wx.ID_OK:
            col_idx = dlg.GetSelection()
            old_name = col_names[col_idx]

            # Get new name
            name_dlg = wx.TextEntryDialog(self, f"Enter new name for '{old_name}':", "Edit Header", old_name)
            if name_dlg.ShowModal() == wx.ID_OK:
                new_name = name_dlg.GetValue().strip()
                if new_name:
                    # Check if name already exists
                    if new_name != old_name and new_name in col_names:
                        wx.MessageBox("Column name already exists", "Error", wx.OK | wx.ICON_ERROR)
                    else:
                        # Save grid state before change
                        self.save_grid_state()
                        grid.SetColLabelValue(col_idx, new_name)
                        # Trigger auto-save after header change
                        self.on_save(None)
            name_dlg.Destroy()

        dlg.Destroy()

    def on_header_double_click(self, event):
        """Handle double-click on column header"""
        if event.GetRow() == -1 and event.GetCol() >= 0:
            grid = self.get_current_grid()
            if not grid:
                event.Skip()
                return

            col_idx = event.GetCol()
            old_name = grid.GetColLabelValue(col_idx)

            dlg = wx.TextEntryDialog(self, f"Enter new name for '{old_name}':", "Edit Header", old_name)
            if dlg.ShowModal() == wx.ID_OK:
                new_name = dlg.GetValue().strip()
                if new_name:
                    # Check if name already exists
                    existing_cols = [grid.GetColLabelValue(i) for i in range(grid.GetNumberCols())]
                    if new_name != old_name and new_name in existing_cols:
                        wx.MessageBox("Column name already exists", "Error", wx.OK | wx.ICON_ERROR)
                    else:
                        # Save grid state before change
                        self.save_grid_state()
                        grid.SetColLabelValue(col_idx, new_name)
                        # Trigger auto-save after header change
                        self.on_save(None)
            dlg.Destroy()
            # Don't call event.Skip() to prevent further processing
            return

        event.Skip()

    def on_edit_x_label(self, event):
        """Edit X axis label"""
        if not self.current_profile:
            wx.MessageBox("No profile selected", "Error", wx.OK | wx.ICON_ERROR)
            return

        current_label = self.x_labels.get(self.current_profile, 'Number')
        dlg = wx.TextEntryDialog(self, "Enter X-Axis Label:", "X Label", current_label)

        if dlg.ShowModal() == wx.ID_OK:
            self.x_labels[self.current_profile] = dlg.GetValue()

        dlg.Destroy()

    def on_edit_y_label(self, event):
        """Edit Y axis label"""
        if not self.current_profile:
            wx.MessageBox("No profile selected", "Error", wx.OK | wx.ICON_ERROR)
            return

        current_label = self.y_labels.get(self.current_profile, 'Value')
        dlg = wx.TextEntryDialog(self, "Enter Y-Axis Label:", "Y Label", current_label)

        if dlg.ShowModal() == wx.ID_OK:
            self.y_labels[self.current_profile] = dlg.GetValue()

        dlg.Destroy()

    def on_save(self, event):
        """Save all changes to all profiles"""
        import numpy as np

        try:
            for profile_name, grid in self.grids.items():
                # Get grid data
                num_rows = grid.GetNumberRows()
                num_cols = grid.GetNumberCols()

                if num_cols == 0:
                    continue

                # First pass: identify which columns have data
                columns_with_data = []
                for col_idx in range(num_cols):
                    col_name = grid.GetColLabelValue(col_idx)

                    # Always include Number column
                    if col_name == 'Number':
                        columns_with_data.append(col_idx)
                        continue

                    # Skip columns with default names (Col5, Col6, etc.) that have no data
                    if col_name.startswith('Col') and col_name[3:].isdigit():
                        has_data = False
                        for row_idx in range(num_rows):
                            value = grid.GetCellValue(row_idx, col_idx).strip()
                            if value:
                                has_data = True
                                break
                        if not has_data:
                            continue

                    # Include all other columns
                    columns_with_data.append(col_idx)

                # Store data as separate arrays for each column
                profile_data = {}

                for col_idx in columns_with_data:
                    col_name = grid.GetColLabelValue(col_idx)

                    # Collect only non-empty values for this column
                    col_data = []
                    x_data = []  # Corresponding X values (row numbers)

                    for row_idx in range(num_rows):
                        value = grid.GetCellValue(row_idx, col_idx).strip()
                        if value:
                            try:
                                float_val = float(value)
                                col_data.append(float(f"{float_val:.2f}"))
                                x_data.append(row_idx)  # Use row index as X value
                            except:
                                pass  # Skip invalid values

                    # Store both the Y values and their corresponding X positions
                    if col_data:  # Only store if there's actual data
                        profile_data[col_name] = {
                            'x_values': x_data,
                            'y_values': col_data
                        }

                # Update in Data
                if profile_name not in self.parent.Data['Core levels']:
                    self.parent.Data['Core levels'][profile_name] = {}

                self.parent.Data['Core levels'][profile_name].update({
                    'X_Axis_Label': self.x_labels.get(profile_name, 'Number'),
                    'Y_Axis_Label': self.y_labels.get(profile_name, 'Value'),
                    'Profile Data': profile_data
                })

            # Refresh the plot if current sheet is a profile
            sheet_name = self.parent.sheet_combobox.GetValue()
            if sheet_name.startswith('zzProfile'):
                self.parent.plot_manager.plot_data(self.parent)

        except Exception as e:
            wx.MessageBox(f"Error saving profiles: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()


    def on_cancel(self, event):
        """Close without saving"""
        response = wx.MessageBox("Close without saving changes?", "Confirm",
                                 wx.YES_NO | wx.ICON_QUESTION)
        if response == wx.YES:
            self.Destroy()

    def on_grid_key_down(self, event):
        """Handle keyboard events for grid: Delete, Copy, Paste"""
        keycode = event.GetKeyCode()

        # Try to get the grid
        grid = event.GetEventObject()
        if not isinstance(grid, wx.grid.Grid):
            focused = wx.Window.FindFocus()
            if isinstance(focused, wx.grid.Grid):
                grid = focused
            else:
                event.Skip()
                return

        # Check for Ctrl+Z (Undo)
        if keycode == ord('Z') and event.ControlDown():
            self.on_grid_undo(event)
            return

        # Check for Ctrl+Y (Redo)
        if keycode == ord('Y') and event.ControlDown():
            self.on_grid_redo(event)
            return

        # DELETE key - clear current cell
        if keycode == wx.WXK_DELETE or keycode == wx.WXK_BACK:
            row = grid.GetGridCursorRow()
            col = grid.GetGridCursorCol()
            if row >= 0 and col >= 0:
                grid.SetCellValue(row, col, "")
                grid.ForceRefresh()

                # Manually trigger save and plot update
                try:
                    self.on_save(None)

                    # Update plot if active sheet
                    profile_name = None
                    if hasattr(self, 'current_profile') and self.current_profile:
                        profile_name = self.current_profile
                    elif hasattr(self, 'notebook'):
                        page_index = self.notebook.GetSelection()
                        profile_name = self.notebook.GetPageText(page_index)

                    if profile_name:
                        current_sheet = self.parent.sheet_combobox.GetValue()
                        if current_sheet == profile_name:
                            self.parent.plot_manager.plot_data(self.parent)
                            self.parent.canvas.draw_idle()
                except:
                    pass
            return

        # CTRL+C - Copy
        elif event.ControlDown() and keycode == 67:  # 'C'
            self.copy_grid_selection(grid)
            return

        # CTRL+V - Paste
        elif event.ControlDown() and keycode == 86:  # 'V'
            self.paste_grid_selection(grid)
            return

        event.Skip()

    def copy_grid_selection(self, grid):
        """Copy selected cells to clipboard"""
        # Just copy current cell for simplicity
        row = grid.GetGridCursorRow()
        col = grid.GetGridCursorCol()
        value = grid.GetCellValue(row, col)

        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(wx.TextDataObject(value))
            wx.TheClipboard.Close()

    def paste_grid_selection(self, grid):
        """Paste from clipboard to grid"""
        if not wx.TheClipboard.Open():
            return

        data = wx.TextDataObject()
        success = wx.TheClipboard.GetData(data)
        wx.TheClipboard.Close()

        if not success:
            return

        text = data.GetText()
        rows = text.split('\n')

        # Start at current cursor position
        start_row = grid.GetGridCursorRow()
        start_col = grid.GetGridCursorCol()

        for i, row_text in enumerate(rows):
            if not row_text.strip():
                continue
            cols = row_text.split('\t')
            for j, value in enumerate(cols):
                target_row = start_row + i
                target_col = start_col + j

                # Check bounds
                if target_row < grid.GetNumberRows() and target_col < grid.GetNumberCols():
                    grid.SetCellValue(target_row, target_col, value)

        grid.ForceRefresh()

    def on_create_empty_profile(self, event):
        """Create an empty zzProfile sheet"""
        import pandas as pd
        import os
        import json

        try:
            # Find next available profile name
            counter = 0
            while True:
                if counter == 0:
                    profile_sheet_name = "zzProfile"
                else:
                    profile_sheet_name = f"zzProfile{counter}"

                if profile_sheet_name not in self.parent.Data['Core levels']:
                    break
                counter += 1

            # Create empty profile data with just Number column (matching zzProfile template)
            profile_data = {'Number': [0.00]}

            # Create DataFrame
            df = pd.DataFrame(profile_data)

            # Save to Excel
            file_path = self.parent.Data['FilePath']

            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=profile_sheet_name, index=False)

            # Update Data structure with empty profile metadata (using proper zzProfile template)
            self.parent.Data['Core levels'][profile_sheet_name] = {
                'Name': profile_sheet_name,
                'B.E.': [0.00],
                'Raw Data': [0.00],
                'Profile Data': profile_data,
                'Y_Axis_Label': 'Value',
                'X_Axis_Label': 'Number',
                'Profile_Type': 'Custom',
                'Background': {
                    'Bkg Type': '',
                    'Bkg Low': '',
                    'Bkg High': '',
                    'Bkg Offset Low': '',
                    'Bkg Offset High': '',
                    'Bkg Y': [0.00]
                }
            }

            # Save to JSON
            json_file_path = os.path.splitext(file_path)[0] + '.json'
            from libraries.FileMenu.Save import convert_to_serializable_and_round
            json_data = convert_to_serializable_and_round(self.parent.Data)
            with open(json_file_path, 'w') as json_file:
                json.dump(json_data, json_file, indent=2)

            # Update sheet combobox
            if profile_sheet_name not in self.parent.sheet_combobox.GetStrings():
                self.parent.sheet_combobox.Append(profile_sheet_name)

            # Add the new profile tab to the editor window
            self.add_profile_tab(profile_sheet_name)

            # Select the newly created tab
            for i in range(self.notebook.GetPageCount()):
                if self.notebook.GetPageText(i) == profile_sheet_name:
                    self.notebook.SetSelection(i)
                    self.current_profile = profile_sheet_name
                    break

            wx.MessageBox(
                f"Empty profile '{profile_sheet_name}' created successfully!",
                "Success",
                wx.OK | wx.ICON_INFORMATION
            )

        except Exception as e:
            wx.MessageBox(f"Error creating empty profile: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

    def on_grid_right_click(self, event):
        """Handle right-click on grid to show context menu"""
        row = event.GetRow()
        col = event.GetCol()
        grid = event.GetEventObject()

        # Get column name
        col_name = grid.GetColLabelValue(col)

        # Create context menu
        menu = wx.Menu()

        # Add menu items
        insert_col_item = menu.Append(wx.ID_ANY, "Insert Column")
        menu.AppendSeparator()
        multiply_item = menu.Append(wx.ID_ANY, f"Multiply Column by...")
        multiply_selected_item = menu.Append(wx.ID_ANY, f"Multiply Selected by...")
        divide_item = menu.Append(wx.ID_ANY, f"Divide Column by Another Column...")
        make_array_item = menu.Append(wx.ID_ANY, f"Create Array in Selected...")
        menu.AppendSeparator()
        copy_item = menu.Append(wx.ID_ANY, "Copy")
        paste_item = menu.Append(wx.ID_ANY, "Paste")
        menu.AppendSeparator()
        undo_item = menu.Append(wx.ID_ANY, "Undo\tCtrl+Z")
        redo_item = menu.Append(wx.ID_ANY, "Redo\tCtrl+Y")

        # Enable/disable based on state
        undo_item.Enable(len(self.grid_history) > 1)
        redo_item.Enable(len(self.grid_redo_stack) > 0)

        # Bind events
        self.Bind(wx.EVT_MENU, lambda evt: self.on_insert_column(evt, col), insert_col_item)
        self.Bind(wx.EVT_MENU, lambda evt: self.on_multiply_column(evt, col), multiply_item)
        self.Bind(wx.EVT_MENU, lambda evt: self.on_copy_cell(evt, grid, row, col), copy_item)
        self.Bind(wx.EVT_MENU, lambda evt: self.on_paste_cell(evt, grid, row, col), paste_item)
        self.Bind(wx.EVT_MENU, lambda evt: self.on_multiply_selected(evt, col), multiply_selected_item)
        self.Bind(wx.EVT_MENU, lambda evt: self.on_make_array_selected(evt, col), make_array_item)
        self.Bind(wx.EVT_MENU, lambda evt: self.on_divide_column(evt, col), divide_item)
        self.Bind(wx.EVT_MENU, self.on_grid_undo, undo_item)
        self.Bind(wx.EVT_MENU, self.on_grid_redo, redo_item)

        # Show the menu
        grid.PopupMenu(menu)
        menu.Destroy()

    def on_insert_column(self, event, col_idx):
        """Insert a new column at the specified index and renumber default headers to the right"""
        grid = self.get_current_grid()
        if not grid:
            return

        insert_at = col_idx + 1
        grid.InsertCols(insert_at, 1)

        import re
        pattern = re.compile(r'^Col(\d+)$')
        last_col = grid.GetNumberCols() - 1
        for c in range(last_col, insert_at, -1):
            old_label = grid.GetColLabelValue(c)
            m = pattern.match(old_label)
            if m:
                grid.SetColLabelValue(c, f"Col{int(m.group(1)) + 1}")

        grid.SetColLabelValue(insert_at, f"Col{insert_at}")
        grid.SetColSize(insert_at, 40)

        grid.AutoSizeColumns()  # uncomment if you prefer autosizing after insert

        self.on_save(None)

    def on_copy_cell(self, event, grid, row, col):
        """Copy selected cell or cells to clipboard"""
        # Handle copying
        if grid.IsSelection():
            # Get the selection
            selected_cells = []
            for selection_block in grid.GetSelectedBlocks():
                top_left = (selection_block.GetTopRow(), selection_block.GetLeftCol())
                bottom_right = (selection_block.GetBottomRow(), selection_block.GetRightCol())

                rows = range(top_left[0], bottom_right[0] + 1)
                cols = range(top_left[1], bottom_right[1] + 1)

                # Format cells as tab-separated values with newlines between rows
                data = []
                for r in rows:
                    row_data = []
                    for c in cols:
                        value = grid.GetCellValue(r, c)
                        # Format as .2f if it's a number
                        try:
                            value = f"{float(value):.2f}"
                        except ValueError:
                            pass
                        row_data.append(value)
                    data.append("\t".join(row_data))

                clipboard_text = "\n".join(data)

                # Set clipboard data
                if wx.TheClipboard.Open():
                    wx.TheClipboard.SetData(wx.TextDataObject(clipboard_text))
                    wx.TheClipboard.Close()
        else:
            # Copy single cell
            value = grid.GetCellValue(row, col)
            # Format as .2f if it's a number
            try:
                value = f"{float(value):.2f}"
            except ValueError:
                pass

            if wx.TheClipboard.Open():
                wx.TheClipboard.SetData(wx.TextDataObject(value))
                wx.TheClipboard.Close()

    def on_paste_cell(self, event, grid, row, col):
        """Paste clipboard content to selected cell"""
        if wx.TheClipboard.Open():
            if wx.TheClipboard.IsSupported(wx.DataFormat(wx.DF_TEXT)):
                data = wx.TextDataObject()
                wx.TheClipboard.GetData(data)

                text = data.GetText()

                # Check if we have tab-separated data (multiple cells)
                if "\t" in text or "\n" in text:
                    rows = text.split("\n")
                    for r_idx, row_text in enumerate(rows):
                        if row + r_idx >= grid.GetNumberRows():
                            break

                        cols = row_text.split("\t")
                        for c_idx, cell_text in enumerate(cols):
                            if col + c_idx >= grid.GetNumberCols():
                                break

                            # Format as .2f if it's a number
                            try:
                                cell_text = f"{float(cell_text):.2f}"
                            except ValueError:
                                pass

                            grid.SetCellValue(row + r_idx, col + c_idx, cell_text)
                else:
                    # Single cell paste
                    # Format as .2f if it's a number
                    try:
                        text = f"{float(text):.2f}"
                    except ValueError:
                        pass

                    grid.SetCellValue(row, col, text)

            wx.TheClipboard.Close()

        # Auto-save after paste
        self.on_save(None)

    # def on_close(self, event):
    #     """Handle window close"""
    #     response = wx.MessageBox("Close without saving changes?", "Confirm",
    #                              wx.YES_NO | wx.ICON_QUESTION)
    #     if response == wx.YES:
    #         self.Destroy()
    #     else:
    #         event.Veto()