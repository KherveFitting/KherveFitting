# KherveFitting - XPS Data Analysis Software
# Copyright (C) 2024-2026 Gwilherm Kerherve <g.kerherve@ic.ac.uk>
#
# KherveFitting is dual-licensed:
#   - GNU GPL v3.0 (see LICENSE-GPL.txt) for open-source use
#   - Commercial Licence (see LICENSE-COMMERCIAL.txt) for proprietary use
# SPDX-License-Identifier: GPL-3.0-only OR LicenseRef-KherveFitting-Commercial

import wx
import re


class ExportResultsWindow(wx.Frame):
    def __init__(self, parent):
        super(ExportResultsWindow, self).__init__(
            parent,
            style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX | wx.MINIMIZE_BOX) | wx.STAY_ON_TOP
        )
        self.parent = parent
        self.SetTitle("Export to Results Grid")
        self.SetSize((273, 340))

        # Light red background color
        panel = wx.Panel(self)
        panel.SetBackgroundColour(wx.Colour(250, 200, 200))

        self.init_ui(panel)
        self.populate_core_levels_list()

        from libraries.ConfigFile import set_consistent_fonts
        set_consistent_fonts(self)

    def init_ui(self, panel):
        """Initialize the UI"""
        main_sizer = wx.GridBagSizer(0, 0)

        # Progress indicator
        progress_label = wx.StaticText(panel, label="Select Core Levels:")
        main_sizer.Add(progress_label, pos=(0, 0), span=(1, 2), flag=wx.ALL | wx.EXPAND, border=5)

        # Select All and Unselect All buttons
        select_all_button = wx.Button(panel, label="Select All")
        if 'wxMac' in wx.PlatformInfo:
            select_all_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            select_all_button.SetMinSize((125, 35))
        else:
            select_all_button.SetMinSize((125, 35))
        select_all_button.Bind(wx.EVT_BUTTON, self.on_select_all)
        main_sizer.Add(select_all_button, pos=(1, 0), flag=wx.ALL | wx.EXPAND, border=1)

        unselect_all_button = wx.Button(panel, label="Unselect All")
        if 'wxMac' in wx.PlatformInfo:
            unselect_all_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            unselect_all_button.SetMinSize((125, 35))
        else:
            unselect_all_button.SetMinSize((125, 35))
        unselect_all_button.Bind(wx.EVT_BUTTON, self.on_unselect_all)
        main_sizer.Add(unselect_all_button, pos=(1, 1), flag=wx.ALL | wx.EXPAND, border=1)

        # Select this core level and Select this row buttons
        select_core_level_button = wx.Button(panel, label="Select This Core Level")
        if 'wxMac' in wx.PlatformInfo:
            select_core_level_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            select_core_level_button.SetMinSize((125, 35))
        else:
            select_core_level_button.SetMinSize((125, 35))
        select_core_level_button.Bind(wx.EVT_BUTTON, self.on_select_this_core_level)
        main_sizer.Add(select_core_level_button, pos=(2, 0), flag=wx.ALL | wx.EXPAND, border=1)

        select_row_button = wx.Button(panel, label="Select This Row")
        if 'wxMac' in wx.PlatformInfo:
            select_row_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            select_row_button.SetMinSize((125, 35))
        else:
            select_row_button.SetMinSize((125, 35))
        select_row_button.Bind(wx.EVT_BUTTON, self.on_select_this_row)
        main_sizer.Add(select_row_button, pos=(2, 1), flag=wx.ALL | wx.EXPAND, border=1)

        # Core levels checklist box
        self.core_levels_checklist = wx.CheckListBox(panel)
        self.core_levels_checklist.SetMinSize((250, 150))
        self.core_levels_checklist.Bind(wx.EVT_CONTEXT_MENU, self.on_core_levels_context_menu)
        main_sizer.Add(self.core_levels_checklist, pos=(3, 0), span=(1, 2),
                       flag=wx.ALL | wx.EXPAND, border=3)

        # Export Selected button
        export_selected_button = wx.Button(panel, label="Export\nto Results Grid")
        if 'wxMac' in wx.PlatformInfo:
            export_selected_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            export_selected_button.SetMinSize((125, 35))
        else:
            export_selected_button.SetMinSize((125, 35))
        export_selected_button.Bind(wx.EVT_BUTTON, self.on_export_selected)
        main_sizer.Add(export_selected_button, pos=(4, 0), flag=wx.ALL | wx.EXPAND, border=1)

        # Clean and Export button
        clean_export_button = wx.Button(panel, label="Clean and Export\nto Results Grid")
        if 'wxMac' in wx.PlatformInfo:
            clean_export_button.SetMinSize((125, 30))
        elif 'wxGTK' in wx.PlatformInfo:
            clean_export_button.SetMinSize((125, 35))
        else:
            clean_export_button.SetMinSize((125, 35))
        clean_export_button.Bind(wx.EVT_BUTTON, self.on_clean_and_export)
        main_sizer.Add(clean_export_button, pos=(4, 1), flag=wx.ALL | wx.EXPAND, border=1)

        panel.SetSizer(main_sizer)
        self.Bind(wx.EVT_CLOSE, self.on_close)

    def populate_core_levels_list(self):
        """Populate the checklist with available core levels"""
        self.core_levels_checklist.Clear()

        if 'Core levels' in self.parent.Data:
            core_levels = list(self.parent.Data['Core levels'].keys())
            self.core_levels_checklist.AppendItems(core_levels)

    def on_select_all(self, event):
        """Select all core levels"""
        for i in range(self.core_levels_checklist.GetCount()):
            self.core_levels_checklist.Check(i, True)

    def on_unselect_all(self, event):
        """Unselect all core levels"""
        for i in range(self.core_levels_checklist.GetCount()):
            self.core_levels_checklist.Check(i, False)

    def on_select_this_core_level(self, event):
        """Select only the current active/plotted core level"""
        current_core_level = self.parent.sheet_combobox.GetValue()

        if not current_core_level or current_core_level not in self.parent.Data['Core levels']:
            wx.MessageBox("No valid core level currently selected", "No Selection", wx.OK | wx.ICON_INFORMATION)
            return

        # Uncheck all first
        for i in range(self.core_levels_checklist.GetCount()):
            self.core_levels_checklist.Check(i, False)

        # Check only the current core level
        for i in range(self.core_levels_checklist.GetCount()):
            if self.core_levels_checklist.GetString(i) == current_core_level:
                self.core_levels_checklist.Check(i, True)
                break

    def on_select_this_row_OLD(self, event):
        """Select all core levels on the same row as the current core level"""
        current_core_level = self.parent.sheet_combobox.GetValue()

        if not current_core_level or current_core_level not in self.parent.Data['Core levels']:
            wx.MessageBox("No valid core level currently selected", "No Selection", wx.OK | wx.ICON_INFORMATION)
            return

        # Uncheck all first
        for i in range(self.core_levels_checklist.GetCount()):
            self.core_levels_checklist.Check(i, False)

        # Extract row number from current core level
        current_row = self.extract_row_number(current_core_level)

        # Select all core levels with the same row number
        for i in range(self.core_levels_checklist.GetCount()):
            core_level = self.core_levels_checklist.GetString(i)
            if self.extract_row_number(core_level) == current_row:
                self.core_levels_checklist.Check(i, True)

    def on_select_this_row(self, event):
        """Select all core levels on the same row as the current core level (excluding zzProfiles)"""
        current_core_level = self.parent.sheet_combobox.GetValue()

        if not current_core_level or current_core_level not in self.parent.Data['Core levels']:
            wx.MessageBox("No valid core level currently selected", "No Selection", wx.OK | wx.ICON_INFORMATION)
            return

        # Uncheck all first
        for i in range(self.core_levels_checklist.GetCount()):
            self.core_levels_checklist.Check(i, False)

        # Extract row number from current core level
        current_row = self.extract_row_number(current_core_level)

        # Select all core levels with the same row number (excluding zzProfiles)
        for i in range(self.core_levels_checklist.GetCount()):
            core_level = self.core_levels_checklist.GetString(i)

            # Skip zzProfile entries
            if core_level.startswith('zzProfile'):
                continue

            if self.extract_row_number(core_level) == current_row:
                self.core_levels_checklist.Check(i, True)
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
            self.Bind(wx.EVT_MENU, lambda evt, cl=core_levels_of_type: self.select_core_levels(cl), select_item)
            self.Bind(wx.EVT_MENU, lambda evt, cl=core_levels_of_type: self.unselect_core_levels(cl), unselect_item)

            menu.AppendSeparator()

        # Show the menu
        self.core_levels_checklist.PopupMenu(menu)
        menu.Destroy()

    def extract_row_number(self, core_level_name):
        """Extract row number from core level name (e.g., Sr3d1 -> 1, Survey19 -> 19, Sr3d -> 0)"""
        cleaned_name = core_level_name.strip()

        # Extract trailing digits from the end of the name
        match = re.search(r'(\d+)$', cleaned_name)

        if match:
            return int(match.group(1))

        # No trailing digits means row 0
        return 0
    def extract_column_type(self, core_level_name):
        """Extract the core level type from the name (e.g., C1s from C1s_1)"""
        # Remove leading/trailing whitespace
        cleaned_name = core_level_name.strip()

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

    def select_core_levels(self, core_levels_list):
        """Select specific core levels in the checklist"""
        for i in range(self.core_levels_checklist.GetCount()):
            core_level = self.core_levels_checklist.GetString(i)
            if core_level in core_levels_list:
                self.core_levels_checklist.Check(i, True)

    def unselect_core_levels(self, core_levels_list):
        """Unselect specific core levels in the checklist"""
        for i in range(self.core_levels_checklist.GetCount()):
            core_level = self.core_levels_checklist.GetString(i)
            if core_level in core_levels_list:
                self.core_levels_checklist.Check(i, False)

    def on_export_selected(self, event):
        """Export selected core levels to results grid"""
        selected_sheets = []
        for i in range(self.core_levels_checklist.GetCount()):
            if self.core_levels_checklist.IsChecked(i):
                selected_sheets.append(self.core_levels_checklist.GetString(i))

        if not selected_sheets:
            wx.MessageBox("No core levels selected", "Warning", wx.OK | wx.ICON_WARNING)
            return

        self.export_multiple_sheets(selected_sheets, clean_first=False)

    def on_clean_and_export(self, event):
        """Clean results grid then export selected core levels"""
        selected_sheets = []
        for i in range(self.core_levels_checklist.GetCount()):
            if self.core_levels_checklist.IsChecked(i):
                selected_sheets.append(self.core_levels_checklist.GetString(i))

        if not selected_sheets:
            wx.MessageBox("No core levels selected", "Warning", wx.OK | wx.ICON_WARNING)
            return

        self.export_multiple_sheets(selected_sheets, clean_first=True)

    def export_multiple_sheets(self, sheet_names, clean_first=False):
        """Export multiple sheets to results grid"""
        from libraries.FileMenu.Export import export_results

        # Clean results grid if requested
        if clean_first:
            # Only delete if there are rows to delete
            if self.parent.results_grid.GetNumberRows() > 0:
                self.parent.results_grid.DeleteRows(0, self.parent.results_grid.GetNumberRows())

            # Also clear the Data structure for all results tables
            keys_to_clear = [key for key in self.parent.Data.keys() if key.startswith('Results Table')]
            for key in keys_to_clear:
                if 'Peak' in self.parent.Data[key]:
                    self.parent.Data[key]['Peak'].clear()

        # Store current sheet
        original_sheet = self.parent.sheet_combobox.GetValue()

        # Export each selected sheet
        for sheet_name in sheet_names:
            # Switch to the sheet
            self.parent.sheet_combobox.SetValue(sheet_name)
            from libraries.Sheet_Operations import on_sheet_selected
            on_sheet_selected(self.parent, None)

            # Export the current sheet
            export_results(self.parent)

        # Restore original sheet
        self.parent.sheet_combobox.SetValue(original_sheet)
        on_sheet_selected(self.parent, None)

        # wx.MessageBox(f"Exported {len(sheet_names)} core levels to Results Grid",
        #               "Export Complete", wx.OK | wx.ICON_INFORMATION)

    def on_close(self, event):
        """Handle window close"""
        self.Destroy()

    def OnClose(self, event):
        self.Destroy()