# KherveFitting - XPS Data Analysis Software
# Copyright (C) 2024-2026 Gwilherm Kerherve <g.kerherve@ic.ac.uk>
#
# KherveFitting is dual-licensed:
#   - GNU GPL v3.0 (see LICENSE-GPL.txt) for open-source use
#   - Commercial Licence (see LICENSE-COMMERCIAL.txt) for proprietary use
# SPDX-License-Identifier: GPL-3.0-only OR LicenseRef-KherveFitting-Commercial

"""
Edit Data Window - allows editing of raw data from Excel
"""
import wx
import wx.grid
import pandas as pd
import numpy as np
import os


class EditDataWindow(wx.Frame):
    """Window for editing raw spectroscopy data"""

    def __init__(self, parent):
        super().__init__(parent, title="Edit Data", size=(400, 400),
                         style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP)
        self.parent = parent
        self.current_sheet = parent.sheet_combobox.GetValue()

        self.Centre()
        self.init_ui()
        self.load_data()

    def init_ui(self):
        """Initialize the user interface"""
        panel = wx.Panel(self)

        def detect_dark_mode():
            if 'wxMac' in wx.PlatformInfo:
                return wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW).GetLuminance() < 0.5
            elif 'wxMSW' in wx.PlatformInfo:
                return wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW).GetLuminance() < 0.5
            elif 'wxGTK' in wx.PlatformInfo:
                return wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW).GetLuminance() < 0.5
            return False

        # if not detect_dark_mode():
        #     panel.SetBackgroundColour(wx.Colour(250, 220, 240))

        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Create toolbar
        self.toolbar = wx.ToolBar(panel, style=wx.TB_HORIZONTAL | wx.TB_FLAT | wx.TB_NODIVIDER)
        self.toolbar.SetToolBitmapSize(wx.Size(25, 25))

        # if not detect_dark_mode():
        #     self.toolbar.SetBackgroundColour(wx.Colour(230, 200, 220))

        # Get icon path
        current_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Icons")

        # Save tool
        save_icon = os.path.join(icon_path, "Save_Json-3.png")
        if os.path.exists(save_icon):
            save_bmp = wx.Bitmap(save_icon)
        else:
            save_bmp = wx.ArtProvider.GetBitmap(wx.ART_FILE_SAVE, wx.ART_TOOLBAR)
        save_tool = self.toolbar.AddTool(wx.ID_ANY, "Save", save_bmp, "Save Changes")
        self.Bind(wx.EVT_TOOL, self.on_save, save_tool)

        self.toolbar.Realize()
        main_sizer.Add(self.toolbar, 0, wx.EXPAND)

        # # Add sheet name label
        # sheet_label = wx.StaticText(panel, label=f"Editing: {self.current_sheet}")
        # font = sheet_label.GetFont()
        # font.PointSize += 2
        # font = font.Bold()
        # sheet_label.SetFont(font)
        # main_sizer.Add(sheet_label, 0, wx.ALL | wx.CENTER, 10)

        # Create grid
        self.grid = wx.grid.Grid(panel)
        self.grid.CreateGrid(0, 4)  # Start with 0 rows, 4 columns

        self.grid.SetRowLabelSize(30)

        # Set column labels
        self.grid.SetColLabelValue(0, "Col A")
        self.grid.SetColLabelValue(1, "Col B")
        self.grid.SetColLabelValue(2, "Col C")
        self.grid.SetColLabelValue(3, "Col D")

        # # Set column widths
        # for col in range(4):
        #     self.grid.SetColSize(col, 150)

        main_sizer.Add(self.grid, 1, wx.ALL | wx.EXPAND, 5)

        panel.SetSizer(main_sizer)

    def load_data(self):
        """Load data from Excel file into the grid"""
        file_path = self.parent.Data.get('FilePath', '')
        if not file_path or not os.path.exists(file_path):
            wx.MessageBox("No file path found", "Error", wx.OK | wx.ICON_ERROR)
            return

        try:
            # Read Excel file
            df = pd.read_excel(file_path, sheet_name=self.current_sheet)

            # Get number of rows
            num_rows = len(df)

            # Clear existing rows
            if self.grid.GetNumberRows() > 0:
                self.grid.DeleteRows(0, self.grid.GetNumberRows())

            # Add rows
            self.grid.AppendRows(num_rows)

            # Fill grid with data (4 columns: A, B, C, D)
            for row_idx in range(num_rows):
                # Column A (X / B.E.)
                if row_idx < len(df.iloc[:, 0]):
                    val = df.iloc[row_idx, 0]
                    if pd.notna(val):
                        self.grid.SetCellValue(row_idx, 0, f"{float(val):.2f}")

                # Column B (Y / Raw Data)
                if df.shape[1] > 1 and row_idx < len(df.iloc[:, 1]):
                    val = df.iloc[row_idx, 1]
                    if pd.notna(val):
                        self.grid.SetCellValue(row_idx, 1, f"{float(val):.2f}")

                # Column C
                if df.shape[1] > 2 and row_idx < len(df.iloc[:, 2]):
                    val = df.iloc[row_idx, 2]
                    if pd.notna(val):
                        self.grid.SetCellValue(row_idx, 2, f"{float(val):.2f}")

                # Column D
                if df.shape[1] > 3 and row_idx < len(df.iloc[:, 3]):
                    val = df.iloc[row_idx, 3]
                    if pd.notna(val):
                        self.grid.SetCellValue(row_idx, 3, f"{float(val):.2f}")

            self.grid.ForceRefresh()

        except Exception as e:
            wx.MessageBox(f"Error loading data: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

    def on_save(self, event):
        """Save the edited data back to Excel and window.Data"""
        file_path = self.parent.Data.get('FilePath', '')
        if not file_path or not os.path.exists(file_path):
            wx.MessageBox("No file path found", "Error", wx.OK | wx.ICON_ERROR)
            return

        try:
            # Collect data from grid
            num_rows = self.grid.GetNumberRows()

            col_a_data = []
            col_b_data = []
            col_c_data = []
            col_d_data = []

            for row_idx in range(num_rows):
                # Column A (X / B.E.)
                val_a = self.grid.GetCellValue(row_idx, 0)
                if val_a:
                    try:
                        col_a_data.append(float(val_a))
                    except ValueError:
                        col_a_data.append(0.0)
                else:
                    col_a_data.append(0.0)

                # Column B (Y / Raw Data)
                val_b = self.grid.GetCellValue(row_idx, 1)
                if val_b:
                    try:
                        col_b_data.append(float(val_b))
                    except ValueError:
                        col_b_data.append(0.0)
                else:
                    col_b_data.append(0.0)

                # Column C
                val_c = self.grid.GetCellValue(row_idx, 2)
                if val_c:
                    try:
                        col_c_data.append(float(val_c))
                    except ValueError:
                        col_c_data.append(0.0)
                else:
                    col_c_data.append(0.0)

                # Column D
                val_d = self.grid.GetCellValue(row_idx, 3)
                if val_d:
                    try:
                        col_d_data.append(float(val_d))
                    except ValueError:
                        col_d_data.append(1.0)
                else:
                    col_d_data.append(1.0)

            # Read original Excel to get column names
            df_original = pd.read_excel(file_path, sheet_name=self.current_sheet)
            column_names = df_original.columns.tolist()

            # Ensure we have at least 4 column names
            while len(column_names) < 4:
                column_names.append(f"Column{len(column_names) + 1}")

            # Create DataFrame with formatted values (.2f)
            df_new = pd.DataFrame({
                column_names[0]: [float(f"{val:.2f}") for val in col_a_data],
                column_names[1]: [float(f"{val:.2f}") for val in col_b_data],
                column_names[2]: [float(f"{val:.2f}") for val in col_c_data],
                column_names[3]: [float(f"{val:.2f}") for val in col_d_data]
            })

            # Save to Excel
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a',
                                if_sheet_exists='replace') as writer:
                df_new.to_excel(writer, sheet_name=self.current_sheet, index=False)

            # Update window.Data
            if self.current_sheet in self.parent.Data['Core levels']:
                # Update B.E. and Raw Data
                self.parent.Data['Core levels'][self.current_sheet]['B.E.'] = col_a_data
                self.parent.Data['Core levels'][self.current_sheet]['Raw Data'] = col_b_data

                # Store column C and D data as metadata for preservation
                self.parent.Data['Core levels'][self.current_sheet]['column_C_data'] = col_c_data
                self.parent.Data['Core levels'][self.current_sheet]['column_D_data'] = col_d_data

                # Update x_values and y_values if this is the current sheet
                if self.parent.sheet_combobox.GetValue() == self.current_sheet:
                    self.parent.x_values = np.array(col_a_data)
                    self.parent.y_values = np.array(col_b_data)

                    # Replot the data
                    self.parent.plot_manager.plot_data(self.parent)

            # Save JSON
            import json
            from libraries.FileMenu.Save import convert_to_serializable_and_round

            json_file_path = os.path.splitext(file_path)[0] + '.json'
            json_data = convert_to_serializable_and_round(self.parent.Data)
            with open(json_file_path, 'w') as json_file:
                json.dump(json_data, json_file, indent=2)

            wx.MessageBox("Data saved successfully!", "Success", wx.OK | wx.ICON_INFORMATION)

        except Exception as e:
            wx.MessageBox(f"Error saving data: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()