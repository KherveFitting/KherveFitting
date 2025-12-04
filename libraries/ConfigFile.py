
# CONFIG FILE FOR XPS DATA ------------------------------------
# -------------------------------------------------------------
import pandas as pd
import wx
import matplotlib.pyplot as plt


def Init_Measurement_Data(window):
    Data = {
        'FilePath': '',
        'Number of Core levels': 0,
        'Core levels': {},
    }

    # Initialize 10 Results Tables
    for i in range(10):
        Data[f'Results Table{i}'] = {
            'Peak': {}
        }

    return Data


def add_core_level_Data(Data, window, file_path, sheet_name):
    """
    Add core level data from Excel sheet to Data structure, including experimental info
    """
    import openpyxl
    import pandas as pd
    import numpy as np

    try:
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

        # ========== Handle EDX~Plot sheets ==========
        if sheet_name == 'EDX~Plot' or sheet_name.startswith('EDX~Plot'):
            be_values = []
            raw_data = []

            for index, row in df.iterrows():
                if index == 0:
                    continue

                be_val = row.iloc[0]
                raw_val = row.iloc[1]

                if pd.isna(be_val) or pd.isna(raw_val):
                    continue

                try:
                    be_values.append(float(be_val))
                    raw_data.append(float(raw_val))
                except (ValueError, TypeError):
                    continue

            core_level = {
                'Name': sheet_name,
                'B.E.': be_values,
                'Raw Data': raw_data,
                '_EDX_display_max': 20,
                '_EDX_type': 'plot',
                'Background': {}
            }

            Data['Core levels'][sheet_name] = core_level
            return Data

        # ========== Handle EDX~Map sheets ==========
        if sheet_name == 'EDX~Map':
            # For EDX~Map, just store metadata - actual map will be loaded from HDF5
            energy_range = None

            for index, row in df.iterrows():
                if index == 0:
                    header_text = str(row.iloc[0])
                    if 'Range:' in header_text:
                        energy_range = header_text.split('Range:')[1].strip()
                    break  # Only need the header

            # Determine HDF5 file path
            hdf5_path = file_path.replace('_EDX.xlsx', '_EDX.hdf5')
            if not os.path.exists(hdf5_path):
                hdf5_path = file_path.replace('_EDX.xlsx', '.hdf5')
            if not os.path.exists(hdf5_path):
                hdf5_path = file_path.replace('_EDX.xlsx', '.h5')

            core_level = {
                'Name': sheet_name,
                'Energy_Range': energy_range if energy_range else 'N/A',
                '_EDX_type': 'map',
                '_HDF5_Path': hdf5_path if os.path.exists(hdf5_path) else None
            }

            Data['Core levels'][sheet_name] = core_level
            return Data

        # Extract B.E. and Raw Data columns
        be_values = []
        raw_data = []
        corrected_data = []
        transmission = []

        # Skip the header row and extract data
        for index, row in df.iterrows():
            if index == 0:  # Skip header
                continue

            be_val = row.iloc[0]  # Column A (B.E.)
            raw_val = row.iloc[1]  # Column B (Raw Data)

            # Handle missing or invalid data
            if pd.isna(be_val) or pd.isna(raw_val):
                continue

            try:
                be_values.append(float(be_val))
                raw_data.append(float(raw_val))

                # Check if corrected data column exists
                if len(row) > 2 and not pd.isna(row.iloc[2]):
                    corrected_data.append(float(row.iloc[2]))
                else:
                    corrected_data.append(float(raw_val))

                # Check if transmission column exists
                if len(row) > 3 and not pd.isna(row.iloc[3]):
                    transmission.append(float(row.iloc[3]))
                else:
                    transmission.append(1.0)

            except (ValueError, TypeError):
                continue

        # Extract experimental description data from Excel file
        experimental_info = {}

        # Load workbook to access experimental description columns
        wb = openpyxl.load_workbook(file_path)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            # Search for experimental description column (typically around column 45-50)
            exp_col = None
            for col in range(40, min(61, ws.max_column + 1)):
                cell_value = ws.cell(row=1, column=col).value
                if cell_value and "Experimental Description" in str(cell_value):
                    exp_col = col
                    break

            if exp_col:
                # Read experimental description data
                for row in range(2, ws.max_row + 1):
                    param_cell = ws.cell(row=row, column=exp_col)
                    value_cell = ws.cell(row=row, column=exp_col + 1)

                    if param_cell.value is not None and str(param_cell.value).strip():
                        param_name = str(param_cell.value).strip()
                        param_value = str(value_cell.value).strip() if value_cell.value is not None else ""
                        experimental_info[param_name] = param_value

        # Create the core level data structure with .2f formatting
        core_level_data = {
            'B.E.': [float(f"{val:.2f}") for val in be_values],
            'Raw Data': [float(f"{val:.2f}") for val in raw_data],
            'Corrected Data': [float(f"{val:.2f}") for val in corrected_data],
            'Transmission': [float(f"{val:.2f}") for val in transmission],
            'Name': sheet_name
        }

        # Add experimental info to the core level data if found
        if experimental_info:
            core_level_data['ExperimentalInfo'] = experimental_info

        # Initialize background structure
        if be_values:
            core_level_data['Background'] = {
                'Bkg Y': core_level_data['Raw Data'],
                'Bkg Type': '',
                'Bkg Low': float(f"{min(be_values):.2f}"),
                'Bkg High': float(f"{max(be_values):.2f}"),
                'Bkg Offset Low': 0,
                'Bkg Offset High': 0
            }

        # Add to Data structure
        if 'Core levels' not in Data:
            Data['Core levels'] = {}

        Data['Core levels'][sheet_name] = core_level_data
        Data['Number of Core levels'] = len(Data['Core levels'])

        return Data

    except Exception as e:
        print(f"Error adding core level data for {sheet_name}: {e}")
        return Data


def add_peak_to_core_level_Data(data, core_name, peak_data):
    if core_name in data['Core levels']:
        fitting = data['Core levels'][core_name]['Fitting']
        fitting.update(peak_data)
    else:
        print(f"Core level {core_name} does not exist.")


def set_consistent_fonts(window):
    if 'wxMac' in wx.PlatformInfo:
        default_font = 'Helvetica'
        # default_font = 'Calibri'
        # default_font = 'Arial'

        STANDARD_FONT_SIZE = 11
        # STANDARD_FONT_SIZE = 5
        HEADER_FONT_SIZE = 12
        # HEADER_FONT_SIZE = 5
        LABEL_FONT_SIZE = 11
        # LABEL_FONT_SIZE = 5
    elif 'wxGTK' in wx.PlatformInfo:  # wxGTK is the Linux version of wxWidgets
        default_font = 'DejaVu Sans'  # A common Linux font
        STANDARD_FONT_SIZE = 9
        HEADER_FONT_SIZE = 10
        LABEL_FONT_SIZE = 9
    else:
        # default_font = 'Arial'
        # default_font = 'Helvetica'
        default_font = 'Calibri'

        STANDARD_FONT_SIZE = 9
        # STANDARD_FONT_SIZE = 5
        HEADER_FONT_SIZE = 10
        # HEADER_FONT_SIZE = 5
        LABEL_FONT_SIZE = 9
        # LABEL_FONT_SIZE = 5

    window.SetFont(wx.Font(STANDARD_FONT_SIZE, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL,
                           faceName=default_font))

    # Handle all wx widgets including buttons
    def set_font_recursive(widget):
        widget.SetFont(wx.Font(STANDARD_FONT_SIZE, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL,
                               faceName=default_font))
        if isinstance(widget, wx.Button):
            widget.SetFont(wx.Font(STANDARD_FONT_SIZE, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL,
                                   faceName=default_font))
        for child in widget.GetChildren():
            set_font_recursive(child)

    set_font_recursive(window)

    # Handle matplotlib
    if hasattr(window, 'ax'):
        plt.rcParams['font.family'] = default_font
        window.ax.tick_params(axis='both', labelsize=STANDARD_FONT_SIZE)
        window.ax.set_xlabel(window.ax.get_xlabel(), fontsize=STANDARD_FONT_SIZE, fontname=default_font)
        window.ax.set_ylabel(window.ax.get_ylabel(), fontsize=STANDARD_FONT_SIZE, fontname=default_font)
        if window.ax.get_title():
            window.ax.set_title(window.ax.get_title(), fontsize=HEADER_FONT_SIZE, fontname=default_font)

    # Handle grids
    if hasattr(window, 'peak_params_grid'):
        window.peak_params_grid.SetDefaultCellFont(
            wx.Font(STANDARD_FONT_SIZE, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL,
                    faceName=default_font))
        window.peak_params_grid.SetLabelFont(
            wx.Font(LABEL_FONT_SIZE, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL,
                    faceName=default_font))

    if hasattr(window, 'results_grid'):
        window.results_grid.SetDefaultCellFont(
            wx.Font(STANDARD_FONT_SIZE, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL,
                    faceName=default_font))
        window.results_grid.SetLabelFont(
            wx.Font(LABEL_FONT_SIZE, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL,
                    faceName=default_font))

    if hasattr(window, 'fitting_window') and window.fitting_window:
        set_consistent_fonts(window.fitting_window)

    if hasattr(window, 'background_window') and window.background_window:
        set_consistent_fonts(window.background_window)

    if hasattr(window, 'canvas'):
        window.canvas.draw_idle()