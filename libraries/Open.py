import wx
import os
import json
import openpyxl
import xlrd
import wx
import re
import sys
import pandas as pd
import struct
from pathlib import Path
import shutil
from vamas import Vamas
from openpyxl import Workbook
import numpy as np
from scipy.interpolate import interp1d
import pandas as pd
from openpyxl.styles import Alignment
from yadg.extractors.phi.spe import extract  # NOTE THIS LIBRARY HAS BEEN TRANSFORMED

from libraries.ConfigFile import Init_Measurement_Data, add_core_level_Data
from libraries.Save import update_undo_redo_state, save_state
from libraries.Sheet_Operations import on_sheet_selected
from libraries.Grid_Operations import populate_results_grid


class ExcelDropTarget(wx.FileDropTarget):
    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self.window = window

    def OnDropFiles(self, x, y, filenames):
        from libraries.Open import open_xlsx_file, open_vamas_file
        for file in filenames:
            if not any(file.lower().endswith(ext) for ext in ['.xlsx', '.xls', '.vms', '.kal', '.avg', '.spe']):
                wx.MessageBox(f"Only .xlsx/.xls (Khervefitting or Avantage), .vms (Vamas), "
                              f".kal (Kratos), .avg (Thermo) and .spe (Phi) files can be dropped.", "Invalid File Type",
                              wx.OK | wx.ICON_ERROR)
                return False
            if file.lower().endswith('.xlsx'):
                try:
                    wb = openpyxl.load_workbook(file)
                    if "Titles" in wb.sheetnames:
                        wx.CallAfter(import_avantage_file_direct, self.window, file)
                    else:
                        wx.CallAfter(open_xlsx_file, self.window, file)
                    return True
                except Exception:
                    return False
            elif file.lower().endswith('.xls'):
                try:
                    wb = xlrd.open_workbook(file)
                    if "Titles" in wb.sheet_names():
                        wx.CallAfter(import_avantage_file_direct_xls, self.window, file)
                    else:
                        wx.CallAfter(open_xlsx_file, self.window, file)
                    return True
                except Exception:
                    print("Error opening Excel file:", sys.exc_info())
                    return False
            elif file.lower().endswith('.vms'):
                wx.CallAfter(open_vamas_file, self.window, file)
                return True
            elif file.lower().endswith('.kal'):
                wx.CallAfter(open_kal_file, self.window, file)
                return True
            elif file.lower().endswith('.avg'):
                wx.CallAfter(open_avg_file_direct, self.window, file)
                return True
            elif file.lower().endswith('.spe'):
                wx.CallAfter(open_spe_file, self.window, file)
                return True
        return False


def load_library_data_WITHEXCEL():
   wb = openpyxl.load_workbook('KherveFitting_library.xlsx')
   sheet = wb['Library']
   data = {}
   for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
       element, orbital, full_name, auger, ke_be, position, ds, rsf, instrument = row
       key = (element, orbital)
       if key not in data:
           data[key] = {}
       data[key][instrument] = {
           'position': position,
           'ds': ds,
           'rsf': rsf,
           'row': row_idx
       }
   return data


def load_library_data():
    with open('KherveFitting_library.json', 'r') as f:
        json_data = json.load(f)

    # Convert string keys back to tuples
    data = {}
    for k, v in json_data.items():
        element, orbital = k.split('_')
        data[(element, orbital)] = v
    return data


def load_library_data_NEWBUTNO():
    wb = openpyxl.load_workbook('KherveFitting_library.xlsx')
    sheet = wb['Library']
    data = {}
    instruments = set()  # Create set for unique instruments

    for row in sheet.iter_rows(min_row=2, values_only=True):
        element, orbital, full_name, auger, ke_be, position, ds, rsf, instrument = row
        instruments.add(instrument)  # Add each instrument to set
        key = (element, orbital)
        if key not in data:
            data[key] = {}
        data[key][instrument] = {
            'position': position,
            'ds': ds,
            'rsf': rsf
        }
    return data, sorted(list(instruments))  # Return data and sorted instruments list

def load_recent_files_from_config(window):
    config = window.load_config()
    window.recent_files = config.get('recent_files', [])
    update_recent_files_menu(window)

def update_recent_files(window, file_path):
    if file_path in window.recent_files:
        window.recent_files.remove(file_path)
    window.recent_files.insert(0, file_path)
    window.recent_files = window.recent_files[:window.max_recent_files]
    update_recent_files_menu(window)
    window.save_config()  # Call save_config directly on the window object


def open_spe_file_olderPhi(window, file_path):
    try:
        from yadg.extractors.phi.spe import extract
        import openpyxl

        # Extract data from SPE file
        data = extract(fn=file_path)

        # Create new Excel workbook
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        # Process each core level
        for core_level in data:
            # Create sheet with core level name
            ws = wb.create_sheet(title=core_level)

            # Get energy and intensity values
            energy = data[core_level].coords["E"].values
            intensity = data[core_level].data_vars["y"].values

            # Set column headers
            ws["A1"] = "BE"
            ws["B1"] = "Corrected Data"
            ws["C1"] = "Raw Data"
            ws["D1"] = "Transmission"

            # Fill data
            for i, (e, inten) in enumerate(zip(energy, intensity), start=2):
                ws[f"A{i}"] = e
                ws[f"B{i}"] = inten
                ws[f"C{i}"] = inten
                ws[f"D{i}"] = 1.0

        # Create Experimental Description sheet
        exp_sheet = wb.create_sheet("Experimental description")
        exp_sheet.column_dimensions['A'].width = 30
        exp_sheet.column_dimensions['B'].width = 50

        # Read header information using binary mode
        with open(file_path, 'rb') as f:
            header_lines = []
            in_header = False
            for line in f:
                try:
                    decoded_line = line.decode('utf-8', errors='ignore').strip()
                    if decoded_line == 'SOFH':
                        in_header = True
                        continue
                    if decoded_line == 'EOFH':
                        break
                    if in_header and decoded_line:
                        header_lines.append(decoded_line)
                except UnicodeDecodeError:
                    continue

        # Write header info to exp sheet
        for i, line in enumerate(header_lines, start=1):
            if ':' in line:
                key, value = line.split(':', 1)
                exp_sheet[f"A{i}"] = key.strip()
                exp_sheet[f"B{i}"] = value.strip()

        # Save Excel file
        excel_path = os.path.splitext(file_path)[0] + ".xlsx"
        wb.save(excel_path)

        # Open the created Excel file using existing function
        open_xlsx_file(window, excel_path)

    except Exception as e:
        # wx.MessageBox(f"Error processing SPE file: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
        window.show_popup_message2("Error", f"Error processing SPE file: {str(e)}")

def open_spe_file_BEST_SURVEY(window, file_path):
    import numpy as np
    import openpyxl
    import re
    import os
    import struct

    def extract_float_sequences(data, min_length=100):
        sequences = []
        positions = []

        for offset in range(0, 512):
            values = []
            i = offset
            while i + 4 <= len(data):
                try:
                    val = struct.unpack('<f', data[i:i + 4])[0]
                    if 0 <= val < 1e6:
                        values.append(val)
                    else:
                        if len(values) >= min_length:
                            sequences.append(values)
                            positions.append(i)
                        values = []
                    i += 4
                except:
                    i += 4
            if len(values) >= min_length:
                sequences.append(values)
                positions.append(i)
        return sequences, positions

    try:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        with open(file_path, 'rb') as f:
            content = f.read()

        header_match = re.search(rb'SOFH(.*?)EOFH', content, re.DOTALL)
        if not header_match:
            raise ValueError("Cannot find header section (SOFH...EOFH) in SPE file")

        header_text = header_match.group(1).decode('utf-8', errors='ignore')
        header_lines = [line.strip() for line in header_text.strip().split('\n')]

        a, b = 31.826, 0.229
        for line in header_lines:
            if 'IntensityCalCoeff:' in line:
                _, coeffs = line.split(':', 1)
                a, b = map(float, coeffs.strip().split())
                break

        spectral_regions = []
        region_count = 0
        for line in header_lines:
            if line.startswith('NoSpectralReg:'):
                region_count = int(line.split(':')[1].strip())
            if line.startswith('SpectralRegDef:'):
                parts = line.split()
                is_active = int(parts[2])
                if is_active == 1:
                    # Get the name and CHANGE 1: Replace 'Su1s' with 'Survey'
                    name = parts[3]
                    if name == 'Su1s':
                        name = 'Survey'

                    spectral_regions.append({
                        'index': int(parts[1]),
                        'name': name,
                        'start_energy': float(parts[7]),
                        'end_energy': float(parts[8]),
                        'step_size': float(parts[5])
                    })

        data_bytes = content[content.find(b'EOFH') + 4:]

        float_sequences, _ = extract_float_sequences(data_bytes)
        if float_sequences:
            intensity_values = max(float_sequences, key=lambda x: max(x))
        else:
            raise ValueError("No valid intensity data found in binary section")

        for region in spectral_regions:
            sheet_name = region['name']
            ws = wb.create_sheet(title=sheet_name)
            ws["A1"] = "BE"
            ws["B1"] = "Corrected Data"
            ws["C1"] = "Raw Data"
            ws["D1"] = "Transmission"

            start_energy = region['start_energy']
            end_energy = region['end_energy']
            num_points = len(intensity_values)

            # CHANGE 2: Swap start and end energies to correct BE scale orientation
            # Note: We're using end_energy first, then start_energy to flip the BE scale
            energy_values = np.linspace(start_energy, end_energy, num_points)

            ke_values = 1486.6 - energy_values
            transmission = a * np.power(ke_values, -b)
            corrected_intensity = np.array(intensity_values) / transmission

            for i, (e, ci, ri, t) in enumerate(zip(energy_values, corrected_intensity, intensity_values, transmission),
                                               start=2):
                ws[f"A{i}"] = float(e)
                ws[f"B{i}"] = float(ci)
                ws[f"C{i}"] = float(ri)
                ws[f"D{i}"] = float(t)

        exp_sheet = wb.create_sheet("Experimental description")
        exp_sheet.column_dimensions['A'].width = 30
        exp_sheet.column_dimensions['B'].width = 50
        for i, line in enumerate(header_lines, start=1):
            if ':' in line:
                key, value = line.split(':', 1)
                exp_sheet[f"A{i}"] = key.strip()
                exp_sheet[f"B{i}"] = value.strip()

        excel_path = os.path.splitext(file_path)[0] + ".xlsx"
        wb.save(excel_path)

        from libraries.Open import open_xlsx_file
        open_xlsx_file(window, excel_path)

    except Exception as e:
        import traceback
        traceback.print_exc()
        window.show_popup_message2("Error", f"Error processing SPE file: {str(e)}")


def parse_spe_binary_data_B2(data_bytes, min_length=100, start_offset=0):
    """Extract next valid intensity sequence from SPE binary data"""
    import struct

    print(f"\n🔍 Looking for sequence of at least {min_length} values from offset {start_offset}")

    for offset in range(start_offset, len(data_bytes) - 4):
        values = []
        pos = offset

        while pos + 4 <= len(data_bytes):
            try:
                val = struct.unpack('<f', data_bytes[pos:pos + 4])[0]
                if 0 <= val < 100000 and val > 0.01:
                    values.append(val)
                else:
                    break
            except Exception as e:
                print(f"  ❌ Unpacking failed at pos {pos}: {e}")
                break
            pos += 4

        if len(values) >= min_length:
            avg = sum(values) / len(values)
            max_val = max(values)
            min_val = min(values)
            range_val = max_val - min_val
            score = range_val * len(values)

            if score > 1000:
                print(f"✅ Found valid sequence at offset {offset}")
                print(f"  ➤ Length: {len(values)} | Range: {min_val:.1f} – {max_val:.1f} | Avg: {avg:.1f}")
                print(f"  ➤ First 3 values: {values[:3]}")
                print(f"  ➤ Last  3 values: {values[-3:]}")
                print(f"  ➤ Ending byte position: {pos}")
                return values[:min_length], pos

    print("❌ No valid intensity data found in this section")
    return [], start_offset


def open_spe_file_B2(window, file_path):
    """Process a PHI SPE file and convert it to Excel format"""
    import numpy as np
    import openpyxl
    import re
    import os
    import struct

    try:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        with open(file_path, 'rb') as f:
            content = f.read()

        print(f"\n=== Processing SPE file: {file_path} ===")
        print(f"File size: {len(content)} bytes")

        # Extract header
        header_match = re.search(rb'SOFH(.*?)EOFH', content, re.DOTALL)
        if not header_match:
            raise ValueError("Cannot find header section (SOFH...EOFH) in SPE file")

        header_text = header_match.group(1).decode('utf-8', errors='ignore')
        header_lines = [line.strip() for line in header_text.strip().split('\n')]
        print(f"Header found: {len(header_lines)} lines")

        # Extract intensity calibration coefficients
        a, b = 31.826, 0.229  # Default values
        for line in header_lines:
            if 'IntensityCalCoeff:' in line:
                _, coeffs = line.split(':', 1)
                a, b = map(float, coeffs.strip().split())
                print(f"Intensity calibration: a={a}, b={b}")
                break

        # Parse active regions and their parameters
        active_regions = []
        region_info = {}

        for line in header_lines:
            if line.startswith('SpectralRegDef:'):
                parts = line.split()
                if len(parts) >= 12 and int(parts[2]) == 1:  # Active region
                    idx = int(parts[1])
                    name = parts[3]
                    if name == 'Su1s':
                        name = 'Survey'

                    point_count = int(parts[5])
                    step_size = float(parts[6])
                    start_energy = float(parts[7])
                    end_energy = float(parts[8])

                    region_info[idx] = {
                        'name': name,
                        'points': point_count,
                        'step_size': step_size,
                        'start_energy': start_energy,
                        'end_energy': end_energy
                    }
                    active_regions.append(idx)
                    print(f"Region {idx}: {name}, {point_count} points, {start_energy}-{end_energy} eV")

        # Get binary data after EOFH
        # After the header parsing section:
        data_start = content.find(b'EOFH') + 4
        binary_data = content[data_start:]
        print(f"📊 Binary data section starts at byte {data_start}, size: {len(binary_data)} bytes")

        # For more reliable data extraction, scan the entire binary section
        for region in active_regions:
            sheet_name = region['name']
            points = region['points']
            start_energy = region['start_energy']
            end_energy = region['end_energy']

            print(f"\n🎯 Processing region: {sheet_name} ({points} points)")

            # Create sheet
            ws = wb.create_sheet(title=sheet_name)
            ws["A1"] = "BE"
            ws["B1"] = "Corrected Data"
            ws["C1"] = "Raw Data"
            ws["D1"] = "Transmission"

            # Find intensity values by scanning all possible offsets and scoring each sequence
            best_sequence = None
            best_score = 0
            best_offset = 0

            # Use a larger step size to speed up the scan, then refine
            for base_offset in range(0, len(binary_data) - points * 4, 16):
                try:
                    values = []
                    for i in range(points):
                        if base_offset + i * 4 + 4 > len(binary_data):
                            break
                        val = struct.unpack('<f', binary_data[base_offset + i * 4:base_offset + i * 4 + 4])[0]
                        if 0 <= val < 1e6:  # Reasonable intensity range
                            values.append(val)
                        else:
                            break

                    if len(values) == points:  # Found complete sequence
                        # Score based on range and statistical properties
                        max_val = max(values)
                        range_val = max_val - min(values)
                        score = range_val * (sum(1 for v in values if v > 1) / len(values))

                        if score > best_score:
                            best_score = score
                            best_sequence = values
                            best_offset = base_offset
                            print(f"✅ Found candidate data at offset {base_offset} with score {score:.1f}")
                except:
                    continue

            if best_sequence:
                intensity_values = best_sequence
                print(f"✅ Using data from offset {best_offset} (score: {best_score:.1f})")
                print(f"   ➤ Range: {min(intensity_values):.1f} – {max(intensity_values):.1f}")
            else:
                print(f"❌ No valid data found for {sheet_name}, using placeholder values")
                intensity_values = [1.0] * points

            # Energy scale and transmission
            energy_values = np.linspace(start_energy, end_energy, points)
            ke_values = 1486.6 - energy_values
            transmission = a * np.power(ke_values, -b)
            # corrected = np.array(intensity_values) / transmission
            corrected_intensity = np.array(intensity_values) / transmission

            for i, (e, ci, ri, t) in enumerate(zip(energy_values, corrected_intensity, intensity_values, transmission),
                                               start=2):
                ws[f"A{i}"] = float(e)
                ws[f"B{i}"] = float(ci)
                ws[f"C{i}"] = float(ri)
                ws[f"D{i}"] = float(t)

        # Add experimental description sheet
        exp_sheet = wb.create_sheet("Experimental description")
        exp_sheet.column_dimensions['A'].width = 30
        exp_sheet.column_dimensions['B'].width = 50
        for i, line in enumerate(header_lines, start=1):
            if ':' in line:
                key, value = line.split(':', 1)
                exp_sheet[f"A{i}"] = key.strip()
                exp_sheet[f"B{i}"] = value.strip()

        excel_path = os.path.splitext(file_path)[0] + ".xlsx"
        wb.save(excel_path)
        print(f"Saved Excel file: {excel_path}")

        from libraries.Open import open_xlsx_file
        open_xlsx_file(window, excel_path)

    except Exception as e:
        import traceback
        traceback.print_exc()
        window.show_popup_message2("Error", f"Error processing SPE file: {str(e)}")

def open_spe_file(window, file_path):
    import numpy as np
    import openpyxl
    import re
    import os
    import struct

    def parse_spe_binary_data(data_bytes, min_length=100, start_offset=0):
        """Extract next valid intensity sequence from SPE binary data"""
        import struct

        print(f"🔍 Looking for sequence of at least {min_length} values from offset {start_offset}")

        # For first region, do a thorough search for valid data
        if start_offset <= 2:
            # Known good starting positions based on observed patterns
            # Adding offset 114 for survey scans with single regions
            priority_offsets = [306, 2, 114, 310, 314]

            # Try priority offsets first for the first region
            for offset in priority_offsets:
                if offset + 4 * min_length > len(data_bytes):
                    continue

                values = []
                pos = offset

                # Try to extract min_length values
                for i in range(min_length):
                    if pos + 4 > len(data_bytes):
                        values = []
                        break

                    try:
                        val = struct.unpack('<f', data_bytes[pos:pos + 4])[0]
                        if 0 <= val < 10000 and (i > 0 or val > 0.01):  # Skip tiny first values
                            values.append(val)
                        else:
                            values = []  # Reset on invalid values
                            break
                    except:
                        values = []
                        break
                    pos += 4

                if len(values) == min_length:
                    print(f"✅ Found valid sequence at offset {offset}")
                    print(f"  ➤ Length: {len(values)} | Range: {min(values):.1f} – {max(values):.1f}")
                    print(f"  ➤ First 3 values: {values[:3]}")
                    print(f"  ➤ Last  3 values: {values[-3:]}")
                    print(f"  ➤ Ending byte position: {pos}")
                    return values, pos

        # For subsequent regions, use the next consecutive data
        else:
            values = []
            pos = start_offset

            # Try to extract min_length values from the exact starting position
            for i in range(min_length):
                if pos + 4 > len(data_bytes):
                    break

                try:
                    val = struct.unpack('<f', data_bytes[pos:pos + 4])[0]
                    if 0 <= val < 10000:  # Reasonable values
                        values.append(val)
                    else:
                        values = []
                        break
                except:
                    values = []
                    break
                pos += 4

            if len(values) == min_length:
                print(f"✅ Found consecutive sequence at offset {start_offset}")
                print(f"  ➤ Length: {len(values)} | Range: {min(values):.1f} – {max(values):.1f}")
                print(f"  ➤ First 3 values: {values[:3]}")
                print(f"  ➤ Last  3 values: {values[-3:]}")
                print(f"  ➤ Ending byte position: {pos}")
                return values, pos

        # Fall back to scanning the entire data block if needed
        print("🛟 Fallback: scanning entire data block")
        for offset in range(2, len(data_bytes) - 4 * min_length, 4):
            values = []
            pos = offset

            for i in range(min_length):
                if pos + 4 > len(data_bytes):
                    break

                try:
                    val = struct.unpack('<f', data_bytes[pos:pos + 4])[0]
                    if 0 <= val < 10000:
                        values.append(val)
                    else:
                        values = []
                        break
                except:
                    values = []
                    break
                pos += 4

            if len(values) == min_length:
                print(f"✅ Found sequence at offset {offset} using fallback")
                print(f"  ➤ Length: {len(values)} | Range: {min(values):.1f} – {max(values):.1f}")
                print(f"  ➤ First 3 values: {values[:3]}")
                print(f"  ➤ Last  3 values: {values[-3:]}")
                print(f"  ➤ Ending byte position: {pos}")
                return values, pos

        print("❌ No valid intensity data found")
        return [], start_offset

    try:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        with open(file_path, 'rb') as f:
            content = f.read()

        print(f"\n== Processing SPE file: {file_path} ==")
        print(f"📦 File size: {len(content)} bytes")

        # ————— Parse header —————
        header_match = re.search(rb'SOFH(.*?)EOFH', content, re.DOTALL)
        if not header_match:
            raise ValueError("❗ Cannot find header section (SOFH...EOFH) in SPE file")

        header_text = header_match.group(1).decode('utf-8', errors='ignore')
        header_lines = [line.strip() for line in header_text.strip().split('\n')]
        print(f"📄 Header found: {len(header_lines)} lines")

        a, b = 31.826, 0.229  # Default
        for line in header_lines:
            if 'IntensityCalCoeff:' in line:
                _, coeffs = line.split(':', 1)
                a, b = map(float, coeffs.strip().split())
                print(f"🔧 Intensity calibration: a={a}, b={b}")
                break

        region_info = {}
        active_regions = []

        for line in header_lines:
            if line.startswith('SpectralRegDef:'):
                parts = line.split()
                if len(parts) >= 12 and int(parts[2]) == 1:
                    idx = int(parts[1])
                    name = parts[3]
                    if name == 'Su1s':
                        name = 'Survey'
                    region_info[idx] = {
                        'name': name,
                        'points': int(parts[5]),
                        'step_size': float(parts[6]),
                        'start_energy': float(parts[7]),
                        'end_energy': float(parts[8])
                    }
                    active_regions.append(idx)
                    print(f"🔬 Region {idx}: {name}, {region_info[idx]['points']} pts, "
                          f"{region_info[idx]['start_energy']}-{region_info[idx]['end_energy']} eV")

        data_start = content.find(b'EOFH') + 4
        data_bytes = content[data_start:]
        print(f"📊 Binary data section starts at byte {data_start}, size: {len(data_bytes)} bytes")

        # ————— Extract intensity data —————
        intensity_data = {}
        data_offset = 0

        for idx in active_regions:
            info = region_info[idx]
            name, points = info['name'], info['points']
            print(f"\n🎯 Trying to find data for {name} ({points} points)")

            sequence, new_offset = parse_spe_binary_data(data_bytes, points, start_offset=data_offset)

            if sequence and len(sequence) >= points:
                intensity_data[idx] = sequence
                data_offset = new_offset
                print(f"✅ Assigned {points} points for {name}")
            else:
                print(f"🛟 Fallback: scanning entire data block for orphaned {name}")
                sequence, _ = parse_spe_binary_data(data_bytes, points, start_offset=0)
                if sequence and len(sequence) >= points:
                    intensity_data[idx] = sequence
                    print(f"⚠️  Found orphaned region for {name} using fallback scan")
                else:
                    print(f"❌ No valid data found for {name}, using placeholder values")
                    intensity_data[idx] = [1.0] * points

        # ————— Write to Excel —————
        for idx in active_regions:
            info = region_info[idx]
            sheet_name = info['name']
            start_energy = info['start_energy']
            end_energy = info['end_energy']
            intensity_values = intensity_data[idx]

            print(f"\n📄 Creating sheet for {sheet_name}")
            ws = wb.create_sheet(title=sheet_name)
            ws["A1"] = "BE"
            ws["B1"] = "Corrected Data"
            ws["C1"] = "Raw Data"
            ws["D1"] = "Transmission"

            energy_values = np.linspace(start_energy, end_energy, len(intensity_values))
            ke_values = 1486.6 - energy_values
            transmission = a * np.power(ke_values, -b)
            corrected = np.array(intensity_values) / transmission

            print(f"⚙️  Using {len(intensity_values)} data points for {sheet_name}")
            print(f"   ➤ Energy scale: {start_energy} to {end_energy}")
            print(f"   ➤ First 5 energies: {energy_values[:5]}")
            print(f"   ➤ Last 5 energies: {energy_values[-5:]}")

            for i, (e, ci, ri, t) in enumerate(zip(energy_values, corrected, intensity_values, transmission), start=2):
                ws[f"A{i}"] = float(e)
                ws[f"B{i}"] = float(ci)
                ws[f"C{i}"] = float(ri)
                ws[f"D{i}"] = float(t)

        # ————— Experimental description sheet —————
        exp_sheet = wb.create_sheet("Experimental description")
        exp_sheet.column_dimensions['A'].width = 30
        exp_sheet.column_dimensions['B'].width = 50
        for i, line in enumerate(header_lines, start=1):
            if ':' in line:
                key, value = line.split(':', 1)
                exp_sheet[f"A{i}"] = key.strip()
                exp_sheet[f"B{i}"] = value.strip()

        excel_path = os.path.splitext(file_path)[0] + ".xlsx"
        wb.save(excel_path)
        print(f"💾 Saved Excel file: {excel_path}")

        from libraries.Open import open_xlsx_file
        open_xlsx_file(window, excel_path)

    except Exception as e:
        import traceback
        traceback.print_exc()
        window.show_popup_message2("Error", f"Error processing SPE file: {str(e)}")


def open_spe_file_NONONON(window, file_path):
    import numpy as np
    import openpyxl
    import re
    import os
    import struct

    try:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        with open(file_path, 'rb') as f:
            content = f.read()

        print(f"\n== Processing SPE file: {file_path} ==")
        print(f"📦 File size: {len(content)} bytes")

        # ————— Parse header —————
        header_match = re.search(rb'SOFH(.*?)EOFH', content, re.DOTALL)
        if not header_match:
            raise ValueError("❗ Cannot find header section (SOFH...EOFH) in SPE file")

        header_text = header_match.group(1).decode('utf-8', errors='ignore')
        header_lines = [line.strip() for line in header_text.strip().split('\n')]
        print(f"📄 Header found: {len(header_lines)} lines")

        # Extract intensity calibration coefficients
        a, b = 31.826, 0.229  # Default values
        for line in header_lines:
            if 'IntensityCalCoeff:' in line:
                _, coeffs = line.split(':', 1)
                a, b = map(float, coeffs.strip().split())
                print(f"🔧 Intensity calibration: a={a}, b={b}")
                break

        # Parse active regions
        active_regions = []
        for line in header_lines:
            if line.startswith('SpectralRegDef:'):
                parts = line.split()
                if len(parts) >= 12 and int(parts[2]) == 1:  # Active region
                    region_name = parts[3]
                    if region_name == 'Su1s':
                        region_name = 'Survey'

                    active_regions.append({
                        'index': int(parts[1]),
                        'name': region_name,
                        'points': int(parts[5]),
                        'step': float(parts[6]),
                        'start_energy': float(parts[7]),
                        'end_energy': float(parts[8])
                    })
                    print(f"🔬 Region {len(active_regions)}: {region_name}, {int(parts[5])} pts, "
                          f"{float(parts[7])}-{float(parts[8])} eV")

        # ————— Extract binary data —————
        data_start = content.find(b'EOFH') + 4
        binary_data = content[data_start:]
        print(f"📊 Binary data section starts at byte {data_start}, size: {len(binary_data)} bytes")

        # Binary header is usually 24 bytes (6 integers)
        header_size = 24

        # Process each region in order
        for region in active_regions:
            sheet_name = region['name']
            points = region['points']
            start_energy = region['start_energy']
            end_energy = region['end_energy']

            print(f"\n🎯 Processing region: {sheet_name} ({points} points)")

            # Create sheet
            ws = wb.create_sheet(title=sheet_name)
            ws["A1"] = "BE"
            ws["B1"] = "Corrected Data"
            ws["C1"] = "Raw Data"
            ws["D1"] = "Transmission"

            # Calculate offset for this region's data
            # First binary header (24 bytes) followed by floats (4 bytes each)
            # Each region has its own binary data block
            region_index = region['index'] - 1  # 0-based index

            # Try different known offsets where data might start
            offsets_to_try = [
                header_size,  # Standard offset
                header_size + region_index * 4,  # Simple sequential
                header_size + region_index * points * 4  # Full blocks
            ]

            # For more complex files with multiple regions
            if region_index > 0:
                # Add additional potential offsets
                previous_points_sum = sum(r['points'] for r in active_regions[:region_index])
                offsets_to_try.extend([
                    header_size + previous_points_sum * 4,
                    header_size + (region_index * header_size) + (previous_points_sum * 4)
                ])

            # Try each offset until we find valid data
            intensity_values = None
            used_offset = None

            for offset in offsets_to_try:
                if offset + (points * 4) > len(binary_data):
                    continue

                # Try to read data block
                try:
                    test_values = []
                    for i in range(points):
                        pos = offset + (i * 4)
                        value = struct.unpack('<f', binary_data[pos:pos + 4])[0]
                        # Only accept reasonable intensity values
                        if 0 <= value < 100000:
                            test_values.append(value)
                        else:
                            test_values = []
                            break

                    # If we got enough valid values, use this offset
                    if len(test_values) == points:
                        intensity_values = test_values
                        used_offset = offset
                        print(f"✅ Found valid data at offset {offset}")
                        print(f"   ➤ Range: {min(intensity_values):.1f} – {max(intensity_values):.1f}")
                        break
                except Exception as e:
                    pass

            # If we didn't find valid data at any offset, do a more thorough scan
            if intensity_values is None:
                print(f"⚠️ Could not find data at standard offsets, performing full scan")
                for offset in range(0, len(binary_data) - (points * 4), 4):
                    try:
                        test_values = []
                        for i in range(points):
                            pos = offset + (i * 4)
                            value = struct.unpack('<f', binary_data[pos:pos + 4])[0]
                            if 0 <= value < 100000:
                                test_values.append(value)
                            else:
                                test_values = []
                                break

                        if len(test_values) == points:
                            max_val = max(test_values)
                            if max_val > 10:  # Must have some significant counts
                                intensity_values = test_values
                                used_offset = offset
                                print(f"✅ Found valid data in full scan at offset {offset}")
                                print(f"   ➤ Range: {min(intensity_values):.1f} – {max(intensity_values):.1f}")
                                break
                    except:
                        pass

            # If we still didn't find data, use placeholder values
            if intensity_values is None:
                print(f"❌ No valid data found for {sheet_name}, using placeholder values")
                intensity_values = [1.0] * points

            # Energy scale and transmission function
            energy_values = np.linspace(start_energy, end_energy, points)
            ke_values = 1486.6 - energy_values
            transmission = a * np.power(ke_values, -b)
            corrected = np.array(intensity_values) / transmission

            # Add data to sheet
            for i, (e, ci, ri, t) in enumerate(zip(energy_values, corrected, intensity_values, transmission), start=2):
                ws[f"A{i}"] = float(e)
                ws[f"B{i}"] = float(ci)
                ws[f"C{i}"] = float(ri)
                ws[f"D{i}"] = float(t)

            print(f"💾 Added {points} data points to sheet {sheet_name}")

        # ————— Experimental description sheet —————
        exp_sheet = wb.create_sheet("Experimental description")
        exp_sheet.column_dimensions['A'].width = 30
        exp_sheet.column_dimensions['B'].width = 50

        for i, line in enumerate(header_lines, start=1):
            if ':' in line:
                key, value = line.split(':', 1)
                exp_sheet[f"A{i}"] = key.strip()
                exp_sheet[f"B{i}"] = value.strip()

        # Save file and open in the application
        excel_path = os.path.splitext(file_path)[0] + ".xlsx"
        wb.save(excel_path)
        print(f"💾 Saved Excel file: {excel_path}")

        from libraries.Open import open_xlsx_file
        open_xlsx_file(window, excel_path)

    except Exception as e:
        import traceback
        traceback.print_exc()
        window.show_popup_message2("Error", f"Error processing SPE file: {str(e)}")



def open_spe_file_dialog(window):
    with wx.FileDialog(window, "Open SPE file", wildcard="PHI files (*.spe)|*.spe",
                       style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
        if fileDialog.ShowModal() == wx.ID_CANCEL:
            return
        file_path = fileDialog.GetPath()
        open_spe_file(window, file_path)



def validate_data(df):
    """Validate the processed data"""
    issues = []

    if df['Binding Energy'].isna().any():
        issues.append("Missing binding energy values detected")

    if df['Raw Data'].isna().any():
        issues.append("Missing intensity values detected")

    if not np.allclose(df['Raw Data'], df['Corrected Data']):
        issues.append("Mismatch between raw and corrected data")

    if not np.allclose(df['Transmission'], 1.0):
        issues.append("Transmission values not all set to 1.0")

    return issues


def import_mrs_file(window):
    import wx
    from openpyxl import Workbook
    import os

    with wx.FileDialog(window, "Open MRS file", wildcard="MRS files (*.mrs)|*.mrs",
                       style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
        if fileDialog.ShowModal() == wx.ID_CANCEL:
            return
        file_path = fileDialog.GetPath()

    try:
        # Create new Excel workbook
        wb = Workbook()
        wb.remove(wb.active)

        raw_data = []
        be_start = None
        be_end = None
        step_size = None

        with open(file_path, 'r') as f:
            lines = f.readlines()

        # Extract metadata and data block
        in_data_section = False
        delimiter_count = 0  # Track the number of "!" encountered

        for line in lines:
            if line.strip() == "!":
                delimiter_count += 1
                if delimiter_count == 2:  # Stop processing after the second "!"
                    break
                in_data_section = True
                continue

            if in_data_section and line.strip().isdigit():
                raw_data.append(int(line.strip()))

            # Parse metadata
            if 'lo_be=' in line:
                be_start = float(line.split('=')[1].strip())
            elif 'up_be=' in line:
                be_end = float(line.split('=')[1].strip())

        # Calculate step size if not directly provided
        if be_start is not None and be_end is not None:
            step_size = (be_end - be_start) / (len(raw_data) - 1)

        if be_start is None or (be_end is None and step_size is None):
            raise ValueError("Missing essential metadata: 'lo_be', 'up_be', or step size.")

        # Compute BE scale
        be_values = [be_start + i * step_size for i in range(len(raw_data))]

        # Create Excel sheet
        sheet_name = os.path.splitext(os.path.basename(file_path))[0]
        ws = wb.create_sheet(sheet_name)

        # Add headers
        ws['A1'] = 'BE'
        ws['B1'] = 'Raw Data'

        # Add data
        for i, (be, intensity) in enumerate(zip(be_values, raw_data), start=2):
            ws[f'A{i}'] = be
            ws[f'B{i}'] = intensity

        # Save Excel file
        excel_path = os.path.splitext(file_path)[0] + ".xlsx"
        wb.save(excel_path)

        # Open the created Excel file
        open_xlsx_file(window, excel_path)

    except Exception as e:
        # wx.MessageBox(f"Error processing MRS file: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
        self.parent.show_popup_message2("Error", f"Error processing MRS file: {str(e)}")


def import_avantage_file_direct(window, file_path):
    import re
    import openpyxl

    wb = openpyxl.load_workbook(file_path)
    new_file_path = os.path.splitext(file_path)[0] + "_Kfitting.xlsx"
    new_wb = openpyxl.Workbook()
    new_wb.remove(new_wb.active)

    sheets_to_process = []
    for sheet_name in wb.sheetnames:
        if "Survey" in sheet_name or "Scan" in sheet_name:
            sheets_to_process.append(sheet_name)

    for sheet_name in sheets_to_process:
        sheet = wb[sheet_name]

        # Extract element name (e.g., C1s, O1s)
        if "Survey" in sheet_name or "survey" in sheet_name:
            base_name = "Survey"
        else:
            parts = sheet_name.split()
            base_name = parts[0]  # Get element (C1s, O1s, etc.)

        # Check if multi-sample (D8 cell not empty and > 1)
        multi_sample = False
        num_samples = 1
        if sheet.cell(row=8, column=4).value is not None:
            try:
                num_samples = int(sheet.cell(row=8, column=4).value)
                if num_samples > 1:
                    multi_sample = True
            except (ValueError, TypeError):
                pass

        # Find data start row (default to 19 for Thermo files)
        start_row = 19
        for row_idx in range(1, sheet.max_row + 1):
            if sheet.cell(row=row_idx, column=1).value == "eV":
                start_row = row_idx + 1
                break

        if multi_sample:
            # Process each sample into its own sheet
            for sample_idx in range(num_samples):
                sample_name = f"{base_name}{sample_idx if sample_idx > 0 else ''}"
                new_sheet = new_wb.create_sheet(sample_name)
                new_sheet['A1'] = "Binding Energy"
                new_sheet['B1'] = "Raw Data"

                # A column is binding energy (col 1), data is in col C+sample_idx (col 3+sample_idx)
                data_col = 3 + sample_idx  # C is 3, D is 4, etc.

                # Copy data
                row_new = 2
                for row in range(start_row, sheet.max_row + 1):
                    be_value = sheet.cell(row=row, column=1).value
                    intensity_value = sheet.cell(row=row, column=data_col).value

                    if be_value is None or intensity_value is None:
                        continue

                    new_sheet.cell(row=row_new, column=1, value=be_value)
                    new_sheet.cell(row=row_new, column=2, value=intensity_value)
                    row_new += 1
        else:
            # Single sample - process as before
            new_name = base_name
            # Extract number if present in format like "C1s Scan (2)"
            number_match = re.search(r'\((\d+)\)', sheet_name)
            if number_match:
                number = number_match.group(1)
                new_name = f"{base_name}{number}"

            new_sheet = new_wb.create_sheet(new_name)
            new_sheet['A1'] = "Binding Energy"
            new_sheet['B1'] = "Raw Data"

            for row in range(start_row, sheet.max_row + 1):
                be_value = sheet.cell(row=row, column=1).value
                intensity_value = sheet.cell(row=row, column=3).value  # Column C

                if be_value is None or intensity_value is None:
                    continue

                new_sheet['A{}'.format(row - start_row + 2)] = be_value
                new_sheet['B{}'.format(row - start_row + 2)] = intensity_value

    new_wb.save(new_file_path)
    open_xlsx_file(window, new_file_path)

def import_avantage_file_direct_xls(window, file_path):
    import xlrd
    wb_xls = xlrd.open_workbook(file_path)
    wb_new = openpyxl.Workbook()
    wb_new.remove(wb_new.active)

    for sheet_name in wb_xls.sheet_names():
        sheet = wb_xls.sheet_by_name(sheet_name)
        if "Survey" in sheet_name or "Scan" in sheet_name:
            # Handle Survey sheets
            if "Survey" in sheet_name or "survey" in sheet_name:
                # Extract number if present in "Survey Scan (2)" format
                number_match = re.search(r'\((\d+)\)', sheet_name)
                if number_match:
                    number = number_match.group(1)
                    new_name = f"Survey{number}"
                else:
                    new_name = "Survey"
            else:
                # Extract element name and number for patterns like "C1s Scan (1)"
                parts = sheet_name.split()
                element = parts[0]
                # Check if there's a number in parentheses
                number_match = re.search(r'\((\d+)\)', sheet_name)
                if number_match:
                    number = number_match.group(1)
                    new_name = f"{element}{number}"
                else:
                    new_name = element

            wb.create_sheet(new_name)
            new_sheet = wb[new_name]
            new_sheet['A1'] = "Binding Energy"
            new_sheet['B1'] = "Raw Data"

            start_row = 17
            for row_idx in range(sheet.nrows):
                if sheet.cell_value(row_idx, 0) == "eV":
                    start_row = row_idx + 1
                    break

            for row_idx in range(start_row, sheet.nrows):
                row_values = sheet.row_values(row_idx)
                new_sheet.append([row_values[0]] + row_values[2:])

            for col in new_sheet.iter_cols(min_col=3, max_col=24):
                for cell in col:
                    cell.value = None

    new_file_path = os.path.splitext(file_path)[0] + "_Kfitting.xlsx"
    wb_new.save(new_file_path)
    open_xlsx_file(window, new_file_path)


def open_avg_file_direct(window, avg_file_path):

    # Get the basename without extension
    raw_sheet_name = os.path.basename(avg_file_path).split('.')[0]

    # Normalize the sheet name to follow KherveFitting conventions
    sheet_name = normalize_sheet_name(raw_sheet_name)

    photon_energy, start_energy, width, num_points, y_values = parse_avg_file(avg_file_path)
    be_values = [photon_energy - (start_energy + i * width) for i in range(num_points)]

    df = pd.DataFrame({
        'BE': be_values,
        'Intensity': y_values[:num_points]
    })

    output_path = avg_file_path.rsplit('.', 1)[0] + '.xlsx'

    # Extract metadata
    metadata = extract_metadata_from_avg(avg_file_path)

    # Fields in the order shown in the example
    field_order = [
        'Sample ID', 'Date', 'Time', 'Technique', 'Species & Transition',
        'Number of scans', 'Source Label', 'Source Energy', 'Source width X',
        'Source width Y', 'Pass Energy', 'Work Function', 'Analyzer Mode',
        'Sputtering Energy', 'Take-off Polar Angle', 'Take-off Azimuth',
        'Target Bias', 'Analysis Width X', 'Analysis Width Y', 'X Label',
        'X Units', 'X Start', 'X Step', 'Num Y Values', 'Num Scans',
        'Collection Time', 'Time Correction', 'Y Unit'
    ]

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Use the normalized sheet name instead of the raw file name
        df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Add metadata to the main sheet
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Add experimental description at column 50
        exp_col = 50
        worksheet.cell(row=1, column=exp_col, value="Experimental Description")

        for i, field in enumerate(field_order, start=2):
            worksheet.cell(row=i, column=exp_col, value=field)
            worksheet.cell(row=i, column=exp_col + 1, value=metadata.get(field, "N/A"))

        # Set column widths
        from openpyxl.utils import get_column_letter
        worksheet.column_dimensions[get_column_letter(exp_col)].width = 25
        worksheet.column_dimensions[get_column_letter(exp_col + 1)].width = 40

        # Create separate experimental description sheet
        exp_sheet = workbook.create_sheet(title="Experimental description")
        exp_sheet.column_dimensions['A'].width = 30
        exp_sheet.column_dimensions['B'].width = 50

        exp_sheet['A1'] = "Experimental Description"

        for row, field in enumerate(field_order, start=2):
            exp_sheet[f'A{row}'] = field
            exp_sheet[f'B{row}'] = metadata.get(field, "N/A")

    return output_path

def import_avantage_file(window):
    # Opens a file dialog for user to select file
    with wx.FileDialog(window, "Open Avantage Excel file", wildcard="Excel files (*.xlsx;*.xls)|*.xlsx;*.xls",
                       style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
        if fileDialog.ShowModal() == wx.ID_CANCEL:
            return
        file_path = fileDialog.GetPath()

        # Call appropriate function based on extension
        if file_path.lower().endswith('.xlsx'):
            import_avantage_file_direct_xlsx(window, file_path)
        elif file_path.lower().endswith('.xls'):
            import_avantage_file_direct_xls(window, file_path)


def parse_avg_file_OLD(file_path):
    with open(file_path, 'r') as file:
        content = file.read()

    photon_energy = float(re.search(r'DS_SOPROPID_ENERGY\s+:\s+VT_R4\s+=\s+(\d+\.\d+)', content).group(1))
    start_energy, width, num_points = map(float, re.search(r'\$SPACEAXES=1\s+0=\s+(\d+\.\d+),\s+(\d+\.\d+),\s+(\d+),',
                                                           content).groups())

    # Modified part to handle multiple numbers per line
    y_values = []
    for match in re.findall(r'LIST@\s+\d+=\s+([\d., ]+)', content):
        values = [float(val.strip()) for val in match.split(',')]
        y_values.extend(values)

    return photon_energy, start_energy, width, int(num_points), y_values


def parse_avg_file(file_path):
    with open(file_path, 'r') as file:
        content = file.read()

    photon_energy = float(re.search(r'DS_SOPROPID_ENERGY\s+:\s+VT_R4\s+=\s+(\d+\.\d+)', content).group(1))
    start_energy, width, num_points = map(float, re.search(r'\$SPACEAXES=1\s+0=\s+(\d+\.\d+),\s+(\d+\.\d+),\s+(\d+),',
                                                           content).groups())

    # Extract all intensity values using a different approach
    y_values = []
    in_list_section = False
    list_counter = 0

    for line in content.split('\n'):
        if line.strip().startswith('LIST@'):
            in_list_section = True
            # Get values from the first line
            values_part = line.split('=', 1)[1].strip()
            # Process values from this line
            for val in values_part.split(','):
                if val.strip():
                    try:
                        y_values.append(float(val.strip()))
                    except ValueError:
                        pass
        elif in_list_section:
            # Check if this line likely contains data values
            if ',' in line and not line.strip().startswith('$') and not line.strip().startswith('DS_'):
                # Process values from continuation lines
                for val in line.split(','):
                    if val.strip():
                        try:
                            y_values.append(float(val.strip()))
                        except ValueError:
                            pass
            else:
                # If we hit a line that doesn't look like data, we're probably out of the LIST section
                in_list_section = False

    # Ensure we have the correct number of points
    if len(y_values) > int(num_points):
        y_values = y_values[:int(num_points)]

    # If we still don't have enough points, print warning
    if len(y_values) < int(num_points):
        print(f"Warning: Expected {int(num_points)} points but found only {len(y_values)}.")
        # This is serious - don't silently pad with zeros as it would give misleading data
        # Instead, let's just return what we have and make it known there's an issue

    return photon_energy, start_energy, width, int(num_points), y_values


def create_excel_from_avg(avg_file_path):
    from openpyxl.utils import get_column_letter

    # Extract the raw sheet name from the file name
    raw_sheet_name = os.path.basename(avg_file_path).split()[0]

    # Normalize the sheet name to follow KherveFitting conventions
    sheet_name = normalize_sheet_name(raw_sheet_name)

    photon_energy, start_energy, width, num_points, y_values = parse_avg_file(avg_file_path)

    be_values = [photon_energy - (start_energy + i * width) for i in range(num_points)]

    df = pd.DataFrame({
        'BE': be_values,
        'Intensity': y_values[:num_points]
    })

    output_path = avg_file_path.rsplit('.', 1)[0] + '.xlsx'

    # Extract metadata
    metadata = extract_metadata_from_avg(avg_file_path)

    # Fields in the order shown in the example
    field_order = [
        'Sample ID', 'Date', 'Time', 'Technique', 'Species & Transition',
        'Number of scans', 'Source Label', 'Source Energy', 'Source width X',
        'Source width Y', 'Pass Energy', 'Work Function', 'Analyzer Mode',
        'Sputtering Energy', 'Take-off Polar Angle', 'Take-off Azimuth',
        'Target Bias', 'Analysis Width X', 'Analysis Width Y', 'X Label',
        'X Units', 'X Start', 'X Step', 'Num Y Values', 'Num Scans',
        'Collection Time', 'Time Correction', 'Y Unit'
    ]

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Use normalized sheet name
        df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Add metadata to the main sheet
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Add experimental description at column 50
        exp_col = 50
        worksheet.cell(row=1, column=exp_col, value="Experimental Description")

        for i, field in enumerate(field_order, start=2):
            worksheet.cell(row=i, column=exp_col, value=field)
            worksheet.cell(row=i, column=exp_col + 1, value=metadata.get(field, "N/A"))

        # Set column widths
        worksheet.column_dimensions[get_column_letter(exp_col)].width = 25
        worksheet.column_dimensions[get_column_letter(exp_col + 1)].width = 40

        # Create separate experimental description sheet
        exp_sheet = workbook.create_sheet(title="Experimental description")
        exp_sheet.column_dimensions['A'].width = 30
        exp_sheet.column_dimensions['B'].width = 50

        exp_sheet['A1'] = "Experimental Description"

        for row, field in enumerate(field_order, start=2):
            exp_sheet[f'A{row}'] = field
            exp_sheet[f'B{row}'] = metadata.get(field, "N/A")

    return output_path


def extract_metadata_from_avg(file_path):
    with open(file_path, 'r') as file:
        content = file.read()

    metadata = {
        'Sample ID': "Group",
        'Date': "N/A",
        'Time': "N/A",
        'Technique': "XPS",
        'Species & Transition': "C 1s",
        'Number of scans': "20",
        'Source Label': "KA1486",
        'Source Energy': "1486.68",
        'Source width X': "1E+37",
        'Source width Y': "1E+37",
        'Pass Energy': "20",
        'Work Function': "1E+37",
        'Analyzer Mode': "FAT",
        'Sputtering Energy': "N/A",
        'Take-off Polar Angle': "1E+37",
        'Take-off Azimuth': "1E+37",
        'Target Bias': "1E+37",
        'Analysis Width X': "1E+37",
        'Analysis Width Y': "1E+37",
        'X Label': "Kinetic Energy",
        'X Units': "eV",
        'X Start': "1188.6",
        'X Step': "0.1",
        'Num Y Values': "382",
        'Num Scans': "20",
        'Collection Time': "0.05",
        'Time Correction': "1E+37",
        'Y Unit': "d"
    }

    # Extract values from file where available
    created_match = re.search(r'DS_EXT_SUPROPID_CREATED\s+:\s+VT_DATE\s+=\s+(\d+)/(\d+)/(\d+)\s+(\d+):(\d+):(\d+)',
                              content)
    if created_match:
        day, month, year = created_match.group(1), created_match.group(2), created_match.group(3)
        hour, minute, second = created_match.group(4), created_match.group(5), created_match.group(6)
        metadata['Date'] = f"{year}/{month}/{day}"
        metadata['Time'] = f"{hour}:{minute}:{second}"

    # Update other metadata from file where patterns match
    # Patterns already exist in the parse_avg_file function

    return metadata


def open_avg_file(window):
    with wx.FileDialog(window, "Open AVG file", wildcard="AVG files (*.avg)|*.avg",
                       style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
        if fileDialog.ShowModal() == wx.ID_CANCEL:
            return

        avg_file_path = fileDialog.GetPath()

    try:
        # Use create_excel_from_avg which we need to modify to normalize sheet names
        excel_file_path = create_excel_from_avg(avg_file_path)

        from libraries.Open import open_xlsx_file
        open_xlsx_file(window, excel_file_path)
    except Exception as e:
        wx.MessageBox(f"Error processing AVG file: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

def import_multiple_avg_files(window):
    with wx.DirDialog(window, "Choose a directory containing AVG files",
                      style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST) as dirDialog:

        if dirDialog.ShowModal() == wx.ID_CANCEL:
            return

        root_folder_path = dirDialog.GetPath()

    try:
        # Function to process a single folder
        def process_folder(folder_path):
            folder_name = os.path.basename(folder_path)
            avg_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.avg')]

            if not avg_files:
                return None  # No AVG files in this folder

            # Create Excel file for this folder
            excel_file_path = os.path.join(folder_path, f"{folder_name}.xlsx")
            wb = openpyxl.Workbook()
            wb.remove(wb.active)  # Remove default sheet

            # Process each AVG file in this folder
            for avg_file in avg_files:
                avg_file_path = os.path.join(folder_path, avg_file)

                # Extract raw sheet name from file name (without extension)
                raw_sheet_name = os.path.splitext(avg_file)[0]

                # Normalize the sheet name according to KherveFitting nomenclature
                sheet_name = normalize_sheet_name(raw_sheet_name)

                # Handle duplicate sheet names if needed
                suffix = 1
                original_sheet_name = sheet_name
                while sheet_name in wb.sheetnames:
                    sheet_name = f"{original_sheet_name}{suffix}"
                    suffix += 1

                # Parse AVG file
                try:
                    photon_energy, start_energy, width, expected_points, y_values = parse_avg_file(avg_file_path)

                    # Use the actual number of points we retrieved, not the expected number
                    actual_points = len(y_values)

                    # If we got fewer points than expected, adjust the width of the energy range
                    if actual_points < expected_points:
                        print(f"Warning: {avg_file} - Expected {expected_points} points but found {actual_points}.")
                        # Generate BE values based on the actual number of points we have
                        be_values = [photon_energy - (start_energy + i * width) for i in range(actual_points)]
                    else:
                        be_values = [photon_energy - (start_energy + i * width) for i in range(actual_points)]

                    # Create new sheet
                    ws = wb.create_sheet(title=sheet_name)

                    # Add headers
                    ws.append(["BE", "Intensity"])

                    # Add data - only use the points we actually have
                    for be, intensity in zip(be_values, y_values):
                        ws.append([be, intensity])

                    # Extract metadata if available
                    metadata = extract_metadata_from_avg(avg_file_path)

                    # Add metadata starting at column 50
                    exp_col = 50
                    ws.cell(row=1, column=exp_col, value="Experimental Description")

                    for i, (key, value) in enumerate(metadata.items(), start=2):
                        ws.cell(row=i, column=exp_col, value=key)
                        ws.cell(row=i, column=exp_col + 1, value=value)

                except Exception as e:
                    print(f"Error processing {avg_file}: {str(e)}")
                    continue

            # Save the Excel file if any sheets were created
            if len(wb.sheetnames) > 0:
                wb.save(excel_file_path)
                return excel_file_path
            else:
                return None

        # Process the root folder and all subfolders
        excel_files = []

        # Walk through directory tree
        for dirpath, dirnames, filenames in os.walk(root_folder_path):
            # Process current folder
            excel_path = process_folder(dirpath)
            if excel_path:
                excel_files.append(excel_path)

        # Show results
        if excel_files:
            message = f"Created {len(excel_files)} Excel files:\n" + "\n".join(excel_files)
            wx.MessageBox(message, "Success", wx.OK | wx.ICON_INFORMATION)

            # Open the first Excel file
            from libraries.Open import open_xlsx_file
            open_xlsx_file(window, excel_files[0])
        else:
            wx.MessageBox("No AVG files found in the selected folder or subfolders.",
                          "Information", wx.OK | wx.ICON_INFORMATION)

    except Exception as e:
        wx.MessageBox(f"Error processing AVG files: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

def open_xlsx_file(window, file_path=None):
    """
    Opens an Excel file, loads its data, and updates the application's state accordingly.
    If a corresponding JSON file exists, it loads data from there instead.
    """
    print("Starting open_xlsx_file function")
    if file_path is None:
        with wx.FileDialog(window, "Open XLSX file", wildcard="Excel files (*.xlsx)|*.xlsx",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                file_path = dlg.GetPath()
            else:
                return

    # Store reference to file manager if open
    file_manager_was_open = False
    file_manager_position = None
    try:
        if hasattr(window, 'file_manager') and window.file_manager is not None and window.file_manager.IsShown():
            file_manager_was_open = True
            file_manager_position = window.file_manager.GetPosition()
            window.file_manager.Close()
            window.file_manager = None
    except RuntimeError:
        # The file_manager window was deleted but the reference still exists
        window.file_manager = None
        file_manager_was_open = False
        file_manager_position = None

    window.SetStatusText(f"Selected File: {file_path}", 0)

    try:
        # Clear undo and redo history
        window.history = []
        window.redo_stack = []
        update_undo_redo_state(window)

        # Clear the results grid
        window.results_grid.ClearGrid()
        if window.results_grid.GetNumberRows() > 0:
            window.results_grid.DeleteRows(0, window.results_grid.GetNumberRows())

        # Look for corresponding .json file
        json_file = os.path.splitext(file_path)[0] + '.json'
        if os.path.exists(json_file):
            print(f"Found corresponding .json file: {json_file}")
            with open(json_file, 'r') as f:
                loaded_data = json.load(f)

            # Convert data structure without changing types
            window.Data = convert_from_serializable(loaded_data)

            print("Loaded data from .json file")

            # Populate the results grid
            populate_results_grid(window)
        else:
            print("No corresponding .json file found. Initializing new data.")
            # Initialize the measurement data
            window.Data = Init_Measurement_Data(window)

        # Read the Excel file
        excel_file = pd.ExcelFile(file_path)
        all_sheet_names = excel_file.sheet_names
        sheet_names = [name for name in all_sheet_names if
                       name.lower() not in ["results table", "experimental description"]]

        # Check for invalid sheet names before proceeding
        invalid_sheets = [name for name in sheet_names if name.startswith('Sheet')]
        if invalid_sheets:
            wx.MessageBox(
                f"File contains invalid sheet names: {', '.join(invalid_sheets)}\n\nAll sheets must be named after "
                f"their core level (e.g., C1s, O1s) using one word in proper format without spaces.",
                "Invalid Sheet Names", wx.OK | wx.ICON_WARNING)

            return

        # Check first row values in each sheet
        # Modify the column header check to accept both XPS and Raman formats
        for sheet_name in sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            col1_value = str(df.iloc[0, 0]).strip().upper()
            col2_value = str(df.iloc[0, 1]).strip().upper()

            # Check if it's XPS data
            xps_valid = ('BE' in col1_value or 'BINDING' in col1_value) and \
                        ('RAW DATA' in col2_value or 'CORRECTED DATA' in col2_value or 'INTENSITY' in col2_value)

            # Check if it's Raman data
            raman_valid = ('WAVENUMBER' in col1_value or 'CM-1' in col1_value) and \
                          ('RAW DATA' in col2_value or 'INTENSITY' in col2_value)

            if not (xps_valid or raman_valid):
                wx.MessageBox(
                    f"Sheet '{sheet_name}' has invalid column labels in row 1.\n"
                    f"Column A should contain 'BE'/'Binding Energy' (for XPS) or 'Wavenumber'/'Wavenumber (cm-1)' (for Raman)\n"
                    f"Column B should contain 'Raw Data', 'Corrected Data' or 'Intensity'",
                    "Invalid Column Labels", wx.OK | wx.ICON_WARNING)
                return

        results_table_index = -1
        for i, name in enumerate(all_sheet_names):
            if name.lower() == "results table":
                results_table_index = i
                break

        if results_table_index != -1:
            sheet_names = sheet_names[:results_table_index]

        print(f"Number of sheets: {len(sheet_names)}")

        # Update file path
        window.Data['FilePath'] = file_path

        # If we didn't load from json, populate the data from Excel
        if 'Core levels' not in window.Data or not window.Data['Core levels']:
            window.Data['Number of Core levels'] = 0
            for sheet_name in sheet_names:
                window.Data = add_core_level_Data(window.Data, window, file_path, sheet_name)

        print(f"Final number of core levels: {window.Data['Number of Core levels']}")

        # Load BE correction
        window.load_be_correction()

        # Update sheet names in the combobox
        window.sheet_combobox.Clear()
        window.sheet_combobox.AppendItems(sheet_names)

        # Set the first sheet as the selected one
        first_sheet = sheet_names[0]
        window.sheet_combobox.SetValue(first_sheet)

        # Use on_sheet_selected to update peak parameter grid and plot
        event = wx.CommandEvent(wx.EVT_COMBOBOX.typeId)
        event.SetString(first_sheet)
        window.plot_config.plot_limits.clear()
        on_sheet_selected(window, event)

        # undo and redo
        save_state(window)

        # Update recent files list
        update_recent_files(window, file_path)

        if hasattr(window, 'setup_backup_timer'):
            window.setup_backup_timer()
            print("Auto backup timer checked/updated after file loaded")

        # Refresh any open FileManager windows
        for top_window in wx.GetTopLevelWindows():
            if hasattr(top_window, '__class__') and top_window.__class__.__name__ == 'FileManagerWindow':
                # Force a complete refresh by getting new core levels
                top_window.core_levels = top_window.get_unique_core_levels()
                max_row_index = top_window.get_max_core_level_row_index()
                num_rows = max(10, max_row_index + 1)

                # Clear the grid completely
                if top_window.grid.GetNumberRows() > 0:
                    top_window.grid.DeleteRows(0, top_window.grid.GetNumberRows())
                if top_window.grid.GetNumberCols() > 0:
                    top_window.grid.DeleteCols(0, top_window.grid.GetNumberCols())

                # Add new columns and rows
                num_levels = len(top_window.core_levels)
                top_window.grid.AppendCols(num_levels)
                top_window.grid.AppendRows(num_rows)

                # Set column labels
                for i, level in enumerate(top_window.core_levels):
                    top_window.grid.SetColLabelValue(i, level)

                # Set row labels
                for i in range(num_rows):
                    top_window.grid.SetRowLabelValue(i, str(i))

                # Set column width and row height
                default_col_width = 50
                default_row_height = 20

                for i in range(num_levels):
                    top_window.grid.SetColSize(i, default_col_width)
                for i in range(num_rows):
                    top_window.grid.SetRowSize(i, default_row_height)

                # Set cell alignment
                for row in range(num_rows):
                    for col in range(num_levels):
                        top_window.grid.SetCellAlignment(row, col, wx.ALIGN_CENTER, wx.ALIGN_CENTER)

                # Now populate the grid with new data
                top_window.populate_grid()

        # Create backup before opening file
        from libraries.Utilities import perform_auto_backup
        perform_auto_backup(window)

        # Refresh Sheet
        from libraries.Save import refresh_sheets
        refresh_sheets(window, on_sheet_selected)

        # Reopen file manager if it was open
        if file_manager_was_open:
            from libraries.FileManager import FileManagerWindow
            window.file_manager = FileManagerWindow(window)
            if file_manager_position:
                window.file_manager.SetPosition(file_manager_position)
            window.file_manager.Show()

        print("open_xlsx_file function completed successfully")
    except Exception as e:
        print(f"Error in open_xlsx_file: {str(e)}")
        import traceback
        traceback.print_exc()
        wx.MessageBox(f"Error reading file: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

def convert_from_serializable(obj):
    """
    Recursively converts a serializable object (list or dict) back into its original structure.
    This is used to restore complex data structures that were serialized to JSON.
    """
    if isinstance(obj, list):
        return [convert_from_serializable(item) for item in obj]
    elif isinstance(obj, dict):
        return {k: convert_from_serializable(v) for k, v in obj.items()}
    else:
        return obj

def update_recent_files(window, file_path):
    """
    Updates the list of recently opened files, adding the new file to the top and removing duplicates.
    """
    if file_path in window.recent_files:
        window.recent_files.remove(file_path)
    window.recent_files.insert(0, file_path)
    window.recent_files = window.recent_files[:window.max_recent_files]
    update_recent_files_menu(window)
    save_recent_files_to_config(window)

def update_recent_files_menu(window):
    """
    Refreshes the 'Recent Files' menu with the updated list of recently opened files.
    """
    if hasattr(window, 'recent_files_menu'):
        # Remove all existing menu items
        for item in window.recent_files_menu.GetMenuItems():
            window.recent_files_menu.DestroyItem(item)

        # Add new menu items for recent files
        for i, file_path in enumerate(window.recent_files):
            item = window.recent_files_menu.Append(wx.ID_ANY, os.path.basename(file_path))
            window.Bind(wx.EVT_MENU, lambda evt, fp=file_path: open_xlsx_file(window, fp), item)

def save_recent_files_to_config(window):
    """
    Saves the updated list of recently opened files to the application's configuration.
    """
    window.recent_files = window.recent_files[:window.max_recent_files]
    window.save_config()


def normalize_sheet_name(name):
    """Normalize core level name to standard format."""
    import re
    new_name = name

    lower_name = name.lower()
    if 'xps survey' in lower_name:
        new_name = 'Survey'
    elif 'survey' in lower_name:
        new_name = 'Survey'
    elif 'survey scan' in lower_name:
        new_name = 'Survey'
    elif 'wide' in lower_name:
        new_name = 'Wide'
    elif 'wide scan' in lower_name:
        new_name = 'Wide'
    elif 'su1s' in lower_name:
        new_name = 'Survey'
    else:
        # Remove spaces between element and orbital (e.g., "C 1s" → "C1s")

        match = re.search(r'([A-Z][a-z]?)\s+(\d+[spdf])', name)
        if match:
            element, orbital = match.groups()
            new_name = f"{element}{orbital}"

        # Simplify names like "C1s Scan" to just "C1s"
        match = re.search(r'([A-Z][a-z]?\d+[spdf])', new_name)
        if match and len(new_name) > len(match.group(1)):
            new_name = match.group(1)

    # Preserve sample number suffix if it exists
    suffix_match = re.search(r'(\d+)$', name)
    if suffix_match and not re.search(r'\d+$', new_name):
        new_name = f"{new_name}{suffix_match.group(1)}"

    return new_name


def open_vamas_file_dialog(window):
    """
    Open a file dialog for selecting a VAMAS file and process it.

    This function displays a file dialog for the user to select a VAMAS file,
    then calls open_vamas_file to process the selected file.

    Args:
    window: The main application window object.
    """
    with wx.FileDialog(window, "Open VAMAS file", wildcard="VAMAS files (*.vms)|*.vms",
                       style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
        if fileDialog.ShowModal() == wx.ID_CANCEL:
            return
        file_path = fileDialog.GetPath()
        open_vamas_file(window, file_path)

def open_vamas_file(window, file_path):
    """
    Open and process a VAMAS file, converting it to an Excel file format.
    This function reads a VAMAS file, extracts its data and metadata,
    and creates a new Excel file with multiple sheets for each data block
    and an additional sheet for experimental description.

    Args:
    window: The main application window object.
    file_path: The path to the VAMAS file to be opened.
    """
    try:
        # Clear undo and redo history
        window.history = []
        window.redo_stack = []
        update_undo_redo_state(window)

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"The file {file_path} does not exist.")

        # Copy VAMAS file to current working directory
        vamas_filename = os.path.basename(file_path)
        destination_path = os.path.join(os.getcwd(), vamas_filename)
        shutil.copy2(file_path, destination_path)

        # Read VAMAS file
        vamas_data = Vamas(vamas_filename)

        # Create new Excel workbook
        wb = Workbook()
        wb.remove(wb.active)

        exp_data = []  # Store experimental description data

        # Process each block in the VAMAS file
        for i, block in enumerate(vamas_data.blocks, start=1):
            # # Determine sheet name
            # if block.species_label.lower() == "wide" or block.transition_or_charge_state_label.lower() == "none":
            #     sheet_name = block.species_label
            # else:
            #     sheet_name = f"{block.species_label}{block.transition_or_charge_state_label}"
            # sheet_name = sheet_name.replace("/", "_")
            #
            # # Create new sheet
            # ws = wb.create_sheet(title=sheet_name)

            # Determine original sheet name
            if block.species_label.lower() == "wide" or block.transition_or_charge_state_label.lower() == "none":
                raw_sheet_name = block.species_label
            else:
                raw_sheet_name = f"{block.species_label}{block.transition_or_charge_state_label}"
            raw_sheet_name = raw_sheet_name.replace("/", "_")

            # Normalize the sheet name
            sheet_name = normalize_sheet_name(raw_sheet_name)

            # Handle duplicate sheet names
            if sheet_name in wb.sheetnames:
                # Find a unique name by appending a number
                count = 1
                while f"{sheet_name}{count}" in wb.sheetnames:
                    count += 1
                sheet_name = f"{sheet_name}{count}"

            # Create new sheet with normalized name
            ws = wb.create_sheet(title=sheet_name)

            # Extract and process data
            num_points = block.num_y_values
            x_start = block.x_start
            x_step = block.x_step
            x_values = [x_start + i * x_step for i in range(num_points)]
            y_values = block.corresponding_variables[0].y_values
            y_unit = block.corresponding_variables[0].unit
            num_scans = block.num_scans_to_compile_block

            # Convert counts to counts per second if necessary
            if y_unit != "c/s":
                y_values = [y / num_scans for y in y_values]

            # Convert to Binding Energy if necessary
            if block.x_label.lower() in ["kinetic energy", "ke"]:
                x_values = [window.photons - x - window.workfunction for x in x_values]
                x_label = "Binding Energy"
            else:
                x_label = block.x_label

            # Write data to sheet
            ws.append([x_label, "Corrected Data", "Raw Data", "Transmission"])

            # Get transmission data if it exists
            transmission_data = None
            if hasattr(block, 'corresponding_variables') and len(block.corresponding_variables) > 1:
                transmission_data = block.corresponding_variables[1].y_values
            else:
                transmission_data = [1.0] * len(y_values)

            # Write data row by row
            for i, (x, y) in enumerate(zip(x_values, y_values)):
                trans = transmission_data[i]
                corrected_y = y / trans
                ws.append([x, corrected_y, y, trans])

            # Store experimental setup data
            block_exp_data = [
                f"Block {i}",
                block.sample_identifier,
                f"{block.year}/{block.month}/{block.day}",
                f"{block.hour}:{block.minute}:{block.second}",
                block.technique,
                f"{block.species_label} {block.transition_or_charge_state_label}",
                block.num_scans_to_compile_block,
                block.analysis_source_label,
                block.analysis_source_characteristic_energy,
                block.analysis_source_beam_width_x,
                block.analysis_source_beam_width_y,
                block.analyzer_pass_energy_or_retard_ratio_or_mass_res,
                block.analyzer_work_function_or_acceptance_energy,
                block.analyzer_mode,
                block.sputtering_source_energy if hasattr(block, 'sputtering_source_energy') else 'N/A',
                block.analyzer_axis_take_off_polar_angle,
                block.analyzer_axis_take_off_azimuth,
                block.target_bias,
                block.analysis_width_x,
                block.analysis_width_y,
                block.x_label,
                block.x_units,
                block.x_start,
                block.x_step,
                block.num_y_values,
                block.num_scans_to_compile_block,
                block.signal_collection_time,
                block.signal_time_correction,
                y_unit,
                block.num_lines_block_comment,
                block.block_comment
            ]
            exp_data.append(block_exp_data)

            # Add experimental description data to this sheet starting at column 50
            # This is safely beyond the peak parameters grid which typically ends around column 40-45
            exp_col = 50
            ws.cell(row=1, column=exp_col, value="Experimental Description")

            exp_labels = [
                "Sample ID", "Date", "Time", "Technique", "Species & Transition", "Number of scans",
                "Source Label", "Source Energy", "Source width X", "Source width Y", "Pass Energy", "Work Function",
                "Analyzer Mode", "Sputtering Energy", "Take-off Polar Angle", "Take-off Azimuth", "Target Bias",
                "Analysis Width X", "Analysis Width Y", "X Label", "X Units", "X Start", "X Step", "Num Y Values",
                "Num Scans", "Collection Time", "Time Correction", "Y Unit", "# Comment Lines", "Block Comment"
            ]

            for j, (label, value) in enumerate(zip(exp_labels, block_exp_data[1:])):
                ws.cell(row=j + 2, column=exp_col, value=label)
                ws.cell(row=j + 2, column=exp_col + 1, value=value)

            # Set column width for experimental data
            ws.column_dimensions[chr(64 + exp_col)].width = 25
            ws.column_dimensions[chr(64 + exp_col + 1)].width = 40

        # Create "Experimental description" sheet (keep this for backward compatibility)
        exp_sheet = wb.create_sheet(title="Experimental description")
        exp_sheet.column_dimensions['A'].width = 50
        exp_sheet.column_dimensions['B'].width = 100
        left_aligned = Alignment(horizontal='left')

        # Add VAMAS header information
        exp_sheet.append(["VAMAS Header Information"])
        for item in [
            ("Format Identifier", vamas_data.header.format_identifier),
            ("Institution Identifier", vamas_data.header.institution_identifier),
            ("Instrument Model", vamas_data.header.instrument_model_identifier),
            ("Operator Identifier", vamas_data.header.operator_identifier),
            ("Experiment Identifier", vamas_data.header.experiment_identifier),
            ("Number of Comment Lines", vamas_data.header.num_lines_comment),
            ("Comment", vamas_data.header.comment),
            ("Experiment Mode", vamas_data.header.experiment_mode),
            ("Scan Mode", vamas_data.header.scan_mode),
            ("Number of Spectral Regions", vamas_data.header.num_spectral_regions),
            ("Number of Analysis Positions", vamas_data.header.num_analysis_positions),
            ("Number of Discrete X Coordinates", vamas_data.header.num_discrete_x_coords_in_full_map),
            ("Number of Discrete Y Coordinates", vamas_data.header.num_discrete_y_coords_in_full_map)
        ]:
            exp_sheet.append(item)

        exp_sheet.append([])  # Add a blank row for separation

        # Define the order of block information
        block_info_order = [
            "Sample ID", "Year/Month/Day", "Time HH,MM,SS", "Technique", "Species & Transition", "Number of scans",
            "Source Label", "Source Energy", "Source width X", "Source width Y", "Pass Energy", "Work Function",
            "Analyzer Mode", "Sputtering Energy", "Take-off Polar Angle", "Take-off Azimuth", "Target Bias",
            "Analysis Width X", "Analysis Width Y", "X Label", "X Units", "X Start", "X Step", "Num Y Values",
            "Num Scans", "Collection Time", "Time Correction", "Y Unit", "# Comment Lines", "Block Comment"
        ]

        # Add block information
        for i, block_data in enumerate(exp_data, start=1):
            exp_sheet.append([f"Block {i}", ""])
            for j, info in enumerate(block_info_order):
                exp_sheet.append([info, block_data[j + 1]])
            exp_sheet.append([])  # Add a blank row between blocks

        # Set alignment for all cells in column B
        for row in exp_sheet.iter_rows(min_row=1, max_row=exp_sheet.max_row, min_col=2, max_col=2):
            for cell in row:
                cell.alignment = left_aligned

        # Save Excel file
        excel_filename = os.path.splitext(vamas_filename)[0] + ".xlsx"
        excel_path = os.path.join(os.path.dirname(file_path), excel_filename)
        wb.save(excel_path)

        # Remove temporary VAMAS file
        os.remove(destination_path)

        # Update window.Data with the new Excel file
        window.Data = Init_Measurement_Data(window)
        window.Data['FilePath'] = excel_path

        # Open the Excel file and populate window.Data
        open_xlsx_file_vamas(window, excel_path)

    except FileNotFoundError as e:
        wx.MessageBox(f"File not found: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
    except Exception as e:
        wx.MessageBox(f"Error processing VAMAS file: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

def open_xlsx_file_vamas(window, file_path):
    """
    Open and process an Excel file created from a VAMAS file.

    This function initializes the data structure, reads the Excel file,
    populates the window.Data dictionary with core level information,
    updates the GUI elements, and plots the data for the first sheet.

    Args:
    window: The main application window object.
    file_path: The path to the Excel file to be opened.
    """
    try:
        # Update status bar with the selected file path
        window.SetStatusText(f"Selected File: {file_path}", 0)

        # Initialize the measurement data structure
        window.Data = Init_Measurement_Data(window)
        window.Data['FilePath'] = file_path

        # Read the Excel file
        excel_file = pd.ExcelFile(file_path)

        # Get sheet names, excluding the "Experimental description" sheet
        sheet_names = [name for name in excel_file.sheet_names if name != "Experimental description"]

        # Initialize the number of core levels
        window.Data['Number of Core levels'] = 0

        # Process each sheet (core level) in the Excel file
        for sheet_name in sheet_names:
            window.Data = add_core_level_Data(window.Data, window, file_path, sheet_name)

        print(f"Final number of core levels: {window.Data['Number of Core levels']}")

        # Update GUI elements
        window.sheet_combobox.Clear()
        window.sheet_combobox.AppendItems(sheet_names)
        window.sheet_combobox.SetValue(sheet_names[0])  # Set first sheet as default

        # Plot the data for the first sheet
        window.plot_manager.plot_data(window)

    except Exception as e:
        # Handle any errors that occur during file processing
        print(f"Error in open_xlsx_file_vamas: {str(e)}")
        import traceback
        traceback.print_exc()
        wx.MessageBox(f"Error reading Excel file: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)


def extract_transmission_data(block):
    """
    Extract transmission function data from a Kratos .kal file block.

    Args:
        block (str): Text block containing transmission data

    Returns:
        tuple: (ke_trans, trans_values) arrays of kinetic energies and transmission values
    """
    if 'Transmission Function Object (ke,t)' in block:
        trans_section = block.split('Transmission Function Object (ke,t)')[1].split('Dataset filename')[0]

        ke_line = trans_section.split('Transmission Function Kinetic Energy =')[1].split('\n')[0]
        ke_str = ke_line.strip().strip('{}').strip()
        ke_trans = np.array([float(x.strip()) for x in ke_str.split(',')])

        val_line = trans_section.split('Transmission Function Value          =')[1].split('\n')[0]
        val_str = val_line.strip().strip('{}').strip()
        trans_values = np.array([float(x.strip()) for x in val_str.split(',')])

        return ke_trans, trans_values
    return None, None


def convert_kal_to_excel_OLD(file_path):
    """
    Convert Kratos .kal file to Excel format suitable for KherveFitting.

    Args:
        file_path (str): Path to the .kal file

    Returns:
        str: Path to the created Excel file
    """
    with open(file_path, 'r') as f:

        content = f.read()

    blocks = content.split('Dataset filename')
    spectra = {}
    PHOTON_ENERGY = 1486.67

    for block in blocks:
        if 'Ordinate values' in block and 'Object name' in block:
            name = block.split('Object name               = ')[1].split('/')[0].strip()

            ke_trans, trans_values = extract_transmission_data(block)
            if ke_trans is not None and trans_values is not None:
                trans_func = interp1d(ke_trans, trans_values, kind='linear', bounds_error=False,
                                      fill_value='extrapolate')

                start_ke = float(block.split('Spectrum scan start')[1].split('=')[1].split('eV')[0].strip())
                step = float(block.split('Spectrum scan step size')[1].split('=')[1].split('eV')[0].strip())
                raw_str = block.split('Ordinate values')[1].split('=')[1].split('}')[0].strip().strip('{').strip()
                raw_data = np.array([float(x.strip()) for x in raw_str.split(',')])

                num_points = len(raw_data)
                ke_values = np.linspace(start_ke, start_ke + (num_points - 1) * step, num_points)
                be_values = PHOTON_ENERGY - ke_values

                transmission = trans_func(ke_values)
                corrected_data = raw_data / transmission

                df = pd.DataFrame({
                    'BE': be_values,
                    'Corrected Data': corrected_data,
                    'Raw Data': raw_data,
                    'Transmission': transmission
                })

                spectra[name] = df

    output_file = file_path.replace('.kal', '.xlsx')
    with pd.ExcelWriter(output_file) as writer:
        for name, df in spectra.items():
            sheet_name = name.replace(' ', '')
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    return output_file


def convert_kal_to_excel(file_path):
    """
    Convert Kratos .kal file to Excel format suitable for KherveFitting.

    Args:
        file_path (str): Path to the .kal file

    Returns:
        str: Path to the created Excel file
    """
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Alignment

    with open(file_path, 'r') as f:
        content = f.read()

    blocks = content.split('Dataset filename')
    spectra = {}
    PHOTON_ENERGY = 1486.67

    # Extract sample ID
    sample_id = "Unknown"
    for block in blocks:
        for line in block.split('\n'):
            if 'Stage Position Name' in line:
                sample_id = line.split('=')[1].strip()
                break
        if sample_id != "Unknown":
            break

    # Process each spectrum block
    for block in blocks:
        if 'Ordinate values' not in block or 'Object name' not in block:
            continue

        # Get spectrum name and number
        object_name_line = next((line for line in block.split('\n') if 'Object name' in line), None)
        if not object_name_line:
            continue

        name_parts = object_name_line.split('=')[1].strip().split('/')
        name = name_parts[0].strip()
        spectrum_id = name_parts[1].strip() if len(name_parts) > 1 else ""

        # Skip non-spectrum blocks
        if 'Sample Position' in name or 'Counter' in name:
            continue

        # Initialize metadata dictionary with default values
        metadata = {
            'Sample ID': sample_id,
            'Date': '',
            'Time': '',
            'Technique': 'XPS',
            'Species & Transition': '',
            'Number of scans': '',
            'Source Label': 'Al mono',
            'Source Energy': '1486.6',
            'Source width X': '1E+37',
            'Source width Y': '1E+37',
            'Pass Energy': '',
            'Work Function': '1E+37',
            'Analyzer Mode': 'FAT',
            'Sputtering Energy': 'N/A',
            'Take-off Polar Angle': '1E+37',
            'Take-off Azimuth': '1E+37',
            'Target Bias': '1E+37',
            'Analysis Width X': '1E+37',
            'Analysis Width Y': '1E+37',
            'X Label': '',
            'X Units': '',
            'X Start': '',
            'X Step': '',
            'Num Y Values': '',
            'Num Scans': '',
            'Collection Time': '',
            'Time Correction': '1E+37',
            'Y Unit': 'd',
            '# Comment Lines': '0',
            'Block Comment': ''
        }

        # Extract metadata from block
        for line in block.split('\n'):
            line = line.strip()

            # Date and Time
            if 'Date Acquired' in line:
                parts = line.split('=')[1].strip().split()
                if len(parts) >= 2:
                    metadata['Date'] = parts[0]
                    metadata['Time'] = parts[1]

            # Species & Transition
            elif 'Chemical symbol or formula' in line:
                metadata['species'] = line.split('=')[1].strip()
            elif 'Transition or charge state' in line:
                metadata['transition'] = line.split('=')[1].strip()

            # Number of scans
            elif '# Sweeps completed' in line:
                num_scans = line.split('=')[1].strip()
                metadata['Number of scans'] = num_scans
                metadata['Num Scans'] = num_scans

            # Pass Energy
            elif 'Pass energy' in line:
                metadata['Pass Energy'] = line.split('=')[1].strip()

            # X Label, Units, Start, Step
            elif 'Abscissa label' in line:
                metadata['X Label'] = line.split('=')[1].strip()
            elif 'Abscissa units' in line:
                metadata['X Units'] = line.split('=')[1].strip()
            elif 'Spectrum scan start' in line:
                metadata['X Start'] = line.split('=')[1].split('eV')[0].strip()
            elif 'Spectrum scan step size' in line:
                metadata['X Step'] = line.split('=')[1].split('eV')[0].strip()

            # Collection Time
            elif 'Dwell time' in line:
                metadata['Collection Time'] = line.split('=')[1].replace('seconds', '').strip()

        # Combine species and transition
        if 'species' in metadata and 'transition' in metadata:
            metadata['Species & Transition'] = f"{metadata['species']} {metadata['transition']}"
            metadata.pop('species')
            metadata.pop('transition')

        # Count Y values
        if 'Ordinate values' in block:
            try:
                values_str = block.split('Ordinate values')[1].split('=')[1].split('}')[0].strip().strip('{').strip()
                values = values_str.split(',')
                metadata['Num Y Values'] = str(len(values))
            except:
                pass

        # Extract block comment
        comment_lines = []
        in_comment = False
        for line in block.split('\n'):
            if 'Casa Info Follows' in line:
                in_comment = True
                comment_lines.append(line.strip())
            elif in_comment and line.strip():
                comment_lines.append(line.strip())

        if comment_lines:
            metadata['# Comment Lines'] = str(len(comment_lines))
            metadata['Block Comment'] = '"' + '\n'.join(comment_lines) + '"'

        # Process spectral data
        ke_trans, trans_values = extract_transmission_data(block)
        if ke_trans is not None and trans_values is not None:
            # Create interpolation function
            trans_func = interp1d(ke_trans, trans_values, kind='linear', bounds_error=False,
                                  fill_value='extrapolate')

            start_ke = float(metadata['X Start'])
            step = float(metadata['X Step'])
            raw_str = block.split('Ordinate values')[1].split('=')[1].split('}')[0].strip().strip('{').strip()
            raw_data = np.array([float(x.strip()) for x in raw_str.split(',')])

            num_points = len(raw_data)
            ke_values = np.linspace(start_ke, start_ke + (num_points - 1) * step, num_points)
            be_values = PHOTON_ENERGY - ke_values

            transmission = trans_func(ke_values)
            corrected_data = raw_data / transmission

            df = pd.DataFrame({
                'BE': be_values,
                'Corrected Data': corrected_data,
                'Raw Data': raw_data,
                'Transmission': transmission
            })

            spectra[name] = {
                'data': df,
                'metadata': metadata,
                'id': spectrum_id
            }

    # Create Excel workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Create sheets for each spectrum
    for name, info in spectra.items():
        sheet_name = name.replace(' ', '')
        df = info['data']
        metadata = info['metadata']

        # Create sheet and write data
        ws = wb.create_sheet(sheet_name)

        # Write headers
        for col_idx, col_name in enumerate(df.columns, 1):
            ws.cell(row=1, column=col_idx, value=col_name)

        # Write data values
        for row_idx, row_data in enumerate(df.values, 2):
            for col_idx, value in enumerate(row_data, 1):
                ws.cell(row=row_idx, column=col_idx, value=value)

        # Add experimental description (column 50)
        exp_col = 50
        ws.cell(row=1, column=exp_col, value="Experimental Description")

        # Metadata fields in specified order
        fields = [
            'Sample ID', 'Date', 'Time', 'Technique', 'Species & Transition',
            'Number of scans', 'Source Label', 'Source Energy', 'Source width X',
            'Source width Y', 'Pass Energy', 'Work Function', 'Analyzer Mode',
            'Sputtering Energy', 'Take-off Polar Angle', 'Take-off Azimuth',
            'Target Bias', 'Analysis Width X', 'Analysis Width Y', 'X Label',
            'X Units', 'X Start', 'X Step', 'Num Y Values', 'Num Scans',
            'Collection Time', 'Time Correction', 'Y Unit', '# Comment Lines',
            'Block Comment'
        ]

        # Write metadata in order
        for i, field in enumerate(fields, 2):
            ws.cell(row=i, column=exp_col, value=field)
            ws.cell(row=i, column=exp_col + 1, value=metadata.get(field, ''))

        # Set column widths
        ws.column_dimensions[get_column_letter(exp_col)].width = 25
        ws.column_dimensions[get_column_letter(exp_col + 1)].width = 50

    # Create Experimental description sheet
    exp_sheet = wb.create_sheet("Experimental description")
    exp_sheet.column_dimensions['A'].width = 30
    exp_sheet.column_dimensions['B'].width = 100
    left_aligned = Alignment(horizontal='left')

    # Write metadata to Experimental description sheet
    row = 1
    for name, info in spectra.items():
        metadata = info['metadata']
        spectrum_id = info['id']

        # Add spectrum header
        header = exp_sheet.cell(row=row, column=1, value=f"Spectrum: {name}/{spectrum_id}")
        header.alignment = left_aligned
        row += 1

        # Write metadata fields
        for field in fields:
            cell_a = exp_sheet.cell(row=row, column=1, value=field)
            cell_b = exp_sheet.cell(row=row, column=2, value=metadata.get(field, ''))

            cell_a.alignment = left_aligned
            cell_b.alignment = left_aligned

            row += 1

        # Add space between spectra
        row += 1

    # Save Excel file
    output_file = file_path.replace('.kal', '.xlsx')
    wb.save(output_file)

    return output_file

def open_kal_file_dialog(window):
    """Open a file dialog for selecting a Kratos .kal file"""
    with wx.FileDialog(window, "Open Kratos file", wildcard="Kratos files (*.kal)|*.kal",
                      style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
        if fileDialog.ShowModal() == wx.ID_CANCEL:
            return
        file_path = fileDialog.GetPath()
        open_kal_file(window, file_path)

def open_kal_file(window, file_path):
    """Process a Kratos .kal file and convert it to Excel format"""

    try:
        output_excel = convert_kal_to_excel(file_path)
        open_xlsx_file(window, output_excel)
    except Exception as e:
        wx.MessageBox(f"Error processing Kratos file: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)


def import_raman_txt_file(window):
    import wx
    import os
    import pandas as pd
    import openpyxl
    from libraries.Open import open_xlsx_file

    with wx.FileDialog(window, "Open Raman text file", wildcard="Text files (*.txt)|*.txt",
                       style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
        if fileDialog.ShowModal() == wx.ID_CANCEL:
            return
        file_path = fileDialog.GetPath()

    try:
        # Get filename without extension for sheet name
        base_filename = os.path.basename(file_path).split('.')[0]
        sheet_name = f"Raman_{base_filename}"

        # Read data from txt file
        data = []
        with open(file_path, 'r') as f:
            for line in f:
                if line.strip() and not line.startswith('#'):
                    parts = line.strip().split()
                    if len(parts) >= 2:
                        try:
                            wavenumber = float(parts[0])
                            intensity = float(parts[1])
                            data.append([wavenumber, intensity])
                        except ValueError:
                            continue

        if not data:
            window.show_popup_message2("Error", "No valid data found in the file.")
            return

        # Create new Excel file if needed
        output_dir = os.path.dirname(file_path)
        excel_path = os.path.join(output_dir, f"{base_filename}.xlsx")

        if os.path.exists(excel_path):
            # If Excel file exists, load it
            wb = openpyxl.load_workbook(excel_path)
        else:
            # Create new Excel file
            wb = openpyxl.Workbook()
            # Remove default sheet
            if "Sheet" in wb.sheetnames:
                wb.remove(wb["Sheet"])

        # Create or replace the sheet
        if sheet_name in wb.sheetnames:
            wb.remove(wb[sheet_name])
        ws = wb.create_sheet(sheet_name)

        # Add headers
        ws["A1"] = "Wavenumber (cm-1)"
        ws["B1"] = "Raw Data"

        # Add data
        for i, (wavenumber, intensity) in enumerate(data, start=2):
            ws[f"A{i}"] = wavenumber
            ws[f"B{i}"] = intensity

        # Save Excel file
        wb.save(excel_path)

        # Open the Excel file in Khervefitting
        open_xlsx_file(window, excel_path)

    except Exception as e:
        import traceback
        traceback.print_exc()
        window.show_popup_message2("Error", f"Error processing file: {str(e)}")


def import_multiple_raman_files(window):
    import wx
    import os
    import pandas as pd
    import openpyxl
    from libraries.Open import open_xlsx_file

    with wx.DirDialog(window, "Choose a directory containing Raman text files",
                      style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST) as dirDialog:
        if dirDialog.ShowModal() == wx.ID_CANCEL:
            return
        dir_path = dirDialog.GetPath()

    try:
        # Find all txt files in the directory
        txt_files = [f for f in os.listdir(dir_path) if f.lower().endswith('.txt')]

        if not txt_files:
            window.show_popup_message2("Information", "No .txt files found in the selected folder.")
            return

        # Create an Excel file for each group of files
        excel_path = os.path.join(dir_path, "Raman_Data.xlsx")
        wb = openpyxl.Workbook()

        # Remove default sheet
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

        # Process each txt file
        for txt_file in txt_files:
            file_path = os.path.join(dir_path, txt_file)
            base_filename = os.path.splitext(txt_file)[0]
            sheet_name = f"Raman_{base_filename}"

            # Ensure sheet name is valid (max 31 chars)
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]

            # Handle duplicate sheet names
            count = 1
            original_name = sheet_name
            while sheet_name in wb.sheetnames:
                sheet_name = f"{original_name[:27]}_{count}"
                count += 1

            # Read data from txt file
            data = []
            with open(file_path, 'r') as f:
                for line in f:
                    if line.strip() and not line.startswith('#'):
                        parts = line.strip().split()
                        if len(parts) >= 2:
                            try:
                                wavenumber = float(parts[0])
                                intensity = float(parts[1])
                                data.append([wavenumber, intensity])
                            except ValueError:
                                continue

            if not data:
                continue  # Skip files with no valid data

            # Create sheet
            ws = wb.create_sheet(sheet_name)

            # Add headers
            ws["A1"] = "Wavenumber (cm-1)"
            ws["B1"] = "Raw Data"

            # Add data
            for i, (wavenumber, intensity) in enumerate(data, start=2):
                ws[f"A{i}"] = wavenumber
                ws[f"B{i}"] = intensity

        # Save Excel file if sheets were created
        if len(wb.sheetnames) > 0:
            wb.save(excel_path)
            open_xlsx_file(window, excel_path)
            window.show_popup_message2("Success", f"Created Excel file with {len(wb.sheetnames)} Raman data sheets.")
        else:
            window.show_popup_message2("Information", "No valid data found in any of the text files.")

    except Exception as e:
        import traceback
        traceback.print_exc()
        window.show_popup_message2("Error", f"Error processing files: {str(e)}")

def import_xps_asc_file(window):
    import wx
    import os
    import pandas as pd
    import openpyxl
    from libraries.Open import open_xlsx_file

    with wx.FileDialog(window, "Open XPS ASC file", wildcard="ASC files (*.asc)|*.asc",
                       style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
        if fileDialog.ShowModal() == wx.ID_CANCEL:
            return
        file_path = fileDialog.GetPath()

    try:
        # Get filename without extension for sheet name
        base_filename = os.path.basename(file_path).split('.')[0]
        sheet_name = base_filename

        # Read data from asc file
        data = []
        with open(file_path, 'r') as f:
            for line in f:
                if line.strip() and not line.startswith('#'):
                    parts = line.strip().split()
                    if len(parts) >= 2:
                        try:
                            binding_energy = float(parts[0])
                            intensity = float(parts[1])
                            data.append([binding_energy, intensity])
                        except ValueError:
                            continue

        if not data:
            window.show_popup_message2("Error", "No valid data found in the file.")
            return

        # Create new Excel file if needed
        output_dir = os.path.dirname(file_path)
        excel_path = os.path.join(output_dir, f"{base_filename}.xlsx")

        if os.path.exists(excel_path):
            # If Excel file exists, load it
            wb = openpyxl.load_workbook(excel_path)
        else:
            # Create new Excel file
            wb = openpyxl.Workbook()
            # Remove default sheet
            if "Sheet" in wb.sheetnames:
                wb.remove(wb["Sheet"])

        # Create or replace the sheet
        if sheet_name in wb.sheetnames:
            wb.remove(wb[sheet_name])
        ws = wb.create_sheet(sheet_name)

        # Add headers
        ws["A1"] = "BE"
        ws["B1"] = "Raw Data"

        # Add data
        for i, (binding_energy, intensity) in enumerate(data, start=2):
            ws[f"A{i}"] = binding_energy
            ws[f"B{i}"] = intensity

        # Save Excel file
        wb.save(excel_path)

        # Open the Excel file in Khervefitting
        open_xlsx_file(window, excel_path)

    except Exception as e:
        import traceback
        traceback.print_exc()
        window.show_popup_message2("Error", f"Error processing file: {str(e)}")


def import_multiple_xps_asc_files(window):
    import wx
    import os
    import pandas as pd
    import openpyxl
    from libraries.Open import open_xlsx_file

    with wx.DirDialog(window, "Choose a directory containing XPS ASC files",
                      style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST) as dirDialog:
        if dirDialog.ShowModal() == wx.ID_CANCEL:
            return
        dir_path = dirDialog.GetPath()

    try:
        # Find all asc files in the directory
        asc_files = [f for f in os.listdir(dir_path) if f.lower().endswith('.asc')]

        if not asc_files:
            window.show_popup_message2("Information", "No .asc files found in the selected folder.")
            return

        # Create an Excel file for each group of files
        excel_path = os.path.join(dir_path, "XPS_Data.xlsx")
        wb = openpyxl.Workbook()

        # Remove default sheet
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

        # Process each asc file
        for asc_file in asc_files:
            file_path = os.path.join(dir_path, asc_file)
            base_filename = os.path.splitext(asc_file)[0]
            sheet_name = base_filename

            # Ensure sheet name is valid (max 31 chars)
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]

            # Handle duplicate sheet names
            count = 1
            original_name = sheet_name
            while sheet_name in wb.sheetnames:
                sheet_name = f"{original_name[:27]}_{count}"
                count += 1

            # Read data from asc file
            data = []
            with open(file_path, 'r') as f:
                for line in f:
                    if line.strip() and not line.startswith('#'):
                        parts = line.strip().split()
                        if len(parts) >= 2:
                            try:
                                binding_energy = float(parts[0])
                                intensity = float(parts[1])
                                data.append([binding_energy, intensity])
                            except ValueError:
                                continue

            if not data:
                continue  # Skip files with no valid data

            # Create sheet
            ws = wb.create_sheet(sheet_name)

            # Add headers
            ws["A1"] = "BE"
            ws["B1"] = "Raw Data"

            # Add data
            for i, (binding_energy, intensity) in enumerate(data, start=2):
                ws[f"A{i}"] = binding_energy
                ws[f"B{i}"] = intensity

        # Save Excel file if sheets were created
        if len(wb.sheetnames) > 0:
            wb.save(excel_path)
            open_xlsx_file(window, excel_path)
            window.show_popup_message2("Success", f"Created Excel file with {len(wb.sheetnames)} XPS data sheets.")
        else:
            window.show_popup_message2("Information", "No valid data found in any of the ASC files.")

    except Exception as e:
        import traceback
        traceback.print_exc()
        window.show_popup_message2("Error", f"Error processing files: {str(e)}")

def import_xps_csv_file(window):
    import wx
    import os
    import pandas as pd
    import openpyxl
    from libraries.Open import open_xlsx_file

    with wx.FileDialog(window, "Open XPS CSV file", wildcard="CSV files (*.csv)|*.csv",
                       style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
        if fileDialog.ShowModal() == wx.ID_CANCEL:
            return
        file_path = fileDialog.GetPath()

    try:
        # Get filename without extension for sheet name
        base_filename = os.path.basename(file_path).split('.')[0]
        sheet_name = base_filename

        # Read data from csv file
        df = pd.read_csv(file_path, delimiter=',', header=None)
        if df.shape[1] < 2:
            # Try other common separators if comma doesn't work
            for sep in [';', '\t', ' ']:
                test_df = pd.read_csv(file_path, delimiter=sep, header=None)
                if test_df.shape[1] >= 2:
                    df = test_df
                    break

        if df.shape[1] < 2:
            window.show_popup_message2("Error", "CSV file must contain at least two columns (BE and Intensity).")
            return

        # Extract the first two columns (assuming BE and intensity)
        data = df.iloc[:, :2].values

        if len(data) == 0:  # Use len() instead of direct boolean check
            window.show_popup_message2("Error", "No valid data found in the file.")
            return

        # Create new Excel file if needed
        output_dir = os.path.dirname(file_path)
        excel_path = os.path.join(output_dir, f"{base_filename}.xlsx")

        if os.path.exists(excel_path):
            # If Excel file exists, load it
            wb = openpyxl.load_workbook(excel_path)
        else:
            # Create new Excel file
            wb = openpyxl.Workbook()
            # Remove default sheet
            if "Sheet" in wb.sheetnames:
                wb.remove(wb["Sheet"])

        # Create or replace the sheet
        if sheet_name in wb.sheetnames:
            wb.remove(wb[sheet_name])
        ws = wb.create_sheet(sheet_name)

        # Add headers
        ws["A1"] = "Binding Energy (eV)"
        ws["B1"] = "Raw Data"

        # Add data
        for i, (binding_energy, intensity) in enumerate(data, start=2):
            ws[f"A{i}"] = float(binding_energy)
            ws[f"B{i}"] = float(intensity)

        # Save Excel file
        wb.save(excel_path)

        # Open the Excel file in Khervefitting
        open_xlsx_file(window, excel_path)

    except Exception as e:
        import traceback
        traceback.print_exc()
        window.show_popup_message2("Error", f"Error processing file: {str(e)}")

def import_multiple_xps_csv_files(window):
    import wx
    import os
    import pandas as pd
    import openpyxl
    from libraries.Open import open_xlsx_file

    with wx.DirDialog(window, "Choose a directory containing XPS CSV files",
                      style=wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST) as dirDialog:
        if dirDialog.ShowModal() == wx.ID_CANCEL:
            return
        dir_path = dirDialog.GetPath()

    try:
        # Find all csv files in the directory
        csv_files = [f for f in os.listdir(dir_path) if f.lower().endswith('.csv')]

        if not csv_files:
            window.show_popup_message2("Information", "No .csv files found in the selected folder.")
            return

        # Create an Excel file for all files
        excel_path = os.path.join(dir_path, "XPS_Data_CSV.xlsx")
        wb = openpyxl.Workbook()

        # Remove default sheet
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])

        # Process each csv file
        for csv_file in csv_files:
            file_path = os.path.join(dir_path, csv_file)
            base_filename = os.path.splitext(csv_file)[0]
            sheet_name = base_filename

            # Ensure sheet name is valid (max 31 chars)
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]

            # Handle duplicate sheet names
            count = 1
            original_name = sheet_name
            while sheet_name in wb.sheetnames:
                sheet_name = f"{original_name[:27]}_{count}"
                count += 1

            # Read data from csv file
            try:
                df = pd.read_csv(file_path, delimiter=',', header=None)
                if df.shape[1] < 2:
                    # Try other common separators if comma doesn't work
                    for sep in [';', '\t', ' ']:
                        test_df = pd.read_csv(file_path, delimiter=sep, header=None)
                        if test_df.shape[1] >= 2:
                            df = test_df
                            break

                if df.shape[1] < 2:
                    continue  # Skip files without at least two columns

                # Extract the first two columns (assuming BE and intensity)
                data = df.iloc[:, :2].values

                if len(data) == 0:  # Use len() instead of direct boolean check
                    continue  # Skip files with no valid data
            except Exception as e:
                print(f"Error reading {csv_file}: {e}")
                continue

            # Create sheet
            ws = wb.create_sheet(sheet_name)

            # Add headers
            ws["A1"] = "Binding Energy (eV)"
            ws["B1"] = "Raw Data"

            # Add data
            for i, (binding_energy, intensity) in enumerate(data, start=2):
                ws[f"A{i}"] = float(binding_energy)
                ws[f"B{i}"] = float(intensity)

        # Save Excel file if sheets were created
        if len(wb.sheetnames) > 0:
            wb.save(excel_path)
            open_xlsx_file(window, excel_path)
            window.show_popup_message2("Success", f"Created Excel file with {len(wb.sheetnames)} XPS data sheets.")
        else:
            window.show_popup_message2("Information", "No valid data found in any of the CSV files.")

    except Exception as e:
        import traceback
        traceback.print_exc()
        window.show_popup_message2("Error", f"Error processing files: {str(e)}")

def open_file_location(window):
    if 'FilePath' in window.Data:
        file_path = window.Data['FilePath']
        folder_path = os.path.dirname(file_path)
        if os.path.exists(folder_path):
            if sys.platform == 'win32':
                os.startfile(folder_path)
            elif sys.platform == 'darwin':
                os.system(f'open "{folder_path}"')
            else:
                os.system(f'xdg-open "{folder_path}"')

# ------------------ HISTORRY DEF ---------------------------------------------------
# -----------------------------------------------------------------------------------
