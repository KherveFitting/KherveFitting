# KherveFitting - XPS Data Analysis Software
# Copyright (C) 2024-2026 Gwilherm Kerherve <g.kerherve@ic.ac.uk>
#
# KherveFitting is dual-licensed:
#   - GNU GPL v3.0 (see LICENSE-GPL.txt) for open-source use
#   - Commercial Licence (see LICENSE-COMMERCIAL.txt) for proprietary use
# SPDX-License-Identifier: GPL-3.0-only OR LicenseRef-KherveFitting-Commercial

import re
import wx


def on_be_correction_change(window, event):
    new_correction = window.be_correction_spinbox.GetValue()

    # Save to BEcorrection for backward compatibility
    window.Data['BEcorrection'] = new_correction

    # Extract row number from sheet name
    sheet_name = window.sheet_combobox.GetValue()
    sample_row = None

    # Extract the numeric suffix from the sheet name
    match = re.search(r'(\d+)$', sheet_name)
    if match:
        sample_row = match.group(1)
    else:
        # If no numeric suffix, assume it's row 0
        sample_row = "0"

    # Update BEcorrections
    if 'BEcorrections' not in window.Data:
        window.Data['BEcorrections'] = {}

    window.Data['BEcorrections'][sample_row] = new_correction

    # Update file manager grid if it's open
    if hasattr(window, 'file_manager') and window.file_manager is not None:
        try:
            grid = window.file_manager.grid
            if grid and grid.IsShown():
                be_col = len(window.file_manager.core_levels) + 1
                try:
                    row = int(sample_row)
                    if row < grid.GetNumberRows():
                        grid.SetCellValue(row, be_col, str(new_correction))
                except (ValueError, IndexError):
                    pass
        except (RuntimeError, wx.PyDeadObjectError):
            pass

    # Apply correction to the current sheet
    apply_be_correction(window, new_correction)


def on_auto_be(window, event):
    import wx
    c1s_correction = calculate_c1s_correction(window)
    if c1s_correction is not None:
        # Update the spinbox value
        window.be_correction_spinbox.SetValue(c1s_correction)

        # Create and post a spin control event to trigger the change handler
        spin_event = wx.SpinDoubleEvent(wx.EVT_SPINCTRLDOUBLE.typeId, window.be_correction_spinbox.GetId())
        spin_event.SetEventObject(window.be_correction_spinbox)
        spin_event.SetValue(c1s_correction)
        wx.PostEvent(window.be_correction_spinbox, spin_event)
    else:
        # No reference peak found - don't change anything
        wx.MessageBox(f"Reference peak '{window.ref_peak_name}' not found in the current row's sheets.",
                      "No Reference Peak", wx.OK | wx.ICON_INFORMATION)


def apply_be_correction(window, correction):
    """
    Apply binding energy correction to all data from the same sample row.

    Args:
        correction (float): New BE correction value in eV
    """
    from libraries.Sheet_Operations import on_sheet_selected

    delta_correction = correction - window.be_correction
    window.be_correction = correction
    window.Data['BEcorrection'] = correction

    # Get the current sheet and extract its row number
    current_sheet = window.sheet_combobox.GetValue()
    current_row = None

    # Extract row number from current sheet name using regex
    match = re.search(r'(\d+)$', current_sheet)
    if match:
        current_row = match.group(1)
    else:
        current_row = "0"  # Default to row 0 if no number found

    # Update BEcorrections for this sample
    if 'BEcorrections' not in window.Data:
        window.Data['BEcorrections'] = {}
    window.Data['BEcorrections'][current_row] = correction

    # Process all sheets to find and update ones from the same row
    for sheet_name, sheet_data in window.Data['Core levels'].items():
        # Check if this sheet belongs to the same row (no number suffix)
        sheet_match = re.search(r'(\d+)$', sheet_name)

        # Only process sheets that either:
        # 1. Have no number suffix (base core levels), or
        # 2. Have the same row number as current_row
        sheet_row = sheet_match.group(1) if sheet_match else "0"

        if sheet_row == current_row:
            # Update B.E. values in memory
            sheet_data['B.E.'] = [be + delta_correction for be in sheet_data['B.E.']]

            # Update Background range
            if 'Background' in sheet_data:
                bg = sheet_data['Background']
                if 'Bkg Low' in bg and bg['Bkg Low'] != '':
                    bg['Bkg Low'] = round(float(bg['Bkg Low']) + delta_correction, 2)
                if 'Bkg High' in bg and bg['Bkg High'] != '':
                    bg['Bkg High'] = round(float(bg['Bkg High']) + delta_correction, 2)

                # Shift the BE x-axis stored alongside the background so it stays
                # aligned with the corrected B.E. array.
                if 'Bkg X' in bg and bg['Bkg X']:
                    bg['Bkg X'] = [x + delta_correction for x in bg['Bkg X']]

                # Each recorded range tuple is (offset_h, offset_l, min_BE, max_BE).
                # Only the last two entries are binding energies and need shifting.
                if 'Recorded_Ranges' in bg and bg['Recorded_Ranges']:
                    shifted = []
                    for r in bg['Recorded_Ranges']:
                        try:
                            shifted.append((
                                float(r[0]),
                                float(r[1]),
                                float(r[2]) + delta_correction,
                                float(r[3]) + delta_correction,
                            ))
                        except (TypeError, ValueError, IndexError):
                            shifted.append(r)
                    bg['Recorded_Ranges'] = shifted

            # Update peak positions
            if 'Fitting' in sheet_data and 'Peaks' in sheet_data['Fitting']:
                for peak in sheet_data['Fitting']['Peaks'].values():
                    peak['Position'] += delta_correction
                    # Each peak also stores its own Bkg Low / Bkg High snapshot
                    # that the peak_params_grid reads from. Shift those too,
                    # otherwise on_sheet_selected re-populates the grid with
                    # the un-corrected values and the user sees no change.
                    for k in ('Bkg Low', 'Bkg High'):
                        if k in peak and peak[k] not in ('', None):
                            try:
                                peak[k] = round(float(peak[k]) + delta_correction, 2)
                            except (TypeError, ValueError):
                                pass
                    if 'Constraints' in peak:
                        pos_constraint = peak['Constraints'].get('Position', '')
                        if pos_constraint and ',' in pos_constraint and not any(
                                c in pos_constraint for c in 'ABCDEFGHIJKLMNOP'):
                            min_val, max_val = map(float, pos_constraint.split(','))
                            peak['Constraints'][
                                'Position'] = f"{min_val + delta_correction:.2f},{max_val + delta_correction:.2f}"

            # Update plot limits
            if sheet_name in window.plot_config.plot_limits:
                limits = window.plot_config.plot_limits[sheet_name]
                limits['Xmin'] += delta_correction
                limits['Xmax'] += delta_correction

    # Update peak_params_grid with corrected positions (col 2) and the
    # per-peak Bkg Low / Bkg High columns (15, 16) shown in the grid.
    if current_sheet == window.sheet_combobox.GetValue():
        num_peaks = window.peak_params_grid.GetNumberRows() // 2
        for i in range(num_peaks):
            row = i * 2
            try:
                pos = float(window.peak_params_grid.GetCellValue(row, 2))
                window.peak_params_grid.SetCellValue(row, 2, f"{pos + delta_correction:.2f}")
            except ValueError:
                pass
            for col in (15, 16):
                try:
                    val = float(window.peak_params_grid.GetCellValue(row, col))
                    if val != 0:
                        window.peak_params_grid.SetCellValue(
                            row, col, f"{val + delta_correction:.2f}")
                except ValueError:
                    continue

    # Update Results grid if it corresponds to the current row
    for row in range(window.results_grid.GetNumberRows()):
        sheet_name = window.results_grid.GetCellValue(row, 21)  # Sheetname column
        match = re.search(r'(\d+)$', sheet_name)
        grid_row = match.group(1) if match else "0"

        if grid_row == current_row:
            try:
                pos = float(window.results_grid.GetCellValue(row, 1))
                window.results_grid.SetCellValue(row, 1, f"{pos + delta_correction:.2f}")
            except ValueError:
                continue

    # Update current sheet display
    on_sheet_selected(window, current_sheet)

    # If the Peak Fitting window is open, push the corrected Bkg Low/High values
    # into its range textboxes so the user sees the shift immediately.
    if hasattr(window, 'fitting_window') and window.fitting_window is not None:
        try:
            cur_bg = window.Data['Core levels'].get(current_sheet, {}).get('Background', {})
            bg_low = cur_bg.get('Bkg Low', '')
            bg_high = cur_bg.get('Bkg High', '')
            fw = window.fitting_window
            if hasattr(fw, 'min_range_text') and bg_low != '':
                fw.min_range_text.SetValue(f"{float(bg_low):.2f}")
            if hasattr(fw, 'max_range_text') and bg_high != '':
                fw.max_range_text.SetValue(f"{float(bg_high):.2f}")
            if hasattr(fw, 'update_range_boxes'):
                fw.update_range_boxes()
        except Exception as e:
            print(f"Error refreshing Fitting Screen ranges after BE correction: {e}")

    # Update FileManager's BE corrections if open
    if hasattr(window, 'file_manager') and window.file_manager is not None:
        try:
            window.file_manager.save_be_corrections()
        except Exception as e:
            print(f"Error updating FileManager: {e}")


def calculate_c1s_correction(window):
    # Get the current sheet and extract its row number
    current_sheet = window.sheet_combobox.GetValue()
    current_row = "0"  # Default to row 0

    # Extract row number using regex
    match = re.search(r'(\d+)$', current_sheet)
    if match:
        current_row = match.group(1)

    # Look only in sheets from the same row
    matching_sheets = []
    for sheet_name in window.Data['Core levels']:
        sheet_match = re.search(r'(\d+)$', sheet_name)
        sheet_row = sheet_match.group(1) if sheet_match else "0"

        if sheet_row == current_row:
            matching_sheets.append(sheet_name)

    # Check each matching sheet for the reference peak
    for ref_sheet in matching_sheets:
        if 'Fitting' in window.Data['Core levels'][ref_sheet] and 'Peaks' in window.Data['Core levels'][ref_sheet][
            'Fitting']:
            peaks = window.Data['Core levels'][ref_sheet]['Fitting']['Peaks']
            ref_peak = next((peak for label, peak in peaks.items() if window.ref_peak_name in label), None)
            if ref_peak:
                return window.ref_peak_be - ref_peak['Position']

    # If no reference peak found in this row, return None
    return None

def load_be_correction(window):
    if 'BEcorrection' in window.Data:
        window.be_correction = window.Data['BEcorrection']
        window.be_correction_spinbox.SetValue(window.be_correction)