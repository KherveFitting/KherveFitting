import wx
import numpy as np
from scipy.signal import find_peaks
from scipy.integrate import trapz
from libraries.FileMenu.Save import save_state
from libraries.FileMenu.Open import load_library_data


class AutoSurveyID:
    def __init__(self, parent):
        self.parent = parent
        self.library_data = load_library_data()

        # Priority lists from AreaFit_Screen
        self.priority_1_elements = {
            'C': '1s', 'N': '1s', 'O': '1s', 'F': '1s', 'Na': '1s', 'Mg': '1s',
            'K': '2p', 'Ca': '2p', 'Si': '2p', 'P': '2p', 'S': '2p', 'Cl': '2p', 'I': '3d'
        }

        self.priority_2_elements = {
            'Ti': '2p', 'V': '2p', 'Cr': '2p', 'Mn': '2p', 'Fe': '2p', 'Co': '2p',
            'Ni': '2p', 'Cu': '2p', 'Zn': '2p', 'Ag': '3d', 'Au': '4f', 'In': '3d', 'Sn': '3d'
        }

        self.priority_3_elements = {
            'Sb': '3d', 'Ba': '3d', 'Ta': '4f', 'W': '4f', 'Re': '4f', 'Ir': '4f', 'Pt': '4f',
            'La': '3d', 'Ce': '3d', 'Pr': '3d', 'Sm': '3d', 'Nd': '3d', 'Eu': '3d',
            'Gd': '3d', 'Dy': '3d', 'Cs': '3d'
        }

        # Orbital hierarchy for verification
        self.orbital_hierarchy = {
            '2p': ['2s'],
            '3d': ['3p', '3s'],
            '4f': ['4d', '4p', '4s'],
            '4d': ['4p', '4s']
        }

        # Minimum intensity threshold (2% of max)
        self.min_intensity_threshold = 0.02

    def run(self):
        """Main execution method"""
        try:
            sheet_name = self.parent.sheet_combobox.GetValue()

            # Check if this is a survey
            if not any(x in sheet_name.lower() for x in ['survey', 'wide']):
                wx.MessageBox("Please select a Survey or Wide scan sheet", "Info")
                return

            # Get survey data
            if sheet_name not in self.parent.Data['Core levels']:
                wx.MessageBox("No data found for selected sheet", "Error")
                return

            x_values = np.array(self.parent.Data['Core levels'][sheet_name]['B.E.'])
            y_values_raw = np.array(self.parent.Data['Core levels'][sheet_name]['Raw Data'])

            # Apply Gaussian smoothing with width 1
            from scipy.ndimage import gaussian_filter1d
            y_values = gaussian_filter1d(y_values_raw, sigma=1.0)
            print(f"Applied Gaussian smoothing (sigma=1.0) to survey data")

            print(f"\n=== AutoID Starting for {sheet_name} ===")
            print(f"Data range: {np.min(x_values):.2f} to {np.max(x_values):.2f} eV")

            # Save state before making changes
            save_state(self.parent)

            # Store vLine positions BEFORE any operations that might destroy them
            vline1_x = None
            vline2_x = None
            if hasattr(self.parent, 'vline1') and self.parent.vline1 is not None:
                vline1_x = self.parent.vline1.get_xdata()[0]
            if hasattr(self.parent, 'vline2') and self.parent.vline2 is not None:
                vline2_x = self.parent.vline2.get_xdata()[0]

            # DELETE WHOLE PREVIOUS BACKGROUND - reset to raw data
            if 'Background' in self.parent.Data['Core levels'][sheet_name]:
                self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = y_values_raw.tolist()
                print("Cleared previous background - reset to raw data")

            # Clear existing peaks in grid
            self.clear_peak_grid()

            # Find peaks in survey
            peaks_found = self.find_survey_peaks(x_values, y_values)

            # Identify elements based on peaks
            identified_elements = self.identify_elements(peaks_found, x_values, y_values)

            # Create peaks and measure areas
            self.create_peaks_and_measure(identified_elements, x_values, y_values, sheet_name)

            # Force grid refresh to show new data
            self.parent.peak_params_grid.ForceRefresh()

            # Update the plot to show new peaks and backgrounds
            print(f"\n=== AutoID: Updating plot display ===")

            # First plot the raw data
            self.parent.plot_manager.plot_data(self.parent)

            # Then trigger a complete replot with fits and residuals
            self.parent.clear_and_replot()

            # Update legend to include new peaks
            self.parent.plot_manager.update_legend(self.parent)

            # REINITIALIZE VLINES after clear_and_replot destroys them
            if (hasattr(self.parent, 'background_window') and
                    self.parent.background_window is not None and
                    hasattr(self.parent, 'area_tab_selected') and
                    self.parent.area_tab_selected):

                # Reinitialize vLines using the area screen method
                if hasattr(self.parent.background_window, 'initialize_or_restore_area_vlines'):
                    self.parent.background_window.initialize_or_restore_area_vlines()

                # If we had stored positions, restore them
                if vline1_x is not None and vline2_x is not None:
                    if self.parent.vline1 is not None:
                        self.parent.vline1.set_xdata([vline1_x, vline1_x])
                    if self.parent.vline2 is not None:
                        self.parent.vline2.set_xdata([vline2_x, vline2_x])

                    # Update text labels
                    if hasattr(self.parent.background_window, 'update_vline_text_labels'):
                        self.parent.background_window.update_vline_text_labels()

                    # Update range controls
                    if hasattr(self.parent.background_window, 'update_range_controls_from_data'):
                        self.parent.background_window.update_range_controls_from_data()

                # Reset mouse interaction system
                if hasattr(self.parent, 'mouse_handler'):
                    self.parent.mouse_handler.cleanup_vline_handlers()
                    self.parent.moving_vline = None

            # Force canvas redraw
            self.parent.canvas.draw_idle()

            print(f"=== AutoID Complete: {len(identified_elements)} elements identified ===")

            wx.MessageBox(f"Auto ID completed.\n\nFound {len(identified_elements)} elements:\n" +
                          "\n".join([f"• {key}: {data['peak_position']:.2f} eV"
                                     for key, data in identified_elements.items()]),
                          "AutoID Results", wx.OK | wx.ICON_INFORMATION)

        except Exception as e:
            print(f"AutoID Error: {str(e)}")
            import traceback
            traceback.print_exc()
            wx.MessageBox(f"Auto ID failed: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def find_survey_peaks(self, x_data, y_data):
        """Find peaks in survey spectrum with enhanced O1s detection and debugging"""
        print(f"\n=== AutoID Peak Detection Debug ===")
        print(f"Data range: {np.min(x_data):.2f} to {np.max(x_data):.2f} eV")
        print(f"Data points: {len(x_data)}")



        # Normalize y_data
        y_norm = y_data / np.max(y_data)
        print(f"Max intensity: {np.max(y_data):.2f}")

        # Use lower prominence threshold for better O1s detection
        base_prominence = max(0.005, self.min_intensity_threshold * 0.3)  # Even lower threshold
        print(f"Base prominence threshold: {base_prominence:.4f}")

        # Find peaks with adjusted parameters for better detection
        peaks, properties = find_peaks(y_norm,
                                       # prominence=base_prominence,
                                       prominence= 0.01,
                                       # width=1,  # Reduced from 2
                                       width=0.6,
                                       distance=3)  # Reduced from 5

        # Get peak positions and intensities
        peak_positions = x_data[peaks]
        peak_intensities = y_norm[peaks]
        peak_prominences = properties['prominences']

        print(f"Initial peaks found: {len(peaks)}")

        # Special debugging for O1s region (528-536 eV)
        o1s_region_mask = (peak_positions >= 528) & (peak_positions <= 536)
        o1s_peaks = peaks[o1s_region_mask]
        o1s_positions = peak_positions[o1s_region_mask]
        o1s_intensities = peak_intensities[o1s_region_mask]

        print(f"\n=== O1s Region Analysis (528-536 eV) ===")
        print(f"Peaks in O1s region: {len(o1s_peaks)}")

        if len(o1s_peaks) > 0:
            for i, (pos, intensity) in enumerate(zip(o1s_positions, o1s_intensities)):
                print(f"  O1s candidate {i + 1}: {pos:.2f} eV, intensity: {intensity:.4f}")

            # Force include the strongest peak in O1s region
            o1s_peak_idx = np.argmax(o1s_intensities)
            strongest_o1s_peak = o1s_peaks[o1s_peak_idx]
            strongest_o1s_pos = o1s_positions[o1s_peak_idx]

            print(f"  Strongest O1s peak: {strongest_o1s_pos:.2f} eV")

            # Boost O1s prominence to ensure it's included
            peak_idx_in_full_array = np.where(peaks == strongest_o1s_peak)[0][0]
            original_prominence = peak_prominences[peak_idx_in_full_array]
            peak_prominences[peak_idx_in_full_array] = max(original_prominence, base_prominence * 5)
            print(
                f"  Boosted prominence from {original_prominence:.4f} to {peak_prominences[peak_idx_in_full_array]:.4f}")
        else:
            print("  WARNING: No peaks found in O1s region!")

            # Try to find ANY signal in O1s region, even if below threshold
            o1s_mask = (x_data >= 526) & (x_data <= 545)
            if np.any(o1s_mask):
                o1s_data = y_norm[o1s_mask]
                o1s_x_data = x_data[o1s_mask]
                max_idx = np.argmax(o1s_data)
                max_pos = o1s_x_data[max_idx]
                max_intensity = o1s_data[max_idx]

                print(f"  Manual O1s search: max at {max_pos:.2f} eV, intensity: {max_intensity:.4f}")

                # If there's any signal above 1% of max, add it as a peak
                if max_intensity > 0.01:
                    print(f"  Adding manual O1s peak at {max_pos:.2f} eV")
                    # Add this as a new peak
                    peaks = np.append(peaks, o1s_mask.nonzero()[0][max_idx])
                    peak_positions = np.append(peak_positions, max_pos)
                    peak_intensities = np.append(peak_intensities, max_intensity)
                    peak_prominences = np.append(peak_prominences, base_prominence * 3)

        # Check for other important elements
        important_regions = {
            'C1s': (280, 290),
            'N1s': (395, 405),
            'Si2p': (98, 105),
            'Al2p': (70, 78)
        }

        print(f"\n=== Other Important Elements ===")
        for element, (low, high) in important_regions.items():
            element_mask = (peak_positions >= low) & (peak_positions <= high)
            element_peaks = len(peak_positions[element_mask])
            print(f"{element} region ({low}-{high} eV): {element_peaks} peaks")

        # Sort by prominence (most prominent first)
        sorted_indices = np.argsort(peak_prominences)[::-1]

        peaks_found = []
        for idx in sorted_indices[:50]:  # Limit to top 50 peaks
            peaks_found.append({
                'position': peak_positions[idx],
                'intensity': peak_intensities[idx],
                'prominence': peak_prominences[idx],
                'index': peaks[idx]
            })

        print(f"\nFinal peaks for analysis: {len(peaks_found)}")
        print(f"Top 10 peaks by prominence:")
        for i, peak in enumerate(peaks_found[:10]):
            print(f"  {i + 1}. {peak['position']:.2f} eV, prominence: {peak['prominence']:.4f}")

        return peaks_found

    def identify_elements(self, peaks_found, x_data, y_data):
        """Identify elements from found peaks using AreaFit_Screen database approach"""
        print(f"\n=== Element Identification (AreaFit_Screen Database) ===")

        identified = {}
        used_peaks = set()
        detected_elements = set()  # Track elements we've already identified

        # Use the hardcoded elements database like AreaFit_Screen
        # Use existing method and filter out high energy orbitals (>1400 eV)
        elements_db_raw = self.get_core_level_ranges()
        elements_db = self.filter_high_energy_orbitals(elements_db_raw, max_energy=1400.0)

        # Convert to simple list format for compatibility with find_best_element_match
        element_db = []
        for element, orbitals in elements_db.items():
            for orbital, (be_min, be_max) in orbitals.items():
                be_center = (be_min + be_max) / 2
                element_db.append({
                    'element': element,
                    'orbital': orbital,
                    'position': be_center,
                    'rsf': 1.0  # Default RSF, will be updated later
                })

        print(f"Using AreaFit_Screen database with {len(element_db)} entries")

        if len(element_db) == 0:
            print("ERROR: Element database is empty!")
            return {}

        # Step 1: Always check for O1s first (around 532 eV)
        print(f"\n=== Step 1: O1s Detection ===")
        o1s_tolerance = 4.0
        o1s_expected = 532.0

        for peak in peaks_found:
            if abs(peak['position'] - o1s_expected) <= o1s_tolerance:
                print(f"  Found O1s at {peak['position']:.2f} eV (expected: {o1s_expected:.2f})")
                identified['O1s'] = {
                    'element': 'O',
                    'orbital': '1s',
                    'position': o1s_expected,
                    'peak_position': peak['position'],
                    'intensity': peak['intensity'],
                    'prominence': peak['prominence'],
                    'priority': 1
                }
                used_peaks.add(peak['index'])
                detected_elements.add('O')  # Remember we found oxygen
                print(f"  Found O1s at {peak['position']:.2f} eV - O added to detected elements")
                break

        # Step 2: Always check for C1s second (around 285 eV)
        print(f"\n=== Step 2: C1s Detection ===")
        c1s_tolerance = 4.0
        c1s_expected = 285.0

        for peak in peaks_found:
            if peak['index'] in used_peaks:
                continue
            if abs(peak['position'] - c1s_expected) <= c1s_tolerance:
                print(f"  Found C1s at {peak['position']:.2f} eV (expected: {c1s_expected:.2f})")
                identified['C1s'] = {
                    'element': 'C',
                    'orbital': '1s',
                    'position': c1s_expected,
                    'peak_position': peak['position'],
                    'intensity': peak['intensity'],
                    'prominence': peak['prominence'],
                    'priority': 1
                }
                used_peaks.add(peak['index'])
                detected_elements.add('C')  # Remember we found carbon
                print(f"  Found C1s at {peak['position']:.2f} eV - C added to detected elements")
                break

        # Step 3: Process remaining peaks by prominence (highest first)
        print(f"\n=== Step 3: Process Remaining Peaks by Prominence ===")
        remaining_peaks = [p for p in peaks_found if p['index'] not in used_peaks]
        remaining_peaks.sort(key=lambda x: x['prominence'], reverse=True)

        for i, peak in enumerate(remaining_peaks):
            if peak['index'] in used_peaks:
                continue

            print(f"\nPeak {i + 1}: {peak['position']:.2f} eV, prominence: {peak['prominence']:.4f}")

            # Find ALL possible matches within tolerance (not just the best one)
            all_matches = self.find_all_element_matches(peak['position'], element_db)

            if len(all_matches) == 0:
                print(f"  ✗ No matches found within tolerance")
                continue

            print(f"  Found {len(all_matches)} possible matches:")
            for j, match in enumerate(all_matches):
                diff = abs(match['position'] - peak['position'])
                print(
                    f"    {j + 1}. {match['element']}{match['orbital']}: expected {match['position']:.2f} eV (diff: {diff:.2f} eV)")

            # Try each match in order of best fit
            peak_assigned = False
            for match_idx, best_match in enumerate(all_matches):
                if peak_assigned:
                    break

                element = best_match['element']
                orbital = best_match['orbital']

                print(f"\n  Trying match {match_idx + 1}: {element}{orbital} (expected: {best_match['position']:.2f})")

                # Check if already identified
                key = f"{element}{orbital}"
                if key in identified:
                    print(f"    ✗ {key} already identified - trying next match")
                    continue

                # Check if element is already detected or if this is a main orbital
                is_detected_element = element in detected_elements
                is_main_orbital = self.is_main_orbital(element, orbital)

                print(f"    Element {element}: detected={is_detected_element}, main_orbital={is_main_orbital}")

                # Handle different orbital types
                if orbital in ['1s']:
                    # 1s peaks - no companion required
                    if is_main_orbital or is_detected_element:
                        identified[key] = {
                            'element': element,
                            'orbital': orbital,
                            'position': best_match['position'],
                            'peak_position': peak['position'],
                            'intensity': peak['intensity'],
                            'prominence': peak['prominence'],
                            'priority': self.get_priority_level(element, orbital)
                        }
                        used_peaks.add(peak['index'])
                        detected_elements.add(element)  # Add to detected elements
                        reason = "detected element" if is_detected_element else "main orbital"
                        print(f"    ✓ ASSIGNED as {key} (no companion required - {reason})")
                        peak_assigned = True
                    else:
                        print(
                            f"    ✗ {orbital} is not main orbital for {element} and element not yet detected - trying next match")

                elif orbital in ['2s', '3s', '4s']:
                    # s orbitals - allow if element is detected, otherwise look for corresponding p orbital
                    if is_detected_element:
                        # Element already detected - assign directly
                        identified[key] = {
                            'element': element,
                            'orbital': orbital,
                            'position': best_match['position'],
                            'peak_position': peak['position'],
                            'intensity': peak['intensity'],
                            'prominence': peak['prominence'],
                            'priority': self.get_priority_level(element, orbital)
                        }
                        used_peaks.add(peak['index'])
                        print(f"    ✓ ASSIGNED as {key} (element {element} already detected)")
                        peak_assigned = True
                    else:
                        # Look for corresponding p orbital
                        p_orbital = orbital.replace('s', 'p')  # 2s->2p, 3s->3p, etc.
                        print(f"    Found {element}{orbital} - looking for corresponding {element}{p_orbital}...")
                        companion_p = self.find_companion_peak_simple(element, p_orbital, remaining_peaks,
                                                                      used_peaks, element_db)
                        if companion_p:
                            key_s = f"{element}{orbital}"
                            key_p = f"{element}{p_orbital}"

                            if key_s not in identified and key_p not in identified:
                                identified[key_s] = {
                                    'element': element,
                                    'orbital': orbital,
                                    'position': best_match['position'],
                                    'peak_position': peak['position'],
                                    'intensity': peak['intensity'],
                                    'prominence': peak['prominence'],
                                    'priority': self.get_priority_level(element, orbital)
                                }

                                identified[key_p] = {
                                    'element': element,
                                    'orbital': p_orbital,
                                    'position': companion_p['expected_position'],
                                    'peak_position': companion_p['peak']['position'],
                                    'intensity': companion_p['peak']['intensity'],
                                    'prominence': companion_p['peak']['prominence'],
                                    'priority': self.get_priority_level(element, p_orbital)
                                }

                                used_peaks.add(peak['index'])
                                used_peaks.add(companion_p['peak']['index'])
                                detected_elements.add(element)  # Add to detected elements
                                print(f"    ✓ ASSIGNED as {key_s} with {key_p} companion")
                                peak_assigned = True
                            else:
                                print(f"    ✗ {key_s} or {key_p} already identified - trying next match")
                        else:
                            print(f"    ✗ No {element}{p_orbital} companion found - trying next match")

                elif orbital in ['2p', '3p', '4p']:
                    # p orbitals - handle based on main orbital status or detected element
                    if orbital == '2p' and (is_main_orbital or is_detected_element):
                        # 2p peak - MUST find 2s companion (unless element already detected)
                        if is_detected_element:
                            # Element already detected - assign directly
                            identified[key] = {
                                'element': element,
                                'orbital': orbital,
                                'position': best_match['position'],
                                'peak_position': peak['position'],
                                'intensity': peak['intensity'],
                                'prominence': peak['prominence'],
                                'priority': self.get_priority_level(element, orbital)
                            }
                            used_peaks.add(peak['index'])
                            print(f"    ✓ ASSIGNED as {key} (element {element} already detected)")
                            peak_assigned = True
                        else:
                            # Look for required 2s companion
                            print(f"    Found {element}2p - looking for required {element}2s companion...")
                            companion_2s = self.find_companion_peak_simple(element, '2s', remaining_peaks,
                                                                           used_peaks, element_db)
                            if companion_2s:
                                key_2p = f"{element}2p"
                                key_2s = f"{element}2s"

                                if key_2p not in identified and key_2s not in identified:
                                    identified[key_2p] = {
                                        'element': element,
                                        'orbital': '2p',
                                        'position': best_match['position'],
                                        'peak_position': peak['position'],
                                        'intensity': peak['intensity'],
                                        'prominence': peak['prominence'],
                                        'priority': self.get_priority_level(element, '2p')
                                    }

                                    identified[key_2s] = {
                                        'element': element,
                                        'orbital': '2s',
                                        'position': companion_2s['expected_position'],
                                        'peak_position': companion_2s['peak']['position'],
                                        'intensity': companion_2s['peak']['intensity'],
                                        'prominence': companion_2s['peak']['prominence'],
                                        'priority': self.get_priority_level(element, '2s')
                                    }

                                    used_peaks.add(peak['index'])
                                    used_peaks.add(companion_2s['peak']['index'])
                                    detected_elements.add(element)  # Add to detected elements
                                    print(f"    ✓ ASSIGNED as {key_2p} with required {key_2s} companion")
                                    peak_assigned = True
                                else:
                                    print(f"    ✗ {key_2p} or {key_2s} already identified - trying next match")
                            else:
                                print(f"    ✗ No required {element}2s companion found - trying next match")
                    elif is_detected_element:
                        # Other p orbitals (3p, 4p) - assign if element already detected
                        identified[key] = {
                            'element': element,
                            'orbital': orbital,
                            'position': best_match['position'],
                            'peak_position': peak['position'],
                            'intensity': peak['intensity'],
                            'prominence': peak['prominence'],
                            'priority': self.get_priority_level(element, orbital)
                        }
                        used_peaks.add(peak['index'])
                        print(f"    ✓ ASSIGNED as {key} (element {element} already detected)")
                        peak_assigned = True
                    else:
                        print(
                            f"    ✗ {orbital} is not main orbital for {element} and element not yet detected - trying next match")

                elif orbital == '3d':
                    # 3d peak - MUST find 3p companion (unless element already detected)
                    if is_main_orbital or is_detected_element:
                        if is_detected_element:
                            # Element already detected - assign directly
                            identified[key] = {
                                'element': element,
                                'orbital': orbital,
                                'position': best_match['position'],
                                'peak_position': peak['position'],
                                'intensity': peak['intensity'],
                                'prominence': peak['prominence'],
                                'priority': self.get_priority_level(element, orbital)
                            }
                            used_peaks.add(peak['index'])
                            print(f"    ✓ ASSIGNED as {key} (element {element} already detected)")
                            peak_assigned = True
                        else:
                            # Look for required 3p companion
                            print(f"    Found {element}3d - looking for required {element}3p companion...")
                            companion_3p = self.find_companion_peak_simple(element, '3p', remaining_peaks,
                                                                           used_peaks, element_db)
                            if companion_3p:
                                key_3d = f"{element}3d"
                                key_3p = f"{element}3p"

                                if key_3d not in identified and key_3p not in identified:
                                    identified[key_3d] = {
                                        'element': element,
                                        'orbital': '3d',
                                        'position': best_match['position'],
                                        'peak_position': peak['position'],
                                        'intensity': peak['intensity'],
                                        'prominence': peak['prominence'],
                                        'priority': self.get_priority_level(element, '3d')
                                    }

                                    identified[key_3p] = {
                                        'element': element,
                                        'orbital': '3p',
                                        'position': companion_3p['expected_position'],
                                        'peak_position': companion_3p['peak']['position'],
                                        'intensity': companion_3p['peak']['intensity'],
                                        'prominence': companion_3p['peak']['prominence'],
                                        'priority': self.get_priority_level(element, '3p')
                                    }

                                    used_peaks.add(peak['index'])
                                    used_peaks.add(companion_3p['peak']['index'])
                                    detected_elements.add(element)  # Add to detected elements
                                    print(f"    ✓ ASSIGNED as {key_3d} with required {key_3p} companion")
                                    peak_assigned = True
                                else:
                                    print(f"    ✗ {key_3d} or {key_3p} already identified - trying next match")
                            else:
                                print(f"    ✗ No required {element}3p companion found - trying next match")
                    else:
                        print(
                            f"    ✗ 3d is not main orbital for {element} and element not yet detected - trying next match")

                elif orbital in ['4d', '4f']:
                    # 4d/4f orbitals - handle based on detected element status
                    if orbital == '4f' and (is_main_orbital or is_detected_element):
                        if is_detected_element:
                            # Element already detected - assign directly
                            identified[key] = {
                                'element': element,
                                'orbital': orbital,
                                'position': best_match['position'],
                                'peak_position': peak['position'],
                                'intensity': peak['intensity'],
                                'prominence': peak['prominence'],
                                'priority': self.get_priority_level(element, orbital)
                            }
                            used_peaks.add(peak['index'])
                            print(f"    ✓ ASSIGNED as {key} (element {element} already detected)")
                            peak_assigned = True
                        else:
                            # Look for required 4d companion
                            print(f"    Found {element}4f - looking for required {element}4d companion...")
                            companion_4d = self.find_companion_peak_simple(element, '4d', remaining_peaks,
                                                                           used_peaks, element_db)
                            if companion_4d:
                                key_4f = f"{element}4f"
                                key_4d = f"{element}4d"

                                if key_4f not in identified and key_4d not in identified:
                                    identified[key_4f] = {
                                        'element': element,
                                        'orbital': '4f',
                                        'position': best_match['position'],
                                        'peak_position': peak['position'],
                                        'intensity': peak['intensity'],
                                        'prominence': peak['prominence'],
                                        'priority': self.get_priority_level(element, '4f')
                                    }

                                    identified[key_4d] = {
                                        'element': element,
                                        'orbital': '4d',
                                        'position': companion_4d['expected_position'],
                                        'peak_position': companion_4d['peak']['position'],
                                        'intensity': companion_4d['peak']['intensity'],
                                        'prominence': companion_4d['peak']['prominence'],
                                        'priority': self.get_priority_level(element, '4d')
                                    }

                                    used_peaks.add(peak['index'])
                                    used_peaks.add(companion_4d['peak']['index'])
                                    detected_elements.add(element)  # Add to detected elements
                                    print(f"    ✓ ASSIGNED as {key_4f} with required {key_4d} companion")
                                    peak_assigned = True
                                else:
                                    print(f"    ✗ {key_4f} or {key_4d} already identified - trying next match")
                            else:
                                print(f"    ✗ No required {element}4d companion found - trying next match")
                    elif is_detected_element:
                        # 4d orbital - assign if element already detected
                        identified[key] = {
                            'element': element,
                            'orbital': orbital,
                            'position': best_match['position'],
                            'peak_position': peak['position'],
                            'intensity': peak['intensity'],
                            'prominence': peak['prominence'],
                            'priority': self.get_priority_level(element, orbital)
                        }
                        used_peaks.add(peak['index'])
                        print(f"    ✓ ASSIGNED as {key} (element {element} already detected)")
                        peak_assigned = True
                    else:
                        print(
                            f"    ✗ {orbital} is not main orbital for {element} and element not yet detected - trying next match")


                elif is_detected_element and self.is_auger_orbital(orbital):
                    # Auger peaks - only assign if element already detected
                    identified[key] = {
                        'element': element,
                        'orbital': orbital,
                        'position': best_match['position'],
                        'peak_position': peak['position'],
                        'intensity': peak['intensity'],
                        'prominence': peak['prominence'],
                        'priority': self.get_priority_level(element, orbital)
                    }
                    used_peaks.add(peak['index'])
                    print(f"    ✓ ASSIGNED as {key} (Auger peak - element {element} already detected)")
                    peak_assigned = True
                else:
                    orbital_type = "Auger peak" if self.is_auger_orbital(orbital) else "orbital type"
                    if self.is_auger_orbital(orbital) and not is_detected_element:
                        print(
                            f"    ✗ {orbital} is an Auger peak but element {element} not yet detected - trying next match")
                    else:
                        print(f"    ✗ {orbital} is not a supported {orbital_type} - trying next match")

        print(f"\n=== Detected Elements ===")
        if detected_elements:
            print(f"Elements found in sample: {sorted(detected_elements)}")
        else:
            print("No elements detected")

        print(f"\n=== Final Results: {len(identified)} peaks identified ===")
        for key, data in identified.items():
            print(f"  {key}: {data['peak_position']:.2f} eV, priority {data['priority']}")

        return identified

    def is_auger_orbital(self, orbital):
        """Check if orbital represents an Auger peak"""
        orbital_lower = orbital.lower()

        # Common Auger peak patterns
        auger_patterns = [
            'kll', 'klm', 'kmm', 'lmm', 'lmv', 'mnn', 'mvv', 'nvv',
            'noo', 'moo', 'lll', 'mmm', 'nnn'
        ]

        # Check for Auger patterns
        for pattern in auger_patterns:
            if pattern in orbital_lower:
                return True

        # Check for numbered Auger transitions (like mn1, mn2, etc.)
        import re
        if re.match(r'^[klmno]{2,3}\d*$', orbital_lower):
            return True

        return False

    def find_all_element_matches(self, peak_position, element_db):
        """Find all matching elements for given peak position, sorted by best fit"""
        tolerance = 6.0  # eV tolerance
        matches = []

        for entry in element_db:
            diff = abs(entry['position'] - peak_position)
            if diff <= tolerance:
                matches.append({
                    'element': entry['element'],
                    'orbital': entry['orbital'],
                    'position': entry['position'],
                    'rsf': entry['rsf'],
                    'diff': diff
                })

        # Sort by difference (best fit first)
        matches.sort(key=lambda x: x['diff'])
        return matches

    def find_companion_peak_simple(self, element, companion_orbital, remaining_peaks, used_peaks, element_db):
        """Find companion peak (e.g., 2s for 2p, 3p for 3d) - simplified version"""
        # Find expected position for companion orbital
        expected_position = None
        for entry in element_db:
            if entry['element'] == element and entry['orbital'] == companion_orbital:
                expected_position = entry['position']
                break

        if not expected_position:
            return None

        # Look for peak near expected position
        tolerance = 4.0  # Increased tolerance
        for peak in remaining_peaks:
            if peak['index'] in used_peaks:
                continue
            if abs(peak['position'] - expected_position) <= tolerance:
                return {
                    'peak': peak,
                    'expected_position': expected_position
                }

        return None

    def build_element_database(self):
        """Build element database from library data"""
        database = []

        for (element, orbital), data in self.library_data.items():
            # Get instrument data
            instrument = self.parent.current_instrument
            if instrument not in data:
                instrument = 'Al1486' if 'Al1486' in data else next(iter(data))

            if 'position' in data[instrument]:
                position = float(data[instrument]['position'])
                rsf = float(data[instrument].get('rsf', 1.0))

                database.append({
                    'element': element,
                    'orbital': orbital,
                    'position': position,
                    'rsf': rsf
                })

        return database

    def find_best_element_match(self, peak_position, element_db):
        """Find best matching element for given peak position"""
        tolerance = 3.0  # eV tolerance
        best_match = None
        min_diff = tolerance

        for entry in element_db:
            diff = abs(entry['position'] - peak_position)
            if diff < min_diff:
                min_diff = diff
                best_match = entry

        return best_match

    def is_main_orbital(self, element, orbital):
        """Check if this is the main orbital for the element"""
        # Check priority lists
        if element in self.priority_1_elements:
            return orbital == self.priority_1_elements[element]
        elif element in self.priority_2_elements:
            return orbital == self.priority_2_elements[element]
        elif element in self.priority_3_elements:
            return orbital == self.priority_3_elements[element]

        # Default main orbitals
        if orbital in ['1s', '2p', '3d', '4f']:
            return True

        return False

    def verify_orbital_hierarchy(self, element, orbital, peaks_found, element_db):
        """Verify orbital hierarchy rules"""
        if orbital not in self.orbital_hierarchy:
            return True

        required_orbitals = self.orbital_hierarchy[orbital]

        # For 2p, 3d, 4f - must find supporting orbitals
        for req_orbital in required_orbitals:
            found = False
            search_position = None

            # Find expected position for required orbital
            for entry in element_db:
                if entry['element'] == element and entry['orbital'] == req_orbital:
                    search_position = entry['position']
                    break

            if search_position:
                # Look for peak near this position
                for peak in peaks_found:
                    if abs(peak['position'] - search_position) < 5.0:
                        found = True
                        break

            # For main orbitals, lack of supporting orbital means reject
            if not found and orbital in ['2p', '3d', '4f']:
                return False

        return True

    def get_priority_level(self, element, orbital):
        """Get priority level for element-orbital combination"""

        # Auger peaks get lower priority (higher number)
        if self.is_auger_orbital(orbital):
            return 4  # Lower priority than photoelectron peaks

        if element in self.priority_1_elements and self.priority_1_elements[element] == orbital:
            return 1
        elif element in self.priority_2_elements and self.priority_2_elements[element] == orbital:
            return 2
        elif element in self.priority_3_elements and self.priority_3_elements[element] == orbital:
            return 3
        else:
            return 5

    def filter_by_priority(self, identified):
        """Filter identified elements by priority and intensity"""
        filtered = {}

        # Sort by priority and prominence
        sorted_items = sorted(identified.items(),
                              key=lambda x: (x[1]['priority'], -x[1]['prominence']))

        # Always include priority 1 elements if found
        for key, data in sorted_items:
            if data['priority'] == 1:
                filtered[key] = data

        # Add priority 2 if prominence is sufficient
        for key, data in sorted_items:
            if data['priority'] == 2 and data['prominence'] > 0.01:
                filtered[key] = data

        # Add priority 3 if prominence is significant
        for key, data in sorted_items:
            if data['priority'] == 3 and data['prominence'] > 0.02:
                filtered[key] = data

        return filtered

    def create_peaks_and_measure(self, identified_elements, x_data, y_data, sheet_name):
        """Create peaks in grid and measure areas with Multi-Regions Smart background"""
        print(f"\n=== AutoID: Creating peaks for {len(identified_elements)} identified elements ===")

        # Initialize peak fitting data structure
        if 'Fitting' not in self.parent.Data['Core levels'][sheet_name]:
            self.parent.Data['Core levels'][sheet_name]['Fitting'] = {}
        if 'Peaks' not in self.parent.Data['Core levels'][sheet_name]['Fitting']:
            self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = {}

        peaks_data = self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']

        # Initialize background data structure
        if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
            self.parent.Data['Core levels'][sheet_name]['Background'] = {}
        if 'Bkg Y' not in self.parent.Data['Core levels'][sheet_name]['Background']:
            self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = y_data.tolist()

        # Sort by binding energy (high to low)
        sorted_elements = sorted(identified_elements.items(),
                                 key=lambda x: x[1]['peak_position'],
                                 reverse=True)

        row = 0
        for peak_name, element_data in sorted_elements:
            print(f"\nProcessing peak: {peak_name}")
            print(f"  Position: {element_data['peak_position']:.2f} eV")
            print(f"  Priority: {element_data['priority']}")
            print(f"  Prominence: {element_data['prominence']:.4f}")

            # Calculate area and height with Multi-Regions Smart background
            area, bg_low, bg_high, peak_height = self.calculate_peak_area_and_height_with_background(
                element_data['peak_position'],
                x_data,
                y_data,
                sheet_name
            )

            # Format peak name like AreaFit_Screen (with space and dot)
            formatted_peak_name = f"{peak_name} ."
            print(f"  Formatted name: {formatted_peak_name}")
            print(f"  Calculated area: {area:.2f}")
            print(f"  Calculated height: {peak_height:.2f}")

            # Get RSF value
            rsf = self.get_rsf_value(element_data['element'], element_data['orbital'])

            # Add to grid
            letter_id = chr(65 + row // 2)  # A, B, C, etc.

            # Ensure we have enough rows
            if row >= self.parent.peak_params_grid.GetNumberRows():
                self.parent.peak_params_grid.AppendRows(2)

            # Set main row values with .2f formatting
            self.parent.peak_params_grid.SetCellValue(row, 0, letter_id)
            self.parent.peak_params_grid.SetCellValue(row, 1, formatted_peak_name)
            self.parent.peak_params_grid.SetCellValue(row, 2, f"{element_data['peak_position']:.2f}")
            self.parent.peak_params_grid.SetCellValue(row, 3, f"{peak_height:.2f}")  # Height
            self.parent.peak_params_grid.SetCellValue(row, 4, "2.00")  # FWHM
            self.parent.peak_params_grid.SetCellValue(row, 5, "30.00")  # L/G
            self.parent.peak_params_grid.SetCellValue(row, 6, f"{area:.2f}")  # Area
            self.parent.peak_params_grid.SetCellValue(row, 7, "0.00")  # Sigma
            self.parent.peak_params_grid.SetCellValue(row, 8, "0.00")  # Gamma
            self.parent.peak_params_grid.SetCellValue(row, 9, "0.00")  # Skew
            self.parent.peak_params_grid.SetCellValue(row, 13, "Unfitted")  # Model
            self.parent.peak_params_grid.SetCellValue(row, 14, "Multi-Regions Smart")  # Bkg Type
            self.parent.peak_params_grid.SetCellValue(row, 15, f"{bg_low:.2f}")  # Bkg Low
            self.parent.peak_params_grid.SetCellValue(row, 16, f"{bg_high:.2f}")  # Bkg High
            self.parent.peak_params_grid.SetCellValue(row, 17, "0.00")  # Offset Low
            self.parent.peak_params_grid.SetCellValue(row, 18, "0.00")  # Offset High

            # Set constraint row background color
            for col in range(self.parent.peak_params_grid.GetNumberCols()):
                self.parent.peak_params_grid.SetCellBackgroundColour(row + 1, col, wx.Colour(200, 245, 228))

            # Set constraints with .2f formatting
            self.parent.peak_params_grid.SetCellValue(row + 1, 2, "Fixed")  # Position
            self.parent.peak_params_grid.SetCellValue(row + 1, 3, "Fixed")  # Height
            self.parent.peak_params_grid.SetCellValue(row + 1, 4, "Fixed")  # FWHM
            self.parent.peak_params_grid.SetCellValue(row + 1, 5, "Fixed")  # L/G


            # Add to Data structure
            peaks_data[formatted_peak_name] = {
                'Position': float(f"{element_data['peak_position']:.2f}"),
                'Height': float(f"{peak_height:.2f}"),
                'FWHM': 2.00,
                'L/G': 30.00,
                'Area': float(f"{area:.2f}"),
                'Sigma': 0.00,
                'Gamma': 0.00,
                'Skew': 0.00,
                'Fitting Model': 'Unfitted',
                'Bkg Type': 'Multi-Regions Smart',
                'Bkg Low': float(f"{bg_low:.2f}"),
                'Bkg High': float(f"{bg_high:.2f}"),
                'Bkg Offset Low': 0.00,
                'Bkg Offset High': 0.00,
                'Constraints': {
                    'position': 'Fixed',
                    'height': 'Fixed',
                    'fwhm': 'Fixed',
                    'lg_ratio': 'Fixed',
                    'area': 'Fixed'
                }
            }

            row += 2

        # Update peak count
        self.parent.peak_count = row // 2
        print(f"\n=== AutoID: Created {self.parent.peak_count} peaks total ===")

    def calculate_peak_area_and_height_with_background(self, peak_position, x_data, y_data, sheet_name):
        """Calculate area and height with Multi-Regions Smart background using defined ranges"""
        from libraries.Peak_Functions import BackgroundCalculations

        # Get core level ranges database
        core_level_ranges = self.get_core_level_ranges()

        # Find which element/orbital this peak belongs to by checking ranges
        bg_low = peak_position - 15.0  # Default fallback
        bg_high = peak_position + 15.0
        found_range = False

        for element, orbitals in core_level_ranges.items():
            for orbital, (range_low, range_high) in orbitals.items():
                if range_low <= peak_position <= range_high:
                    bg_low = range_low
                    bg_high = range_high
                    found_range = True
                    print(f"    Using {element}{orbital} range: {bg_low:.2f} to {bg_high:.2f} eV")
                    break
            if found_range:
                break

        if not found_range:
            print(f"    Using default ±15 eV range: {bg_low:.2f} to {bg_high:.2f} eV")

        # Get current background or initialize
        if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
            self.parent.Data['Core levels'][sheet_name]['Background'] = {}
        if 'Bkg Y' not in self.parent.Data['Core levels'][sheet_name]['Background']:
            self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = y_data.tolist()

        current_background = np.array(self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Y'])

        # Calculate Multi-Regions Smart background for this range
        adaptive_range = (bg_low, bg_high)
        offset_h = 0.0
        offset_l = 0.0

        background_filtered = BackgroundCalculations.calculate_adaptive_smart_background(
            x_data, y_data, adaptive_range, current_background, offset_h, offset_l
        )

        # Update background in data structure
        self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = background_filtered.tolist()

        # Calculate area and height using background-subtracted data
        mask = (x_data >= bg_low) & (x_data <= bg_high)
        x_range = x_data[mask]
        y_range = y_data[mask]
        bg_range = background_filtered[mask]

        if len(x_range) < 3:
            print(f"    WARNING: Insufficient data points in range")
            return 0.0, bg_low, bg_high, 0.0

        # Calculate background-subtracted data
        y_minus_bg = y_range - bg_range

        # Sort data for proper integration
        sorted_indices = np.argsort(x_range)
        x_sorted = x_range[sorted_indices]
        y_minus_bg_sorted = y_minus_bg[sorted_indices]

        # Calculate area - same as AreaFit_Screen
        area = np.trapz(y_minus_bg_sorted, x_sorted)

        # Calculate height - same as AreaFit_Screen
        peak_index = np.argmax(y_minus_bg)
        peak_height = y_minus_bg[peak_index]

        area = abs(area)
        peak_height = max(0.0, peak_height)

        print(f"    Raw area: {area:.2f}, Height: {peak_height:.2f}")

        return area, bg_low, bg_high, peak_height

    def calculate_peak_area(self, peak_position, x_data, y_data):
        """Calculate area under peak with background subtraction"""
        # Define integration window based on element type
        window_width = 15.0  # eV on each side

        # Find indices for integration window
        idx_center = np.argmin(np.abs(x_data - peak_position))
        idx_low = np.argmin(np.abs(x_data - (peak_position - window_width)))
        idx_high = np.argmin(np.abs(x_data - (peak_position + window_width)))

        # Ensure correct order
        if idx_low > idx_high:
            idx_low, idx_high = idx_high, idx_low

        # Extract region
        x_region = x_data[idx_low:idx_high + 1]
        y_region = y_data[idx_low:idx_high + 1]

        if len(x_region) < 3:
            return 0.0, peak_position - window_width, peak_position + window_width

        # Calculate linear background
        background = np.linspace(y_region[0], y_region[-1], len(y_region))

        # Calculate area above background
        y_corrected = y_region - background
        y_corrected[y_corrected < 0] = 0  # Remove negative values

        # Calculate area using trapezoidal rule
        area = trapz(y_corrected, x_region)

        return abs(area), x_region[0], x_region[-1]

    def get_rsf_value(self, element, orbital):
        """Get RSF value for element-orbital combination"""
        key = (element, orbital)

        if key in self.library_data:
            instrument = self.parent.current_instrument
            if instrument not in self.library_data[key]:
                instrument = 'Al1486' if 'Al1486' in self.library_data[key] else next(iter(self.library_data[key]))

            if 'rsf' in self.library_data[key][instrument]:
                return float(self.library_data[key][instrument]['rsf'])

        return 1.00  # Default RSF

    def clear_peak_grid(self):
        """Clear existing peaks in peak fitting grid"""
        # Clear grid cells
        for row in range(self.parent.peak_params_grid.GetNumberRows()):
            for col in range(self.parent.peak_params_grid.GetNumberCols()):
                if col == 0 and row % 2 == 0:
                    # Keep letter IDs
                    letter = chr(65 + row // 2)
                    self.parent.peak_params_grid.SetCellValue(row, col, letter)
                else:
                    self.parent.peak_params_grid.SetCellValue(row, col, "")

        # Clear data structure
        sheet_name = self.parent.sheet_combobox.GetValue()
        if sheet_name in self.parent.Data['Core levels']:
            if 'Fitting' in self.parent.Data['Core levels'][sheet_name]:
                if 'Peaks' in self.parent.Data['Core levels'][sheet_name]['Fitting']:
                    self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'].clear()

        # Reset peak count
        self.parent.peak_count = 0

    def filter_high_energy_orbitals(self, elements_db, max_energy=1400.0):
        """Filter out orbitals with binding energies above max_energy"""
        filtered_db = {}

        for element, orbitals in elements_db.items():
            filtered_orbitals = {}

            for orbital, (be_min, be_max) in orbitals.items():
                # Only include orbitals where the minimum BE is below max_energy
                if be_min <= max_energy:
                    filtered_orbitals[orbital] = (be_min, be_max)

            # Only include elements that have at least one valid orbital
            if filtered_orbitals:
                filtered_db[element] = filtered_orbitals

        print(f"Filtered database: {len(filtered_db)} elements (removed high energy orbitals > {max_energy} eV)")

        # Print some examples of what was kept
        example_count = 0
        for element, orbitals in filtered_db.items():
            if example_count < 5:
                orbital_list = list(orbitals.keys())
                print(f"  {element}: {orbital_list}")
                example_count += 1

        return filtered_db


    def get_core_level_ranges(self, max_energy=1400.0):
        """Hardcoded database of core levels with binding energy ranges for fine-tuning"""

        """Get core level ranges, excluding orbitals above max_energy eV"""
        elements_db = {
            # Light elements
            'C': {'1s': (276.0, 295.0), '2s': (20.0, 25.0)},
            'N': {'1s': (395.0, 405.0),'2s': (25.0, 30.0)},
            'O': {'1s': (521.0, 538.0), '2s': (35.0, 45.0)},
            'F': {'1s': (680.0, 695.0),'2s': (45.0, 55.0)},
            'Na': {'1s': (1057.0, 1081.0),'kll': (492.0, 501.0)},
            'Mg': {'1s': (1300.0, 1310.0),'2s': (85.0, 95.0),'2p': (45.0, 55.0)},
            'Al': {'2s': (115.0, 125.0),'2p': (70.0, 78.0)},
            'Si': {'2s': (145.0, 155.0),'2p': (98.0, 106.0)},
            'P': {'2s': (180.0, 200.0),'2p': (125.0, 141.0)},
            'S': {'2s': (225.0, 235.0),'2p': (158.0, 170.0)},
            'Cl': {'2s': (265.0, 275.0),'2p': (195.0, 205.0)},
            'K': {'2s': (375.0, 385.0),'2p': (290.0, 300.0)},
            'Ca': {'2s': (428.0, 452.0),'2p': (343.0, 356.0)},

            # Transition metals
            'Ti': {'2s': (556.0, 571.0),'2p': (450.0, 481.0), '3s': (60.00, 61.00),'3p': (37.00, 38.00)},
            'V': {'2s': (610.0, 620.0),'2p': (510.0, 530.0),'3s': (65.0, 75.0),'3p': (35.0, 45.0)},
            'Cr': {'2s': (685.0, 695.0),'2p': (570.0, 590.0),'3s': (70.0, 80.0),'3p': (40.0, 50.0)},
            'Mn': {'2s': (765.0, 775.0),'2p': (631.0, 661.0),'3s': (80.0, 90.0),'3p': (45.0, 55.0)},
            'Fe': {'2s': (845.0, 855.0),'2p': (700.0, 745.0)},
            'Co': {'2s': (925.0, 935.0),'2p': (775.0, 795.0),'3s': (100.0, 110.0),'3p': (55.0, 65.0)},
            'Ni': {'2s': (1005.0, 1015.0),'2p': (850.0, 870.0),'3s': (110.0, 120.0),'3p': (65.0, 75.0)},
            'Cu': {'2s': (1090.0, 1100.0),'2p': (930.0, 950.0),'3s': (120.0, 130.0),'3p': (70.0, 80.0)},
            'Zn': {'2s': (1190.0, 1200.0),'2p': (1015.0, 1025.0),'3s': (135.0, 145.0),'3p': (85.0, 95.0)},

            'Ga': {'2s': (1298.00, 1300.00), '2p': (1116.00, 1120.00), '3s': (159.00, 161.00), '3p': (103.00, 105.00)},
            'Ge': {'2p': (1217.00, 1221.00), '3s': (180.00, 182.00), '3p': (120.00, 122.00)},
            'As': {'2p': (1323.00, 1327.00), '3s': (204.00, 206.00), '3p': (141.00, 143.00)},
            'Se': {'3s': (229.00, 231.00), '3p': (160.00, 162.00), '3d': (55.00, 59.00)},
            'Br': {'3s': (257.00, 259.00), '3p': (182.00, 184.00),'3d': (67.00, 71.00)},
            'Sr': {'3s': (358.00, 360.00), '3p': (269.00, 271.00),'3d': (132.00, 136.00)},
            'Y': {'3s': (392.00, 394.00), '3p': (298.00, 300.00),'3d': (158.00, 160.00)},
            'Nb': {'3s': (468.00, 470.00), '3p': (360.00, 362.00),'3d': (202.00, 204.00)},
            'Zr': {'3s': (430.00, 432.00), '3p': (343.00, 345.00),'3d': (177.00, 181.00)},
            'Mo': {'3s': (506.00, 508.00), '3p': (412.00, 414.00),'3d': (226.00, 230.00)},
            'Ru': {'3s': (586.00, 588.00), '3p': (461.00, 463.00),'3d': (280.00, 284.00)},
            'Rh': {'3s': (628.00, 630.00), '3p': (496.00, 498.00),'3d': (307.00, 311.00)},
            'Pd': {'3s': (671.00, 673.00), '3p': (532.00, 534.00),'3d': (335.00, 339.00)},
            'Ag': {'3s': (365.0, 375.0), '3p': (365.0, 375.0), '3d': (358.0, 385.0)},
            'Cd': {'3s': (772.00, 774.00), '3p': (618.00, 620.00),'3d': (405.00, 409.00)},
            'In': {'3s': (820.0, 830.0),'3p': (660.0, 670.0),'3d': (440.0, 450.0)},
            'Sn': {'3s': (880.0, 890.0),'3p': (710.0, 720.0),'3d': (480.0, 490.0)},
            'Sb': {'3s': (940.0, 950.0),'3p': (760.0, 770.0),'3d': (525.0, 535.0)},
            'Te': {'3s': (1006.00, 1008.00), '3p': (820.00, 822.00),'3d': (573.00, 577.00)},
            'I': {'3s': (1065.0, 1075.0),'3p': (870.0, 880.0),'3d': (615.0, 625.0)},

            # Heavy elements (removed high energy orbitals > 9000eV)
            'Cs': {'3s': (1217.00, 1219.00),'3p': (1065.00, 1067.00), '3d': (726.00, 740.00)},
            'Ba': {'3s': (1293.00, 1295.00),'3p': (1137.00, 1139.00), '3d': (780.00, 795.00)},

            # Lanthanides (removed high energy orbitals > 9000eV)
            'La': {'3p': (1209.00, 1211.00), '3d': (836.00, 851.00)},
            'Ce': {'3p': (1274.00, 1276.00), '3d': (879.00, 915.00)},
            'Pr': {'3d': (932.00, 950.00)},
            'Nd': {'3d': (981.00, 1000.00)},
            'Pm': {'3d': (1027.00, 1048.00)},
            'Sm': {'3d': (1081.00, 1101.00)},
            'Eu': {'3d': (1131.00, 1151.00)},
            'Gd': {'3d': (1185.00, 1205.00)},
            'Tb': {'3d': (1241.00, 1261.00)},

            # Heavy transition metals (< 9000eV only)
            'Hf': {'4s': (660.00, 665.00),'4p': (380.00, 385.00), '4d': (220.00, 225.00), '4f': (14.00, 17.00)},
            'Ta': {'4s': (708.00, 713.00),'4p': (405.00, 410.00), '4d': (230.00, 235.00), '4f': (22.00, 25.00)},
            'W': {'4s': (756.00, 761.00),'4p': (423.00, 428.00), '4d': (243.00, 248.00), '4f': (31.00, 34.00)},
            'Re': {'4s': (625.00, 630.00),
                   '4p': (445.00, 450.00), '4d': (255.00, 260.00), '4f': (40.00, 43.00)},
            'Os': {'4s': (658.00, 663.00),'4p': (470.00, 475.00), '4d': (278.00, 283.00), '4f': (51.00, 54.00)},
            'Ir': {'4s': (691.00, 696.00),'4p': (495.00, 500.00), '4d': (311.00, 316.00), '4f': (61.00, 64.00)},
            'Pt': {'4s': (725.00, 730.00),'4p': (519.00, 524.00), '4d': (314.00, 319.00), '4f': (71.00, 74.00)},
            'Au': {'4s': (745.0, 755.0),'4p': (545.0, 555.0),'4f': (79.0, 95.0)},
            'Hg': {'4s': (800.00, 805.00),'4p': (572.00, 577.00), '4d': (358.00, 363.00), '4f': (101.00, 104.00)},
            'Tl': {'4s': (846.00, 851.00),'4p': (608.00, 613.00), '4d': (386.00, 391.00), '4f': (118.00, 121.00)},
            'Pb': {'4s': (894.00, 899.00),'4p': (644.00, 649.00), '4d': (412.00, 417.00), '4f': (137.00, 140.00)},
            'Bi': {'4s': (938.00, 943.00),'4p': (680.00, 685.00), '4d': (440.00, 445.00), '4f': (157.00, 160.00)}
        }

        # Filter out high energy orbitals
        filtered_db = {}
        for element, orbitals in elements_db.items():
            filtered_orbitals = {}
            for orbital, (be_min, be_max) in orbitals.items():
                if be_min <= max_energy:
                    filtered_orbitals[orbital] = (be_min, be_max)

            if filtered_orbitals:
                filtered_db[element] = filtered_orbitals

        return filtered_db