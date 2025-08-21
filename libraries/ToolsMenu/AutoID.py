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
            'O': '1s', 'C': '1s', 'N': '1s', 'Na': ['1s', 'kll'], 'F': '1s', 'Mg': '1s',
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
            all_matches = self.find_all_element_matches(peak['position'], element_db, tolerance=6.0)

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

    def find_all_element_matches(self, peak_position, element_db, tolerance=6.0):
        """Find all matching elements for given peak position, sorted by best fit"""
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

        # ALWAYS reset background data structure to raw data (delete any existing background)
        if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
            self.parent.Data['Core levels'][sheet_name]['Background'] = {}

        # ALWAYS reset background to raw data - this clears any existing background processing
        self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = y_data.tolist()
        print(f"  Background reset to raw data for clean AutoID processing")

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

        # Always use the current background data (which was reset in create_peaks_and_measure)
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

    def calculate_confidence_score(self, peak, assignment_type, companion_found=False, distance_to_expected=0.0):
        """Calculate confidence score for peak assignment"""
        confidence = 0.0

        # Base score from prominence (0-40 points)
        prominence_score = min(40, peak['prominence'] * 1000)  # Scale prominence
        confidence += prominence_score

        # Distance penalty (0-30 points, higher is better for closer matches)
        distance_score = max(0, 30 - distance_to_expected * 5)
        confidence += distance_score

        # Companion bonus (0-30 points)
        if companion_found:
            confidence += 30

        # Assignment type bonus
        if assignment_type == "priority_1":
            confidence += 10
        elif assignment_type == "priority_2":
            confidence += 5

        return min(100, confidence)  # Cap at 100

    def find_auger_companion(self, element, main_peak_position, remaining_peaks, tolerance=4.0):
        """Find Auger companion peaks (like KLL for 1s)"""
        elements_db = self.get_core_level_ranges()

        if element not in elements_db:
            return None

        # Look for Auger transitions
        for orbital, (be_min, be_max) in elements_db[element].items():
            if 'kll' in orbital.lower():
                be_center = (be_min + be_max) / 2

                # Find peak near expected Auger position
                for peak in remaining_peaks:
                    if abs(peak['position'] - be_center) <= tolerance:
                        return {
                            'peak': peak,
                            'orbital': orbital,
                            'expected_position': be_center,
                            'distance': abs(peak['position'] - be_center)
                        }
        return None

    def find_companion_orbital(self, element, main_orbital, remaining_peaks, tolerance=4.0):
        """Find companion orbital for given element and main orbital"""
        companion_map = {
            '2p': '2s',
            '3d': '3p',
            '4f': '4d'
        }

        companion_orbital = companion_map.get(main_orbital)
        if not companion_orbital:
            return None

        elements_db = self.get_core_level_ranges()

        if element not in elements_db or companion_orbital not in elements_db[element]:
            return None

        be_min, be_max = elements_db[element][companion_orbital]
        be_center = (be_min + be_max) / 2

        # Find peak near expected companion position
        for peak in remaining_peaks:
            if abs(peak['position'] - be_center) <= tolerance:
                return {
                    'peak': peak,
                    'orbital': companion_orbital,
                    'expected_position': be_center,
                    'distance': abs(peak['position'] - be_center)
                }
        return None

    def find_low_energy_validation(self, element, remaining_peaks, max_energy=70.0):
        """Find validation peaks below specified energy for element confirmation"""
        elements_db = self.get_core_level_ranges()

        if element not in elements_db:
            return []

        validation_peaks = []

        for orbital, (be_min, be_max) in elements_db[element].items():
            be_center = (be_min + be_max) / 2

            if be_center <= max_energy:
                # Look for peaks in this energy range
                for peak in remaining_peaks:
                    if abs(peak['position'] - be_center) <= 3.0:  # Tighter tolerance for validation
                        validation_peaks.append({
                            'peak': peak,
                            'orbital': orbital,
                            'expected_position': be_center,
                            'distance': abs(peak['position'] - be_center)
                        })

        return validation_peaks

    def get_possible_assignments(self, peak_position, tolerance=6.0):
        """Get all possible core level assignments for a peak position, sorted by likelihood"""
        elements_db = self.get_core_level_ranges()
        possible = []

        for element, orbitals in elements_db.items():
            for orbital, (be_min, be_max) in orbitals.items():
                be_center = (be_min + be_max) / 2
                distance = abs(peak_position - be_center)

                if distance <= tolerance:
                    # Calculate likelihood based on priority and distance
                    priority_score = 0
                    if element in self.priority_1_elements:
                        priority_score = 100
                    elif element in self.priority_2_elements:
                        priority_score = 50
                    elif element in self.priority_3_elements:
                        priority_score = 25
                    else:
                        priority_score = 1

                    likelihood = priority_score - distance

                    possible.append({
                        'element': element,
                        'orbital': orbital,
                        'position': be_center,
                        'distance': distance,
                        'likelihood': likelihood,
                        'assignment': f"{element}{orbital}"
                    })

        # Sort by likelihood (highest first)
        possible.sort(key=lambda x: x['likelihood'], reverse=True)
        return possible


    def get_core_level_ranges(self, max_energy=1400.0):
        """Hardcoded database of core levels with binding energy ranges for fine-tuning"""

        """Get core level ranges, excluding orbitals above max_energy eV"""
        elements_db = {
            # Light elements
            'C': {'1s': (276.0, 295.0),'kll' :(1205,1235) , '2s': (20.0, 25.0)},        # DONE
            'N': {'1s': (395.0, 405.0),'2s': (25.0, 30.0)},
            'O': {'1s': (521.0, 538.0), 'kll': (969,980) ,'2s': (35.0, 45.0)},          # DONE
            'F': {'1s': (675.00, 700.00), '2s': (45.0, 55.0)},
            'Na': {'1s': (1060.50, 1081.50), 'kll': (491.00, 501.00)},                  # DONE
            'Mg': {'1s': (1292.50, 1317.50), '2s': (85.0, 95.0), '2p': (45.0, 55.0)},
            'Al': {'2s': (115.0, 125.0), '2p': (70.0, 78.0)},
            'Si': {'2s': (145.0, 155.0), '2p': (98.0, 106.0)},  # DONE
            'P': {'2s': (180.0, 200.0), '2p': (125.0, 141.0)},
            'S': {'2s': (217.50, 242.50), '2p': (158.0, 170.0), 'lmm': (1330, 1350)},   # DONE
            'Cl': {'2s': (257.50, 282.50), '2p': (187.50, 212.50)},
            'K': {'2s': (367.50, 392.50), '2p': (282.50, 307.50)},
            'Ca': {'lmm':(1187,1208),'2s': (428.0, 452.0), '2p': (343.0, 362.0)},       # DONE

            # Transition metals
            'Ti': {'2s': (551.00, 576.00), '2p': (453.00, 478.00), '3s': (60.00, 61.00), '3p': (37.00, 38.00)},
            'V': {'2s': (602.50, 627.50), '2p': (507.50, 532.50), '3s': (65.0, 75.0), '3p': (35.0, 45.0)},
            'Cr': {'2s': (677.50, 702.50), '2p': (567.50, 592.50), '3s': (70.0, 80.0), '3p': (40.0, 50.0)},
            'Mn': {'2s': (757.50, 782.50), '2p': (633.50, 658.50), '3s': (80.0, 90.0), '3p': (45.0, 55.0)},
            'Fe': {'2s': (837.50, 862.50), '2p': (710.00, 735.00)},
            'Co': {'2s': (917.50, 942.50), '2p': (772.50, 797.50), '3s': (100.0, 110.0), '3p': (55.0, 65.0)},
            'Ni': {'2s': (993.0, 1026.0), '2p': (840.0, 885.0), '3s': (110.0, 120.0), '3p': (65.0, 75.0)},  # DONE
            'Cu': {'2s': (1082.50, 1107.50), '2p': (927.50, 952.50), '3s': (120.0, 130.0), '3p': (70.0, 80.0)},
            'Zn': {'2s': (1182.50, 1207.50), '2p': (1007.50, 1032.50), '3s': (135.0, 145.0), '3p': (85.0, 95.0)},

            'Ga': {'2s': (1286.50, 1311.50), '2p': (1105.50, 1130.50), '3s': (159.00, 161.00), '3p': (103.00, 105.00)},
            'Ge': {'2p': (1206.50, 1231.50), '3s': (180.00, 182.00), '3p': (120.00, 122.00)},
            'As': {'2p': (1312.50, 1337.50), '3s': (192.50, 217.50), '3p': (141.00, 143.00)},
            'Se': {'3s': (217.50, 242.50), '3p': (160.00, 162.00), '3d': (55.00, 59.00)},
            'Br': {'3s': (245.50, 270.50), '3p': (182.00, 184.00), '3d': (67.00, 71.00)},
            'Sr': {'3s': (346.50, 371.50), '3p': (257.50, 282.50), '3d': (132.00, 136.00)},
            'Y': {'3s': (380.50, 405.50), '3p': (286.50, 311.50), '3d': (158.00, 160.00)},
            'Nb': {'3s': (456.50, 481.50), '3p': (348.50, 373.50), '3d': (190.50, 215.50)},
            'Zr': {'3s': (418.50, 443.50), '3p': (331.50, 356.50), '3d': (177.00, 181.00)},
            'Mo': {'3s': (494.00, 520.00), '3p': (378.00, 425.00), '3d': (221.00, 238.00)},  # DONE
            'Ru': {'3s': (574.50, 599.50), '3p': (449.50, 474.50), '3d': (269.50, 294.50)},
            'Rh': {'3s': (616.50, 641.50), '3p': (484.50, 509.50), '3d': (296.50, 321.50)},
            'Pd': {'3s': (659.50, 684.50), '3p': (520.50, 545.50), '3d': (324.50, 349.50)},
            'Ag': {'3s': (357.50, 382.50), '3p': (357.50, 382.50), '3d': (359.00, 384.00)},
            'Cd': {'3s': (760.50, 785.50), '3p': (606.50, 631.50), '3d': (394.50, 419.50)},
            'In': {'3s': (812.50, 837.50), '3p': (652.50, 677.50), '3d': (432.50, 457.50)},
            'Sn': {'3s': (872.50, 897.50), '3p': (702.50, 727.50), '3d': (472.50, 497.50)},
            'Sb': {'3s': (932.50, 957.50), '3p': (752.50, 777.50), '3d': (517.50, 542.50)},
            'Te': {'3s': (994.50, 1019.50), '3p': (808.50, 833.50), '3d': (562.50, 587.50)},
            'I': {'3s': (1057.50, 1082.50), '3p': (862.50, 887.50), '3d': (607.50, 632.50)},

            # Heavy elements (removed high energy orbitals > 9000eV)
            'Cs': {'3s': (1205.50, 1230.50), '3p': (1053.50, 1078.50), '3d': (720.50, 745.50)},
            'Ba': {'3s': (1281.50, 1306.50), '3p': (1125.50, 1150.50), '3d': (775.00, 800.00)},

            # Lanthanides (removed high energy orbitals > 9000eV)
            'La': {'3p': (1197.50, 1222.50), '3d': (831.00, 856.00)},
            'Ce': {'3p': (1262.50, 1287.50), '3d': (884.50, 909.50)},
            'Pr': {'3d': (928.50, 953.50)},
            'Nd': {'3d': (978.00, 1003.00)},
            'Pm': {'3d': (1025.00, 1050.00)},
            'Sm': {'3d': (1078.50, 1103.50)},
            'Eu': {'3d': (1128.50, 1153.50)},
            'Gd': {'3d': (1182.50, 1207.50)},
            'Tb': {'3d': (1238.50, 1263.50)},

            # Heavy transition metals (< 9000eV only)
            'Hf': {'4s': (650.00, 675.00), '4p': (370.00, 395.00), '4d': (210.00, 235.00), '4f': (14.00, 17.00)},
            'Ta': {'4s': (698.00, 723.00), '4p': (395.00, 420.00), '4d': (220.00, 245.00), '4f': (22.00, 25.00)},
            'W': {'4s': (746.00, 771.00), '4p': (413.00, 438.00), '4d': (233.00, 258.00), '4f': (31.00, 34.00)},
            'Re': {'4s': (615.00, 640.00),
                   '4p': (435.00, 460.00), '4d': (245.00, 270.00), '4f': (40.00, 43.00)},
            'Os': {'4s': (648.00, 673.00), '4p': (460.00, 485.00), '4d': (268.00, 293.00), '4f': (51.00, 54.00)},
            'Ir': {'4s': (681.00, 706.00), '4p': (485.00, 510.00), '4d': (301.00, 326.00), '4f': (61.00, 64.00)},
            'Pt': {'4s': (715.00, 740.00), '4p': (509.00, 534.00), '4d': (304.00, 329.00), '4f': (71.00, 74.00)},
            'Au': {'4s': (737.50, 762.50), '4p': (537.50, 562.50), '4f': (79.0, 95.0)},
            'Hg': {'4s': (790.00, 815.00), '4p': (562.00, 587.00), '4d': (348.00, 373.00), '4f': (101.00, 104.00)},
            'Tl': {'4s': (836.00, 861.00), '4p': (598.00, 623.00), '4d': (376.00, 401.00), '4f': (118.00, 121.00)},
            'Pb': {'4s': (884.00, 909.00), '4p': (634.00, 659.00), '4d': (402.00, 427.00), '4f': (137.00, 140.00)},
            'Bi': {'4s': (928.00, 953.00), '4p': (670.00, 695.00), '4d': (430.00, 455.00), '4f': (157.00, 160.00)}
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


class AutoIDWindow(wx.Frame):
    """Auto ID configuration and step-by-step identification window"""

    def __init__(self, parent):
        super().__init__(None, title="Automatic Element Identification -- Beta --",
                         style=wx.DEFAULT_FRAME_STYLE)
        self.parent = parent
        self.auto_survey_id = AutoSurveyID(parent)

        # Parameters for peak finding
        self.prominence = 0.01
        self.width = 0.6
        self.width_max = 70.0
        self.distance = 5.0
        self.tolerance = 12

        # Step data storage
        self.all_peaks = []  # All found peaks
        self.assigned_peaks = set()  # Track assigned peak indices
        self.step_assignments = [{} for _ in range(5)]  # Assignments for each step
        self.peak_widths = {}  # Store calculated peak widths

        self.init_ui()
        self.position_window()

    def init_ui(self):
        """Initialize the user interface"""
        panel = wx.Panel(self)
        panel.SetBackgroundColour(wx.Colour(250, 250, 230))

        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Method selection
        method_box = wx.StaticBox(panel, label="Identification Method")
        method_sizer = wx.StaticBoxSizer(method_box, wx.VERTICAL)

        self.method_choice = wx.Choice(panel, choices=["Improved Systematic Method"])
        self.method_choice.SetSelection(0)
        method_sizer.Add(self.method_choice, 0, wx.EXPAND | wx.ALL, 5)

        # Peak finding parameters - MORE COMPACT (2 per row)
        param_box = wx.StaticBox(panel, label="Peak Finding Parameters")
        param_sizer = wx.StaticBoxSizer(param_box, wx.VERTICAL)

        param_grid = wx.FlexGridSizer(3, 4, 5, 5)  # 3 rows, 4 columns
        param_grid.AddGrowableCol(1)
        param_grid.AddGrowableCol(3)

        # Row 1: Prominence and Width (min)
        param_grid.Add(wx.StaticText(panel, label="Prominence:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.prominence_ctrl = wx.TextCtrl(panel, value=f"{self.prominence:.2f}")
        param_grid.Add(self.prominence_ctrl, 0, wx.EXPAND)

        param_grid.Add(wx.StaticText(panel, label="Width (min):"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.width_ctrl = wx.TextCtrl(panel, value=f"{self.width:.2f}")
        param_grid.Add(self.width_ctrl, 0, wx.EXPAND)

        # Row 2: Width (max) and Distance
        param_grid.Add(wx.StaticText(panel, label="Width (max):"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.width_max_ctrl = wx.TextCtrl(panel, value=f"{self.width_max:.1f}")
        param_grid.Add(self.width_max_ctrl, 0, wx.EXPAND)

        param_grid.Add(wx.StaticText(panel, label="Distance:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.distance_ctrl = wx.TextCtrl(panel, value=f"{self.distance:.1f}")
        param_grid.Add(self.distance_ctrl, 0, wx.EXPAND)

        # Row 3: Tolerance
        param_grid.Add(wx.StaticText(panel, label="Tolerance (eV):"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.tolerance_ctrl = wx.TextCtrl(panel, value=f"{self.tolerance:.1f}")
        param_grid.Add(self.tolerance_ctrl, 0, wx.EXPAND)
        param_grid.Add(wx.StaticText(panel, label=""), 0)  # Empty cell
        param_grid.Add(wx.StaticText(panel, label=""), 0)  # Empty cell

        param_sizer.Add(param_grid, 0, wx.EXPAND | wx.ALL, 5)

        # Control buttons
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.run_btn = wx.Button(panel, label="Run Identification")
        self.run_btn.Bind(wx.EVT_BUTTON, self.on_run)
        button_sizer.Add(self.run_btn, 0, wx.ALL, 5)

        self.create_regions_btn = wx.Button(panel, label="Create Background / Area")
        self.create_regions_btn.Bind(wx.EVT_BUTTON, self.on_create_regions)
        self.create_regions_btn.Enable(False)  # Disabled until peaks are identified
        button_sizer.Add(self.create_regions_btn, 0, wx.ALL, 5)

        # SELECT ALL / DESELECT ALL buttons
        self.select_all_btn = wx.Button(panel, label="Select All")
        self.select_all_btn.Bind(wx.EVT_BUTTON, self.on_select_all_global)
        button_sizer.Add(self.select_all_btn, 0, wx.ALL, 5)

        self.deselect_all_btn = wx.Button(panel, label="Deselect All")
        self.deselect_all_btn.Bind(wx.EVT_BUTTON, self.on_deselect_all_global)
        button_sizer.Add(self.deselect_all_btn, 0, wx.ALL, 5)

        # Manual peak addition
        manual_box = wx.StaticBox(panel, label="Manual Peak Addition")
        manual_sizer = wx.StaticBoxSizer(manual_box, wx.HORIZONTAL)

        manual_sizer.Add(wx.StaticText(panel, label="Add Peak at:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        self.manual_peak_ctrl = wx.TextCtrl(panel, value="")
        manual_sizer.Add(self.manual_peak_ctrl, 1, wx.EXPAND | wx.ALL, 5)

        self.add_peak_btn = wx.Button(panel, label="Add Peak")
        self.add_peak_btn.Bind(wx.EVT_BUTTON, self.on_add_manual_peak)
        manual_sizer.Add(self.add_peak_btn, 0, wx.ALL, 5)

        # Step lists area
        lists_box = wx.StaticBox(panel, label="Identification Steps")
        lists_sizer = wx.StaticBoxSizer(lists_box, wx.VERTICAL)

        # Create notebook for step tabs
        self.notebook = wx.Notebook(panel)

        # Create step tabs with new logic
        self.step_pages = []
        step_names = ["1. Find Peaks", "2. Peak Width", "3. Usual's", "4. Most commons",
                      "5. Others", "6. Final Checked", "7. Results"]

        for i, name in enumerate(step_names):
            page = self.create_step_page(self.notebook, name)
            self.notebook.AddPage(page, name)
            self.step_pages.append(page)

        # Set Final tab as active by default
        self.notebook.SetSelection(6)  # Index 6 = Final tab

        lists_sizer.Add(self.notebook, 1, wx.EXPAND | wx.ALL, 5)

        # Status bar
        self.status_text = wx.StaticText(panel, label="Ready to run identification...")

        # Layout
        main_sizer.Add(method_sizer, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(param_sizer, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(manual_sizer, 0, wx.EXPAND | wx.ALL, 5)

        main_sizer.Add(lists_sizer, 1, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(button_sizer, 0, wx.EXPAND | wx.ALL, 5)
        main_sizer.Add(self.status_text, 0, wx.EXPAND | wx.ALL, 5)


        panel.SetSizer(main_sizer)
        self.SetSize((650, 700))

    def create_step_page(self, parent, title):
        """Create a page for a step with enhanced peak list including checkboxes"""
        page = wx.Panel(parent)

        sizer = wx.BoxSizer(wx.VERTICAL)

        # Peak list with reordered columns (no Manual column)
        peak_box = wx.StaticBox(page, label=f"{title} - Peak Analysis")
        peak_sizer = wx.StaticBoxSizer(peak_box, wx.VERTICAL)

        peak_list = wx.ListCtrl(page, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        peak_list.InsertColumn(0, "T", width=20)
        peak_list.InsertColumn(1, "BE (eV)", width=60)
        peak_list.InsertColumn(2, "Width (eV)", width=70)
        peak_list.InsertColumn(3, "Prominence", width=85)
        peak_list.InsertColumn(4, "Confidence", width=85)  # MOVED AFTER PROMINENCE
        peak_list.InsertColumn(5, "Assigned To", width=85)
        peak_list.InsertColumn(6, "Companion", width=85)
        peak_list.InsertColumn(7, "Possibility", width=85)
        # REMOVED Manual column

        peak_sizer.Add(peak_list, 1, wx.EXPAND | wx.ALL, 5)

        # Assignment controls
        assign_box = wx.StaticBox(page, label="Manual Assignment Controls")
        assign_sizer = wx.StaticBoxSizer(assign_box, wx.HORIZONTAL)

        assign_sizer.Add(wx.StaticText(page, label="Force assignment:"), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
        core_level_choice = wx.Choice(page, choices=[])
        assign_sizer.Add(core_level_choice, 1, wx.EXPAND | wx.ALL, 5)

        assign_btn = wx.Button(page, label="Apply")
        assign_sizer.Add(assign_btn, 0, wx.ALL, 5)

        remove_btn = wx.Button(page, label="Remove")
        assign_sizer.Add(remove_btn, 0, wx.ALL, 5)

        # Layout
        sizer.Add(peak_sizer, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(assign_sizer, 0, wx.EXPAND | wx.ALL, 5)

        page.SetSizer(sizer)

        # Store references
        page.peak_list = peak_list
        page.core_level_choice = core_level_choice
        page.assign_btn = assign_btn
        page.remove_btn = remove_btn

        # Bind events
        assign_btn.Bind(wx.EVT_BUTTON, lambda evt: self.on_force_assignment(evt, page))
        remove_btn.Bind(wx.EVT_BUTTON, lambda evt: self.on_remove_peak(evt, page))
        peak_list.Bind(wx.EVT_LIST_ITEM_SELECTED, lambda evt: self.on_peak_selected(evt, page))

        # BIND ENTER KEY TO TOGGLE CHECKBOX
        peak_list.Bind(wx.EVT_KEY_DOWN, lambda evt: self.on_key_down(evt, page))

        return page

    def position_window(self):
        """Position window relative to parent"""
        if self.parent:
            parent_pos = self.parent.GetPosition()
            parent_size = self.parent.GetSize()
            new_x = parent_pos.x + parent_size.width + 10
            new_y = parent_pos.y
            self.SetPosition((new_x, new_y))
        else:
            self.Center()

    def on_key_down(self, event, page):
        """Handle key press events - Enter toggles checkbox"""
        key_code = event.GetKeyCode()

        if key_code == wx.WXK_RETURN or key_code == wx.WXK_NUMPAD_ENTER:
            # Toggle checkbox for selected item
            selected = page.peak_list.GetFirstSelected()
            if selected != wx.NOT_FOUND:
                # Toggle checkbox
                current_text = page.peak_list.GetItem(selected, 0).GetText()
                new_text = "☐" if current_text == "☑" else "☑"
                page.peak_list.SetItem(selected, 0, new_text)

                # Update the peak data
                position_text = page.peak_list.GetItem(selected, 1).GetText()
                try:
                    position = float(position_text)
                    for peak in self.all_peaks:
                        if abs(peak['position'] - position) < 0.1:
                            peak['create_region'] = (new_text == "☑")
                            break
                except ValueError:
                    pass
        else:
            event.Skip()

    def on_select_all_global(self, event):
        """Select all checkboxes in current tab"""
        current_page = self.step_pages[self.notebook.GetSelection()]
        peak_list = current_page.peak_list

        for row in range(peak_list.GetItemCount()):
            peak_list.SetItem(row, 0, "☑")

            # Update peak data
            position_text = peak_list.GetItem(row, 1).GetText()
            try:
                position = float(position_text)
                for peak in self.all_peaks:
                    if abs(peak['position'] - position) < 0.1:
                        peak['create_region'] = True
                        break
            except ValueError:
                pass

    def on_deselect_all_global(self, event):
        """Deselect all checkboxes in current tab"""
        current_page = self.step_pages[self.notebook.GetSelection()]
        peak_list = current_page.peak_list

        for row in range(peak_list.GetItemCount()):
            peak_list.SetItem(row, 0, "☐")

            # Update peak data
            position_text = peak_list.GetItem(row, 1).GetText()
            try:
                position = float(position_text)
                for peak in self.all_peaks:
                    if abs(peak['position'] - position) < 0.1:
                        peak['create_region'] = False
                        break
            except ValueError:
                pass

    def on_run(self, event):
        """Run the improved identification process"""
        try:
            # Update parameters
            self.prominence = float(self.prominence_ctrl.GetValue())
            self.width = float(self.width_ctrl.GetValue())
            self.width_max = float(self.width_max_ctrl.GetValue())
            self.distance = float(self.distance_ctrl.GetValue())
            self.tolerance = float(self.tolerance_ctrl.GetValue())

            # Get survey data
            sheet_name = self.parent.sheet_combobox.GetValue()
            if not any(x in sheet_name.lower() for x in ['survey', 'wide']):
                wx.MessageBox("Please select a Survey or Wide scan sheet", "Info")
                return

            if sheet_name not in self.parent.Data['Core levels']:
                wx.MessageBox("No data found for selected sheet", "Error")
                return

            x_values = np.array(self.parent.Data['Core levels'][sheet_name]['B.E.'])
            y_values_raw = np.array(self.parent.Data['Core levels'][sheet_name]['Raw Data'])

            from scipy.ndimage import gaussian_filter1d
            y_values = gaussian_filter1d(y_values_raw, sigma=1.0)

            # Find all peaks
            self.all_peaks = self.find_peaks_with_params(x_values, y_values)
            self.assigned_peaks = set()
            self.step_assignments = [{} for _ in range(7)]  # Assignments for each step

            # Run the 4-step identification process
            self.run_systematic_identification()

            # Enable the Create Regions button
            self.create_regions_btn.Enable(True)

            self.status_text.SetLabel(f"Identification complete. {len(self.all_peaks)} peaks analyzed.")

        except ValueError as e:
            wx.MessageBox(f"Invalid parameter: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
        except Exception as e:
            wx.MessageBox(f"Identification failed: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def find_peaks_with_params(self, x_data, y_data):
        """Find peaks and calculate peak widths"""
        y_norm = y_data / np.max(y_data)

        peaks, properties = find_peaks(y_norm,
                                       prominence=self.prominence,
                                       width=self.width,
                                       distance=self.distance)

        peaks_found = []
        peaks_dismissed = []  # Track dismissed peaks
        self.peak_widths = {}

        for i, peak_idx in enumerate(peaks):
            # Calculate peak width from properties or estimate
            if 'widths' in properties:
                width_points = properties['widths'][i]
                # Convert width in points to width in eV
                if len(x_data) > 1:
                    point_spacing = abs(x_data[1] - x_data[0])
                    peak_width = width_points * point_spacing
                else:
                    peak_width = 2.0  # Default fallback
            else:
                # Estimate width using FWHM calculation
                peak_width = self.estimate_peak_width(x_data, y_data, peak_idx)

            peak_data = {
                'index': i,
                'position': x_data[peak_idx],
                'intensity': y_data[peak_idx],
                'prominence': properties['prominences'][i],
                'width': peak_width,
                'manual': False,
                'create_region': True  # Default to checked
            }

            # Filter out peaks wider than max width
            if peak_width > self.width_max:
                peak_data['dismissed'] = True
                peak_data['dismiss_reason'] = f"Width {peak_width:.2f}eV > {self.width_max:.1f}eV"
                peaks_dismissed.append(peak_data)
            else:
                peaks_found.append(peak_data)
                self.peak_widths[i] = peak_width

        # Store dismissed peaks for step 2
        self.dismissed_peaks = peaks_dismissed

        # Sort by prominence (highest first) - ALL TABS ORDER BY PROMINENCE
        peaks_found.sort(key=lambda x: x['prominence'], reverse=True)

        # Update indices after sorting
        for i, peak in enumerate(peaks_found):
            peak['index'] = i

        return peaks_found

    def estimate_peak_width(self, x_data, y_data, peak_idx, threshold=0.5):
        """Estimate peak width at half maximum"""
        if peak_idx >= len(y_data) or peak_idx < 0:
            return 2.0  # Default width

        peak_height = y_data[peak_idx]
        half_max = peak_height * threshold

        # Find left side of peak
        left_idx = peak_idx
        while left_idx > 0 and y_data[left_idx] > half_max:
            left_idx -= 1

        # Find right side of peak
        right_idx = peak_idx
        while right_idx < len(y_data) - 1 and y_data[right_idx] > half_max:
            right_idx += 1

        # Calculate width in eV
        if right_idx > left_idx:
            width = abs(x_data[right_idx] - x_data[left_idx])
            return max(0.5, width)  # Minimum width of 0.5 eV
        else:
            return 2.0  # Default fallback

    def run_systematic_identification(self):
        """Run the systematic 6-step identification process"""

        # Step 1: Show all peaks with possibilities (no assignments) - ORDERED BY PROMINENCE
        self.step1_all_peaks()

        # Step 2: Peak dismissal for width filtering
        self.step2_peak_dismissal()

        # Step 3: Priority 1 elements with companions - ORDERED BY PROMINENCE
        self.step3_priority1_with_companions()

        # Step 4: Priority 2&3 elements and validation - ORDERED BY PROMINENCE
        self.step4_priority23_and_validation()

        # Step 5: Remaining unassigned peaks - ORDERED BY PROMINENCE
        self.step5_remaining_peaks()

        # Step 6: Peak assignment dismissal - NEW STEP
        self.step6_peak_assignment_dismissal()

        # Final: Compile all assignments - ORDERED BY PROMINENCE
        self.step7_final_compilation()

    def step1_all_peaks(self):
        """Step 1: List all peaks with possible assignments, ordered by prominence"""
        self.status_text.SetLabel("Step 1: Analyzing all peaks...")

        step1_data = []

        # ALL PEAKS ALREADY SORTED BY PROMINENCE in find_peaks_with_params
        for peak in self.all_peaks:
            # Get possible assignments
            possible = self.auto_survey_id.get_possible_assignments(peak['position'], self.tolerance)
            possible_text = ", ".join([p['assignment'] for p in possible[:5]])

            step1_data.append({
                'peak': peak,
                'assigned': "",
                'confidence': 0,
                'companion': "",
                'possible': possible_text
            })

        self.populate_step_page(0, step1_data)

    def step2_peak_dismissal(self):
        """Step 2: Show dismissed peaks that are too wide"""
        self.status_text.SetLabel("Step 2: Peak dismissal - width filtering...")

        step2_data = []

        if hasattr(self, 'dismissed_peaks'):
            for peak in self.dismissed_peaks:
                step2_data.append({
                    'peak': peak,
                    'assigned': "DISMISSED",
                    'confidence': 0,
                    'companion': peak['dismiss_reason'],
                    'possible': "Too wide for XPS"
                })

        self.populate_step_page(1, step2_data)

    def step3_priority1_with_companions(self):
        """Step 2: Priority 1 elements with companion verification - ordered by prominence"""
        self.status_text.SetLabel("Step 2: Processing Priority 1 elements...")

        step2_data = []
        # Start with peaks ordered by prominence
        remaining_peaks = [p for p in self.all_peaks if p['index'] not in self.assigned_peaks]
        remaining_peaks.sort(key=lambda x: x['prominence'], reverse=True)  # ENSURE PROMINENCE ORDER

        # Process Priority 1 elements
        for peak in remaining_peaks[:]:
            if peak['index'] in self.assigned_peaks:
                continue

            possible = self.auto_survey_id.get_possible_assignments(peak['position'], self.tolerance)

            for possibility in possible:
                element = possibility['element']
                orbital = possibility['orbital']

                # Check if this is a Priority 1 element
                if element in self.auto_survey_id.priority_1_elements:
                    main_orbital = self.auto_survey_id.priority_1_elements[element]

                    if orbital == main_orbital:
                        companion_found = False
                        companion_text = ""
                        confidence = 50  # Base confidence

                        # Look for companions
                        if orbital == '1s':
                            # Look for Auger companion
                            auger = self.auto_survey_id.find_auger_companion(element, peak['position'], remaining_peaks)
                            if auger:
                                companion_found = True
                                companion_text = f"{element}{auger['orbital']}"
                                self.assigned_peaks.add(auger['peak']['index'])
                                confidence = 90

                        elif orbital == '2p':
                            # Look for 2s companion
                            companion = self.auto_survey_id.find_companion_orbital(element, orbital, remaining_peaks)
                            if companion:
                                companion_found = True
                                companion_text = f"{element}{companion['orbital']}"
                                self.assigned_peaks.add(companion['peak']['index'])
                                confidence = 85

                        # Calculate final confidence
                        final_confidence = self.auto_survey_id.calculate_confidence_score(
                            peak, "priority_1", companion_found, possibility['distance']
                        )

                        # Assign this peak
                        if companion_found or peak['prominence'] > 0.02:
                            self.assigned_peaks.add(peak['index'])
                            assignment = f"{element}{orbital}"

                            step2_data.append({
                                'peak': peak,
                                'assigned': assignment,
                                'confidence': int(final_confidence),
                                'companion': companion_text,
                                'possible': f"{element}{orbital} (assigned)"
                            })
                            break

        # SORT STEP2 DATA BY PROMINENCE
        step2_data.sort(key=lambda x: x['peak']['prominence'], reverse=True)
        self.populate_step_page(2, step2_data)

    def step4_priority23_and_validation(self):
        """Step 4: Priority 2&3 elements with validation and main peak checking - ordered by prominence"""
        self.status_text.SetLabel("Step 4: Processing Priority 2&3 elements...")

        step4_data = []
        remaining_peaks = [p for p in self.all_peaks if p['index'] not in self.assigned_peaks]

        # ENSURE PROMINENCE ORDER
        remaining_peaks.sort(key=lambda x: x['prominence'], reverse=True)

        for peak in remaining_peaks[:]:
            if peak['index'] in self.assigned_peaks:
                continue

            possible = self.auto_survey_id.get_possible_assignments(peak['position'], self.tolerance)

            for possibility in possible:
                element = possibility['element']
                orbital = possibility['orbital']

                # Check Priority 2 and 3
                is_priority23 = (element in self.auto_survey_id.priority_2_elements or
                                 element in self.auto_survey_id.priority_3_elements)

                if is_priority23:
                    # Check if main peak required (for 4d and others)
                    main_peak_required = self.requires_main_peak(element, orbital)

                    if main_peak_required:
                        main_orbital = self.get_main_orbital(element, orbital)
                        if not self.main_peak_found(element, main_orbital):
                            continue  # Skip assignment if main peak not found

                    companion_found = False
                    companion_text = ""
                    confidence = 30

                    # Look for companions
                    if orbital == '2p':
                        companion = self.auto_survey_id.find_companion_orbital(element, orbital, remaining_peaks)
                        if companion:
                            companion_found = True
                            companion_text = f"{element}{companion['orbital']}"
                            self.assigned_peaks.add(companion['peak']['index'])
                            confidence = 75

                    # For most prominent peaks, look for low energy validation
                    if peak['prominence'] > 0.015:
                        validation_peaks = self.auto_survey_id.find_low_energy_validation(element, remaining_peaks)
                        if validation_peaks:
                            companion_text += f" +{len(validation_peaks)} validation"
                            confidence += 15

                    # Calculate final confidence
                    final_confidence = self.auto_survey_id.calculate_confidence_score(
                        peak, "priority_23", companion_found, possibility['distance']
                    )

                    # Assign if confidence is high enough
                    if final_confidence > 50 or companion_found:
                        self.assigned_peaks.add(peak['index'])
                        assignment = f"{element}{orbital}"

                        step4_data.append({
                            'peak': peak,
                            'assigned': assignment,
                            'confidence': int(final_confidence),
                            'companion': companion_text,
                            'possible': f"{element}{orbital} (assigned)"
                        })
                        break

        # SORT STEP4 DATA BY PROMINENCE
        step4_data.sort(key=lambda x: x['peak']['prominence'], reverse=True)
        self.populate_step_page(3, step4_data)

    def step5_remaining_peaks(self):
        """Step 4: Handle remaining unassigned peaks - ordered by prominence"""
        self.status_text.SetLabel("Step 4: Processing remaining peaks...")

        step4_data = []
        remaining_peaks = [p for p in self.all_peaks if p['index'] not in self.assigned_peaks]

        # ENSURE PROMINENCE ORDER
        remaining_peaks.sort(key=lambda x: x['prominence'], reverse=True)

        for peak in remaining_peaks:
            possible = self.auto_survey_id.get_possible_assignments(peak['position'], self.tolerance)
            possible_text = ", ".join([p['assignment'] for p in possible[:5]])

            # For remaining peaks, assign tentatively if prominence is high
            assignment = ""
            confidence = 0

            if possible and peak['prominence'] > 0.005:
                best_match = possible[0]
                assignment = best_match['assignment']
                confidence = int(self.auto_survey_id.calculate_confidence_score(
                    peak, "remaining", False, best_match['distance']
                ))

            step4_data.append({
                'peak': peak,
                'assigned': assignment,
                'confidence': confidence,
                'companion': "",
                'possible': possible_text
            })

        self.populate_step_page(4, step4_data)

    def step6_peak_assignment_dismissal(self):
        """Step 6: Dismiss non-main orbital assignments if main peak not found"""
        self.status_text.SetLabel("Step 6: Peak assignment dismissal...")

        step6_data = []
        dismissed_assignments = []

        # Define which orbitals require main orbitals - FIXED MAP
        secondary_to_main_map = {
            '3p': '3d',  # Ba3p requires Ba3d
            '4d': '4f',  # Ir4d requires Ir4f
            '2s': '2p',  # F2s requires F2p (NOT F1s)
            '3s': '3p',  # 3s requires 3p
            '4p': '4d',  # 4p requires 4d - THIS CATCHES Au4p
            '5d': '5f',  # 5d requires 5f
            '2p1': '2p3',  # 2p1/2 requires 2p3/2
        }

        # Collect all currently assigned peaks with their assignments
        assigned_elements_orbitals = set()
        print("=== Step 6: Collecting all assignments ===")

        # Check steps 2, 3, 4 (Usual's, Most commons, Others)
        for step_idx in [2, 3, 4]:
            if step_idx >= len(self.step_pages):
                continue

            page = self.step_pages[step_idx]
            peak_list = page.peak_list
            step_name = ["", "", "Usual's", "Most commons", "Others"][step_idx]

            for row in range(peak_list.GetItemCount()):
                assigned = peak_list.GetItem(row, 5).GetText()  # Column 5 = Assigned To
                if assigned and assigned != "" and assigned != "DISMISSED":
                    assigned_elements_orbitals.add(assigned)
                    print(f"  Found assignment: {assigned} (from {step_name})")

        print(f"  Total assignments found: {assigned_elements_orbitals}")

        # Now check assignments for dismissal and MODIFY original lists
        # Now check assignments for dismissal and MODIFY original lists
        print("=== Step 6: Starting dismissal checks ===")
        for step_idx in [2, 3, 4]:
            if step_idx >= len(self.step_pages):
                continue

            page = self.step_pages[step_idx]
            peak_list = page.peak_list
            step_name = ["", "", "Usual's", "Most commons", "Others"][step_idx]

            print(f"  Checking step {step_idx} ({step_name}) - {peak_list.GetItemCount()} rows")

            for row in range(peak_list.GetItemCount()):
                assigned = peak_list.GetItem(row, 5).GetText()  # Column 5 = Assigned To

                if assigned and assigned != "" and assigned != "DISMISSED":

                    # Debug the assignment string
                    print(f"    Row {row}: assigned='{assigned}', len={len(assigned)}")

                    # Simple parsing - find where digits start
                    digit_start = -1
                    for i, char in enumerate(assigned):
                        if char.isdigit():
                            digit_start = i
                            break

                    if digit_start > 0:
                        element = assigned[:digit_start]
                        orbital = assigned[digit_start:]

                        # Normalize orbital (remove /2 variations)
                        orbital_normalized = orbital.replace('/2', '').replace('3/2', '3').replace('1/2', '1')

                        print(f"    ✓ Parsed: element='{element}', orbital='{orbital_normalized}'")

                        # Check if this orbital requires a main orbital
                        if orbital_normalized in secondary_to_main_map:
                            main_orbital = secondary_to_main_map[orbital_normalized]
                            main_assignment = f"{element}{main_orbital}"

                            print(f"      Secondary orbital '{orbital_normalized}' requires main: '{main_assignment}'")

                            # Check if main orbital has been assigned
                            main_found = False
                            for existing in assigned_elements_orbitals:
                                # Parse existing assignment
                                existing_digit_start = -1
                                for i, char in enumerate(existing):
                                    if char.isdigit():
                                        existing_digit_start = i
                                        break

                                if existing_digit_start > 0:
                                    existing_element = existing[:existing_digit_start]
                                    existing_orbital = existing[existing_digit_start:].replace('/2', '').replace('3/2',
                                                                                                                 '3').replace(
                                        '1/2', '1')

                                    if existing_element == element and existing_orbital == main_orbital:
                                        main_found = True
                                        print(f"        ✓ Found main orbital: {existing}")
                                        break

                            if not main_found:
                                print(f"      ✗ Main orbital '{main_assignment}' NOT found - DISMISSING {assigned}")

                                # MODIFY the original list - mark as DISMISSED
                                peak_list.SetItem(row, 5, "DISMISSED")  # Update Assigned To column
                                peak_list.SetItem(row, 6, f"No {main_assignment} found")  # Update Companion

                                # Color it red
                                peak_list.SetItemBackgroundColour(row, wx.Colour(255, 200, 200))

                                # Track dismissal
                                dismissed_assignments.append(assigned)

                                # Remove from assigned peaks set and add to step6 data
                                position_text = peak_list.GetItem(row, 1).GetText()
                                position = float(position_text)
                                for peak in self.all_peaks:
                                    if abs(peak['position'] - position) < 0.1:
                                        self.assigned_peaks.discard(peak['index'])
                                        step6_data.append({
                                            'peak': peak,
                                            'assigned': "DISMISSED",
                                            'confidence': 0,
                                            'companion': f"No {main_assignment} found",
                                            'possible': f"{assigned} (dismissed - missing main peak)"
                                        })
                                        break
                            else:
                                print(f"      ✓ Main orbital '{main_assignment}' found - keeping {assigned}")
                        else:
                            print(f"      ℹ Orbital '{orbital_normalized}' doesn't require main peak")
                    else:
                        print(f"    ✗ No digits found in assignment: {assigned}")

        print("=== Step 6: Dismissal checks complete ===")

        # SORT STEP6 DATA BY PROMINENCE
        step6_data.sort(key=lambda x: x['peak']['prominence'], reverse=True)
        self.populate_step_page(5, step6_data)  # Tab 5: "6. Final Checked"

        if dismissed_assignments:
            print(f"=== Step 6: Dismissed {len(dismissed_assignments)} assignments ===")
            for assignment in dismissed_assignments:
                print(f"  Dismissed: {assignment}")
        else:
            print("=== Step 6: No assignments dismissed ===")

    def step7_final_compilation(self):
        """Step 7: Final compilation of all assignments - ordered by prominence"""
        self.status_text.SetLabel("Step 7: Compiling final results...")

        final_data = []
        seen_positions = set()  # Track positions to avoid duplicates

        # Collect all assigned peaks from all steps (skip step 1 - peak dismissal and step 6 - assignment dismissal)
        for step_idx in range(7):  # Now we have 7 total steps (0-6)
            if step_idx in [1, 5]:  # Skip peak dismissal (step 1) and assignment dismissal (step 5)
                continue

            if step_idx >= len(self.step_pages):
                continue

            page = self.step_pages[step_idx]
            peak_list = page.peak_list

            for row in range(peak_list.GetItemCount()):
                # FIXED COLUMN INDICES:
                # 0: Create, 1: Position, 2: Width, 3: Prominence, 4: Confidence, 5: Assigned To, 6: Companion, 7: Possible
                assigned = peak_list.GetItem(row, 5).GetText()  # Column 5 = Assigned To

                if assigned and assigned != "" and assigned != "DISMISSED":
                    position_text = peak_list.GetItem(row, 1).GetText()
                    position = float(position_text)

                    # Avoid duplicates - skip if we've already seen this position
                    position_key = round(position, 1)  # Round to avoid floating point issues
                    if position_key in seen_positions:
                        continue
                    seen_positions.add(position_key)

                    width = peak_list.GetItem(row, 2).GetText()
                    prominence = peak_list.GetItem(row, 3).GetText()
                    confidence_text = peak_list.GetItem(row, 4).GetText()  # Column 4 = Confidence
                    companion = peak_list.GetItem(row, 6).GetText()  # Column 6 = Companion

                    # Find the original peak data
                    original_peak = None
                    for peak in self.all_peaks:
                        if abs(peak['position'] - position) < 0.1:
                            original_peak = peak
                            break

                    if original_peak:
                        confidence = int(confidence_text) if confidence_text.isdigit() else 0

                        # Set create_region based on confidence - low confidence peaks appear but aren't ticked
                        original_peak['create_region'] = confidence >= 40  # Only tick if confidence >= 40

                        final_data.append({
                            'peak': original_peak,
                            'assigned': assigned,
                            'confidence': confidence,
                            'companion': companion,
                            'possible': assigned  # Use the assignment as the "possible" text
                        })

        # SORT FINAL DATA BY PROMINENCE (highest first)
        final_data.sort(key=lambda x: x['peak']['prominence'], reverse=True)

        self.populate_step_page(6, final_data)  # Tab 6: "7. Final Results"

    def populate_step_page(self, step_index, data_list):
        """Populate a step page with peak data including checkboxes"""
        if step_index >= len(self.step_pages):
            return

        page = self.step_pages[step_index]
        peak_list = page.peak_list

        # Clear existing items
        peak_list.DeleteAllItems()

        # Add data - ALL DATA ALREADY SORTED BY PROMINENCE
        for i, data in enumerate(data_list):
            peak = data['peak']

            # Check if peak should be ticked based on confidence and create_region setting
            should_tick = peak.get('create_region', True) and data['confidence'] >= 30
            index = peak_list.InsertItem(i, "☑" if should_tick else "☐")
            peak_list.SetItem(index, 1, f"{peak['position']:.2f}")
            peak_list.SetItem(index, 2, f"{peak['width']:.2f}")
            peak_list.SetItem(index, 3, f"{peak['prominence']:.4f}")
            peak_list.SetItem(index, 4, str(data['confidence']))  # CONFIDENCE AFTER PROMINENCE
            peak_list.SetItem(index, 5, data['assigned'])  # ASSIGNED TO
            peak_list.SetItem(index, 6, data['companion'])  # COMPANION
            peak_list.SetItem(index, 7, data['possible'])  # POSSIBLE CORE LEVELS
            # REMOVED Manual column

            # Color code by confidence
            if data['confidence'] >= 80:
                peak_list.SetItemBackgroundColour(index, wx.Colour(200, 255, 200))  # Green
            elif data['confidence'] >= 60:
                peak_list.SetItemBackgroundColour(index, wx.Colour(255, 255, 200))  # Yellow
            elif data['confidence'] >= 40:
                peak_list.SetItemBackgroundColour(index, wx.Colour(255, 230, 200))  # Orange
            elif data['confidence'] > 0:
                peak_list.SetItemBackgroundColour(index, wx.Colour(255, 200, 200))  # Red

    def requires_main_peak(self, element, orbital):
        """Check if orbital requires main peak to be found first"""
        main_peak_map = {
            '4d': '4f',
            '3p': '3d',
            '2s': '2p'
        }
        return orbital in main_peak_map

    def get_main_orbital(self, element, orbital):
        """Get the main orbital for a given orbital"""
        main_peak_map = {
            '4d': '4f',
            '3p': '3d',
            '2s': '2p'
        }
        return main_peak_map.get(orbital, orbital)

    def main_peak_found(self, element, main_orbital):
        """Check if main peak has already been assigned"""
        main_key = f"{element}{main_orbital}"

        # Check in all step assignments
        for step_assignments in self.step_assignments:
            if main_key in step_assignments:
                return True

        # Check in current assignments by looking at assigned peaks
        for peak in self.all_peaks:
            if peak['index'] in self.assigned_peaks:
                possible = self.auto_survey_id.get_possible_assignments(peak['position'], self.tolerance)
                for p in possible:
                    if p['element'] == element and p['orbital'] == main_orbital:
                        return True
        return False

    def on_list_click(self, event, page):
        """Handle clicks on the list to toggle checkboxes"""
        # Get the item that was clicked
        pos = event.GetPosition()
        item, flags = page.peak_list.HitTest(pos)

        if item != wx.NOT_FOUND:
            # Check if click was in the first column (checkbox column)
            rect = page.peak_list.GetItemRect(item)
            if pos.x <= 50:  # First column width
                # Toggle checkbox
                current_text = page.peak_list.GetItem(item, 0).GetText()
                new_text = "☐" if current_text == "☑" else "☑"
                page.peak_list.SetItem(item, 0, new_text)

                # Update the peak data
                position_text = page.peak_list.GetItem(item, 1).GetText()
                try:
                    position = float(position_text)
                    for peak in self.all_peaks:
                        if abs(peak['position'] - position) < 0.1:
                            peak['create_region'] = (new_text == "☑")
                            break
                except ValueError:
                    pass  # Skip if position is not a valid float

        event.Skip()

    def on_select_all(self, event, page):
        """Handle select/deselect all checkbox"""
        select_all = event.GetEventObject().GetValue()
        symbol = "☑" if select_all else "☐"

        peak_list = page.peak_list
        for row in range(peak_list.GetItemCount()):
            peak_list.SetItem(row, 0, symbol)

            # Update peak data
            position = float(peak_list.GetItem(row, 1).GetText())
            for peak in self.all_peaks:
                if abs(peak['position'] - position) < 0.1:
                    peak['create_region'] = select_all
                    break

    def on_checkbox_toggle(self, event, page):
        """Toggle checkbox on double-click or activation"""
        item = event.GetIndex()
        if item != wx.NOT_FOUND:
            # Toggle checkbox
            current_text = page.peak_list.GetItem(item, 0).GetText()
            new_text = "☐" if current_text == "☑" else "☑"
            page.peak_list.SetItem(item, 0, new_text)

            # Update the peak data
            position_text = page.peak_list.GetItem(item, 1).GetText()
            try:
                position = float(position_text)
                for peak in self.all_peaks:
                    if abs(peak['position'] - position) < 0.1:
                        peak['create_region'] = (new_text == "☑")
                        break
            except ValueError:
                pass

    def on_create_regions(self, event):
        """Create background/area regions for selected peaks"""
        try:
            # Get selected peaks from current tab
            current_page = self.step_pages[self.notebook.GetSelection()]
            peak_list = current_page.peak_list

            selected_peaks = []
            for row in range(peak_list.GetItemCount()):
                checkbox = peak_list.GetItem(row, 0).GetText()  # Column 0 = Create checkbox
                if checkbox == "☑":
                    position = float(peak_list.GetItem(row, 1).GetText())  # Column 1 = Position
                    assigned = peak_list.GetItem(row, 5).GetText()  # Column 5 = Assigned To

                    if assigned:  # Only create regions for assigned peaks
                        selected_peaks.append({
                            'position': position,
                            'assignment': assigned
                        })

            if not selected_peaks:
                wx.MessageBox("No assigned peaks selected for region creation", "No Selection",
                              wx.OK | wx.ICON_INFORMATION)
                return

            # Import AutoSurveyID and use its create_peaks_and_measure method
            from libraries.FileMenu.Save import save_state

            # Save state before making changes
            save_state(self.parent)

            # CLEAR ALL EXISTING BACKGROUND DATA AND PEAK DATA
            sheet_name = self.parent.sheet_combobox.GetValue()

            # Clear existing peaks from grid
            if self.parent.peak_params_grid.GetNumberRows() > 0:
                self.parent.peak_params_grid.DeleteRows(0, self.parent.peak_params_grid.GetNumberRows())

            # Clear existing peaks from data structure
            if sheet_name in self.parent.Data['Core levels']:
                if 'Fitting' in self.parent.Data['Core levels'][sheet_name]:
                    if 'Peaks' in self.parent.Data['Core levels'][sheet_name]['Fitting']:
                        self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = {}

                # CLEAR EXISTING BACKGROUND DATA - RESET TO RAW DATA
                if 'Background' in self.parent.Data['Core levels'][sheet_name]:
                    # Reset background to raw data to clear any previous processing
                    raw_data = self.parent.Data['Core levels'][sheet_name]['Raw Data']
                    self.parent.Data['Core levels'][sheet_name]['Background']['Bkg Y'] = raw_data.copy()
                    print(f"=== AutoID: Reset background to raw data ===")

            # Reset peak count
            self.parent.peak_count = 0

            print(f"=== AutoID: Cleared existing peaks and background data ===")

            # Create identified elements dictionary from selected peaks
            identified_elements = {}
            for peak_data in selected_peaks:
                # Parse assignment like "C1s" into element and orbital
                assignment = peak_data['assignment']
                element = ''.join([c for c in assignment if c.isalpha()])
                orbital = ''.join([c for c in assignment if not c.isalpha()])

                identified_elements[assignment] = {
                    'element': element,
                    'orbital': orbital,
                    'peak_position': peak_data['position'],
                    'intensity': 1000.0,  # Default
                    'prominence': 0.01,  # Default
                    'priority': 1
                }

            # Get current data
            x_values = np.array(self.parent.Data['Core levels'][sheet_name]['B.E.'])
            y_values_raw = np.array(self.parent.Data['Core levels'][sheet_name]['Raw Data'])

            from scipy.ndimage import gaussian_filter1d
            y_values = gaussian_filter1d(y_values_raw, sigma=1.0)

            # Create peaks and measure areas - use RAW data for background calculations
            self.auto_survey_id.create_peaks_and_measure(identified_elements, x_values, y_values_raw, sheet_name)

            # Update the main window - PRESERVE VLINES
            self.parent.peak_params_grid.ForceRefresh()

            # Store vLine state before operations
            vlines_visible = hasattr(self.parent.plot_manager,
                                     'vlines_visible') and self.parent.plot_manager.vlines_visible

            # Update plot
            self.parent.plot_manager.plot_data(self.parent)
            self.parent.clear_and_replot()
            self.parent.plot_manager.update_legend(self.parent)

            # Restore vLine state if they were visible
            if vlines_visible:
                if hasattr(self.parent.plot_manager, 'draw_vlines'):
                    self.parent.plot_manager.draw_vlines(self.parent)

            self.parent.canvas.draw_idle()

            wx.MessageBox(f"Created {len(selected_peaks)} peak regions successfully!",
                          "Regions Created", wx.OK | wx.ICON_INFORMATION)

        except Exception as e:
            wx.MessageBox(f"Failed to create regions: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)
            import traceback
            traceback.print_exc()

    def on_add_manual_peak(self, event):
        """Add a manual peak and re-run identification"""
        try:
            position = float(self.manual_peak_ctrl.GetValue())

            manual_peak = {
                'index': len(self.all_peaks),
                'position': position,
                'intensity': 0.0,
                'prominence': 0.01,
                'width': 2.0,  # Default width
                'manual': True,
                'create_region': True
            }

            self.all_peaks.append(manual_peak)
            self.manual_peak_ctrl.SetValue("")

            # Re-run identification
            self.run_systematic_identification()

        except ValueError:
            wx.MessageBox("Please enter a valid binding energy", "Invalid Input", wx.OK | wx.ICON_ERROR)

    def on_peak_selected(self, event, page):
        """Handle peak selection"""
        selected = event.GetIndex()
        if selected >= 0:
            possible_text = page.peak_list.GetItem(selected, 7).GetText()  # Column 7 = Possible Core Levels
            possible_levels = [level.strip() for level in possible_text.split(",") if level.strip()]

            page.core_level_choice.Clear()
            page.core_level_choice.AppendItems(possible_levels)
            if possible_levels:
                page.core_level_choice.SetSelection(0)

    def on_force_assignment(self, event, page):
        """Force assignment of selected peak"""
        selected_peak = page.peak_list.GetFirstSelected()
        if selected_peak >= 0:
            selected_core_level = page.core_level_choice.GetStringSelection()
            if selected_core_level:
                page.peak_list.SetItem(selected_peak, 5, selected_core_level)  # Column 5 = Assigned To
                page.peak_list.SetItem(selected_peak, 4, "100")  # Column 4 = Confidence
                wx.MessageBox(f"Peak manually assigned to {selected_core_level}", "Assignment Updated",
                              wx.OK | wx.ICON_INFORMATION)
        else:
            wx.MessageBox("Please select a peak first", "No Peak Selected", wx.OK | wx.ICON_WARNING)

    def on_remove_peak(self, event, page):
        """Remove selected peak"""
        selected_peak = page.peak_list.GetFirstSelected()
        if selected_peak >= 0:
            position = page.peak_list.GetItem(selected_peak, 1).GetText()  # Column 1 = Position
            result = wx.MessageBox(f"Remove peak at {position} eV?", "Confirm Removal",
                                   wx.YES_NO | wx.ICON_QUESTION)
            if result == wx.YES:
                position_float = float(position)
                self.all_peaks = [p for p in self.all_peaks if abs(p['position'] - position_float) > 0.1]
                self.run_systematic_identification()