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
        self.use_main_library = False  # Default to hardcoded ranges

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

    def get_core_level_ranges_from_library(self, max_energy=1400.0, tolerance=15.0):
        """Get core level ranges from main library data with calculated ranges"""
        elements_db = {}

        for (element, orbital), data in self.library_data.items():
            # Get instrument data
            instrument = self.parent.current_instrument
            if instrument not in data:
                instrument = 'Al1486' if 'Al1486' in data else next(iter(data))

            if 'position' in data[instrument]:
                position = float(data[instrument]['position'])

                if position <= max_energy:
                    if element not in elements_db:
                        elements_db[element] = {}

                    # Create range around position with tolerance
                    be_min = position - tolerance
                    be_max = position + tolerance
                    elements_db[element][orbital] = (be_min, be_max)

        return elements_db

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
            # y_values = y_values_raw
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

    def calculate_decision_value(self, element, orbital, distance):
        """Calculate decision value based on distance and priorities (lower is better)."""
        # Get individual priorities - match priority system from AreaFit_Screen
        common_elements = {
            'C': 1, 'O': 1, 'N': 1, 'Si': 2, 'P': 2, 'Ca': 2, 'Na': 2, 'Fe': 3, 'Cr': 3, 'Ti': 3,
            'Cu': 3, 'Ag': 3, 'Au': 3, 'Zn': 3, 'Co': 3, 'Ni': 3, 'Mn': 3, 'Al': 3, 'Mg': 3,
            'F': 1, 'K': 2, 'S': 2, 'Cl': 2, 'V': 3, 'Sb': 4, 'Ba': 4,
            'Ta': 3, 'W': 3, 'Re': 4, 'Ir': 3, 'Pt': 4, 'La': 3, 'Ce': 3, 'Pr': 4, 'Nd': 4, 'Hf': 4
        }
        priority_element = common_elements.get(element, 4)

        orbital_priorities = {
            '1s': 1, '2p': 1, '3d': 1, '4f': 1, 'kll': 2, '2s': 2, '3p': 2, '4d': 2,
            '3s': 3, '4p': 3, '4s': 4, 'lmm': 3
        }
        priority_orbital = orbital_priorities.get(orbital, 5)

        # Handle division by zero and negative denominators
        orbital_denominator = max(0.1, 5 - priority_orbital)  # Higher priority = larger denominator
        element_denominator = max(0.1, 5 - priority_element)  # Higher priority = larger denominator

        # Calculate decision value (lower is better)
        decision = (distance / orbital_denominator) + (distance / element_denominator)

        # return decision
        return distance


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

    def get_possible_assignments_OLD(self, peak_position, tolerance=6.0):
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

    def get_possible_assignments(self, peak_position, tolerance=6.0):
        """Get all possible core level assignments for a peak position, sorted by decision value"""
        elements_db = self.get_core_level_ranges()
        possible = []

        for element, orbitals in elements_db.items():
            for orbital, (be_min, be_max) in orbitals.items():
                be_center = (be_min + be_max) / 2
                distance = abs(peak_position - be_center)

                if distance <= tolerance:
                    # Calculate decision value using the improved method (lower is better)
                    decision = self.calculate_decision_value(element, orbital, distance)

                    possible.append({
                        'element': element,
                        'orbital': orbital,
                        'position': be_center,
                        'distance': distance,
                        'decision': decision,  # Changed from 'likelihood'
                        'assignment': f"{element}{orbital}"
                    })

        # Sort by decision value (lower is better)
        possible.sort(key=lambda x: x['decision'])
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

            'Ti': {'2s': (554.20, 568.20), '2p': (451.70, 465.70), '2p1/2': (453.20, 467.20), '2p3/2': (447.00, 461.00), '3s': (52.00, 66.00), '3p': (26.40, 40.40), 'LMN3': (1099.70, 1113.70), 'LMN4': (1121.70, 1135.70)},
            'V': {'2s': (602.50, 627.50), '2p': (507.50, 532.50), '2p1/2': (512.90, 526.90), '2p3/2': (505.20, 519.20), '3s': (65.0, 75.0), '3p': (35.0, 45.0), 'LMN1': (974.30, 988.30), 'LMN2': (1011.30, 1025.30), 'LMN3': (1045.80, 1059.80), 'LMN4': (1052.10, 1066.10), 'LMN5': (1072.00, 1086.00), 'LMN6': (1084.10, 1098.10)},
            'Cr': {'2s': (677.50, 702.50), '2p': (567.50, 592.50), '2p1/2': (576.70, 590.70), '2p3/2': (567.30, 581.30), '3s': (70.0, 80.0), '3p': (40.0, 50.0), 'LMN1': (913.10, 927.10), 'LMN2': (955.10, 969.10), 'LMN3': (995.10, 1009.10), 'LMN4': (1025.10, 1039.10), 'LMN5': (1037.10, 1051.10)},
            'Mn': {'2s': (757.50, 782.50), '2p': (633.50, 658.50), '2p1/2': (643.00, 657.00), '2p3/2': (631.80, 645.80), '3s': (80.0, 90.0), '3p': (45.0, 55.0), 'LMN1': (849.10, 863.10), 'LMN2': (897.10, 911.10), 'LMN3': (942.10, 956.10)},
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
            'Mo': {'4s': (52.40, 72.40), '4p': (25.50, 45.50),'3s': (494.00, 520.00), '3p': (378.00, 425.00),
            '3p3/2': (392.00, 396.00), '3p1/2': (409.00,413.00), '3d': (221.00, 238.00)},
            # DONE
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
        self.width_max = 10.0
        self.distance = 60.0
        self.tolerance = 12

        # Step data storage
        self.all_peaks = []  # All found peaks
        self.assigned_peaks = set()  # Track assigned peak indices
        self.step_assignments = [{} for _ in range(5)]  # Assignments for each step
        self.peak_widths = {}  # Store calculated peak widths

        # Initialize highlight tracking
        self.peak_highlight = None
        self.peak_center_line = None

        self.init_ui()
        self.position_window()

        # Bind peak list events AFTER creating all pages
        self.bind_peak_list_events()

        # Bind close event
        self.Bind(wx.EVT_CLOSE, self.OnClose)

    def init_ui(self):
        """Initialize the user interface"""
        panel = wx.Panel(self)
        panel.SetBackgroundColour(wx.Colour(250, 250, 230))

        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Method selection
        method_box = wx.StaticBox(panel, label="Identification Method")
        method_sizer = wx.StaticBoxSizer(method_box, wx.VERTICAL)

        self.method_choice = wx.Choice(panel, choices=["Method 1 GK 25/08/25"])
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
        """Create a page for a step with enhanced peak list including clickable columns and checkbox toggling"""
        page = wx.Panel(parent)

        sizer = wx.BoxSizer(wx.VERTICAL)

        # Peak list with UPDATED column names and sorting capability
        peak_box = wx.StaticBox(page, label=f"{title} - Peak Analysis")
        peak_sizer = wx.StaticBoxSizer(peak_box, wx.VERTICAL)

        peak_list = wx.ListCtrl(page, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        peak_list.InsertColumn(0, "", width=25)  # Tick/Untick column
        peak_list.InsertColumn(1, "BE (eV)", width=60)
        peak_list.InsertColumn(2, "Width (eV)", width=70)
        peak_list.InsertColumn(3, "Signal", width=60)  # ← Changed from "Prominence"
        peak_list.InsertColumn(4, "Score", width=55)  # ← Changed from "Confidence"
        peak_list.InsertColumn(5, "Peak", width=65)  # ← Changed from "Assigned To"
        peak_list.InsertColumn(6, "Companion", width=65)
        peak_list.InsertColumn(7, "Possibility", width=120)
        peak_list.InsertColumn(8, "Distance", width=120)
        peak_list.InsertColumn(9, "Decision", width=120)

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

        # Add sorting state
        page.sort_column = -1
        page.sort_ascending = True
        page.original_data = []  # Store original data for sorting

        # Bind events
        assign_btn.Bind(wx.EVT_BUTTON, lambda evt: self.on_force_assignment(evt, page))
        remove_btn.Bind(wx.EVT_BUTTON, lambda evt: self.on_remove_peak(evt, page))
        peak_list.Bind(wx.EVT_LIST_ITEM_SELECTED, lambda evt: self.on_peak_selected(evt, page))
        peak_list.Bind(wx.EVT_KEY_DOWN, lambda evt: self.on_key_down(evt, page))

        # NEW: Add column click for sorting
        peak_list.Bind(wx.EVT_LIST_COL_CLICK, lambda evt: self.on_column_click(evt, page))

        # NEW: Add left click for checkbox toggling
        peak_list.Bind(wx.EVT_LEFT_UP, lambda evt: self.on_list_click(evt, page))

        return page

    def position_window(self):
        """Position window centered on main frame"""
        if self.parent:
            parent_pos = self.parent.GetPosition()
            parent_size = self.parent.GetSize()
            my_size = self.GetSize()

            # Calculate center position
            new_x = parent_pos.x + (parent_size.width - my_size.width) // 2
            new_y = parent_pos.y + (parent_size.height - my_size.height) // 2

            # Ensure window stays on screen
            new_x = max(0, new_x)
            new_y = max(0, new_y)

            self.SetPosition((new_x, new_y))

            # Keep window on top
            self.SetWindowStyle(self.GetWindowStyle() | wx.STAY_ON_TOP)

            print(f"Positioned AutoID window at center: ({new_x}, {new_y})")

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
            # y_values = y_values_raw

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

    def get_original_possibilities(self, peak, assigned_element_orbital):
        """Get original possibilities with assigned one moved to front"""
        if not hasattr(self, 'original_possibilities'):
            return assigned_element_orbital if assigned_element_orbital else ""

        original = self.original_possibilities.get(peak['index'], "")

        if not original or not assigned_element_orbital:
            return original

        if assigned_element_orbital not in original:
            return original

        # Split possibilities and reorder to put assigned one first
        possibilities = [p.strip() for p in original.split(',')]

        # Remove assigned from list and put it at the front
        if assigned_element_orbital in possibilities:
            possibilities.remove(assigned_element_orbital)

        reordered = [assigned_element_orbital] + possibilities
        return ", ".join(reordered)

    def calculate_peak_metrics(self, peak, assigned_element_orbital=""):
        """Calculate distance and decision values for ALL possibilities of a peak"""

        # Get all possible assignments for this peak
        possible = self.auto_survey_id.get_possible_assignments(peak['position'], self.tolerance)

        if not possible:
            return "N/A", "N/A"

        detailed_possibilities = []

        # Calculate distance and decision for ALL possibilities (up to 5)
        for possibility in possible[:5]:
            assignment = possibility['assignment']
            distance = possibility['distance']

            # Parse assignment to get element and orbital
            element = None
            orbital = None

            # Try digit-based parsing first (Ti2p, Ca2p, etc.)
            digit_start = -1
            for i, char in enumerate(assignment):
                if char.isdigit():
                    digit_start = i
                    break

            if digit_start > 0:
                # Standard photoelectron peak
                element = assignment[:digit_start]
                orbital = assignment[digit_start:]
            else:
                # Check for Auger peaks
                auger_suffixes = ['kll', 'lmm', 'mnn', 'mvv', 'mnv']
                for suffix in auger_suffixes:
                    if assignment.lower().endswith(suffix):
                        element = assignment[:-len(suffix)]
                        orbital = suffix
                        break

            if element and orbital:
                # Calculate decision value using existing method
                decision = self.auto_survey_id.calculate_decision_value(element, orbital, distance)

                # Mark assigned peaks with *
                marker = "*" if assignment == assigned_element_orbital else ""
                detailed_possibilities.append(f"{assignment}{marker}({distance:.2f}/{decision:.2f})")
            else:
                # Fallback if parsing fails
                marker = "*" if assignment == assigned_element_orbital else ""
                detailed_possibilities.append(f"{assignment}{marker}({distance:.2f}/N/A)")

        # Join all detailed possibilities
        combined_info = ", ".join(detailed_possibilities)

        # Return as single combined column for distance, and separate decision values
        distance_values = [f"{p['distance']:.2f}" for p in possible[:5]]
        decision_values = []

        for possibility in possible[:5]:
            assignment = possibility['assignment']
            distance = possibility['distance']

            # Parse and calculate decision (same logic as above)
            element = None
            orbital = None
            digit_start = -1
            for i, char in enumerate(assignment):
                if char.isdigit():
                    digit_start = i
                    break

            if digit_start > 0:
                element = assignment[:digit_start]
                orbital = assignment[digit_start:]
            else:
                auger_suffixes = ['kll', 'lmm', 'mnn', 'mvv', 'mnv']
                for suffix in auger_suffixes:
                    if assignment.lower().endswith(suffix):
                        element = assignment[:-len(suffix)]
                        orbital = suffix
                        break

            if element and orbital:
                decision = self.auto_survey_id.calculate_decision_value(element, orbital, distance)
                decision_values.append(f"{decision:.2f}")
            else:
                decision_values.append("N/A")

        return ", ".join(distance_values), ", ".join(decision_values)
    def step1_all_peaks(self):
        """Step 1: List all peaks with possible assignments, ordered by prominence"""
        self.status_text.SetLabel("Step 1: Analyzing all peaks...")

        step1_data = []

        # Initialize storage for original possibilities
        self.original_possibilities = {}

        # ALL PEAKS ALREADY SORTED BY PROMINENCE in find_peaks_with_params
        for peak in self.all_peaks:

            # FILTER OUT LOW ENERGY PEAKS (below 25 eV)
            if peak['position'] < 25.0:
                print(f"  Skipping peak at {peak['position']:.2f} eV (below 25 eV threshold)")
                continue

            # Get possible assignments
            possible = self.auto_survey_id.get_possible_assignments(peak['position'], self.tolerance)
            possible_text = ", ".join([p['assignment'] for p in possible[:5]])

            # Store original possibilities for this peak
            self.original_possibilities[peak['index']] = possible_text

            step1_data.append({
                'peak': peak,
                'assigned': "",
                'confidence': 0,
                'companion': "",
                'possible': possible_text
            })

        self.populate_step_page(0, step1_data)

    def step2_peak_dismissal(self):
        """Step 2: Element frequency analysis + dismissed peaks"""
        self.status_text.SetLabel("Step 2: Element frequency analysis and peak dismissal...")

        step2_data = []

        # ===== NEW: ELEMENT FREQUENCY ANALYSIS =====
        print("=== Step 2: Element Frequency Analysis ===")

        # Count element occurrences across all peak possibilities
        element_frequency = {}  # element -> {peaks: [peak_info], orbitals: [orbitals]}

        for peak in self.all_peaks:
            if peak['position'] < 20.0:  # Skip low energy peaks
                continue

            # Get possibilities for this peak
            possible = self.auto_survey_id.get_possible_assignments(peak['position'], self.tolerance)

            for possibility in possible[:5]:  # Top 5 possibilities
                assignment = possibility['assignment']

                # Parse element from assignment
                element = None
                digit_start = -1
                for i, char in enumerate(assignment):
                    if char.isdigit():
                        digit_start = i
                        break

                if digit_start > 0:
                    # Standard photoelectron peak (Ti2p, Mo3d, etc.)
                    element = assignment[:digit_start]
                    orbital = assignment[digit_start:]
                else:
                    # Check for Auger peaks
                    auger_suffixes = ['kll', 'lmm', 'mnn', 'mvv', 'mnv']
                    for suffix in auger_suffixes:
                        if assignment.lower().endswith(suffix):
                            element = assignment[:-len(suffix)]
                            orbital = suffix
                            break

                if element:
                    if element not in element_frequency:
                        element_frequency[element] = {'peaks': [], 'orbitals': set(), 'assignments': []}

                    # Avoid counting same element from same peak multiple times
                    peak_already_counted = any(p['index'] == peak['index'] for p in element_frequency[element]['peaks'])
                    if not peak_already_counted:
                        element_frequency[element]['peaks'].append(peak)
                        element_frequency[element]['orbitals'].add(orbital)
                        element_frequency[element]['assignments'].append(assignment)

        # Analyze frequency and assign confidence levels
        high_priority_elements = set()  # 2 occurrences
        certainty_elements = set()  # 3+ occurrences

        print("Element frequency analysis:")
        for element, data in element_frequency.items():
            count = len(data['peaks'])
            orbitals = list(data['orbitals'])

            if count >= 3:
                certainty_elements.add(element)
                priority_level = "CERTAINTY"
                confidence_boost = 30
            elif count >= 2:
                high_priority_elements.add(element)
                priority_level = "HIGH PRIORITY"
                confidence_boost = 20
            else:
                priority_level = "normal"
                confidence_boost = 0

            if count >= 2:
                print(
                    f"  {element}: {count} peaks, orbitals: {orbitals} → {priority_level} (+{confidence_boost} confidence)")

        # Store frequency data for use in later steps
        self.element_frequency_data = {
            'high_priority': high_priority_elements,
            'certainty': certainty_elements,
            'all_frequency': element_frequency
        }

        # ===== FREQUENCY-BASED AUTO ASSIGNMENTS =====
        print("=== Step 2: Frequency-based assignments ===")

        # Auto-assign CERTAINTY elements (3+ occurrences) - ASSIGN ALL PEAKS
        for element in certainty_elements:
            element_data = element_frequency[element]

            # Sort peaks by prominence but assign ALL of them
            element_peaks = sorted(element_data['peaks'], key=lambda x: x['prominence'], reverse=True)
            assigned_count = 0

            for peak in element_peaks:
                if peak['index'] not in self.assigned_peaks:
                    # Get the best assignment for this element from this peak
                    possible = self.auto_survey_id.get_possible_assignments(peak['position'], self.tolerance)

                    best_assignment = None
                    best_distance = float('inf')

                    # Find the best assignment for this element based on distance only
                    for possibility in possible:
                        if possibility['assignment'].startswith(element):
                            if possibility['distance'] < best_distance:
                                best_assignment = possibility['assignment']
                                best_distance = possibility['distance']

                    if best_assignment:
                        # High confidence assignment due to frequency
                        confidence = 95  # Very high confidence for certainty elements
                        assigned_count += 1

                        step2_data.append({
                            'peak': peak,
                            'assigned': best_assignment,
                            'confidence': confidence,
                            'companion': f"CERTAINTY ({len(element_peaks)} peaks) - Peak {assigned_count}",
                            'possible': self.get_original_possibilities(peak, best_assignment)
                        })

                        self.assigned_peaks.add(peak['index'])
                        print(
                            f"  ✓ AUTO-ASSIGNED {best_assignment} at {peak['position']:.2f} eV (CERTAINTY - Peak {assigned_count}/{len(element_peaks)})")
                        # Removed break - now assigns ALL peaks for this element

        # ===== ORIGINAL DISMISSED PEAKS LOGIC =====
        print("=== Step 2: Dismissed peaks (width filtering) ===")

        if hasattr(self, 'dismissed_peaks'):
            for peak in self.dismissed_peaks:
                step2_data.append({
                    'peak': peak,
                    'assigned': "DISMISSED",
                    'confidence': 0,
                    'companion': peak['dismiss_reason'],
                    'possible': self.get_original_possibilities(peak, "DISMISSED")
                })

        # SORT BY PROMINENCE
        step2_data.sort(key=lambda x: x['peak']['prominence'], reverse=True)
        self.populate_step_page(1, step2_data)

    def step3_priority1_with_companions(self):
        """Step 3: Priority 1 elements with companion verification - FIXED for Na"""
        self.status_text.SetLabel("Step 3: Processing Priority 1 elements...")

        step2_data = []
        remaining_peaks = [p for p in self.all_peaks if p['index'] not in self.assigned_peaks]
        remaining_peaks.sort(key=lambda x: x['prominence'], reverse=True)

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

                    # FIXED: Handle Na special case where main_orbital is a list
                    is_main_orbital = False
                    if isinstance(main_orbital, list):
                        # Na case: ['1s', 'kll']
                        is_main_orbital = orbital in main_orbital
                        print(
                            f"  Checking {element}{orbital}: orbital '{orbital}' in {main_orbital} = {is_main_orbital}")
                    else:
                        # Normal case: single orbital string
                        is_main_orbital = orbital == main_orbital

                    if is_main_orbital:
                        companion_found = False
                        companion_text = ""
                        confidence = 50  # Base confidence
                        assignment = f"{element}{orbital}"  # ← DEFINE assignment EARLY

                        # Look for companions
                        if orbital == '1s':
                            # Look for Auger companion
                            auger = self.auto_survey_id.find_auger_companion(element, peak['position'], remaining_peaks)
                            if auger:
                                companion_found = True
                                companion_text = f"{element}{auger['orbital']}"
                                confidence = 90

                                # ADD AUGER PEAK AS SEPARATE ENTRY (Fix for Bug 2)
                                auger_assignment = f"{element}{auger['orbital']}"
                                step2_data.append({
                                    'peak': auger['peak'],
                                    'assigned': auger_assignment,  # ← FIXED: Use auger_assignment
                                    'confidence': 85,
                                    'companion': f"Companion to {assignment}",
                                    'possible': self.get_original_possibilities(auger['peak'], auger_assignment)
                                })
                                self.assigned_peaks.add(auger['peak']['index'])

                        elif orbital == '2p':
                            # Look for 2s companion
                            companion = self.auto_survey_id.find_companion_orbital(element, orbital, remaining_peaks)
                            if companion:
                                companion_found = True
                                companion_text = f"{element}{companion['orbital']}"
                                confidence = 85

                                # ADD COMPANION PEAK AS SEPARATE ENTRY (Fix for Bug 2)
                                companion_assignment = f"{element}{companion['orbital']}"
                                step2_data.append({
                                    'peak': companion['peak'],
                                    'assigned': companion_assignment,  # ← FIXED: Use companion_assignment
                                    'confidence': 80,
                                    'companion': f"Companion to {assignment}",
                                    'possible': self.get_original_possibilities(companion['peak'], companion_assignment)
                                })
                                self.assigned_peaks.add(companion['peak']['index'])

                        elif orbital == 'kll':
                            # Na Auger peak found - look for Na1s companion
                            na1s_companion = None
                            for comp_peak in remaining_peaks:
                                if abs(comp_peak['position'] - 1071) <= 4.0:  # Na1s around 1071 eV
                                    na1s_companion = comp_peak
                                    break

                            if na1s_companion:
                                companion_found = True
                                companion_text = f"{element}1s"
                                confidence = 90

                                # ADD Na1s PEAK AS SEPARATE ENTRY
                                na1s_assignment = f"{element}1s"
                                step2_data.append({
                                    'peak': na1s_companion,
                                    'assigned': na1s_assignment,  # ← FIXED: Use na1s_assignment
                                    'confidence': 85,
                                    'companion': f"Companion to {assignment}",
                                    'possible': self.get_original_possibilities(na1s_companion, na1s_assignment)
                                })
                                self.assigned_peaks.add(na1s_companion['index'])

                        # Calculate final confidence
                        final_confidence = self.auto_survey_id.calculate_confidence_score(
                            peak, "priority_1", companion_found, possibility['distance']
                        )

                        # Assign this peak
                        if companion_found or peak['prominence'] > 0.02:
                            self.assigned_peaks.add(peak['index'])

                            step2_data.append({
                                'peak': peak,
                                'assigned': assignment,  # ← NOW assignment is defined
                                'confidence': int(final_confidence),  # ← NOW final_confidence is defined
                                'companion': companion_text,
                                'possible': self.get_original_possibilities(peak, assignment)
                            })

                            print(f"✓ ASSIGNED {assignment} at {peak['position']:.2f} eV (Priority 1)")
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
                            'possible': self.get_original_possibilities(peak, assignment) # Use original possibilities
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
                'possible': self.get_original_possibilities(peak, assignment) if assignment else possible_text  # ← NEW
            })

        self.populate_step_page(4, step4_data)

    def step6_peak_assignment_dismissal(self):
        """Step 6: Enhanced reassignment to detected elements + dismissal checks"""
        self.status_text.SetLabel("Step 6: Peak reassignment and dismissal...")

        step6_data = []
        dismissed_assignments = []
        reassigned_peaks = []

        # Define hierarchical main orbital requirements (any of these will satisfy)
        secondary_orbital_hierarchy = {
            '3p': ['3d', '2p', '1s'],  # 3p needs 3d OR 2p OR 1s
            '3s': ['3d', '2p', '1s'],  # 3s needs 3d OR 2p OR 1s
            '4d': ['4f', '3d', '2p', '1s'],  # 4d needs 4f OR 3d OR 2p OR 1s
            '4p': ['4f', '4d', '3d', '2p', '1s'],  # 4p needs any main orbital
            '2s': ['2p', '1s'],  # 2s needs 2p OR 1s
            '5d': ['5f', '4f', '3d', '2p', '1s'],  # 5d needs any main orbital
            '2p1': ['2p3', '2p'],  # 2p1/2 needs 2p3/2 OR 2p
        }

        # Collect all currently assigned peaks with their assignments
        assigned_elements_orbitals = set()
        detected_elements = set()  # Track which elements have been detected

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

                    # Extract element name and add to detected elements
                    digit_start = -1
                    for i, char in enumerate(assigned):
                        if char.isdigit():
                            digit_start = i
                            break

                    if digit_start > 0:
                        element = assigned[:digit_start]
                        detected_elements.add(element)

                    print(f"  Found assignment: {assigned} (from {step_name}) -> element {element} detected")

        print(f"  Total assignments found: {assigned_elements_orbitals}")
        print(f"  Detected elements: {detected_elements}")

        # ===== NEW: REASSIGNMENT TO DETECTED ELEMENTS =====
        print("=== Step 6: Checking for reassignments to detected elements ===")

        for step_idx in [2, 3, 4]:
            if step_idx >= len(self.step_pages):
                continue

            page = self.step_pages[step_idx]
            peak_list = page.peak_list
            step_name = ["", "", "Usual's", "Most commons", "Others"][step_idx]

            for row in range(peak_list.GetItemCount()):
                assigned = peak_list.GetItem(row, 5).GetText()  # Column 5 = Assigned To
                position_text = peak_list.GetItem(row, 1).GetText()  # Column 1 = Position
                possibilities_text = peak_list.GetItem(row, 7).GetText()  # Column 7 = Possibilities

                if assigned and assigned != "" and assigned != "DISMISSED":

                    # Parse current assignment
                    digit_start = -1
                    for i, char in enumerate(assigned):
                        if char.isdigit():
                            digit_start = i
                            break

                    if digit_start <= 0:
                        continue

                    current_element = assigned[:digit_start]
                    current_orbital = assigned[digit_start:]

                    print(f"    Checking {assigned} (current: {current_element}{current_orbital})")
                    print(f"      Possibilities: {possibilities_text}")

                    # BUILD DETECTED ELEMENTS EXCLUDING THE CURRENT ASSIGNMENT
                    detected_from_others = set()
                    for other_assignment in assigned_elements_orbitals:
                        if other_assignment != assigned:  # ← EXCLUDE current assignment
                            # Parse other assignment
                            other_digit_start = -1
                            for i, char in enumerate(other_assignment):
                                if char.isdigit():
                                    other_digit_start = i
                                    break

                            if other_digit_start > 0:
                                other_element = other_assignment[:other_digit_start]
                                detected_from_others.add(other_element)

                    print(f"      Elements detected from OTHER peaks: {detected_from_others}")

                    # Parse possibilities to find detected elements
                    if possibilities_text:
                        possibilities = [p.strip() for p in possibilities_text.split(',')]

                        best_detected_match = None
                        best_confidence_boost = 0

                        for possibility in possibilities:
                            # IMPROVED PARSING: Handle both photoelectron (Ti2p) and Auger (Calmm) peaks
                            poss_element = None
                            poss_orbital = None

                            # First try digit-based parsing (Ti2p, Ca2p, etc.)
                            poss_digit_start = -1
                            for i, char in enumerate(possibility):
                                if char.isdigit():
                                    poss_digit_start = i
                                    break

                            if poss_digit_start > 0:
                                # Standard photoelectron peak (e.g., Ti2p, Ca2p)
                                poss_element = possibility[:poss_digit_start]
                                poss_orbital = possibility[poss_digit_start:]
                            else:
                                # Check for Auger peaks (element + auger_type)
                                auger_suffixes = ['kll', 'lmm', 'mnn', 'mvv', 'mnv', 'kln', 'klm']

                                for suffix in auger_suffixes:
                                    if possibility.lower().endswith(suffix):
                                        poss_element = possibility[:-len(suffix)]
                                        poss_orbital = suffix
                                        print(
                                            f"          → Parsed Auger: '{possibility}' = element '{poss_element}' + orbital '{poss_orbital}'")
                                        break

                            if poss_element and poss_orbital:
                                print(f"          → Parsed: element='{poss_element}', orbital='{poss_orbital}'")
                                print(
                                    f"          → Element '{poss_element}' in detected? {poss_element in detected_from_others}")

                                # Check if this element is detected from OTHER peaks
                                if poss_element in detected_from_others and poss_element != current_element:
                                    print(
                                        f"        ✓ Found better match: {possibility} (element {poss_element} detected from other peaks)")

                                    # Calculate priority boost based on orbital importance
                                    orbital_priority = {
                                        '1s': 50, '2p': 45, '3d': 40, '4f': 35,
                                        '2s': 30, '3p': 25, '4d': 20, '3s': 15, '4p': 10,
                                        # Add Auger priorities
                                        'kll': 35, 'lmm': 30, 'mnn': 25, 'mvv': 20, 'mnv': 15
                                    }

                                    orbital_clean = poss_orbital.replace('/2', '').replace('3/2', '3').replace('1/2',
                                                                                                               '1')
                                    confidence_boost = orbital_priority.get(orbital_clean, 5)

                                    if confidence_boost > best_confidence_boost:
                                        best_detected_match = possibility
                                        best_confidence_boost = confidence_boost
                                        print(
                                            f"          → New best: {best_detected_match} (boost: +{best_confidence_boost})")
                                else:
                                    print(f"          ✗ Element '{poss_element}' not detected from other peaks")
                            else:
                                print(f"          ⚠ Could not parse possibility: '{possibility}'")

                        # If we found a better match with a detected element
                        if best_detected_match:
                            old_assignment = assigned
                            new_assignment = best_detected_match

                            print(
                                f"      ✓ REASSIGNING: {old_assignment} → {new_assignment} (element detected from other peaks)")

                            # Update the assignment
                            peak_list.SetItem(row, 5, new_assignment)  # Update Assigned To

                            # Calculate new confidence (boost for detected element)
                            old_confidence_text = peak_list.GetItem(row, 4).GetText()
                            try:
                                old_confidence = int(old_confidence_text)
                                new_confidence = min(95, old_confidence + best_confidence_boost)
                                peak_list.SetItem(row, 4, str(new_confidence))
                            except ValueError:
                                new_confidence = 80
                                peak_list.SetItem(row, 4, "80")

                            # Update companion info - FIXED to handle Auger peaks
                            new_element = None

                            # Try digit-based parsing first (Ti2p, Ca2p, etc.)
                            digit_pos = -1
                            for i, char in enumerate(new_assignment):
                                if char.isdigit():
                                    digit_pos = i
                                    break

                            if digit_pos > 0:
                                # Standard photoelectron peak
                                new_element = new_assignment[:digit_pos]
                            else:
                                # Check for Auger peaks
                                auger_suffixes = ['kll', 'lmm', 'mnn', 'mvv', 'mnv', 'kln', 'klm']
                                for suffix in auger_suffixes:
                                    if new_assignment.lower().endswith(suffix):
                                        new_element = new_assignment[:-len(suffix)]
                                        break

                            if new_element:
                                peak_list.SetItem(row, 6,
                                                  f"Reassigned to detected {new_element} (+{best_confidence_boost})")
                            else:
                                peak_list.SetItem(row, 6, f"Reassigned (+{best_confidence_boost})")

                            # Color it blue for reassignment
                            peak_list.SetItemBackgroundColour(row, wx.Colour(200, 220, 255))

                            # Track reassignment
                            reassigned_peaks.append({
                                'old': old_assignment,
                                'new': new_assignment,
                                'position': position_text,
                                'confidence_boost': best_confidence_boost
                            })

                            # Update our tracking sets
                            assigned_elements_orbitals.discard(old_assignment)
                            assigned_elements_orbitals.add(new_assignment)
                        else:
                            print(f"      → No better matches with elements detected from other peaks")
                    else:
                        print(f"      → No possibilities available")

        # ===== NEW: REMOVE DUPLICATE ASSIGNMENTS =====
        print("=== Step 6: Checking for duplicate assignments ===")

        # Collect all current assignments with their peak info
        assignment_to_peaks = {}  # assignment -> [(peak_info, step_idx, row), ...]

        for step_idx in [2, 3, 4]:
            if step_idx >= len(self.step_pages):
                continue

            page = self.step_pages[step_idx]
            peak_list = page.peak_list
            step_name = ["", "", "Usual's", "Most commons", "Others"][step_idx]

            for row in range(peak_list.GetItemCount()):
                assigned = peak_list.GetItem(row, 5).GetText()  # Column 5 = Assigned To
                position_text = peak_list.GetItem(row, 1).GetText()  # Column 1 = Position
                prominence_text = peak_list.GetItem(row, 3).GetText()  # Column 3 = Prominence

                if assigned and assigned != "" and assigned != "DISMISSED":
                    try:
                        position = float(position_text)
                        prominence = float(prominence_text)

                        if assigned not in assignment_to_peaks:
                            assignment_to_peaks[assigned] = []

                        assignment_to_peaks[assigned].append({
                            'position': position,
                            'prominence': prominence,
                            'step_idx': step_idx,
                            'row': row,
                            'page': page,
                            'peak_list': peak_list
                        })
                    except ValueError:
                        pass

        # Find and resolve duplicates
        for assignment, peaks_info in assignment_to_peaks.items():
            if len(peaks_info) > 1:
                print(f"  Found duplicate assignment '{assignment}' in {len(peaks_info)} peaks:")

                # Sort by prominence (highest first)
                peaks_info.sort(key=lambda x: x['prominence'], reverse=True)

                # Keep the most prominent, dismiss the rest
                for i, peak_info in enumerate(peaks_info):
                    if i == 0:
                        print(
                            f"    ✓ KEEPING: {assignment} at {peak_info['position']:.2f} eV (prominence: {peak_info['prominence']:.4f})")
                    else:
                        print(
                            f"    ✗ DISMISSING: {assignment} at {peak_info['position']:.2f} eV (prominence: {peak_info['prominence']:.4f}) - duplicate")

                        # Dismiss the duplicate
                        peak_info['peak_list'].SetItem(peak_info['row'], 5, "DISMISSED")
                        peak_info['peak_list'].SetItem(peak_info['row'], 6, f"Duplicate {assignment} (less prominent)")
                        peak_info['peak_list'].SetItemBackgroundColour(peak_info['row'], wx.Colour(255, 200, 200))

                        dismissed_assignments.append(f"{assignment} (duplicate)")

        # ===== EXISTING: ENHANCEMENT AND DISMISSAL LOGIC =====
        print("=== Step 6: Starting enhancement and dismissal checks ===")

        for step_idx in [2, 3, 4]:
            if step_idx >= len(self.step_pages):
                continue

            page = self.step_pages[step_idx]
            peak_list = page.peak_list
            step_name = ["", "", "Usual's", "Most commons", "Others"][step_idx]

            for row in range(peak_list.GetItemCount()):
                assigned = peak_list.GetItem(row, 5).GetText()  # Column 5 = Assigned To
                position_text = peak_list.GetItem(row, 1).GetText()  # Column 1 = Position

                if assigned and assigned != "" and assigned != "DISMISSED":

                    # Parse assignment
                    digit_start = -1
                    for i, char in enumerate(assigned):
                        if char.isdigit():
                            digit_start = i
                            break

                    if digit_start <= 0:
                        continue

                    element = assigned[:digit_start]
                    orbital = assigned[digit_start:]
                    orbital_normalized = orbital.replace('/2', '').replace('3/2', '3').replace('1/2', '1')

                    # Skip if this was already processed in reassignment (has blue color)
                    try:
                        current_color = peak_list.GetItemBackgroundColour(row)
                        if current_color == wx.Colour(200, 220, 255):  # Blue = reassigned
                            continue
                    except:
                        pass

                    # ENHANCEMENT: If element is detected, boost confidence
                    if element in detected_elements:
                        current_confidence = peak_list.GetItem(row, 4).GetText()
                        try:
                            confidence_val = int(current_confidence)
                            boosted_confidence = min(95, confidence_val + 15)  # +15 boost, max 95
                            peak_list.SetItem(row, 4, str(boosted_confidence))
                            peak_list.SetItem(row, 6, f"Element {element} detected (+15 confidence)")
                            peak_list.SetItemBackgroundColour(row, wx.Colour(200, 255, 200))  # Green
                            print(
                                f"      ✓ BOOSTED confidence: {confidence_val} → {boosted_confidence} (element detected)")
                        except ValueError:
                            pass

                    # DISMISSAL: Check if secondary orbital needs main orbital
                    elif orbital_normalized in secondary_orbital_hierarchy:
                        required_orbitals = secondary_orbital_hierarchy[orbital_normalized]

                        # Check if ANY of the required main orbitals has been assigned
                        main_found = False
                        found_main = None

                        for required_orbital in required_orbitals:
                            main_assignment = f"{element}{required_orbital}"

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

                                    if existing_element == element and existing_orbital == required_orbital:
                                        main_found = True
                                        found_main = existing
                                        break

                            if main_found:
                                break

                        if not main_found:
                            print(f"      ✗ NO main orbital found for {element} - DISMISSING {assigned}")

                            # MODIFY the original list - mark as DISMISSED
                            peak_list.SetItem(row, 5, "DISMISSED")  # Update Assigned To column
                            peak_list.SetItem(row, 6, f"No main {element} peak found")  # Update Companion

                            # Color it red
                            peak_list.SetItemBackgroundColour(row, wx.Colour(255, 200, 200))

                            # Track dismissal
                            dismissed_assignments.append(assigned)

        # SORT STEP6 DATA BY PROMINENCE
        step6_data.sort(key=lambda x: x['peak']['prominence'], reverse=True)
        self.populate_step_page(5, step6_data)  # Tab 5: "6. Final Checked"

        # Print summary
        if reassigned_peaks:
            print(f"=== Step 6: Reassigned {len(reassigned_peaks)} peaks to detected elements ===")
            for reassignment in reassigned_peaks:
                print(
                    f"  Reassigned: {reassignment['old']} → {reassignment['new']} (+{reassignment['confidence_boost']} confidence)")

        if dismissed_assignments:
            print(f"=== Step 6: Dismissed {len(dismissed_assignments)} assignments ===")
            for assignment in dismissed_assignments:
                print(f"  Dismissed: {assignment}")

        if not reassigned_peaks and not dismissed_assignments:
            print("=== Step 6: No reassignments or dismissals ===")

    def step7_final_compilation(self):
        """Step 7: Final compilation of all assignments - ordered by prominence"""
        self.status_text.SetLabel("Step 7: Compiling final results...")

        final_data = []
        seen_positions = set()  # Track positions to avoid duplicates

        # Collect all assigned peaks from all steps
        for step_idx in range(7):  # Now we have 7 total steps (0-6)
            if step_idx in [0,
                            5]:  # FIXED: Skip only step 0 (Find Peaks - no assignments) and step 5 (Final Checked - duplicate processing)
                continue

            if step_idx >= len(self.step_pages):
                continue

            page = self.step_pages[step_idx]
            peak_list = page.peak_list
            step_name = ["Find Peaks", "Peak Width", "Usual's", "Most commons", "Others", "Final Checked", "Results"][
                step_idx]

            print(f"  Collecting from step {step_idx} ({step_name}): {peak_list.GetItemCount()} rows")

            for row in range(peak_list.GetItemCount()):
                # UPDATED COLUMN INDICES for new column layout:
                # 0: T, 1: BE (eV), 2: Width (eV), 3: Signal, 4: Score, 5: Peak, 6: Companion, 7: Possibility, 8: Distance, 9: Decision
                assigned = peak_list.GetItem(row, 5).GetText()  # Column 5 = Peak (assigned to)

                if assigned and assigned != "" and assigned != "DISMISSED":
                    position_text = peak_list.GetItem(row, 1).GetText()
                    position = float(position_text)

                    # Avoid duplicates - skip if we've already seen this position
                    position_key = round(position, 1)  # Round to avoid floating point issues
                    if position_key in seen_positions:
                        print(f"    Skipping duplicate at {position:.1f} eV: {assigned}")
                        continue
                    seen_positions.add(position_key)

                    width = peak_list.GetItem(row, 2).GetText()
                    prominence = peak_list.GetItem(row, 3).GetText()
                    confidence_text = peak_list.GetItem(row, 4).GetText()  # Column 4 = Score
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
                        original_peak['create_region'] = confidence >= 60  # Updated threshold

                        # GET ORIGINAL POSSIBILITIES FROM STEP 1 (NOT MODIFIED ONES)
                        original_possibilities = self.original_possibilities.get(original_peak['index'], assigned)

                        final_data.append({
                            'peak': original_peak,
                            'assigned': assigned,
                            'confidence': confidence,
                            'companion': companion,
                            'possible': original_possibilities  # Use original from step 1
                        })

                        print(f"    Added: {assigned} at {position:.2f} eV (confidence: {confidence}) from {step_name}")

        print(f"  Total peaks collected: {len(final_data)}")

        # SORT FINAL DATA BY PROMINENCE (highest first)
        final_data.sort(key=lambda x: x['peak']['prominence'], reverse=True)

        self.populate_step_page(6, final_data)  # Tab 6: "7. Results"

    def populate_step_page(self, step_index, data_list):
        """Populate a step page with peak data including checkboxes, distance, and decision"""
        if step_index >= len(self.step_pages):
            return

        page = self.step_pages[step_index]
        peak_list = page.peak_list

        # DETECT COMPANIONS AFTER ALL ASSIGNMENTS ARE MADE
        data_list = self.detect_companions_post_assignment(data_list)

        # Store original data for sorting
        page.original_data = data_list.copy()

        # Clear existing items
        peak_list.DeleteAllItems()

        # Add data - ALL DATA ALREADY SORTED BY PROMINENCE
        for i, data in enumerate(data_list):
            peak = data['peak']
            assigned = data['assigned']

            # Calculate distance and decision values
            distance_str, decision_str = self.calculate_peak_metrics(peak, assigned)

            # Check if peak should be ticked - UPDATED THRESHOLD TO 60
            should_tick = peak.get('create_region', True) and data['confidence'] >= 60
            tick_text = "☑" if should_tick else "☐"

            index = peak_list.InsertItem(i, tick_text)
            peak_list.SetItem(index, 1, f"{peak['position']:.2f}")
            peak_list.SetItem(index, 2, f"{peak['width']:.2f}")
            peak_list.SetItem(index, 3, f"{peak['prominence']:.4f}")
            peak_list.SetItem(index, 4, str(data['confidence']))
            peak_list.SetItem(index, 5, data['assigned'])
            peak_list.SetItem(index, 6, self.format_companion_text(data['companion']))  # ← Now with better logic
            peak_list.SetItem(index, 7, data['possible'])
            peak_list.SetItem(index, 8, distance_str)
            peak_list.SetItem(index, 9, decision_str)

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

    def on_peak_selected_OLD(self, event, page):
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
        """Handle force assignment button click"""
        selected_index = page.peak_list.GetFirstSelected()
        selected_assignment = page.core_level_choice.GetStringSelection()

        if selected_index == -1:
            wx.MessageBox("Please select a peak first", "No Peak Selected", wx.OK | wx.ICON_WARNING)
            return

        if not selected_assignment:
            wx.MessageBox("Please select an assignment from the dropdown", "No Assignment Selected",
                          wx.OK | wx.ICON_WARNING)
            return

        # Update the assignment in the list
        page.peak_list.SetItem(selected_index, 5, selected_assignment)  # Peak column

        # Calculate new metrics for the forced assignment
        if hasattr(page, 'original_data') and selected_index < len(page.original_data):
            peak_data = page.original_data[selected_index]
            peak = peak_data['peak']

            # Update the underlying data
            page.original_data[selected_index]['assigned'] = selected_assignment

            # Recalculate distance and decision
            distance_str, decision_str = self.calculate_peak_metrics(peak, selected_assignment)
            page.peak_list.SetItem(selected_index, 8, distance_str)
            page.peak_list.SetItem(selected_index, 9, decision_str)

            # Set high confidence for manual assignments
            page.peak_list.SetItem(selected_index, 4, "95")  # Score
            page.original_data[selected_index]['confidence'] = 95

            # Update companion info
            page.peak_list.SetItem(selected_index, 6, "Manual assignment")

            # Color it differently for manual assignments
            page.peak_list.SetItemBackgroundColour(selected_index, wx.Colour(200, 200, 255))  # Blue

            print(f"Force assigned peak {selected_index} to {selected_assignment}")

        event.Skip()

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

    def bind_peak_list_events(self):
        """Bind click events to all peak lists"""
        for page in self.step_pages:
            if hasattr(page, 'peak_list'):
                page.peak_list.Bind(wx.EVT_LIST_ITEM_SELECTED, self.on_peak_selected)

    def on_peak_selected(self, event, page=None):  # ← Add default page=None
        """Handle peak selection and populate force assignment combo"""

        # Handle case where page isn't passed (shouldn't happen but safety check)
        if page is None:
            print("WARNING: on_peak_selected called without page argument")
            event.Skip()
            return

        selected_index = event.GetIndex()

        if selected_index != -1 and hasattr(page, 'original_data') and selected_index < len(page.original_data):
            # Get the selected peak data
            peak_data = page.original_data[selected_index]
            peak = peak_data['peak']

            # Get all possible assignments for this peak
            possible = self.auto_survey_id.get_possible_assignments(peak['position'], self.tolerance)

            # Populate the combo box with possibilities
            choices = [p['assignment'] for p in possible[:10]]  # Top 10 possibilities

            # Clear and populate combo box
            page.core_level_choice.Clear()
            for choice in choices:
                page.core_level_choice.Append(choice)

            # Highlight the peak region on the plot
            position_text = page.peak_list.GetItem(selected_index, 1).GetText()
            width_text = page.peak_list.GetItem(selected_index, 2).GetText()

            try:
                position = float(position_text)
                width = float(width_text)
                self.highlight_peak_region(position, width)
            except ValueError:
                pass

            print(f"Selected peak at {peak['position']:.2f} eV - {len(choices)} possibilities loaded")

        event.Skip()

    def highlight_peak_region(self, center, width):
        """Highlight peak region on the plot with transparent pink overlay"""
        try:
            # Clear any existing highlights
            self.clear_peak_highlight()

            # Calculate region bounds (center ± width/2)
            half_width = width / 2.0
            x_min = center - 2*width
            x_max = center + width

            # Get the plot axes
            ax = self.parent.figure.gca()

            # Get y-axis limits for the highlight
            y_min, y_max = ax.get_ylim()

            # Create transparent pink highlight
            self.peak_highlight = ax.axvspan(x_min, x_max,
                                             alpha=0.3,
                                             color='hotpink',
                                             zorder=1)

            # # Add vertical lines at peak center
            # self.peak_center_line = ax.axvline(center,
            #                                    color='red',
            #                                    linestyle='--',
            #                                    alpha=0.7,
            #                                    linewidth=2,
            #                                    zorder=2)

            # Refresh the plot
            self.parent.canvas.draw()

            print(f"Highlighted peak: {center:.2f} eV (width: {width:.2f} eV)")

        except Exception as e:
            print(f"Error highlighting peak: {e}")

    def clear_peak_highlight(self):
        """Clear any existing peak highlights"""
        try:
            if hasattr(self, 'peak_highlight') and self.peak_highlight:
                self.peak_highlight.remove()
                self.peak_highlight = None

            if hasattr(self, 'peak_center_line') and self.peak_center_line:
                self.peak_center_line.remove()
                self.peak_center_line = None

            # Refresh the plot
            self.parent.canvas.draw()

        except Exception as e:
            print(f"Error clearing peak highlight: {e}")

    def on_column_click(self, event, page):
        """Handle column header clicks for sorting"""
        col = event.GetColumn()

        # Toggle sort order if same column, otherwise sort ascending
        if page.sort_column == col:
            page.sort_ascending = not page.sort_ascending
        else:
            page.sort_column = col
            page.sort_ascending = True

        # Sort the data
        self.sort_page_data(page, col, page.sort_ascending)

        print(f"Sorted by column {col} ({'ascending' if page.sort_ascending else 'descending'})")

    def sort_page_data(self, page, col, ascending):
        """Sort page data by specified column"""
        if not hasattr(page, 'original_data') or not page.original_data:
            return

        data_list = page.original_data.copy()

        # Define sort keys for each column
        sort_keys = {
            0: lambda x: x['peak'].get('create_region', False),  # T (ticked status)
            1: lambda x: x['peak']['position'],  # BE (eV)
            2: lambda x: x['peak']['width'],  # Width (eV)
            3: lambda x: x['peak']['prominence'],  # Signal
            4: lambda x: x['confidence'],  # Score
            5: lambda x: x['assigned'],  # Peak
            6: lambda x: x['companion'],  # Companion
            7: lambda x: x['possible'],  # Possibility
            8: lambda x: self.get_first_distance(x['peak']),  # Distance (first value)
            9: lambda x: self.get_first_decision(x['peak'], x['assigned'])  # Decision (first value)
        }

        if col in sort_keys:
            try:
                data_list.sort(key=sort_keys[col], reverse=not ascending)

                # Repopulate the list with sorted data
                self.populate_sorted_page(page, data_list)

            except Exception as e:
                print(f"Error sorting: {e}")

    def get_first_distance(self, peak):
        """Get first distance value for sorting"""
        try:
            possible = self.auto_survey_id.get_possible_assignments(peak['position'], self.tolerance)
            return possible[0]['distance'] if possible else 999
        except:
            return 999

    def get_first_decision(self, peak, assigned):
        """Get first decision value for sorting"""
        try:
            distance_str, decision_str = self.calculate_peak_metrics(peak, assigned)
            first_decision = decision_str.split(',')[0].strip()
            return float(first_decision) if first_decision != "N/A" else 999
        except:
            return 999

    def populate_sorted_page(self, page, data_list):
        """Repopulate page with sorted data"""
        peak_list = page.peak_list
        peak_list.DeleteAllItems()

        # Update original_data to match the new sorted order
        page.original_data = data_list.copy()

        # Use the same logic as populate_step_page but with sorted data
        for i, data in enumerate(data_list):
            peak = data['peak']
            assigned = data['assigned']

            # Calculate distance and decision values
            distance_str, decision_str = self.calculate_peak_metrics(peak, assigned)

            # Check if peak should be ticked - UPDATED THRESHOLD
            should_tick = peak.get('create_region', True) and data['confidence'] >= 60  # ← Changed from 30 to 60
            tick_text = "☑" if should_tick else "☐"

            index = peak_list.InsertItem(i, tick_text)
            peak_list.SetItem(index, 1, f"{peak['position']:.2f}")
            peak_list.SetItem(index, 2, f"{peak['width']:.2f}")
            peak_list.SetItem(index, 3, f"{peak['prominence']:.4f}")
            peak_list.SetItem(index, 4, str(data['confidence']))
            peak_list.SetItem(index, 5, data['assigned'])
            peak_list.SetItem(index, 6, self.format_companion_text(data['companion']))  # ← NEW formatting
            peak_list.SetItem(index, 7, data['possible'])
            peak_list.SetItem(index, 8, distance_str)
            peak_list.SetItem(index, 9, decision_str)

            # Color code by confidence
            if data['confidence'] >= 80:
                peak_list.SetItemBackgroundColour(index, wx.Colour(200, 255, 200))  # Green
            elif data['confidence'] >= 60:
                peak_list.SetItemBackgroundColour(index, wx.Colour(255, 255, 200))  # Yellow
            elif data['confidence'] >= 40:
                peak_list.SetItemBackgroundColour(index, wx.Colour(255, 200, 200))  # Light red
            else:
                peak_list.SetItemBackgroundColour(index, wx.Colour(255, 150, 150))  # Red

    def on_list_click(self, event, page):
        """Handle mouse clicks on list items for checkbox toggling"""
        pos = event.GetPosition()
        item, flags = page.peak_list.HitTest(pos)

        if item != wx.NOT_FOUND:
            # Check if click was in the T column (first column)
            rect = page.peak_list.GetItemRect(item)
            if pos.x <= 30:  # First column width + margin
                # Toggle checkbox
                current_text = page.peak_list.GetItem(item, 0).GetText()
                new_text = "☐" if current_text == "☑" else "☑"
                page.peak_list.SetItem(item, 0, new_text)

                # Update the underlying data
                if hasattr(page, 'original_data') and item < len(page.original_data):
                    page.original_data[item]['peak']['create_region'] = (new_text == "☑")

                print(f"Toggled peak {item}: {current_text} → {new_text}")

        event.Skip()  # Allow other events to process

    def format_companion_text(self, companion_text):
        """Simple pass-through for companion text - no reformatting needed"""
        if not companion_text:
            return ""

        # Just return as-is since detect_companions_post_assignment handles all formatting
        return companion_text

    def detect_companions_post_assignment(self, step_data):
        """Count total peaks per element and show +N for additional peaks"""

        # Count total assigned peaks per element
        element_counts = {}

        for data in step_data:
            assigned = data['assigned']
            if assigned and assigned != "DISMISSED":
                # Parse element from assignment
                element = None
                digit_start = -1
                for i, char in enumerate(assigned):
                    if char.isdigit():
                        digit_start = i
                        break

                if digit_start > 0:
                    # Standard photoelectron peak (Ti2p, Mo3d, etc.)
                    element = assigned[:digit_start]
                else:
                    # Check for Auger peaks (Calmm, Nakll, etc.)
                    auger_suffixes = ['kll', 'lmm', 'mnn', 'mvv', 'mnv']
                    for suffix in auger_suffixes:
                        if assigned.lower().endswith(suffix):
                            element = assigned[:-len(suffix)]
                            break

                if element:
                    element_counts[element] = element_counts.get(element, 0) + 1

        # Update companion text for each peak
        for data in step_data:
            assigned = data['assigned']
            if assigned and assigned != "DISMISSED":
                # Parse element from assignment
                element = None
                digit_start = -1
                for i, char in enumerate(assigned):
                    if char.isdigit():
                        digit_start = i
                        break

                if digit_start > 0:
                    element = assigned[:digit_start]
                else:
                    auger_suffixes = ['kll', 'lmm', 'mnn', 'mvv', 'mnv']
                    for suffix in auger_suffixes:
                        if assigned.lower().endswith(suffix):
                            element = assigned[:-len(suffix)]
                            break

                if element:
                    total_peaks = element_counts.get(element, 1)
                    additional_peaks = total_peaks - 1  # Subtract self

                    # Set companion text based on additional peaks
                    if additional_peaks > 0:
                        data['companion'] = f"+{additional_peaks}"
                    else:
                        # Check for special cases (reassignments, etc.)
                        original_companion = data.get('companion', '')
                        if 'Reassigned' in original_companion:
                            data['companion'] = original_companion.replace("Reassigned to detected ", "→ ")
                        elif 'CERTAINTY' in original_companion:
                            data['companion'] = original_companion
                        else:
                            data['companion'] = ""

        return step_data

    def OnClose(self, event):
        """Handle dialog close event"""
        # Clear any peak highlights before closing
        self.clear_peak_highlight()

        self.Destroy()