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
        self.min_intensity_threshold = 0.01

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
            y_values = np.array(self.parent.Data['Core levels'][sheet_name]['Raw Data'])

            # Save state before making changes
            save_state(self.parent)

            # Clear existing peaks in grid
            self.clear_peak_grid()

            # Find peaks in survey
            peaks_found = self.find_survey_peaks(x_values, y_values)

            # Identify elements based on peaks
            identified_elements = self.identify_elements(peaks_found, x_values, y_values)

            # Create peaks and measure areas
            self.create_peaks_and_measure(identified_elements, x_values, y_values, sheet_name)

            # Update display
            self.parent.clear_and_replot()

            wx.MessageBox(f"Auto ID completed. Found {len(identified_elements)} elements.", "Success")

        except Exception as e:
            wx.MessageBox(f"Auto ID failed: {str(e)}", "Error")

    def find_survey_peaks(self, x_data, y_data):
        """Find peaks in survey spectrum with improved O1s detection"""
        # Normalize y_data
        y_norm = y_data / np.max(y_data)

        # Use lower prominence threshold for better O1s detection
        prominence = max(0.01, self.min_intensity_threshold * 0.5)  # Lower threshold

        # Find peaks with adjusted parameters for better detection
        peaks, properties = find_peaks(y_norm,
                                       prominence=prominence,
                                       width=1,  # Reduced from 2
                                       distance=3)  # Reduced from 5

        # Get peak positions and intensities
        peak_positions = x_data[peaks]
        peak_intensities = y_norm[peaks]
        peak_prominences = properties['prominences']

        # Special handling for O1s region (around 532 eV)
        o1s_region_mask = (peak_positions >= 525) & (peak_positions <= 545)
        o1s_peaks = peaks[o1s_region_mask]

        if len(o1s_peaks) > 0:
            # Force include the strongest peak in O1s region
            o1s_peak_idx = np.argmax(y_norm[o1s_peaks])
            o1s_peak = o1s_peaks[o1s_peak_idx]
            # Boost O1s prominence to ensure it's included
            peak_prominences[np.where(peaks == o1s_peak)[0][0]] = max(
                peak_prominences[np.where(peaks == o1s_peak)[0][0]], prominence * 2
            )

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

        return peaks_found

    def identify_elements(self, peaks_found, x_data, y_data):
        """Identify elements from found peaks using priority strategy"""
        identified = {}
        used_peaks = set()

        # Build element database from library_data
        element_db = self.build_element_database()

        # Process peaks by prominence
        for peak in peaks_found:
            if peak['index'] in used_peaks:
                continue

            # Find best match for this peak
            best_match = self.find_best_element_match(peak['position'], element_db)

            if best_match:
                element = best_match['element']
                orbital = best_match['orbital']

                # Check if this is the main orbital for the element
                if self.is_main_orbital(element, orbital):
                    # Verify orbital hierarchy if needed
                    if orbital != '1s' and not self.verify_orbital_hierarchy(element, orbital, peaks_found, element_db):
                        continue

                    # Add to identified elements
                    key = f"{element}{orbital}"
                    if key not in identified:
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

        # Filter by priority and minimum criteria
        filtered = self.filter_by_priority(identified)

        return filtered

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
        if element in self.priority_1_elements and self.priority_1_elements[element] == orbital:
            return 1
        elif element in self.priority_2_elements and self.priority_2_elements[element] == orbital:
            return 2
        elif element in self.priority_3_elements and self.priority_3_elements[element] == orbital:
            return 3
        else:
            return 4

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
        """Create peaks in grid and measure areas"""
        # Initialize peak fitting data structure
        if 'Fitting' not in self.parent.Data['Core levels'][sheet_name]:
            self.parent.Data['Core levels'][sheet_name]['Fitting'] = {}
        if 'Peaks' not in self.parent.Data['Core levels'][sheet_name]['Fitting']:
            self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = {}

        peaks_data = self.parent.Data['Core levels'][sheet_name]['Fitting']['Peaks']

        # Sort by binding energy (high to low)
        sorted_elements = sorted(identified_elements.items(),
                                 key=lambda x: x[1]['peak_position'],
                                 reverse=True)

        row = 0
        for peak_name, element_data in sorted_elements:
            # Calculate area and height for this peak
            area, bg_low, bg_high, peak_height = self.calculate_peak_area_and_height(
                element_data['peak_position'],
                x_data,
                y_data
            )

            # Format peak name like AreaFit_Screen (with space and dot)
            formatted_peak_name = f"{peak_name} ."

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

            # Add to Data structure
            peaks_data[formatted_peak_name] = {
                'Position': element_data['peak_position'],
                'Height': peak_height,
                'FWHM': 2.0,
                'L/G': 30,
                'Area': area,
                'Sigma': 0,
                'Gamma': 0,
                'Skew': 0,
                'Fitting Model': 'Unfitted',
                'Bkg Type': 'Multi-Regions Smart',
                'Bkg Low': bg_low,
                'Bkg High': bg_high,
                'Bkg Offset Low': 0,
                'Bkg Offset High': 0,
                'Constraints': {
                    'position': '',
                    'height': '',
                    'fwhm': '',
                    'lg_ratio': '',
                    'area': '',
                    'sigma': '',
                    'gamma': ''
                }
            }

            row += 2

        # Update peak count
        self.parent.peak_count = row // 2

        # Update background if needed
        if 'Background' not in self.parent.Data['Core levels'][sheet_name]:
            self.parent.Data['Core levels'][sheet_name]['Background'] = {
                'Bkg Y': y_data.tolist()
            }

    def calculate_peak_area_and_height(self, peak_position, x_data, y_data):
        """Calculate area under peak with background subtraction and height - same method as AreaFit_Screen"""
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
            return 0.0, peak_position - window_width, peak_position + window_width, 0.0

        # Calculate linear background
        background = np.linspace(y_region[0], y_region[-1], len(y_region))

        # Calculate background-subtracted data
        y_minus_bg = y_region - background

        # Sort data for proper integration
        sorted_indices = np.argsort(x_region)
        x_sorted = x_region[sorted_indices]
        y_minus_bg_sorted = y_minus_bg[sorted_indices]

        # Calculate area - same as AreaFit_Screen
        area = np.trapz(y_minus_bg_sorted, x_sorted)

        # Calculate height - same as AreaFit_Screen
        peak_index = np.argmax(y_minus_bg)
        peak_height = y_minus_bg[peak_index]

        return abs(area), peak_position - window_width, peak_position + window_width, max(0.0, peak_height)

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