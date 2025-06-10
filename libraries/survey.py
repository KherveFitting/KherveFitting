import wx
import numpy as np
from libraries.Sheet_Operations import on_sheet_selected
import os
import sys

# Import the ElementTile class and try to import KherveDB methods
try:
    from libraries.kherveDB_wxpython import ElementTile, PeriodicTableXPS
    KHERVE_AVAILABLE = True
except ImportError:
    KHERVE_AVAILABLE = False
    print("Warning: Could not import from kherveDB_wxpython")

def get_libraryid_path():
    """Get the path to the LibraryID.py file"""
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        base_path = os.path.dirname(sys.executable)
    else:
        # Running as script
        current_dir = os.path.dirname(os.path.abspath(__file__))
        base_path = os.path.dirname(current_dir)  # Go up one level to the root

    # Try both locations - root directory and libraries directory
    possible_paths = [
        os.path.join(base_path, "LibraryID.py"),
        os.path.join(base_path, "libraries", "LibraryID.py")
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return path

    return None


class PeriodicTableHelper:
    """Helper class to use KherveDB methods without importing the full class"""

    def __init__(self):
        pass

    def get_element_positions(self):
        """Define positions for elements in the periodic table grid"""
        positions = {}
        # Period 1
        positions['H'] = (0, 0)
        positions['He'] = (0, 17)
        # Period 2
        positions['Li'] = (1, 0)
        positions['Be'] = (1, 1)
        positions['B'] = (1, 12)
        positions['C'] = (1, 13)
        positions['N'] = (1, 14)
        positions['O'] = (1, 15)
        positions['F'] = (1, 16)
        positions['Ne'] = (1, 17)
        # Period 3
        positions['Na'] = (2, 0)
        positions['Mg'] = (2, 1)
        positions['Al'] = (2, 12)
        positions['Si'] = (2, 13)
        positions['P'] = (2, 14)
        positions['S'] = (2, 15)
        positions['Cl'] = (2, 16)
        positions['Ar'] = (2, 17)
        # Period 4
        positions['K'] = (3, 0)
        positions['Ca'] = (3, 1)
        for i, symbol in enumerate(['Sc', 'Ti', 'V', 'Cr', 'Mn', 'Fe', 'Co', 'Ni', 'Cu', 'Zn']):
            positions[symbol] = (3, i + 2)
        positions['Ga'] = (3, 12)
        positions['Ge'] = (3, 13)
        positions['As'] = (3, 14)
        positions['Se'] = (3, 15)
        positions['Br'] = (3, 16)
        positions['Kr'] = (3, 17)
        # Period 5
        positions['Rb'] = (4, 0)
        positions['Sr'] = (4, 1)
        for i, symbol in enumerate(['Y', 'Zr', 'Nb', 'Mo', 'Tc', 'Ru', 'Rh', 'Pd', 'Ag', 'Cd']):
            positions[symbol] = (4, i + 2)
        positions['In'] = (4, 12)
        positions['Sn'] = (4, 13)
        positions['Sb'] = (4, 14)
        positions['Te'] = (4, 15)
        positions['I'] = (4, 16)
        positions['Xe'] = (4, 17)
        # Period 6
        positions['Cs'] = (5, 0)
        positions['Ba'] = (5, 1)
        positions['La'] = (5, 2)
        for i, symbol in enumerate(['Hf', 'Ta', 'W', 'Re', 'Os', 'Ir', 'Pt', 'Au', 'Hg']):
            positions[symbol] = (5, i + 3)
        positions['Tl'] = (5, 12)
        positions['Pb'] = (5, 13)
        positions['Bi'] = (5, 14)
        positions['Po'] = (5, 15)
        positions['At'] = (5, 16)
        positions['Rn'] = (5, 17)
        # Period 7
        positions['Fr'] = (6, 0)
        positions['Ra'] = (6, 1)
        positions['Ac'] = (6, 2)
        for i, symbol in enumerate(['Rf', 'Db', 'Sg', 'Bh', 'Hs', 'Mt', 'Ds', 'Rg', 'Cn']):
            positions[symbol] = (6, i + 3)
        positions['Nh'] = (6, 12)
        positions['Fl'] = (6, 13)
        positions['Mc'] = (6, 14)
        positions['Lv'] = (6, 15)
        positions['Ts'] = (6, 16)
        positions['Og'] = (6, 17)
        # Lanthanides
        lanthanides = ['La', 'Ce', 'Pr', 'Nd', 'Pm', 'Sm', 'Eu', 'Gd', 'Tb', 'Dy', 'Ho', 'Er', 'Tm', 'Yb', 'Lu']
        for i, symbol in enumerate(lanthanides):
            positions[symbol] = (8, i + 2)
        # Actinides
        actinides = ['Ac', 'Th', 'Pa', 'U', 'Np', 'Pu', 'Am', 'Cm', 'Bk', 'Cf', 'Es', 'Fm', 'Md', 'No', 'Lr']
        for i, symbol in enumerate(actinides):
            positions[symbol] = (9, i + 2)
        return positions

    def get_element_categories(self):
        """Return element categories for coloring"""
        element_categories = {
            'H': 'nonmetal', 'He': 'noble_gas',
            'Li': 'alkali_metal', 'Be': 'alkaline_earth',
            'B': 'metalloid', 'C': 'nonmetal', 'N': 'nonmetal', 'O': 'nonmetal',
            'F': 'halogen', 'Ne': 'noble_gas',
            'Na': 'alkali_metal', 'Mg': 'alkaline_earth',
            'Al': 'post_transition', 'Si': 'metalloid', 'P': 'nonmetal',
            'S': 'nonmetal', 'Cl': 'halogen', 'Ar': 'noble_gas',
            'K': 'alkali_metal', 'Ca': 'alkaline_earth',
            'Sc': 'transition_metal', 'Ti': 'transition_metal', 'V': 'transition_metal',
            'Cr': 'transition_metal', 'Mn': 'transition_metal', 'Fe': 'transition_metal',
            'Co': 'transition_metal', 'Ni': 'transition_metal', 'Cu': 'transition_metal',
            'Zn': 'transition_metal', 'Ga': 'post_transition', 'Ge': 'metalloid',
            'As': 'metalloid', 'Se': 'nonmetal', 'Br': 'halogen', 'Kr': 'noble_gas',
            'Rb': 'alkali_metal', 'Sr': 'alkaline_earth',
            'Y': 'transition_metal', 'Zr': 'transition_metal', 'Nb': 'transition_metal',
            'Mo': 'transition_metal', 'Tc': 'transition_metal', 'Ru': 'transition_metal',
            'Rh': 'transition_metal', 'Pd': 'transition_metal', 'Ag': 'transition_metal',
            'Cd': 'transition_metal', 'In': 'post_transition', 'Sn': 'post_transition',
            'Sb': 'metalloid', 'Te': 'metalloid', 'I': 'halogen', 'Xe': 'noble_gas',
            'Cs': 'alkali_metal', 'Ba': 'alkaline_earth',
            'Hf': 'transition_metal', 'Ta': 'transition_metal', 'W': 'transition_metal',
            'Re': 'transition_metal', 'Os': 'transition_metal', 'Ir': 'transition_metal',
            'Pt': 'transition_metal', 'Au': 'transition_metal', 'Hg': 'transition_metal',
            'Tl': 'post_transition', 'Pb': 'post_transition', 'Bi': 'post_transition',
            'Po': 'metalloid', 'At': 'halogen', 'Rn': 'noble_gas',
            'Fr': 'alkali_metal', 'Ra': 'alkaline_earth',
            'Rf': 'transition_metal', 'Db': 'transition_metal', 'Sg': 'transition_metal',
            'Bh': 'transition_metal', 'Hs': 'transition_metal', 'Mt': 'transition_metal',
            'Ds': 'transition_metal', 'Rg': 'transition_metal', 'Cn': 'transition_metal',
            'Nh': 'post_transition', 'Fl': 'post_transition', 'Mc': 'post_transition',
            'Lv': 'post_transition', 'Ts': 'halogen', 'Og': 'noble_gas'
        }

        # Add lanthanides and actinides
        lanthanides = ['La', 'Ce', 'Pr', 'Nd', 'Pm', 'Sm', 'Eu', 'Gd', 'Tb', 'Dy', 'Ho', 'Er', 'Tm', 'Yb', 'Lu']
        actinides = ['Ac', 'Th', 'Pa', 'U', 'Np', 'Pu', 'Am', 'Cm', 'Bk', 'Cf', 'Es', 'Fm', 'Md', 'No', 'Lr']

        for element in lanthanides:
            element_categories[element] = 'lanthanide'
        for element in actinides:
            element_categories[element] = 'actinide'

        return element_categories


class PeriodicTableWindow(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, title="Survey Identification / Labelling",
                         style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX),
                         size=(940, 450))

        # Set window properties
        self.SetMinSize((940, 450))
        self.SetMaxSize((940, 450))
        self.Centre()

        self.parent_window = parent
        self.library_data = self.parent_window.library_data
        self.button_states = {}
        self.element_lines = {}
        self.element_buttons = {}
        self.original_colors = {}
        self.element_lines = {}  # For plot lines
        self.core_level_data = {}  # For list box data

        self.intensity_scale = 0.6

        # Create a KherveDB instance to access its methods
        if KHERVE_AVAILABLE:
            self.kherve_instance = PeriodicTableXPS()
            # We need to initialize some attributes that KherveDB expects
            self.kherve_instance.elements = set(self.get_available_elements())

        self.InitUI()
        self.Bind(wx.EVT_CLOSE, self.OnClose)

    def get_available_elements(self):
        """Get list of elements available in your library data"""
        elements = set()
        for (elem, orbital), data in self.library_data.items():
            elements.add(elem)
        return elements

    def InitUI(self):
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Info text
        self.info_text1 = wx.StaticText(panel, style=wx.ALIGN_CENTER | wx.ST_NO_AUTORESIZE)
        self.info_text1.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        main_sizer.Add(self.info_text1, 0, wx.EXPAND | wx.ALL, 0)

        self.info_text2 = wx.StaticText(panel, style=wx.ALIGN_CENTER | wx.ST_NO_AUTORESIZE)
        self.info_text2.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        main_sizer.Add(self.info_text2, 0, wx.EXPAND | wx.ALL, 0)

        hsizer = wx.BoxSizer(wx.HORIZONTAL)

        # Left side: Use KherveDB periodic table
        if KHERVE_AVAILABLE:
            self.create_kherve_periodic_table(panel, hsizer)
        else:
            self.create_fallback_periodic_table(panel, hsizer)

        # Right side: Core Level List and Buttons
        right_sizer = wx.BoxSizer(wx.VERTICAL)

        self.core_level_list = wx.ListBox(panel, style=wx.LB_MULTIPLE, size=(170, -1))
        right_sizer.Add(self.core_level_list, 1, wx.EXPAND | wx.ALL, 5)

        # Buttons
        button_sizer = wx.GridBagSizer(1, 1)

        self.add_labels_btn = wx.Button(panel, label="Add Labels")
        self.add_peak_btn = wx.Button(panel, label="Add to Grid")
        self.remove_selected_btn = wx.Button(panel, label="Clear Selected")
        self.remove_all_btn = wx.Button(panel, label="Clear All List")
        self.remove_last_label_btn = wx.Button(panel, label="Clear Last Label")
        self.remove_all_labels_btn = wx.Button(panel, label="Clear All Labels")

        # Bind events
        self.add_labels_btn.Bind(wx.EVT_BUTTON, self.OnAddLabels)
        self.add_peak_btn.Bind(wx.EVT_BUTTON, self.OnAddPeak)
        self.remove_selected_btn.Bind(wx.EVT_BUTTON, self.OnRemoveSelected)
        self.remove_all_btn.Bind(wx.EVT_BUTTON, self.OnRemoveAll)
        self.remove_last_label_btn.Bind(wx.EVT_BUTTON, self.OnRemoveLastLabel)
        self.remove_all_labels_btn.Bind(wx.EVT_BUTTON, self.OnRemoveAllLabels)

        button_sizer.Add(self.add_labels_btn, pos=(0, 0), flag=wx.EXPAND)
        button_sizer.Add(self.add_peak_btn, pos=(0, 1), flag=wx.EXPAND)
        button_sizer.Add(self.remove_selected_btn, pos=(1, 0), flag=wx.EXPAND)
        button_sizer.Add(self.remove_all_btn, pos=(1, 1), flag=wx.EXPAND)
        button_sizer.Add(self.remove_last_label_btn, pos=(2, 0), flag=wx.EXPAND)
        button_sizer.Add(self.remove_all_labels_btn, pos=(2, 1), flag=wx.EXPAND)

        right_sizer.Add(button_sizer, 0, wx.ALL | wx.EXPAND, 5)

        # ADD NEW INTENSITY CONTROL HERE
        intensity_sizer = wx.BoxSizer(wx.HORIZONTAL)

        intensity_label = wx.StaticText(panel, label="Line Intensity:")
        intensity_sizer.Add(intensity_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 2)

        # Decrease button
        self.intensity_down_btn = wx.Button(panel, label="-", size=(25, 25))
        self.intensity_down_btn.Bind(wx.EVT_BUTTON, self.OnIntensityDecrease)
        intensity_sizer.Add(self.intensity_down_btn, 0, wx.ALL, 2)

        # Display current value
        self.intensity_display = wx.StaticText(panel, label="0.6", size=(30, -1), style=wx.ALIGN_CENTER)
        self.intensity_display.SetBackgroundColour(wx.WHITE)
        intensity_sizer.Add(self.intensity_display, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 2)

        # Increase button
        self.intensity_up_btn = wx.Button(panel, label="+", size=(25, 25))
        self.intensity_up_btn.Bind(wx.EVT_BUTTON, self.OnIntensityIncrease)
        intensity_sizer.Add(self.intensity_up_btn, 0, wx.ALL, 2)

        right_sizer.Add(intensity_sizer, 0, wx.ALL | wx.EXPAND, 5)

        hsizer.Add(right_sizer, 0, wx.EXPAND)

        main_sizer.Add(hsizer, 1, wx.EXPAND)
        panel.SetSizer(main_sizer)

    def create_kherve_periodic_table(self, parent, sizer):
        """Create periodic table using actual KherveDB ElementTiles"""
        # Create frame for periodic table
        pt_panel = wx.Panel(parent, style=wx.BORDER_SUNKEN)
        pt_panel.SetBackgroundColour(wx.Colour(230, 230, 230))

        # Use grid sizer for periodic table layout
        pt_sizer = wx.GridBagSizer(1, 1)

        # Get data using KherveDB methods
        element_positions = self.kherve_instance.get_element_positions()
        element_categories = self.kherve_instance.get_element_categories()

        # Define color schemes (from KherveDB)
        colors = {
            'alkali_metal': "#FF6666",
            'alkaline_earth': "#FFDEAD",
            'transition_metal': "#FFC0CB",
            'post_transition': "#CCCCCC",
            'metalloid': "#97FFFF",
            'nonmetal': "#A0FFA0",
            'halogen': "#FFFF99",
            'noble_gas': "#C8A2C8",
            'lanthanide': "#FFBFFF",
            'actinide': "#FF99CC",
            'unknown': "#E8E8E8"
        }

        # Create ElementTiles for each element
        for element, (row, col) in element_positions.items():
            category = element_categories.get(element, 'unknown')
            color = colors.get(category, colors['unknown'])

            # Use KherveDB's methods to get element data
            atomic_number = self.kherve_instance.get_atomic_number(element)
            core_level = self.kherve_instance.get_main_core_level(element)
            binding_energy = self.kherve_instance.get_main_core_binding_energy(element)

            # Check if element is in our dataset
            enabled = element in self.get_available_elements()

            # Create ElementTile using KherveDB's ElementTile class
            tile = ElementTile(pt_panel, element, color, enabled, atomic_number, core_level, binding_energy)

            # Set callbacks to use our methods instead of KherveDB's
            tile.set_click_callback(self.on_element_click_survey)
            tile.set_double_click_callback(self.on_element_double_click_survey)

            # Bind hover events
            tile.Bind(wx.EVT_ENTER_WINDOW, self.OnElementHover)
            tile.Bind(wx.EVT_LEAVE_WINDOW, self.OnElementLeave)

            # Add to grid
            pt_sizer.Add(tile, pos=(row, col), flag=wx.EXPAND)
            self.element_buttons[element] = tile

            # Store original color and button state
            self.original_colors[element] = color
            self.button_states[element] = False

        # Set the sizer for the panel
        pt_panel.SetSizer(pt_sizer)
        sizer.Add(pt_panel, 0, wx.EXPAND | wx.ALL, 5)

    def on_element_click_survey(self, element):
        """Handle element tile clicks - pass element directly"""
        print(f"DEBUG: on_element_click_survey called with element: '{element}'")

        if not element or element.strip() == '':
            print("DEBUG: Element is empty in callback!")
            return

        # Instead of creating a mock event, let's call the logic directly
        self.handle_element_selection(element, is_tile=True)

    def handle_element_selection(self, element, is_tile=True):
        """Handle element selection logic - matches backup.py"""
        print(f"Element clicked: {element}")

        # Get the actual object
        if element in self.element_buttons:
            obj = self.element_buttons[element]
        else:
            print(f"DEBUG: Element {element} not found in element_buttons")
            return

        # Initialize button state if not exists
        if element not in self.button_states:
            self.button_states[element] = False

        # Toggle the button state
        self.button_states[element] = not self.button_states[element]

        if self.button_states[element]:
            # Set to selected color (green) and show lines
            if is_tile:
                obj.color = wx.Colour(0, 255, 0)  # Green
                obj.Refresh()
            else:
                obj.SetBackgroundColour(wx.GREEN)

            # Show red lines on plot
            self.plot_element_lines(element)

            # Get filtered transitions (this is the key fix)
            transitions = self.get_element_transitions(element)

            # Add transitions to list without clearing existing items (like backup.py)
            existing_items = [self.core_level_list.GetString(i) for i in range(self.core_level_list.GetCount())]
            for orbital, be in transitions:
                item = f"{element}{orbital}: {be:.1f} eV"
                if item not in existing_items:
                    self.core_level_list.Append(item)

        else:
            # Reset to original color and remove lines
            if is_tile:
                obj.color = self.original_colors.get(element, wx.Colour(200, 200, 200))
                obj.Refresh()
            else:
                obj.SetBackgroundColour(self.original_colors.get(element, wx.Colour(200, 200, 200)))

            # Remove element lines from plot
            self.remove_element_lines(element)

        # Update element info
        self.UpdateElementInfo(element)

    def OnElementClick_ElementTile(self, event):
        """Modified OnElementClick to work with ElementTiles"""
        tile = event.GetEventObject()
        element = tile.element  # ElementTile stores element as .element attribute

        print(f"Element clicked: {element}")

        # Handle tile color changes (ElementTiles use different methods)
        if self.button_states.get(element, False):
            # Reset to original color
            tile.color = self.original_colors[element]
            self.button_states[element] = False
        else:
            # Set to selected color (green)
            tile.color = wx.Colour(0, 255, 0)  # Green
            self.button_states[element] = True

        # Refresh the tile to show color change
        tile.Refresh()

        # Clear the core level list
        self.core_level_list.Clear()
        self.element_lines.clear()

        # Find transitions for this element
        transitions = []
        for (elem, orbital), data in self.library_data.items():
            if elem == element:
                # Check if 'Al1486' key exists, if not, use the first available instrument
                instrument = 'Al1486' if 'Al1486' in data else next(iter(data))
                if 'position' in data[instrument]:
                    be_value = float(data[instrument]['position'])
                    transitions.append((orbital, be_value))

        # Sort by binding energy
        transitions.sort(key=lambda x: x[1])

        # Add to list
        for orbital, be in transitions:
            display_text = f"{element} {orbital}: {be:.1f} eV"
            self.core_level_list.Append(display_text)
            self.element_lines[display_text] = (element, orbital, be)

        # Update element info
        self.UpdateElementInfo(element)

    def on_element_right_click(self, event):
        """Handle right-click on an element button"""
        button = event.GetEventObject()
        element = button.GetLabel()

        # Create a context menu
        menu = wx.Menu()
        item = menu.Append(wx.ID_ANY, f"Show {element} Properties")
        self.Bind(wx.EVT_MENU, lambda evt: self.show_element_properties(element), item)

        # Show the context menu at the button position
        button_pos = button.GetPosition()
        self.PopupMenu(menu, button_pos)
        menu.Destroy()

    def on_element_double_click_survey(self, element):
        """Handle element tile double-clicks if needed"""
        pass

    def create_fallback_periodic_table(self, parent, sizer):
        """Fallback if KherveDB import fails - use simple buttons"""
        # Your existing simple button code as fallback
        pass


    def show_element_properties(self, element):
        """Show properties for the given element using a wxPython window"""
        try:
            # Get element data using the static method from LibraryID
            element_data = PeriodicTableXPS.get_element_properties(None, element)

            # Create a wxPython dialog
            properties_window = wx.Dialog(self, title=f"Properties for {element}", size=(580, 600))

            # Create a panel
            panel = wx.Panel(properties_window)

            # Main sizer
            main_sizer = wx.BoxSizer(wx.VERTICAL)

            # Header with element symbol and name
            header_panel = wx.Panel(panel, style=wx.BORDER_NONE)
            header_panel.SetBackgroundColour("#f0f0f0")

            header_sizer = wx.BoxSizer(wx.HORIZONTAL)

            # Element symbol in large font
            symbol_label = wx.StaticText(header_panel, label=element)
            symbol_label.SetFont(wx.Font(40, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
            header_sizer.Add(symbol_label, 0, wx.ALL, 20)

            # Element name and atomic number
            info_sizer = wx.BoxSizer(wx.VERTICAL)

            name_label = wx.StaticText(header_panel, label=element_data["Name"])
            name_label.SetFont(wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
            info_sizer.Add(name_label, 0, wx.BOTTOM, 5)

            atomic_label = wx.StaticText(header_panel, label=f"Atomic Number: {element_data['Atomic Number']}")
            atomic_label.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
            info_sizer.Add(atomic_label, 0)

            header_sizer.Add(info_sizer, 0, wx.TOP, 20)
            header_panel.SetSizer(header_sizer)

            main_sizer.Add(header_panel, 0, wx.EXPAND)

            # Create a scrolled window for properties
            scroll_win = wx.ScrolledWindow(panel, style=wx.VSCROLL)
            scroll_win.SetScrollRate(0, 10)

            scroll_sizer = wx.BoxSizer(wx.VERTICAL)

            # Define property groups
            property_groups = {
                "Physical Properties": ["Atomic Mass", "Density", "Melting Point", "Boiling Point", "State at 20°C"],
                "Atomic Properties": ["Electron Configuration", "Electronegativity", "Atomic Radius",
                                      "Ionization Energy"],
                "General Information": ["Group", "Period", "Category", "Discovered By", "Year of Discovery"],
                "XPS Information": ["Common Core Levels", "Most Intense Line", "Typical FWHM", "Chemical Shift Range"]
            }

            # Add each property group
            for group_name, properties in property_groups.items():
                # Group header
                group_label = wx.StaticText(scroll_win, label=group_name)
                group_label.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
                group_label.SetBackgroundColour("#e0e0e0")
                scroll_sizer.Add(group_label, 0, wx.EXPAND | wx.ALL, 5)

                # Properties grid
                grid = wx.FlexGridSizer(rows=0, cols=2, vgap=5, hgap=10)
                grid.AddGrowableCol(1)

                for prop in properties:
                    prop_label = wx.StaticText(scroll_win, label=f"{prop}:")
                    value = element_data.get(prop, "N/A")
                    value_text = wx.StaticText(scroll_win, label=str(value))

                    grid.Add(prop_label, 0, wx.ALIGN_RIGHT | wx.ALL, 3)
                    grid.Add(value_text, 0, wx.EXPAND | wx.ALL, 3)

                scroll_sizer.Add(grid, 0, wx.EXPAND | wx.ALL, 5)

            # XPS Binding Energies Summary Section
            be_label = wx.StaticText(scroll_win, label="XPS Binding Energies")
            be_label.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
            be_label.SetBackgroundColour("#e0e0e0")
            scroll_sizer.Add(be_label, 0, wx.EXPAND | wx.ALL, 5)

            # Get transitions from core levels
            transitions = self.get_element_transitions(element)
            if transitions:
                # Create table header (similar to LibraryID implementation)
                table_sizer = wx.BoxSizer(wx.VERTICAL)

                # Headers
                header_grid = wx.FlexGridSizer(rows=1, cols=5, vgap=0, hgap=0)

                headers = ["Line", "Avg BE (eV)", "Min BE (eV)", "Max BE (eV)", "Count"]
                header_widths = [80, 80, 80, 80, 60]

                for i, header_text in enumerate(headers):
                    header_panel = wx.Panel(scroll_win, size=(header_widths[i], -1))
                    header_panel.SetBackgroundColour("#e8e8e8")
                    header_sizer = wx.BoxSizer(wx.VERTICAL)
                    header = wx.StaticText(header_panel, label=header_text, style=wx.ALIGN_CENTER)
                    header.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
                    header_sizer.Add(header, 1, wx.EXPAND | wx.ALL, 2)
                    header_panel.SetSizer(header_sizer)
                    header_grid.Add(header_panel, 0, wx.EXPAND | wx.ALL, 1)

                table_sizer.Add(header_grid, 0, wx.EXPAND)

                # Group transitions by line
                lines_data = {}
                for orbital, be in transitions:
                    line = orbital
                    if line not in lines_data:
                        lines_data[line] = {'values': [], 'count': 0}
                    lines_data[line]['values'].append(be)
                    lines_data[line]['count'] += 1

                # Create table rows
                for line, data in lines_data.items():
                    values = data['values']
                    avg_be = sum(values) / len(values) if values else 0
                    min_be = min(values) if values else 0
                    max_be = max(values) if values else 0
                    count = data['count']

                    row_grid = wx.FlexGridSizer(rows=1, cols=5, vgap=0, hgap=0)

                    # Create each cell in the row
                    cell_data = [
                        line,
                        f"{avg_be:.2f}",
                        f"{min_be:.2f}",
                        f"{max_be:.2f}",
                        str(count)
                    ]

                    for i, text in enumerate(cell_data):
                        cell_panel = wx.Panel(scroll_win, size=(header_widths[i], -1))
                        cell_panel.SetBackgroundColour(wx.WHITE)
                        cell_sizer = wx.BoxSizer(wx.VERTICAL)
                        cell = wx.StaticText(cell_panel, label=text)
                        cell_sizer.Add(cell, 1, wx.EXPAND | wx.ALL, 2)
                        cell_panel.SetSizer(cell_sizer)
                        row_grid.Add(cell_panel, 0, wx.EXPAND | wx.ALL, 1)

                    table_sizer.Add(row_grid, 0, wx.EXPAND)

                scroll_sizer.Add(table_sizer, 0, wx.EXPAND | wx.ALL, 5)

                # Now add individual transitions grid
                grid_label = wx.StaticText(scroll_win, label="Individual Core Level Transitions")
                grid_label.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
                scroll_sizer.Add(grid_label, 0, wx.EXPAND | wx.ALL, 5)

                transitions_grid = wx.FlexGridSizer(rows=0, cols=2, vgap=5, hgap=10)
                transitions_grid.AddGrowableCol(1)

                for orbital, be in transitions:
                    orbital_label = wx.StaticText(scroll_win, label=f"{orbital}:")
                    be_text = wx.StaticText(scroll_win, label=f"{be:.2f} eV")

                    transitions_grid.Add(orbital_label, 0, wx.ALIGN_RIGHT | wx.ALL, 3)
                    transitions_grid.Add(be_text, 0, wx.EXPAND | wx.ALL, 3)

                scroll_sizer.Add(transitions_grid, 0, wx.EXPAND | wx.ALL, 5)
            else:
                no_data = wx.StaticText(scroll_win, label="No XPS data available for this element")
                scroll_sizer.Add(no_data, 0, wx.ALL, 10)

            scroll_win.SetSizer(scroll_sizer)
            main_sizer.Add(scroll_win, 1, wx.EXPAND | wx.ALL, 10)

            # Close button
            close_btn = wx.Button(panel, wx.ID_CLOSE, "Close")
            close_btn.Bind(wx.EVT_BUTTON, lambda evt: properties_window.Close())
            main_sizer.Add(close_btn, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)

            panel.SetSizer(main_sizer)
            properties_window.ShowModal()
            properties_window.Destroy()

        except Exception as e:
            wx.MessageBox(f"Error showing element properties: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)




    def OnRemoveLastLabel(self, event):
        sheet_name = self.parent_window.sheet_combobox.GetValue()
        if 'Labels' in self.parent_window.Data['Core levels'][sheet_name]:
            labels = self.parent_window.Data['Core levels'][sheet_name]['Labels']
            if labels:
                labels.pop()
                # Clear all existing text annotations
                for text in self.parent_window.ax.texts[:]:
                    text.remove()
                # Redraw remaining labels
                maxY = max(self.parent_window.y_values)
                for label_data in labels:
                    self.parent_window.ax.text(
                        label_data['x'],
                        label_data['y'],
                        label_data['text'],
                        rotation=90,
                        va='bottom',
                        ha='center'
                    )
                self.parent_window.canvas.draw_idle()

    def OnRemoveAllLabels(self, event):
        sheet_name = self.parent_window.sheet_combobox.GetValue()
        if 'Labels' in self.parent_window.Data['Core levels'][sheet_name]:
            self.parent_window.Data['Core levels'][sheet_name]['Labels'] = []
            # Clear all text annotations
            for text in self.parent_window.ax.texts[:]:
                text.remove()
            self.parent_window.canvas.draw_idle()

    def get_element_transitions_OLD(self, element):
        allowed_orbitals = ['1s', '2s', '2p', '3s', '3p', '3d', '4s', '4p', '4d', '4f', '5s', '5p', '5d', '5f']
        transitions = {}
        photon_energy = self.parent_window.photons  # Get photon energy

        for (elem, orbital), data in self.library_data.items():
            if elem == element:
                # Choose instrument based on whether it's an Auger line
                if 'C-Any' in data:
                    instrument = 'C-Any'  # For Auger lines
                else:
                    instrument = 'C-Al1486' if 'C-Al1486' in data else next(iter(data))

                if 'position' in data[instrument] and float(data[instrument]['position']) >= 20:
                    # Check if it's a core level or Auger transition
                    is_auger = instrument == 'C-Any'
                    # is_auger = data[instrument]['Auger'] == '1'  # Corrected this line
                    orbital_lower = orbital.lower()

                    if is_auger:
                        # Subtract photon energy for Auger transitions
                        kinetic_energy = float(data[instrument]['position'])
                        binding_energy = photon_energy - kinetic_energy  # Convert KE to BE
                        transitions[orbital_lower] = binding_energy
                    else:
                        # For core levels
                        main_orbital = ''.join([c for c in orbital_lower if c.isalpha() or c.isdigit()])[:2]
                        if main_orbital in allowed_orbitals:
                            energy = float(data[instrument]['position'])
                            if main_orbital not in transitions or energy > transitions[main_orbital]:
                                transitions[main_orbital] = energy

        sorted_transitions = sorted(transitions.items(), key=lambda x: x[1])
        return sorted_transitions

    def get_element_transitions(self, element):
        """Get filtered transitions for an element - matches backup.py filtering"""
        allowed_orbitals = ['1s', '2s', '2p', '3s', '3p', '3d', '4s', '4p', '4d', '4f', '5s', '5p', '5d', '5f']
        transitions = {}
        photon_energy = getattr(self.parent_window, 'photons', 1486.6)  # Default Al Ka energy

        for (elem, orbital), data in self.library_data.items():
            if elem == element:
                # Choose instrument based on whether it's an Auger line
                if 'C-Any' in data:
                    instrument = 'C-Any'  # For Auger lines
                elif 'Al1486' in data:
                    instrument = 'Al1486'
                else:
                    instrument = next(iter(data))

                if 'position' in data[instrument] and float(data[instrument]['position']) >= 20:
                    orbital_lower = orbital.lower()

                    # Check if it's an Auger transition
                    is_auger = instrument == 'C-Any' or any(
                        orbital_lower.endswith(x) for x in ['kll', 'mnn', 'mvv', 'mnv', 'lmm'])

                    if is_auger:
                        # For Auger lines - filter to main types only (KLL not KLL1)
                        auger_main = None
                        if 'kll' in orbital_lower:
                            auger_main = 'kll'
                        elif 'mnn' in orbital_lower:
                            auger_main = 'mnn'
                        elif 'mvv' in orbital_lower:
                            auger_main = 'mvv'
                        elif 'mnv' in orbital_lower:
                            auger_main = 'mnv'
                        elif 'lmm' in orbital_lower:
                            auger_main = 'lmm'

                        if auger_main:
                            if instrument == 'C-Any':
                                # Kinetic energy for Auger
                                kinetic_energy = float(data[instrument]['position'])
                                binding_energy = photon_energy - kinetic_energy
                            else:
                                binding_energy = float(data[instrument]['position'])

                            # Only keep the highest energy for each main Auger type
                            if auger_main not in transitions or binding_energy > transitions[auger_main]:
                                transitions[auger_main] = binding_energy
                    else:
                        # For core levels - extract main orbital only (4p not 4p3/2)
                        main_orbital = ''.join([c for c in orbital_lower if c.isalpha() or c.isdigit()])

                        # Further filter to get just the main part (e.g., "4p" from "4p3/2")
                        import re
                        match = re.match(r'(\d+[spdf])', main_orbital)
                        if match:
                            main_orbital = match.group(1)

                            if main_orbital in allowed_orbitals:
                                energy = float(data[instrument]['position'])
                                # Only keep the highest energy for each main orbital
                                if main_orbital not in transitions or energy > transitions[main_orbital]:
                                    transitions[main_orbital] = energy

        # Sort transitions by binding energy
        sorted_transitions = sorted(transitions.items(), key=lambda x: x[1])
        return sorted_transitions

    def OnElementClick(self, event):
        """Handle element clicks - works with both buttons and ElementTiles"""
        obj = event.GetEventObject()

        print(f"DEBUG: Object type: {type(obj)}")
        print(f"DEBUG: Object attributes: {dir(obj)}")

        # Determine if it's a button or ElementTile and get the element
        if hasattr(obj, 'element'):  # ElementTile
            element = obj.element
            is_tile = True
            print(f"DEBUG: ElementTile element: '{element}'")
        elif hasattr(obj, 'GetLabel'):  # Regular button
            element = obj.GetLabel()
            is_tile = False
            print(f"DEBUG: Button label: '{element}'")
        else:
            print("DEBUG: Unknown object type in OnElementClick")
            return

        # Check if element is empty or None
        if not element or element.strip() == '':
            print("DEBUG: Element is empty! Aborting.")
            return

        print(f"Element clicked: {element}")

        # Initialize button state if not exists
        if element not in self.button_states:
            self.button_states[element] = False

        # Handle color changes based on object type
        if self.button_states[element]:
            # Reset to original color
            if is_tile:
                obj.color = self.original_colors.get(element, wx.Colour(200, 200, 200))
                obj.Refresh()
            else:
                obj.SetBackgroundColour(self.original_colors.get(element, wx.Colour(200, 200, 200)))
            self.button_states[element] = False
        else:
            # Set to selected color (green)
            if is_tile:
                obj.color = wx.Colour(0, 255, 0)  # Green
                obj.Refresh()
            else:
                obj.SetBackgroundColour(wx.GREEN)
            self.button_states[element] = True

        # Clear the core level list
        self.core_level_list.Clear()
        self.element_lines.clear()

        # Find transitions for this element
        transitions = []
        for (elem, orbital), data in self.library_data.items():
            if elem == element:
                # Check if 'Al1486' key exists, if not, use the first available instrument
                instrument = 'Al1486' if 'Al1486' in data else next(iter(data))
                if 'position' in data[instrument]:
                    be_value = float(data[instrument]['position'])
                    transitions.append((orbital, be_value))

        # Sort by binding energy
        transitions.sort(key=lambda x: x[1])

        # Add to list
        for orbital, be in transitions:
            display_text = f"{element} {orbital}: {be:.1f} eV"
            self.core_level_list.Append(display_text)
            self.element_lines[display_text] = (element, orbital, be)

        # Update element info
        self.UpdateElementInfo(element)

    def OnAddLabels_OLD(self, event):
        selections = self.core_level_list.GetSelections()
        for selection in selections:
            label = self.core_level_list.GetString(selection)
            element_orbital, be_str = label.split(':')
            print(f'Element Orbital: {element_orbital}')
            be = float(be_str.replace(' eV', ''))
            formatted_label = ".."

            # Extract element and orbital correctly
            import re
            match = re.match(r'([A-Z][a-z]*)(\d+[spdf])', element_orbital)
            if match:
                element, orbital = match.groups()
                formatted_label = f"{element} {orbital[0]}{orbital[1]}"  # e.g., "C 1 s"
            elif any(element_orbital.endswith(x) for x in ['kll', 'mnn', 'mvv', 'mnv', 'lmm']):
                auger = element_orbital[-3:]
                formatted_label = f"{element_orbital[:-3]} {auger.upper()}"

            if formatted_label != "..":
                # Get max intensity in ±5 eV range
                x_values = self.parent_window.x_values
                y_values = self.parent_window.y_values
                maxY = max(y_values)
                mask = (x_values >= be - 5) & (x_values <= be + 5)
                if np.any(mask):
                    local_max = np.max(y_values[mask])
                    # Add label at 1.2 times the local maximum height
                    self.parent_window.ax.text(be, local_max +0.05*maxY, formatted_label,
                                               rotation=90, va='bottom', ha='center')
                    self.parent_window.canvas.draw_idle()

                sheet_name = self.parent_window.sheet_combobox.GetValue()

                if 'Labels' not in self.parent_window.Data['Core levels'][sheet_name]:
                    # print('Labels not in sheetname')
                    self.parent_window.Data['Core levels'][sheet_name]['Labels'] = []

                self.parent_window.Data['Core levels'][sheet_name]['Labels'].append({
                    'text': formatted_label,
                    'x':be,
                    'y': local_max +0.05*maxY,
                    'rotation': 90  # Force 90 degrees
                })

    def OnAddLabels(self, event):
        selections = self.core_level_list.GetSelections()

        if not hasattr(self.parent_window, 'ax') or not hasattr(self.parent_window, 'x_values'):
            print("ERROR: Parent window doesn't have plot data")
            return

        for selection in selections:
            label = self.core_level_list.GetString(selection)
            element_orbital, be_str = label.split(':')
            print(f'Element Orbital: {element_orbital}')
            be = float(be_str.replace(' eV', '').strip())

            # Extract element and orbital correctly
            import re
            match = re.match(r'([A-Z][a-z]*)(\s*)(\d*[spdf]+)', element_orbital.strip())
            if match:
                element, space, orbital = match.groups()
                formatted_label = f"{element} {orbital}"
            else:
                formatted_label = element_orbital.strip()

            print(f"Adding label: {formatted_label} at {be} eV")

            # Get max intensity in ±5 eV range
            try:
                x_values = self.parent_window.x_values
                y_values = self.parent_window.y_values
                maxY = max(y_values)

                # Create mask for the region around the peak
                mask = (x_values >= be - 5) & (x_values <= be + 5)
                if np.any(mask):
                    local_max = np.max(y_values[mask])
                    # Add label at 1.2 times the local maximum height
                    label_y = local_max + 0.05 * maxY

                    self.parent_window.ax.text(be, label_y, formatted_label,
                                               rotation=90, va='bottom', ha='center',
                                               fontsize=8, color='blue')
                    self.parent_window.canvas.draw_idle()

                    # Store label data
                    sheet_name = self.parent_window.sheet_combobox.GetValue()
                    if 'Labels' not in self.parent_window.Data['Core levels'][sheet_name]:
                        self.parent_window.Data['Core levels'][sheet_name]['Labels'] = []

                    self.parent_window.Data['Core levels'][sheet_name]['Labels'].append({
                        'text': formatted_label,
                        'x': be,
                        'y': label_y,
                        'rotation': 90
                    })

                    print(f"Successfully added label {formatted_label} at ({be}, {label_y})")
                else:
                    print(f"No data points found near {be} eV")

            except Exception as e:
                print(f"Error adding label: {str(e)}")


    def add_peak_to_grid(self, peak_name):

        # Add peak to window.Data first
        sheet_name = self.parent_window.sheet_combobox.GetValue()

        # Initialize the full structure if it doesn't exist
        if 'Fitting' not in self.parent_window.Data['Core levels'][sheet_name]:
            self.parent_window.Data['Core levels'][sheet_name]['Fitting'] = {}
        if 'Peaks' not in self.parent_window.Data['Core levels'][sheet_name]['Fitting']:
            self.parent_window.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = {}

        # Better element and orbital extraction
        import re
        match = re.match(r'([A-Z][a-z]*)(\d+[spdf])', peak_name)
        if match:
            element, orbital = match.groups()
        else:
            return False

        position = None
        for (elem, orb), data in self.library_data.items():
            if elem == element and orb.lower() == orbital.lower():
                instrument = 'Al' if 'Al' in data else next(iter(data))
                if 'position' in data[instrument]:
                    position = float(data[instrument]['position'])
                    break

        if position:
            # Add to window.Data
            peak_data = {
                'Position': position,
                'Height': 0,
                'FWHM': 2.0,
                'L/G': 30,
                'Area': 0,
                'Fitting Model': 'SurveyID'
            }
            self.parent_window.Data['Core levels'][sheet_name]['Fitting']['Peaks'][peak_name] = peak_data

            # Add to grid
            current_rows = self.parent_window.peak_params_grid.GetNumberRows()

            self.parent_window.peak_params_grid.AppendRows(2)

            # Make sure to add the letter ID
            letter_id = chr(65 + (current_rows // 2))  # A, B, C, etc.

            self.parent_window.peak_params_grid.SetCellValue(current_rows, 0, letter_id)
            self.parent_window.peak_params_grid.SetCellValue(current_rows, 1, peak_name)
            self.parent_window.peak_params_grid.SetCellValue(current_rows, 2, f"{position:.2f}")
            self.parent_window.peak_params_grid.SetCellValue(current_rows, 4, "2.0")  # FWHM
            self.parent_window.peak_params_grid.SetCellValue(current_rows, 5, "30")  # L/G
            self.parent_window.peak_params_grid.SetCellValue(current_rows, 13, "SurveyID")

            # Set constraint row background color
            for col in range(self.parent_window.peak_params_grid.GetNumberCols()):
                self.parent_window.peak_params_grid.SetCellBackgroundColour(current_rows + 1, col,
                                                                            wx.Colour(200, 245, 228))

            # Update peak count
            self.parent_window.peak_count = current_rows // 2 + 1

            self.parent_window.peak_params_grid.ForceRefresh()
            return True
        else:
            print(f"No position found for peak {peak_name}")
            return False

    def OnAddPeak(self, event):
        selections = self.core_level_list.GetSelections()
        sheet_name = self.parent_window.sheet_combobox.GetValue().lower()

        if any(x in sheet_name for x in ['survey', 'wide']):
            for selection in selections:
                label = self.core_level_list.GetString(selection)
                element = label.split(':')[0]
                # Check if it's a main core level
                if any(element.endswith(x) for x in ['1s', '2p', '3d', '4f']):
                    self.add_peak_to_grid(element)

    def OnRemoveSelected(self, event):
        selections = list(self.core_level_list.GetSelections())
        selections.reverse()  # Remove from bottom to top to avoid index issues
        for selection in selections:
            self.core_level_list.Delete(selection)

    def OnRemoveAll(self, event):
        self.core_level_list.Clear()


    def get_main_core_level(self, element):
        main_core_levels = {
            'Li': '1s', 'Be': '1s', 'B': '1s', 'C': '1s', 'N': '1s', 'O': '1s', 'F': '1s', 'Ne': '1s',
            'Na': '1s', 'Mg': '1s', 'Al': '2p', 'Si': '2p', 'P': '2p', 'S': '2p', 'Cl': '2p',
            'K': '2p', 'Ca': '2p', 'Sc': '2p', 'Ti': '2p', 'V': '2p', 'Cr': '2p', 'Mn': '2p',
            'Fe': '2p', 'Co': '2p', 'Ni': '2p', 'Cu': '2p', 'Zn': '2p', 'Ga': '2p', 'Ge': '3d',
            'As': '3d', 'Se': '3d', 'Br': '3d', 'Sr': '3d', 'Y': '3d', 'Zr': '3d', 'Nb': '3d',
            'Mo': '3d', 'Tc': '3d', 'Ru': '3d', 'Rh': '3d', 'Pd': '3d', 'Ag': '3d', 'Cd': '3d',
            'In': '3d', 'Sn': '3d', 'Sb': '3d', 'Te': '3d', 'I': '3d', 'Xe': '3d', 'Cs': '3d',
            'Ba': '3d', 'La': '3d', 'W': '4f'
        }
        return main_core_levels.get(element)

    def plot_element_lines(self, element):
        transitions = self.get_element_transitions(element)

        if transitions:
            xmin, xmax = self.parent_window.ax.get_xlim()
            ymin, ymax = self.parent_window.ax.get_ylim()

            # Filter transitions within xmin and xmax
            valid_transitions = [t for t in transitions if xmax <= t[1] <= xmin]

            if valid_transitions:
                # Get RSF values for each transition
                orbital_list = [t[0] for t in valid_transitions]
                rsf_values = self.get_rsf_values(element, orbital_list)

                if rsf_values:
                    max_rsf = max(rsf_values)

                    # Initialize element_lines if not exists
                    if element not in self.element_lines:
                        self.element_lines[element] = []

                    for (orbital, be), rsf in zip(valid_transitions, rsf_values):
                        # For RSF = 0, use 0.1 of the max scale; otherwise normalize normally
                        if rsf == 0:
                            intensity = 0.1 * self.intensity_scale * (ymax - ymin)
                        else:
                            intensity = (rsf / max_rsf) * self.intensity_scale * (ymax - ymin)

                        # Draw the vertical line
                        line = self.parent_window.ax.vlines(be, ymin - intensity, ymin + intensity,
                                                            color='red', linewidth=1)
                        self.element_lines[element].append(line)

                        # Add text label
                        text_y = ymin + 0.01 * (ymax - ymin)
                        text_label = f"{element}{orbital}"

                        text = self.parent_window.ax.text(
                            be + 0.1,  # Slightly to the right of the line
                            text_y,  # At the bottom of the plot
                            text_label,
                            rotation=90,
                            va='bottom',
                            ha='right',
                            fontsize=7,
                            color='red'
                        )
                        self.element_lines[element].append(text)

                    self.parent_window.canvas.draw_idle()

    def remove_element_lines(self, element):
        if element in self.element_lines:
            for line in self.element_lines[element]:
                line.remove()
            del self.element_lines[element]
            self.parent_window.canvas.draw_idle()

    def reset_all_buttons(self):
        for element, button in self.button_states.items():
            if button:
                self.button_states[element] = False
                btn = self.FindWindowByLabel(element)
                if btn:
                    # Restore original color
                    original_color = self.original_colors.get(element, wx.WHITE)
                    btn.SetBackgroundColour(original_color)
                    btn.Refresh()

        for element, lines in self.element_lines.items():
            for line in lines:
                line.remove()
        self.element_lines.clear()
        self.parent_window.canvas.draw_idle()

    def get_rsf_values(self, element, orbitals):
        rsf_values = []

        for requested_orbital in orbitals:
            found_rsf = None

            for (elem, orbital), data in self.library_data.items():
                if elem == element:
                    orbital_lower = orbital.lower()
                    requested_lower = requested_orbital.lower()

                    if orbital_lower == requested_lower or requested_lower in orbital_lower:
                        # Check if 'Al1486' key exists, if not, use the first available instrument
                        instrument = 'Al1486' if 'Al1486' in data else next(iter(data))

                        if 'rsf' in data[instrument]:
                            rsf_value = float(data[instrument]['rsf'])
                            found_rsf = rsf_value
                            break

            if found_rsf is not None:
                rsf_values.append(found_rsf)
            else:
                rsf_values.append(0.1)  # Default value

        return rsf_values

    def OnElementHover(self, event):
        obj = event.GetEventObject()

        # Handle both ElementTiles and regular buttons
        if hasattr(obj, 'element'):  # ElementTile
            element = obj.element
        elif hasattr(obj, 'GetLabel'):  # Regular button
            element = obj.GetLabel()
        else:
            return

        self.UpdateElementInfo(element)

    def OnElementLeave(self, event):
        self.info_text1.SetLabelMarkup("")
        self.info_text2.SetLabelMarkup("")
        self.Layout()

    def OnIntensityIncrease(self, event):
        """Increase the intensity scaling factor by 0.1"""
        self.intensity_scale = min(2.0, self.intensity_scale + 0.1)  # Cap at 2.0
        self.intensity_display.SetLabel(f"{self.intensity_scale:.1f}")
        self.update_all_element_lines()

    def OnIntensityDecrease(self, event):
        """Decrease the intensity scaling factor by 0.1"""
        self.intensity_scale = max(0.1, self.intensity_scale - 0.1)  # Minimum 0.1
        self.intensity_display.SetLabel(f"{self.intensity_scale:.1f}")
        self.update_all_element_lines()

    def update_all_element_lines(self):
        """Redraw all currently visible element lines with new intensity"""
        # Get all currently selected elements
        selected_elements = [element for element, state in self.button_states.items() if state]

        if selected_elements:
            # Remove all existing lines
            for element in selected_elements:
                if element in self.element_lines:
                    for line_obj in self.element_lines[element]:
                        line_obj.remove()
                    del self.element_lines[element]

            # Redraw with new intensity
            for element in selected_elements:
                self.plot_element_lines(element)

            print(f"Updated line intensities to scale factor: {self.intensity_scale}")

    def UpdateElementInfo(self, element):
        element_names = {
            'H': 'Hydrogen', 'He': 'Helium', 'Li': 'Lithium', 'Be': 'Beryllium', 'B': 'Boron',
            'C': 'Carbon', 'N': 'Nitrogen', 'O': 'Oxygen', 'F': 'Fluorine', 'Ne': 'Neon',
            'Na': 'Sodium', 'Mg': 'Magnesium', 'Al': 'Aluminum', 'Si': 'Silicon', 'P': 'Phosphorus',
            'S': 'Sulfur', 'Cl': 'Chlorine', 'Ar': 'Argon', 'K': 'Potassium', 'Ca': 'Calcium',
            'Sc': 'Scandium', 'Ti': 'Titanium', 'V': 'Vanadium', 'Cr': 'Chromium', 'Mn': 'Manganese',
            'Fe': 'Iron', 'Co': 'Cobalt', 'Ni': 'Nickel', 'Cu': 'Copper', 'Zn': 'Zinc',
            'Ga': 'Gallium', 'Ge': 'Germanium', 'As': 'Arsenic', 'Se': 'Selenium', 'Br': 'Bromine',
            'Kr': 'Krypton', 'Rb': 'Rubidium', 'Sr': 'Strontium', 'Y': 'Yttrium', 'Zr': 'Zirconium',
            'Nb': 'Niobium', 'Mo': 'Molybdenum', 'Ru': 'Ruthenium', 'Rh': 'Rhodium', 'Pd': 'Palladium',
            'Ag': 'Silver', 'Cd': 'Cadmium', 'In': 'Indium', 'Sn': 'Tin', 'Sb': 'Antimony',
            'Te': 'Tellurium', 'I': 'Iodine', 'Xe': 'Xenon', 'Cs': 'Cesium', 'Ba': 'Barium',
            'La': 'Lanthanum', 'Ce': 'Cerium', 'Pr': 'Praseodymium', 'Nd': 'Neodymium', 'Pm': 'Promethium',
            'Sm': 'Samarium', 'Eu': 'Europium', 'Gd': 'Gadolinium', 'Tb': 'Terbium', 'Dy': 'Dysprosium',
            'Ho': 'Holmium', 'Er': 'Erbium', 'Tm': 'Thulium', 'Yb': 'Ytterbium', 'Lu': 'Lutetium',
            'Hf': 'Hafnium', 'Ta': 'Tantalum', 'W': 'Tungsten', 'Re': 'Rhenium', 'Os': 'Osmium',
            'Ir': 'Iridium', 'Pt': 'Platinum', 'Au': 'Gold', 'Hg': 'Mercury', 'Tl': 'Thallium',
            'Pb': 'Lead', 'Bi': 'Bismuth', 'At': 'Astatine', 'Rn': 'Radon', 'Ra': 'Radium',
            'Th': 'Thorium', 'U': 'Uranium', 'Np': 'Neptunium', 'Pu': 'Plutonium', 'Am': 'Americium',
            'Cm': 'Curium'
        }

        transitions = self.get_element_transitions(element)
        if transitions:
            info1 = f"<b>{element_names.get(element, element)}</b>: "
            info2 = ", ".join(f"{orbital}: {be:.1f} eV" for orbital, be in transitions)

            # Split info2 into two lines if it's too long
            max_line_length = 60
            if len(info2) > max_line_length:
                split_point = info2.rfind(", ", 0, max_line_length) + 2
                info1 += info2[:split_point]
                info2 = info2[split_point:]
            else:
                info1 += info2
                info2 = ""

            self.info_text1.SetLabelMarkup(info1)
            self.info_text2.SetLabelMarkup(info2)
        else:
            self.info_text1.SetLabelMarkup(f"<b>{element_names.get(element, element)}</b>: No BE transitions found")
            self.info_text2.SetLabelMarkup("")
        self.Layout()

    def Close(self, force=False):
        self.reset_all_buttons()
        super().Close(force)

    def OnClose(self, event):
        self.reset_all_buttons()
        self.Destroy()


def open_periodic_table(parent):
    periodic_table = PeriodicTableWindow(parent)
    periodic_table.Show()