#!/usr/bin/env python3
import wx
import wx.grid
import pandas as pd
import webbrowser
import re
from typing import Dict, List, Tuple
import os
import sys
import pyperclip
import matplotlib
import wx.adv

matplotlib.use('WXAgg')
from matplotlib.figure import Figure
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.backends.backend_wxagg import NavigationToolbar2WxAgg as NavigationToolbar
import numpy as np
import platform
import wx.html2


class PeriodicTableXPS(wx.Frame):
    def __init__(self):
        super().__init__(None, title="KherveDB Library: How I wish NIST would look like",
                         size=(620, 720))

        if platform.system() == 'Darwin':  # Mac OS
            window_size = (680, 720)
            # Set minimum and maximum sizes
            self.SetMinSize((620, 660))
            self.SetMaxSize((620, 10000))

        else:
            window_size = (705, 720)
            # Set minimum and maximum sizes
            self.SetMinSize((705, 660))
            self.SetMaxSize((705, 10000))

        # Center the window
        self.Centre()

        # Create menu bar
        self.create_menu()

        # Load data
        self.load_data()

        # Create main panel
        self.panel = wx.Panel(self)
        self.main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Create UI components
        self.create_periodic_table()
        self.create_search_area()
        self.create_results_table()

        # Set sizer
        self.panel.SetSizer(self.main_sizer)

        # Initialize variables
        self.selected_element = None
        self.selected_line = None

        # Track property dialog
        self.property_dialog = None
        self.property_dialog_position = None

        # Track property dialog
        self.property_dialog = None
        self.property_dialog_position = None
        self.property_dialog_tab_index = 0  # Remember which tab was selected

        # Bind close event
        self.Bind(wx.EVT_CLOSE, self.on_close)

    def on_close(self, event):
        """Handle window close event"""
        self.Destroy()

    def detect_mac_os(self):
        """Detect if running on macOS"""
        return platform.system() == 'Darwin'

    def create_menu(self):
        """Create the application menu bar"""
        menubar = wx.MenuBar()

        # File menu
        file_menu = wx.Menu()
        exit_item = file_menu.Append(wx.ID_EXIT, 'E&xit\tCtrl+Q', 'Exit application')
        self.Bind(wx.EVT_MENU, lambda e: self.Close(), exit_item)
        menubar.Append(file_menu, '&File')

        # Help menu
        help_menu = wx.Menu()
        about_item = help_menu.Append(wx.ID_ABOUT, '&About', 'About this application')
        self.Bind(wx.EVT_MENU, self.show_about, about_item)
        menubar.Append(help_menu, '&Help')

        self.SetMenuBar(menubar)

    def show_about(self, event):
        """Display information about the application"""
        about_text = """My KherveDB Library

This application provides access to the NIST X-ray Photoelectron Spectroscopy (XPS) 
binding energy database recorded in 2019, before it was shut down during the Trump 
administration.

The original NIST database is an invaluable resource for scientists and researchers 
in materials science, chemistry, and physics fields, providing standard reference 
data for XPS analysis.

This application aims to preserve and provide easy access to this important 
scientific resource.

This application also provide rapid access to the website XPSfitting from M. Biesinger and to
the webite of Thermo Knowledge.

Developer: Gwilherm Kerherve
Version: 1.1"""

        wx.MessageBox(about_text, "About My KherveDB Library",
                      wx.OK | wx.ICON_INFORMATION)

    def load_data(self):
        """Load XPS data from file"""
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            base_path = os.path.dirname(current_dir)

        possible_paths = [
            os.path.join(base_path, "NIST_BE.parquet"),
            os.path.join(base_path, "libraries", "NIST_BE.parquet"),
            os.path.join(base_path, "..", "Resources", "NIST_BE.parquet"),  # Mac app bundle
            os.path.join(base_path, "NIST_BE.xlsx"),
            os.path.join(base_path, "libraries", "NIST_BE.xlsx"),
            os.path.join(base_path, "..", "Resources", "NIST_BE.xlsx")  # Mac app bundle
        ]

        data_found = False
        for data_path in possible_paths:
            if os.path.exists(data_path):
                try:
                    if data_path.endswith('.parquet'):
                        self.df = pd.read_parquet(data_path)
                        print(f'Loaded the .parquet NIST library')
                    else:
                        self.df = pd.read_excel(data_path)
                        print(f'Loaded the .xlsx NIST library')

                    self.elements = sorted(self.df['Element'].unique())
                    self.lines = sorted(self.df['Line'].unique())
                    data_found = True
                    break
                except Exception as e:
                    continue

        if not data_found:
            wx.MessageBox("Failed to load data: NIST_BE file not found",
                          "Error", wx.OK | wx.ICON_ERROR)
            self.Close()

    def create_periodic_table(self):
        """Create the periodic table with colored buttons"""
        # Create frame for periodic table
        pt_panel = wx.Panel(self.panel, style=wx.BORDER_SUNKEN)
        pt_panel.SetBackgroundColour(wx.Colour(230, 230, 230))

        # Use grid sizer for periodic table layout
        self.pt_sizer = wx.GridBagSizer(1, 1)

        # Define element positions
        self.element_positions = self.get_element_positions()

        # Define color schemes
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

        # Define element categories
        element_categories = self.get_element_categories()

        # Create buttons for each element
        self.element_buttons = {}
        for element, (row, col) in self.element_positions.items():
            category = element_categories.get(element, 'unknown')
            color = colors.get(category, colors['unknown'])

            # Create button
            btn = self.create_element_button(pt_panel, element, color)

            # Check if element is in our dataset
            if element not in self.elements:
                btn.Enable(False)

            # Add to grid
            self.pt_sizer.Add(btn, pos=(row, col), flag=wx.EXPAND)
            self.element_buttons[element] = btn

        # Add labels for lanthanides and actinides
        label1 = wx.StaticText(pt_panel, label="*")
        label2 = wx.StaticText(pt_panel, label="**")
        self.pt_sizer.Add(label1, pos=(6, 2))
        self.pt_sizer.Add(label2, pos=(7, 2))

        pt_panel.SetSizer(self.pt_sizer)
        self.main_sizer.Add(pt_panel, 0, wx.EXPAND, 0)

    def get_atomic_number(self, element_symbol):
        """Get atomic number for an element"""
        atomic_numbers = {
            'H': 1, 'He': 2, 'Li': 3, 'Be': 4, 'B': 5, 'C': 6, 'N': 7, 'O': 8, 'F': 9, 'Ne': 10,
            'Na': 11, 'Mg': 12, 'Al': 13, 'Si': 14, 'P': 15, 'S': 16, 'Cl': 17, 'Ar': 18,
            'K': 19, 'Ca': 20, 'Sc': 21, 'Ti': 22, 'V': 23, 'Cr': 24, 'Mn': 25, 'Fe': 26,
            'Co': 27, 'Ni': 28, 'Cu': 29, 'Zn': 30, 'Ga': 31, 'Ge': 32, 'As': 33, 'Se': 34,
            'Br': 35, 'Kr': 36, 'Rb': 37, 'Sr': 38, 'Y': 39, 'Zr': 40, 'Nb': 41, 'Mo': 42,
            'Tc': 43, 'Ru': 44, 'Rh': 45, 'Pd': 46, 'Ag': 47, 'Cd': 48, 'In': 49, 'Sn': 50,
            'Sb': 51, 'Te': 52, 'I': 53, 'Xe': 54, 'Cs': 55, 'Ba': 56, 'La': 57, 'Ce': 58,
            'Pr': 59, 'Nd': 60, 'Pm': 61, 'Sm': 62, 'Eu': 63, 'Gd': 64, 'Tb': 65, 'Dy': 66,
            'Ho': 67, 'Er': 68, 'Tm': 69, 'Yb': 70, 'Lu': 71, 'Hf': 72, 'Ta': 73, 'W': 74,
            'Re': 75, 'Os': 76, 'Ir': 77, 'Pt': 78, 'Au': 79, 'Hg': 80, 'Tl': 81, 'Pb': 82,
            'Bi': 83, 'Po': 84, 'At': 85, 'Rn': 86, 'Fr': 87, 'Ra': 88, 'Ac': 89, 'Th': 90,
            'Pa': 91, 'U': 92, 'Np': 93, 'Pu': 94, 'Am': 95, 'Cm': 96, 'Bk': 97, 'Cf': 98,
            'Es': 99, 'Fm': 100, 'Md': 101, 'No': 102, 'Lr': 103, 'Rf': 104, 'Db': 105,
            'Sg': 106, 'Bh': 107, 'Hs': 108, 'Mt': 109, 'Ds': 110, 'Rg': 111, 'Cn': 112,
            'Nh': 113, 'Fl': 114, 'Mc': 115, 'Lv': 116, 'Ts': 117, 'Og': 118
        }
        return atomic_numbers.get(element_symbol, 0)

    def get_main_core_level(self, element_symbol):
        """Get the main XPS core level for an element"""
        core_levels = {
            'H': 'N.D.', 'He': 'N.D.', 'Li': '1s', 'Be': '1s', 'B': '1s', 'C': '1s', 'N': '1s', 'O': '1s', 'F': '1s',
            'Ne': '1s',
            'Na': '1s', 'Mg': '2p', 'Al': '2p', 'Si': '2p', 'P': '2p3/2', 'S': '2p3/2', 'Cl': '2p3/2', 'Ar': '2p3/2',
            'K': '2p3/2', 'Ca': '2p3/2', 'Sc': '2p3/2', 'Ti': '2p3/2', 'V': '2p3/2', 'Cr': '2p3/2', 'Mn': '2p3/2',
            'Fe': '2p3/2',
            'Co': '2p3/2', 'Ni': '2p3/2', 'Cu': '2p3/2', 'Zn': '2p3/2', 'Ga': '2p3/2', 'Ge': '3d', 'As': '3d5/2',
            'Se': '3d5/2',
            'Br': '3d5/2', 'Kr': '3d5/2', 'Rb': '3d5/2', 'Sr': '3d5/2', 'Y': '3d5/2', 'Zr': '3d5/2', 'Nb': '3d5/2',
            'Mo': '3d5/2',
            'Tc': '3d5/2', 'Ru': '3d5/2', 'Rh': '3d5/2', 'Pd': '3d5/2', 'Ag': '3d5/2', 'Cd': '3d5/2', 'In': '3d5/2',
            'Sn': '3d5/2',
            'Sb': '3d5/2', 'Te': '3d5/2', 'I': '3d5/2', 'Xe': '3d5/2', 'Cs': '3d5/2', 'Ba': '3d5/2', 'La': '3d5/2',
            'Ce': '3d5/2',
            'Pr': '3d5/2', 'Nd': '3d5/2', 'Pm': '3d5/2', 'Sm': '3d5/2', 'Eu': '3d5/2', 'Gd': '4d', 'Tb': '4d',
            'Dy': '4d',
            'Ho': '4d', 'Er': '4d', 'Tm': '4d', 'Yb': '4d', 'Lu': '4f7/2', 'Hf': '4f7/2', 'Ta': '4f7/2', 'W': '4f7/2',
            'Re': '4f7/2', 'Os': '4f7/2', 'Ir': '4f7/2', 'Pt': '4f7/2', 'Au': '4f7/2', 'Hg': '4f7/2', 'Tl': '4f7/2',
            'Pb': '4f7/2',
            'Bi': '4f7/2', 'Po': '4f7/2', 'At': '4f7/2', 'Rn': '4f7/2', 'Fr': '4f7/2', 'Ra': '4f7/2', 'Ac': '4f7/2',
            'Th': '4f7/2',
            'Pa': '4f7/2', 'U': '4f7/2', 'Np': '4f7/2', 'Pu': '4f7/2', 'Am': '4f7/2', 'Cm': '4f7/2', 'Bk': '4f7/2',
            'Cf': '4f7/2',
            'Es': '4f7/2', 'Fm': '4f7/2', 'Md': '4f7/2', 'No': '4f7/2', 'Lr': '4f7/2'
        }
        return core_levels.get(element_symbol, 'N.D.')

    def get_main_core_binding_energy(self, element_symbol):
        """Get the binding energy (eV) for the main XPS core level of an element"""
        binding_energies = {
            'H': 'N.D.', 'He': 'N.D.', 'Li': '56', 'Be': '112', 'B': '189', 'C': '285', 'N': '398', 'O': '531',
            'F': '685', 'Ne': '863',
            'Na': '1072', 'Mg': '50', 'Al': '73', 'Si': '99', 'P': '130', 'S': '164', 'Cl': '199', 'Ar': '242',
            'K': '294', 'Ca': '347', 'Sc': '399', 'Ti': '454', 'V': '512', 'Cr': '574', 'Mn': '639', 'Fe': '707',
            'Co': '778', 'Ni': '853', 'Cu': '933', 'Zn': '1022', 'Ga': '1117', 'Ge': '29', 'As': '42', 'Se': '57',
            'Br': '69', 'Kr': '87', 'Rb': '111', 'Sr': '134', 'Y': '158', 'Zr': '179', 'Nb': '202', 'Mo': '228',
            'Tc': 'N.D.', 'Ru': '280', 'Rh': '307', 'Pd': '335', 'Ag': '368', 'Cd': '405', 'In': '444', 'Sn': '485',
            'Sb': '528', 'Te': '573', 'I': '619', 'Xe': '670', 'Cs': '726', 'Ba': '781', 'La': '836', 'Ce': '884',
            'Pr': '932', 'Nd': '981', 'Pm': '1033', 'Sm': '1081', 'Eu': '1126', 'Gd': '140', 'Tb': '146', 'Dy': '152',
            'Ho': '160', 'Er': '167', 'Tm': '175', 'Yb': '182', 'Lu': '7', 'Hf': '14', 'Ta': '22', 'W': '31',
            'Re': '40', 'Os': '51', 'Ir': '61', 'Pt': '71', 'Au': '84', 'Hg': '101', 'Tl': '118', 'Pb': '137',
            'Bi': '157', 'Po': 'N.D.', 'At': 'N.D.', 'Rn': 'N.D.', 'Fr': 'N.D.', 'Ra': 'N.D.', 'Ac': '307', 'Th': '333',
            'Pa': '358', 'U': '377', 'Np': '403', 'Pu': '425', 'Am': '448', 'Cm': '472', 'Bk': '499', 'Cf': '523',
            'Es': '550', 'Fm': 'N.D.', 'Md': 'N.D.', 'No': 'N.D.', 'Lr': 'N.D.'
        }
        return binding_energies.get(element_symbol, 'N.D.')

    def create_element_button(self, parent, element, color):
        """Create a colored tile for an element"""
        # Check if element is in our dataset
        enabled = element in self.elements

        # Get atomic number
        atomic_number = self.get_atomic_number(element)
        core_level = self.get_main_core_level(element)
        binding_energy = self.get_main_core_binding_energy(element)

        #  Create custom tile with all information
        tile = ElementTile(parent, element, color, enabled, atomic_number, core_level, binding_energy)

        # Set callbacks
        tile.set_click_callback(self.select_element)
        tile.set_double_click_callback(self.on_element_double_click)

        return tile

    def get_element_positions(self) -> Dict[str, Tuple[int, int]]:
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
            'Lv': 'post_transition', 'Ts': 'post_transition', 'Og': 'noble_gas'
        }

        # Add lanthanides and actinides
        lanthanides = ['La', 'Ce', 'Pr', 'Nd', 'Pm', 'Sm', 'Eu', 'Gd', 'Tb', 'Dy', 'Ho', 'Er', 'Tm', 'Yb', 'Lu']
        actinides = ['Ac', 'Th', 'Pa', 'U', 'Np', 'Pu', 'Am', 'Cm', 'Bk', 'Cf', 'Es', 'Fm', 'Md', 'No', 'Lr']

        for element in lanthanides:
            element_categories[element] = 'lanthanide'
        for element in actinides:
            element_categories[element] = 'actinide'

        return element_categories

    def create_search_area(self):
        """Create the search interface"""
        search_panel = wx.Panel(self.panel, style=wx.BORDER_SUNKEN)
        search_panel.SetBackgroundColour(wx.Colour(224, 224, 224))

        search_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Left side controls
        left_sizer = wx.GridBagSizer(5, 5)

        # Selected element
        left_sizer.Add(wx.StaticText(search_panel, label="Selected Element:"),
                       pos=(0, 0), flag=wx.ALIGN_CENTER_VERTICAL)
        self.element_label = wx.StaticText(search_panel, label="None",
                                           style=wx.BORDER_SUNKEN | wx.ST_NO_AUTORESIZE)
        self.element_label.SetBackgroundColour(wx.WHITE)
        self.element_label.SetMinSize((60, -1))
        left_sizer.Add(self.element_label, pos=(0, 1), flag=wx.EXPAND)

        # XPS Line selection
        left_sizer.Add(wx.StaticText(search_panel, label="XPS Line:"),
                       pos=(1, 0), flag=wx.ALIGN_CENTER_VERTICAL)
        self.line_combo = wx.ComboBox(search_panel, choices=['All Lines'] + list(self.lines),
                                      style=wx.CB_READONLY)
        self.line_combo.SetSelection(0)
        self.line_combo.Bind(wx.EVT_COMBOBOX, self.on_line_selected)
        left_sizer.Add(self.line_combo, pos=(1, 1), flag=wx.EXPAND)

        # Right side controls
        right_sizer = wx.GridBagSizer(5, 5)

        # Formula search
        right_sizer.Add(wx.StaticText(search_panel, label="Search Formula:"),
                        pos=(0, 0), flag=wx.ALIGN_CENTER_VERTICAL)
        self.formula_search = wx.TextCtrl(search_panel)
        self.formula_search.Bind(wx.EVT_TEXT, self.on_search_change)
        right_sizer.Add(self.formula_search, pos=(0, 1), flag=wx.EXPAND)

        # Name search
        right_sizer.Add(wx.StaticText(search_panel, label="Search Name:"),
                        pos=(1, 0), flag=wx.ALIGN_CENTER_VERTICAL)
        self.name_search = wx.TextCtrl(search_panel)
        self.name_search.Bind(wx.EVT_TEXT, self.on_search_change)
        right_sizer.Add(self.name_search, pos=(1, 1), flag=wx.EXPAND)

        # Buttons
        self.properties_btn = wx.Button(search_panel, label="Properties")
        self.properties_btn.Bind(wx.EVT_BUTTON, self.show_element_properties)
        right_sizer.Add(self.properties_btn, pos=(0, 2))

        self.plot_btn = wx.Button(search_panel, label="Plot Results")
        self.plot_btn.Bind(wx.EVT_BUTTON, self.plot_results)
        right_sizer.Add(self.plot_btn, pos=(1, 2))

        # Add to main sizer
        search_sizer.Add(left_sizer, 0, wx.ALL, 10)
        search_sizer.Add(right_sizer, 1, wx.ALL | wx.EXPAND, 10)

        search_panel.SetSizer(search_sizer)
        self.main_sizer.Add(search_panel, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 1)


    def create_results_table(self):
        """Create the results table"""
        # Create panel for results
        results_panel = wx.Panel(self.panel)
        results_sizer = wx.BoxSizer(wx.VERTICAL)

        # Create grid
        self.results_grid = wx.grid.Grid(results_panel)
        self.results_grid.CreateGrid(0, 6)

        # Set font for the grid (decrease default size by 1)
        default_font = self.results_grid.GetDefaultCellFont()
        new_font_size = default_font.GetPointSize() - 1
        grid_font = wx.Font(new_font_size, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)

        # Apply font to cells
        self.results_grid.SetDefaultCellFont(grid_font)

        # Set smaller height for column labels (header row)
        self.results_grid.SetColLabelSize(20)  # Adjust this value as needed (default is usually around 32)

        # Hide row labels (row numbers)
        self.results_grid.HideRowLabels()

        # Set column labels and platform-specific widths
        col_labels = ["", "Line", "BE (eV)", "Formula", "Name", "Journal"]

        # Platform-specific column widths
        import platform
        if platform.system() == 'Darwin':  # macOS
            col_widths = [25, 50, 60, 100, 170, 205]
        else:  # Windows and other systems
            col_widths = [25, 50, 60, 110, 180, 220]  # Slightly wider for Windows

        for i, (label, width) in enumerate(zip(col_labels, col_widths)):
            self.results_grid.SetColLabelValue(i, label)
            self.results_grid.SetColSize(i, width)

        # Make grid read-only
        self.results_grid.EnableEditing(False)

        # Bind events
        self.results_grid.Bind(wx.grid.EVT_GRID_CELL_LEFT_DCLICK, self.on_grid_double_click)
        self.results_grid.Bind(wx.grid.EVT_GRID_CELL_RIGHT_CLICK, self.on_grid_right_click)
        self.results_grid.Bind(wx.grid.EVT_GRID_LABEL_LEFT_CLICK, self.on_column_click)

        # Status bar
        self.status_text = wx.StaticText(results_panel, label="Ready")

        # Add to sizer
        results_sizer.Add(self.results_grid, 1, wx.ALL | wx.EXPAND, 5)
        results_sizer.Add(self.status_text, 0, wx.ALL | wx.EXPAND, 5)

        results_panel.SetSizer(results_sizer)
        self.main_sizer.Add(results_panel, 1, wx.ALL | wx.EXPAND, 1)

        # Sort tracking
        self.sort_column = None
        self.sort_ascending = True

    def select_element(self, element):
        """Handle element selection"""
        # Close existing property dialog if open and remember position and tab
        if self.property_dialog:
            self.property_dialog_position = self.property_dialog.GetPosition()
            # Remember the selected tab
            try:
                self.property_dialog_tab_index = self.property_dialog.notebook.GetSelection()
            except:
                pass
            self.property_dialog.Close()
            self.property_dialog = None

        self.selected_element = element
        self.element_label.SetLabel(element)

        # Update line dropdown
        element_lines = ['All Lines'] + sorted(
            self.df[self.df['Element'] == element]['Line'].unique().tolist()
        )
        self.line_combo.Set(element_lines)
        self.line_combo.SetSelection(0)

        self.update_results()

        # Auto-open properties dialog for new element at same position with same tab
        if self.property_dialog_position:  # Only if a dialog was previously open
            self.show_element_properties(None)

    def on_element_double_click(self, element):
        """Handle double-click on element"""
        if element in self.elements:
            self.select_element(element)
            # Always show properties on double-click, even if no previous dialog
            self.show_element_properties(None)

    def on_line_selected(self, event):
        """Handle line selection"""
        self.update_results()

    def on_search_change(self, event):
        """Handle search text change"""
        self.update_results()

    def get_filtered_data(self):
        """Get filtered dataframe based on current selections"""
        filtered_df = self.df.copy()

        # Filter by element
        if self.selected_element:
            filtered_df = filtered_df[filtered_df['Element'] == self.selected_element]

        # Filter by line
        selected_line = self.line_combo.GetStringSelection()
        if selected_line != 'All Lines':
            filtered_df = filtered_df[filtered_df['Line'] == selected_line]

        # Filter by formula
        formula_search = self.formula_search.GetValue().strip().lower()
        if formula_search:
            formula_search = re.escape(formula_search)
            filtered_df = filtered_df[
                filtered_df['Formula'].str.lower().str.contains(formula_search, na=False)
            ]

        # Filter by name
        name_search = self.name_search.GetValue().strip().lower()
        if name_search:
            name_search = re.escape(name_search)
            filtered_df = filtered_df[
                filtered_df['Name'].str.lower().str.contains(name_search, na=False)
            ]

        return filtered_df

    def update_results(self):
        """Update the results grid"""
        # Get filtered data
        filtered_df = self.get_filtered_data()

        # Apply sorting if needed
        if self.sort_column is not None:
            col_name = self.results_grid.GetColLabelValue(self.sort_column)
            # Remove sort indicators from column name
            col_name = col_name.replace(" ▲", "").replace(" ▼", "")

            # Map display name to actual dataframe column name
            if col_name == "BE (eV)":
                col_name = "BE (eV)"  # Already correct
            elif col_name == "Element":
                col_name = "Element"
            elif col_name == "Line":
                col_name = "Line"
            elif col_name == "Formula":
                col_name = "Formula"
            elif col_name == "Name":
                col_name = "Name"
            elif col_name == "Journal":
                col_name = "Journal"

            filtered_df = filtered_df.sort_values(
                by=col_name,
                ascending=self.sort_ascending
            )
        else:
            filtered_df = filtered_df.sort_values(by='BE (eV)')

        # Clear existing rows
        if self.results_grid.GetNumberRows() > 0:
            self.results_grid.DeleteRows(0, self.results_grid.GetNumberRows())

        # Add new rows
        for _, row in filtered_df.iterrows():
            self.results_grid.AppendRows(1)
            row_num = self.results_grid.GetNumberRows() - 1

            self.results_grid.SetCellValue(row_num, 0, str(row['Element']))
            self.results_grid.SetCellValue(row_num, 1, str(row['Line']))
            self.results_grid.SetCellValue(row_num, 2,
                                           f"{row['BE (eV)']:.2f}" if pd.notnull(row['BE (eV)']) else "")
            self.results_grid.SetCellValue(row_num, 3,
                                           str(row['Formula']) if pd.notnull(row['Formula']) else "")
            self.results_grid.SetCellValue(row_num, 4,
                                           str(row['Name']) if pd.notnull(row['Name']) else "")
            self.results_grid.SetCellValue(row_num, 5,
                                           str(row['Journal']) if pd.notnull(row['Journal']) else "")

        # Update status
        self.status_text.SetLabel(f"{len(filtered_df)} results found")

    def on_column_click(self, event):
        """Handle column header click for sorting"""
        col = event.GetCol()
        if col == -1:  # Row label clicked
            return

        # Toggle sort direction if same column
        if self.sort_column == col:
            self.sort_ascending = not self.sort_ascending
        else:
            self.sort_column = col
            self.sort_ascending = True

        # Update column labels to show sort direction
        for i in range(self.results_grid.GetNumberCols()):
            label = self.results_grid.GetColLabelValue(i)
            label = label.replace(" ▲", "").replace(" ▼", "")
            if i == col:
                label += " ▲" if self.sort_ascending else " ▼"
            self.results_grid.SetColLabelValue(i, label)

        self.update_results()

    def on_grid_double_click(self, event):
        """Handle double-click on grid cell"""
        self.show_full_info()

    def on_grid_right_click(self, event):
        """Handle right-click on grid"""
        row = event.GetRow()
        if row < 0:
            return

        # Select the row
        self.results_grid.SelectRow(row)

        # Create context menu
        menu = wx.Menu()

        copy_ref = menu.Append(wx.ID_ANY, "Copy Full Reference")
        copy_journal = menu.Append(wx.ID_ANY, "Copy Journal Only")
        search_scholar = menu.Append(wx.ID_ANY, "Search in Google Scholar")
        menu.AppendSeparator()
        show_info = menu.Append(wx.ID_ANY, "Show Full Information")

        # Bind menu events
        self.Bind(wx.EVT_MENU, self.copy_reference, copy_ref)
        self.Bind(wx.EVT_MENU, self.copy_journal_only, copy_journal)
        self.Bind(wx.EVT_MENU, self.search_google_scholar, search_scholar)
        self.Bind(wx.EVT_MENU, lambda e: self.show_full_info(), show_info)

        # Show menu
        self.PopupMenu(menu)
        menu.Destroy()

    def copy_reference(self, event):
        """Copy full reference to clipboard"""
        row = self.results_grid.GetSelectedRows()[0] if self.results_grid.GetSelectedRows() else -1
        if row < 0:
            return

        # Build reference string
        element = self.results_grid.GetCellValue(row, 0)
        line = self.results_grid.GetCellValue(row, 1)
        be = self.results_grid.GetCellValue(row, 2)
        formula = self.results_grid.GetCellValue(row, 3)
        name = self.results_grid.GetCellValue(row, 4)
        journal = self.results_grid.GetCellValue(row, 5)

        reference = f"{element} {line} - {be} eV - {formula} - {name} - {journal}"

        try:
            pyperclip.copy(reference)
            self.status_text.SetLabel("Reference copied to clipboard")
        except:
            self.status_text.SetLabel("Failed to copy reference")

    def copy_journal_only(self, event):
        """Copy journal to clipboard"""
        row = self.results_grid.GetSelectedRows()[0] if self.results_grid.GetSelectedRows() else -1
        if row < 0:
            return

        journal = self.results_grid.GetCellValue(row, 5)

        try:
            pyperclip.copy(journal)
            self.status_text.SetLabel("Journal copied to clipboard")
        except:
            self.status_text.SetLabel("Failed to copy journal")

    def search_google_scholar(self, event):
        """Search journal in Google Scholar"""
        row = self.results_grid.GetSelectedRows()[0] if self.results_grid.GetSelectedRows() else -1
        if row < 0:
            return

        journal = self.results_grid.GetCellValue(row, 5)

        if journal.strip():
            import urllib.parse
            query = urllib.parse.quote(journal)
            scholar_url = f"https://scholar.google.com/scholar?q={query}"
            webbrowser.open(scholar_url)
            self.status_text.SetLabel("Opened in Google Scholar")
        else:
            self.status_text.SetLabel("No journal information to search")

    def show_full_info(self):
        """Show full information dialog"""
        row = self.results_grid.GetSelectedRows()[0] if self.results_grid.GetSelectedRows() else -1
        if row < 0:
            return

        # Get row data
        element = self.results_grid.GetCellValue(row, 0)
        line = self.results_grid.GetCellValue(row, 1)
        be_str = self.results_grid.GetCellValue(row, 2)
        formula = self.results_grid.GetCellValue(row, 3)

        # Find matching row in dataframe
        be = float(be_str) if be_str else None

        if be is not None:
            matches = self.df[(self.df['Element'] == element) &
                              (self.df['Line'] == line) &
                              (abs(self.df['BE (eV)'] - be) < 0.01) &
                              (self.df['Formula'] == formula)]
        else:
            matches = self.df[(self.df['Element'] == element) &
                              (self.df['Line'] == line) &
                              (self.df['Formula'] == formula)]

        if matches.empty:
            self.status_text.SetLabel("No detailed information found")
            return

        row_data = matches.iloc[0]

        # Create dialog
        dlg = wx.Dialog(self, title=f"Full Information: {element} {line} - {be_str} eV",
                        size=(400, 500))

        panel = wx.Panel(dlg)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Create grid for details
        detail_grid = wx.grid.Grid(panel)
        detail_grid.CreateGrid(0, 2)
        detail_grid.SetColLabelValue(0, "Property")
        detail_grid.SetColLabelValue(1, "Value")
        detail_grid.SetColSize(0, 150)
        detail_grid.SetColSize(1, 200)
        detail_grid.EnableEditing(False)

        # Add data
        for col in self.df.columns:
            value = row_data[col]
            if pd.notnull(value):
                detail_grid.AppendRows(1)
                row_num = detail_grid.GetNumberRows() - 1
                detail_grid.SetCellValue(row_num, 0, col)
                detail_grid.SetCellValue(row_num, 1, str(value))

        # Close button
        close_btn = wx.Button(panel, wx.ID_CLOSE, "Close")
        close_btn.Bind(wx.EVT_BUTTON, lambda e: dlg.Close())

        sizer.Add(detail_grid, 1, wx.ALL | wx.EXPAND, 10)
        sizer.Add(close_btn, 0, wx.ALL | wx.CENTER, 10)

        panel.SetSizer(sizer)
        dlg.ShowModal()
        dlg.Destroy()

    def plot_results(self, event):
        """Create matplotlib plot of binding energies"""
        filtered_df = self.get_filtered_data()

        if filtered_df.empty:
            wx.MessageBox("There is no data to plot.", "No Data",
                          wx.OK | wx.ICON_INFORMATION)
            return

        binding_energies = filtered_df['BE (eV)'].dropna().values

        if len(binding_energies) == 0:
            wx.MessageBox("No binding energy values to plot.", "No Data",
                          wx.OK | wx.ICON_INFORMATION)
            return

        # Create plot window
        plot_frame = PlotFrame(self, binding_energies, self.selected_element,
                               self.line_combo.GetStringSelection())
        plot_frame.Show()

    def show_element_properties(self, event):
        """Show element properties dialog"""
        if not self.selected_element:
            wx.MessageBox("Please select an element from the periodic table first.",
                          "No Element Selected", wx.OK | wx.ICON_INFORMATION)
            return

        # Close existing property dialog if open and remember position and tab
        if self.property_dialog:
            self.property_dialog_position = self.property_dialog.GetPosition()
            # Remember the selected tab
            try:
                self.property_dialog_tab_index = self.property_dialog.notebook.GetSelection()
            except:
                pass
            self.property_dialog.Close()
            self.property_dialog = None

        # Create new properties dialog
        self.property_dialog = ElementPropertiesDialog(self, self.selected_element, self.df)

        # Set position if we have a saved one
        if self.property_dialog_position:
            self.property_dialog.SetPosition(self.property_dialog_position)
        else:
            self.property_dialog.CenterOnParent()

        # Set the remembered tab selection
        try:
            wx.CallAfter(self.property_dialog.notebook.SetSelection, self.property_dialog_tab_index)
        except:
            pass

        # Show as non-modal dialog
        self.property_dialog.Show()


class PlotFrame(wx.Frame):
    """Frame for displaying binding energy plots"""

    def __init__(self, parent, binding_energies, element, line):
        super().__init__(parent, title="Binding Energy Distribution", size=(800, 700))

        self.binding_energies = binding_energies
        self.element = element
        self.line = line

        # Create panel
        panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Control panel
        control_panel = wx.Panel(panel)
        control_panel.SetBackgroundColour(wx.Colour(224, 224, 224))
        control_sizer = wx.BoxSizer(wx.HORIZONTAL)

        control_sizer.Add(wx.StaticText(control_panel, label="Histogram Resolution:"),
                          0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.resolution_combo = wx.ComboBox(control_panel,
                                            choices=["0.1", "0.2", "0.3", "0.4", "0.5",
                                                     "0.6", "0.7", "0.8", "0.9", "1.0"],
                                            value="0.1", style=wx.CB_READONLY)
        self.resolution_combo.Bind(wx.EVT_COMBOBOX, self.on_resolution_change)
        control_sizer.Add(self.resolution_combo, 0, wx.ALL, 5)

        control_sizer.Add(wx.StaticText(control_panel, label="eV"),
                          0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        control_panel.SetSizer(control_sizer)

        # Create matplotlib figure
        self.figure = Figure()
        self.canvas = FigureCanvas(panel, -1, self.figure)
        self.toolbar = NavigationToolbar(self.canvas)

        # Layout
        sizer.Add(control_panel, 0, wx.ALL | wx.EXPAND, 5)
        sizer.Add(self.canvas, 1, wx.ALL | wx.EXPAND, 5)
        sizer.Add(self.toolbar, 0, wx.ALL | wx.EXPAND, 5)

        panel.SetSizer(sizer)

        # Initial plot
        self.update_plot()

        # Center on parent
        self.CenterOnParent()

    def on_resolution_change(self, event):
        """Handle resolution change"""
        self.update_plot()

    def update_plot(self):
        """Update the plot with current resolution"""
        self.figure.clear()
        ax = self.figure.add_subplot(111)

        bin_width = float(self.resolution_combo.GetValue())

        # Calculate bins
        min_energy = np.floor(np.min(self.binding_energies) * 10) / 10
        max_energy = np.ceil(np.max(self.binding_energies) * 10) / 10
        bins = np.arange(min_energy, max_energy + bin_width, bin_width)

        # Create histogram
        counts, bins, patches = ax.hist(self.binding_energies, bins=bins,
                                        alpha=0.7, color='skyblue', edgecolor='black')

        # Add smooth curve
        try:
            from scipy.stats import gaussian_kde
            kde = gaussian_kde(self.binding_energies, bw_method='silverman')
            x = np.linspace(min_energy, max_energy, 1000)
            y = kde(x) * len(self.binding_energies) * bin_width
            ax.plot(x, y, 'r-', linewidth=2)
        except ImportError:
            pass

        # Labels and title
        element_str = f" for {self.element}" if self.element else ""
        line_str = f" ({self.line})" if self.line != "All Lines" else ""

        ax.set_xlabel('Binding Energy (eV)')
        ax.set_ylabel('Number of References')
        ax.set_title(f'Binding Energy Distribution{element_str}{line_str}')
        ax.grid(True, linestyle='--', alpha=0.7)

        # Add text annotations
        ax.text(0.98, 0.95, f'Total References: {len(self.binding_energies)}',
                transform=ax.transAxes, ha='right', va='top',
                bbox=dict(boxstyle='round', facecolor='white', alpha=0.8))

        ax.text(0.98, 0.90, f'Resolution: {bin_width} eV',
                transform=ax.transAxes, ha='right', va='top',
                bbox=dict(boxstyle='round', facecolor='white', alpha=0.8))

        self.canvas.draw()


class ElementPropertiesDialog(wx.Dialog):
    """Dialog for showing element properties"""

    def __init__(self, parent, element, df):
        super().__init__(parent, title=f"Properties for {element}",
                         size=(800, 800), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)

        self.element = element
        self.df = df

        # Get element properties
        self.properties = self.get_element_properties(self.element)

        # Create UI
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Header
        header_panel = wx.Panel(panel)
        header_panel.SetBackgroundColour(wx.Colour(240, 240, 240))
        header_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Element symbol
        symbol_text = wx.StaticText(header_panel, label=element)
        symbol_font = wx.Font(40, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        symbol_text.SetFont(symbol_font)
        header_sizer.Add(symbol_text, 0, wx.ALL, 20)

        # Element name and info
        info_sizer = wx.BoxSizer(wx.VERTICAL)

        name_text = wx.StaticText(header_panel, label=self.properties.get("Name", element))
        name_font = wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        name_text.SetFont(name_font)
        info_sizer.Add(name_text, 0, wx.ALL, 5)

        atomic_text = wx.StaticText(header_panel,
                                    label=f"Atomic Number: {self.properties.get('Atomic Number', 'N/A')}")
        info_sizer.Add(atomic_text, 0, wx.ALL, 5)

        header_sizer.Add(info_sizer, 1, wx.ALL | wx.EXPAND, 20)
        header_panel.SetSizer(header_sizer)

        # Properties notebook
        notebook = wx.Notebook(panel)

        # Properties notebook
        self.notebook = wx.Notebook(panel)  # Make sure this is self.notebook

        # Create tabs
        self.create_properties_tab(self.notebook)
        self.create_xps_tab(self.notebook)
        self.create_xps_fitting_tab(self.notebook)
        self.create_thermo_tab(self.notebook)

        # Bind notebook page change event to track selections
        self.notebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.on_tab_changed)


        # Layout - links come after header, before notebook
        main_sizer.Add(header_panel, 0, wx.ALL | wx.EXPAND, 0)
        main_sizer.Add(self.notebook, 1, wx.ALL | wx.EXPAND, 0)

        panel.SetSizer(main_sizer)

        # Center on parent
        self.CenterOnParent()

    def on_tab_changed(self, event):
        """Handle tab change to update parent's memory"""
        try:
            # Update parent's tab index memory
            if hasattr(self.GetParent(), 'property_dialog_tab_index'):
                self.GetParent().property_dialog_tab_index = event.GetSelection()
        except:
            pass
        event.Skip()  # Allow the event to continue processing

    def create_properties_tab(self, notebook):
        """Create general properties tab"""
        panel = wx.Panel(notebook)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Create scrolled window
        scrolled = wx.ScrolledWindow(panel)
        scrolled.SetScrollRate(20, 20)
        grid_sizer = wx.FlexGridSizer(cols=2, hgap=10, vgap=5)

        # Property groups
        property_groups = {
            "Physical Properties": ["Atomic Mass", "Density", "Melting Point", "Boiling Point", "State at 20°C"],
            "Atomic Properties": ["Electron Configuration", "Electronegativity", "Atomic Radius", "Ionization Energy"],
            "General Information": ["Group", "Period", "Category", "Discovered By", "Year of Discovery"],
            "Online Resources": ["XPS Fitting URL", "Thermo Fisher URL"]
        }

        for group_name, properties in property_groups.items():
            # Group header
            header = wx.StaticText(scrolled, label=group_name)
            header_font = wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
            header.SetFont(header_font)
            grid_sizer.Add(header, 0, wx.ALL | wx.EXPAND, 5)
            grid_sizer.Add(wx.StaticText(scrolled, label=""), 0)  # Empty cell

            # Properties
            for prop in properties:
                label = wx.StaticText(scrolled, label=f"{prop}:")

                # Handle URL properties specially
                if prop.endswith("URL"):
                    url = str(self.properties.get(prop, "N/A"))
                    if url != "N/A":
                        # Create clickable link
                        link_label = "XPS Fitting" if "xpsfitting" in url else "Thermo Fisher"
                        value = wx.adv.HyperlinkCtrl(scrolled, label=link_label, url=url)
                    else:
                        value = wx.StaticText(scrolled, label="N/A")
                else:
                    value = wx.StaticText(scrolled, label=str(self.properties.get(prop, "N/A")))

                grid_sizer.Add(label, 0, wx.ALL | wx.ALIGN_RIGHT, 3)
                grid_sizer.Add(value, 0, wx.ALL, 3)

        scrolled.SetSizer(grid_sizer)
        sizer.Add(scrolled, 1, wx.ALL | wx.EXPAND, 5)
        panel.SetSizer(sizer)

        notebook.AddPage(panel, "General Properties")

    def create_xps_tab(self, notebook):
        """Create XPS data tab"""
        panel = wx.Panel(notebook)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Get XPS data for element
        element_data = self.df[self.df['Element'] == self.element]

        if not element_data.empty:
            # Create grid
            grid = wx.grid.Grid(panel)

            # Hide row labels (row numbers) - ADD THIS LINE
            grid.HideRowLabels()

            # Group by line
            lines_summary = element_data.groupby('Line')['BE (eV)'].agg(
                ['mean', 'count', 'min', 'max']).reset_index()

            # Create grid
            grid.CreateGrid(len(lines_summary), 5)
            grid.SetColLabelValue(0, "Line")
            grid.SetColLabelValue(1, "Avg BE (eV)")
            grid.SetColLabelValue(2, "Min BE (eV)")
            grid.SetColLabelValue(3, "Max BE (eV)")
            grid.SetColLabelValue(4, "N# of References")

            # Set column widths
            grid.SetColSize(0, 80)
            grid.SetColSize(1, 100)
            grid.SetColSize(2, 100)
            grid.SetColSize(3, 100)
            grid.SetColSize(4, 80)

            # Populate grid
            for i, (_, row) in enumerate(lines_summary.iterrows()):
                grid.SetCellValue(i, 0, str(row['Line']))
                grid.SetCellValue(i, 1, f"{row['mean']:.2f}")
                grid.SetCellValue(i, 2, f"{row['min']:.2f}")
                grid.SetCellValue(i, 3, f"{row['max']:.2f}")
                grid.SetCellValue(i, 4, str(int(row['count'])))

            grid.EnableEditing(False)
            sizer.Add(grid, 1, wx.ALL | wx.EXPAND, 5)
        else:
            label = wx.StaticText(panel, label="No XPS data available for this element")
            sizer.Add(label, 0, wx.ALL | wx.CENTER, 20)

        panel.SetSizer(sizer)
        notebook.AddPage(panel, "XPS Data")

    def create_thermo_tab(self, notebook):
        """Create Thermo Knowledge tab with embedded web browser"""
        panel = wx.Panel(notebook)
        sizer = wx.BoxSizer(wx.VERTICAL)

        try:
            # Create toolbar with refresh and zoom buttons
            toolbar_panel = wx.Panel(panel)
            toolbar_sizer = wx.BoxSizer(wx.HORIZONTAL)

            refresh_btn = wx.Button(toolbar_panel, label="Refresh")
            refresh_btn.Bind(wx.EVT_BUTTON, self.on_refresh_thermo)

            zoom_in_btn = wx.Button(toolbar_panel, label="+", size=(30, -1))
            zoom_in_btn.Bind(wx.EVT_BUTTON, self.on_thermo_zoom_in)

            zoom_out_btn = wx.Button(toolbar_panel, label="-", size=(30, -1))
            zoom_out_btn.Bind(wx.EVT_BUTTON, self.on_thermo_zoom_out)

            url_label = wx.StaticText(toolbar_panel, label="Thermo Fisher Knowledge Base")
            font = url_label.GetFont()
            font.SetWeight(wx.FONTWEIGHT_BOLD)
            url_label.SetFont(font)

            toolbar_sizer.Add(url_label, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
            toolbar_sizer.Add(zoom_out_btn, 0, wx.ALL, 2)
            toolbar_sizer.Add(zoom_in_btn, 0, wx.ALL, 2)
            toolbar_sizer.Add(refresh_btn, 0, wx.ALL, 5)
            toolbar_panel.SetSizer(toolbar_sizer)

            # Create web view control
            self.web_view = wx.html2.WebView.New(panel)

            # Get the Thermo Fisher URL for this element
            self.thermo_url = self.get_thermo_url(self.element)

            # Load the webpage
            self.web_view.LoadURL(self.thermo_url)

            # Add loading indicator
            loading_text = wx.StaticText(panel, label="Loading Thermo Fisher knowledge page...")
            loading_text.SetForegroundColour(wx.Colour(100, 100, 100))

            # Bind events to handle loading
            self.web_view.Bind(wx.html2.EVT_WEBVIEW_LOADED, self.on_page_loaded)
            self.web_view.Bind(wx.html2.EVT_WEBVIEW_ERROR, self.on_page_error)

            sizer.Add(toolbar_panel, 0, wx.ALL | wx.EXPAND, 2)
            sizer.Add(loading_text, 0, wx.ALL | wx.EXPAND, 5)
            sizer.Add(self.web_view, 1, wx.ALL | wx.EXPAND, 5)

            # Store reference to loading text for removal later
            self.loading_text = loading_text

        except Exception as e:
            # Fallback if WebView is not available
            error_text = wx.StaticText(panel,
                                       label=f"Web browser not available.\n\nPlease visit:\n{self.get_thermo_url(self.element)}")
            error_text.Wrap(400)
            sizer.Add(error_text, 1, wx.ALL | wx.EXPAND, 20)

        panel.SetSizer(sizer)
        notebook.AddPage(panel, "Thermo Knowledge")

    def on_refresh_thermo(self, event):
        """Refresh the Thermo Fisher page"""
        try:
            self.loading_text.SetLabel("Refreshing page...")
            self.loading_text.Show()
            self.web_view.Reload()
            self.Layout()
        except:
            pass

    def create_xps_fitting_tab(self, notebook):
        """Create XPS Fitting tab with embedded web browser"""
        panel = wx.Panel(notebook)
        sizer = wx.BoxSizer(wx.VERTICAL)

        try:
            # Create toolbar with refresh and zoom buttons
            toolbar_panel = wx.Panel(panel)
            toolbar_sizer = wx.BoxSizer(wx.HORIZONTAL)

            refresh_btn = wx.Button(toolbar_panel, label="Refresh")
            refresh_btn.Bind(wx.EVT_BUTTON, self.on_refresh_xps)

            zoom_in_btn = wx.Button(toolbar_panel, label="+", size=(30, -1))
            zoom_in_btn.Bind(wx.EVT_BUTTON, self.on_xps_zoom_in)

            zoom_out_btn = wx.Button(toolbar_panel, label="-", size=(30, -1))
            zoom_out_btn.Bind(wx.EVT_BUTTON, self.on_xps_zoom_out)

            url_label = wx.StaticText(toolbar_panel, label="XPS Fitting Database")
            font = url_label.GetFont()
            font.SetWeight(wx.FONTWEIGHT_BOLD)
            url_label.SetFont(font)

            toolbar_sizer.Add(url_label, 1, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
            toolbar_sizer.Add(zoom_out_btn, 0, wx.ALL, 2)
            toolbar_sizer.Add(zoom_in_btn, 0, wx.ALL, 2)
            toolbar_sizer.Add(refresh_btn, 0, wx.ALL, 5)
            toolbar_panel.SetSizer(toolbar_sizer)
            # Create web view control
            self.xps_web_view = wx.html2.WebView.New(panel)

            # Get the XPS Fitting URL for this element
            self.xps_url = self.get_xps_fitting_url(self.element)

            # Load the webpage
            self.xps_web_view.LoadURL(self.xps_url)

            # Add loading indicator
            xps_loading_text = wx.StaticText(panel, label="Loading XPS Fitting database page...")
            xps_loading_text.SetForegroundColour(wx.Colour(100, 100, 100))

            # Bind events to handle loading
            self.xps_web_view.Bind(wx.html2.EVT_WEBVIEW_LOADED, self.on_xps_page_loaded)
            self.xps_web_view.Bind(wx.html2.EVT_WEBVIEW_ERROR, self.on_xps_page_error)

            sizer.Add(toolbar_panel, 0, wx.ALL | wx.EXPAND, 2)
            sizer.Add(xps_loading_text, 0, wx.ALL | wx.EXPAND, 5)
            sizer.Add(self.xps_web_view, 1, wx.ALL | wx.EXPAND, 5)

            # Store reference to loading text for removal later
            self.xps_loading_text = xps_loading_text

        except Exception as e:
            # Fallback if WebView is not available
            error_text = wx.StaticText(panel,
                                       label=f"Web browser not available.\n\nPlease visit:\n{self.get_xps_fitting_url(self.element)}")
            error_text.Wrap(400)
            sizer.Add(error_text, 1, wx.ALL | wx.EXPAND, 20)

        panel.SetSizer(sizer)
        notebook.AddPage(panel, "XPS Fitting")

    def on_refresh_xps(self, event):
        """Refresh the XPS Fitting page"""
        try:
            self.xps_loading_text.SetLabel("Refreshing page...")
            self.xps_loading_text.Show()
            self.xps_web_view.Reload()
            self.Layout()
        except:
            pass

    def on_xps_page_loaded(self, event):
        """Handle successful XPS page loading"""
        try:
            # Hide loading text when page is loaded
            self.xps_loading_text.Hide()
            self.Layout()
        except:
            pass

    def on_xps_page_error(self, event):
        """Handle XPS page loading errors"""
        try:
            # Show error message
            self.xps_loading_text.SetLabel("Failed to load webpage. Please check your internet connection.")
            self.xps_loading_text.SetForegroundColour(wx.Colour(200, 0, 0))
        except:
            pass

    def on_xps_zoom_in(self, event):
        """Zoom in the XPS web view using JavaScript"""
        try:
            script = """
            document.body.style.zoom = (parseFloat(document.body.style.zoom || 1) * 1.1).toString();
            """
            self.xps_web_view.RunScript(script)
        except:
            pass

    def on_xps_zoom_out(self, event):
        """Zoom out the XPS web view using JavaScript"""
        try:
            script = """
            var currentZoom = parseFloat(document.body.style.zoom || 1);
            if (currentZoom > 0.5) {
                document.body.style.zoom = (currentZoom / 1.1).toString();
            }
            """
            self.xps_web_view.RunScript(script)
        except:
            pass

    def on_thermo_zoom_in(self, event):
        """Zoom in the Thermo web view using JavaScript"""
        try:
            script = """
            document.body.style.zoom = (parseFloat(document.body.style.zoom || 1) * 1.1).toString();
            """
            self.web_view.RunScript(script)
        except:
            pass

    def on_thermo_zoom_out(self, event):
        """Zoom out the Thermo web view using JavaScript"""
        try:
            script = """
            var currentZoom = parseFloat(document.body.style.zoom || 1);
            if (currentZoom > 0.5) {
                document.body.style.zoom = (currentZoom / 1.1).toString();
            }
            """
            self.web_view.RunScript(script)
        except:
            pass

    def on_page_loaded(self, event):
        """Handle successful page loading"""
        try:
            # Hide loading text when page is loaded
            self.loading_text.Hide()
            self.Layout()
        except:
            pass

    def on_page_error(self, event):
        """Handle page loading errors"""
        try:
            # Show error message
            self.loading_text.SetLabel("Failed to load webpage. Please check your internet connection.")
            self.loading_text.SetForegroundColour(wx.Colour(200, 0, 0))
        except:
            pass

    def get_xps_fitting_url(self, element_symbol):
        """Generate XPS Fitting URL for element"""
        # Map element symbols to exact names used on xpsfitting.com
        element_names = {
            'H': 'Hydrogen', 'He': 'Helium', 'Li': 'Lithium', 'Be': 'Beryllium', 'B': 'Boron',
            'C': 'carbon', 'N': 'Nitrogen', 'O': 'Oxygen', 'F': 'Fluorine', 'Ne': 'Neon',
            'Na': 'Sodium', 'Mg': 'Magnesium', 'Al': 'Aluminium', 'Si': 'Silicon', 'P': 'Phosphorus',
            'S': 'Sulfur', 'Cl': 'Chlorine', 'Ar': 'Argon', 'K': 'Potassium', 'Ca': 'Calcium',
            'Sc': 'Scandium', 'Ti': 'Titanium', 'V': 'Vanadium', 'Cr': 'Chromium', 'Mn': 'Manganese',
            'Fe': 'Iron', 'Co': 'Cobalt', 'Ni': 'Nickel', 'Cu': 'Copper', 'Zn': 'Zinc',
            'Ga': 'Gallium', 'Ge': 'Germanium', 'As': 'Arsenic', 'Se': 'Selenium', 'Br': 'Bromine',
            'Kr': 'Krypton', 'Rb': 'Rubidium', 'Sr': 'Strontium', 'Y': 'Yttrium', 'Zr': 'Zirconium',
            'Nb': 'Niobium', 'Mo': 'Molybdenum', 'Tc': 'Technetium', 'Ru': 'Ruthenium', 'Rh': 'Rhodium',
            'Pd': 'Palladium', 'Ag': 'Silver', 'Cd': 'Cadmium', 'In': 'Indium', 'Sn': 'Tin',
            'Sb': 'Antimony', 'Te': 'Tellurium', 'I': 'Iodine', 'Xe': 'Xenon', 'Cs': 'Caesium',
            'Ba': 'Barium', 'La': 'Lanthanum', 'Ce': 'Cerium', 'Pr': 'Praseodymium', 'Nd': 'Neodymium',
            'Pm': 'Promethium', 'Sm': 'Samarium', 'Eu': 'Europium', 'Gd': 'Gadolinium', 'Tb': 'Terbium',
            'Dy': 'Dysprosium', 'Ho': 'Holmium', 'Er': 'Erbium', 'Tm': 'Thulium', 'Yb': 'Ytterbium',
            'Lu': 'Lutetium', 'Hf': 'Hafnium', 'Ta': 'Tantalum', 'W': 'Tungsten', 'Re': 'Rhenium',
            'Os': 'Osmium', 'Ir': 'Iridium', 'Pt': 'Platinum', 'Au': 'Gold', 'Hg': 'Mercury',
            'Tl': 'Thallium', 'Pb': 'Lead', 'Bi': 'Bismuth', 'Po': 'Polonium', 'At': 'Astatine',
            'Rn': 'Radon', 'Fr': 'Francium', 'Ra': 'Radium', 'Ac': 'Actinium', 'Th': 'Thorium',
            'Pa': 'Protactinium', 'U': 'Uranium', 'Np': 'Neptunium', 'Pu': 'Plutonium', 'Am': 'Americium',
            'Cm': 'Curium', 'Bk': 'Berkelium', 'Cf': 'Californium', 'Es': 'Einsteinium', 'Fm': 'Fermium',
            'Md': 'Mendelevium', 'No': 'Nobelium', 'Lr': 'Lawrencium'
        }

        element_name = element_names.get(element_symbol, element_symbol)
        return f"https://www.xpsfitting.com/search/label/{element_name}"

    def get_thermo_url(self, element_symbol):
        """Generate Thermo Fisher URL for element"""
        # Map element symbols to their category paths for Thermo URLs
        element_categories = {
            'H': 'nonmetal/hydrogen', 'He': 'noble-gas/helium', 'Li': 'alkali-metal/lithium',
            'Be': 'alkaline-earth-metal/beryllium', 'B': 'metalloid/boron', 'C': 'non-metal/carbon',
            'N': 'non-metal/nitrogen', 'O': 'non-metal/oxygen', 'F': 'halogen/fluorine',
            'Ne': 'noble-gas/neon', 'Na': 'alkali-metal/sodium', 'Mg': 'alkaline-earth-metal/magnesium',
            'Al': 'post-transition-metal/aluminum', 'Si': 'metalloid/silicon', 'P': 'non-metal/phosphorus',
            'S': 'nonmetal/sulfur', 'Cl': 'halogen/chlorine', 'Ar': 'noble-gas/argon',
            'K': 'alkali-metal/potassium', 'Ca': 'alkaline-earth-metal/calcium',
            'Sc': 'transition-metal/scandium', 'Ti': 'transition-metal/titanium',
            'V': 'transition-metal/vanadium', 'Cr': 'transition-metal/chromium',
            'Mn': 'transition-metal/manganese', 'Fe': 'transition-metal/iron',
            'Co': 'transition-metal/cobalt', 'Ni': 'transition-metal/nickel',
            'Cu': 'transition-metal/copper', 'Zn': 'transition-metal/zinc',
            'Ga': 'post-transition-metal/gallium', 'Ge': 'metalloid/germanium',
            'As': 'metalloid/arsenic', 'Se': 'nonmetal/selenium', 'Br': 'halogen/bromine',
            'Kr': 'noble-gas/krypton', 'Rb': 'alkali-metal/rubidium',
            'Sr': 'alkaline-earth-metal/strontium', 'Y': 'transition-metal/yttrium',
            'Zr': 'transition-metal/zirconium', 'Nb': 'transition-metal/niobium',
            'Mo': 'transition-metal/molybdenum', 'Tc': 'transition-metal/technetium',
            'Ru': 'transition-metal/ruthenium', 'Rh': 'transition-metal/rhodium',
            'Pd': 'transition-metal/palladium', 'Ag': 'transition-metal/silver',
            'Cd': 'transition-metal/cadmium', 'In': 'post-transition-metal/indium',
            'Sn': 'post-transition-metal/tin', 'Sb': 'metalloid/antimony',
            'Te': 'metalloid/tellurium', 'I': 'halogen/iodine', 'Xe': 'noble-gas/xenon',
            'Cs': 'alkali-metal/cesium', 'Ba': 'alkaline-earth-metal/barium',
            'La': 'lanthanide-rare-earth/lanthanum', 'Ce': 'lanthanide-rare-earth/cerium', 'Pr': 'lanthanide-rare-earth/praseodymium',
            'Nd': 'lanthanide-rare-earth/neodymium', 'Pm': 'lanthanide-rare-earth/promethium', 'Sm': 'lanthanide-rare-earth/samarium',
            'Eu': 'lanthanide-rare-earth/europium', 'Gd': 'lanthanide-rare-earth/gadolinium', 'Tb': 'lanthanide-rare-earth/terbium',
            'Dy': 'lanthanide-rare-earth/dysprosium', 'Ho': 'lanthanide-rare-earth/holmium', 'Er': 'lanthanide-rare-earth/erbium',
            'Tm': 'lanthanide-rare-earth/thulium', 'Yb': 'lanthanide-rare-earth/ytterbium', 'Lu': 'lanthanide-rare-earth/lutetium',
            'Hf': 'transition-metal/hafnium', 'Ta': 'transition-metal/tantalum',
            'W': 'transition-metal/tungsten', 'Re': 'transition-metal/rhenium',
            'Os': 'transition-metal/osmium', 'Ir': 'transition-metal/iridium',
            'Pt': 'transition-metal/platinum', 'Au': 'transition-metal/gold',
            'Hg': 'transition-metal/mercury', 'Tl': 'post-transition-metal/thallium',
            'Pb': 'post-transition-metal/lead', 'Bi': 'post-transition-metal/bismuth',
            'Po': 'metalloid/polonium', 'At': 'halogen/astatine', 'Rn': 'noble-gas/radon',
            'Fr': 'alkali-metal/francium', 'Ra': 'alkaline-earth-metal/radium',
            'Ac': 'actinide/actinium', 'Th': 'actinide/thorium', 'Pa': 'actinide/protactinium',
            'U': 'actinide/uranium', 'Np': 'actinide/neptunium', 'Pu': 'actinide/plutonium',
            'Am': 'actinide/americium', 'Cm': 'actinide/curium', 'Bk': 'actinide/berkelium',
            'Cf': 'actinide/californium', 'Es': 'actinide/einsteinium', 'Fm': 'actinide/fermium',
            'Md': 'actinide/mendelevium', 'No': 'actinide/nobelium', 'Lr': 'actinide/lawrencium'
        }

        category_path = element_categories.get(element_symbol, f'unknown/{element_symbol.lower()}')
        return f"https://www.thermofisher.com/uk/en/home/materials-science/learning-center/periodic-table/{category_path}.html"

    def get_element_properties(self, element_symbol):
        """Get properties for a specific element"""
        # Comprehensive database of element properties based on reference tables
        element_properties = {
            "H": {
                "Name": "Hydrogen",
                "Atomic Number": 1,
                "Atomic Mass": "1.00794 u",
                "Density": "0.0708 g/cm³",
                "Melting Point": "-259.34 °C",
                "Boiling Point": "-252.87 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "1s¹",
                "Ground State": "²S₁/₂",
                "Electronegativity": 2.20,
                "Atomic Radius": "53 pm",
                "Ionization Energy": "13.598 eV",
                "Specific Heat": "14.304 J/(g·K)",
                "Group": 1,
                "Period": 1,
                "Category": "Nonmetal",
                "Common Core Levels": "1s",
                "Most Intense Line": "1s",
                "Typical FWHM": "0.9-1.2 eV",
                "Chemical Shift Range": "0-3 eV"
            },
            "He": {
                "Name": "Helium",
                "Atomic Number": 2,
                "Atomic Mass": "4.002602 u",
                "Density": "0.122 g/cm³",
                "Melting Point": "— °C",
                "Boiling Point": "-268.93 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "1s²",
                "Ground State": "¹S₀",
                "Electronegativity": "N/A",
                "Atomic Radius": "31 pm",
                "Ionization Energy": "24.587 eV",
                "Specific Heat": "5.193 J/(g·K)",
                "Group": 18,
                "Period": 1,
                "Category": "Noble Gas"
            },
            "Li": {
                "Name": "Lithium",
                "Atomic Number": 3,
                "Atomic Mass": "6.941 u",
                "Density": "0.534 g/cm³",
                "Melting Point": "180.50 °C",
                "Boiling Point": "1342 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "1s² 2s¹",
                "Ground State": "²S₁/₂",
                "Electronegativity": 0.98,
                "Atomic Radius": "167 pm",
                "Ionization Energy": "5.392 eV",
                "Specific Heat": "3.582 J/(g·K)",
                "Group": 1,
                "Period": 2,
                "Category": "Alkali Metal"
            },
            "Be": {
                "Name": "Beryllium",
                "Atomic Number": 4,
                "Atomic Mass": "9.012182 u",
                "Density": "1.848 g/cm³",
                "Melting Point": "1287 °C",
                "Boiling Point": "2471 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "1s² 2s²",
                "Ground State": "¹S₀",
                "Electronegativity": 1.57,
                "Atomic Radius": "112 pm",
                "Ionization Energy": "9.323 eV",
                "Specific Heat": "1.825 J/(g·K)",
                "Group": 2,
                "Period": 2,
                "Category": "Alkaline Earth Metal"
            },
            "B": {
                "Name": "Boron",
                "Atomic Number": 5,
                "Atomic Mass": "10.811 u",
                "Density": "2.34 g/cm³",
                "Melting Point": "2075 °C",
                "Boiling Point": "4000 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "1s² 2s² 2p¹",
                "Ground State": "²P°₁/₂",
                "Electronegativity": 2.04,
                "Atomic Radius": "87 pm",
                "Ionization Energy": "8.298 eV",
                "Specific Heat": "1.026 J/(g·K)",
                "Group": 13,
                "Period": 2,
                "Category": "Metalloid"
            },
            "C": {
                "Name": "Carbon",
                "Atomic Number": 6,
                "Atomic Mass": "12.0107 u",
                "Density": "1.9-2.3 (graphite) g/cm³",
                "Melting Point": "4492 (10.3 MPa) °C",
                "Boiling Point": "3825ᵇ °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "1s² 2s² 2p²",
                "Ground State": "³P₀",
                "Electronegativity": 2.55,
                "Atomic Radius": "67 pm",
                "Ionization Energy": "11.260 eV",
                "Specific Heat": "0.709 J/(g·K)",
                "Group": 14,
                "Period": 2,
                "Category": "Nonmetal",
                "Common Core Levels": "1s, 2s, 2p",
                "Most Intense Line": "1s",
                "Typical FWHM": "0.8-1.2 eV",
                "Chemical Shift Range": "0-10 eV"
            },
            "N": {
                "Name": "Nitrogen",
                "Atomic Number": 7,
                "Atomic Mass": "14.00674 u",
                "Density": "0.808 g/cm³",
                "Melting Point": "-210.00 °C",
                "Boiling Point": "-195.79 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "1s² 2s² 2p³",
                "Ground State": "⁴S°₃/₂",
                "Electronegativity": 3.04,
                "Atomic Radius": "56 pm",
                "Ionization Energy": "14.534 eV",
                "Specific Heat": "1.040 J/(g·K)",
                "Group": 15,
                "Period": 2,
                "Category": "Nonmetal"
            },
            "O": {
                "Name": "Oxygen",
                "Atomic Number": 8,
                "Atomic Mass": "15.9994 u",
                "Density": "1.14 g/cm³",
                "Melting Point": "-218.79 °C",
                "Boiling Point": "-182.95 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "1s² 2s² 2p⁴",
                "Ground State": "³P₂",
                "Electronegativity": 3.44,
                "Atomic Radius": "48 pm",
                "Ionization Energy": "13.618 eV",
                "Specific Heat": "0.918 J/(g·K)",
                "Group": 16,
                "Period": 2,
                "Category": "Nonmetal",
                "Common Core Levels": "1s, 2s, 2p",
                "Most Intense Line": "1s",
                "Typical FWHM": "0.9-1.3 eV",
                "Chemical Shift Range": "0-8 eV"
            },
            "F": {
                "Name": "Fluorine",
                "Atomic Number": 9,
                "Atomic Mass": "18.9984032 u",
                "Density": "1.50 g/cm³",
                "Melting Point": "-219.62 °C",
                "Boiling Point": "-188.12 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "1s² 2s² 2p⁵",
                "Ground State": "²P°₃/₂",
                "Electronegativity": 3.98,
                "Atomic Radius": "42 pm",
                "Ionization Energy": "17.423 eV",
                "Specific Heat": "0.824 J/(g·K)",
                "Group": 17,
                "Period": 2,
                "Category": "Halogen"
            },
            "Ne": {
                "Name": "Neon",
                "Atomic Number": 10,
                "Atomic Mass": "20.1797 u",
                "Density": "1.207 g/cm³",
                "Melting Point": "-248.59 °C",
                "Boiling Point": "-246.08 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "1s² 2s² 2p⁶",
                "Ground State": "¹S₀",
                "Electronegativity": "N/A",
                "Atomic Radius": "38 pm",
                "Ionization Energy": "21.565 eV",
                "Specific Heat": "1.030 J/(g·K)",
                "Group": 18,
                "Period": 2,
                "Category": "Noble Gas"
            },
            "Na": {
                "Name": "Sodium",
                "Atomic Number": 11,
                "Atomic Mass": "22.989770 u",
                "Density": "0.971 g/cm³",
                "Melting Point": "97.80 °C",
                "Boiling Point": "883 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ne] 3s¹",
                "Ground State": "²S₁/₂",
                "Electronegativity": 0.93,
                "Atomic Radius": "190 pm",
                "Ionization Energy": "5.139 eV",
                "Specific Heat": "1.228 J/(g·K)",
                "Group": 1,
                "Period": 3,
                "Category": "Alkali Metal"
            },
            "Mg": {
                "Name": "Magnesium",
                "Atomic Number": 12,
                "Atomic Mass": "24.3050 u",
                "Density": "1.738 g/cm³",
                "Melting Point": "650 °C",
                "Boiling Point": "1090 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ne] 3s²",
                "Ground State": "¹S₀",
                "Electronegativity": 1.31,
                "Atomic Radius": "145 pm",
                "Ionization Energy": "7.646 eV",
                "Specific Heat": "1.023 J/(g·K)",
                "Group": 2,
                "Period": 3,
                "Category": "Alkaline Earth Metal"
            },
            "Al": {
                "Name": "Aluminum",
                "Atomic Number": 13,
                "Atomic Mass": "26.981538 u",
                "Density": "2.6989 g/cm³",
                "Melting Point": "660.32 °C",
                "Boiling Point": "2519 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ne] 3s² 3p¹",
                "Ground State": "²P°₁/₂",
                "Electronegativity": 1.61,
                "Atomic Radius": "118 pm",
                "Ionization Energy": "5.986 eV",
                "Specific Heat": "0.897 J/(g·K)",
                "Group": 13,
                "Period": 3,
                "Category": "Post-Transition Metal"
            },
            "Si": {
                "Name": "Silicon",
                "Atomic Number": 14,
                "Atomic Mass": "28.0855 u",
                "Density": "2.3325 g/cm³",
                "Melting Point": "1414 °C",
                "Boiling Point": "3265 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ne] 3s² 3p²",
                "Ground State": "³P₀",
                "Electronegativity": 1.90,
                "Atomic Radius": "111 pm",
                "Ionization Energy": "8.152 eV",
                "Specific Heat": "0.705 J/(g·K)",
                "Group": 14,
                "Period": 3,
                "Category": "Metalloid"
            },
            "P": {
                "Name": "Phosphorus",
                "Atomic Number": 15,
                "Atomic Mass": "30.973761 u",
                "Density": "1.82 g/cm³",
                "Melting Point": "44.15 °C",
                "Boiling Point": "280.5 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ne] 3s² 3p³",
                "Ground State": "⁴S°₃/₂",
                "Electronegativity": 2.19,
                "Atomic Radius": "98 pm",
                "Ionization Energy": "10.487 eV",
                "Specific Heat": "0.769 J/(g·K)",
                "Group": 15,
                "Period": 3,
                "Category": "Nonmetal"
            },
            "S": {
                "Name": "Sulfur",
                "Atomic Number": 16,
                "Atomic Mass": "32.066 u",
                "Density": "2.07 g/cm³",
                "Melting Point": "119.6 °C",
                "Boiling Point": "444.60 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ne] 3s² 3p⁴",
                "Ground State": "³P₂",
                "Electronegativity": 2.58,
                "Atomic Radius": "88 pm",
                "Ionization Energy": "10.360 eV",
                "Specific Heat": "0.710 J/(g·K)",
                "Group": 16,
                "Period": 3,
                "Category": "Nonmetal"
            },
            "Cl": {
                "Name": "Chlorine",
                "Atomic Number": 17,
                "Atomic Mass": "35.4527 u",
                "Density": "1.56 (3.3, 6) g/cm³",
                "Melting Point": "-101.5 °C",
                "Boiling Point": "-34.04 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "[Ne] 3s² 3p⁵",
                "Ground State": "²P°₃/₂",
                "Electronegativity": 3.16,
                "Atomic Radius": "79 pm",
                "Ionization Energy": "12.968 eV",
                "Specific Heat": "0.479 J/(g·K)",
                "Group": 17,
                "Period": 3,
                "Category": "Halogen"
            },
            "Ar": {
                "Name": "Argon",
                "Atomic Number": 18,
                "Atomic Mass": "39.948 u",
                "Density": "1.40 g/cm³",
                "Melting Point": "-189.35 °C",
                "Boiling Point": "-185.85 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "[Ne] 3s² 3p⁶",
                "Ground State": "¹S₀",
                "Electronegativity": "N/A",
                "Atomic Radius": "71 pm",
                "Ionization Energy": "15.760 eV",
                "Specific Heat": "0.520 J/(g·K)",
                "Group": 18,
                "Period": 3,
                "Category": "Noble Gas"
            },
            "K": {
                "Name": "Potassium",
                "Atomic Number": 19,
                "Atomic Mass": "39.0983 u",
                "Density": "0.862 g/cm³",
                "Melting Point": "63.5 °C",
                "Boiling Point": "759 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 4s¹",
                "Ground State": "²S₁/₂",
                "Electronegativity": 0.82,
                "Atomic Radius": "243 pm",
                "Ionization Energy": "4.341 eV",
                "Specific Heat": "0.757 J/(g·K)",
                "Group": 1,
                "Period": 4,
                "Category": "Alkali Metal"
            },
            "Ca": {
                "Name": "Calcium",
                "Atomic Number": 20,
                "Atomic Mass": "40.078 u",
                "Density": "1.55 g/cm³",
                "Melting Point": "842 °C",
                "Boiling Point": "1484 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 4s²",
                "Ground State": "¹S₀",
                "Electronegativity": 1.00,
                "Atomic Radius": "194 pm",
                "Ionization Energy": "6.113 eV",
                "Specific Heat": "0.647 J/(g·K)",
                "Group": 2,
                "Period": 4,
                "Category": "Alkaline Earth Metal"
            },
            "Sc": {
                "Name": "Scandium",
                "Atomic Number": 21,
                "Atomic Mass": "44.955912 u",
                "Density": "2.989 g/cm³",
                "Melting Point": "1541 °C",
                "Boiling Point": "2836 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d¹ 4s²",
                "Ground State": "²D₃/₂",
                "Electronegativity": 1.36,
                "Atomic Radius": "184 pm",
                "Ionization Energy": "6.561 eV",
                "Specific Heat": "0.568 J/(g·K)",
                "Group": 3,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Lars Fredrik Nilson",
                "Year of Discovery": 1879
            },
            "Ti": {
                "Name": "Titanium",
                "Atomic Number": 22,
                "Atomic Mass": "47.867 u",
                "Density": "4.54 g/cm³",
                "Melting Point": "1668 °C",
                "Boiling Point": "3287 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d² 4s²",
                "Ground State": "³F₂",
                "Electronegativity": 1.54,
                "Atomic Radius": "176 pm",
                "Ionization Energy": "6.828 eV",
                "Specific Heat": "0.523 J/(g·K)",
                "Group": 4,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "William Gregor",
                "Year of Discovery": 1791
            },
            "V": {
                "Name": "Vanadium",
                "Atomic Number": 23,
                "Atomic Mass": "50.9415 u",
                "Density": "6.11 g/cm³",
                "Melting Point": "1910 °C",
                "Boiling Point": "3407 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d³ 4s²",
                "Ground State": "⁴F₃/₂",
                "Electronegativity": 1.63,
                "Atomic Radius": "171 pm",
                "Ionization Energy": "6.746 eV",
                "Specific Heat": "0.489 J/(g·K)",
                "Group": 5,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Andrés Manuel del Río",
                "Year of Discovery": 1801
            },
            "Cr": {
                "Name": "Chromium",
                "Atomic Number": 24,
                "Atomic Mass": "51.9961 u",
                "Density": "7.15 g/cm³",
                "Melting Point": "1907 °C",
                "Boiling Point": "2671 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d⁵ 4s¹",
                "Ground State": "⁷S₃",
                "Electronegativity": 1.66,
                "Atomic Radius": "166 pm",
                "Ionization Energy": "6.767 eV",
                "Specific Heat": "0.449 J/(g·K)",
                "Group": 6,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Louis Nicolas Vauquelin",
                "Year of Discovery": 1797
            },
            "Mn": {
                "Name": "Manganese",
                "Atomic Number": 25,
                "Atomic Mass": "54.938045 u",
                "Density": "7.44 g/cm³",
                "Melting Point": "1246 °C",
                "Boiling Point": "2061 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d⁵ 4s²",
                "Ground State": "⁶S₅/₂",
                "Electronegativity": 1.55,
                "Atomic Radius": "161 pm",
                "Ionization Energy": "7.434 eV",
                "Specific Heat": "0.479 J/(g·K)",
                "Group": 7,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Johan Gottlieb Gahn",
                "Year of Discovery": 1774
            },
            "Fe": {
                "Name": "Iron",
                "Atomic Number": 26,
                "Atomic Mass": "55.845 u",
                "Density": "7.874 g/cm³",
                "Melting Point": "1538 °C",
                "Boiling Point": "2861 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d⁶ 4s²",
                "Ground State": "⁵D₄",
                "Electronegativity": 1.83,
                "Atomic Radius": "156 pm",
                "Ionization Energy": "7.902 eV",
                "Specific Heat": "0.449 J/(g·K)",
                "Group": 8,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Co": {
                "Name": "Cobalt",
                "Atomic Number": 27,
                "Atomic Mass": "58.933200 u",
                "Density": "8.9 g/cm³",
                "Melting Point": "1495 °C",
                "Boiling Point": "2927 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d⁷ 4s²",
                "Ground State": "⁴F₉/₂",
                "Electronegativity": 1.88,
                "Atomic Radius": "152 pm",
                "Ionization Energy": "7.881 eV",
                "Specific Heat": "0.421 J/(g·K)",
                "Group": 9,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Georg Brandt",
                "Year of Discovery": 1735
            },
            "Ni": {
                "Name": "Nickel",
                "Atomic Number": 28,
                "Atomic Mass": "58.6934 u",
                "Density": "8.90225 g/cm³",
                "Melting Point": "1455 °C",
                "Boiling Point": "2913 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d⁸ 4s²",
                "Ground State": "³F₄",
                "Electronegativity": 1.91,
                "Atomic Radius": "149 pm",
                "Ionization Energy": "7.640 eV",
                "Specific Heat": "0.444 J/(g·K)",
                "Group": 10,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Axel Fredrik Cronstedt",
                "Year of Discovery": 1751
            },
            "Cu": {
                "Name": "Copper",
                "Atomic Number": 29,
                "Atomic Mass": "63.546 u",
                "Density": "8.96 g/cm³",
                "Melting Point": "1084.62 °C",
                "Boiling Point": "2562 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d¹⁰ 4s¹",
                "Ground State": "²S₁/₂",
                "Electronegativity": 1.90,
                "Atomic Radius": "145 pm",
                "Ionization Energy": "7.726 eV",
                "Specific Heat": "0.385 J/(g·K)",
                "Group": 11,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Zn": {
                "Name": "Zinc",
                "Atomic Number": 30,
                "Atomic Mass": "65.39 u",
                "Density": "7.13325 g/cm³",
                "Melting Point": "419.53 °C",
                "Boiling Point": "907 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d¹⁰ 4s²",
                "Ground State": "¹S₀",
                "Electronegativity": 1.65,
                "Atomic Radius": "142 pm",
                "Ionization Energy": "9.394 eV",
                "Specific Heat": "0.388 J/(g·K)",
                "Group": 12,
                "Period": 4,
                "Category": "Transition Metal",
                "Discovered By": "Andreas Sigismund Marggraf",
                "Year of Discovery": 1746
            },
            # Continue with all other elements...
            # I'll provide a few more key examples, but you should follow this pattern for all elements:

            "Ga": {
                "Name": "Gallium",
                "Atomic Number": 31,
                "Atomic Mass": "69.723 u",
                "Density": "5.90429.6 g/cm³",
                "Melting Point": "29.76 °C",
                "Boiling Point": "2204 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d¹⁰ 4s² 4p¹",
                "Ground State": "²P₁/₂",
                "Electronegativity": 1.81,
                "Atomic Radius": "136 pm",
                "Ionization Energy": "5.999 eV",
                "Specific Heat": "0.371 J/(g·K)",
                "Group": 13,
                "Period": 4,
                "Category": "Post-transition Metal",
                "Discovered By": "Lecoq de Boisbaudran",
                "Year of Discovery": 1875
            },
            "Ge": {
            "Name": "Germanium",
            "Atomic Number": 32,
            "Atomic Mass": "72.61 u",
            "Density": "5.32325 g/cm³",
            "Melting Point": "938.25 °C",
            "Boiling Point": "2833 °C",
            "State at 20°C": "Solid",
            "Electron Configuration": "[Ar] 3d¹⁰ 4s² 4p²",
            "Ground State": "³P₀",
            "Electronegativity": 2.01,
            "Atomic Radius": "125 pm",
            "Ionization Energy": "7.899 eV",
            "Specific Heat": "0.320 J/(g·K)",
            "Group": 14,
            "Period": 4,
            "Category": "Metalloid",
            "Discovered By": "Clemens Winkler",
            "Year of Discovery": 1886
            },
            "As": {
                "Name": "Arsenic",
                "Atomic Number": 33,
                "Atomic Mass": "74.92160 u",
                "Density": "5.73 g/cm³",
                "Melting Point": "817 °C (at 28 atm)",
                "Boiling Point": "603 °C (sublimes)",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d¹⁰ 4s² 4p³",
                "Ground State": "⁴S³/₂",
                "Electronegativity": 2.18,
                "Atomic Radius": "114 pm",
                "Ionization Energy": "9.789 eV",
                "Specific Heat": "0.329 J/(g·K)",
                "Group": 15,
                "Period": 4,
                "Category": "Metalloid",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Se": {
                "Name": "Selenium",
                "Atomic Number": 34,
                "Atomic Mass": "78.96 u",
                "Density": "4.79 g/cm³",
                "Melting Point": "220.5 °C",
                "Boiling Point": "685 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Ar] 3d¹⁰ 4s² 4p⁴",
                "Ground State": "³P₂",
                "Electronegativity": 2.55,
                "Atomic Radius": "103 pm",
                "Ionization Energy": "9.752 eV",
                "Specific Heat": "0.321 J/(g·K)",
                "Group": 16,
                "Period": 4,
                "Category": "Nonmetal",
                "Discovered By": "Jöns Jakob Berzelius",
                "Year of Discovery": 1817
            },
            "Br": {
                "Name": "Bromine",
                "Atomic Number": 35,
                "Atomic Mass": "79.904 u",
                "Density": "3.12 g/cm³",
                "Melting Point": "-7.2 °C",
                "Boiling Point": "58.8 °C",
                "State at 20°C": "Liquid",
                "Electron Configuration": "[Ar] 3d¹⁰ 4s² 4p⁵",
                "Ground State": "²P₃/₂",
                "Electronegativity": 2.96,
                "Atomic Radius": "94 pm",
                "Ionization Energy": "11.814 eV",
                "Specific Heat": "0.226 J/(g·K)",
                "Group": 17,
                "Period": 4,
                "Category": "Halogen",
                "Discovered By": "Antoine Jérôme Balard",
                "Year of Discovery": 1826
            },
            "Kr": {
                "Name": "Krypton",
                "Atomic Number": 36,
                "Atomic Mass": "83.80 u",
                "Density": "2.16 g/cm³",
                "Melting Point": "-157.37 °C (at 73.2 kPa)",
                "Boiling Point": "-153.22 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "[Ar] 3d¹⁰ 4s² 4p⁶",
                "Ground State": "¹S₀",
                "Electronegativity": 3.00,
                "Atomic Radius": "88 pm",
                "Ionization Energy": "14.000 eV",
                "Specific Heat": "0.248 J/(g·K)",
                "Group": 18,
                "Period": 4,
                "Category": "Noble Gas",
                "Discovered By": "William Ramsay and Morris Travers",
                "Year of Discovery": 1898
            },
            "Rb": {
                "Name": "Rubidium",
                "Atomic Number": 37,
                "Atomic Mass": "85.4678 u",
                "Density": "1.532 g/cm³",
                "Melting Point": "39.30 °C",
                "Boiling Point": "688 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 5s¹",
                "Ground State": "²S₁/₂",
                "Electronegativity": 0.82,
                "Atomic Radius": "265 pm",
                "Ionization Energy": "4.177 eV",
                "Specific Heat": "0.363 J/(g·K)",
                "Group": 1,
                "Period": 5,
                "Category": "Alkali Metal",
                "Discovered By": "Robert Bunsen and Gustav Kirchhoff",
                "Year of Discovery": 1861
            },
            "Sr": {
                "Name": "Strontium",
                "Atomic Number": 38,
                "Atomic Mass": "87.62 u",
                "Density": "2.54 g/cm³",
                "Melting Point": "777 °C",
                "Boiling Point": "1382 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 5s²",
                "Ground State": "¹S₀",
                "Electronegativity": 0.95,
                "Atomic Radius": "219 pm",
                "Ionization Energy": "5.695 eV",
                "Specific Heat": "0.301 J/(g·K)",
                "Group": 2,
                "Period": 5,
                "Category": "Alkaline Earth Metal",
                "Discovered By": "Adair Crawford",
                "Year of Discovery": 1790
            },
            "Y": {
                "Name": "Yttrium",
                "Atomic Number": 39,
                "Atomic Mass": "88.90585 u",
                "Density": "4.46925 g/cm³",
                "Melting Point": "1522 °C",
                "Boiling Point": "3345 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹ 5s²",
                "Ground State": "²D₃/₂",
                "Electronegativity": 1.22,
                "Atomic Radius": "212 pm",
                "Ionization Energy": "6.217 eV",
                "Specific Heat": "0.298 J/(g·K)",
                "Group": 3,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Johan Gadolin",
                "Year of Discovery": 1794
            },
            "Zr": {
                "Name": "Zirconium",
                "Atomic Number": 40,
                "Atomic Mass": "91.224 u",
                "Density": "6.506 g/cm³",
                "Melting Point": "1855 °C",
                "Boiling Point": "4409 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d² 5s²",
                "Ground State": "³F₂",
                "Electronegativity": 1.33,
                "Atomic Radius": "206 pm",
                "Ionization Energy": "6.634 eV",
                "Specific Heat": "0.278 J/(g·K)",
                "Group": 4,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Martin Heinrich Klaproth",
                "Year of Discovery": 1789
            },
            "Nb": {
                "Name": "Niobium",
                "Atomic Number": 41,
                "Atomic Mass": "92.90638 u",
                "Density": "8.57 g/cm³",
                "Melting Point": "2477 °C",
                "Boiling Point": "4744 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d⁴ 5s¹",
                "Ground State": "⁶D₁/₂",
                "Electronegativity": 1.6,
                "Atomic Radius": "198 pm",
                "Ionization Energy": "6.759 eV",
                "Specific Heat": "0.265 J/(g·K)",
                "Group": 5,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Charles Hatchett",
                "Year of Discovery": 1801
            },
            "Mo": {
                "Name": "Molybdenum",
                "Atomic Number": 42,
                "Atomic Mass": "95.94 u",
                "Density": "10.22 g/cm³",
                "Melting Point": "2623 °C",
                "Boiling Point": "4639 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d⁵ 5s¹",
                "Ground State": "⁷S₃",
                "Electronegativity": 2.16,
                "Atomic Radius": "190 pm",
                "Ionization Energy": "7.092 eV",
                "Specific Heat": "0.251 J/(g·K)",
                "Group": 6,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Carl Wilhelm Scheele",
                "Year of Discovery": 1778
            },
            "Tc": {
                "Name": "Technetium",
                "Atomic Number": 43,
                "Atomic Mass": "(98) u",
                "Density": "11.50 g/cm³",
                "Melting Point": "2157 °C",
                "Boiling Point": "4265 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d⁵ 5s²",
                "Ground State": "⁶S₅/₂",
                "Electronegativity": 1.9,
                "Atomic Radius": "183 pm",
                "Ionization Energy": "7.28 eV",
                "Specific Heat": "—",
                "Group": 7,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Carlo Perrier and Emilio Segrè",
                "Year of Discovery": 1937
            },
            "Ru": {
                "Name": "Ruthenium",
                "Atomic Number": 44,
                "Atomic Mass": "101.07 u",
                "Density": "12.41 g/cm³",
                "Melting Point": "2334 °C",
                "Boiling Point": "4150 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d⁷ 5s¹",
                "Ground State": "⁵F₅",
                "Electronegativity": 2.2,
                "Atomic Radius": "178 pm",
                "Ionization Energy": "7.360 eV",
                "Specific Heat": "0.238 J/(g·K)",
                "Group": 8,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Karl Ernst Claus",
                "Year of Discovery": 1844
            },
            "Rh": {
                "Name": "Rhodium",
                "Atomic Number": 45,
                "Atomic Mass": "102.90550 u",
                "Density": "12.41 g/cm³",
                "Melting Point": "1964 °C",
                "Boiling Point": "3695 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d⁸ 5s¹",
                "Ground State": "⁴F₉/₂",
                "Electronegativity": 2.28,
                "Atomic Radius": "173 pm",
                "Ionization Energy": "7.459 eV",
                "Specific Heat": "0.243 J/(g·K)",
                "Group": 9,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "William Hyde Wollaston",
                "Year of Discovery": 1804
            },
            "Pd": {
                "Name": "Palladium",
                "Atomic Number": 46,
                "Atomic Mass": "106.42 u",
                "Density": "12.02 g/cm³",
                "Melting Point": "1554.9 °C",
                "Boiling Point": "2963 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰",
                "Ground State": "¹S₀",
                "Electronegativity": 2.20,
                "Atomic Radius": "169 pm",
                "Ionization Energy": "8.337 eV",
                "Specific Heat": "0.246 J/(g·K)",
                "Group": 10,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "William Hyde Wollaston",
                "Year of Discovery": 1803
            },
            "Ag": {
                "Name": "Silver",
                "Atomic Number": 47,
                "Atomic Mass": "107.8682 u",
                "Density": "10.50 g/cm³",
                "Melting Point": "961.78 °C",
                "Boiling Point": "2162 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s¹",
                "Ground State": "²S₁/₂",
                "Electronegativity": 1.93,
                "Atomic Radius": "165 pm",
                "Ionization Energy": "7.576 eV",
                "Specific Heat": "0.235 J/(g·K)",
                "Group": 11,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Cd": {
                "Name": "Cadmium",
                "Atomic Number": 48,
                "Atomic Mass": "112.411 u",
                "Density": "8.65 g/cm³",
                "Melting Point": "321.07 °C",
                "Boiling Point": "767 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s²",
                "Ground State": "¹S₀",
                "Electronegativity": 1.69,
                "Atomic Radius": "161 pm",
                "Ionization Energy": "8.994 eV",
                "Specific Heat": "0.232 J/(g·K)",
                "Group": 12,
                "Period": 5,
                "Category": "Transition Metal",
                "Discovered By": "Friedrich Stromeyer",
                "Year of Discovery": 1817
            },
            "In": {
                "Name": "Indium",
                "Atomic Number": 49,
                "Atomic Mass": "114.818 u",
                "Density": "7.31 g/cm³",
                "Melting Point": "156.60 °C",
                "Boiling Point": "2072 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s² 5p¹",
                "Ground State": "²P₁/₂",
                "Electronegativity": 1.78,
                "Atomic Radius": "156 pm",
                "Ionization Energy": "5.786 eV",
                "Specific Heat": "0.233 J/(g·K)",
                "Group": 13,
                "Period": 5,
                "Category": "Post-transition Metal",
                "Discovered By": "Ferdinand Reich and Hieronymous Theodor Richter",
                "Year of Discovery": 1863
            },
            "Sn": {
                "Name": "Tin",
                "Atomic Number": 50,
                "Atomic Mass": "118.710 u",
                "Density": "7.31 g/cm³",
                "Melting Point": "231.93 °C",
                "Boiling Point": "2602 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s² 5p²",
                "Ground State": "³P₀",
                "Electronegativity": 1.96,
                "Atomic Radius": "145 pm",
                "Ionization Energy": "7.344 eV",
                "Specific Heat": "0.228 J/(g·K)",
                "Group": 14,
                "Period": 5,
                "Category": "Post-transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Sb": {
                "Name": "Antimony",
                "Atomic Number": 51,
                "Atomic Mass": "121.760 u",
                "Density": "6.691 g/cm³",
                "Melting Point": "630.73 °C",
                "Boiling Point": "1587 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s² 5p³",
                "Ground State": "⁴S₃/₂",
                "Electronegativity": 2.05,
                "Atomic Radius": "133 pm",
                "Ionization Energy": "8.608 eV",
                "Specific Heat": "0.207 J/(g·K)",
                "Group": 15,
                "Period": 5,
                "Category": "Metalloid",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Te": {
                "Name": "Tellurium",
                "Atomic Number": 52,
                "Atomic Mass": "127.60 u",
                "Density": "6.24 g/cm³",
                "Melting Point": "449.51 °C",
                "Boiling Point": "988 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s² 5p⁴",
                "Ground State": "³P₂",
                "Electronegativity": 2.1,
                "Atomic Radius": "123 pm",
                "Ionization Energy": "9.010 eV",
                "Specific Heat": "0.202 J/(g·K)",
                "Group": 16,
                "Period": 5,
                "Category": "Metalloid",
                "Discovered By": "Franz-Joseph Müller von Reichenstein",
                "Year of Discovery": 1782
            },
            "I": {
                "Name": "Iodine",
                "Atomic Number": 53,
                "Atomic Mass": "126.90447 u",
                "Density": "4.93 g/cm³",
                "Melting Point": "113.7 °C",
                "Boiling Point": "184.4 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s² 5p⁵",
                "Ground State": "²P₃/₂",
                "Electronegativity": 2.66,
                "Atomic Radius": "115 pm",
                "Ionization Energy": "10.451 eV",
                "Specific Heat": "0.145 J/(g·K)",
                "Group": 17,
                "Period": 5,
                "Category": "Halogen",
                "Discovered By": "Bernard Courtois",
                "Year of Discovery": 1811
            },
            "Xe": {
                "Name": "Xenon",
                "Atomic Number": 54,
                "Atomic Mass": "131.29 u",
                "Density": "3.52 g/cm³",
                "Melting Point": "-111.79 °C (at 81.6 kPa)",
                "Boiling Point": "-108.12 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "[Kr] 4d¹⁰ 5s² 5p⁶",
                "Ground State": "¹S₀",
                "Electronegativity": 2.60,
                "Atomic Radius": "108 pm",
                "Ionization Energy": "12.130 eV",
                "Specific Heat": "0.158 J/(g·K)",
                "Group": 18,
                "Period": 5,
                "Category": "Noble Gas",
                "Discovered By": "William Ramsay and Morris Travers",
                "Year of Discovery": 1898
            },
            "Cs": {
                "Name": "Cesium",
                "Atomic Number": 55,
                "Atomic Mass": "132.90545 u",
                "Density": "1.873 g/cm³",
                "Melting Point": "28.5 °C",
                "Boiling Point": "671 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 6s¹",
                "Ground State": "²S₁/₂",
                "Electronegativity": 0.79,
                "Atomic Radius": "298 pm",
                "Ionization Energy": "3.894 eV",
                "Specific Heat": "0.242 J/(g·K)",
                "Group": 1,
                "Period": 6,
                "Category": "Alkali Metal",
                "Discovered By": "Robert Bunsen and Gustav Kirchhoff",
                "Year of Discovery": 1860
            },
            "Ba": {
                "Name": "Barium",
                "Atomic Number": 56,
                "Atomic Mass": "137.327 u",
                "Density": "3.5 g/cm³",
                "Melting Point": "727 °C",
                "Boiling Point": "1897 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 6s²",
                "Ground State": "¹S₀",
                "Electronegativity": 0.89,
                "Atomic Radius": "253 pm",
                "Ionization Energy": "5.212 eV",
                "Specific Heat": "0.204 J/(g·K)",
                "Group": 2,
                "Period": 6,
                "Category": "Alkaline Earth Metal",
                "Discovered By": "Humphry Davy",
                "Year of Discovery": 1808
            },
            "La": {
                "Name": "Lanthanum",
                "Atomic Number": 57,
                "Atomic Mass": "138.9055 u",
                "Density": "6.14525 g/cm³",
                "Melting Point": "918 °C",
                "Boiling Point": "3464 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 5d¹ 6s²",
                "Ground State": "²D₃/₂",
                "Electronegativity": 1.10,
                "Atomic Radius": "240 pm",
                "Ionization Energy": "5.577 eV",
                "Specific Heat": "0.195 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Carl Gustaf Mosander",
                "Year of Discovery": 1839
            },
            "Ce": {
                "Name": "Cerium",
                "Atomic Number": 58,
                "Atomic Mass": "140.116 u",
                "Density": "6.77025 g/cm³",
                "Melting Point": "798 °C",
                "Boiling Point": "3443 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹ 5d¹ 6s²",
                "Ground State": "¹G₄",
                "Electronegativity": 1.12,
                "Atomic Radius": "235 pm",
                "Ionization Energy": "5.539 eV",
                "Specific Heat": "0.192 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Martin Heinrich Klaproth, Jöns Jakob Berzelius, Wilhelm Hisinger",
                "Year of Discovery": 1803
            },
            "Pr": {
                "Name": "Praseodymium",
                "Atomic Number": 59,
                "Atomic Mass": "140.90765 u",
                "Density": "6.773 g/cm³",
                "Melting Point": "931 °C",
                "Boiling Point": "3520 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f³ 6s²",
                "Ground State": "⁴I₉/₂",
                "Electronegativity": 1.13,
                "Atomic Radius": "239 pm",
                "Ionization Energy": "5.473 eV",
                "Specific Heat": "0.193 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Carl Auer von Welsbach",
                "Year of Discovery": 1885
            },
            "Nd": {
                "Name": "Neodymium",
                "Atomic Number": 60,
                "Atomic Mass": "144.24 u",
                "Density": "7.00825 g/cm³",
                "Melting Point": "1021 °C",
                "Boiling Point": "3074 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f⁴ 6s²",
                "Ground State": "⁵I₄",
                "Electronegativity": 1.14,
                "Atomic Radius": "229 pm",
                "Ionization Energy": "5.525 eV",
                "Specific Heat": "0.190 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Carl Auer von Welsbach",
                "Year of Discovery": 1885
            },
            "Pm": {
                "Name": "Promethium",
                "Atomic Number": 61,
                "Atomic Mass": "(145) u",
                "Density": "7.26425 g/cm³",
                "Melting Point": "1042 °C",
                "Boiling Point": "3000 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f⁵ 6s²",
                "Ground State": "⁶H₅/₂",
                "Electronegativity": 1.13,
                "Atomic Radius": "236 pm",
                "Ionization Energy": "5.582 eV",
                "Specific Heat": "—",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Jacob A. Marinsky, Lawrence E. Glendenin, Charles D. Coryell",
                "Year of Discovery": 1945
            },
            "Sm": {
                "Name": "Samarium",
                "Atomic Number": 62,
                "Atomic Mass": "150.36 u",
                "Density": "7.52025 g/cm³",
                "Melting Point": "1074 °C",
                "Boiling Point": "1794 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f⁶ 6s²",
                "Ground State": "⁷F₀",
                "Electronegativity": 1.17,
                "Atomic Radius": "229 pm",
                "Ionization Energy": "5.644 eV",
                "Specific Heat": "0.197 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Paul Émile Lecoq de Boisbaudran",
                "Year of Discovery": 1879
            },
            "Eu": {
                "Name": "Europium",
                "Atomic Number": 63,
                "Atomic Mass": "151.964 u",
                "Density": "5.24425 g/cm³",
                "Melting Point": "822 °C",
                "Boiling Point": "1529 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f⁷ 6s²",
                "Ground State": "⁸S₇/₂",
                "Electronegativity": 1.2,
                "Atomic Radius": "233 pm",
                "Ionization Energy": "5.670 eV",
                "Specific Heat": "0.182 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Eugène-Anatole Demarçay",
                "Year of Discovery": 1901
            },
            "Gd": {
                "Name": "Gadolinium",
                "Atomic Number": 64,
                "Atomic Mass": "157.25 u",
                "Density": "7.90125 g/cm³",
                "Melting Point": "1313 °C",
                "Boiling Point": "3273 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f⁷ 5d¹ 6s²",
                "Ground State": "⁹D₂",
                "Electronegativity": 1.20,
                "Atomic Radius": "237 pm",
                "Ionization Energy": "6.150 eV",
                "Specific Heat": "0.236 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Jean Charles Galissard de Marignac",
                "Year of Discovery": 1880
            },
            "Tb": {
                "Name": "Terbium",
                "Atomic Number": 65,
                "Atomic Mass": "158.92534 u",
                "Density": "8.230 g/cm³",
                "Melting Point": "1356 °C",
                "Boiling Point": "3230 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f⁹ 6s²",
                "Ground State": "⁶H₁₅/₂",
                "Electronegativity": 1.10,
                "Atomic Radius": "221 pm",
                "Ionization Energy": "5.864 eV",
                "Specific Heat": "0.182 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Carl Gustaf Mosander",
                "Year of Discovery": 1843
            },
            "Dy": {
                "Name": "Dysprosium",
                "Atomic Number": 66,
                "Atomic Mass": "162.50 u",
                "Density": "8.55125 g/cm³",
                "Melting Point": "1412 °C",
                "Boiling Point": "2567 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁰ 6s²",
                "Ground State": "⁵I₈",
                "Electronegativity": 1.22,
                "Atomic Radius": "229 pm",
                "Ionization Energy": "5.939 eV",
                "Specific Heat": "0.170 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Paul Émile Lecoq de Boisbaudran",
                "Year of Discovery": 1886
            },
            "Ho": {
                "Name": "Holmium",
                "Atomic Number": 67,
                "Atomic Mass": "164.93032 u",
                "Density": "8.79525 g/cm³",
                "Melting Point": "1474 °C",
                "Boiling Point": "2700 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹¹ 6s²",
                "Ground State": "⁴I₁₅/₂",
                "Electronegativity": 1.23,
                "Atomic Radius": "216 pm",
                "Ionization Energy": "6.022 eV",
                "Specific Heat": "0.165 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Marc Delafontaine, Jacques-Louis Soret",
                "Year of Discovery": 1878
            },
            "Er": {
                "Name": "Erbium",
                "Atomic Number": 68,
                "Atomic Mass": "167.26 u",
                "Density": "9.06625 g/cm³",
                "Melting Point": "1529 °C",
                "Boiling Point": "2868 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹² 6s²",
                "Ground State": "³H₆",
                "Electronegativity": 1.24,
                "Atomic Radius": "235 pm",
                "Ionization Energy": "6.108 eV",
                "Specific Heat": "0.168 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Carl Gustaf Mosander",
                "Year of Discovery": 1843
            },
            "Tm": {
                "Name": "Thulium",
                "Atomic Number": 69,
                "Atomic Mass": "168.93421 u",
                "Density": "9.32125 g/cm³",
                "Melting Point": "1545 °C",
                "Boiling Point": "1950 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹³ 6s²",
                "Ground State": "²F₇/₂",
                "Electronegativity": 1.25,
                "Atomic Radius": "227 pm",
                "Ionization Energy": "6.184 eV",
                "Specific Heat": "0.160 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Per Teodor Cleve",
                "Year of Discovery": 1879
            },
            "Yb": {
                "Name": "Ytterbium",
                "Atomic Number": 70,
                "Atomic Mass": "173.04 u",
                "Density": "6.966 g/cm³",
                "Melting Point": "819 °C",
                "Boiling Point": "1196 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 6s²",
                "Ground State": "¹S₀",
                "Electronegativity": 1.1,
                "Atomic Radius": "242 pm",
                "Ionization Energy": "6.254 eV",
                "Specific Heat": "0.155 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Jean Charles Galissard de Marignac",
                "Year of Discovery": 1878
            },
            "Lu": {
                "Name": "Lutetium",
                "Atomic Number": 71,
                "Atomic Mass": "174.967 u",
                "Density": "9.84125 g/cm³",
                "Melting Point": "1663 °C",
                "Boiling Point": "3402 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹ 6s²",
                "Ground State": "²D₃/₂",
                "Electronegativity": 1.27,
                "Atomic Radius": "221 pm",
                "Ionization Energy": "5.426 eV",
                "Specific Heat": "0.154 J/(g·K)",
                "Group": 3,
                "Period": 6,
                "Category": "Lanthanide",
                "Discovered By": "Georges Urbain",
                "Year of Discovery": 1907
            },
            "Hf": {
                "Name": "Hafnium",
                "Atomic Number": 72,
                "Atomic Mass": "178.49 u",
                "Density": "13.31 g/cm³",
                "Melting Point": "2233 °C",
                "Boiling Point": "4603 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d² 6s²",
                "Ground State": "³F₂",
                "Electronegativity": 1.3,
                "Atomic Radius": "212 pm",
                "Ionization Energy": "6.825 eV",
                "Specific Heat": "0.144 J/(g·K)",
                "Group": 4,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Dirk Coster and George de Hevesy",
                "Year of Discovery": 1923
            },
            "Ta": {
                "Name": "Tantalum",
                "Atomic Number": 73,
                "Atomic Mass": "180.9479 u",
                "Density": "16.654 g/cm³",
                "Melting Point": "3017 °C",
                "Boiling Point": "5458 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d³ 6s²",
                "Ground State": "⁴F₃/₂",
                "Electronegativity": 1.5,
                "Atomic Radius": "206 pm",
                "Ionization Energy": "7.550 eV",
                "Specific Heat": "0.140 J/(g·K)",
                "Group": 5,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Anders Gustaf Ekeberg",
                "Year of Discovery": 1802
            },
            "W": {
                "Name": "Tungsten",
                "Atomic Number": 74,
                "Atomic Mass": "183.84 u",
                "Density": "19.3 g/cm³",
                "Melting Point": "3422 °C",
                "Boiling Point": "5555 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d⁴ 6s²",
                "Ground State": "⁵D₀",
                "Electronegativity": 2.36,
                "Atomic Radius": "200 pm",
                "Ionization Energy": "7.864 eV",
                "Specific Heat": "0.132 J/(g·K)",
                "Group": 6,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Fausto and Juan José de Elhuyar",
                "Year of Discovery": 1783
            },
            "Re": {
                "Name": "Rhenium",
                "Atomic Number": 75,
                "Atomic Mass": "186.207 u",
                "Density": "21.02 g/cm³",
                "Melting Point": "3186 °C",
                "Boiling Point": "5596 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d⁵ 6s²",
                "Ground State": "⁶S₅/₂",
                "Electronegativity": 1.9,
                "Atomic Radius": "197 pm",
                "Ionization Energy": "7.834 eV",
                "Specific Heat": "0.137 J/(g·K)",
                "Group": 7,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Masataka Ogawa, Walter Noddack, Ida Tacke, Otto Berg",
                "Year of Discovery": 1925
            },
            "Os": {
                "Name": "Osmium",
                "Atomic Number": 76,
                "Atomic Mass": "190.23 u",
                "Density": "22.57 g/cm³",
                "Melting Point": "3033 °C",
                "Boiling Point": "5012 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d⁶ 6s²",
                "Ground State": "⁵D₄",
                "Electronegativity": 2.2,
                "Atomic Radius": "193 pm",
                "Ionization Energy": "8.438 eV",
                "Specific Heat": "0.130 J/(g·K)",
                "Group": 8,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Smithson Tennant",
                "Year of Discovery": 1803
            },
            "Ir": {
                "Name": "Iridium",
                "Atomic Number": 77,
                "Atomic Mass": "192.217 u",
                "Density": "22.4217 g/cm³",
                "Melting Point": "2446 °C",
                "Boiling Point": "4428 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d⁷ 6s²",
                "Ground State": "⁴F₉/₂",
                "Electronegativity": 2.20,
                "Atomic Radius": "190 pm",
                "Ionization Energy": "8.967 eV",
                "Specific Heat": "0.131 J/(g·K)",
                "Group": 9,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Smithson Tennant",
                "Year of Discovery": 1803
            },
            "Pt": {
                "Name": "Platinum",
                "Atomic Number": 78,
                "Atomic Mass": "195.078 u",
                "Density": "21.45 g/cm³",
                "Melting Point": "1768.4 °C",
                "Boiling Point": "3825 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d⁹ 6s¹",
                "Ground State": "³D₃",
                "Electronegativity": 2.28,
                "Atomic Radius": "187 pm",
                "Ionization Energy": "8.959 eV",
                "Specific Heat": "0.133 J/(g·K)",
                "Group": 10,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Antonio de Ulloa",
                "Year of Discovery": 1735
            },
            "Au": {
                "Name": "Gold",
                "Atomic Number": 79,
                "Atomic Mass": "196.96655 u",
                "Density": "19.3 g/cm³",
                "Melting Point": "1064.18 °C",
                "Boiling Point": "2856 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s¹",
                "Ground State": "²S₁/₂",
                "Electronegativity": 2.54,
                "Atomic Radius": "184 pm",
                "Ionization Energy": "9.226 eV",
                "Specific Heat": "0.129 J/(g·K)",
                "Group": 11,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Hg": {
                "Name": "Mercury",
                "Atomic Number": 80,
                "Atomic Mass": "200.59 u",
                "Density": "13.546 g/cm³",
                "Melting Point": "-38.83 °C",
                "Boiling Point": "356.73 °C",
                "State at 20°C": "Liquid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s²",
                "Ground State": "¹S₀",
                "Electronegativity": 2.00,
                "Atomic Radius": "181 pm",
                "Ionization Energy": "10.438 eV",
                "Specific Heat": "0.140 J/(g·K)",
                "Group": 12,
                "Period": 6,
                "Category": "Transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Tl": {
                "Name": "Thallium",
                "Atomic Number": 81,
                "Atomic Mass": "204.3833 u",
                "Density": "11.85 g/cm³",
                "Melting Point": "304 °C",
                "Boiling Point": "1473 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s² 6p¹",
                "Ground State": "²P₁/₂",
                "Electronegativity": 1.62,
                "Atomic Radius": "156 pm",
                "Ionization Energy": "6.108 eV",
                "Specific Heat": "0.129 J/(g·K)",
                "Group": 13,
                "Period": 6,
                "Category": "Post-transition Metal",
                "Discovered By": "William Crookes",
                "Year of Discovery": 1861
            },
            "Pb": {
                "Name": "Lead",
                "Atomic Number": 82,
                "Atomic Mass": "207.2 u",
                "Density": "11.35 g/cm³",
                "Melting Point": "327.46 °C",
                "Boiling Point": "1749 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s² 6p²",
                "Ground State": "³P₀",
                "Electronegativity": 1.87,
                "Atomic Radius": "154 pm",
                "Ionization Energy": "7.417 eV",
                "Specific Heat": "0.129 J/(g·K)",
                "Group": 14,
                "Period": 6,
                "Category": "Post-transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Bi": {
                "Name": "Bismuth",
                "Atomic Number": 83,
                "Atomic Mass": "208.98038 u",
                "Density": "9.747 g/cm³",
                "Melting Point": "271.40 °C",
                "Boiling Point": "1564 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s² 6p³",
                "Ground State": "⁴S₃/₂",
                "Electronegativity": 2.02,
                "Atomic Radius": "143 pm",
                "Ionization Energy": "7.286 eV",
                "Specific Heat": "0.122 J/(g·K)",
                "Group": 15,
                "Period": 6,
                "Category": "Post-transition Metal",
                "Discovered By": "Known since ancient times",
                "Year of Discovery": "prehistoric"
            },
            "Po": {
                "Name": "Polonium",
                "Atomic Number": 84,
                "Atomic Mass": "(209) u",
                "Density": "9.32 g/cm³",
                "Melting Point": "254 °C",
                "Boiling Point": "962 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s² 6p⁴",
                "Ground State": "³P₂",
                "Electronegativity": 2.0,
                "Atomic Radius": "135 pm",
                "Ionization Energy": "8.417 eV",
                "Specific Heat": "—",
                "Group": 16,
                "Period": 6,
                "Category": "Metalloid",
                "Discovered By": "Marie and Pierre Curie",
                "Year of Discovery": 1898
            },
            "At": {
                "Name": "Astatine",
                "Atomic Number": 85,
                "Atomic Mass": "(210) u",
                "Density": "—",
                "Melting Point": "302 °C",
                "Boiling Point": "—",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s² 6p⁵",
                "Ground State": "²P₃/₂",
                "Electronegativity": 2.2,
                "Atomic Radius": "127 pm",
                "Ionization Energy": "—",
                "Specific Heat": "—",
                "Group": 17,
                "Period": 6,
                "Category": "Halogen",
                "Discovered By": "Dale R. Corson, Kenneth Ross MacKenzie, Emilio Segrè",
                "Year of Discovery": 1940
            },
            "Rn": {
                "Name": "Radon",
                "Atomic Number": 86,
                "Atomic Mass": "(222) u",
                "Density": "—",
                "Melting Point": "-71 °C",
                "Boiling Point": "-61.7 °C",
                "State at 20°C": "Gas",
                "Electron Configuration": "[Xe] 4f¹⁴ 5d¹⁰ 6s² 6p⁶",
                "Ground State": "¹S₀",
                "Electronegativity": 2.2,
                "Atomic Radius": "120 pm",
                "Ionization Energy": "10.748 eV",
                "Specific Heat": "0.094 J/(g·K)",
                "Group": 18,
                "Period": 6,
                "Category": "Noble Gas",
                "Discovered By": "Friedrich Ernst Dorn",
                "Year of Discovery": 1900
            },
            "Fr": {
                "Name": "Francium",
                "Atomic Number": 87,
                "Atomic Mass": "(223) u",
                "Density": "—",
                "Melting Point": "27 °C",
                "Boiling Point": "—",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Rn] 7s¹",
                "Ground State": "²S₁/₂",
                "Electronegativity": 0.79,
                "Atomic Radius": "348 pm",
                "Ionization Energy": "4.073 eV",
                "Specific Heat": "—",
                "Group": 1,
                "Period": 7,
                "Category": "Alkali Metal",
                "Discovered By": "Marguerite Perey",
                "Year of Discovery": 1939
            },
            "Ra": {
                "Name": "Radium",
                "Atomic Number": 88,
                "Atomic Mass": "(226) u",
                "Density": "—",
                "Melting Point": "700 °C",
                "Boiling Point": "—",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Rn] 7s²",
                "Ground State": "¹S₀",
                "Electronegativity": 0.9,
                "Atomic Radius": "283 pm",
                "Ionization Energy": "5.278 eV",
                "Specific Heat": "—",
                "Group": 2,
                "Period": 7,
                "Category": "Alkaline Earth Metal",
                "Discovered By": "Marie and Pierre Curie",
                "Year of Discovery": 1898
            },
            "Ac": {
                "Name": "Actinium",
                "Atomic Number": 89,
                "Atomic Mass": "(227) u",
                "Density": "—",
                "Melting Point": "1051 °C",
                "Boiling Point": "3198 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Rn] 6d¹ 7s²",
                "Ground State": "²D₃/₂",
                "Electronegativity": 1.1,
                "Atomic Radius": "260 pm",
                "Ionization Energy": "5.17 eV",
                "Specific Heat": "0.120 J/(g·K)",
                "Group": 3,
                "Period": 7,
                "Category": "Actinide",
                "Discovered By": "André-Louis Debierne, Friedrich Oskar Giesel",
                "Year of Discovery": 1899
            },
            "Th": {
                "Name": "Thorium",
                "Atomic Number": 90,
                "Atomic Mass": "232.0381 u",
                "Density": "11.72 g/cm³",
                "Melting Point": "1750 °C",
                "Boiling Point": "4788 °C",
                "State at 20°C": "Solid",
                "Electron Configuration": "[Rn] 6d² 7s²",
                "Ground State": "³F₂",
                "Electronegativity": 1.3,
                "Atomic Radius": "237 pm",
                "Ionization Energy": "6.307 eV",
                "Specific Heat": "0.113 J/(g·K)",
                "Group": 3,
                "Period": 7,
                "Category": "Actinide",
                "Discovered By": "Jöns Jakob Berzelius",
                "Year of Discovery": 1829
            }

        }

        # Add more elements as needed...

        # Default properties for elements not explicitly defined
        default_properties = {
            "Name": element_symbol,
            "Atomic Number": "N/A",
            "Atomic Mass": "N/A",
            "Density": "N/A",
            "Melting Point": "N/A",
            "Boiling Point": "N/A",
            "State at 20°C": "N/A",
            "Electron Configuration": "N/A",
            "Ground State": "N/A",
            "Electronegativity": "N/A",
            "Ionization Energy": "N/A",
            "Specific Heat": "N/A",
            "Group": "N/A",
            "Period": "N/A",
            "Category": "N/A",
            "Common Core Levels": "N/A",
            "Most Intense Line": "N/A",
            "Typical FWHM": "N/A",
            "Chemical Shift Range": "N/A"
        }

        # Get known properties or default set
        properties = element_properties.get(element_symbol, default_properties)

        # Add URLs to properties
        properties["XPS Fitting URL"] = self.get_xps_fitting_url(element_symbol)
        properties["Thermo Fisher URL"] = self.get_thermo_url(element_symbol)

        return properties


class ElementTile(wx.Panel):
    """Custom widget for periodic table element tiles"""

    def __init__(self, parent, element, color, enabled=True, atomic_number=None, core_level=None, binding_energy=None):
        super().__init__(parent, size=(37, 37))
        self.element = element
        self.color = color
        self.enabled = enabled
        self.atomic_number = atomic_number or 0  # Default to 0 if not provided
        self.core_level = core_level or 'N.D.'
        self.binding_energy = binding_energy or 'N.D.'
        self.hover = False
        self.pressed = False

        # Bind paint and mouse events
        self.Bind(wx.EVT_PAINT, self.on_paint)
        self.Bind(wx.EVT_LEFT_DOWN, self.on_mouse_down)
        self.Bind(wx.EVT_LEFT_UP, self.on_mouse_up)
        self.Bind(wx.EVT_LEFT_DCLICK, self.on_double_click)
        self.Bind(wx.EVT_ENTER_WINDOW, self.on_enter)
        self.Bind(wx.EVT_LEAVE_WINDOW, self.on_leave)
        self.Bind(wx.EVT_MOTION, self.on_motion)

        # Callbacks
        self.click_callback = None
        self.double_click_callback = None

    def on_paint(self, event):
        """Draw the element tile with atomic number, element symbol, core level, and binding energy"""
        dc = wx.PaintDC(self)
        gc = wx.GraphicsContext.Create(dc)

        width, height = self.GetSize()

        # Determine the actual color to use (existing color logic)
        if not self.enabled:
            base_color = wx.Colour(self.color)
            r, g, b = base_color.Red(), base_color.Green(), base_color.Blue()
            r = min(255, int(r * 1.3))
            g = min(255, int(g * 1.3))
            b = min(255, int(b * 1.3))
            actual_color = wx.Colour(r, g, b)
        elif self.pressed and self.hover:
            base_color = wx.Colour(self.color)
            r, g, b = base_color.Red(), base_color.Green(), base_color.Blue()
            r = max(0, int(r * 0.8))
            g = max(0, int(g * 0.8))
            b = max(0, int(b * 0.8))
            actual_color = wx.Colour(r, g, b)
        elif self.hover:
            base_color = wx.Colour(self.color)
            r, g, b = base_color.Red(), base_color.Green(), base_color.Blue()
            r = max(0, int(r * 0.9))
            g = max(0, int(g * 0.9))
            b = max(0, int(b * 0.9))
            actual_color = wx.Colour(r, g, b)
        else:
            actual_color = wx.Colour(self.color)

        # Draw rectangle with thin border
        gc.SetPen(wx.Pen(wx.Colour(100, 100, 100), 1))
        gc.SetBrush(wx.Brush(actual_color))
        gc.DrawRoundedRectangle(0, 0, width - 1, height - 1, 2)

        # Set up fonts and colors - PLATFORM SPECIFIC
        text_color = wx.BLACK if self.enabled else wx.Colour(136, 136, 136)

        import platform
        if platform.system() == 'Darwin':  # macOS
            small_font_size = 7
            element_font_size = 13
            small_font = wx.Font(small_font_size, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
            element_font = wx.Font(element_font_size, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        else:  # Windows and other systems
            small_font_size = 7
            element_font_size = 11
            # Use Segoe UI which is clearer on Windows
            small_font = wx.Font(small_font_size, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL,
                                 faceName="Segoe UI")
            element_font = wx.Font(element_font_size, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD,
                                   faceName="Segoe UI")

        # 1. Draw atomic number in top-left corner
        if self.atomic_number and self.atomic_number > 0:
            gc.SetFont(small_font, text_color)
            gc.DrawText(str(self.atomic_number), 1, 1)

        # 2. Draw binding energy in top-right corner
        if self.binding_energy and self.binding_energy != 'N.D.':
            gc.SetFont(small_font, text_color)
            be_width, be_height = gc.GetTextExtent(self.binding_energy)
            be_x = width - be_width - 1  # Right-aligned with small margin
            gc.DrawText(self.binding_energy, be_x, 1)
        elif self.binding_energy == 'N.D.':
            gc.SetFont(small_font, text_color)
            be_width, be_height = gc.GetTextExtent('N.D.')
            be_x = width - be_width - 1
            gc.DrawText('N.D.', be_x, 1)

        # 3. Draw element symbol in center
        gc.SetFont(element_font, text_color)
        element_width, element_height = gc.GetTextExtent(self.element)
        element_x = (width - element_width) / 2
        element_y = (height - element_height) / 2 - 2  # Slightly higher to make room for core level
        gc.DrawText(self.element, element_x, element_y+1)

        # 4. Draw core level below element symbol
        if self.core_level and self.core_level != 'N.D.':
            gc.SetFont(small_font, text_color)
            core_width, core_height = gc.GetTextExtent(self.core_level)
            core_x = (width - core_width) / 2
            core_y = element_y + element_height + 0  # Below the element symbol
            gc.DrawText(self.core_level, core_x, core_y)
        elif self.core_level == 'N.D.':
            # Show N.D. for elements with no data
            gc.SetFont(small_font, text_color)
            core_width, core_height = gc.GetTextExtent('N.D.')
            core_x = (width - core_width) / 2
            core_y = element_y + element_height + 0
            gc.DrawText('N.D.', core_x, core_y)

    def on_mouse_down(self, event):
        """Handle mouse down"""
        if self.enabled:
            self.pressed = True
            self.Refresh()

    def on_mouse_up(self, event):
        """Handle mouse up"""
        if self.enabled and self.pressed:
            self.pressed = False
            self.Refresh()
            # Check if mouse is still over the tile
            x, y = event.GetPosition()
            width, height = self.GetSize()
            if 0 <= x <= width and 0 <= y <= height:
                if self.click_callback:
                    self.click_callback(self.element)

    def on_double_click(self, event):
        """Handle double click"""
        if self.enabled and self.double_click_callback:
            self.double_click_callback(self.element)

    def on_enter(self, event):
        """Handle mouse enter"""
        if self.enabled:
            self.hover = True
            self.Refresh()

    def on_leave(self, event):
        """Handle mouse leave"""
        self.hover = False
        self.pressed = False
        self.Refresh()

    def on_motion(self, event):
        """Handle mouse motion"""
        # This ensures hover state is maintained during mouse movement
        pass

    def set_click_callback(self, callback):
        """Set the click callback"""
        self.click_callback = callback

    def set_double_click_callback(self, callback):
        """Set the double-click callback"""
        self.double_click_callback = callback

def main():
    app = wx.App()
    frame = PeriodicTableXPS()
    frame.Show()
    app.MainLoop()


if __name__ == "__main__":
    main()