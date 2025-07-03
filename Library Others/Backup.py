class PeriodicTableWindow_OLD(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, title="Survey Identification / Labelling",
                         style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX))
        self.parent_window = parent  # Store the parent window
        # self.SetBackgroundColour(wx.WHITE)


        self.library_data = self.parent_window.library_data
        self.button_states = {}
        self.element_lines = {}
        self.InitUI()

        # Bind the close event
        self.Bind(wx.EVT_CLOSE, self.OnClose)

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

        # Split the window horizontally
        hsizer = wx.BoxSizer(wx.HORIZONTAL)

        # Left side: Periodic Table
        periodic_sizer = wx.BoxSizer(wx.VERTICAL)
        grid = wx.GridSizer(10, 18, 0, 0)

        elements = [
            "H", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "He",
            "Li", "Be", "", "", "", "", "", "", "", "", "", "", "B", "C", "N", "O", "F", "Ne",
            "Na", "Mg", "", "", "", "", "", "", "", "", "", "", "Al", "Si", "P", "S", "Cl", "Ar",
            "K", "Ca", "Sc", "Ti", "V", "Cr", "Mn", "Fe", "Co", "Ni", "Cu", "Zn", "Ga", "Ge", "As", "Se", "Br", "Kr",
            "Rb", "Sr", "Y", "Zr", "Nb", "Mo", "", "Ru", "Rh", "Pd", "Ag", "Cd", "In", "Sn", "Sb", "Te", "I", "Xe",
            "Cs", "Ba", "", "Hf", "Ta", "W", "Re", "Os", "Ir", "Pt", "Au", "Hg", "Tl", "Pb", "Bi", "", "At", "Rn",
            "", "Ra", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
            "", "", "La", "Ce", "Pr", "Nd", "Pm", "Sm", "Eu", "Gd", "Tb", "Dy", "Ho", "Er", "Tm", "Yb", "Lu", "",
            "", "", "", "Th", "", "U", "Np", "Pu", "Am", "Cm", "", "", "", "", "", "", "", ""
        ]

        # Define color schemes based on element categories
        colors = {
            'alkali_metal': "#FF6666",  # Light red
            'alkaline_earth': "#FFDEAD",  # Navajo white
            'transition_metal': "#FFC0CB",  # Pink
            'post_transition': "#CCCCCC",  # Light gray
            'metalloid': "#97FFFF",  # Light cyan
            'nonmetal': "#A0FFA0",  # Light green
            'halogen': "#FFFF99",  # Light yellow
            'noble_gas': "#C8A2C8",  # Lilac
            'lanthanide': "#FFBFFF",  # Light magenta
            'actinide': "#FF99CC",  # Light pink
            'unknown': "#E8E8E8"  # Very light gray
        }

        # Define element categories
        element_categories = {
            'H': 'nonmetal',
            'He': 'noble_gas',
            'Li': 'alkali_metal', 'Na': 'alkali_metal', 'K': 'alkali_metal',
            'Rb': 'alkali_metal', 'Cs': 'alkali_metal', 'Fr': 'alkali_metal',
            'Be': 'alkaline_earth', 'Mg': 'alkaline_earth', 'Ca': 'alkaline_earth',
            'Sr': 'alkaline_earth', 'Ba': 'alkaline_earth', 'Ra': 'alkaline_earth',
            'Sc': 'transition_metal', 'Ti': 'transition_metal', 'V': 'transition_metal',
            'Cr': 'transition_metal', 'Mn': 'transition_metal', 'Fe': 'transition_metal',
            'Co': 'transition_metal', 'Ni': 'transition_metal', 'Cu': 'transition_metal',
            'Zn': 'transition_metal', 'Y': 'transition_metal', 'Zr': 'transition_metal',
            'Nb': 'transition_metal', 'Mo': 'transition_metal', 'Tc': 'transition_metal',
            'Ru': 'transition_metal', 'Rh': 'transition_metal', 'Pd': 'transition_metal',
            'Ag': 'transition_metal', 'Cd': 'transition_metal', 'Hf': 'transition_metal',
            'Ta': 'transition_metal', 'W': 'transition_metal', 'Re': 'transition_metal',
            'Os': 'transition_metal', 'Ir': 'transition_metal', 'Pt': 'transition_metal',
            'Au': 'transition_metal', 'Hg': 'transition_metal',
            'B': 'metalloid', 'Si': 'metalloid', 'Ge': 'metalloid',
            'As': 'metalloid', 'Sb': 'metalloid', 'Te': 'metalloid', 'Po': 'metalloid',
            'C': 'nonmetal', 'N': 'nonmetal', 'O': 'nonmetal', 'P': 'nonmetal',
            'S': 'nonmetal', 'Se': 'nonmetal',
            'F': 'halogen', 'Cl': 'halogen', 'Br': 'halogen', 'I': 'halogen', 'At': 'halogen',
            'Ne': 'noble_gas', 'Ar': 'noble_gas', 'Kr': 'noble_gas',
            'Xe': 'noble_gas', 'Rn': 'noble_gas',
            'Al': 'post_transition', 'Ga': 'post_transition', 'In': 'post_transition',
            'Sn': 'post_transition', 'Tl': 'post_transition', 'Pb': 'post_transition',
            'Bi': 'post_transition',
        }

        # Add lanthanides and actinides
        lanthanides = ['La', 'Ce', 'Pr', 'Nd', 'Pm', 'Sm', 'Eu', 'Gd', 'Tb', 'Dy', 'Ho', 'Er', 'Tm', 'Yb', 'Lu']
        actinides = ['Ac', 'Th', 'Pa', 'U', 'Np', 'Pu', 'Am', 'Cm', 'Bk', 'Cf', 'Es', 'Fm', 'Md', 'No', 'Lr']

        for element in lanthanides:
            element_categories[element] = 'lanthanide'
        for element in actinides:
            element_categories[element] = 'actinide'

        button_size = (30, 30)
        font = wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)

        # Store original colors for restoration when buttons are deselected
        self.original_colors = {}

        # # Create context menu
        # self.create_context_menu()

        for element in elements:
            if element:
                # Get color based on element category
                category = element_categories.get(element, 'unknown')
                color = colors.get(category, colors['unknown'])

                btn = wx.Button(panel, label=element, size=button_size)
                btn.SetFont(font)
                btn.Bind(wx.EVT_ENTER_WINDOW, self.OnElementHover)
                btn.Bind(wx.EVT_LEAVE_WINDOW, self.OnElementLeave)
                btn.Bind(wx.EVT_BUTTON, self.OnElementClick)
                btn.Bind(wx.EVT_RIGHT_DOWN, self.on_element_right_click)
                btn.SetBackgroundColour(color)  # Set the initial color based on category
                self.original_colors[element] = color  # Store the original color
                self.button_states[element] = False
            else:
                btn = wx.StaticText(panel, label="")
            grid.Add(btn, 0, wx.EXPAND)

        periodic_sizer.Add(grid, 0, wx.EXPAND | wx.ALL, 0)
        hsizer.Add(periodic_sizer, 0, wx.EXPAND, 0)

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

        hsizer.Add(right_sizer, 0, wx.EXPAND)

        main_sizer.Add(hsizer, 1, wx.EXPAND)
        panel.SetSizer(main_sizer)
        self.SetSize(790, 320)

    # Add this import at the top of your file to import just the function
    from libraries.LibraryID import PeriodicTableXPS

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

    def get_element_transitions(self, element):
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

    def OnElementClick(self, event):
        element = event.GetEventObject().GetLabel()
        self.button_states[element] = not self.button_states[element]

        if self.button_states[element]:
            event.GetEventObject().SetBackgroundColour(wx.GREEN)
            self.plot_element_lines(element)
            transitions = self.get_element_transitions(element)

            # Add transitions to list without clearing existing items
            existing_items = [self.core_level_list.GetString(i) for i in range(self.core_level_list.GetCount())]
            for orbital, be in transitions:
                item = f"{element}{orbital}: {be:.1f} eV"
                if item not in existing_items:
                    self.core_level_list.Append(item)
        else:
            # Restore the original color when deselected
            original_color = self.original_colors.get(element, wx.WHITE)
            event.GetEventObject().SetBackgroundColour(original_color)
            self.remove_element_lines(element)

        event.GetEventObject().Refresh()

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
            formatted_label = ".."

            # Extract element and orbital correctly - using original logic
            import re
            match = re.match(r'([A-Z][a-z]*)(\d+[spdf])', element_orbital.strip())
            if match:
                element, orbital = match.groups()
                formatted_label = f"{element} {orbital[0]} {orbital[1:]}"  # e.g., "C 1 s"
            elif any(element_orbital.strip().endswith(x) for x in ['kll', 'mnn', 'mvv', 'mnv', 'lmm']):
                # Handle Auger lines - only specific ones allowed
                auger = element_orbital.strip()[-3:]
                formatted_label = f"{element_orbital.strip()[:-3]} {auger.upper()}"

            # Only proceed if we have a valid formatted label (preserves original restriction)
            if formatted_label != "..":
                try:
                    # Get max intensity in ±5 eV range
                    x_values = self.parent_window.x_values
                    y_values = self.parent_window.y_values
                    maxY = max(y_values)
                    mask = (x_values >= be - 5) & (x_values <= be + 5)

                    if np.any(mask):
                        local_max = np.max(y_values[mask])
                        label_y = local_max + 0.05 * maxY

                        # Add label at 1.2 times the local maximum height
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
                            'rotation': 90  # Force 90 degrees
                        })

                        print(f"Successfully added label: {formatted_label} at {be} eV")
                    else:
                        print(f"No data points found near {be} eV for {formatted_label}")

                except Exception as e:
                    print(f"Error adding label for {element_orbital}: {str(e)}")
            else:
                print(f"Skipping {element_orbital} - not a valid core level or allowed Auger line")

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
        print(f"transitions: {transitions}")

        if transitions:
            xmin, xmax = self.parent_window.ax.get_xlim()
            ymin, ymax = self.parent_window.ax.get_ylim()


            # Filter transitions within xmin and xmax
            valid_transitions = [t for t in transitions if xmax <= t[1] <= xmin]

            if valid_transitions:
                # Get RSF values for each transition
                rsf_values = self.get_rsf_values(element, [t[0] for t in valid_transitions])
                # print("RSF  " + str(rsf_values))

                if rsf_values:
                    max_rsf = max(rsf_values)

                    # Initialize the list for this element
                    self.element_lines[element] = []

                    for (orbital, be), rsf in zip(valid_transitions, rsf_values):
                        intensity = (rsf / max_rsf) * 0.6 * (ymax - ymin)
                        line = self.parent_window.ax.vlines(be, ymin - intensity, ymin + intensity, color='red',
                                                            linewidth=1)
                        self.element_lines[element].append(line)

                        # Add text label
                        text = self.parent_window.ax.text(
                            be + 0.1,  # Slightly to the left of the line
                            ymin+0.01*(ymax-ymin),  # At the bottom of the plot
                            element+""+orbital,
                            rotation=90,
                            va='bottom',
                            ha='right',
                            fontsize=7,
                            color='red'
                        )
                        self.element_lines[element].append(text)

                    self.parent_window.canvas.draw_idle()
                else:
                    print(f"No RSF values found for {element}")
            else:
                print(f"No valid transitions found for {element} in the current plot range")

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
        print(f"Searching RSF values for element: {element}")
        print(f"Orbitals to search for: {orbitals}")

        for (elem, orbital), data in self.library_data.items():
            if elem == element:
                orbital_lower = orbital.lower()
                if orbital_lower in [o.lower() for o in orbitals]:
                    # Check if 'Al' key exists, if not, use the first available instrument
                    instrument = 'Al1486' if 'Al1486' in data else next(iter(data))
                    print(data[instrument])
                    if 'rsf' in data[instrument]:
                        rsf_values.append(float(data[instrument]['rsf']))
                        print(f"Added RSF value: {data[instrument]['rsf']} for orbital {orbital}")

        print(f"Final RSF values: {rsf_values}")
        return rsf_values

    def OnElementHover(self, event):
        element = event.GetEventObject().GetLabel()
        self.UpdateElementInfo(element)

    def OnElementLeave(self, event):
        self.info_text1.SetLabelMarkup("")
        self.info_text2.SetLabelMarkup("")
        self.Layout()

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