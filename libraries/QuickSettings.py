import wx
import os


class QuickSettings:
    def __init__(self, parent):
        self.parent = parent
        # Quick settings options
        self.quick_options = [
            ("Al Kα + A-ALTHERMO1", 1486.67, "A-ALTHERMO1"),
            ("Al Kα + C-KratosF1s-Al1486", 1486.67, "C-KratosF1s-Al1486"),
            ("Al Kα + C-KratosC1s-Al1486", 1486.67, "C-KratosC1s-Al1486"),
            ("Ag Lα + C-KratosC1s-Al1486", 2984.3, "C-KratosC1s-Al1486"),
            ("Ga Kα + O-ScientaGCs-Ga", 9251, "O-ScientaGCs-Ga")
        ]

    def create_quick_settings_tool(self, toolbar):
        """Create quick settings dropdown tool in toolbar"""
        if not getattr(self.parent, 'enable_quick_settings', False):
            return None

        current_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(current_dir, "Icons")
        if not os.path.exists(icon_path):
            icon_path = os.path.join(os.path.dirname(current_dir), "Icons")

        # Create quick settings button
        quick_settings_tool = toolbar.AddTool(
            wx.ID_ANY,
            'Quick Settings',
            wx.Bitmap(os.path.join(icon_path, "QuickSettings-3.png"), wx.BITMAP_TYPE_PNG),
            shortHelp="Quick Setting for X-ray Source & RSF Library"
        )

        # Bind to show menu
        self.parent.Bind(wx.EVT_TOOL, self.on_quick_settings_click, quick_settings_tool)
        return quick_settings_tool

    def on_quick_settings_click(self, event):
        """Show quick settings popup menu"""
        menu = wx.Menu()

        for i, (label, photon_energy, instrument) in enumerate(self.quick_options):
            item = menu.Append(wx.ID_ANY, label)
            self.parent.Bind(wx.EVT_MENU,
                             lambda evt, pe=photon_energy, inst=instrument: self.apply_quick_setting(pe, inst),
                             item)

        # Show menu at toolbar position
        toolbar = event.GetEventObject()
        toolbar.PopupMenu(menu)
        menu.Destroy()

    def apply_quick_setting(self, photon_energy, instrument):
        """Apply the selected quick setting"""
        # Create brief console window
        parent_pos = self.parent.GetPosition()
        parent_size = self.parent.GetSize()
        console_frame = wx.Frame(self.parent, title="Quick Settings", size=(300, 150))
        console_frame.SetPosition((
            parent_pos.x + (parent_size.width - 300) // 2,
            parent_pos.y + (parent_size.height - 150) // 2
        ))
        console_text = wx.TextCtrl(console_frame, style=wx.TE_MULTILINE | wx.TE_READONLY)
        console_frame.Show()

        def update_console(message):
            console_text.AppendText(message + '\n')
            console_text.Update()
            wx.SafeYield()

        update_console("Applying quick settings...")

        # Update photon energy
        self.parent.photons = photon_energy
        update_console(f"X-ray source: {photon_energy} eV")

        # Update instrument
        self.parent.current_instrument = instrument
        update_console(f"Instrument: {instrument}")

        # Save config
        self.parent.save_config()
        update_console("Configuration saved")

        # Update any open preference windows
        self.update_preference_windows(photon_energy, instrument)
        update_console("Settings applied successfully!")

        # Close console after 1 second
        wx.CallLater(1000, console_frame.Close)

    def update_preference_windows(self, photon_energy, instrument):
        """Update any open preference windows with new settings"""
        from libraries.PreferenceWindow import PreferenceWindow

        for window in wx.GetTopLevelWindows():
            if isinstance(window, PreferenceWindow):
                window.custom_photon.SetValue(photon_energy)
                window.instrument_combo.SetValue(instrument)

                # Update photon source dropdown based on energy
                if photon_energy == 1486.67:
                    window.photon_combo.SetSelection(0)  # Al Kα
                elif photon_energy == 1253.6:
                    window.photon_combo.SetSelection(1)  # Mg Kα
                elif photon_energy == 2984.3:
                    window.photon_combo.SetSelection(2)  # Ag Lα
                elif photon_energy == 9251:
                    window.photon_combo.SetSelection(3)  # Ga Kα
                elif photon_energy == 5417:
                    window.photon_combo.SetSelection(4)  # Cr Kα
                else:
                    window.photon_combo.SetSelection(5)  # Custom