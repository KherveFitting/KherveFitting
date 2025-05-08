import wx
import wx.richtext
import requests
import re
import os
import json


# Replace with your actual Google Form URL
GOOGLE_FORM_URL = 'https://docs.google.com/forms/u/0/d/1_PbjsaQGuhN5_K0-x90QIf75gcfCFiiv5snqju8nUBE/formResponse'

# Replace with your actual entry IDs
FORM_FIELDS = {
    'Firstname': 'entry.1651471049',
    'Surname': 'entry.853870358',
    'Email': 'entry.1419570696',
    'University/Company': 'entry.1822967727',
    'Country': 'entry.1299197141',
    'Usage': 'entry.431362163',
    'Supplier Name': 'entry.745225078',
    'Discovery': 'entry.968199044'  # Replace with actual field ID
}

# List of countries for dropdown
COUNTRIES = ["Afghanistan", "Albania", "Algeria", "Andorra", "Angola", "Antigua and Barbuda", "Argentina", "Armenia",
             "Australia", "Austria", "Azerbaijan", "Bahamas", "Bahrain", "Bangladesh", "Barbados", "Belarus", "Belgium",
             "Belize", "Benin", "Bhutan", "Bolivia", "Bosnia and Herzegovina", "Botswana", "Brazil", "Brunei",
             "Bulgaria", "Burkina Faso", "Burundi", "Cabo Verde", "Cambodia", "Cameroon", "Canada",
             "Central African Republic", "Chad", "Chile", "China", "Colombia", "Comoros", "Congo", "Costa Rica",
             "Croatia", "Cuba", "Cyprus", "Czech Republic", "Denmark", "Djibouti", "Dominica", "Dominican Republic",
             "Ecuador", "Egypt", "El Salvador", "Equatorial Guinea", "Eritrea", "Estonia", "Eswatini", "Ethiopia",
             "Fiji", "Finland", "France", "Gabon", "Gambia", "Georgia", "Germany", "Ghana", "Greece", "Grenada",
             "Guatemala", "Guinea", "Guinea-Bissau", "Guyana", "Haiti", "Honduras", "Hungary", "Iceland", "India",
             "Indonesia", "Iran", "Iraq", "Ireland", "Israel", "Italy", "Jamaica", "Japan", "Jordan", "Kazakhstan",
             "Kenya", "Kiribati", "Kosovo", "Kuwait", "Kyrgyzstan", "Laos", "Latvia", "Lebanon", "Lesotho", "Liberia",
             "Libya", "Liechtenstein", "Lithuania", "Luxembourg", "Madagascar", "Malawi", "Malaysia", "Maldives",
             "Mali", "Malta", "Marshall Islands", "Mauritania", "Mauritius", "Mexico", "Micronesia", "Moldova",
             "Monaco", "Mongolia", "Montenegro", "Morocco", "Mozambique", "Myanmar", "Namibia", "Nauru", "Nepal",
             "Netherlands", "New Zealand", "Nicaragua", "Niger", "Nigeria", "North Korea", "North Macedonia", "Norway",
             "Oman", "Pakistan", "Palau", "Palestine", "Panama", "Papua New Guinea", "Paraguay", "Peru", "Philippines",
             "Poland", "Portugal", "Qatar", "Romania", "Russia", "Rwanda", "Saint Kitts and Nevis", "Saint Lucia",
             "Saint Vincent and the Grenadines", "Samoa", "San Marino", "Sao Tome and Principe", "Saudi Arabia",
             "Senegal", "Serbia", "Seychelles", "Sierra Leone", "Singapore", "Slovakia", "Slovenia", "Solomon Islands",
             "Somalia", "South Africa", "South Korea", "South Sudan", "Spain", "Sri Lanka", "Sudan", "Suriname",
             "Sweden", "Switzerland", "Syria", "Taiwan", "Tajikistan", "Tanzania", "Thailand", "Timor-Leste", "Togo",
             "Tonga", "Trinidad and Tobago", "Tunisia", "Turkey", "Turkmenistan", "Tuvalu", "Uganda", "Ukraine",
             "United Arab Emirates", "United Kingdom", "United States", "Uruguay", "Uzbekistan", "Vanuatu",
             "Vatican City", "Venezuela", "Vietnam", "Yemen", "Zambia", "Zimbabwe"]

USAGES = ["Multiple times in a day", "Daily", "Weekly", "Monthly", "Yearly", "Rarely / One off"]


class RegistrationForm(wx.Frame):
    def __init__(self, parent):
        custom_style = wx.CAPTION | wx.SYSTEM_MENU | wx.RESIZE_BORDER | wx.STAY_ON_TOP

        super().__init__(parent=parent,
                         title="KherveFitting Questionnaire",
                         size=(550, 650),
                         style=custom_style)

        self.parent = parent
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)


        # Description text
        # description = wx.StaticText(panel,
        #                             label="KherveFitting is an open-source software developed for XPS data analysis. This registration helps us track the distribution and size of the research community using this tool. The data collected will only be used for statistical purposes and to "
        #                                   "support continued development of scientific software for the XPS community. "
        #                                   "No personal information will be shared with third parties.")
        # description.Wrap(530)  # Wrap text to fit the panel
        # vbox.Add(description, flag=wx.ALL | wx.EXPAND, border=10)
        # Create a RichTextCtrl instead of StaticText
        description = wx.richtext.RichTextCtrl(
            panel,
            style=wx.BORDER_NONE | wx.richtext.RE_READONLY
        )

        # Set the text with proper formatting
        description.SetValue(
            "KherveFitting is an open-source / Free software developed for XPS and now Raman data analysis. "
            "This registration will helps me track the distribution and size of the research "
            "community using this tool. The data collected will only be used for statistical "
            "purposes and to support continued development of KherveFitting. No personal information will be shared "
            "with third parties. If you have any queries, please use the Bug Report function in the Help menu or "
            "contact me directly at  g.kerherve@ic.ac.uk\n\n"
            
            "NOTE: At the end of the submission, the application will have to be restarted. "
        )

        # Make the background match the panel
        description.SetBackgroundColour(panel.GetBackgroundColour())

        # Disable the control to make it look more like static text
        description.SetEditable(False)

        # Set a good height for the control (adjust as needed)
        description.SetMinSize((-1, 145))

        # Add to your sizer with some padding
        vbox.Add(description, flag=wx.ALL | wx.EXPAND, border=10)

        # Add a line separator
        line = wx.StaticLine(panel)
        vbox.Add(line, flag=wx.EXPAND | wx.ALL, border=5)

        # Create form fields
        self.fields = {}
        field_width = 200  # Standard width for all input fields

        # Name field
        self.add_text_field(panel, vbox, "Firstname", field_width)

        # Name field
        self.add_text_field(panel, vbox, "Surname", field_width)

        # Email field with validation
        self.add_text_field(panel, vbox, "Email", field_width)

        # University field
        self.add_text_field(panel, vbox, "University/Company", field_width)

        # Country dropdown
        country_label = wx.StaticText(panel, label="Country:")
        self.country_combo = wx.ComboBox(panel, choices=COUNTRIES, style=wx.CB_READONLY, size=(field_width, -1))

        hbox = wx.BoxSizer(wx.HORIZONTAL)
        hbox.Add(country_label, proportion=0, flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)

        # Add a spacer for alignment
        label_width = country_label.GetSize().GetWidth()
        spacer_width = max(15, 130 - label_width)
        hbox.Add((spacer_width, 1))

        hbox.Add(self.country_combo, proportion=0, flag=wx.ALL, border=5)
        vbox.Add(hbox, flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=5)
        self.fields["Country"] = self.country_combo

        # Usage dropdown
        Usage_label = wx.StaticText(panel, label="Usage:")
        self.Usage_combo = wx.ComboBox(panel, choices=USAGES, style=wx.CB_READONLY, size=(field_width, -1))

        hbox = wx.BoxSizer(wx.HORIZONTAL)
        hbox.Add(Usage_label, proportion=0, flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)

        # Add a spacer for alignment
        label_width = Usage_label.GetSize().GetWidth()
        spacer_width = max(15, 130 - label_width)
        hbox.Add((spacer_width, 1))

        hbox.Add(self.Usage_combo, proportion=0, flag=wx.ALL, border=5)
        vbox.Add(hbox, flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=5)
        self.fields["Usage"] = self.Usage_combo

        # Add a line separator before the Purpose field
        line2 = wx.StaticLine(panel)
        vbox.Add(line2, flag=wx.EXPAND | wx.ALL, border=5)

        # purpose field - multiline
        purpose_hbox = wx.BoxSizer(wx.VERTICAL)
        purpose_label = wx.StaticText(panel, label="Please provide the name of the supplier and the model of your XPS system")
        purpose_text = wx.TextCtrl(panel, style=wx.TE_MULTILINE, size=(field_width, 40))

        purpose_hbox.Add(purpose_label, flag=wx.ALL, border=5)
        purpose_hbox.Add(purpose_text, flag=wx.ALL | wx.EXPAND, border=5)
        vbox.Add(purpose_hbox, flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=10)
        self.fields["Supplier Name"] = purpose_text

        # Discovery field - multiline
        discovery_hbox = wx.BoxSizer(wx.VERTICAL)
        discovery_label = wx.StaticText(panel, label="How did you find out about KherveFitting?\n(LinkedIn / Internet / Colleagues / Friends / Others)")
        discovery_text = wx.TextCtrl(panel, style=wx.TE_MULTILINE, size=(field_width, 40))

        discovery_hbox.Add(discovery_label, flag=wx.ALL, border=5)
        discovery_hbox.Add(discovery_text, flag=wx.ALL | wx.EXPAND, border=5)
        vbox.Add(discovery_hbox, flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=10)
        self.fields["Discovery"] = discovery_text

        # Submit button - with standard size (110, 40)
        submit_btn = wx.Button(panel, label="Submit")
        submit_btn.SetInitialSize(wx.Size(110, 40))
        submit_btn.SetMinSize(wx.Size(110, 40))
        submit_btn.SetMaxSize(wx.Size(110, 40))
        vbox.Add(submit_btn, flag=wx.ALL | wx.CENTER, border=15)

        submit_btn.Bind(wx.EVT_BUTTON, self.on_submit)

        panel.SetSizer(vbox)
        self.Centre()
        self.Show()

    def add_text_field(self, panel, sizer, field_name, field_width):
        hbox = wx.BoxSizer(wx.HORIZONTAL)
        label = wx.StaticText(panel, label=field_name + ":")
        text_ctrl = wx.TextCtrl(panel, size=(field_width, -1))

        hbox.Add(label, proportion=0, flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)

        # Add a spacer for better alignment
        label_width = label.GetSize().GetWidth()
        spacer_width = max(15, 130 - label_width)
        hbox.Add((spacer_width, 1))

        hbox.Add(text_ctrl, proportion=0, flag=wx.ALL, border=5)
        sizer.Add(hbox, flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=5)

        self.fields[field_name] = text_ctrl

    def validate_email(self, email):
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None


    def on_submit(self, event):
        data = {}
        for label, control in self.fields.items():
            value = control.GetValue()
            if not value:
                wx.MessageBox(f"Please enter your {label}.", "Missing Information", wx.OK | wx.ICON_WARNING)
                return

            # Validate email
            if label == "Email" and not self.validate_email(value):
                wx.MessageBox("Please enter a valid email address.", "Invalid Email", wx.OK | wx.ICON_WARNING)
                return

            data[FORM_FIELDS[label]] = value

        try:
            response = requests.post(GOOGLE_FORM_URL, data=data)
            if response.status_code == 200:
                # Save registration state and restart application
                self.save_registration_state()
            else:
                wx.MessageBox(f"Failed to submit form. Status code: {response.status_code}", "Error",
                              wx.OK | wx.ICON_ERROR)
        except Exception as e:
            wx.MessageBox(f"An error occurred:\n{e}", "Error", wx.OK | wx.ICON_ERROR)


    def restart_application(self):
        """Restart the entire application"""
        import os
        import sys

        # Close the registration form and any parent windows
        self.Close()

        # Restart the program
        python = sys.executable
        os.execl(python, python, *sys.argv)

    def save_registration_state(self):
        """Save registration state to config file"""
        # Use the same config file as the main application
        import os
        import sys
        import json

        config_file = 'config.json'
        config = {}

        if os.path.exists(config_file):
            try:
                with open(config_file, 'r') as f:
                    config = json.load(f)
            except:
                config = {}

        # Set registered flag
        config['registered'] = True

        # Save config
        with open(config_file, 'w') as f:
            json.dump(config, f, indent=2)

        # Show success message
        wx.MessageBox("Thank you for registering!\nPlease Restart the application", "Completed",
                      wx.OK | wx.ICON_INFORMATION)

        # Now restart the entire application
        self.restart_application()


def check_registration_needed():
    """Check if registration is needed by looking at config file"""
    config_file = 'config.json'

    # Default to True (registration needed) if file doesn't exist
    if not os.path.exists(config_file):
        return True

    try:
        with open(config_file, 'r') as f:
            config = json.load(f)
            # If registered key exists and is True, no registration needed
            return not config.get('registered', False)
    except:
        # If error reading file, default to registration needed
        return True


def show_registration_form():
    """Show the registration form"""
    app = wx.App(False)
    frame = RegistrationForm(None)
    app.MainLoop()
    # return True  # Registration complete



# Function to be called on first run or from Help menu
def launch_registration_form(parent=None):
    """Launch the registration form, optionally as modal if parent is provided"""
    if parent:
        # If called from Help menu with a parent window
        dialog = RegistrationForm()
        dialog.Show()
    else:
        # If called on first run
        show_registration_form()


if __name__ == "__main__":
    show_registration_form()