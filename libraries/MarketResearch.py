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

DISCOVERY_OPTIONS = ["LinkedIn", "Internet Search", "Conference", "Colleagues", "Friends", "Publication", "Other"]


# Dictionary mapping countries to their universities
DEFAULT_UNIVERSITIES = {
    "United States": [
        "Massachusetts Institute of Technology (MIT)", "Harvard University", "Stanford University",
        "California Institute of Technology (Caltech)", "Princeton University", "Yale University",
        "University of Chicago", "Columbia University", "University of Pennsylvania",
        "Johns Hopkins University", "University of California, Berkeley", "Cornell University",
        "University of California, Los Angeles (UCLA)", "Duke University", "University of Michigan",
        "Northwestern University", "New York University (NYU)", "Carnegie Mellon University",
        "University of California, San Diego", "University of Washington", "University of Illinois at Urbana-Champaign",
        "University of Wisconsin-Madison", "Washington University in St. Louis", "University of Texas at Austin",
        "Georgia Institute of Technology", "University of North Carolina at Chapel Hill", "Boston University",
        "Emory University", "University of California, Davis", "University of California, Santa Barbara",
        "University of Florida", "University of Minnesota", "University of Maryland, College Park",
        "University of Southern California", "Ohio State University", "Purdue University",
        "Pennsylvania State University", "University of Pittsburgh", "Rice University",
        "University of Rochester", "University of California, Irvine", "Vanderbilt University",
        "Dartmouth College", "University of Virginia", "University of Colorado Boulder",
        "University of Arizona", "Michigan State University", "Brown University",
        "University of California, Santa Cruz", "Arizona State University", "Rutgers University",
        "University of Miami", "University of Notre Dame", "Case Western Reserve University",
        "University of Illinois Chicago", "Georgetown University", "Tufts University",
        "University of Massachusetts Amherst", "University of California, Riverside",
        "University of Connecticut", "University of Hawaii at Manoa", "University of Iowa",
        "North Carolina State University", "University of Tennessee, Knoxville",
        "Florida State University", "Virginia Tech", "University of Delaware",
        "Indiana University Bloomington", "Texas A&M University", "University of Oregon",
        "University of Vermont", "University of Alabama at Birmingham", "Northeastern University",
        "Colorado State University", "University of Kansas", "Oregon State University",
        "University of New Mexico", "Syracuse University", "Stony Brook University",
        "University of Utah", "University of South Florida", "Iowa State University",
        "University of Oklahoma", "University of Missouri", "University at Buffalo",
        "University of Kentucky", "University of Georgia", "University of Nebraska-Lincoln",
        "University of Arkansas", "University of Cincinnati", "Brandeis University",
        "University of Houston", "Tulane University", "Auburn University",
        "Boston College", "George Washington University", "Lehigh University",
        "University of Nevada, Reno", "University of South Carolina", "Wayne State University",
        "Brigham Young University", "University of Texas Dallas", "Rensselaer Polytechnic Institute",
        "Stevens Institute of Technology", "University of Denver", "Other"
    ],
    "Senegal": [
        "Université Cheikh Anta Diop de Dakar", "Université Gaston Berger de Saint-Louis",
        "Université Alioune Diop de Bambey", "Université Assane Seck de Ziguinchor",
        "Université virtuelle du Sénégal", "Université du Sahel", "École Polytechnique de Thiès",
        "École Supérieure Polytechnique de Dakar", "Institut Africain de Management",
        "CESAG Business School", "Université Amadou Hampâté Bâ", "Institut Supérieur de Management",
        "École Nationale d'Administration", "École Nationale d'Économie Appliquée",
        "École Supérieure Multinationale des Télécommunications",
        "Institut Supérieur d'Entrepreneurship et de Gestion", "Suffolk University Dakar Campus",
        "Université de l'Entreprise", "École Nationale des Arts et Métiers de Dakar",
        "École Supérieure de Commerce de Dakar", "Bordeaux Management School Dakar",
        "Institut Supérieur des Sciences de l'Information et de la Communication",
        "Institut des Sciences de la Terre", "Centre Africain d'Études Supérieures en Gestion",
        "Other"
    ],
    "United Kingdom": [
        "University of Oxford", "University of Cambridge", "Imperial College London",
        "University College London", "London School of Economics and Political Science",
        "University of Edinburgh", "King's College London", "University of Manchester",
        "University of Bristol", "University of Warwick", "University of Glasgow",
        "Durham University", "University of Sheffield", "University of Birmingham",
        "University of Southampton", "University of Leeds", "University of Nottingham",
        "Queen Mary University of London", "Lancaster University", "University of York",
        "University of St Andrews", "University of Exeter", "University of Bath",
        "University of Liverpool", "Cardiff University", "Newcastle University",
        "University of Aberdeen", "Queen's University Belfast", "University of Sussex",
        "University of Reading", "University of Leicester", "University of East Anglia",
        "Heriot-Watt University", "Loughborough University", "Royal Holloway, University of London",
        "University of Surrey", "University of Strathclyde", "Swansea University",
        "University of Dundee", "Brunel University London", "Birkbeck, University of London",
        "SOAS University of London", "City, University of London", "Goldsmiths, University of London",
        "Aston University", "Keele University", "University of Kent", "University of Essex",
        "Oxford Brookes University", "University of the Arts London", "Manchester Metropolitan University",
        "University of Plymouth", "Coventry University", "University of Portsmouth",
        "Nottingham Trent University", "University of Westminster", "University of Brighton",
        "Liverpool John Moores University", "University of Central Lancashire", "University of Huddersfield",
        "University of Lincoln", "Edinburgh Napier University", "University of Northumbria at Newcastle",
        "Other"
    ],
    "Canada": [
        "University of Toronto", "McGill University", "University of British Columbia",
        "University of Alberta", "McMaster University", "University of Waterloo",
        "University of Montreal", "University of Calgary", "Western University",
        "Queen's University", "University of Ottawa", "Dalhousie University",
        "Simon Fraser University", "University of Victoria", "York University", "Other"
    ],
    "Australia": [
        "University of Melbourne", "Australian National University", "University of Sydney",
        "University of Queensland", "University of New South Wales (UNSW)", "Monash University",
        "University of Western Australia", "University of Adelaide", "University of Technology Sydney",
        "Macquarie University", "RMIT University", "Curtin University", "Deakin University",
        "Queensland University of Technology", "University of Newcastle", "Other"
    ],
    "Germany": [
        "Technical University of Munich", "Ludwig Maximilian University of Munich",
        "Heidelberg University", "Humboldt University of Berlin", "Freie Universität Berlin",
        "RWTH Aachen University", "Technical University of Berlin", "Charité - Universitätsmedizin Berlin",
        "University of Tübingen", "University of Bonn", "University of Freiburg",
        "University of Hamburg", "University of Göttingen", "University of Cologne",
        "Karlsruhe Institute of Technology", "University of Mannheim", "University of Stuttgart",
        "University of Münster", "University of Erlangen-Nuremberg", "University of Frankfurt",
        "University of Jena", "University of Kiel", "Technical University of Darmstadt",
        "University of Konstanz", "University of Duisburg-Essen", "University of Ulm",
        "University of Mainz", "University of Bayreuth", "University of Bremen",
        "University of Bochum", "Technical University of Dresden", "University of Würzburg",
        "University of Potsdam", "University of Marburg", "University of Regensburg",
        "University of Giessen", "Technical University of Braunschweig", "Hannover Medical School",
        "University of Halle-Wittenberg", "University of Greifswald", "University of Kassel",
        "University of Rostock", "University of Leipzig", "University of Hohenheim",
        "University of Osnabrück", "University of Siegen", "University of Magdeburg",
        "University of Saarland", "University of Bamberg", "University of Passau",
        "University of Lübeck", "Jacobs University Bremen", "Technical University of Hamburg",
        "Technical University of Kaiserslautern", "University of Applied Sciences Munich",
        "Zeppelin University", "Frankfurt School of Finance and Management", "WHU – Otto Beisheim School of Management",
        "Other"
    ],
    "France": [
        "Paris Sciences et Lettres – PSL Research University", "Sorbonne University",
        "Institut Polytechnique de Paris", "Université Paris-Saclay", "École Polytechnique",
        "École Normale Supérieure, Paris", "Sciences Po Paris", "École des Ponts ParisTech",
        "CentraleSupélec", "Université Grenoble Alpes", "Université de Paris",
        "Université de Strasbourg", "Aix-Marseille Université", "Université Claude Bernard Lyon 1",
        "École Normale Supérieure de Lyon", "Université de Bordeaux", "École Centrale de Lyon",
        "Université Paris 1 Panthéon-Sorbonne", "Université de Montpellier",
        "Université de Lorraine", "Université Toulouse III - Paul Sabatier", "Université de Lille",
        "Université de Rennes 1", "Université Paris-Est Créteil Val de Marne",
        "Université de Versailles Saint-Quentin-en-Yvelines", "Université de Nantes",
        "Université de Nice Sophia Antipolis", "Université Paris-Nanterre",
        "École des Hautes Études en Sciences Sociales", "Institut National des Sciences Appliquées de Lyon",
        "Université Lumière Lyon 2", "Université Toulouse 1 Capitole", "Université Paris Diderot",
        "Université Jean Moulin Lyon 3", "Université de Poitiers", "INSA Toulouse",
        "Toulouse Business School", "Université de Bourgogne", "Université de Caen Normandie",
        "Université d'Angers", "Université de Cergy-Pontoise", "EDHEC Business School",
        "Université de Franche-Comté", "KEDGE Business School", "Université du Havre",
        "Université de La Rochelle", "ESC Rennes School of Business", "Université de Reims Champagne-Ardenne",
        "EM Lyon Business School", "Université Jean Monnet", "Other"
    ],

    "Japan": [
        "University of Tokyo", "Kyoto University", "Tohoku University", "Tokyo Institute of Technology",
        "Osaka University", "Nagoya University", "Hokkaido University", "Kyushu University",
        "Waseda University", "Keio University", "Hiroshima University", "Tsukuba University",
        "Kobe University", "Okayama University", "Chiba University", "Other"
    ],
    "China": [
        "Tsinghua University", "Peking University", "Fudan University", "Shanghai Jiao Tong University",
        "Zhejiang University", "University of Science and Technology of China",
        "Nanjing University", "Wuhan University", "Sun Yat-sen University",
        "Beijing Normal University", "Tongji University", "Harbin Institute of Technology",
        "Huazhong University of Science and Technology", "Xi'an Jiaotong University",
        "Nankai University", "East China Normal University", "Other"
    ],
    "South Korea": [
        "Seoul National University", "KAIST", "Pohang University of Science and Technology (POSTECH)",
        "Korea University", "Sungkyunkwan University", "Yonsei University", "Hanyang University",
        "Kyunghee University", "Ewha Womans University", "Pusan National University", "Other"
    ],
    "Singapore": [
        "National University of Singapore", "Nanyang Technological University",
        "Singapore Management University", "Singapore University of Technology and Design", "Other"
    ],
    "Netherlands": [
        "Delft University of Technology", "University of Amsterdam", "Eindhoven University of Technology",
        "Utrecht University", "Leiden University", "University of Groningen", "Wageningen University",
        "Erasmus University Rotterdam", "Maastricht University", "Vrije Universiteit Amsterdam", "Other"
    ],
    "Switzerland": [
        "ETH Zurich", "École Polytechnique Fédérale de Lausanne (EPFL)", "University of Zurich",
        "University of Geneva", "University of Basel", "University of Bern", "University of Lausanne", "Other"
    ],
    "Italy": [
        "Politecnico di Milano", "University of Bologna", "Sapienza University of Rome",
        "University of Padua", "University of Milan", "University of Naples Federico II",
        "University of Turin", "University of Pisa", "University of Florence", "Politecnico di Torino", "Other"
    ],
    "Spain": [
        "University of Barcelona", "Autonomous University of Madrid", "Autonomous University of Barcelona",
        "Complutense University of Madrid", "Pompeu Fabra University", "University of Navarra",
        "Universitat Politècnica de Catalunya", "University of Valencia", "University of Granada",
        "Universidad Carlos III de Madrid", "University of the Basque Country", "University of Zaragoza",
        "Polytechnic University of Valencia", "University of Salamanca", "University of Seville",
        "University of Oviedo", "University of Santiago de Compostela", "Universidad de Deusto",
        "University of Malaga", "University of Alicante", "University of Murcia",
        "University of Vigo", "University of La Laguna", "University of Valladolid",
        "University of Alcalá", "Universidad Politécnica de Madrid", "University of Castilla–La Mancha",
        "University of Rovira i Virgili", "University of La Coruña", "University of Las Palmas de Gran Canaria",
        "University of Jaén", "University of Cadiz", "University of León", "University of Cantabria",
        "University of Girona", "University of Córdoba", "University of Almería",
        "University of the Balearic Islands", "University of Extremadura", "ESADE Business School",
        "IE Business School", "IESE Business School", "University of Lleida",
        "Jaume I University", "Miguel Hernández University of Elche", "University of Huelva",
        "University of La Rioja", "University of Burgos", "Pablo de Olavide University",
        "Public University of Navarre", "Other"
    ],
    "Sweden": [
        "Karolinska Institute", "Lund University", "Uppsala University", "KTH Royal Institute of Technology",
        "Stockholm University", "Chalmers University of Technology", "University of Gothenburg", "Other"
    ],
    "Denmark": [
        "University of Copenhagen", "Technical University of Denmark", "Aarhus University",
        "Aalborg University", "University of Southern Denmark", "Copenhagen Business School", "Other"
    ],
    "Belgium": [
        "KU Leuven", "Ghent University", "Université Catholique de Louvain",
        "Vrije Universiteit Brussel", "University of Liège", "University of Antwerp", "Other"
    ],
    "Russia": [
        "Lomonosov Moscow State University", "Saint Petersburg State University",
        "Novosibirsk State University", "Moscow Institute of Physics and Technology",
        "HSE University", "Tomsk State University", "ITMO University", "Other"
    ],
    "Brazil": [
        "University of São Paulo", "University of Campinas", "Federal University of Rio de Janeiro",
        "São Paulo State University", "Federal University of Minas Gerais",
        "Federal University of Rio Grande do Sul", "Federal University of São Paulo", "Other"
    ],
    "India": [
        "Indian Institute of Science", "Indian Institute of Technology Bombay",
        "Indian Institute of Technology Delhi", "Indian Institute of Technology Madras",
        "Indian Institute of Technology Kanpur", "Indian Institute of Technology Kharagpur",
        "Indian Institute of Technology Roorkee", "Indian Institute of Technology Guwahati",
        "Indian Institute of Technology Hyderabad", "Indian Institute of Technology Gandhinagar",
        "Indian Institute of Technology Indore", "Indian Institute of Technology Mandi",
        "Indian Institute of Technology (Banaras Hindu University) Varanasi",
        "Indian Institute of Technology Ropar", "Indian Institute of Technology Patna",
        "Indian Institute of Technology Bhubaneswar", "Indian Institute of Technology Jodhpur",
        "Indian Institute of Technology Dhanbad", "Indian Institute of Technology Tirupati",
        "Indian Institute of Technology Bhilai", "Indian Institute of Technology Goa",
        "Indian Institute of Technology Dharwad", "Indian Institute of Technology Palakkad",
        "Indian Institute of Technology Jammu", "Jawaharlal Nehru University",
        "Delhi University", "Anna University", "Jadavpur University", "University of Hyderabad",
        "Banaras Hindu University", "Aligarh Muslim University", "Savitribai Phule Pune University",
        "Amrita Vishwa Vidyapeetham", "Jamia Millia Islamia", "Birla Institute of Technology and Science, Pilani",
        "Calcutta University", "Indian Institute of Science Education and Research Pune",
        "National Institute of Technology Tiruchirappalli", "Indian Institute of Science Education and Research Kolkata",
        "Indian Institute of Science Education and Research Mohali", "Vellore Institute of Technology",
        "National Institute of Technology Karnataka", "National Institute of Technology Rourkela",
        "Indian Statistical Institute", "SRM Institute of Science and Technology",
        "Indian Institute of Science Education and Research Bhopal", "National Institute of Technology Warangal",
        "Indian Institute of Technology (ISM) Dhanbad", "Thapar Institute of Engineering and Technology",
        "Tata Institute of Fundamental Research", "National Institute of Technology Surathkal",
        "Amity University", "Tezpur University", "National Institute of Technology Durgapur",
        "Indian Institute of Foreign Trade", "Indian Institute of Management Ahmedabad",
        "Indian Institute of Management Bangalore", "Indian Institute of Management Calcutta",
        "Indian Institute of Management Lucknow", "Indian Institute of Management Indore",
        "Indian Institute of Management Kozhikode", "XLRI - Xavier School of Management",
        "Indian Institute of Management Udaipur", "Indian School of Business",
        "Indian Institute of Management Tiruchirappalli", "Indian Institute of Management Ranchi",
        "Indian Institute of Management Raipur", "Indian Institute of Management Rohtak",
        "Indian Institute of Management Kashipur", "Indian Institute of Management Nagpur",
        "Indian Institute of Management Amritsar", "Indian Institute of Management Bodh Gaya",
        "Indian Institute of Management Jammu", "Indian Institute of Management Sambalpur",
        "Indian Institute of Management Sirmaur", "Indian Institute of Management Visakhapatnam",
        "Other"
    ],
    "Israel": [
        "Weizmann Institute of Science", "Hebrew University of Jerusalem", "Tel Aviv University",
        "Technion – Israel Institute of Technology", "Ben-Gurion University of the Negev",
        "Bar-Ilan University", "University of Haifa", "Other"
    ],
    "South Africa": [
        "University of Cape Town", "University of the Witwatersrand", "Stellenbosch University",
        "University of Pretoria", "University of Johannesburg", "University of KwaZulu-Natal", "Other"
    ],
    "New Zealand": [
        "University of Auckland", "University of Otago", "University of Canterbury",
        "Victoria University of Wellington", "University of Waikato", "Massey University", "Other"
    ],
    "Mexico": [
        "National Autonomous University of Mexico", "Tecnológico de Monterrey",
        "Instituto Politécnico Nacional", "Universidad Iberoamericana", "Other"
    ],
    "Argentina": [
        "University of Buenos Aires", "National University of La Plata",
        "National University of Córdoba", "Austral University", "Other"
    ],
    "Chile": [
        "Pontifical Catholic University of Chile", "University of Chile",
        "Federico Santa María Technical University", "University of Concepción", "Other"
    ],
    "Egypt": [
        "Cairo University", "American University in Cairo", "Alexandria University",
        "Ain Shams University", "Mansoura University", "Other"
    ],
    "Saudi Arabia": [
        "King Abdulaziz University", "King Fahd University of Petroleum & Minerals",
        "King Saud University", "King Abdullah University of Science and Technology", "Other"
    ],
    "United Arab Emirates": [
        "Khalifa University", "United Arab Emirates University",
        "American University of Sharjah", "University of Sharjah", "Other"
    ],
    "Turkey": [
        "Koç University", "Sabancı University", "Bilkent University",
        "Middle East Technical University", "Boğaziçi University", "Istanbul Technical University", "Other"
    ],
    "Malaysia": [
        "University of Malaya", "Universiti Putra Malaysia", "Universiti Sains Malaysia",
        "Universiti Kebangsaan Malaysia", "Universiti Teknologi Malaysia", "Other"
    ],
    "Indonesia": [
        "University of Indonesia", "Bandung Institute of Technology", "Gadjah Mada University",
        "Airlangga University", "Bogor Agricultural University", "Other"
    ],
    "Thailand": [
        "Chulalongkorn University", "Mahidol University", "King Mongkut's University of Technology Thonburi",
        "Chiang Mai University", "Kasetsart University", "Other"
    ],
    "Vietnam": [
        "Vietnam National University, Hanoi", "Vietnam National University, Ho Chi Minh City",
        "Hanoi University of Science and Technology", "Can Tho University", "Other"
    ]
}

# For any country not listed above, provide a generic entry
for country in COUNTRIES:
    if country not in DEFAULT_UNIVERSITIES:
        DEFAULT_UNIVERSITIES[country] = ["University (Please specify)", "Research Institute", "Company", "Other"]




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

        # # University field
        # self.add_text_field(panel, vbox, "University/Company", field_width)

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

        # Bind country selection event
        self.country_combo.Bind(wx.EVT_COMBOBOX, self.on_country_selected)

        # University/Company combobox with status text
        university_label = wx.StaticText(panel, label="University/Company:")
        self.university_combo = wx.ComboBox(panel, style=wx.CB_DROPDOWN, size=(field_width, -1))

        uni_hbox = wx.BoxSizer(wx.HORIZONTAL)
        uni_hbox.Add(university_label, proportion=0, flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)

        # Add a spacer for alignment
        uni_label_width = university_label.GetSize().GetWidth()
        uni_spacer_width = max(15, 130 - uni_label_width)
        uni_hbox.Add((uni_spacer_width, 1))

        uni_hbox.Add(self.university_combo, proportion=0, flag=wx.ALL, border=5)
        vbox.Add(uni_hbox, flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=5)
        self.fields["University/Company"] = self.university_combo

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
        purpose_label = wx.StaticText(panel,
                                      label="Please provide the name of the supplier and the model of your XPS system")
        purpose_text = wx.TextCtrl(panel, style=wx.TE_MULTILINE, size=(field_width, 40))

        purpose_hbox.Add(purpose_label, flag=wx.ALL, border=5)
        purpose_hbox.Add(purpose_text, flag=wx.ALL | wx.EXPAND, border=5)
        vbox.Add(purpose_hbox, flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=10)
        self.fields["Supplier Name"] = purpose_text

        # Discovery field - now as combobox
        discovery_hbox = wx.BoxSizer(wx.HORIZONTAL)
        discovery_label = wx.StaticText(panel, label="How did you find out about KherveFitting?")
        self.discovery_combo = wx.ComboBox(panel, choices=DISCOVERY_OPTIONS, style=wx.CB_DROPDOWN,
                                           size=(field_width, -1))

        discovery_hbox.Add(discovery_label, proportion=0, flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL, border=5)

        # Add a spacer for alignment
        disc_label_width = discovery_label.GetSize().GetWidth()
        disc_spacer_width = max(15, 130 - disc_label_width)
        discovery_hbox.Add((disc_spacer_width, 1))

        discovery_hbox.Add(self.discovery_combo, proportion=0, flag=wx.ALL, border=5)
        vbox.Add(discovery_hbox, flag=wx.EXPAND | wx.LEFT | wx.RIGHT, border=10)
        self.fields["Discovery"] = self.discovery_combo

        # Submit button - with standard size (110, 40)
        submit_btn = wx.Button(panel, label="Submit and Restart")
        submit_btn.SetInitialSize(wx.Size(110, 40))
        submit_btn.SetMinSize(wx.Size(110, 40))
        submit_btn.SetMaxSize(wx.Size(110, 40))

        # Add close button
        close_btn = wx.Button(panel, label="Submit Later")
        close_btn.SetInitialSize(wx.Size(110, 40))
        close_btn.SetMinSize(wx.Size(110, 40))
        close_btn.SetMaxSize(wx.Size(110, 40))

        # Create horizontal box for buttons
        buttons_hbox = wx.BoxSizer(wx.HORIZONTAL)
        buttons_hbox.Add(close_btn, flag=wx.RIGHT, border=10)
        buttons_hbox.Add(submit_btn)

        # Add the button box to the main vertical box
        vbox.Add(buttons_hbox, flag=wx.ALL | wx.CENTER, border=15)

        submit_btn.Bind(wx.EVT_BUTTON, self.on_submit)
        close_btn.Bind(wx.EVT_BUTTON, self.on_close)

        panel.SetSizer(vbox)
        self.Centre()
        self.Show()

    def on_country_selected(self, event):
        """Update university list when a country is selected"""
        country = self.country_combo.GetValue()

        # If a country is selected, populate university dropdown from our DEFAULT_UNIVERSITIES dictionary
        if country:
            self.university_combo.Clear()

            # Get the list of universities for the selected country
            universities = DEFAULT_UNIVERSITIES.get(country, ["Other - Please specify"])

            # Add all universities to the dropdown
            for university in universities:
                self.university_combo.Append(university)

            # Select the first option by default
            if self.university_combo.GetCount() > 0:
                self.university_combo.SetSelection(0)


    def on_close(self, event):
        self.Close()

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
        wx.MessageBox("Thank you for registering!\nKherveFitting will now Shutdown\n\nClose all error windows and restart the application",
                      "Completed",
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
