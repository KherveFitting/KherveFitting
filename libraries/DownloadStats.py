import wx
import requests
import json
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.patches import Circle
import math
from datetime import datetime, timedelta


class DownloadStatsWindow(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, title="Download Stats", size=(910, 580))
        self.parent = parent
        self.SetMinSize((900, 580))
        self.Center()

        # Create main panel
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Add strength control
        self.geographic_pull_strength = 0.03  # Default value

        # Left panel for buttons
        left_panel = wx.Panel(panel)
        left_sizer = wx.BoxSizer(wx.VERTICAL)

        time_periods = [
            ("Since Sept 2024", self.plot_sept_2024),
            ("Last Year", self.plot_last_year),
            ("Last 6 Months", self.plot_last_6_months),
            ("Last 3 Months", self.plot_last_3_months),
            ("Last 2 Months", self.plot_last_2_months),
            ("Last Month", self.plot_last_month),
            ("Last 3 Weeks", self.plot_last_3_weeks),
            ("Last 2 Weeks", self.plot_last_2_weeks),
            ("Last Week", self.plot_last_week),
            ("Last 2 Days", self.plot_last_2_days),
            ("Yesterday", self.plot_yesterday),
            ("Today", self.plot_today),
            ("Daily (Line)", self.plot_daily_line),
            ("Weekly (Line)", self.plot_weekly_line),
            ("Monthly (Line)", self.plot_monthly_line),
        ]

        for label, callback in time_periods:
            btn = wx.Button(left_panel, label=label, size=(90, 30))
            btn.Bind(wx.EVT_BUTTON, callback)
            left_sizer.Add(btn, 0, wx.ALL | wx.EXPAND, 1)

        # Add strength control after the time period buttons, before setting the sizer
        left_sizer.AddSpacer(10)  # Add some space

        # Add strength control
        strength_sizer = wx.BoxSizer(wx.HORIZONTAL)
        strength_label = wx.StaticText(left_panel, label="Pull:")
        strength_sizer.Add(strength_label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 2)

        self.strength_spin = wx.SpinCtrlDouble(left_panel)
        self.strength_spin.SetRange(0.001, 0.200)
        self.strength_spin.SetIncrement(0.005)
        self.strength_spin.SetValue(0.03)
        self.strength_spin.SetDigits(3)
        self.strength_spin.Bind(wx.EVT_SPINCTRLDOUBLE, self.on_strength_changed)
        strength_sizer.Add(self.strength_spin, 1, wx.ALL, 2)

        left_sizer.Add(strength_sizer, 0, wx.ALL | wx.EXPAND, 2)

        left_panel.SetSizer(left_sizer)

        # Create matplotlib figure and canvas
        # self.figure = plt.Figure(figsize=(8, 5))
        self.figure = plt.Figure()
        self.canvas = FigureCanvas(panel, -1, self.figure)
        self.ax = self.figure.add_subplot(111)

        main_sizer.Add(left_panel, 0, wx.ALL | wx.EXPAND, 10)
        main_sizer.Add(self.canvas, 1, wx.ALL | wx.EXPAND, 5)

        panel.SetSizer(main_sizer)
        self.Show()

        # Initialize with Sept 2024 data
        self.plot_sept_2024(None)
        self.figure.subplots_adjust(top=0.98, bottom=0.08, left=0.08, right=0.95)

    def on_strength_changed(self, event):
        self.geographic_pull_strength = self.strength_spin.GetValue()
        # Optionally refresh the current plot
        # self.plot_world_map_bubbles(self.current_data)  # if you store current_data

    def show_welcome_message(self):
        """Show welcome message"""
        self.ax.clear()
        self.ax.text(0.5, 0.5, 'KherveFitting Download Statistics\n\nSelect a time period above to view download data',
                     ha='center', va='center', fontsize=16, transform=self.ax.transAxes)
        self.ax.set_xlim(0, 1)
        self.ax.set_ylim(0, 1)
        self.ax.axis('off')
        self.canvas.draw()
        self.figure.subplots_adjust(top=0.99, bottom=0.01)

    def get_sourceforge_data(self, start_date, end_date):
        """Fetch data from SourceForge API"""
        try:
            api_url = f"https://sourceforge.net/projects/khervefitting/files/stats/json?start_date={start_date}&end_date={end_date}"
            response = requests.get(api_url)
            response.raise_for_status()
            return response.json()
        except:
            # Fallback sample data if API fails
            return self.get_sample_data()

    def get_sample_data(self):
        """Fallback sample data"""
        return {
            'countries': [
                ['United States', 8234],
                ['United Kingdom', 3456],
                ['Germany', 2987],
                ['France', 2341],
                ['Canada', 1876],
                ['Australia', 1543],
                ['Japan', 1234],
                ['China', 987],
                ['Brazil', 876],
                ['India', 654],
                ['Netherlands', 543],
                ['Sweden', 432],
                ['Italy', 398],
                ['Spain', 321],
                ['Korea', 287]
            ]
        }

    def plot_world_map_bubbles(self, data):
        """Plot world map bubbles exactly like statAPI.py"""
        countries_data = data.get('countries', [])
        if not countries_data:
            return

        # Create country download dictionary
        country_downloads = {}
        for entry in countries_data:
            country_downloads[entry[0]] = entry[1]

        # DEBUG: Print all countries received from API
        print("=== DEBUG: Countries from API ===")
        for country, downloads in country_downloads.items():
            print(f"'{country}': {downloads}")
        print("================================")

        # Country name to acronym mapping
        country_acronyms = {
            'United States': 'US', 'US': 'US',
            'Canada': 'CA', 'CA': 'CA',
            'Mexico': 'MX', 'MX': 'MX',
            'United Kingdom': 'UK', 'GB': 'UK',
            'Germany': 'DE', 'DE': 'DE',
            'France': 'FR', 'FR': 'FR',
            'Italy': 'IT', 'IT': 'IT',
            'Spain': 'ES', 'ES': 'ES',
            'Netherlands': 'NL', 'NL': 'NL',
            'Belgium': 'BE', 'BE': 'BE',
            'Switzerland': 'CH', 'CH': 'CH',
            'Austria': 'AT', 'AT': 'AT',
            'Poland': 'PL', 'PL': 'PL',
            'Czech Republic': 'CZ', 'CZ': 'CZ',
            'Slovakia': 'SK', 'SK': 'SK',
            'Hungary': 'HU', 'HU': 'HU',
            'Romania': 'RO', 'RO': 'RO',
            'Bulgaria': 'BG', 'BG': 'BG',
            'Greece': 'GR', 'GR': 'GR',
            'Portugal': 'PT', 'PT': 'PT',
            'Sweden': 'SE', 'SE': 'SE',
            'Norway': 'NO', 'NO': 'NO',
            'Finland': 'FI', 'FI': 'FI',
            'Denmark': 'DK', 'DK': 'DK',
            'Ireland': 'IE', 'IE': 'IE',
            'Estonia': 'EE', 'EE': 'EE',
            'Iraq': 'IQ', 'IQ': 'IQ',
            'Latvia': 'LV', 'LV': 'LV',
            'Lithuania': 'LT', 'LT': 'LT',
            'Croatia': 'HR', 'HR': 'HR',
            'Slovenia': 'SI', 'SI': 'SI',
            'Ukraine': 'UA', 'UA': 'UA',
            'Turkey': 'TR', 'TR': 'TR',
            'Russia': 'RU', 'RU': 'RU',
            'China': 'CN', 'CN': 'CN',
            'Japan': 'JP', 'JP': 'JP',
            'India': 'IN', 'IN': 'IN',
            'Korea': 'KR', 'KR': 'KR',
            'Indonesia': 'ID', 'ID': 'ID',
            'Thailand': 'TH', 'TH': 'TH',
            'Malaysia': 'MY', 'MY': 'MY',
            'Singapore': 'SG', 'SG': 'SG',
            'Philippines': 'PH', 'PH': 'PH',
            'Viet Nam': 'VN', 'VN': 'VN',
            'Israel': 'IL', 'IL': 'IL',
            'United Arab Emirates': 'AE', 'AE': 'AE',
            'Saudi Arabia': 'SA', 'SA': 'SA',
            'Pakistan': 'PK', 'PK': 'PK',
            'South Africa': 'ZA', 'ZA': 'ZA',
            'Egypt': 'EG', 'EG': 'EG',
            'Nigeria': 'NG', 'NG': 'NG',
            'Kenya': 'KE', 'KE': 'KE',
            'Morocco': 'MA', 'MA': 'MA',
            'Algeria': 'DZ', 'DZ': 'DZ',
            'Tunisia': 'TN', 'TN': 'TN',
            'Brazil': 'BR', 'BR': 'BR',
            'Argentina': 'AR', 'AR': 'AR',
            'Chile': 'CL', 'CL': 'CL',
            'Colombia': 'CO', 'CO': 'CO',
            'Peru': 'PE', 'PE': 'PE',
            'Venezuela': 'VE', 'VE': 'VE',
            'Ecuador': 'EC', 'EC': 'EC',
            'Australia': 'AU', 'AU': 'AU',
            'New Zealand': 'NZ', 'NZ': 'NZ',
            'Hong Kong': 'HK', 'HK': 'HK',
            'Taiwan': 'TW', 'TW': 'TW',
            'Senegal': 'SN', 'SN': 'SN',
            'Iran': 'IR', 'IR': 'IR',
        }

        # Country coordinates (longitude, latitude)
        country_coords = {
            'United States': (-95, 39), 'US': (-95, 39),
            'Canada': (-106, 56), 'CA': (-106, 56),
            'Mexico': (-102, 23), 'MX': (-102, 23),
            'United Kingdom': (-0.13, 71.5), 'GB': (-0.13, 71.5),
            'Germany': (10.45, 51.16), 'DE': (10.45, 51.16),
            'France': (2.21, 46.60), 'FR': (2.21, 46.60),
            'Italy': (12.56, 41.87), 'IT': (12.56, 41.87),
            'Spain': (-3.74, 40.46), 'ES': (-3.74, 40.46),
            'Netherlands': (5.29, 52.13), 'NL': (5.29, 52.13),
            'Belgium': (7.47, 50.50), 'BE': (7.47, 50.50),
            'Switzerland': (8.23, 46.82), 'CH': (8.23, 46.82),
            'Austria': (14.55, 47.52), 'AT': (14.55, 47.52),
            'Poland': (19.15, 51.92), 'PL': (19.15, 51.92),
            'Czech Republic': (15.47, 49.82), 'CZ': (15.47, 49.82),
            'Slovakia': (19.70, 48.67), 'SK': (19.70, 48.67),
            'Hungary': (19.50, 47.16), 'HU': (19.50, 47.16),
            'Romania': (24.97, 45.94), 'RO': (24.97, 45.94),
            'Bulgaria': (25.49, 42.73), 'BG': (25.49, 42.73),
            'Greece': (21.82, 39.07), 'GR': (21.82, 39.07),
            'Portugal': (-9.14, 38.74), 'PT': (-9.14, 38.74),
            'Sweden': (18.64, 70.13), 'SE': (18.64, 70.13),
            'Norway': (8.47, 70.47), 'NO': (8.47, 70.47),
            'Finland': (25.75, 71.92), 'FI': (25.75, 71.92),
            'Denmark': (9.50, 56.26), 'DK': (9.50, 56.26),
            'Ireland': (-8.24, 73.41), 'IE': (-8.24, 73.41),
            'Estonia': (25.01, 58.60), 'EE': (25.01, 58.60),
            'Latvia': (24.60, 56.88), 'LV': (24.60, 56.88),
            'Lithuania': (23.88, 55.17), 'LT': (23.88, 55.17),
            'Croatia': (15.20, 45.10), 'HR': (15.20, 45.10),
            'Slovenia': (14.99, 46.15), 'SI': (14.99, 46.15),
            'Ukraine': (31.17, 48.38), 'UA': (31.17, 48.38),
            'Turkey': (35.24, 38.96), 'TR': (35.24, 38.96),
            'Russia': (105.32, 61.52), 'RU': (105.32, 61.52),
            'China': (104, 35), 'CN': (104, 35),
            'Japan': (138, 36), 'JP': (138, 36),
            'India': (78, 20), 'IN': (78, 20),
            'Korea': (127, 37), 'KR': (127, 37),
            'Indonesia': (113, -5), 'ID': (113, -5),
            'Thailand': (100, 15), 'TH': (100, 15),
            'Malaysia': (101, 4), 'MY': (101, 4),
            'Singapore': (103, 1), 'SG': (103, 1),
            'Philippines': (122, 12), 'PH': (122, 12),
            'Viet Nam': (108, 14), 'VN': (108, 14),
            'Iraq': (44, 33), 'IQ': (44, 33),
            'Israel': (34, 31), 'IL': (34, 31),
            'United Arab Emirates': (53, 23), 'AE': (53, 23),
            'Saudi Arabia': (45, 24), 'SA': (45, 24),
            'Pakistan': (70, 30), 'PK': (70, 30),
            'South Africa': (24, -49), 'ZA': (24, -49),
            'Egypt': (30, -36), 'EG': (30, -36),
            'Nigeria': (8, -29), 'NG': (8, -29),
            'Kenya': (37, -31), 'KE': (37, -31),
            'Morocco': (-7, -21), 'MA': (-7, -21),
            'Algeria': (1, -22), 'DZ': (1, -22),
            'Tunisia': (9, -23), 'TN': (9, -23),
            'Brazil': (-47, -14), 'BR': (-47, -14),
            'Argentina': (-63, -38), 'AR': (-63, -38),
            'Chile': (-71, -35), 'CL': (-71, -35),
            'Colombia': (-74, 4), 'CO': (-74, 4),
            'Peru': (-75, -9), 'PE': (-75, -9),
            'Venezuela': (-66, 10), 'VE': (-66, 10),
            'Ecuador': (-78, -1), 'EC': (-78, -1),
            'Australia': (133, -27), 'AU': (133, -27),
            'New Zealand': (138, -31), 'NZ': (138, -31),
            'Hong Kong': (114, 22), 'HK': (114, 22),
            'Taiwan': (121, 24), 'TW': (121, 24),
            'Senegal': (-14, -26), 'SN': (-14, -26),
            'Iran': (53, 32), 'IR': (53, 32),
        }

        # Prepare bubble data
        bubbles = []
        max_downloads = max(country_downloads.values()) if country_downloads else 1

        for country, downloads in country_downloads.items():
            if country in country_coords:
                lon, lat = country_coords[country]
                download_ratio = downloads / max_downloads
                min_radius = 4.0
                max_radius = 30.0
                radius = min_radius + (max_radius - min_radius) * math.sqrt(download_ratio)
                color_intensity = download_ratio
                acronym = country_acronyms.get(country, country[:3].upper())

                bubbles.append({
                    'country': country,
                    'acronym': acronym,
                    'downloads': downloads,
                    'x': lon,
                    'y': lat,
                    'original_x': lon,
                    'original_y': lat,
                    'radius': radius,
                    'color_intensity': color_intensity
                })

        # # Sort bubbles by size (largest first)
        # bubbles.sort(key=lambda b: b['radius'], reverse=True)
        # Sort bubbles by downloads (largest first) - gives priority to bigger countries
        bubbles.sort(key=lambda b: b['downloads'], reverse=True)

        # Collision resolution
        def distance(b1, b2):
            return math.sqrt((b1['x'] - b2['x']) ** 2 + (b1['y'] - b2['y']) ** 2)

        def check_overlap(b1, b2):
            min_distance = b1['radius'] + b2['radius']
            return distance(b1, b2) < min_distance

        def resolve_overlap_BEFORE(b1, b2):
            current_dist = distance(b1, b2)
            min_dist = b1['radius'] + b2['radius']

            if current_dist < min_dist and current_dist > 0:
                dx = b2['x'] - b1['x']
                dy = b2['y'] - b1['y']
                factor = min_dist / current_dist
                total_radius = b1['radius'] + b2['radius']
                b1_weight = b2['radius'] / total_radius
                b2_weight = b1['radius'] / total_radius
                move_x = dx * (factor - 1) * 0.5
                move_y = dy * (factor - 1) * 0.5
                b1['x'] -= move_x * b1_weight
                b1['y'] -= move_y * b1_weight
                b2['x'] += move_x * b2_weight
                b2['y'] += move_y * b2_weight

        def resolve_overlap(b1, b2):
            current_dist = distance(b1, b2)
            min_dist = b1['radius'] + b2['radius']

            if current_dist < min_dist and current_dist > 0:
                dx = b2['x'] - b1['x']
                dy = b2['y'] - b1['y']
                factor = min_dist / current_dist
                total_radius = b1['radius'] + b2['radius']

                # Add mass factor to favor larger countries
                mass_factor = math.sqrt(b1['downloads'] / b2['downloads']) if b2['downloads'] > 0 else 1.0
                mass_factor = max(0.1, min(10.0, mass_factor))  # Clamp between 0.1 and 10

                # Adjust weights based on both radius and mass
                b1_weight = (b2['radius'] / total_radius) / mass_factor
                b2_weight = (b1['radius'] / total_radius) * mass_factor

                # Normalize weights
                total_weight = b1_weight + b2_weight
                b1_weight /= total_weight
                b2_weight /= total_weight

                move_x = dx * (factor - 1) * 0.5
                move_y = dy * (factor - 1) * 0.5
                b1['x'] -= move_x * b1_weight
                b1['y'] -= move_y * b1_weight
                b2['x'] += move_x * b2_weight
                b2['y'] += move_y * b2_weight

        def apply_geographic_pull(bubble, strength=0.03):
            dx = bubble['original_x'] - bubble['x']
            dy = bubble['original_y'] - bubble['y']
            bubble['x'] += dx * strength
            bubble['y'] += dy * strength

        # Iterative collision resolution
        for iteration in range(150):
            overlaps_resolved = True
            for i in range(len(bubbles)):
                for j in range(i + 1, len(bubbles)):
                    if check_overlap(bubbles[i], bubbles[j]):
                        resolve_overlap(bubbles[i], bubbles[j])
                        overlaps_resolved = False
            for bubble in bubbles:
                apply_geographic_pull(bubble, strength=self.geographic_pull_strength)
            if overlaps_resolved:
                break

        # Clear plot and create bubble chart
        self.ax.clear()
        self.ax.set_xlim(-130, 170)
        self.ax.set_ylim(-60, 130)
        # self.figure.tight_layout()
        # self.figure.subplots_adjust(top=0.99, bottom=0.01, left=0.05, right=0.95)
        self.ax.set_aspect('equal')
        self.ax.grid(True, alpha=0.2, linestyle='--')

        # Add continent reference areas
        continent_boxes = [
            {'box': (-170, -50, 15, 85), 'color': 'lightblue', 'alpha': 0.1, 'label': 'North America'},
            {'box': (-85, -60, -30, 15), 'color': 'lightgreen', 'alpha': 0.1, 'label': 'South America'},
            {'box': (-15, 30, 40, 75), 'color': 'lightcoral', 'alpha': 0.1, 'label': 'Europe'},
            {'box': (-20, -50, 55, -10), 'color': 'lightyellow', 'alpha': 0.1, 'label': 'Africa'},
            {'box': (25, -15, 180, 80), 'color': 'lightpink', 'alpha': 0.1, 'label': 'Asia'},
            {'box': (110, -50, 180, -10), 'color': 'lightgray', 'alpha': 0.1, 'label': 'Oceania'},
        ]

        for continent in continent_boxes:
            x1, y1, x2, y2 = continent['box']
            width = x2 - x1
            height = y2 - y1
            rect = plt.Rectangle((x1, y1), width, height,
                                 facecolor=continent['color'],
                                 alpha=continent['alpha'],
                                 edgecolor='gray', linewidth=0.5)
            self.ax.add_patch(rect)
            self.ax.text(x1 + width / 2, y2 - 5, continent['label'],
                         ha='center', va='top', fontsize=8,
                         alpha=0.6, style='italic')

        def calculate_font_size(radius, text_length):
            base_size = max(6, min(16, int(radius * 1.2)))
            if text_length > 6:
                base_size = max(6, int(base_size * 0.8))
            elif text_length > 4:
                base_size = max(8, int(base_size * 0.9))
            return base_size

        def format_downloads(downloads):
            if downloads >= 1000000:
                return f"{downloads / 1000000:.1f}M"
            elif downloads >= 1000:
                return f"{downloads / 1000:.1f}K"
            else:
                return str(downloads)

        # Draw bubbles
        for bubble in bubbles:
            circle = Circle((bubble['x'], bubble['y']), bubble['radius'],
                            facecolor=plt.cm.YlOrRd(bubble['color_intensity']),
                            edgecolor='black', linewidth=2, alpha=0.85)
            self.ax.add_patch(circle)

            acronym = bubble['acronym']
            downloads_text = format_downloads(bubble['downloads'])
            display_text = f"{acronym}\n{downloads_text}"
            font_size = calculate_font_size(bubble['radius'], len(display_text.replace('\n', '')))
            text_color = 'white' if bubble['color_intensity'] > 0.5 else 'black'

            self.ax.text(bubble['x'], bubble['y'], display_text,
                         ha='center', va='center', fontsize=font_size, fontweight='bold',
                         color=text_color,
                         bbox=dict(boxstyle='round,pad=0.1', facecolor='none', edgecolor='none', alpha=0.7))

        # Add statistics
        total_downloads = sum(country_downloads.values())
        stats_text = f'Downloads: {format_downloads(total_downloads)}\n'
        stats_text += f'Countries: {len(bubbles)}\n'
        stats_text += f'Largest: {bubbles[0]["acronym"]} ({format_downloads(bubbles[0]["downloads"])})\n'

        # Add last updated time
        last_updated = data.get('stats_updated', 'Unknown')
        if last_updated != 'Unknown':
            try:
                # Try to format the date nicely
                from datetime import datetime
                dt = datetime.fromisoformat(last_updated.replace('Z', '+00:00'))
                formatted_date = dt.strftime('%H:%M')
                stats_text += f'Updated: {formatted_date}\n'
                stats_text += f'Date: {dt.strftime('%d-%m-%y')}'
            except:
                stats_text += f'Updated: {last_updated}'
        else:
            stats_text += f'Updated: {last_updated}'

        self.ax.text(0.85, 0.98, stats_text,
                     transform=self.ax.transAxes, fontsize=8,
                     bbox=dict(boxstyle='round,pad=0.4', facecolor='white', alpha=0.9),
                     verticalalignment='top', zorder=100)

        self.ax.set_xlabel('Longitude', fontsize=12)
        self.ax.set_ylabel('Latitude', fontsize=12)
        self.canvas.draw()

    def plot_sept_2024(self, event):
        data = self.get_sourceforge_data("2024-09-01", datetime.now().strftime("%Y-%m-%d"))
        self.plot_world_map_bubbles(data)

    def plot_last_year(self, event):
        end_date = datetime.now()
        start_date = end_date - timedelta(days=365)
        data = self.get_sourceforge_data(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
        self.plot_world_map_bubbles(data)

    def plot_last_6_months(self, event):
        end_date = datetime.now()
        start_date = end_date - timedelta(days=180)
        data = self.get_sourceforge_data(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
        self.plot_world_map_bubbles(data)

    def plot_last_3_months(self, event):
        end_date = datetime.now()
        start_date = end_date - timedelta(days=90)
        data = self.get_sourceforge_data(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
        self.plot_world_map_bubbles(data)

    def plot_last_2_months(self, event):
        end_date = datetime.now()
        start_date = end_date - timedelta(days=60)
        data = self.get_sourceforge_data(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
        self.plot_world_map_bubbles(data)

    def plot_last_month(self, event):
        end_date = datetime.now()
        start_date = end_date - timedelta(days=30)
        data = self.get_sourceforge_data(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
        self.plot_world_map_bubbles(data)

    def plot_last_3_weeks(self, event):
        end_date = datetime.now()
        start_date = end_date - timedelta(days=21)
        data = self.get_sourceforge_data(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
        self.plot_world_map_bubbles(data)

    def plot_last_2_weeks(self, event):
        end_date = datetime.now()
        start_date = end_date - timedelta(days=14)
        data = self.get_sourceforge_data(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
        self.plot_world_map_bubbles(data)

    def plot_last_week(self, event):
        end_date = datetime.now()
        start_date = end_date - timedelta(days=7)
        data = self.get_sourceforge_data(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
        self.plot_world_map_bubbles(data)

    def plot_yesterday(self, event):
        yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        data = self.get_sourceforge_data(yesterday, yesterday)
        self.plot_world_map_bubbles(data)

    def plot_last_2_days(self, event):
        end_date = datetime.now()
        start_date = end_date - timedelta(days=2)
        data = self.get_sourceforge_data(start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d"))
        self.plot_world_map_bubbles(data)

    def plot_today(self, event):
        today = datetime.now().strftime("%Y-%m-%d")
        data = self.get_sourceforge_data(today, today)
        self.plot_world_map_bubbles(data)

    def plot_weekly_line(self, event):
        try:
            api_url = f"https://sourceforge.net/projects/khervefitting/files/stats/json?start_date=2024-09-01&end_date={datetime.now().strftime('%Y-%m-%d')}&period=weekly"
            response = requests.get(api_url)
            response.raise_for_status()
            data = response.json()

            self.figure.clear()
            self.ax = self.figure.add_subplot(111)

            # Get download data - expecting a list of lists [date, count]
            downloads_data = data.get('downloads', [])

            if not downloads_data:
                self.ax.text(0.5, 0.5, 'No downloaded data available',
                             horizontalalignment='center', verticalalignment='center')
                self.canvas.draw()
                return

            dates = []
            download_counts = []

            for entry in downloads_data:
                date_str = entry[0]
                # Try to extract time info if available
                try:
                    date_obj = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
                    dates.append(date_obj)
                except ValueError:
                    # If only date is provided, use it
                    dates.append(datetime.strptime(date_str, "%Y-%m-%d"))

                download_counts.append(entry[1])

            # Convert to pandas for easier handling
            import pandas as pd
            df = pd.DataFrame({'date': dates, 'downloads': download_counts})
            df = df.sort_values('date')

            self.ax.plot(df['date'], df['downloads'], marker='o')
            # self.ax.set_title('Weekly Downloads Since Sept 2024')
            self.ax.set_xlabel('Date')
            self.ax.set_ylabel('Downloads')
            self.ax.grid(True)

            # Add total downloads annotation
            total_downloads = sum(download_counts)
            self.ax.annotate(f'Total Downloads: {total_downloads}',
                             xy=(0.02, 0.95), xycoords='axes fraction',
                             fontsize=10, bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="gray", alpha=0.8))

            # Format x-axis dates
            self.figure.autofmt_xdate()

            # Update the canvas
            self.canvas.draw()

        except Exception as e:
            print(f"Error details: {e}")
            import traceback
            traceback.print_exc()
            self.figure.clear()
            self.ax = self.figure.add_subplot(111)
            self.ax.text(0.5, 0.5, f'Error: {str(e)}', ha='center', va='center')
            self.canvas.draw()

    def plot_daily_line(self, event):
        try:
            # Last 3 months instead of since Sept 2024
            start_date = (datetime.now() - timedelta(days=90)).strftime('%Y-%m-%d')
            api_url = f"https://sourceforge.net/projects/khervefitting/files/stats/json?start_date={start_date}&end_date={datetime.now().strftime('%Y-%m-%d')}&period=daily"
            response = requests.get(api_url)
            response.raise_for_status()
            data = response.json()

            self.figure.clear()
            self.ax = self.figure.add_subplot(111)

            # Get download data - expecting a list of lists [date, count]
            downloads_data = data.get('downloads', [])

            if not downloads_data:
                self.ax.text(0.5, 0.5, 'No downloaded data available',
                             horizontalalignment='center', verticalalignment='center')
                self.canvas.draw()
                return

            dates = []
            download_counts = []

            for entry in downloads_data:
                date_str = entry[0]
                # Try to extract time info if available
                try:
                    date_obj = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
                    dates.append(date_obj)
                except ValueError:
                    # If only date is provided, use it
                    dates.append(datetime.strptime(date_str, "%Y-%m-%d"))

                download_counts.append(entry[1])

            # Convert to pandas for easier handling
            import pandas as pd
            df = pd.DataFrame({'date': dates, 'downloads': download_counts})
            df = df.sort_values('date')

            self.ax.plot(df['date'], df['downloads'], marker='o')
            # self.ax.set_title('Daily Downloads - Last 3 Months')
            self.ax.set_xlabel('Date')
            self.ax.set_ylabel('Downloads')
            self.ax.grid(True)

            # Add total downloads annotation
            total_downloads = sum(download_counts)
            self.ax.annotate(f'Total Downloads: {total_downloads}',
                             xy=(0.02, 0.95), xycoords='axes fraction',
                             fontsize=10, bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="gray", alpha=0.8))

            # Format x-axis dates
            self.figure.autofmt_xdate()

            # Update the canvas
            self.canvas.draw()

        except Exception as e:
            print(f"Error details: {e}")
            import traceback
            traceback.print_exc()
            self.figure.clear()
            self.ax = self.figure.add_subplot(111)
            self.ax.text(0.5, 0.5, f'Error: {str(e)}', ha='center', va='center')
            self.canvas.draw()

    def plot_monthly_line(self, event):
        try:
            api_url = f"https://sourceforge.net/projects/khervefitting/files/stats/json?start_date=2024-09-01&end_date={datetime.now().strftime('%Y-%m-%d')}&period=monthly"
            response = requests.get(api_url)
            response.raise_for_status()
            data = response.json()

            self.figure.clear()
            self.ax = self.figure.add_subplot(111)

            # Get download data - expecting a list of lists [date, count]
            downloads_data = data.get('downloads', [])

            if not downloads_data:
                self.ax.text(0.5, 0.5, 'No downloaded data available',
                             horizontalalignment='center', verticalalignment='center')
                self.canvas.draw()
                return

            dates = []
            download_counts = []

            for entry in downloads_data:
                date_str = entry[0]
                # Try to extract time info if available
                try:
                    date_obj = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
                    dates.append(date_obj)
                except ValueError:
                    # If only date is provided, use it
                    dates.append(datetime.strptime(date_str, "%Y-%m-%d"))

                download_counts.append(entry[1])

            # Convert to pandas for easier handling
            import pandas as pd
            df = pd.DataFrame({'date': dates, 'downloads': download_counts})
            df = df.sort_values('date')

            self.ax.plot(df['date'], df['downloads'], marker='o')
            # self.ax.set_title('Monthly Downloads Since Sept 2024')
            self.ax.set_xlabel('Date')
            self.ax.set_ylabel('Downloads')
            self.ax.grid(True)

            # Add total downloads annotation
            total_downloads = sum(download_counts)
            self.ax.annotate(f'Total Downloads: {total_downloads}',
                             xy=(0.02, 0.95), xycoords='axes fraction',
                             fontsize=10, bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="gray", alpha=0.8))

            # Format x-axis dates
            self.figure.autofmt_xdate()

            # Update the canvas
            self.canvas.draw()

        except Exception as e:
            print(f"Error details: {e}")
            import traceback
            traceback.print_exc()
            self.figure.clear()
            self.ax = self.figure.add_subplot(111)
            self.ax.text(0.5, 0.5, f'Error: {str(e)}', ha='center', va='center')
            self.canvas.draw()


def show_download_stats_window(parent):
    """Show the download stats window"""
    DownloadStatsWindow(parent)