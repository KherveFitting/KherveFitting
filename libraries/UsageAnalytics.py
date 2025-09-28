import wx
import requests
import json
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.patches import Circle
import math
import csv
import io
from datetime import datetime, timedelta

# Google Form configuration
GOOGLE_SHEETS_ID = '1mP23kwwg-H1Gx5ppp2YQNv6X1lYhJj2B6_PagQ0yABw'
GOOGLE_SHEETS_CSV_URL = f'https://docs.google.com/spreadsheets/d/{GOOGLE_SHEETS_ID}/export?format=csv&gid=213903690'


class UsageStatsWindow(wx.Frame):
    def __init__(self, parent):
        super().__init__(parent, title="Usage Stats", size=(1100, 650))
        self.SetMinSize((1100, 650))
        self.SetMaxSize((1100, 650))  # Fix to this size, prevent resizing

        self.parent = parent
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
            ("Last 4 Days", self.plot_last_4_days),
            ("Last 3 Days", self.plot_last_3_days),
            ("Last 2 Days", self.plot_last_2_days),
            ("2 Days Ago", self.plot_2_days_ago),
            ("Yesterday", self.plot_yesterday),
            ("Today", self.plot_today),
            ("All Usage Data", self.plot_all_usage),
            ("Google Form", self.plot_google_form),
        ]

        for label, callback in time_periods:
            btn = wx.Button(left_panel, label=label, size=(90, 30))
            btn.Bind(wx.EVT_BUTTON, callback)
            left_sizer.Add(btn, 0, wx.ALL | wx.EXPAND, 0)

        # Add strength control after the time period buttons
        left_sizer.AddSpacer(10)

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
        self.figure = plt.Figure()
        self.canvas = FigureCanvas(panel, -1, self.figure)
        self.ax = self.figure.add_subplot(111)

        main_sizer.Add(left_panel, 0, wx.ALL | wx.EXPAND, 10)
        main_sizer.Add(self.canvas, 1, wx.ALL | wx.EXPAND, 5)

        panel.SetSizer(main_sizer)
        self.Show()

        # Initialize with all usage data
        self.plot_all_usage(None)
        self.figure.subplots_adjust(top=0.98, bottom=0.08, left=0.08, right=0.95)

        panel.Layout()
        self.Layout()
        self.Refresh()

        if hasattr(self, 'canvas'):
            self.canvas.SetSize((950, 610))
            self.canvas.draw()

    def on_strength_changed(self, event):
        self.geographic_pull_strength = self.strength_spin.GetValue()

    def get_google_form_data(self, start_date=None, end_date=None):
        """Fetch usage data from Google Form with optional date filtering"""
        try:
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }

            response = requests.get(GOOGLE_SHEETS_CSV_URL, headers=headers)
            if response.status_code != 200:
                return self.get_sample_usage_data()

            csv_data = response.content.decode('utf-8')
            reader = csv.DictReader(io.StringIO(csv_data))

            country_usage = {}

            for row in reader:
                # Parse timestamp
                timestamp_str = row.get('Timestamp', '')
                if timestamp_str and start_date and end_date:
                    try:
                        # Parse Google Form timestamp (usually in format: "MM/DD/YYYY HH:MM:SS")
                        timestamp = datetime.strptime(timestamp_str, "%m/%d/%Y %H:%M:%S")

                        # Check if timestamp is within date range
                        if not (start_date <= timestamp <= end_date):
                            continue  # Skip this row if outside date range

                    except ValueError:
                        # Try alternative timestamp formats
                        try:
                            timestamp = datetime.strptime(timestamp_str, "%Y-%m-%d %H:%M:%S")
                            if not (start_date <= timestamp <= end_date):
                                continue
                        except ValueError:
                            try:
                                timestamp = datetime.strptime(timestamp_str, "%d/%m/%Y %H:%M:%S")
                                if not (start_date <= timestamp <= end_date):
                                    continue
                            except ValueError:
                                # If we can't parse the timestamp, include the row
                                pass

                country = row.get('Country', 'Unknown')
                if ',' in country:
                    country = country.split(',')[0].strip()

                if country in country_usage:
                    country_usage[country] += 1
                else:
                    country_usage[country] = 1

            # Convert to format expected by plot_world_map_bubbles
            countries_list = [[country, count] for country, count in country_usage.items()]
            countries_list.sort(key=lambda x: x[1], reverse=True)

            return {
                'countries': countries_list,
                'stats_updated': datetime.now().isoformat()
            }

        except Exception as e:
            print(f"Error reading Google Form data: {e}")
            return self.get_sample_usage_data()

    def get_date_range(self, period_type):
        """Calculate start and end dates for different periods"""
        now = datetime.now()

        if period_type == "today":
            start_date = now.replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = now.replace(hour=23, minute=59, second=59, microsecond=999999)
        elif period_type == "yesterday":
            yesterday = now - timedelta(days=1)
            start_date = yesterday.replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = yesterday.replace(hour=23, minute=59, second=59, microsecond=999999)
        elif period_type == "2_days_ago":
            two_days_ago = now - timedelta(days=2)
            start_date = two_days_ago.replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = two_days_ago.replace(hour=23, minute=59, second=59, microsecond=999999)
        elif period_type == "last_2_days":
            start_date = (now - timedelta(days=2)).replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = now
        elif period_type == "last_3_days":
            start_date = (now - timedelta(days=3)).replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = now
        elif period_type == "last_4_days":
            start_date = (now - timedelta(days=4)).replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = now
        elif period_type == "last_week":
            start_date = now - timedelta(days=7)
            end_date = now
        elif period_type == "last_2_weeks":
            start_date = now - timedelta(days=14)
            end_date = now
        elif period_type == "last_3_weeks":
            start_date = now - timedelta(days=21)
            end_date = now
        elif period_type == "last_month":
            start_date = now - timedelta(days=30)
            end_date = now
        elif period_type == "last_2_months":
            start_date = now - timedelta(days=60)
            end_date = now
        elif period_type == "last_3_months":
            start_date = now - timedelta(days=90)
            end_date = now
        elif period_type == "last_6_months":
            start_date = now - timedelta(days=180)
            end_date = now
        elif period_type == "last_year":
            start_date = now - timedelta(days=365)
            end_date = now
        elif period_type == "since_sept_2024":
            start_date = datetime(2024, 9, 1)
            end_date = now
        else:
            # Default: no date filtering
            start_date = None
            end_date = None

        return start_date, end_date

    def get_sample_usage_data(self):
        """Fallback sample data if Google Form is unavailable"""
        return {
            'countries': [
                ['United States', 45],
                ['United Kingdom', 32],
                ['Germany', 28],
                ['France', 23],
                ['Canada', 18],
                ['Australia', 15],
                ['Japan', 12],
                ['China', 9],
                ['Brazil', 8],
                ['India', 6],
                ['Netherlands', 5],
                ['Sweden', 4],
                ['Italy', 3],
                ['Spain', 3],
                ['Korea', 2]
            ],
            'stats_updated': datetime.now().isoformat()
        }

    def plot_world_map_bubbles(self, data):
        """Plot world map bubbles exactly like DownloadStats.py"""
        countries_data = data.get('countries', [])
        if not countries_data:
            return

        # Create country usage dictionary
        country_usage = {}
        for entry in countries_data:
            country_usage[entry[0]] = entry[1]

        # Country name to acronym mapping (same as DownloadStats.py)
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
            'South Korea': 'KR', 'Korea': 'KR', 'KR': 'KR',
            'Indonesia': 'ID', 'ID': 'ID',
            'Thailand': 'TH', 'TH': 'TH',
            'Malaysia': 'MY', 'MY': 'MY',
            'Singapore': 'SG', 'SG': 'SG',
            'Philippines': 'PH', 'PH': 'PH',
            'Vietnam': 'VN', 'VN': 'VN',
            'Israel': 'IL', 'IL': 'IL',
            'UAE': 'AE', 'AE': 'AE',
            'Saudi Arabia': 'SA', 'SA': 'SA',
            'Pakistan': 'PK', 'PK': 'PK',
            'South Africa': 'ZA', 'ZA': 'ZA',
            'Egypt': 'EG', 'EG': 'EG',
            'Nigeria': 'NG', 'NG': 'NG',
            'Kenya': 'KE', 'KE': 'KE',
            'Morocco': 'MA', 'MA': 'MA',
            'Algeria': 'DZ', 'DZ': 'DZ',
            'Tunisia': 'TN', 'TN': 'TN',
            'Argentina': 'AR', 'AR': 'AR',
            'Chile': 'CL', 'CL': 'CL',
            'Colombia': 'CO', 'CO': 'CO',
            'Peru': 'PE', 'PE': 'PE',
            'Venezuela': 'VE', 'VE': 'VE',
            'Ecuador': 'EC', 'EC': 'EC',
            'Australia': 'AU', 'AU': 'AU',
            'New Zealand': 'NZ', 'NZ': 'NZ',
        }

        # Country coordinates (same as DownloadStats.py)
        country_coords = {
            'United States': (-95, 39), 'US': (-95, 39),
            'Canada': (-106, 56), 'CA': (-106, 56),
            'Mexico': (-102, 23), 'MX': (-102, 23),
            'United Kingdom': (-3, 55), 'GB': (-3, 55),
            'Germany': (10, 51), 'DE': (10, 51),
            'France': (2, 46), 'FR': (2, 46),
            'Italy': (12, 42), 'IT': (12, 42),
            'Spain': (-4, 40), 'ES': (-4, 40),
            'Netherlands': (5, 52), 'NL': (5, 52),
            'Belgium': (4, 50), 'BE': (4, 50),
            'Switzerland': (8, 47), 'CH': (8, 47),
            'Austria': (14, 48), 'AT': (14, 48),
            'Poland': (19, 52), 'PL': (19, 52),
            'Czech Republic': (15, 50), 'CZ': (15, 50),
            'Slovakia': (20, 49), 'SK': (20, 49),
            'Hungary': (20, 47), 'HU': (20, 47),
            'Romania': (25, 46), 'RO': (25, 46),
            'Bulgaria': (25, 43), 'BG': (25, 43),
            'Greece': (22, 39), 'GR': (22, 39),
            'Portugal': (-9, 39), 'PT': (-9, 39),
            'Sweden': (18, 60), 'SE': (18, 60),
            'Norway': (8, 60), 'NO': (8, 60),
            'Finland': (26, 62), 'FI': (26, 62),
            'Denmark': (9, 56), 'DK': (9, 56),
            'Ireland': (-8, 53), 'IE': (-8, 53),
            'Estonia': (25, 59), 'EE': (25, 59),
            'Latvia': (25, 57), 'LV': (25, 57),
            'Lithuania': (24, 55), 'LT': (24, 55),
            'Croatia': (15, 45), 'HR': (15, 45),
            'Slovenia': (15, 46), 'SI': (15, 46),
            'Ukraine': (31, 48), 'UA': (31, 48),
            'Turkey': (35, 39), 'TR': (35, 39),
            'Russia': (105, 62), 'RU': (105, 62),
            'China': (104, 35), 'CN': (104, 35),
            'Japan': (138, 36), 'JP': (138, 36),
            'India': (78, 20), 'IN': (78, 20),
            'South Korea': (127, 37), 'Korea': (127, 37), 'KR': (127, 37),
            'Indonesia': (113, -5), 'ID': (113, -5),
            'Thailand': (100, 15), 'TH': (100, 15),
            'Malaysia': (101, 4), 'MY': (101, 4),
            'Singapore': (103, 1), 'SG': (103, 1),
            'Philippines': (122, 12), 'PH': (122, 12),
            'Vietnam': (108, 14), 'VN': (108, 14),
            'Israel': (34, 31), 'IL': (34, 31),
            'UAE': (53, 23), 'AE': (53, 23),
            'Saudi Arabia': (45, 24), 'SA': (45, 24),
            'Pakistan': (70, 30), 'PK': (70, 30),
            'South Africa': (24, -49), 'ZA': (24, -49),
            'Egypt': (30, -36), 'EG': (30, -36),
            'Nigeria': (8, -29), 'NG': (8, -29),
            'Kenya': (37, -31), 'KE': (37, -31),
            'Morocco': (-7, -21), 'MA': (-7, -21),
            'Algeria': (1, -22), 'DZ': (1, -22),
            'Tunisia': (9, -23), 'TN': (9, -23),
            'Argentina': (-64, 34), 'AR': (-64, 34),
            'Chile': (-71, 30), 'CL': (-71, 30),
            'Colombia': (-74, 4), 'CO': (-74, 4),
            'Peru': (-76, 10), 'PE': (-76, 10),
            'Venezuela': (-66, 10), 'VE': (-66, 10),
            'Ecuador': (-78, -1), 'EC': (-78, -1),
            'Australia': (133, -27), 'AU': (133, -27),
            'New Zealand': (138, -31), 'NZ': (138, -31),
        }

        # Prepare bubble data
        bubbles = []
        max_usage = max(country_usage.values()) if country_usage else 1

        for country, usage_count in country_usage.items():
            if country in country_coords:
                lon, lat = country_coords[country]
                usage_ratio = usage_count / max_usage
                min_radius = 4.0
                max_radius = 30.0
                radius = min_radius + (max_radius - min_radius) * math.sqrt(usage_ratio)
                color_intensity = usage_ratio
                acronym = country_acronyms.get(country, country[:3].upper())

                bubbles.append({
                    'country': country,
                    'acronym': acronym,
                    'usage': usage_count,
                    'x': lon,
                    'y': lat,
                    'original_x': lon,
                    'original_y': lat,
                    'radius': radius,
                    'color_intensity': color_intensity
                })

        # Sort bubbles by usage (largest first)
        bubbles.sort(key=lambda b: b['usage'], reverse=True)

        # Collision resolution (same as DownloadStats.py)
        def distance(b1, b2):
            return math.sqrt((b1['x'] - b2['x']) ** 2 + (b1['y'] - b2['y']) ** 2)

        def check_overlap(b1, b2):
            min_distance = b1['radius'] + b2['radius']
            return distance(b1, b2) < min_distance

        def resolve_overlap(b1, b2):
            current_dist = distance(b1, b2)
            min_dist = b1['radius'] + b2['radius']

            if current_dist < min_dist and current_dist > 0:
                dx = b2['x'] - b1['x']
                dy = b2['y'] - b1['y']
                factor = min_dist / current_dist
                total_radius = b1['radius'] + b2['radius']

                mass_factor = math.sqrt(b1['usage'] / b2['usage']) if b2['usage'] > 0 else 1.0
                mass_factor = max(0.1, min(10.0, mass_factor))

                b1_weight = (b2['radius'] / total_radius) / mass_factor
                b2_weight = (b1['radius'] / total_radius) * mass_factor

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
        self.ax.set_aspect('equal')
        self.ax.grid(True, alpha=0.2, linestyle='--')

        # Add continent reference areas (same as DownloadStats.py)
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

        def format_usage(usage_count):
            if usage_count >= 1000000:
                return f"{usage_count / 1000000:.1f}M"
            elif usage_count >= 1000:
                return f"{usage_count / 1000:.1f}K"
            else:
                return str(usage_count)

        # Draw bubbles
        for bubble in bubbles:
            circle = Circle((bubble['x'], bubble['y']), bubble['radius'],
                            facecolor=plt.cm.YlOrRd(bubble['color_intensity']),
                            edgecolor='black', linewidth=2, alpha=0.85)
            self.ax.add_patch(circle)

            acronym = bubble['acronym']
            usage_text = format_usage(bubble['usage'])
            display_text = f"{acronym}\n{usage_text}"
            font_size = calculate_font_size(bubble['radius'], len(display_text.replace('\n', '')))
            text_color = 'white' if bubble['color_intensity'] > 0.5 else 'black'

            self.ax.text(bubble['x'], bubble['y'], display_text,
                         ha='center', va='center', fontsize=font_size, fontweight='bold',
                         color=text_color,
                         bbox=dict(boxstyle='round,pad=0.1', facecolor='none', edgecolor='none', alpha=0.7))

        # Add statistics
        total_usage = sum(country_usage.values())
        stats_text = f'Usage: {format_usage(total_usage)}\n'
        stats_text += f'Countries: {len(bubbles)}\n'
        if bubbles:
            stats_text += f'Largest: {bubbles[0]["acronym"]} ({format_usage(bubbles[0]["usage"])})\n'

        # Add last updated time
        last_updated = data.get('stats_updated', 'Unknown')
        if last_updated != 'Unknown':
            try:
                dt = datetime.fromisoformat(last_updated.replace('Z', '+00:00'))
                formatted_date = dt.strftime('%H:%M')
                stats_text += f'Updated: {formatted_date}\n'
                stats_text += f'Date: {dt.strftime("%d-%m-%y")}'
            except:
                stats_text += f'Updated: {last_updated}'
        else:
            stats_text += f'Updated: {last_updated}'

        self.ax.text(0.85, 0.98, stats_text,
                     transform=self.ax.transAxes, fontsize=8,
                     bbox=dict(boxstyle='round,pad=0.4', facecolor='white', alpha=0.9),
                     verticalalignment='top', zorder=100)

        self.ax.set_xlabel('~Longitude (depends on usage and pull)', fontsize=12)
        self.ax.set_ylabel('~Latitude (depends on usage and pull)', fontsize=12)
        self.canvas.draw()

    # Button callback methods with proper date filtering
    def plot_today(self, event):
        start_date, end_date = self.get_date_range("today")
        data = self.get_google_form_data(start_date, end_date)
        self.plot_world_map_bubbles(data)

    def plot_yesterday(self, event):
        start_date, end_date = self.get_date_range("yesterday")
        data = self.get_google_form_data(start_date, end_date)
        self.plot_world_map_bubbles(data)

    def plot_2_days_ago(self, event):
        start_date, end_date = self.get_date_range("2_days_ago")
        data = self.get_google_form_data(start_date, end_date)
        self.plot_world_map_bubbles(data)

    def plot_last_2_days(self, event):
        start_date, end_date = self.get_date_range("last_2_days")
        data = self.get_google_form_data(start_date, end_date)
        self.plot_world_map_bubbles(data)

    def plot_last_3_days(self, event):
        start_date, end_date = self.get_date_range("last_3_days")
        data = self.get_google_form_data(start_date, end_date)
        self.plot_world_map_bubbles(data)

    def plot_last_4_days(self, event):
        start_date, end_date = self.get_date_range("last_4_days")
        data = self.get_google_form_data(start_date, end_date)
        self.plot_world_map_bubbles(data)

    def plot_last_week(self, event):
        start_date, end_date = self.get_date_range("last_week")
        data = self.get_google_form_data(start_date, end_date)
        self.plot_world_map_bubbles(data)

    def plot_last_2_weeks(self, event):
        start_date, end_date = self.get_date_range("last_2_weeks")
        data = self.get_google_form_data(start_date, end_date)
        self.plot_world_map_bubbles(data)

    def plot_last_3_weeks(self, event):
        start_date, end_date = self.get_date_range("last_3_weeks")
        data = self.get_google_form_data(start_date, end_date)
        self.plot_world_map_bubbles(data)

    def plot_last_month(self, event):
        start_date, end_date = self.get_date_range("last_month")
        data = self.get_google_form_data(start_date, end_date)
        self.plot_world_map_bubbles(data)

    def plot_last_2_months(self, event):
        start_date, end_date = self.get_date_range("last_2_months")
        data = self.get_google_form_data(start_date, end_date)
        self.plot_world_map_bubbles(data)

    def plot_last_3_months(self, event):
        start_date, end_date = self.get_date_range("last_3_months")
        data = self.get_google_form_data(start_date, end_date)
        self.plot_world_map_bubbles(data)

    def plot_last_6_months(self, event):
        start_date, end_date = self.get_date_range("last_6_months")
        data = self.get_google_form_data(start_date, end_date)
        self.plot_world_map_bubbles(data)

    def plot_last_year(self, event):
        start_date, end_date = self.get_date_range("last_year")
        data = self.get_google_form_data(start_date, end_date)
        self.plot_world_map_bubbles(data)

    def plot_sept_2024(self, event):
        start_date, end_date = self.get_date_range("since_sept_2024")
        data = self.get_google_form_data(start_date, end_date)
        self.plot_world_map_bubbles(data)

    def plot_all_usage(self, event):
        data = self.get_google_form_data()  # No date filtering
        self.plot_world_map_bubbles(data)

    def plot_google_form(self, event):
        data = self.get_google_form_data()  # No date filtering
        self.plot_world_map_bubbles(data)


def show_usage_stats_window(parent):
    """Function to launch Usage Stats window"""
    UsageStatsWindow(parent)


if __name__ == "__main__":
    app = wx.App()
    show_usage_stats_window(None)
    app.MainLoop()