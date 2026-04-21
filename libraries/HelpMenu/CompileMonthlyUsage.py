# KherveFitting - XPS Data Analysis Software
# Copyright (C) 2024-2026 Gwilherm Kerherve <g.kerherve@ic.ac.uk>
#
# KherveFitting is dual-licensed:
#   - GNU GPL v3.0 (see LICENSE-GPL.txt) for open-source use
#   - Commercial Licence (see LICENSE-COMMERCIAL.txt) for proprietary use
# SPDX-License-Identifier: GPL-3.0-only OR LicenseRef-KherveFitting-Commercial

"""
Monthly Usage Data Compiler
This program reads usage tracking data from Google Sheets and compiles it into monthly summaries.
Outputs Python dictionary code that you can copy-paste into UsageAnalytics.py
"""

import pandas as pd
import json
from datetime import datetime
from collections import defaultdict
import sys
import os

# Google Sheets CSV export URL
GOOGLE_SHEETS_ID = '1mP23kwwg-H1Gx5ppp2YQNv6X1lYhJj2B6_PagQ0yABw'
GOOGLE_SHEETS_CSV_URL = f'https://docs.google.com/spreadsheets/d/{GOOGLE_SHEETS_ID}/export?format=csv&gid=213903690'


def parse_timestamp(timestamp_str):
    """Parse timestamp string to datetime object"""
    try:
        return datetime.strptime(timestamp_str, "%m/%d/%Y %H:%M:%S")
    except:
        try:
            return datetime.strptime(timestamp_str, "%d/%m/%Y %H:%M:%S")
        except:
            try:
                return datetime.strptime(timestamp_str, "%Y-%m-%d %H:%M:%S")
            except:
                return None


def compile_month_data(df, year, month):
    """Compile usage data for a specific month - countries only"""
    countries = {}

    for idx, row in df.iterrows():
        timestamp = parse_timestamp(row['Timestamp'])
        if timestamp and timestamp.year == year and timestamp.month == month:
            # Get country directly from Country column
            country = str(row['Country']).strip()

            # Extract just the country name (before comma if present)
            # "United States, Kentucky" -> "United States"
            # "France, Île-de-France" -> "France"
            if ',' in country:
                country = country.split(',')[0].strip()

            # Handle empty or NaN values
            if country == '' or pd.isna(country) or country == 'nan':
                country = 'Unknown'

            if country in countries:
                countries[country] += 1
            else:
                countries[country] = 1

    return countries


def get_all_months_in_data(df):
    """Find all unique year-month combinations in the dataset"""
    months = set()

    for idx, row in df.iterrows():
        timestamp = parse_timestamp(row['Timestamp'])
        if timestamp:
            months.add((timestamp.year, timestamp.month))

    return sorted(list(months))


def compile_all_months():
    """Main function to automatically compile all months"""

    print("Automatic Monthly Usage Data Compiler")
    print("=" * 70)
    print("Loading data from Google Sheets...")

    try:
        df = pd.read_csv(GOOGLE_SHEETS_CSV_URL)
        print(f"Successfully loaded {len(df)} records from Google Sheets")
    except Exception as e:
        print(f"Error loading data from Google Sheets: {e}")
        return

    all_months = get_all_months_in_data(df)

    if not all_months:
        print("No valid data found in the spreadsheet")
        return

    print(f"\nFound data for {len(all_months)} months")
    print("=" * 70)

    # Build the complete dictionary
    monthly_data = {}

    for year, month in all_months:
        key = f"{year}_{month:02d}"
        countries = compile_month_data(df, year, month)
        monthly_data[key] = countries

        total = sum(countries.values())
        print(f"  {month:02d}/{year}: {total} events, {len(countries)} countries")

    # Generate Python code
    print("\n" + "=" * 70)
    print("COPY THE CODE BELOW INTO UsageAnalytics.py")
    print("=" * 70)
    print()

    # Write to output file
    output_file = "monthly_usage_dictionary.txt"
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("# Compressed monthly usage data - Updated: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "\n")
        f.write("MONTHLY_USAGE_DATA = {\n")

        for key in sorted(monthly_data.keys()):
            countries = monthly_data[key]
            f.write(f"    '{key}': {{\n")

            # Sort countries by count (descending)
            sorted_countries = sorted(countries.items(), key=lambda x: x[1], reverse=True)
            for country, count in sorted_countries:
                f.write(f"        '{country}': {count},\n")

            f.write("    },\n")

        f.write("}\n")

    # Also print to console
    with open(output_file, 'r', encoding='utf-8') as f:
        print(f.read())

    print("=" * 70)
    print(f"✓ Dictionary saved to: {output_file}")
    print("✓ Copy the dictionary above and paste it into UsageAnalytics.py")
    print("=" * 70)


if __name__ == "__main__":
    compile_all_months()