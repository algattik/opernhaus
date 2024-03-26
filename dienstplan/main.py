#!/usr/bin/env python3

import sys
from datetime import datetime, timedelta
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font
from collections import defaultdict

# Global constants
DATE_FORMAT = "%d.%m.%Y"
COLOR1 = PatternFill(start_color="00DDFFDD", fill_type="solid")
COLOR2 = PatternFill(start_color="00FFDDDD", fill_type="solid")

# Global styles
THIN_BORDER = Border(top=Side(style='thin'), bottom=Side(style='thin'))
THICK_BORDER = Border(top=Side(style='thick'), bottom=Side(style='thin'))
BOLD_FONT = Font(bold=True)
TOP_ALIGNMENT = Alignment(horizontal="left", vertical="top", wrap_text=True)

# Workbook initialization
output_workbook = Workbook()
first_sheet = output_workbook.active

# Utility function to iterate over days
def iterate_days(start_date, end_date):
    current_date = start_date
    while current_date <= end_date:
        yield current_date
        current_date += timedelta(days=1)

# Command-line argument handling
if len(sys.argv) != 2:
    print("Usage: main.py file.xlsx")
    print(sys.argv)
    exit(1)

[script_name, input_excel_file] = sys.argv
shifts_data = defaultdict(lambda: defaultdict(dict))
artists_data = defaultdict(int)

# Load workbook
input_workbook = load_workbook(filename=input_excel_file)

# Process sheets
for sheet in input_workbook:
    if not sheet.title.startswith("W"):
        continue

    print(sheet.title)

    season_info = sheet["A1"].value
    season_year = int(season_info.split(" ")[1].split("/")[0])

    calendar_week_info = sheet["A4"].value
    calendar_week_value = int(calendar_week_info.split(" ")[1])
    season_week = int(sheet["A6"].value)
    if season_week > calendar_week_value:
        season_year += 1

    days_row = sheet[3]
    hours_row = sheet[4]
    type_row = sheet[5]
    show_row = sheet[6]
    points_row = sheet[7]

    days_cell_range = sheet['D3':'BG3'][0]
    previous_date = None
    artist_days = []
    for cell in days_cell_range:
        day_value = cell.value
        if day_value is not None:
            day_parts = day_value.split(" ")
            current_date = datetime.strptime(f"{day_parts[1]}.{season_year}", DATE_FORMAT)
            previous_date = current_date
        artist_days.append(previous_date)

    for row in sheet.iter_rows(min_row=9, min_col=1, max_col=60):
        artist_name = row[0].value
        for cell in row:
            if cell.value not in ['D', 'DK', 'E', 'EK']: #D = Dienst, DK = Dienst Krank, #E = Ersatz, EK = Ersatz Krank. F = Fiktive Dienst (not for planning)
                continue
            shift_hours = str(hours_row[cell.column - 1].value)
            shift_day = artist_days[cell.column - 4]
            shifts_data[shift_day.isoformat()][artist_name][shift_hours] = [type_row[cell.column - 1].value, show_row[cell.column - 1].value, points_row[cell.column - 1].value]
            artists_data[artist_name] += 1

min_shift_date = datetime.fromisoformat(min(shifts_data.keys()))
max_shift_date = datetime.fromisoformat(max(shifts_data.keys()))

# Function to add a row to a worksheet
def add_row(worksheet, row_data, fill_color):
    worksheet.append(row_data)
    for coln in range(0, len(row_data)):
        cell = worksheet.cell(row=worksheet.max_row, column=coln+1)
        cell.alignment = TOP_ALIGNMENT
        if coln == 0:
            cell.fill = fill_color

# Process artists and create sheets
for artist_name in sorted(artists_data, key=lambda k: -len(k)):
    current_worksheet = output_workbook.create_sheet(title=' '.join(artist_name.split()[:2]))
    for shift_day in iterate_days(min_shift_date, max_shift_date + timedelta(days=90)):
        artist_shifts = shifts_data[shift_day.isoformat()]
        shift_day_date = shift_day.date()
        if artist_name not in artist_shifts:
            add_row(current_worksheet, [shift_day_date], COLOR1)
            continue
        sorted_shifts = sorted(artist_shifts[artist_name])

        row_hours = []
        row_types = []
        row_info = []
        row_points = 0
        for shift_hours in sorted_shifts:
            [shift_type, show_info, points] = artist_shifts[artist_name][shift_hours]
            row_hours.append(shift_hours.replace('.', ':'))
            row_types.append(shift_type or 'VST')
            row_info.append(show_info or '')
            if points:
                row_points += points

        def merge(row_data):
            return '\n'.join(row_data)

        add_row(current_worksheet, [shift_day_date, merge(row_hours), merge(row_types), merge(row_info), row_points], COLOR2)

    def contains_substring(input_string, substrings):
        for substring in substrings:
            if substring in input_string:
                return True
        return False

    for row_data in current_worksheet.iter_rows():
        if row_data[0].value is None:
            continue
        shift_day_is_monday = row_data[0].value.weekday() == 0
        for cell in row_data:
            cell.border = THICK_BORDER if shift_day_is_monday else THIN_BORDER
        if contains_substring(row_data[2].value or '', ['VST', 'GP', 'WA', 'OHP', 'Pr-A', 'Pr-B']):
            row_data[2].font = BOLD_FONT
            row_data[3].font = BOLD_FONT

    current_worksheet.column_dimensions["A"].width = 10
    current_worksheet.column_dimensions["B"].width = 7
    current_worksheet.column_dimensions["C"].width = 5
    current_worksheet.column_dimensions["D"].width = 20

# Save workbook
output_workbook.remove(first_sheet)
output_workbook.save("dienstplan.xlsx")
