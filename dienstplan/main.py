#!/usr/bin/env python3
import pandas as pd
import sys

#exec(open("main.py").read())

if len(sys.argv) != 4:
    print("Usage: main.py file.xlsx startYear 'ARTIST NAME'")
    print(sys.argv)
    exit(1)

[script_name, excel_file, year, artist] = sys.argv

year = int(year)


def parse(xl, sheet_name):
    print(sheet_name)
    worksheet = xl.parse(sheet_name, header=None)  # read a specific worksheet to DataFrame

    calendar_week = int(worksheet.iloc[3, 0].split(' ')[1])
    season_week = int(sheet_name.split('W')[1])
    add_to_year = 1 if season_week > calendar_week else 0

    worksheet_unmerged = worksheet.copy()
    worksheet_unmerged.iloc[2] = worksheet_unmerged.iloc[2].ffill()
    worksheet_timeslots = worksheet_unmerged.dropna(axis=1, subset=3).set_index(0)

    #worksheet_timeslots.iloc[2].ffill(inplace=True)
    #e = worksheet_timeslots.set_index(0)
    time_slots = worksheet_timeslots.iloc[2] + '.' + str(year + add_to_year) + ' ' + worksheet_timeslots.iloc[3]
    worksheet_timeslots.columns = pd.to_datetime(time_slots, format="%d.%m.%Y %H.%M", exact=False)
    artist_slots = worksheet_timeslots.dropna(axis=1, subset=[artist])
    slot_content = artist_slots.iloc[[4, 5]].transpose()
    slot_content.columns = ['type', 'show']
    return slot_content


xl = pd.ExcelFile(excel_file)

week_sheet_names = list(filter(lambda n: n.startswith('W'), xl.sheet_names))
all_slots = pd.concat(map(lambda s: parse(xl, s), week_sheet_names))

show_days = pd.Series(all_slots.index).dt.normalize()
cal_days = pd.date_range(show_days.min(), show_days.max())

free_days = pd.DataFrame({'day': cal_days[~cal_days.isin(show_days)]}).set_index('day')

cal = pd.concat([all_slots, free_days]).sort_index()
cal_dates = pd.Series(cal.index).dt

# cal['show'].fillna(value='FREE', inplace=True)

mcal = cal.copy()
mcal.index = pd.MultiIndex.from_frame(pd.DataFrame({
    'day': cal_dates.day_name().str[0:3] + ' ' + cal_dates.strftime('%d.%m.%y'),
    'time': cal_dates.strftime('%H:%M').replace('00:00', '')
}))

mcal.to_excel('results.xlsx')
