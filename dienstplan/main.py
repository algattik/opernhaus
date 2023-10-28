#!/usr/bin/env python3
import pandas as pd
import sys
from collections import defaultdict

#exec(open("main.py").read())

if len(sys.argv) != 3:
    print("Usage: main.py file.xlsx startYear")
    print(sys.argv)
    exit(1)

[script_name, excel_file, year] = sys.argv

year = int(year)


def parse(sheet_name, worksheet, artist):
    print(sheet_name)
    print(sheet_name)
    print(sheet_name)
    print(sheet_name)
    print(worksheet)
    print(artist)
    calendar_week = int(worksheet.iloc[3, 0].split(' ')[1])
    season_week = int(sheet_name.split('W')[1])
    add_to_year = 1 if season_week > calendar_week else 0

    worksheet_unmerged = worksheet.copy()
    worksheet_unmerged.iloc[2] = worksheet_unmerged.iloc[2].ffill()
    worksheet_timeslots = worksheet_unmerged.dropna(axis=1, subset=3).set_index(0)

    time_slots = worksheet_timeslots.iloc[2] + '.' + str(year + add_to_year) + ' ' + worksheet_timeslots.iloc[3]
    worksheet_timeslots.columns = pd.to_datetime(time_slots, format="%d.%m.%Y %H.%M", exact=False)
    artist_slots = worksheet_timeslots.dropna(axis=1, subset=[artist])
    slot_content = artist_slots.iloc[[4, 5]].transpose()
    slot_content.columns = ['type', 'show']
    return slot_content


xl = pd.ExcelFile(excel_file)

week_sheet_names = list(filter(lambda n: n.startswith('W'), xl.sheet_names))
a = list(map(lambda w: [w, xl.parse(w, header=None)], week_sheet_names))

artist_weeks = defaultdict(list)
for s in a:
    week = s[0]
    sheet = s[1]
    for artist in sheet[sheet.isin(['D']).any(axis=1)][0]:
        if artist:
            artist_weeks[artist].append(week)

# artists = ['SACRAMENTO NUNES Filipa Magarida']

with pd.ExcelWriter('result.xlsx', engine='xlsxwriter') as writer:

    for artist, weeks in artist_weeks.items():
        print(artist)
        print(weeks)

        all_slots = pd.concat(map(lambda s: parse(s[0], s[1], artist), filter(lambda w: w[0] in weeks, a)))

        show_days = pd.Series(all_slots.index).dt.normalize()
        cal_days = pd.date_range(show_days.min(), show_days.max())

        free_days = pd.DataFrame({'day': cal_days[~cal_days.isin(show_days)]}).set_index('day')

        cal = pd.concat([all_slots, free_days]).sort_index()
        cal_dates = pd.Series(cal.index).dt

        multilevel_cal = cal.copy()
        multilevel_cal.index = pd.MultiIndex.from_frame(pd.DataFrame({
            'day': cal_dates.day_name().str[0:3] + ' ' + cal_dates.strftime('%d.%m.%y'),
            'time': cal_dates.strftime('%H:%M').replace('00:00', '')
        }))

        multilevel_cal.to_excel(writer, sheet_name=' '.join(artist.split()[:2]))
