from icalendar import Calendar, Event
import numpy as np
import pandas as pd
from pathlib import Path
from sys import argv
from zoneinfo import ZoneInfo

TZINFO = ZoneInfo("Europe/Paris")
JOUR_NUM = {
    "Lundi": 1,
    "Mardi": 2,
    "Mercredi": 3,
    "Jeudi": 4,
    "Vendredi": 5,
    "Samedi": 6,
    "Dimanche": 7
}
START_HOURS = {
    "Administratif": (9, 0),
    "Préparation marché bio": (9, 0),
    "Cultures": (9, 0),
    "Préparation boulange": (12, 0),
    "Pépinière": (9, 0),
    "Pain": (5, 30),
    "Récoltes marché bio (1/2) + marché Pontorson": (8, 30),
    "Chargement marché bio (1/2)": (8, 30),
    "Chargement marché Pontorson": (8, 30),
    "Marché bio": (14, 45),
    "Marché Pontorson": (6, 30),
    "Récolte marché Rocabey + AMAP": (9, 0),
    "Chargement marché Rocabey jeudi": (14, 30),
    "Chargement AMAP": (9, 0),
    "AMAP": (17, 0),
    "Marché Rocabey jeudi": (5, 15),
    "Récolte marché à la ferme + marché Rocabey": (8, 30),
    "Préparation marché à la ferme": (8, 30),
    "Préparation marché Rocabey samedi": (8, 30),
    "Marché à la ferme": (16,30),
    "Chargement marché Rocabey samedi": (8, 30),
    "Marché Rocabey samedi": (5, 15)
}
    
COLORS = {
    'Alexis': "black",
    'Christophe': "green",
    'Lalo': "red",
    'Marie': "fuchsia",
    'Mathieu M': "purple",
    'Mylène': "yellow",
    'Mathieu V': "maroon",
    'Jérôme': "olive",
    'Isabelle': "lime",
    'Seydou': "navy",
    'Aide 1': "teal",
    'Aide 2': "aqua"
}
CATEGORIES = ["Professionnel"]

READ_EXCEL_KWARGS = dict(
    index_col=[0, 1],
    usecols=range(14),
    skiprows=[1]
)

YEAR = 2026

def main():
    ics_root = Path(argv[-1])

    # Get calendars
    all_calendars = {}
    for ods_filepath in argv[1:-1]:
        file_calendars = get_calendars_from_file(ods_filepath)
        for name, cal in file_calendars.items():
            if name not in all_calendars:
                all_calendars[name] = []
            all_calendars[name].append(cal)

    merged_calendars = {}
    for name, cals in all_calendars.items():
        # Merge calendars
        merged_cal = merge_calendars(cals)
        merged_calendars[name] = merged_cal

        # Print people calendar
        print('\n', name, ":")
        print(display(merged_cal))

        # Write people calendar
        ics_filepath = Path(
            ics_root.parent, f"{ics_root.stem}_{name}.ics")
        f = open(ics_filepath, 'wb')
        f.write(merged_cal.to_ical())
        f.close()

def get_calendars_from_file(
        ods_filepath, read_excel_kwargs=READ_EXCEL_KWARGS, year=YEAR,
        tzinfo=TZINFO, start_hours=START_HOURS, colors=COLORS,
        categories=CATEGORIES):
    """Get calendar from file."""
    df = pd.read_excel(ods_filepath, **read_excel_kwargs)
    week = pd.read_excel(ods_filepath, usecols=[0], nrows=1).squeeze()

    calendars = {}
    for num, (name, df_name) in enumerate(df.items()):
        cal_name = Calendar()
        cal_name.color = colors[name]

        for jour, df_new in df_name.groupby(level=0):
            df_name_jour = df_new.droplevel(0)

            duration = pd.Timedelta(hours=df_name_jour.sum())

            if duration.value > 0:
                date = pd.Timestamp.fromisocalendar(
                    year, week, JOUR_NUM[jour])
                start_hour = get_start_hour(df_name_jour, start_hours)
                start = pd.Timestamp(date.year, date.month, date.day,
                                     *start_hour, tzinfo=tzinfo)
                end = start + duration
                summary = ", ".join(
                    [f"{atelier} ({int(heures)})"
                     for atelier, heures in df_name_jour.items()
                     if heures > 0]
                )
                description = "\n".join(
                    [f"- {atelier}: {int(heures)} h"
                     for atelier, heures in df_name_jour.items()
                     if heures > 0]
                )

                event = Event()
                event.color = colors[name]
                event.categories = categories
                event.start = start
                event.end = end
                event["summary"] = summary
                event["description"] = description

                cal_name.add_component(event)
        
        calendars[name] = cal_name

    return calendars

def merge_calendars(cals):
    """Merge calendars."""
    merged_cal = Calendar()

    for cal in cals:
        for component in cal.walk():
            if component.name != "VCALENDAR":
                merged_cal.add_component(component)

    return merged_cal

def display(cal):
    """Display calendar."""
    return cal.to_ical().decode("utf-8").replace('\r\n', '\n').strip()

def get_start_hour(df_name_jour, start_hours):
    """Get start hour."""
    start_minutes = np.min([
        start_hours[atelier][0] * 60 + start_hours[atelier][1]
        for atelier, heures in df_name_jour.items()
        if heures > 0
    ])
    start_hour = (start_minutes // 60, np.mod(start_minutes, 60))

    return start_hour

if __name__ == "__main__":
    main()

