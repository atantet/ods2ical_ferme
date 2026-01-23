"""
python ods2ical_emploi_du_temps_prévisionnel.py ../emploi_du_temps_prévisionnel_paire.ods ../emploi_du_temps_prévisionnel_impaire.ods ../test.ics
"""
from icalendar import Calendar, Event, Timezone
import numpy as np
import pandas as pd
from pathlib import Path
from sys import argv
from zoneinfo import ZoneInfo


TZINFO = ZoneInfo("Europe/Paris")
JOUR_NUM = {
    "LUNDI": 1,
    "MARDI": 2,
    "MERCREDI": 3,
    "JEUDI": 4,
    "VENDREDI": 5,
    "SAMEDI": 6,
    "DIMANCHE": 7
}
START_HOURS = {
    "administratif": (9, 0),
    "administratif+autres": (9, 0),
    "divers": (9, 0),
    "Divers": (9, 0),
    "(récoltes)": (9, 0),
    "prépa marché bio": (9, 0),
    "Cultures": (9, 0),
    "prépa pain": (12, 0),
    "Pépinière": (9, 0),
    "pain ou administratif": (5, 30),
    "pain": (5, 30),
    "récoltes Marchés": (8, 30),
    "récolte (MB) + M": (8, 30),
    "betteraves": (8, 30),
    "chargement Mbio": (8, 30),
    "chargement Mpon": (9, 0),
    "crgt Mpon et rocabey": (9, 0),
    "marché bio (Mbio)": (14, 45),
    "retour Mbio administratif": (14, 45),
    "retour Mbio": (14, 45),
    "marché pontorson (Mpon)": (6, 30),
    "Déchargement (Mpon)": (14, 30),
    "Déchargement Mbio": (9, 0),
    "récolte Mroc et amap": (9, 0),
    "chargement Mroc + com": (14, 30),
    "chargement amap": (9, 0),
    "amap": (17, 0),
    "marché Rocabey (Mroc)": (5, 15),
    "Déchargement amap": (9, 0),
    "(récolte) divers": (9, 0),
    "récolte Mferme + Mroc": (8, 30),
    "prépa yourte": (8, 30),
    "Préparation marché Rocabey samedi": (8, 30),
    "marché à la ferme (Mroc)": (16,30),
    "chargement Mroc": (8, 30),
    "marché Rocabey": (5, 15),
    "Aller, chrgt, retour Mroc et déchargement ": (5, 15)
}
    
COLORS = {
    'Alexis': "blue",
    'Christophe': "green",
    'Lalo': "red",
    'Marie': "fuchsia",
    'Mathieu M': "purple",
    'Mylène': "yellow",
    'Mathieu V': "maroon",
    'Jérôme': "olive",
    'Isa': "lime",
    'Seydou': "navy",
    'Léandre': "teal",
    'Clothilde': "aqua",
    'Clément': "teal",
    'Elsa': "aqua",
    'Liam': "teal",
    'Josepha': "aqua",
    'Gwendal': "teal",
    "E’ouann": "aqua",
    'Aide 1': "teal",
    'Aide 2': "aqua"
}
CATEGORIES = ["Professionnel"]

READ_EXCEL_KWARGS = dict(
    index_col=[0, 1],
    usecols=[0, 2] + np.arange(5, 23).tolist()
)

YEAR = 2026

def main():
    ics_root = Path(argv[-1])

    all_calendars = get_calendars_from_file(argv[1])

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

    return


def get_calendars_from_file(
        ods_filepath, read_excel_kwargs=READ_EXCEL_KWARGS, year=YEAR,
        tzinfo=TZINFO, start_hours=START_HOURS, colors=COLORS,
        categories=CATEGORIES):
    """Get calendars for different weeks from sheets of ODS file."""
    d = pd.read_excel(ods_filepath, sheet_name=None, **read_excel_kwargs)

    # Get calendars
    all_calendars = {}
    for sheet_name, df0 in d.items():
        file_calendars = get_calendars_from_frame(
            df0, year=YEAR,
            tzinfo=TZINFO, start_hours=START_HOURS, colors=COLORS,
            categories=CATEGORIES)

        for name, cal in file_calendars.items():
            if name not in all_calendars:
                all_calendars[name] = []
            all_calendars[name].append(cal)

    return all_calendars

def get_calendars_from_frame(
        df0, year=YEAR,
        tzinfo=TZINFO, start_hours=START_HOURS, colors=COLORS,
        categories=CATEGORIES, week_row=0):
    """Get calendar from file."""
    calendars = {}
    
    week = df0.index[week_row][0]
    df = df0.drop(labels=df0.index[week_row])

    # Vérification de la qualité
    df = df.where(df != 'o', 0)

    for num, (name, df_name) in enumerate(df.items()):
        cal_name = Calendar()
        cal_name.color = colors[name]

        for jour, df_new in df_name.groupby(level=0):
            df_name_jour = df_new.droplevel(0)

            duration = pd.Timedelta(hours=df_name_jour.sum())

            if duration.value > 0:
                dtstamp = pd.Timestamp.now(tz=tzinfo)
                uid = f'{week}/{name}/{jour}/{dtstamp}'
                date = pd.Timestamp.fromisocalendar(
                    year, week, JOUR_NUM[jour.strip(" ")])
                start_hour = get_start_hour(df_name_jour, start_hours)
                dtstart = pd.Timestamp(date.year, date.month, date.day,
                                       *start_hour, tzinfo=tzinfo)
                dtend = dtstart + duration
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
                color = colors[name]

                event = Event(uid=uid)
                
                event.add('dtstamp', dtstamp)
                event.add('dtstart', dtstart)
                event.add('dtend', dtend)
                event.add('summary', summary)
                event.add('description', description)
                event.add('categories', categories)
                event.add('color', color)

                cal_name.add_component(event)

        cal_name.add_missing_timezones()
        
        calendars[name] = cal_name

    return calendars

def merge_calendars(cals):
    """Merge calendars."""
    merged_cal = Calendar()

    # Some properties are required to be compliant.
    merged_cal.add('prodid', '-//atantet//ods2ical_christophe/')
    merged_cal.add('version', '2.0')

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

