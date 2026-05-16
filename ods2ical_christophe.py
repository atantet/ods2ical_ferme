"""
python ods2ical_emploi_du_temps_prévisionnel.py ../emploi_du_temps_prévisionnel_paire.ods ../emploi_du_temps_prévisionnel_impaire.ods ../test.ics
"""
from icalendar import Calendar, Event, Timezone
import numpy as np
import pandas as pd
from pathlib import Path
import tomllib
from zoneinfo import ZoneInfo

CONFIG_FILEPATH = Path("config.toml")

JOUR_NUM = {
    "LUNDI": 1,
    "MARDI": 2,
    "MERCREDI": 3,
    "JEUDI": 4,
    "VENDREDI": 5,
    "SAMEDI": 6,
    "DIMANCHE": 7
}

def main():
    config = load_config()
    ics_root = Path(config["destination"]["ics_root"])

    all_calendars = get_calendars_from_file(config)

    merged_calendars = {}
    for name, cals in all_calendars.items():
        # Merge calendars
        merged_cal = merge_calendars(cals, config)
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


def get_calendars_from_file(config):
    """Get calendars for different weeks from sheets of ODS file."""
    d = pd.read_excel(config["source"]["ods_filepath"], sheet_name=None)

    # Get calendars
    all_calendars = {}
    for sheet_name, df0_raw in d.items():
        if sheet_name.title()[:7] == "Semaine":
            df0 = df0_raw.set_index(["sem", "TRAVAUX"]).iloc[:, 3:-1]
            file_calendars = get_calendars_from_frame(df0, config)

            for name, cal in file_calendars.items():
                if name not in all_calendars:
                    all_calendars[name] = []
                all_calendars[name].append(cal)

    return all_calendars

def get_calendars_from_frame(df0, config):
    """Get calendar from file."""
    calendars = {}
    
    week = df0.index[config["source"]["week_row"]][0]
    df = df0.drop(labels=df0.index[config["source"]["week_row"]])

    # Vérification de la qualité
    df = df.where(df != 'o', 0)

    for num, (name, df_name) in enumerate(df.items()):
        cal_name = Calendar()
        color = config["people"]["colors"][name]
        cal_name.color = color

        for jour, df_new in df_name.groupby(level=0):
            df_name_jour = df_new.droplevel(0)

            duration = pd.Timedelta(hours=df_name_jour.sum())

            if duration.value > 0:
                tzinfo = ZoneInfo(config["calendar"]["timezone"])
                dtstamp = pd.Timestamp.now(tz=tzinfo)
                uid = f'{week}/{name}/{jour}/{dtstamp}'
                date = pd.Timestamp.fromisocalendar(
                    config["source"]["year"], week,
                    JOUR_NUM[jour.strip(" ")])
                start_hour = get_start_hour(
                    df_name_jour, config["operations"]["start_hours"])
                dtstart = pd.Timestamp(date.year, date.month, date.day,
                                       *start_hour, tzinfo=tzinfo)
                dtend = dtstart + duration
                summary = ", ".join(
                    [f"{operation} ({int(heures)})"
                     for operation, heures in df_name_jour.items()
                     if heures > 0]
                )
                description = "\n".join(
                    [f"- {operation}: {int(heures)} h"
                     for operation, heures in df_name_jour.items()
                     if heures > 0]
                )

                event = Event(uid=uid)
                
                event.add('dtstamp', dtstamp)
                event.add('dtstart', dtstart)
                event.add('dtend', dtend)
                event.add('summary', summary)
                event.add('description', description)
                event.add('categories', config["events"]["categories"])
                event.add('color', color)

                cal_name.add_component(event)

        cal_name.add_missing_timezones()
        
        calendars[name] = cal_name

    return calendars

def load_config():
    with open(CONFIG_FILEPATH, "rb") as f:
        config = tomllib.load(f)

    return config

def merge_calendars(cals, config):
    """Merge calendars."""
    merged_cal = Calendar()

    # Some properties are required to be compliant.
    merged_cal.add('prodid', config["calendar"]["prodid"])
    merged_cal.add('version', config["calendar"]["version"])

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
        start_hours[operation][0] * 60 + start_hours[operation][1]
        for operation, heures in df_name_jour.items()
        if heures > 0
    ])
    start_hour = (start_minutes // 60, np.mod(start_minutes, 60))

    return start_hour

if __name__ == "__main__":
    main()

