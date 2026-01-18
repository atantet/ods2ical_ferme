from icalendar import Calendar, Event
import pandas as pd
from pathlib import Path
from sys import argv
from zoneinfo import ZoneInfo

def display(cal):
    return cal.to_ical().decode("utf-8").replace('\r\n', '\n').strip()

TZINFO = ZoneInfo("Europe/Paris")
START_HOUR = 8
JOUR_NUM = {
    "Lundi": 1,
    "Mardi": 2,
    "Mercredi": 3,
    "Jeudi": 4,
    "Vendredi": 5,
    "Samedi": 6,
    "Dimanche": 7
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

ODS_FILEPATH = Path("..", "emploi_du_temps_prévisionnel_paire.ods")
INDEX_COL = [0, 1]
USECOLS = range(14)
SKIPROWS = [1]

YEAR = 2026

df = pd.read_excel(
    ODS_FILEPATH,
    index_col=INDEX_COL, usecols=USECOLS, skiprows=SKIPROWS
)
week = pd.read_excel(ODS_FILEPATH, usecols=[0], nrows=1).squeeze()

for num, (name, df_name) in enumerate(df.items()):
    ICS_FILEPATH = Path(ODS_FILEPATH.parent, 
                        ODS_FILEPATH.stem + "_" + name + ".ics")
    cal = Calendar()
    cal.color = COLORS[name]

    for jour, df_new in df_name.groupby(level=0):
        df_name_jour = df_new.droplevel(0)

        duration = pd.Timedelta(hours=df_name_jour.sum())

        if duration.value > 0:
            date = pd.Timestamp.fromisocalendar(YEAR, week, JOUR_NUM[jour])
            start = pd.Timestamp(date.year, date.month, date.day,
                                 START_HOUR, tzinfo=TZINFO)
            end = start + duration
            summary = ", ".join(
                [f"{atelier} ({int(heures)})"
                 for atelier, heures in df_name_jour.items() if heures > 0]
            )
            description = "\n".join(
                [f"- {atelier}: {int(heures)} h"
                 for atelier, heures in df_name_jour.items() if heures > 0]
            )

            event = Event()
            event.color = COLORS[name]
            event.categories = CATEGORIES
            event.start = start
            event.end = end
            event["summary"] = summary
            event["description"] = description

            cal.add_component(event)
            
    print('\n', name, ":")
    print(display(cal))

    f = open(ICS_FILEPATH, 'wb')
    f.write(cal.to_ical())
    f.close()
