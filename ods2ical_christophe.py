"""Convertit un planning ODS (un onglet "Semaine N" par semaine, avec des
heures par personne et par opération) en un ensemble de fichiers iCalendar
(.ics), un par personne.

Usage :
    python ods2ical_christophe.py

La configuration (chemin du fichier source, année, racine des fichiers de
sortie, couleurs par personne, heures de début par opération…) est lue
depuis ``config.toml`` placé à côté du script.
"""
from icalendar import Calendar, Event
import numpy as np
import pandas as pd
from pathlib import Path
import tomllib
from zoneinfo import ZoneInfo

CONFIG_FILEPATH = Path("config.toml")

JOUR_NUM = {
    "lundi": 1,
    "mardi": 2,
    "mercredi": 3,
    "jeudi": 4,
    "vendredi": 5,
    "samedi": 6,
    "dimanche": 7,
}


def _normalize(s):
    """Normalise une chaîne (strip + casefold) pour des comparaisons
    robustes aux variations de casse et d'espaces dans les libellés du
    tableur."""
    return s.strip().casefold() if isinstance(s, str) else s


def _lookup_normalized(mapping, key):
    """Recherche `key` dans `mapping` avec normalisation côté clés ET côté
    valeur cherchée. Renvoie la valeur trouvée ou KeyError."""
    norm_key = _normalize(key)
    for k, v in mapping.items():
        if _normalize(k) == norm_key:
            return v
    raise KeyError(key)

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
        # Ne traiter que les feuilles "Semaine N" où N est un entier (ISO week).
        # On exclut les modèles "Semaine paire" / "Semaine impaire" et les
        # feuilles vides (Feuille15, …).
        sheet_lower = sheet_name.lower()
        if not sheet_lower.startswith("semaine "):
            continue
        suffix = sheet_lower[len("semaine "):].strip()
        if not suffix.isdigit():
            continue
        if df0_raw.empty:
            continue
        # Le numéro de semaine ISO est dans le header de la 2ème colonne ;
        # on vérifie qu'il est cohérent avec le nom de feuille.
        week_header = df0_raw.columns[1]
        if not isinstance(week_header, (int, np.integer)):
            continue
        week = int(week_header)
        if week != int(suffix):
            print(
                f"Attention : feuille {sheet_name!r} a un numéro de semaine "
                f"en en-tête ({week}) qui ne correspond pas au nom de feuille "
                f"({suffix}). Utilisation de la valeur de l'en-tête."
            )

        # Structure du fichier (après lecture pandas, header sur ligne 0) :
        #   col 0 : jour de la semaine (Lundi, Mardi, ...)
        #   col 1 : numéro du jour du mois (non utilisé)
        #   col 2 : opération (Unnamed: 2)
        #   col 3..N-3 : colonnes par personne
        #   col N-2 : TOTAL
        #   col N-1 : TOTAL Jour
        # Dans le tableur, la colonne "jour" n'a une valeur que sur la
        # 1ère ligne de chaque jour (cellules fusionnées) ; les lignes
        # suivantes sont NaN. On forward-fill avant l'indexation.
        jour_col = df0_raw.columns[0]
        op_col = df0_raw.columns[2]
        df0_raw = df0_raw.copy()
        df0_raw[jour_col] = df0_raw[jour_col].ffill()
        df0 = df0_raw.set_index([jour_col, op_col]).iloc[:, 1:-2]

        file_calendars = get_calendars_from_frame(df0, week, config)

        for name, cal in file_calendars.items():
            if name not in all_calendars:
                all_calendars[name] = []
            all_calendars[name].append(cal)

    return all_calendars

def get_calendars_from_frame(df0, week, config):
    """Get calendar for a single week sheet."""
    calendars = {}

    # La 1ère ligne d'un onglet de semaine est une ligne TOTAL ("TOTAL", "dates")
    # avec des valeurs agrégées : on la supprime pour ne garder que les
    # (jour, opération) réels.
    df = df0[df0.index.get_level_values(0) != "TOTAL"]

    # Certaines cellules contiennent 'o' (présence sans durée) ; on les
    # traite comme 0.
    df = df.where(df != 'o', 0)
    # Forcer les colonnes à du numérique (les NaN restent NaN, les sommes
    # ignorent NaN).
    df = df.apply(pd.to_numeric, errors='coerce')

    for num, (name, df_name) in enumerate(df.items()):
        cal_name = Calendar()
        try:
            color = _lookup_normalized(config["people"]["colors"], name)
        except KeyError as e:
            raise KeyError(
                f"Personne absente de [people].colors : {name!r}. "
                f"Compléter config.toml."
            ) from e
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
                    JOUR_NUM[_normalize(jour)])
                start_hour = get_start_hour(
                    df_name_jour, config["operations"]["start_hours"])
                dtstart = pd.Timestamp(date.year, date.month, date.day,
                                       *start_hour, tzinfo=tzinfo)
                dtend = dtstart + duration
                summary = ", ".join(
                    [f"{operation} ({_format_hours(heures)})"
                     for operation, heures in df_name_jour.items()
                     if heures > 0]
                )
                description = "\n".join(
                    [f"- {operation}: {_format_hours(heures)} h"
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
    """Get start hour as the earliest start time across operations of the day."""
    start_minutes_list = []
    for operation, heures in df_name_jour.items():
        if heures > 0:
            try:
                hh, mm = _parse_hhmm(_lookup_normalized(start_hours, operation))
            except KeyError as e:
                raise KeyError(
                    f"Opération absente de [operations].start_hours : "
                    f"{operation!r}. Compléter config.toml."
                ) from e
            start_minutes_list.append(hh * 60 + mm)
    start_minutes = int(np.min(start_minutes_list))
    return (start_minutes // 60, start_minutes % 60)


def _parse_hhmm(value):
    """Parse une heure au format 'H:MM' ou 'HH:MM' en (h, m)."""
    if isinstance(value, (list, tuple)) and len(value) == 2:
        return int(value[0]), int(value[1])
    parts = str(value).split(":")
    return int(parts[0]), int(parts[1])


def _format_hours(h):
    """Formate une durée en heures : '7' si entier, '0.75' sinon."""
    if h == int(h):
        return str(int(h))
    # On supprime les zéros de fin (ex: 1.50 -> 1.5).
    return f"{h:g}"

if __name__ == "__main__":
    main()

