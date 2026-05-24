# ods2ical_ferme

Conversion d'un planning de ferme tenu dans un tableur **ODS** (LibreOffice
Calc) en un ensemble de calendriers **iCalendar** (`.ics`), un par personne,
importables dans n'importe quel agenda (LibreOffice Calendar, Thunderbird,
Google Calendar, Apple Calendar, etc.).

Chaque journée travaillée par une personne devient un événement, dont :

- l'**heure de début** est la plus matinale parmi les opérations de la
  journée (configurable par opération) ;
- la **durée** est la somme des heures déclarées dans le tableur ;
- le **résumé** liste les opérations effectuées avec leur durée ;
- la **couleur** est celle attribuée à la personne dans la configuration.

## Installation

L'environnement Python est décrit dans `environment.yaml` :

```bash
conda env create -f environment.yaml
conda activate ods2ical_ferme
```

Dépendances : Python ≥ 3.11 (pour `tomllib` de la stdlib), `pandas`,
`numpy`, `odfpy` (lecture des `.ods`), `icalendar`.

## Format du fichier ODS attendu

Le script `ods2ical_christophe.py` lit un classeur ODS contenant une feuille
par semaine, nommée `Semaine N` (`N` entier, ex. `Semaine 21`). Les
feuilles dont le nom ne correspond pas à ce motif (`Semaine paire`,
`Semaine impaire`, `Feuille15`…) sont ignorées.

La structure d'une feuille semaine est :

| Colonne | Contenu                                                          |
|---------|------------------------------------------------------------------|
| 1       | Jour de la semaine (`Lundi`, `Mardi`, …) — cellules fusionnées   |
| 2       | Numéro ISO de la semaine sur la 1ʳᵉ ligne, jour du mois ensuite  |
| 3       | Nom de l'opération (`Cultures`, `Pain`, `Marché de Pontorson`…)  |
| 4..N-2  | Une colonne par personne, contenant les heures travaillées        |
| N-1     | `TOTAL` (somme par ligne)                                         |
| N       | `TOTAL Jour` (somme par journée)                                  |

La 1ʳᵉ ligne contient le numéro ISO de la semaine (en en-tête de la 2ᵉ
colonne) et une ligne agrégée `TOTAL / dates` qui sert de référence et est
ignorée par le script.

Le script tolère :

- les cellules fusionnées sur la colonne « jour » (forward-fill automatique) ;
- les valeurs `'o'` (présence sans durée, traitées comme 0) ;
- les variations de casse et d'espaces sur les libellés d'opérations
  (`marché bio de Dol` ≡ `Marché bio de Dol` ≡ ` Marché bio de Dol`).

## Configuration

Toute la configuration vit dans `config.toml`, placé à côté du script.
Toutes les sections sont obligatoires.

```toml
[source]
ods_filepath = "../202605_planing.ods"   # chemin du tableur source
year = 2026                              # année ISO des semaines

[destination]
ics_root = "../202605_planing"           # racine des .ics ; un suffixe
                                          # "_<Personne>.ics" est ajouté

[calendar]
prodid = "-//atantet//ods2ical_christophe/"
timezone = "Europe/Paris"
version = "2.0"

[events]
categories = ["Professionnel"]           # appliquées à tous les événements

[operations]
# Heure de début par opération, au format "H:MM" ou "HH:MM".
# L'heure de début de l'événement d'une journée est la plus matinale
# parmi les opérations réalisées ce jour-là.
start_hours = { "Administratif" = "9:00", "Pain" = "5:30", ... }

[people]
# Couleur (nom CSS) par personne. Le nom doit correspondre à l'en-tête de
# colonne du tableur (la comparaison est insensible à la casse et aux
# espaces).
colors = { "Alexis" = "blue", "Christophe" = "green", ... }
```

Si une opération du tableur n'a pas d'entrée dans `[operations].start_hours`,
ou si une personne n'a pas de couleur dans `[people].colors`, le script
lève une `KeyError` explicite indiquant ce qu'il faut compléter.

## Utilisation

Depuis le dossier contenant `ods2ical_christophe.py` et `config.toml` :

```bash
python ods2ical_christophe.py
```

Le script écrit un fichier `<ics_root>_<Personne>.ics` par personne listée
dans `[people].colors`, agrégeant les événements de toutes les semaines
trouvées dans le tableur. Le contenu de chaque calendrier est aussi
affiché sur la sortie standard.

## Import dans un agenda

Les `.ics` produits sont des calendriers iCalendar 2.0 standards, avec
fuseau horaire `Europe/Paris` intégré. Pour les importer :

- **Thunderbird / Lightning** : *Fichier → Ouvrir → Fichier calendrier* ;
- **Google Calendar** : *Paramètres → Importer et exporter → Importer* ;
- **Apple Calendar** : double-clic sur le `.ics` ou *Fichier → Importer*.

## Scripts

- `ods2ical_christophe.py` — script principal, piloté par `config.toml`.
- `ods2ical_emploi_du_temps_prévisionnel.py` — version historique
  paramétrée en ligne de commande (constantes en dur, deux fichiers ODS
  paire/impaire passés en arguments). Conservé à titre de référence.

## Licence

GNU General Public License v3.0 — voir [`LICENSE`](LICENSE).
