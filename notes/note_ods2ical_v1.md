# Note de travail — ods2ical_ferme

Note pour mémoire des sessions futures sur le pipeline de conversion du
planning ODS de la ferme en calendriers ICS.

## Contexte

- **Dépôt** : https://github.com/atantet/ods2ical_ferme
- **Fichier source du projet** : `202605_planing.ods` (planning de mai/juin
  2026, semaines 21 à 26).
- **Objectif** : générer un `.ics` par personne, importable dans n'importe
  quel agenda, à partir du tableur.

Le dépôt contient deux scripts :

- `ods2ical_christophe.py` — script principal, piloté par `config.toml`.
  C'est celui à maintenir.
- `ods2ical_emploi_du_temps_prévisionnel.py` — version historique avec
  constantes en dur et CLI à deux fichiers paire/impaire. Conservée mais
  pas maintenue. Ne pas y toucher sauf demande explicite.

## État après la session du 24/05/2026

Première mise au point du script pour qu'il traite correctement
`202605_planing.ods`. Le user a poussé sur GitHub : le dépôt actuel
contient les corrections décrites ci-dessous. Le README et
l'`environment.yaml` ont aussi été produits durant cette session.

### Modifications structurelles apportées au script

1. **Lecture des feuilles par position et non par nom de colonne**
   — l'ancien `set_index(["sem", "TRAVAUX"])` faisait référence à une mise
   en forme antérieure du tableur. Désormais : `col[0]` = jour, `col[2]` =
   opération, `iloc[:, 1:-2]` pour ne garder que les colonnes des
   personnes (saute le numéro de jour, exclut `TOTAL` et `TOTAL Jour`).
2. **Numéro de semaine ISO lu dans le header `col[1]`** (entier comme
   `21`, `22`…), avec vérification de cohérence avec le suffixe du nom de
   feuille `Semaine N`. La clé `[source].week_row` du TOML est devenue
   inutile et a été retirée.
3. **`ffill()` sur la colonne jour** — dans le tableur, les cellules de la
   colonne « jour » sont fusionnées : seule la 1ʳᵉ ligne d'un jour porte
   `Lundi`, `Mardi`, etc., les suivantes sont `NaN`. Sans cette
   correction, ~80 % des heures tombent dans un groupe `NaN` et sont
   silencieusement perdues. **C'est le piège le plus pernicieux du
   projet** : les calendriers sont produits, ils ne lèvent aucune erreur,
   mais ils sont incomplets. Si jamais un futur fichier produit des
   calendriers avec des totaux faibles inexplicables, vérifier ce point
   en premier.
4. **Filtrage strict des feuilles** : on ne traite que les feuilles dont
   le nom matche `Semaine <entier>`. Cela exclut `Semaine paire`,
   `Semaine impaire` (qui ont un en-tête `19`/`20` et seraient sinon
   traitées comme semaines 19/20 — cf. piège §3 ci-dessous) et les
   `Feuille15..37` vides.
5. **Normalisation des libellés** (`strip + casefold`) pour la
   correspondance entre tableur et config, sur les opérations comme sur
   les personnes. Le tableur contient à la fois `Marché bio de Dol` et
   ` marché bio de Dol` ; une seule entrée dans le TOML les couvre.
6. **`JOUR_NUM` en minuscules** — était en majuscules dans la version
   antérieure, alors que le tableur écrit `Lundi`, `Mardi`. Couplé à la
   normalisation.
7. **`_parse_hhmm()`** — la config TOML stocke les heures de départ comme
   strings `"9:00"`, alors que la version originale les attendait en
   tuples Python. La fonction accepte les deux.
8. **`_format_hours()`** — remplace `int(heures)` qui affichait `(0)` pour
   une opération de 45 minutes (0.75 h). Désormais : `0.75` ou `7` selon
   le cas, fidèle au tableur.

### Validation effectuée

Pour chaque (personne, semaine), la somme des durées d'événements ICS
égale la somme des heures de la colonne correspondante dans le tableur.
72 contrôles (12 personnes × 6 semaines), tous OK. Bilan : **192
événements** au total répartis sur 12 calendriers (Marie est à 0 partout
dans le fichier source, son `.ics` est vide). C'est l'état de référence
auquel se comparer si une régression est suspectée.

## État après la session du 22/06/2026

Nouveau fichier source : `2026_emploi_du_temps.ods` (renommé depuis
`emploi_du_temps_prévisionnel modèle Alexis(1).ods`), **semaines 21 à
30** (10 semaines au lieu de 6). `config.toml` mis à jour :
`ods_filepath` et `ics_root` pointent désormais sur
`../2026_emploi_du_temps[.ods]`.

**Aucune modification du script n'a été nécessaire** : la structure
interne des feuilles de semaine est identique à l'ancien fichier (col 0 =
jour fusionné, col 1 = n° semaine / jour du mois, col 2 = opération, 12
colonnes personnes, `TOTAL`, `TOTAL Jour`). Le nouveau fichier ajoute des
onglets de synthèse (`CONSTANTES GLOBALES`, `Définition des débouchés`,
`Par jour`, `Par personne`, `Par atelier`, `Coûts de commercialisation`)
qui sont tous ignorés par le filtre `Semaine <entier>`. À noter : une
feuille s'appelle `semaine 21` en minuscule — gérée car le filtre
abaisse la casse.

### Validation effectuée

Même méthodo qu'au 24/05 : pour chaque (personne, semaine), somme des
durées ICS == somme de la colonne tableur. **120 contrôles (12 personnes
× 10 semaines), tous OK. 317 événements** au total. Marie reste à 0
partout (ICS vide, 86 octets). C'est le nouvel état de référence.

Le dossier de notes a été renommé `IA/` → `notes/` pour coller à la
convention des autres projets de `~/projets/`.

## Format du fichier ODS — choses à savoir

La structure attendue (et tolérée) est documentée dans le README. Points
qui ne sont **pas** évidents à la lecture du script :

- La 1ʳᵉ ligne d'une feuille de semaine est une ligne `TOTAL / dates`
  agrégée, ignorée explicitement (`df.index.get_level_values(0) != "TOTAL"`).
- Une cellule contenant `'o'` est traitée comme `0` (présence sans
  durée). Cas observé dans certains plannings, conservé pour rétro-
  compatibilité.
- Les libellés d'opérations sont normalisés (`strip + casefold`) côté
  lookup. Donc inutile de dupliquer dans le TOML les variations de
  casse/espace ; **une seule entrée par opération** suffit.
- Les couleurs des personnes utilisent les noms CSS (`blue`, `green`,
  `fuchsia`, `purple`, `maroon`…). Le champ `COLOR` d'un VEVENT n'est pas
  affiché par tous les clients d'agenda, mais figure bien dans les ICS.

## Pièges connus pour les futures sessions

1. **Nouvelle opération dans le tableur**. Si une opération absente de
   `[operations].start_hours` apparaît dans une feuille, le script lève
   `KeyError` avec un message indiquant quoi compléter. **Action** :
   ajouter une entrée dans le TOML avec l'heure de début habituelle de
   cette opération. Ne pas être tenté de mettre une heure par défaut côté
   script : l'erreur explicite est plus utile.

2. **Nouvelle personne dans le tableur**. Même comportement, sur
   `[people].colors`. Choisir une couleur CSS qui ne soit pas déjà prise
   pour éviter les confusions d'affichage.

3. **Feuilles `Semaine paire` / `Semaine impaire`**. Ce sont des modèles
   à copier. Elles portent en en-tête un numéro de semaine plausible
   (`20` et `19` dans le fichier actuel). Le filtre `Semaine <entier>`
   les exclut, donc tant qu'elles ne sont pas renommées en `Semaine 19`
   ou `Semaine 20` elles n'auront aucun effet. Si le user crée
   effectivement des feuilles `Semaine 19/20` en plus, vérifier qu'elles
   sont bien remplies — sinon on pourrait avoir des doubles événements
   ou des ICS vides pour ces semaines.

4. **Année ISO incohérente**. `config.toml` fixe `year = 2026`. Si le
   tableur contient des semaines en chevauchement d'année (semaine 1 ou
   52/53), la cohérence ISO doit être vérifiée. Pour l'instant non
   pertinent (le fichier ne contient que les semaines 21-26).

5. **Encodage du nom de personne dans le nom de fichier ICS**. Le script
   écrit `202605_planing_Aide 1.ics` (avec espace) et
   `202605_planing_Jérôme.ics` (avec accent). Marche sur Linux/macOS et
   sur la plupart des outils d'import. Si jamais un user signale des
   problèmes (Windows en encodage non-UTF-8, anciens clients), envisager
   un `slugify` du nom dans le chemin du fichier — pas dans le
   `SUMMARY`/`COLOR` du VEVENT, ces champs restant les libellés exacts.

6. **Le `Calendar.color` est défini deux fois** dans la chaîne actuelle :
   sur le `cal_name` à l'intérieur de chaque semaine, puis perdu lors du
   merge dans `merged_cal` (où seuls `prodid` et `version` sont propagés).
   Conséquence : la propriété `color` du calendrier merged n'est pas
   définie, seul le `COLOR` de chaque VEVENT l'est. Si un user demande à
   ce que le calendrier entier porte une couleur (pour Thunderbird par
   exemple), il faudra propager `color` dans `merge_calendars`. Non
   bloquant aujourd'hui.

## Reproductibilité

L'environnement est dans `environment.yaml` : Python ≥ 3.11 (pour
`tomllib` de la stdlib), `pandas`, `numpy`, `odfpy`, `icalendar`. Tout
sur conda-forge. Testé à blanc dans un venv vide : 4 paquets pip
suffisent à reproduire la chaîne complète.

Commande standard :

```bash
conda env create -f environment.yaml
conda activate ods2ical_ferme
python ods2ical_christophe.py
```

Pour itérer rapidement : `rm -f *.ics && python ods2ical_christophe.py`
depuis le dossier qui contient le tableur et la config.

## À faire si la conversation reprend

- Vérifier l'état actuel du dépôt GitHub (le user peut avoir poussé
  d'autres modifs depuis).
- Vérifier si un nouveau fichier ODS est arrivé dans le projet — s'il a
  une mise en forme différente, repartir de la liste de pièges ci-dessus
  pour diagnostiquer.
- Si le user demande d'ajouter une feature (par ex. export en un seul
  ICS multi-personnes, ou filtrage par semaine), c'est le script
  `ods2ical_christophe.py` à modifier, pas l'autre.
