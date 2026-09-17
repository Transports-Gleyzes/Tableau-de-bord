#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
maj_mensuelle_flotte.py
========================
Met à jour automatiquement TABLEAU_DE_BORD_FLOTTE.xlsx à partir de
l'arborescence SOURCES du serveur (OneDrive) :

    TABLEAU DE BORD/SOURCES/
        CA CAMION/CA_CAMION.xlsx                (1 onglet par mois : "CA <MOIS> <AAAA>")
        CARBURANT/GLEYZES/...                    (pas encore de source -> 0€, cf. DONNEES_MANQUANTES)
        CARBURANT/LPB/*.xlsx                     (colonnes DATE / IMMATRICULATION (texte libre) / MONTANTS €)
        CHARGES FIXES MUTUALISEES/CHARGES_FIXES_MUTUALISEES.xlsx  (statique, jamais réimporté)
        KM CAMION/Liste_KM_AAAA-MM.xlsx
        PAIES/AAAA-MM/*.pdf                      (bulletins "clarifiés", texte OCR intégré ou à défaut OCR Tesseract)
        PEAGES/GLEYZES/*.pdf et PEAGES/LPB/*.pdf  (relevés Axxès, 2 formats possibles)

Principes (méthode HPPH) :
  - REF_Camions n'est JAMAIS modifié par ce script.
  - Ne jamais inventer un 0 : une donnée introuvable reste vide et part dans
    DONNEES_MANQUANTES plutôt que d'être comblée silencieusement.
  - Déduplication : une ligne déjà importée (même Mois + Camion_ID + Société
    + montants) n'est pas réinjectée si le script est relancé sur le même mois.
  - Ajout de lignes = vraie extension de la Table Excel (ListObject), pas
    une simple écriture de cellules : les formules des colonnes calculées
    sont recopiées automatiquement sur les nouvelles lignes, exactement
    comme si l'utilisateur appuyait sur Tab dans Excel.
  - En cas de fichier verrouillé par OneDrive/Excel (PermissionError), le
    résultat est sauvegardé sous un nom de secours *_SAUVEGARDE_SECOURS.xlsx.

Usage :
    python3 maj_mensuelle_flotte.py --classeur TABLEAU_DE_BORD_FLOTTE.xlsx \
        --sources "TABLEAU DE BORD/SOURCES" [--mois 2026-08]

Si --mois est omis, le script traite le(s) mois le(s) plus récent(s) trouvé(s)
dans KM CAMION (source la plus fiable pour détecter un nouveau mois).
"""
import argparse
import glob
import os
import re
import sys
import zipfile
from collections import defaultdict
from datetime import date, datetime

import openpyxl
from openpyxl.worksheet.table import Table
from openpyxl.utils import get_column_letter

try:
    import pdfplumber
except Exception:
    # pdfplumber est optionnel (seuls les PDF natifs, non déjà OCRisés, en ont besoin) ; on tolère
    # ici toute erreur d'import, pas seulement ImportError, car une dépendance transitive cassée
    # (ex: cryptography/cffi mal installé sous Windows) lève une autre exception et faisait
    # planter tout le script dès le chargement du module, avant même d'afficher quoi que ce soit
    # (bug corrigé le 17/09/2026).
    pdfplumber = None

MOIS_FR = {"JANVIER": 1, "FEVRIER": 2, "FÉVRIER": 2, "MARS": 3, "AVRIL": 4, "MAI": 5, "JUIN": 6,
           "JUILLET": 7, "AOUT": 8, "AOÛT": 8, "SEPTEMBRE": 9, "OCTOBRE": 10,
           "NOVEMBRE": 11, "DECEMBRE": 12, "DÉCEMBRE": 12}

CORRECTIONS_PLAQUES = {  # erreurs de saisie/OCR connues, à enrichir au fil de l'eau
    "6D042ZC": "GD042ZC",
    "GD042RC": "GD042ZC",  # coquille RC/ZC repérée dans TABLEAU_TICKET_CARB (31/08/2026)
    "GC978GW": "GQ978GW",
    "BX490XG": "GD042ZC",   # alias badge Axxès confirmé par Adeline (26/08/2026)
}
NOM_CHAUFFEUR_VERS_PLAQUE = {"GARCIA": "GD042ZC"}  # complété depuis REF_Camions au chargement


# ---------------------------------------------------------------------------
# Utilitaires génériques
# ---------------------------------------------------------------------------
def normalise_prenom(txt):
    """Compare les prénoms sans tenir compte des accents, tirets ou espaces
    (bug corrigé le 31/08/2026 : 'Loic' ne matchait pas 'Loïc', 'Jean Baptiste'
    ne matchait pas 'Jean-Baptiste')."""
    import unicodedata
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return re.sub(r"[\s\-]+", "", txt).upper()


def societe_canonique(s):
    """Ramène une société à sa forme canonique 'Gleyzes' ou 'LPB'.
    Bug corrigé le 31/08/2026 : REF_Camions stocke 'LPB (except.)' pour GD042ZC
    (lisibilité humaine), mais ce libellé ne doit JAMAIS se retrouver comme
    valeur de la colonne Société dans les tables de données (sinon les clés de
    dédoublonnage 'LPB' vs 'LPB (except.)' ne correspondent plus -> doublons).
    Bug corrigé le 03/09/2026 : normalise aussi la casse ('GLEYZES' -> 'Gleyzes'),
    certains fichiers sources (CHARGES_FIXES_MUTUALISEES) utilisant des majuscules."""
    if not s:
        return s
    s = str(s).split("(")[0].strip()
    if s.upper() == "GLEYZES":
        return "Gleyzes"
    if s.upper() == "LPB":
        return "LPB"
    return s


def normalise_plaque(txt):
    """Extrait et normalise une immatriculation depuis un texte libre."""
    txt_up = str(txt).upper().replace("-", "").replace(" ", "")
    for wrong, right in CORRECTIONS_PLAQUES.items():
        if wrong in txt_up:
            return right
    m = re.search(r"([A-Z]{2}\d{2,3}[A-Z]{2})", txt_up)
    if m:
        return CORRECTIONS_PLAQUES.get(m.group(1), m.group(1))
    for nom, plaque in NOM_CHAUFFEUR_VERS_PLAQUE.items():
        if nom in txt_up:
            return plaque
    return None


def lire_texte_pdf(path):
    """Lit le texte d'un PDF, qu'il s'agisse d'un vrai PDF ou d'une archive
    OCR (jpeg + txt zippés, format des bulletins/relevés déjà OCRisés)."""
    try:
        with zipfile.ZipFile(path) as z:
            noms = sorted([n for n in z.namelist() if n.endswith(".txt")],
                          key=lambda x: int(x.split(".")[0]) if x.split(".")[0].isdigit() else 0)
            return "\n".join(z.read(n).decode("utf-8", errors="ignore") for n in noms)
    except zipfile.BadZipFile:
        if pdfplumber is None:
            raise RuntimeError("pdfplumber requis pour lire un PDF natif (pip install pdfplumber)")
        with pdfplumber.open(path) as pdf:
            return "\n".join((p.extract_text() or "") for p in pdf.pages)


def sauvegarder_avec_secours(wb, chemin):
    """Sauvegarde le classeur ; si le fichier est verrouillé (OneDrive/Excel
    ouvert), sauvegarde sous un nom de secours plutôt que d'échouer."""
    try:
        wb.save(chemin)
        return chemin
    except PermissionError:
        base, ext = os.path.splitext(chemin)
        secours = f"{base}_SAUVEGARDE_SECOURS{ext}"
        wb.save(secours)
        print(f"  [!] {chemin} est verrouillé (OneDrive/Excel ouvert) -> sauvegardé sous {secours}")
        return secours


def _cle_normalisee(valeur):
    """Normalise une valeur pour la comparaison de clé de dédoublonnage.
    Bug corrigé le 31/08/2026 : openpyxl relit une date stockée comme
    datetime.datetime alors que le script fournit un datetime.date pour le
    même jour -> comparaison de tuple qui échoue silencieusement -> doublons.
    On convertit systématiquement tout objet date/datetime en chaîne AAAA-MM-JJ,
    et on met les textes en majuscules sans espaces superflus."""
    if isinstance(valeur, (datetime, date)):
        return valeur.strftime("%Y-%m-%d")
    if isinstance(valeur, str):
        return valeur.strip().upper()
    return valeur


def ajouter_lignes_table(ws, table_name, nouvelles_lignes, cle_dedup, colonnes_formule=None):
    """
    Étend une vraie Table Excel (ListObject) avec de nouvelles lignes, ET
    complète les lignes déjà existantes dont une case est restée vide.
    - nouvelles_lignes : liste de dict {nom_colonne: valeur} (colonnes calculées omises)
    - cle_dedup : tuple de noms de colonnes formant la clé d'unicité (ex: ("Mois","Camion_ID","Société"))
    - colonnes_formule : {nom_colonne: fonction(r) -> texte de formule}
      Formules CLASSIQUES (références de cellules/colonnes), pas de références
      structurées [@Col] ni Table[Col] — abandonnées le 31/08/2026 après échec
      répété dans Excel réel (tables supprimées à l'ouverture, #REF! sur des
      formules triviales). r = numéro de ligne Excel (1-based) de la nouvelle ligne.

    Bug corrigé le 03/09/2026 : une ligne déjà présente (ex: le mois de juillet
    créé à l'avance avec KM vide) était reconnue comme "déjà importée" par la
    clé de dédoublonnage et donc jamais complétée, même quand la vraie donnée
    devenait disponible ensuite. Désormais, pour une clé déjà existante, toute
    cellule actuellement vide est complétée avec la nouvelle valeur si elle
    existe ; une cellule déjà remplie n'est en revanche jamais écrasée (pour ne
    pas effacer une correction manuelle ou une saisie antérieure sans confirmation).

    Bug corrigé le 17/09/2026 : les colonnes de colonnes_formule n'étaient
    écrites qu'à la création d'une ligne, jamais revérifiées ensuite. Une
    formule corrompue une seule fois (ex: plage décalée lors d'un ajout
    manuel, ou par une version antérieure du script) restait donc fausse
    indéfiniment, avec des montants qui ne "reportaient" plus le même total
    d'un mois à l'autre. Contrairement aux colonnes de données, une formule
    n'est jamais une saisie manuelle à préserver : elle est désormais
    toujours resynchronisée avec colonnes_formule sur les lignes existantes,
    ce qui corrige aussi automatiquement toute corruption passée.

    Retourne (lignes_ajoutees, lignes_completees, formules_reparees).
    """
    tbl = ws.tables[table_name]
    ref = tbl.ref
    from openpyxl.utils.cell import range_boundaries
    min_col, min_row, max_col, max_row = range_boundaries(ref)
    headers = [ws.cell(row=min_row, column=c).value for c in range(min_col, max_col + 1)]
    col_index = {h: i for i, h in enumerate(headers)}
    colonnes_formule = colonnes_formule or {}

    # Construire la correspondance clé -> ligne pour les clés déjà présentes
    existantes = {}
    for r in range(min_row + 1, max_row + 1):
        vals = tuple(_cle_normalisee(ws.cell(row=r, column=min_col + col_index[k]).value)
                      for k in cle_dedup if k in col_index)
        if any(v is not None for v in vals):
            existantes[vals] = r

    # Trouver la première ligne "vide" du buffer pré-provisionné, sinon on ajoutera à la fin
    ligne_libre = None
    for r in range(min_row + 1, max_row + 1):
        cle_vals = tuple(ws.cell(row=r, column=min_col + col_index[k]).value for k in cle_dedup if k in col_index)
        if all(v is None for v in cle_vals):
            ligne_libre = r
            break

    ajoutees, completees, formules_reparees = 0, 0, 0
    for ligne in nouvelles_lignes:
        cle = tuple(_cle_normalisee(ligne.get(k)) for k in cle_dedup)
        if cle in existantes:
            r = existantes[cle]
            a_complete = False
            for h, i in col_index.items():
                if h in colonnes_formule or h in cle_dedup:
                    continue
                c = min_col + i
                if ws.cell(row=r, column=c).value is None and ligne.get(h) is not None:
                    ws.cell(row=r, column=c, value=ligne.get(h))
                    a_complete = True
            if a_complete:
                completees += 1
            for h, fn in colonnes_formule.items():
                if h not in col_index:
                    continue
                c = min_col + col_index[h]
                nouvelle_formule = fn(r)
                if ws.cell(row=r, column=c).value != nouvelle_formule:
                    ws.cell(row=r, column=c, value=nouvelle_formule)
                    formules_reparees += 1
            continue
        if ligne_libre is not None and ligne_libre <= max_row:
            r = ligne_libre
            ligne_libre += 1
        else:
            r = max_row + 1
            max_row = r
        for h, i in col_index.items():
            c = min_col + i
            if h in colonnes_formule:
                ws.cell(row=r, column=c, value=colonnes_formule[h](r))
            else:
                ws.cell(row=r, column=c, value=ligne.get(h))
        existantes[cle] = r
        ajoutees += 1

    if max_row != range_boundaries(ref)[3]:
        new_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
        tbl.ref = new_ref
    return ajoutees, completees, formules_reparees


def get_table_ws(wb, table_name):
    for ws in wb.worksheets:
        if table_name in ws.tables:
            return ws
    raise KeyError(f"Table {table_name} introuvable dans le classeur")


def lire_referentiel(wb):
    ws = get_table_ws(wb, "T_Referentiel")
    from openpyxl.utils.cell import range_boundaries
    min_col, min_row, max_col, max_row = range_boundaries(ws.tables["T_Referentiel"].ref)
    headers = [ws.cell(row=min_row, column=c).value for c in range(min_col, max_col + 1)]
    idx = {h: i for i, h in enumerate(headers)}
    ref = {}
    for r in range(min_row + 1, max_row + 1):
        row = [ws.cell(row=r, column=min_col + i).value for i in range(len(headers))]
        cid = row[idx["Camion_ID"]]
        if cid:
            ref[cid] = dict(zip(headers, row))
    return ref


# ---------------------------------------------------------------------------
# Détection du mois traité
# ---------------------------------------------------------------------------
def _mois_depuis_nom(texte):
    """Cherche un motif AAAA-MM (séparateur libre) dans un nom de fichier/dossier."""
    m = re.search(r"(\d{4})[-_ ]?(\d{2})\b", str(texte))
    return f"{m.group(1)}-{m.group(2)}" if m else None


def _mois_depuis_texte_fr(texte):
    """Cherche un nom de mois français + une année à 4 chiffres dans un texte libre
    (ex: nom d'onglet 'CA AOUT 2026', nom de fichier de bulletin 'PAIE_SEPTEMBRE_2026.pdf')."""
    texte_up = str(texte).upper()
    m_annee = re.search(r"(20\d{2})", texte_up)
    if not m_annee:
        return None
    annee = m_annee.group(1)
    for nom, num in MOIS_FR.items():
        if nom in texte_up:
            return f"{annee}-{num:02d}"
    return None


def _mois_depuis_dates_xlsx(fichiers):
    """Parcourt la colonne DATE (format dd.mm.yy, cf. CARBURANT/PEAGES) de chaque
    fichier et retourne l'ensemble des mois (AAAA-MM) rencontrés."""
    mois = set()
    for f in fichiers:
        try:
            wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
        except Exception:
            continue
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            try:
                d = datetime.strptime(str(row[0]), "%d.%m.%y")
            except ValueError:
                continue
            mois.add(f"{d.year}-{d.month:02d}")
        wb.close()
    return mois


def detecter_mois(sources_dir, mois_force=None):
    """Détecte le mois à traiter en cherchant dans TOUTES les sources (pas
    seulement KM CAMION comme avant) le mois le plus récent qui y apparaît.

    Bug corrigé le 17/09/2026 : la détection ne regardait que les noms de
    fichiers de KM CAMION. Résultat, quand l'utilisateur déposait de nouveaux
    documents (CA, CARBURANT, PAIES, PEAGES) pour un mois dont le fichier KM
    n'était pas encore arrivé, le script retombait sur le dernier mois déjà
    traité (déjà entièrement rempli) et n'ajoutait/complétait donc rien, sans
    la moindre erreur — donnant l'impression que « la mise à jour ne se fait
    pas ». On agrège désormais les mois vus dans KM CAMION, CA CAMION (noms
    d'onglets), PAIES (sous-dossiers et noms de fichiers) et CARBURANT/PEAGES
    (colonne DATE des tickets), et on retient le plus récent trouvé n'importe où.
    """
    if mois_force:
        return mois_force

    candidats = defaultdict(set)  # mois -> {sources qui l'ont révélé}

    for f in glob.glob(os.path.join(sources_dir, "KM CAMION", "**", "*.xlsx"), recursive=True):
        m = _mois_depuis_nom(os.path.basename(f))
        if m:
            candidats[m].add("KM CAMION")

    for f in glob.glob(os.path.join(sources_dir, "CA CAMION", "**", "*.xlsx"), recursive=True):
        try:
            wb = openpyxl.load_workbook(f, read_only=True)
        except Exception:
            continue
        for nom in wb.sheetnames:
            m = _mois_depuis_texte_fr(nom)
            if m:
                candidats[m].add("CA CAMION")
        wb.close()

    for d in glob.glob(os.path.join(sources_dir, "PAIES", "*")):
        if os.path.isdir(d):
            m = _mois_depuis_nom(os.path.basename(d))
            if m:
                candidats[m].add("PAIES")
    for f in glob.glob(os.path.join(sources_dir, "PAIES", "**", "*.pdf"), recursive=True):
        m = _mois_depuis_texte_fr(os.path.basename(f))
        if m:
            candidats[m].add("PAIES")

    carb_files = glob.glob(os.path.join(sources_dir, "CARBURANT", "**", "*.xlsx"), recursive=True)
    for m in _mois_depuis_dates_xlsx(carb_files):
        candidats[m].add("CARBURANT")

    peage_tickets = glob.glob(os.path.join(sources_dir, "PEAGES", "**", "TICKETS_*.xlsx"), recursive=True)
    for m in _mois_depuis_dates_xlsx(peage_tickets):
        candidats[m].add("PEAGES (tickets)")

    if not candidats:
        raise RuntimeError("Impossible de détecter le mois automatiquement : aucun fichier daté trouvé dans "
                            "KM CAMION, CA CAMION, PAIES ou CARBURANT. Fournir --mois AAAA-MM.")

    mois_retenu = sorted(candidats)[-1]
    print(f"Mois auto-détecté : {mois_retenu} (vu dans : {', '.join(sorted(candidats[mois_retenu]))})")
    if "KM CAMION" not in candidats[mois_retenu]:
        print(f"  [!] Aucun fichier KM CAMION pour {mois_retenu} pour l'instant : les KM resteront vides "
              f"jusqu'à son dépôt (le reste des données sera bien mis à jour).")
    return mois_retenu


# ---------------------------------------------------------------------------
# Parsers par source (réutilisent la logique validée manuellement en amont)
# ---------------------------------------------------------------------------
def parse_km(sources_dir, mois):
    """Recherche RÉCURSIVE (sous-dossiers inclus) et tolérante sur le nom de
    fichier : accepte Liste_KM_2026-07.xlsx, Liste_KM_2026_07.xlsx,
    Liste_KM_07-2026.xlsx, avec ou sans suffixe/copie ' (1)', etc. Bug corrigé
    le 03/09/2026 : la recherche précédente n'acceptait qu'un nom de fichier
    exact au premier niveau du dossier KM CAMION."""
    annee, mm = mois.split("-")
    tous = glob.glob(os.path.join(sources_dir, "KM CAMION", "**", "*.xlsx"), recursive=True)
    candidats = [f for f in tous
                 if re.search(rf"{annee}[-_ ]?{mm}\b", os.path.basename(f))
                 or re.search(rf"{mm}[-_ ]?{annee}\b", os.path.basename(f))]
    if not candidats:
        return {}, [f"KM CAMION : fichier du mois {mois} introuvable (recherche récursive, "
                     f"{len(tous)} fichier(s) .xlsx examiné(s) dans KM CAMION)"]
    wb = openpyxl.load_workbook(candidats[0], data_only=True)
    ws = wb.active
    res, manquants = {}, []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0] or str(row[0]).strip().upper() == "TOTAL":
            continue
        plaque = normalise_plaque(row[0])
        km = row[-1]
        if plaque and isinstance(km, (int, float)):
            res[plaque] = km
    return res, manquants


def parse_ca(sources_dir, mois, referentiel):
    fichiers = glob.glob(os.path.join(sources_dir, "CA CAMION", "**", "*.xlsx"), recursive=True)
    if not fichiers:
        return {}, ["CA CAMION : aucun fichier .xlsx trouvé dans le dossier CA CAMION"]
    mois_nom = {v: k for k, v in MOIS_FR.items()}[int(mois.split("-")[1])]
    annee = mois.split("-")[0]
    prenom_vers_camion = {}
    for cid, infos in referentiel.items():
        chauffeur = (infos.get("Chauffeur") or "").strip()
        if chauffeur and " " in chauffeur:
            prenom = chauffeur.rsplit(" ", 1)[0]  # tout sauf le dernier mot (le nom de famille)
            prenom_vers_camion[normalise_prenom(prenom)] = (cid, societe_canonique(infos.get("Société")))

    res, manquants = {}, []
    for fpath in fichiers:
        wb = openpyxl.load_workbook(fpath, data_only=True)
        cible = None
        for name in wb.sheetnames:
            if mois_nom in name.upper() and annee in name:
                cible = name
                break
        if not cible:
            manquants.append(f"CA CAMION : onglet du mois {mois} introuvable dans {os.path.basename(fpath)}")
            continue
        ws = wb[cible]
        for row in ws.iter_rows(values_only=True):
            if not row:
                continue
            nom = str(row[2]).strip() if row[2] else None
            if not nom or nom.upper() in ("CHAUFFEUR", "") or "TOTAL" in nom.upper() or "DIFFERENCE" in normalise_prenom(nom):
                continue
            if normalise_prenom(nom) in ("DENIS", "DANIEL"):
                continue  # Denis exclu (demande Adeline) ; Daniel = agent d'exploitation Gleyzes, pas chauffeur
            ca_gleyzes = row[4] or 0
            ca_lpb = row[7] or 0
            match = prenom_vers_camion.get(normalise_prenom(nom))
            if not match:
                manquants.append(f"CA CAMION : chauffeur '{row[2]}' non trouvé dans le référentiel "
                                  f"(fichier {os.path.basename(fpath)})")
                continue
            cid, societe_ref = match
            if ca_gleyzes and ca_lpb:
                res.setdefault(cid, []).extend([("Gleyzes", ca_gleyzes), ("LPB", ca_lpb)])
            else:
                montant = ca_gleyzes if ca_gleyzes else ca_lpb
                res.setdefault(cid, []).append((societe_ref, montant))
    return res, manquants


def parse_montant(val):
    """Convertit une valeur de montant potentiellement mal formée en float,
    sans jamais planter tout le script sur une seule cellule. Bug corrigé le
    03/09/2026 : des montants du type '1 025.04' (espace = séparateur de
    milliers) faisaient planter float() et interrompaient tout le traitement
    du mois sans message clair. Retourne None si vraiment inexploitable."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(" ", "").replace("\u202f", "").replace("\xa0", "")
    if "," in s and "." in s:
        s = s.replace(",", "")  # virgule = séparateur de milliers, point = décimale
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def parse_carburant(sources_dir, mois):
    """Lit tous les fichiers de tickets carburant, quel que soit le sous-dossier
    (GLEYZES ou LPB) : en pratique, certains camions Gleyzes se retrouvent dans
    le fichier déposé côté LPB (cf. GB050RT/GB045RT/GE606BF dans
    TABLEAU_TICKET_CARB_LPB) — on ne présuppose donc plus la société à partir
    de l'emplacement du fichier, seule l'immatriculation fait foi. Bug corrigé
    le 31/08/2026 : le carburant Gleyzes était auparavant forcé à 0 dans main(),
    ignorant toute donnée réelle déposée entre-temps."""
    fichiers = glob.glob(os.path.join(sources_dir, "CARBURANT", "**", "*.xlsx"), recursive=True)
    totals = defaultdict(float)
    manquants = []
    y, m = mois.split("-")
    for f in fichiers:
        wb = openpyxl.load_workbook(f, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None or row[2] is None:
                continue
            try:
                d = datetime.strptime(str(row[0]), "%d.%m.%y")
            except ValueError:
                continue
            if f"{d.year}" != y or f"{d.month:02d}" != m:
                continue
            plaque = normalise_plaque(row[1])
            if not plaque:
                manquants.append(f"CARBURANT : immatriculation illisible '{row[1]}' ({os.path.basename(f)})")
                continue
            montant = parse_montant(row[2])
            if montant is None:
                manquants.append(f"CARBURANT : montant illisible '{row[2]}' pour {plaque} ({os.path.basename(f)})")
                continue
            totals[plaque] += montant
    return dict(totals), manquants


def parse_peages_axxes(path):
    """Retourne {plaque: montant_ttc} pour un relevé Axxès (2 formats possibles)."""
    return parse_peages_axxes_depuis_texte(lire_texte_pdf(path))


def parse_peages_axxes_depuis_texte(full):
    """Retourne {plaque: montant_HT} pour un relevé Axxès (2 formats possibles).
    Bug corrigé le 03/09/2026 : le script prenait le montant TTC (3e colonne,
    'Flotte HT TTC') au lieu du montant HT (2e colonne) demandé par Adeline."""
    res = {}
    for m in re.finditer(
        r"V[ée]hicule\s*/\s*Vehicle\s*[-:]?\s*([A-Z0-9]{5,8})\s*-\s*Classe.*?Flotte\s+(-?[\d\s]+,\d{2})\s+(-?[\d\s]+,\d{2})",
        full):
        plaque = normalise_plaque(m.group(1)) or m.group(1)
        ht = float(m.group(2).replace(" ", "").replace(",", "."))
        res[plaque] = res.get(plaque, 0.0) + ht
    if res:
        return res
    parts = re.split(r"(V[ée]hicule\s*/\s*Vehicle\s*:\s*-\s*([A-Z0-9]{5,8}))", full)
    i = 1
    while i < len(parts) - 2:
        plaque_brute, chunk = parts[i + 1], parts[i + 2]
        plaque = normalise_plaque(plaque_brute) or plaque_brute
        mm = re.search(r"Total\s*/\s*Total\s+([\d\s]+,\d)\s+([\d\s]+,\d{2})\s+([\d\s]+,\d{2})", chunk)
        if mm:
            res[plaque] = res.get(plaque, 0.0) + float(mm.group(2).replace(" ", "").replace(",", "."))
        i += 3
    return res


def extraire_mois_relevé(full_text):
    """Retourne le mois (AAAA-MM) couvert par un relevé Axxès, à partir du
    texte du PDF, en cherchant en priorité une période explicite, sinon la
    date du détail. Renvoie None si aucune date exploitable n'est trouvée.
    Bug corrigé le 31/08/2026 : sans ce contrôle, un PDF resté dans le dossier
    d'un mois précédent était réimporté dans le mois en cours (double compte)."""
    m = re.search(r"P[ée]riode\s+du.*?(\d{2})/(\d{2})/(\d{4}).*?au.*?(\d{2})/(\d{2})/(\d{4})", full_text)
    if m:
        # on retient le mois de la date de FIN de période (convention de facturation Axxès)
        return f"{m.group(6)}-{m.group(5)}"
    m = re.search(r"[ée]tail\s+no\.?.*?du.*?(\d{2})/(\d{2})/(\d{4})", full_text)
    if m:
        return f"{m.group(3)}-{m.group(2)}"
    return None


def parse_peages(sources_dir, mois):
    """Combine relevés Axxès (Gleyzes + LPB) et tickets manuels LPB pour le mois.
    Recherche RÉCURSIVE dans les sous-dossiers (ex: PEAGES/GLEYZES/GLEYZES-2026-07/) :
    bug corrigé le 03/09/2026, la recherche ne regardait auparavant que le premier
    niveau du dossier et ratait tout fichier rangé dans un sous-dossier mensuel."""
    totals = defaultdict(lambda: defaultdict(float))  # plaque -> société -> montant
    manquants = []
    for societe, dossier in (("Gleyzes", "GLEYZES"), ("LPB", "LPB")):
        pdfs = glob.glob(os.path.join(sources_dir, "PEAGES", dossier, "**", "*.pdf"), recursive=True)
        if not pdfs:
            manquants.append(f"PEAGES {dossier} : aucun relevé Axxès trouvé pour {mois}")
            continue
        au_moins_un_du_mois = False
        for p in pdfs:
            try:
                full = lire_texte_pdf(p)
            except Exception as e:
                manquants.append(f"PEAGES {dossier} : échec lecture {p} ({e})")
                continue
            mois_du_pdf = extraire_mois_relevé(full)
            if mois_du_pdf is None:
                manquants.append(f"PEAGES {dossier} : période introuvable dans {os.path.basename(p)}, "
                                  f"fichier ignoré par précaution (ne pas fausser un autre mois)")
                continue
            if mois_du_pdf != mois:
                continue  # relevé d'un autre mois resté dans le dossier : on l'ignore silencieusement
            au_moins_un_du_mois = True
            try:
                r = parse_peages_axxes_depuis_texte(full)
            except Exception as e:
                manquants.append(f"PEAGES {dossier} : échec analyse {p} ({e})")
                continue
            if not r:
                manquants.append(f"PEAGES {dossier} : {os.path.basename(p)} lu et période reconnue ({mois}), "
                                  f"mais AUCUNE immatriculation détectée dedans — le format de ce PDF diffère "
                                  f"probablement des deux formats Axxès connus du script. Coller le texte du "
                                  f"fichier au consultant pour ajuster la reconnaissance.")
                continue
            for plaque, montant in r.items():
                totals[plaque][societe] += montant
        if not au_moins_un_du_mois:
            manquants.append(f"PEAGES {dossier} : aucun relevé correspondant au mois {mois} trouvé "
                              f"parmi les {len(pdfs)} fichier(s) du dossier")
    # tickets manuels LPB (complément non Axxès)
    tickets = list(set(
        glob.glob(os.path.join(sources_dir, "PEAGES", "LPB", "**", "TICKETS_*.xlsx"), recursive=True) +
        glob.glob(os.path.join(sources_dir, "PEAGES", "**", "TICKETS_*.xlsx"), recursive=True)
    ))
    y, m = mois.split("-")
    for f in tickets:
        wb = openpyxl.load_workbook(f, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None or row[2] is None:
                continue
            try:
                d = datetime.strptime(str(row[0]), "%d.%m.%y")
            except ValueError:
                continue
            if f"{d.year}" != y or f"{d.month:02d}" != m:
                continue
            plaque = normalise_plaque(row[1])
            if plaque:
                montant = parse_montant(row[2])
                if montant is None:
                    manquants.append(f"PEAGES : montant illisible '{row[2]}' pour {plaque} ({os.path.basename(f)})")
                    continue
                totals[plaque]["LPB"] += montant
    return {p: dict(v) for p, v in totals.items()}, manquants


PERSONNEL_NON_CHAUFFEUR = {"BRUNA", "DANIEL"}  # agents d'exploitation mutualisés, pas de camion associé


def lire_charges_mutualisees(sources_dir, referentiel):
    """Recalcule la charge mutualisée par chauffeur, par société, à partir du
    fichier source (recherché récursivement, nom de fichier flexible) plutôt
    que d'une valeur codée en dur. Bug corrigé le 03/09/2026 : les montants
    5807,58 €/2576,54 € étaient figés dans le script, ignorant toute mise à
    jour ultérieure du fichier ou évolution du nombre de chauffeurs actifs."""
    # Motif avec '*' final : le dossier réel s'appelle "CHARGES FIXES MUTUALISEES(ne pas toucher)"
    # (bug corrigé le 17/09/2026 : la recherche exigeait auparavant un nom de dossier exact et ne
    # trouvait donc jamais ce fichier -> charge mutualisée systématiquement en DONNEES_MANQUANTES).
    fichiers = [f for f in glob.glob(os.path.join(sources_dir, "CHARGES FIXES MUTUALISEES*", "**", "*.xlsx"),
                                      recursive=True)
                if "~$" not in os.path.basename(f)]
    if not fichiers:
        return {}, ["CHARGES FIXES MUTUALISEES : aucun fichier trouvé — charge mutualisée non calculée"]
    wb = openpyxl.load_workbook(fichiers[0], data_only=True)
    ws = wb.active
    totaux = defaultdict(float)
    totaux_fichier = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[1] is None or row[2] is None:
            continue
        montant = parse_montant(row[2])
        if montant is None:
            continue
        poste = str(row[0] or "").strip().upper()
        soc = societe_canonique(row[1])
        if poste.startswith("TOTAL PAR CHAUFFEUR"):
            continue
        if poste.startswith("TOTAL"):
            if soc in ("Gleyzes", "LPB"):
                totaux_fichier[soc] = montant
            continue  # ne pas recompter la ligne de total elle-même dans le détail
        if soc in ("Gleyzes", "LPB"):
            totaux[soc] += montant

    manquants_coherence = []
    for soc, total_fichier in totaux_fichier.items():
        recalcule = totaux.get(soc, 0)
        if abs(recalcule - total_fichier) > 0.01:
            manquants_coherence.append(
                f"CHARGES FIXES MUTUALISEES : incohérence pour {soc} — la ligne TOTAL du fichier indique "
                f"{total_fichier:.2f} € mais la somme du détail donne {recalcule:.2f} € (écart {recalcule-total_fichier:+.2f} €). "
                f"Le calcul utilise le détail (plus fiable qu'un total qui peut être resté figé après une modification)."
            )

    nb_chauffeurs = defaultdict(set)
    for cid, infos in referentiel.items():
        if infos.get("Actif") != "O":
            continue
        chauffeur = (infos.get("Chauffeur") or "").strip()
        soc = societe_canonique(infos.get("Société"))
        if chauffeur and chauffeur != "-" and soc in ("Gleyzes", "LPB"):
            nb_chauffeurs[soc].add(chauffeur)

    resultat, manquants = {}, list(manquants_coherence)
    for soc in ("Gleyzes", "LPB"):
        if soc not in totaux:
            manquants.append(f"CHARGES FIXES MUTUALISEES : aucun montant trouvé pour {soc}")
            continue
        n = len(nb_chauffeurs.get(soc, []))
        if n == 0:
            manquants.append(f"CHARGES FIXES MUTUALISEES : aucun chauffeur actif identifié pour {soc}, "
                              f"division impossible")
            continue
        resultat[soc] = totaux[soc] / n
    return resultat, manquants


def parse_salaires(sources_dir, mois, referentiel):
    """Robuste à toute année/mois futur et à toute variante de nommage du
    sous-dossier (bug corrigé le 03/09/2026, même famille que KM/PEAGES/CA) :
    1) cherche un sous-dossier de PAIES dont le nom contient AAAA-MM (séparateur
       libre : -, _, espace, ou rien) ;
    2) à défaut, cherche récursivement tout PDF de PAIES dont le nom contient
       le mois en toutes lettres + l'année (ex: 'JUILLET_2026'), sans exiger de
       sous-dossier particulier.
    """
    annee, mm = mois.split("-")
    mois_nom_fr = {v: k for k, v in MOIS_FR.items() if len(k) > 3}.get(int(mm))  # nom long (MAI a un homonyme court, ok)

    pdfs = []
    for d in glob.glob(os.path.join(sources_dir, "PAIES", "*")):
        if os.path.isdir(d) and re.search(rf"{annee}[-_ ]?{mm}\b", os.path.basename(d)):
            pdfs = glob.glob(os.path.join(d, "**", "*.pdf"), recursive=True)
            break
    if not pdfs and mois_nom_fr:
        pdfs = [f for f in glob.glob(os.path.join(sources_dir, "PAIES", "**", "*.pdf"), recursive=True)
                if mois_nom_fr in os.path.basename(f).upper() and annee in os.path.basename(f)]

    nom_vers_camion = {}
    for cid, infos in referentiel.items():
        chauffeur = (infos.get("Chauffeur") or "").split()[-1] if infos.get("Chauffeur") else None
        if chauffeur:
            nom_vers_camion[chauffeur.upper()] = cid
    res, manquants = {}, []
    for f in pdfs:
        full = lire_texte_pdf(f)
        val = None
        for m in re.finditer(r"TOTAL VERSE PAR L.EMPLOYEUR\s*([\d\s]*[\d],\d{2})", full):
            val = float(m.group(1).replace(" ", "").replace(",", "."))
            break
        if val is None:
            manquants.append(f"PAIES : 'TOTAL VERSE PAR L'EMPLOYEUR' introuvable dans {os.path.basename(f)} "
                              f"(PDF non-OCR ? lancer un OCR Tesseract -l fra manuellement)")
            continue
        base = os.path.basename(f).upper()
        camion = None
        for nom, cid in nom_vers_camion.items():
            if nom in base:
                camion = cid
                break
        if not camion:
            base_norm = normalise_prenom(base)
            if any(nom in base_norm for nom in PERSONNEL_NON_CHAUFFEUR):
                continue  # personnel mutualisé, pas un chauffeur de camion : pas une anomalie
            manquants.append(f"PAIES : chauffeur non identifié pour le fichier {os.path.basename(f)}")
            continue
        # Bug corrigé le 17/09/2026 : certains chauffeurs ont leur paie scindée en deux
        # bulletins dans le même mois (ex: "1 AU 23 AOUT" + "24 AU 31 AOUT"). L'ancien code
        # écrasait la valeur du premier bulletin avec celle du second au lieu de les
        # additionner, sous-évaluant le salaire de moitié pour ces chauffeurs.
        res[camion] = res.get(camion, 0) + val
    if not pdfs:
        manquants.append(f"PAIES : aucun bulletin trouvé pour {mois} (ni sous-dossier PAIES/{mois}, "
                          f"ni fichier contenant '{mois_nom_fr or mois}' + '{annee}' dans PAIES/)")
    return res, manquants


# ---------------------------------------------------------------------------
# Programme principal
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Mise à jour mensuelle du tableau de bord flotte Gleyzes/LPB")
    ap.add_argument("--classeur", default="TABLEAU_DE_BORD_FLOTTE.xlsx")
    ap.add_argument("--sources", default="TABLEAUX DE BORD/SOURCES")
    ap.add_argument("--mois", default=None, help="Format AAAA-MM, ex: 2026-08. Auto-détecté si omis.")
    args = ap.parse_args()

    mois = detecter_mois(args.sources, args.mois)
    mois_date = date(int(mois.split("-")[0]), int(mois.split("-")[1]), 1)
    print(f"=== Mise à jour du mois {mois} ===")

    wb = openpyxl.load_workbook(args.classeur)
    referentiel = lire_referentiel(wb)
    actifs = {cid: v for cid, v in referentiel.items() if v.get("Actif") == "O"}

    manquants_globaux = []
    total_maj = 0  # lignes ajoutées + complétées + formules réparées, toutes tables (hors DONNEES_MANQUANTES)

    def _rapport(n, maj_c, rep):
        msg = f"{n} lignes ajoutées, {maj_c} complétées"
        if rep:
            msg += f", {rep} formule(s) resynchronisée(s) (corruption passée corrigée)"
        return msg

    # --- KM ---
    km, manq = parse_km(args.sources, mois)
    manquants_globaux += manq
    ws = get_table_ws(wb, "T_KM")
    lignes = [{"Mois": mois_date, "Camion_ID": cid, "Société": societe_canonique(v.get("Société")), "KM_Parcourus": km.get(cid)}
              for cid, v in actifs.items()]
    n, maj_c, rep = ajouter_lignes_table(ws, "T_KM", lignes, cle_dedup=("Mois", "Camion_ID"))
    total_maj += n + maj_c + rep
    print(f"KM : {_rapport(n, maj_c, rep)}")

    # --- CA ---
    ca, manq = parse_ca(args.sources, mois, referentiel)
    manquants_globaux += manq
    ws = get_table_ws(wb, "T_CA")
    lignes = []
    for cid, v in actifs.items():
        entrees = ca.get(cid)
        if entrees:
            for societe, montant in entrees:
                lignes.append({"Mois": mois_date, "Camion_ID": cid, "Société": societe, "CA": montant})
        else:
            lignes.append({"Mois": mois_date, "Camion_ID": cid, "Société": societe_canonique(v.get("Société")), "CA": None})
    n, maj_c, rep = ajouter_lignes_table(ws, "T_CA", lignes, cle_dedup=("Mois", "Camion_ID", "Société"))
    total_maj += n + maj_c + rep
    print(f"CA : {_rapport(n, maj_c, rep)}")

    # --- Carburant + Péages + Salaires -> Charges_Variables ---
    carb, manq = parse_carburant(args.sources, mois)
    manquants_globaux += manq
    peages, manq = parse_peages(args.sources, mois)
    manquants_globaux += manq
    salaires, manq = parse_salaires(args.sources, mois, referentiel)
    manquants_globaux += manq

    ws_cv = get_table_ws(wb, "T_ChargesVariables")
    formule_total_cv = lambda r: f"=SUM(D{r}:G{r})"
    lignes_cv = []
    for cid, v in actifs.items():
        societe_ppal = societe_canonique(v.get("Société"))
        peage_info = peages.get(cid, {})
        if cid == "GD042ZC":
            if "Gleyzes" in peage_info:
                lignes_cv.append({"Mois": mois_date, "Camion_ID": cid, "Société": "Gleyzes",
                                   "Carburant": None, "Peages": peage_info.get("Gleyzes"),
                                   "Salaire_Chauffeur": None, "Entretien_Hors_Contrat": None,
                                   "Remarque": "Péages réglés par Gleyzes sur ce camion"})
            lignes_cv.append({"Mois": mois_date, "Camion_ID": cid, "Société": "LPB",
                               "Carburant": carb.get(cid), "Peages": peage_info.get("LPB"),
                               "Salaire_Chauffeur": salaires.get(cid),
                               "Entretien_Hors_Contrat": None,
                               "Remarque": "Salaire GARCIA (LPB) + péages réglés par LPB"})
        else:
            carburant_val = carb.get(cid)
            peage_val = peage_info.get(societe_ppal) or (list(peage_info.values())[0] if peage_info else None)
            lignes_cv.append({"Mois": mois_date, "Camion_ID": cid, "Société": societe_ppal,
                               "Carburant": carburant_val, "Peages": peage_val,
                               "Salaire_Chauffeur": salaires.get(cid), "Entretien_Hors_Contrat": None,
                               "Remarque": "" if salaires.get(cid) is not None else "Bulletin de paie non fourni"})
    n, maj_c, rep = ajouter_lignes_table(ws_cv, "T_ChargesVariables", lignes_cv,
                              cle_dedup=("Mois", "Camion_ID", "Société"),
                              colonnes_formule={"Total": formule_total_cv})
    total_maj += n + maj_c + rep
    print(f"Charges_Variables : {_rapport(n, maj_c, rep)}")

    # --- Charges_Fixes (loyers + taxe essieu + charge mutualisée, stables -> copiés depuis REF) ---
    ws_cf = get_table_ws(wb, "T_ChargesFixes")
    formule_total_cf = lambda r: f"=SUM(D{r}:G{r})"
    CHARGE_MUT, manq = lire_charges_mutualisees(args.sources, referentiel)
    manquants_globaux += manq
    lignes_cf = []
    for cid, v in actifs.items():
        loyer_t = v.get("Loyer_Tracteur_HT") or 0
        loyer_r = v.get("Loyer_Remorque") or 0
        if cid == "GD042ZC":
            lignes_cf.append({"Mois": mois_date, "Camion_ID": cid, "Société": "Gleyzes",
                               "Loyer_Tracteur": loyer_t, "Loyer_Remorque": loyer_r,
                               "Taxe_Essieu_Mensuelle": 516 / 12, "Charge_Mutualisee_Chauffeur": 0,
                               "Remarque": "Financement à charge Gleyzes"})
            lignes_cf.append({"Mois": mois_date, "Camion_ID": cid, "Société": "LPB",
                               "Loyer_Tracteur": 0, "Loyer_Remorque": 0, "Taxe_Essieu_Mensuelle": 0,
                               "Charge_Mutualisee_Chauffeur": CHARGE_MUT.get("LPB"),
                               "Remarque": "Charge mutualisée chauffeur GARCIA (LPB)"})
        else:
            societe = societe_canonique(v.get("Société"))
            lignes_cf.append({"Mois": mois_date, "Camion_ID": cid, "Société": societe,
                               "Loyer_Tracteur": loyer_t, "Loyer_Remorque": loyer_r,
                               "Taxe_Essieu_Mensuelle": 516 / 12,
                               "Charge_Mutualisee_Chauffeur": CHARGE_MUT.get(societe), "Remarque": ""})
    n, maj_c, rep = ajouter_lignes_table(ws_cf, "T_ChargesFixes", lignes_cf,
                              cle_dedup=("Mois", "Camion_ID", "Société"),
                              colonnes_formule={"Total": formule_total_cf})
    total_maj += n + maj_c + rep
    print(f"Charges_Fixes : {_rapport(n, maj_c, rep)}")

    # --- Synthese_Camion / Societe / Entreprise : ajout des lignes Mois/Camion (le reste = formules) ---
    ws_sc = get_table_ws(wb, "T_SyntheseCamion")
    # Bug corrigé le 17/09/2026 : Charges_Fixes/Charges_Variables/Charges_Mutualisees/Charges_Totales
    # n'avaient pas de garde "N/D" comme KM et CA. Un mois sans encore aucune charge importée
    # affichait donc 0,00 € (résultat naturel d'un SUMIFS sans correspondance), ce qui donne
    # l'impression trompeuse que le mois a été traité avec un coût nul plutôt que "pas encore de
    # données". Toutes les colonnes calculées à partir de Charges_Fixes/Charges_Variables
    # basculent maintenant sur "N/D" tant qu'aucune ligne source n'existe pour ce mois/camion,
    # et propagent ce N/D en cascade (Cout_au_km, Resultat, Marge_%).
    formules_sc = {
        "Société": lambda r: f'=IFERROR(INDEX(REF_Camions!$B:$B,MATCH(B{r},REF_Camions!$A:$A,0)),"")',
        "Charges_Fixes": lambda r: (f'=IF(COUNTIFS(Charges_Fixes!$B:$B,B{r},Charges_Fixes!$A:$A,A{r})=0,"N/D",'
                                     f"SUMIFS(Charges_Fixes!$I:$I,Charges_Fixes!$B:$B,B{r},Charges_Fixes!$A:$A,A{r})-F{r})"),
        "Charges_Variables": lambda r: (f'=IF(COUNTIFS(Charges_Variables!$B:$B,B{r},Charges_Variables!$A:$A,A{r})=0,"N/D",'
                                         f"SUMIFS(Charges_Variables!$I:$I,Charges_Variables!$B:$B,B{r},Charges_Variables!$A:$A,A{r}))"),
        "Charges_Mutualisees": lambda r: (f'=IF(COUNTIFS(Charges_Fixes!$B:$B,B{r},Charges_Fixes!$A:$A,A{r})=0,"N/D",'
                                           f"SUMIFS(Charges_Fixes!$G:$G,Charges_Fixes!$B:$B,B{r},Charges_Fixes!$A:$A,A{r}))"),
        "Charges_Totales": lambda r: f'=IF(OR(D{r}="N/D",E{r}="N/D",F{r}="N/D"),"N/D",D{r}+E{r}+F{r})',
        "KM": lambda r: (f'=IF(COUNTIFS(KM!$B:$B,B{r},KM!$A:$A,A{r},KM!$D:$D,"<>")=0,"N/D",'
                         f"SUMIFS(KM!$D:$D,KM!$B:$B,B{r},KM!$A:$A,A{r}))"),
        "Cout_au_km": lambda r: f'=IF(OR(G{r}="N/D",H{r}="N/D",H{r}=0),"N/D",G{r}/H{r})',
        "CA": lambda r: (f'=IF(COUNTIFS(CA!$B:$B,B{r},CA!$A:$A,A{r},CA!$D:$D,"<>")=0,"N/D",'
                         f"SUMIFS(CA!$D:$D,CA!$B:$B,B{r},CA!$A:$A,A{r}))"),
        "Resultat": lambda r: f'=IF(OR(J{r}="N/D",G{r}="N/D"),"N/D",J{r}-G{r})',
        "Marge_%": lambda r: f'=IF(OR(K{r}="N/D",J{r}=0),"N/D",K{r}/J{r})',
    }
    lignes_sc = [{"Mois": mois_date, "Camion_ID": cid} for cid in actifs]
    n, maj_c, rep = ajouter_lignes_table(ws_sc, "T_SyntheseCamion", lignes_sc, cle_dedup=("Mois", "Camion_ID"),
                              colonnes_formule=formules_sc)
    total_maj += n + maj_c + rep
    print(f"Synthese_Camion : {_rapport(n, maj_c, rep)}")

    ws_ss = get_table_ws(wb, "T_SyntheseSociete")
    formules_ss = {
        "Charges_Fixes": lambda r: (f'=IF(COUNTIFS(Charges_Fixes!$C:$C,B{r},Charges_Fixes!$A:$A,A{r})=0,"N/D",'
                                     f"SUMIFS(Charges_Fixes!$I:$I,Charges_Fixes!$C:$C,B{r},Charges_Fixes!$A:$A,A{r})-E{r})"),
        "Charges_Variables": lambda r: (f'=IF(COUNTIFS(Charges_Variables!$C:$C,B{r},Charges_Variables!$A:$A,A{r})=0,"N/D",'
                                         f"SUMIFS(Charges_Variables!$I:$I,Charges_Variables!$C:$C,B{r},Charges_Variables!$A:$A,A{r}))"),
        "Charges_Mutualisees": lambda r: (f'=IF(COUNTIFS(Charges_Fixes!$C:$C,B{r},Charges_Fixes!$A:$A,A{r})=0,"N/D",'
                                           f"SUMIFS(Charges_Fixes!$G:$G,Charges_Fixes!$C:$C,B{r},Charges_Fixes!$A:$A,A{r}))"),
        "Charges_Totales": lambda r: f'=IF(OR(C{r}="N/D",D{r}="N/D",E{r}="N/D"),"N/D",C{r}+D{r}+E{r})',
        "CA": lambda r: (f'=IF(COUNTIFS(CA!$C:$C,B{r},CA!$A:$A,A{r},CA!$D:$D,"<>")=0,"N/D",'
                         f"SUMIFS(CA!$D:$D,CA!$C:$C,B{r},CA!$A:$A,A{r}))"),
        "Resultat": lambda r: f'=IF(OR(G{r}="N/D",F{r}="N/D"),"N/D",G{r}-F{r})',
        "Marge_%": lambda r: f'=IF(OR(H{r}="N/D",G{r}=0),"N/D",H{r}/G{r})',
    }
    lignes_ss = [{"Mois": mois_date, "Société": s} for s in ("Gleyzes", "LPB")]
    n, maj_c, rep = ajouter_lignes_table(ws_ss, "T_SyntheseSociete", lignes_ss, cle_dedup=("Mois", "Société"),
                              colonnes_formule=formules_ss)
    total_maj += n + maj_c + rep
    print(f"Synthese_Societe : {_rapport(n, maj_c, rep)}")

    ws_se = get_table_ws(wb, "T_SyntheseEntreprise")
    formules_se = {
        "CA_Global": lambda r: f'=IF(COUNTIFS(CA!$A:$A,A{r},CA!$D:$D,"<>")=0,"N/D",SUMIFS(CA!$D:$D,CA!$A:$A,A{r}))',
        "Charges_Mutualisees": lambda r: (f'=IF(COUNTIFS(Charges_Fixes!$A:$A,A{r})=0,"N/D",'
                                           f"SUMIFS(Charges_Fixes!$G:$G,Charges_Fixes!$A:$A,A{r}))"),
        "Charges_Totales_Flotte": lambda r: (
            f'=IF(AND(COUNTIFS(Charges_Fixes!$A:$A,A{r})=0,COUNTIFS(Charges_Variables!$A:$A,A{r})=0),"N/D",'
            f"SUMIFS(Charges_Fixes!$I:$I,Charges_Fixes!$A:$A,A{r})"
            f"+SUMIFS(Charges_Variables!$I:$I,Charges_Variables!$A:$A,A{r}))"),
        "Resultat": lambda r: f'=IF(OR(B{r}="N/D",D{r}="N/D"),"N/D",B{r}-D{r})',
        "Marge_%": lambda r: f'=IF(OR(E{r}="N/D",B{r}=0),"N/D",E{r}/B{r})',
    }
    n, maj_c, rep = ajouter_lignes_table(ws_se, "T_SyntheseEntreprise", [{"Mois": mois_date}], cle_dedup=("Mois",),
                              colonnes_formule=formules_se)
    total_maj += n + maj_c + rep
    print(f"Synthese_Entreprise : {_rapport(n, maj_c, rep)}")

    # --- DONNEES_MANQUANTES : journalisation des trous détectés ce mois-ci ---
    if manquants_globaux:
        ws_dm = get_table_ws(wb, "T_DonneesManquantes")
        lignes_dm = [{"Type de donnée": "Auto-détection script", "Camion(s) concerné(s)": "-",
                      "Période": mois, "Constat": m, "Action recommandée": "À vérifier"} for m in manquants_globaux]
        n, maj_c, rep = ajouter_lignes_table(ws_dm, "T_DonneesManquantes", lignes_dm,
                                  cle_dedup=("Constat", "Période"))
        print(f"DONNEES_MANQUANTES : {n} anomalies journalisées")
        for m in manquants_globaux:
            print("  [!]", m)

    if total_maj == 0:
        print(f"\n[!] ATTENTION : aucune ligne ajoutée ni complétée pour {mois} — le classeur ne va pas changer.")
        print(f"    Causes possibles : (1) {mois} est déjà entièrement à jour dans le classeur ; "
              f"(2) les documents que vous venez de déposer concernent en fait un autre mois que {mois} "
              f"(relancez avec --mois AAAA-MM pour forcer le bon mois) ; "
              f"(3) les fichiers déposés ne sont pas au bon endroit/format (voir DONNEES_MANQUANTES).")

    chemin_final = sauvegarder_avec_secours(wb, args.classeur)
    print(f"\nClasseur mis à jour : {chemin_final}")


if __name__ == "__main__":
    main()
