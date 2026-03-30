#!/usr/bin/env python3
"""
Normalisiert contracts_extracted_v2.json gegen Stammdaten (Aktive Firmen + Aktive Projekte).
- Firmennamen-Matching (exakt → normalisiert → fuzzy) mit Nummern-Schutz
- Zusatzdaten anreichern (IBAN, USt-IdNr, Handelsregister, Stammdaten-Adresse)
- Adressen aus Verträgen NICHT überschreiben (nur als stammdaten_adresse ergänzen)
- Projekt-IDs zuordnen
- Output: contracts_extracted_v3.json + normalization_report.txt
"""

import json
import re
from difflib import SequenceMatcher
from pathlib import Path
import pandas as pd

# --- Pfade ---
BASE_DIR = Path(__file__).parent
STAMMDATEN_DIR = BASE_DIR.parent / "Stammdaten"
JSON_INPUT = BASE_DIR / "data" / "contracts_extracted_v2.json"
JSON_OUTPUT = BASE_DIR / "data" / "contracts_extracted_v3.json"
REPORT_OUTPUT = BASE_DIR / "normalization_report.txt"

FIRMEN_XLSX = STAMMDATEN_DIR / "Aktive Firmen - Details 18_03_2026 15-23-59.xlsx"
PROJEKTE_XLSX = STAMMDATEN_DIR / "02 Aktive Projekte 18_03_2026 15-25-04.xlsx"

# --- Schwellenwerte ---
FUZZY_THRESHOLD = 0.88  # Höher als vorher wegen False-Positive-Risiko
FUZZY_REVIEW_THRESHOLD = 0.70

# --- Manuelles Mapping für bekannte Sonderfälle ---
# JSON-Name → Stammdaten-Name (oder None = kein Match, beibehalten)
MANUAL_MAPPING = {
    # Nummerierte Gesellschaften die fuzzy verwechselt werden
    "Windtaler III GmbH & Co. KG": "Windtaler 03 GmbH & Co. KG",
    "Windtaler III GmbH u. Co KG": "Windtaler 03 GmbH & Co. KG",
    "Windtaler IV GmbH & Co. KG": "Windtaler 04 GmbH & Co. KG",
    "Windtaler VII-1 GmbH u. Co KG": "Windtaler 07-1 GmbH & Co. KG",
    "Sonnentaler I GbR": "Sonnentaler 01 GbR",
    "Sonnentaler I GmbH & Co. KG": "Sonnentaler 01 GmbH & Co. KG",
    "Sonnentaler II GmbH & Co.KG": "Sonnentaler 02 GbR",
    "Sonnentaler III": "Sonnentaler 03 GmbH & Co. KG",
    "Sonnentaler III GbR": "Sonnentaler 03 GmbH & Co. KG",
    "Sonnentaler V GmbH & Co.KG": "Sonnentaler V GmbH & Co. KG",
    "Sonnentaler V GmbH u CoKG": "Sonnentaler V GmbH & Co. KG",
    "Sonnentaler IX GmbH & Co.KG": "Sonnentaler IX GmbH & Co. KG",
    "Sonnentaler XI GmbH & Co. KG": "Sonnentaler XI GmbH & Co. KG",
    "Sonnentaler XI GmbH u CoKG": "Sonnentaler XI GmbH & Co. KG",
    "Sonnentaler XI GmbH u. Co.KG": "Sonnentaler XI GmbH & Co. KG",
    "Sonnentaler XII GmbH u. Co.KG": "Sonnentaler XII GmbH & Co. KG",
    "Sonnentaler XV GmbH & Co. KG": "Sonnentaler XV GmbH & Co. KG",
    "Sonnentaler XV GmbH & Co.KG": "Sonnentaler XV GmbH & Co. KG",
    "Sonnentaler XVI GmbH & Co.KG": "Sonnentaler XVI GmbH & Co. KG",
    "Sonnentaler XVII GmbH & Co.KG": "Sonnentaler XVII GmbH & Co. KG",
    # Sonderfälle
    "Sonnentaler": None,  # Zu unspezifisch
    "Niedenhof Andrea": "Niedenhof Andrea",  # Nicht "Niedenhof Andrea FW"
    "Niedenhof Detlef / Taler GmbH & Co. KG": None,  # Kombination, kein einzelner Match
    "Ingenieurbüro Wintering": "Ingenieurbüro Wintering",  # Beibehalten, nicht KINDE
    "Peters, Niedenhof GbR": "Peters-Niedenhof GbR",
    "Spark-Niedenhof GbR": "Spark Niedenhof GbR",
    "Bioenergie Witte-Moor GmbH & Co. KG": "Bioenergie Witte Moor GmbH & Co. KG",
    "Biogas Büter und Heymann GmbH & Co KG": "Biogas Büter & Heymann GmbH & Co.KG",
    "Biogas Büter und Heymann GmbH & Co. KG": "Biogas Büter & Heymann GmbH & Co.KG",
    "Wilken Wilken, Niedenhof GbR": "Wilken Wilken Niedenhof Untiedt GbR",
    "Wilken & Spark Agrar GmbH": "Wilken und Spark Agrar GmbH",
    "Wilken und Spark Agrar GmbH": "Wilken und Spark Agrar GmbH",
    "Kirschner, Niedenhof GbR": "Kirschner Niedenhof GbR",
    "Flerlage Gefluegelzucht GmbH": "Flerlage Geflügelzucht GmbH",
    "Flerlage Geflügelzucht GmbH": "Flerlage Geflügelzucht GmbH",
    "Lohner Putenmast GmbH & Co.KG": "Lohner Putenmast GmbH & Co. KG",
    "Stefan Scholübbers GmbH & Co. KG": "Stefan Scholübbers GmbH & Co. KG",
    "Maschinenbau Manfred Kaiser": "Maschinenbau Manfred Kaiser GmbH",
    "Maschinenbau Manfred Kaiser GmbH": "Maschinenbau Manfred Kaiser GmbH",
    "HRN GbR": "HRN Solar GbR",
    "HRN Solar GbR": "HRN Solar GbR",
    "Hümmlinger Volksbank eG": "Hümmlinger Volksbank eG",
    "LHH Rühlertwist": "LHH Rühlertwist",
    "Hollander Jacob": "Hollander Jacob",
    "Hollander Jakob": "Hollander Jakob",
    "Hollander Werner": "Hollander Werner",
    "Gemeinde Vrees": "Gemeinde Vrees",
    "Stadtwerke Wesel GmbH": "Stadtwerke Wesel GmbH",
    "Quappen Immobilien GmbH & Co. KG": "Quappen Immobilien GmbH & Co. KG",
    "Ralf Schulte GmbH & Co. KG": "Ralf Schulte GmbH & Co. KG",
    "GESEVO GmbH": "GESEVO GmbH",
    "BWV Biowärme Vrees GmbH & Co. KG": "BWV Biowärme Vrees GmbH & Co. KG",
    "BBV GmbH & Co. KG": "BBV GmbH & Co. KG",
    "Buterei GmbH & Co. KG": "Buterei GmbH & Co. KG",
    "Nutraferm PetFood GmbH": "Nutraferm PetFood GmbH",
    "EDEKA Meyer": "EDEKA Meyer",
    "Investment-Gemeinschaft Sigiltrastrasse GbR": "Investment-Gemeinschaft Sigiltrastrasse GbR",
    "Goenniger Agrar e.K. / Reulmann Andreas": "Goenniger Agrar e.K. / Reulmann Andreas",
}


def normalize_rechtsform(name: str) -> str:
    """Vereinheitlicht Rechtsformen und Schreibweisen."""
    s = name.strip()
    s = s.replace("\u00ad", "")  # soft hyphen
    s = re.sub(r'GmbH\s*&\s*Co\s*\.?\s*KG', 'GmbH & Co. KG', s)
    s = re.sub(r'GmbH\s+u\.?\s*Co\s*\.?\s*KG', 'GmbH & Co. KG', s)
    s = re.sub(r'GmbH\s+und\s+Co\s*\.?\s*KG', 'GmbH & Co. KG', s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip()


def extract_number_part(name: str) -> str | None:
    """Extrahiert die Nummer/römische Ziffer aus Gesellschaftsnamen.
    z.B. 'Sonnentaler V GmbH' → 'V', 'Windtaler 03 GmbH' → '03', 'Windtaler VII-1' → 'VII-1'
    """
    patterns = [
        r'(?:Sonnentaler|Windtaler|Taler Wind)\s+([\dIVXLC][\d\-IVXLCa-z]*)',
    ]
    for pat in patterns:
        m = re.search(pat, name, re.IGNORECASE)
        if m:
            return m.group(1).upper().replace(' ', '')
    return None


ROMAN_TO_ARABIC = {
    'I': '01', 'II': '02', 'III': '03', 'IV': '04', 'V': '05',
    'VI': '06', 'VII': '07', 'VIII': '08', 'IX': '09', 'X': '10',
    'XI': '11', 'XII': '12', 'XIII': '13', 'XIV': '14', 'XV': '15',
    'XVI': '16', 'XVII': '17', 'XVIII': '18', 'XIX': '19', 'XX': '20',
}


def numbers_match(num_a: str | None, num_b: str | None) -> bool:
    """Prüft ob zwei Nummern/Ziffern equivalent sind (z.B. 'III' == '03')."""
    if num_a is None or num_b is None:
        return num_a is None and num_b is None
    a = num_a.upper()
    b = num_b.upper()
    if a == b:
        return True
    # Römisch zu Arabisch konvertieren
    a_arab = ROMAN_TO_ARABIC.get(a, a)
    b_arab = ROMAN_TO_ARABIC.get(b, b)
    # Zero-padded vergleichen
    try:
        return int(a_arab) == int(b_arab)
    except ValueError:
        return a_arab == b_arab


def normalize_for_comparison(name: str) -> str:
    """Aggressivere Normalisierung nur für Vergleichszwecke."""
    s = normalize_rechtsform(name).lower()
    for rf in ['gmbh & co. kg', 'gmbh', 'gbr', 'e.k.', 'ohg', 'eg', 'e.v.']:
        s = s.replace(rf, '')
    s = re.sub(r'[^a-zäöüß0-9 ]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def is_numbered_entity(name: str) -> bool:
    """Prüft ob der Name eine nummerierte Gesellschaft ist."""
    return bool(re.search(r'(?:Sonnentaler|Windtaler|Taler Wind)\s+[\dIVXLC]', name, re.IGNORECASE))


def build_firmen_lookup(df_firmen: pd.DataFrame) -> dict:
    """Baut Lookup-Dict für Firmennamen."""
    lookup = {}
    for _, row in df_firmen.iterrows():
        name = str(row.get("Firmenname", "")).strip()
        if not name:
            continue
        entry = {
            "firmenname": name,
            "firmenname_kurz": str(row.get("Firmenname (kurz)", "")).strip() if pd.notna(row.get("Firmenname (kurz)")) else None,
            "ort": str(row.get("Adresse 1: Ort", "")).strip() if pd.notna(row.get("Adresse 1: Ort")) else None,
            "strasse": None,
            "plz": str(row.get("Adresse 1: Postleitzahl", "")).strip() if pd.notna(row.get("Adresse 1: Postleitzahl")) else None,
            "iban": str(row.get("IBAN 1", "")).strip() if pd.notna(row.get("IBAN 1")) else None,
            "ust_id": str(row.get("USt-IdNr.", "")).strip() if pd.notna(row.get("USt-IdNr.")) else None,
            "handelsregister": str(row.get("Handelsregister", "")).strip() if pd.notna(row.get("Handelsregister")) else None,
        }
        # Strasse-Spalte finden
        for col in df_firmen.columns:
            if "stra" in col.lower() and "1" in col:
                val = row.get(col)
                if pd.notna(val):
                    entry["strasse"] = str(val).strip()
                break

        # Adresse zusammenbauen
        parts = []
        if entry["strasse"]:
            parts.append(entry["strasse"])
        if entry["plz"] and entry["ort"]:
            parts.append(f"{entry['plz']} {entry['ort']}")
        elif entry["ort"]:
            parts.append(entry["ort"])
        entry["adresse_komplett"] = ", ".join(parts) if parts else None

        lookup[name] = entry
    return lookup


def match_firma(json_name: str, firmen_lookup: dict) -> tuple:
    """
    Matcht einen JSON-Firmennamen gegen die Stammdaten.
    Returns: (stammdaten_name, match_type, confidence)
    """
    if not json_name:
        return (None, 'none', 0.0)

    # 0. Manuelles Mapping prüfen
    if json_name in MANUAL_MAPPING:
        mapped = MANUAL_MAPPING[json_name]
        if mapped is None:
            return (None, 'manual_skip', 1.0)
        # Prüfen ob das Ziel in den Stammdaten existiert
        if mapped in firmen_lookup:
            return (mapped, 'manual', 1.0)
        else:
            # Manuelles Mapping zeigt auf Namen der nicht in Stammdaten ist
            # → Rechtsform-normalisiert trotzdem als Match verwenden
            return (mapped, 'manual_nodata', 1.0)

    # 1. Exakter Match
    if json_name in firmen_lookup:
        return (json_name, 'exact', 1.0)

    # 2. Normalisierter Match (Rechtsform)
    json_norm = normalize_rechtsform(json_name)
    for stamm_name in firmen_lookup:
        if normalize_rechtsform(stamm_name) == json_norm:
            return (stamm_name, 'normalized', 0.98)

    # 3. Aggressiv normalisierter Match (ohne Rechtsformen)
    json_comp = normalize_for_comparison(json_name)
    for stamm_name in firmen_lookup:
        stamm_comp = normalize_for_comparison(stamm_name)
        if json_comp and stamm_comp and json_comp == stamm_comp:
            return (stamm_name, 'normalized', 0.95)

    # 4. Match gegen Kurzname
    for stamm_name, data in firmen_lookup.items():
        if data["firmenname_kurz"]:
            kurz = data["firmenname_kurz"]
            if json_name.lower() == kurz.lower() or normalize_rechtsform(json_name) == normalize_rechtsform(kurz):
                return (stamm_name, 'normalized', 0.93)

    # 5. Fuzzy Match — mit Nummern-Schutz
    json_is_numbered = is_numbered_entity(json_name)
    json_num = extract_number_part(json_name) if json_is_numbered else None

    best_match = None
    best_score = 0.0
    for stamm_name in firmen_lookup:
        stamm_comp = normalize_for_comparison(stamm_name)
        if not json_comp or not stamm_comp:
            continue

        # Nummern-Schutz: Bei nummerierten Entities müssen die Nummern übereinstimmen
        if json_is_numbered and is_numbered_entity(stamm_name):
            stamm_num = extract_number_part(stamm_name)
            if not numbers_match(json_num, stamm_num):
                continue  # Nummern passen nicht → Skip

        score = similarity(json_comp, stamm_comp)
        if score > best_score:
            best_score = score
            best_match = stamm_name

    if best_score >= FUZZY_THRESHOLD:
        return (best_match, 'fuzzy', best_score)
    elif best_score >= FUZZY_REVIEW_THRESHOLD:
        return (best_match, 'review', best_score)

    return (None, 'none', best_score)


def extract_projekt_id(projektkennung: str) -> str:
    m = re.match(r'(P\d{4})', str(projektkennung))
    return m.group(1) if m else None


def match_projekt(contract: dict, projekte_df: pd.DataFrame, firmen_mapping: dict) -> list:
    """Matcht einen Vertrag gegen Projekte basierend auf Firmenname + Standort/Anlage."""
    matches = []

    contract_firmen = set()
    for p_key in ['vertragspartner_1', 'vertragspartner_2']:
        partner = contract.get(p_key)
        if partner and partner.get('name'):
            contract_firmen.add(partner['name'].lower())
            mapped = firmen_mapping.get(partner['name'])
            if mapped:
                contract_firmen.add(mapped.lower())
    for field in ['auftraggeber', 'auftragnehmer']:
        val = contract.get(field)
        if val:
            contract_firmen.add(val.lower())

    standort = (contract.get('standort') or '').lower()
    anlage = (contract.get('anlage_bezeichnung') or '').lower()

    for _, proj in projekte_df.iterrows():
        proj_id = extract_projekt_id(proj.get('Projektkennung', ''))
        proj_kunde = str(proj.get('Kunde', '')).lower()
        proj_kennung = str(proj.get('Projektkennung', '')).lower()

        if not proj_id:
            continue

        # Kunde muss matchen
        kunde_match = False
        for firma in contract_firmen:
            if not firma:
                continue
            if firma == proj_kunde:
                kunde_match = True
                break
            norm_firma = normalize_for_comparison(firma)
            norm_kunde = normalize_for_comparison(proj_kunde)
            if norm_firma and norm_kunde and similarity(norm_firma, norm_kunde) > 0.85:
                kunde_match = True
                break

        if not kunde_match:
            continue

        # Standort/Anlage matchen für Spezifität
        if anlage and ('windtaler' in anlage or 'windpark' in anlage):
            anlage_terms = [t for t in anlage.split() if len(t) > 3]
            if any(term in proj_kennung for term in anlage_terms):
                matches.append(proj_id)
                continue

        if standort:
            standort_parts = [p.strip() for p in re.split(r'[,/]', standort) if len(p.strip()) > 3]
            for part in standort_parts:
                if part in proj_kennung:
                    matches.append(proj_id)
                    break

    return list(set(matches))


def main():
    print("=== Normalisierung starten ===\n")

    # 1. Daten laden
    print("Lade Stammdaten...")
    df_firmen = pd.read_excel(FIRMEN_XLSX)
    df_projekte = pd.read_excel(PROJEKTE_XLSX)
    print(f"  Firmen: {len(df_firmen)} Einträge")
    print(f"  Projekte: {len(df_projekte)} Einträge")

    print("Lade Verträge...")
    with open(JSON_INPUT, 'r', encoding='utf-8') as f:
        contracts = json.load(f)
    print(f"  Verträge: {len(contracts)}")

    # 2. Firmen-Lookup aufbauen
    firmen_lookup = build_firmen_lookup(df_firmen)
    print(f"  Firmen-Lookup: {len(firmen_lookup)} eindeutige Namen")

    # 3. Alle JSON-Firmennamen sammeln und matchen
    print("\nFirmennamen-Matching...")
    json_firmen = set()
    for v in contracts:
        for p_key in ['vertragspartner_1', 'vertragspartner_2', 'vertragspartner_3', 'vertragspartner_4', 'vertragspartner_5']:
            partner = v.get(p_key)
            if partner and partner.get('name'):
                json_firmen.add(partner['name'])

    firmen_mapping = {}
    match_details = {}

    for json_name in sorted(json_firmen):
        stamm_name, match_type, confidence = match_firma(json_name, firmen_lookup)
        firmen_mapping[json_name] = stamm_name
        match_details[json_name] = (stamm_name, match_type, confidence)

    # Statistiken
    stats = {}
    for _, (_, mtype, _) in match_details.items():
        stats[mtype] = stats.get(mtype, 0) + 1

    for k in ['exact', 'normalized', 'manual', 'manual_nodata', 'manual_skip', 'fuzzy', 'review', 'none']:
        if stats.get(k, 0) > 0:
            print(f"  {k}: {stats[k]}")

    # 4. Verträge normalisieren
    print("\nVerträge normalisieren...")
    changes_log = []
    auto_match_types = ('exact', 'normalized', 'fuzzy', 'manual', 'manual_nodata')

    for v in contracts:
        vid = v.get('vertrags_id', 'unknown')

        for p_key in ['vertragspartner_1', 'vertragspartner_2', 'vertragspartner_3', 'vertragspartner_4', 'vertragspartner_5']:
            partner = v.get(p_key)
            if not partner or not partner.get('name'):
                continue

            old_name = partner['name']
            stamm_name = firmen_mapping.get(old_name)
            match_info = match_details.get(old_name, (None, 'none', 0))

            if stamm_name and match_info[1] in auto_match_types:
                stamm_data = firmen_lookup.get(stamm_name, {})

                # Name normalisieren
                if old_name != stamm_name:
                    partner['name'] = stamm_name
                    changes_log.append(f"[{vid}] {p_key}.name: '{old_name}' → '{stamm_name}' ({match_info[1]}, {match_info[2]:.0%})")

                # Stammdaten-Adresse als ZUSÄTZLICHES Feld (Originaladresse NICHT überschreiben)
                if stamm_data.get('adresse_komplett'):
                    partner['stammdaten_adresse'] = stamm_data['adresse_komplett']

                # Zusatzdaten anreichern (nur wenn in Stammdaten vorhanden)
                if stamm_data.get('iban') and not partner.get('iban'):
                    partner['iban'] = stamm_data['iban']
                if stamm_data.get('ust_id') and not partner.get('ust_id'):
                    partner['ust_id'] = stamm_data['ust_id']
                if stamm_data.get('handelsregister') and not partner.get('handelsregister'):
                    partner['handelsregister'] = stamm_data['handelsregister']

        # Auftraggeber/Auftragnehmer normalisieren
        for field in ['auftraggeber', 'auftragnehmer']:
            old_val = v.get(field)
            if old_val and old_val in firmen_mapping:
                new_val = firmen_mapping[old_val]
                match_info = match_details.get(old_val, (None, 'none', 0))
                if new_val and old_val != new_val and match_info[1] in auto_match_types:
                    v[field] = new_val
                    changes_log.append(f"[{vid}] {field}: '{old_val}' → '{new_val}'")

        # Projekt-Matching
        matched_projects = match_projekt(v, df_projekte, firmen_mapping)
        if matched_projects:
            v['projekt_ids'] = sorted(matched_projects)

    # 5. Output schreiben
    print(f"\nSchreibe {JSON_OUTPUT}...")
    with open(JSON_OUTPUT, 'w', encoding='utf-8') as f:
        json.dump(contracts, f, ensure_ascii=False, indent=2)

    # 6. Report schreiben
    print(f"Schreibe {REPORT_OUTPUT}...")
    with open(REPORT_OUTPUT, 'w', encoding='utf-8') as f:
        f.write("=" * 80 + "\n")
        f.write("NORMALISIERUNGS-REPORT\n")
        f.write(f"Verträge: {len(contracts)} | Eindeutige Firmennamen: {len(json_firmen)}\n")
        f.write("=" * 80 + "\n\n")

        f.write("--- MATCHING-STATISTIK ---\n")
        for k in ['exact', 'normalized', 'manual', 'manual_nodata', 'manual_skip', 'fuzzy', 'review', 'none']:
            if stats.get(k, 0) > 0:
                f.write(f"  {k:20s} {stats[k]:>4}\n")
        f.write(f"  {'GESAMT':20s} {len(json_firmen):>4}\n\n")

        # Exakte + Normalisierte
        f.write("--- EXAKTE & NORMALISIERTE MATCHES ---\n")
        for name, (stamm, mtype, conf) in sorted(match_details.items()):
            if mtype == 'exact':
                f.write(f"  ✓ {name}\n")
            elif mtype == 'normalized':
                f.write(f"  ✓ '{name}' → '{stamm}'\n")
        f.write("\n")

        # Manuelle Matches
        f.write("--- MANUELLE ZUORDNUNGEN ---\n")
        for name, (stamm, mtype, conf) in sorted(match_details.items()):
            if mtype in ('manual', 'manual_nodata'):
                marker = " [nicht in Stammdaten]" if mtype == 'manual_nodata' else ""
                if name != stamm:
                    f.write(f"  ✓ '{name}' → '{stamm}'{marker}\n")
                else:
                    f.write(f"  ✓ '{name}' (beibehalten){marker}\n")
            elif mtype == 'manual_skip':
                f.write(f"  ⊘ '{name}' (übersprungen — zu unspezifisch)\n")
        f.write("\n")

        # Fuzzy
        f.write("--- FUZZY MATCHES (auto-akzeptiert, ≥88%) ---\n")
        for name, (stamm, mtype, conf) in sorted(match_details.items()):
            if mtype == 'fuzzy':
                f.write(f"  ~ '{name}' → '{stamm}' ({conf:.0%})\n")
        f.write("\n")

        # Review
        f.write("--- REVIEW NÖTIG (70-88%, NICHT automatisch übernommen) ---\n")
        for name, (stamm, mtype, conf) in sorted(match_details.items()):
            if mtype == 'review':
                f.write(f"  ? '{name}' → '{stamm}' ({conf:.0%}) ← BITTE PRÜFEN\n")
        f.write("\n")

        # Kein Match
        f.write("--- KEIN MATCH GEFUNDEN ---\n")
        for name, (stamm, mtype, conf) in sorted(match_details.items()):
            if mtype == 'none':
                f.write(f"  ✗ '{name}'\n")
        f.write("\n")

        # Änderungsprotokoll
        f.write("--- ÄNDERUNGSPROTOKOLL ---\n")
        f.write(f"Anzahl Änderungen: {len(changes_log)}\n\n")
        for change in changes_log:
            f.write(f"  {change}\n")
        f.write("\n")

        # Projekt-Zuordnungen
        f.write("--- PROJEKT-ZUORDNUNGEN ---\n")
        proj_count = 0
        for v in contracts:
            if v.get('projekt_ids'):
                proj_count += 1
                vid = v.get('vertrags_id', 'unknown')
                f.write(f"  {vid}: {', '.join(v['projekt_ids'])}\n")
        f.write(f"\n{proj_count} von {len(contracts)} Verträgen mit Projekt-ID zugeordnet.\n")

    print(f"\n=== Fertig! ===")
    print(f"  Änderungen: {len(changes_log)}")
    print(f"  Output: {JSON_OUTPUT}")
    print(f"  Report: {REPORT_OUTPUT}")


if __name__ == '__main__':
    main()
