"""
optimize_graph.py — Post-Processing für TERP Graph-Datenmodell
Liest vertices.json, edges.json und contracts_metadata.json,
optimiert die Daten und schreibt sie zurück.

Pipeline: normalize_contracts.py → generate_graph.py → optimize_graph.py → app.py

Phasen:
  1. Vertragsarten normalisieren (Nachtrag-Suffix entfernen, Subtypen extrahieren)
  2. Multi-Personen-Einträge aufsplitten
  3. Projekt-P-Nummern zuweisen (via Firma→Kunde→Projekte-Excel)
  4. Firma/Person-Klassifizierung bereinigen
"""

import json
import os
import re
import shutil
from collections import defaultdict
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
INPUT_DIR = BASE_DIR / "input"

# ── PHASE 1: Vertragsarten-Normalisierung ──

VERTRAGSART_NORMALIZE = {
    "Gestattungsvertrag (PV-Dachanlage)": {
        "vertragsart": "Gestattungsvertrag",
        "vertragsart_subtyp": "PV-Dachanlage",
    },
    "Gestattungsvertrag (Nachtrag)": {
        "vertragsart": "Gestattungsvertrag",
        "ist_nachtrag": True,
    },
    "Wärmeliefervertrag (Nachtrag)": {
        "vertragsart": "Wärmeliefervertrag",
        "ist_nachtrag": True,
    },
    "Stromliefervertrag (Nachtrag)": {
        "vertragsart": "Stromliefervertrag",
        "ist_nachtrag": True,
    },
    "Kaufvertrag (Nachtrag)": {
        "vertragsart": "Kaufvertrag",
        "ist_nachtrag": True,
    },
    "Jahresabrechnung": {
        "vertragsart": "Sonstige",
        "dokument_typ": "Jahresabrechnung",
    },
}

# ── PHASE 2: Personen-Split-Map ──

PERSON_SPLIT_MAP = {
    # Person-Vertices
    "Büsing Detlev, Büsing Petra": ["Büsing Detlev", "Büsing Petra"],
    "Hoesen Karl Heinz / Hoesen Hans-Jürgen": ["Hoesen Karl Heinz", "Hoesen Hans-Jürgen"],
    "Kemper Berendina / Nüsse Traute Johanne": ["Kemper Berendina", "Nüsse Traute Johanne"],
    "Kremers Gezienus und Kiena": ["Kremers Gezienus", "Kremers Kiena"],
    "Kleene Christine, König Henning": ["Kleene Christine", "König Henning"],
    "Steineck Manfred, Mular Irena": ["Steineck Manfred", "Mular Irena"],
    "Lünswilken Josef u. Sabine": ["Lünswilken Josef", "Lünswilken Sabine"],
}

# Firma-Vertices die Personen-Splits brauchen (nur echte Multi-Personen, keine GbR/GmbH)
FIRMA_PERSON_SPLIT_MAP = {
    "Vagedes Bernhard / Vagedes Guido (Gebaeudeigentuemer)": ["Vagedes Bernhard", "Vagedes Guido"],
    "Wraga Iris und Matthias": ["Wraga Iris", "Wraga Matthias"],
}

# ── PHASE 4: Firma/Person-Klassifizierung ──

# Firma-Einträge die NICHT als natürliche Person markiert werden (sind echte Firmen)
FIRMA_KEEP_AS_COMPANY = {
    "Janssen, Niedenhof GbR",
    "Niedenhof Detlef / Taler GmbH & Co. KG",
    "Wilken und Spark Agrar GmbH",
    "Erbengemeinschaft/Besitzgesellschaft Rosen Gerda, Rosen Ralf",
    "Goenniger Agrar e.K. / Reulmann Andreas",
    "Andreas & Hans Janssen",
}


def make_id(prefix: str, name: str) -> str:
    """Generate vertex ID from prefix and name (same logic as generate_graph.py)."""
    slug = name.lower()
    slug = re.sub(r"[^a-z0-9äöüß]+", "_", slug)
    slug = slug.strip("_")[:60]
    return f"{prefix}:{slug}"


def load_data():
    """Load vertices, edges, contracts_metadata from JSON."""
    with open(DATA_DIR / "vertices.json", "r", encoding="utf-8") as f:
        vertices = json.load(f)
    with open(DATA_DIR / "edges.json", "r", encoding="utf-8") as f:
        edges = json.load(f)
    with open(DATA_DIR / "contracts_metadata.json", "r", encoding="utf-8") as f:
        metadata = json.load(f)
    return vertices, edges, metadata


def save_data(vertices, edges, metadata):
    """Write optimized data back to JSON."""
    with open(DATA_DIR / "vertices.json", "w", encoding="utf-8") as f:
        json.dump(vertices, f, ensure_ascii=False, indent=2)
    with open(DATA_DIR / "edges.json", "w", encoding="utf-8") as f:
        json.dump(edges, f, ensure_ascii=False, indent=2)
    with open(DATA_DIR / "contracts_metadata.json", "w", encoding="utf-8") as f:
        json.dump(metadata, f, ensure_ascii=False, indent=2)


def backup_data():
    """Create backup of original data files."""
    for fname in ["vertices.json", "edges.json", "contracts_metadata.json"]:
        src = DATA_DIR / fname
        dst = DATA_DIR / fname.replace(".json", "_pre_optimize.json")
        if src.exists() and not dst.exists():
            shutil.copy2(src, dst)
            print(f"  Backup: {fname} -> {dst.name}")


def load_projekte_excel():
    """Load Projekte Excel and build Kunde→P-Nummer mapping."""
    excel_files = list(INPUT_DIR.glob("Projekte*.xlsx"))
    if not excel_files:
        print("  WARNUNG: Keine Projekte-Excel gefunden in input/")
        return {}, {}

    df = pd.read_excel(excel_files[0])
    print(f"  Projekte-Excel geladen: {len(df)} Zeilen")

    # Build Kunde → list of (P-number, Projektkennung)
    kunde_projects = defaultdict(list)
    pnum_details = {}

    for _, row in df.iterrows():
        pk = str(row.get("Projektkennung", ""))
        m = re.match(r"(P\d{4})", pk)
        if not m:
            continue
        pnum = m.group(1)
        kunde = str(row.get("Kunde", "")).strip()
        beschr = str(row.get("Beschreibung", "")).strip()
        bereich = str(row.get("Projektbereich", "")).strip()

        # Additional fields from Excel
        leistung = row.get("Leistung")
        abrechnung_evu = str(row.get("Abrechnung EVU", "")).strip()
        vcom_id = str(row.get("VCOM ID", "")).strip()
        soll_erzeugung = str(row.get("Soll Beleg Erzeugung", "")).strip()
        erzeugung_datum = str(row.get("Erzeugung Datum", "")).strip()
        soll_direktvermarktung = str(row.get("Soll Beleg Direktvermarktung", "")).strip()
        direktvermarktung_datum = str(row.get("Direktvermarktung Datum", "")).strip()
        soll_redispatch = str(row.get("Soll Beleg Redispatch", "")).strip()
        redispatch_datum = str(row.get("Redispatch Datum", "")).strip()

        kunde_projects[kunde].append(pnum)
        pnum_details[pnum] = {
            "projektkennung": pk,
            "kunde": kunde,
            "beschreibung": beschr,
            "bereich": bereich,
            "leistung_kwp": float(leistung) if pd.notna(leistung) else None,
            "abrechnung_evu": abrechnung_evu if abrechnung_evu and abrechnung_evu != "nan" else None,
            "vcom_id": vcom_id if vcom_id and vcom_id != "nan" else None,
            "soll_erzeugung": soll_erzeugung if soll_erzeugung != "nan" else None,
            "erzeugung_datum": erzeugung_datum if erzeugung_datum != "nan" else None,
            "soll_direktvermarktung": soll_direktvermarktung if soll_direktvermarktung != "nan" else None,
            "direktvermarktung_datum": direktvermarktung_datum if direktvermarktung_datum != "nan" else None,
            "soll_redispatch": soll_redispatch if soll_redispatch != "nan" else None,
            "redispatch_datum": redispatch_datum if redispatch_datum != "nan" else None,
        }

    return kunde_projects, pnum_details


# =============================================
# PHASE 1: Vertragsarten normalisieren
# =============================================

def phase1_normalize_vertragsarten(vertices, metadata):
    """Normalize Vertragsart values and extract Nachtrag/Subtyp as separate properties."""
    count = 0

    for v in vertices:
        if v.get("label") != "Vertrag":
            continue
        art = v.get("vertragsart", "")
        if art in VERTRAGSART_NORMALIZE:
            updates = VERTRAGSART_NORMALIZE[art]
            v["vertragsart"] = updates["vertragsart"]
            if "ist_nachtrag" in updates:
                v["ist_nachtrag"] = True
            if "vertragsart_subtyp" in updates:
                v["vertragsart_subtyp"] = updates["vertragsart_subtyp"]
            if "dokument_typ" in updates:
                v["dokument_typ"] = updates["dokument_typ"]
            count += 1

    # Same for contracts_metadata
    meta_count = 0
    for m in metadata:
        art = m.get("vertragsart", "")
        if art in VERTRAGSART_NORMALIZE:
            updates = VERTRAGSART_NORMALIZE[art]
            m["vertragsart"] = updates["vertragsart"]
            if "ist_nachtrag" in updates:
                m["ist_nachtrag"] = True
            if "vertragsart_subtyp" in updates:
                m["vertragsart_subtyp"] = updates["vertragsart_subtyp"]
            if "dokument_typ" in updates:
                m["dokument_typ"] = updates["dokument_typ"]
            meta_count += 1

    return count, meta_count


# =============================================
# PHASE 2: Multi-Personen aufsplitten
# =============================================

def phase2_split_persons(vertices, edges, metadata):
    """Split multi-person entries into individual person vertices."""
    splits_done = 0
    vertex_id_map = {}  # old_id → [new_id1, new_id2, ...]

    # Process Person splits
    for combined_name, individual_names in PERSON_SPLIT_MAP.items():
        old_id = make_id("person", combined_name)
        # Find vertex
        old_vertex = next((v for v in vertices if v["id"] == old_id), None)
        if not old_vertex:
            # Try finding by name match
            old_vertex = next(
                (v for v in vertices if v.get("name") == combined_name and v.get("label") == "Person"),
                None,
            )
        if not old_vertex:
            print(f"  WARNUNG: Person '{combined_name}' nicht gefunden")
            continue

        old_id = old_vertex["id"]
        new_ids = []
        for name in individual_names:
            new_id = make_id("person", name)
            new_vertex = {
                "id": new_id,
                "label": "Person",
                "name": name,
            }
            # Copy additional properties
            for k, val in old_vertex.items():
                if k not in ("id", "label", "name"):
                    new_vertex[k] = val
            vertices.append(new_vertex)
            new_ids.append(new_id)

        vertex_id_map[old_id] = new_ids
        vertices.remove(old_vertex)
        splits_done += 1

    # Process Firma→Person splits
    for combined_name, individual_names in FIRMA_PERSON_SPLIT_MAP.items():
        old_vertex = next(
            (v for v in vertices if v.get("name") == combined_name and v.get("label") == "Firma"),
            None,
        )
        if not old_vertex:
            print(f"  WARNUNG: Firma '{combined_name}' nicht gefunden")
            continue

        old_id = old_vertex["id"]
        new_ids = []
        for name in individual_names:
            new_id = make_id("person", name)
            new_vertex = {
                "id": new_id,
                "label": "Person",
                "name": name,
                "ist_natuerliche_person": True,
            }
            if old_vertex.get("adresse"):
                new_vertex["adresse"] = old_vertex["adresse"]
            vertices.append(new_vertex)
            new_ids.append(new_id)

        vertex_id_map[old_id] = new_ids
        vertices.remove(old_vertex)
        splits_done += 1

    # Duplicate edges for split vertices
    edges_added = 0
    edges_to_add = []
    edges_to_remove = []

    for edge in edges:
        src = edge["src"]
        dst = edge["dst"]
        if src in vertex_id_map:
            edges_to_remove.append(edge)
            for new_id in vertex_id_map[src]:
                new_edge = dict(edge)
                new_edge["src"] = new_id
                edges_to_add.append(new_edge)
                edges_added += 1
        elif dst in vertex_id_map:
            edges_to_remove.append(edge)
            for new_id in vertex_id_map[dst]:
                new_edge = dict(edge)
                new_edge["dst"] = new_id
                edges_to_add.append(new_edge)
                edges_added += 1

    for e in edges_to_remove:
        edges.remove(e)
    edges.extend(edges_to_add)

    # Update metadata parteien
    for m in metadata:
        parteien = m.get("parteien", [])
        new_parteien = []
        for p in parteien:
            pname = p.get("name", "")
            # Check all split maps
            found_split = False
            for combined, individuals in {**PERSON_SPLIT_MAP, **FIRMA_PERSON_SPLIT_MAP}.items():
                if pname == combined:
                    for ind_name in individuals:
                        new_p = dict(p)
                        new_p["name"] = ind_name
                        new_p["typ"] = "Person"
                        new_parteien.append(new_p)
                    found_split = True
                    break
            if not found_split:
                new_parteien.append(p)
        m["parteien"] = new_parteien

    return splits_done, edges_added


# =============================================
# PHASE 3: Projekt P-Nummern zuweisen
# =============================================

def phase3_assign_projekt_numbers(vertices, edges, metadata):
    """Assign P-numbers to Projekt vertices via Firma→Kunde→Projekte mapping."""
    kunde_projects, pnum_details = load_projekte_excel()
    if not kunde_projects:
        return 0, 0

    # Build lookup: vertex_id → vertex for fast access
    vertex_by_id = {v["id"]: v for v in vertices}

    # Build reverse edge lookups
    # HAT_STANDORT: vertrag → standort
    standort_to_vertraege = defaultdict(list)  # standort_id → [vertrag_id, ...]
    for e in edges:
        if e["relationship"] == "HAT_STANDORT":
            standort_to_vertraege[e["dst"]].append(e["src"])

    # PARTEI_VON: firma/person → vertrag
    vertrag_to_firmen = defaultdict(list)  # vertrag_id → [firma_name, ...]
    for e in edges:
        if e["relationship"] == "PARTEI_VON":
            src_v = vertex_by_id.get(e["src"])
            if src_v and src_v.get("label") == "Firma":
                vertrag_to_firmen[e["dst"]].append(src_v.get("name", ""))

    matched = 0
    unmatched_counter = 9000  # P9001, P9002, ...

    for v in vertices:
        if v.get("label") != "Projekt":
            continue

        standort_id = v["id"]
        original_name = v.get("name", "")

        # Find connected Firmen via: Standort ← HAT_STANDORT ← Vertrag ← PARTEI_VON ← Firma
        connected_firmen = set()
        for vertrag_id in standort_to_vertraege.get(standort_id, []):
            for firma_name in vertrag_to_firmen.get(vertrag_id, []):
                connected_firmen.add(firma_name)

        # Try to match Firma → Kunde in Projekte Excel
        best_pnum = None
        best_pk = None

        for firma_name in connected_firmen:
            projects = kunde_projects.get(firma_name, [])
            if not projects:
                continue

            if len(projects) == 1:
                # Unique match
                best_pnum = projects[0]
                best_pk = pnum_details[best_pnum]["projektkennung"]
                break
            else:
                # Multiple projects for this Kunde → try address/name matching
                standort_lower = original_name.lower()
                for pnum in projects:
                    pk = pnum_details[pnum]["projektkennung"].lower()
                    # Check if standort words appear in Projektkennung
                    standort_words = [w for w in re.split(r"[\s,/]+", standort_lower) if len(w) > 3]
                    match_score = sum(1 for w in standort_words if w in pk)
                    if match_score >= 2:
                        best_pnum = pnum
                        best_pk = pnum_details[pnum]["projektkennung"]
                        break

                if best_pnum:
                    break

        if best_pnum:
            details = pnum_details[best_pnum]
            # Extract clean display name from Projektkennung
            pk_clean = re.sub(r"^P\d{4}\s*-\s*", "", details["projektkennung"]).strip()
            v["projekt_id"] = best_pnum
            v["display_name"] = f"{best_pnum} - {pk_clean}"
            v["standort_original"] = original_name
            # All enrichment fields from Projekte Excel
            v["projektbereich"] = details.get("bereich") or ""
            v["beschreibung"] = details.get("beschreibung") or ""
            if details.get("leistung_kwp"):
                v["leistung_kwp"] = details["leistung_kwp"]
            if details.get("abrechnung_evu"):
                v["abrechnung_evu"] = details["abrechnung_evu"]
            if details.get("vcom_id"):
                v["vcom_id"] = details["vcom_id"]
            if details.get("soll_erzeugung"):
                v["soll_erzeugung"] = details["soll_erzeugung"]
            if details.get("erzeugung_datum"):
                v["erzeugung_datum"] = details["erzeugung_datum"]
            if details.get("soll_direktvermarktung"):
                v["soll_direktvermarktung"] = details["soll_direktvermarktung"]
            if details.get("direktvermarktung_datum"):
                v["direktvermarktung_datum"] = details["direktvermarktung_datum"]
            if details.get("soll_redispatch"):
                v["soll_redispatch"] = details["soll_redispatch"]
            if details.get("redispatch_datum"):
                v["redispatch_datum"] = details["redispatch_datum"]
            matched += 1
        else:
            unmatched_counter += 1
            pnum = f"P{unmatched_counter}"
            v["projekt_id"] = pnum
            v["display_name"] = f"{pnum} - {original_name}"
            v["standort_original"] = original_name

    # Parse address components from standort_original
    address_parsed = 0
    for v in vertices:
        if v.get("label") != "Projekt":
            continue
        name = v.get("standort_original", v.get("name", ""))
        plz_match = re.search(r"\b(\d{5})\b", name)
        if plz_match:
            v["plz"] = plz_match.group(1)
            address_parsed += 1
        gemarkung_match = re.search(r"Gemarkung\s+([^,]+)", name, re.IGNORECASE)
        if gemarkung_match:
            v["gemarkung"] = gemarkung_match.group(1).strip()
        flur_match = re.search(r"Flur\s+(\d+)", name, re.IGNORECASE)
        if flur_match:
            v["flur"] = flur_match.group(1)

        # Extract Ort (first word/phrase before comma or PLZ)
        ort_match = re.match(r"(?:\d{5}\s+)?([A-Za-zÄÖÜäöüß\-/]+(?:\s+[A-Za-zÄÖÜäöüß\-/]+)?)", name)
        if ort_match:
            ort = ort_match.group(1).strip()
            if ort and ort not in ("Windpark",):
                v["ort"] = ort

    unmatched = sum(1 for v in vertices if v.get("label") == "Projekt" and v.get("projekt_id", "").startswith("P9"))
    return matched, unmatched


# =============================================
# PHASE 4: Firma/Person-Klassifizierung
# =============================================

def phase4_fix_firma_classification(vertices):
    """Add ist_natuerliche_person flag to Privatperson Firma entries."""
    count = 0
    for v in vertices:
        if v.get("label") != "Firma":
            continue

        name = v.get("name", "")
        rechtsform = v.get("rechtsform", "") or ""

        # Skip entries that are actual companies
        if name in FIRMA_KEEP_AS_COMPANY:
            continue

        # Check for Privatperson indicators
        is_natural = False
        if "Privatperson" in rechtsform:
            is_natural = True
        elif "natürliche Person" in rechtsform:
            is_natural = True
        elif rechtsform in (None, "", "None", "Unternehmen"):
            # Check if name looks like a person name (Lastname Firstname pattern)
            # Skip if it contains company indicators
            company_indicators = ["GmbH", "GbR", "KG", "AG", "eG", "e.K.", "Solar", "Agrar",
                                  "Bioenergie", "Biogas", "Windpark", "Energiepark", "Erben",
                                  "Infrastruktur", "EDEKA", "Stadtwerke"]
            if not any(ind in name for ind in company_indicators):
                # Looks like a person name
                is_natural = True

        if is_natural:
            v["ist_natuerliche_person"] = True
            count += 1

    return count


# =============================================
# MAIN
# =============================================

def main():
    print("=" * 60)
    print("TERP Graph-Datenmodell Optimierung")
    print("=" * 60)

    # Load data
    print("\n[1] Daten laden...")
    vertices, edges, metadata = load_data()
    print(f"  Vertices: {len(vertices)}, Edges: {len(edges)}, Metadata: {len(metadata)}")

    # Backup
    print("\n[2] Backup erstellen...")
    backup_data()

    # Phase 1
    print("\n-- Phase 1: Vertragsarten normalisieren --")
    v_count, m_count = phase1_normalize_vertragsarten(vertices, metadata)
    print(f"  Vertices normalisiert: {v_count}")
    print(f"  Metadata normalisiert: {m_count}")

    # Show final Vertragsarten
    from collections import Counter
    arts = Counter(v.get("vertragsart") for v in vertices if v.get("label") == "Vertrag")
    print("  Ergebnis:")
    for art, c in arts.most_common():
        print(f"    {art}: {c}")

    # Phase 2
    print("\n-- Phase 2: Multi-Personen aufsplitten --")
    splits, edges_added = phase2_split_persons(vertices, edges, metadata)
    print(f"  Personen-Splits: {splits}")
    print(f"  Kanten dupliziert: {edges_added}")

    # Phase 3
    print("\n-- Phase 3: Projekt P-Nummern zuweisen --")
    matched, unmatched = phase3_assign_projekt_numbers(vertices, edges, metadata)
    print(f"  Zugewiesen (aus Excel): {matched}")
    print(f"  Nicht zugewiesen (P9xxx): {unmatched}")

    # Phase 4
    print("\n-- Phase 4: Firma/Person-Klassifizierung --")
    nat_count = phase4_fix_firma_classification(vertices)
    print(f"  Als natürliche Person markiert: {nat_count}")

    # Save
    print("\n[SAVE] Optimierte Daten speichern...")
    save_data(vertices, edges, metadata)

    # Final stats
    from collections import Counter
    labels = Counter(v.get("label") for v in vertices)
    print(f"\n{'=' * 60}")
    print("ERGEBNIS")
    print(f"{'=' * 60}")
    print(f"  Vertices: {len(vertices)}")
    for l, c in labels.most_common():
        print(f"    {l}: {c}")
    print(f"  Edges: {len(edges)}")
    print(f"  Metadata: {len(metadata)}")

    # Show P-number assignment summary
    projekte = [v for v in vertices if v.get("label") == "Projekt"]
    with_p = [v for v in projekte if v.get("projekt_id") and not v["projekt_id"].startswith("P9")]
    without_p = [v for v in projekte if v.get("projekt_id", "").startswith("P9")]
    print(f"\n  Projekte mit P-Nummer: {len(with_p)}")
    print(f"  Projekte ohne Match (P9xxx): {len(without_p)}")
    if without_p:
        print("  Unmatched Projekte:")
        for p in sorted(without_p, key=lambda x: x.get("projekt_id", "")):
            print(f"    {p['projekt_id']}: {p.get('standort_original', p.get('name', ''))}")

    print(f"\n[DONE] Optimierung abgeschlossen!")


if __name__ == "__main__":
    main()
