# Changelog - TERP Vertragsmanagement

## v1.5.1 - 2026-03-26: Projekt-Enrichment + neue Filter

### Aenderungen
- **Projekt-Daten angereichert**: Alle Felder aus Projekte-Excel werden auf Projekt-Vertices uebernommen:
  Projektbereich, Beschreibung, Leistung kWp, Abrechnung EVU, VCOM-ID, Erzeugungsdatum,
  Direktvermarktungsdatum, Redispatch-Datum
- **Neuer Filter: Projektbereich** (PV, WIND, BIOGAS, IMMOBILIE, SONSTIGE, VERSORGUNG, ARCHITEKTUR)
  - Filtert automatisch den Graphen und die Tabellen, ohne dass ein Projekt manuell ausgewaehlt werden muss
- **Neuer Filter: Abrechnung EVU** (Nein, Monatlich, Jaehrlich)
  - Gleiche Funktionsweise wie Projektbereich
- **Projekt-Dropdown fix**: Zeigt jetzt korrekt `display_name` (P-Nummer + Beschreibung) statt roher Adressen
- **Multiselect-Chips CSS**: Kleinere Schrift (11px), volle Breite, Zeilenumbruch statt Abschneiden
- **Tooltips erweitert**: Leistung kWp, VCOM-ID, Abrechnung EVU, Erzeugungsdatum bei Projekt-Nodes

---

## v1.5.0 - 2026-03-26: Datenmodell-Optimierung

### Neues Skript: `optimize_graph.py`
Post-Processing-Skript das nach `generate_graph.py` laeuft und die Graph-Daten bereinigt.
Pipeline: `normalize_contracts.py` -> `generate_graph.py` -> `optimize_graph.py` -> `app.py`

### Aenderungen

**Phase 1: Vertragsarten normalisiert**
- "(Nachtrag)"-Suffix entfernt, stattdessen `ist_nachtrag: true` als eigene Property
- "Gestattungsvertrag (PV-Dachanlage)" -> "Gestattungsvertrag" + `vertragsart_subtyp: "PV-Dachanlage"`
- "Jahresabrechnung" -> "Sonstige" + `dokument_typ: "Jahresabrechnung"`
- Ergebnis: 16 Vertragsarten auf 11 saubere Typen reduziert

**Phase 2: Multi-Personen aufgesplittet**
- 9 Eintraege mit mehreren Personen in einer Zeile in Einzelpersonen aufgeteilt
- z.B. "Buesing Detlev, Buesing Petra" -> 2 separate Person-Vertices
- Alle Kanten vom kombinierten Vertex auf Einzel-Personen dupliziert (20 Kanten)

**Phase 3: Projekt P-Nummern zugewiesen**
- Projekte ueber Firma->Kunde->Projekte-Excel (274 Projekte) mit P-Nummern verknuepft
- 15 Projekte erfolgreich zugewiesen (z.B. P0031, P0035, P0067 etc.)
- 62 Projekte ohne Match als P9001-P9062 nummeriert
- Neue Properties: `projekt_id`, `display_name`, `standort_original`, `plz`, `ort`, `gemarkung`

**Phase 4: Firma/Person-Klassifizierung**
- 9 Firma-Eintraege mit rechtsform="Privatperson" als `ist_natuerliche_person: true` markiert
- Label bleibt "Firma" (sind formelle Vertragsparteien)

**Phase 5: App-Anpassungen**
- Projekt-Dropdown zeigt jetzt `display_name` (P-Nummer + Beschreibung) statt roher Adressen
- Predefined Query "Gestattungsvertraege Sonnentaler" aktualisiert
- Tooltips erweitert: Projekt-Nr, Bereich, Standort, PLZ, Ort, Gemarkung
- Firma-Tooltip zeigt "Natuerliche Person" wenn zutreffend
- Vertragsart-Tooltip zeigt Subtyp (z.B. "PV-Dachanlage")

### Dateien
- `optimize_graph.py` - NEU: Post-Processing-Skript
- `normalize_contracts.py` - aus PoC kopiert (Referenz)
- `generate_graph.py` - aus PoC kopiert (Referenz)
- `data/vertices.json` - optimiert (684 -> 693 Vertices nach Person-Splits)
- `data/edges.json` - optimiert (814 -> 824 Edges nach Kanten-Duplizierung)
- `data/contracts_metadata.json` - optimiert (Vertragsarten + Parteien)
- `data/*_pre_optimize.json` - Backups der Original-Daten
- `app.py` - Filter/Tooltip/Query-Anpassungen

### Bekannte Einschraenkungen
- P-Nummern-Zuordnung: nur 15 von 77 Projekten automatisch zugewiesen
  - Verbesserungsmoeglichkeit: erweitertes Fuzzy-Matching auf Standort/Adresse
- Person-Split: nur bekannte Faelle per Hardcoded-Map, kein automatischer Scan

---

## v1.4.8 - vorher: Fix sidebar toggle, improve tooltips, cleanup filters

## v1.4.1 - vorher: Fix filter reset button styling (TERP purple, white text)

## v1.4.0 - vorher: Fix slider/radio/button CSS, query_params compat, upgrade streamlit
