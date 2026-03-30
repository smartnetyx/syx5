# Plan: TERP Graph-Datenmodell Optimierung

**Datum**: 2026-03-26
**Status**: Geplant

## Context
Die Daten im TERP Vertragsmanagement-Graph sind für eine Kundenpräsentation nicht sauber genug: Projekte haben keine P-Nummern (nur Adressen), Vertragsarten sind nicht normalisiert (Nachtrag-Suffix), mehrere Personen stehen in einer Zeile, und Privatpersonen sind als Firma klassifiziert. Das Ziel ist ein deutlich aufgeräumteres Datenmodell.

## Ansatz: Neues `optimize_graph.py` Skript (Post-Processing)

Statt `generate_graph.py` zu modifizieren, ein neues Skript das NACH der Graph-Generierung läuft. Damit bleibt die Pipeline reversibel.

**Pipeline**: `normalize_contracts.py` → `generate_graph.py` → **`optimize_graph.py`** (NEU) → `app.py`

**Alle Dateien leben im syx5 Streamlit Deployment Repository.**
Relevante Skripte aus dem PoC-Ordner werden hierher kopiert.

**Ordnerstruktur (Ziel)**:
```
syx5 Streamlit Deployment/
├── app.py                          ← Streamlit App (v1.4.8, wird angepasst)
├── optimize_graph.py               ← NEU: Post-Processing-Skript
├── normalize_contracts.py          ← aus PoC kopiert (Referenz/Re-Run)
├── generate_graph.py               ← aus PoC kopiert (Referenz/Re-Run)
├── data/
│   ├── vertices.json               ← wird optimiert
│   ├── edges.json                  ← wird optimiert
│   └── contracts_metadata.json     ← wird optimiert
├── input/
│   ├── Firmen 18_03_2026 15-23-59.xlsx
│   └── Projekte 18_03_2026 15-25-04.xlsx
├── Pläne/
│   └── 20260326_Graph-Datenmodell-Optimierung.md  ← dieser Plan
├── requirements.txt
├── README.md
└── .streamlit/
```

---

## Phase 1: Vertragsarten normalisieren

| Aktuell | Wird zu | Eigenschaften |
|---|---|---|
| Gestattungsvertrag (PV-Dachanlage) | Gestattungsvertrag | vertragsart_subtyp="PV-Dachanlage" |
| Gestattungsvertrag (Nachtrag) | Gestattungsvertrag | ist_nachtrag=true |
| Wärmeliefervertrag (Nachtrag) | Wärmeliefervertrag | ist_nachtrag=true |
| Stromliefervertrag (Nachtrag) | Stromliefervertrag | ist_nachtrag=true |
| Kaufvertrag (Nachtrag) | Kaufvertrag | ist_nachtrag=true |
| Jahresabrechnung | Sonstige | dokument_typ="Jahresabrechnung" |

→ 16 Werte werden zu ~10 saubere Typen. Auch `contracts_metadata.json` aktualisieren.

## Phase 2: Personen aufsplitten

Hardcoded Split-Map:

| Aktuell | Wird zu |
|---|---|
| Büsing Detlev, Büsing Petra | 2 Personen |
| Hoesen Karl Heinz / Hoesen Hans-Jürgen | 2 Personen |
| Kemper Berendina / Nüsse Traute Johanne | 2 Personen |
| Kremers Gezienus und Kiena | Kremers Gezienus + Kremers Kiena |
| Kleene Christine, König Henning | 2 Personen |
| Steineck Manfred, Mular Irena | 2 Personen |

Auch in Firma-Vertices prüfen (z.B. "Vagedes Bernhard / Vagedes Guido"). Bei Split: alle Kanten vom kombinierten Vertex auf die Einzel-Personen duplizieren.

## Phase 3: Projekt P-Nummern zuweisen

**Algorithmus** (Firma→Kunde→Projekte Beziehung):
1. Projekte-Excel laden (`input/Projekte *.xlsx`, 274 Projekte, Spalte "Kunde")
2. Für jeden Projekt-Vertex in vertices.json:
   a. Über HAT_STANDORT-Kanten → Vertrag → PARTEI_VON-Kanten → verbundene Firma finden
   b. Firma-Name in Projekte-Excel nachschlagen (über "Kunde"-Spalte)
   c. **Match gefunden**: P-Nummer übernehmen (z.B. P0042)
   d. **Mehrere Projekte für selbe Firma**: Adress-/Standort-Matching zur Disambiguierung
   e. **Kein Match**: P9001, P9002, P9003, ... zuweisen
3. Neue Vertex-Properties:
   - `projekt_id`: "P0042"
   - `display_name`: "P0042 - PV-Anlage Dornum"
   - `standort_original`: Original-Adress-String behalten
   - Optional: `plz`, `ort`, `strasse`, `gemarkung` als separate Properties (für Filter)

## Phase 4: Firma/Person-Klassifizierung

75 Firma-Einträge mit rechtsform="Privatperson" → Property `ist_natuerliche_person: true` hinzufügen. Label bleibt "Firma" (sind formelle Vertragsparteien). Ausnahmen: GbR, e.K. etc. bleiben ohne Flag.

## Phase 5: App-Anpassungen

Datei: `app.py`
- Projekt-Dropdown: `display_name` statt roher Adresse anzeigen
- Predefined Queries: alte Vertragsart-Strings aktualisieren
- Node-Tooltips: neue Properties anzeigen (projekt_id, vertragsart_subtyp, ist_natuerliche_person)

---

## Ausführungsreihenfolge

1. Relevante Skripte aus PoC-Ordner nach syx5 kopieren (`normalize_contracts.py`, `generate_graph.py`)
2. `optimize_graph.py` erstellen mit allen Transformationen
3. Phase 1-4 implementieren
4. Skript ausführen im syx5-Ordner, Output prüfen
5. Phase 5: App-Anpassungen in `app.py`
6. Lokal testen im Streamlit
7. Git commit + push → Live-Deployment auf Streamlit Cloud + Azure

## Kritische Dateien (alle im syx5 Streamlit Deployment Repo)

- `optimize_graph.py` — **NEU**, Kernlogik
- `data/vertices.json` — Input/Output
- `data/edges.json` — Input/Output (Kanten bei Person-Splits)
- `data/contracts_metadata.json` — Input/Output
- `input/Projekte 18_03_2026 15-25-04.xlsx` — Master-Daten
- `input/Firmen 18_03_2026 15-23-59.xlsx` — Master-Daten
- `app.py` — Filter/Display-Anpassungen

## Verifizierung

1. `python optimize_graph.py` ausführen → Report prüfen
2. Backup-Dateien werden automatisch angelegt (`vertices_pre_optimize.json` etc.)
3. `streamlit run app.py --server.headless true` → alle Filter-Tabs durchklicken
4. Prüfen: Projekt-Dropdown zeigt P-Nummern, Vertragsart-Filter ist sauber, keine Multi-Personen mehr
