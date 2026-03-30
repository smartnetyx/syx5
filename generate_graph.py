import json, hashlib, re, os

base = os.path.dirname(os.path.abspath(__file__))

contracts = json.load(open(os.path.join(base, 'data', 'contracts_extracted_v3.json'), 'r', encoding='utf-8'))

# ── NORMALIZATION MAPS ──

NAME_NORMALIZE = {
    # v3 hat bereits alle Namen normalisiert — hier nur noch Restfälle
}

ROLLE_NORMALIZE = {
    "Grundeigentümer": "Eigentümer",
    "Grundstückseigentümer": "Eigentümer",
    "Gebäudeeigentümer": "Eigentümer",
    "Eigentuemer / Gestattungsgeber": "Eigentümer",
    "Eigentümer / Gestattungsgeber": "Eigentümer",
    "Gesellschafter 1": "Gesellschafter",
    "Gesellschafter 2": "Gesellschafter",
    "Gesellschafter 3": "Gesellschafter",
    "Gesellschafter 4": "Gesellschafter",
    "Gesellschafter 5": "Gesellschafter",
    "Nutzer / Gestattungsnehmer": "Nutzer",
    "Pächter/Nutzer": "Nutzer",
    "Verpächter/Eigentümer": "Eigentümer",
    "Erwerber/Käufer": "Käufer",
    "Veräußerer/Verkäufer": "Verkäufer",
    "Verkäufer und Nutzungsberechtigter": "Verkäufer",
    "Vermieterin": "Vermieter",
    "Kunde (Wärmeabnehmer)": "Abnehmer",
    "Unternehmen (Wärmelieferant)": "Wärmelieferant",
    "Direktvermarkter (e2m)": "Direktvermarkter",
}

STANDORT_NORMALIZE = {
    # Vrees mit PLZ -> Vrees
    "49757 Vrees": "Vrees",
    # Surwold Schreibvarianten
    "Surwold, Schleusenstrasse 82": "Surwold, Schleusenstraße 82",
    # Twist Varianten
    "Twist, Alt-Röhlertwist 76A": "Twist, Alt-Röhlertwist 76 A",
    # Sögel Varianten
    "Sögel": "Sögel",
    # Vrees Gewerbegebiet
    "Vrees, Gewerbegebiet Peheimer Straße": "Vrees, Gewerbegebiet, Peheimer Straße",
    "49757 Vrees, Gewerbegebiet, Peheimer Straße": "Vrees, Gewerbegebiet, Peheimer Straße",
    # Vrees Bischofsbr. Weg Varianten
    "Vrees, Bischofsbriicker Weg": "Vrees, Bischofsbr. Weg",
    "49757 Vrees, Bischofsbröcker Weg": "Vrees, Bischofsbr. Weg",
    # Windpark Varianten
    "Windpark Börger / Werlte, 49757 Vrees": "Windpark Börger",
    "Windpark Werlte, 49757 Vrees": "Windpark Werlte",
    "49757 Werlte": "Werlte",
}

def normalize_standort(standort):
    """Normalize location strings: strip PLZ prefix, unify spelling"""
    if not standort:
        return standort
    # Direct mapping
    if standort in STANDORT_NORMALIZE:
        return STANDORT_NORMALIZE[standort]
    # Strip leading PLZ (5 digits + space) if the rest is a known place
    m = re.match(r'^\d{5}\s+(.+)$', standort)
    if m:
        stripped = m.group(1)
        # Check if stripped version or standort is in normalize map
        if stripped in STANDORT_NORMALIZE:
            return STANDORT_NORMALIZE[stripped]
        # Keep stripped version if original had PLZ (but keep PLZ in detailed addresses)
        rest = stripped.split(',')
        if len(rest) == 1:
            return stripped  # Just "Vrees" instead of "49757 Vrees"
    return standort

def normalize_name(name):
    return NAME_NORMALIZE.get(name, name)

def normalize_rolle(rolle):
    return ROLLE_NORMALIZE.get(rolle, rolle)

# ── HELPERS ──

def make_id(prefix, name):
    clean = re.sub(r'[^a-z0-9_]', '_', name.lower().strip())
    clean = re.sub(r'_+', '_', clean).strip('_')
    if len(clean) > 60:
        clean = clean[:60]
    return f"{prefix}:{clean}"

def make_vertrag_id(contract):
    vid = contract.get('vertrags_id')
    if vid:
        return vid
    h = hashlib.md5(contract.get('dateiname', '').encode()).hexdigest()[:8]
    return f"vertrag:{h}"

def determine_entity_type(name, rechtsform=None):
    firma_indicators = ['gmbh', 'kg', 'gbr', 'ag', 'ohg', 'e.v.', 'solar', 'taler', 'wind',
                        'bioenergie', 'bioenenergie', 'biogas', 'stadtwerk', 'gemeinde', 'samtgemeinde',
                        'edeka', 'volksbank', 'sparkasse', 'immobilien', 'erbengemeinschaft',
                        'maschinenbau', 'holztechnik', 'gefluegelzucht', 'putenmast',
                        'hera', 'gts', 'bbv', 'snw', 'hrn', 'lhh',
                        'sonnentaler', 'marka', 'beeke', 'arenberg', 'quappen']
    name_lower = name.lower() if name else ''
    rf_lower = (rechtsform or '').lower()

    for ind in firma_indicators:
        if ind in name_lower or ind in rf_lower:
            return 'Firma'

    if rf_lower and any(x in rf_lower for x in ['gmbh', 'kg', 'gbr', 'ag', 'einzelunternehmen', 'e.v.']):
        return 'Firma'

    return 'Person'

def get_kategorie(dateipfad, vertragsart):
    if dateipfad and '/' in dateipfad:
        return dateipfad.split('/')[0]
    return vertragsart or 'Sonstige'

def extract_date_string(contract):
    for field in ['datum_abschluss', 'datum_beginn', 'datum_ende']:
        val = contract.get(field)
        if val:
            return val
    return ''

def extract_year_from_date(date_str):
    """Extract year from various date formats like '14.11.2017', '2024-12-31', '31.12.2041'"""
    if not date_str or not isinstance(date_str, str):
        return None
    # Try DD.MM.YYYY
    m = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', date_str)
    if m:
        return int(m.group(3))
    # Try YYYY-MM-DD
    m = re.search(r'(\d{4})-\d{2}-\d{2}', date_str)
    if m:
        return int(m.group(1))
    # Try standalone year
    m = re.search(r'(\d{4})', date_str)
    if m:
        y = int(m.group(1))
        if 1990 <= y <= 2100:
            return y
    return None

def extract_vertragswert(contract):
    vw = contract.get('vertragswert_jaehrlich')
    if vw:
        if isinstance(vw, (int, float)):
            return vw
        nums = re.findall(r'[\d.,]+', str(vw).replace('.', '').replace(',', '.'))
        if nums:
            try:
                return float(nums[0])
            except:
                pass
    return None

# ── BUILD GRAPH ──

vertices = []
edges = []
metadata = []
seen_vertices = set()
seen_edges = set()

def add_vertex(v):
    if v['id'] not in seen_vertices:
        seen_vertices.add(v['id'])
        vertices.append(v)

def add_edge(e):
    key = (e['src'], e['dst'], e['relationship'])
    if key not in seen_edges:
        seen_edges.add(key)
        edges.append(e)

for c in contracts:
    vid = make_vertrag_id(c)
    dateiname = c.get('dateiname', '')
    dateipfad = c.get('dateipfad', '')
    vertragsart = c.get('vertragsart', 'Unbekannt')
    kategorie = get_kategorie(dateipfad, vertragsart)

    # Vertrag vertex
    add_vertex({
        'id': vid,
        'label': 'Vertrag',
        'name': dateiname.replace('.pdf', ''),
        'vertragsart': vertragsart,
        'status': c.get('status', 'aktiv'),
        'ist_nachtrag': c.get('ist_nachtrag', False),
        'kategorie': kategorie,
        'seitenzahl': c.get('seitenzahl'),
        'standorte': c.get('standort', '') or '',
        'vertragswert': extract_vertragswert(c),
        'datum_gefunden': extract_date_string(c),
        'vertragstitel': c.get('vertragstitel', ''),
        'datum_ende': c.get('datum_ende'),
        'datum_ende_jahr': extract_year_from_date(c.get('datum_ende')),
    })

    # Dokument vertex
    dok_id = make_id('dokument', dateiname.replace('.pdf', ''))
    add_vertex({
        'id': dok_id,
        'label': 'Dokument',
        'name': dateiname,
        'dateipfad': dateipfad,
        'seitenzahl': c.get('seitenzahl'),
    })
    add_edge({'src': vid, 'dst': dok_id, 'relationship': 'HAT_DOKUMENT', 'dokument_typ': 'Original'})

    # Extract parties (with normalization)
    parteien_meta = []
    for pkey in ['vertragspartner_1', 'vertragspartner_2', 'vertragspartner_3',
                 'vertragspartner_4', 'vertragspartner_5']:
        p = c.get(pkey)
        if not p or not isinstance(p, dict):
            continue
        pname = normalize_name(p.get('name', ''))
        if not pname:
            continue

        rechtsform = p.get('rechtsform', '')
        entity_type = determine_entity_type(pname, rechtsform)
        prefix = 'firma' if entity_type == 'Firma' else 'person'
        pid = make_id(prefix, pname)

        add_vertex({
            'id': pid,
            'label': entity_type,
            'name': pname,
            'rechtsform': rechtsform,
            'adresse': p.get('adresse', ''),
        })

        rolle = normalize_rolle(p.get('rolle', 'Vertragspartei'))
        add_edge({'src': pid, 'dst': vid, 'relationship': 'PARTEI_VON', 'rolle': rolle})
        parteien_meta.append({'name': pname, 'typ': entity_type, 'rolle': rolle})

    # Standort vertex if available (normalized)
    standort = normalize_standort(c.get('standort', ''))
    if standort and len(standort) > 3:
        sid = make_id('standort', standort)
        add_vertex({
            'id': sid,
            'label': 'Projekt',
            'name': standort,
        })
        add_edge({'src': vid, 'dst': sid, 'relationship': 'HAT_STANDORT'})

    # MaLo nodes
    malo_ids = c.get('malo_ids') or c.get('malo_id')
    if malo_ids:
        if isinstance(malo_ids, str):
            malo_ids = [malo_ids]
        for malo_id in malo_ids:
            if malo_id and len(malo_id) > 5:
                mid = f"malo:{malo_id}"
                add_vertex({
                    'id': mid,
                    'label': 'MaLo',
                    'name': malo_id,
                    'malo_leistung_mw': c.get('malo_leistung_mw'),
                })
                add_edge({'src': vid, 'dst': mid, 'relationship': 'HAT_MALO'})

    # Anlage nodes (from PV data or vertragsgegenstand)
    pv_kwp = c.get('pv_leistung_kwp')
    if pv_kwp and pv_kwp != 'null':
        anlage_name = f"PV-Anlage {pv_kwp} kWp"
        standort_kurz = (standort or dateiname)[:30]
        aid = make_id('anlage', f"pv_{standort_kurz}_{pv_kwp}")
        add_vertex({
            'id': aid,
            'label': 'Anlage',
            'name': anlage_name,
            'typ': 'PV',
            'leistung_kwp': pv_kwp,
            'standort': standort,
        })
        add_edge({'src': vid, 'dst': aid, 'relationship': 'BETRIFFT_ANLAGE'})

    # Nachtrag reference
    nref = c.get('nachtrag_referenz')
    if nref and c.get('ist_nachtrag'):
        for c2 in contracts:
            if c2.get('dateiname') == nref or c2.get('vertrags_id') == nref:
                ref_vid = make_vertrag_id(c2)
                add_edge({'src': vid, 'dst': ref_vid, 'relationship': 'NACHTRAG_ZU'})
                break

    # Metadata entry (extended)
    metadata.append({
        'vertrags_id': vid,
        'dateiname': dateiname,
        'dateipfad': dateipfad,
        'kategorie': kategorie,
        'vertragsart': vertragsart,
        'ist_nachtrag': c.get('ist_nachtrag', False),
        'seitenzahl': c.get('seitenzahl'),
        'parteien': parteien_meta,
        'datum': {
            'abschluss': c.get('datum_abschluss'),
            'beginn': c.get('datum_beginn'),
            'ende': c.get('datum_ende'),
            'ende_jahr': extract_year_from_date(c.get('datum_ende')),
        },
        'finanzen': {
            'vertragswert_jaehrlich': c.get('vertragswert_jaehrlich'),
            'verguetung': c.get('verguetung'),
            'verguetungsbasis': c.get('verguetungsbasis'),
            'waehrung': c.get('waehrung'),
            'zahlungsrhythmus': c.get('zahlungsrhythmus'),
            'arbeitspreis_ct_kwh': c.get('arbeitspreis_ct_kwh'),
            'grundpreis_eur_monat': c.get('grundpreis_eur_monat'),
            'strompreis_ct_kwh': c.get('strompreis_ct_kwh'),
            'kaltmiete_eur': c.get('kaltmiete_eur'),
            'nebenkosten_eur': c.get('nebenkosten_eur'),
            'gesamtmiete_eur': c.get('gesamtmiete_eur'),
        },
        'energie': {
            'pv_leistung_kwp': c.get('pv_leistung_kwp'),
            'anschlussleistung_kw': c.get('anschlussleistung_kw'),
            'malo_ids': c.get('malo_ids') or (c.get('malo_id') and [c.get('malo_id')]) or [],
            'malo_leistung_mw': c.get('malo_leistung_mw'),
            'eeg_refs': c.get('eeg_refs') or [],
        },
        'standorte': [c.get('standort')] if c.get('standort') else [],
        'laufzeit': c.get('laufzeit_text'),
        'auto_verlaengerung': c.get('auto_verlaengerung'),
        'kuendigungsfrist': c.get('kuendigungsfrist'),
        'kuendigung_form': c.get('kuendigung_form'),
        'kuendigung_ausserordentlich': c.get('kuendigung_ausserordentlich'),
        'gerichtsstand': c.get('gerichtsstand'),
        'vertragstitel': c.get('vertragstitel'),
        'vertragsgegenstand': c.get('vertragsgegenstand'),
    })

# Write output files
for fname, data in [('vertices.json', vertices), ('edges.json', edges), ('contracts_metadata.json', metadata)]:
    for subdir in ['data', os.path.join('app', 'data')]:
        outpath = os.path.join(base, subdir, fname)
        with open(outpath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

# Stats
label_counts = {}
for v in vertices:
    l = v['label']
    label_counts[l] = label_counts.get(l, 0) + 1

print(f"Vertices: {len(vertices)}")
for label, count in sorted(label_counts.items(), key=lambda x: -x[1]):
    print(f"  {label}: {count}")
print(f"Edges: {len(edges)}")
print(f"Metadata: {len(metadata)}")

# Stats on roles
rolle_counts = {}
for e in edges:
    if e['relationship'] == 'PARTEI_VON':
        r = e.get('rolle', '?')
        rolle_counts[r] = rolle_counts.get(r, 0) + 1
print(f"\nRollen (normalisiert):")
for r, cnt in sorted(rolle_counts.items(), key=lambda x: -x[1]):
    print(f"  {r}: {cnt}")
