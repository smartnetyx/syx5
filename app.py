"""
TERP Contract Management Graph Visualization
Streamlit Application for Contract Network Analysis
"""

import streamlit as st
import pandas as pd
import json
import plotly.graph_objects as go
import plotly.express as px
import networkx as nx
from pathlib import Path
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import warnings
warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(
    page_title="TERP Vertragsmanagement",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for larger tab labels
st.markdown("""<style>
button[data-baseweb="tab"] > div > p { font-size: 18px !important; }
</style>""", unsafe_allow_html=True)

# Color scheme for node types
NODE_COLORS = {
    'Vertrag': '#4A90D9',      # Blue
    'Firma': '#27AE60',        # Green
    'Person': '#E67E22',       # Orange
    'Anlage': '#F1C40F',       # Yellow
    'Projekt': '#E74C3C',      # Red
    'Dokument': '#95A5A6',     # Gray
    'MaLo': '#9B59B6'          # Purple
}

# Load data
@st.cache_resource
def load_data():
    """Load vertices, edges, and contracts metadata"""
    data_dir = Path(__file__).parent / 'data'

    with open(data_dir / 'vertices.json', 'r', encoding='utf-8') as f:
        vertices = json.load(f)

    with open(data_dir / 'edges.json', 'r', encoding='utf-8') as f:
        edges = json.load(f)

    with open(data_dir / 'contracts_metadata.json', 'r', encoding='utf-8') as f:
        contracts = json.load(f)

    return vertices, edges, contracts

@st.cache_data
def build_network_graph(vertices, edges):
    """Build NetworkX graph from vertices and edges"""
    G = nx.DiGraph()

    # Add nodes with attributes
    for vertex in vertices:
        G.add_node(
            vertex['id'],
            label=vertex.get('label', 'Unbekannt'),
            name=vertex.get('name', vertex['id']),
            **{k: v for k, v in vertex.items() if k not in ['id', 'label', 'name']}
        )

    # Add edges with attributes
    for edge in edges:
        G.add_edge(
            edge['src'],
            edge['dst'],
            relationship=edge.get('relationship', 'UNKNOWN'),
            **{k: v for k, v in edge.items() if k not in ['src', 'dst', 'relationship']}
        )

    return G

@st.cache_data
def get_firmen_list(vertices):
    """Get unique list of Firmen"""
    firmen = [v for v in vertices if v.get('label') == 'Firma']
    return sorted([f.get('name', f['id']) for f in firmen])

@st.cache_data
def get_personen_list(vertices):
    """Get unique list of Personen"""
    personen = [v for v in vertices if v.get('label') == 'Person']
    return sorted([p.get('name', p['id']) for p in personen])

@st.cache_data
def get_vertragsart_list(contracts):
    """Get unique list of Vertragsart"""
    return sorted(set(c.get('vertragsart', 'Unbekannt') for c in contracts if c.get('vertragsart')))

@st.cache_data
def get_projekte_list(vertices):
    """Get unique list of Projekte"""
    projekte = [v for v in vertices if v.get('label') == 'Projekt']
    return sorted([p.get('name', p['id']) for p in projekte])

def filter_contracts(contracts, vertices, edges, firmen_filter, projekt_filter, vertragsart_filter, personen_filter):
    """Filter contracts based on the 4 global filters"""
    filtered = list(contracts)

    if vertragsart_filter:
        filtered = [c for c in filtered if c.get('vertragsart') in vertragsart_filter]

    if firmen_filter:
        filtered = [c for c in filtered
                   if any(p.get('name') in firmen_filter for p in c.get('parteien', []))]

    if personen_filter:
        filtered = [c for c in filtered
                   if any(p.get('name') in personen_filter for p in c.get('parteien', []))]

    if projekt_filter:
        projekt_nodes = [v['id'] for v in vertices
                        if v.get('label') == 'Projekt' and v.get('name') in projekt_filter]
        contract_ids_with_projekt = set()
        for edge in edges:
            if edge['dst'] in projekt_nodes and 'vertrag:' in edge['src']:
                contract_ids_with_projekt.add(edge['src'])
            if edge['src'] in projekt_nodes and 'vertrag:' in edge['dst']:
                contract_ids_with_projekt.add(edge['dst'])
        filtered = [c for c in filtered if c.get('vertrags_id') in contract_ids_with_projekt]

    return filtered

def filter_graph_by_criteria(G, vertices, edges, firmen_filter, vertragsart_filter, projekt_filter, personen_filter=None):
    """Filter graph based on selected criteria"""
    contracts_filtered = [v['id'] for v in vertices
                         if v.get('label') == 'Vertrag']

    if vertragsart_filter:
        contracts_filtered = [c for c in contracts_filtered
                            if any(v['id'] == c and v.get('vertragsart') in vertragsart_filter
                                   for v in vertices)]

    if firmen_filter:
        firma_nodes = [v['id'] for v in vertices
                      if v.get('label') == 'Firma' and v.get('name') in firmen_filter]
        contracts_with_firmen = set()
        for contract in contracts_filtered:
            for edge in edges:
                if ((edge['src'] in firma_nodes and edge['dst'] == contract) or
                    (edge['src'] == contract and edge['dst'] in firma_nodes)):
                    contracts_with_firmen.add(contract)
        contracts_filtered = list(contracts_with_firmen)

    if personen_filter:
        person_nodes = [v['id'] for v in vertices
                       if v.get('label') == 'Person' and v.get('name') in personen_filter]
        contracts_with_personen = set()
        for contract in contracts_filtered:
            for edge in edges:
                if ((edge['src'] in person_nodes and edge['dst'] == contract) or
                    (edge['src'] == contract and edge['dst'] in person_nodes)):
                    contracts_with_personen.add(contract)
        contracts_filtered = list(contracts_with_personen)

    if projekt_filter:
        projekt_nodes = [v['id'] for v in vertices
                        if v.get('label') == 'Projekt' and v.get('name') in projekt_filter]
        contracts_with_projekten = set()
        for contract in contracts_filtered:
            for edge in edges:
                if ((edge['src'] == contract and edge['dst'] in projekt_nodes) or
                    (edge['src'] in projekt_nodes and edge['dst'] == contract)):
                    contracts_with_projekten.add(contract)
        contracts_filtered = list(contracts_with_projekten)

    relevant_nodes = set(contracts_filtered)
    for contract in contracts_filtered:
        for edge in edges:
            if edge['src'] == contract:
                relevant_nodes.add(edge['dst'])
            if edge['dst'] == contract:
                relevant_nodes.add(edge['src'])

    subgraph_vertices = [v for v in vertices if v['id'] in relevant_nodes]
    subgraph_edges = [e for e in edges
                     if e['src'] in relevant_nodes and e['dst'] in relevant_nodes]

    subG = nx.DiGraph()
    for vertex in subgraph_vertices:
        subG.add_node(
            vertex['id'],
            **{k: v for k, v in vertex.items() if k != 'id'}
        )

    for edge in subgraph_edges:
        subG.add_edge(edge['src'], edge['dst'],
                     **{k: v for k, v in edge.items() if k not in ['src', 'dst']})

    return subG

def create_pyvis_graph(G):
    """Create interactive pyvis visualization"""
    import pyvis.network as net

    g = net.Network(
        height='900px',
        width='100%',
        directed=True,
        notebook=False,
        layout=None
    )

    for node in G.nodes():
        node_data = G.nodes[node]
        label = node_data.get('label', 'Unbekannt')
        color = NODE_COLORS.get(label, '#95A5A6')
        name = node_data.get('name', node)[:50]

        tooltip_lines = [f"<b>{node_data.get('name', node)}</b>"]
        tooltip_lines.append(f"Typ: {label}")
        for k, v in node_data.items():
            if k not in ('name', 'label') and v:
                tooltip_lines.append(f"{k}: {v}")
        title = "<br>".join(tooltip_lines)

        g.add_node(
            node,
            label=name,
            title=title,
            color=color,
            size=25 if label == 'Vertrag' else 20,
            font={'size': 12}
        )

    for src, dst in G.edges():
        edge_data = G.edges[src, dst]
        relationship = edge_data.get('relationship', 'UNKNOWN')

        tooltip_lines = [f"<b>{relationship}</b>"]
        for k, v in edge_data.items():
            if k != 'relationship' and v:
                tooltip_lines.append(f"{k}: {v}")
        edge_title = "<br>".join(tooltip_lines)

        g.add_edge(src, dst, label=relationship,
                  title=edge_title, font={'size': 10})

    g.toggle_physics(True)
    g.show_buttons(filter_=['physics'])

    return g

def create_contracts_table(contracts_list):
    """Create comprehensive contracts table"""
    df = pd.DataFrame()
    df['Vertrags-ID'] = [c.get('vertrags_id', '') for c in contracts_list]
    df['Dateiname'] = [c.get('dateiname', '') for c in contracts_list]
    df['Vertragsart'] = [c.get('vertragsart', '') for c in contracts_list]
    df['Seiten'] = [c.get('seitenzahl', '') for c in contracts_list]
    df['Status'] = ['Nachtrag' if c.get('ist_nachtrag') else 'Original' for c in contracts_list]
    df['Parteien'] = [', '.join([p.get('name', '') for p in c.get('parteien', [])])
                     for c in contracts_list]
    df['Standorte'] = [', '.join(c.get('standorte', [])) if isinstance(c.get('standorte'), list) else c.get('standorte', '') for c in contracts_list]

    return df

def get_contracts_by_firma(contracts, selected_firmen):
    """Get contracts grouped by company"""
    result = {}
    for firma in selected_firmen:
        firma_contracts = []
        for c in contracts:
            for party in c.get('parteien', []):
                if party.get('name') == firma:
                    firma_contracts.append(c)
                    break
        result[firma] = firma_contracts
    return result

def get_contracts_by_vertragsart(contracts, selected_arten):
    """Get contracts grouped by type"""
    result = {}
    for art in selected_arten:
        result[art] = [c for c in contracts if c.get('vertragsart') == art]
    return result

def get_contracts_by_projekt(vertices, edges, contracts):
    """Get contracts grouped by project"""
    result = {}
    projekte = [v for v in vertices if v.get('label') == 'Projekt']

    for projekt in projekte:
        projekt_name = projekt.get('name')
        projekt_contracts = []

        for edge in edges:
            if edge['dst'] == projekt['id'] and 'vertrag:' in edge['src']:
                vertrag_id = edge['src']
                matching = [c for c in contracts if c.get('vertrags_id') == vertrag_id]
                if matching:
                    projekt_contracts.append(matching[0])

        if projekt_contracts:
            result[projekt_name] = projekt_contracts

    return result

def get_parteien_ubersicht(contracts):
    """Get overview of all parties and their contracts"""
    parteien_map = {}

    for c in contracts:
        for party in c.get('parteien', []):
            party_name = party.get('name')
            if party_name not in parteien_map:
                parteien_map[party_name] = {
                    'Typ': party.get('typ'),
                    'Anzahl Verträge': 0,
                    'Verträge': []
                }
            parteien_map[party_name]['Anzahl Verträge'] += 1
            parteien_map[party_name]['Verträge'].append(c.get('dateiname', ''))

    return parteien_map

def get_finanz_ubersicht(contracts):
    """Get financial overview"""
    df = pd.DataFrame()
    df['Vertrags-ID'] = [c.get('vertrags_id', '') for c in contracts]
    df['Dateiname'] = [c.get('dateiname', '') for c in contracts]
    df['Vertragsart'] = [c.get('vertragsart', '') for c in contracts]

    finanzen_list = [c.get('finanzen', {}) for c in contracts]
    non_empty = [f for f in finanzen_list if f]
    if non_empty:
        for key in set().union(*[set(f.keys()) for f in non_empty]):
            df[key] = [f.get(key, '') for f in finanzen_list]

    return df

def get_waermeabnehmer_ubersicht(edges, vertices):
    """Get overview of heat supply customers"""
    waermeabnehmer = []

    for edge in edges:
        if edge.get('relationship') == 'LIEFERT_WAERME':
            src_vertex = next((v for v in vertices if v['id'] == edge['src']), None)
            dst_vertex = next((v for v in vertices if v['id'] == edge['dst']), None)

            if src_vertex and dst_vertex:
                waermeabnehmer.append({
                    'Lieferant': src_vertex.get('name', src_vertex['id']),
                    'Abnehmer': dst_vertex.get('name', dst_vertex['id']),
                    'Beziehung': edge.get('beziehung', 'Wärmelieferung')
                })

    return waermeabnehmer

def get_beziehungs_matrix(vertices, edges):
    """Get relationship matrix between companies"""
    firmen = {v['id']: v.get('name') for v in vertices if v.get('label') == 'Firma'}

    matrix_data = []
    for edge in edges:
        src = edge['src']
        dst = edge['dst']

        if src in firmen and dst in firmen:
            matrix_data.append({
                'Firma 1': firmen[src],
                'Firma 2': firmen[dst],
                'Beziehung': edge.get('relationship', 'UNKNOWN'),
                'Details': edge.get('rolle', '')
            })

    return pd.DataFrame(matrix_data) if matrix_data else pd.DataFrame()

def get_gesellschafter_ubersicht(vertices, edges):
    """Get overview of shareholders and participations"""
    gesellschafter = []

    for edge in edges:
        if edge.get('relationship') == 'GESELLSCHAFTER_VON':
            src_vertex = next((v for v in vertices if v['id'] == edge['src']), None)
            dst_vertex = next((v for v in vertices if v['id'] == edge['dst']), None)

            if src_vertex and dst_vertex:
                gesellschafter.append({
                    'Gesellschafter': src_vertex.get('name', src_vertex['id']),
                    'Gesellschaft': dst_vertex.get('name', dst_vertex['id']),
                    'Typ Gesellschafter': src_vertex.get('label', ''),
                    'Typ Gesellschaft': dst_vertex.get('label', '')
                })

    return pd.DataFrame(gesellschafter) if gesellschafter else pd.DataFrame()

def search_contracts(contracts, search_term):
    """Search contracts by free text"""
    search_term_lower = search_term.lower()
    results = []

    for c in contracts:
        searchable_fields = [
            c.get('dateiname', ''),
            c.get('vertragsart', ''),
            ', '.join([p.get('name', '') for p in c.get('parteien', [])]),
            ', '.join(c.get('standorte', [])) if isinstance(c.get('standorte'), list) else c.get('standorte', ''),
            c.get('text_preview', '')
        ]

        if any(search_term_lower in field.lower() for field in searchable_fields):
            results.append(c)

    return results

def get_ego_graph(G, contract_id, depth=1):
    """Get ego graph for a specific contract"""
    if contract_id not in G:
        return nx.DiGraph()

    ego_nodes = set([contract_id])

    for _ in range(depth):
        new_nodes = set()
        for node in ego_nodes:
            new_nodes.update(G.predecessors(node))
            new_nodes.update(G.successors(node))
        ego_nodes.update(new_nodes)

    return G.subgraph(ego_nodes)

def export_to_csv(df):
    """Export dataframe to CSV"""
    return df.to_csv(index=False).encode('utf-8')

def export_to_excel(dfs_dict):
    """Export multiple dataframes to Excel"""
    wb = Workbook()
    wb.remove(wb.active)

    for sheet_name, df in dfs_dict.items():
        ws = wb.create_sheet(sheet_name[:31])

        for col_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.value = col_name
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row_idx, row in enumerate(df.values, 2):
            for col_idx, value in enumerate(row, 1):
                ws.cell(row=row_idx, column=col_idx).value = value

        for column in ws.columns:
            max_length = 20
            column_letter = column[0].column_letter
            ws.column_dimensions[column_letter].width = min(max_length, 50)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# Main app
def main():
    """Main Streamlit application"""

    # Compact header
    st.markdown("""
    <h2 style='text-align: center; color: #2C3E50; margin: 0; padding: 5px 0;'>
        🔗 TERP Vertragsmanagement
    </h2>
    """, unsafe_allow_html=True)

    # Load data
    vertices, edges, contracts = load_data()
    G = build_network_graph(vertices, edges)

    # ── GLOBAL CASCADING FILTERS (above tabs) ──
    # Each filter's options are computed from contracts filtered by the OTHER filters.
    # This means selecting a Firma narrows down available Personen/Vertragsarten/Projekte.

    # Helper: get available options from a set of contracts
    def get_options_from_contracts(c_list):
        firmen = sorted(set(
            p.get('name') for c in c_list for p in c.get('parteien', []) if p.get('typ') == 'Firma'
        ))
        personen = sorted(set(
            p.get('name') for c in c_list for p in c.get('parteien', []) if p.get('typ') == 'Person'
        ))
        arten = sorted(set(c.get('vertragsart') for c in c_list if c.get('vertragsart')))
        # Projects need edges
        c_ids = set(c.get('vertrags_id') for c in c_list)
        projekt_ids = set()
        for edge in edges:
            if edge['src'] in c_ids and edge.get('relationship') == 'GEHOERT_ZU':
                projekt_ids.add(edge['dst'])
        projekte = sorted(
            v.get('name') for v in vertices
            if v.get('label') == 'Projekt' and v['id'] in projekt_ids
        )
        return firmen, projekte, arten, personen

    # Step 1: Firma filter — always shows all options (no dependency)
    all_firmen = get_firmen_list(vertices)
    filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)
    with filter_col1:
        selected_firmen = st.multiselect("Firma", all_firmen, key="global_firmen")

    # Step 2: Apply Firma filter to narrow options for remaining filters
    if selected_firmen:
        contracts_after_firma = filter_contracts(contracts, vertices, edges, selected_firmen, [], [], [])
    else:
        contracts_after_firma = contracts

    # Step 3: Projekt filter — options narrowed by Firma
    _, available_projekte, _, _ = get_options_from_contracts(contracts_after_firma)
    with filter_col2:
        selected_projekte = st.multiselect("Projekt", available_projekte, key="global_projekte")

    # Step 4: Apply Firma + Projekt to narrow Vertragsart options
    if selected_firmen or selected_projekte:
        contracts_after_firma_projekt = filter_contracts(contracts, vertices, edges, selected_firmen, selected_projekte, [], [])
    else:
        contracts_after_firma_projekt = contracts

    _, _, available_arten, _ = get_options_from_contracts(contracts_after_firma_projekt)
    with filter_col3:
        selected_arten = st.multiselect("Vertragsart", available_arten, key="global_arten")

    # Step 5: Apply Firma + Projekt + Vertragsart to narrow Person options
    if selected_firmen or selected_projekte or selected_arten:
        contracts_after_three = filter_contracts(contracts, vertices, edges, selected_firmen, selected_projekte, selected_arten, [])
    else:
        contracts_after_three = contracts

    _, _, _, available_personen = get_options_from_contracts(contracts_after_three)
    with filter_col4:
        selected_personen = st.multiselect("Person", available_personen, key="global_personen")

    # Final filtered data
    any_filter_active = selected_firmen or selected_arten or selected_projekte or selected_personen

    if any_filter_active:
        filtered_contracts = filter_contracts(
            contracts, vertices, edges,
            selected_firmen, selected_projekte, selected_arten, selected_personen
        )
        filtered_graph = filter_graph_by_criteria(
            G, vertices, edges,
            selected_firmen, selected_arten, selected_projekte, selected_personen
        )
    else:
        filtered_contracts = contracts
        filtered_graph = G

    # Create tabs
    tab1, tab2, tab3, tab4 = st.tabs([
        "🔗 Graph-Visualisierung",
        "📋 Tabellen-Generator",
        "📈 Dashboard-Übersicht",
        "🔍 Detail-Ansicht"
    ])

    # ── TAB 1: GRAPH VISUALIZATION ──
    with tab1:
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Knoten", len(filtered_graph.nodes()))
        with col2:
            st.metric("Kanten", len(filtered_graph.edges()))
        with col3:
            st.metric("Verträge", len([v for v in filtered_graph.nodes()
                                       if any(n['id'] == v and n.get('label') == 'Vertrag'
                                             for n in vertices)]))
        with col4:
            st.metric("Firmen", len([v for v in filtered_graph.nodes()
                                    if any(n['id'] == v and n.get('label') == 'Firma'
                                         for n in vertices)]))

        # Color legend
        legend_html = "<div style='display:flex; gap:15px; flex-wrap:wrap; margin:10px 0;'>"
        for lbl, color in NODE_COLORS.items():
            legend_html += f"<span style='display:flex;align-items:center;gap:5px;'><span style='width:14px;height:14px;border-radius:50%;background:{color};display:inline-block;'></span>{lbl}</span>"
        legend_html += "</div>"
        st.markdown(legend_html, unsafe_allow_html=True)

        try:
            net_graph = create_pyvis_graph(filtered_graph)
            html_string = net_graph.generate_html()
            st.components.v1.html(html_string, height=900)
        except Exception as e:
            st.error(f"Fehler bei Graphvisualisierung: {e}")

    # ── TAB 2: TABLES GENERATOR ──
    with tab2:
        col1, col2 = st.columns([2, 1])

        with col1:
            table_type = st.selectbox(
                "Tabellen-Typ",
                [
                    "Alle Verträge",
                    "Verträge pro Firma",
                    "Verträge pro Vertragsart",
                    "Verträge pro Projekt",
                    "Parteien-Übersicht",
                    "Finanz-Übersicht",
                    "Gesellschafter-Übersicht",
                    "Wärmeabnehmer-Übersicht",
                    "Beziehungs-Matrix",
                    "Custom Query"
                ]
            )

        with col2:
            st.markdown("#### Export")

        if table_type == "Alle Verträge":
            df = create_contracts_table(filtered_contracts)
            st.dataframe(df, use_container_width=True, height=900)

            col1, col2 = st.columns(2)
            with col1:
                csv_data = export_to_csv(df)
                st.download_button("📥 CSV herunterladen", csv_data, "vertraege.csv", "text/csv")
            with col2:
                excel_data = export_to_excel({"Verträge": df})
                st.download_button("📥 Excel herunterladen", excel_data, "vertraege.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        elif table_type == "Verträge pro Firma":
            available_firmen = sorted(set(
                p.get('name') for c in filtered_contracts for p in c.get('parteien', []) if p.get('typ') == 'Firma'
            ))
            selected_firmen_tab = st.multiselect("Firmen wählen", available_firmen, key="table_firmen_2")

            if selected_firmen_tab:
                contracts_by_firma = get_contracts_by_firma(filtered_contracts, selected_firmen_tab)

                for firma, firma_contracts in contracts_by_firma.items():
                    with st.expander(f"{firma} ({len(firma_contracts)} Verträge)"):
                        df = pd.DataFrame({
                            'Dateiname': [c.get('dateiname') for c in firma_contracts],
                            'Vertragsart': [c.get('vertragsart') for c in firma_contracts],
                            'Seiten': [c.get('seitenzahl') for c in firma_contracts],
                            'Status': ['Nachtrag' if c.get('ist_nachtrag') else 'Original' for c in firma_contracts]
                        })
                        st.dataframe(df, use_container_width=True, height=900)

        elif table_type == "Verträge pro Vertragsart":
            available_arten = sorted(set(c.get('vertragsart') for c in filtered_contracts if c.get('vertragsart')))
            selected_arten_tab = st.multiselect("Vertragsarten wählen", available_arten, key="table_arten_2")

            if selected_arten_tab:
                contracts_by_art = get_contracts_by_vertragsart(filtered_contracts, selected_arten_tab)

                for art, art_contracts in contracts_by_art.items():
                    with st.expander(f"{art} ({len(art_contracts)} Verträge)"):
                        df = pd.DataFrame({
                            'Dateiname': [c.get('dateiname') for c in art_contracts],
                            'Seiten': [c.get('seitenzahl') for c in art_contracts],
                            'Parteien': [', '.join([p.get('name') for p in c.get('parteien', [])]) for c in art_contracts]
                        })
                        st.dataframe(df, use_container_width=True, height=900)

        elif table_type == "Verträge pro Projekt":
            contracts_by_projekt = get_contracts_by_projekt(vertices, edges, filtered_contracts)

            if contracts_by_projekt:
                for projekt, projekt_contracts in contracts_by_projekt.items():
                    with st.expander(f"{projekt} ({len(projekt_contracts)} Verträge)"):
                        df = pd.DataFrame({
                            'Dateiname': [c.get('dateiname') for c in projekt_contracts],
                            'Vertragsart': [c.get('vertragsart') for c in projekt_contracts],
                            'Seiten': [c.get('seitenzahl') for c in projekt_contracts],
                            'Parteien': [', '.join([p.get('name') for p in c.get('parteien', [])]) for c in projekt_contracts]
                        })
                        st.dataframe(df, use_container_width=True, height=900)
            else:
                st.info("Keine Projekte mit Verträgen gefunden.")

        elif table_type == "Parteien-Übersicht":
            parteien_ubersicht = get_parteien_ubersicht(filtered_contracts)

            df = pd.DataFrame([
                {
                    'Partei': name,
                    'Typ': info.get('Typ'),
                    'Anzahl Verträge': info.get('Anzahl Verträge')
                }
                for name, info in parteien_ubersicht.items()
            ]).sort_values('Anzahl Verträge', ascending=False)

            st.dataframe(df, use_container_width=True, height=900)

            col1, col2 = st.columns(2)
            with col1:
                csv_data = export_to_csv(df)
                st.download_button("📥 CSV herunterladen", csv_data, "parteien.csv", "text/csv")
            with col2:
                excel_data = export_to_excel({"Parteien": df})
                st.download_button("📥 Excel herunterladen", excel_data, "parteien.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        elif table_type == "Finanz-Übersicht":
            df = get_finanz_ubersicht(filtered_contracts)
            st.dataframe(df, use_container_width=True, height=900)

            col1, col2 = st.columns(2)
            with col1:
                csv_data = export_to_csv(df)
                st.download_button("📥 CSV herunterladen", csv_data, "finanzen.csv", "text/csv")
            with col2:
                excel_data = export_to_excel({"Finanzen": df})
                st.download_button("📥 Excel herunterladen", excel_data, "finanzen.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        elif table_type == "Gesellschafter-Übersicht":
            df = get_gesellschafter_ubersicht(vertices, edges)

            if not df.empty:
                st.dataframe(df, use_container_width=True, height=900)

                col1, col2 = st.columns(2)
                with col1:
                    csv_data = export_to_csv(df)
                    st.download_button("📥 CSV herunterladen", csv_data, "gesellschafter.csv", "text/csv")
                with col2:
                    excel_data = export_to_excel({"Gesellschafter": df})
                    st.download_button("📥 Excel herunterladen", excel_data, "gesellschafter.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("Keine Gesellschafter-Beziehungen gefunden.")

        elif table_type == "Wärmeabnehmer-Übersicht":
            waermeabnehmer = get_waermeabnehmer_ubersicht(edges, vertices)

            if waermeabnehmer:
                df = pd.DataFrame(waermeabnehmer)
                st.dataframe(df, use_container_width=True, height=900)

                col1, col2 = st.columns(2)
                with col1:
                    csv_data = export_to_csv(df)
                    st.download_button("📥 CSV herunterladen", csv_data, "waermeabnehmer.csv", "text/csv")
                with col2:
                    excel_data = export_to_excel({"Wärmeabnehmer": df})
                    st.download_button("📥 Excel herunterladen", excel_data, "waermeabnehmer.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("Keine Wärmelieferungsbeziehungen gefunden.")

        elif table_type == "Beziehungs-Matrix":
            df = get_beziehungs_matrix(vertices, edges)

            if not df.empty:
                st.dataframe(df, use_container_width=True, height=900)

                col1, col2 = st.columns(2)
                with col1:
                    csv_data = export_to_csv(df)
                    st.download_button("📥 CSV herunterladen", csv_data, "beziehungen.csv", "text/csv")
                with col2:
                    excel_data = export_to_excel({"Beziehungen": df})
                    st.download_button("📥 Excel herunterladen", excel_data, "beziehungen.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("Keine Unternehmensbeziehungen gefunden.")

        elif table_type == "Custom Query":
            search_term = st.text_input("Suchbegriff eingeben", "")

            if search_term:
                results = search_contracts(filtered_contracts, search_term)

                if results:
                    st.success(f"{len(results)} Verträge gefunden")
                    df = pd.DataFrame({
                        'Dateiname': [c.get('dateiname') for c in results],
                        'Vertragsart': [c.get('vertragsart') for c in results],
                        'Parteien': [', '.join([p.get('name') for p in c.get('parteien', [])]) for c in results]
                    })
                    st.dataframe(df, use_container_width=True, height=900)
                else:
                    st.info("Keine Verträge gefunden.")

    # ── TAB 3: DASHBOARD ──
    with tab3:
        # KPIs based on filtered data
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric("📋 Verträge", len(filtered_contracts))

        with col2:
            firmen_in_filtered = set(
                p.get('name') for c in filtered_contracts for p in c.get('parteien', []) if p.get('typ') == 'Firma'
            )
            st.metric("🏢 Firmen", len(firmen_in_filtered))

        with col3:
            # Count projects linked to filtered contracts
            filtered_ids = set(c.get('vertrags_id') for c in filtered_contracts)
            projekt_set = set()
            for edge in edges:
                if edge['src'] in filtered_ids and edge.get('relationship') == 'GEHOERT_ZU':
                    projekt_set.add(edge['dst'])
            st.metric("🔨 Projekte", len(projekt_set))

        with col4:
            personen_in_filtered = set(
                p.get('name') for c in filtered_contracts for p in c.get('parteien', []) if p.get('typ') == 'Person'
            )
            st.metric("👤 Personen", len(personen_in_filtered))

        # Charts
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("#### Verträge pro Vertragsart")
            vertragsart_counts = {}
            for c in filtered_contracts:
                art = c.get('vertragsart', 'Unbekannt')
                vertragsart_counts[art] = vertragsart_counts.get(art, 0) + 1

            if vertragsart_counts:
                fig = go.Figure(data=[
                    go.Pie(labels=list(vertragsart_counts.keys()),
                          values=list(vertragsart_counts.values()),
                          hole=0.3)
                ])
                fig.update_layout(height=450, showlegend=True)
                st.plotly_chart(fig, use_container_width=True)

        with col2:
            st.markdown("#### Vertragsstatistiken")
            nachtrag_count = len([c for c in filtered_contracts if c.get('ist_nachtrag')])
            original_count = len(filtered_contracts) - nachtrag_count

            stats_data = {
                'Status': ['Originale', 'Nachträge'],
                'Anzahl': [original_count, nachtrag_count]
            }

            fig = go.Figure(data=[
                go.Bar(x=stats_data['Status'], y=stats_data['Anzahl'], marker_color=['#4ECDC4', '#FF6B6B'])
            ])
            fig.update_layout(height=450, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

        # Projekte chart
        st.markdown("#### Verträge pro Projekt")
        contracts_by_projekt = get_contracts_by_projekt(vertices, edges, filtered_contracts)

        if contracts_by_projekt:
            projekt_names = list(contracts_by_projekt.keys())
            projekt_counts = [len(c) for c in contracts_by_projekt.values()]

            fig = go.Figure(data=[
                go.Bar(x=projekt_names, y=projekt_counts, marker_color='#3498DB')
            ])
            fig.update_layout(height=450, xaxis_title="Projekt", yaxis_title="Anzahl Verträge")
            st.plotly_chart(fig, use_container_width=True)

    # ── TAB 4: DETAIL VIEW ──
    with tab4:
        contract_names = [c.get('dateiname', c.get('vertrags_id')) for c in filtered_contracts]
        selected_contract_name = st.selectbox("Vertrag wählen", contract_names)

        if selected_contract_name:
            selected_contract = next((c for c in filtered_contracts
                                     if c.get('dateiname') == selected_contract_name), None)

            if selected_contract:
                col1, col2 = st.columns([2, 1])

                with col1:
                    st.markdown("#### Vertragsdetails")

                    info_cols = st.columns(3)
                    with info_cols[0]:
                        st.markdown(f"**ID:** {selected_contract.get('vertrags_id', 'N/A')}")
                        st.markdown(f"**Vertragsart:** {selected_contract.get('vertragsart', 'N/A')}")
                        st.markdown(f"**Status:** {'Nachtrag' if selected_contract.get('ist_nachtrag') else 'Original'}")
                    with info_cols[1]:
                        st.markdown(f"**Seiten:** {selected_contract.get('seitenzahl', 'N/A')}")
                        st.markdown(f"**Datei:** {selected_contract.get('dateiname', 'N/A')}")
                        standorte = selected_contract.get('standorte', [])
                        if isinstance(standorte, list):
                            standorte = ', '.join(standorte)
                        st.markdown(f"**Standorte:** {standorte or 'N/A'}")
                    with info_cols[2]:
                        # Datum
                        datum = selected_contract.get('datum', {})
                        found_dates = datum.get('found_dates', []) if datum else []
                        if found_dates:
                            st.markdown(f"**Datum:** {', '.join(found_dates)}")

                with col2:
                    st.markdown("#### Parteien")
                    for party in selected_contract.get('parteien', []):
                        st.markdown(f"• {party.get('name')} ({party.get('typ')})")

                # Finanzen & Energie
                finanzen = selected_contract.get('finanzen', {})
                energie = selected_contract.get('energie', {})

                if finanzen or energie:
                    fin_col, ene_col = st.columns(2)

                    with fin_col:
                        if finanzen:
                            st.markdown("#### Finanzen")
                            if finanzen.get('betraege'):
                                st.markdown(f"**Beträge:** {', '.join(finanzen['betraege'])} €")
                            if finanzen.get('preis_ct_kwh'):
                                st.markdown(f"**Preis (ct/kWh):** {', '.join(finanzen['preis_ct_kwh'])}")
                            if finanzen.get('prozente'):
                                st.markdown(f"**Prozente:** {', '.join(finanzen['prozente'])}%")

                    with ene_col:
                        if energie:
                            st.markdown("#### Energie")
                            if energie.get('malo_ids'):
                                st.markdown(f"**MaLo-IDs:** {', '.join(energie['malo_ids'])}")
                            if energie.get('eeg_refs'):
                                st.markdown(f"**EEG-Referenzen:** {', '.join(energie['eeg_refs'])}")

                # Ego Graph
                st.markdown("#### Ego-Graph (Direkte Verbindungen)")

                contract_id = selected_contract.get('vertrags_id')

                if contract_id in G:
                    ego = get_ego_graph(G, contract_id)

                    if ego.number_of_nodes() > 0:
                        try:
                            ego_vis = create_pyvis_graph(ego)
                            html_string = ego_vis.generate_html()
                            st.components.v1.html(html_string, height=900)
                        except Exception as e:
                            st.error(f"Fehler bei Ego-Graph: {e}")
                    else:
                        st.info("Keine direkten Verbindungen gefunden.")

                # Text preview
                if selected_contract.get('text_preview') and selected_contract.get('text_preview') != "[KEIN TEXT]":
                    st.markdown("#### Textvorschau")
                    st.text_area("Vorschau", selected_contract.get('text_preview', ''), height=200, disabled=True)

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #7F8C8D; font-size: 12px;'>
        TERP Vertragsmanagement | Powered by Streamlit
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
