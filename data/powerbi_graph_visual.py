# ============================================================
# Power BI Python Visual Script
# TERP Vertragsmanagement — Graph-Visualisierung
# ============================================================
# Dieses Script wird im Power BI Python Visual verwendet.
#
# SETUP in Power BI:
# 1. Daten abrufen → CSV → vertices_powerbi.csv laden
# 2. Daten abrufen → CSV → edges_powerbi.csv laden
# 3. Python Visual auf Report-Seite ziehen
# 4. Aus der edges-Tabelle: src, dst, relationship ins Visual ziehen
# 5. Dieses Script im Script-Editor einfügen
# ============================================================

import networkx as nx
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

# --- Power BI übergibt die Daten als 'dataset' DataFrame ---
# Spalten: src, dst, relationship

# Graph aufbauen
G = nx.Graph()

for _, row in dataset.iterrows():
    src = str(row['src'])
    dst = str(row['dst'])
    rel = str(row.get('relationship', ''))

    # Knotentyp aus der ID ableiten (format: "typ:hash")
    src_type = src.split(':')[0] if ':' in src else 'unknown'
    dst_type = dst.split(':')[0] if ':' in dst else 'unknown'

    # Kurzname für die Anzeige (nach dem Doppelpunkt, max 20 Zeichen)
    src_short = src.split(':')[1][:15] if ':' in src else src[:15]
    dst_short = dst.split(':')[1][:15] if ':' in dst else dst[:15]

    G.add_node(src, node_type=src_type, short_name=src_short)
    G.add_node(dst, node_type=dst_type, short_name=dst_short)
    G.add_edge(src, dst, relationship=rel)

# --- Farben und Größen nach Knotentyp (Streamlit-App-Farben) ---
color_map = {
    'vertrag':  '#FABBC3',   # Pastell Rosa
    'firma':    '#CAEEE9',   # Pastell Mint
    'person':   '#FFEBC0',   # Pastell Gelb
    'projekt':  '#D8D8EE',   # Pastell Lila
    'dokument': '#E5E5E5',   # Hell Grau
    'malo':     '#E6E0F0',   # Pastell Lila
    'anlage':   '#D4E7F6',   # Pastell Blau
}

border_map = {
    'vertrag':  '#E8949F',
    'firma':    '#9DD5CC',
    'person':   '#E8CF97',
    'projekt':  '#B3B3D6',
    'dokument': '#C0C0C0',
    'malo':     '#C5BAD9',
    'anlage':   '#A8CCE6',
}

size_map = {
    'vertrag':  250,
    'firma':    400,
    'person':   150,
    'projekt':  350,
    'dokument': 50,
    'malo':     50,
    'anlage':   300,
}

node_colors = [color_map.get(G.nodes[n].get('node_type', ''), '#BDC3C7') for n in G.nodes()]
node_sizes = [size_map.get(G.nodes[n].get('node_type', ''), 100) for n in G.nodes()]

# --- Layout berechnen ---
pos = nx.spring_layout(G, k=1.5, iterations=50, seed=42)

# --- Zeichnen ---
fig, ax = plt.subplots(1, 1, figsize=(14, 10))
fig.patch.set_facecolor('#FFFFFF')
ax.set_facecolor('#FAFAFA')

# Kanten zeichnen (dünn, transparent)
nx.draw_networkx_edges(G, pos, ax=ax,
                       edge_color='#D5D8DC',
                       alpha=0.4,
                       width=0.5)

# Knoten zeichnen
nx.draw_networkx_nodes(G, pos, ax=ax,
                       node_color=node_colors,
                       node_size=node_sizes,
                       alpha=0.85,
                       edgecolors='white',
                       linewidths=0.5)

# Labels nur für große Knoten (Firma, Projekt)
labels_to_show = {n: G.nodes[n].get('short_name', '')
                  for n in G.nodes()
                  if G.nodes[n].get('node_type') in ('firma', 'projekt')}

nx.draw_networkx_labels(G, pos, labels_to_show, ax=ax,
                        font_size=6,
                        font_color='#2C3E50',
                        font_weight='bold')

# --- Legende ---
legend_items = [
    mpatches.Patch(color='#E74C3C', label=f'Vertrag ({sum(1 for n in G.nodes() if G.nodes[n].get("node_type")=="vertrag")})'),
    mpatches.Patch(color='#3498DB', label=f'Firma ({sum(1 for n in G.nodes() if G.nodes[n].get("node_type")=="firma")})'),
    mpatches.Patch(color='#F39C12', label=f'Person ({sum(1 for n in G.nodes() if G.nodes[n].get("node_type")=="person")})'),
    mpatches.Patch(color='#2ECC71', label=f'Projekt ({sum(1 for n in G.nodes() if G.nodes[n].get("node_type")=="projekt")})'),
    mpatches.Patch(color='#95A5A6', label=f'Dokument ({sum(1 for n in G.nodes() if G.nodes[n].get("node_type")=="dokument")})'),
]
ax.legend(handles=legend_items, loc='upper left', fontsize=8,
          framealpha=0.9, edgecolor='#E0E0E0')

# Titel
ax.set_title(f'TERP Vertragsgraph — {G.number_of_nodes()} Knoten, {G.number_of_edges()} Kanten',
             fontsize=12, fontweight='bold', color='#2C3E50', pad=15)

ax.axis('off')
plt.tight_layout()
plt.show()
