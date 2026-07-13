# ============================================================
# INTERACTIVE GISN EXPLORER
# 20 national tourism strategies in a single interactive HTML
#
# Output:
#   GISN_interactive/GISN_interactive_explorer.html
#
# Required packages:
#   pip install pandas openpyxl numpy networkx plotly
# ============================================================

from __future__ import annotations

import json
import os
import re
import unicodedata
from pathlib import Path
from typing import Any

import networkx as nx
import numpy as np
import pandas as pd

try:
    from plotly.offline import get_plotlyjs
except ImportError as exc:
    raise ImportError(
        "Plotly is required. Install it with: pip install plotly"
    ) from exc


# ============================================================
# 1. CONFIGURATION
# ============================================================

EXCEL_FILE = "Datos_finales.xlsx"
EDGE_SHEET = "Sustainability Integration Edge"
NODE_SHEET = "GISN_Node_Attributes"

OUTPUT_DIR = Path("GISN_interactive")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

OUTPUT_HTML = OUTPUT_DIR / "GISN_interactive_explorer.html"

# Common visual scales across all 20 countries
MIN_NODE_DIAMETER = 16.0
MAX_NODE_DIAMETER = 54.0
NODE_SIZE_EXPONENT = 1.30

EDGE_WIDTH_MAP = {
    1: 1.2,
    2: 2.5,
    3: 4.0,
    4: 6.0,
}

EDGE_COLOR_MAP = {
    1: "#BDBDBD",
    2: "#666666",
    3: "#E5B82E",
    4: "#27AE60",
}

TYPOLOGY_COLORS = {
    "Mature Integrated Hub": "#1f77b4",
    "Emerging Hub": "#ff7f0e",
    "Mature but Isolated": "#2ca02c",
    "Marginal Theme": "#d62728",
    "Mixed / tied": "#8c8c8c",
    "Unknown": "#bdbdbd",
}


# ============================================================
# 2. STRATEGY METADATA AND OFFICIAL LINKS
# ============================================================
# These fields populate the information card above the network.
# Edit any title, institution or URL here if a source changes.

STRATEGY_METADATA: dict[str, dict[str, str]] = {
    "Austria": {
        "title": "Plan T – Master Plan for Tourism",
        "published": "2019",
        "institution": "Federal Ministry for Sustainability and Tourism (BMNT)",
        "language": "English (official translation; German original)",
        "period": "Not specified",
        "source_url": "https://www.bmwet.gv.at/en/Topics/tourism/plan-t.html",
        "document_url": "https://www.bmwet.gv.at/dam/jcr%3A0ea14456-ac84-4d66-ac69-d507317cd3f2/PLAN%20T%20-%20MASTER%20PLAN%20FOR%20TOURISM.pdf",
    },
    "Bulgaria": {
        "title": "Updated National Strategy for Sustainable Tourism Development in the Republic of Bulgaria 2014–2030",
        "published": "2017",
        "institution": "Ministry of Tourism",
        "language": "Bulgarian",
        "period": "2014–2030",
        "source_url": "https://www.tourism.government.bg/bg/kategorii/strategicheski-dokumenti/aktualizirana-nacionalna-strategiya-za-ustoychivo-razvitie-na",
        "document_url": "https://www.tourism.government.bg/sites/tourism.government.bg/files/documents/2018-01/nsurtb_2014-2030.pdf",
    },
    "Croatia": {
        "title": "Sustainable Tourism Development Strategy until 2030",
        "published": "2022",
        "institution": "Ministry of Tourism and Sport",
        "language": "Croatian",
        "period": "2022–2030",
        "source_url": "https://mint.gov.hr/strategija-razvoja-odrzivog-turizma-do-2030-godine/11411",
        "document_url": "https://narodne-novine.nn.hr/clanci/sluzbeni/full/2023_01_2_18.html",
    },
    "Czech Republic": {
        "title": "Tourism Development Strategy of the Czech Republic 2021–2030",
        "published": "2021",
        "institution": "Ministry of Regional Development",
        "language": "Czech",
        "period": "2021–2030",
        "source_url": "https://mmr.gov.cz/cs/ministerstvo/cestovni-ruch/pro-profesionaly/koncepce-strategie/strategie-rozvoje-cestovniho-ruchu-cr-2021-2030",
        "document_url": "https://mmr.gov.cz/getattachment/Ministerstvo/Cestovni-ruch/Pro-profesionaly/Koncepce-Strategie/Strategie-rozvoje-cestovniho-ruchu-CR-2021-2030/Dokumenty/Strategie-CR-2021-2030/Strategie-rozvoje-CR-CR-2021-2030.pdf.aspx?ext=.pdf&lang=cs-CZ",
    },
    "Denmark": {
        "title": "National Strategy for Sustainable Growth in Danish Tourism",
        "published": "2022",
        "institution": "National Tourism Forum / Ministry of Industry, Business and Financial Affairs",
        "language": "Danish",
        "period": "2022–2030",
        "source_url": "https://www.em.dk/aktuelt/udgivelser-og-aftaler/2022/jun/national-strategi-for-baeredygtig-vaekst-i-dansk-turisme",
        "document_url": "https://www.em.dk/media/15417/skaerm_national-strategi-for-baeredygtig-vaekst-i-dansk-turisme_090822.pdf",
    },
    "Estonia": {
        "title": "Tourism Strategy 2022–2025: It’s about time",
        "published": "2022",
        "institution": "Ministry of Economic Affairs and Communications",
        "language": "Estonian",
        "period": "2022–2025",
        "source_url": "https://mkm.ee/ettevotlus-ja-innovatsioon/turism",
        "document_url": "https://www.mkm.ee/media/7613/download",
    },
    "Finland": {
        "title": "Achieving more together – sustainable growth and renewal in Finnish tourism",
        "published": "2025",
        "institution": "Ministry of Economic Affairs and Employment",
        "language": "Finnish",
        "period": "2025–2028",
        "source_url": "https://tem.fi/en/finland-tourism-strategy",
        "document_url": "https://urn.fi/URN:ISBN:978-952-327-805-9",
    },
    "France": {
        "title": "Destination France: Plan de reconquête et de transformation du tourisme",
        "published": "2021",
        "institution": "Government of France / Ministry responsible for Tourism",
        "language": "French",
        "period": "2021–2030",
        "source_url": "https://www.atout-france.fr/fr/plan-destination-france",
        "document_url": "https://www.atout-france.fr/sites/default/files/2023-10/Dossier%20de%20presse%20Plan%20Destination%20France.pdf",
    },
    "Hungary": {
        "title": "National Tourism Development Strategy 2030 – Tourism 2.0",
        "published": "2021",
        "institution": "Hungarian Tourism Agency",
        "language": "Hungarian",
        "period": "2021–2030",
        "source_url": "https://mtu.gov.hu/cikkek/strategia/",
        "document_url": "https://mtu.gov.hu/dokumentumok/NTS2030_Turizmus2.0-Strategia.pdf",
    },
    "Ireland": {
        "title": "Tourism Policy Framework 2025–2030",
        "published": "2024",
        "institution": "Department of Tourism, Culture, Arts, Gaeltacht, Sport and Media",
        "language": "English",
        "period": "2025–2030",
        "source_url": "https://www.gov.ie/en/department-of-enterprise-tourism-and-employment/publications/tourism-policy-framework/",
        "document_url": "https://assets.gov.ie/static/documents/tourism-policy-framework-20252030-a38dae8b-f1e6-4719-8929-bba74b7b5b22.pdf",
    },
    "Italy": {
        "title": "Strategic Tourism Plan 2023–2027",
        "published": "2023",
        "institution": "Ministry of Tourism",
        "language": "Italian",
        "period": "2023–2027",
        "source_url": "https://www.ministeroturismo.gov.it/piano-strategico-del-turismo-pst/",
        "document_url": "https://www.ministeroturismo.gov.it/wp-content/uploads/2024/09/Volume_PST_Settembre_2024_web_B.pdf",
    },
    "Lithuania": {
        "title": "Lithuanian Tourism Roadmap",
        "published": "2024",
        "institution": "Ministry of the Economy and Innovation",
        "language": "Lithuanian",
        "period": "2024–2030",
        "source_url": "https://eimin.lrv.lt/en/structure-and-contacts/news-1/eimin-lithuanias-tourism-roadmap-approved/",
        "document_url": "https://eimin.lrv.lt/media/viesa/saugykla/2024/6/FS59Q67N6oE.pdf",
    },
    "Luxembourg": {
        "title": "People, Regions and Economy: Luxembourg Tourism Strategy",
        "published": "2022",
        "institution": "Ministry of the Economy, Directorate-General for Tourism",
        "language": "German (title also presented in Luxembourgish)",
        "period": "Not specified",
        "source_url": "https://gouvernement.lu/fr/actualites/toutes_actualites/communiques/2022/05-mai/20-delles-strategie-touristique.html",
        "document_url": "https://gouvernement.lu/dam-assets/documents/actualites/2022/05-mai/20-delles-tourisus/tourismus-strategie.pdf",
    },
    "Malta": {
        "title": "Malta’s Tourism Strategy 2021–2030: Recover, Rethink and Revitalise",
        "published": "2021",
        "institution": "Ministry for Tourism and Consumer Protection / Malta Tourism Authority",
        "language": "English",
        "period": "2021–2030",
        "source_url": "https://tourism.gov.mt/resources/",
        "document_url": "https://tourism.gov.mt/wp-content/uploads/2023/04/National-Tourism-Strategy-2021-2030.pdf",
    },
    "Netherlands": {
        "title": "Perspective Destination Netherlands 2030",
        "published": "2019",
        "institution": "Netherlands Board of Tourism & Conventions (NBTC)",
        "language": "English (official edition; Dutch original)",
        "period": "2019–2030",
        "source_url": "https://www.nbtc.nl/en/site/about-us/what-we-do-at-nbtc/perspective-destination-netherlands-2030",
        "document_url": "https://www.nbtc.nl/en/site/download/perspective-destination-nl-2030-en?disposition=inline",
    },
    "Portugal": {
        "title": "Tourism Strategy 2027",
        "published": "2017",
        "institution": "Turismo de Portugal, I.P. / Government of Portugal",
        "language": "Portuguese",
        "period": "2017–2027",
        "source_url": "https://www.turismodeportugal.pt/pt/Turismo_Portugal/Estrategia/Estrategia_2027/Paginas/default.aspx",
        "document_url": "https://www.turismodeportugal.pt/SiteCollectionDocuments/estrategia/estrategia-turismo-2027.pdf",
    },
    "Romania": {
        "title": "National Strategy of Romania for Tourism Development 2024–2035",
        "published": "2024",
        "institution": "Ministry of Economy, Entrepreneurship and Tourism",
        "language": "Romanian",
        "period": "2024–2035",
        "source_url": "https://economie.gov.ro/proiect-de-hotarare-privind-aprobarea-strategiei-nationale-a-romaniei-pentru-dezvoltarea-turismului-2024-2035/",
        "document_url": "https://economie.gov.ro/wp-content/uploads/2024/02/SNDT-2024-2035.docx",
    },
    "Slovenia": {
        "title": "Slovenian Tourism Strategy 2022–2028",
        "published": "2022",
        "institution": "Ministry of Economic Development and Technology",
        "language": "English (official translation; Slovenian original)",
        "period": "2022–2028",
        "source_url": "https://www.gov.si/en/news/2022-05-10-government-adopts-the-new-seven-year-slovenian-tourism-strategy-2022-2028/",
        "document_url": "https://www.gov.si/assets/ministrstva/MGTS/Dokumenti/DTUR/Strategija-slovenskega-turizma-20222028-/SLOVENIAN-TOURISM-STRATEGY-2022-2028-v2.pdf",
    },
    "Spain": {
        "title": "General Guidelines for the Sustainable Tourism Strategy of Spain 2030",
        "published": "2019",
        "institution": "Secretariat of State for Tourism, Ministry of Industry, Trade and Tourism",
        "language": "Spanish",
        "period": "2019–2030",
        "source_url": "https://www.segittur.es/sala-de-prensa/planes-nacionales-de-turismo/",
        "document_url": "https://www.mintur.gob.es/Documents/Estrategia-Espana-Turismo-2030.pdf",
    },
    "Sweden": {
        "title": "Strategy for Sustainable Tourism and a Growing Visitor Economy",
        "published": "2021",
        "institution": "Government Offices of Sweden / Ministry of Enterprise and Innovation",
        "language": "Swedish",
        "period": "2021–2030",
        "source_url": "https://www.regeringen.se/contentassets/c8d7bec3060141ed934d94b73bdd80d5/regeringskansliets-arsbok-2021.pdf",
        "document_url": "https://s3-eu-west-1.amazonaws.com/images.corporate.visitsweden.com/documents/strategi-for-hallbar-turism-och-vaxande-besoksnaring_slutlig.pdf",
    },
}


# ============================================================
# 3. HELPERS
# ============================================================

def normalise_text(value: Any) -> str:
    value = str(value).strip().lower()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(
        character
        for character in value
        if not unicodedata.combining(character)
    )
    return re.sub(r"\s+", " ", value)


def standardise_typology(value: Any) -> str:
    replacements = {
        "institutionalised hub": "Mature Integrated Hub",
        "institutionalized hub": "Mature Integrated Hub",
        "mature integrated hub": "Mature Integrated Hub",
        "emerging hub": "Emerging Hub",
        "mature but isolated": "Mature but Isolated",
        "marginal/peripheral theme": "Marginal Theme",
        "marginal / peripheral theme": "Marginal Theme",
        "marginal theme": "Marginal Theme",
        "mixed / tied": "Mixed / tied",
    }
    normalised = normalise_text(value)
    return replacements.get(normalised, str(value).strip())


def canonical_pair(theme_a: Any, theme_b: Any) -> tuple[str, str]:
    return tuple(sorted((str(theme_a).strip(), str(theme_b).strip())))


def edge_category(weight: float) -> int:
    if weight <= 1:
        return 1
    if weight <= 2:
        return 2
    if weight <= 3:
        return 3
    return 4


def metadata_for_country(country: str) -> dict[str, str]:
    lookup = {
        normalise_text(name): metadata
        for name, metadata in STRATEGY_METADATA.items()
    }
    return lookup.get(
        normalise_text(country),
        {
            "title": country,
            "published": "Not available",
            "institution": "Not available",
            "language": "Not available",
            "period": "Not available",
            "source_url": "",
            "document_url": "",
        },
    )


def json_for_html(value: Any) -> str:
    # Prevent an accidental closing script tag inside embedded JSON.
    return json.dumps(value, ensure_ascii=False).replace("</", "<\\/")


# ============================================================
# 4. LOAD AND CLEAN THE EXCEL DATA
# ============================================================

try:
    edges = pd.read_excel(EXCEL_FILE, sheet_name=EDGE_SHEET)
    nodes = pd.read_excel(EXCEL_FILE, sheet_name=NODE_SHEET)
except FileNotFoundError as exc:
    raise FileNotFoundError(
        f"Could not find '{EXCEL_FILE}'. Place the script in the same "
        "folder as the workbook or edit EXCEL_FILE."
    ) from exc
except ValueError as exc:
    raise ValueError(
        f"Could not read '{EDGE_SHEET}' or '{NODE_SHEET}'. "
        f"Original error: {exc}"
    ) from exc

edges.columns = edges.columns.astype(str).str.strip()
nodes.columns = nodes.columns.astype(str).str.strip()

required_edge_columns = {"Country", "Theme A", "Theme B", "Weight"}
required_node_columns = {
    "Country",
    "Theme",
    "Weighted Degree",
    "GISN Typology",
}

missing_edges = required_edge_columns - set(edges.columns)
missing_nodes = required_node_columns - set(nodes.columns)

if missing_edges:
    raise KeyError(
        "Missing columns in the edge sheet: " + ", ".join(sorted(missing_edges))
    )
if missing_nodes:
    raise KeyError(
        "Missing columns in the node sheet: " + ", ".join(sorted(missing_nodes))
    )

for column in ["Country", "Theme A", "Theme B"]:
    edges[column] = edges[column].astype("string").str.strip()

for column in ["Country", "Theme", "GISN Typology"]:
    nodes[column] = nodes[column].astype("string").str.strip()

edges = edges.dropna(subset=["Country", "Theme A", "Theme B"]).copy()
nodes = nodes.dropna(subset=["Country", "Theme"]).copy()

edges["Weight"] = pd.to_numeric(edges["Weight"], errors="coerce").fillna(0)
nodes["Weighted Degree"] = pd.to_numeric(
    nodes["Weighted Degree"], errors="coerce"
).fillna(0)

if "Maturity Score" in nodes.columns:
    nodes["Maturity Score"] = pd.to_numeric(
        nodes["Maturity Score"], errors="coerce"
    )

edges = edges.loc[edges["Weight"] > 0].copy()
nodes["GISN Typology Standardised"] = nodes["GISN Typology"].apply(
    standardise_typology
)

countries = sorted(nodes["Country"].dropna().unique().tolist())
themes = sorted(nodes["Theme"].dropna().unique().tolist())

if len(countries) != 20:
    print(
        f"Warning: expected 20 countries, but found {len(countries)}: {countries}"
    )
if len(themes) != 15:
    print(f"Warning: expected 15 themes, but found {len(themes)}.")


# ============================================================
# 5. BUILD COUNTRY GRAPHS
# ============================================================

def build_country_graph(country_name: str) -> nx.Graph:
    country_nodes = nodes.loc[nodes["Country"] == country_name].copy()
    country_edges = edges.loc[edges["Country"] == country_name].copy()

    graph = nx.Graph()

    for theme in themes:
        matching = country_nodes.loc[country_nodes["Theme"] == theme]

        if matching.empty:
            attributes = {
                "weighted_degree": 0.0,
                "degree_reported": np.nan,
                "maturity_score": np.nan,
                "maturity_stage": "Not available",
                "typology": "Unknown",
            }
        else:
            row = matching.iloc[0]
            attributes = {
                "weighted_degree": float(row["Weighted Degree"]),
                "degree_reported": (
                    float(row["Degree"])
                    if "Degree" in row.index and pd.notna(row["Degree"])
                    else np.nan
                ),
                "maturity_score": (
                    float(row["Maturity Score"])
                    if "Maturity Score" in row.index
                    and pd.notna(row["Maturity Score"])
                    else np.nan
                ),
                "maturity_stage": (
                    str(row["Maturity Stage"])
                    if "Maturity Stage" in row.index
                    and pd.notna(row["Maturity Stage"])
                    else "Not available"
                ),
                "typology": str(row["GISN Typology Standardised"]),
            }

        graph.add_node(theme, **attributes)

    aggregated_edges: dict[tuple[str, str], float] = {}

    for _, row in country_edges.iterrows():
        theme_a = str(row["Theme A"]).strip()
        theme_b = str(row["Theme B"]).strip()
        weight = float(row["Weight"])

        if theme_a not in graph or theme_b not in graph:
            continue

        pair = canonical_pair(theme_a, theme_b)
        aggregated_edges[pair] = max(aggregated_edges.get(pair, 0.0), weight)

    for (theme_a, theme_b), weight in aggregated_edges.items():
        graph.add_edge(theme_a, theme_b, weight=weight)

    return graph


country_graphs = {
    country: build_country_graph(country)
    for country in countries
}


# ============================================================
# 6. COMMON NODE SCALE AND COMMON POSITIONS
# ============================================================

global_weighted_degrees = np.array(
    [
        graph.nodes[node].get("weighted_degree", 0.0)
        for graph in country_graphs.values()
        for node in graph.nodes()
    ],
    dtype=float,
)

global_wd_min = float(global_weighted_degrees.min())
global_wd_max = float(global_weighted_degrees.max())


def scale_node_diameter(weighted_degree: float) -> float:
    if global_wd_max == global_wd_min:
        return 30.0

    normalised = (weighted_degree - global_wd_min) / (
        global_wd_max - global_wd_min
    )
    normalised = float(np.clip(normalised, 0.0, 1.0))

    return MIN_NODE_DIAMETER + (
        MAX_NODE_DIAMETER - MIN_NODE_DIAMETER
    ) * normalised**NODE_SIZE_EXPONENT


union_graph = nx.Graph()
union_graph.add_nodes_from(themes)

for graph in country_graphs.values():
    union_graph.add_edges_from(graph.edges())

common_pos = nx.spring_layout(
    union_graph,
    seed=42,
    weight=None,
    k=1.40,
    iterations=6000,
    scale=3.7,
)

x_values = np.array([common_pos[theme][0] for theme in themes], dtype=float)
y_values = np.array([common_pos[theme][1] for theme in themes], dtype=float)

x_range = max(float(x_values.max() - x_values.min()), 1.0)
y_range = max(float(y_values.max() - y_values.min()), 1.0)

x_padding = x_range * 0.32
y_padding = y_range * 0.32

common_x_range = [
    float(x_values.min() - x_padding),
    float(x_values.max() + x_padding),
]
common_y_range = [
    float(y_values.min() - y_padding),
    float(y_values.max() + y_padding),
]

centre_x = float(np.mean(x_values))
centre_y = float(np.mean(y_values))


# ============================================================
# 7. SERIALISE THE NETWORKS FOR JAVASCRIPT
# ============================================================

networks: dict[str, dict[str, Any]] = {}

for country, graph in country_graphs.items():
    node_records: list[dict[str, Any]] = []

    for theme in themes:
        attributes = graph.nodes[theme]
        x, y = common_pos[theme]
        dx = float(x - centre_x)
        dy = float(y - centre_y)

        if abs(dx) >= abs(dy):
            text_position = "middle right" if dx >= 0 else "middle left"
        else:
            text_position = "top center" if dy >= 0 else "bottom center"

        degree_value = attributes.get("degree_reported", np.nan)
        if pd.isna(degree_value):
            degree_value = float(graph.degree(theme))

        maturity_score = attributes.get("maturity_score", np.nan)

        node_records.append(
            {
                "theme": theme,
                "x": float(x),
                "y": float(y),
                "size": float(
                    scale_node_diameter(
                        float(attributes.get("weighted_degree", 0.0))
                    )
                ),
                "color": TYPOLOGY_COLORS.get(
                    str(attributes.get("typology", "Unknown")),
                    TYPOLOGY_COLORS["Unknown"],
                ),
                "typology": str(attributes.get("typology", "Unknown")),
                "weighted_degree": float(
                    attributes.get("weighted_degree", 0.0)
                ),
                "degree": float(degree_value),
                "maturity_score": (
                    None if pd.isna(maturity_score) else float(maturity_score)
                ),
                "maturity_stage": str(
                    attributes.get("maturity_stage", "Not available")
                ),
                "text_position": text_position,
            }
        )

    edge_records: list[dict[str, Any]] = []

    for theme_a, theme_b, attributes in graph.edges(data=True):
        weight = float(attributes.get("weight", 1.0))
        category = edge_category(weight)
        x0, y0 = common_pos[theme_a]
        x1, y1 = common_pos[theme_b]

        edge_records.append(
            {
                "theme_a": theme_a,
                "theme_b": theme_b,
                "x0": float(x0),
                "y0": float(y0),
                "x1": float(x1),
                "y1": float(y1),
                "weight": weight,
                "width": EDGE_WIDTH_MAP[category],
                "color": EDGE_COLOR_MAP[category],
            }
        )

    networks[country] = {
        "metadata": metadata_for_country(country),
        "nodes": node_records,
        "edges": edge_records,
    }


# ============================================================
# 8. LEGEND VALUES
# ============================================================

legend_values = np.linspace(0.0, float(np.ceil(global_wd_max)), 4)
legend_values = [float(round(value, 1)) for value in legend_values]
legend_diameters = [
    float(round(scale_node_diameter(value), 1))
    for value in legend_values
]


# ============================================================
# 9. BUILD THE STANDALONE HTML
# ============================================================

plotly_javascript = get_plotlyjs()
network_json = json_for_html(networks)
country_options = "\n".join(
    f'<option value="{country}">{country}</option>'
    for country in countries
)

node_legend_html = "\n".join(
    f"""
    <div class="size-item">
        <span class="size-circle" style="width:{diameter:.1f}px;height:{diameter:.1f}px"></span>
        <span>{value:.1f}</span>
    </div>
    """
    for value, diameter in zip(legend_values, legend_diameters)
)

html_document = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Interactive GISN Explorer</title>
<style>
    :root {{
        --border: #d7d7d7;
        --text: #202124;
        --muted: #666;
        --panel: #ffffff;
        --background: #f5f6f8;
        --accent: #1f5f99;
    }}

    * {{ box-sizing: border-box; }}

    body {{
        margin: 0;
        font-family: Arial, Helvetica, sans-serif;
        color: var(--text);
        background: var(--background);
    }}

    .page {{
        max-width: 1500px;
        margin: 0 auto;
        padding: 22px;
    }}

    header {{
        text-align: center;
        margin-bottom: 16px;
    }}

    h1 {{
        margin: 0;
        font-size: clamp(24px, 3vw, 38px);
    }}

    .subtitle {{
        margin-top: 7px;
        color: var(--muted);
        font-size: 15px;
    }}

    .controls {{
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        align-items: center;
        padding: 14px 16px;
        border: 1px solid var(--border);
        border-radius: 12px;
        background: var(--panel);
        margin-bottom: 14px;
    }}

    label {{ font-weight: 700; }}

    select, button {{
        border: 1px solid #aeb4bb;
        border-radius: 8px;
        padding: 9px 12px;
        background: white;
        font-size: 14px;
    }}

    select {{ min-width: 210px; }}
    button {{ cursor: pointer; }}
    button:hover {{ background: #f1f3f4; }}

    .check-label {{
        display: inline-flex;
        align-items: center;
        gap: 7px;
        font-weight: 400;
    }}

    .strategy-card {{
        background: var(--panel);
        border: 1px solid var(--border);
        border-radius: 12px;
        padding: 16px 18px;
        margin-bottom: 14px;
    }}

    .strategy-card h2 {{
        margin: 0 0 8px 0;
        font-size: 20px;
    }}

    .meta-grid {{
        display: grid;
        grid-template-columns: repeat(4, minmax(150px, 1fr));
        gap: 8px 18px;
        font-size: 13px;
        color: #3c4043;
    }}

    .meta-wide {{ grid-column: span 2; }}

    .links {{
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        margin-top: 14px;
    }}

    .link-button {{
        display: inline-block;
        text-decoration: none;
        color: white;
        background: var(--accent);
        border-radius: 8px;
        padding: 9px 13px;
        font-weight: 700;
        font-size: 13px;
    }}

    .link-button.secondary {{ background: #555f6a; }}
    .link-button[aria-disabled="true"] {{
        pointer-events: none;
        opacity: 0.45;
    }}

    .content-grid {{
        display: grid;
        grid-template-columns: minmax(0, 1fr) 280px;
        gap: 14px;
        align-items: stretch;
    }}

    .plot-card, .legend-card {{
        background: var(--panel);
        border: 1px solid var(--border);
        border-radius: 12px;
    }}

    .plot-card {{ padding: 4px; }}
    #gisnPlot {{ width: 100%; height: 760px; }}

    .legend-card {{
        padding: 18px;
        font-size: 13px;
    }}

    .legend-card h3 {{
        text-align: center;
        margin: 0 0 16px 0;
        font-size: 18px;
    }}

    .legend-section {{
        border-top: 1px solid #ddd;
        padding-top: 14px;
        margin-top: 14px;
    }}

    .legend-section strong {{ display: block; margin-bottom: 10px; }}

    .legend-row {{
        display: flex;
        align-items: center;
        gap: 9px;
        margin: 9px 0;
    }}

    .colour-dot {{
        width: 17px;
        height: 17px;
        border-radius: 50%;
        border: 1px solid black;
        flex: 0 0 auto;
    }}

       .size-scale {{
        position: relative;
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        min-height: 118px;
        padding: 6px 0 0;
    }}

    /* Cono gris simétrico y moderado */
    .size-scale::before {{
        content: "";
        position: absolute;
        left: 4%;
        right: 4%;
        top: 10px;
        height: 76px;
        background: rgba(180, 180, 180, 0.20);

        clip-path: polygon(
            0% 47%,
            100% 34%,
            100% 66%,
            0% 53%
        );

        z-index: 0;
        pointer-events: none;
    }}

    .size-item {{
        position: relative;
        z-index: 1;

        display: grid;
        grid-template-rows: 78px 24px;
        justify-items: center;
        align-items: center;

        text-align: center;
        color: #444;
        font-size: 12px;
    }}

    .size-circle {{
        display: block;
        margin: 0;

        border-radius: 50%;
        border: 1px solid black;
        background: white;

        max-width: 52px;
        max-height: 52px;
    }}

    .edge-line {{
        width: 68px;
        display: inline-block;
        border-radius: 999px;
        flex: 0 0 auto;
    }}

    .edge-line {{
        width: 68px;
        display: inline-block;
        border-radius: 999px;
        flex: 0 0 auto;
    }}

    .figure-note {{
        color: var(--muted);
        font-size: 12.5px;
        line-height: 1.45;
        margin-top: 14px;
        text-align: center;
    }}

    @media (max-width: 950px) {{
        .content-grid {{ grid-template-columns: 1fr; }}
        .legend-card {{ order: 2; }}
        .meta-grid {{ grid-template-columns: 1fr 1fr; }}
        .meta-wide {{ grid-column: span 2; }}
        #gisnPlot {{ height: 650px; }}
    }}

    @media (max-width: 600px) {{
        .page {{ padding: 10px; }}
        .meta-grid {{ grid-template-columns: 1fr; }}
        .meta-wide {{ grid-column: span 1; }}
        #gisnPlot {{ height: 560px; }}
    }}
</style>
<script>{plotly_javascript}</script>
</head>
<body>
<div class="page">
    <header>
        <h1>Governance-Integrated Sustainability Network Explorer</h1>
        <div class="subtitle">Interactive comparison of 20 national tourism strategies</div>
    </header>

    <section class="controls">
        <label for="countrySelect">Country</label>
        <select id="countrySelect">{country_options}</select>

        <label class="check-label">
            <input type="checkbox" id="labelToggle" checked>
            Show theme labels
        </label>

        <button id="resetButton" type="button">Reset view</button>
    </section>

    <section class="strategy-card">
        <h2 id="strategyHeading"></h2>
        <div class="meta-grid">
            <div><strong>Published:</strong> <span id="published"></span></div>
            <div><strong>Policy period:</strong> <span id="period"></span></div>
            <div><strong>Language:</strong> <span id="language"></span></div>
            <div class="meta-wide"><strong>Responsible institution:</strong> <span id="institution"></span></div>
        </div>
        <div class="links">
            <a id="sourceLink" class="link-button" target="_blank" rel="noopener noreferrer">Official source webpage</a>
            <a id="documentLink" class="link-button secondary" target="_blank" rel="noopener noreferrer">Open strategy document</a>
        </div>
    </section>

    <main class="content-grid">
        <section class="plot-card">
            <div id="gisnPlot"></div>
        </section>

        <aside class="legend-card">
            <h3>Legend</h3>

            <div class="legend-section">
                <strong>GISN category — node colour</strong>
                <div class="legend-row"><span class="colour-dot" style="background:#1f77b4"></span>Mature Integrated Hub</div>
                <div class="legend-row"><span class="colour-dot" style="background:#ff7f0e"></span>Emerging Hub</div>
                <div class="legend-row"><span class="colour-dot" style="background:#2ca02c"></span>Mature but Isolated</div>
                <div class="legend-row"><span class="colour-dot" style="background:#d62728"></span>Marginal Theme</div>
            </div>

            <div class="legend-section">
                <strong>Weighted Degree — node size</strong>
                <div class="size-scale">{node_legend_html}</div>
            </div>

            <div class="legend-section">
                <strong>Coded thematic connections — edge width and colour</strong>
                <div class="legend-row"><span class="edge-line" style="height:1.2px;background:#BDBDBD"></span>1 coded connection</div>
                <div class="legend-row"><span class="edge-line" style="height:2.5px;background:#666666"></span>2 coded connections</div>
                <div class="legend-row"><span class="edge-line" style="height:4px;background:#E5B82E"></span>3 coded connections</div>
                <div class="legend-row"><span class="edge-line" style="height:6px;background:#27AE60"></span>4 or more coded connections</div>
            </div>
        </aside>
    </main>

    <div class="figure-note">
        Nodes represent the 15 sustainability themes. Node colour indicates the GISN category, and node size represents Weighted Degree. Edge width and colour indicate the number of coded thematic connections between each pair of themes. Identical node positions and visual scales are used across all country networks.
    </div>
</div>

<script>
const NETWORKS = {network_json};
const COMMON_X_RANGE = {json_for_html(common_x_range)};
const COMMON_Y_RANGE = {json_for_html(common_y_range)};

const countrySelect = document.getElementById('countrySelect');
const labelToggle = document.getElementById('labelToggle');
const resetButton = document.getElementById('resetButton');
const plotElement = document.getElementById('gisnPlot');

function setLink(anchor, url) {{
    if (url) {{
        anchor.href = url;
        anchor.setAttribute('aria-disabled', 'false');
    }} else {{
        anchor.removeAttribute('href');
        anchor.setAttribute('aria-disabled', 'true');
    }}
}}

function updateMetadata(country) {{
    const metadata = NETWORKS[country].metadata;

    document.getElementById('strategyHeading').textContent = `${{country}} — ${{metadata.title}}`;
    document.getElementById('published').textContent = metadata.published || 'Not available';
    document.getElementById('period').textContent = metadata.period || 'Not available';
    document.getElementById('language').textContent = metadata.language || 'Not available';
    document.getElementById('institution').textContent = metadata.institution || 'Not available';

    setLink(document.getElementById('sourceLink'), metadata.source_url);
    setLink(document.getElementById('documentLink'), metadata.document_url);
}}

function maturityText(node) {{
    const score = node.maturity_score === null ? 'Not available' : node.maturity_score.toFixed(1);
    return `<b>${{node.theme}}</b><br>` +
           `GISN category: ${{node.typology}}<br>` +
           `Weighted Degree: ${{node.weighted_degree.toFixed(1)}}<br>` +
           `Degree: ${{node.degree.toFixed(0)}}<br>` +
           `Maturity Score: ${{score}}<br>` +
           `Maturity Stage: ${{node.maturity_stage}}`;
}}

function buildTraces(country) {{
    const network = NETWORKS[country];
    const traces = [];

    network.edges.forEach(edge => {{
        traces.push({{
            type: 'scatter',
            mode: 'lines',
            x: [edge.x0, edge.x1],
            y: [edge.y0, edge.y1],
            line: {{
                color: edge.color,
                width: edge.width
            }},
            text: [
                `<b>${{edge.theme_a}} ↔ ${{edge.theme_b}}</b><br>Coded thematic connections: ${{edge.weight.toFixed(0)}}`,
                `<b>${{edge.theme_a}} ↔ ${{edge.theme_b}}</b><br>Coded thematic connections: ${{edge.weight.toFixed(0)}}`
            ],
            hovertemplate: '%{{text}}<extra></extra>',
            showlegend: false
        }});
    }});

    const nodes = network.nodes;
    const showLabels = labelToggle.checked;

    traces.push({{
        type: 'scatter',
        mode: showLabels ? 'markers+text' : 'markers',
        x: nodes.map(node => node.x),
        y: nodes.map(node => node.y),
        text: nodes.map(node => node.theme),
        textposition: nodes.map(node => node.text_position),
        textfont: {{
            family: 'Arial, sans-serif',
            size: 11,
            color: '#111'
        }},
        marker: {{
            size: nodes.map(node => node.size),
            color: nodes.map(node => node.color),
            line: {{color: '#111', width: 1.1}},
            opacity: 0.98
        }},
        customdata: nodes.map(node => maturityText(node)),
        hovertemplate: '%{{customdata}}<extra></extra>',
        cliponaxis: false,
        showlegend: false
    }});

    return traces;
}}

function buildLayout(country) {{
    return {{
        title: {{
            text: country,
            x: 0.02,
            xanchor: 'left',
            font: {{size: 22}}
        }},
        margin: {{l: 30, r: 30, t: 58, b: 30}},
        paper_bgcolor: '#ffffff',
        plot_bgcolor: '#ffffff',
        hovermode: 'closest',
        dragmode: 'pan',
        xaxis: {{
            visible: false,
            range: COMMON_X_RANGE,
            fixedrange: false
        }},
        yaxis: {{
            visible: false,
            range: COMMON_Y_RANGE,
            fixedrange: false,
            scaleanchor: 'x',
            scaleratio: 1
        }},
        uirevision: country,
        showlegend: false
    }};
}}

const plotConfig = {{
    responsive: true,
    displaylogo: false,
    scrollZoom: true,
    toImageButtonOptions: {{
        format: 'png',
        filename: 'GISN_network',
        height: 1000,
        width: 1500,
        scale: 2
    }}
}};

function renderCountry(country) {{
    updateMetadata(country);
    Plotly.react(
        plotElement,
        buildTraces(country),
        buildLayout(country),
        plotConfig
    );
}}

countrySelect.addEventListener('change', () => {{
    renderCountry(countrySelect.value);
}});

labelToggle.addEventListener('change', () => {{
    renderCountry(countrySelect.value);
}});

resetButton.addEventListener('click', () => {{
    Plotly.relayout(plotElement, {{
        'xaxis.range': COMMON_X_RANGE,
        'yaxis.range': COMMON_Y_RANGE
    }});
}});

renderCountry(countrySelect.value);
</script>
</body>
</html>
"""

OUTPUT_HTML.write_text(html_document, encoding="utf-8")

print("Interactive GISN explorer generated successfully.")
print(f"Countries included: {len(countries)}")
print(f"Output: {OUTPUT_HTML.resolve()}")
print("Open the HTML file in a browser or publish it with GitHub Pages.")
