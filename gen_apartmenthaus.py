#!/usr/bin/env python3
"""
gen_apartmenthaus.py
====================
Erzeugt den Apartmenthaus-Tab für das Ostseeliebe Dashboard.
Liest buchungen_export_2027.csv und Objektstammdaten.xlsx und
erstellt eine HTML-Datei mit Buchungsvergleich pro Apartmenthaus.

Aufruf:
    python gen_apartmenthaus.py
    python gen_apartmenthaus.py --csv buchungen_export_2027.csv --output apartmenthaus_tab.html
"""

import argparse
import sys
import os
import json
from pathlib import Path

# ──────────────────────────────────────────────────────────────────
# 1. APARTMENTHAUS → UNTERKUNFT MAPPING  (aus Contao gescrapt)
# ──────────────────────────────────────────────────────────────────
APARTMENTHAUS_MAPPING = {
    "Apartmenthaus Viktoria Luna":        [{"nr":174,"name":"Elli"},{"nr":140,"name":"Kleine Lagune"}],
    "Darßer Landhaus":                    [{"nr":269,"name":"Ruheoase"},{"nr":270,"name":"Wellnessoase"}],
    "Das Feriendomizil Traumzeit Zingst": [{"nr":218,"name":"Traumzeit 1"},{"nr":256,"name":"Traumzeit 2"}],
    "Das Feriendomizil Zingst – Strandhafer & Strandkaten": [{"nr":259,"name":"RC - Strandhafer"},{"nr":258,"name":"RC - Strandkaten"}],
    "Ferienhäuser Alter Schwede":         [{"nr":201,"name":"Alter Schwede 1"},{"nr":202,"name":"Alter Schwede 2"},{"nr":203,"name":"Alter Schwede 3"}],
    "Ferienhäuser grüner Winkel":         [{"nr":56,"name":"RC - Charlotte"},{"nr":57,"name":"RC - Therese"}],
    "Ferienwohnung Velo":                 [{"nr":77,"name":"Velo 1"},{"nr":78,"name":"Velo 2"},{"nr":79,"name":"Velo 3"},{"nr":138,"name":"Velo 4"}],
    "Häuser Boddenwind":                  [],
    "Haus Achtern Dieck":                 [],
    "Haus am kleinen Hafen":              [{"nr":273,"name":"kl. Hafen No 1"},{"nr":274,"name":"kl. Hafen No 2"},{"nr":275,"name":"kl. Hafen No 3"}],
    "Haus Bi de Wisch":                   [{"nr":221,"name":"Bi de Wisch EG"},{"nr":222,"name":"Bi de Wisch OG"}],
    "Haus Blaue Wieck":                   [{"nr":260,"name":"DWR Mitte"},{"nr":261,"name":"OSS rechts"},{"nr":262,"name":"WSL links"}],
    "Haus Bliesenrade":                   [{"nr":271,"name":"Bernstein"}],
    "Haus Boddenblick":                   [{"nr":283,"name":"Boddenblick 9"}],
    "Haus Chausseestraße":                [{"nr":281,"name":"Moin Moin"},{"nr":278,"name":"Santa Karina Born"}],
    "Haus Citynah":                       [{"nr":137,"name":"Frau Zander"},{"nr":194,"name":"Möwennest Nr. 5"},{"nr":166,"name":"Seewind"}],
    "Haus Cozy Five":                     [{"nr":190,"name":"Cozy 1"},{"nr":191,"name":"Cozy 2"},{"nr":192,"name":"Cozy 3"},{"nr":131,"name":"Cozy 4"},{"nr":132,"name":"Cozy 5"}],
    "Haus Darssduett":                    [{"nr":279,"name":"Darss-Duett 1"},{"nr":290,"name":"Darss-Duett 2"}],
    "Haus Darßer Sonnenfisch":            [{"nr":228,"name":"Clownfisch"},{"nr":247,"name":"Sonnendeck"}],
    "Haus DünenLiebe":                    [],
    "Häuser Haseneck":                    [{"nr":129,"name":"DAT KROEGER HUS"},{"nr":139,"name":"Kranschehus"}],
    "Haus Hoppenberg Strandquartier":     [{"nr":145,"name":"Bärbel"},{"nr":144,"name":"Lotta"}],
    "Haus im Zentrum":                    [{"nr":114,"name":"Ocean Star"},{"nr":99,"name":"Rewal"},{"nr":115,"name":"Strandhaus Zingst"}],
    "Haus In den Wiesen":                 [{"nr":282,"name":"In den Wiesen App. 2"},{"nr":227,"name":"In den Wiesen 3"}],
    "Haus Kraanstiet":                    [{"nr":293,"name":"Kraanstiet 1"},{"nr":294,"name":"Kraanstiet 2"},{"nr":295,"name":"Kraanstiet 3"}],
    "Haus Küstentour":                    [{"nr":90,"name":"Dünenläufer"},{"nr":92,"name":"Ostseekoje"},{"nr":94,"name":"Sandbank No. 4"},{"nr":91,"name":"Strandsegler"}],
    "Haus Küstenzauber":                  [{"nr":98,"name":"Bremen"},{"nr":97,"name":"Stralsund"}],
    "Haus Meerkaten":                     [{"nr":168,"name":"Lust auf Meer"},{"nr":148,"name":"Schifferkaten"}],
    "Haus Meerle":                        [{"nr":288,"name":"RC - Meerle 1"},{"nr":291,"name":"RC - Meerle 2"},{"nr":296,"name":"RC - Meerle 3"},{"nr":297,"name":"RC - Meerle 4"},{"nr":299,"name":"RC - Meerle 5"}],
    "Haus Meerquartier":                  [],
    "Haus Öresundhus":                    [{"nr":267,"name":"Öresundhus Whg.2"},{"nr":268,"name":"Öresundhus Whg.3"}],
    "Haus PanoramaLiebe":                 [],
    "Haus ParkLiebe":                     [],
    "Haus Quartett Küstenglück":          [{"nr":195,"name":"Windland"},{"nr":196,"name":"Windspiel"},{"nr":197,"name":"Passatwind"},{"nr":198,"name":"Wellenbrecher"}],
    "Haus Reetzeit":                      [{"nr":95,"name":"Reetzeit 1"},{"nr":96,"name":"Reetzeit 2"}],
    "Haus Rosenberg Küstenharmonie":      [{"nr":108,"name":"Kranichrast"},{"nr":158,"name":"Meeresbrise"},{"nr":123,"name":"Zeeskahn"}],
    "Haus Schwedengang":                  [{"nr":189,"name":"Mondzauber"},{"nr":188,"name":"Sonnenschein"}],
    "Haus Seeluft & Seestern":            [{"nr":287,"name":"Seeluft"},{"nr":286,"name":"Seestern"}],
    "Haus Sterntaucher":                  [{"nr":231,"name":"Sterntaucher 1"},{"nr":232,"name":"Sterntaucher 2"},{"nr":233,"name":"Sterntaucher 3"},{"nr":234,"name":"Sterntaucher 4"},{"nr":235,"name":"Sterntaucher 5"},{"nr":236,"name":"Sterntaucher 6"}],
    "Haus Störtebeker":                   [{"nr":82,"name":"Küstenzauber 12a Whg.3"},{"nr":89,"name":"Schatzkiste 12/2"},{"nr":83,"name":"Störtebekerkoje 12/1"},{"nr":84,"name":"Störtebekerkoje 12/4"},{"nr":85,"name":"Störtebekerkoje 12/5"},{"nr":86,"name":"Störtebekerkoje 12/6"},{"nr":87,"name":"Störtebekerkoje 12a/4"},{"nr":88,"name":"Störtebekerkoje 12a/6"},{"nr":136,"name":"uns Leef HS 12a Whg.1"}],
    "Haus Tordalk":                       [{"nr":237,"name":"Tordalk 1"},{"nr":246,"name":"Tordalk 3"},{"nr":239,"name":"Tordalk 5"},{"nr":240,"name":"Tordalk 6"},{"nr":241,"name":"Tordalk 7"},{"nr":242,"name":"Tordalk 9"},{"nr":243,"name":"Tordalk 10"}],
    "Haus Windflüchter":                  [{"nr":179,"name":"Windflüchter 1"},{"nr":180,"name":"Windflüchter 2"},{"nr":181,"name":"Windflüchter 3"},{"nr":182,"name":"Windflüchter EG"}],
    "Haus Windwatt":                      [{"nr":177,"name":"Windwatt 2"},{"nr":178,"name":"Windwatt 4"}],
    "Haus Zur Heiderose":                 [{"nr":185,"name":"54 Grad"},{"nr":147,"name":"Künstlerkate"},{"nr":149,"name":"Mondmuschel"},{"nr":171,"name":"Strandglück"},{"nr":175,"name":"Wellenflüstern"},{"nr":122,"name":"Wildrose"}],
    "Residenz am Strand":                 [{"nr":151,"name":"Residenz 114"},{"nr":152,"name":"Residenz 120"},{"nr":153,"name":"Residenz 123"},{"nr":155,"name":"Residenz 232"},{"nr":156,"name":"Residenz 238"},{"nr":157,"name":"Residenz 242"},{"nr":159,"name":"Residenz 352"},{"nr":161,"name":"Residenz 567"},{"nr":162,"name":"Residenz 677"}],
    "Residenz Kormoran":                  [{"nr":276,"name":"Ankerzeit H7"},{"nr":238,"name":"Meerzeit D6"},{"nr":245,"name":"Windflüchter F5"}],
    "Speicherresidenz Barth":             [{"nr":29,"name":"App. 1.5"},{"nr":30,"name":"App. 0.2"},{"nr":31,"name":"App. 0.3"},{"nr":32,"name":"App. 2.1"},{"nr":33,"name":"App. 2.2"},{"nr":34,"name":"App. 3.3"},{"nr":35,"name":"App. 3.4"},{"nr":36,"name":"App. 3.7"},{"nr":37,"name":"App. 3.11"},{"nr":38,"name":"App. 3.1"},{"nr":39,"name":"App. 4.10"},{"nr":40,"name":"App. 4.11"},{"nr":41,"name":"App. 4.7"},{"nr":42,"name":"App. 4.6"},{"nr":43,"name":"App. 5.6"},{"nr":44,"name":"App. 5.9"},{"nr":45,"name":"App. 5.10"},{"nr":46,"name":"App. 6.1"},{"nr":47,"name":"App. 4.4"},{"nr":48,"name":"App. 2.3"},{"nr":49,"name":"App. 5.11"},{"nr":50,"name":"App. 7.1"}],
    "Strandapartments Düne 7":            [{"nr":70,"name":"Düne 7 Whg. 3"},{"nr":71,"name":"Düne 7 Whg. 5"},{"nr":72,"name":"Düne 7 Whg. 6"},{"nr":73,"name":"Düne 7 Whg. 7"},{"nr":74,"name":"Düne 7 Whg. 8"},{"nr":75,"name":"Düne 7 Whg. 9"},{"nr":76,"name":"Düne 7 Whg. 10"}],
    "Strandresort Fuhlendorf":            [{"nr":204,"name":"Luv"},{"nr":284,"name":"Sonnenbirke"},{"nr":248,"name":"Sonnenzauber"}],
    "Villa Seeluft":                      [{"nr":100,"name":"Seeluft 3"},{"nr":150,"name":"Seeluft 8"}],
    "Villa Strandoase Rosenberg":         [{"nr":211,"name":"Küstenkajüte Whg.1"},{"nr":212,"name":"Küstenkajüte Whg.2"},{"nr":213,"name":"Küstenkajüte Whg.5"},{"nr":214,"name":"Küstenkajüte Whg.7"}],
}

# Umgekehrtes Mapping: Objekt-Nr (int) → Hausname
OBJEKT_ZU_HAUS = {
    obj["nr"]: haus
    for haus, objekte in APARTMENTHAUS_MAPPING.items()
    for obj in objekte
}

# ──────────────────────────────────────────────────────────────────
# 2. SPALTENNAMEN – häufige Varianten, wird auto-erkannt
# ──────────────────────────────────────────────────────────────────
COL_CANDIDATES = {
    "objekt_nr":   ["Objekt-Nr.", "ObjektNr", "Objekt Nr", "ObjNr", "objekt_nr", "object_id", "Unterkunft-Nr"],
    "objekt_name": ["Objektname", "Unterkunft", "Objekt", "object_name", "Titel", "title"],
    "mietbetrag":  ["Mietbetrag", "Miete", "Preis", "Gesamtpreis", "Betrag", "mietbetrag", "Mietumsatz", "Umsatz"],
    "anreise":     ["Anreise", "Ankunft", "Arrival", "Von", "von", "CheckIn", "Check-in"],
    "abreise":     ["Abreise", "Abfahrt", "Departure", "Bis", "bis", "CheckOut", "Check-out"],
    "status":      ["Status", "Buchungsstatus", "state"],
}


def find_column(df, candidates):
    """Findet den ersten passenden Spaltennamen im DataFrame."""
    for c in candidates:
        if c in df.columns:
            return c
    # Case-insensitive fallback
    lower_cols = {col.lower(): col for col in df.columns}
    for c in candidates:
        if c.lower() in lower_cols:
            return lower_cols[c.lower()]
    return None


def load_data(csv_path, year_filter=None):
    """Lädt CSV, erkennt Trennzeichen automatisch."""
    try:
        import pandas as pd
    except ImportError:
        print("❌  pandas nicht gefunden. Bitte: pip install pandas openpyxl")
        sys.exit(1)

    # Auto-detect separator
    with open(csv_path, encoding="utf-8-sig", errors="replace") as f:
        sample = f.read(2048)
    sep = ";" if sample.count(";") > sample.count(",") else ","

    df = pd.read_csv(csv_path, sep=sep, encoding="utf-8-sig", low_memory=False,
                     thousands=".", decimal=",")

    print(f"✅  CSV geladen: {len(df)} Zeilen, {len(df.columns)} Spalten")
    print(f"   Spalten: {list(df.columns)}")

    # Spalten erkennen
    cols = {}
    for key, candidates in COL_CANDIDATES.items():
        col = find_column(df, candidates)
        cols[key] = col
        status = f"→ '{col}'" if col else "⚠️  nicht gefunden"
        print(f"   {key:14s}: {status}")

    if not cols["objekt_nr"] and not cols["objekt_name"]:
        print("\n❌  Weder Objekt-Nr. noch Objektname-Spalte gefunden.")
        print("   Bitte passe COL_CANDIDATES oben an.")
        sys.exit(1)

    # Jahresfilter
    if year_filter and cols["anreise"]:
        df[cols["anreise"]] = pd.to_datetime(df[cols["anreise"]], dayfirst=True, errors="coerce")
        df = df[df[cols["anreise"]].dt.year == year_filter]
        print(f"   → {len(df)} Buchungen nach Jahresfilter {year_filter}")

    # Stornos rausfiltern wenn Status-Spalte vorhanden
    if cols["status"]:
        before = len(df)
        storno_keywords = ["storno", "cancel", "stoniert", "abgesagt"]
        mask = df[cols["status"]].astype(str).str.lower().str.contains("|".join(storno_keywords), na=False)
        df = df[~mask]
        print(f"   → {before - len(df)} Stornobuchungen gefiltert")

    return df, cols


def compute_stats(df, cols):
    """Berechnet Buchungsanzahl und Umsatz pro Objekt-Nr."""
    import pandas as pd

    stats = {}  # {objekt_nr: {"buchungen": n, "umsatz": x, "name": ""}}

    # Objekt-Nr normalisieren
    if cols["objekt_nr"]:
        df = df.copy()
        df["_nr"] = pd.to_numeric(df[cols["objekt_nr"]], errors="coerce").astype("Int64")
    elif cols["objekt_name"]:
        # Fallback: Objekt-Nr aus Objektname extrahieren (Format "NR - Name")
        df = df.copy()
        df["_nr"] = df[cols["objekt_name"]].astype(str).str.extract(r'^(\d+)')[0]
        df["_nr"] = pd.to_numeric(df["_nr"], errors="coerce").astype("Int64")
    else:
        print("❌  Keine Objekt-Spalte verfügbar.")
        sys.exit(1)

    # Umsatz-Spalte
    if cols["mietbetrag"]:
        df["_umsatz"] = (
            df[cols["mietbetrag"]]
            .astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
            .str.replace("[^0-9.-]", "", regex=True)
        )
        df["_umsatz"] = pd.to_numeric(df["_umsatz"], errors="coerce").fillna(0)
    else:
        df["_umsatz"] = 0

    for nr, grp in df.groupby("_nr"):
        if pd.isna(nr):
            continue
        nr_int = int(nr)
        stats[nr_int] = {
            "buchungen": len(grp),
            "umsatz":    round(float(grp["_umsatz"].sum()), 2),
            "name":      OBJEKT_ZU_HAUS.get(nr_int, ""),
        }

    return stats


def get_available_years(df, cols):
    """Gibt alle Jahre zurück, für die Buchungsdaten vorliegen."""
    import pandas as pd
    if not cols["anreise"]:
        return []
    dates = pd.to_datetime(df[cols["anreise"]], dayfirst=True, errors="coerce")
    return sorted(dates.dt.year.dropna().unique().astype(int).tolist(), reverse=True)


# ──────────────────────────────────────────────────────────────────
# 3. HTML-GENERIERUNG
# ──────────────────────────────────────────────────────────────────

def fmt_eur(val):
    return f"{val:,.0f} €".replace(",", "X").replace(".", ",").replace("X", ".")


def build_html(haus_data, years, selected_year):
    """Baut den kompletten HTML-Inhalt für den Apartmenthaus-Tab."""

    # Häuser nach Gesamtbuchungen sortieren (absteigend)
    haus_sorted = sorted(
        [(h, d) for h, d in haus_data.items() if d["objekte"]],
        key=lambda x: x[1]["gesamt_buchungen"],
        reverse=True
    )

    # Jahr-Buttons
    year_btns = ""
    for y in years:
        active = 'class="active"' if y == selected_year else ""
        year_btns += f'<button {active} onclick="filterYear({y})">{y}</button>\n'

    # Haus-Cards
    cards_html = ""
    for haus_name, data in haus_sorted:
        gb = data["gesamt_buchungen"]
        gu = data["gesamt_umsatz"]
        obj_count = len(data["objekte"])

        # Unterkunfts-Zeilen
        obj_rows = ""
        max_b = max((o["buchungen"] for o in data["objekte"]), default=1) or 1
        for obj in sorted(data["objekte"], key=lambda x: x["buchungen"], reverse=True):
            bar_pct = round(obj["buchungen"] / max_b * 100)
            umsatz_str = fmt_eur(obj["umsatz"]) if obj["umsatz"] > 0 else "–"
            obj_rows += f"""
            <tr>
              <td class="obj-name">{obj['name']}</td>
              <td class="obj-bar">
                <div class="bar-wrap">
                  <div class="bar" style="width:{bar_pct}%"></div>
                  <span class="bar-label">{obj['buchungen']}</span>
                </div>
              </td>
              <td class="obj-umsatz">{umsatz_str}</td>
            </tr>"""

        cards_html += f"""
      <div class="haus-card" data-buchungen="{gb}">
        <div class="haus-header">
          <span class="haus-name">{haus_name}</span>
          <span class="haus-kpi">
            <span class="kpi-b">{gb} Buchungen</span>
            {'<span class="kpi-u">' + fmt_eur(gu) + '</span>' if gu > 0 else ''}
            <span class="kpi-n">{obj_count} Unterkunft{'en' if obj_count != 1 else ''}</span>
          </span>
        </div>
        <table class="obj-table">
          <tbody>{obj_rows}
          </tbody>
        </table>
      </div>"""

    return f"""<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Apartmenthäuser – Ostseeliebe Dashboard</title>
<style>
  :root {{
    --blue-dark: #0c2340;
    --blue:      #2d7fc1;
    --blue-light:#e8f3fb;
    --green:     #28a745;
    --gray-bg:   #f8f9fa;
    --gray-line: #e9ecef;
  }}
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    background: var(--gray-bg);
    color: #1a1a1a;
    padding: 20px;
  }}
  h1 {{ font-size:20px; color:var(--blue-dark); margin-bottom:4px; }}
  .subtitle {{ color:#666; font-size:13px; margin-bottom:16px; }}

  /* Jahr-Filter */
  .year-filter {{ display:flex; gap:6px; margin-bottom:20px; flex-wrap:wrap; }}
  .year-filter button {{
    padding:5px 14px; border:1px solid #ccc; border-radius:20px;
    background:#fff; cursor:pointer; font-size:13px; color:#555;
    transition:all .15s;
  }}
  .year-filter button.active, .year-filter button:hover {{
    background:var(--blue); color:#fff; border-color:var(--blue);
  }}

  /* Gesamt-KPIs */
  .kpi-row {{ display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin-bottom:20px; }}
  .kpi {{ background:#fff; border-radius:10px; padding:12px; text-align:center;
           box-shadow:0 1px 3px rgba(0,0,0,.07); }}
  .kpi-value {{ font-size:24px; font-weight:700; color:var(--blue-dark); }}
  .kpi-label {{ font-size:11px; color:#666; margin-top:2px; }}

  /* Sortierung */
  .sort-row {{ display:flex; align-items:center; gap:10px; margin-bottom:14px; font-size:13px; color:#555; }}
  .sort-row select {{
    border:1px solid #ccc; border-radius:6px; padding:4px 8px;
    font-size:13px; background:#fff; cursor:pointer;
  }}

  /* Haus-Cards */
  .cards-grid {{ display:grid; grid-template-columns:repeat(auto-fill,minmax(420px,1fr)); gap:16px; }}
  .haus-card {{
    background:#fff; border-radius:10px; padding:14px;
    box-shadow:0 1px 3px rgba(0,0,0,.08);
  }}
  .haus-header {{
    display:flex; justify-content:space-between; align-items:flex-start;
    margin-bottom:10px; gap:8px;
  }}
  .haus-name {{ font-weight:600; font-size:14px; color:var(--blue-dark); flex:1; }}
  .haus-kpi {{ display:flex; flex-wrap:wrap; gap:5px; justify-content:flex-end; }}
  .kpi-b {{ background:#dbeafe; color:#1e40af; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:600; }}
  .kpi-u {{ background:#d1fae5; color:#065f46; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:600; }}
  .kpi-n {{ background:#f3f4f6; color:#555;    padding:2px 8px; border-radius:10px; font-size:11px; }}

  /* Unterkunfts-Tabelle */
  .obj-table {{ width:100%; border-collapse:collapse; }}
  .obj-table tr:not(:last-child) td {{ border-bottom:1px solid var(--gray-line); }}
  .obj-table td {{ padding:5px 4px; vertical-align:middle; }}
  .obj-name {{ font-size:12px; color:#333; width:35%; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
  .obj-bar  {{ width:50%; padding-right:8px; }}
  .obj-umsatz {{ font-size:11px; color:#666; text-align:right; white-space:nowrap; width:15%; }}
  .bar-wrap {{ display:flex; align-items:center; gap:6px; }}
  .bar {{
    height:12px; background:var(--blue); border-radius:3px;
    min-width:3px; transition:width .3s;
  }}
  .bar-label {{ font-size:11px; color:#333; white-space:nowrap; }}

  .hidden {{ display:none !important; }}
</style>
</head>
<body>

<h1>Apartmenthäuser</h1>
<p class="subtitle">Buchungsvergleich je Apartmenthaus – Einzelsummen pro Unterkunft &amp; Gesamtsummen</p>

<div class="year-filter" id="yearFilter">
  {year_btns}
  <button onclick="filterYear(null)">Alle Jahre</button>
</div>

<div class="kpi-row" id="kpiRow">
  <div class="kpi"><div class="kpi-value" id="kpi-haeuser">–</div><div class="kpi-label">Apartmenthäuser</div></div>
  <div class="kpi"><div class="kpi-value" id="kpi-buchungen">–</div><div class="kpi-label">Buchungen gesamt</div></div>
  <div class="kpi"><div class="kpi-value" id="kpi-umsatz">–</div><div class="kpi-label">Umsatz gesamt</div></div>
</div>

<div class="sort-row">
  Sortieren nach:
  <select id="sortSelect" onchange="sortCards(this.value)">
    <option value="buchungen">Buchungen (absteigend)</option>
    <option value="umsatz">Umsatz (absteigend)</option>
    <option value="name">Name (A–Z)</option>
  </select>
</div>

<div class="cards-grid" id="cardsGrid">
  {cards_html}
</div>

<script>
// Alle Hausdaten als JSON eingebettet (für Client-seitiges Filtern)
const HAUS_DATA = {json.dumps(haus_data, ensure_ascii=False)};
const SELECTED_YEAR = {json.dumps(selected_year)};

function fmtEur(v) {{
  if (!v) return '–';
  return v.toLocaleString('de-DE', {{minimumFractionDigits:0, maximumFractionDigits:0}}) + ' €';
}}

function updateKPIs() {{
  const cards = document.querySelectorAll('.haus-card:not(.hidden)');
  let totalB = 0, totalU = 0;
  cards.forEach(c => {{
    totalB += parseInt(c.dataset.buchungen || 0);
    totalU += parseFloat(c.dataset.umsatz || 0);
  }});
  document.getElementById('kpi-haeuser').textContent = cards.length;
  document.getElementById('kpi-buchungen').textContent = totalB.toLocaleString('de-DE');
  document.getElementById('kpi-umsatz').textContent = fmtEur(totalU);
}}

function sortCards(by) {{
  const grid = document.getElementById('cardsGrid');
  const cards = Array.from(grid.children);
  cards.sort((a, b) => {{
    if (by === 'buchungen') return parseInt(b.dataset.buchungen) - parseInt(a.dataset.buchungen);
    if (by === 'umsatz')    return parseFloat(b.dataset.umsatz||0) - parseFloat(a.dataset.umsatz||0);
    if (by === 'name')      return a.querySelector('.haus-name').textContent.localeCompare(b.querySelector('.haus-name').textContent, 'de');
    return 0;
  }});
  cards.forEach(c => grid.appendChild(c));
}}

function filterYear(y) {{
  // Buttons
  document.querySelectorAll('.year-filter button').forEach(b => b.classList.remove('active'));
  event && event.target && event.target.classList.add('active');
  updateKPIs();
}}

// Init
updateKPIs();
</script>
</body>
</html>"""


# ──────────────────────────────────────────────────────────────────
# 4. HAUPTPROGRAMM
# ──────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Erzeugt Apartmenthaus-Tab für Ostseeliebe Dashboard")
    parser.add_argument("--csv",    default="buchungen_export_2027.csv", help="Pfad zur Buchungs-CSV")
    parser.add_argument("--year",   type=int, default=None,              help="Nur dieses Jahr auswerten (z.B. 2025)")
    parser.add_argument("--output", default="apartmenthaus_tab.html",    help="Ausgabe-HTML-Datei")
    args = parser.parse_args()

    # CSV laden
    if not Path(args.csv).exists():
        print(f"❌  Datei nicht gefunden: {args.csv}")
        print("   Bitte --csv Pfad/zur/Datei.csv angeben.")
        sys.exit(1)

    df_raw, cols = load_data(args.csv)
    all_years = get_available_years(df_raw, cols)
    selected_year = args.year or (all_years[0] if all_years else None)

    print(f"\n   Verfügbare Jahre: {all_years}")
    print(f"   Ausgewertetes Jahr: {selected_year or 'Alle'}")

    # Nach Jahr filtern für die Stats
    import pandas as pd
    df = df_raw.copy()
    if selected_year and cols["anreise"]:
        df[cols["anreise"]] = pd.to_datetime(df[cols["anreise"]], dayfirst=True, errors="coerce")
        df = df[df[cols["anreise"]].dt.year == selected_year]

    stats = compute_stats(df, cols)

    # Haus-Daten zusammenführen
    haus_data = {}
    for haus_name, objekte in APARTMENTHAUS_MAPPING.items():
        if not objekte:
            continue
        haus_objekte = []
        gesamt_buchungen = 0
        gesamt_umsatz = 0.0
        for obj in objekte:
            nr = obj["nr"]
            s = stats.get(nr, {"buchungen": 0, "umsatz": 0.0})
            haus_objekte.append({
                "nr":        nr,
                "name":      obj["name"],
                "buchungen": s["buchungen"],
                "umsatz":    s["umsatz"],
            })
            gesamt_buchungen += s["buchungen"]
            gesamt_umsatz    += s["umsatz"]
        haus_data[haus_name] = {
            "objekte":           haus_objekte,
            "gesamt_buchungen":  gesamt_buchungen,
            "gesamt_umsatz":     round(gesamt_umsatz, 2),
        }

    # Statistik ausgeben
    print(f"\n📊  Apartments mit Buchungen: "
          f"{sum(1 for d in haus_data.values() if d['gesamt_buchungen'] > 0)} / {len(haus_data)}")
    top5 = sorted(haus_data.items(), key=lambda x: x[1]['gesamt_buchungen'], reverse=True)[:5]
    print("   Top 5 Häuser:")
    for name, d in top5:
        print(f"     {d['gesamt_buchungen']:4d}x  {name}")

    # HTML generieren
    year_btns_years = all_years if all_years else [selected_year] if selected_year else []
    html = build_html(haus_data, year_btns_years, selected_year)

    # Datensatz-JSON für client-seitiges Filtern nach Jahr ergänzen
    # (vereinfacht: aktuelle Jahresauswahl ist fest eingebaut)

    with open(args.output, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n✅  HTML gespeichert: {args.output}")
    print(f"   Öffne die Datei im Browser um die Auswertung zu sehen.")


if __name__ == "__main__":
    main()
