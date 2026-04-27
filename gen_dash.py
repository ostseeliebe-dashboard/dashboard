#!/usr/bin/env python3
"""
Generates the Ostseeliebe Buchungs-Dashboard HTML from a CSV booking export.
No external dependencies required - uses only Python standard library.
"""

import csv
import json
import os
import sys
from collections import defaultdict
from datetime import datetime

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ---------------------------------------------------------------------------
# Paths - resolve the mapped session paths to real filesystem paths
# ---------------------------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# Default paths - can be overridden via command-line args
DEFAULT_CSV = os.path.expanduser("~/Claude_14.04.2026/buchungen_export_2027.csv")
DEFAULT_OUT = os.path.expanduser("~/Claude_14.04.2026/ostseeliebe-dashboard.html")


def parse_german_number(s):
    """Parse a German-formatted number (e.g. '1.234,56') to float."""
    if not s or not s.strip():
        return 0.0
    s = s.strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def format_german_number(n, decimals=2):
    """Format a number in German style: 1.234,56"""
    if decimals == 0:
        formatted = f"{n:,.0f}"
    else:
        formatted = f"{n:,.{decimals}f}"
    # Swap . and , for German formatting
    formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
    return formatted


def format_euro(n):
    """Format as Euro amount."""
    return f"{format_german_number(n)} \u20ac"


def read_objektstammdaten(xlsx_path):
    """Read property master data (Wohnfläche, Zimmer, Schlafplätze) from Objektstammdaten.xlsx."""
    if not HAS_OPENPYXL or not os.path.exists(xlsx_path):
        return None

    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb["Objektstammdaten"]

    properties = []
    orte_summary = defaultdict(lambda: {"count": 0, "wohnflaeche": 0, "zimmer": 0,
                                         "schlafzimmer": 0, "badzimmer": 0, "max_personen": 0,
                                         "sauna": 0, "hund": 0, "kamin": 0})
    totals = {"count": 0, "wohnflaeche": 0, "zimmer": 0, "schlafzimmer": 0,
              "badzimmer": 0, "max_personen": 0, "sauna": 0, "hund": 0, "kamin": 0}

    for row in ws.iter_rows(min_row=4, values_only=True):
        name = row[0]
        objnr = row[1]
        if not objnr or not name or not str(name).strip():
            continue  # Skip section headers

        ort = str(row[3]).strip() if row[3] else ""
        wf = row[4] if row[4] else 0
        zi = row[5] if row[5] else 0
        sz = row[6] if row[6] else 0
        bz = row[7] if row[7] else 0
        mp = row[8] if row[8] else 0
        sauna = 1 if row[11] and str(row[11]).strip() not in ("–", "-", "") else 0
        hund = 1 if row[12] and str(row[12]).strip() not in ("–", "-", "") else 0
        kamin = 1 if row[13] and str(row[13]).strip() not in ("–", "-", "") else 0

        properties.append({
            "name": str(name).strip(), "objnr": str(objnr).strip(),
            "ort": ort, "wohnflaeche": wf, "zimmer": zi,
            "schlafzimmer": sz, "badzimmer": bz, "max_personen": mp,
            "sauna": sauna, "hund": hund, "kamin": kamin,
        })

        totals["count"] += 1
        totals["wohnflaeche"] += wf
        totals["zimmer"] += zi
        totals["schlafzimmer"] += sz
        totals["badzimmer"] += bz
        totals["max_personen"] += mp
        totals["sauna"] += sauna
        totals["hund"] += hund
        totals["kamin"] += kamin

        o = orte_summary[ort]
        o["count"] += 1
        o["wohnflaeche"] += wf
        o["zimmer"] += zi
        o["schlafzimmer"] += sz
        o["badzimmer"] += bz
        o["max_personen"] += mp
        o["sauna"] += sauna
        o["hund"] += hund
        o["kamin"] += kamin

    wb.close()
    return {
        "properties": properties,
        "totals": totals,
        "orte": dict(orte_summary),
    }


def _derive_profiles(zusatz_names):
    """Derive travel profiles from a set of Zusatzkosten category names."""
    profiles = []
    if any("Kinderreisebett" in n for n in zusatz_names) or any("Kinderhochstuhl" in n for n in zusatz_names):
        profiles.append("Familie mit Kleinkind")
    if any("Hund" in n for n in zusatz_names):
        profiles.append("Urlaub mit Hund")
    if any("Mitreisende" in n and "Aufschlag" in n for n in zusatz_names):
        profiles.append("Größere Reisegruppe")
    if any("Sauna" in n or "Whirlpool" in n for n in zusatz_names):
        profiles.append("Wellness-Gäste")
    if any("Wallbox" in n for n in zusatz_names):
        profiles.append("E-Auto Reisende")
    if not profiles:
        profiles.append("Paare/Einzelreisende")
    return profiles


def load_cache(cache_path):
    """Load historical bookings from JSON cache and return as booking list."""
    import json
    with open(cache_path, "r", encoding="utf-8") as f:
        cache = json.load(f)
    bookings = []
    for year_str, year_bookings in cache["bookings"].items():
        for b in year_bookings:
            try:
                anreise = datetime.strptime(b["anreise"].strip(), "%d.%m.%Y")
            except ValueError:
                continue
            zusatz_names = set(b.get("zusatzkosten", {}).keys())
            bookings.append({
                "objekt_nr": b["objekt_nr"],
                "unterkunft": b["unterkunft"],
                "ort": b["ort"],
                "buchungsdatum": b["buchungsdatum"],
                "anreise": anreise,
                "abreise_str": b["abreise_str"],
                "naechte": b["naechte"],
                "vorgang": b["vorgang"],
                "vertriebskanal": b["vertriebskanal"],
                "reisepreis": b["reisepreis"],
                "provision_pct": b.get("provision_pct", ""),
                "miete_gesamt": b["miete_gesamt"],
                "miete_vermittler": b["miete_vermittler"],
                "miete_eigentuemer": b["miete_eigentuemer"],
                "zusatzkosten": b.get("zusatzkosten", {}),
                "profiles": _derive_profiles(zusatz_names),
            })
    return bookings, cache.get("years", [])


def read_bookings(csv_path, from_year=None):
    """Read and parse bookings from the CSV file, including Zusatzkosten.
    If from_year is set, only rows with Anreise >= from_year are read."""
    bookings = []
    zusatz_categories = []  # list of (col_index, category_name)

    with open(csv_path, "r", encoding="utf-8") as f:
        reader = csv.reader(f, delimiter=";")
        header1 = next(reader)
        header2 = next(reader)

        # Parse Zusatzkosten column pairs (Vermittler at odd idx, Eigentümer at idx+1)
        for i in range(17, len(header1), 2):
            name = header1[i].strip()
            if name:
                zusatz_categories.append((i, name))

        for row in reader:
            if len(row) < 17:
                continue
            status = row[9].strip()
            if status != "Buchung":
                continue
            try:
                anreise = datetime.strptime(row[4].strip(), "%d.%m.%Y")
            except ValueError:
                continue

            # Skip rows before from_year if cache is used
            if from_year and anreise.year < from_year:
                continue

            # Parse Zusatzkosten for this row
            zusatz = {}
            zusatz_names = set()
            for col_idx, cat_name in zusatz_categories:
                v = parse_german_number(row[col_idx]) if col_idx < len(row) else 0.0
                e = parse_german_number(row[col_idx + 1]) if col_idx + 1 < len(row) else 0.0
                if v != 0 or e != 0:
                    zusatz[cat_name] = {"vermittler": v, "eigentuemer": e}
                    zusatz_names.add(cat_name)

            bookings.append({
                "objekt_nr": row[0].strip(),
                "unterkunft": row[1].strip(),
                "ort": row[2].strip(),
                "buchungsdatum": row[3].strip(),
                "anreise": anreise,
                "abreise_str": row[5].strip(),
                "naechte": parse_german_number(row[6]),
                "vorgang": row[7].strip(),
                "vertriebskanal": row[8].strip(),
                "reisepreis": parse_german_number(row[11]),
                "provision_pct": row[13].strip() if len(row) > 13 else "",
                "miete_gesamt": parse_german_number(row[14]),
                "miete_vermittler": parse_german_number(row[15]),
                "miete_eigentuemer": parse_german_number(row[16]),
                "zusatzkosten": zusatz,
                "profiles": _derive_profiles(zusatz_names),
            })
    return bookings


def compute_data(bookings):
    """Compute all aggregated data for the dashboard."""
    # --- Per-year KPIs ---
    years = sorted(set(b["anreise"].year for b in bookings))
    kpis = {}
    for y in years:
        yb = [b for b in bookings if b["anreise"].year == y]
        kpis[y] = {
            "buchungen": len(yb),
            "naechte": sum(b["naechte"] for b in yb),
            "reisepreis": sum(b["reisepreis"] for b in yb),
            "miete_gesamt": sum(b["miete_gesamt"] for b in yb),
            "miete_vermittler": sum(b["miete_vermittler"] for b in yb),
            "miete_eigentuemer": sum(b["miete_eigentuemer"] for b in yb),
        }

    # --- Monthly nights per year (for line chart) ---
    monthly = defaultdict(lambda: defaultdict(float))
    for b in bookings:
        y = b["anreise"].year
        m = b["anreise"].month
        monthly[y][m] += b["naechte"]

    monthly_data = {}
    for y in years:
        monthly_data[y] = [monthly[y].get(m, 0) for m in range(1, 13)]

    # --- Monthly bookings count per year (for comparison table) ---
    monthly_count = defaultdict(lambda: defaultdict(int))
    for b in bookings:
        monthly_count[b["anreise"].year][b["anreise"].month] += 1

    monthly_count_data = {}
    for y in years:
        monthly_count_data[y] = [monthly_count[y].get(m, 0) for m in range(1, 13)]

    # --- Top 10 properties by booking count ---
    prop_counts = defaultdict(int)
    for b in bookings:
        prop_counts[b["unterkunft"]] += 1
    top_props = sorted(prop_counts.items(), key=lambda x: -x[1])[:10]

    # --- Reiseprofile / Zielgruppen ---
    profile_order = ["Paare/Einzelreisende", "Urlaub mit Hund", "Größere Reisegruppe",
                     "Familie mit Kleinkind", "Wellness-Gäste", "E-Auto Reisende"]
    profile_colors = {
        "Paare/Einzelreisende": "#0066cc",
        "Urlaub mit Hund": "#ff6b6b",
        "Größere Reisegruppe": "#ffa500",
        "Familie mit Kleinkind": "#4ecdc4",
        "Wellness-Gäste": "#aa96da",
        "E-Auto Reisende": "#95e1d3",
    }

    # Per year totals
    profiles_by_year = defaultdict(lambda: defaultdict(int))
    for b in bookings:
        y = b["anreise"].year
        for p in b["profiles"]:
            profiles_by_year[y][p] += 1

    # Per property: profile breakdown (all years combined + per year)
    prop_profiles = defaultdict(lambda: {"total": defaultdict(int), "per_year": defaultdict(lambda: defaultdict(int)), "buchungen": 0})
    for b in bookings:
        y = b["anreise"].year
        name = b["unterkunft"]
        prop_profiles[name]["buchungen"] += 1
        for p in b["profiles"]:
            prop_profiles[name]["total"][p] += 1
            prop_profiles[name]["per_year"][y][p] += 1

    # --- Sales channels ---
    channel_counts = defaultdict(int)
    for b in bookings:
        channel_counts[b["vertriebskanal"]] += 1
    channels_sorted = sorted(channel_counts.items(), key=lambda x: -x[1])

    # --- Sales channels per year (Ostseeliebe vs. Portalbuchungen grouping) ---
    OSTSEELIEBE_CHANNELS = {
        "Webseite", "Telefon", "Email", "E-Mail", "Büro", "Buchungsanfrage",
        "Online-Reservierung", "Whatsapp", "Eigentümer-Login", "Kundenangebot",
        "Umbuchungen", "Buchung durch Eigentümer", "Newsletter",
        "vorherige Agentur", "Wiederbucher",
    }
    channels_by_year = {}
    for y in years:
        yb = [b for b in bookings if b["anreise"].year == y]
        year_total = len(yb)
        raw = defaultdict(int)
        for b in yb:
            raw[b["vertriebskanal"]] += 1

        # Classify each channel
        ostl_sub = []
        portal_sub = []
        for k, v in sorted(raw.items(), key=lambda x: -x[1]):
            if not k:
                continue
            entry = {"name": k, "count": v, "pct": round(100 * v / year_total, 1) if year_total else 0}
            if k in OSTSEELIEBE_CHANNELS:
                ostl_sub.append(entry)
            else:
                portal_sub.append(entry)

        grouped = []
        ostl_total = sum(e["count"] for e in ostl_sub)
        if ostl_total > 0:
            grouped.append({"name": "Ostseeliebe", "count": ostl_total,
                            "pct": round(100 * ostl_total / year_total, 1) if year_total else 0,
                            "sub": ostl_sub})
        portal_total = sum(e["count"] for e in portal_sub)
        if portal_total > 0:
            grouped.append({"name": "Portalbuchungen", "count": portal_total,
                            "pct": round(100 * portal_total / year_total, 1) if year_total else 0,
                            "sub": portal_sub})
        channels_by_year[y] = {"total": year_total, "channels": grouped}

    # --- Locations ---
    ort_counts = defaultdict(int)
    for b in bookings:
        ort_counts[b["ort"]] += 1
    orte_sorted = sorted(ort_counts.items(), key=lambda x: -x[1])

    # --- Zusatzkosten aggregation ---
    # Structure: {category: {year: {vermittler, eigentuemer, count}}}
    zusatz_by_cat_year = defaultdict(lambda: defaultdict(lambda: {"vermittler": 0.0, "eigentuemer": 0.0, "count": 0}))
    for b in bookings:
        y = b["anreise"].year
        for cat_name, vals in b["zusatzkosten"].items():
            zusatz_by_cat_year[cat_name][y]["vermittler"] += vals["vermittler"]
            zusatz_by_cat_year[cat_name][y]["eigentuemer"] += vals["eigentuemer"]
            zusatz_by_cat_year[cat_name][y]["count"] += 1

    # Build sorted list by absolute total across all years
    zusatz_sorted = []
    for cat, year_data in zusatz_by_cat_year.items():
        total_v = sum(d["vermittler"] for d in year_data.values())
        total_e = sum(d["eigentuemer"] for d in year_data.values())
        total_c = sum(d["count"] for d in year_data.values())
        zusatz_sorted.append({
            "name": cat,
            "total": total_v + total_e,
            "vermittler": total_v,
            "eigentuemer": total_e,
            "count": total_c,
            "per_year": {y: {"vermittler": year_data[y]["vermittler"],
                             "eigentuemer": year_data[y]["eigentuemer"],
                             "gesamt": year_data[y]["vermittler"] + year_data[y]["eigentuemer"],
                             "count": year_data[y]["count"]} for y in years}
        })
    zusatz_sorted.sort(key=lambda x: abs(x["total"]), reverse=True)

    # Per-year Zusatzkosten totals (for Übersicht summary)
    zusatz_year_totals = {}
    for y in years:
        v_sum = sum(z["per_year"][y]["vermittler"] for z in zusatz_sorted)
        e_sum = sum(z["per_year"][y]["eigentuemer"] for z in zusatz_sorted)
        zusatz_year_totals[y] = {"vermittler": v_sum, "eigentuemer": e_sum, "gesamt": v_sum + e_sum}

    # --- Per-property detailed data ---
    property_data = {}  # {name: {...}}
    prop_bookings = defaultdict(list)
    for b in bookings:
        prop_bookings[b["unterkunft"]].append(b)

    for prop_name, pbs in sorted(prop_bookings.items()):
        ort = pbs[0]["ort"] if pbs else ""
        p_years = {}
        for y in years:
            ybs = [b for b in pbs if b["anreise"].year == y]
            if not ybs:
                p_years[y] = {
                    "buchungen": 0, "naechte": 0, "reisepreis": 0, "miete_gesamt": 0,
                    "miete_vermittler": 0, "miete_eigentuemer": 0, "provision_pct": 0,
                    "avg_preis_nacht": 0, "avg_aufenthalt": 0,
                    "belegung_pct": 0, "channels": {}, "zusatzkosten": {}
                }
                continue
            n_buch = len(ybs)
            n_naechte = sum(b["naechte"] for b in ybs)
            rp = sum(b["reisepreis"] for b in ybs)
            mg = sum(b["miete_gesamt"] for b in ybs)
            mv = sum(b["miete_vermittler"] for b in ybs)
            me = sum(b["miete_eigentuemer"] for b in ybs)
            prov_pct = round(mv / mg * 100, 1) if mg > 0 else 0
            avg_pn = rp / n_naechte if n_naechte > 0 else 0
            avg_auf = n_naechte / n_buch if n_buch > 0 else 0
            beleg = round(n_naechte / 365 * 100, 1) if n_naechte > 0 else 0

            # Channels for this property/year
            ch = defaultdict(int)
            for b in ybs:
                ch[b["vertriebskanal"]] += 1
            ch_sorted = sorted(ch.items(), key=lambda x: -x[1])

            # Zusatzkosten for this property/year
            zk = defaultdict(lambda: {"vermittler": 0.0, "eigentuemer": 0.0, "count": 0})
            for b in ybs:
                for cat, vals in b["zusatzkosten"].items():
                    zk[cat]["vermittler"] += vals["vermittler"]
                    zk[cat]["eigentuemer"] += vals["eigentuemer"]
                    zk[cat]["count"] += 1
            zk_list = [{"name": k, "vermittler": v["vermittler"], "eigentuemer": v["eigentuemer"],
                         "gesamt": v["vermittler"] + v["eigentuemer"], "count": v["count"]}
                        for k, v in zk.items() if v["vermittler"] != 0 or v["eigentuemer"] != 0]
            zk_list.sort(key=lambda x: abs(x["gesamt"]), reverse=True)

            p_years[y] = {
                "buchungen": n_buch, "naechte": n_naechte, "reisepreis": rp,
                "miete_gesamt": mg, "miete_vermittler": mv, "miete_eigentuemer": me,
                "provision_pct": prov_pct,
                "avg_preis_nacht": round(avg_pn, 2), "avg_aufenthalt": round(avg_auf, 1),
                "belegung_pct": beleg,
                "channels": ch_sorted, "zusatzkosten": zk_list
            }

        property_data[prop_name] = {"ort": ort, "years": p_years}

    # --- Provision summary per property (for Provisionen tab) ---
    provision_by_prop = {}
    for prop_name, pd in sorted(property_data.items()):
        prop_prov = {}
        for y in years:
            yd = pd["years"].get(y, {})
            if yd.get("buchungen", 0) == 0:
                continue
            zk_verm = sum(z["vermittler"] for z in yd.get("zusatzkosten", []))
            prop_prov[y] = {
                "buchungen": yd["buchungen"],
                "miete_gesamt": yd["miete_gesamt"],
                "miete_vermittler": yd["miete_vermittler"],
                "zusatz_vermittler": round(zk_verm, 2),
                "provision_gesamt": round(yd["miete_vermittler"] + zk_verm, 2),
                "provision_pct": yd["provision_pct"],
            }
        if prop_prov:
            provision_by_prop[prop_name] = {"ort": pd["ort"], "years": prop_prov}

    return {
        "years": years,
        "kpis": kpis,
        "monthly_data": monthly_data,
        "monthly_count_data": monthly_count_data,
        "top_props": top_props,
        "channels": channels_sorted,
        "channels_by_year": channels_by_year,
        "orte": orte_sorted,
        "zusatz_sorted": zusatz_sorted,
        "zusatz_year_totals": zusatz_year_totals,
        "property_data": property_data,
        "provision_by_prop": provision_by_prop,
        "profile_order": profile_order,
        "profile_colors": profile_colors,
        "profiles_by_year": profiles_by_year,
        "prop_profiles": prop_profiles,
    }


def generate_html(data):
    """Generate the complete dashboard HTML."""
    years = data["years"]
    # Only display years >= 2026, current year first
    current_year = datetime.now().year
    display_years = sorted([y for y in years if y >= 2026], reverse=True)
    kpis = data["kpis"]
    monthly_data = data["monthly_data"]
    monthly_count_data = data["monthly_count_data"]
    top_props = data["top_props"]
    channels = data["channels"]
    channels_by_year = data["channels_by_year"]
    orte = data["orte"]
    zusatz_sorted = data["zusatz_sorted"]
    zusatz_year_totals = data["zusatz_year_totals"]
    property_data = data["property_data"]
    provision_by_prop = data["provision_by_prop"]
    profile_order = data["profile_order"]
    profile_colors = data["profile_colors"]
    profiles_by_year = data["profiles_by_year"]
    prop_profiles = data["prop_profiles"]
    stammdaten = data.get("stammdaten")

    update_date = datetime.now().strftime("%d.%m.%Y %H:%M")
    month_names = [
        "Jan", "Feb", "Mär", "Apr", "Mai", "Jun",
        "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"
    ]
    colors = [
        "#0066cc", "#00aaaa", "#ff6b6b", "#ffa500", "#4ecdc4",
        "#95e1d3", "#f38181", "#aa96da", "#fcbad3", "#a8dadc"
    ]

    # --- Build Bestandskennzahlen HTML (from Objektstammdaten) ---
    bestand_html = ""
    if stammdaten:
        t = stammdaten["totals"]
        orte_stamm = stammdaten["orte"]

        # Ort rows for the table
        ort_rows = ""
        for ort_name in sorted(orte_stamm.keys()):
            o = orte_stamm[ort_name]
            ort_rows += f'''<tr>
                <td>{ort_name}</td>
                <td class="zk-num">{o["count"]}</td>
                <td class="zk-num">{format_german_number(o["wohnflaeche"], 0)} m²</td>
                <td class="zk-num">{format_german_number(o["zimmer"], 0)}</td>
                <td class="zk-num">{o["schlafzimmer"]}</td>
                <td class="zk-num">{o["badzimmer"]}</td>
                <td class="zk-num">{o["max_personen"]}</td>
                <td class="zk-num">{o["sauna"]}</td>
                <td class="zk-num">{o["hund"]}</td>
                <td class="zk-num">{o["kamin"]}</td>
            </tr>'''

        bestand_html = f'''
        <div class="year-section">
            <h3 class="year-title">Bestand – Buchbare Unterkünfte (Online)</h3>
            <div class="kpi-grid">
                <div class="kpi-card" style="border-left:4px solid #28a745">
                    <div class="kpi-label">Unterkünfte</div>
                    <div class="kpi-value" style="color:#28a745">{t["count"]}</div>
                </div>
                <div class="kpi-card" style="border-left:4px solid #28a745">
                    <div class="kpi-label">Wohnfläche gesamt</div>
                    <div class="kpi-value" style="color:#28a745">{format_german_number(t["wohnflaeche"], 0)} m²</div>
                </div>
                <div class="kpi-card" style="border-left:4px solid #28a745">
                    <div class="kpi-label">Schlafzimmer</div>
                    <div class="kpi-value" style="color:#28a745">{t["schlafzimmer"]}</div>
                </div>
                <div class="kpi-card" style="border-left:4px solid #28a745">
                    <div class="kpi-label">Schlafplätze (max.)</div>
                    <div class="kpi-value" style="color:#28a745">{t["max_personen"]}</div>
                </div>
                <div class="kpi-card" style="border-left:4px solid #28a745">
                    <div class="kpi-label">Zimmer gesamt</div>
                    <div class="kpi-value" style="color:#28a745">{format_german_number(t["zimmer"], 0)}</div>
                </div>
            </div>
            <div class="kpi-grid" style="grid-template-columns:repeat(3,1fr);margin-top:8px">
                <div class="kpi-card" style="border-left:4px solid #aa96da">
                    <div class="kpi-label">mit Sauna</div>
                    <div class="kpi-value" style="color:#aa96da">{t["sauna"]}</div>
                </div>
                <div class="kpi-card" style="border-left:4px solid #ff6b6b">
                    <div class="kpi-label">Hund erlaubt</div>
                    <div class="kpi-value" style="color:#ff6b6b">{t["hund"]}</div>
                </div>
                <div class="kpi-card" style="border-left:4px solid #ffa500">
                    <div class="kpi-label">mit Kamin</div>
                    <div class="kpi-value" style="color:#ffa500">{t["kamin"]}</div>
                </div>
            </div>
            <div class="zusatz-summary" style="margin-top:12px">
                <table class="zusatz-table">
                    <thead>
                        <tr>
                            <th>Ort</th>
                            <th class="zk-num">Objekte</th>
                            <th class="zk-num">Wohnfl.</th>
                            <th class="zk-num">Zimmer</th>
                            <th class="zk-num">Schlafz.</th>
                            <th class="zk-num">Bäder</th>
                            <th class="zk-num">Schlafpl.</th>
                            <th class="zk-num">Sauna</th>
                            <th class="zk-num">Hund</th>
                            <th class="zk-num">Kamin</th>
                        </tr>
                    </thead>
                    <tbody>
                        {ort_rows}
                        <tr class="zk-total">
                            <td><strong>Gesamt</strong></td>
                            <td class="zk-num"><strong>{t["count"]}</strong></td>
                            <td class="zk-num"><strong>{format_german_number(t["wohnflaeche"], 0)} m²</strong></td>
                            <td class="zk-num"><strong>{format_german_number(t["zimmer"], 0)}</strong></td>
                            <td class="zk-num"><strong>{t["schlafzimmer"]}</strong></td>
                            <td class="zk-num"><strong>{t["badzimmer"]}</strong></td>
                            <td class="zk-num"><strong>{t["max_personen"]}</strong></td>
                            <td class="zk-num"><strong>{t["sauna"]}</strong></td>
                            <td class="zk-num"><strong>{t["hund"]}</strong></td>
                            <td class="zk-num"><strong>{t["kamin"]}</strong></td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>'''

    # --- Build KPI cards HTML (with Zusatzkosten summary) ---
    # Determine top Zusatzkosten categories (those with abs total > 100)
    top_zusatz = [z for z in zusatz_sorted if abs(z["total"]) > 100]

    kpi_html_parts = []
    for y in display_years:
        k = kpis[y]
        zt = zusatz_year_totals[y]

        # Build Zusatzkosten mini-table for this year
        zusatz_rows = ""
        for z in top_zusatz:
            zy = z["per_year"][y]
            if zy["gesamt"] == 0 and zy["vermittler"] == 0 and zy["eigentuemer"] == 0:
                continue
            zusatz_rows += f'''<tr>
                <td>{z["name"]}</td>
                <td class="zk-num">{format_euro(zy["gesamt"])}</td>
                <td class="zk-num zk-sub">{format_euro(zy["vermittler"])}</td>
                <td class="zk-num zk-sub">{format_euro(zy["eigentuemer"])}</td>
            </tr>'''

        zusatz_total_row = f'''<tr class="zk-total">
            <td><strong>Summe Zusatzkosten</strong></td>
            <td class="zk-num"><strong>{format_euro(zt["gesamt"])}</strong></td>
            <td class="zk-num zk-sub"><strong>{format_euro(zt["vermittler"])}</strong></td>
            <td class="zk-num zk-sub"><strong>{format_euro(zt["eigentuemer"])}</strong></td>
        </tr>'''

        kpi_html_parts.append(f'''
        <div class="year-section">
            <h3 class="year-title">{y}</h3>
            <div class="kpi-grid">
                <div class="kpi-card">
                    <div class="kpi-label">Buchungen</div>
                    <div class="kpi-value">{k["buchungen"]}</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">N\u00e4chte</div>
                    <div class="kpi-value">{format_german_number(k["naechte"], 0)}</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Reisepreis</div>
                    <div class="kpi-value">{format_euro(k["reisepreis"])}</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Miete gesamt</div>
                    <div class="kpi-value">{format_euro(k["miete_gesamt"])}</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Miete Eigent\u00fcmer</div>
                    <div class="kpi-value">{format_euro(k["miete_eigentuemer"])}</div>
                </div>
            </div>
            <div class="zusatz-summary">
                <table class="zusatz-table">
                    <thead>
                        <tr>
                            <th>Zusatzkosten</th>
                            <th class="zk-num">Gesamt</th>
                            <th class="zk-num zk-sub">Vermittler</th>
                            <th class="zk-num zk-sub">Eigent\u00fcmer</th>
                        </tr>
                    </thead>
                    <tbody>
                        {zusatz_rows}
                        {zusatz_total_row}
                    </tbody>
                </table>
            </div>
        </div>''')
    kpi_html = "\n".join(kpi_html_parts)

    # --- Build comparison table HTML ---
    table_header = "<tr><th>Monat</th>" + "".join(f"<th>{y}</th>" for y in display_years) + "</tr>"
    table_rows = []
    for m_idx, m_name in enumerate(month_names):
        row = f"<tr><td>{m_name}</td>"
        for y in display_years:
            val = monthly_count_data[y][m_idx]
            row += f"<td>{val}</td>"
        row += "</tr>"
        table_rows.append(row)
    # Totals row
    totals_row = "<tr class='total-row'><td><strong>Gesamt</strong></td>"
    for y in display_years:
        totals_row += f"<td><strong>{sum(monthly_count_data[y])}</strong></td>"
    totals_row += "</tr>"
    table_rows.append(totals_row)
    comparison_table = f"<table class='comparison-table'><thead>{table_header}</thead><tbody>{''.join(table_rows)}</tbody></table>"

    # --- Build Zusatzkosten detail tab HTML ---
    # Full detail table with all categories and years (display_years only)
    zusatz_detail_header = "<tr><th>Kategorie</th><th class='zk-num'>Anzahl</th>"
    for y in display_years:
        zusatz_detail_header += f"<th class='zk-num'>{y} Ges.</th><th class='zk-num zk-sub'>{y} Verm.</th><th class='zk-num zk-sub'>{y} Eig.</th>"
    zusatz_detail_header += "<th class='zk-num'>Gesamt</th><th class='zk-num zk-sub'>Vermittler</th><th class='zk-num zk-sub'>Eigent\u00fcmer</th></tr>"

    zusatz_detail_rows = ""
    for z in zusatz_sorted:
        zusatz_detail_rows += f"<tr><td>{z['name']}</td><td class='zk-num'>{z['count']}</td>"
        for y in display_years:
            zy = z["per_year"].get(y, {"gesamt": 0, "vermittler": 0, "eigentuemer": 0})
            zusatz_detail_rows += f"<td class='zk-num'>{format_euro(zy['gesamt'])}</td>"
            zusatz_detail_rows += f"<td class='zk-num zk-sub'>{format_euro(zy['vermittler'])}</td>"
            zusatz_detail_rows += f"<td class='zk-num zk-sub'>{format_euro(zy['eigentuemer'])}</td>"
        zusatz_detail_rows += f"<td class='zk-num'><strong>{format_euro(z['total'])}</strong></td>"
        zusatz_detail_rows += f"<td class='zk-num zk-sub'>{format_euro(z['vermittler'])}</td>"
        zusatz_detail_rows += f"<td class='zk-num zk-sub'>{format_euro(z['eigentuemer'])}</td></tr>"

    # Totals row
    zusatz_detail_rows += "<tr class='zk-total'><td><strong>SUMME</strong></td><td></td>"
    for y in display_years:
        zt = zusatz_year_totals[y]
        zusatz_detail_rows += f"<td class='zk-num'><strong>{format_euro(zt['gesamt'])}</strong></td>"
        zusatz_detail_rows += f"<td class='zk-num zk-sub'><strong>{format_euro(zt['vermittler'])}</strong></td>"
        zusatz_detail_rows += f"<td class='zk-num zk-sub'><strong>{format_euro(zt['eigentuemer'])}</strong></td>"
    grand_v = sum(zusatz_year_totals[y]["vermittler"] for y in display_years)
    grand_e = sum(zusatz_year_totals[y]["eigentuemer"] for y in display_years)
    zusatz_detail_rows += f"<td class='zk-num'><strong>{format_euro(grand_v + grand_e)}</strong></td>"
    zusatz_detail_rows += f"<td class='zk-num zk-sub'><strong>{format_euro(grand_v)}</strong></td>"
    zusatz_detail_rows += f"<td class='zk-num zk-sub'><strong>{format_euro(grand_e)}</strong></td></tr>"

    zusatz_detail_table = f"""<div style="overflow-x:auto;"><table class="zusatz-detail-table">
        <thead>{zusatz_detail_header}</thead>
        <tbody>{zusatz_detail_rows}</tbody>
    </table></div>"""

    # Chart data: top 10 Zusatzkosten by absolute total (bar chart)
    top10_zusatz = zusatz_sorted[:10]
    zusatz_chart_labels = json.dumps([z["name"] for z in top10_zusatz])
    zusatz_chart_vermittler = json.dumps([round(z["vermittler"], 2) for z in top10_zusatz])
    zusatz_chart_eigentuemer = json.dumps([round(z["eigentuemer"], 2) for z in top10_zusatz])

    # Stacked bar chart per year for top 5 categories
    top5_zusatz = zusatz_sorted[:5]
    zusatz_yearly_labels = json.dumps([z["name"] for z in top5_zusatz])
    zusatz_yearly_datasets = []
    for i, y in enumerate(display_years):
        color = colors[i % len(colors)]
        zusatz_yearly_datasets.append({
            "label": str(y),
            "data": [round(z["per_year"][y]["gesamt"], 2) for z in top5_zusatz],
            "backgroundColor": color,
            "borderRadius": 4,
        })
    json_zusatz_yearly_datasets = json.dumps(zusatz_yearly_datasets)

    # --- Build Provisionen tab HTML ---
    prov_tab_parts = []
    for y in display_years:
        # Provision KPIs for this year
        k = kpis[y]
        mv_total = k["miete_vermittler"]
        mg_total = k["miete_gesamt"]
        prov_rate_avg = round(mv_total / mg_total * 100, 1) if mg_total > 0 else 0
        # Sum Zusatzkosten Vermittler for this year
        zk_verm_year = zusatz_year_totals[y]["vermittler"]
        prov_incl_zk = mv_total + zk_verm_year

        prov_tab_parts.append(f'''
        <div class="prop-section">
            <h4 style="color:#0066cc;border-bottom:2px solid #e0e0e0;padding-bottom:6px;font-size:18px;">{y}</h4>
            <div class="prop-detail-grid" style="margin-bottom:16px;">
                <div class="prop-kpi"><div class="pk-label">Miete Vermittler</div><div class="pk-value">{format_euro(mv_total)}</div></div>
                <div class="prop-kpi"><div class="pk-label">+ Zusatzk. Vermittler</div><div class="pk-value">{format_euro(zk_verm_year)}</div></div>
                <div class="prop-kpi"><div class="pk-label">Provision gesamt</div><div class="pk-value green">{format_euro(prov_incl_zk)}</div></div>
                <div class="prop-kpi"><div class="pk-label">\u00d8 Provisionssatz</div><div class="pk-value">{format_german_number(prov_rate_avg, 1)} %</div></div>
            </div>
            <div style="overflow-x:auto;">
            <table class="prov-table">
                <thead>
                    <tr>
                        <th>Unterkunft</th>
                        <th>Ort</th>
                        <th class="num">Buchungen</th>
                        <th class="num">Miete ges.</th>
                        <th class="num">Miete Verm.</th>
                        <th class="num">Zusatzk. Verm.</th>
                        <th class="num">Provision ges.</th>
                        <th class="num">Prov.satz</th>
                    </tr>
                </thead>
                <tbody>''')
        # Rows per property for this year, sorted by provision desc
        year_rows = []
        sum_mg = sum_mv = sum_zv = sum_pg = 0
        for pname, pprov in sorted(provision_by_prop.items()):
            if y not in pprov["years"]:
                continue
            py = pprov["years"][y]
            year_rows.append((pname, pprov["ort"], py))
            sum_mg += py["miete_gesamt"]
            sum_mv += py["miete_vermittler"]
            sum_zv += py["zusatz_vermittler"]
            sum_pg += py["provision_gesamt"]

        year_rows.sort(key=lambda x: -x[2]["provision_gesamt"])
        for pname, ort, py in year_rows:
            prov_tab_parts.append(f'''
                    <tr>
                        <td>{pname}</td>
                        <td>{ort}</td>
                        <td class="num">{py["buchungen"]}</td>
                        <td class="num">{format_euro(py["miete_gesamt"])}</td>
                        <td class="num">{format_euro(py["miete_vermittler"])}</td>
                        <td class="num">{format_euro(py["zusatz_vermittler"])}</td>
                        <td class="num"><strong>{format_euro(py["provision_gesamt"])}</strong></td>
                        <td class="num">{format_german_number(py["provision_pct"], 1)} %</td>
                    </tr>''')

        sum_pct = round(sum_mv / sum_mg * 100, 1) if sum_mg > 0 else 0
        prov_tab_parts.append(f'''
                    <tr class="prov-total">
                        <td colspan="2"><strong>SUMME {y}</strong></td>
                        <td></td>
                        <td class="num"><strong>{format_euro(sum_mg)}</strong></td>
                        <td class="num"><strong>{format_euro(sum_mv)}</strong></td>
                        <td class="num"><strong>{format_euro(sum_zv)}</strong></td>
                        <td class="num"><strong>{format_euro(sum_pg)}</strong></td>
                        <td class="num"><strong>{format_german_number(sum_pct, 1)} %</strong></td>
                    </tr>
                </tbody>
            </table>
            </div>
        </div>''')

    prov_tab_html = "\n".join(prov_tab_parts)

    # --- Build Reiseprofile tab HTML ---
    # Stacked bar chart data: profiles per year
    profile_chart_datasets = []
    for p in profile_order:
        vals = [profiles_by_year[y].get(p, 0) for y in display_years]
        if sum(vals) == 0:
            continue
        profile_chart_datasets.append({
            "label": p,
            "data": vals,
            "backgroundColor": profile_colors.get(p, "#ccc"),
            "borderRadius": 4,
        })
    json_profile_datasets = json.dumps(profile_chart_datasets)

    # Property table: sorted by dog percentage (most interesting for targeting)
    profile_table_rows = ""
    prop_profile_list = []
    for pname, pd in prop_profiles.items():
        if pd["buchungen"] < 5:
            continue
        total_b = pd["buchungen"]
        hund_b = pd["total"].get("Urlaub mit Hund", 0)
        familie_b = pd["total"].get("Familie mit Kleinkind", 0)
        gruppe_b = pd["total"].get("Größere Reisegruppe", 0)
        wellness_b = pd["total"].get("Wellness-Gäste", 0)
        eauto_b = pd["total"].get("E-Auto Reisende", 0)
        paar_b = pd["total"].get("Paare/Einzelreisende", 0)
        prop_profile_list.append({
            "name": pname,
            "total": total_b,
            "hund": hund_b, "hund_pct": round(hund_b / total_b * 100, 1),
            "familie": familie_b, "familie_pct": round(familie_b / total_b * 100, 1),
            "gruppe": gruppe_b, "gruppe_pct": round(gruppe_b / total_b * 100, 1),
            "wellness": wellness_b, "wellness_pct": round(wellness_b / total_b * 100, 1),
            "eauto": eauto_b, "eauto_pct": round(eauto_b / total_b * 100, 1),
            "paar": paar_b, "paar_pct": round(paar_b / total_b * 100, 1),
        })
    prop_profile_list.sort(key=lambda x: -x["hund_pct"])

    for pp in prop_profile_list:
        # Visual bar for profile distribution
        bar_html = ""
        for label, key, color in [
            ("Paar", "paar", "#0066cc"), ("Hund", "hund", "#ff6b6b"),
            ("Gruppe", "gruppe", "#ffa500"), ("Familie", "familie", "#4ecdc4"),
            ("Wellness", "wellness", "#aa96da"), ("E-Auto", "eauto", "#95e1d3")
        ]:
            pct = pp[f"{key}_pct"]
            if pct > 0:
                bar_html += f'<div style="width:{max(pct, 2)}%;background:{color};height:100%;display:inline-block;" title="{label}: {pp[key]} ({pct}%)"></div>'

        profile_table_rows += f'''<tr>
            <td>{pp["name"]}</td>
            <td class="num">{pp["total"]}</td>
            <td class="num" style="color:#ff6b6b;font-weight:600;">{pp["hund"]} <small>({pp["hund_pct"]}%)</small></td>
            <td class="num" style="color:#4ecdc4;">{pp["familie"]} <small>({pp["familie_pct"]}%)</small></td>
            <td class="num" style="color:#ffa500;">{pp["gruppe"]} <small>({pp["gruppe_pct"]}%)</small></td>
            <td class="num" style="color:#aa96da;">{pp["wellness"]} <small>({pp["wellness_pct"]}%)</small></td>
            <td class="num" style="color:#0066cc;">{pp["paar"]} <small>({pp["paar_pct"]}%)</small></td>
            <td style="min-width:200px;"><div style="display:flex;height:16px;border-radius:4px;overflow:hidden;">{bar_html}</div></td>
        </tr>'''

    # Top Hund properties chart (top 15)
    top_hund = sorted(prop_profile_list, key=lambda x: -x["hund_pct"])[:15]
    json_hund_labels = json.dumps([p["name"] for p in top_hund])
    json_hund_pcts = json.dumps([p["hund_pct"] for p in top_hund])
    json_hund_counts = json.dumps([p["hund"] for p in top_hund])

    # --- Build property data JSON for embedded JS ---
    prop_json_data = {}
    for pname, pdata in property_data.items():
        pj = {"ort": pdata["ort"], "years": {}}
        for y, yd in pdata["years"].items():
            pj["years"][y] = {
                "buchungen": yd["buchungen"],
                "naechte": yd["naechte"],
                "reisepreis": round(yd["reisepreis"], 2),
                "miete_gesamt": round(yd["miete_gesamt"], 2),
                "miete_vermittler": round(yd["miete_vermittler"], 2),
                "miete_eigentuemer": round(yd["miete_eigentuemer"], 2),
                "provision_pct": yd["provision_pct"],
                "avg_preis_nacht": yd["avg_preis_nacht"],
                "avg_aufenthalt": yd["avg_aufenthalt"],
                "belegung_pct": yd["belegung_pct"],
                "channels": [list(c) for c in yd["channels"]] if isinstance(yd["channels"], list) and yd["channels"] and isinstance(yd["channels"][0], tuple) else yd["channels"],
                "zusatzkosten": yd["zusatzkosten"]
            }
        prop_json_data[pname] = pj
    json_property_data = json.dumps(prop_json_data, ensure_ascii=False)

    property_options = "\n".join(
        f'<option value="{pname}">{pname} ({pdata["ort"]})</option>'
        for pname, pdata in sorted(property_data.items())
    )

    # --- JSON data for charts ---
    json_years = json.dumps(display_years)
    json_month_names = json.dumps(month_names)
    json_colors = json.dumps(colors)

    # Monthly line chart datasets
    line_datasets = []
    for i, y in enumerate(display_years):
        color = colors[i % len(colors)]
        line_datasets.append({
            "label": str(y),
            "data": [round(v, 1) for v in monthly_data[y]],
            "borderColor": color,
            "backgroundColor": color + "33",
            "tension": 0.3,
            "fill": False,
            "borderWidth": 2,
            "pointRadius": 4,
        })
    json_line_datasets = json.dumps(line_datasets)

    # Top properties bar chart
    prop_labels = json.dumps([p[0] for p in top_props])
    prop_values = json.dumps([p[1] for p in top_props])

    # Channels doughnut (legacy, kept for overall chart)
    channel_labels = json.dumps([c[0] for c in channels])
    channel_values = json.dumps([c[1] for c in channels])
    channel_colors = json.dumps([colors[i % len(colors)] for i in range(len(channels))])

    # --- Build Vertriebskanäle per-year HTML ---
    channel_year_html_parts = []
    channel_chart_js_parts = []
    for idx, y in enumerate(display_years):
        cy = channels_by_year[y]
        total = cy["total"]
        chs = cy["channels"]

        # Table rows with expandable groups (Ostseeliebe + Portalbuchungen)
        table_rows = ""
        for gi, ch in enumerate(chs):
            # Safe CSS class from group name
            grp_cls = f"ch-g{gi}-{y}"
            bg_color = "#f0f7ff" if gi == 0 else "#fff7f0"
            sub_bg = "#f8fbff" if gi == 0 else "#fffbf8"
            table_rows += f'''<tr class="ch-group" onclick="this.parentElement.querySelectorAll('.{grp_cls}').forEach(function(r){{r.style.display=r.style.display==='none'?'table-row':'none'}});" style="cursor:pointer;background:{bg_color}">
                <td><strong>&#9654; {ch["name"]}</strong></td>
                <td class="zk-num"><strong>{format_german_number(ch["count"], 0)}</strong></td>
                <td class="zk-num"><strong>{ch["pct"]} %</strong></td>
            </tr>'''
            for sub in ch["sub"]:
                table_rows += f'''<tr class="{grp_cls}" style="display:none;background:{sub_bg}">
                    <td style="padding-left:28px">{sub["name"]}</td>
                    <td class="zk-num">{format_german_number(sub["count"], 0)}</td>
                    <td class="zk-num">{sub["pct"]} %</td>
                </tr>'''
        table_rows += f'''<tr class="zk-total">
            <td><strong>Gesamt</strong></td>
            <td class="zk-num"><strong>{format_german_number(total, 0)}</strong></td>
            <td class="zk-num"><strong>100 %</strong></td>
        </tr>'''

        # Chart data for this year (Ostseeliebe vs Portalbuchungen)
        chart_id = f"channelChart{y}"
        ch_labels = json.dumps([ch["name"] for ch in chs])
        ch_values = json.dumps([ch["count"] for ch in chs])
        ch_colors = json.dumps(["#0066cc", "#ff6b6b"][:len(chs)])

        channel_year_html_parts.append(f'''
        <div class="year-section">
            <h3 class="year-title">{y}</h3>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;align-items:start">
                <div class="zusatz-summary">
                    <table class="zusatz-table">
                        <thead><tr><th>Vertriebskanal</th><th class="zk-num">Buchungen</th><th class="zk-num">Anteil</th></tr></thead>
                        <tbody>{table_rows}</tbody>
                    </table>
                    <div style="font-size:11px;color:#888;margin-top:6px">&#9654; Klicke auf eine Gruppe um die Einzelkanäle aufzuklappen</div>
                </div>
                <div class="chart-wrapper doughnut-chart" style="max-width:360px;margin:0 auto">
                    <canvas id="{chart_id}"></canvas>
                </div>
            </div>
        </div>''')

        channel_chart_js_parts.append(f'''
        new Chart(document.getElementById('{chart_id}'), {{
            type: 'doughnut',
            data: {{
                labels: {ch_labels},
                datasets: [{{
                    data: {ch_values},
                    backgroundColor: {ch_colors},
                    borderWidth: 1
                }}]
            }},
            options: {{
                responsive: true,
                plugins: {{
                    legend: {{ position: 'bottom', labels: {{ font: {{ size: 11 }} }} }},
                    tooltip: {{ callbacks: {{ label: function(c) {{
                        var total = c.dataset.data.reduce(function(a,b){{return a+b}},0);
                        var pct = (c.raw/total*100).toFixed(1);
                        return c.label + ': ' + c.raw.toLocaleString('de-DE') + ' (' + pct + ' %)';
                    }} }} }}
                }}
            }}
        }});''')

    channel_year_html = "\n".join(channel_year_html_parts)
    channel_chart_js = "\n".join(channel_chart_js_parts)

    # Locations bar
    ort_labels = json.dumps([o[0] for o in orte])
    ort_values = json.dumps([o[1] for o in orte])
    ort_colors = json.dumps([colors[i % len(colors)] for i in range(len(orte))])

    html = f'''<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Ostseeliebe - Buchungs-Dashboard</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, sans-serif;
            background: #f5f7fa;
            color: #333;
        }}
        .header {{
            background: linear-gradient(135deg, #0066cc, #00aaaa);
            color: white;
            padding: 30px 40px;
            text-align: center;
        }}
        .header h1 {{
            font-size: 28px;
            margin-bottom: 8px;
        }}
        .header .subtitle {{
            font-size: 14px;
            opacity: 0.85;
        }}
        .tabs {{
            display: flex;
            background: white;
            border-bottom: 2px solid #e0e0e0;
            padding: 0 20px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }}
        .tab {{
            padding: 14px 24px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            color: #666;
            border-bottom: 3px solid transparent;
            transition: all 0.2s;
        }}
        .tab:hover {{
            color: #0066cc;
        }}
        .tab.active {{
            color: #0066cc;
            border-bottom-color: #0066cc;
        }}
        .content {{
            max-width: 1400px;
            margin: 0 auto;
            padding: 30px 20px;
        }}
        .tab-content {{
            display: none;
        }}
        .tab-content.active {{
            display: block;
        }}
        .year-section {{
            margin-bottom: 24px;
        }}
        .year-title {{
            font-size: 20px;
            color: #0066cc;
            margin-bottom: 12px;
            padding-bottom: 6px;
            border-bottom: 2px solid #e0e0e0;
        }}
        .kpi-grid {{
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 16px;
        }}
        .kpi-card {{
            background: white;
            border-radius: 12px;
            padding: 20px;
            text-align: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            transition: transform 0.2s;
        }}
        .kpi-card:hover {{
            transform: translateY(-2px);
            box-shadow: 0 4px 16px rgba(0,0,0,0.12);
        }}
        .kpi-label {{
            font-size: 12px;
            color: #888;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
        }}
        .kpi-value {{
            font-size: 20px;
            font-weight: 700;
            color: #0066cc;
        }}
        .chart-container {{
            background: white;
            border-radius: 12px;
            padding: 24px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            margin-bottom: 24px;
        }}
        .chart-container h3 {{
            margin-bottom: 16px;
            color: #333;
        }}
        .chart-wrapper {{
            position: relative;
            width: 100%;
        }}
        .chart-wrapper.line-chart {{
            height: 400px;
        }}
        .chart-wrapper.bar-chart {{
            height: 400px;
        }}
        .chart-wrapper.hbar-chart {{
            height: 500px;
        }}
        .chart-wrapper.doughnut-chart {{
            height: 450px;
            max-width: 600px;
            margin: 0 auto;
        }}
        .comparison-table {{
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }}
        .comparison-table th,
        .comparison-table td {{
            padding: 10px 16px;
            text-align: center;
            border-bottom: 1px solid #eee;
            font-size: 14px;
        }}
        .comparison-table th {{
            background: #0066cc;
            color: white;
            font-weight: 600;
        }}
        .comparison-table tr:hover {{
            background: #f0f7ff;
        }}
        .total-row {{
            background: #f5f7fa !important;
        }}
        .total-row td {{
            border-top: 2px solid #0066cc;
        }}
        .prov-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
            white-space: nowrap;
        }}
        .prov-table th {{
            background: #0066cc;
            color: white;
            font-weight: 600;
            padding: 10px 12px;
            text-align: left;
            position: sticky;
            top: 0;
        }}
        .prov-table th.num {{
            text-align: right;
        }}
        .prov-table td {{
            padding: 7px 12px;
            border-bottom: 1px solid #eee;
        }}
        .prov-table td.num {{
            text-align: right;
            font-variant-numeric: tabular-nums;
        }}
        .prov-table tr:hover {{
            background: #f0f7ff;
        }}
        .prov-total {{
            background: #f0f4f8 !important;
            border-top: 2px solid #0066cc;
        }}
        .pk-value.green {{
            color: #2e7d32 !important;
        }}
        .prop-select {{
            width: 100%;
            padding: 12px 16px;
            font-size: 16px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            background: white;
            cursor: pointer;
            outline: none;
            transition: border-color 0.2s;
        }}
        .prop-select:focus {{
            border-color: #0066cc;
        }}
        .prop-detail-grid {{
            display: grid;
            grid-template-columns: repeat(5, 1fr);
            gap: 12px;
            margin-bottom: 16px;
        }}
        .prop-kpi {{
            background: white;
            border-radius: 10px;
            padding: 16px;
            text-align: center;
            box-shadow: 0 1px 6px rgba(0,0,0,0.07);
        }}
        .prop-kpi .pk-label {{
            font-size: 11px;
            color: #888;
            text-transform: uppercase;
            letter-spacing: 0.4px;
            margin-bottom: 6px;
        }}
        .prop-kpi .pk-value {{
            font-size: 18px;
            font-weight: 700;
            color: #0066cc;
        }}
        .prop-kpi .pk-value.green {{
            color: #2e7d32;
        }}
        .prop-section {{
            background: white;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            margin-bottom: 20px;
        }}
        .prop-section h4 {{
            margin-bottom: 12px;
            color: #333;
            font-size: 16px;
        }}
        .prop-channel-table,
        .prop-zk-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }}
        .prop-channel-table th,
        .prop-zk-table th {{
            background: #e8f0fe;
            color: #0066cc;
            font-weight: 600;
            padding: 8px 10px;
            text-align: left;
        }}
        .prop-channel-table td,
        .prop-zk-table td {{
            padding: 6px 10px;
            border-bottom: 1px solid #f0f0f0;
        }}
        .prop-channel-table .num,
        .prop-zk-table .num {{
            text-align: right;
            font-variant-numeric: tabular-nums;
        }}
        .prop-zk-table .sub {{
            color: #888;
            font-size: 12px;
        }}
        .zusatz-summary {{
            margin-top: 16px;
        }}
        .zusatz-table {{
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 1px 4px rgba(0,0,0,0.06);
            font-size: 13px;
        }}
        .zusatz-table th {{
            background: #e8f0fe;
            color: #0066cc;
            font-weight: 600;
            padding: 8px 12px;
            text-align: left;
        }}
        .zusatz-table td {{
            padding: 6px 12px;
            border-bottom: 1px solid #f0f0f0;
        }}
        .zusatz-table .zk-num {{
            text-align: right;
            font-variant-numeric: tabular-nums;
        }}
        .zusatz-table .zk-sub {{
            color: #888;
            font-size: 12px;
        }}
        .zusatz-table th.zk-num {{
            text-align: right;
        }}
        .zusatz-table th.zk-sub {{
            color: #5588bb;
        }}
        .zusatz-table .zk-total {{
            background: #f5f7fa;
            border-top: 2px solid #0066cc;
        }}
        .zusatz-detail-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
            white-space: nowrap;
        }}
        .zusatz-detail-table th {{
            background: #0066cc;
            color: white;
            font-weight: 600;
            padding: 8px 10px;
            text-align: left;
            position: sticky;
            top: 0;
        }}
        .zusatz-detail-table td {{
            padding: 6px 10px;
            border-bottom: 1px solid #eee;
        }}
        .zusatz-detail-table .zk-num {{
            text-align: right;
            font-variant-numeric: tabular-nums;
        }}
        .zusatz-detail-table .zk-sub {{
            color: #888;
        }}
        .zusatz-detail-table th.zk-num {{
            text-align: right;
        }}
        .zusatz-detail-table th.zk-sub {{
            color: #aaccff;
        }}
        .zusatz-detail-table .zk-total {{
            background: #f0f4f8;
            border-top: 2px solid #0066cc;
        }}
        .zusatz-detail-table tr:hover {{
            background: #f0f7ff;
        }}
        @media (max-width: 900px) {{
            .kpi-grid {{
                grid-template-columns: repeat(2, 1fr);
            }}
            .tabs {{
                flex-wrap: wrap;
            }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>Ostseeliebe - Buchungs-Dashboard</h1>
        <div class="subtitle">Aktualisiert am {update_date}</div>
    </div>

    <div class="tabs">
        <div class="tab active" data-tab="uebersicht">\u00dcbersicht</div>
        <div class="tab" data-tab="jahresvergleich">Jahresvergleich</div>
        <div class="tab" data-tab="reiseprofile">Reiseprofile</div>
        <div class="tab" data-tab="vertriebskanaele">Vertriebskan\u00e4le</div>
        <div class="tab" data-tab="orte">Orte</div>
        <div class="tab" data-tab="zusatzkosten">Zusatzkosten</div>
        <div class="tab" data-tab="provisionen">Provisionen</div>
        <div class="tab" data-tab="unterkunft_detail">Unterkunft Detail</div>
    </div>

    <div class="content">
        <!-- Uebersicht -->
        <div class="tab-content active" id="uebersicht">
            {bestand_html}
            {kpi_html}
        </div>

        <!-- Jahresvergleich -->
        <div class="tab-content" id="jahresvergleich">
            <div class="chart-container">
                <h3>\u00dcbernachtungen pro Monat (Jahresvergleich)</h3>
                <div class="chart-wrapper line-chart">
                    <canvas id="monthlyChart"></canvas>
                </div>
            </div>
            <div class="chart-container">
                <h3>Buchungen pro Monat (Vergleichstabelle)</h3>
                {comparison_table}
            </div>
        </div>

        <!-- Reiseprofile -->
        <div class="tab-content" id="reiseprofile">
            <div class="chart-container">
                <h3>Reiseprofile \u2013 Jahresvergleich</h3>
                <p style="color:#666;font-size:13px;margin-bottom:12px;">Abgeleitet aus gebuchten Zusatzleistungen: Kinderreisebett/Hochstuhl = Familie, Hund-Zuschlag = Hundeurlaub, Aufschlag Mitreisende = Gruppe, Sauna/Whirlpool = Wellness, Wallbox = E-Auto.</p>
                <div class="chart-wrapper bar-chart">
                    <canvas id="profileYearChart"></canvas>
                </div>
            </div>
            <div class="chart-container">
                <h3>Top 15 Unterk\u00fcnfte \u2013 Hundeurlaub-Anteil</h3>
                <div class="chart-wrapper hbar-chart">
                    <canvas id="hundChart"></canvas>
                </div>
            </div>
            <div class="chart-container">
                <h3>Alle Unterk\u00fcnfte \u2013 Zielgruppen-Verteilung</h3>
                <p style="color:#666;font-size:13px;margin-bottom:12px;">Sortiert nach h\u00f6chstem Hundeurlaub-Anteil. Der farbige Balken zeigt die Verteilung:
                    <span style="color:#0066cc;">\u25cf Paare</span>
                    <span style="color:#ff6b6b;">\u25cf Hund</span>
                    <span style="color:#ffa500;">\u25cf Gruppe</span>
                    <span style="color:#4ecdc4;">\u25cf Familie</span>
                    <span style="color:#aa96da;">\u25cf Wellness</span>
                    <span style="color:#95e1d3;">\u25cf E-Auto</span>
                </p>
                <div style="overflow-x:auto;">
                <table class="prov-table">
                    <thead>
                        <tr>
                            <th>Unterkunft</th>
                            <th class="num">Buchungen</th>
                            <th class="num">Hund</th>
                            <th class="num">Familie</th>
                            <th class="num">Gruppe</th>
                            <th class="num">Wellness</th>
                            <th class="num">Paare/Einzel</th>
                            <th>Verteilung</th>
                        </tr>
                    </thead>
                    <tbody>
                        {profile_table_rows}
                    </tbody>
                </table>
                </div>
            </div>
        </div>

        <!-- Vertriebskanaele -->
        <div class="tab-content" id="vertriebskanaele">
            {channel_year_html}
        </div>

        <!-- Orte -->
        <div class="tab-content" id="orte">
            <div class="chart-container">
                <h3>Buchungen nach Ort</h3>
                <div class="chart-wrapper bar-chart">
                    <canvas id="ortChart"></canvas>
                </div>
            </div>
        </div>

        <!-- Provisionen -->
        <div class="tab-content" id="provisionen">
            <div class="chart-container">
                <h3>Provisionseinnahmen Ostseeliebe \u2013 nach Unterkunft</h3>
                <p style="color:#666;font-size:13px;margin-bottom:16px;">Miete Vermittler + Zusatzkosten Vermittler = Provision gesamt. Sortiert nach h\u00f6chster Provision.</p>
            </div>
            {prov_tab_html}
        </div>

        <!-- Unterkunft Detail -->
        <div class="tab-content" id="unterkunft_detail">
            <div class="chart-container">
                <h3>Unterkunft ausw\u00e4hlen</h3>
                <select id="propSelect" class="prop-select">
                    <option value="">-- Bitte Unterkunft w\u00e4hlen --</option>
                    {property_options}
                </select>
            </div>
            <div id="propDetailContent"></div>
        </div>

        <!-- Zusatzkosten -->
        <div class="tab-content" id="zusatzkosten">
            <div class="chart-container">
                <h3>Top 10 Zusatzkosten \u2013 Vermittler vs. Eigent\u00fcmer</h3>
                <div class="chart-wrapper hbar-chart">
                    <canvas id="zusatzChart"></canvas>
                </div>
            </div>
            <div class="chart-container">
                <h3>Top 5 Zusatzkosten \u2013 Jahresvergleich</h3>
                <div class="chart-wrapper bar-chart">
                    <canvas id="zusatzYearChart"></canvas>
                </div>
            </div>
            <div class="chart-container">
                <h3>Alle Zusatzkosten \u2013 Detailtabelle</h3>
                {zusatz_detail_table}
            </div>
        </div>
    </div>

    <script>
        // Tab switching
        document.querySelectorAll('.tab').forEach(function(tab) {{
            tab.addEventListener('click', function() {{
                document.querySelectorAll('.tab').forEach(function(t) {{ t.classList.remove('active'); }});
                document.querySelectorAll('.tab-content').forEach(function(c) {{ c.classList.remove('active'); }});
                tab.classList.add('active');
                document.getElementById(tab.dataset.tab).classList.add('active');
            }});
        }});

        // Chart.js defaults
        Chart.defaults.font.family = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif";

        // Monthly Line Chart
        new Chart(document.getElementById('monthlyChart'), {{
            type: 'line',
            data: {{
                labels: {json_month_names},
                datasets: {json_line_datasets}
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ position: 'top' }}
                }},
                scales: {{
                    y: {{
                        beginAtZero: true,
                        title: {{ display: true, text: 'N\u00e4chte' }}
                    }}
                }}
            }}
        }});

        // Reiseprofile Year Chart (stacked bar)
        new Chart(document.getElementById('profileYearChart'), {{
            type: 'bar',
            data: {{
                labels: {json_years},
                datasets: {json_profile_datasets}
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ position: 'top' }}
                }},
                scales: {{
                    x: {{ stacked: true }},
                    y: {{
                        stacked: true,
                        beginAtZero: true,
                        title: {{ display: true, text: 'Anzahl Buchungen' }}
                    }}
                }}
            }}
        }});

        // Top 15 Hund Properties Chart
        new Chart(document.getElementById('hundChart'), {{
            type: 'bar',
            data: {{
                labels: {json_hund_labels},
                datasets: [{{
                    label: 'Hundeurlaub-Anteil (%)',
                    data: {json_hund_pcts},
                    backgroundColor: '#ff6b6b',
                    borderRadius: 4
                }}]
            }},
            options: {{
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ display: false }},
                    tooltip: {{
                        callbacks: {{
                            label: function(ctx) {{
                                var counts = {json_hund_counts};
                                return ctx.raw + '% (' + counts[ctx.dataIndex] + ' Buchungen)';
                            }}
                        }}
                    }}
                }},
                scales: {{
                    x: {{
                        beginAtZero: true,
                        max: 100,
                        title: {{ display: true, text: 'Anteil Hundeurlaub (%)' }},
                        ticks: {{ callback: function(v) {{ return v + '%'; }} }}
                    }}
                }}
            }}
        }});

        // Sales Channels Doughnut Charts (per year)
        {channel_chart_js}

        // Locations Bar Chart
        new Chart(document.getElementById('ortChart'), {{
            type: 'bar',
            data: {{
                labels: {ort_labels},
                datasets: [{{
                    label: 'Buchungen',
                    data: {ort_values},
                    backgroundColor: {ort_colors},
                    borderRadius: 6
                }}]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ display: false }}
                }},
                scales: {{
                    y: {{
                        beginAtZero: true,
                        title: {{ display: true, text: 'Anzahl Buchungen' }}
                    }}
                }}
            }}
        }});

        // Zusatzkosten Stacked Horizontal Bar (Vermittler vs Eigentuemer)
        new Chart(document.getElementById('zusatzChart'), {{
            type: 'bar',
            data: {{
                labels: {zusatz_chart_labels},
                datasets: [
                    {{
                        label: 'Vermittler',
                        data: {zusatz_chart_vermittler},
                        backgroundColor: '#0066cc',
                        borderRadius: 4
                    }},
                    {{
                        label: 'Eigent\u00fcmer',
                        data: {zusatz_chart_eigentuemer},
                        backgroundColor: '#00aaaa',
                        borderRadius: 4
                    }}
                ]
            }},
            options: {{
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ position: 'top' }},
                    tooltip: {{
                        callbacks: {{
                            label: function(ctx) {{
                                return ctx.dataset.label + ': ' + new Intl.NumberFormat('de-DE', {{style:'currency',currency:'EUR'}}).format(ctx.raw);
                            }}
                        }}
                    }}
                }},
                scales: {{
                    x: {{
                        stacked: true,
                        ticks: {{
                            callback: function(v) {{ return new Intl.NumberFormat('de-DE', {{style:'currency',currency:'EUR',maximumFractionDigits:0}}).format(v); }}
                        }}
                    }},
                    y: {{
                        stacked: true
                    }}
                }}
            }}
        }});

        // Zusatzkosten Yearly Comparison (grouped bar)
        new Chart(document.getElementById('zusatzYearChart'), {{
            type: 'bar',
            data: {{
                labels: {zusatz_yearly_labels},
                datasets: {json_zusatz_yearly_datasets}
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                plugins: {{
                    legend: {{ position: 'top' }},
                    tooltip: {{
                        callbacks: {{
                            label: function(ctx) {{
                                return ctx.dataset.label + ': ' + new Intl.NumberFormat('de-DE', {{style:'currency',currency:'EUR'}}).format(ctx.raw);
                            }}
                        }}
                    }}
                }},
                scales: {{
                    y: {{
                        beginAtZero: true,
                        ticks: {{
                            callback: function(v) {{ return new Intl.NumberFormat('de-DE', {{style:'currency',currency:'EUR',maximumFractionDigits:0}}).format(v); }}
                        }}
                    }}
                }}
            }}
        }});
        // --- Unterkunft Detail ---
        var propData = {json_property_data};
        var allYears = {json_years};
        var propChartInstance = null;

        function fmtEuro(n) {{
            return new Intl.NumberFormat('de-DE', {{style:'currency',currency:'EUR'}}).format(n);
        }}
        function fmtNum(n, d) {{
            return new Intl.NumberFormat('de-DE', {{minimumFractionDigits:d||0, maximumFractionDigits:d||0}}).format(n);
        }}

        function renderProperty(name) {{
            var container = document.getElementById('propDetailContent');
            if (!name || !propData[name]) {{
                container.innerHTML = '<p style="color:#888;padding:20px;">Bitte eine Unterkunft ausw\u00e4hlen.</p>';
                return;
            }}
            var p = propData[name];
            var html = '<h3 style="margin:20px 0 6px;font-size:22px;color:#0066cc;">' + name + '</h3>';
            html += '<p style="color:#666;margin-bottom:20px;">Ort: ' + p.ort + '</p>';

            allYears.forEach(function(y) {{
                var yd = p.years[y];
                if (!yd || yd.buchungen === 0) return;
                html += '<div class="prop-section">';
                html += '<h4 style="color:#0066cc;border-bottom:2px solid #e0e0e0;padding-bottom:6px;">' + y + '</h4>';

                // KPI Grid
                html += '<div class="prop-detail-grid">';
                html += '<div class="prop-kpi"><div class="pk-label">Buchungen</div><div class="pk-value">' + yd.buchungen + '</div></div>';
                html += '<div class="prop-kpi"><div class="pk-label">N\u00e4chte</div><div class="pk-value">' + fmtNum(yd.naechte) + '</div></div>';
                html += '<div class="prop-kpi"><div class="pk-label">Reisepreis</div><div class="pk-value">' + fmtEuro(yd.reisepreis) + '</div></div>';
                html += '<div class="prop-kpi"><div class="pk-label">Miete gesamt</div><div class="pk-value">' + fmtEuro(yd.miete_gesamt) + '</div></div>';
                html += '<div class="prop-kpi"><div class="pk-label">Miete Eigent\u00fcmer</div><div class="pk-value green">' + fmtEuro(yd.miete_eigentuemer) + '</div></div>';
                html += '<div class="prop-kpi"><div class="pk-label">Provision (Verm.)</div><div class="pk-value" style="color:#e65100;">' + fmtEuro(yd.miete_vermittler) + '</div></div>';
                html += '<div class="prop-kpi"><div class="pk-label">Prov.satz</div><div class="pk-value" style="color:#e65100;">' + fmtNum(yd.provision_pct, 1) + ' %</div></div>';
                html += '<div class="prop-kpi"><div class="pk-label">\u00d8 Preis/Nacht</div><div class="pk-value">' + fmtEuro(yd.avg_preis_nacht) + '</div></div>';
                html += '<div class="prop-kpi"><div class="pk-label">\u00d8 Aufenthalt</div><div class="pk-value">' + fmtNum(yd.avg_aufenthalt, 1) + ' N.</div></div>';
                html += '<div class="prop-kpi"><div class="pk-label">Auslastung</div><div class="pk-value">' + fmtNum(yd.belegung_pct, 1) + ' %</div></div>';
                html += '</div>';

                // Vertriebskanäle
                if (yd.channels && yd.channels.length > 0) {{
                    html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:12px;">';
                    html += '<div>';
                    html += '<h4 style="font-size:14px;margin-bottom:8px;">Vertriebskan\u00e4le</h4>';
                    html += '<table class="prop-channel-table"><thead><tr><th>Kanal</th><th class="num">Buchungen</th></tr></thead><tbody>';
                    yd.channels.forEach(function(ch) {{
                        html += '<tr><td>' + (ch[0] || '(leer)') + '</td><td class="num">' + ch[1] + '</td></tr>';
                    }});
                    html += '</tbody></table></div>';

                    // Zusatzkosten
                    html += '<div>';
                    html += '<h4 style="font-size:14px;margin-bottom:8px;">Zusatzkosten</h4>';
                    if (yd.zusatzkosten && yd.zusatzkosten.length > 0) {{
                        html += '<table class="prop-zk-table"><thead><tr><th>Kategorie</th><th class="num">Gesamt</th><th class="num sub">Vermittler</th><th class="num sub">Eigent\u00fcmer</th></tr></thead><tbody>';
                        var zkTotal = 0, zkV = 0, zkE = 0;
                        yd.zusatzkosten.forEach(function(z) {{
                            html += '<tr><td>' + z.name + '</td><td class="num">' + fmtEuro(z.gesamt) + '</td><td class="num sub">' + fmtEuro(z.vermittler) + '</td><td class="num sub">' + fmtEuro(z.eigentuemer) + '</td></tr>';
                            zkTotal += z.gesamt; zkV += z.vermittler; zkE += z.eigentuemer;
                        }});
                        html += '<tr style="background:#f5f7fa;border-top:2px solid #0066cc;"><td><strong>Summe</strong></td><td class="num"><strong>' + fmtEuro(zkTotal) + '</strong></td><td class="num sub"><strong>' + fmtEuro(zkV) + '</strong></td><td class="num sub"><strong>' + fmtEuro(zkE) + '</strong></td></tr>';
                        html += '</tbody></table>';
                    }} else {{
                        html += '<p style="color:#aaa;">Keine Zusatzkosten</p>';
                    }}
                    html += '</div></div>';
                }}

                html += '</div>'; // prop-section
            }});

            // Jahresvergleich mini chart
            html += '<div class="prop-section"><h4>Jahresvergleich \u2013 Reisepreis</h4><canvas id="propYearChart" style="max-height:300px;"></canvas></div>';
            container.innerHTML = html;

            // Render mini chart
            if (propChartInstance) propChartInstance.destroy();
            var ctx = document.getElementById('propYearChart');
            if (ctx) {{
                var labels = []; var rpData = []; var meData = [];
                allYears.forEach(function(y) {{
                    var yd = p.years[y];
                    if (yd && yd.buchungen > 0) {{
                        labels.push(y);
                        rpData.push(yd.reisepreis);
                        meData.push(yd.miete_eigentuemer);
                    }}
                }});
                propChartInstance = new Chart(ctx, {{
                    type: 'bar',
                    data: {{
                        labels: labels,
                        datasets: [
                            {{ label: 'Reisepreis', data: rpData, backgroundColor: '#0066cc', borderRadius: 4 }},
                            {{ label: 'Miete Eigent\u00fcmer', data: meData, backgroundColor: '#00aaaa', borderRadius: 4 }}
                        ]
                    }},
                    options: {{
                        responsive: true,
                        plugins: {{
                            legend: {{ position: 'top' }},
                            tooltip: {{
                                callbacks: {{
                                    label: function(ctx) {{ return ctx.dataset.label + ': ' + fmtEuro(ctx.raw); }}
                                }}
                            }}
                        }},
                        scales: {{
                            y: {{
                                beginAtZero: true,
                                ticks: {{ callback: function(v) {{ return fmtEuro(v); }} }}
                            }}
                        }}
                    }}
                }});
            }}
        }}

        document.getElementById('propSelect').addEventListener('change', function() {{
            renderProperty(this.value);
        }});
    </script>
</body>
</html>'''
    return html


def generate_property_html(prop_name, pdata, years):
    """Generate a standalone HTML file for a single property."""
    update_date = datetime.now().strftime("%d.%m.%Y %H:%M")

    year_sections = ""
    chart_labels = []
    chart_rp = []
    chart_me = []

    for y in years:
        yd = pdata["years"].get(y)
        if not yd or yd["buchungen"] == 0:
            continue

        chart_labels.append(y)
        chart_rp.append(round(yd["reisepreis"], 2))
        chart_me.append(round(yd["miete_eigentuemer"], 2))
        chart_mv = round(yd.get("miete_vermittler", 0), 2)

        # Channels table
        ch_rows = ""
        for ch in yd["channels"]:
            ch_name = ch[0] if isinstance(ch, (list, tuple)) else ch
            ch_count = ch[1] if isinstance(ch, (list, tuple)) else 0
            ch_rows += f'<tr><td>{ch_name or "(leer)"}</td><td class="num">{ch_count}</td></tr>'

        # Zusatzkosten table
        zk_rows = ""
        zk_total_g = zk_total_v = zk_total_e = 0
        for z in yd["zusatzkosten"]:
            zk_rows += f'<tr><td>{z["name"]}</td><td class="num">{format_euro(z["gesamt"])}</td><td class="num sub">{format_euro(z["vermittler"])}</td><td class="num sub">{format_euro(z["eigentuemer"])}</td></tr>'
            zk_total_g += z["gesamt"]
            zk_total_v += z["vermittler"]
            zk_total_e += z["eigentuemer"]
        if zk_rows:
            zk_rows += f'<tr style="background:#f5f7fa;border-top:2px solid #0066cc;"><td><strong>Summe</strong></td><td class="num"><strong>{format_euro(zk_total_g)}</strong></td><td class="num sub"><strong>{format_euro(zk_total_v)}</strong></td><td class="num sub"><strong>{format_euro(zk_total_e)}</strong></td></tr>'

        year_sections += f'''
        <div class="section">
            <h2>{y}</h2>
            <div class="kpi-grid">
                <div class="kpi"><div class="kl">Buchungen</div><div class="kv">{yd["buchungen"]}</div></div>
                <div class="kpi"><div class="kl">N\u00e4chte</div><div class="kv">{format_german_number(yd["naechte"], 0)}</div></div>
                <div class="kpi"><div class="kl">Reisepreis</div><div class="kv">{format_euro(yd["reisepreis"])}</div></div>
                <div class="kpi"><div class="kl">Miete gesamt</div><div class="kv">{format_euro(yd["miete_gesamt"])}</div></div>
                <div class="kpi"><div class="kl">Miete Eigent\u00fcmer</div><div class="kv green">{format_euro(yd["miete_eigentuemer"])}</div></div>
                <div class="kpi"><div class="kl">Provision (Verm.)</div><div class="kv" style="color:#e65100;">{format_euro(yd.get("miete_vermittler", 0))}</div></div>
                <div class="kpi"><div class="kl">Prov.satz</div><div class="kv" style="color:#e65100;">{format_german_number(yd.get("provision_pct", 0), 1)} %</div></div>
                <div class="kpi"><div class="kl">\u00d8 Preis/Nacht</div><div class="kv">{format_euro(yd["avg_preis_nacht"])}</div></div>
                <div class="kpi"><div class="kl">\u00d8 Aufenthalt</div><div class="kv">{format_german_number(yd["avg_aufenthalt"], 1)} N.</div></div>
                <div class="kpi"><div class="kl">Auslastung</div><div class="kv">{format_german_number(yd["belegung_pct"], 1)} %</div></div>
            </div>
            <div class="two-col">
                <div>
                    <h3>Vertriebskan\u00e4le</h3>
                    <table><thead><tr><th>Kanal</th><th class="num">Buchungen</th></tr></thead><tbody>{ch_rows}</tbody></table>
                </div>
                <div>
                    <h3>Zusatzkosten</h3>
                    {"<table><thead><tr><th>Kategorie</th><th class='num'>Gesamt</th><th class='num sub'>Vermittler</th><th class='num sub'>Eigent.</th></tr></thead><tbody>" + zk_rows + "</tbody></table>" if zk_rows else "<p style='color:#aaa;'>Keine Zusatzkosten</p>"}
                </div>
            </div>
        </div>'''

    json_labels = json.dumps(chart_labels)
    json_rp = json.dumps(chart_rp)
    json_me = json.dumps(chart_me)

    return f'''<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{prop_name} \u2013 Umsatz\u00fcbersicht</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
    <style>
        * {{ margin:0; padding:0; box-sizing:border-box; }}
        body {{ font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif; background:#f5f7fa; color:#333; }}
        .header {{ background:linear-gradient(135deg,#0066cc,#00aaaa); color:white; padding:30px 40px; }}
        .header h1 {{ font-size:24px; margin-bottom:4px; }}
        .header .sub {{ opacity:0.85; font-size:13px; }}
        .container {{ max-width:1100px; margin:0 auto; padding:24px 20px; }}
        .section {{ background:white; border-radius:12px; padding:24px; box-shadow:0 2px 8px rgba(0,0,0,0.08); margin-bottom:20px; }}
        .section h2 {{ color:#0066cc; font-size:20px; border-bottom:2px solid #e0e0e0; padding-bottom:6px; margin-bottom:16px; }}
        .kpi-grid {{ display:grid; grid-template-columns:repeat(5,1fr); gap:12px; margin-bottom:16px; }}
        .kpi {{ background:#f8fafc; border-radius:10px; padding:14px; text-align:center; }}
        .kl {{ font-size:11px; color:#888; text-transform:uppercase; letter-spacing:.4px; margin-bottom:4px; }}
        .kv {{ font-size:17px; font-weight:700; color:#0066cc; }}
        .kv.green {{ color:#2e7d32; }}
        .two-col {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; }}
        h3 {{ font-size:14px; margin-bottom:8px; color:#555; }}
        table {{ width:100%; border-collapse:collapse; font-size:13px; }}
        th {{ background:#e8f0fe; color:#0066cc; font-weight:600; padding:8px 10px; text-align:left; }}
        td {{ padding:6px 10px; border-bottom:1px solid #f0f0f0; }}
        .num {{ text-align:right; font-variant-numeric:tabular-nums; }}
        .sub {{ color:#888; font-size:12px; }}
        .chart-section {{ background:white; border-radius:12px; padding:24px; box-shadow:0 2px 8px rgba(0,0,0,0.08); margin-bottom:20px; }}
        .chart-section h3 {{ font-size:16px; color:#333; margin-bottom:12px; }}
        @media(max-width:800px) {{ .kpi-grid {{ grid-template-columns:repeat(2,1fr); }} .two-col {{ grid-template-columns:1fr; }} }}
    </style>
</head>
<body>
    <div class="header">
        <h1>{prop_name}</h1>
        <div class="sub">{pdata["ort"]} \u2014 Umsatz\u00fcbersicht \u2014 Stand {update_date}</div>
    </div>
    <div class="container">
        <div class="chart-section">
            <h3>Jahresvergleich \u2013 Reisepreis & Miete Eigent\u00fcmer</h3>
            <canvas id="yearChart" style="max-height:300px;"></canvas>
        </div>
        {year_sections}
    </div>
    <script>
        new Chart(document.getElementById('yearChart'), {{
            type: 'bar',
            data: {{
                labels: {json_labels},
                datasets: [
                    {{ label: 'Reisepreis', data: {json_rp}, backgroundColor: '#0066cc', borderRadius: 4 }},
                    {{ label: 'Miete Eigent\u00fcmer', data: {json_me}, backgroundColor: '#00aaaa', borderRadius: 4 }}
                ]
            }},
            options: {{
                responsive: true,
                plugins: {{
                    legend: {{ position: 'top' }},
                    tooltip: {{ callbacks: {{ label: function(c) {{ return c.dataset.label + ': ' + new Intl.NumberFormat('de-DE',{{style:'currency',currency:'EUR'}}).format(c.raw); }} }} }}
                }},
                scales: {{
                    y: {{
                        beginAtZero: true,
                        ticks: {{ callback: function(v) {{ return new Intl.NumberFormat('de-DE',{{style:'currency',currency:'EUR',maximumFractionDigits:0}}).format(v); }} }}
                    }}
                }}
            }}
        }});
    </script>
</body>
</html>'''


def main():
    csv_path = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_CSV
    out_path = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_OUT

    if not os.path.exists(csv_path):
        print(f"Fehler: CSV-Datei nicht gefunden: {csv_path}")
        sys.exit(1)

    # Objektstammdaten (optional)
    stamm_path = os.path.join(os.path.dirname(csv_path), "Objektstammdaten.xlsx")
    if len(sys.argv) > 3:
        stamm_path = sys.argv[3]

    # Historical cache: load frozen 2024/2025 data if available
    cache_path = os.path.join(os.path.dirname(os.path.abspath(csv_path)), "historical_cache.json")
    cached_bookings = []
    from_year = None
    if os.path.exists(cache_path):
        cached_bookings, cached_years = load_cache(cache_path)
        from_year = max(cached_years) + 1 if cached_years else None
        print(f"  Cache geladen: {len(cached_bookings)} Buchungen aus {cached_years}")

    print(f"Lese Buchungen aus: {csv_path}" + (f" (nur ab {from_year})" if from_year else ""))
    bookings = read_bookings(csv_path, from_year=from_year)
    bookings = cached_bookings + bookings
    print(f"  {len(bookings)} Buchungen gesamt (davon {len(bookings) - len(cached_bookings)} aus CSV)")

    stammdaten = read_objektstammdaten(stamm_path)
    if stammdaten:
        t = stammdaten["totals"]
        print(f"  Objektstammdaten: {t['count']} Unterkünfte, {t['wohnflaeche']} m², "
              f"{t['schlafzimmer']} Schlafzimmer, {t['max_personen']} Schlafplätze")
    else:
        print("  Objektstammdaten nicht gefunden oder openpyxl nicht installiert – überspringe Bestandskennzahlen")

    data = compute_data(bookings)
    data["stammdaten"] = stammdaten
    print(f"  Jahre: {data['years']}")
    for y in data["years"]:
        k = data["kpis"][y]
        print(f"    {y}: {k['buchungen']} Buchungen, {k['naechte']:.0f} Naechte, "
              f"Reisepreis {k['reisepreis']:.2f}, Miete ges. {k['miete_gesamt']:.2f}, "
              f"Miete Eig. {k['miete_eigentuemer']:.2f}")

    html = generate_html(data)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\nDashboard geschrieben: {out_path}")
    print(f"  Dateigroesse: {os.path.getsize(out_path):,} Bytes")

    # Generate individual property HTML files
    out_dir = os.path.dirname(out_path)
    prop_dir = os.path.join(out_dir, "unterkuenfte")
    os.makedirs(prop_dir, exist_ok=True)

    property_data = data["property_data"]
    years = data["years"]
    count = 0
    for prop_name, pdata in sorted(property_data.items()):
        # Skip properties with no bookings
        if all(pdata["years"].get(y, {}).get("buchungen", 0) == 0 for y in years):
            continue
        # Safe filename
        safe_name = prop_name.replace("/", "-").replace("\\", "-").replace(" ", "_")
        safe_name = "".join(c for c in safe_name if c.isalnum() or c in "-_").strip("_")
        if not safe_name:
            safe_name = f"unterkunft_{count}"
        filepath = os.path.join(prop_dir, f"{safe_name}.html")
        prop_html = generate_property_html(prop_name, pdata, years)
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(prop_html)
        count += 1

    print(f"\n{count} Unterkunft-Einzelseiten geschrieben: {prop_dir}/")


if __name__ == "__main__":
    main()
