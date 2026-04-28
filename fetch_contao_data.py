#!/usr/bin/env python3
"""
fetch_contao_data.py
====================
Loggt sich in das Contao-Backend ein und lädt die Buchungs-CSV herunter.
Wird täglich von der GitHub Action ausgeführt.

Quellen:
  1. key=journal      → abgerechnete Buchungen (historisch vollständig)
  2. key=salesbooking → ALLE Buchungen aktuelles + nächstes Jahr
                        (inkl. noch nicht abgerechnete)

Beide werden zusammengeführt; Duplikate (gleiche Vorgangsnummer) werden
nur einmal gezählt, wobei die Journal-Version Vorrang hat.

Benötigt folgende Umgebungsvariablen (als GitHub Secrets hinterlegt):
    CONTAO_URL      z.B. https://www.ostseeliebe-ferienwohnungen.de
    CONTAO_USER     Benutzername für Contao-Login
    CONTAO_PASS     Passwort für Contao-Login

Aufruf:
    python fetch_contao_data.py
    python fetch_contao_data.py --from 01.01.2017 --to 31.12.2027
"""
import os
import sys
import csv
import io
import argparse
import re
from datetime import datetime, timedelta
from pathlib import Path

try:
    import requests
except ImportError:
    print("❌  requests fehlt. Bitte: pip install requests")
    sys.exit(1)

# ──────────────────────────────────────────────────────────────────
# Konfiguration
# ──────────────────────────────────────────────────────────────────
BASE_URL = os.environ.get("CONTAO_URL",  "https://www.ostseeliebe-ferienwohnungen.de")
USERNAME = os.environ.get("CONTAO_USER", "")
PASSWORD = os.environ.get("CONTAO_PASS", "")
OUT_CSV  = "buchungen_export_2027.csv"

# Felder die exportiert werden (für Dashboard relevant)
EXPORT_FIELDS = [
    # Pflichtfelder
    "objnr", "city", "arrival", "departure", "nights",
    "bookingid", "schannel", "totalrent", "paid", "rent_total",
    # Gebühren
    "touristtax", "shortstay", "provisionfee", "cancelfee",
    # Leistungskategorien (aktiv)
    "group_1",   # 01. Endreinigung
    "group_2",   # 02. Zusatzleistungen (buchbar)
    "group_23",  # 03. Zusatzleistungen (nur Ostseeliebe)
    "group_26",  # 04. Zusatzleistungen (provisionsfähig)
    "group_6",   # 05. Zusatzleistungen (inklusive)
    "group_9",   # 06. Nebenkosten (obligatorisch)
    "group_33",  # 07. Nebenkosten (teilw. Eigentümer)
    "group_22",  # 08. Kaution
    "group_32",  # 09. Divers
]


# ──────────────────────────────────────────────────────────────────
# Hilfsfunktionen
# ──────────────────────────────────────────────────────────────────

def get_request_token(session: requests.Session, url: str) -> str:
    """Holt den REQUEST_TOKEN von einer Contao-Seite."""
    resp = session.get(url, timeout=30)
    resp.raise_for_status()
    match = re.search(r'name="REQUEST_TOKEN"\s+value="([^"]+)"', resp.text)
    if not match:
        match = re.search(r'<meta name="request-token" content="([^"]+)"', resp.text)
    if not match:
        raise RuntimeError("REQUEST_TOKEN nicht gefunden auf: " + url)
    return match.group(1)


def contao_login(session: requests.Session) -> bool:
    """Loggt sich in das Contao-Backend ein."""
    login_url = f"{BASE_URL}/contao/login"
    print(f"🔐  Login bei {BASE_URL} …")
    token = get_request_token(session, login_url)
    resp = session.post(login_url, data={
        "REQUEST_TOKEN": token,
        "username":      USERNAME,
        "password":      PASSWORD,
    }, allow_redirects=True, timeout=30)
    if "do=main" in resp.url or ("/contao" in resp.url and "login" not in resp.url):
        print("✅  Login erfolgreich")
        return True
    if 'name="username"' in resp.text:
        print("❌  Login fehlgeschlagen – Benutzername/Passwort prüfen")
        return False
    print("✅  Login erfolgreich (Redirect erkannt)")
    return True


def _extract_csv(resp: requests.Response, label: str) -> str | None:
    """
    Versucht einen CSV-String aus einem HTTP-Response zu extrahieren.
    Gibt den CSV-Text zurück oder None wenn kein CSV erkennbar.
    """
    content_type = resp.headers.get("content-type", "")
    text = resp.text

    # Direkte CSV-Antwort
    if "text/csv" in content_type or "application/csv" in content_type:
        print(f"✅  {label}: CSV direkt ({len(resp.content):,} Bytes)")
        return text

    # Plausibilitätsprüfung: Aussehen einer CSV
    lines = text.strip().split("\n")
    if len(lines) > 5 and ";" in lines[0]:
        print(f"✅  {label}: CSV erkannt ({len(lines)} Zeilen)")
        return text

    # Download-Link in HTML-Antwort
    match = re.search(r'href="([^"]*\.csv[^"]*)"', text)
    if match:
        csv_url = match.group(1)
        if not csv_url.startswith("http"):
            csv_url = BASE_URL + "/" + csv_url.lstrip("/")
        print(f"📎  {label}: Download-Link → {csv_url}")
        csv_resp = resp.connection.send(
            resp.request.__class__("GET", csv_url), timeout=60
        )
        # Einfacher: neuer GET-Request (Session wird von Aufrufer übergeben)
        return None  # Fallback – wird vom Aufrufer gehandelt

    print(f"⚠️  {label}: Kein CSV erkannt")
    print(f"   Content-Type: {content_type}")
    print(f"   Erste 200 Zeichen: {text[:200]!r}")
    return None


# ──────────────────────────────────────────────────────────────────
# CSV-Fetch: Journal (abgerechnete Buchungen)
# ──────────────────────────────────────────────────────────────────

def fetch_journal_csv(session: requests.Session, date_from: str, date_to: str) -> str:
    """
    Lädt abgerechnete Buchungen über den Journal-Export.
    date_from / date_to: Format 'DD.MM.YYYY'
    """
    stats_url = f"{BASE_URL}/contao?do=fewoOffice_stats&key=journal"
    print(f"📥  Journal-Export: {date_from} – {date_to} …")
    token = get_request_token(session, stats_url)

    post_data = [
        ("FORM_ACTION",   "doExport"),
        ("REQUEST_TOKEN", token),
        ("period",        f"{date_from} - {date_to}"),
        ("selectorfield", "departure"),  # nach Abreisedatum filtern
        ("object",        ""),
        ("house",         ""),
        ("agent",         ""),
        ("owner",         ""),
        ("output",        "details"),
        ("schannel",      ""),
    ] + [("exportFields[]", f) for f in EXPORT_FIELDS]

    resp = session.post(stats_url, data=post_data, timeout=60)
    resp.raise_for_status()

    content_type = resp.headers.get("content-type", "")
    text = resp.text

    if "text/csv" in content_type or "application/csv" in content_type:
        print(f"✅  Journal-CSV: {len(resp.content):,} Bytes")
        return text

    lines = text.strip().split("\n")
    if len(lines) > 5 and ";" in (lines[0] if lines else ""):
        print(f"✅  Journal-CSV: {len(lines)} Zeilen")
        return text

    match = re.search(r'href="([^"]*\.csv[^"]*)"', text)
    if match:
        csv_url = match.group(1)
        if not csv_url.startswith("http"):
            csv_url = BASE_URL + "/" + csv_url.lstrip("/")
        print(f"📎  Journal-CSV via Link: {csv_url}")
        csv_resp = session.get(csv_url, timeout=60)
        csv_resp.raise_for_status()
        print(f"✅  Journal-CSV: {len(csv_resp.content):,} Bytes")
        return csv_resp.text

    print("⚠️  Journal: Unbekanntes Format – speichere trotzdem")
    return text


# ──────────────────────────────────────────────────────────────────
# CSV-Fetch: Salesbooking (alle Buchungen, auch nicht abgerechnet)
# ──────────────────────────────────────────────────────────────────

def fetch_salesbooking_csv(session: requests.Session, year: int) -> str | None:
    """
    Lädt ALLE Buchungen eines Jahres (inkl. noch nicht abgerechnete).

    Die Salesbooking-Statistikseite bietet einen direkten CSV-Download über
    den Parameter export=csv – gleiche URL wie die HTML-Ansicht, nur ohne
    submit=Anwenden und mit &export=csv stattdessen.

    URL-Muster (aus dem Browser-Link "export CSV" ermittelt):
        fewoOffice_stats?key=salesbooking&export=csv&year=YYYY&...
    """
    export_url = (
        f"{BASE_URL}/contao?do=fewoOffice_stats&key=salesbooking"
        f"&export=csv&issuer=&owner=&houses=&object=&year={year}&months="
    )
    print(f"📥  Salesbooking-CSV-Export: Jahr {year} …")

    try:
        resp = session.get(export_url, timeout=120)
        resp.raise_for_status()

        ct   = resp.headers.get("content-type", "")
        text = resp.text
        lines = [l for l in text.strip().split("\n") if l.strip()]

        # ── Erfolgscheck ─────────────────────────────────────────────
        is_csv = (
            "text/csv" in ct
            or "application/csv" in ct
            or (len(lines) > 3 and sum(1 for l in lines[:3] if ";" in l) >= 2)
        )

        if not is_csv:
            print(f"⚠️  Salesbooking {year}: Kein CSV (Content-Type: {ct!r})")
            print(f"   Erste 300 Zeichen: {text[:300]!r}")
            return None

        print(f"✅  Salesbooking {year}: {len(resp.content):,} Bytes, "
              f"{len(lines)} Zeilen")
        # Erste Zeile (Header) zur Diagnose loggen
        print(f"   Header: {lines[0][:200]!r}")
        if len(lines) > 1:
            print(f"   Zeile1: {lines[1][:200]!r}")
        return text

    except Exception as exc:
        print(f"⚠️  Salesbooking {year}: Fehler – {exc}")
        return None


# ──────────────────────────────────────────────────────────────────
# CSV-Zusammenføhrung
# ──────────────────────────────────────────────────────────────────

def _parse_raw_csv(text: str) -> tuple[list[str], list[str], list[list[str]]]:
    """
    Parst den rohen CSV-Text.
    Gibt (header1, header2, data_rows) zurück.
    Erwartet 2 Kopfzeilen wie beim Journal-Export.
    """
    reader = csv.reader(io.StringIO(text), delimiter=";")
    rows = list(reader)
    if len(rows) < 3:
        return [], [], rows
    return rows[0], rows[1], rows[2:]


def _vorgang_id(row: list[str]) -> str:
    """
    Eindeutiger Schlüssel für Deduplizierung.

    Spalte 7 ("Vorgang") enthält den Buchungstyp ("Buchung", "Stornierung" …),
    KEINE eindeutige ID. Als Schlüssel verwenden wir deshalb:
        Objekt-Nr. + Anreise + Abreise
    Damit werden identische Buchungen (gleiche Unterkunft, gleicher Zeitraum)
    zuverlässig erkannt und nicht doppelt gezählt.
    """
    obj = row[0].strip() if len(row) > 0 else ""
    arr = row[4].strip() if len(row) > 4 else ""
    dep = row[5].strip() if len(row) > 5 else ""
    return f"{obj}_{arr}_{dep}"


def merge_csvs(journal_csv: str, salesbooking_results: dict[int, str]) -> str:
    """
    Führt Journal-CSV und Salesbooking-CSVs zusammen.

    - Journal-Einträge haben immer Vorrang (vollständigere Finanzdaten).
    - Aus Salesbooking werden nur Zeilen übernommen, deren Vorgangsnummer
      noch NICHT im Journal vorkommt.
    - Das Format des Journal-CSV wird beibehalten (2 Kopfzeilen).
    """
    header1, header2, journal_rows = _parse_raw_csv(journal_csv)

    # Alle Vorgangsnummern aus dem Journal merken
    seen_ids: set[str] = set()
    valid_journal_rows = []
    for row in journal_rows:
        if len(row) < 10:
            continue
        vid = _vorgang_id(row)
        seen_ids.add(vid)
        valid_journal_rows.append(row)

    print(f"📊  Journal: {len(valid_journal_rows)} Buchungen")

    new_rows: list[list[str]] = []
    for year, sb_csv in sorted(salesbooking_results.items()):
        if not sb_csv:
            continue
        sb_h1, sb_h2, sb_data = _parse_raw_csv(sb_csv)
        added = 0
        skipped_dup = 0
        skipped_short = 0

        for row in sb_data:
            if len(row) < 10:
                skipped_short += 1
                continue
            # Status-Check: Stornos øberspringen
            status = row[9].strip() if len(row) > 9 else ""
            if status in ("Storno", "Stornierung", "Stonierung", "Storniert", "cancelled"):
                continue
            vid = _vorgang_id(row)
            if vid in seen_ids:
                skipped_dup += 1
                continue
            # Zeile auf Journal-Spaltenbreite normalisieren
            if header2:
                target_len = len(header2)
                if len(row) < target_len:
                    row = row + [""] * (target_len - len(row))
                elif len(row) > target_len:
                    row = row[:target_len]
            new_rows.append(row)
            seen_ids.add(vid)
            added += 1

        print(f"📊  Salesbooking {year}: {added} neue Buchungen ergänzt "
              f"({skipped_dup} Duplikate, {skipped_short} ungültige Zeilen)")

    all_rows = valid_journal_rows + new_rows
    print(f"✅  Gesamt nach Merge: {len(all_rows)} Buchungen")

    # CSV rekonstruieren (mit den originalen 2 Kopfzeilen)
    output = io.StringIO()
    writer = csv.writer(output, delimiter=";", quoting=csv.QUOTE_MINIMAL,
                        lineterminator="\n")
    if header1:
        writer.writerow(header1)
    if header2:
        writer.writerow(header2)
    for row in all_rows:
        writer.writerow(row)

    return output.getvalue()


# ──────────────────────────────────────────────────────────────────
# Hauptprogramm
# ──────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Lädt Buchungs-CSV aus Contao (Journal + Salesbooking)"
    )
    parser.add_argument("--from", dest="date_from", default="01.01.2017",
                        help="Startdatum DD.MM.YYYY (default: 01.01.2017)")
    parser.add_argument("--to", dest="date_to",
                        default=(datetime.today() + timedelta(days=730)).strftime("%d.%m.%Y"),
                        help="Enddatum DD.MM.YYYY (default: heute + 2 Jahre)")
    parser.add_argument("--out", default=OUT_CSV,
                        help=f"Ausgabedatei (default: {OUT_CSV})")
    parser.add_argument("--no-salesbooking", action="store_true",
                        help="Salesbooking-Fetch überspringen (nur Journal)")
    args = parser.parse_args()

    if not USERNAME or not PASSWORD:
        print("❌  CONTAO_USER und CONTAO_PASS müssen gesetzt sein.")
        sys.exit(1)

    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (compatible; OstseeliebeBot/1.0)",
    })

    # 1. Login
    if not contao_login(session):
        sys.exit(1)

    # 2. Journal-CSV (abgerechnete Buchungen, gesamter Zeitraum)
    journal_csv = fetch_journal_csv(session, args.date_from, args.date_to)

    # 3. Salesbooking før aktuelles + nächstes Jahr (nicht abgerechnete Buchungen)
    salesbooking_results: dict[int, str] = {}
    if not args.no_salesbooking:
        current_year = datetime.today().year
        for year in [current_year, current_year + 1]:
            result = fetch_salesbooking_csv(session, year)
            if result:
                salesbooking_results[year] = result
    else:
        print("ℹ️  Salesbooking übersprungen (--no-salesbooking)")

    # 4. Zusammenführen
    if salesbooking_results:
        merged_csv = merge_csvs(journal_csv, salesbooking_results)
    else:
        print("ℹ️  Kein Salesbooking verføgbar – nur Journal-Daten")
        merged_csv = journal_csv

    # 5. Speichern
    out_path = Path(args.out)
    out_path.write_text(merged_csv, encoding="utf-8-sig")
    line_count = merged_csv.strip().count("\n")
    print(f"💾  Gespeichert: {out_path} ({line_count} Zeilen)")


if __name__ == "__main__":
    main()
