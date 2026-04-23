#!/usr/bin/env python3
"""
fetch_contao_data.py
====================
Loggt sich in das Contao-Backend ein und lädt die Buchungs-CSV herunter.
Wird täglich von der GitHub Action ausgeführt.

Benötigt folgende Umgebungsvariablen (als GitHub Secrets hinterlegt):
    CONTAO_URL      z.B. https://www.ostseeliebe-ferienwohnungen.de
    CONTAO_USER     Benutzername für Contao-Login
    CONTAO_PASS     Passwort für Contao-Login

Aufruf:
    python fetch_contao_data.py
    python fetch_contao_data.py --from 2020-01-01 --to 2027-12-31
"""

import os
import sys
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
BASE_URL   = os.environ.get("CONTAO_URL",  "https://www.ostseeliebe-ferienwohnungen.de")
USERNAME   = os.environ.get("CONTAO_USER", "")
PASSWORD   = os.environ.get("CONTAO_PASS", "")
OUT_CSV    = "buchungen_export_2027.csv"   # Ausgabedatei (im Repo-Root)

# Felder die exportiert werden (für Dashboard relevant)
EXPORT_FIELDS = [
    # Pflichtfelder
    "objnr", "city", "arrival", "departure", "nights",
    "bookingid", "schannel", "totalrent", "paid", "rent_total",
    # Gebuehren
    "touristtax", "shortstay", "provisionfee", "cancelfee",
    # Leistungskategorien (aktiv)
    "group_1",   # 01. Endreinigung
    "group_2",   # 02. Zusatzleistungen (buchbar)
    "group_23",  # 03. Zusatzleistungen (nur Ostseeliebe)
    "group_26",  # 04. Zusatzleistungen (provisionsfaehig)
    "group_6",   # 05. Zusatzleistungen (inklusive)
    "group_9",   # 06. Nebenkosten (obligatorisch)
    "group_33",  # 07. Nebenkosten (teilw. Eigentuemer)
    "group_22",  # 08. Kaution
    "group_32",  # 09. Divers
]


def get_request_token(session: requests.Session, url: str) -> str:
    """Holt den REQUEST_TOKEN von einer Contao-Seite."""
    resp = session.get(url, timeout=30)
    resp.raise_for_status()
    match = re.search(r'name="REQUEST_TOKEN"\s+value="([^"]+)"', resp.text)
    if not match:
        # Fallback: Meta-Tag
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

    # Erfolg: wir landen auf dem Dashboard (kein Login-Formular mehr)
    if "do=main" in resp.url or "/contao" in resp.url and "login" not in resp.url:
        print("✅  Login erfolgreich")
        return True

    # Prüfen ob noch ein Login-Formular vorhanden
    if 'name="username"' in resp.text:
        print("❌  Login fehlgeschlagen – Benutzername/Passwort prüfen")
        return False

    print("✅  Login erfolgreich (Redirect erkannt)")
    return True


def fetch_buchungen_csv(session: requests.Session, date_from: str, date_to: str) -> str:
    """
    Lädt die Buchungs-CSV über die 'Individuelle Auswertungen'-Export-Form.

    date_from, date_to: Format 'DD.MM.YYYY'
    Gibt den CSV-Inhalt als String zurück.
    """
    stats_url = f"{BASE_URL}/contao?do=fewoOffice_stats&key=journal"
    print(f"📥  Exportiere Buchungen {date_from} – {date_to} …")

    token = get_request_token(session, stats_url)
    period = f"{date_from} - {date_to}"

    # Alle Export-Felder als Liste von Tupeln
    post_data = [
        ("FORM_ACTION",    "doExport"),
        ("REQUEST_TOKEN",  token),
        ("period",         period),
        ("selectorfield",  "arrival"),   # nach Anreisedatum filtern
        ("object",         ""),
        ("house",          ""),
        ("agent",          ""),
        ("owner",          ""),
        ("output",         "details"),
        ("schannel",       ""),
    ] + [("exportFields[]", f) for f in EXPORT_FIELDS]

    resp = session.post(stats_url, data=post_data, timeout=60)
    resp.raise_for_status()

    # Prüfen ob wir CSV oder HTML zurückbekommen haben
    content_type = resp.headers.get("content-type", "")
    if "text/csv" in content_type or "application/csv" in content_type:
        print(f"✅  CSV erhalten ({len(resp.content)} Bytes)")
        return resp.text

    # Manchmal liefert Contao die CSV als Download ohne Content-Type
    if resp.text.startswith('"') or resp.text.startswith("Objekt") or "\n" in resp.text[:100]:
        lines = resp.text.strip().split("\n")
        if len(lines) > 5:
            print(f"✅  CSV erhalten: {len(lines)} Zeilen")
            return resp.text

    # HTML-Response → Link zum Download suchen
    match = re.search(r'href="([^"]*\.csv[^"]*)"', resp.text)
    if match:
        csv_url = match.group(1)
        if not csv_url.startswith("http"):
            csv_url = BASE_URL + "/" + csv_url.lstrip("/")
        print(f"📎  Download-Link gefunden: {csv_url}")
        csv_resp = session.get(csv_url, timeout=60)
        csv_resp.raise_for_status()
        print(f"✅  CSV heruntergeladen: {len(csv_resp.content)} Bytes")
        return csv_resp.text

    # Fallback: gesamten Response-Text als CSV speichern und prüfen lassen
    print("⚠️   Unbekanntes Format – speichere Response-Text als CSV")
    print(f"   Content-Type: {content_type}")
    print(f"   Erste 200 Zeichen: {resp.text[:200]}")
    return resp.text


def main():
    parser = argparse.ArgumentParser(description="Lädt Buchungs-CSV aus Contao herunter")
    parser.add_argument("--from", dest="date_from", default="01.01.2017",
                        help="Startdatum DD.MM.YYYY (default: 01.01.2017)")
    parser.add_argument("--to", dest="date_to",
                        default=(datetime.today() + timedelta(days=730)).strftime("%d.%m.%Y"),
                        help="Enddatum DD.MM.YYYY (default: heute + 2 Jahre)")
    parser.add_argument("--out", default=OUT_CSV,
                        help=f"Ausgabedatei (default: {OUT_CSV})")
    args = parser.parse_args()

    if not USERNAME or not PASSWORD:
        print("❌  CONTAO_USER und CONTAO_PASS müssen als Umgebungsvariablen gesetzt sein.")
        print("   Lokal: export CONTAO_USER=dein_user && export CONTAO_PASS=dein_pass")
        print("   GitHub: als Repository Secrets hinterlegen")
        sys.exit(1)

    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (compatible; OstseeliebeBot/1.0)",
    })

    # Login
    if not contao_login(session):
        sys.exit(1)

    # CSV abrufen
    csv_content = fetch_buchungen_csv(session, args.date_from, args.date_to)

    # Speichern
    out_path = Path(args.out)
    out_path.write_text(csv_content, encoding="utf-8-sig")
    lines = csv_content.strip().count("\n")
    print(f"💾  Gespeichert: {out_path} ({lines} Zeilen)")


if __name__ == "__main__":
    main()
