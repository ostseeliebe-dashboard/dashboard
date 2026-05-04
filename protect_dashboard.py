#!/usr/bin/env python3
"""
protect_dashboard.py — Fügt einen Passwortschutz zur index.html hinzu.

Das Passwort wird als SHA-256-Hash direkt im HTML gespeichert.
Der Browser prüft die Eingabe clientseitig via Web Crypto API.
localStorage merkt sich die Anmeldung für 7 Tage.

Passwort: Ostsee2026
"""

import hashlib
import os

# Passwort aus Umgebungsvariable oder Standardwert
PASSWORD = os.environ.get("DASH_PASSWORD", "Ostsee2026")
PW_HASH  = hashlib.sha256(PASSWORD.encode()).hexdigest()

OVERLAY = f"""
<!-- ===== PASSWORTSCHUTZ ===== -->
<div id="pw-overlay" style="
    position:fixed; inset:0; z-index:99999;
    background:#f0f4f8;
    display:flex; align-items:center; justify-content:center;
    font-family:'Segoe UI',Arial,sans-serif;
">
  <div style="
      background:#fff; border-radius:16px;
      box-shadow:0 8px 40px rgba(0,0,0,0.13);
      padding:48px 56px; max-width:380px; width:90%;
      text-align:center;
  ">
    <div style="font-size:2.2rem; margin-bottom:8px;">🌊</div>
    <h2 style="margin:0 0 4px; font-size:1.3rem; color:#1a3a5c;">Ostseeliebe Dashboard</h2>
    <p style="margin:0 0 28px; color:#7a8fa6; font-size:0.9rem;">Bitte Passwort eingeben</p>
    <input id="pw-input" type="password"
        placeholder="Passwort"
        style="
            width:100%; box-sizing:border-box;
            padding:11px 16px; border:1.5px solid #dde3ea;
            border-radius:8px; font-size:1rem; outline:none;
            margin-bottom:12px; transition:border-color .2s;
        "
        onkeydown="if(event.key==='Enter') checkPw()"
        onfocus="this.style.borderColor='#3b82f6'"
        onblur="this.style.borderColor='#dde3ea'"
    />
    <div id="pw-error" style="
        color:#e53e3e; font-size:0.82rem;
        margin-bottom:8px; min-height:18px;
    "></div>
    <button onclick="checkPw()" style="
        width:100%; padding:11px; background:#1a3a5c;
        color:#fff; border:none; border-radius:8px;
        font-size:1rem; cursor:pointer; transition:background .2s;
    "
        onmouseover="this.style.background='#2563a0'"
        onmouseout="this.style.background='#1a3a5c'"
    >Anmelden</button>
  </div>
</div>
<script>
(function () {{
  const HASH = '{PW_HASH}';
  const KEY  = 'dash_auth';
  const TTL  = 7 * 24 * 60 * 60 * 1000; // 7 Tage

  // Bereits angemeldet?
  try {{
    const stored = JSON.parse(localStorage.getItem(KEY) || 'null');
    if (stored && stored.hash === HASH && Date.now() < stored.expires) {{
      document.getElementById('pw-overlay').style.display = 'none';
    }}
  }} catch(e) {{}}

  window.checkPw = async function () {{
    const input = document.getElementById('pw-input').value;
    const enc   = new TextEncoder().encode(input);
    const buf   = await crypto.subtle.digest('SHA-256', enc);
    const hex   = Array.from(new Uint8Array(buf))
                       .map(b => b.toString(16).padStart(2,'0')).join('');
    if (hex === HASH) {{
      localStorage.setItem(KEY, JSON.stringify({{
        hash: HASH,
        expires: Date.now() + TTL
      }}));
      document.getElementById('pw-overlay').style.display = 'none';
    }} else {{
      document.getElementById('pw-error').textContent = 'Falsches Passwort – bitte erneut versuchen.';
      document.getElementById('pw-input').value = '';
      document.getElementById('pw-input').focus();
    }}
  }};
}})();
</script>
<!-- ===== ENDE PASSWORTSCHUTZ ===== -->
"""


def main():
    html_path = "index.html"

    if not os.path.exists(html_path):
        print(f"⚠️  {html_path} nicht gefunden – überspringe Passwortschutz.")
        return

    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()

    # Alten Overlay entfernen (falls vorhanden)
    import re
    html = re.sub(
        r"<!-- ===== PASSWORTSCHUTZ =====.*?<!-- ===== ENDE PASSWORTSCHUTZ ===== -->",
        "",
        html,
        flags=re.DOTALL,
    )

    # Overlay nach <body> einfügen
    if "<body" in html:
        html = re.sub(r"(<body[^>]*>)", r"\1\n" + OVERLAY, html, count=1)
    else:
        html = OVERLAY + html

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅ Passwortschutz angewendet (Hash: {PW_HASH[:12]}…)")
    print(f"   Passwort: {PASSWORD}")


if __name__ == "__main__":
    main()
