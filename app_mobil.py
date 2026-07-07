"""
Ázsia Gastro B2B Rendelési Rendszer – MOBIL verzió
====================================================
Indítás: python app_mobil.py
Böngészőben: http://localhost:5001  (más port mint az asztali!)

Szükséges fájlok (ugyanabba a mappába):
  - Termekek_Tárhely.csv   (Kulcs-Soft termékexport)
  - partnerek.csv          (partner_kod;partner_nev;jelszo)
  - Kepek/                 (termékkép mappa, opcionális)
"""

from flask import Flask, render_template_string, request, send_file, session, redirect, url_for
import pandas as pd
import io
import os

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "azsiaGastroB2B2024xK9")

import urllib.request
import json

# ── Konfiguráció ──────────────────────────────────────────────────────────────
# A termékfájlt több néven/formátumban is keresi, ebben a sorrendben:
TERMEK_JELOLTEK = ["Termekek_Tárhely.csv", "Termekek_Tárhely.xlsx", "Termekek.xlsx"]
PARTNER_CSV = "partnerek.csv"
RAKTAR_NEV  = "Bandó"
PORT        = int(os.environ.get("PORT", 5001))
KOTELEZO_OSZLOPOK = ["Termék név", "Termék kód", "Mee.", "Nettó egységár",
                      "Vonalkód", "Cikkszám", "Kategória", "Részlegszám"]

# Képek: publikus GitHub repóból, jsDelivr CDN-en keresztül (gyors, nincs rate limit)
KEP_REPO_TULAJDONOS = "lajos90b-byte"
KEP_REPO_NEV        = "azsia-gastro-kepek"
KEP_REPO_AG         = "main"
KEP_ENGEDELYEZETT   = {"jpg", "jpeg", "png", "webp"}
# ─────────────────────────────────────────────────────────────────────────────


def _fajlok_kigyujtese(csomopont, gyujto):
    """jsDelivr fastruktúra rekurzív bejárása (README.md-t és almappákat is kezeli)."""
    for elem in csomopont:
        if elem.get("type") == "file":
            gyujto.append(elem.get("name", ""))
        elif elem.get("type") == "directory" and "files" in elem:
            _fajlok_kigyujtese(elem["files"], gyujto)


def kepek_terkep_betoltese():
    """Lekéri a képes repó teljes fájllistáját (előbb jsDelivr-től, ha az nem
    válaszol, tartalékként a GitHub API-tól), és termékkód → raw GitHub URL
    térképet épít belőle."""
    fajlnevek = []

    # 1. próbálkozás: jsDelivr (gyors CDN, jellemzően nincs korlátozás)
    try:
        url = f"https://data.jsdelivr.com/v1/packages/gh/{KEP_REPO_TULAJDONOS}/{KEP_REPO_NEV}@{KEP_REPO_AG}"
        req = urllib.request.Request(url, headers={"User-Agent": "azsia-gastro-b2b"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            adat = json.load(resp)
        _fajlok_kigyujtese(adat.get("files", []), fajlnevek)
    except Exception as e1:
        print(f"jsDelivr listázás sikertelen ({e1}), próbálkozás GitHub API-val...")
        # 2. próbálkozás: GitHub API (tartalék)
        try:
            url = (f"https://api.github.com/repos/{KEP_REPO_TULAJDONOS}/{KEP_REPO_NEV}"
                   f"/git/trees/{KEP_REPO_AG}?recursive=1")
            fejlecek = {"User-Agent": "azsia-gastro-b2b", "Accept": "application/vnd.github+json"}
            token = os.environ.get("GITHUB_TOKEN")
            if token:
                fejlecek["Authorization"] = f"Bearer {token}"
            req = urllib.request.Request(url, headers=fejlecek)
            with urllib.request.urlopen(req, timeout=15) as resp:
                adat = json.load(resp)
            fajlnevek = [elem.get("path", "") for elem in adat.get("tree", [])
                         if elem.get("type") == "blob"]
        except Exception as e2:
            print(f"FIGYELEM: a képek listázása mindkét módszerrel sikertelen "
                  f"(jsDelivr: {e1} / GitHub API: {e2}) - képek nélkül indul az app")
            return {}

    terkep = {}
    for nev in fajlnevek:
        if "." not in nev:
            continue
        kod, kiterjesztes = nev.rsplit(".", 1)
        if kiterjesztes.lower() in KEP_ENGEDELYEZETT:
            terkep.setdefault(
                kod,
                f"https://raw.githubusercontent.com/{KEP_REPO_TULAJDONOS}/{KEP_REPO_NEV}/{KEP_REPO_AG}/{nev}"
            )
    print(f"✓ {len(terkep)} kép betöltve a képes repóból")
    return terkep


KEPEK_TERKEP = kepek_terkep_betoltese()


def kep_url(termekkod):
    if not termekkod or str(termekkod).strip() in ("", "nan"):
        return ""
    return KEPEK_TERKEP.get(str(termekkod).strip(), "")


def termekfajl_beolvasasa():
    """Megkeresi az első létező termékfájlt a TERMEK_JELOLTEK listából,
    és beolvassa (cp1250 csv vagy xlsx). Visszaadja a fájlnevet is."""
    for nev in TERMEK_JELOLTEK:
        if os.path.exists(nev):
            if nev.lower().endswith(".csv"):
                df = pd.read_csv(nev, encoding="cp1250", sep=";", dtype=str)
            else:
                df = pd.read_excel(nev, dtype=str)
            return df, nev
    keresett = ", ".join(f"'{n}'" for n in TERMEK_JELOLTEK)
    raise FileNotFoundError(f"Egyik termékfájlt sem találom ezek közül: {keresett}")


def termekek_betoltese():
    df, forras_nev = termekfajl_beolvasasa()

    hianyzo = [o for o in KOTELEZO_OSZLOPOK if o not in df.columns]
    if hianyzo:
        raise ValueError(
            f"A(z) '{forras_nev}' fájlból hiányoznak ezek az oszlopok: {', '.join(hianyzo)}. "
            f"Valószínűleg rossz exportot használsz - a 'tárhelyes' (Részlegszám oszlopot "
            f"tartalmazó) Kulcs-Soft terméklistát kell exportálni, nem az általános törzsadat exportot."
        )

    df = df[df["Részlegszám"].notna() & (df["Részlegszám"].str.strip() != "")]
    df["Nettó egységár"] = pd.to_numeric(df["Nettó egységár"], errors="coerce").fillna(0).astype(int)
    df["Kategória_rövid"] = df["Kategória"].str.replace("Cikkcsoport/", "", regex=False).str.strip()

    def vonalkod_tisztit(v):
        if pd.isna(v) or str(v).strip() == "":
            return ""
        v = str(v).replace(",", ".")
        try:
            return str(int(float(v)))
        except:
            return str(v).strip()

    df["Vonalkód"] = df["Vonalkód"].apply(vonalkod_tisztit)
    df["Termék kód"] = df["Termék kód"].fillna("").astype(str).str.strip()
    df["kep_url"] = df["Termék kód"].apply(kep_url)
    print(f"  (forrás: {forras_nev})")
    return df.reset_index(drop=True)


def partnerek_betoltese():
    df = pd.read_csv(PARTNER_CSV, encoding="utf-8", sep=";", dtype=str)
    partnerek = {}
    for _, sor in df.iterrows():
        kod = str(sor["partner_kod"]).strip()
        partnerek[kod] = {
            "nev":    str(sor["partner_nev"]).strip(),
            "jelszo": str(sor["jelszo"]).strip()
        }
    return partnerek


TERMEK_HIBA = None
try:
    TERMEKEK = termekek_betoltese()
    print(f"OK: {len(TERMEKEK)} termek betoltve")
except Exception as e:
    TERMEK_HIBA = str(e)
    print(f"HIBA: {TERMEK_HIBA}")
    TERMEKEK = pd.DataFrame()

try:
    PARTNEREK = partnerek_betoltese()
    print(f"✓ {len(PARTNEREK)} partner betöltve")
except FileNotFoundError:
    print(f"HIBA: Nem találom: '{PARTNER_CSV}'")
    PARTNEREK = {}


# ══════════════════════════════════════════════════════════════════════════════
# LOGIN SABLON
# ══════════════════════════════════════════════════════════════════════════════

LOGIN_HTML = """
<!DOCTYPE html>
<html lang="hu">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
  <title>Ázsia Gastro – Belépés</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; -webkit-tap-highlight-color: transparent; }
    html, body {
      height: 100%;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
      background: #1a3a2a;
    }
    .wrap {
      min-height: 100%;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 24px 20px;
    }
    .logo-ikon { font-size: 48px; margin-bottom: 12px; }
    h1 { color: white; font-size: 22px; font-weight: 600; margin-bottom: 4px; }
    .alcim { color: rgba(255,255,255,0.5); font-size: 14px; margin-bottom: 36px; }

    .card {
      background: white;
      border-radius: 20px;
      padding: 28px 24px;
      width: 100%;
      max-width: 400px;
    }
    label { display: block; font-size: 12px; color: #666; margin-bottom: 6px; margin-top: 18px; }
    label:first-child { margin-top: 0; }
    input[type=text], input[type=password] {
      width: 100%;
      padding: 14px 16px;
      border: 1.5px solid #e0e0e0;
      border-radius: 12px;
      font-size: 16px;
      color: #333;
      outline: none;
      transition: border-color 0.2s;
    }
    input:focus { border-color: #1a3a2a; }
    .btn {
      display: block;
      width: 100%;
      margin-top: 24px;
      padding: 16px;
      background: #1a3a2a;
      color: white;
      border: none;
      border-radius: 14px;
      font-size: 16px;
      font-weight: 600;
      cursor: pointer;
      text-align: center;
    }
    .btn:active { background: #2d6248; }
    .hiba {
      background: #fff0f0;
      border: 1px solid #fcc;
      color: #cc0000;
      border-radius: 10px;
      padding: 12px 14px;
      font-size: 14px;
      margin-top: 16px;
    }
    .hint { color: rgba(255,255,255,0.35); font-size: 12px; margin-top: 20px; text-align: center; }
  </style>
</head>
<body>
<div class="wrap">
  <div class="logo-ikon">🥢</div>
  <h1>Ázsia Gastro</h1>
  <p class="alcim">B2B Rendelési Rendszer</p>
  <div class="card">
    <form method="POST" action="/belepes">
      <label>Partner kód</label>
      <input type="text" name="partner_kod" placeholder="pl. FISH BOX"
             autocomplete="username" autocapitalize="off" required>
      <label>Jelszó</label>
      <input type="password" name="jelszo" placeholder="••••••••"
             autocomplete="current-password" required>
      <button type="submit" class="btn">Belépés →</button>
      {% if hiba %}<div class="hiba">{{ hiba }}</div>{% endif %}
    </form>
  </div>
  <p class="hint">Hozzáférési problémánál hívja értékesítőjét.</p>
</div>
</body>
</html>
"""


# ══════════════════════════════════════════════════════════════════════════════
# FŐOLDAL SABLON (kártya alapú, mobilra)
# ══════════════════════════════════════════════════════════════════════════════

FOOLDAL_HTML = """
<!DOCTYPE html>
<html lang="hu">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
  <title>Ázsia Gastro – Rendelés</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; -webkit-tap-highlight-color: transparent; }
    html { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif; }
    body { background: #f2f2f7; font-size: 14px; color: #1c1c1e; }

    /* ── HEADER ── */
    .hdr {
      background: #1a3a2a;
      color: white;
      padding: 14px 16px 12px;
      position: sticky;
      top: 0;
      z-index: 100;
    }
    .hdr-sor1 {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 10px;
    }
    .partner-badge {
      background: rgba(255,255,255,0.15);
      border-radius: 20px;
      padding: 4px 12px;
      font-size: 13px;
    }
    .kilepes-btn {
      color: rgba(255,255,255,0.6);
      font-size: 13px;
      text-decoration: none;
    }
    .search-wrap {
      background: rgba(255,255,255,0.12);
      border-radius: 12px;
      display: flex;
      align-items: center;
      padding: 10px 14px;
      gap: 8px;
    }
    .search-wrap svg { opacity: 0.7; flex-shrink: 0; }
    .search-wrap input {
      background: none;
      border: none;
      outline: none;
      color: white;
      font-size: 15px;
      width: 100%;
    }
    .search-wrap input::placeholder { color: rgba(255,255,255,0.5); }

    /* ── KATEGÓRIA CSÍK ── */
    .kat-csik {
      display: flex;
      gap: 8px;
      padding: 10px 16px;
      overflow-x: auto;
      scrollbar-width: none;
      background: white;
      border-bottom: 1px solid #e5e5ea;
      position: sticky;
      top: 94px;
      z-index: 99;
    }
    .kat-csik::-webkit-scrollbar { display: none; }
    .kat-chip {
      background: #f2f2f7;
      border: 1px solid #e5e5ea;
      border-radius: 20px;
      padding: 6px 14px;
      font-size: 12px;
      white-space: nowrap;
      color: #3a3a3c;
      cursor: pointer;
      user-select: none;
      flex-shrink: 0;
    }
    .kat-chip.aktiv {
      background: #1a3a2a;
      border-color: #1a3a2a;
      color: white;
    }

    /* ── TERMÉKLISTA ── */
    .termek-lista {
      padding: 12px 12px 120px;
      display: flex;
      flex-direction: column;
      gap: 8px;
    }
    .termek-card {
      background: white;
      border-radius: 14px;
      padding: 12px;
      display: flex;
      align-items: center;
      gap: 12px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    .termek-kep {
      width: 60px;
      height: 60px;
      border-radius: 10px;
      border: 1px solid #f0f0f0;
      object-fit: contain;
      flex-shrink: 0;
      background: #fafafa;
    }
    .termek-placeholder {
      width: 60px;
      height: 60px;
      border-radius: 10px;
      border: 1px solid #f0f0f0;
      background: #fafafa;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 26px;
      flex-shrink: 0;
    }
    .termek-info { flex: 1; min-width: 0; }
    .termek-nev {
      font-size: 13px;
      font-weight: 500;
      line-height: 1.35;
      color: #1c1c1e;
      display: -webkit-box;
      -webkit-line-clamp: 2;
      -webkit-box-orient: vertical;
      overflow: hidden;
    }
    .termek-kat {
      font-size: 11px;
      color: #8e8e93;
      margin-top: 3px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .termek-ar {
      font-size: 14px;
      font-weight: 600;
      color: #1a3a2a;
      margin-top: 5px;
    }
    .termek-jobb { display: flex; flex-direction: column; align-items: center; gap: 6px; flex-shrink: 0; }
    .add-btn {
      width: 36px;
      height: 36px;
      background: #1a3a2a;
      border-radius: 10px;
      border: none;
      color: white;
      font-size: 22px;
      display: flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      font-weight: 300;
    }
    .add-btn:active { background: #2d6248; transform: scale(0.95); }
    .kosarban-szam {
      font-size: 12px;
      font-weight: 600;
      color: #c8a000;
      min-height: 16px;
    }

    /* ── KOSÁR SÁTOR (alul fix) ── */
    .kosar-sor {
      position: fixed;
      bottom: 0;
      left: 0;
      right: 0;
      background: #c8a000;
      padding: 14px 16px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      z-index: 200;
      display: none;
    }
    .kosar-sor.latszik { display: flex; }
    .kosar-sor-bal { color: white; }
    .kosar-sor-bal .cim { font-size: 15px; font-weight: 600; }
    .kosar-sor-bal .reszlet { font-size: 12px; opacity: 0.85; margin-top: 2px; }
    .kosar-sor-gomb {
      background: white;
      color: #c8a000;
      border: none;
      border-radius: 12px;
      padding: 10px 20px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
    }
    .kosar-sor-gomb:active { transform: scale(0.97); }

    /* ── KOSÁR OLDAL (overlay) ── */
    .kosar-overlay {
      position: fixed;
      inset: 0;
      background: rgba(0,0,0,0.45);
      z-index: 300;
      display: none;
    }
    .kosar-overlay.latszik { display: block; }
    .kosar-panel {
      position: absolute;
      bottom: 0;
      left: 0;
      right: 0;
      background: #f2f2f7;
      border-radius: 20px 20px 0 0;
      max-height: 90vh;
      display: flex;
      flex-direction: column;
      overflow: hidden;
    }
    .kosar-hdr {
      background: #1a3a2a;
      color: white;
      padding: 16px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-shrink: 0;
    }
    .kosar-hdr h2 { font-size: 17px; font-weight: 600; }
    .kilepes-x {
      background: rgba(255,255,255,0.15);
      border: none;
      color: white;
      font-size: 18px;
      width: 32px;
      height: 32px;
      border-radius: 50%;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    .kosar-tartalom {
      flex: 1;
      overflow-y: auto;
      padding: 12px;
      display: flex;
      flex-direction: column;
      gap: 8px;
    }
    .kosar-item {
      background: white;
      border-radius: 14px;
      padding: 12px;
      display: flex;
      align-items: center;
      gap: 10px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    .k-kep {
      width: 46px;
      height: 46px;
      border-radius: 8px;
      border: 1px solid #f0f0f0;
      object-fit: contain;
      background: #fafafa;
      flex-shrink: 0;
    }
    .k-placeholder {
      width: 46px;
      height: 46px;
      border-radius: 8px;
      border: 1px solid #f0f0f0;
      background: #fafafa;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 20px;
      flex-shrink: 0;
    }
    .k-info { flex: 1; min-width: 0; }
    .k-nev { font-size: 13px; font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .k-ar { font-size: 12px; color: #8e8e93; margin-top: 2px; }
    .k-menny {
      display: flex;
      align-items: center;
      gap: 10px;
      flex-shrink: 0;
    }
    .k-btn {
      width: 30px;
      height: 30px;
      background: #f2f2f7;
      border: none;
      border-radius: 8px;
      font-size: 18px;
      color: #1c1c1e;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    .k-btn:active { background: #e5e5ea; }
    .k-szam { font-size: 15px; font-weight: 600; min-width: 20px; text-align: center; }

    .kosar-alap {
      background: white;
      border-radius: 16px;
      margin: 0 12px 12px;
      padding: 16px;
      flex-shrink: 0;
    }
    .megjegyzes-label { font-size: 12px; color: #8e8e93; margin-bottom: 6px; }
    .megjegyzes-input {
      width: 100%;
      border: 1.5px solid #e5e5ea;
      border-radius: 10px;
      padding: 10px 12px;
      font-size: 14px;
      font-family: inherit;
      resize: none;
      outline: none;
      margin-bottom: 14px;
      color: #1c1c1e;
    }
    .megjegyzes-input:focus { border-color: #1a3a2a; }
    .ossz-sor {
      display: flex;
      justify-content: space-between;
      font-size: 13px;
      color: #8e8e93;
      margin-bottom: 5px;
    }
    .ossz-sor.total {
      font-size: 16px;
      font-weight: 600;
      color: #1c1c1e;
      margin-bottom: 14px;
    }
    .letolt-btn {
      width: 100%;
      background: #c8a000;
      color: white;
      border: none;
      border-radius: 14px;
      padding: 16px;
      font-size: 16px;
      font-weight: 600;
      cursor: pointer;
    }
    .letolt-btn:active { background: #a88000; }

    /* ── TOAST ── */
    .toast {
      position: fixed;
      bottom: 80px;
      left: 50%;
      transform: translateX(-50%);
      background: #1c1c1e;
      color: white;
      padding: 10px 20px;
      border-radius: 20px;
      font-size: 13px;
      white-space: nowrap;
      z-index: 500;
      opacity: 0;
      transition: opacity 0.2s;
      pointer-events: none;
    }
    .toast.latszik { opacity: 1; }

    .ures-uzenet { text-align: center; color: #8e8e93; padding: 40px 20px; font-size: 14px; }
  </style>
</head>
<body>

<!-- HEADER -->
<div class="hdr">
  <div class="hdr-sor1">
    <span class="partner-badge">👤 {{ partner_nev }}</span>
    <a href="/kilepes" class="kilepes-btn">Kilépés</a>
  </div>
  <div class="search-wrap">
    <svg width="16" height="16" fill="none" viewBox="0 0 24 24" stroke="white" stroke-width="2">
      <circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/>
    </svg>
    <input type="text" id="kereses" placeholder="Keresés termékben, cikkszámban..." oninput="szur()">
  </div>
</div>

<!-- KATEGÓRIA CSÍK -->
<div class="kat-csik" id="katCsik">
  <div class="kat-chip aktiv" onclick="katValaszt(this, '')">Mind</div>
  {% for kat in kategoriak %}
  <div class="kat-chip" onclick="katValaszt(this, '{{ kat }}')">{{ kat }}</div>
  {% endfor %}
</div>

<!-- TERMÉKLISTA -->
<div class="termek-lista" id="termekLista">
  {% for t in termekek %}
  <div class="termek-card"
    data-cikkszam="{{ t.Cikkszám }}"
    data-vonalkod="{{ t.Vonalkód }}"
    data-nev="{{ t['Termék név'] }}"
    data-me="{{ t['Mee.'] }}"
    data-ar="{{ t['Nettó egységár'] }}"
    data-tarhely="{{ t.Részlegszám }}"
    data-kat="{{ t.Kategória_rövid }}"
    data-termekkod="{{ t['Termék kód'] }}"
    data-kepurl="{{ t.kep_url }}"
  >
    {% if t.kep_url %}
    <img class="termek-kep" src="{{ t.kep_url }}" alt="{{ t['Termék név'] }}" loading="lazy">
    {% else %}
    <div class="termek-placeholder">🥢</div>
    {% endif %}
    <div class="termek-info">
      <div class="termek-nev">{{ t['Termék név'] }}</div>
      <div class="termek-kat">{{ t.Kategória_rövid }} · {{ t.Cikkszám }}</div>
      <div class="termek-ar">{{ "{:,}".format(t['Nettó egységár']).replace(",", " ") }} Ft</div>
    </div>
    <div class="termek-jobb">
      <button class="add-btn" onclick="hozzaad(this)">+</button>
      <div class="kosarban-szam" id="db_{{ t.Cikkszám }}"></div>
    </div>
  </div>
  {% endfor %}
</div>

<!-- KOSÁR SÁTOR (alul fix) -->
<div class="kosar-sor" id="kosarSor">
  <div class="kosar-sor-bal">
    <div class="cim">🛒 <span id="kosarCimSzam">0</span> tétel</div>
    <div class="reszlet" id="kosarCimOssz"></div>
  </div>
  <button class="kosar-sor-gomb" onclick="kosarMegnyit()">Kosár megtekintése</button>
</div>

<!-- KOSÁR OVERLAY -->
<div class="kosar-overlay" id="kosarOverlay" onclick="kosarBezar()">
  <div class="kosar-panel" onclick="event.stopPropagation()">
    <div class="kosar-hdr">
      <h2>Kosár</h2>
      <button class="kilepes-x" onclick="kosarBezar()">✕</button>
    </div>
    <div class="kosar-tartalom" id="kosarTartalom"></div>
    <div class="kosar-alap">
      <div class="megjegyzes-label">Megjegyzés (opcionális)</div>
      <textarea class="megjegyzes-input" id="megjegyzes" rows="2"
                placeholder="pl. szállítási instrukciók..."></textarea>
      <div class="ossz-sor"><span>Tételek száma</span><span id="ossz-db">0 db</span></div>
      <div class="ossz-sor total"><span>Összesen (nettó)</span><span id="ossz-ar">0 Ft</span></div>
      <button class="letolt-btn" onclick="letolt()">⬇ Import CSV letöltése</button>
    </div>
  </div>
</div>

<!-- TOAST -->
<div class="toast" id="toast"></div>

<script>
let kosar = {};
let aktKat = "";

function katValaszt(el, kat) {
  aktKat = kat;
  document.querySelectorAll(".kat-chip").forEach(c => c.classList.remove("aktiv"));
  el.classList.add("aktiv");
  szur();
}

function szur() {
  const q   = document.getElementById("kereses").value.toLowerCase().trim();
  const kartok = document.querySelectorAll(".termek-card");
  kartok.forEach(k => {
    const nev  = k.dataset.nev.toLowerCase();
    const cikk = k.dataset.cikkszam.toLowerCase();
    const kat  = k.dataset.kat;
    const nevOk = !q || nev.includes(q) || cikk.includes(q);
    const katOk = !aktKat || kat === aktKat;
    k.style.display = (nevOk && katOk) ? "flex" : "none";
  });
}

function hozzaad(btn) {
  const k = btn.closest(".termek-card");
  const c = k.dataset.cikkszam;
  if (kosar[c]) {
    kosar[c].menny++;
  } else {
    kosar[c] = {
      nev:      k.dataset.nev,
      me:       k.dataset.me,
      ar:       parseInt(k.dataset.ar),
      vonalkod: k.dataset.vonalkod,
      tarhely:  k.dataset.tarhely,
      kepurl:   k.dataset.kepurl,
      menny:    1
    };
  }
  dbFrissit(c);
  kosarSatorFrissit();
  toast("✓ " + k.dataset.nev.substring(0, 32) + (k.dataset.nev.length > 32 ? "…" : ""));
}

function novel(c) { kosar[c].menny++; dbFrissit(c); kosarFrissit(); kosarSatorFrissit(); }
function csokken(c) {
  kosar[c].menny--;
  if (kosar[c].menny <= 0) { delete kosar[c]; }
  dbFrissit(c);
  kosarFrissit();
  kosarSatorFrissit();
}

function dbFrissit(c) {
  const el = document.getElementById("db_" + c);
  if (!el) return;
  el.textContent = kosar[c] ? kosar[c].menny + " db" : "";
}

function kosarSatorFrissit() {
  const tetelek = Object.entries(kosar);
  const sor = document.getElementById("kosarSor");
  if (!tetelek.length) { sor.classList.remove("latszik"); return; }
  sor.classList.add("latszik");
  const ossz = tetelek.reduce((s, [, t]) => s + t.ar * t.menny, 0);
  document.getElementById("kosarCimSzam").textContent = tetelek.length;
  document.getElementById("kosarCimOssz").textContent = fmt(ossz) + " összesen";
}

function kosarMegnyit() {
  kosarFrissit();
  document.getElementById("kosarOverlay").classList.add("latszik");
  document.body.style.overflow = "hidden";
}
function kosarBezar() {
  document.getElementById("kosarOverlay").classList.remove("latszik");
  document.body.style.overflow = "";
}

function kosarFrissit() {
  const tetelek = Object.entries(kosar);
  const tartalom = document.getElementById("kosarTartalom");
  if (!tetelek.length) {
    tartalom.innerHTML = '<div class="ures-uzenet">A kosár üres.<br>Adj hozzá termékeket!</div>';
    document.getElementById("ossz-db").textContent = "0 db";
    document.getElementById("ossz-ar").textContent = "0 Ft";
    return;
  }
  let ossz = 0, html = "";
  tetelek.forEach(([c, t]) => {
    ossz += t.ar * t.menny;
    const kepHtml = t.kepurl
      ? `<img class="k-kep" src="${t.kepurl}" alt="">`
      : `<div class="k-placeholder">🥢</div>`;
    html += `<div class="kosar-item">
      ${kepHtml}
      <div class="k-info">
        <div class="k-nev">${t.nev.substring(0, 40)}${t.nev.length > 40 ? "…" : ""}</div>
        <div class="k-ar">${fmt(t.ar)} / ${t.me}</div>
      </div>
      <div class="k-menny">
        <button class="k-btn" onclick="csokken('${c}')">−</button>
        <span class="k-szam">${t.menny}</span>
        <button class="k-btn" onclick="novel('${c}')">+</button>
      </div>
    </div>`;
  });
  tartalom.innerHTML = html;
  document.getElementById("ossz-db").textContent = tetelek.length + " tétel";
  document.getElementById("ossz-ar").textContent = fmt(ossz);
}

function letolt() {
  const tetelek = Object.entries(kosar);
  if (!tetelek.length) { toast("A kosár üres!"); return; }
  const sorok = tetelek.map(([c, t]) =>
    [c, t.vonalkod, t.nev, "{{ raktar_nev }}", "", "", "", t.tarhely, t.me, t.menny, t.ar, 0].join(";")
  );
  fetch("/letoltes", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ sorok, megjegyzes: document.getElementById("megjegyzes").value.trim() })
  })
  .then(r => r.blob())
  .then(blob => {
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    const datum = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = "rendeles_{{ partner_kod }}_" + datum + ".csv";
    a.click();
    URL.revokeObjectURL(url);
    kosarBezar();
    toast("✓ CSV letöltve – importálható a Kulcs-Softba!");
  });
}

function fmt(n) { return n.toLocaleString("hu-HU") + " Ft"; }
function toast(uzenet) {
  const el = document.getElementById("toast");
  el.textContent = uzenet;
  el.classList.add("latszik");
  clearTimeout(el._t);
  el._t = setTimeout(() => el.classList.remove("latszik"), 2500);
}
</script>
</body>
</html>
"""


# ══════════════════════════════════════════════════════════════════════════════
# ÚTVONALAK
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/")
def fooldal():
    if "partner_kod" not in session:
        return redirect(url_for("belepes_oldal"))
    partner_kod = session["partner_kod"]
    partner_nev = PARTNEREK.get(partner_kod, {}).get("nev", partner_kod)
    if TERMEKEK.empty:
        return f"<h2>Hiba a termekfajl betoltesekor:</h2><p>{TERMEK_HIBA or 'Ismeretlen hiba.'}</p>"
    kategoriak = sorted(TERMEKEK["Kategória_rövid"].dropna().unique().tolist())
    termekek   = TERMEKEK.to_dict(orient="records")
    return render_template_string(
        FOOLDAL_HTML,
        termekek=termekek,
        kategoriak=kategoriak,
        partner_nev=partner_nev,
        partner_kod=partner_kod,
        raktar_nev=RAKTAR_NEV
    )


@app.route("/belepes", methods=["GET", "POST"])
def belepes_oldal():
    hiba = None
    if request.method == "POST":
        kod    = request.form.get("partner_kod", "").strip()
        jelszo = request.form.get("jelszo", "").strip()
        partner = PARTNEREK.get(kod)
        if partner and partner["jelszo"] == jelszo:
            session["partner_kod"] = kod
            return redirect(url_for("fooldal"))
        else:
            hiba = "Hibás partner kód vagy jelszó."
    return render_template_string(LOGIN_HTML, hiba=hiba)


@app.route("/kilepes")
def kilepes():
    session.clear()
    return redirect(url_for("belepes_oldal"))


@app.route("/letoltes", methods=["POST"])
def letoltes():
    if "partner_kod" not in session:
        return "Nincs bejelentkezve", 403
    adatok = request.get_json()
    sorok  = adatok.get("sorok", [])
    csv_szoveg = "\r\n".join(sorok)
    try:
        csv_bytes = csv_szoveg.encode("cp1250")
    except UnicodeEncodeError:
        csv_bytes = csv_szoveg.encode("utf-8")
    return send_file(
        io.BytesIO(csv_bytes),
        mimetype="text/csv",
        as_attachment=True,
        download_name="rendeles.csv"
    )


@app.route("/frissit")
def frissit():
    global TERMEKEK, PARTNEREK, TERMEK_HIBA, KEPEK_TERKEP
    try:
        KEPEK_TERKEP = kepek_terkep_betoltese()
        TERMEKEK    = termekek_betoltese()
        TERMEK_HIBA = None
        PARTNEREK   = partnerek_betoltese()
        return f"OK: {len(TERMEKEK)} termek, {len(PARTNEREK)} partner es {len(KEPEK_TERKEP)} kep ujratoltve."
    except Exception as e:
        TERMEK_HIBA = str(e)
        return f"Hiba: {e}"


# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print(f"\n{'='*50}")
    print(f"  Ázsia Gastro B2B – MOBIL verzió")
    print(f"{'='*50}")
    print(f"  Böngészőben: http://localhost:{PORT}")
    print(f"  Leállítás:   Ctrl+C")
    print(f"{'='*50}\n")
    app.run(debug=False, port=PORT, host="0.0.0.0")
