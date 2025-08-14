import io
import os
import time
import base64
import requests
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from datetime import datetime

# =========================
# CONFIG
# =========================
REPO_OWNER = "lbianco11197"
REPO_NAME  = "Avanzamento-economico"
BRANCH     = "main"
XLSX_PATH  = "Avanzamento.xlsx"      # nome e percorso esatto nel repo
SHEET_NAME = ""                      # "" => primo foglio
HOME_URL   = "https://homeeuroirte.streamlit.app/"
LOGO_PATH  = "LogoEuroirte.jpg"      # opzionale: presente nel repo dell'app

# Token opzionale (consigliato per evitare rate limit API)
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", os.getenv("GITHUB_TOKEN", None))

PAGE_TITLE = "Avanzamento mensile €/h per Tecnico - Euroirte s.r.l."

# Endpoint API GitHub (NO CDN)
API_URL  = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/{XLSX_PATH}"
COMMITS_API = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/commits"

# =========================
# PAGE SETUP & THEME
# =========================
st.set_page_config(layout="wide", page_title=PAGE_TITLE, page_icon=":bar_chart:")

st.markdown("""
<style>
:root { color-scheme: light !important; }
@media (prefers-color-scheme: dark){ :root { color-scheme: light !important; } }
html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"],
[data-testid="stHeader"], [data-testid="stSidebar"] {
  background:#fff !important; color:#000 !important;
}

/* Selectbox con bordo nero */
div[data-baseweb="select"] > div {
  border: 2px solid #000 !important;
  border-radius: 8px !important;
  background: #fff !important;
}
div[data-baseweb="select"] > div:hover,
div[data-baseweb="select"] > div:focus-within,
div[data-baseweb="select"][aria-expanded="true"] > div {
  border-color: #000 !important;
  box-shadow: 0 0 0 3px rgba(0,0,0,0.12) !important;
}
div[data-baseweb="select"] * { color:#000 !important; }
div[data-baseweb="select"] svg { stroke:#000 !important; fill:#000 !important; }

/* Bottoni */
.stButton > button{
  background:#fff !important; color:#000 !important;
  border:1px solid #999 !important; border-radius:8px; padding:.5rem .9rem;
}

/* Tabelle */
.stTable, .stDataFrame table, .stDataFrame th, .stDataFrame td { background:#fff !important; color:#000 !important; }
header [data-testid="theme-toggle"]{ display:none; }
</style>
""", unsafe_allow_html=True)

# =========================
# HEADER
# =========================
col_logo, col_title, col_btn = st.columns([0.12, 0.68, 0.20], vertical_alignment="center")
with col_logo:
    try:
        st.image(LOGO_PATH, use_container_width=True)
    except Exception:
        pass
with col_title:
    st.markdown(f"# 📊 {PAGE_TITLE}")
with col_btn:
    try:
        st.link_button("🏠 Torna alla Home", HOME_URL)
    except Exception:
        st.markdown(f"[🏠 Torna alla Home]({HOME_URL})")

st.divider()

# =========================
# CACHE HELPERS
# =========================
def _headers():
    h = {"Accept": "application/vnd.github+json"}
    if GITHUB_TOKEN:
        h["Authorization"] = f"token {GITHUB_TOKEN}"
    return h

@st.cache_data(show_spinner=True, ttl=60)
def fetch_excel_bytes_via_api():
    """
    Legge Avanzamento.xlsx direttamente dall'API GitHub (NO CDN).
    Ritorna (sha, bytes, last_modified_human).
    """
    # 1) Metadati + contenuto
    r = requests.get(API_URL, headers=_headers(), params={"ref": BRANCH}, timeout=30)
    r.raise_for_status()
    j = r.json()

    sha = j.get("sha") or str(int(time.time()))
    content = j.get("content")
    encoding = j.get("encoding")
    download_url = j.get("download_url")

    # 2) Data ultimo commit sul file
    r2 = requests.get(COMMITS_API, headers=_headers(),
                      params={"path": XLSX_PATH, "per_page": 1, "sha": BRANCH}, timeout=30)
    last_human = None
    if r2.ok:
        lst = r2.json()
        if isinstance(lst, list) and lst:
            iso = lst[0]["commit"]["committer"]["date"]
            try:
                dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
                last_human = dt.strftime("%d/%m/%Y %H:%M")
            except Exception:
                last_human = iso

    # 3) Decodifica (preferibilmente base64 dall'API)
    if content and encoding == "base64":
        data = base64.b64decode(content)
        return sha, data, last_human

    # Fallback: usa download_url (sempre API; niente raw CDN)
    if download_url:
        r3 = requests.get(download_url, timeout=30, headers={"Cache-Control": "no-cache", "Pragma": "no-cache"})
        r3.raise_for_status()
        return sha, r3.content, last_human

    raise RuntimeError("Impossibile ottenere il contenuto del file da GitHub.")

def load_avanzamento_df_from_bytes(xls_bytes: bytes) -> pd.DataFrame:
    """Parsa l'Excel (valori calcolati dalle formule) e normalizza le 3 colonne richieste."""
    bio = io.BytesIO(xls_bytes)
    wb = load_workbook(bio, data_only=True, read_only=True)
    ws = wb[SHEET_NAME] if SHEET_NAME and SHEET_NAME in wb.sheetnames else wb[wb.sheetnames[0]]

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return pd.DataFrame(columns=["Tecnico","Ore lavorate","Avanzamento €/h"])

    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    df = pd.DataFrame(rows[1:], columns=header)

    # Normalizzazione nomi colonne
    rename_map = {}
    for c in df.columns:
        cn = str(c).strip()
        key = cn.lower().replace(" ", "")
        if key == "tecnico":
            rename_map[c] = "Tecnico"
        elif key in ("orelavorate","orelavorate(h)","ore"):
            rename_map[c] = "Ore lavorate"
        elif key in ("avanzamento€/h","avanzamentoeuro/ora","avanzamentoeuroh","avanzamento"):
            rename_map[c] = "Avanzamento €/h"
    df = df.rename(columns=rename_map)

    # Colonne minime
    for needed in ["Tecnico", "Ore lavorate", "Avanzamento €/h"]:
        if needed not in df.columns:
            df[needed] = None

    # Tipi & pulizia
    df = df[["Tecnico","Ore lavorate","Avanzamento €/h"]].copy()
    df["Ore lavorate"] = pd.to_numeric(df["Ore lavorate"], errors="coerce")
    df["Avanzamento €/h"] = pd.to_numeric(df["Avanzamento €/h"], errors="coerce")
    df = df.dropna(how="all")
    if "Tecnico" in df.columns:
        df = df[df["Tecnico"].notna() & (df["Tecnico"].astype(str).str.strip() != "")]
    return df.reset_index(drop=True)

# Pulsante refresh manuale
cols = st.columns([0.2, 0.8])
with cols[0]:
    if st.button("🔄 Aggiorna dati"):
        st.cache_data.clear()

# Carica bytes via API e costruisci DataFrame
try:
    version_sha, xls_bytes, last_update_date = fetch_excel_bytes_via_api()
    df = load_avanzamento_df_from_bytes(xls_bytes)
except Exception as e:
    st.error(f"Errore nel caricamento del file da GitHub: {e}")
    st.stop()

# Mostra data ultimo aggiornamento (da commit)
if last_update_date:
    st.caption(f"📅 Dati aggiornati al {last_update_date}")

# =========================
# SELECTBOX (nessun default)
# =========================
PLACEHOLDER = "— Seleziona un tecnico —"
tecnici = sorted(df["Tecnico"].astype(str).dropna().unique().tolist())
options = [PLACEHOLDER] + tecnici
selezionato = st.selectbox("👷‍♂️ Seleziona un tecnico", options, index=0)

# =========================
# TABELLA CON LOGICA SEMAFORICA
# =========================
def color_semaforo(val):
    try:
        v = float(val)
    except (ValueError, TypeError):
        return ""
    if v < 25:
        return "background-color: #ff4d4d;"   # rosso
    elif 25 <= v <= 30:
        return "background-color: #ffff99;"   # giallo
    else:
        return "background-color: #b3ffb3;"   # verde

if selezionato != PLACEHOLDER:
    df_sel = df[df["Tecnico"].astype(str) == str(selezionato)][
        ["Tecnico", "Ore lavorate", "Avanzamento €/h"]
    ].copy()

    styler = (
        df_sel.style
        .format({
            "Ore lavorate": "{:.0f}",        # intero senza decimali
            "Avanzamento €/h": "€{:.2f}/h",
        })
        .applymap(color_semaforo, subset=["Avanzamento €/h"])
    )

    st.subheader("Dettaglio")
    st.table(styler)
else:
    st.info("Seleziona un tecnico per visualizzare il dettaglio.")
