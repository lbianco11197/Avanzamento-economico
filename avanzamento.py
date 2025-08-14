import io
import os
import requests
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# =========================
# CONFIG
# =========================
REPO_OWNER = "lbianco11197"
REPO_NAME  = "Avanzamento-economico"
BRANCH     = "main"
XLSX_PATH  = "Avanzamento.xlsx"
SHEET_NAME = ""                               # "" = primo foglio
HOME_URL   = "https://homeeuroirte.streamlit.app/"
LOGO_PATH  = "LogoEuroirte.jpg"               # se presente nel repo dell'app

# Se in Streamlit Cloud hai secrets, li usa (non necessari per repo pubblico)
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", os.getenv("GITHUB_TOKEN", None))

PAGE_TITLE = "Avanzamento mensile €/h per Tecnico - Euroirte s.r.l."

# --- Endpoints GitHub per versioning ---
API_URL = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/contents/{XLSX_PATH}"
RAW_BASE = "https://raw.githubusercontent.com"

# =========================
# PAGE SETUP & THEME
# =========================
st.set_page_config(layout="wide", page_title=PAGE_TITLE, page_icon=":bar_chart:")

# Tema chiaro + bordo nero selectbox
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
@st.cache_data(show_spinner=False, ttl=60)
def get_file_sha(ref: str = BRANCH) -> str:
    """Ritorna lo SHA corrente del file su GitHub (cambia ad ogni upload/commit)."""
    headers = {"Accept": "application/vnd.github+json"}
    if GITHUB_TOKEN:
        headers["Authorization"] = f"token {GITHUB_TOKEN}"
    r = requests.get(API_URL, headers=headers, params={"ref": ref}, timeout=30)
    r.raise_for_status()
    return r.json()["sha"]

@st.cache_data(show_spinner=True)
def load_avanzamento_df(version_sha: str) -> pd.DataFrame:
    """Scarica Avanzamento.xlsx versionato con ?v=<sha>, legge i valori calcolati (data_only)."""
    raw_url = f"{RAW_BASE}/{REPO_OWNER}/{REPO_NAME}/{BRANCH}/{XLSX_PATH}?v={version_sha}"
    headers = {"Cache-Control": "no-cache", "Pragma": "no-cache"}
    if GITHUB_TOKEN:
        headers["Authorization"] = f"token {GITHUB_TOKEN}"
    r = requests.get(raw_url, headers=headers, timeout=30)
    r.raise_for_status()

    bio = io.BytesIO(r.content)
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

# Pulsante per refresh manuale
cols = st.columns([0.2, 0.8])
with cols[0]:
    if st.button("🔄 Aggiorna dati"):
        st.cache_data.clear()

# Carica dati versionati
version_sha = get_file_sha()
df = load_avanzamento_df(version_sha)

# =========================
# SELECTBOX (senza default)
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
