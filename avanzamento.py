import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit.components.v1 import html
import io
import os
import requests
from openpyxl import load_workbook

st.set_page_config(layout="wide")

# Imposta sfondo bianco e testo nero
st.markdown("""
    <style>
    html, body, [data-testid="stApp"] {
        background-color: white !important;
        color: black !important;
    }

    /* Stile menù a tendina (selectbox) con bordo colorato */
div[data-baseweb="select"] > div {
  border: 2px solid #0d6efd !important;      /* colore bordo normale */
  border-radius: 8px !important;
  background: #fff !important;
}

div[data-baseweb="select"] > div:hover {
  border-color: #0b5ed7 !important;          /* bordo più scuro in hover */
}

div[data-baseweb="select"] > div:focus-within,
div[data-baseweb="select"][aria-expanded="true"] > div {
  border-color: #0a58ca !important;
  box-shadow: 0 0 0 3px rgba(13,110,253,0.15) !important;
}

div[data-baseweb="select"] * {
  color: #000 !important;
}

div[data-baseweb="select"] svg {
  stroke: #0d6efd !important;
  fill: #0d6efd !important;
}

    /* Forza stile nelle tabelle */
    .stDataFrame, .stDataFrame table, .stDataFrame th, .stDataFrame td {
        background-color: white !important;
        color: black !important;
    }

    /* Pulsanti */
    .stButton > button {
        background-color: white !important;
        color: black !important;
        border: 1px solid #999 !important;
        border-radius: 6px;
    }

    /* Radio button */
    div[data-baseweb="radio"] label span {
        color: black !important;
        font-weight: 600 !important;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <style>
        header [data-testid="theme-toggle"] {
            display: none;
        }
    </style>
""", unsafe_allow_html=True)

# --- Titolo ---
st.title("📊 Avanzamento mensile €/h per Tecnico - Euroirte s.r.l.")

# Intestazione con logo e bottone
# Logo in alto
st.image("LogoEuroirte.jpg", width=180)

# Bottone sotto il logo
st.link_button("🏠 Torna alla Home", url="https://homeeuroirte.streamlit.app/")

# ---------------------------
# CONFIG (modifica qui se serve)
# ---------------------------
REPO_OWNER   = os.getenv("AE_REPO_OWNER", "lbianco11197")
REPO_NAME    = os.getenv("AE_REPO_NAME",  "Avanzamento-economico")
BRANCH       = os.getenv("AE_BRANCH",     "main")
XLSX_PATH    = os.getenv("AE_XLSX_PATH",  "Avanzamento.xlsx")   # percorso nel repo
SHEET_NAME   = os.getenv("AE_SHEET_NAME", "")                   # "" = prima sheet
HOME_URL     = os.getenv("AE_HOME_URL",   "https://euroirte.it")# link per "Torna alla Home"
LOGO_URL     = os.getenv("AE_LOGO_URL",   "")                   # opzionale: URL raw del logo
# Se il repo è privato, imposta un token qui o in .streamlit/secrets.toml
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", os.getenv("GITHUB_TOKEN", None))



# ---------------------------
# DATA LOADING (valori calcolati da formule)
# ---------------------------
def raw_url(path: str) -> str:
    from urllib.parse import quote
    return f"https://raw.githubusercontent.com/{REPO_OWNER}/{REPO_NAME}/{BRANCH}/{quote(path)}"

@st.cache_data(show_spinner=True, ttl=600)
def load_avanzamento_df() -> pd.DataFrame:
    url = raw_url(XLSX_PATH)
    headers = {}
    if GITHUB_TOKEN:
        headers["Authorization"] = f"token {GITHUB_TOKEN}"
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()

    # Legge workbook con data_only=True per usare i valori calcolati delle formule
    bio = io.BytesIO(r.content)
    wb = load_workbook(bio, data_only=True, read_only=True)
    ws = wb[SHEET_NAME] if SHEET_NAME and SHEET_NAME in wb.sheetnames else wb[wb.sheetnames[0]]

    # Prima riga = intestazioni
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return pd.DataFrame(columns=["Tecnico","Ore lavorate","Avanzamento €/h"])

    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    data = rows[1:]
    df = pd.DataFrame(data, columns=header)

    # Normalizza nomi colonne attese
    rename_map = {}
    for c in df.columns:
        cn = str(c).strip()
        if cn.lower().replace(" ", "") in ("tecnico",):
            rename_map[c] = "Tecnico"
        elif cn.lower().replace(" ", "") in ("orelavorate","orelavorate(h)","ore"):
            rename_map[c] = "Ore lavorate"
        elif cn.lower().replace(" ", "") in ("avanzamento€/h","avanzamentoeuro/ora","avanzamentoeuroh","avanzamento"):
            rename_map[c] = "Avanzamento €/h"
    df = df.rename(columns=rename_map)

    # Tieni solo le 3 colonne richieste (se mancano, crearle vuote per evitare crash)
    for needed in ["Tecnico", "Ore lavorate", "Avanzamento €/h"]:
        if needed not in df.columns:
            df[needed] = None

    # Pulisci
    df = df[["Tecnico","Ore lavorate","Avanzamento €/h"]].copy()
    # cast numerici sicuri
    df["Ore lavorate"] = pd.to_numeric(df["Ore lavorate"], errors="coerce")
    df["Avanzamento €/h"] = pd.to_numeric(df["Avanzamento €/h"], errors="coerce")
    # Elimina righe completamente vuote
    df = df.dropna(how="all")
    # Rimuovi totali/sommari se presenti (opzionale: filtra dove Tecnico non è vuoto)
    if "Tecnico" in df.columns:
        df = df[df["Tecnico"].notna() & (df["Tecnico"].astype(str).str.strip() != "")]
    return df.reset_index(drop=True)

df = load_avanzamento_df()
if df.empty:
    st.warning("Nessun dato trovato in Avanzamento.xlsx.")
    st.stop()

# ---------------------------
# UI: seleziona tecnico
# ---------------------------
tecnici = sorted(df["Tecnico"].astype(str).dropna().unique().tolist())
selezionato = st.selectbox("👷‍♂️ Seleziona un tecnico", tecnici)

df_sel = df[df["Tecnico"].astype(str) == str(selezionato)].copy()

# Format e tabella
st.subheader("Dettaglio")
st.dataframe(
    df_sel.style.format({
        "Ore lavorate": "{:.2f}",
        "Avanzamento €/h": "€{:.2f}/h",
    }),
    use_container_width=True,
    hide_index=True
)

