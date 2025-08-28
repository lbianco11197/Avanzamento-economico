import io
import os
import time
import base64
import requests
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from datetime import datetime
from pathlib import Path

# =========================
# CONFIG
# =========================
REPO_OWNER = "lbianco11197"
REPO_NAME  = "Avanzamento-economico"
BRANCH     = "main"
XLSX_PATH  = "Avanzamento.xlsx"      # nome e percorso esatto nel repo
SHEET_NAME = ""                      # "" => primo foglio
HOME_URL   = "https://homeeuroirte.streamlit.app/"
LOGO_PATH  = "LogoEuroirte.png"      # opzionale: presente nel repo dell'app

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

# --- SFONDO FULL-SCREEN: funzione riutilizzabile ---
def set_page_background(image_path: str):
    p = Path(image_path)
    if not p.exists():
        # tentativo robusto: cerca accanto al file corrente
        alt = Path(__file__).parent / image_path
        if alt.exists():
            p = alt
        else:
            st.warning(f"Background non trovato: {image_path}")
            return

    encoded = base64.b64encode(p.read_bytes()).decode()
    css = f"""
    <style>
      [data-testid="stAppViewContainer"] {{
        background: url("data:image/png;base64,{encoded}") center/cover no-repeat fixed;
      }}
      [data-testid="stHeader"], [data-testid="stSidebar"] {{
        background-color: rgba(255,255,255,0.0) !important;
      }}
      html, body, [data-testid="stApp"] {{
        color: #0b1320 !important;
      }}
      .stDataFrame, .stTable, .stSelectbox div[data-baseweb="select"],
      .stTextInput, .stNumberInput, .stDateInput, .stMultiSelect,
      .stRadio, .stCheckbox, .stSlider, .stFileUploader, .stTextArea {{
        background-color: rgba(255,255,255,0.88) !important;
        border-radius: 10px;
        backdrop-filter: blur(0.5px);
      }}
      .stDataFrame table, .stDataFrame th, .stDataFrame td {{
        color: #0b1320 !important;
        background-color: rgba(255,255,255,0.0) !important;
      }}
      .stButton > button, .stDownloadButton > button, .stLinkButton > a {{
        background-color: #ffffff !important;
        color: #0b1320 !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px;
      }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)

# Imposta lo sfondo (metti il nome file che usi nel progetto)
set_page_background("sfondo.png")  # usa il PNG dello sfondo bianco/soft con glow

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
                dt = datetime.fromisoformat(iso.replace("Z", "+02:00"))
                last_human = dt.strftime("%d/%m/%Y")
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
    """
    Parsa l'Excel (valori calcolati dalle formule) e **mantiene** anche
    le colonne 'Mail' e 'Data aggiornamento' oltre a Tecnico, Ore lavorate, Avanzamento €/h.
    """
    bio = io.BytesIO(xls_bytes)
    wb = load_workbook(bio, data_only=True, read_only=True)
    ws = wb[SHEET_NAME] if SHEET_NAME and SHEET_NAME in wb.sheetnames else wb[wb.sheetnames[0]]

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return pd.DataFrame(columns=["Tecnico","Data aggiornamento","Ore lavorate","Avanzamento €/h","Mail"])

    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    df = pd.DataFrame(rows[1:], columns=header)

    # --- Normalizzazione nomi colonne (includiamo Mail e Data) ---
    rename_map = {}
    for c in df.columns:
        key = str(c).strip().lower().replace(" ", "")
        if key.startswith("tecnico"):
            rename_map[c] = "Tecnico"
        elif key.startswith("data"):
            rename_map[c] = "Data aggiornamento"
        elif key in ("orelavorate","ore","orelavorate(h)"):
            rename_map[c] = "Ore lavorate"
        elif key in ("avanzamento€/h","avanzamento","avanzamentoeuro/ora","avanzamentoeuroh"):
            rename_map[c] = "Avanzamento €/h"
        elif key in ("mail","email","e-mail"):
            rename_map[c] = "Mail"
    df = df.rename(columns=rename_map)

    # --- Colonne minime che vogliamo tenere ---
    wanted = ["Tecnico", "Data aggiornamento", "Ore lavorate", "Avanzamento €/h", "Mail"]
    for w in wanted:
        if w not in df.columns:
            df[w] = None

    # --- Tipi & pulizia ---
    df = df[wanted].copy()
    df["Ore lavorate"] = pd.to_numeric(df["Ore lavorate"], errors="coerce")
    df["Avanzamento €/h"] = pd.to_numeric(df["Avanzamento €/h"], errors="coerce")
    # Data in gg/mm/aaaa se presente
    df["Data aggiornamento"] = pd.to_datetime(df["Data aggiornamento"], errors="coerce").dt.strftime("%d/%m/%Y")

    # Pulizia Tecnico (come da tuo file)
    df = df.dropna(how="all")
    if "Tecnico" in df.columns:
        df = df[df["Tecnico"].notna() & (df["Tecnico"].astype(str).str.strip() != "")]
        df["Tecnico"] = (
            df["Tecnico"].astype(str)
            .str.strip()
            .str.replace(r"\s+", " ", regex=True)
            .str.upper()
        )
        df = df[df["Tecnico"] != ""]
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
# INVIO EMAIL MANUALE (STREAMLIT) — versione pulita
# =========================
import smtplib
from email.mime.text import MIMEText

SMTP_HOST = "mail.euroirte.it"
SMTP_PORT = 465
SMTP_USER = st.secrets["SMTP_USER"]
SMTP_PASS = st.secrets["SMTP_PASS"]
SMTP_FROM = st.secrets.get("SMTP_FROM", SMTP_USER)
MAIL_SUBJECT = "Aggiornamento settimanale"
DATA_RIF = last_update_date or ""

st.subheader("📧 Invio email personalizzate")

# Trova colonne
def _norm(s): 
    return str(s).replace("\u00a0"," ").strip().lower()

email_col   = next((c for c in df.columns if _norm(c) in {"mail","email","e-mail"}), None)
tecnico_col = next((c for c in df.columns if _norm(c) == "tecnico"), None)
ore_col     = next((c for c in df.columns if _norm(c) in {"ore lavorate","orelavorate","ore"}), None)
av_col      = next((c for c in df.columns if _norm(c).startswith("avanzamento")), None)
data_col    = next((c for c in df.columns if _norm(c).startswith("data")), None)

if email_col is None or tecnico_col is None or ore_col is None or av_col is None:
    st.error("Servono almeno le colonne: 'Tecnico', 'Ore lavorate', 'Avanzamento €/h' e 'Mail'.")
else:
    # Anteprima (prime 5)
    anteprima = []
    for _, r in df.head(5).iterrows():
        nome  = str(r.get(tecnico_col, "")).strip()
        # Rimuovi prefisso ID tipo IRTE0000001 -
        if len(nome) > 14:
            nome = nome[14:].strip()
        data  = str(r.get(data_col, DATA_RIF)) if data_col else DATA_RIF
        avanz = r.get(av_col, 0)
        ore   = r.get(ore_col, 0)
        to    = str(r.get(email_col, "")).strip()
        corpo = (
            f"Ciao {nome},\n\n"
            f"il tuo avanzamento economico aggiornato al {data} è di {avanz:.2f} €/h "
            f"e il totale delle ore lavorate è {ore:.0f}.\n"
        )
        anteprima.append({"Destinatario": to, "Messaggio": corpo})
    st.caption("Anteprima (prime 5 righe)")
    st.dataframe(pd.DataFrame(anteprima), use_container_width=True)

    if st.button("✉️ Invia email a tutti"):
        risultati = []
        try:
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
                server.login(SMTP_USER, SMTP_PASS)
                for _, r in df.iterrows():
                    to = str(r.get(email_col, "")).strip()
                    if not to:
                        risultati.append(("❌", "(manca Mail)"))
                        continue
                    nome  = str(r.get(tecnico_col, "")).strip()
                    if len(nome) > 14:
                        nome = nome[14:].strip()
                    data  = str(r.get(data_col, DATA_RIF)) if data_col else DATA_RIF
                    avanz = r.get(av_col, 0)
                    ore   = r.get(ore_col, 0)
                    corpo = (
                        f"Ciao {nome},\n\n"
                        f"il tuo avanzamento economico aggiornato al {data} è di {avanz:.2f} €/h "
                        f"e il totale delle ore lavorate è {ore:.0f}.\n"
                    )
                    msg = MIMEText(corpo, "plain", "utf-8")
                    msg["Subject"] = MAIL_SUBJECT
                    msg["From"] = SMTP_FROM
                    msg["To"] = to
                    server.send_message(msg)
                    risultati.append(("✅", to))
            st.success(f"Email inviate: {sum(1 for s,_ in risultati if s=='✅')}")
            st.write(pd.DataFrame(risultati, columns=["Stato","Destinatario"]))
        except Exception as e:
            st.error(f"Errore durante l'invio: {e}")

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
