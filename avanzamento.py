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
import smtplib
from email.mime.text import MIMEText
import re

# =========================
# CONFIG
# =========================
REPO_OWNER = "lbianco11197"
REPO_NAME  = "Avanzamento-economico"
BRANCH     = "main"
XLSX_PATH  = "Avanzamento.xlsx"      # nome e percorso nel repo
SHEET_NAME = ""                      # "" => primo foglio
HOME_URL   = "https://homeeuroirte.streamlit.app/"
LOGO_PATH  = "LogoEuroirte.png"      # opzionale, se presente accanto allo script

# Token opzionale (consigliato per rate limit API)
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", os.getenv("GITHUB_TOKEN", None))

PAGE_TITLE = "Avanzamento mensile €/h per Tecnico - Euroirte s.r.l."

# Endpoint API GitHub
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

# Imposta lo sfondo (metti il nome del file che usi nel progetto)
set_page_background("sfondo.png")  # immagine soft con glow

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
    Legge Avanzamento.xlsx direttamente dall'API GitHub (no CDN).
    Ritorna (sha, bytes, last_modified_human).
    """
    r = requests.get(API_URL, headers=_headers(), params={"ref": BRANCH}, timeout=30)
    r.raise_for_status()
    j = r.json()

    sha = j.get("sha") or str(int(time.time()))
    content = j.get("content")
    encoding = j.get("encoding")
    download_url = j.get("download_url")

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

    if content and encoding == "base64":
        data = base64.b64decode(content)
        return sha, data, last_human

    if download_url:
        r3 = requests.get(download_url, timeout=30, headers={"Cache-Control": "no-cache", "Pragma": "no-cache"})
        r3.raise_for_status()
        return sha, r3.content, last_human

    raise RuntimeError("Impossibile ottenere il contenuto del file da GitHub.")

def load_avanzamento_df_from_bytes(xls_bytes: bytes) -> pd.DataFrame:
    """
    Parsa l'Excel (valori calcolati) e **mantiene** 'Mail' e 'Data aggiornamento'
    oltre a 'Tecnico', 'Ore lavorate', 'Avanzamento €/h'.
    """
    bio = io.BytesIO(xls_bytes)
    wb = load_workbook(bio, data_only=True, read_only=True)
    ws = wb[SHEET_NAME] if SHEET_NAME and SHEET_NAME in wb.sheetnames else wb[wb.sheetnames[0]]

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return pd.DataFrame(columns=["Tecnico","Data aggiornamento","Ore lavorate","Avanzamento €/h","Mail"])

    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    df = pd.DataFrame(rows[1:], columns=header)

    # Normalizzazione nomi colonne
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

    wanted = ["Tecnico", "Data aggiornamento", "Ore lavorate", "Avanzamento €/h", "Mail"]
    for w in wanted:
        if w not in df.columns:
            df[w] = None

    df = df[wanted].copy()
    df["Ore lavorate"] = pd.to_numeric(df["Ore lavorate"], errors="coerce")
    df["Avanzamento €/h"] = pd.to_numeric(df["Avanzamento €/h"], errors="coerce")
    df["Data aggiornamento"] = pd.to_datetime(df["Data aggiornamento"], errors="coerce").dt.strftime("%d/%m/%Y")

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

# Data ultimo aggiornamento (da commit) + messaggio spostato qui
if last_update_date:
    st.caption(f"📅 Dati aggiornati al {last_update_date}")
st.info("Seleziona un tecnico per visualizzare il dettaglio.")

# =========================
# SELECTBOX (nessun default)
# =========================
PLACEHOLDER = "— Seleziona un tecnico —"
tecnici = sorted(df["Tecnico"].astype(str).dropna().unique().tolist())
options = [PLACEHOLDER] + tecnici
selezionato = st.selectbox("👷‍♂️ Seleziona un tecnico", options, index=0)

# =========================
# INVIO EMAIL MANUALE (robusto + diagnostica)
# =========================
SMTP_HOST = "mail.euroirte.it"
MAIL_SUBJECT = "EUROIRTE - Avanzamento Economico"   # ← oggetto aggiornato

SMTP_USER = str(st.secrets["SMTP_USER"]).strip()
SMTP_PASS = str(st.secrets["SMTP_PASS"]).strip()
SMTP_FROM = str(st.secrets.get("SMTP_FROM", SMTP_USER)).strip()

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

st.subheader("📧 Invio email personalizzate")

# Avviso se mittente diverso dall'utente
if SMTP_FROM.lower() != SMTP_USER.lower():
    st.warning(f"Attenzione: mittente ({SMTP_FROM}) diverso dall'utente SMTP ({SMTP_USER}). Alcuni server lo vietano.")

# Tester rapido SMTP
def _smtp_login_test():
    try:
        with smtplib.SMTP_SSL(SMTP_HOST, 465, timeout=20) as s:
            s.noop()
            s.login(SMTP_USER, SMTP_PASS)
        return True, "Login OK via SSL (465)"
    except Exception as e_ssl:
        try:
            with smtplib.SMTP(SMTP_HOST, 587, timeout=20) as s:
                s.ehlo(); s.starttls(); s.ehlo()
                s.login(SMTP_USER, SMTP_PASS)
            return True, "Login OK via STARTTLS (587)"
        except Exception as e_tls:
            return False, f"Errore login.\nSSL465: {e_ssl}\nSTARTTLS587: {e_tls}"

if st.button("🔍 Test connessione SMTP"):
    ok, msg = _smtp_login_test()
    (st.success if ok else st.error)(msg)

# Trova colonne (nomi tolleranti)
def _norm(s): 
    return str(s).replace("\u00a0"," ").strip().lower()

email_col   = next((c for c in df.columns if _norm(c) in {"mail","email","e-mail"}), None)
tecnico_col = next((c for c in df.columns if _norm(c) == "tecnico"), None)
ore_col     = next((c for c in df.columns if _norm(c) in {"ore lavorate","orelavorate","ore"}), None)
av_col      = next((c for c in df.columns if _norm(c).startswith("avanzamento")), None)
data_col    = next((c for c in df.columns if _norm(c).startswith("data")), None)
DATA_RIF = last_update_date or ""

if email_col is None or tecnico_col is None or ore_col is None or av_col is None:
    st.error("Servono almeno: 'Tecnico', 'Ore lavorate', 'Avanzamento €/h' e 'Mail'.")
else:
    # Anteprima (prime 5)
    anteprima = []
    for _, r in df.head(5).iterrows():
        nome = str(r.get(tecnico_col, "")).strip()
        if len(nome) > 14:  # rimuovi prefisso tipo "IRTE0000001 - "
            nome = nome[14:].strip()
        data  = str(r.get(data_col, DATA_RIF)) if data_col else DATA_RIF
        avanz = float(pd.to_numeric(r.get(av_col, 0), errors="coerce") or 0)
        ore   = float(pd.to_numeric(r.get(ore_col, 0), errors="coerce") or 0)
        to    = ("" if pd.isna(r.get(email_col, "")) else str(r.get(email_col, ""))).strip()
        corpo = (
            f"Ciao {nome},\n\n"
            f"il tuo avanzamento economico aggiornato al {data} è di {avanz:.2f} €/h "
            f"e il totale delle ore lavorate è {ore:.0f}.\n"
        )
        anteprima.append({"Destinatario": to, "Messaggio": corpo})
    st.caption("Anteprima (prime 5 righe)")
    st.dataframe(pd.DataFrame(anteprima), use_container_width=True)

    # ---------- INVIO ----------
if st.button("✉️ Invia email a tutti"):
    risultati = []          # -> conterrà tuple: ("✅"/"❌", nome, email)
    invalidi  = []          # -> conterrà tuple: (nome, email_originale)
    try:
        # tenta SSL465, poi fallback 587
        try:
            server = smtplib.SMTP_SSL(SMTP_HOST, 465, timeout=30)
            mode = "SSL465"
        except Exception:
            server = smtplib.SMTP(SMTP_HOST, 587, timeout=30)
            server.ehlo(); server.starttls(); server.ehlo()
            mode = "STARTTLS587"

        with server:
            server.login(SMTP_USER, SMTP_PASS)
            for _, r in df.iterrows():
                # nome tecnico (senza prefisso)
                nome = str(r.get(tecnico_col, "")).strip()
                if len(nome) > 14:
                    nome = nome[14:].strip()

                # email raw + normalizzazione
                raw_to = r.get(email_col, "")
                to = ("" if pd.isna(raw_to) else str(raw_to)).strip()
                to_l = to.lower()

                # valida email: se non valida, accoda con NOME e continua
                if (not to) or (to_l in {"nan", "none", "<na>", "na"}) or (not EMAIL_RE.match(to)):
                    invalidi.append((nome, str(raw_to)))
                    continue

                # dati messaggio
                data  = str(r.get(data_col, DATA_RIF)) if data_col else DATA_RIF
                avanz = float(pd.to_numeric(r.get(av_col, 0), errors="coerce") or 0)
                ore   = float(pd.to_numeric(r.get(ore_col, 0), errors="coerce") or 0)

                corpo = (
                    f"Ciao {nome},\n\n"
                    f"il tuo avanzamento economico aggiornato al {data} è di {avanz:.2f} €/h "
                    f"e il totale delle ore lavorate è {ore:.0f}.\n"
                )
                msg = MIMEText(corpo, "plain", "utf-8")
                msg["Subject"] = MAIL_SUBJECT
                msg["From"] = SMTP_FROM
                msg["To"] = to

                refused = server.send_message(msg)
                if refused:
                    # rifiuti dal server: mostriamo nome + motivo
                    for dest, info in refused.items():
                        risultati.append(("❌", nome, f"{dest} ({info[0]} {str(info[1])})"))
                else:
                    risultati.append(("✅", nome, to))

        # riepilogo
        inviate = sum(1 for s,_,_ in risultati if s == "✅")
        st.success(f"Inviate {inviate} email (modalità {mode}).")

        if invalidi:
            st.warning(f"Ignorati {len(invalidi)} destinatari con email non valida/assente.")
            df_invalidi = pd.DataFrame(invalidi, columns=["Tecnico", "Email non valida"])
            st.dataframe(df_invalidi, use_container_width=True)

        df_ris = pd.DataFrame(risultati, columns=["Stato", "Tecnico", "Destinatario"])
        st.dataframe(df_ris, use_container_width=True)

    except smtplib.SMTPAuthenticationError as e:
        st.error(f"Autenticazione fallita: {e}")
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
            "Ore lavorate": "{:.0f}",
            "Avanzamento €/h": "€{:.2f}/h",
        })
        .applymap(color_semaforo, subset=["Avanzamento €/h"])
    )
    st.subheader("Dettaglio")
    st.table(styler)
