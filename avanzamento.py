# app.py
import io
import pandas as pd
import streamlit as st
import smtplib, re
from email.mime.text import MIMEText
from pandas import ExcelWriter

# =========================
# CONFIG SMTP
# =========================
SMTP_HOST = "mail.euroirte.it"
MAIL_SUBJECT = "EUROIRTE - Avanzamento Economico"

SMTP_USER = str(st.secrets["SMTP_USER"]).strip()
SMTP_PASS = str(st.secrets["SMTP_PASS"]).strip()
SMTP_FROM = str(st.secrets.get("SMTP_FROM", SMTP_USER)).strip()
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# =========================
# Caricamento file
# =========================
st.set_page_config(page_title="Avanzamento Mensile", layout="wide")

uploaded = st.sidebar.file_uploader("Carica Avanzamento.xlsx", type=["xlsx"])
if not uploaded:
    st.stop()

df = pd.read_excel(uploaded, engine="openpyxl")
df = df.rename(columns={c: c.strip() for c in df.columns})

req = ["Data aggiornamento", "Tecnico", "Ore lavorate", "Avanzamento €/h", "Mail"]
for r in req:
    if r not in df.columns:
        df[r] = None

df["Data aggiornamento"] = pd.to_datetime(df["Data aggiornamento"], errors="coerce", dayfirst=True)
df["Mese"] = df["Data aggiornamento"].dt.to_period("M").dt.to_timestamp()
df_all = df.copy()

# =========================
# Filtro globale mese
# =========================
mesi = sorted(df["Mese"].dropna().unique())
mese_scelto = st.selectbox("📅 Mese globale", mesi, index=len(mesi)-1,
                           format_func=lambda x: x.strftime("%m/%Y"))
mostra_tutti = st.toggle("Mostra tutti i mesi", False)

if not mostra_tutti:
    df = df[df["Mese"] == mese_scelto].copy()

df["Data aggiornamento_str"] = df["Data aggiornamento"].dt.strftime("%d/%m/%Y")

# =========================
# INVIO EMAIL PER MESE SCELTO
# =========================
st.divider()
st.subheader("📧 Invia email ai tecnici")

mesi_email = sorted(df_all["Mese"].dropna().unique(), reverse=True)
mese_email = st.selectbox("Mese da inviare", mesi_email,
                          index=mesi_email.index(mese_scelto))

if st.button("✉️ Invia email per il mese scelto"):
    df_email = df_all[df_all["Mese"] == mese_email].copy()
    if df_email.empty:
        st.warning("Nessun dato per questo mese")
        st.stop()

    try:
        # tenta SSL465, poi fallback 587
        try:
            server = smtplib.SMTP_SSL(SMTP_HOST, 465, timeout=30)
            mode = "SSL465"
        except Exception:
            server = smtplib.SMTP(SMTP_HOST, 587, timeout=30)
            server.ehlo(); server.starttls(); server.ehlo()
            mode = "STARTTLS587"

        risultati, invalidi = [], []
        with server:
            server.login(SMTP_USER, SMTP_PASS)
            for _, r in df_email.iterrows():
                nome = str(r["Tecnico"]).strip()
                to = str(r["Mail"]).strip() if pd.notna(r["Mail"]) else ""
                if not to or not EMAIL_RE.match(to):
                    invalidi.append((nome, to))
                    continue
                data = r["Data aggiornamento_str"]
                avanz = float(pd.to_numeric(r["Avanzamento €/h"], errors="coerce") or 0)
                ore = float(pd.to_numeric(r["Ore lavorate"], errors="coerce") or 0)

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
                    risultati.append(("❌", nome, to))
                else:
                    risultati.append(("✅", nome, to))

        st.success(f"Invio completato ({mode})")
        st.dataframe(pd.DataFrame(risultati, columns=["Stato", "Tecnico", "Email"]))
        if invalidi:
            st.warning("Email non valide/assenti")
            st.dataframe(pd.DataFrame(invalidi, columns=["Tecnico", "Email"]))
    except Exception as e:
        st.error(f"Errore invio: {e}")
