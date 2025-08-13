import io
import os
import re
import requests
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(layout="wide", page_title="Avanzamento economico €/h")

# Forza tema chiaro anche su mobile
st.markdown("""
<style>
:root { color-scheme: light !important; }
@media (prefers-color-scheme: dark){ :root { color-scheme: light !important; } }
html, body, [data-testid="stApp"], [data-testid="stAppViewContainer"],
[data-testid="stHeader"], [data-testid="stSidebar"] { background:#fff !important; color:#000 !important; }
div[data-baseweb="select"], input, textarea, select { background:#fff !important; color:#000 !important; }
.stButton > button{ background:#fff !important; color:#000 !important; border:1px solid #999 !important; border-radius:6px; }
.stDataFrame [role="grid"], .stTable, .stDataFrame table, .stDataFrame th, .stDataFrame td { background:#fff !important; color:#000 !important; }
header [data-testid="theme-toggle"]{ display:none; }
</style>
""", unsafe_allow_html=True)

st.title("📊 Avanzamento mensile €/h per Tecnico (da GitHub)")

# =========================
# CONFIG REPO (ADATTATA)
# =========================
REPO_OWNER   = "lbianco11197"
REPO_NAME    = "Avanzamento-economico"
BRANCH       = "main"

# I file Excel sono nella root del repo con questi nomi:
PATH_PRES    = "Presenze.xlsx"
PATH_DEL_TIM = "Delivery TIM.xlsx"
PATH_ASS_TIM = "Assurance TIM.xlsx"
PATH_DEL_OF  = "Delivery OF.xlsx"

# Se il repo è PRIVATO, inserisci un token in st.secrets["GITHUB_TOKEN"]
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", os.getenv("GITHUB_TOKEN", None))

# Fattori economici (modificabili)
F_DEL_TIM_FTTH = 100
F_DEL_TIM_NON  = 40
F_ASS_TIM      = 20
F_DEL_OF       = 100

def raw_url(path: str) -> str:
    return f"https://raw.githubusercontent.com/{REPO_OWNER}/{REPO_NAME}/{BRANCH}/{path}"

@st.cache_data(show_spinner=False, ttl=600)
def fetch_excel_from_github(path: str) -> pd.DataFrame:
    url = raw_url(path)
    headers = {}
    if GITHUB_TOKEN:
        headers["Authorization"] = f"token {GITHUB_TOKEN}"
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    return pd.read_excel(io.BytesIO(r.content))

# ---------- Utilità per date ----------
def find_date_col(df: pd.DataFrame):
    if df is None: return None
    for c in df.columns:
        if str(c).strip().lower() in ("data", "date"):
            return c
    for c in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            return c
    return None

def ensure_datetime(df: pd.DataFrame, col: str) -> pd.DataFrame:
    if df is None or col is None: return df
    dfx = df.copy()
    if not pd.api.types.is_datetime64_any_dtype(dfx[col]):
        dfx[col] = pd.to_datetime(dfx[col], dayfirst=True, errors="coerce")
    return dfx

def month_options(*dfs):
    opts = []
    for df in dfs:
        if df is None: continue
        c = find_date_col(df)
        if c is None: continue
        dfx = ensure_datetime(df, c)
        if dfx[c].notna().any():
            ms = dfx[c].dt.to_period("M").dropna().unique().astype(str).tolist()
            opts.extend(ms)
    return sorted(set(opts))

# ---------- Caricamento file ----------
try:
    df_ore     = fetch_excel_from_github(PATH_PRES)
    df_del_tim = fetch_excel_from_github(PATH_DEL_TIM)
    df_ass_tim = fetch_excel_from_github(PATH_ASS_TIM)
    df_del_of  = fetch_excel_from_github(PATH_DEL_OF)
except Exception as e:
    st.error(f"Errore nel caricamento dai raw GitHub: {e}")
    st.stop()

# ---------- Filtro mese opzionale ----------
options = month_options(df_ore, df_del_tim, df_ass_tim, df_del_of)
selected_period = None
if options:
    selected_period = st.selectbox("📅 Mese (se disponibile nei file):", options, index=len(options)-1)

def filter_period(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or not selected_period: return df
    c = find_date_col(df)
    if c is None: return df
    dfx = ensure_datetime(df, c)
    per = pd.Period(selected_period, freq="M")
    return dfx[dfx[c].dt.to_period("M") == per]

df_ore     = filter_period(df_ore)
df_del_tim = filter_period(df_del_tim)
df_ass_tim = filter_period(df_ass_tim)
df_del_of  = filter_period(df_del_of)

# ---------- Presenze ----------
df_ore = df_ore.rename(columns=lambda c: str(c).strip())
if not {"Tecnico","Totale"}.issubset(df_ore.columns):
    st.error("Nel file Presenze servono le colonne: 'Tecnico' e 'Totale'.")
    st.stop()
df_ore = df_ore[["Tecnico","Totale"]].rename(columns={"Tecnico":"tecnico","Totale":"ore_totali"})
df_ore["tecnico"] = df_ore["tecnico"].astype(str).str.strip().str.lower()
df_ore["ore_totali"] = pd.to_numeric(df_ore["ore_totali"], errors="coerce").fillna(0)
df_ore = df_ore.groupby("tecnico", as_index=False)["ore_totali"].sum()

# ---------- Delivery TIM: rilevamento robusto colonne FTTH / NON FTTH ----------
def normalize(s: str) -> str:
    s = str(s).lower()
    s = s.replace("\u2260", "!=")   # ≠
    s = s.replace("â‰", "!=")       # mis-encoding frequente
    s = s.replace(" ftth", "ftth")
    s = s.replace("≠", "!=")
    s = s.replace("  ", " ")
    s = re.sub(r"\s+", "", s)       # rimuovi spazi
    s = re.sub(r"[^a-z0-9!=]", "", s)  # solo alfanum e ! =
    return s

def find_ftth_col(columns):
    candidates = []
    for col in columns:
        n = normalize(col)
        if "impiantiespletati" in n and "ftth" in n and "non" not in n and "!=" not in n:
            candidates.append(col)
    # fallback: qualunque col con ftth
    if not candidates:
        for col in columns:
            n = normalize(col)
            if "ftth" in n and "non" not in n and "!=" not in n:
                candidates.append(col)
    return candidates[0] if candidates else None

def find_non_ftth_col(columns):
    for col in columns:
        n = normalize(col)
        if ("impiantiespletati" in n and ("nonftth" in n or "!=ftth" in n)) or \
           ("ftth" in n and ("nonftth" in n or "!=ftth" in n)):
            return col
    # Alcuni file usano "FTTC" per indicare non FTTH: prova last-resort
    for col in columns:
        n = normalize(col)
        if "impiantiespletati" in n and "fttc" in n:
            return col
    return None

df_del_tim = df_del_tim.rename(columns=lambda c: str(c).strip())
# Individua colonne
tec_col = None
for c in df_del_tim.columns:
    if normalize(c) in ("tecnico",):
        tec_col = c
        break
if tec_col is None and "Tecnico" in df_del_tim.columns:
    tec_col = "Tecnico"
if tec_col is None:
    st.error("Nel file Delivery TIM manca la colonna 'Tecnico'.")
    st.stop()

ftth_col = find_ftth_col(df_del_tim.columns)
non_ftth_col = find_non_ftth_col(df_del_tim.columns)

# Rinomina quanto trovato
rename_map = {tec_col: "tecnico"}
if ftth_col:     rename_map[ftth_col] = "del_tim_ftth"
if non_ftth_col: rename_map[non_ftth_col] = "del_tim_non_ftth"
df_del_tim = df_del_tim.rename(columns=rename_map)

# Se mancano, crea colonne a 0 per proseguire senza errore
if "del_tim_ftth" not in df_del_tim.columns:
    st.warning("⚠️ Colonna FTTH non trovata in Delivery TIM: impostata a 0.")
    df_del_tim["del_tim_ftth"] = 0
if "del_tim_non_ftth" not in df_del_tim.columns:
    st.warning("⚠️ Colonna NON FTTH non trovata in Delivery TIM: impostata a 0.")
    df_del_tim["del_tim_non_ftth"] = 0

df_del_tim["tecnico"] = df_del_tim["tecnico"].astype(str).str.strip().str.lower()
for c in ["del_tim_ftth","del_tim_non_ftth"]:
    df_del_tim[c] = pd.to_numeric(df_del_tim[c], errors="coerce").fillna(0)
df_del_tim = df_del_tim.groupby("tecnico", as_index=False)[["del_tim_ftth","del_tim_non_ftth"]].sum()

# ---------- Assurance TIM ----------
df_ass_tim = df_ass_tim.rename(columns=lambda c: str(c).strip())
tec_col = "Referente" if "Referente" in df_ass_tim.columns else ("Tecnico" if "Tecnico" in df_ass_tim.columns else None)
if tec_col is None or "ProduttiviCount" not in df_ass_tim.columns:
    st.error("Nel file Assurance TIM servono: 'Referente' (o 'Tecnico') e 'ProduttiviCount'.")
    st.stop()
df_ass_tim = df_ass_tim.rename(columns={tec_col:"tecnico","ProduttiviCount":"ass_tim"})
df_ass_tim["tecnico"] = df_ass_tim["tecnico"].astype(str).str.strip().str.lower()
df_ass_tim["ass_tim"] = pd.to_numeric(df_ass_tim["ass_tim"], errors="coerce").fillna(0)
df_ass_tim = df_ass_tim.groupby("tecnico", as_index=False)["ass_tim"].sum()

# ---------- Delivery OF ----------
df_del_of = df_del_of.rename(columns=lambda c: str(c).strip())
if not {"Tecnico","Impianti espletati"}.issubset(df_del_of.columns):
    st.error("Nel file Delivery OF servono: 'Tecnico' e 'Impianti espletati'.")
    st.stop()
df_del_of = df_del_of.rename(columns={"Tecnico":"tecnico","Impianti espletati":"del_of"})
df_del_of["tecnico"] = df_del_of["tecnico"].astype(str).str.strip().str.lower()
df_del_of["del_of"] = pd.to_numeric(df_del_of["del_of"], errors="coerce").fillna(0)
df_del_of = df_del_of.groupby("tecnico", as_index=False)["del_of"].sum()

# ---------- Merge + calcoli ----------
df = df_ore.merge(df_del_tim, on="tecnico", how="left") \
           .merge(df_ass_tim, on="tecnico", how="left") \
           .merge(df_del_of, on="tecnico", how="left") \
           .fillna(0)

ore = df["ore_totali"].replace(0, np.nan)
df["Resa Delivery TIM FTTH"]     = (df["del_tim_ftth"]     * F_DEL_TIM_FTTH) / ore
df["Resa Delivery TIM non FTTH"] = (df["del_tim_non_ftth"] * F_DEL_TIM_NON)  / ore
df["Resa Assurance TIM"]         = (df["ass_tim"]          * F_ASS_TIM)      / ore
df["Resa Delivery OF"]           = (df["del_of"]           * F_DEL_OF)       / ore
df = df.replace({np.nan: 0})

df_out = df.rename(columns={"tecnico":"Nome Tecnico"})[
    ["Nome Tecnico","Resa Delivery TIM FTTH","Resa Delivery TIM non FTTH","Resa Assurance TIM","Resa Delivery OF"]
].sort_values("Nome Tecnico")

st.subheader("Riepilogo €/h per Tecnico")
st.dataframe(
    df_out.style.format({
        "Resa Delivery TIM FTTH":"€{:.2f}/h",
        "Resa Delivery TIM non FTTH":"€{:.2f}/h",
        "Resa Assurance TIM":"€{:.2f}/h",
        "Resa Delivery OF":"€{:.2f}/h",
    }),
    use_container_width=True,
    hide_index=True
)

csv = df_out.to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Scarica CSV", data=csv, file_name="avanzamento_euro_ora.csv", mime="text/csv")

with st.expander("Sorgenti GitHub"):
    st.code(f"Presenze:      {raw_url(PATH_PRES)}", language="text")
    st.code(f"Delivery TIM:  {raw_url(PATH_DEL_TIM)}", language="text")
    st.code(f"Assurance TIM: {raw_url(PATH_ASS_TIM)}", language="text")
    st.code(f"Delivery OF:   {raw_url(PATH_DEL_OF)}", language="text")
