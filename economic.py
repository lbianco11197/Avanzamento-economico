import io
import os
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
    # usa raw.githubusercontent.com
    return f"https://raw.githubusercontent.com/{REPO_OWNER}/{REPO_NAME}/{BRANCH}/{path}"

@st.cache_data(show_spinner=False, ttl=600)
def fetch_excel_from_github(path: str) -> pd.DataFrame:
    url = raw_url(path)
    headers = {}
    if GITHUB_TOKEN:
        # Abilita accesso a repo privati con GitHub token
        headers["Authorization"] = f"token {GITHUB_TOKEN}"
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    # pandas legge da BytesIO
    return pd.read_excel(io.BytesIO(r.content))

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

# Carica i 4 file direttamente dal repo
try:
    df_ore     = fetch_excel_from_github(PATH_PRES)
    df_del_tim = fetch_excel_from_github(PATH_DEL_TIM)
    df_ass_tim = fetch_excel_from_github(PATH_ASS_TIM)
    df_del_of  = fetch_excel_from_github(PATH_DEL_OF)
except Exception as e:
    st.error(f"Errore nel caricamento dai raw GitHub: {e}")
    st.stop()

# Filtro mese opzionale (se esistono colonne Data)
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

# --- Presenze: Tecnico, Totale (ore)
df_ore = df_ore.rename(columns=lambda c: str(c).strip())
if not {"Tecnico","Totale"}.issubset(df_ore.columns):
    st.error("Nel file Presenze servono le colonne: 'Tecnico' e 'Totale'.")
    st.stop()
df_ore = df_ore[["Tecnico","Totale"]].rename(columns={"Tecnico":"tecnico","Totale":"ore_totali"})
df_ore["tecnico"] = df_ore["tecnico"].astype(str).str.strip().str.lower()
df_ore["ore_totali"] = pd.to_numeric(df_ore["ore_totali"], errors="coerce").fillna(0)
df_ore = df_ore.groupby("tecnico", as_index=False)["ore_totali"].sum()

# --- Delivery TIM: colonne con varianti di "≠ FTTH"
df_del_tim = df_del_tim.rename(columns=lambda c: str(c).strip())
cols_map = {
    "Tecnico":"tecnico",
    "Impianti espletati FTTH":"del_tim_ftth",
    "Impianti espletati ≠ FTTH":"del_tim_non_ftth",
    "Impianti espletati != FTTH":"del_tim_non_ftth",
    "Impianti espletati â‰  FTTH":"del_tim_non_ftth",   # mis-encoding frequente
    "Impianti espletati â‰ FTTH":"del_tim_non_ftth",
    "Impianti espletati â‰  FTTH":"del_tim_non_ftth",
}
df_del_tim = df_del_tim.rename(columns=cols_map)
need = {"tecnico","del_tim_ftth","del_tim_non_ftth"}
miss = need - set(df_del_tim.columns)
if miss:
    st.error(f"Nel file Delivery TIM mancano colonne: {', '.join(miss)}")
    st.stop()
df_del_tim["tecnico"] = df_del_tim["tecnico"].astype(str).str.strip().str.lower()
for c in ["del_tim_ftth","del_tim_non_ftth"]:
    df_del_tim[c] = pd.to_numeric(df_del_tim[c], errors="coerce").fillna(0)
df_del_tim = df_del_tim.groupby("tecnico", as_index=False)[["del_tim_ftth","del_tim_non_ftth"]].sum()

# --- Assurance TIM: Referente oppure Tecnico + ProduttiviCount
df_ass_tim = df_ass_tim.rename(columns=lambda c: str(c).strip())
tec_col = "Referente" if "Referente" in df_ass_tim.columns else ("Tecnico" if "Tecnico" in df_ass_tim.columns else None)
if tec_col is None or "ProduttiviCount" not in df_ass_tim.columns:
    st.error("Nel file Assurance TIM servono: 'Referente' (o 'Tecnico') e 'ProduttiviCount'.")
    st.stop()
df_ass_tim = df_ass_tim.rename(columns={tec_col:"tecnico","ProduttiviCount":"ass_tim"})
df_ass_tim["tecnico"] = df_ass_tim["tecnico"].astype(str).str.strip().str.lower()
df_ass_tim["ass_tim"] = pd.to_numeric(df_ass_tim["ass_tim"], errors="coerce").fillna(0)
df_ass_tim = df_ass_tim.groupby("tecnico", as_index=False)["ass_tim"].sum()

# --- Delivery OF: Tecnico + Impianti espletati
df_del_of = df_del_of.rename(columns=lambda c: str(c).strip())
if not {"Tecnico","Impianti espletati"}.issubset(df_del_of.columns):
    st.error("Nel file Delivery OF servono: 'Tecnico' e 'Impianti espletati'.")
    st.stop()
df_del_of = df_del_of.rename(columns={"Tecnico":"tecnico","Impianti espletati":"del_of"})
df_del_of["tecnico"] = df_del_of["tecnico"].astype(str).str.strip().str.lower()
df_del_of["del_of"] = pd.to_numeric(df_del_of["del_of"], errors="coerce").fillna(0)
df_del_of = df_del_of.groupby("tecnico", as_index=False)["del_of"].sum()

# --- Merge
df = df_ore.merge(df_del_tim, on="tecnico", how="left") \
           .merge(df_ass_tim, on="tecnico", how="left") \
           .merge(df_del_of, on="tecnico", how="left") \
           .fillna(0)

# --- Calcolo €/h
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

# Esporta CSV
csv = df_out.to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Scarica CSV", data=csv, file_name="avanzamento_euro_ora.csv", mime="text/csv")

# Info sorgenti
with st.expander("Sorgenti GitHub"):
    st.code(f"Presenze:      {raw_url(PATH_PRES)}", language="text")
    st.code(f"Delivery TIM:  {raw_url(PATH_DEL_TIM)}", language="text")
    st.code(f"Assurance TIM: {raw_url(PATH_ASS_TIM)}", language="text")
    st.code(f"Delivery OF:   {raw_url(PATH_DEL_OF)}", language="text")
