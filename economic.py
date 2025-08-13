import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(layout="wide", page_title="Avanzamento economico €/h")

# --- Tema chiaro forzato anche su mobile ---
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

st.title("📊 Avanzamento mensile €/h per Tecnico")

with st.expander("Carica i 4 file (Excel)"):
    f_presenze   = st.file_uploader("1) Presenze.xlsx (colonne: Tecnico, Totale, facoltativa Data)", type=["xlsx"], key="presenze")
    f_del_tim    = st.file_uploader("2) Delivery TIM.xlsx (colonne: Tecnico, Impianti espletati FTTH, Impianti espletati ≠ FTTH)", type=["xlsx"], key="del_tim")
    f_ass_tim    = st.file_uploader("3) Assurance TIM.xlsx (colonne: Referente o Tecnico, ProduttiviCount)", type=["xlsx"], key="ass_tim")
    f_del_of     = st.file_uploader("4) Delivery OF.xlsx (colonne: Tecnico, Impianti espletati)", type=["xlsx"], key="del_of")

# --- Helper: trova una colonna data (facoltativo) e filtra per mese/anno ---
def find_date_col(df):
    if df is None: return None
    # preferisci una colonna chiamata Data/Date
    for c in df.columns:
        if str(c).strip().lower() in ("data","date"):
            return c
    # altrimenti prima colonna datetime
    for c in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            return c
    return None

def ensure_datetime(df, col):
    if df is None or col is None: return df
    df = df.copy()
    if not pd.api.types.is_datetime64_any_dtype(df[col]):
        df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
    return df

def month_filter_ui(dfs):
    # raccogli mesi disponibili da qualunque df che abbia una data
    options = []
    for df in dfs:
        if df is None: continue
        col = find_date_col(df)
        if col is None: continue
        dfx = ensure_datetime(df, col)
        if dfx[col].notna().any():
            ms = dfx[col].dt.to_period("M").dropna().unique().astype(str).tolist()
            options.extend(ms)
    options = sorted(set(options))
    if options:
        sel = st.selectbox("📅 Filtra per mese (se i file hanno una colonna Data):", options, index=len(options)-1)
        return sel
    return None

# --- Loader robusto per i 4 file ---
def load_xlsx(uploaded):
    if uploaded is None: return None
    try:
        return pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Errore nel leggere {uploaded.name if hasattr(uploaded,'name') else 'file'}: {e}")
        return None

df_ore     = load_xlsx(f_presenze)
df_del_tim = load_xlsx(f_del_tim)
df_ass_tim = load_xlsx(f_ass_tim)
df_del_of  = load_xlsx(f_del_of)

# Filtro mese opzionale
selected_period = month_filter_ui([df_ore, df_del_tim, df_ass_tim, df_del_of])

def filter_by_period(df):
    if df is None or selected_period is None: return df
    col = find_date_col(df)
    if col is None: return df
    dfx = ensure_datetime(df, col)
    per = pd.Period(selected_period, freq="M")
    return dfx[dfx[col].dt.to_period("M") == per]

df_ore     = filter_by_period(df_ore)
df_del_tim = filter_by_period(df_del_tim)
df_ass_tim = filter_by_period(df_ass_tim)
df_del_of  = filter_by_period(df_del_of)

if None in (df_ore, df_del_tim, df_ass_tim, df_del_of):
    st.info("➡️ Carica tutti e quattro i file per vedere i risultati.")
    st.stop()

# --- Normalizzazioni colonne/naming ---
# Presenze: Tecnico, Totale (ore)
df_ore = df_ore.rename(columns=lambda c: str(c).strip())
if not {"Tecnico","Totale"}.issubset(df_ore.columns):
    st.error("Nel file Presenze servono le colonne: 'Tecnico' e 'Totale'.")
    st.stop()
df_ore = df_ore[["Tecnico","Totale"]].rename(columns={"Tecnico":"tecnico","Totale":"ore_totali"})
df_ore["tecnico"] = df_ore["tecnico"].astype(str).str.strip().str.lower()
df_ore["ore_totali"] = pd.to_numeric(df_ore["ore_totali"], errors="coerce").fillna(0)
df_ore = df_ore.groupby("tecnico", as_index=False)["ore_totali"].sum()

# Delivery TIM: gestisci ≠ FTTH in varie grafie/encoding
df_del_tim = df_del_tim.rename(columns=lambda c: str(c).strip())
cols_map = {
    "Tecnico":"tecnico",
    "Impianti espletati FTTH":"del_tim_ftth",
    "Impianti espletati ≠ FTTH":"del_tim_non_ftth",
    "Impianti espletati != FTTH":"del_tim_non_ftth",
    "Impianti espletati â‰  FTTH":"del_tim_non_ftth",  # mis-encoding frequente
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

# Assurance TIM: Referente oppure Tecnico, ProduttiviCount
df_ass_tim = df_ass_tim.rename(columns=lambda c: str(c).strip())
tec_col = "Referente" if "Referente" in df_ass_tim.columns else ("Tecnico" if "Tecnico" in df_ass_tim.columns else None)
if tec_col is None or "ProduttiviCount" not in df_ass_tim.columns:
    st.error("Nel file Assurance TIM servono: 'Referente' (o 'Tecnico') e 'ProduttiviCount'.")
    st.stop()
df_ass_tim = df_ass_tim.rename(columns={tec_col:"tecnico","ProduttiviCount":"ass_tim"})
df_ass_tim["tecnico"] = df_ass_tim["tecnico"].astype(str).str.strip().str.lower()
df_ass_tim["ass_tim"] = pd.to_numeric(df_ass_tim["ass_tim"], errors="coerce").fillna(0)
df_ass_tim = df_ass_tim.groupby("tecnico", as_index=False)["ass_tim"].sum()

# Delivery OF: Tecnico, Impianti espletati
df_del_of = df_del_of.rename(columns=lambda c: str(c).strip())
if not {"Tecnico","Impianti espletati"}.issubset(df_del_of.columns):
    st.error("Nel file Delivery OF servono: 'Tecnico' e 'Impianti espletati'.")
    st.stop()
df_del_of = df_del_of.rename(columns={"Tecnico":"tecnico","Impianti espletati":"del_of"})
df_del_of["tecnico"] = df_del_of["tecnico"].astype(str).str.strip().str.lower()
df_del_of["del_of"] = pd.to_numeric(df_del_of["del_of"], errors="coerce").fillna(0)
df_del_of = df_del_of.groupby("tecnico", as_index=False)["del_of"].sum()

# --- Merge finale ---
df = df_ore.merge(df_del_tim, on="tecnico", how="left") \
           .merge(df_ass_tim, on="tecnico", how="left") \
           .merge(df_del_of, on="tecnico", how="left") \
           .fillna(0)

# --- Fattori economici ---
F_DEL_TIM_FTTH   = 100
F_DEL_TIM_NON    = 40
F_ASS_TIM        = 20
F_DEL_OF         = 100

# --- Calcolo €/h (evita divisioni per zero) ---
ore = df["ore_totali"].replace(0, np.nan)
df["Resa Delivery TIM FTTH"]     = (df["del_tim_ftth"]   * F_DEL_TIM_FTTH) / ore
df["Resa Delivery TIM non FTTH"] = (df["del_tim_non_ftth"] * F_DEL_TIM_NON) / ore
df["Resa Assurance TIM"]         = (df["ass_tim"]        * F_ASS_TIM) / ore
df["Resa Delivery OF"]           = (df["del_of"]         * F_DEL_OF) / ore
df = df.replace({np.nan: 0})

# --- Output tabella richiesta ---
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

# Download CSV
csv = df_out.to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Scarica CSV", data=csv, file_name="avanzamento_euro_ora.csv", mime="text/csv")
