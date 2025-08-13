import streamlit as st
import pandas as pd
from datetime import datetime
from streamlit.components.v1 import html

st.set_page_config(layout="wide")

# Imposta sfondo bianco e testo nero
st.markdown("""
    <style>
    html, body, [data-testid="stApp"] {
        background-color: white !important;
        color: black !important;
    }

    /* Forza colore dei testi nei menu a discesa */
    .stSelectbox div[data-baseweb="select"] {
        background-color: white !important;
        color: black !important;
    }

    .stSelectbox span, .stSelectbox label {
        color: black !important;
        font-weight: 500;
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
st.title("📊 Dashboard di Produttività €/h - Euroirte s.r.l.")

# Intestazione con logo e bottone
# Logo in alto
st.image("LogoEuroirte.jpg", width=180)

# Bottone sotto il logo
st.link_button("🏠 Torna alla Home", url="https://homeeuroirte.streamlit.app/")

# --- INSERISCI QUI I TUOI URL "RAW" DI GITHUB ---
# Sostituisci queste stringhe di esempio con i tuoi link reali
URL_PRESENZE = "https://raw.githubusercontent.com/lbianco11197/Avanzamento-economico/raw/refs/heads/main/Presenze.xlsx"
URL_DEL_TIM = "https://raw.githubusercontent.com/lbianco11197/Avanzamento-economico/raw/refs/heads/main/Delivery%20TIM.xlsx"
URL_ASS_TIM = "https://raw.githubusercontent.com/lbianco11197/Avanzamento-economico/raw/refs/heads/main/Assurance%20TIM.xlsx"
URL_DEL_OF = "https://raw.githubusercontent.com/lbianco11197/Avanzamento-economico/raw/refs/heads/main/Delivery%20OF.xlsx"
# ----------------------------------------------------

# Funzione per caricare un file da un URL, gestendo errori
@st.cache_data(ttl=600) # Mette in cache i dati per 10 minuti
def carica_dati_da_url(url):
    try:
        # Determina se il file è Excel o CSV basandosi sull'URL
        if '.xlsx' in url:
            return pd.read_excel(url)
        else:
            return pd.read_csv(url)
    except Exception as e:
        st.error(f"Impossibile caricare il file da: {url}\nErrore: {e}")
        return None

# --- Logica Principale ---
try:
    # 1. Caricamento e Preparazione Dati
    df_ore = carica_dati_da_url(URL_PRESENZE)
    df_del_tim = carica_dati_da_url(URL_DEL_TIM)
    df_ass_tim = carica_dati_da_url(URL_ASS_TIM)
    df_del_of = carica_dati_da_url(URL_DEL_OF)

    # Controlla se tutti i file sono stati caricati correttamente
    if all(df is not None for df in [df_ore, df_del_tim, df_ass_tim, df_del_of]):
        
        # Preparazione file Ore
        df_ore = df_ore[['Tecnico', 'Totale']].rename(columns={'Tecnico': 'tecnico', 'Totale': 'ore_totali'})
        df_ore['tecnico'] = df_ore['tecnico'].str.strip().str.lower()

        # Preparazione Delivery TIM
        df_del_tim = df_del_tim.rename(columns={
            'Tecnico': 'tecnico',
            'Impianti espletati FTTH': 'del_tim_ftth',
            'Impianti espletati != FTTH': 'del_tim_non_ftth',
            'Impianti espletati â‰  FTTH': 'del_tim_non_ftth'
        })
        df_del_tim = df_del_tim[['tecnico', 'del_tim_ftth', 'del_tim_non_ftth']]
        df_del_tim['tecnico'] = df_del_tim['tecnico'].str.strip().str.lower()
        df_del_tim = df_del_tim.groupby('tecnico').sum().reset_index()
        
        # Preparazione Assurance TIM
        df_ass_tim = df_ass_tim.rename(columns={'Referente': 'tecnico', 'ProduttiviCount': 'ass_tim'})
        df_ass_tim = df_ass_tim[['tecnico', 'ass_tim']]
        df_ass_tim['tecnico'] = df_ass_tim['tecnico'].str.strip().str.lower()
        df_ass_tim = df_ass_tim.groupby('tecnico').sum().reset_index()
        
        # Preparazione Delivery OF
        df_del_of = df_del_of.rename(columns={'Tecnico': 'tecnico', 'Impianti espletati': 'del_of'})
        df_del_of = df_del_of[['tecnico', 'del_of']]
        df_del_of['tecnico'] = df_del_of['tecnico'].str.strip().str.lower()
        df_del_of = df_del_of.groupby('tecnico').sum().reset_index()

        # 2. Unione dei Dati (Merge)
        df_finale = pd.merge(df_ore, df_del_tim, on='tecnico', how='left')
        df_finale = pd.merge(df_finale, df_ass_tim, on='tecnico', how='left')
        df_finale = pd.merge(df_finale, df_del_of, on='tecnico', how='left')
        df_finale = df_finale.fillna(0)

        # 3. Calcolo delle Performance €/h
        F_DEL_TIM_FTTH, F_DEL_TIM_NON_FTTH, F_ASS_TIM, F_DEL_OF = 100, 40, 20, 100
        
        df_finale['Resa Delivery TIM FTTH'] = np.divide((df_finale['del_tim_ftth'] * F_DEL_TIM_FTTH), df_finale['ore_totali'], where=df_finale['ore_totali']!=0)
        df_finale['Resa Delivery TIM non FTTH'] = np.divide((df_finale['del_tim_non_ftth'] * F_DEL_TIM_NON_FTTH), df_finale['ore_totali'], where=df_finale['ore_totali']!=0)
        df_finale['Resa Assurance TIM'] = np.divide((df_finale['ass_tim'] * F_ASS_TIM), df_finale['ore_totali'], where=df_finale['ore_totali']!=0)
        df_finale['Resa Delivery OF'] = np.divide((df_finale['del_of'] * F_DEL_OF), df_finale['ore_totali'], where=df_finale['ore_totali']!=0)

        # 4. Visualizzazione della Tabella Finale
        st.header("Riepilogo Performance €/h per Tecnico")
        df_display = df_finale[['tecnico', 'Resa Delivery TIM FTTH', 'Resa Delivery TIM non FTTH', 'Resa Assurance TIM', 'Resa Delivery OF']]
        
        st.dataframe(
            df_display.style.format({
                'Resa Delivery TIM FTTH': '€{:.2f}/h',
                'Resa Delivery TIM non FTTH': '€{:.2f}/h',
                'Resa Assurance TIM': '€{:.2f}/h',
                'Resa Delivery OF': '€{:.2f}/h'
            }).background_gradient(
                cmap='viridis', 
                subset=['Resa Delivery TIM FTTH', 'Resa Delivery TIM non FTTH', 'Resa Assurance TIM', 'Resa Delivery OF']
            ),
            use_container_width=True,
            hide_index=True
        )

except Exception as e:
    st.error(f"Si è verificato un errore generale durante l'elaborazione dei dati. Controlla che gli URL siano corretti e che i file non siano corrotti. Errore: {e}")
