import pandas as pd
import streamlit as st

# Nastavení stránky
st.set_page_config(
    page_title="Evidence závad lokomotiv", layout="wide", page_icon="🚆"
)

st.title("🚆 Evidence a přehled závad lokomotiv")


# Načtení dat ze souboru
@st.cache_data
def load_data():
    # Načte přímo vyčištěný soubor připravený pro web
    df = pd.read_excel("PREDAVKA_ELEKTRONICI_PRO_APPSHEET.xlsx")
    return df


df = load_data()

# Postranní panel s filtry
st.sidebar.header("🔍 Filtry")

# Filtr podle lokomotivy
seznam_loko = sorted([str(x) for x in df["Lokomotiva"].dropna().unique()])
vybrane_loko = st.sidebar.multiselect(
    "Vyberte lokomotivu:", options=seznam_loko
)

# Textové vyhledávání
vyhledavani = st.sidebar.text_input("Hledat v popisu závady nebo poznámce:")

# Aplikace filtrů
filtr_df = df.copy()

if vybrane_loko:
    filtr_df = filtr_df[filtr_df["Lokomotiva"].astype(str).isin(vybrane_loko)]

if vyhledavani:
    maska = filtr_df["Popis závady"].astype(str).str.contains(
        vyhledavani, case=False, na=False
    ) | filtr_df["Poznámka"].astype(str).str.contains(
        vyhledavani, case=False, na=False
    )
    filtr_df = filtr_df[maska]

# Přehledové metriky
col1, col2 = st.columns(2)
col1.metric("Počet zobrazených záznamů", len(filtr_df))
col2.metric("Celkem lokomotiv v databázi", len(seznam_loko))

st.markdown("---")

# Tabulka s daty
st.subheader("📋 Přehled závad")
st.dataframe(filtr_df, use_container_width=True, height=500)