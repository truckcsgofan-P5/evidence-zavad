import io
import github
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Evidence závad lokomotiv", layout="wide", page_icon="🚆"
)

FILE_PATH = "PREDAVKA_ELEKTRONICI_PRO_APPSHEET.xlsx"


def formatuj_datum_str(val):
    """Pomocná funkce pro sjednocení formátu na DD.MM.RRRR"""
    if pd.isna(val) or not val:
        return ""
    try:
        # Pokud je již v datetimu nebo standardním řetězci
        dt = pd.to_datetime(val, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%d.%m.%Y")
    except Exception:
        pass
    return str(val).strip()


# Načtení dat přímo z GitHubu
@st.cache_data(ttl=5)
def load_data():
    if "GITHUB_TOKEN" in st.secrets:
        g = github.Github(st.secrets["GITHUB_TOKEN"])
        repo = g.get_repo(st.secrets["REPO_NAME"])
        file_content = repo.get_contents(FILE_PATH)
        df = pd.read_excel(io.BytesIO(file_content.decoded_content))
    else:
        df = pd.read_excel(FILE_PATH)

    # Sjednocení formátu data u všech načtených řádků na DD.MM.RRRR
    if "Datum" in df.columns:
        df["Datum"] = df["Datum"].apply(formatuj_datum_str)

    return df


df = load_data()

# Záložky v aplikaci
tab_prehled, tab_novy = st.tabs(
    ["📋 Přehled závad", "➕ Přidat novou závadu"]
)

# TAB 1: Přehled
with tab_prehled:
    st.title("🚆 Přehled závad lokomotiv")

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        seznam_loko = sorted(
            [str(x) for x in df["Lokomotiva"].dropna().unique()]
        )
        vybrane_loko = st.multiselect(
            "Filtr podle lokomotivy:", options=seznam_loko
        )
    with col_f2:
        vyhledavani = st.text_input("Hledat v popisu nebo poznámce:")

    filtr_df = df.copy()
    if vybrane_loko:
        filtr_df = filtr_df[
            filtr_df["Lokomotiva"].astype(str).isin(vybrane_loko)
        ]
    if vyhledavani:
        maska = filtr_df["Popis závady"].astype(str).str.contains(
            vyhledavani, case=False, na=False
        ) | filtr_df["Poznámka"].astype(str).str.contains(
            vyhledavani, case=False, na=False
        )
        filtr_df = filtr_df[maska]

    st.dataframe(filtr_df, use_container_width=True, height=500)

# TAB 2: Formulář pro zadání nové závady
with tab_novy:
    st.title("➕ Zapsat novou závadu")

    with st.form("form_zavada", clear_on_submit=True):
        loko_input = st.text_input("Označení lokomotivy (např. 814 190):")
        datum_input = st.date_input(
            "Datum zjištění závady:", format="DD.MM.YYYY"
        )
        popis_input = st.text_area("Popis závady:")
        poznamka_input = st.text_input("Poznámka (volitelné):")

        submit = st.form_submit_button("Uložit závadu")

    if submit:
        if not loko_input or not popis_input:
            st.error("Vyplňte prosím lokomotivu a popis závady.")
        else:
            try:
                # Výpočet nového ID
                nove_id = (
                    int(df["ID"].max()) + 1
                    if not df.empty and "ID" in df
                    else 1
                )

                # Přísný formát DD.MM.RRRR (např. 05.06.2026)
                datum_str = datum_input.strftime("%d.%m.%Y")

                # Vytvoření nového řádku
                novy_radek = pd.DataFrame(
                    [
                        {
                            "ID": nove_id,
                            "Lokomotiva": loko_input.strip(),
                            "Datum": datum_str,
                            "Popis závady": popis_input.strip(),
                            "Poznámka": poznamka_input.strip(),
                        }
                    ]
                )

                upraveny_df = pd.concat([df, novy_radek], ignore_index=True)

                # Uložení do GitHub repozitáře
                if "GITHUB_TOKEN" in st.secrets:
                    output = io.BytesIO()
                    with pd.ExcelWriter(
                        output, engine="openpyxl"
                    ) as writer:
                        upraveny_df.to_excel(writer, index=False)

                    g = github.Github(st.secrets["GITHUB_TOKEN"])
                    repo = g.get_repo(st.secrets["REPO_NAME"])
                    contents = repo.get_contents(FILE_PATH)

                    repo.update_file(
                        contents.path,
                        f"Přidána nová závada ID {nove_id}",
                        output.getvalue(),
                        contents.sha,
                    )
                    st.success(
                        f"Závada pro lokomotivu {loko_input} byla úspěšně uložena s datem {datum_str} pod ID {nove_id}!"
                    )
                    st.cache_data.clear()
                else:
                    st.warning(
                        "Aplikace běží lokálně bez nastaveného GitHub tokenu. Data nebyla zapsána na server."
                    )

            except Exception as e:
                st.error(f"Chyba při ukládání: {e}")
