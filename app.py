from datetime import datetime
import io
import github
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Evidence závad lokomotiv", layout="wide", page_icon="🚆"
)

FILE_PATH = "PREDAVKA_ELEKTRONICI_PRO_APPSHEET.xlsx"


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

    # Převod na datový typ datetime pro chronologické řazení
    if "Datum" in df.columns:
        df["Datum"] = pd.to_datetime(
            df["Datum"], dayfirst=True, errors="coerce"
        )

    return df


df = load_data()

# 4 záložky v aplikaci
tab_prehled, tab_novy, tab_edit, tab_smazat = st.tabs(
    [
        "📋 Přehled závad",
        "➕ Přidat novou závadu",
        "✏️ Úprava závady",
        "🗑️ Smazat závadu",
    ]
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

    # Zobrazení s českým formátem a chronologickým řazením
    st.dataframe(
        filtr_df,
        use_container_width=True,
        height=500,
        column_config={
            "Datum": st.column_config.DateColumn(
                "Datum", format="DD.MM.YYYY"
            ),
            "ID": st.column_config.NumberColumn("ID", format="%d"),
        },
    )

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
                nove_id = (
                    int(df["ID"].max()) + 1
                    if not df.empty and "ID" in df
                    else 1
                )

                novy_radek = pd.DataFrame(
                    [
                        {
                            "ID": nove_id,
                            "Lokomotiva": loko_input.strip(),
                            "Datum": pd.to_datetime(datum_input),
                            "Popis závady": popis_input.strip(),
                            "Poznámka": poznamka_input.strip(),
                        }
                    ]
                )

                upraveny_df = pd.concat([df, novy_radek], ignore_index=True)

                if "GITHUB_TOKEN" in st.secrets:
                    output = io.BytesIO()
                    with pd.ExcelWriter(
                        output, engine="openpyxl"
                    ) as writer:
                        upraveny_df.to_excel(
                            writer, index=False, date_format="DD.MM.YYYY"
                        )

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
                        f"Závada pro lokomotivu {loko_input} byla úspěšně uložena pod ID {nove_id}!"
                    )
                    st.cache_data.clear()
            except Exception as e:
                st.error(f"Chyba při ukládání: {e}")

# TAB 3: Formulář pro úpravu stávající závady
with tab_edit:
    st.title("✏️ Úprava existující závady")

    if df.empty or "ID" not in df.columns:
        st.warning("V databázi nejsou žádné záznamy k úpravě.")
    else:
        seznam_id = df["ID"].dropna().astype(int).tolist()
        vybrane_id = st.selectbox(
            "Vyberte ID závady, kterou chcete upravit:",
            options=seznam_id,
            key="select_edit_id",
        )

        radek = df[df["ID"] == vybrane_id].iloc[0]

        puvodni_loko = (
            str(radek["Lokomotiva"]) if pd.notna(radek["Lokomotiva"]) else ""
        )

        puvodni_datum = (
            radek["Datum"].date()
            if pd.notna(radek["Datum"])
            else datetime.today().date()
        )
        puvodni_popis = (
            str(radek["Popis závady"])
            if pd.notna(radek["Popis závady"])
            else ""
        )
        puvodni_poznamka = (
            str(radek["Poznámka"]) if pd.notna(radek["Poznámka"]) else ""
        )

        with st.form("form_edit_zavada"):
            st.info(f"Úprava závady ID: **{vybrane_id}**")
            loko_edit = st.text_input("Lokomotiva:", value=puvodni_loko)
            datum_edit = st.date_input(
                "Datum:", value=puvodni_datum, format="DD.MM.YYYY"
            )
            popis_edit = st.text_area("Popis závady:", value=puvodni_popis)
            poznamka_edit = st.text_input(
                "Poznámka (např. stav opravy):", value=puvodni_poznamka
            )

            submit_edit = st.form_submit_button("Uložit změny")

        if submit_edit:
            if not loko_edit or not popis_edit:
                st.error("Lokomotiva a popis závady nesmí být prázdné.")
            else:
                try:
                    idx = df[df["ID"] == vybrane_id].index[0]
                    df.at[idx, "Lokomotiva"] = loko_edit.strip()
                    df.at[idx, "Datum"] = pd.to_datetime(datum_edit)
                    df.at[idx, "Popis závady"] = popis_edit.strip()
                    df.at[idx, "Poznámka"] = poznamka_edit.strip()

                    if "GITHUB_TOKEN" in st.secrets:
                        output = io.BytesIO()
                        with pd.ExcelWriter(
                            output, engine="openpyxl"
                        ) as writer:
                            df.to_excel(
                                writer, index=False, date_format="DD.MM.YYYY"
                            )

                        g = github.Github(st.secrets["GITHUB_TOKEN"])
                        repo = g.get_repo(st.secrets["REPO_NAME"])
                        contents = repo.get_contents(FILE_PATH)

                        repo.update_file(
                            contents.path,
                            f"Úprava závady ID {vybrane_id}",
                            output.getvalue(),
                            contents.sha,
                        )
                        st.success(
                            f"Závada ID {vybrane_id} byla úspěšně aktualizována!"
                        )
                        st.cache_data.clear()
                except Exception as e:
                    st.error(f"Chyba při ukládání změn: {e}")

# TAB 4: Smazat závadu
with tab_smazat:
    st.title("🗑️ Odstranění závady")

    if df.empty or "ID" not in df.columns:
        st.warning("V databázi nejsou žádné záznamy ke smazání.")
    else:
        seznam_id_del = df["ID"].dropna().astype(int).tolist()
        vybrane_id_del = st.selectbox(
            "Vyberte ID závady, kterou chcete trvale smazat:",
            options=seznam_id_del,
            key="select_del_id",
        )

        radek_del = df[df["ID"] == vybrane_id_del].iloc[0]

        datum_zobraz = (
            radek_del["Datum"].strftime("%d.%m.%Y")
            if pd.notna(radek_del["Datum"])
            else ""
        )

        st.warning(
            f"**Chystáte se smazat závadu ID {vybrane_id_del}:**\n\n"
            f"* **Lokomotiva:** {radek_del['Lokomotiva']}\n"
            f"* **Datum:** {datum_zobraz}\n"
            f"* **Popis:** {radek_del['Popis závady']}\n"
            f"* **Poznámka:** {radek_del['Poznámka']}"
        )

        potvrzeni = st.checkbox(
            f"Rozumím, opravdu chci trvale smazat závadu ID {vybrane_id_del}"
        )

        if st.button("🗑️ Trvale smazat záznam", type="primary"):
            if not potvrzeni:
                st.error(
                    "Pro smazání musíte nejprve zaškrtnout potvrzovací políčko."
                )
            else:
                try:
                    upraveny_df = df[df["ID"] != vybrane_id_del].copy()

                    if "GITHUB_TOKEN" in st.secrets:
                        output = io.BytesIO()
                        with pd.ExcelWriter(
                            output, engine="openpyxl"
                        ) as writer:
                            upraveny_df.to_excel(
                                writer, index=False, date_format="DD.MM.YYYY"
                            )

                        g = github.Github(st.secrets["GITHUB_TOKEN"])
                        repo = g.get_repo(st.secrets["REPO_NAME"])
                        contents = repo.get_contents(FILE_PATH)

                        repo.update_file(
                            contents.path,
                            f"Smazána závada ID {vybrane_id_del}",
                            output.getvalue(),
                            contents.sha,
                        )
                        st.success(
                            f"Závada ID {vybrane_id_del} byla úspěšně smazána!"
                        )
                        st.cache_data.clear()
                except Exception as e:
                    st.error(f"Chyba při mazání záznamu: {e}")
