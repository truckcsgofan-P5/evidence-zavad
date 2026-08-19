from datetime import datetime
import io
import github
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Evidence závad lokomotiv", layout="wide", page_icon="🚆"
)

# Skrytí vrchního systémového menu a zápatí Streamlitu
st.markdown(
    """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
    """,
    unsafe_allow_index=True,
)

FILE_PATH = "PREDAVKA_ELEKTRONICI_PRO_APPSHEET.xlsx"

# --- DEFINICE SEZNAMU KATEGORIÍ ZÁVAD ---
KATEGORIE_LIST = [
    "Elektrická výzbroj",
    "Mechanická část",
    "IS - Informační systém",
    "Brzdový systém",
    "Spalovací motor / Pohon",
    "Sdělovací a zabezpečovací technika",
    "Klimatizace / Topení",
    "Ostatní",
]


# --- PŘIHLAŠOVACÍ SYSTÉM ---
def prihlaseni_uzivatele():
    """Zobrazí přihlašovací formulář, pokud uživatel není přihlášen."""
    if "prihlasen" not in st.session_state:
        st.session_state["prihlasen"] = False

    if st.session_state["prihlasen"]:
        return True

    st.title("🔒 Přihlášení do aplikace")
    st.info("Pro přístup k evidenci závad se prosím přihlaste.")

    with st.form("login_form"):
        uzivatel = st.text_input(
            "Uživatelské jméno:", placeholder="Zadejte uživatelské jméno"
        )
        heslo = st.text_input(
            "Heslo:", type="password", placeholder="Zadejte heslo"
        )
        submit_login = st.form_submit_button("Přihlásit se")

        if submit_login:
            povoleni_uzivatele = st.secrets.get("users", {})
            if (
                uzivatel in povoleni_uzivatele
                and str(povoleni_uzivatele[uzivatel]) == heslo
            ):
                st.session_state["prihlasen"] = True
                st.session_state["uzivatel_jmeno"] = uzivatel
                st.rerun()
            else:
                st.error("❌ Nesprávné uživatelské jméno nebo heslo.")

    return False


# Pokud uživatel není přihlášen, aplikace dál nepokračuje
if not prihlaseni_uzivatele():
    st.stop()

# --- SIDEBAR: INFORMACE O UŽIVATELI A ODHLÁŠENÍ ---
with st.sidebar:
    st.write(
        f"👤 Přihlášen: **{st.session_state.get('uzivatel_jmeno', 'Uživatel')}**"
    )
    if st.button("🚪 Odhlásit se"):
        st.session_state["prihlasen"] = False
        st.rerun()


# --- POMOCNÉ FUNKCE ---
def formatuj_lokomotivu(text):
    """Sjednotí formát označení lokomotivy tak, aby za prvními 3 číslicemi byla mezera."""
    if not text:
        return ""
    cisty_text = str(text).replace(" ", "").strip()
    if len(cisty_text) > 3:
        return f"{cisty_text[:3]} {cisty_text[3:]}"
    return cisty_text


def ulozit_df_do_bytes(df_to_save):
    """Bezpečně převede DataFrame na bajty Excelu s českým formátem data."""
    df_copy = df_to_save.copy()

    if "Datum" in df_copy.columns:
        df_copy["Datum"] = pd.to_datetime(
            df_copy["Datum"], errors="coerce"
        ).dt.strftime("%d.%m.%Y")

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")
    df_copy.to_excel(writer, sheet_name="Sheet1", index=False)
    writer.close()
    return output.getvalue()


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

    if "Datum" in df.columns:
        df["Datum"] = pd.to_datetime(
            df["Datum"], dayfirst=True, errors="coerce"
        )

    if "Lokomotiva" in df.columns:
        df["Lokomotiva"] = df["Lokomotiva"].apply(formatuj_lokomotivu)

    # Kontrola existence sloupce Kategorie (pokud v starém Excelu chybí)
    if "Kategorie" not in df.columns:
        df["Kategorie"] = "Neuvedeno"
    else:
        df["Kategorie"] = df["Kategorie"].fillna("Neuvedeno")

    return df


df = load_data()

# Záložky aplikace
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
    st.title("📋 Přehled závad lokomotiv")

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        seznam_loko = sorted(
            [str(x) for x in df["Lokomotiva"].dropna().unique()]
        )
        vybrane_loko = st.multiselect(
            "Filtr podle lokomotivy:",
            options=seznam_loko,
            placeholder="Vyberte lokomotivy...",
        )
    with col_f2:
        vybrane_kategorie = st.multiselect(
            "Filtr podle kategorie:",
            options=KATEGORIE_LIST,
            placeholder="Vyberte kategorie...",
        )
    with col_f3:
        vyhledavani = st.text_input(
            "Hledat v popisu nebo poznámce:",
            placeholder="Napište hledaný text...",
        )

    filtr_df = df.copy()
    if vybrane_loko:
        filtr_df = filtr_df[
            filtr_df["Lokomotiva"].astype(str).isin(vybrane_loko)
        ]
    if vybrane_kategorie:
        filtr_df = filtr_df[filtr_df["Kategorie"].isin(vybrane_kategorie)]
    if vyhledavani:
        maska = filtr_df["Popis závady"].astype(str).str.contains(
            vyhledavani, case=False, na=False
        ) | filtr_df["Poznámka"].astype(str).str.contains(
            vyhledavani, case=False, na=False
        )
        filtr_df = filtr_df[maska]

    st.dataframe(
        filtr_df,
        use_container_width=True,
        height=500,
        column_config={
            "Datum": st.column_config.DateColumn(
                "Datum", format="DD.MM.YYYY"
            ),
            "ID": st.column_config.NumberColumn("ID", format="%d"),
            "Kategorie": st.column_config.SelectboxColumn(
                "Kategorie", options=KATEGORIE_LIST
            ),
        },
    )

# TAB 2: Formulář pro zadání nové závady
with tab_novy:
    st.title("➕ Zapsat novou závadu")

    with st.form("form_zavada", clear_on_submit=True):
        col_n1, col_n2 = st.columns(2)
        with col_n1:
            loko_input = st.text_input(
                "Označení lokomotivy (např. 814 190):",
                placeholder="Např. 814 190",
            )
            kategorie_input = st.selectbox(
                "Kategorie závady:", options=KATEGORIE_LIST
            )
        with col_n2:
            datum_input = st.date_input(
                "Datum zjištění závady:", format="DD.MM.YYYY"
            )

        popis_input = st.text_area(
            "Popis závady:", placeholder="Detailní popis zjištěné závady..."
        )
        poznamka_input = st.text_input(
            "Poznámka (volitelné):",
            placeholder="Např. objednané díly, způsob opravy...",
        )

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
                loko_formatted = formatuj_lokomotivu(loko_input)

                novy_radek = pd.DataFrame(
                    [
                        {
                            "ID": nove_id,
                            "Lokomotiva": loko_formatted,
                            "Kategorie": kategorie_input,
                            "Datum": pd.to_datetime(datum_input),
                            "Popis závady": popis_input.strip(),
                            "Poznámka": poznamka_input.strip(),
                        }
                    ]
                )

                upraveny_df = pd.concat([df, novy_radek], ignore_index=True)

                if "GITHUB_TOKEN" in st.secrets:
                    excel_bytes = ulozit_df_do_bytes(upraveny_df)

                    g = github.Github(st.secrets["GITHUB_TOKEN"])
                    repo = g.get_repo(st.secrets["REPO_NAME"])
                    contents = repo.get_contents(FILE_PATH)

                    repo.update_file(
                        contents.path,
                        f"Přidána nová závada ID {nove_id} (autor: {st.session_state.get('uzivatel_jmeno')})",
                        excel_bytes,
                        contents.sha,
                    )
                    st.success(
                        f"Závada pro lokomotivu {loko_formatted} byla úspěšně uložena pod ID {nove_id}!"
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
            formatuj_lokomotivu(radek["Lokomotiva"])
            if pd.notna(radek["Lokomotiva"])
            else ""
        )

        puvodni_kat = (
            str(radek["Kategorie"]) if pd.notna(radek["Kategorie"]) else ""
        )
        kat_index = (
            KATEGORIE_LIST.index(puvodni_kat)
            if puvodni_kat in KATEGORIE_LIST
            else 0
        )

        if pd.notna(radek["Datum"]):
            try:
                puvodni_datum = pd.to_datetime(radek["Datum"]).date()
            except Exception:
                puvodni_datum = datetime.today().date()
        else:
            puvodni_datum = datetime.today().date()

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
            col_e1, col_e2 = st.columns(2)
            with col_e1:
                loko_edit = st.text_input("Lokomotiva:", value=puvodni_loko)
                kategorie_edit = st.selectbox(
                    "Kategorie závady:", options=KATEGORIE_LIST, index=kat_index
                )
            with col_e2:
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
                    df.at[idx, "Lokomotiva"] = formatuj_lokomotivu(loko_edit)
                    df.at[idx, "Kategorie"] = kategorie_edit
                    df.at[idx, "Datum"] = pd.to_datetime(datum_edit)
                    df.at[idx, "Popis závady"] = popis_edit.strip()
                    df.at[idx, "Poznámka"] = poznamka_edit.strip()

                    if "GITHUB_TOKEN" in st.secrets:
                        excel_bytes = ulozit_df_do_bytes(df)

                        g = github.Github(st.secrets["GITHUB_TOKEN"])
                        repo = g.get_repo(st.secrets["REPO_NAME"])
                        contents = repo.get_contents(FILE_PATH)

                        repo.update_file(
                            contents.path,
                            f"Úprava závady ID {vybrane_id} (autor: {st.session_state.get('uzivatel_jmeno')})",
                            excel_bytes,
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

        if pd.notna(radek_del["Datum"]):
            try:
                datum_zobraz = pd.to_datetime(radek_del["Datum"]).strftime(
                    "%d.%m.%Y"
                )
            except Exception:
                datum_zobraz = str(radek_del["Datum"])
        else:
            datum_zobraz = ""

        st.warning(
            f"**Chystáte se smazat závadu ID {vybrane_id_del}:**\n\n"
            f"* **Lokomotiva:** {radek_del['Lokomotiva']}\n"
            f"* **Kategorie:** {radek_del.get('Kategorie', 'Neuvedeno')}\n"
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
                        excel_bytes = ulozit_df_do_bytes(upraveny_df)

                        g = github.Github(st.secrets["GITHUB_TOKEN"])
                        repo = g.get_repo(st.secrets["REPO_NAME"])
                        contents = repo.get_contents(FILE_PATH)

                        repo.update_file(
                            contents.path,
                            f"Smazána závada ID {vybrane_id_del} (autor: {st.session_state.get('uzivatel_jmeno')})",
                            excel_bytes,
                            contents.sha,
                        )
                        st.success(
                            f"Závada ID {vybrane_id_del} byla úspěšně smazána!"
                        )
                        st.cache_data.clear()
                except Exception as e:
                    st.error(f"Chyba při mazání záznamu: {e}")
