from datetime import datetime
import hmac
import io
import json
import urllib.request
import github
from google import genai
from google.genai import types
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Evidence závad lokomotiv", layout="wide", page_icon="🚆"
)

st.markdown(
    """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)

FILE_PATH = "PREDAVKA_ELEKTRONICI_PRO_APPSHEET.xlsx"

KATEGORIE_LIST = [
    "Elektrická výzbroj",
    "Mechanická část",
    "Brzdový systém",
    "IS - Infosystém",
    "Dobíjení",
    "Spalovací motor / Pohon",
    "Sdělovací a zabezpečovací technika",
    "Klimatizace / Topení",
    "WC - systém",
    "Ostatní",
]


# --- POMOCNÁ FUNKCE PRO DATUM, SVÁTEK A POČASÍ ---
@st.cache_data(ttl=1800)
def ziskej_info_hlavicka():
    dnes = datetime.now()
    datum_str = dnes.strftime("%d.%m.%Y")

    # Načtení svátku s primárním a záložním zdrojem
    svatek_jmeno = "Neznámo"

    # 1. Pokus: SvatkyAPI.cz
    try:
        req = urllib.request.Request(
            "https://svatkyapi.cz/api/day",
            headers={"User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, timeout=4) as resp:
            data = json.loads(resp.read().decode())
            svatek_jmeno = data.get("name", "Neznámo")
    except Exception:
        # 2. Pokus (Záloha): svatek.jdem.cz
        try:
            req = urllib.request.Request(
                "https://svatek.jdem.cz/json",
                headers={"User-Agent": "Mozilla/5.0"},
            )
            with urllib.request.urlopen(req, timeout=4) as resp:
                data = json.loads(resp.read().decode())
                if isinstance(data, list) and len(data) > 0:
                    svatek_jmeno = data[0].get("name", "Neznámo")
        except Exception:
            pass

    # Načtení počasí pro Valašské Meziříčí (souřadnice: 49.4718, 17.9712)
    pocasi_str = "Neznámo"
    try:
        url_pocasi = "https://api.open-meteo.com/v1/forecast?latitude=49.4718&longitude=17.9712&current_weather=true"
        req_poc = urllib.request.Request(
            url_pocasi, headers={"User-Agent": "Mozilla/5.0"}
        )
        with urllib.request.urlopen(req_poc, timeout=4) as resp:
            p_data = (
                json.loads(resp.read().decode())
                .get("current_weather", {})
                .get("temperature")
            )
            if p_data is not None:
                pocasi_str = f"{p_data} °C"
    except Exception:
        pass

    return datum_str, svatek_jmeno, pocasi_str


# --- GEMINI AI POMOCNÉ FUNKCE ---
def získej_gemini_klient():
    api_key = st.secrets.get("GEMINI_API_KEY")
    if not api_key:
        st.error("❌ V `secrets.toml` chybí `GEMINI_API_KEY`!")
        return None
    return genai.Client(api_key=api_key)


def analyzuj_zavadu_gemini(popis_raw):
    """Pomocí Gemini vybere kategorii a upraví neformální text na odborný."""
    client = získej_gemini_klient()
    if not client:
        return None

    prompt = f"""
    Jsi expert na železniční kolejová vozidla a údržbu lokomotiv.
    Uživatel zadal následující neformální popis závady: "{popis_raw}"

    Úkoly:
    1. Vyber nejvhodnější kategorii výhradně z tohoto seznamu: {KATEGORIE_LIST}
    2. Přeformuluj popis do spisovné, profesionální a stručné technické češtiny.

    Vrať odpověď výhradně jako platný JSON objekt s klíči "kategorie" a "upraveny_popis".
    """

    try:
        response = client.models.generate_content(
            model="gemini-3.6-flash",
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json"
            ),
        )
        return json.loads(response.text)
    except Exception as e:
        st.error(f"Chyba při komunikaci s Gemini API: {e}")
        return None


def dotaz_na_gemini(dotaz, df):
    """Položí dotaz modelu Gemini s kontextem celé databáze v CSV."""
    client = získej_gemini_klient()
    if not client:
        return "Není k dispozici API klíč."

    csv_data = df.to_csv(index=False)
    prompt = f"""
    Jsi inteligentní asistent správce lokomotivního parku.
    Zde jsou aktuální data o závadách ve formátu CSV:

    {csv_data}

    Odpověz věcně, přesně a přehledně v češtině na dotaz uživatele:
    "{dotaz}"
    """

    try:
        response = client.models.generate_content(
            model="gemini-3.6-flash", contents=prompt
        )
        return response.text
    except Exception as e:
        return f"Chyba při zpracování dotazu: {e}"


# --- AUTENTIZACE ---
def prihlaseni_uzivatele():
    if st.session_state.get("prihlasen", False):
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
            if uzivatel in povoleni_uzivatele:
                ulozene_heslo = str(povoleni_uzivatele[uzivatel])
                if hmac.compare_digest(ulozene_heslo, heslo):
                    st.session_state["prihlasen"] = True
                    st.session_state["uzivatel_jmeno"] = uzivatel
                    st.rerun()

            st.error("❌ Nesprávné uživatelské jméno nebo heslo.")
    return False


if not prihlaseni_uzivatele():
    st.stop()

# --- COMPACT LIŠTA S UŽIVATELEM A INFORMACEMI ---
datum_dnes, svatek_dnes, pocasi_valmez = ziskej_info_hlavicka()

col_info, col_btn = st.columns([5, 1], vertical_alignment="bottom")

with col_info:
    st.caption(
        f"👤 Přihlášen: **{st.session_state.get('uzivatel_jmeno', 'Uživatel')}**  |  "
        f"📅 {datum_dnes}  |  "
        f"🎉 Svátek: {svatek_dnes}  |  "
        f"🌤️ Počasí (Val. Meziříčí): {pocasi_valmez}"
    )

with col_btn:
    if st.button("🚪 Odhlásit", key="logout_top", use_container_width=True):
        st.session_state["prihlasen"] = False
        st.rerun()

st.divider()


# --- POMOCNÉ FUNKCE PRO SOUBORY ---
def formatuj_lokomotivu(text):
    if not text:
        return ""
    cisty_text = str(text).replace(" ", "").strip()
    return (
        f"{cisty_text[:3]} {cisty_text[3:]}"
        if len(cisty_text) > 3
        else cisty_text
    )


def ulozit_df_do_bytes(df_to_save):
    df_copy = df_to_save.copy()
    if "Datum" in df_copy.columns:
        df_copy["Datum"] = pd.to_datetime(
            df_copy["Datum"], errors="coerce"
        ).dt.strftime("%d.%m.%Y")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_copy.to_excel(writer, sheet_name="Sheet1", index=False)
    return output.getvalue()


def ulozit_databazi(df_to_save, commit_msg):
    try:
        excel_bytes = ulozit_df_do_bytes(df_to_save)
        autor = st.session_state.get("uzivatel_jmeno", "Neznámý")

        if "GITHUB_TOKEN" in st.secrets:
            g = github.Github(st.secrets["GITHUB_TOKEN"])
            repo = g.get_repo(st.secrets["REPO_NAME"])
            contents = repo.get_contents(FILE_PATH)
            repo.update_file(
                contents.path,
                f"{commit_msg} (autor: {autor})",
                excel_bytes,
                contents.sha,
            )
        else:
            with open(FILE_PATH, "wb") as f:
                f.write(excel_bytes)

        st.cache_data.clear()
        return True, None
    except Exception as e:
        return False, str(e)


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

    df["Kategorie"] = (
        df["Kategorie"].fillna("Neuvedeno")
        if "Kategorie" in df.columns
        else "Neuvedeno"
    )
    return df


df = load_data()

# --- ZÁLOŽKY APLIKACE ---
tab_prehled, tab_novy, tab_edit, tab_smazat, tab_ai = st.tabs(
    [
        "📋 Přehled a úprava",
        "➕ Přidat novou závadu",
        "✏️ Detailní úprava",
        "🗑️ Smazat závadu",
        "🤖 Gemini AI Asistent",
    ]
)

# TAB 1: Přehled
with tab_prehled:
    st.title("📋 Přehled a úprava závad")
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        seznam_loko = sorted(
            [str(x) for x in df["Lokomotiva"].dropna().unique()]
        )
        vybrane_loko = st.multiselect(
            "Filtr podle lokomotivy:", options=seznam_loko
        )
    with col_f2:
        vybrane_kategorie = st.multiselect(
            "Filtr podle kategorie:", options=KATEGORIE_LIST
        )
    with col_f3:
        vyhledavani = st.text_input("Hledat v popisu nebo poznámce:")

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

    edited_df = st.data_editor(
        filtr_df,
        use_container_width=True,
        height=500,
        num_rows="fixed",
        disabled=["ID"],
        column_config={
            "Datum": st.column_config.DateColumn(
                "Datum", format="DD.MM.YYYY"
            ),
            "ID": st.column_config.NumberColumn("ID", format="%d"),
            "Kategorie": st.column_config.SelectboxColumn(
                "Kategorie", options=KATEGORIE_LIST
            ),
        },
        key="editor_zavad",
    )

    if st.button("💾 Uložit změny v tabulce", type="primary"):
        for idx, row in edited_df.iterrows():
            main_idx = df[df["ID"] == row["ID"]].index
            if not main_idx.empty:
                i = main_idx[0]
                df.loc[i, "Lokomotiva"] = formatuj_lokomotivu(
                    row["Lokomotiva"]
                )
                df.loc[i, "Kategorie"] = row["Kategorie"]
                df.loc[i, "Datum"] = pd.to_datetime(row["Datum"])
                df.loc[i, "Popis závady"] = (
                    str(row["Popis závady"]).strip()
                    if pd.notna(row["Popis závady"])
                    else ""
                )
                df.loc[i, "Poznámka"] = (
                    str(row["Poznámka"]).strip()
                    if pd.notna(row["Poznámka"])
                    else ""
                )

        ok, err = ulozit_databazi(df, "Hromadná úprava z tabulky")
        if ok:
            st.success("✅ Všechny změny byly uloženy!")
            st.rerun()
        else:
            st.error(f"Chyba při ukládání: {err}")

# TAB 2: Nová závada s podporou Gemini
with tab_novy:
    st.title("➕ Zapsat novou závadu")

    default_kat = st.session_state.get("ai_kategorie", KATEGORIE_LIST[0])
    default_popis = st.session_state.get("ai_popis", "")
    kat_idx = (
        KATEGORIE_LIST.index(default_kat)
        if default_kat in KATEGORIE_LIST
        else 0
    )

    with st.form("form_zavada"):
        col_n1, col_n2 = st.columns(2)
        with col_n1:
            loko_input = st.text_input(
                "Označení lokomotivy:", placeholder="Např. 814 190"
            )
            kategorie_input = st.selectbox(
                "Kategorie závady:", options=KATEGORIE_LIST, index=kat_idx
            )
        with col_n2:
            datum_input = st.date_input(
                "Datum zjištění závady:", format="DD.MM.YYYY"
            )

        popis_input = st.text_area(
            "Popis závady:",
            value=default_popis,
            placeholder="Můžete zadat i nespisovně, např.: 'bliká kontrolka tlaku oleje a píská to'...",
        )
        poznamka_input = st.text_input(
            "Poznámka (volitelné):", placeholder="Např. objednané díly..."
        )

        col_b1, col_b2 = st.columns([1, 1])
        with col_b1:
            submit = st.form_submit_button(
                "💾 Uložit závadu", type="primary", use_container_width=True
            )
        with col_b2:
            ai_btn = st.form_submit_button(
                "🪄 Analyzovat text přes Gemini AI", use_container_width=True
            )

    if ai_btn:
        if not popis_input:
            st.warning("Před analýzou vyplňte popis závady.")
        else:
            with st.spinner("Gemini analyzuje text..."):
                res = analyzuj_zavadu_gemini(popis_input)
                if res:
                    st.session_state["ai_kategorie"] = res.get(
                        "kategorie", default_kat
                    )
                    st.session_state["ai_popis"] = res.get(
                        "upraveny_popis", popis_input
                    )
                    st.success("✅ Text byl upraven a kategorie navržena!")
                    st.rerun()

    if submit:
        if not loko_input or not popis_input:
            st.error("Vyplňte prosím lokomotivu a popis závady.")
        else:
            nove_id = (
                int(df["ID"].max()) + 1 if not df.empty and "ID" in df else 1
            )
            novy_radek = pd.DataFrame(
                [
                    {
                        "ID": nove_id,
                        "Lokomotiva": formatuj_lokomotivu(loko_input),
                        "Kategorie": kategorie_input,
                        "Datum": pd.to_datetime(datum_input),
                        "Popis závady": popis_input.strip(),
                        "Poznámka": poznamka_input.strip(),
                    }
                ]
            )

            upraveny_df = pd.concat([df, novy_radek], ignore_index=True)
            ok, err = ulozit_databazi(
                upraveny_df, f"Přidána nová závada ID {nove_id}"
            )
            if ok:
                st.session_state["ai_popis"] = ""
                st.session_state["ai_kategorie"] = KATEGORIE_LIST[0]
                st.success(f"Závada byla uložena pod ID {nove_id}!")
                st.rerun()
            else:
                st.error(f"Chyba při ukládání: {err}")

# TAB 3: Detailní úprava
with tab_edit:
    st.title("✏️ Úprava existující závady")
    if df.empty or "ID" not in df.columns:
        st.warning("V databázi nejsou žádné záznamy k úpravě.")
    else:
        seznam_id = df["ID"].dropna().astype(int).tolist()
        vybrane_id = st.selectbox(
            "Vyberte ID závady k úpravě:", options=seznam_id
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

        puvodni_datum = (
            pd.to_datetime(radek["Datum"]).date()
            if pd.notna(radek["Datum"])
            and not pd.isna(pd.to_datetime(radek["Datum"]))
            else datetime.today().date()
        )

        with st.form("form_edit_zavada"):
            col_e1, col_e2 = st.columns(2)
            with col_e1:
                loko_edit = st.text_input("Lokomotiva:", value=puvodni_loko)
                kategorie_edit = st.selectbox(
                    "Kategorie:", options=KATEGORIE_LIST, index=kat_index
                )
            with col_e2:
                datum_edit = st.date_input(
                    "Datum:", value=puvodni_datum, format="DD.MM.YYYY"
                )

            popis_edit = st.text_area(
                "Popis závady:", value=str(radek.get("Popis závady", ""))
            )
            poznamka_edit = st.text_input(
                "Poznámka:", value=str(radek.get("Poznámka", ""))
            )

            submit_edit = st.form_submit_button("Uložit změny")

        if submit_edit:
            idx = df[df["ID"] == vybrane_id].index[0]
            df.at[idx, "Lokomotiva"] = formatuj_lokomotivu(loko_edit)
            df.at[idx, "Kategorie"] = kategorie_edit
            df.at[idx, "Datum"] = pd.to_datetime(datum_edit)
            df.at[idx, "Popis závady"] = popis_edit.strip()
            df.at[idx, "Poznámka"] = poznamka_edit.strip()

            ok, err = ulozit_databazi(df, f"Úprava závady ID {vybrane_id}")
            if ok:
                st.success(f"Závada ID {vybrane_id} byla aktualizována!")
                st.rerun()
            else:
                st.error(f"Chyba při ukládání: {err}")

# TAB 4: Smazat
with tab_smazat:
    st.title("🗑️ Odstranění závady")
    if df.empty or "ID" not in df.columns:
        st.warning("V databázi nejsou žádné záznamy ke smazání.")
    else:
        seznam_id_del = df["ID"].dropna().astype(int).tolist()
        vybrane_id_del = st.selectbox(
            "Vyberte ID závady k smazání:", options=seznam_id_del
        )
        radek_del = df[df["ID"] == vybrane_id_del].iloc[0]

        st.warning(
            f"**Smazat závadu ID {vybrane_id_del} pro lokomotivu {radek_del['Lokomotiva']}?**"
        )
        potvrzeni = st.checkbox("Rozumím, opravdu chci trvale smazat")

        if st.button("🗑️ Trvale smazat záznam", type="primary"):
            if not potvrzeni:
                st.error("Zaškrtněte potvrzovací políčko.")
            else:
                upraveny_df = df[df["ID"] != vybrane_id_del].copy()
                ok, err = ulozit_databazi(
                    upraveny_df, f"Smazána závada ID {vybrane_id_del}"
                )
                if ok:
                    st.success("Záznam byl smazán!")
                    st.rerun()
                else:
                    st.error(f"Chyba: {err}")

# TAB 5: Gemini AI Chat nad databází
with tab_ai:
    st.title("🤖 Gemini AI Asistent")
    st.caption(
        "Ptejte se na statistiky, historii oprav nebo doporučení k celému parku lokomotiv."
    )

    dotaz_user = st.text_input(
        "Váš dotaz pro AI:",
        placeholder="Např. Jaké byly nejčastější závady na lokomotivách v kategorii Klimatizace?",
    )

    if st.button("💬 Zeptat se Gemini", type="primary"):
        if dotaz_user:
            with st.spinner("Gemini analyzuje databázi..."):
                odpoved = dotaz_na_gemini(dotaz_user, df)
                st.markdown("### Odpověď Gemini:")
                st.info(odpoved)
        else:
            st.warning("Napište dotaz.")
