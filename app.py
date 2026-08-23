import base64
from datetime import datetime, timedelta
import hmac
import io
import json
import urllib.request

from google import genai
from google.genai import types
import openpyxl
import pandas as pd
import requests
import streamlit as st
from github import Github, GithubException
from streamlit_cookies_controller import CookieController
from streamlit_pdf_viewer import pdf_viewer

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


# --- POMOCNÁ FUNKCE PRO IMGBB ---
def nahraj_na_imgbb(image_bytes):
    """Nahraje obrázek na ImgBB a vrátí jeho URL adresu."""
    api_key = st.secrets.get("IMGBB_API_KEY")
    if not api_key:
        st.error("❌ V `secrets.toml` chybí `IMGBB_API_KEY`!")
        return None
        
    url = "https://api.imgbb.com/1/upload"
    payload = {
        "key": api_key,
        "image": base64.b64encode(image_bytes).decode('utf-8')
    }
    
    try:
        response = requests.post(url, data=payload)
        if response.status_code == 200:
            return response.json()['data']['url']
        else:
            st.error(f"Chyba ImgBB API: {response.text}")
            return None
    except Exception as e:
        st.error(f"Chyba při nahrávání fotky: {e}")
        return None


# --- POMOCNÁ FUNKCE PRO DATUM, SVÁTEK A POČASÍ ---
@st.cache_data(ttl=1800)
def ziskej_info_hlavicka():
    dnes = datetime.now()
    datum_str = dnes.strftime("%d.%m.%Y")

    svatek_jmeno = "Neznámo"
    try:
        req = urllib.request.Request(
            "https://svatkyapi.cz/api/day",
            headers={"User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, timeout=4) as resp:
            data = json.loads(resp.read().decode())
            svatek_jmeno = data.get("name", "Neznámo")
    except Exception:
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


# --- INICIALIZACE COOKIES PRO ZAPAMATOVÁNÍ ---
controller = CookieController()

# Název cookie pro uložení přihlášeného uživatele
COOKIE_NAME = "evidence_zavad_user"


# --- AUTENTIZACE ---
def prihlaseni_uzivatele():
    # 1. Kontrola, zda již máme uloženo v session_state
    if st.session_state.get("prihlasen", False):
        return True

    # 2. Kontrola, zda existuje platná cookie v prohlížeči
    saved_user = controller.get(COOKIE_NAME)
    povoleni_uzivatele = st.secrets.get("users", {})

    if saved_user and saved_user in povoleni_uzivatele:
        st.session_state["prihlasen"] = True
        st.session_state["uzivatel_jmeno"] = saved_user
        return True

    # 3. Zobrazení přihlašovacího formuláře
    st.title("🔒 Přihlášení do aplikace")
    st.info("Pro přístup k evidenci závad se prosím přihlaste.")

    with st.form("login_form"):
        uzivatel = st.text_input(
            "Uživatelské jméno:", placeholder="Zadejte uživatelské jméno"
        )
        heslo = st.text_input(
            "Heslo:", type="password", placeholder="Zadejte heslo"
        )
        zapamatovat = st.checkbox("Zapamatovat si přihlášení (na 30 dní)")

        submit_login = st.form_submit_button("Přihlásit se")

        if submit_login:
            if uzivatel in povoleni_uzivatele:
                ulozene_heslo = str(povoleni_uzivatele[uzivatel])
                if hmac.compare_digest(ulozene_heslo, heslo):
                    st.session_state["prihlasen"] = True
                    st.session_state["uzivatel_jmeno"] = uzivatel

                    # Pokud zaškrtl zapamatování, uložíme uživatele do cookie na 30 dní
                    if zapamatovat:
                        datum_expirace = datetime.now() + timedelta(days=30)
                        controller.set(
                            COOKIE_NAME,
                            uzivatel,
                            expires=datum_expirace,
                            same_site="lax",
                        )

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
        f"👤 Přihlášen: **{st.session_state.get('uzivatel_jmeno', 'Uživatel')}**  |  "
        f"📅 {datum_dnes}  |  "
        f"🎉 Svátek: {svatek_dnes}  |  "
        f"🌤️ Počasí (Val. Meziříčí): {pocasi_valmez}"
    )

with col_btn:
    if st.button("🚪 Odhlásit", key="logout_top", use_container_width=True):
        st.session_state["prihlasen"] = False
        st.session_state["uzivatel_jmeno"] = None
        # Smazání uložené cookie při odhlášení
        controller.remove(COOKIE_NAME)
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
            g = Github(st.secrets["GITHUB_TOKEN"])
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
        g = Github(st.secrets["GITHUB_TOKEN"])
        repo = g.get_repo(st.secrets["REPO_NAME"])
        file_content = repo.get_contents(FILE_PATH)
        df = pd.read_excel(io.BytesIO(file_content.decoded_content))
    else:
        df = pd.read_excel(FILE_PATH)
    return df

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
    
    if "Fotka" not in df.columns:
        df["Fotka"] = ""
        
    return df


df = load_data()

# --- ZÁLOŽKY APLIKACE ---
tab_prehled, tab_novy, tab_edit, tab_smazat, tab_pdf, tab_foto, tab_ai = st.tabs(
    [
        "📋 Přehled a úprava",
        "➕ Přidat závadu",
        "✏️ Detailní úprava",
        "🗑️ Smazat závadu",
        "📄 Technická dokumentace",
        "🖼️ Fotodokumentace",
        "🤖 Gemini Asistent",
    ]
)

# TAB 1: Přehled
with tab_prehled:
    st.title("📋 Přehled a úprava závad")
    
    if "msg_tab1" in st.session_state:
        st.success(st.session_state["msg_tab1"])
        del st.session_state["msg_tab1"]

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

    # Převeďte sloupec Datum na datetime, aby byl kompatibilní s DateColumn
    if "Datum" in filtr_df.columns:
        filtr_df["Datum"] = pd.to_datetime(filtr_df["Datum"], errors="coerce")

    edited_df = st.data_editor(
        filtr_df,
        use_container_width=True,
        height=500,
        num_rows="fixed",
        disabled=["ID"],
        column_config={
            "Datum": st.column_config.DateColumn("Datum", format="DD.MM.YYYY"),
            "ID": st.column_config.NumberColumn("ID", format="%d"),
            "Kategorie": st.column_config.SelectboxColumn(
                "Kategorie", options=KATEGORIE_LIST
            ),
            "Fotka": st.column_config.LinkColumn("Fotka"),
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
                df.loc[i, "Datum"] = pd.to_datetime(row["Datum"]).strftime("%Y-%m-%d")
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
                df.loc[i, "Fotka"] = (
                    str(row["Fotka"]).strip()
                    if pd.notna(row.get("Fotka"))
                    else ""
                )

        ok, err = ulozit_databazi(df, "Hromadná úprava z tabulky")
        if ok:
            st.session_state["msg_tab1"] = "✅ Všechny změny z tabulky byly úspěšně uloženy!"
            st.rerun()
        else:
            st.error(f"Chyba při ukládání: {err}")

# TAB 2: Nová závada s podporou Gemini a ImgBB
with tab_novy:
    st.title("➕ Zapsat novou závadu")

    if "msg_tab2" in st.session_state:
        st.success(st.session_state["msg_tab2"])
        del st.session_state["msg_tab2"]

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
            fotka_input = st.file_uploader(
                "Nahrát fotku závady (volitelné):", 
                type=["png", "jpg", "jpeg"]
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
            url_fotky = ""
            if fotka_input is not None:
                with st.spinner("Nahrávám fotku na ImgBB..."):
                    obrazek_bytes = fotka_input.getvalue()
                    imgbb_url = nahraj_na_imgbb(obrazek_bytes)
                    if imgbb_url:
                        url_fotky = imgbb_url

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
                        "Fotka": url_fotky,
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
                st.session_state["msg_tab2"] = f"✅ Závada byla úspěšně uložena pod ID {nove_id}!"
                st.rerun()
            else:
                st.error(f"Chyba při ukládání: {err}")

# TAB 3: Detailní úprava
with tab_edit:
    st.title("✏️ Úprava existující závady")
    
    if "msg_tab3" in st.session_state:
        st.success(st.session_state["msg_tab3"])
        del st.session_state["msg_tab3"]

    if df.empty or "ID" not in df.columns:
        st.warning("V databázi nejsou žádné záznamy k úpravě.")
    else:
        seznam_id = df["ID"].dropna().astype(int).tolist()
        vybrane_id = st.selectbox(
            "Vyberte ID závady k úpravě:", options=seznam_id, key="select_edit_id"
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

        puvodni_fotka = str(radek.get("Fotka", ""))

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
                fotka_edit_file = st.file_uploader(
                    "Nahrát novou fotku (nahradí původní):", 
                    type=["png", "jpg", "jpeg"]
                )

            smazat_fotku_checkbox = False
            if puvodni_fotka and puvodni_fotka != "nan":
                st.markdown(f"📎 Aktuální odkaz na fotku: [{puvodni_fotka}]({puvodni_fotka})")
                smazat_fotku_checkbox = st.checkbox("❌ Smazat stávající fotku (odstranit odkaz)")

            popis_edit = st.text_area(
                "Popis závady:", value=str(radek.get("Popis závady", ""))
            )
            poznamka_edit = st.text_input(
                "Poznámka:", value=str(radek.get("Poznámka", ""))
            )

            submit_edit = st.form_submit_button("Uložit změny", type="primary")

        if submit_edit:
            cilova_fotka_url = puvodni_fotka
            
            if smazat_fotku_checkbox:
                cilova_fotka_url = ""
            elif fotka_edit_file is not None:
                with st.spinner("Nahrávám novou fotku na ImgBB..."):
                    novy_obrazek_bytes = fotka_edit_file.getvalue()
                    nove_imgbb_url = nahraj_na_imgbb(novy_obrazek_bytes)
                    if nove_imgbb_url:
                        cilova_fotka_url = nove_imgbb_url

            idx = df[df["ID"] == vybrane_id].index[0]
            df.at[idx, "Lokomotiva"] = formatuj_lokomotivu(loko_edit)
            df.at[idx, "Kategorie"] = kategorie_edit
            df.at[idx, "Datum"] = pd.to_datetime(datum_edit).strftime("%Y-%m-%d")
            df.at[idx, "Popis závady"] = popis_edit.strip()
            df.at[idx, "Poznámka"] = poznamka_edit.strip()
            df.at[idx, "Fotka"] = cilova_fotka_url.strip()

            ok, err = ulozit_databazi(df, f"Úprava závady ID {vybrane_id}")
            if ok:
                st.session_state["msg_tab3"] = f"✅ Závada ID {vybrane_id} byla úspěšně aktualizována!"
                st.rerun()
            else:
                st.error(f"Chyba při ukládání: {err}")

# TAB 4: Smazat
with tab_smazat:
    st.title("🗑️ Odstranění závady")
    
    if "msg_tab4" in st.session_state:
        st.success(st.session_state["msg_tab4"])
        del st.session_state["msg_tab4"]

    if df.empty or "ID" not in df.columns:
        st.warning("V databázi nejsou žádné záznamy ke smazání.")
    else:
        seznam_id_del = df["ID"].dropna().astype(int).tolist()
        vybrane_id_del = st.selectbox(
            "Vyberte ID závady k smazání:", options=seznam_id_del
        )
        radek_del = df[df["ID"] == vybrane_id_del].iloc[0]

        datum_zobraz = (
            pd.to_datetime(radek_del["Datum"]).strftime("%d.%m.%Y")
            if pd.notna(radek_del.get("Datum"))
            and not pd.isna(pd.to_datetime(radek_del.get("Datum")))
            else "Neuvedeno"
        )

        st.markdown("### 📄 Detail vybraného záznamu k odstranění:")
        st.info(
            f"**ID závady:** {radek_del.get('ID', '')}\n\n"
            f"**Lokomotiva:** {radek_del.get('Lokomotiva', '')}\n\n"
            f"**Kategorie:** {radek_del.get('Kategorie', '')}\n\n"
            f"**Datum zjištění:** {datum_zobraz}\n\n"
            f"**Popis závady:** {radek_del.get('Popis závady', 'Bez popisu')}\n\n"
            f"**Poznámka:** {radek_del.get('Poznámka', 'Bez poznámky')}\n\n"
            f"**Fotka:** {radek_del.get('Fotka', 'Bez fotky')}"
        )

        st.warning(
            f"⚠️ Opravdu chcete trvale smazat tuto závadu pro lokomotivu **{radek_del['Lokomotiva']}**?"
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
                    st.session_state["msg_tab4"] = "✅ Záznam byl úspěšně trvale smazán!"
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

# --- TAB: Dokumentace (PDF + Word) ---
with tab_pdf:
    st.title("📄 Technická dokumentace (PDF a Word)")
    st.caption(
        "Ukládání a prohlížení PDF i Word dokumentů v podsložkách podle řad."
    )

    RADY_LOKOMOTIV = ["844", "842", "814", "954", "Ostatní"]

    # Načtení konfiguračních údajů ze Streamlit Secrets
    try:
        github_token = st.secrets["GITHUB_TOKEN"]
        repo_name = st.secrets["GITHUB_REPO"]
        g = Github(github_token)
        repo = g.get_repo(repo_name)
    except Exception as e:
        st.error(
            "⚠️ Nepodařilo se načíst GITHUB_TOKEN nebo GITHUB_REPO ze Secrets."
        )
        st.stop()

    # ---------------------------------------------------------
    # 1. NAHRÁVÁNÍ DOKUMENTŮ A SPRÁVA PODSLOŽEK
    # ---------------------------------------------------------
    st.subheader("➕ Nahrát nový dokument (PDF nebo Word)")

    col_rada, col_sub = st.columns([1, 1])

    with col_rada:
        zvolena_rada_pdf = st.selectbox(
            "Vyberte řadu lokomotivy:", RADY_LOKOMOTIV, key="upload_rada_pdf"
        )

    # Načtení existujících podsložek z GitHubu pro danou řadu
    base_path_upload = f"docs_pdf/{zvolena_rada_pdf}"
    existujici_podslozky = []

    try:
        items = repo.get_contents(base_path_upload)
        existujici_podslozky = [
            item.name for item in items if item.type == "dir"
        ]
    except GithubException:
        existujici_podslozky = []

    moznosti_podslozek = [
        "-- Vytvořit novou podsložku --"
    ] + existujici_podslozky

    with col_sub:
        vybrana_podslozka = st.selectbox(
            "Vyberte podsložku:", moznosti_podslozek, key="upload_sub_pdf"
        )

    # Pokud uživatel zvolí vytvoření nové podsložky
    nazev_podslozky = ""
    if vybrana_podslozka == "-- Vytvořit novou podsložku --":
        nazev_podslozky = st.text_input(
            "Název nové podsložky (např. Motor nebo Schémata):",
            key="new_sub_input_pdf",
        ).strip()
    else:
        nazev_podslozky = vybrana_podslozka

    uploaded_doc = st.file_uploader(
        "Vyberte PDF nebo Word soubor:",
        type=["pdf", "doc", "docx"],
        key="github_doc_uploader",
    )

    if uploaded_doc is not None:
        if not nazev_podslozky:
            st.warning("⚠️ Prosím zadejte nebo vyberte podsložku!")
        else:
            safe_sub = (
                nazev_podslozky.replace(" ", "_")
                .replace("/", "_")
                .replace("\\", "_")
            )
            file_path_doc = (
                f"docs_pdf/{zvolena_rada_pdf}/{safe_sub}/{uploaded_doc.name}"
            )
            file_bytes_doc = uploaded_doc.getvalue()

            if st.button("🚀 Uložit dokument na GitHub", key="btn_save_doc"):
                with st.spinner("Ukládám dokument do GitHub repozitáře..."):
                    try:
                        try:
                            contents = repo.get_contents(file_path_doc)
                            repo.update_file(
                                path=file_path_doc,
                                message=f"Aktualizace dokumentu {uploaded_doc.name} v {safe_sub}",
                                content=file_bytes_doc,
                                sha=contents.sha,
                            )
                            st.success(
                                f"✅ Dokument '{uploaded_doc.name}' byl aktualizován!"
                            )
                        except GithubException:
                            repo.create_file(
                                path=file_path_doc,
                                message=f"Přidán dokument {uploaded_doc.name} do {safe_sub}",
                                content=file_bytes_doc,
                            )
                            st.success(
                                f"✅ Dokument '{uploaded_doc.name}' byl úspěšně uložen do složky **{safe_sub}**!"
                            )

                        st.rerun()
                    except Exception as ex:
                        st.error(f"Při ukládání došlo k chybě: {ex}")

    st.divider()

    # ---------------------------------------------------------
    # 2. PROHLÍŽEČ A MAZÁNÍ DOKUMENTŮ
    # ---------------------------------------------------------
    st.subheader("📂 Prohlížet uložené dokumenty")

    col_view1, col_view2 = st.columns([1, 1])

    with col_view1:
        vybrana_rada_view_pdf = st.selectbox(
            "Zobrazit řadu:", RADY_LOKOMOTIV, key="view_rada_pdf"
        )

    # Načtení podsložek pro prohlížení
    base_path_view = f"docs_pdf/{vybrana_rada_view_pdf}"
    podslozky_view = []

    try:
        items_view = repo.get_contents(base_path_view)
        podslozky_view = [
            item.name for item in items_view if item.type == "dir"
        ]
    except GithubException:
        podslozky_view = []

    with col_view2:
        if podslozky_view:
            vybrana_sub_view = st.selectbox(
                "Vyberte podsložku ke zobrazení:",
                podslozky_view,
                key="view_sub_pdf",
            )
        else:
            vybrana_sub_view = None
            st.info("Pro tuto řadu zatím neexistují žádné podsložky.")

    if vybrana_sub_view:
        target_doc_folder = (
            f"docs_pdf/{vybrana_rada_view_pdf}/{vybrana_sub_view}"
        )

        try:
            folder_docs = repo.get_contents(target_doc_folder)
            valid_extensions = (".pdf", ".doc", ".docx")
            doc_files = [
                f
                for f in folder_docs
                if f.name.lower().endswith(valid_extensions)
            ]
        except GithubException:
            doc_files = []

        if doc_files:
            soubor_dict = {f.name: f for f in doc_files}
            zvoleny_nazev = st.selectbox(
                "Vyberte konkrétní dokument k zobrazení:",
                list(soubor_dict.keys()),
            )

            selected_file_obj = soubor_dict[zvoleny_nazev]

            # Načtení souboru z GitHubu
            file_url = selected_file_obj.download_url
            headers = {"Authorization": f"token {github_token}"}
            response = requests.get(file_url, headers=headers)

            if response.status_code == 200:
                file_data = response.content
                is_pdf = zvoleny_nazev.lower().endswith(".pdf")
                cdn_url = f"https://cdn.jsdelivr.net/gh/{repo_name}@main/{selected_file_obj.path}"

                # Ovládací tlačítka
                col_btn1, col_btn2, col_btn3 = st.columns([1.5, 1, 1])

                with col_btn1:
                    if is_pdf:
                        st.link_button(
                            "🔗 Otevřít PDF v novém okně", url=cdn_url
                        )
                    else:
                        # Google Docs Viewer pro Word dokumenty (.doc/.docx)
                        google_viewer_url = (
                            f"https://docs.google.com/viewer?url={cdn_url}"
                        )
                        st.link_button(
                            "🔗 Otevřít Word v novém okně", url=google_viewer_url
                        )

                with col_btn2:
                    st.download_button(
                        label="💾 Stáhnout",
                        data=file_data,
                        file_name=zvoleny_nazev,
                        mime="application/pdf"
                        if is_pdf
                        else "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="dl_doc_btn",
                    )

                with col_btn3:
                    if st.button(
                        "🗑️ Smazat", key="del_doc_btn", type="secondary"
                    ):
                        try:
                            repo.delete_file(
                                path=selected_file_obj.path,
                                message=f"Smazán dokument {zvoleny_nazev} ze složky {vybrana_sub_view}",
                                sha=selected_file_obj.sha,
                            )
                            st.success(
                                f"Dokument '{zvoleny_nazev}' byl smazán."
                            )
                            st.rerun()
                        except Exception as del_err:
                            st.error(f"Chyba při mazání: {del_err}")

                st.divider()

                # NÁHLED UVNITŘ APLIKACE
                if is_pdf:
                    pdf_viewer(input=file_data, width=700, height=800)
                else:
                    st.info(
                        f"📄 **Soubor '{zvoleny_nazev}' je dokument Microsoft Word.**\n\n"
                        "Z bezpečnostních důvodů nelze Word zobrazit přímo v malém náhledu. "
                        "Použijte tlačítko **🔗 Otevřít Word v novém okně** výše pro plné zobrazení bez stahování, nebo tlačítko **💾 Stáhnout**."
                    )
            else:
                st.error("Dokument se nepodařilo načíst z GitHubu.")
        else:
            st.info(
                f"Ve složce **{vybrana_sub_view}** zatím nejsou žádné dokumenty."
            )
# --- TAB: Fotodokumentace ---
with tab_foto:
    st.title("🖼️ Fotodokumentace")
    st.caption("Ukládání a správa fotografií v podsložkách podle řad na GitHubu.")

    RADY_LOKOMOTIV = ["844", "842", "814", "954", "Ostatní"]

    # Načtení konfiguračních údajů ze Streamlit Secrets
    try:
        github_token = st.secrets["GITHUB_TOKEN"]
        repo_name = st.secrets["GITHUB_REPO"]
        g = Github(github_token)
        repo = g.get_repo(repo_name)
    except Exception as e:
        st.error(
            "⚠️ Nepodařilo se načíst GITHUB_TOKEN nebo GITHUB_REPO ze Secrets."
        )
        st.stop()

    # ---------------------------------------------------------
    # 1. NAHRÁVÁNÍ FOTOGRAFIÍ A SPRÁVA PODSLOŽEK
    # ---------------------------------------------------------
    st.subheader("➕ Nahrát novou fotografii")

    col_rada, col_sub = st.columns([1, 1])

    with col_rada:
        zvolena_rada_foto = st.selectbox(
            "Vyberte řadu lokomotivy:", RADY_LOKOMOTIV, key="upload_rada_foto"
        )

    # Načtení existujících podsložek z GitHubu pro danou řadu
    base_path_upload = f"docs_foto/{zvolena_rada_foto}"
    existujici_podslozky = []

    try:
        items = repo.get_contents(base_path_upload)
        existujici_podslozky = [item.name for item in items if item.type == "dir"]
    except GithubException:
        existujici_podslozky = []

    moznosti_podslozek = ["-- Vytvořit novou podsložku --"] + existujici_podslozky

    with col_sub:
        vybrana_podslozka = st.selectbox(
            "Vyberte podsložku:", moznosti_podslozek, key="upload_sub_foto"
        )

    # Pokud uživatel zvolí vytvoření nové podsložky
    nazev_podslozky = ""
    if vybrana_podslozka == "-- Vytvořit novou podsložku --":
        nazev_podslozky = st.text_input(
            "Název nové podsložky (např. 844-001 nebo Prevodovka):",
            key="new_sub_input",
        ).strip()
    else:
        nazev_podslozky = vybrana_podslozka

    uploaded_img = st.file_uploader(
        "Vyberte fotografii:",
        type=["jpg", "jpeg", "png", "webp"],
        key="github_img_uploader",
    )

    if uploaded_img is not None:
        if not nazev_podslozky:
            st.warning("⚠️ Prosím zadejte nebo vyberte podsložku!")
        else:
            # Očištění názvu podsložky od nepovolených znaků
            safe_sub = (
                nazev_podslozky.replace(" ", "_")
                .replace("/", "_")
                .replace("\\", "_")
            )
            file_path_foto = (
                f"docs_foto/{zvolena_rada_foto}/{safe_sub}/{uploaded_img.name}"
            )
            file_bytes_foto = uploaded_img.getvalue()

            if st.button("🚀 Uložit fotku na GitHub", key="btn_save_foto"):
                with st.spinner("Ukládám fotografii do GitHub repozitáře..."):
                    try:
                        try:
                            contents = repo.get_contents(file_path_foto)
                            repo.update_file(
                                path=file_path_foto,
                                message=f"Aktualizace fotky {uploaded_img.name} v {safe_sub}",
                                content=file_bytes_foto,
                                sha=contents.sha,
                            )
                            st.success(
                                f"✅ Fotografie '{uploaded_img.name}' byla aktualizována!"
                            )
                        except GithubException:
                            repo.create_file(
                                path=file_path_foto,
                                message=f"Přidána fotka {uploaded_img.name} do {safe_sub}",
                                content=file_bytes_foto,
                            )
                            st.success(
                                f"✅ Fotografie '{uploaded_img.name}' byla úspěšně uložena do složky **{safe_sub}**!"
                            )

                        st.rerun()
                    except Exception as ex:
                        st.error(f"Při ukládání došlo k chybě: {ex}")

    st.divider()

   # ---------------------------------------------------------
   # 2. PROHLÍŽEČ A MAZÁNÍ FOTOGRAFIÍ
   # ---------------------------------------------------------
    st.subheader("🖼️ Prohlížet fotodokumentaci")

    col_view1, col_view2 = st.columns([1, 1])

    with col_view1:
        vybrana_rada_view_foto = st.selectbox(
            "Zobrazit řadu:", RADY_LOKOMOTIV, key="view_rada_foto"
        )

    # Načtení podsložek pro prohlížení
    base_path_view = f"docs_foto/{vybrana_rada_view_foto}"
    podslozky_view = []

    try:
        items_view = repo.get_contents(base_path_view)
        podslozky_view = [
            item.name for item in items_view if item.type == "dir"
        ]
    except GithubException:
        podslozky_view = []

    with col_view2:
        if podslozky_view:
            vybrana_sub_view = st.selectbox(
                "Vyberte podsložku ke zobrazení:",
                podslozky_view,
                key="view_sub_foto",
            )
        else:
            vybrana_sub_view = None
            st.info("Pro tuto řadu zatím neexistují žádné podsložky.")

    if vybrana_sub_view:
        target_foto_folder = (
            f"docs_foto/{vybrana_rada_view_foto}/{vybrana_sub_view}"
        )

        try:
            folder_imgs = repo.get_contents(target_foto_folder)
            img_extensions = (".jpg", ".jpeg", ".png", ".webp")
            image_files = [
                f
                for f in folder_imgs
                if f.name.lower().endswith(img_extensions)
            ]
        except GithubException:
            image_files = []

        if image_files:
            st.write(
                f"Nalezeno **{len(image_files)}** fotografií ve složce **{vybrana_sub_view}**:"
            )

            # Zobrazení fotografií v mřížce (2 sloupce vedle sebe)
            cols = st.columns(2)
            headers = {"Authorization": f"token {github_token}"}

            for idx, img_obj in enumerate(image_files):
                col = cols[idx % 2]

                res = requests.get(img_obj.download_url, headers=headers)
                if res.status_code == 200:
                    view_url_img = f"https://cdn.jsdelivr.net/gh/{repo_name}@main/{img_obj.path}"

                    with col:
                        st.image(
                            res.content,
                            caption=img_obj.name,
                            use_container_width=True,
                        )

                        c1, c2, c3 = st.columns([1, 1, 1])
                        with c1:
                            st.link_button(
                                "🔗 Otevřít",
                                url=view_url_img,
                                use_container_width=True,
                            )
                        with c2:
                            st.download_button(
                                label="💾 Stáhnout",
                                data=res.content,
                                file_name=img_obj.name,
                                mime="image/jpeg",
                                key=f"dl_{img_obj.sha}",
                                use_container_width=True,
                            )
                        with c3:
                            # Tlačítko pro smazání fotografíe
                            if st.button(
                                "🗑️ Smazat",
                                key=f"del_{img_obj.sha}",
                                type="secondary",
                                use_container_width=True,
                            ):
                                try:
                                    repo.delete_file(
                                        path=img_obj.path,
                                        message=f"Smazána fotka {img_obj.name} ze složky {vybrana_sub_view}",
                                        sha=img_obj.sha,
                                    )
                                    st.success(
                                        f"Fotografie '{img_obj.name}' byla smazána."
                                    )
                                    st.rerun()
                                except Exception as del_err:
                                    st.error(
                                        f"Chyba při mazání souboru: {del_err}"
                                    )

                        st.write("---")
        else:
            st.info(
                f"Ve složce **{vybrana_sub_view}** zatím nejsou žádné obrázky."
            )
