import base64
from datetime import datetime, timedelta
import hmac
import io
import json
import time
import os
import tempfile
from moviepy import VideoFileClip
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
from zoneinfo import ZoneInfo  # Import pro české časové pásmo

# --- TOTO DEJTE ÚPLNĚ NA ZAČÁTEK SOUBORU (před vytváření tabů) ---
try:
    github_token = st.secrets["GITHUB_TOKEN"]
    repo_name = st.secrets["GITHUB_REPO"]
    g = Github(github_token)
    repo = g.get_repo(repo_name)
except Exception as e:
    st.error("⚠️ Nepodařilo se načíst GITHUB_TOKEN nebo GITHUB_REPO ze Secrets.")
    st.stop()

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

CHAT_FILE_PATH = "chat_messages.json"

def nacti_chat_z_githubu():
    """Načte historii chatu z GitHub repozitáře."""
    try:
        file_content = repo.get_contents(CHAT_FILE_PATH)
        json_data = file_content.decoded_content.decode("utf-8")
        return json.loads(json_data)
    except Exception:
        # Pokud soubor ještě neexistuje nebo nastala chyba, vrátíme prázdný seznam
        return []


def uloz_chat_na_github(zpravy_list, autor="Neznámý"):
    """Uloží nový seznam zpráv do JSON na GitHub."""
    json_data = json.dumps(zpravy_list, ensure_ascii=False, indent=2)
    try:
        try:
            file_content = repo.get_contents(CHAT_FILE_PATH)
            repo.update_file(
                path=CHAT_FILE_PATH,
                message=f"Chat: nová zpráva od {autor}",
                content=json_data,
                sha=file_content.sha,
            )
        except Exception:
            # Vytvoření souboru, pokud na GitHubu ještě neexistuje
            repo.create_file(
                path=CHAT_FILE_PATH,
                message=f"Chat: vytvořen soubor chatu ({autor})",
                content=json_data,
            )
        return True
    except Exception as e:
        st.error(f"Chyba při ukládání chatu na GitHub: {e}")
        return False

KATEGORIE_LIST = [
    "Elektrická výzbroj",
    "Mechanická část",
    "Brzdový systém",
    "IS - Infosystém",
    "Dobíjení",
    "Spalovací motor / Pohon",
    "VZ - Radiostanice",
    "Klimatizace / Topení",
    "WC - systém",
    "Ostatní",
]


# --- POMOCNÁ FUNKCE PRO IMGBB ---
def nahrat_na_imgbb(file_bytes, api_key):
    """Nahrání fotky na ImgBB API"""
    try:
        url = "https://api.imgbb.com/1/upload"
        payload = {
            "key": api_key,
            "image": base64.b64encode(file_bytes).decode("utf-8")
        }
        res = requests.post(url, data=payload)
        if res.status_code == 200:
            return res.json()["data"]["url"]
        else:
            st.error(f"Chyba při nahrávání fotky na ImgBB: {res.text}")
            return None
    except Exception as e:
        st.error(f"Chyba spojení s ImgBB: {e}")
        return None

def nahrat_video_na_github(file_bytes, file_name, repo_name, token):
    """Nahrání videa přímo do složky videa/ v GitHub repozitáři"""
    try:
        g = Github(token)
        repo = g.get_repo(repo_name)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        cesta_v_repo = f"videa/{timestamp}_{file_name}"
        
        repo.create_file(
            path=cesta_v_repo,
            message=f"Upload videa: {file_name}",
            content=file_bytes,
            branch="main"
        )
        
        # Sestavení přímé URL adresy na raw soubor v GitHubu
        raw_url = f"https://raw.githubusercontent.com/{repo_name}/main/{cesta_v_repo}"
        return raw_url
    except Exception as e:
        st.error(f"Chyba při nahrávání videa na GitHub: {e}")
        return None

    # ==================== 1. POMOCNÁ FUNKCE PRO UPLOAD A KONVERZI VIDEA ====================


def zpracuj_a_nahraj_video(uploaded_file, repo, slozka_na_githubu="videa"):
    """Převede .mov na .mp4 (pokud je potřeba) a nahraje soubor na GitHub."""
    file_bytes = uploaded_file.read()
    puvodni_nazev = uploaded_file.name
    pripona = puvodni_nazev.split(".")[-1].lower()

    if pripona == "mov":
        st.info(
            "⏳ Nahráno video ve formátu MOV. Probíhá automatická konverze na MP4..."
        )

        with tempfile.NamedTemporaryFile(
            suffix=".mov", delete=False
        ) as tmp_mov:
            tmp_mov.write(file_bytes)
            tmp_mov_path = tmp_mov.name

        tmp_mp4_path = tmp_mov_path.replace(".mov", ".mp4")

        try:
            clip = VideoFileClip(tmp_mov_path)
            clip.write_videofile(
                tmp_mp4_path,
                codec="libx264",
                audio_codec="aac",
                preset="ultrafast",
                logger=None,
            )
            clip.close()

            with open(tmp_mp4_path, "rb") as f:
                final_bytes = f.read()

            novy_nazev = puvodni_nazev.rsplit(".", 1)[0] + ".mp4"
        finally:
            if os.path.exists(tmp_mov_path):
                os.remove(tmp_mov_path)
            if os.path.exists(tmp_mp4_path):
                os.remove(tmp_mp4_path)
    else:
        final_bytes = file_bytes
        novy_nazev = puvodni_nazev

    cesta_na_githubu = f"{slozka_na_githubu}/{novy_nazev}"
    repo.create_file(
        path=cesta_na_githubu,
        message=f"Upload videa: {novy_nazev}",
        content=final_bytes,
    )

    return f"https://raw.githubusercontent.com/{repo.full_name}/main/{cesta_na_githubu}"

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
    if st.session_state.get("prihlasen", False):
        return True

    saved_user = controller.get(COOKIE_NAME)
    povoleni_uzivatele = st.secrets.get("users", {})

    def over_uzivatele(jmeno, zapsane_heslo):
        if jmeno not in povoleni_uzivatele:
            return False, None

        zaznam = povoleni_uzivatele[jmeno]

        # 1. Nový formát: pepa = { password = "123", role = "admin" }
        # Ověřujeme přítomnost klíče/atributu password
        if hasattr(zaznam, "password") or (isinstance(zaznam, dict) and "password" in zaznam):
            ulozene_heslo = str(zaznam.get("password") if isinstance(zaznam, dict) else zaznam.password)
            role = str(zaznam.get("role", "viewer") if isinstance(zaznam, dict) else getattr(zaznam, "role", "viewer"))

            if hmac.compare_digest(ulozene_heslo, zapsane_heslo):
                return True, role

        # 2. Starý formát: pepa = "123"
        else:
            ulozene_heslo = str(zaznam)
            if hmac.compare_digest(ulozene_heslo, zapsane_heslo):
                return True, "admin"

        return False, None

    # Automatické přihlášení z Cookie
    if saved_user and saved_user in povoleni_uzivatele:
        zaznam = povoleni_uzivatele[saved_user]
        if hasattr(zaznam, "role") or (isinstance(zaznam, dict) and "role" in zaznam):
            role = str(zaznam.get("role") if isinstance(zaznam, dict) else zaznam.role)
        else:
            role = "admin"

        st.session_state["prihlasen"] = True
        st.session_state["uzivatel_jmeno"] = saved_user
        st.session_state["uzivatel_role"] = role
        return True

    # Formulář pro přihlášení
    st.title("🔒 Přihlášení do aplikace")
    st.info("Pro přístup k evidenci závad se prosím přihlaste.")

    with st.form("login_form"):
        uzivatel = st.text_input("Uživatelské jméno:", placeholder="Zadejte uživatelské jméno")
        heslo = st.text_input("Heslo:", type="password", placeholder="Zadejte heslo")
        zapamatovat = st.checkbox("Zapamatovat si přihlášení (na 30 dní)", value=True)

        submit_login = st.form_submit_button("Přihlásit se")

        if submit_login:
            ok, role = over_uzivatele(uzivatel, heslo)
            if ok:
                st.session_state["prihlasen"] = True
                st.session_state["uzivatel_jmeno"] = uzivatel
                st.session_state["uzivatel_role"] = role

                if zapamatovat:
                    datum_expirace = datetime.now() + timedelta(days=30)
                    controller.set(
                        COOKIE_NAME,
                        uzivatel,
                        expires=datum_expirace,
                        same_site="lax",
                    )
                    time.sleep(0.5)
                st.rerun()
            else:
                st.error("❌ Nesprávné uživatelské jméno nebo heslo.")

    return False


if not prihlaseni_uzivatele():
    st.stop()

# --- COMPACT LIŠTA S UŽIVATELEM A INFORMACEMI ---
datum_dnes, svatek_dnes, pocasi_valmez = ziskej_info_hlavicka()

# Načtení role a přiřazení přehledné ikony
role_user = st.session_state.get("uzivatel_role", "viewer")

# Pomocná práva (vyhodnotí se jako True/False)
is_admin = role_user in ["admin", "SAdmin"]  # True pro Admina i SAdmina
is_sadmin = role_user == "SAdmin"            # True pouze pro SAdmina

ikony_roli = {
    "SAdmin": "🔑 SAdmin",
    "admin": "🔑 Admin",
    "editor": "✏️ Editor",
    "viewer": "👁️ Pouze čtení"
}
zobrazena_role = ikony_roli.get(role_user, role_user.capitalize())

col_info, col_btn = st.columns([5, 1], vertical_alignment="bottom")

with col_info:
    st.caption(
        f"👤 Přihlášen: **{st.session_state.get('uzivatel_jmeno', 'Uživatel')}** ({zobrazena_role})  |  "
        f"📅 {datum_dnes}  |  "
        f"🎉 Svátek: {svatek_dnes}  |  "
        f"🌤️ Počasí (Val. Meziříčí): {pocasi_valmez}"
    )

with col_btn:
    if st.button("🚪 Odhlásit", key="logout_top", use_container_width=True):
        st.session_state["prihlasen"] = False
        st.session_state["uzivatel_jmeno"] = None
        st.session_state["uzivatel_role"] = None
        
        # Bezpečné smazání uložené cookie při odhlášení
        try:
            if controller.get(COOKIE_NAME):
                controller.remove(COOKIE_NAME)
                time.sleep(0.5)
        except Exception:
            pass
            
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
            df_copy["Datum"], dayfirst=True, errors="coerce" # Přidáno dayfirst=True
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

# Pouze tento jeden řádek – zajistí správný typ pro sloupec, aniž by přepsal data
if "Upravil" in df.columns:
    df["Upravil"] = df["Upravil"].astype(object)

# Načtení role z přihlášení
role_user = st.session_state.get("uzivatel_role", "viewer")

# Načtení role uživatele (pokud ji v této části kódu ještě nemáte načtenou)
# role_user = st.session_state.get("uzivatel_role", "viewer")

# Zobrazení odkazu POUZE pro SAdmina
if role_user == "SAdmin":
    # st.success("👑 Vítejte v režimu Super Admin") # Volitelné: jen aby věděl, že má speciální práva
    
    # VARIANTA 1: Streamlit Multipage (pokud máte stránku ve složce 'pages')
    st.page_link("https://prohlidky.streamlit.app", label="Přehled VZ a Radiostanice", icon="📋")
    
    # VARIANTA 2: Pokud je to odkaz na externí web nebo jinou URL
    # st.markdown("[Přejít na portál SAdmin](https://vasedomena.cz/sadmin)")
    
    st.divider() # Vizuální oddělovač od zbytku stránky (Přehledu)

# =========================================================
# VÝPOČET NEPŘEČTENÝCH ZPRÁV V CHATU
# =========================================================
vsechny_zpravy = nacti_chat_z_githubu()
pocet_zprav_celkem = len(vsechny_zpravy)

# 📌 ZMĚNA ZDE: Pokud uživatel chat ještě neotevřel, výchozí stav přečtených je 0!
if "chat_precteno_pocet" not in st.session_state:
    st.session_state["chat_precteno_pocet"] = 0

# Zjistíme, kolik zpráv uživatel už viděl
pocet_videnych = st.session_state["chat_precteno_pocet"]
neprecteno = max(0, pocet_zprav_celkem - pocet_videnych)

# Dynamický název záložky (přidá červenou tečku a číslo, pokud jsou nové zprávy)
nazev_chat_tab = f"💬 Chat (🔴 {neprecteno})" if neprecteno > 0 else "💬 Chat"

# Logika zobrazení tabů podle role:
# viewer -> vidí jen Přehled, Dokumenty, Fotodokumentaci a AI
# editor -> vidí vše kromě Smazat
# admin  -> vidí úplně vše

if role_user in ["admin", "SAdmin"]:
    # Prvky, které má vidět Admin i Super Admin (např. tlačítko Uložit, st.data_editor apod.)
    tab_prehled, tab_novy, tab_edit, tab_smazat, tab_pdf, tab_foto, tab_ai, tab_chat = st.tabs(
        [
            "📋 Přehled a úprava",
            "➕ Přidat závadu",
            "✏️ Detailní úprava",
            "🗑️ Smazat závadu",
            "📄 Technická dokumentace",
            "🖼️ Fotodokumentace",
            "🤖 Gemini Asistent",
            nazev_chat_tab,  # Proměnná s dynamickým názvem
        ]
    )
elif role_user == "editor":
    tab_prehled, tab_novy, tab_edit, tab_pdf, tab_foto, tab_ai, tab_chat = st.tabs(
        [
            "📋 Přehled a úprava",
            "➕ Přidat závadu",
            "✏️ Detailní úprava",
            "📄 Technická dokumentace",
            "🖼️ Fotodokumentace",
            "🤖 Gemini Asistent",
            nazev_chat_tab,  # Proměnná s dynamickým názvem
        ]
    )
    tab_smazat = None  # Editor nemá tab smazat
else:  # viewer
    tab_prehled, tab_pdf, tab_foto, tab_ai = st.tabs(
        [
            "📋 Přehled (pouze čtení)",
            "📄 Technická dokumentace",
            "🖼️ Fotodokumentace",
            "🤖 Gemini Asistent",
        ]
    )
    tab_novy = None
    tab_edit = None
    tab_smazat = None
    tab_chat = None

# TAB 1: Přehled
with tab_prehled:
    st.title("📋 Přehled a úprava závad")
    
    je_editor = role_user in ["admin", "SAdmin", "editor"]

    if "msg_tab1" in st.session_state:
        st.success(st.session_state["msg_tab1"])
        del st.session_state["msg_tab1"]

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        seznam_loko = sorted([str(x) for x in df["Lokomotiva"].dropna().unique()])
        vybrane_loko = st.multiselect("Filtr podle lokomotivy:", options=seznam_loko)
    with col_f2:
        vybrane_kategorie = st.multiselect("Filtr podle kategorie:", options=KATEGORIE_LIST)
    with col_f3:
        vyhledavani = st.text_input("Hledat v popisu nebo poznámce:")

    filtr_df = df.copy()
    if vybrane_loko:
        filtr_df = filtr_df[filtr_df["Lokomotiva"].astype(str).isin(vybrane_loko)]
    if vybrane_kategorie:
        filtr_df = filtr_df[filtr_df["Kategorie"].isin(vybrane_kategorie)]
    if vyhledavani:
        maska = filtr_df["Popis závady"].astype(str).str.contains(
            vyhledavani, case=False, na=False
        ) | filtr_df["Poznámka"].astype(str).str.contains(
            vyhledavani, case=False, na=False
        )
        filtr_df = filtr_df[maska]

    if "Datum" in filtr_df.columns:
        filtr_df["Datum"] = pd.to_datetime(filtr_df["Datum"], dayfirst=True, errors="coerce")
        filtr_df = filtr_df.sort_values(by="Datum", ascending=False)
        
    filtr_df = filtr_df.fillna("")

    textove_sloupce = ["Lokomotiva", "Popis závady", "Poznámka", "Fotka", "Kategorie", "Vytvořil", "Upravil"]
    for col in textove_sloupce:
        if col in filtr_df.columns:
            filtr_df[col] = filtr_df[col].astype(str).replace({"None": "", "nan": "", "<NA>": ""}) 

    if "Fotka" in filtr_df.columns:
        filtr_df["Fotka"] = filtr_df["Fotka"].astype(str).replace({"bez fotky": "", "Bez fotky": ""}).str.strip()        

    for col_autor in ["Vytvořil", "Upravil"]:
        if col_autor not in filtr_df.columns:
            filtr_df[col_autor] = ""

    # Tabulka (nyní správně vnořená pod tab_prehled)
    edited_df = st.data_editor(
        filtr_df,
        use_container_width=True,
        height=500,
        num_rows="fixed",
        disabled=True if not je_editor else ["ID", "Vytvořil", "Upravil"],
        hide_index=True,
        column_order=["ID", "Datum", "Lokomotiva", "Popis závady", "Poznámka", "Fotka", "Kategorie", "Vytvořil", "Upravil"],
        column_config={
            "ID": st.column_config.NumberColumn("ID", format="%d", width=35),
            "Lokomotiva": st.column_config.Column("Lokomotiva", width=60),
            "Datum": st.column_config.DateColumn("Datum", format="DD.MM.YYYY", width=100),
            "Kategorie": st.column_config.SelectboxColumn("Kategorie", options=KATEGORIE_LIST, width=140),
            "Popis závady": st.column_config.Column("Popis závady", width=330),
            "Poznámka": st.column_config.Column("Poznámka", width=200),
            "Fotka": st.column_config.LinkColumn("Fotka", width=100),
            "Vytvořil": st.column_config.Column("Vytvořil", width=90),
            "Upravil": st.column_config.Column("Upravil", width=90),
        },
        key="editor_zavad",
    )

    # Náhled médií
    df_s_fotkou = filtr_df[filtr_df["Fotka"].astype(str).str.startswith("http", na=False)]
    if not df_s_fotkou.empty:
        st.write("---")
        st.subheader("🖼️ Otevřít / přehrát médium")

        vybrana_zavada_id = st.selectbox(
            "Vyberte závadu pro zobrazení:",
            options=df_s_fotkou["ID"].tolist(),
            format_func=lambda x: f"ID {x} - {df_s_fotkou[df_s_fotkou['ID'] == x]['Lokomotiva'].values[0]} ({df_s_fotkou[df_s_fotkou['ID'] == x]['Popis závady'].values[0][:30]}...)",
            key="ios_media_select",
        )

        url_media = df_s_fotkou[df_s_fotkou["ID"] == vybrana_zavada_id]["Fotka"].values[0]

        col_m1, col_m2 = st.columns([1, 2])
        with col_m1:
            st.link_button("🔗 Otevřít odkaz v novém okně", url_media)
        with col_m2:
            if any(ext in url_media.lower() for ext in [".mp4", ".mov"]):
                st.video(url_media)
            else:
                st.image(url_media, width=300, caption=f"Náhled k ID {vybrana_zavada_id}")

# TAB 2: Nová závada
with tab_nova:
    st.title("➕ Přidat novou závadu")

    with st.form("form_nova_zavada", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            # Lze upravit na selectbox podle vašich lokomotiv, např: st.selectbox("Lokomotiva:", SEZNAM_LOKO)
            loko_input = st.text_input("Lokomotiva (např. 742 001):")
            kategorie_input = st.selectbox("Kategorie:", options=KATEGORIE_LIST)
        with col2:
            datum_input = st.date_input("Datum zjištění:", value=datetime.now())

        popis_input = st.text_area("Popis závady *", placeholder="Detailně popište závadu...")
        poznamka_input = st.text_area("Poznámka (volitelné):", placeholder="Doplňující informace...")

        media_input = st.file_uploader(
            "Připojit fotku nebo video (PNG, JPG, MP4, MOV):", 
            type=["png", "jpg", "jpeg", "mp4", "mov"]
        )

        submit_btn = st.form_submit_button("💾 Uložit závadu")

    if submit_btn:
        if not loko_input.strip() or not popis_input.strip():
            st.warning("⚠️ Vyplňte prosím povinná pole: Lokomotiva a Popis závady!")
        else:
            with st.spinner("Ukládám závadu a zpracovávám média..."):
                url_media = ""

                # Zpracování souboru, pokud byl přiložen
                if media_input is not None:
                    file_bytes = media_input.read()
                    file_ext = media_input.name.split(".")[-1].lower()

                    # 1. Fotka -> ImgBB
                    if file_ext in ["png", "jpg", "jpeg"]:
                        imgbb_key = st.secrets.get("IMGBB_API_KEY", "")
                        if imgbb_key:
                            url_media = nahrat_na_imgbb(file_bytes, imgbb_key)
                        else:
                            st.error("❌ V st.secrets chybí IMGBB_API_KEY!")

                    # 2. Video -> GitHub Repozitář
                    elif file_ext in ["mp4", "mov"]:
                        github_token = st.secrets.get("GITHUB_TOKEN", "")
                        repo_name = st.secrets.get("GITHUB_REPO", "")
                        if github_token and repo_name:
                            url_media = nahrat_video_na_github(
                                file_bytes, 
                                media_input.name, 
                                repo_name, 
                                github_token
                            )
                        else:
                            st.error("❌ V st.secrets chybí GITHUB_TOKEN nebo GITHUB_REPO!")

                # Výpočet nového unikátního ID
                nove_id = 1
                if not df.empty and "ID" in df.columns:
                    ids = pd.to_numeric(df["ID"], errors="coerce").dropna()
                    if not ids.empty:
                        nove_id = int(ids.max()) + 1

                # Autor záznamu ze session_state / cookies
                vytvoril_uzivatel = st.session_state.get(
                    "username", 
                    st.session_state.get("prihlaseny_uzivatel", "Neznámý")
                )

                # Nový řádek databáze
                novy_radek = {
                    "ID": nove_id,
                    "Datum": datum_input.strftime("%d.%m.%Y"),
                    "Lokomotiva": loko_input.strip(),
                    "Kategorie": kategorie_input,
                    "Popis závady": popis_input.strip(),
                    "Poznámka": poznamka_input.strip(),
                    "Fotka": url_media if url_media else "",
                    "Vytvořil": vytvoril_uzivatel,
                    "Upravil": ""
                }

                # Přidání řádku do DataFrame a uložení do CSV na GitHub
                df = pd.concat([df, pd.DataFrame([novy_radek])], ignore_index=True)

                if ulozit_databazi(df):
                    st.session_state["msg_tab1"] = f"✅ Závada ID {nove_id} byla úspěšně uložena!"
                    st.rerun()
            

# TAB 3: Detailní úprava
if tab_edit:    
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
                    # Povolení nahrání fotky i videa
                    fotka_edit_file = st.file_uploader(
                        "Nahrát novou fotku nebo video (nahradí původní):", 
                        type=["png", "jpg", "jpeg", "mp4", "mov", "avi"]
                    )
    
                smazat_fotku_checkbox = False
                if puvodni_fotka and puvodni_fotka != "nan":
                    st.markdown(f"📎 Aktuální odkaz na soubor: [{puvodni_fotka}]({puvodni_fotka})")
                    smazat_fotku_checkbox = st.checkbox("❌ Smazat stávající fotku/video (odstranit odkaz)")
    
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
                    file_ext = fotka_edit_file.name.split(".")[-1].lower()
                    
                    if file_ext in ["mp4", "mov", "avi"]:
                        # -- NAHRÁVÁNÍ VIDEA NA GITHUB --
                        with st.spinner("Nahrávám nové video na GitHub..."):
                            video_bytes = fotka_edit_file.getvalue()
                            safe_name = fotka_edit_file.name.replace(" ", "_")
                            github_path = f"docs_zavady_videa/zavada_{vybrane_id}_{safe_name}"
                            
                            try:
                                github_token = st.secrets["GITHUB_TOKEN"]
                                repo_name = st.secrets["GITHUB_REPO"]
                                g = Github(github_token)
                                temp_repo = g.get_repo(repo_name)
                                
                                temp_repo.create_file(
                                    path=github_path,
                                    message=f"Aktualizace/přidání videa k závadě ID {vybrane_id}",
                                    content=video_bytes
                                )
                                cilova_fotka_url = f"https://cdn.jsdelivr.net/gh/{repo_name}@main/{github_path}"
                            except Exception as e:
                                st.error(f"Chyba při nahrávání videa na GitHub: {e}")
                    else:
                        # -- NAHRÁVÁNÍ FOTKY NA IMGBB --
                        with st.spinner("Nahrávám novou fotku na ImgBB..."):
                            novy_obrazek_bytes = fotka_edit_file.getvalue()
                            nove_imgbb_url = nahraj_na_imgbb(novy_obrazek_bytes)
                            if nove_imgbb_url:
                                cilova_fotka_url = nove_imgbb_url
    
                idx = df[df["ID"] == vybrane_id].index[0]
                df.at[idx, "Lokomotiva"] = formatuj_lokomotivu(loko_edit)
                df.at[idx, "Kategorie"] = kategorie_edit
                df.at[idx, "Datum"] = pd.to_datetime(datum_edit)
                df.at[idx, "Popis závady"] = popis_edit.strip()
                df.at[idx, "Poznámka"] = poznamka_edit.strip()
                df.at[idx, "Fotka"] = cilova_fotka_url.strip()
                # 🟢 POJISTKA PROTI CHYBĚ: Zajištění, že sloupec vezme text
                if "Upravil" in df.columns:
                    df["Upravil"] = df["Upravil"].astype(object)
                
                # Přesný zápis do jednoho řádku (strana obalená do str())
                aktualni_uzivatel = st.session_state.get("uzivatel_jmeno", "Neznámý")
                df.at[idx, "Upravil"] = str(aktualni_uzivatel) 
                
                ok, err = ulozit_databazi(df, f"Úprava závady ID {vybrane_id}")
                if ok:
                    st.session_state["msg_tab3"] = f"✅ Závada ID {vybrane_id} byla úspěšně aktualizována!"
                    st.rerun()
                else:
                    st.error(f"Chyba při ukládání: {err}")           

# TAB 4: Smazat
if tab_smazat:    
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

# --- TAB: Dokumentace (PDF) ---
with tab_pdf:
    st.title("📄 Technická dokumentace (PDF)")
    st.caption(
        "Ukládání a prohlížení PDF dokumentů v podsložkách podle řad."
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
    st.subheader("➕ Nahrát nový dokument (PDF)")

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
        "Vyberte PDF soubor:",
        type=["pdf"],
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
            valid_extensions = (".pdf")
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
    st.title("🖼️ Fotodokumentace a videa")
    st.caption("Ukládání a správa fotografií a krátkých videí v podsložkách podle řad na GitHubu.")

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
    # 1. NAHRÁVÁNÍ FOTOGRAFIÍ/VIDEÍ A SPRÁVA PODSLOŽEK
    # ---------------------------------------------------------
    st.subheader("➕ Nahrát nový soubor (fotku/video)")

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

    # PŘIDÁNY VIDEO FORMÁTY
    uploaded_file = st.file_uploader(
        "Vyberte fotografii nebo krátké video (max 10s):",
        type=["jpg", "jpeg", "png", "webp", "mp4", "mov", "avi"],
        key="github_img_uploader",
    )

    if uploaded_file is not None:
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
                f"docs_foto/{zvolena_rada_foto}/{safe_sub}/{uploaded_file.name}"
            )
            file_bytes_foto = uploaded_file.getvalue()

            if st.button("🚀 Uložit soubor na GitHub", key="btn_save_foto"):
                with st.spinner("Ukládám soubor do GitHub repozitáře (u videa to může chvilku trvat)..."):
                    try:
                        try:
                            contents = repo.get_contents(file_path_foto)
                            repo.update_file(
                                path=file_path_foto,
                                message=f"Aktualizace souboru {uploaded_file.name} v {safe_sub}",
                                content=file_bytes_foto,
                                sha=contents.sha,
                            )
                            st.success(
                                f"✅ Soubor '{uploaded_file.name}' byl aktualizován!"
                            )
                        except GithubException:
                            repo.create_file(
                                path=file_path_foto,
                                message=f"Přidán soubor {uploaded_file.name} do {safe_sub}",
                                content=file_bytes_foto,
                            )
                            st.success(
                                f"✅ Soubor '{uploaded_file.name}' byl úspěšně uložen do složky **{safe_sub}**!"
                            )

                        st.rerun()
                    except Exception as ex:
                        st.error(f"Při ukládání došlo k chybě: {ex}")

    st.divider()

   # ---------------------------------------------------------
   # 2. PROHLÍŽEČ A MAZÁNÍ FOTOGRAFIÍ/VIDEÍ
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
            # ROZŠÍŘENO O VIDEO FORMÁTY
            media_extensions = (".jpg", ".jpeg", ".png", ".webp", ".mp4", ".mov", ".avi")
            media_files = [
                f
                for f in folder_imgs
                if f.name.lower().endswith(media_extensions)
            ]
        except GithubException:
            media_files = []

        if media_files:
            st.write(
                f"Nalezeno **{len(media_files)}** souborů ve složce **{vybrana_sub_view}**:"
            )

            # Zobrazení ve mřížce (2 sloupce vedle sebe)
            cols = st.columns(2)
            headers = {"Authorization": f"token {github_token}"}

            for idx, file_obj in enumerate(media_files):
                col = cols[idx % 2]

                res = requests.get(file_obj.download_url, headers=headers)
                if res.status_code == 200:
                    view_url_file = f"https://cdn.jsdelivr.net/gh/{repo_name}@main/{file_obj.path}"

                    with col:
                        # ROZHODNUTÍ, ZDA ZOBRAZIT FOTKU NEBO VIDEO
                        is_video = file_obj.name.lower().endswith((".mp4", ".mov", ".avi"))
                        
                        if is_video:
                            st.video(res.content)
                            st.caption(file_obj.name) # Přidáme popisek pod video
                        else:
                            st.image(
                                res.content,
                                caption=file_obj.name,
                                use_container_width=True,
                            )

                        c1, c2, c3 = st.columns([1, 1, 1])
                        with c1:
                            st.link_button(
                                "🔗 Otevřít",
                                url=view_url_file,
                                use_container_width=True,
                            )
                        with c2:
                            # Správný typ souboru pro stahování
                            mime_type = "video/mp4" if is_video else "image/jpeg"
                            st.download_button(
                                label="💾 Stáhnout",
                                data=res.content,
                                file_name=file_obj.name,
                                mime=mime_type,
                                key=f"dl_{file_obj.sha}",
                                use_container_width=True,
                            )
                        with c3:
                            # Tlačítko pro smazání
                            if st.button(
                                "🗑️ Smazat",
                                key=f"del_{file_obj.sha}",
                                type="secondary",
                                use_container_width=True,
                            ):
                                try:
                                    repo.delete_file(
                                        path=file_obj.path,
                                        message=f"Smazán soubor {file_obj.name} ze složky {vybrana_sub_view}",
                                        sha=file_obj.sha,
                                    )
                                    st.success(
                                        f"Soubor '{file_obj.name}' byl smazán."
                                    )
                                    st.rerun()
                                except Exception as del_err:
                                    st.error(
                                        f"Chyba při mazání souboru: {del_err}"
                                    )

                    st.write("---")
        else:
            st.info(
                f"Ve složce **{vybrana_sub_view}** zatím nejsou žádné fotky ani videa."
            )

# Tab Chat
if tab_chat:
    with tab_chat:
        st.header("💬 Chat")
        st.caption(
            "Nástěnka pro rychlou komunikaci mezi všemi přihlášenými uživateli."
        )

        # 1. Nejprve načteme zprávy z GitHubu (pokud ještě v session_state nejsou)
        if "chat_zpravy" not in st.session_state:
            st.session_state["chat_zpravy"] = nacti_chat_z_githubu()

        # 2. Až po načtení označíme zprávy jako přečtené
        zpravy = st.session_state["chat_zpravy"]
        st.session_state["chat_precteno_pocet"] = len(zpravy)

        # Tlačítko pro ruční aktualizaci správ
        if st.button("🔄 Obnovit zprávy", key="refresh_chat"):
            st.session_state["chat_zpravy"] = nacti_chat_z_githubu()
            st.rerun()

        st.markdown("---")

        # Kontejner se skrolováním pro zprávy
        chat_container = st.container(height=450)

        with chat_container:
            if not zpravy:
                st.info(
                    "Zatím zde nejsou žádné zprávy. Napište první vzkaz níže!"
                )
            else:
                aktualni_prihlaseny = st.session_state.get(
                    "uzivatel_jmeno", "Neznámý"
                )

                for msg in zpravy:
                    s_uzivatel = msg.get("uzivatel", "Neznámý")
                    s_cas = msg.get("cas", "")
                    s_text = msg.get("zprava", "")

                    # Rozlišení vlastních zpráv od ostatních uživatelů
                    if s_uzivatel == aktualni_prihlaseny:
                        with st.chat_message("user", avatar="👷‍♂️"):
                            st.write(f"**Vy** ({s_cas}):")
                            st.write(s_text)
                    else:
                        with st.chat_message("assistant", avatar="🛠️"):
                            st.write(f"**{s_uzivatel}** ({s_cas}):")
                            st.write(s_text)

        # Vstupní pole pro novou zprávu dole na stránce
        novy_text = st.chat_input("Napište vzkaz kolegovi...")

        if novy_text:
            aktualni_uzivatel = st.session_state.get("uzivatel_jmeno", "Neznámý")
            cas_zpravy = datetime.now(ZoneInfo("Europe/Prague")).strftime("%d.%m. %H:%M")

            nova_zprava = {
                "uzivatel": aktualni_uzivatel,
                "cas": cas_zpravy,
                "zprava": novy_text.strip(),
            }

            # Přidání zprávy do lokálního stavu a okamžité uložení na GitHub
            st.session_state["chat_zpravy"].append(nova_zprava)

            with st.spinner("Odesílám zprávu..."):
                uloz_chat_na_github(
                    st.session_state["chat_zpravy"], autor=aktualni_uzivatel
                )

            st.rerun()
