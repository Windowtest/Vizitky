"""
VEKRA – Generátor vizitek (lokální prototyp)

Spuštění:
    streamlit run app.py

Aplikace běží lokálně v prohlížeči na http://localhost:8501
"""

import base64
import csv
import io
import re
import zipfile
from typing import Optional

import streamlit as st

from generator import (build_vcard, generate_business_card_bytes,
                        render_qr_png)


def _hledat(row: dict, klice: list) -> str:
    """Najde hodnotu ve slovníku podle částečné shody klíče.

    Normalizuje názvy sloupců - odstraní VŠECHNY non-alfanumerické znaky
    (včetně otazníků z rozbité diakritiky a mezer). Klíče se zadávají
    bez diakritiky a bez mezer.

    Příklad: 'Jméno P?íjmení (popř. titul)' -> 'jmenopijmenipoptitul'
    """
    import unicodedata

    def normalize(s: str) -> str:
        if not s:
            return ""
        nfd = unicodedata.normalize("NFD", str(s).lower())
        no_diacritics = "".join(c for c in nfd if not unicodedata.combining(c))
        return re.sub(r"[^a-z0-9]", "", no_diacritics)

    for full_key, val in row.items():
        if not full_key:
            continue
        normalized = normalize(full_key)
        for k in klice:
            if k in normalized:
                return str(val or "").strip()
    return ""


def cti_csv(file_bytes: bytes) -> list[dict]:
    """Načte CSV. Zkusí UTF-8 i CP1250 a oba běžné oddělovače."""
    for encoding in ("utf-8-sig", "cp1250"):
        try:
            text = file_bytes.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = file_bytes.decode("utf-8", errors="replace")

    for delim in (";", ","):
        reader = csv.DictReader(io.StringIO(text), delimiter=delim)
        rows = [r for r in reader if any((v or "").strip() for v in r.values())]
        if len(rows) > 0 and len(reader.fieldnames or []) > 2:
            return rows
    return rows


# === SLEDOVÁNÍ SOUBORU (automatický režim) =========================
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SLEDOVANE_DIR = os.path.join(BASE_DIR, "sledovane")
DEMO_CSV = os.path.join(SLEDOVANE_DIR, "objednavky_demo.csv")
DEMO_HLAVICKA = ["Jmeno Prijmeni", "Pozice", "Telefon",
                 "E-mail1", "Adresa", "Pocet kusu"]


def zajisti_demo_csv():
    """Vytvoří prázdný demo soubor (jen hlavička), pokud neexistuje."""
    os.makedirs(SLEDOVANE_DIR, exist_ok=True)
    if not os.path.exists(DEMO_CSV):
        with open(DEMO_CSV, "w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f, delimiter=";").writerow(DEMO_HLAVICKA)


def nacti_radky(path: str) -> list[dict]:
    """Načte všechny datové řádky z CSV souboru."""
    if not path or not os.path.exists(path):
        return []
    try:
        with open(path, "rb") as f:
            return cti_csv(f.read())
    except Exception:
        return []


def gsheet_to_csv_url(url: str) -> str | None:
    """Převede sdílený odkaz Google Sheets na URL pro CSV export."""
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    if not m:
        return None
    sheet_id = m.group(1)
    gid_m = re.search(r"[#&?]gid=([0-9]+)", url)
    gid = gid_m.group(1) if gid_m else "0"
    return (f"https://docs.google.com/spreadsheets/d/{sheet_id}"
            f"/export?format=csv&gid={gid}")


def nacti_z_gsheet(url: str) -> tuple[list[dict], str]:
    """Stáhne data z Google Sheets. Vrací (rows, chybova_zprava)."""
    csv_url = gsheet_to_csv_url(url)
    if not csv_url:
        return [], ("Neplatný odkaz. Vlož celý odkaz na Google tabulku "
                    "(např. https://docs.google.com/spreadsheets/d/...).")
    try:
        import requests
        r = requests.get(csv_url, timeout=15)
        if r.status_code == 200 and "text/html" in r.headers.get("content-type", ""):
            return [], ("Tabulka není veřejně přístupná. Nastav sdílení na: "
                        "Kdokoli s odkazem → Čtenář.")
        if r.status_code != 200:
            return [], (f"Google vrátil chybu {r.status_code}. "
                        f"Zkontroluj, že je tabulka sdílená.")
        return cti_csv(r.content), ""
    except Exception as e:
        return [], f"Chyba při stahování z Google Sheets: {e}"


def pridej_radek_csv(path: str, radek: list):
    """Přidá řádek do CSV (pro simulaci nového kontaktu)."""
    with open(path, "a", encoding="utf-8-sig", newline="") as f:
        csv.writer(f, delimiter=";").writerow(radek)


def radek_na_data(row: dict) -> dict:
    """Vytáhne z řádku strukturu pro generátor vizitky."""
    adresa = _hledat(row, ["adresa", "address"]).replace(" | ", "\n")
    return {
        "jmeno": _hledat(row, ["prijmeni", "pijmeni", "jmeno", "name"]),
        "pozice": _hledat(row, ["pozice", "position"]),
        "telefon": _hledat(row, ["telefon", "phone"]),
        "email": _hledat(row, ["email1", "email"]),
        "adresa": adresa,
    }


def je_kompletni(data: dict) -> bool:
    """Řádek je kompletní, jen když má vyplněné všechny údaje pro vizitku.
    Brání tomu, aby se vizitka vygenerovala z rozepsaného řádku."""
    povinne = ["jmeno", "pozice", "telefon", "email", "adresa"]
    return all((data.get(k) or "").strip() for k in povinne)


# Outlook integrace je volitelná - jen na Windows s nainstalovaným Outlookem.
# V cloudu (Linux) se vypne a místo ní se ukáže náhled e-mailu.
import platform as _platform
try:
    if _platform.system() == "Windows":
        from outlook_helper import vytvor_email_koncept
        OUTLOOK_DOSTUPNY = True
    else:
        OUTLOOK_DOSTUPNY = False
except ImportError:
    OUTLOOK_DOSTUPNY = False


# === KONFIGURACE STRÁNKY ============================================
st.set_page_config(
    page_title="VEKRA · Vizitky",
    page_icon="📇",
    layout="wide",
)

# Brand styling (firemní červená VEKRA)
st.markdown("""
<style>
    .block-container { padding-top: 2rem; }
    h1 { color: #1A1A1A !important; }
    .stButton button[kind="primary"] {
        background-color: #E30613 !important;
        border-color: #E30613 !important;
    }
    .stButton button[kind="primary"]:hover {
        background-color: #B8050F !important;
        border-color: #B8050F !important;
    }
    /* Karty pro výsledky */
    .vysledek-karta {
        padding: 1rem; border-radius: 8px;
        background: #F8F9FA; border-left: 4px solid #E30613;
        margin: 0.5rem 0;
    }
</style>
""", unsafe_allow_html=True)


# === SESSION STATE ==================================================
if "pdf_bytes" not in st.session_state:
    st.session_state.pdf_bytes = None
if "qr_png" not in st.session_state:
    st.session_state.qr_png = None
if "data" not in st.session_state:
    st.session_state.data = None
if "watching" not in st.session_state:
    st.session_state.watching = False
if "auto_cards" not in st.session_state:
    st.session_state.auto_cards = []
if "processed_rows" not in st.session_state:
    st.session_state.processed_rows = set()
if "watch_path" not in st.session_state:
    st.session_state.watch_path = ""
if "watch_source" not in st.session_state:
    st.session_state.watch_source = "demo"
if "watch_gsheet_url" not in st.session_state:
    st.session_state.watch_gsheet_url = ""
if "batch_rows" not in st.session_state:
    st.session_state.batch_rows = []
if "show_email_preview" not in st.session_state:
    st.session_state.show_email_preview = False


# === HLAVIČKA =======================================================
st.title("📇 Generátor vizitek VEKRA")
st.caption(
    "Prototyp automatizace · Vyplň údaje → vygeneruje se PDF s QR kódem "
    "(vCard) → odešle se do tiskárny e-mailem"
)


# === HLAVNÍ ZÁLOŽKY =================================================
tab_jedna, tab_batch, tab_auto, tab_napoveda = st.tabs(
    ["🎯 Jednotlivá vizitka", "📦 Hromadné zpracování",
     "🔴 Automatický režim", "📖 Jak to funguje"]
)


# -------- TAB 1: JEDNOTLIVÁ VIZITKA --------------------------------
with tab_jedna:
    col_vstup, col_nahled = st.columns([1, 1.2])

    with col_vstup:
        st.markdown("### 1️⃣ Údaje na vizitce")

        jmeno = st.text_input(
            "Jméno (s titulem)",
            value="Mgr. Jakub Červík",
            help="Tituly jako Mgr., Bc., Ing. zůstanou s tečkou; samotné jméno se převede na velká písmena.",
        )
        pozice = st.text_input("Pozice", value="Business Development Manager B2B")
        telefon = st.text_input("Telefon", value="+420 728 449 422")
        email = st.text_input("E-mail", value="jakub.cervik@vekra.cz")
        adresa = st.text_area(
            "Adresa obchodního zastoupení (max. 3 řádky)",
            value="Obchodní zastoupení\nČeské Vrbné 2393\n370 11 České Budějovice",
            height=88,
        )

        st.markdown("### 2️⃣ Objednávka tisku")
        c1, c2 = st.columns(2)
        with c1:
            pocet = st.number_input(
                "Počet kusů", min_value=50, max_value=2000, value=200, step=50
            )
        with c2:
            email_tiskarny = st.text_input(
                "E-mail tiskárny", value="objednavky@tiskarna.cz"
            )

        generovat = st.button(
            "🎨 Generovat vizitku", type="primary", use_container_width=True
        )

        if generovat:
            data = {
                "jmeno": jmeno.strip(),
                "pozice": pozice.strip(),
                "telefon": telefon.strip(),
                "email": email.strip(),
                "adresa": adresa.strip(),
            }
            if not all(data.values()):
                st.error("Vyplň prosím všechny povinné údaje.")
            else:
                with st.spinner("Generuji PDF a QR kód..."):
                    st.session_state.pdf_bytes = generate_business_card_bytes(data)
                    vcard = build_vcard(
                        data["jmeno"], data["pozice"], data["telefon"],
                        data["email"], data["adresa"],
                    )
                    st.session_state.qr_png = render_qr_png(vcard, 400)
                    st.session_state.data = data
                    st.session_state.pocet = int(pocet)
                    st.session_state.email_tiskarny = email_tiskarny

    with col_nahled:
        st.markdown("### 3️⃣ Náhled")

        if st.session_state.pdf_bytes:
            # PDF náhled v iframe
            b64 = base64.b64encode(st.session_state.pdf_bytes).decode()
            st.markdown(
                f'<iframe src="data:application/pdf;base64,{b64}#zoom=160" '
                f'width="100%" height="330" '
                f'style="border:1px solid #ddd; border-radius:6px; background:#fff;">'
                f'</iframe>',
                unsafe_allow_html=True,
            )

            # Akce
            st.markdown("### 4️⃣ Co s vizitkou")
            c_a, c_b = st.columns(2)
            with c_a:
                safe_name = re.sub(r"[^\w\s-]", "",
                                   st.session_state.data["jmeno"], flags=re.UNICODE)
                safe_name = re.sub(r"\s+", "_", safe_name.strip())
                st.download_button(
                    "💾 Stáhnout PDF",
                    data=st.session_state.pdf_bytes,
                    file_name=f"vizitka_{safe_name}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            with c_b:
                if OUTLOOK_DOSTUPNY:
                    if st.button(
                        "📧 Připravit email tiskárně",
                        type="primary",
                        use_container_width=True,
                    ):
                        with st.spinner("Otevírám Outlook..."):
                            ok, msg = vytvor_email_koncept(
                                pdf_bytes=st.session_state.pdf_bytes,
                                filename=f"vizitka_{safe_name}.pdf",
                                prijemce=st.session_state.email_tiskarny,
                                jmeno_zamestnance=st.session_state.data["jmeno"],
                                pocet_kusu=st.session_state.pocet,
                                zobrazit_okno=True,
                            )
                        if ok:
                            st.success(f"✅ {msg}")
                        else:
                            st.error(f"❌ {msg}")
                else:
                    # V cloudu Outlook neběží - ukážeme náhled e-mailu
                    if st.button(
                        "📧 Náhled e-mailu pro tiskárnu",
                        type="primary",
                        use_container_width=True,
                    ):
                        st.session_state.show_email_preview = True

            # Náhled e-mailu (cloud varianta) - mimo sloupce, přes celou šířku
            if not OUTLOOK_DOSTUPNY and st.session_state.get("show_email_preview"):
                st.markdown("##### 📧 Takto by vypadala objednávka tiskárně")
                st.info(
                    "V ostré verzi (na firemním PC nebo přes Power Automate) "
                    "se tento e-mail odešle automaticky s PDF v příloze."
                )
                st.text_input("Komu", st.session_state.email_tiskarny,
                              disabled=True, key="prev_to")
                st.text_input(
                    "Předmět",
                    f"Objednávka tisku vizitek – {st.session_state.data['jmeno']}",
                    disabled=True, key="prev_subj")
                st.text_area(
                    "Text e-mailu",
                    f"Dobrý den,\n\nprosím o vytištění vizitek pro zaměstnance:\n\n"
                    f"  Jméno:       {st.session_state.data['jmeno']}\n"
                    f"  Počet kusů:  {st.session_state.pocet}\n\n"
                    f"Vizitka je v příloze ve formátu PDF, připravená k tisku.\n\n"
                    f"Děkuji.",
                    height=200, disabled=True, key="prev_body")
                st.caption("📎 Příloha: PDF vizitka (stáhneš ji tlačítkem vlevo).")

            # QR kód samostatně (pro demo - lze ho naskenovat)
            with st.expander("📱 QR kód samostatně (pro skenování telefonem na demu)"):
                st.caption(
                    "Naskenuj fotoaparátem telefonu – kontakt se uloží do telefonu."
                )
                st.image(st.session_state.qr_png, width=260)

        else:
            st.info("👈 Vyplň údaje vlevo a klikni na **Generovat vizitku**.")


# -------- TAB 2: BATCH ---------------------------------------------
with tab_batch:
    st.markdown("### Hromadné zpracování z Google tabulky")
    st.caption(
        "Vlož odkaz na Google tabulku s objednávkami. Pro každý řádek se "
        "vygeneruje PDF; všechny si pak stáhneš najednou jako ZIP."
    )

    with st.expander("ℹ️ Jak připravit Google tabulku (jednou)"):
        st.markdown("""
1. Otevři **sheets.google.com** → *Prázdná tabulka*.
2. Do prvního řádku napiš názvy sloupců:
   `Jmeno Prijmeni` · `Pozice` · `Telefon` · `E-mail1` · `Adresa` · `Pocet kusu`
3. Pod ně vyplň kontakty (každý řádek = jedna vizitka).
   V adrese odděl řádky svislítkem `|` (např. `Obchodní zastoupení | Pražská 100 | 110 00 Praha`).
4. Vpravo nahoře **Sdílet** → *Obecný přístup* → **Kdokoli s odkazem → Čtenář** → *Kopírovat odkaz*.
5. Ten odkaz vlož sem dolů. Hotovo.
        """)

    gsheet_url = st.text_input(
        "Odkaz na Google tabulku",
        placeholder="https://docs.google.com/spreadsheets/d/.../edit",
        key="batch_gsheet_url",
    )

    nacist = st.button("📥 Načíst z Google tabulky", type="primary")

    if nacist and gsheet_url:
        with st.spinner("Stahuji data z Google Sheets..."):
            rows, chyba = nacti_z_gsheet(gsheet_url)
        if chyba:
            st.error(chyba)
        else:
            st.session_state.batch_rows = rows

    rows = st.session_state.get("batch_rows", [])
    if rows:
        st.success(f"✅ Načteno **{len(rows)} řádků** z Google tabulky.")

        # Zobrazení jako tabulka
        st.markdown("#### Náhled dat")
        st.dataframe(rows, use_container_width=True, height=240)

        if st.button("🚀 Vygenerovat všechny vizitky"):
            zip_buf = io.BytesIO()
            uspesne = 0
            chyby = []

            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                progress = st.progress(0, "Zpracovávám...")
                for i, row in enumerate(rows):
                    data = radek_na_data(row)
                    if not data["jmeno"]:
                        continue
                    try:
                        pdf = generate_business_card_bytes(data)
                        safe = re.sub(r"[^\w\s-]", "", data["jmeno"], flags=re.UNICODE)
                        safe = re.sub(r"\s+", "_", safe.strip())
                        zf.writestr(f"vizitka_{safe}.pdf", pdf)
                        uspesne += 1
                    except Exception as e:
                        chyby.append(f"{data['jmeno']}: {e}")
                    progress.progress((i + 1) / len(rows),
                                      f"Zpracováno {i+1}/{len(rows)}")

            st.success(f"✅ Vygenerováno {uspesne} vizitek")
            if chyby:
                st.warning(f"Chyby: {len(chyby)}")
                with st.expander("Detail chyb"):
                    for ch in chyby:
                        st.text(ch)

            st.download_button(
                "💾 Stáhnout ZIP se všemi vizitkami",
                data=zip_buf.getvalue(),
                file_name="vizitky.zip",
                mime="application/zip",
            )

    # Alternativa: nahrát CSV
    with st.expander("📄 Nebo nahraj CSV soubor (alternativa ke Google tabulce)"):
        soubor = st.file_uploader("Soubor CSV", type=["csv"])
        if soubor:
            rows_csv = cti_csv(soubor.read())
            if rows_csv:
                st.success(f"Načteno {len(rows_csv)} řádků z CSV.")
                st.session_state.batch_rows = rows_csv
                st.rerun()


# -------- TAB 3: AUTOMATICKÝ REŽIM ---------------------------------
with tab_auto:
    st.markdown("### 🔴 Automatický režim")
    st.caption(
        "Aplikace **sleduje soubor**. Jakmile přibude nový kontakt, "
        "**sama** vygeneruje vizitku – bez jediného kliknutí. Přesně tohle "
        "udělá v produkci Power Automate, když někdo vyplní Microsoft Forms."
    )

    zajisti_demo_csv()

    # Výběr zdroje dat
    zdroj = st.radio(
        "Co sledovat:",
        ["📊 Google tabulka (online)", "🧪 Demo soubor (lokální, s tlačítkem simulace)"],
        horizontal=True,
    )
    je_gsheet = zdroj.startswith("📊")

    watch_url = ""
    if je_gsheet:
        watch_url = st.text_input(
            "Odkaz na Google tabulku",
            placeholder="https://docs.google.com/spreadsheets/d/.../edit",
            value=st.session_state.get("watch_gsheet_url", ""),
            help="Tabulka musí být sdílená: Kdokoli s odkazem → Čtenář.",
        )
        st.caption(
            "💡 Při demu měj tabulku otevřenou ve vedlejší záložce prohlížeče. "
            "Přidáš řádek → během pár sekund se tu objeví vizitka. "
            "(Google data občas chvíli cachuje, drobné zpoždření je normální.)"
        )

    # Ovládání sledování
    sleduje = st.session_state.get("watching", False)
    c_start, c_stop = st.columns(2)
    with c_start:
        if st.button("▶️ Spustit sledování", type="primary",
                     disabled=sleduje, use_container_width=True):
            if je_gsheet:
                # Ověř, že tabulka jde načíst
                test_rows, chyba = nacti_z_gsheet(watch_url)
                if chyba:
                    st.error(chyba)
                else:
                    st.session_state.watching = True
                    st.session_state.watch_source = "gsheet"
                    st.session_state.watch_gsheet_url = watch_url
                    # Existující řádky označíme jako vyřízené - sledujeme jen nové
                    st.session_state.processed_rows = set(range(len(test_rows)))
                    st.session_state.auto_cards = []
                    st.rerun()
            else:
                st.session_state.watching = True
                st.session_state.watch_source = "demo"
                st.session_state.watch_path = DEMO_CSV
                st.session_state.processed_rows = set(range(len(nacti_radky(DEMO_CSV))))
                st.session_state.auto_cards = []
                st.rerun()
    with c_stop:
        if st.button("⏹️ Zastavit", disabled=not sleduje,
                     use_container_width=True):
            st.session_state.watching = False
            st.rerun()

    if st.session_state.get("watching"):
        st.divider()
        je_demo = st.session_state.get("watch_source") == "demo"

        # Simulace nového kontaktu (jen pro demo soubor)
        if je_demo:
            with st.expander("➕ Simulovat nový kontakt (jako vyplnění formuláře)",
                             expanded=True):
                with st.form("novy_kontakt", clear_on_submit=False):
                    fc1, fc2 = st.columns(2)
                    with fc1:
                        n_jmeno = st.text_input("Jméno", "Ing. Petra Nováková")
                        n_pozice = st.text_input("Pozice", "Obchodní zástupce")
                        n_tel = st.text_input("Telefon", "+420 777 123 456")
                    with fc2:
                        n_email = st.text_input("E-mail", "petra.novakova@vekra.cz")
                        n_adresa = st.text_area(
                            "Adresa", "Obchodní zastoupení\nPražská 100\n110 00 Praha",
                            height=80)
                        n_pocet = st.number_input("Počet kusů", 50, 2000, 200, 50)
                    if st.form_submit_button("📤 Odeslat objednávku",
                                             type="primary"):
                        pridej_radek_csv(
                            st.session_state.watch_path,
                            [n_jmeno, n_pozice, n_tel, n_email,
                             n_adresa.replace("\n", " | "), int(n_pocet)],
                        )
                        st.toast(f"✅ Nový kontakt {n_jmeno} přidán!")
        else:
            st.info(
                "Sleduji Google tabulku. Otevři ji v prohlížeči, přidej řádek "
                "s novým kontaktem – vizitka se vygeneruje sama."
            )

        # Auto-refresh fragment - srdce automatického režimu
        @st.fragment(run_every=3)
        def monitor():
            # Načti aktuální řádky podle zdroje
            if st.session_state.get("watch_source") == "gsheet":
                vsechny, _ = nacti_z_gsheet(st.session_state.get("watch_gsheet_url", ""))
            else:
                vsechny = nacti_radky(st.session_state.get("watch_path", ""))

            # Zpracuj jen NOVÉ a KOMPLETNÍ řádky.
            # processed_rows = indexy řádků, které už jsme vyřídili
            # (buď vygenerovali, nebo byly v tabulce už při startu sledování).
            processed = st.session_state.get("processed_rows", set())
            for idx, row in enumerate(vsechny):
                if idx in processed:
                    continue
                data = radek_na_data(row)
                # Dokud řádek není kompletní, počkáme (neoznačíme, nezpracujeme)
                if not je_kompletni(data):
                    continue
                try:
                    pdf = generate_business_card_bytes(data)
                    st.session_state.auto_cards.insert(0, {
                        "data": data, "pdf": pdf,
                    })
                    processed.add(idx)
                except Exception:
                    processed.add(idx)  # ať se chybný řádek nezkouší donekonečna
            st.session_state.processed_rows = processed

            # Status
            pocet = len(st.session_state.get("auto_cards", []))
            zdroj_txt = ("Google tabulku"
                         if st.session_state.get("watch_source") == "gsheet"
                         else "soubor")
            st.markdown(
                f'<div style="padding:0.6rem 1rem; background:#E8F5E9; '
                f'border-radius:6px; border-left:4px solid #2E7D32;">'
                f'🟢 <b>Sleduji {zdroj_txt}…</b> Vygenerováno {pocet} vizitek. '
                f'Vizitka vznikne, až je řádek úplně vyplněný.</div>',
                unsafe_allow_html=True,
            )

            # Vygenerované vizitky
            cards = st.session_state.get("auto_cards", [])
            if cards:
                st.markdown("#### ✨ Automaticky vygenerované vizitky")
                for idx, item in enumerate(cards):
                    data = item["data"]
                    pdf = item["pdf"]
                    with st.container(border=True):
                        st.markdown(
                            f"**Nová vizitka pro {data['jmeno']}**  "
                            f"<span style='color:#888'>· {data['pozice']} "
                            f"({data['email']})</span>",
                            unsafe_allow_html=True,
                        )
                        b64 = base64.b64encode(pdf).decode()
                        st.markdown(
                            f'<iframe src="data:application/pdf;base64,{b64}#zoom=120" '
                            f'width="100%" height="220" '
                            f'style="border:1px solid #eee; border-radius:4px;">'
                            f'</iframe>',
                            unsafe_allow_html=True,
                        )
                        safe = re.sub(r"[^\w\s-]", "", data["jmeno"], flags=re.UNICODE)
                        safe = re.sub(r"\s+", "_", safe.strip())
                        st.download_button(
                            "💾 Stáhnout PDF", pdf,
                            file_name=f"vizitka_{safe}.pdf",
                            mime="application/pdf",
                            key=f"auto_dl_{idx}",
                        )

        monitor()


# -------- TAB 4: NÁPOVĚDA -------------------------------------------
with tab_napoveda:
    st.markdown("""
    ### 🎯 K čemu tento prototyp slouží

    Demonstruje **automatizaci výroby vizitek**: ze vstupních dat (formulář nebo CSV)
    se automaticky vyrobí tiskové PDF s firemním layoutem VEKRA a QR kódem (vCard).

    ### 📋 Co se děje na pozadí

    1. **Vstup** – uživatel vyplní jméno, pozici, telefon, e-mail, adresu a počet kusů.
    2. **Generování PDF** – knihovna `reportlab` vyrobí PDF přesně podle firemního layoutu
       (logo, červený pruh, ořezové značky, bleedy 3 mm).
    3. **QR kód** – vygeneruje se obsah typu **vCard 3.0** se všemi kontaktními údaji.
       Po naskenování telefonem se otevře dialog „Přidat do kontaktů".
    4. **Tisková objednávka** – v Outlooku se připraví koncept e-mailu pro tiskárnu
       s PDF přílohou a počtem kusů v textu. Uživatel ho zkontroluje a sám odešle.

    ### 🌐 Jak to půjde dál do produkce

    Tento prototyp běží **lokálně na jednom PC**. Produkční varianta nahradí:

    - **Streamlit formulář** → **Microsoft Forms** (vstup od uživatelů)
    - **Lokální Python skript** → **Azure Function** (HTTP endpoint)
    - **Outlook desktop integraci** → **Power Automate flow** (automatické odeslání)

    Vstup a výstup zůstávají identické – stejný layout, stejný QR kód, stejný e-mail.

    ### 🛠️ Co je potřeba pro skutečné nasazení

    | Komponenta | Status |
    |---|---|
    | Layout vizitky | ✅ Hotovo (odpovídá originálu) |
    | QR kód s vCard | ✅ Funguje |
    | Generování PDF | ✅ Funguje |
    | Font Geograph | ⚠️ Aktuálně náhrada (Carlito) – pro tisk je třeba dodat originál |
    | Outlook integrace | ✅ Funguje (Windows + Outlook desktop) |
    | E-mailová adresa tiskárny | ⚠️ Doplnit reálnou |
    """)
