# VEKRA · Vizitky · Prototyp

Lokální webová aplikace pro automatickou výrobu vizitek s QR kódem.
Slouží jako **demo automatizace** před nasazením do Azure / Power Automate.

## Co aplikace dělá

1. Uživatel vyplní formulář (nebo nahraje CSV).
2. Aplikace vygeneruje **PDF vizitku** v layoutu VEKRA (logo, červený pruh, ořezové značky, bleedy).
3. Do PDF se vloží **QR kód s vCard** – po naskenování telefonem se kontakt přidá do telefonu.
4. Aplikace připraví **e-mail tiskárně** (otevře koncept v Outlooku s přílohou + počtem kusů).

## Vstupní data: CSV

Prototyp pracuje s **CSV soubory** (žádný Office není potřeba – CSV otevřeš a upravíš
i v Poznámkovém bloku).

Formát CSV: oddělovač středník (`;`), první řádek = názvy sloupců. V balíčku je
ukázkový soubor **`ukazka_vizitky.csv`**, který můžeš rovnou nahrát na záložce
*Hromadné zpracování*.

**V produkci** (přes Power Automate) se data berou přímo ze SharePoint Excelu
automaticky – ruční práce s CSV pak odpadá.

## Instalace (jednorázová, cca 5 minut)

### 1. Nainstaluj Python 3.11

Stáhni z **https://www.python.org/downloads/** – verzi 3.11.x.

⚠️ **Při instalaci zaškrtni *„Add Python to PATH"*.**

### 2. Spusť `instalace.bat`

Dvojklikem na soubor. Skript:
- Vytvoří virtuální prostředí (složka `venv/`)
- Stáhne závislosti (Streamlit, reportlab, pywin32)

Trvá cca 2 minuty. Po dokončení uvidíš zprávu *„Instalace dokončena"*.

## Spuštění aplikace

Dvojklik na **`spustit_demo.bat`**.

- Otevře se okno s konzolovým výstupem (nech ho otevřené, dokud potřebuješ aplikaci).
- Po pár sekundách se v prohlížeči otevře **http://localhost:8501**.
- Pro ukončení stiskni v konzoli **Ctrl+C** nebo zavři okno.

## Jak prezentovat na demu

**Scénář 1 – jednotlivá vizitka (30 sekund):**
1. Vlevo vyplň údaje (nebo nech přednastavené).
2. Klikni *„Generovat vizitku"* → vpravo se objeví **náhled PDF**.
3. Rozbal *„QR kód samostatně"* – pozvi diváka, aby naskenoval QR fotoaparátem telefonu.
   Telefonu naskočí dialog **„Přidat do kontaktů"** s předvyplněnými údaji.
4. Klikni *„Připravit email tiskárně"* → otevře se Outlook s předvyplněným e-mailem,
   přílohou (PDF) a počtem kusů v textu.

**Scénář 2 – automatický režim (nejlepší ukázka automatizace):**
1. Přejdi na záložku *„🔴 Automatický režim"*.
2. Nech zaškrtnutý *„vestavěný demo soubor"* a klikni *„Spustit sledování"*.
3. Aplikace teď **hlídá soubor**. Zobrazí se zelený pruh *„Sleduji soubor…"*.
4. Rozbal *„Simulovat nový kontakt"*, vyplň údaje (nebo nech přednastavené) a klikni
   *„Odeslat objednávku"*.
5. **Do 2 sekund se sama od sebe objeví hotová vizitka** – bez dalšího klikání.
   Tohle je přesně to, co v produkci udělá Power Automate po vyplnění Forms.
6. Můžeš přidat víc kontaktů za sebou – každý se zpracuje automaticky.

**Scénář 3 – hromadné zpracování:**
1. Přepni na záložku *„Hromadné zpracování"*.
2. Nahraj CSV (např. přiložený `ukazka_vizitky.csv`).
3. Klikni *„Vygenerovat všechny vizitky"*.
4. Stáhne se ZIP s PDF pro všechny zaměstnance najednou.

## Co jde do produkce a co zůstává

Tento prototyp ti slouží k demonstraci **konceptu** a layoutu vizitky.
V produkci se nahradí:

| Prototyp | Produkce |
|---|---|
| Streamlit formulář | Microsoft Forms |
| Streamlit web | Power Automate flow |
| Lokální Python | Azure Function (HTTP endpoint) |
| Outlook desktop | Power Automate „Send email" akce |

**Vstup a výstup zůstávají identické** – stejný layout, stejný QR kód.

## Známé limity prototypu

- **Outlook integrace** funguje jen na Windows s nainstalovaným Outlook desktop.
- **Font Geograph** – aktuálně náhrada (Carlito). Pro tisk je třeba dodat originál
  (umístit `Geograph-Regular.ttf` a `Geograph-Bold.ttf` do složky `fonts/` a
  upravit `generator.py` řádky 18-19).
- **CMYK barvy** – PDF je v RGB. Tiskárny si obvykle převedou samy, případně lze
  doplnit konverzi přes Ghostscript.

## Struktura projektu

```
prototyp_demo/
├── app.py                  # Streamlit aplikace
├── generator.py            # Generátor PDF + QR
├── outlook_helper.py       # Integrace s Outlook desktop
├── vekra_logo.png          # Logo
├── fonts/                  # TTF fonty
│   ├── Carlito-Regular.ttf
│   └── Carlito-Bold.ttf
├── requirements.txt        # Python závislosti
├── instalace.bat           # Jednorázová instalace
├── spustit_demo.bat        # Spuštění aplikace
└── README.md               # Tento soubor
```
