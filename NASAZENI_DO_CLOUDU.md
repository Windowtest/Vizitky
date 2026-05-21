# Návod: nasazení prototypu do cloudu (Streamlit Community Cloud)

Po nasazení poběží aplikace na internetu a otevřeš ji na **jakémkoliv PC v prohlížeči**
– bez instalace, bez Pythonu, bez práv administrátora. Ideální pro firemní počítač,
kam nemůžeš nic instalovat.

Budeš potřebovat dva bezplatné účty: **GitHub** a **Streamlit Cloud**.
Celé to zabere zhruba 20 minut a děláš to jen jednou.

---

## Krok 1: Založ GitHub účet

1. Jdi na **github.com** → *Sign up*.
2. Zadej e-mail, heslo, uživatelské jméno. Ověř e-mail.
   (Klidně použij soukromý e-mail, není potřeba firemní.)

## Krok 2: Vytvoř repozitář a nahraj soubory

1. Po přihlášení klikni vpravo nahoře na **+** → **New repository**.
2. **Repository name:** např. `vekra-vizitky`.
3. Zvol **Public** (Streamlit Cloud zdarma vyžaduje veřejný repozitář;
   jsou to jen testovací data a kód, žádné citlivé údaje).
4. Klikni **Create repository**.
5. Na další stránce klikni na odkaz **uploading an existing file**
   (uprostřed textu „…or push an existing repository…" je i tlačítko *upload*).
6. **Přetáhni do okna všechny soubory ze složky `prototyp_demo`** KROMĚ:
   - složky `venv` (pokud ji máš)
   - souborů `instalace.bat` a `spustit_demo.bat` (ty jsou jen pro lokální běh)

   Konkrétně nahraj: `app.py`, `generator.py`, `outlook_helper.py`,
   `requirements.txt`, `runtime.txt`, `vekra_logo.png`, složku `fonts`,
   a klidně i `README.md` a návody.
7. Dole klikni **Commit changes**.

> 💡 Tip: složku `fonts` přetáhni celou – GitHub zachová strukturu.

## Krok 3: Nasaď na Streamlit Cloud

1. Jdi na **share.streamlit.io**.
2. Klikni **Sign in** → **Continue with GitHub** (přihlásíš se přes GitHub účet).
3. Klikni **Create app** → **Deploy a public app from GitHub**.
4. Vyplň:
   - **Repository:** `tvoje-jmeno/vekra-vizitky`
   - **Branch:** `main`
   - **Main file path:** `app.py`
5. Klikni **Deploy**.
6. Počkej 2–5 minut (instaluje balíčky). Pak se aplikace otevře a dostaneš
   **veřejný odkaz** typu `https://vekra-vizitky.streamlit.app`.

## Krok 4: Otevři na firemním PC

Zkopíruj si ten odkaz a otevři ho na firemním PC v prohlížeči. Hotovo –
žádná instalace, funguje to.

---

## Co v cloudu funguje a co ne

| Funkce | Cloud |
|---|---|
| Generování PDF vizitky | ✅ |
| QR kód s vCard | ✅ |
| Google tabulka (hromadné i automatický režim) | ✅ |
| Stažení PDF / ZIP | ✅ |
| Náhled e-mailu pro tiskárnu | ✅ (jako náhled) |
| Automatické otevření Outlooku | ❌ (jen lokálně na Windows) |

Outlook v cloudu neběží – místo tlačítka „Připravit email" se zobrazí **náhled
e-mailu**, který by se v ostré verzi odeslal. Pro demo to stačí: ukážeš PDF,
QR kód i podobu objednávky tiskárně.

## Aktualizace aplikace

Když budu posílat novou verzi `app.py`, nahraješ ji na GitHub (přepíšeš starý soubor)
a Streamlit Cloud aplikaci **sám automaticky aktualizuje** za pár minut. Nic dalšího neděláš.

## Poznámka k bezpečnosti

Odkaz je veřejný – kdokoliv, kdo ho zná, aplikaci otevře. Pro testovací data to nevadí.
Kdybys později chtěl heslo nebo aby vše zůstalo jen ve firmě, řekni – buď přidáme
přihlášení, nebo přejdeme na firemní Azure (to už ale chce IT).
