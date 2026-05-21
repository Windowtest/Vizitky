# Návod: napojení na Google tabulku

Prototyp umí číst data přímo z Google tabulky (Google Sheets). Je to zdarma,
vypadá jako Excel a funguje v prohlížeči bez Office.

## 1. Vytvoř tabulku

1. Otevři **sheets.google.com** (přihlas se Google účtem) → *Prázdná tabulka*.
2. Do **prvního řádku** napiš názvy sloupců (každý do své buňky):

   | Jmeno Prijmeni | Pozice | Telefon | E-mail1 | Adresa | Pocet kusu |
   |---|---|---|---|---|---|

3. Pod ně vyplň kontakty – každý řádek = jedna vizitka.
   - V **Adrese** odděl jednotlivé řádky svislítkem `|`
     (např. `Obchodní zastoupení | Pražská 100 | 110 00 Praha`).

   **Tip:** můžeš naimportovat přiložený `sablona_pro_google_sheets.csv`
   (v Google Sheets: *Soubor → Importovat → Nahrát*), máš rovnou hlavičku i ukázky.

## 2. Nasdílej tabulku (důležité!)

Aby ji aplikace mohla číst, musí být veřejně čitelná:

1. Vpravo nahoře klikni **Sdílet**.
2. V sekci *Obecný přístup* přepni z „Omezený" na **Kdokoli s odkazem**.
3. Role nech na **Čtenář**.
4. Klikni **Kopírovat odkaz** → *Hotovo*.

> Bez tohoto kroku aplikace dostane jen přihlašovací stránku a nahlásí chybu.
> Tabulka je čitelná jen pro toho, kdo má přesný (dlouhý, neuhodnutelný) odkaz.

## 3. Vlož odkaz do aplikace

- **Hromadné zpracování:** vlož odkaz, klikni *Načíst z Google tabulky* → zobrazí se
  data jako tabulka → *Vygenerovat všechny vizitky* → ZIP.
- **Automatický režim:** vyber *Google tabulka*, vlož odkaz, *Spustit sledování*.
  Pak stačí v tabulce přidat řádek a do pár sekund se v aplikaci objeví hotová vizitka.

## Poznámka k rychlosti

Google data po úpravě chvíli cachuje, takže nový řádek se v automatickém režimu
nemusí objevit okamžitě – pár sekund až pár desítek sekund zpoždění je normální.
Na výsledek to nemá vliv, jen na rychlost reakce při demu.

## Jak to bude v produkci

Tady čteš Google tabulku ručně přes odkaz. V produkční variantě (Power Automate)
se to napojí na firemní SharePoint a spustí se automaticky po vyplnění formuláře –
žádné kopírování odkazů. Princip (tabulka → vizitka → e-mail) zůstává stejný.
