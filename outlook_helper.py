"""
Outlook integrace - vytvoří e-mail s PDF přílohou jako koncept (Draft).

Funguje pouze na Windows s nainstalovaným Outlook desktop. Vrací:
- True, "OK" - draft byl vytvořen
- False, "důvod" - na non-Windows, bez Outlooku, atp.

Bezpečnost: e-mail se VŽDY vytvoří jen jako koncept (Drafts).
Uživatel ho musí sám otevřít a kliknout Odeslat. Nic se neodešle automaticky.
"""

import os
import platform
import tempfile


def vytvor_email_koncept(
    pdf_bytes: bytes,
    filename: str,
    prijemce: str,
    jmeno_zamestnance: str,
    pocet_kusu: int,
    zobrazit_okno: bool = True,
) -> tuple[bool, str]:
    """Vytvoří draft v Outlooku s PDF přílohou.

    Args:
        pdf_bytes:    obsah PDF souboru
        filename:     název přílohy (např. "vizitka_Cervik.pdf")
        prijemce:     e-mail tiskárny (To)
        jmeno_zamestnance: pro Subject a Body
        pocet_kusu:   počet kusů k tisku
        zobrazit_okno: True = otevřít okno emailu (Display);
                       False = jen vytvořit draft ve složce Drafts (Save)

    Returns: (success, message)
    """
    if platform.system() != "Windows":
        return False, "Outlook integrace je k dispozici pouze na Windows."

    try:
        import win32com.client  # type: ignore
    except ImportError:
        return False, ("Není nainstalován balíček pywin32. "
                       "Spusť: pip install pywin32")

    # Uložit PDF dočasně - Outlook potřebuje cestu k souboru pro přílohu
    tmp_dir = tempfile.gettempdir()
    pdf_path = os.path.join(tmp_dir, filename)
    with open(pdf_path, "wb") as f:
        f.write(pdf_bytes)

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = prijemce
        mail.Subject = f"Objednávka tisku vizitek – {jmeno_zamestnance}"
        mail.Body = (
            f"Dobrý den,\n\n"
            f"prosím o vytištění vizitek pro zaměstnance:\n\n"
            f"  Jméno:       {jmeno_zamestnance}\n"
            f"  Počet kusů:  {pocet_kusu}\n\n"
            f"Vizitka je v příloze ve formátu PDF, připravená k tisku "
            f"(včetně ořezových značek a 3 mm bleedu).\n\n"
            f"Děkuji.\n\n"
            f"--\n"
            f"Tento e-mail byl automaticky vygenerován ze systému VEKRA."
        )
        mail.Attachments.Add(pdf_path)

        if zobrazit_okno:
            mail.Display(False)  # False = nemodální okno
            msg = "E-mail otevřen v novém okně Outlooku. Zkontrolujte a odešlete."
        else:
            mail.Save()
            msg = "E-mail uložen do složky Koncepty (Drafts) v Outlooku."

        return True, msg

    except Exception as e:
        return False, f"Chyba při komunikaci s Outlookem: {e}"
