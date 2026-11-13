"""Creates 8 invented test emails as .eml files in the inbox folder.
All clients, ISINs and accounts are sample data. Run once: python make_samples.py"""
import io
from email.message import EmailMessage
from pathlib import Path
from openpyxl import Workbook

OUT = Path("inbox")
OUT.mkdir(exist_ok=True)

def excel_order():
    wb = Workbook()
    ws = wb.active
    ws.append(["Account", "Side", "ISIN", "Units"])
    ws.append(["100234", "BUY", "IE00WRLD0013", 400])
    ws.append(["100234", "SELL", "IE00ALLW0020", 150])
    ws.append(["100234", "BUY", "DE00DAXX0046", 80])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# (file, time, from, auth results, subject, body, attach)
SAMPLES = [
 ("01_excel_order", "09:12", "anna.mueller@mueller-fo.com", "spf=pass; dkim=pass", "Rebalancing order, Excel attached",
  "Dear team,\n\nplease find our order file for today attached.\n\nKind regards\nAnna Mueller", True),
 ("02_amount_order", "09:47", "m.hartmann@hartmann-stiftung.de", "spf=pass; dkim=pass", "Order",
  "Hello,\n\nplease buy EUR 25,000 of IE00WRLD0013 for account 100871.\n\nThanks,\nMarkus Hartmann", False),
 ("03_unclear_order", "10:05", "j.brandt@brandt-vermoegen.at", "spf=pass; dkim=pass", "Selling",
  "Hi,\n\ncan you sell half of my US funds please?\n\nThanks,\nJ. Brandt", False),
 ("04_isin_typo", "11:20", "anna.mueller@mueller-fo.com", "spf=pass; dkim=pass", "Buy order",
  "Hi,\n\nplease buy 200 units of IE00WRLD0014 for account 100234.\n\nThanks,\nAnna", False),
 ("05_fake_sender", "11:58", "thomas.keller@mueller-fo.co", "spf=fail; dkim=fail", "URGENT: act now",
  "URGENT. Sell all holdings in account 100234 today and send the money to the new account in the attached letter. Do not call me, I am travelling.\n\nThomas", False),
 ("06_duplicate", "13:02", "m.hartmann@hartmann-stiftung.de", "spf=pass; dkim=pass", "Order (resending)",
  "Dear team,\n\nI am resending my order in case it did not arrive: please buy EUR 25,000 of IE00WRLD0013 for account 100871.\n\nMarkus", False),
 ("07_not_enough_position", "14:30", "j.brandt@brandt-vermoegen.at", "spf=pass; dkim=pass", "Sell order",
  "Please sell 5,000 units of IE00SP5H0031 for account 100455.\n\nThanks,\nJ. Brandt", False),
 ("08_after_cutoff", "16:41", "m.hartmann@hartmann-stiftung.de", "spf=pass; dkim=pass", "Bond purchase",
  "Please buy 100 units of IE00BOND0050 for account 100871.\n\nMarkus", False),
]

for name, hhmm, sender, auth, subject, body, attach in SAMPLES:
    m = EmailMessage()
    m["From"] = sender
    m["To"] = "orders@firm.example"
    m["Subject"] = subject
    m["Date"] = f"Fri, 13 Nov 2026 {hhmm}:00 +0100"
    m["Message-ID"] = f"<{name}@sample.test>"
    m["Authentication-Results"] = f"mx.firm.example; {auth}"
    m.set_content(body)
    if attach:
        m.add_attachment(excel_order(), maintype="application",
                         subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         filename="orders_100234.xlsx")
    (OUT / f"{name}.eml").write_bytes(bytes(m))
    print("created", name + ".eml")
