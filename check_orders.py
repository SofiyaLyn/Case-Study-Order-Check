"""
Order email checker (prototype).

Reads every .eml file in the inbox folder and, for each email:
  1. checks the sender (email authentication, register of signers, fraud pattern)
  2. reads the order (Excel attachment with fixed rules, free text with simple rules)
  3. runs the rule checks (ISIN, instrument, position or cash, duplicate, cut-off)
  4. decides: ready, review, rejected, blocked or hold

A person always approves. Nothing is sent or executed.
All clients, ISINs, prices and positions below are SAMPLE DATA.

Run:  python check_orders.py            (reads ./inbox, writes ./output)
      python check_orders.py my_folder  (reads another folder)
"""
import csv
import io
import json
import re
import sys
from email import policy
from email.parser import BytesParser
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path

from openpyxl import load_workbook

# ----------------------------------------------------------------------------
# Sample reference data (in production: master data and client registers)
# ----------------------------------------------------------------------------
CUTOFF = "16:00"

INSTRUMENTS = {
    "IE00WRLD0013": {"name": "World Equity ETF", "px": 98.40},
    "IE00ALLW0020": {"name": "All-World Equity ETF", "px": 128.10},
    "IE00SP5H0031": {"name": "US 500 ETF", "px": 640.20},
    "DE00DAXX0046": {"name": "German DAX ETF", "px": 185.30},
    "IE00BOND0050": {"name": "Euro Government Bond ETF", "px": 126.90},
}
RESTRICTED = set()

CLIENTS = {
    "100234": {"name": "Mueller Family Office",
               "signers": {"anna.mueller@mueller-fo.com", "thomas.keller@mueller-fo.com"},
               "cash": 250_000, "pos": {"IE00ALLW0020": 900, "IE00WRLD0013": 1500}},
    "100871": {"name": "Hartmann Stiftung",
               "signers": {"m.hartmann@hartmann-stiftung.de"},
               "cash": 80_000, "pos": {"IE00SP5H0031": 300}},
    "100455": {"name": "Brandt Vermoegen",
               "signers": {"j.brandt@brandt-vermoegen.at"},
               "cash": 40_000, "pos": {"IE00SP5H0031": 1200, "IE00BOND0050": 2000}},
}

DECISION_LABEL = {
    "ready": "Ready for approval",
    "review": "Human review",
    "rejected": "Rejected, ask client",
    "blocked": "Blocked, security",
    "hold": "Held: possible duplicate",
}


# ----------------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------------
def valid_isin(isin: str) -> bool:
    """Format check plus check digit (Luhn on the letters-to-digits string)."""
    if not re.fullmatch(r"[A-Z]{2}[A-Z0-9]{9}[0-9]", isin):
        return False
    digits = "".join(str(int(c, 36)) for c in isin)
    total = 0
    for i, ch in enumerate(reversed(digits)):
        n = int(ch)
        if i % 2 == 1:
            n *= 2
            if n > 9:
                n -= 9
        total += n
    return total % 10 == 0


def levenshtein(a: str, b: str) -> int:
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def client_by_sender(address: str):
    for account, c in CLIENTS.items():
        if address in c["signers"]:
            return account
    return None


# ----------------------------------------------------------------------------
# Step 0: read the .eml file
# ----------------------------------------------------------------------------
def read_eml(path: Path) -> dict:
    with path.open("rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)
    body_part = msg.get_body(preferencelist=("plain",))
    attachments = []
    for part in msg.iter_attachments():
        attachments.append({"name": part.get_filename() or "", "data": part.get_payload(decode=True)})
    auth = msg.get("Authentication-Results", "") or ""
    when = parsedate_to_datetime(msg["Date"])
    return {
        "file": path.name,
        "from": parseaddr(msg["From"])[1].lower(),
        "subject": msg["Subject"] or "",
        "time": when.strftime("%H:%M"),
        "sort_key": when,
        "body": body_part.get_content() if body_part else "",
        "attachments": attachments,
        "spf": (re.search(r"spf=(\w+)", auth) or [None, "missing"])[1],
        "dkim": (re.search(r"dkim=(\w+)", auth) or [None, "missing"])[1],
    }


# ----------------------------------------------------------------------------
# Step 2: read the order
# Excel: fixed rules (fully reliable). Free text: simple rules here;
# in production an NLP model would handle messier wording.
# ----------------------------------------------------------------------------
def read_order(mail: dict) -> dict:
    out = {"cls": "order", "source": "Free text", "lines": []}

    for att in mail["attachments"]:
        if att["name"].lower().endswith(".xlsx"):
            out["source"] = "Excel template"
            ws = load_workbook(io.BytesIO(att["data"]), read_only=True).active
            for row in list(ws.iter_rows(values_only=True))[1:]:
                if not row or not row[0]:
                    continue
                account, side, isin, qty = row[:4]
                out["lines"].append({
                    "account": str(account), "side": str(side).upper(), "isin": str(isin),
                    "qty": int(qty), "amount": None, "inferred": False, "all": False, "half": False,
                    "conf": {"account": 1.0, "side": 1.0, "isin": 1.0, "size": 1.0}})
            return out

    text = mail["body"]
    low = text.lower()
    if re.search(r"cancel|withdraw|revoke", low):
        out["cls"] = "cancellation"
    elif not re.search(r"\b(buy|sell|purchase|redeem)\b", low):
        out["cls"] = "question"

    side = "SELL" if re.search(r"\b(sell|redeem)\b", low) else ("BUY" if re.search(r"\b(buy|purchase)\b", low) else None)
    m_isin = re.search(r"\b[A-Z]{2}[A-Z0-9]{9}[0-9]\b", text)
    m_acc = re.search(r"account\s*(?:no\.?|number|#)?\s*:?\s*(\d{6})", text, re.I)
    m_qty = re.search(r"(\d[\d,]*)\s*units", text, re.I)
    m_amt = re.search(r"EUR\s*([\d,]+)", text, re.I)
    half = bool(re.search(r"\bhalf\b", low))
    all_ = bool(re.search(r"\ball (of my |our )?holdings\b", low))

    account, inferred = (m_acc.group(1) if m_acc else None), False
    if not account:
        guess = client_by_sender(mail["from"])
        if guess:
            account, inferred = guess, True

    out["lines"].append({
        "account": account, "side": side,
        "isin": m_isin.group(0) if m_isin else None,
        "qty": int(m_qty.group(1).replace(",", "")) if m_qty else None,
        "amount": int(m_amt.group(1).replace(",", "")) if m_amt else None,
        "inferred": inferred, "all": all_, "half": half,
        "conf": {
            "account": 0.99 if m_acc else (0.8 if inferred else 0.0),
            "side": 0.98 if side else 0.0,
            "isin": 0.99 if m_isin else 0.0,
            "size": 0.97 if (m_qty or m_amt) else (0.6 if all_ else (0.3 if half else 0.0)),
        }})
    return out


# ----------------------------------------------------------------------------
# Steps 1, 3, 4: sender checks, rule checks, decision
# ----------------------------------------------------------------------------
def check_mail(mail: dict, order: dict, earlier: list) -> dict:
    res = {"checks": [], "exceptions": [], "suggestion": "", "reply": "", "decision": "", "reason": ""}

    def add(status, label, detail):
        res["checks"].append((status, label, detail))

    # --- Step 1: sender
    blocked = False
    auth_ok = mail["spf"] == "pass" and mail["dkim"] == "pass"
    add("PASS" if auth_ok else "FAIL", "Email authentication (SPF, DKIM)",
        "both passed" if auth_ok else f"spf={mail['spf']}, dkim={mail['dkim']}: sender domain not verified")
    if not auth_ok:
        blocked = True
        res["exceptions"].append("Sender not verified")

    account_of_sender = client_by_sender(mail["from"])
    if account_of_sender:
        add("PASS", "Sender is an authorised signer", CLIENTS[account_of_sender]["name"])
    else:
        domain = mail["from"].split("@")[1]
        near = None
        for c in CLIENTS.values():
            for s in c["signers"]:
                d2 = s.split("@")[1]
                if d2 != domain and levenshtein(d2, domain) <= 2:
                    near = d2
        add("FAIL", "Sender is an authorised signer",
            "address not in the register" + (f", domain almost identical to {near}" if near else ""))
        blocked = True
        if "Sender not verified" not in res["exceptions"]:
            res["exceptions"].append("Sender not verified")

    if re.search(r"urgent|do not call|new account|new bank", mail["body"], re.I):
        add("FAIL", "Known fraud pattern in the text", "urgency, no call-back, request to pay a new account")
        blocked = True
        res["exceptions"].append("Fraud pattern")

    # --- Step 3: rule checks per order line
    low_conf = dup = failed = False
    late = mail["time"] > CUTOFF
    for n, line in enumerate(order["lines"], 1):
        prefix = f"[line {n}] " if len(order["lines"]) > 1 else ""
        conf = line["conf"]
        min_conf = min(conf.values())
        missing = not (line["account"] and line["side"] and line["isin"]) or (line["qty"] is None and line["amount"] is None) or line["all"]
        if missing or min_conf < 0.8:
            low_conf = True
            unclear = [k for k, v in conf.items() if v < 0.8]
            add("FAIL", prefix + "All fields found with high confidence", "unclear or missing: " + ", ".join(unclear))
            for label in ("ISIN check digit", "Instrument tradable", "Restricted list", "Position or cash", "Duplicate check"):
                add("SKIP", prefix + label, "skipped until a person clarifies the order")
            res["exceptions"].append("Unclear instruction")
            continue

        add("PASS", prefix + "All fields found with high confidence", "account, side, ISIN, size")
        isin, qty, note = line["isin"], line["qty"], ""

        ok_isin = valid_isin(isin)
        add("PASS" if ok_isin else "FAIL", prefix + "ISIN check digit",
            f"{isin} is valid" if ok_isin else f"{isin} fails the check digit test (likely a typo)")
        inst = INSTRUMENTS.get(isin) if ok_isin else None
        if not ok_isin:
            failed = True
            res["exceptions"].append("ISIN error")
            close = [k for k in INSTRUMENTS if levenshtein(k, isin) <= 2]
            if close:
                res["suggestion"] = f"Did the client mean {close[0]} ({INSTRUMENTS[close[0]]['name']})? Ask to confirm, never assume."
            add("SKIP", prefix + "Instrument tradable", "skipped, ISIN invalid")
        else:
            add("PASS" if inst else "FAIL", prefix + "Instrument tradable",
                inst["name"] if inst else "not found in the instrument master")
            if not inst:
                failed = True
                res["exceptions"].append("Unknown instrument")
        add("FAIL" if isin in RESTRICTED else ("PASS" if inst else "SKIP"), prefix + "Restricted list",
            "on the restricted list" if isin in RESTRICTED else ("not on the list" if inst else "skipped"))

        client = CLIENTS.get(line["account"])
        if client and account_of_sender and mail["from"] not in client["signers"]:
            add("FAIL", prefix + "Sender may act on this account", f"not authorised for account {line['account']}")
            failed = True
            res["exceptions"].append("Sender not verified")

        if inst and client:
            px = inst["px"]
            if line["amount"] is not None and qty is None:
                qty = int(line["amount"] // px)
                line["qty"] = qty
                note = f" (EUR {line['amount']:,} at sample price {px:.2f} = {qty} units, to be confirmed)"
            if line["side"] == "SELL":
                have = client["pos"].get(isin, 0)
                ok = have >= qty
                add("PASS" if ok else "FAIL", prefix + "Position is large enough", f"holds {have:,} units, order is {qty:,}{note}")
                if not ok:
                    failed = True
                    res["exceptions"].append("Not enough position")
            else:
                cost = qty * px
                ok = client["cash"] >= cost
                add("PASS" if ok else "FAIL", prefix + "Cash is sufficient", f"needs about EUR {cost:,.0f}, available EUR {client['cash']:,}{note}")
                if not ok:
                    failed = True
                    res["exceptions"].append("Not enough cash")
        else:
            add("SKIP", prefix + "Position or cash", "skipped")

        same = [e for e in earlier for l in e["lines"]
                if (l["account"], l["isin"], l["side"], l["qty"], l["amount"]) ==
                   (line["account"], line["isin"], line["side"], line["qty"], line["amount"])]
        add("FAIL" if same else "PASS", prefix + "Duplicate check",
            f"same order as {same[0]['file']} ({same[0]['time']})" if same else "no earlier identical order today")
        if same:
            dup = True
            res["exceptions"].append("Possible duplicate")

        if late:
            add("WARN", prefix + f"Cut-off {CUTOFF}", f"received {mail['time']}, after the cut-off: executes next trading day, tell the client")
            res["exceptions"].append("After cut-off")
        else:
            add("PASS", prefix + f"Cut-off {CUTOFF}", f"received {mail['time']}, in time")

    # --- Step 4: decision
    if blocked:
        res["exceptions"] = [e for e in res["exceptions"] if e != "Unclear instruction"]
        res["decision"] = "blocked"
        res["reason"] = "Nothing is processed. Security calls the client on a known phone number first."
        res["reply"] = "No reply to the sender. Alert to security and compliance."
    elif order["cls"] != "order":
        res["decision"], res["reason"] = "review", "Not a new order, so a person handles it."
    elif low_conf:
        res["decision"] = "review"
        res["reason"] = "The instruction is unclear. The system does not guess: a person confirms with the client."
        res["reply"] = "Dear client, to process your request we need the exact instrument (ISIN) and quantity. Could you confirm?"
        client = CLIENTS.get(account_of_sender) if account_of_sender else None
        if client and order["lines"][0].get("half"):
            for k, units in client["pos"].items():
                if "US" in INSTRUMENTS[k]["name"]:
                    res["suggestion"] = (f"Possible match from the client's holdings: {INSTRUMENTS[k]['name']} ({k}), "
                                         f"{units:,} units. Half would be {units // 2:,} units. For the reviewer to confirm.")
    elif dup:
        res["decision"] = "hold"
        res["reason"] = "The same order arrived earlier. Hold it so it is not executed twice, and ask the client."
        res["reply"] = "Dear client, we already received this order earlier today. Please tell us if this is a second order."
    elif failed:
        res["decision"] = "rejected"
        res["reason"] = "A check failed. The order does not go to the order system. The client gets a clear reason."
        res["reply"] = "Dear client, we cannot process this order as sent. Could you check the details and resend?"
    else:
        res["decision"] = "ready"
        res["reason"] = "All checks passed. The order is prepared. A person approves it (four-eyes) before release."
        res["reply"] = ("Dear client, your order arrived after the cut-off and will be executed on the next trading day."
                        if late else "Dear client, your order was received and validated. We will confirm execution.")
    return res


# ----------------------------------------------------------------------------
# Run
# ----------------------------------------------------------------------------
def main():
    inbox = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("inbox")
    files = sorted(inbox.glob("*.eml"))
    if not files:
        print(f"No .eml files found in {inbox.resolve()}. Run: python make_samples.py")
        return
    mails = sorted((read_eml(f) for f in files), key=lambda m: m["sort_key"])

    done, rows, audit, check_rows, cases = [], [], [], [], []
    for mail in mails:
        order = read_order(mail)
        mail["lines"] = order["lines"]
        res = check_mail(mail, order, done)
        done.append(mail)

        print("=" * 78)
        print(f"{mail['file']}  |  {mail['time']}  |  from {mail['from']}")
        print(f"Subject: {mail['subject']}   ({order['source']}, type: {order['cls']})")
        for status, label, detail in res["checks"]:
            print(f"  [{status:<4}] {label}: {detail}")
        print(f"  DECISION: {DECISION_LABEL[res['decision']]}")
        print(f"  Why: {res['reason']}")
        if res["suggestion"]:
            print(f"  Suggestion for the reviewer: {res['suggestion']}")
        if res["reply"]:
            print(f"  Draft reply: {res['reply']}")

        total = sum(1 for c in res["checks"] if c[0] != "SKIP")
        passed = sum(1 for c in res["checks"] if c[0] == "PASS")
        audit.append(f"{mail['time']} | {mail['file']} | {res['decision'].upper()} | checks passed {passed}/{total} | lines {len(order['lines'])}")
        for status, label, detail in res["checks"]:
            check_rows.append({"file": mail["file"], "check": label, "status": status, "detail": detail})
        cases.append({"file": mail["file"], "time": mail["time"], "from": mail["from"], "subject": mail["subject"],
                      "source": order["source"], "type": order["cls"], "lines": order["lines"],
                      "checks": [{"status": st, "label": lb, "detail": dt} for st, lb, dt in res["checks"]],
                      "decision": res["decision"], "reason": res["reason"],
                      "suggestion": res["suggestion"], "reply": res["reply"]})
        rows.append({"file": mail["file"], "time": mail["time"], "from": mail["from"], "subject": mail["subject"],
                     "decision": res["decision"], "exceptions": "; ".join(dict.fromkeys(res["exceptions"])),
                     "reason": res["reason"], "suggestion": res["suggestion"], "draft_reply": res["reply"]})

    # summary
    print("=" * 78)
    print("SUMMARY")
    counts = {}
    for r in rows:
        counts[r["decision"]] = counts.get(r["decision"], 0) + 1
    for key, label in DECISION_LABEL.items():
        print(f"  {label:<28} {counts.get(key, 0)}")
    print(f"  {'Total emails':<28} {len(rows)}")

    out = Path("output")
    out.mkdir(exist_ok=True)
    with (out / "results.csv").open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    with (out / "checks.csv").open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["file", "check", "status", "detail"])
        w.writeheader()
        w.writerows(check_rows)
    (out / "cases.json").write_text(json.dumps(cases, indent=2, ensure_ascii=False), encoding="utf-8")
    (out / "audit_log.txt").write_text("\n".join(audit) + "\n", encoding="utf-8")
    print(f"\nWrote {out / 'results.csv'}, {out / 'checks.csv'}, {out / 'cases.json'} and {out / 'audit_log.txt'}")


if __name__ == "__main__":
    main()
