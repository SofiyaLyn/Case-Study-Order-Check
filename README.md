# Order email checker (prototype)

Reads email files (.eml) from a folder and checks each one before any order would be processed.

For every email it does four steps:

1. **Check the sender**: email authentication results (SPF, DKIM), register of authorised signers, fraud patterns
2. **Read the order**: Excel attachment (fixed rules) or free text (simple rules, an NLP model in production), with a confidence score
3. **Run the rule checks**: ISIN check digit, instrument is tradable, restricted list, position or cash, duplicate, cut-off
4. **Decide**: ready for approval, human review, rejected, blocked, or held as a possible duplicate

A person always approves. Nothing is sent or executed.

All clients, ISINs, prices and positions are **sample data**.

## Run it in VS Code

1. Open this folder in VS Code (File, Open Folder).
2. Open a terminal (Terminal, New Terminal) and run:

```
python -m venv .venv
.venv\Scripts\activate          (Windows)   or   source .venv/bin/activate   (Mac, Linux)
pip install -r requirements.txt
python check_orders.py
```

The 8 test emails are already in `inbox/`. To recreate them, run `python make_samples.py`.

Results are printed in the terminal and saved to `output/results.csv` and `output/audit_log.txt`.

## Try your own test emails

1. Write a test email to yourself. Do not use real client or work emails.
2. Open it in Outlook and choose File, Save as, then pick the .eml format (or drag the email into a folder).
3. Put the file in `inbox/` and run the script again.

Note: if the saved file has no `Authentication-Results` header, the script reports SPF and DKIM as "missing" and blocks the email. That is intended: no proof of the sender means no processing.

## Files

| File | Purpose |
|---|---|
| `check_orders.py` | The checker |
| `make_samples.py` | Creates the 8 sample emails |
| `inbox/` | Input emails (.eml) |
| `output/` | Results (created when you run it) |

## Limits

- Free text is read with simple rules. A production version would add an NLP model, with the same checks.
- Reference data (instruments, clients, prices) is hard-coded. In production it comes from master data and registers.
- No connection to Outlook, an order system or Databricks.

## Next steps

- Read a shared mailbox through Microsoft Graph or Power Automate instead of files (read-only, approved by IT and security)
- Run in shadow mode first: label emails only, compare with what the team does by hand
- Write results to a Delta table in Databricks for dashboards and monitoring
