# Gmail Auto-Labeler — rule-based cleanup of a Gmail archive

A Python tool on the **official Gmail API** that puts every conversation of a mailbox under a label, following rules written in a JSON file. Old conversations go to `ARCHIVIO/<year>`; recent ones go to the client they belong to (address book match, or the client's name in the subject) or to a category (banks, tax & social security, e-invoices, utilities, PEC, online shopping, newsletters…).

Used at an Italian tax & payroll advisory firm: **32,434 messages archived by year and 21,427 conversations classified across 243 clients and 21 categories.**

## Safe by design

- **Nothing is ever deleted**: the tool only adds labels and, if configured, removes the conversation from the inbox.
- Only metadata is read (sender, recipients, subject, snippet, date), never the message body.
- **Dry run first**: `--piano` writes the full plan to an Excel sheet with no change to the mailbox; you review it, fix rules, and re-run.
- `--esegui` asks you to type `SI`, can run one year at a time (`--anni 2023,2024`) and resumes where it stopped. API quota errors are retried with backoff.

## How it decides

```
conversation (year of its LAST message)
   │
   ├─ year in an "archive" band ───────────► ARCHIVIO/<year>
   ├─ year in a "classify" band:
   │     1. address in the address book ───► CONTATTI/<name>
   │     2. contact name in subject/snippet► CONTATTI/<name>
   │     3. first matching rule (sender OR subject regex) ► ARCHIVIO PER TIPO/<category>
   │     4. fallbacks: sent by me / personal mailbox / other
   └─ year in a "no action" band ──────────► untouched
```

## Try it offline (no Google account needed)

```
pip install -r requisiti.txt
python prova_offline.py     # 21 checks on a fake mailbox; writes demo/piano_demo.xlsx
```

## Use it on your mailbox

Create your own OAuth desktop credentials in Google Cloud (Gmail API enabled), save them as `credentials.json`, then:

```
python riordino_email.py --login     # authorise
python riordino_email.py --estrai    # read-only: metadata to messaggi.csv (resumable)
python riordino_email.py --piano     # dry run: piano_riordino.xlsx, no changes
python riordino_email.py --esegui    # apply (asks for confirmation)
```

Step-by-step instructions in Italian are in `LEGGIMI.txt`. Rules, year bands and label names are in `configurazione.json`. The address book and all sample messages are **fictitious**; credentials, tokens and extracted data are git-ignored.

## Stack

Python · Gmail API (google-api-python-client, OAuth 2.0) · regex · openpyxl

---

## In italiano

Strumento Python basato sulle API ufficiali di Gmail che classifica le conversazioni di una casella con etichette per anno, per cliente (rubrica o nome nell'oggetto) e per categoria (regole regex su mittente e oggetto). Non cancella mai nulla, legge solo i metadati, ha una prova a secco con piano in Excel e riprende da dove si era fermato. Usato in uno studio di consulenza fiscale e del lavoro: **32.434 messaggi archiviati e 21.427 conversazioni classificate per 243 clienti e 21 categorie.** Collaudo offline con 21 controlli su una casella finta.

Author: Francesco Aversano
