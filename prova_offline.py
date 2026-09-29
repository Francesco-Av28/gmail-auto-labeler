# -*- coding: utf-8 -*-
"""Collaudo OFFLINE di RIORDINO EMAIL: nessun accesso a Gmail.
Crea una casella finta (messaggi inventati, domini .example), calcola il piano con configurazione.json
e rubrica_esempio.csv, controlla che ogni conversazione finisca nell'etichetta attesa e salva una
DEMO del piano in demo\\piano_demo.xlsx (utile per vedere il risultato senza credenziali Google).
Scrive anche demo\\messaggi_esempio.csv: copiato come messaggi.csv permette di provare --piano offline.

    python prova_offline.py
"""
import os
import sys
from datetime import datetime, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import riordino_email as r

IO = "io.stesso@example.com"
cfg = r.carica_config()
rubrica = r.carica_rubrica(cfg)
etichette_esistenti = {"CONTATTI/GIALLI SERVIZI SRL": "Label_1", "VECCHIA ETICHETTA": "Label_2"}

# (thread, data, da, a, oggetto, anteprima, destinazione attesa)
CASI = [
    ("t01", "2015-03-02", "amico@example.net", IO, "Foto vacanze", "", "ARCHIVIO/2015"),
    ("t02", "2022-12-31", IO, "rossi.mario@example.com", "Preventivo", "", "ARCHIVIO/2022"),
    ("t03", "2024-05-10", "Rossi Mario <rossi.mario@example.com>", IO, "Documenti", "", "CONTATTI/ROSSI COSTRUZIONI SRL"),
    ("t04", "2023-02-01", IO, "anna.bianchi@example.org", "Richiesta", "", "CONTATTI/BIANCHI ANNA"),
    ("t05", "2025-07-07", "segreteria@altro.example", IO, "Pratica Verdi Logistica snc", "", "CONTATTI/VERDI LOGISTICA SNC"),
    ("t06", "2024-01-15", "info@fornitore.example", IO, "Contratto Gialli Servizi", "", "CONTATTI/GIALLI SERVIZI SRL"),
    ("t07", "2024-09-01", "avvisi@banca-esempio.example", IO, "Estratto conto trimestrale", "", "ARCHIVIO PER TIPO/BANCHE E PAGAMENTI"),
    ("t08", "2023-11-20", "noreply@servizio.example", IO, "Il tuo modello F24 e' pronto", "", "ARCHIVIO PER TIPO/FISCO E PREVIDENZA"),
    ("t09", "2025-03-03", "fatture@ditta.example", IO, "Fattura n. 12/2025", "", "ARCHIVIO PER TIPO/FATTURE ELETTRONICHE"),
    ("t10", "2024-04-04", "clienti@enel.example", IO, "La tua bolletta di aprile", "", "ARCHIVIO PER TIPO/UTENZE E ABBONAMENTI"),
    ("t11", "2023-06-06", "spedizioni@amazon.example", IO, "Il tuo ordine e' stato spedito", "", "ARCHIVIO PER TIPO/ACQUISTI ONLINE"),
    ("t12", "2024-08-08", "posta@pec.ditta.example", IO, "Comunicazione", "", "ARCHIVIO PER TIPO/PEC"),
    ("t13", "2025-01-01", "offerte@negozio.example", IO, "Newsletter di gennaio", "", "ARCHIVIO PER TIPO/NEWSLETTER E PROMOZIONI"),
    ("t14", "2024-02-02", IO, IO, "Appunti", "", "ARCHIVIO PER TIPO/AUTOINVIATE"),
    ("t15", "2023-03-03", "Luca <luca.privato@gmail.example>", IO, "Ciao", "", "ARCHIVIO PER TIPO/CONTATTI DA CLASSIFICARE"),
    ("t16", "2024-10-10", "info@associazione.example", IO, "Assemblea soci", "", "ARCHIVIO PER TIPO/ALTRO"),
    ("t17", "2026-02-02", "avvisi@banca-esempio.example", IO, "Estratto conto", "", ""),     # 2026: non toccare
    ("t18", "2024-12-12", "info@x.example", IO, "Offerta ROSSI", "", "ARCHIVIO PER TIPO/NEWSLETTER E PROMOZIONI"),  # nome parziale: no
]


def messaggio(i, tid, data, da, a, ogg, snip, inbox=True):
    ts = int(datetime.strptime(data, "%Y-%m-%d").replace(tzinfo=timezone.utc).timestamp() * 1000)
    return {"id": f"m{i:03d}", "threadId": tid, "internalDate": str(ts), "data": f"{data} 10:00",
            "labelIds": "INBOX" if inbox else "Label_2", "from": da, "to": a, "cc": "", "subject": ogg,
            "snippet": snip}


messaggi = [messaggio(i, *c[:6]) for i, c in enumerate(CASI)]
# conversazione su due anni: vale l'ultimo messaggio (2023 -> classificata, non archiviata per anno)
messaggi.append(messaggio(90, "t19", "2022-11-11", "anna.bianchi@example.org", IO, "Pratica", "", inbox=False))
messaggi.append(messaggio(91, "t19", "2023-01-09", IO, "anna.bianchi@example.org", "Re: Pratica", ""))
atteso = {c[0]: c[6] for c in CASI}
atteso["t19"] = "CONTATTI/BIANCHI ANNA"

piano = r.calcola_piano(messaggi, cfg, rubrica, etichette_esistenti, IO)
esiti = []
for p in piano:
    ok = p["destinazione"] == atteso[p["threadId"]]
    esiti.append(ok)
    print(f"[{'OK' if ok else 'FALLITO'}] {p['threadId']} {p['anno']} {p['metodo']:<18} -> "
          f"{p['destinazione'] or '(non toccata)'}" + ("" if ok else f"   ATTESO: {atteso[p['threadId']]}"))

t19 = next(p for p in piano if p["threadId"] == "t19")
esiti.append(t19["n_messaggi"] == 2 and t19["anno"] == 2023 and t19["messageIds"] == "m090|m091")
print(f"[{'OK' if esiti[-1] else 'FALLITO'}] conversazione su due anni: 2 messaggi, anno dell'ultimo (2023)")
esiti.append(len(piano) == len(atteso))
print(f"[{'OK' if esiti[-1] else 'FALLITO'}] una riga di piano per conversazione ({len(piano)})")

os.makedirs(os.path.join(r.BASE, "demo"), exist_ok=True)
# la casella finta anche come CSV: copiandola in messaggi.csv si prova --piano senza Gmail
import csv
with open(os.path.join(r.BASE, "demo", "messaggi_esempio.csv"), "w", encoding="utf-8", newline="") as f:
    w = csv.DictWriter(f, fieldnames=r.CAMPI)
    w.writeheader()
    w.writerows(messaggi)
r.PIANO_XLSX = os.path.join(r.BASE, "demo", "piano_demo.xlsx")
r.scrivi_excel(piano)
print(f"Demo del piano salvata in {r.PIANO_XLSX}")
print(f"\nRISULTATO: {sum(esiti)} controlli superati su {len(esiti)}")
sys.exit(0 if all(esiti) else 1)
