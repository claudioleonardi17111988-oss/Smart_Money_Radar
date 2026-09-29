#!/usr/bin/env python3
"""
Forward test di Money Flow Pro.

Legge data/picks_history.csv (scritto ogni giorno da money_flow_pro.py) e misura, per ogni segnalazione,
il rendimento a +5 / +20 / +60 sedute rispetto a SPY. Confronta:
  - Top 10
  - tutti i verdetti azionabili
  - tutti i candidati
  - quintili di score (Q5 = score piu' alto)
Se Q5 non batte Q1 con continuita', lo score non sta ordinando bene i rendimenti: va ricalibrato.

Convenzioni: ingresso = chiusura della seduta del segnale (approssimazione: nella realta' si entra
il giorno dopo). Nessun costo di transazione. Campioni piccoli = rumore: servono mesi di storico.

Uso: python forward_test.py            (scrive data/forward_test.md)
"""
from __future__ import annotations

import logging
import sys
from collections import Counter
from pathlib import Path

import numpy as np
import pandas as pd

from money_flow_pro import HISTORY_FILE, DATA_DIR, scarica_prezzi

log = logging.getLogger("forward_test")
ORIZZONTI = (5, 20, 60)
OUT_FILE = DATA_DIR / "forward_test.md"


def calcola_forward(storico: pd.DataFrame, closes: dict[str, pd.Series], orizzonti=ORIZZONTI) -> pd.DataFrame:
    """Una riga per (segnalazione, orizzonte) con rendimento e rendimento in eccesso su SPY."""
    spy = closes.get("SPY")
    if spy is None:
        raise ValueError("Serie SPY mancante")
    righe = []
    for r in storico.itertuples(index=False):
        s = closes.get(r.symbol)
        if s is None or s.empty:
            continue
        d = pd.Timestamp(r.date)
        i = s.index.searchsorted(d)
        if i >= len(s) or s.index[i] != d:
            continue
        for h in orizzonti:
            if i + h >= len(s):
                continue
            d_fine = s.index[i + h]
            ret = s.iloc[i + h] / s.iloc[i] - 1
            b0, b1 = spy.asof(d), spy.asof(d_fine)
            if not (np.isfinite(b0) and np.isfinite(b1)) or b0 <= 0:
                continue
            righe.append({"date": r.date, "symbol": r.symbol, "h": h, "ret": ret,
                          "exc": ret - (b1 / b0 - 1),
                          "top10": bool(r.top10), "azionabile": bool(r.azionabile), "q": r.q})
    return pd.DataFrame(righe)


def riepilogo(fw: pd.DataFrame) -> pd.DataFrame:
    gruppi = {
        "Top 10": fw[fw["top10"]],
        "Azionabili": fw[fw["azionabile"]],
        "Tutti i candidati": fw,
    }
    for q in (5, 4, 3, 2, 1):
        gruppi[f"Score Q{q}" + (" (piu' alto)" if q == 5 else " (piu' basso)" if q == 1 else "")] = fw[fw["q"] == q]
    righe = []
    for nome, g in gruppi.items():
        for h, gh in g.groupby("h"):
            if gh.empty:
                continue
            righe.append({"Gruppo": nome, "Orizzonte (sedute)": int(h), "N": len(gh),
                          "Rend. medio %": gh["ret"].mean() * 100, "Rend. mediano %": gh["ret"].median() * 100,
                          "Eccesso vs SPY %": gh["exc"].mean() * 100, "Hit rate vs SPY %": (gh["exc"] > 0).mean() * 100})
    return pd.DataFrame(righe)


def _markdown(df: pd.DataFrame) -> str:
    if df.empty:
        return "_Ancora nessun dato: servono almeno 5 sedute dalla prima segnalazione._\n"
    cols = list(df.columns)
    out = ["| " + " | ".join(cols) + " |", "|" + "|".join(["---"] * len(cols)) + "|"]
    for _, r in df.iterrows():
        cells = [f"{v:.1f}" if isinstance(v, float) else str(v) for v in r]
        out.append("| " + " | ".join(cells) + " |")
    return "\n".join(out) + "\n"


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-7s %(message)s", datefmt="%H:%M:%S")
    if not Path(HISTORY_FILE).exists():
        log.warning("Nessuno storico (%s): niente da analizzare.", HISTORY_FILE)
        return 0
    storico = pd.read_csv(HISTORY_FILE, dtype={"date": str})
    storico["q"] = np.nan
    for d, idx in storico.groupby("date").groups.items():
        if len(idx) >= 5:
            storico.loc[idx, "q"] = pd.qcut(storico.loc[idx, "score"].rank(method="first"), 5, labels=False) + 1

    simboli = sorted(set(storico["symbol"]) | {"SPY"})
    start = (pd.Timestamp(storico["date"].min()) - pd.Timedelta(days=7)).strftime("%Y-%m-%d")
    stats: Counter = Counter()
    prezzi = scarica_prezzi(simboli, stats, batch_size=100, start=start)
    closes = {}
    for tk, d in prezzi.items():
        if "Close" in d.columns:
            s = d["Close"].dropna()
            if getattr(s.index, "tz", None) is not None:
                s.index = s.index.tz_localize(None)
            closes[tk] = s

    fw = calcola_forward(storico, closes)
    tab = riepilogo(fw) if not fw.empty else pd.DataFrame()
    testo = (f"# Forward test Money Flow Pro\n\nSegnalazioni storiche: {len(storico)} "
             f"su {storico['date'].nunique()} sedute (dal {storico['date'].min()} al {storico['date'].max()}).\n\n"
             + _markdown(tab)
             + "\n_Ingresso alla chiusura del segnale, nessun costo. Con pochi campioni i risultati sono rumore._\n")
    OUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    OUT_FILE.write_text(testo, encoding="utf-8")
    print(testo)
    return 0


if __name__ == "__main__":
    sys.exit(main())
