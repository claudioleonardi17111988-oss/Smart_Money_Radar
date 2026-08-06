def main():
  tickers = ottieni_sp500()
  print(
      f"🚀 Avvio Smart Money Radar su {len(tickers)} titoli dell'S&P 500..."
  )

  try:
    df_raw = yf.download(tickers, period="1y", auto_adjust=True, progress=False)
  except Exception as e:
    print(f"Errore critico durante il download dati bulk: {e}")
    invia_telegram(
        CANALE_ACCUMULAZIONE_ID,
        "❌ Errore critico nel download dei dati di borsa.",
    )
    return

  if isinstance(df_raw.columns, pd.MultiIndex):
    if "Close" in df_raw.columns.levels[0]:
      df_close = df_raw["Close"]
      df_volume = df_raw["Volume"]
    elif "Close" in df_raw.columns.levels[1]:
      df_close = df_raw.xs("Close", axis=1, level=1)
      df_volume = df_raw.xs("Volume", axis=1, level=1)
    else:
      df_close = df_raw
      df_volume = None
  else:
    df_close = df_raw
    df_volume = None

  candidati = []

  for ticker in df_close.columns:
    try:
      ticker_str = str(ticker)
      chiusure = df_close[ticker].dropna()

      if len(chiusure) < 200:
        continue

      prezzo_attuale = float(chiusure.iloc[-1])
      massimo_52w = float(chiusure.max())

      storno_pct = ((massimo_52w - prezzo_attuale) / massimo_52w) * 100

      # FILTRO 1: Storno minimo del 15% dai max 52W
      if storno_pct < SOGLIA_STORNO_MINIMA:
        continue

      # CALCOLO MEDIA VOLUMI 1 SETTIMANA (5 GIORNI) vs MEDIA 60 GIORNI
      rvol_5d_pct = 0.0
      if df_volume is not None and ticker in df_volume.columns:
        volumi = df_volume[ticker].dropna()
        if len(volumi) >= 60:
          media_vol_5g = volumi.iloc[-5:].mean()
          media_vol_60g = volumi.iloc[-60:].mean()
          if media_vol_60g > 0:
            rvol_5d_pct = (media_vol_5g / media_vol_60g) * 100

      # FILTRO 2: Volumi della settimana almeno al 115%
      if rvol_5d_pct < SOGLIA_VOLUMI_SETTIMANA:
        continue

      # FILTRO 3: Bilanci sani
      t_obj = yf.Ticker(ticker_str)
      if not verifica_fondamentali_sani(t_obj):
        continue

      rsi_serie = calcola_rsi(chiusure)
      rsi_attuale = float(rsi_serie.iloc[-1])

      sma_200 = float(chiusure.rolling(window=200).mean().iloc[-1])
      is_bear_market = prezzo_attuale < sma_200

      candidati.append({
          "ticker": ticker_str,
          "prezzo": prezzo_attuale,
          "rsi": rsi_attuale,
          "storno": storno_pct,
          "is_bear": is_bear_market,
          "rvol_5d": rvol_5d_pct,
      })
    except Exception:
      continue

  print(
      f"✅ Scansione completata. {len(candidati)} titoli in fase di accumulazione"
      " rilevati oggi!"
  )

  if not candidati:
    invia_telegram(
        CANALE_ACCUMULAZIONE_ID,
        "ℹ️ **Smart Money Radar**: Nessun titolo S&P 500 in storno > 15%"
        " presenta accumulazione di volumi (VOL 1W >= 115%) nella sessione"
        " odierna.",
    )
    return

  candidati_ordinati = sorted(
      candidati, key=lambda x: x["rvol_5d"], reverse=True
  )

  dips_bull_market = []
  bear_market_watchlist = []

  for c in candidati_ordinati:
    info_vol = f"VOL 1W: {c['rvol_5d']:.0f}%"
    info_storno = f"-{c['storno']:.1f}% dai max"

    # FORMATTAZIONE RSI (Grassetto ed evidenziazione se < 30)
    if c["rsi"] < 30:
      info_rsi = f"**RSI: {c['rsi']:.0f} (Ipervenduto)**"
    else:
      info_rsi = f"RSI: {c['rsi']:.0f}"

    if not c["is_bear"]:
      riga_str = (
          f"• 🟢 **{c['ticker']}** (${c['prezzo']:.1f} | {info_vol} |"
          f" {info_storno} | {info_rsi})"
      )
      dips_bull_market.append(riga_str)
    else:
      riga_str = (
          f"• 🔴 **{c['ticker']}** (${c['prezzo']:.1f} | {info_vol} |"
          f" {info_storno} | {info_rsi})"
      )
      bear_market_watchlist.append(riga_str)

  righe = ["📡 **SMART MONEY RADAR - REPORT ACCUMULAZIONE**\n"]

  if dips_bull_market:
    righe.append("🟢 **ACCUMULAZIONE IN BULL TREND (Sopra SMA200)**")
    righe.extend(dips_bull_market)
    righe.append("")

  if bear_market_watchlist:
    righe.append("🔴 **ACCUMULAZIONE IN BEAR TREND (Sotto SMA200)**")
    righe.extend(bear_market_watchlist)

  msg = ""
  for r in righe:
    if len(msg) + len(r) + 1 > 3800:
      invia_telegram(CANALE_ACCUMULAZIONE_ID, msg)
      msg = r + "\n"
    else:
      msg += r + "\n"
  if msg:
    invia_telegram(CANALE_ACCUMULAZIONE_ID, msg)
