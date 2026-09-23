from datetime import datetime
import email
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import smtplib
import numpy as np
import pandas as pd
import requests
import yfinance as yf

# =====================================================================
# CONFIGURAZIONE TELEGRAM, EMAIL & PORTAFOGLIO PERSONALE
# =====================================================================
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
CANALE_ACCUMULAZIONE_ID = os.environ.get(
    "CANALE_ACCUMULAZIONE_ID", "-1003454283658"
)
SENDER_EMAIL = os.environ.get("SENDER_EMAIL")
APP_PASSWORD = os.environ.get("APP_PASSWORD")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL", SENDER_EMAIL)

# ---------------------------------------------------------------------
# 🎯 CONFIGURAZIONE PORTAFOGLIO:
# - 'pmc': Prezzo Medio di Carico in EURO (€)
# - 'core': True se è un titolo strategico di lungo termine (es. Tesi AI / Fotonica)
#           False se è un titolo tattico/speculativo
# ---------------------------------------------------------------------
MEI_PORTAFOGLIO_CONFIG = {
    "APP": {"pmc": 263.53, "core": False},
    "RDDT": {"pmc": 131.10, "core": False},
    "BBWI": {"pmc": 16.34, "core": False},
    "CEG": {"pmc": 236.10, "core": True},
    "MARA": {"pmc": 10.53, "core": True},
    "ON": {"pmc": 70.87, "core": False},
    "VRT": {"pmc": 250.45, "core": True},
    "WMT": {"pmc": 102.48, "core": False},
    "MU": {"pmc": 812.36, "core": True},
    "WULF": {"pmc": 19.35, "core": False},
    "POET": {"pmc": 7.24, "core": True},
    "CLS": {"pmc": 296.91, "core": True},
    "HIVE.TO": {"pmc": 2.50, "core": True},
    "COHR": {"pmc": 261.22, "core": True},
    "CIFR": {"pmc": 21.88, "core": False},
    "AZO": {"pmc": 2525.50, "core": False},
    "FIX": {"pmc": 1495.90, "core": True},
    "NFLX": {"pmc": 65.31, "core": True},
    "NVDA": {"pmc": 184.62, "core": True},
    "QCOM": {"pmc": 146.84, "core": False},
    "TDG": {"pmc": 1071.77, "core": True},
    "VST": {"pmc": 132.48, "core": True},
    "ROL": {"pmc": 31.0, "core": False},
    "NVO": {"pmc": 37.58, "core": True},
    "AMTM": {"pmc": 19.73, "core": True},
    "BMY": {"pmc": 50.60, "core": True},
    "FCT.MI": {"pmc": 15.71, "core": True},
    "SOFI": {"pmc": 14.87, "core": True},
    "NU": {"pmc": 11.98, "core": True},
    "ZENA": {"pmc": 2.03, "core": True},
    "ADUR": {"pmc": 12.91, "core": True},
    "XYL": {"pmc": 97.73, "core": True},
    # "POET": {"pmc": 1.50, "core": True},  # Esempio inserimento POET Technology
}


# =====================================================================
# FUNZIONI DI NOTIFICA & INVIO MAIL
# =====================================================================
def invia_telegram(canale_id, messaggio):
  """Invia il messaggio su Telegram formattato in Markdown."""
  if not TELEGRAM_TOKEN or not canale_id:
    print("⚠️ Telegram non configurato. Stampa a video del messaggio:")
    print(messaggio)
    return
  url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
  payload = {"chat_id": canale_id, "text": messaggio, "parse_mode": "Markdown"}
  try:
    r = requests.post(url, json=payload, timeout=10)
    print(f"Esito invio Telegram: {r.status_code}")
  except Exception as e:
    print(f"Errore invio Telegram: {e}")


def invia_email_con_allegato(oggetto, corpo_testo, file_excel_path):
  """Invia un'e-mail via SMTP Gmail allegando il file Excel generato."""
  if not SENDER_EMAIL or not APP_PASSWORD or not RECEIVER_EMAIL:
    print("⚠️ Credenziali E-mail non configurate. Saltato invio mail.")
    return
  msg = MIMEMultipart()
  msg["From"] = SENDER_EMAIL
  msg["To"] = RECEIVER_EMAIL
  msg["Subject"] = oggetto
  msg.attach(MIMEText(corpo_testo, "plain", "utf-8"))
  if os.path.exists(file_excel_path):
    with open(file_excel_path, "rb") as f:
      part = MIMEApplication(f.read(), Name=os.path.basename(file_excel_path))
      part[
          "Content-Disposition"
      ] = f'attachment; filename="{os.path.basename(file_excel_path)}"'
      msg.attach(part)
  try:
    print(f"📧 Invio e-mail in corso a {RECEIVER_EMAIL}...")
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(SENDER_EMAIL, APP_PASSWORD)
    server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())
    server.quit()
    print("✅ E-mail inviata con successo!")
  except Exception as e:
    print(f"❌ Errore durante l'invio dell'e-mail: {e}")


def ottieni_tasso_cambio_eur_usd():
  """Scarica il tasso di cambio EUR/USD in tempo reale via yfinance."""
  try:
    fx = yf.Ticker("EURUSD=X")
    hist = fx.history(period="1d")
    if not hist.empty:
      return float(hist["Close"].iloc[-1])
  except Exception as e:
    print(f"⚠️ Impossibile scaricare il cambio EUR/USD, uso default 1.08: {e}")
  return 1.08


# =====================================================================
# CALCOLI ANALISI TECNICA AVANZATA
# =====================================================================
def calcola_rsi(chiusure, periodi=14):
  delta = chiusure.diff()
  guadagno = (delta.where(delta > 0, 0)).rolling(window=periodi).mean()
  perdita = (-delta.where(delta < 0, 0)).rolling(window=periodi).mean()
  rs = guadagno / perdita
  return 100 - (100 / (1 + rs))


def calcola_obv(chiusure, volumi):
  obv = (np.sign(chiusure.diff()) * volumi).fillna(0).cumsum()
  return obv


def calcola_cmf(massimi, minimi, chiusure, volumi, periodi=20):
  mf_multiplier = ((chiusure - minimi) - (massimi - chiusure)) / (
      massimi - minimi
  )
  mf_multiplier = mf_multiplier.fillna(0)
  mf_volume = mf_multiplier * volumi
  cmf = mf_volume.rolling(window=periodi).sum() / volumi.rolling(
      window=periodi
  ).sum()
  return cmf


def calcola_close_location_value(chiusura, minimo, massimo):
  range_giornaliero = massimo - minimo
  if range_giornaliero == 0:
    return 0.5
  return (chiusura - minimo) / range_giornaliero


def calcola_volume_poc(chiusure, volumi, periodi=60, bins=10):
  if len(chiusure) < periodi:
    return float(chiusure.iloc[-1])
  sub_close = chiusure.iloc[-periodi:]
  sub_vol = volumi.iloc[-periodi:]
  counts, bin_edges = np.histogram(sub_close, bins=bins, weights=sub_vol)
  max_idx = np.argmax(counts)
  poc_price = (bin_edges[max_idx] + bin_edges[max_idx + 1]) / 2.0
  return float(poc_price)


def ottieni_dati_fondamentali_e_anagrafica(ticker_obj, ticker_str):
  nome_azienda = ""
  fwd_pe = "N/D"
  peg = "N/D"
  short_pct = "N/D"
  try:
    info = ticker_obj.info
    if info:
      nome_azienda = info.get("shortName") or info.get("longName") or ""
      if info.get("forwardPE"):
        fwd_pe = round(info.get("forwardPE"), 2)
      if info.get("pegRatio"):
        peg = round(info.get("pegRatio"), 2)
      if info.get("shortPercentOfFloat"):
        short_pct = round(info.get("shortPercentOfFloat") * 100, 2)
  except Exception:
    pass

  ticker_display = (
      f"{ticker_str} - {nome_azienda}" if nome_azienda else ticker_str
  )
  return ticker_display, fwd_pe, peg, short_pct


# =====================================================================
# MOTORE DI SUGGERIMENTO IBRIDO CON GESTIONE DRAWDOWN CORE
# =====================================================================
def genera_suggerimento_ibrido(c):
  is_core = c["is_core"]
  is_bear = c["is_bear"]
  cmf = c["cmf"]
  vsa = c["vsa_rating"]
  rsi = c["rsi"]
  storno = c["storno"]  # Storno percentuale dai massimi di 52 settimane
  pnl_pct = c["pnl_pct"]
  short_pct = c["short_pct"]

  # 1. GESTIONE SPECIFICA PER TITOLI CORE IN FORTE STORNO (Es. -30%, -40% o peggio)
  if is_core:
    if storno >= 35.0:
      # Se lo storno è profondo (>35%) verifichiamo se c'è capitolazione o supporto istituzionale
      if is_bear and cmf < -0.05:
        return (
            f"⚠️ [ALLARME CORE - ROTTURA STRUTTURALE (-{storno:.1f}%)]:"
            " Il titolo ha subito un drawdown severo ed è sotto la SMA200 con"
            " flussi in uscita. La tesi resta di lungo termine, ma il rischio"
            " di un bear market prolungato è alto: valuta di alleggerire una"
            " quota (es. 30-50%) per proteggere il capitale e rientrare più"
            " in basso."
        )
      else:
        return (
            f"🛡️ [CORE IN PROFONDO SCONTO (-{storno:.1f}%)]:"
            " Drawdown importante ma le mani forti non stanno fuggendo"
            " in modo disordinato. Non svendere la posizione strategica:"
            " sfrutta la debolezza per accumulare gradualmente (DCA)."
        )
    elif storno >= 20.0:
      return (
          f"💎 [CORE IN CORREZIONE (-{storno:.1f}%)]:"
          " Storno fisiologico per un asset di crescita. Mantieni la posizione"
          " core salda, la tesi di fondo non è compromessa."
      )

  # 2. Segnale Short Squeeze Esplosivo (Valido per entrambi se ci sono i presupposti)
  if (
      isinstance(short_pct, (int, float))
      and short_pct > 10
      and cmf > 0.08
      and not is_bear
  ):
    return (
        "🔥 [SHORT SQUEEZE IN ATTO]: Short interest alto + mani forti in forte"
        " accumulo. Lascia correre aggressivo!"
    )

  # 3. SE IL TITOLO "SCOTTA" SUI MASSIMI (Ipercomprato + Distribuzione)
  titolo_scotta = (rsi > 78 or (storno < 3.0 and rsi > 72)) and (
      cmf < -0.02 or vsa == "🔴 DISTRIBUZIONE / VENDITA"
  )

  if titolo_scotta:
    if is_core:
      return (
          f"🛡️ [CORE - ZONA CALDA]: Il titolo scotta (RSI {rsi:.1f}) sui"
          " massimi. Essendo un pilastro strategico, **NON VENDERE LA CORE**:"
          " ignora il rumore di breve e attendi lo storno per incrementare."
      )
    else:
      return (
          f"💰 [ZONA CALDA - PRENDI PROFITTO]: Titolo tirato (RSI {rsi:.1f}) con"
          " distribuzione in corso. Alleggerisci la quota tattica."
      )

  # 4. Gain straordinario senza distribuzione
  if pnl_pct != "N/D" and pnl_pct > 50.0 and not is_bear:
    return (
        f"🚀 [GAIN STRAORDINARIO (+{pnl_pct:.1f}%)]: Trend solido e tesi"
        " intatta. Fai correre i profitti."
    )

  # 5. Trend rialzista sano e accumulo pulito
  if not is_bear and cmf > 0.05 and vsa == "🟢 ACCUMULAZIONE PULITA":
    return (
        "🟢 [ACCUMULO ATTIVO]: Trend e flussi istituzionali sani. Continua ad"
        " accumulare."
    )

  # 6. Storno sano con supporto delle mani forti
  elif cmf >= 0.0 and storno >= 12.0:
    return (
        "💎 [SCONTO STRATEGICO]: Storno salutare in corso con assorbimento"
        " istituzionale. Ottima zona d'acquisto."
    )

  # 7. Distribuzione generale / Trend ribassista
  elif vsa == "🔴 DISTRIBUZIONE / VENDITA" or cmf < -0.05:
    if is_core:
      return (
          "⚠️ [FASE DIFENSIVA CORE]: Pressione ribassista. Mantieni la barra"
          " dritta sul lungo termine monitorando i supporti."
      )
    else:
      return (
          "🔴 [DISTRIBUZIONE]: Flussi negativi. Evita di mediare al ribasso su"
          " questo asset tattico."
      )

  # 8. Default / Neutro
  else:
    if rsi < 30:
      return (
          "🟡 [AREA IPERVENDUTO]: Titolo molto scarico (RSI < 30). Monitorare"
          " per rimbalzo."
      )
    return (
        "🟡 [FASE NEUTRA]: Struttura laterale. Mantenere la posizione senza"
        " fretta."
    )


# =====================================================================
# MAIN FUNCTION
# =====================================================================
def main():
  tickers = [
      t.strip().upper() for t in MEI_PORTAFOGLIO_CONFIG.keys() if t.strip()
  ]
  if not tickers:
    print("⚠️ Nessun ticker inserito nella configurazione del portafoglio.")
    return

  print(f"💱 Recupero tasso di cambio EUR/USD in corso...")
  eur_usd_rate = ottieni_tasso_cambio_eur_usd()
  print(f"ℹ️ Tasso di cambio utilizzato (EUR/USD): {eur_usd_rate:.4f}")

  print(
      f"🚀 Avvio scansione del Portafoglio Ibrido su {len(tickers)} titoli..."
  )
  try:
    df_raw = yf.download(
        tickers, period="1y", auto_adjust=True, progress=False
    )
  except Exception as e:
    print(f"Errore critico durante il download dati bulk: {e}")
    invia_telegram(
        CANALE_ACCUMULAZIONE_ID,
        "❌ Errore critico nel download dei dati del portafoglio.",
    )
    return

  candidati = []
  for ticker_str in tickers:
    try:
      if len(tickers) == 1:
        df_close = df_raw["Close"]
        df_volume = df_raw["Volume"]
        df_high = df_raw["High"]
        df_low = df_raw["Low"]
      else:
        df_close = (
            df_raw["Close"][ticker_str]
            if "Close" in df_raw
            else df_raw.xs(ticker_str, axis=1, level=1)["Close"]
        )
        df_volume = (
            df_raw["Volume"][ticker_str]
            if "Volume" in df_raw
            else df_raw.xs(ticker_str, axis=1, level=1)["Volume"]
        )
        df_high = (
            df_raw["High"][ticker_str]
            if "High" in df_raw
            else df_raw.xs(ticker_str, axis=1, level=1)["High"]
        )
        df_low = (
            df_raw["Low"][ticker_str]
            if "Low" in df_raw
            else df_raw.xs(ticker_str, axis=1, level=1)["Low"]
        )

      chiusure = df_close.dropna()
      if len(chiusure) < 50:
        continue
      volumi = df_volume.dropna()
      massimi = df_high.dropna()
      minimi = df_low.dropna()

      prezzo_attuale = float(chiusure.iloc[-1])
      massimo_52w = float(chiusure.max())
      minimo_52w = float(chiusure.min())
      storno_pct = ((massimo_52w - prezzo_attuale) / massimo_52w) * 100
      dist_min_52w_pct = (
          (prezzo_attuale - minimo_52w) / minimo_52w
      ) * 100

      config_titolo = MEI_PORTAFOGLIO_CONFIG.get(ticker_str, {})
      pmc_eur = config_titolo.get("pmc", 0.0)
      is_core = config_titolo.get("core", False)

      if pmc_eur > 0:
        if ticker_str.endswith(".MI") or ticker_str.endswith(".PA"):
          pmc_usd = pmc_eur
        else:
          pmc_usd = pmc_eur * eur_usd_rate
        pnl_pct = round(
            ((prezzo_attuale - pmc_usd) / pmc_usd) * 100, 2
        )
      else:
        pmc_usd = 0.0
        pnl_pct = "N/D"

      rvol_5d_pct = 0.0
      if len(volumi) >= 60:
        media_vol_5g = volumi.iloc[-5:].mean()
        media_vol_60g = volumi.iloc[-60:].mean()
        if media_vol_60g > 0:
          rvol_5d_pct = (media_vol_5g / media_vol_60g) * 100

      t_obj = yf.Ticker(ticker_str)
      ticker_display, fwd_pe, peg, short_pct = (
          ottieni_dati_fondamentali_e_anagrafica(t_obj, ticker_str)
      )

      minimo_60g = (
          float(chiusure.iloc[-60:].min())
          if len(chiusure) >= 60
          else float(chiusure.min())
      )
      distanza_supporto_pct = (
          (prezzo_attuale - minimo_60g) / minimo_60g
      ) * 100

      sma_50 = (
          float(chiusure.rolling(window=50).mean().iloc[-1])
          if len(chiusure) >= 50
          else prezzo_attuale
      )
      dist_sma50_pct = round(
          ((prezzo_attuale - sma_50) / sma_50) * 100 if sma_50 > 0 else 0.0, 2
      )

      sma_200 = (
          float(chiusure.rolling(window=200).mean().iloc[-1])
          if len(chiusure) >= 200
          else prezzo_attuale
      )
      dist_sma200_pct = round(
          (
              ((prezzo_attuale - sma_200) / sma_200) * 100
              if sma_200 > 0
              else 0.0
          ),
          2,
      )
      is_bear_market = prezzo_attuale < sma_200

      rsi_serie = calcola_rsi(chiusure)
      rsi_attuale = float(rsi_serie.iloc[-1])

      cmf_val = 0.0
      obv_trend = "Neutro"
      clv_val = 0.5
      poc_val = prezzo_attuale
      if len(volumi) >= 20:
        cmf_serie = calcola_cmf(massimi, minimi, chiusure, volumi, periodi=20)
        cmf_val = float(cmf_serie.iloc[-1])
        obv_serie = calcola_obv(chiusure, volumi)
        obv_sma = obv_serie.rolling(20).mean()
        if float(obv_serie.iloc[-1]) > float(obv_sma.iloc[-1]):
          obv_trend = "Rialzista (Accumulo)"
        else:
          obv_trend = "Ribassista (Distribuzione)"
        clv_5d = [
            calcola_close_location_value(
                chiusure.iloc[i], minimi.iloc[i], massimi.iloc[i]
            )
            for i in range(-5, 0)
        ]
        clv_val = float(np.mean(clv_5d))
        poc_val = calcola_volume_poc(
            chiusure, volumi, periodi=min(60, len(chiusure))
        )

      if cmf_val > 0.05 and clv_val >= 0.55:
        vsa_rating = "🟢 ACCUMULAZIONE PULITA"
      elif cmf_val < -0.05 and clv_val <= 0.45:
        vsa_rating = "🔴 DISTRIBUZIONE / VENDITA"
      else:
        vsa_rating = "🟡 NEUTRO / VOLATILITÀ"

      diz_candidato = {
          "ticker_raw": ticker_str,
          "ticker_display": ticker_display,
          "prezzo": prezzo_attuale,
          "pmc_eur": pmc_eur,
          "is_core": is_core,
          "pnl_pct": pnl_pct,
          "rsi": rsi_attuale,
          "storno": storno_pct,
          "is_bear": is_bear_market,
          "rvol_5d": rvol_5d_pct,
          "supporto_60g": minimo_60g,
          "dist_supp_pct": distanza_supporto_pct,
          "resistenza_52w": massimo_52w,
          "minimo_52w": minimo_52w,
          "dist_min_52w_pct": dist_min_52w_pct,
          "sma_50": sma_50,
          "dist_sma50_pct": dist_sma50_pct,
          "sma_200": sma_200,
          "dist_sma200_pct": dist_sma200_pct,
          "cmf": cmf_val,
          "obv_trend": obv_trend,
          "clv": clv_val,
          "poc_60g": poc_val,
          "vsa_rating": vsa_rating,
          "forward_pe": fwd_pe,
          "peg_ratio": peg,
          "short_interest": short_pct,
      }

      diz_candidato["suggerimento"] = genera_suggerimento_ibrido(diz_candidato)
      candidati.append(diz_candidato)

    except Exception as e:
      print(f"⚠️ Errore durante l'elaborazione del ticker {ticker_str}: {e}")
      continue

  if not candidati:
    print("ℹ️ Nessun dato elaborato per i ticker selezionati.")
    return

  # GENERAZIONE FILE EXCEL REPORT PORTAFOGLIO
  data_odierna = datetime.now().strftime("%Y-%m-%d")
  excel_filename = f"Report_Portafoglio_Ibrido_{data_odierna}.xlsx"
  excel_data = []

  for c in candidati:
    stato_trend = (
        "🔴 BEAR TREND (Sotto SMA200)"
        if c["is_bear"]
        else "🟢 BULL TREND (Sopra SMA200)"
    )
    tipo_asset = "⭐ CORE (Lungo Termine)" * c["is_core"] or "💼 Tattico"
    condizione_rsi = (
        "Ipervenduto (<30)" if c["rsi"] < 30 else "Neutro/Normale"
    )
    pnl_display = (
        f"{c['pnl_pct']:+.1f}%"
        if isinstance(c["pnl_pct"], (int, float))
        else "N/D"
    )

    excel_data.append({
        "Ticker": c["ticker_display"],
        "Profilo Asset": tipo_asset,
        "Trend Market": stato_trend,
        "Prezzo Attuale ($)": round(c["prezzo"], 2),
        "PMC Inserito (€)": c["pmc_eur"] if c["pmc_eur"] > 0 else "N/D",
        "Performance P&L (%)": pnl_display,
        "Forward P/E": c["forward_pe"],
        "Short Interest (%)": c["short_interest"],
        "SMA 50 ($)": round(c["sma_50"], 2),
        "Dist. SMA50 (%)": round(c["dist_sma50_pct"] / 100, 4),
        "SMA 200 ($)": round(c["sma_200"], 2),
        "Dist. SMA200 (%)": round(c["dist_sma200_pct"] / 100, 4),
        "Supporto 60G ($)": round(c["supporto_60g"], 2),
        "Dist. Supp. (%)": round(c["dist_supp_pct"] / 100, 4),
        "Storno Max 52W (%)": round(c["storno"] / 100, 4),
        "RSI (14)": round(c["rsi"], 1),
        "Stato RSI": condizione_rsi,
        "CMF (20G)": round(c["cmf"], 3),
        "Analisi VSA": c["vsa_rating"],
        "Suggerimento / Action": c["suggerimento"],
    })

  df_excel = pd.DataFrame(excel_data)
  try:
    with pd.ExcelWriter(excel_filename, engine="openpyxl") as writer:
      df_excel.to_excel(writer, sheet_name="Portafoglio Ibrido", index=False)
    print(f"📊 File Excel generato con successo: {excel_filename}")
  except Exception as e:
    print(f"❌ Errore durante la creazione del file Excel: {e}")

  # NOTIFICHE TELEGRAM & EMAIL
  corpo_email_testo = (
      f"Smart Money Radar - Report Portafoglio Ibrido del {data_odierna}\n"
      f"Tasso cambio EUR/USD applicato: {eur_usd_rate:.4f}\n\n"
  )
  righe = [
      f"📊 **MONITOR PORTAFOGLIO IBRIDO ({data_odierna})** *(Cambio:"
      f" {eur_usd_rate:.4f})*\n"
  ]
  for c in candidati:
    trend_icon = "🔴" if c["is_bear"] else "🟢"
    core_label = "⭐" if c["is_core"] else ""
    righe.append(
        f"• {trend_icon}{core_label} **{c['ticker_raw']}** (${c['prezzo']:.1f})"
        f" ➔ *{c['suggerimento']}*"
    )

  msg = "\n".join(righe)
  invia_telegram(CANALE_ACCUMULAZIONE_ID, msg)
  corpo_email_testo += msg.replace("**", "").replace("*", "")
  corpo_email_testo += (
      "\n\nTrovi in allegato il report in formato Excel con la suddivisione tra"
      " asset Core e Tattici."
  )
  invia_email_con_allegato(
      oggetto=f"📈 Report Portafoglio Ibrido - {data_odierna}",
      corpo_testo=corpo_email_testo,
      file_excel_path=excel_filename,
  )


if __name__ == "__main__":
  main()
