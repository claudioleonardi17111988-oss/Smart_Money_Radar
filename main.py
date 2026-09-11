from datetime import datetime
import email
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import StringIO
import os
import smtplib
import numpy as np
import pandas as pd
import requests
import yfinance as yf

# =====================================================================
# CONFIGURAZIONE TELEGRAM, EMAIL & PARAMETRI RADAR
# =====================================================================
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
CANALE_ACCUMULAZIONE_ID = os.environ.get(
    "CANALE_ACCUMULAZIONE_ID", "-1003454283658"
)
SENDER_EMAIL = os.environ.get("SENDER_EMAIL")
APP_PASSWORD = os.environ.get("APP_PASSWORD")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL", SENDER_EMAIL)
SOGLIA_STORNO_MINIMA = 15.0  # Storno dai max a 52W >= 15%
SOGLIA_VOLUMI_SETTIMANA = 115.0  # Volumi 5 giorni >= 115% della media 60g


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
    print(
        "⚠️ Credenziali E-mail non configurate (SENDER_EMAIL / APP_PASSWORD)."
        " Saltato invio mail."
    )
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


# =====================================================================
# CALCOLI ANALISI TECNICA AVANZATA
# =====================================================================
def calcola_rsi(chiusure, periodi=14):
  """Calcola l'indicatore RSI standard a 14 periodi."""
  delta = chiusure.diff()
  guadagno = (delta.where(delta > 0, 0)).rolling(window=periodi).mean()
  perdita = (-delta.where(delta < 0, 0)).rolling(window=periodi).mean()
  rs = guadagno / perdita
  return 100 - (100 / (1 + rs))


def calcola_obv(chiusure, volumi):
  """Calcola l'On-Balance Volume (OBV)."""
  obv = (np.sign(chiusure.diff()) * volumi).fillna(0).cumsum()
  return obv


def calcola_cmf(massimi, minimi, chiusure, volumi, periodi=20):
  """Calcola il Chaikin Money Flow (CMF) a 20 periodi."""
  mf_multiplier = (
      (chiusure - minimi) - (massimi - chiusure)
  ) / (massimi - minimi)
  mf_multiplier = mf_multiplier.fillna(0)
  mf_volume = mf_multiplier * volumi
  cmf = mf_volume.rolling(window=periodi).sum() / volumi.rolling(
      window=periodi
  ).sum()
  return cmf


def calcola_close_location_value(chiusura, minimo, massimo):
  """Misura dove ha chiuso il prezzo rispetto al range giornaliero (0.0 = sui minimi, 1.0 = sui massimi)."""
  range_giornaliero = massimo - minimo
  if range_giornaliero == 0:
    return 0.5
  return (chiusura - minimo) / range_giornaliero


def calcola_volume_poc(chiusure, volumi, periodi=60, bins=10):
  """Calcola approssimativamente il Point of Control (POC) dei volumi negli ultimi N giorni."""
  if len(chiusure) < periodi:
    return float(chiusure.iloc[-1])
  sub_close = chiusure.iloc[-periodi:]
  sub_vol = volumi.iloc[-periodi:]
  counts, bin_edges = np.histogram(sub_close, bins=bins, weights=sub_vol)
  max_idx = np.argmax(counts)
  poc_price = (bin_edges[max_idx] + bin_edges[max_idx + 1]) / 2.0
  return float(poc_price)


def _fetch_wikipedia_table(url):
  headers = {
      "User-Agent": (
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
          " like Gecko) Chrome/120.0.0.0 Safari/537.36"
      )
  }
  response = requests.get(url, headers=headers)
  response.raise_for_status()
  return pd.read_html(StringIO(response.text))


def ottieni_ticker_usa():
  """Scarica e unisce i ticker di S&P 500, Nasdaq 100 e S&P MidCap 400 senza duplicati."""
  tickers = set()
  # 1. S&P 500
  try:
    url_sp = "https://raw.githubusercontent.com/datasets/s-and-p-500-companies/main/data/constituents.csv"
    df_sp = pd.read_csv(url_sp)
    tickers.update(
        df_sp["Symbol"].str.replace(".", "-", regex=False).str.strip().tolist()
    )
    print("✅ S&P 500 caricato con successo.")
  except Exception as e:
    print(f"⚠️ Errore caricamento S&P 500: {e}")

  # 2. NASDAQ 100
  try:
    url_nasdaq = "https://en.wikipedia.org/wiki/Nasdaq-100"
    tables = _fetch_wikipedia_table(url_nasdaq)
    for df in tables:
      col = next(
          (
              c
              for c in df.columns
              if str(c).lower() in ["ticker", "symbol", "company stock symbol"]
          ),
          None,
      )
      if col:
        raw_nasdaq = (
            df[col]
            .dropna()
            .astype(str)
            .str.replace(".", "-", regex=False)
            .str.strip()
            .tolist()
        )
        tickers.update(raw_nasdaq)
        print("✅ Nasdaq 100 caricato con successo.")
        break
  except Exception as e:
    print(f"⚠️ Errore caricamento Nasdaq 100: {e}")

  # 3. S&P MidCap 400
  try:
    url_midcap = "https://en.wikipedia.org/wiki/List_of_S%26P_400_companies"
    tables = _fetch_wikipedia_table(url_midcap)
    for df in tables:
      col = next(
          (
              c
              for c in df.columns
              if str(c).lower() in ["symbol", "ticker", "company stock symbol"]
          ),
          None,
      )
      if col:
        raw_midcap = (
            df[col]
            .dropna()
            .astype(str)
            .str.replace(".", "-", regex=False)
            .str.strip()
            .tolist()
        )
        tickers.update(raw_midcap)
        print("✅ S&P MidCap 400 caricato con successo.")
        break
  except Exception as e:
    print(f"⚠️ Errore caricamento S&P MidCap 400: {e}")

  lista_finale = list(tickers)
  print(f"🎯 Totale titoli unici da analizzare: {len(lista_finale)}")
  return lista_finale


def ottieni_dati_azienda(ticker_obj, ticker_str):
  """Verifica la salute di bilancio dell'azienda e recupera il nome esteso."""
  nome_azienda = ""
  is_sano = True
  try:
    info = ticker_obj.info
    if info:
      nome_azienda = info.get("shortName") or info.get("longName") or ""
      debt_to_equity = info.get("debtToEquity", None)
      if debt_to_equity is not None and debt_to_equity > 250:
        is_sano = False
      earnings_growth = info.get("earningsGrowth", None)
      revenue_growth = info.get("revenueGrowth", None)
      if earnings_growth is not None and earnings_growth < -0.20:
        is_sano = False
      if revenue_growth is not None and revenue_growth < -0.20:
        is_sano = False
  except Exception as e:
    print(f"⚠️ Impossibile verificare info complete per {ticker_str}: {e}")
    is_sano = True
  ticker_display = (
      f"{ticker_str} - {nome_azienda}" if nome_azienda else ticker_str
  )
  return is_sano, ticker_display


# =====================================================================
# MAIN FUNCTION
# =====================================================================
def main():
  tickers = ottieni_ticker_usa()
  print(f"🚀 Avvio Smart Money Radar su {len(tickers)} titoli USA...")
  try:
    df_raw = yf.download(
        tickers, period="1y", auto_adjust=True, progress=False
    )
  except Exception as e:
    print(f"Errore critico durante il download dati bulk: {e}")
    invia_telegram(
        CANALE_ACCUMULAZIONE_ID,
        "❌ Errore critico nel download dei dati di borsa.",
    )
    return

  # Estrazione DataFrame multi-index
  if isinstance(df_raw.columns, pd.MultiIndex):
    df_close = (
        df_raw["Close"]
        if "Close" in df_raw.columns.levels[0]
        else df_raw.xs("Close", axis=1, level=1)
    )
    df_volume = (
        df_raw["Volume"]
        if "Volume" in df_raw.columns.levels[0]
        else df_raw.xs("Volume", axis=1, level=1)
    )
    df_high = (
        df_raw["High"]
        if "High" in df_raw.columns.levels[0]
        else df_raw.xs("High", axis=1, level=1)
    )
    df_low = (
        df_raw["Low"]
        if "Low" in df_raw.columns.levels[0]
        else df_raw.xs("Low", axis=1, level=1)
    )
  else:
    df_close = df_raw
    df_volume, df_high, df_low = None, None, None

  candidati = []
  for ticker in df_close.columns:
    try:
      ticker_str = str(ticker)
      chiusure = df_close[ticker].dropna()
      if len(chiusure) < 200:
        continue

      volumi = df_volume[ticker].dropna() if df_volume is not None else None
      massimi = df_high[ticker].dropna() if df_high is not None else None
      minimi = df_low[ticker].dropna() if df_low is not None else None

      prezzo_attuale = float(chiusure.iloc[-1])
      massimo_52w = float(chiusure.max())
      minimo_52w = float(chiusure.min())
      storno_pct = ((massimo_52w - prezzo_attuale) / massimo_52w) * 100
      dist_min_52w_pct = (
          (prezzo_attuale - minimo_52w) / minimo_52w
      ) * 100

      # FILTRO 1: Storno minimo del 15% dai max 52W
      if storno_pct < SOGLIA_STORNO_MINIMA:
        continue

      # CALCOLO MEDIA VOLUMI 1 SETTIMANA (5 GIORNI) vs MEDIA 60 GIORNI
      rvol_5d_pct = 0.0
      if volumi is not None and len(volumi) >= 60:
        media_vol_5g = volumi.iloc[-5:].mean()
        media_vol_60g = volumi.iloc[-60:].mean()
        if media_vol_60g > 0:
          rvol_5d_pct = (media_vol_5g / media_vol_60g) * 100

      # FILTRO 2: Volumi della settimana almeno al 115%
      if rvol_5d_pct < SOGLIA_VOLUMI_SETTIMANA:
        continue

      # FILTRO 3: Bilanci sani + Recupero nome esteso
      t_obj = yf.Ticker(ticker_str)
      is_sano, ticker_display = ottieni_dati_azienda(t_obj, ticker_str)
      if not is_sano:
        continue

      # CALCOLI ANALISI TECNICA AVANZATA
      minimo_60g = float(chiusure.iloc[-60:].min())
      distanza_supporto_pct = (
          (prezzo_attuale - minimo_60g) / minimo_60g
      ) * 100
      sma_200 = float(chiusure.rolling(window=200).mean().iloc[-1])
      distanza_sma200_pct = ((prezzo_attuale - sma_200) / sma_200) * 100
      is_bear_market = prezzo_attuale < sma_200

      rsi_serie = calcola_rsi(chiusure)
      rsi_attuale = float(rsi_serie.iloc[-1])

      # Nuovi Indicatori
      cmf_val = 0.0
      obv_trend = "Neutro"
      clv_val = 0.5
      poc_val = prezzo_attuale

      if (
          volumi is not None
          and massimi is not None
          and minimi is not None
          and len(volumi) >= 20
      ):
        # CMF (Chaikin Money Flow)
        cmf_serie = calcola_cmf(massimi, minimi, chiusure, volumi, periodi=20)
        cmf_val = float(cmf_serie.iloc[-1])

        # OBV Trend (confronto OBV attuale vs media 20 giorni)
        obv_serie = calcola_obv(chiusure, volumi)
        obv_sma = obv_serie.rolling(20).mean()
        if float(obv_serie.iloc[-1]) > float(obv_sma.iloc[-1]):
          obv_trend = "Rialzista (Accumulo)"
        else:
          obv_trend = "Ribassista (Distribuzione)"

        # CLV (Close Location Value) media ultimi 5 giorni
        clv_5d = [
            calcola_close_location_value(
                chiusure.iloc[i], minimi.iloc[i], massimi.iloc[i]
            )
            for i in range(-5, 0)
        ]
        clv_val = float(np.mean(clv_5d))

        # Volume POC (Point of Control 60 giorni)
        poc_val = calcola_volume_poc(chiusure, volumi, periodi=60)

      # DETERMINAZIONE PATTERN VSA (Volume Spread Analysis)
      if cmf_val > 0.05 and clv_val >= 0.55:
        vsa_rating = "🟢 ACCUMULAZIONE PULITA"
      elif cmf_val < -0.05 and clv_val <= 0.45:
        vsa_rating = "🔴 DISTRIBUZIONE / VENDITA"
      else:
        vsa_rating = "🟡 NEUTRO / VOLATILITÀ"

      candidati.append({
          "ticker_raw": ticker_str,
          "ticker_display": ticker_display,
          "prezzo": prezzo_attuale,
          "rsi": rsi_attuale,
          "storno": storno_pct,
          "is_bear": is_bear_market,
          "rvol_5d": rvol_5d_pct,
          "supporto_60g": minimo_60g,
          "dist_supp_pct": distanza_supporto_pct,
          "resistenza_52w": massimo_52w,
          "minimo_52w": minimo_52w,
          "dist_min_52w_pct": dist_min_52w_pct,
          "sma_200": sma_200,
          "dist_sma200_pct": distanza_sma200_pct,
          "cmf": cmf_val,
          "obv_trend": obv_trend,
          "clv": clv_val,
          "poc_60g": poc_val,
          "vsa_rating": vsa_rating,
      })
    except Exception:
      continue

  print(
      f"✅ Scansione completata. {len(candidati)} titoli in fase di accumulazione"
      " rilevati oggi!"
  )
  if not candidati:
    msg_vuoto = (
        "ℹ️ **Smart Money Radar**: Nessun titolo in storno > 15% presenta"
        " accumulazione di volumi (VOL 1W >= 115%) nella sessione odierna."
    )
    invia_telegram(CANALE_ACCUMULAZIONE_ID, msg_vuoto)
    return

  candidati_ordinati = sorted(
      candidati, key=lambda x: x["rvol_5d"], reverse=True
  )

  # GENERAZIONE FILE EXCEL CON DATI PER REPORT QUINDICINALE
  data_odierna = datetime.now().strftime("%Y-%m-%d")
  excel_filename = f"Report_Accumulazione_{data_odierna}.xlsx"
  excel_data = []

  for c in candidati_ordinati:
    stato_trend = (
        "🔴 BEAR TREND (Sotto SMA200)"
        if c["is_bear"]
        else "🟢 BULL TREND (Sopra SMA200)"
    )
    condizione_rsi = (
        "Ipervenduto (<30)" if c["rsi"] < 30 else "Neutro/Normale"
    )

    if c["vsa_rating"] == "🟢 ACCUMULAZIONE PULITA":
      valutazione = "🟢 Forte pressione in acquisto: Mani forti in accumulo."
    elif c["vsa_rating"] == "🔴 DISTRIBUZIONE / VENDITA":
      valutazione = "🔴 Attenzione: Elevati volumi in vendita (Distribuzione)."
    else:
      valutazione = "🟡 Frequente volatilità: attendere conferme di inversione."

    excel_data.append({
        "Ticker": c["ticker_display"],
        "Trend Market": stato_trend,
        "Prezzo Attuale ($)": round(c["prezzo"], 2),
        "SMA 200 ($)": round(c["sma_200"], 2),
        "Distanza da SMA200 (%)": round(c["dist_sma200_pct"] / 100, 4),
        "Supporto 60G ($)": round(c["supporto_60g"], 2),
        "Distanza da Supp. (%)": round(c["dist_supp_pct"] / 100, 4),
        "Resistenza 52W ($)": round(c["resistenza_52w"], 2),
        "Storno dai Max 52W (%)": round(c["storno"] / 100, 4),
        "Dist. Dai Min 52W (%)": round(c["dist_min_52w_pct"] / 100, 4),
        "Volumi 1W (% vs media 60g)": round(c["rvol_5d"] / 100, 4),
        "RSI (14)": round(c["rsi"], 1),
        "Stato RSI": condizione_rsi,
        "CMF (20G)": round(c["cmf"], 3),
        "OBV Trend": c["obv_trend"],
        "Close Location (0-1)": round(c["clv"], 2),
        "POC Volumi 60G ($)": round(c["poc_60g"], 2),
        "Analisi VSA (Volume Spread)": c["vsa_rating"],
        "Suggerimento / Action": valutazione,
    })

  df_excel = pd.DataFrame(excel_data)
  try:
    with pd.ExcelWriter(excel_filename, engine="openpyxl") as writer:
      df_excel.to_excel(writer, sheet_name="Accumulazione", index=False)
    print(f"📊 File Excel generato con successo: {excel_filename}")
  except Exception as e:
    print(f"❌ Errore durante la creazione del file Excel: {e}")

  # FORMATTAZIONE E INVIO TELEGRAM & EMAIL
  dips_bull_market = []
  bear_market_watchlist = []
  corpo_email_testo = (
      f"Smart Money Radar - Report Accumulazione del {data_odierna}\n\n"
  )

  for c in candidati_ordinati:
    info_vol = f"VOL 1W: {c['rvol_5d']:.0f}%"
    info_storno = f"-{c['storno']:.1f}% dai max"
    info_vsa = f"VSA: {c['vsa_rating']}"
    info_rsi = (
        f"**RSI: {c['rsi']:.0f} (Ipervenduto)**"
        if c["rsi"] < 30
        else f"RSI: {c['rsi']:.0f}"
    )

    if not c["is_bear"]:
      riga_str = (
          f"• 🟢 **{c['ticker_raw']}** (${c['prezzo']:.1f} | {info_vol} |"
          f" {info_storno} | {info_rsi} | {info_vsa})"
      )
      dips_bull_market.append(riga_str)
    else:
      riga_str = (
          f"• 🔴 **{c['ticker_raw']}** (${c['prezzo']:.1f} | {info_vol} |"
          f" {info_storno} | {info_rsi} | {info_vsa})"
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

  corpo_email_testo += "\n".join(righe).replace("**", "")
  corpo_email_testo += (
      "\n\nTrovi in allegato il report in formato Excel completo di tutte le"
      " metriche e suggerimenti operativi per l'analisi quindicinale."
  )

  invia_email_con_allegato(
      oggetto=f"📈 Smart Money Radar Report - {data_odierna}",
      corpo_testo=corpo_email_testo,
      file_excel_path=excel_filename,
  )


if __name__ == "__main__":
  main()
