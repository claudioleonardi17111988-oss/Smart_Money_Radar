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
# 🎯 INSERISCI QUI I TICKER DA MONITORARE NEL TUO PORTAFOGLIO
# ---------------------------------------------------------------------
MEI_TICKER_PORTAFOGLIO = [
    "APP",  # AppLovin
    "RDDT",  # Reddit
    "BBWI",  # Bath & Body Works
    "CEG", "MARA", "ON", "VRT", "WMT", "CLS", "FIX", "NFLX", "NVDA", "QCOM", "NOW", "TDG", "VST", "EPAM", "ROL", "NVO", "AMTM", "BMY", "FCT.MI", "SOFI", "NU", "ZENA", "ADUR", "XYL"
    # Aggiungi qui i tuoi nuovi ticker tra virgolette separati da virgola
    # es: "AAPL", "NVDA", "MSFT"
]


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
  delta = chiusure.diff()
  guadagno = (delta.where(delta > 0, 0)).rolling(window=periodi).mean()
  perdita = (-delta.where(delta < 0, 0)).rolling(window=periodi).mean()
  rs = guadagno / perdita
  return 100 - (100 / (1 + rs))


def calcola_obv(chiusure, volumi):
  obv = (np.sign(chiusure.diff()) * volumi).fillna(0).cumsum()
  return obv


def calcola_cmf(massimi, minimi, chiusure, volumi, periodi=20):
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


def ottieni_dati_azienda(ticker_obj, ticker_str):
  nome_azienda = ""
  try:
    info = ticker_obj.info
    if info:
      nome_azienda = info.get("shortName") or info.get("longName") or ""
  except Exception:
    pass
  ticker_display = (
      f"{ticker_str} - {nome_azienda}" if nome_azienda else ticker_str
  )
  return ticker_display


# =====================================================================
# MAIN FUNCTION
# =====================================================================
def main():
  tickers = [t.strip().upper() for t in MEI_TICKER_PORTAFOGLIO if t.strip()]

  if not tickers:
    print("⚠️ Nessun ticker inserito nella lista MEI_TICKER_PORTAFOGLIO.")
    return

  print(
      f"🚀 Avvio scansione del Portafoglio Personale su {len(tickers)} titoli:"
      f" {tickers}..."
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

  # Gestione estrazione dati sia per singolo ticker che per multipli
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

      # Volumi della settimana (5 giorni) vs media 60 giorni
      rvol_5d_pct = 0.0
      if len(volumi) >= 60:
        media_vol_5g = volumi.iloc[-5:].mean()
        media_vol_60g = volumi.iloc[-60:].mean()
        if media_vol_60g > 0:
          rvol_5d_pct = (media_vol_5g / media_vol_60g) * 100

      t_obj = yf.Ticker(ticker_str)
      ticker_display = ottieni_dati_azienda(t_obj, ticker_str)

      # Calcoli tecnici
      minimo_60g = (
          float(chiusure.iloc[-60:].min())
          if len(chiusure) >= 60
          else float(chiusure.min())
      )
      distanza_supporto_pct = (
          (prezzo_attuale - minimo_60g) / minimo_60g
      ) * 100

      sma_200 = (
          float(chiusure.rolling(window=200).mean().iloc[-1])
          if len(chiusure) >= 200
          else prezzo_attuale
      )
      distanza_sma200_pct = (
          ((prezzo_attuale - sma_200) / sma_200) * 100 if sma_200 > 0 else 0.0
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

      # VSA Rating
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
    except Exception as e:
      print(f"⚠️ Errore durante l'elaborazione del ticker {ticker_str}: {e}")
      continue

  if not candidati:
    print("ℹ️ Nessun dato elaborato per i ticker selezionati.")
    return

  # GENERAZIONE FILE EXCEL REPORT PORTAFOGLIO
  data_odierna = datetime.now().strftime("%Y-%m-%d")
  excel_filename = f"Report_Portafoglio_{data_odierna}.xlsx"
  excel_data = []

  for c in candidati:
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
      df_excel.to_excel(writer, sheet_name="Portafoglio", index=False)
    print(f"📊 File Excel generato con successo: {excel_filename}")
  except Exception as e:
    print(f"❌ Errore durante la creazione del file Excel: {e}")

  # NOTIFICHE TELEGRAM & EMAIL
  corpo_email_testo = (
      f"Smart Money Radar - Report Portafoglio del {data_odierna}\n\n"
  )
  righe = [f"📊 **MONITOR PORTAFOGLIO PERSONALE ({data_odierna})**\n"]

  for c in candidati:
    info_vol = f"VOL 1W: {c['rvol_5d']:.0f}%"
    info_storno = f"-{c['storno']:.1f}% dai max"
    info_vsa = f"VSA: {c['vsa_rating']}"
    info_rsi = (
        f"**RSI: {c['rsi']:.0f} (Ipervenduto)**"
        if c["rsi"] < 30
        else f"RSI: {c['rsi']:.0f}"
    )
    trend_icon = "🔴" if c["is_bear"] else "🟢"

    righe.append(
        f"• {trend_icon} **{c['ticker_raw']}** (${c['prezzo']:.1f} | {info_vol} |"
        f" {info_storno} | {info_rsi} | {info_vsa})"
    )

  msg = "\n".join(righe)
  invia_telegram(CANALE_ACCUMULAZIONE_ID, msg)

  corpo_email_testo += msg.replace("**", "")
  corpo_email_testo += (
      "\n\nTrovi in allegato il report in formato Excel con tutti i dettagli"
      " tecnici aggiornati dei tuoi titoli."
  )

  invia_email_con_allegato(
      oggetto=f"📈 Report Portafoglio Personale - {data_odierna}",
      corpo_testo=corpo_email_testo,
      file_excel_path=excel_filename,
  )


if __name__ == "__main__":
  main()
