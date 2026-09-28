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
# CONFIGURAZIONE TELEGRAM, EMAIL & PORTAFOGLI FAMIGLIA
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
# - 'proprietario': 'Me', 'Gabri' o 'Greg' (permette il filtro in Excel)
# - 'core': True/False (utilizzato solo se proprietario == 'Me')
# ---------------------------------------------------------------------
MEI_PORTAFOGLIO_CONFIG = {
    # I tuoi titoli personali (con distinzione Core / Tattico)
    "APP": {"pmc": 263.53, "proprietario": "Greg", "core": False},
    "RDDT": {"pmc": 131.10, "proprietario": "Gabri", "core": False},
    "BBWI": {"pmc": 16.34, "proprietario": "Me", "core": False},
    "CEG": {"pmc": 236.10, "proprietario": "Gabri", "core": True},
    "MARA": {"pmc": 10.53, "proprietario": "Gabri", "core": True},
    "ON": {"pmc": 70.87, "proprietario": "Gabri", "core": False},
    "VRT": {"pmc": 250.45, "proprietario": "Gabri", "core": True},
    "WMT": {"pmc": 102.48, "proprietario": "Gabri", "core": False},
    "MU": {"pmc": 812.36, "proprietario": "Gabri", "core": True},
    "WULF": {"pmc": 19.35, "proprietario": "Gabri", "core": False},
    "CLS": {"pmc": 296.91, "proprietario": "Greg", "core": True},
    "HIVE.TO": {"pmc": 2.50, "proprietario": "Greg", "core": True},
    "COHR": {"pmc": 261.22, "proprietario": "Greg", "core": True},
    "CIFR": {"pmc": 21.88, "proprietario": "Greg", "core": False},
    "AZO": {"pmc": 2525.50, "proprietario": "Greg", "core": False},
    "FIX": {"pmc": 1495.90, "proprietario": "Greg", "core": True},
    "NFLX": {"pmc": 65.31, "proprietario": "Greg", "core": True},
    "QCOM": {"pmc": 146.84, "proprietario": "Greg", "core": False},
    "TDG": {"pmc": 1071.77, "proprietario": "Greg", "core": True},
    "VST": {"pmc": 132.48, "proprietario": "Greg", "core": True},
    "ROL": {"pmc": 31.0, "proprietario": "Me", "core": False},
    "NVO": {"pmc": 37.58, "proprietario": "Me", "core": True},
    "AMTM": {"pmc": 19.73, "proprietario": "Me", "core": True},
    "BMY": {"pmc": 50.60, "proprietario": "Me", "core": True},
    "FCT.MI": {"pmc": 15.71, "proprietario": "Me", "core": True},
    "SOFI": {"pmc": 14.87, "proprietario": "Me", "core": True},
    "NU": {"pmc": 11.98, "proprietario": "Me", "core": True},
    "ZENA": {"pmc": 2.03, "proprietario": "Me", "core": True},
    "ADUR": {"pmc": 12.91, "proprietario": "Me", "core": True},
    "XYL": {"pmc": 97.73, "proprietario": "Me", "core": True},
    "HEI": {"pmc": 195.00, "proprietario": "Me", "core": True},
    "SKHY": {"pmc": 168.00, "proprietario": "Me", "core": True},
    "UPST": {"pmc": 23.99, "proprietario": "Me", "core": True},
    "HTZ": {"pmc": 2.83, "proprietario": "Me", "core": True},
    
    # Titoli di Gabri
    "POET": {"pmc": 7.24, "proprietario": "Gabri"},
    
    # Titoli di Greg
    "NVDA": {"pmc": 184.62, "proprietario": "Greg"},
}

# =====================================================================
# FUNZIONI DI NOTIFICA & INVIO MAIL
# =====================================================================
def invia_telegram(canale_id, messaggio):
  if not TELEGRAM_TOKEN or not canale_id:
    print("⚠️ Telegram non configurato. Stampa a video:")
    print(messaggio)
    return
  url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
  payload = {"chat_id": canale_id, "text": messaggio, "parse_mode": "Markdown"}
  try:
    requests.post(url, json=payload, timeout=10)
  except Exception as e:
    print(f"Errore invio Telegram: {e}")

def invia_email_con_allegato(oggetto, corpo_testo, file_excel_path):
  if not SENDER_EMAIL or not APP_PASSWORD or not RECEIVER_EMAIL:
    return
  msg = MIMEMultipart()
  msg["From"] = SENDER_EMAIL
  msg["To"] = RECEIVER_EMAIL
  msg["Subject"] = oggetto
  msg.attach(MIMEText(corpo_testo, "plain", "utf-8"))
  if os.path.exists(file_excel_path):
    with open(file_excel_path, "rb") as f:
      part = MIMEApplication(f.read(), Name=os.path.basename(file_excel_path))
      part["Content-Disposition"] = f'attachment; filename="{os.path.basename(file_excel_path)}"'
      msg.attach(part)
  try:
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(SENDER_EMAIL, APP_PASSWORD)
    server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())
    server.quit()
  except Exception as e:
    print(f"❌ Errore invio e-mail: {e}")

def ottieni_tasso_cambio_eur_usd():
  try:
    fx = yf.Ticker("EURUSD=X")
    hist = fx.history(period="1d")
    if not hist.empty:
      return float(hist["Close"].iloc[-1])
  except Exception:
    pass
  return 1.08

# =====================================================================
# ANALISI TECNICA & VOLUMI
# =====================================================================
def calcola_rsi(chiusure, periodi=14):
  delta = chiusure.diff()
  guadagno = (delta.where(delta > 0, 0)).rolling(window=periodi).mean()
  perdita = (-delta.where(delta < 0, 0)).rolling(window=periodi).mean()
  rs = guadagno / perdita
  return 100 - (100 / (1 + rs))

def calcola_cmf(massimi, minimi, chiusure, volumi, periodi=20):
  mf_multiplier = ((chiusure - minimi) - (massimi - chiusure)) / (massimi - minimi)
  mf_multiplier = mf_multiplier.fillna(0)
  mf_volume = mf_multiplier * volumi
  return mf_volume.rolling(window=periodi).sum() / volumi.rolling(window=periodi).sum()

def analizza_trend_volumi_5g(volumi_ser):
  if len(volumi_ser) < 5:
    return "Stabile ➡️"
  v5 = volumi_ser.iloc[-5:].values
  t_recenti = v5[-2:].mean()
  t_passati = v5[:3].mean()
  if t_passati == 0:
    return "Stabile ➡️"
  ratio = t_recenti / t_passati
  if ratio > 1.10:
    return "In Accelerazione 📈"
  elif ratio < 0.90:
    return "In Raffreddamento 📉"
  else:
    return "Stabile ➡️"

def calcola_pendenza_regressione_volumi(volumi_ser, periodi):
  if len(volumi_ser) < periodi:
    return "N/D"
  y = volumi_ser.iloc[-periodi:].values
  x = np.arange(periodi)
  slope, _ = np.polyfit(x, y, 1)
  media_y = np.mean(y)
  if media_y == 0:
    return "Stabile ➡️"
  slope_pct = (slope / media_y) * 100
  if slope_pct > 0.5:
    return "In Crescita 📈"
  elif slope_pct < -0.5:
    return "In Esaurimento 📉"
  else:
    return "Stabile ➡️"

def ottieni_dati_fondamentali_e_anagrafica(ticker_obj, ticker_str):
  nome_azienda = ""
  fwd_pe = "N/D"
  peg = "N/D"
  short_pct = "N/D"
  try:
    info = ticker_obj.info
    if info:
      nome_azienda = info.get("shortName") or info.get("longName") or ""
      if info.get("forwardPE"): fwd_pe = round(info.get("forwardPE"), 2)
      if info.get("pegRatio"): peg = round(info.get("pegRatio"), 2)
      if info.get("shortPercentOfFloat"): short_pct = round(info.get("shortPercentOfFloat") * 100, 2)
  except Exception:
    pass
  return (f"{ticker_str} - {nome_azienda}" if nome_azienda else ticker_str), fwd_pe, peg, short_pct

# =====================================================================
# MOTORE DI SUGGERIMENTO IBRIDO (CON DCA FINALIZZATO AL PROFITTO)
# =====================================================================
def genera_suggerimento_ibrido(c):
  proprietario = c["proprietario"]
  is_core = c["is_core"]
  is_bear = c["is_bear"]
  cmf = c["cmf"]
  vsa = c["vsa_rating"]
  rsi = c["rsi"]
  storno = c["storno"]  
  pnl_pct = c["pnl_pct"]
  trend_5g = c.get("trend_volumi_5g", "Stabile ➡️")
  
  # 🚨 1. VENDI / PRENDI PROFITTO (FOMO & ESAURIMENTO VOLUMI)
  fomo_esaurimento = (rsi > 76) and (cmf < 0.0 or vsa == "🔴 DISTRIBUZIONE / VENDITA" or trend_5g == "In Raffreddamento 📉")
  if fomo_esaurimento:
    return (
        f"🚨 [{proprietario.upper()} - VENDI / PRENDI PROFITTO (FOMO)]: "
        f"Ipercomprato (RSI {rsi:.1f}) con volumi deboli. Incassa il guadagno prima della correzione!"
    )

  # ⚠️ 2. ALLARME STRUTTURALE / INTERROMPI DCA (Crollo profondo + flussi pesantemente negativi)
  if storno >= 35.0 and (vsa == "🔴 DISTRIBUZIONE / VENDITA" and cmf < -0.05):
    return (
        f"⚠️ [{proprietario.upper()} - INTERROMPI DCA (-{storno:.1f}%)]: "
        f"Storno pesante e flussi istituzionali negativi. Sospendi gli acquisti per proteggere il capitale."
    )

  # 🟢 3. ACCUMULA / DCA FINALIZZATO AL PROFITTO (Sconto sano dai massimi + assenza di panico/distribuzione)
  if storno >= 15.0 and cmf >= -0.03:
    return (
        f"🟢 [{proprietario.upper()} - ACCUMULA / DCA PROFITTO]: "
        f"Sconto sano (-{storno:.1f}% dai massimi) finalizzato al profitto futuro. Ottima area per incrementare (salvo quota massima raggiunta)."
    )

  # 🚀 4. TREND FORTE IN SPINTA
  if not is_bear and cmf > 0.03 and trend_5g == "In Accelerazione 📈":
    return (
        f"🚀 [{proprietario.upper()} - MANTIENI PER PROFITTO]: "
        f"Trend solido con volumi reali. Lascia correre i guadagni."
    )
  
  # 🟡 5. FASE DI PAZIENZA / ATTESA
  return (
      f"🟡 [{proprietario.upper()} - PAZIENTA / MANTIENI]: "
      f"Fase laterale o respiro fisiologico (-{storno:.1f}%). Mantieni la posizione in ottica di rendimento."
  )

# =====================================================================
# MAIN FUNCTION
# =====================================================================
def main():
  tickers = [t.strip().upper() for t in MEI_PORTAFOGLIO_CONFIG.keys() if t.strip()]
  if not tickers:
    print("⚠️ Nessun ticker inserito.")
    return
  
  eur_usd_rate = ottieni_tasso_cambio_eur_usd()
  print(f"🚀 Avvio scansione portafogli familiari (Cambio EUR/USD: {eur_usd_rate:.4f})...")
  
  try:
    df_raw = yf.download(tickers, period="1y", auto_adjust=True, progress=False)
  except Exception as e:
    print(f"Errore download: {e}")
    return
    
  candidati = []
  for ticker_str in tickers:
    try:
      if len(tickers) == 1:
        df_close, df_volume, df_high, df_low = df_raw["Close"], df_raw["Volume"], df_raw["High"], df_raw["Low"]
      else:
        df_close = df_raw["Close"][ticker_str] if "Close" in df_raw else df_raw.xs(ticker_str, axis=1, level=1)["Close"]
        df_volume = df_raw["Volume"][ticker_str] if "Volume" in df_raw else df_raw.xs(ticker_str, axis=1, level=1)["Volume"]
        df_high = df_raw["High"][ticker_str] if "High" in df_raw else df_raw.xs(ticker_str, axis=1, level=1)["High"]
        df_low = df_raw["Low"][ticker_str] if "Low" in df_raw else df_raw.xs(ticker_str, axis=1, level=1)["Low"]
        
      chiusure, volumi, massimi, minimi = df_close.dropna(), df_volume.dropna(), df_high.dropna(), df_low.dropna()
      if len(chiusure) < 50: continue
      
      prezzo_attuale = float(chiusure.iloc[-1])
      massimo_52w = float(chiusure.max())
      storno_pct = ((massimo_52w - prezzo_attuale) / massimo_52w) * 100
      
      config_titolo = MEI_PORTAFOGLIO_CONFIG.get(ticker_str, {})
      pmc_eur = config_titolo.get("pmc", 0.0)
      proprietario = config_titolo.get("proprietario", "Me")
      is_core = config_titolo.get("core", False) if proprietario == "Me" else False
      
      if pmc_eur > 0:
        pmc_usd = pmc_eur if (ticker_str.endswith(".MI") or ticker_str.endswith(".PA")) else pmc_eur * eur_usd_rate
        pnl_pct = round(((prezzo_attuale - pmc_usd) / pmc_usd) * 100, 2)
      else:
        pnl_pct = "N/D"
        
      trend_volumi_5g = analizza_trend_volumi_5g(volumi)
      reg_vol_50g = calcola_pendenza_regressione_volumi(volumi, 50)
      t_obj = yf.Ticker(ticker_str)
      ticker_display, fwd_pe, peg, short_pct = ottieni_dati_fondamentali_e_anagrafica(t_obj, ticker_str)
      
      sma_200 = float(chiusure.rolling(window=200).mean().iloc[-1]) if len(chiusure) >= 200 else prezzo_attuale
      is_bear_market = prezzo_attuale < sma_200
      rsi_attuale = float(calcola_rsi(chiusure).iloc[-1])
      
      cmf_val = 0.0
      if len(volumi) >= 20:
        cmf_val = float(calcola_cmf(massimi, minimi, chiusure, volumi, periodi=20).iloc[-1])
        
      vsa_rating = "🟢 ACCUMULAZIONE PULITA" if cmf_val > 0.03 else ("🔴 DISTRIBUZIONE / VENDITA" if cmf_val < -0.03 else "🟡 NEUTRO / VOLATILITÀ")
        
      diz_candidato = {
          "ticker_raw": ticker_str,
          "ticker_display": ticker_display,
          "proprietario": proprietario,
          "prezzo": prezzo_attuale,
          "pmc_eur": pmc_eur,
          "is_core": is_core,
          "pnl_pct": pnl_pct,
          "rsi": rsi_attuale,
          "storno": storno_pct,
          "is_bear": is_bear_market,
          "trend_volumi_5g": trend_volumi_5g,
          "reg_vol_50g": reg_vol_50g,
          "cmf": cmf_val,
          "vsa_rating": vsa_rating,
      }
      diz_candidato["suggerimento"] = genera_suggerimento_ibrido(diz_candidato)
      candidati.append(diz_candidato)
    except Exception as e:
      print(f"⚠️ Errore su {ticker_str}: {e}")
      continue
      
  if not candidati: return
  
  # EXCEL REPORT
  data_odierna = datetime.now().strftime("%Y-%m-%d")
  excel_filename = f"Report_Portafogli_Famiglia_{data_odierna}.xlsx"
  excel_data = []
  
  for c in candidati:
    excel_data.append({
        "Ticker": c["ticker_display"],
        "Proprietario": c["proprietario"],
        "Profilo Asset": "⭐ CORE (Lungo Termine)" if c["is_core"] else ("🚀 CRESCITA RAPIDA" if c["proprietario"] in ["Gabri", "Greg"] else "💼 Tattico"),
        "Prezzo Attuale ($)": round(c["prezzo"], 2),
        "Performance P&L (%)": f"{c['pnl_pct']:+.1f}%" if isinstance(c["pnl_pct"], (int, float)) else "N/D",
        "Trend Volumi 5G": c["trend_volumi_5g"],
        "Analisi VSA": c["vsa_rating"],
        "Suggerimento / Action": c["suggerimento"],
    })
    
  df_excel = pd.DataFrame(excel_data)
  df_excel.to_excel(excel_filename, index=False)
  print(f"📊 Excel generato con successo: {excel_filename}")
  
  # NOTIFICHE TELEGRAM & EMAIL
  righe = [f"📊 **MONITOR PORTAFOGLI FAMIGLIA - {data_odierna}**\n"]
  for c in candidati:
    tag = f"👶 [{c['proprietario'].upper()}]" if c["proprietario"] in ["Gabri", "Greg"] else "🧑 [ME]"
    righe.append(f"• {tag} **{c['ticker_raw']}** (${c['prezzo']:.1f}) ➔ *{c['suggerimento']}*")
    
  msg = "\n".join(righe)
  invia_telegram(CANALE_ACCUMULAZIONE_ID, msg)
  invia_email_con_allegato(f"📈 Report Portafogli Famiglia - {data_odierna}", msg.replace("**","").replace("*",""), excel_filename)

if __name__ == "__main__":
  main()
