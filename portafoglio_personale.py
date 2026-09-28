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
# ANALISI TECNICA & VOLUMI AVANZATA (Market Screener Integration)
# =====================================================================
def calcola_rsi(chiusure, periodi=14):
  delta = chiusure.diff()
  guadagno = (delta.where(delta > 0, 0)).rolling(window=periodi).mean()
  perdita = (-delta.where(delta < 0, 0)).rolling(window=periodi).mean()
  rs = guadagno / (perdita + 1e-10)
  return 100 - (100 / (1 + rs))

def calcola_cmf(massimi, minimi, chiusure, volumi, periodi=20):
  mf_multiplier = ((chiusure - minimi) - (massimi - chiusure)) / (massimi - minimi + 1e-10)
  mf_volume = mf_multiplier * volumi
  return mf_volume.rolling(window=periodi).sum() / (volumi.rolling(window=periodi).sum() + 1e-10)

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

def rileva_divergenza_cmf(df, lookback=15):
  if len(df) < lookback + 5:
    return "Assente"
  sub_df = df.iloc[-lookback:].copy()
  prices = sub_df["Close"].values
  cmf_vals = sub_df["CMF"].values
  p_min1, p_min2 = np.min(prices[:lookback // 2]), np.min(prices[lookback // 2:])
  c_min1, c_min2 = np.min(cmf_vals[:lookback // 2]), np.min(cmf_vals[lookback // 2:])
  p_max1, p_max2 = np.max(prices[:lookback // 2]), np.max(prices[lookback // 2:])
  c_max1, c_max2 = np.max(cmf_vals[:lookback // 2]), np.max(cmf_vals[lookback // 2:])
  if p_min2 < p_min1 and c_min2 > c_min1:
    return "Rialzista 🟢"
  if p_max2 > p_max1 and c_max2 < c_max1:
    return "Ribassista 🔴"
  return "Assente"

def calcola_volatilita_squeeze(df, length=20):
  try:
    sma = df['Close'].rolling(window=length).mean()
    std = df['Close'].rolling(window=length).std()
    bb_upper, bb_lower = sma + (2 * std), sma - (2 * std)
    atr = (df['High'] - df['Low']).combine((df['High'] - df['Close'].shift()).abs(), max).combine((df['Low'] - df['Close'].shift()).abs(), max).rolling(window=length).mean()
    kc_upper, kc_lower = sma + (1.5 * atr), sma - (1.5 * atr)
    return (bb_upper.iloc[-1] <= kc_upper.iloc[-1]) and (bb_lower.iloc[-1] >= kc_lower.iloc[-1])
  except Exception:
    return False

def verifica_salute_finanziaria_intelligente(t_obj, cmf_val):
  try:
    info = t_obj.info or {}
    revenue_growth = info.get('revenueGrowth', None)
    debt_to_equity = info.get('debtToEquity', None)
    earnings_growth = info.get('earningsGrowth', None)
    if revenue_growth is not None and revenue_growth < -0.30 and cmf_val < 0.10:
      return False, "Crollo ricavi senza supporto istituzionale."
    if debt_to_equity is not None and debt_to_equity > 400 and cmf_val < 0.05:
      return False, "Indice di indebitamento critico."
    if earnings_growth is not None and earnings_growth < -0.35 and cmf_val < 0.05:
      return False, "Utili in forte calo senza flussi a favore."
    return True, "Azienda solida o in investimento."
  except Exception:
    return True, "Dati fondamentali parziali."

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
# MOTORE DI SUGGERIMENTO IBRIDO POTENZIATO
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
  reg_20g = c.get("reg_vol_20g", "Stabile ➡️")
  reg_50g = c.get("reg_vol_50g", "Stabile ➡️")
  divergenza = c.get("divergenza_cmf", "Assente")
  is_squeeze = c.get("is_squeeze", False)
  short_pct = c.get("short_pct", "N/D")
  is_sano = c.get("is_sano", True)
  nota_bilancio = c.get("nota_bilancio", "")

  # 🚨 ALLARME FONDAMENTALE (Se l'azienda mostra crepe strutturali gravi)
  if not is_sano and cmf < 0:
    return f"⚠️ [ALLARME FONDAMENTALE]: {nota_bilancio} Valuta disinvestimento o rotazione immediata."

  # 🚀 LOGICA FIGLI (GABRI / GREG): MASSIMA VELOCITÀ, SQUEEZE & ZERO LATERALITÀ
  if proprietario in ["Gabri", "Greg"]:
    if short_pct != "N/D" and short_pct > 10 and is_squeeze and trend_5g == "In Accelerazione 📈":
      return f"🔥 [{proprietario.upper()} - SQUEEZE ESPLOSIVO]: Volatilità compressa e volumi in forte accelerazione. Tieni duro e fai correre!"
    if divergenza == "Rialzista 🟢" and storno >= 10.0:
      return f"💎 [{proprietario.upper()} - ACCUMULO ISTITUZIONALE SU STORNO]: Storna del -{storno:.1f}% ma c'è divergenza CMF rialzista. Ottimo per tenere o mediare!"
    if is_bear or storno >= 15.0 or cmf < -0.02 or trend_5g == "In Raffreddamento 📉" or reg_20g == "In Esaurimento 📉":
      return (
          f"🚨 [{proprietario.upper()} - VENDI / RUOTA]: Il titolo rallenta o storna (-{storno:.1f}%)."
          " Stop alla lateralizzazione: **vendi adesso** per reinvestire su chi corre!"
      )
    elif rsi > 75 or vsa == "🔴 DISTRIBUZIONE":
      return (
          f"💰 [{proprietario.upper()} - PRENDI PROFITTO]: Area di massimo (RSI {rsi:.1f})."
          " Incassa e ruota sul prossimo cavallo vincente."
      )
    elif not is_bear and cmf > 0.05:
      return (
          f"🚀 [{proprietario.upper()} - IN SPINTA]: Trend solido e flussi a favore."
          " Lascia correre il profitto!"
      )

  # 🧑 LOGICA PERSONALE ("ME"): GESTIONE STANDARD / CORE / TATTICO
  if is_core:
    if storno >= 35.0:
      if is_bear and cmf < -0.05 and reg_50g == "In Esaurimento 📉":
        return f"⚠️ [ALLARME CORE - ROTTURA STRUTTURALE (-{storno:.1f}%)]: Valuta alleggerimento."
      else:
        return f"🛡️ [CORE IN SCONTO (-{storno:.1f}%)]: Mantieni o accumula (DCA)."
    elif storno >= 20.0:
      return f"💎 [CORE IN CORREZIONE (-{storno:.1f}%)]: Storno fisiologico, tesi intatta."
      
  if divergenza == "Rialzista 🟢":
    return f"🚀 [OCCASIONE D'ORO]: Divergenza CMF rialzista attiva in area di storno (-{storno:.1f}%). Accumulo intelligente."
    
  if (rsi > 78 or (storno < 3.0 and rsi > 72)) and (cmf < -0.02 or vsa == "🔴 DISTRIBUZIONE"):
    return f"💰 [ZONA CALDA - PRENDI PROFITTO]: Titolo tirato (RSI {rsi:.1f}). Alleggerisci la quota."
    
  if pnl_pct != "N/D" and pnl_pct > 50.0 and not is_bear:
    return f"🚀 [GAIN STRAORDINARIO (+{pnl_pct:.1f}%)]: Fai correre i profitti."
    
  if not is_bear and cmf > 0.05 and vsa == "🟢 ACCUMULAZIONE PULITA":
    return "🟢 [ACCUMULO ATTIVO]: Trend e flussi istituzionali sani."
  elif cmf >= 0.0 and storno >= 12.0:
    return "💎 [SCONTO STRATEGICO]: Storno salutare con assorbimento istituzionale."
  elif vsa == "🔴 DISTRIBUZIONE" or cmf < -0.05:
    return "🔴 [DISTRIBUZIONE]: Flussi negativi. Evita di mediare."
  else:
    if rsi < 30: return "🟡 [AREA IPERVENDUTO]: Monitorare per rimbalzo."
    return "🟡 [FASE NEUTRA]: Mantenere la posizione senza fretta."

# =====================================================================
# MAIN FUNCTION
# =====================================================================
def main():
  tickers = [t.strip().upper() for t in MEI_PORTAFOGLIO_CONFIG.keys() if t.strip()]
  if not tickers:
    print("⚠️ Nessun ticker inserito.")
    return
  
  eur_usd_rate = ottieni_tasso_cambio_eur_usd()
  print(f"🚀 Avvio scansione portafogli familiari potenziata (Cambio EUR/USD: {eur_usd_rate:.4f})...")
  
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
        
      # Indicatori tecnici avanzati
      trend_volumi_5g = analizza_trend_volumi_5g(volumi)
      reg_vol_20g = calcola_pendenza_regressione_volumi(volumi, 20)
      reg_vol_50g = calcola_pendenza_regressione_volumi(volumi, 50)
      
      t_obj = yf.Ticker(ticker_str)
      ticker_display, fwd_pe, peg, short_pct = ottieni_dati_fondamentali_e_anagrafica(t_obj, ticker_str)
      
      sma_200 = float(chiusure.rolling(window=200).mean().iloc[-1]) if len(chiusure) >= 200 else prezzo_attuale
      is_bear_market = prezzo_attuale < sma_200
      rsi_attuale = float(calcola_rsi(chiusure).iloc[-1])
      
      cmf_val = 0.0
      if len(volumi) >= 20:
        # Prepariamo un dataframe temporaneo per calcolare la divergenza CMF e lo Squeeze
        df_temp = pd.DataFrame({'Close': chiusure, 'High': massimi, 'Low': minimi, 'Volume': volumi})
        df_temp['CMF'] = calcola_cmf(df_temp['High'], df_temp['Low'], df_temp['Close'], df_temp['Volume'])
        cmf_val = float(df_temp['CMF'].iloc[-1])
        divergenza = rileva_divergenza_cmf(df_temp)
        is_squeeze = calcola_volatilita_squeeze(df_temp)
      else:
        divergenza = "Assente"
        is_squeeze = False

      is_sano, nota_bilancio = verifica_salute_finanziaria_intelligente(t_obj, cmf_val)
      vsa_rating = "🟢 ACCUMULAZIONE PULITA" if cmf_val > 0.05 else ("🔴 DISTRIBUZIONE" if cmf_val < -0.05 else "🟡 NEUTRO")
        
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
          "reg_vol_20g": reg_vol_20g,
          "reg_vol_50g": reg_vol_50g,
          "cmf": cmf_val,
          "vsa_rating": vsa_rating,
          "divergenza_cmf": divergenza,
          "is_squeeze": is_squeeze,
          "short_pct": short_pct,
          "is_sano": is_sano,
          "nota_bilancio": nota_bilancio
      }
      diz_candidato["suggerimento"] = genera_suggerimento_ibrido(diz_candidato)
      candidati.append(diz_candidato)
    except Exception as e:
      print(f"⚠️ Errore su {ticker_str}: {e}")
      continue
      
  if not candidati: return
  
  # EXCEL REPORT CON COLONNA DEDICATA AL PROPRIETARIO
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
        "Divergenza CMF": c["divergenza_cmf"],
        "Squeeze Volatilità": "Attivo 🔥" if c["is_squeeze"] else "No",
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
