import os
import yfinance as yf
import requests
from ib_insync import IB, Stock
from dotenv import load_dotenv

# # Load API keys
# load_dotenv()
# FINNHUB_KEY = os.getenv("FINNHUB_API_KEY")
# IEX_KEY = os.getenv("IEX_API_KEY")
# POLYGON_KEY = os.getenv("POLYGON_API_KEY")

symbol = "AAPL"

# Yahoo Finance
def get_yahoo_price(symbol):
    ticker = yf.Ticker(symbol)
    data = ticker.history(period="1d")
    return round(data['Close'].iloc[-1], 2)

# IBKR (requires TWS or IB Gateway running)
def get_ibkr_price(symbol):
    try:
        ib = IB()
        ib.connect("127.0.0.1", 7497, clientId=123)
        contract = Stock(symbol, "SMART", "USD")
        ib.qualifyContracts(contract)
        ticker_data = ib.reqMktData(contract, snapshot=True)
        ib.sleep(2)
        ib.disconnect()
        return round(ticker_data.last, 2)
    except Exception as e:
        print(f"IBKR error: {e}")
        return None

# Finnhub
def get_finnhub_price(symbol):
    url = f"https://finnhub.io/api/v1/quote?symbol={symbol}&token={FINNHUB_KEY}"
    r = requests.get(url)
    data = r.json()
    return round(data["c"], 2) if "c" in data else None

# IEX Cloud
def get_iex_price(symbol):
    url = f"https://cloud.iexapis.com/stable/stock/{symbol}/quote?token={IEX_KEY}"
    r = requests.get(url)
    data = r.json()
    return round(data["latestPrice"], 2) if "latestPrice" in data else None

# Polygon.io
def get_polygon_price(symbol):
    url = f"https://api.polygon.io/v2/last/trade/{symbol}?apiKey={POLYGON_KEY}"
    r = requests.get(url)
    data = r.json()
    return round(data["results"]["p"], 2) if "results" in data else None

# Results
print(f"📈 Yahoo Finance price for {symbol}: ${get_yahoo_price(symbol)}")
print(f"🏦 IBKR price for {symbol}: ${get_ibkr_price(symbol)}")
# print(f"🌐 Finnhub price for {symbol}: ${get_finnhub_price(symbol)}")
# print(f"📡 IEX Cloud price for {symbol}: ${get_iex_price(symbol)}")
# print(f"📊 Polygon.io price for {symbol}: ${get_polygon_price(symbol)}")
