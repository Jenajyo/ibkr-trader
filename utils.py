import logging
CANCEL_ALL_FIRST = True
import pandas as pd
import math
from ib_insync import *
import yfinance as yf
from datetime import datetime

# Suppress ib_insync internal logs
logging.getLogger('ib_insync').setLevel(logging.WARNING)

# Configure logging for your own outputs
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Initialize IB connection
ib = IB()
ib.connect("127.0.0.1", 7497, clientId=1)

# Path to the Excel workbook
excel_file = "c:/Users/jyoti/Downloads/IBKR_TRADER/Orders.xlsx"

# Cancel all open orders
def cancel_all_open_orders():
    try:
        ib.reqGlobalCancel()
        ib.sleep(2)
        ib.reqOpenOrders()
        ib.sleep(2)
        ib.waitOnUpdate(timeout=10)
        open_orders = ib.openOrders()
        for order in open_orders:
            if order.orderId != 0:
                ib.cancelOrder(order)
                logger.info(f"Cancelled open order ID {order.orderId}")
        logger.info("All open orders cancelled.")
    except Exception as e:
        logger.error(f"Error cancelling all open orders: {e}")

# Get the latest market price using IBKR, fallback to Yahoo Finance if needed
def get_market_price(ticker):
    try:
        contract = Stock(ticker, "SMART", "USD")
        ib.qualifyContracts(contract)
        market_data = ib.reqMktData(contract, snapshot=True)
        ib.sleep(1)
        price = market_data.last if market_data.last else (market_data.close or market_data.ask or market_data.bid)
        if price and price > 0:
            return price
        raise ValueError("No valid IBKR price")
    except Exception as e:
        logger.warning(f"IBKR price fetch failed for {ticker}, falling back to Yahoo Finance: {e}")
        stock_info = yf.Ticker(ticker)
        return stock_info.history(period="1d")["Close"].iloc[-1]

# Return a qualified IB contract for the given ticker
def qualify_contract(ticker):
    contract = Stock(ticker, "SMART", "USD")
    ib.qualifyContracts(contract)
    return contract

# Place a GTC market order through IB
def place_market_order(contract, action, quantity):
    order = MarketOrder(action, quantity)
    order.tif = "GTC"
    order.rthOnly = True
    trade = ib.placeOrder(contract, order)
    ib.sleep(1)
    return trade

# Place a limit order through IB at market price + $0.50
def place_limit_order(contract, action, quantity, market_price):
    order = LimitOrder(
        action=action,
        totalQuantity=quantity,
        lmtPrice=round(market_price + 0.50, 2),
        tif="GTC",
        outsideRth=True
    )
    trade = ib.placeOrder(contract, order)
    ib.sleep(1)
    return trade

# Attach a trailing limit order with given offset and percent to a contract
def attach_trailing_limit(contract, action, quantity, market_price, trail_limit_percent):
    reverse_action = "SELL" if action == "BUY" else "BUY"
    trail_stop_price = round(
        market_price * (1 - trail_limit_percent / 100)
        if action == "BUY"
        else market_price * (1 + trail_limit_percent / 100),
        2,
    )
    trailing_order = Order(
        action=reverse_action,
        orderType="TRAIL LIMIT",
        totalQuantity=quantity,
        trailingPercent=trail_limit_percent,
        trailStopPrice=trail_stop_price,
        lmtPriceOffset=0.10,
        tif="GTC",
        outsideRth=True,
    )
    trailing_trade = ib.placeOrder(contract, trailing_order)
    ib.sleep(1)
    return trailing_trade

# Cancel all open limit or trailing limit orders for a given ticker
def cancel_existing_orders(ticker):
    ib.reqOpenOrders()
    ib.sleep(0.5)
    for trade in ib.trades():
        if trade.contract.symbol == ticker and trade.order.orderType in ["LMT", "TRAIL LIMIT"]:
            if trade.orderStatus.status not in ["Cancelled", "Filled"]:
                ib.cancelOrder(trade.order)
                ib.sleep(0.5)

# Check current holdings for a ticker and return the position size
def get_remaining_quantity(ticker):
    ib.sleep(1)
    for pos in ib.portfolio():
        if pos.contract.symbol == ticker and pos.contract.secType == "STK":
            return pos.position
    return 0

# Add a trailing limit stop loss to all currently held positions
def add_trailing_limit_to_holdings(trail_limit_percent=2.5, side="SELL"):
    for pos in ib.portfolio():
        if pos.contract.secType == "STK" and pos.position > 0:
            symbol = pos.contract.symbol
            logger.info(f"[TRAIL-ATTACH] {symbol}: Replacing existing limit/trailing orders.")
            cancel_existing_orders(symbol)
            price = get_market_price(symbol)
            contract = qualify_contract(symbol)
            action = side.upper()
            trail_stop_price = round(
                price * (1 - trail_limit_percent / 100)
                if action == "SELL"
                else price * (1 + trail_limit_percent / 100),
                2
            )
            trailing_order = Order(
                action=action,
                orderType="TRAIL LIMIT",
                totalQuantity=int(pos.position),  # ✅ Ensures no fractional shares
                trailingPercent=trail_limit_percent,
                trailStopPrice=trail_stop_price,
                lmtPriceOffset=0.10,
                tif="GTC",
                outsideRth=True
            )
            trade = ib.placeOrder(contract, trailing_order)
            ib.sleep(1)
            logger.info(f"[TRAIL-ATTACH] {symbol}: Trailing limit placed at {trail_limit_percent}% for {int(pos.position)} shares.")

# Update the given sheet in the Excel file with the latest DataFrame content
def update_sheet_in_excel(sheet_name, df):
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# Append an entry to the "Log" sheet to track order activity
def append_to_log(symbol, action, quantity, price):
    log_entry = {
        "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Symbol": symbol,
        "Type": action,
        "Quantity": quantity,
        "Price": round(quantity * price, 2),
    }
    try:
        existing_log = pd.read_excel(excel_file, sheet_name="Log")
        updated_log = pd.concat([existing_log, pd.DataFrame([log_entry])], ignore_index=True)
    except Exception:
        updated_log = pd.DataFrame([log_entry])

    with pd.ExcelWriter(excel_file, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        updated_log.to_excel(writer, sheet_name="Log", index=False)
