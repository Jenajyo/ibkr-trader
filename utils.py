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

# Initialize IB connection with dynamic mode
ib = IB()
excel_file = None

# Function to set real or paper trading connection
def init_ibkr_connection(Trading_Mode):
    global ib, excel_file
    if Trading_Mode == "Paper":
        ib.connect("127.0.0.1", 7497, clientId=1)
        excel_file = "c:/Users/jyoti/Downloads/Stocks/IBKR_TRADER/PaperTrading/Orders_PaperTrading.xlsx"
    elif Trading_Mode == "Live":
        ib.connect("127.0.0.1", 7497, clientId=1)
        excel_file = "c:/Users/jyoti/Downloads/Stocks/IBKR_TRADER/Orders.xlsx"
    else:
        logger.error("Improper Trading Mode")


# Getter for excel_file path
def get_excel_file():
    return excel_file

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
    # Handle IBKR-specific symbols like BRK.B or BF.B
    if "." in ticker:
        contract = Stock(ticker.replace(".", " "), "SMART", "USD")
        contract.symbol = ticker
        contract.localSymbol = ticker
    else:
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

# Cancel existing LMT or TRAIL LIMIT orders for a ticker
def cancel_existing_orders(ticker):
    ib.reqOpenOrders()
    ib.sleep(0.5)
    for trade in ib.trades():
        if trade.contract.symbol == ticker and trade.order.orderType in ["LMT", "TRAIL LIMIT"]:
            if trade.orderStatus.status not in ["Cancelled", "Filled"]:
                ib.cancelOrder(trade.order)
                ib.sleep(0.5)

# Check current holdings for a ticker
def get_remaining_quantity(ticker):
    ib.sleep(1)
    for pos in ib.portfolio():
        if pos.contract.symbol == ticker and pos.contract.secType == "STK":
            return pos.position
    return 0

# Add trailing limit stop loss to all or specified stocks
def add_trailing_limit_to_holdings(trail_limit_percent=2.5, side="SELL", tickers=[]):
    for pos in ib.portfolio():
        if pos.contract.secType == "STK" and pos.position > 0:
            symbol = pos.contract.symbol
            if tickers and symbol not in tickers:
                continue
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
                totalQuantity=int(pos.position),
                trailingPercent=trail_limit_percent,
                trailStopPrice=trail_stop_price,
                lmtPriceOffset=0.10,
                tif="GTC",
                outsideRth=True
            )
            trade = ib.placeOrder(contract, trailing_order)
            ib.sleep(1)
            logger.info(f"[TRAIL-ATTACH] {symbol}: Trailing limit placed at {trail_limit_percent}% for {int(pos.position)} shares.")

# Update sheet in Excel
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

# Update BUY_USUAL and SELL sheets based on holdings
def update_orders_page(Trading_Mode):
    sheets = pd.read_excel(excel_file, sheet_name=None)

    # Enforce correct dtypes
    dtype_mapping = {
        "Ticker": 'str',
        "Amount": 'float',
        "Quantity": 'float',
        "TrailLimit%": 'float',
        "OrderType": 'str',
        "Status": 'str',
        "Execution": 'str'
    }
    for sheet_name, df in sheets.items():
        if not df.empty:
            for col, dtype in dtype_mapping.items():
                if col in df.columns:
                    df[col] = df[col].astype(dtype, errors='ignore')
            sheets[sheet_name] = df

    holdings = {pos.contract.symbol: pos.position for pos in ib.portfolio() if pos.contract.secType == "STK"}

    # Helper: Find existing TrailLimit% if any
    def get_existing_trail_percent(symbol):
        ib.reqOpenOrders()
        ib.sleep(1)
        for trade in ib.trades():
            if (trade.contract.symbol == symbol 
                and trade.order.orderType == "TRAIL LIMIT" 
                and trade.orderStatus.status not in ["Cancelled", "Filled"]):
                return trade.order.trailingPercent
        return None

    # Helper: Place a new trail limit order
    def place_trailing_limit(symbol, quantity, action, trail_percent):
        contract = qualify_contract(symbol)
        trail_stop_price = round(
            get_market_price(symbol) * (1 - trail_percent / 100)
            if action == "SELL"
            else get_market_price(symbol) * (1 + trail_percent / 100),
            2
        )
        order = Order(
            action=action,
            orderType="TRAIL LIMIT",
            totalQuantity=int(quantity),
            trailingPercent=trail_percent,
            trailStopPrice=trail_stop_price,
            lmtPriceOffset=0.10,
            tif="GTC",
            outsideRth=True
        )
        ib.placeOrder(contract, order)
        ib.sleep(1)

    def sync_sheet(sheet_name, action):
        df = sheets.get(sheet_name, pd.DataFrame())
        if df.empty:
            df = pd.DataFrame({
                "Ticker": pd.Series(dtype='str'),
                "Amount": pd.Series(dtype='float'),
                "Quantity": pd.Series(dtype='float'),
                "TrailLimit%": pd.Series(dtype='float'),
                "OrderType": pd.Series(dtype='str'),
                "Status": pd.Series(dtype='str'),
                "Execution": pd.Series(dtype='str')
            })
        tickers_in_sheet = set(df["Ticker"].astype(str).str.upper())
        updated_tickers = set()

        for symbol, qty in holdings.items():
            if action == "BUY" and qty <= 0:
                continue
            if action == "SELL" and qty > 0:
                continue

            price = get_market_price(symbol)
            trail_limit_percent = get_existing_trail_percent(symbol)

            if trail_limit_percent is None:
                # No existing trail limit order, place one with 5%
                default_trail = 5.0
                order_action = "SELL" if action == "BUY" else "BUY"
                place_trailing_limit(symbol, qty, order_action, default_trail)
                trail_limit_percent = default_trail

            update_data = {
                "Amount": math.ceil(qty * price),
                "Quantity": qty,
                "TrailLimit%": trail_limit_percent,
                "Status": "Open",
                "Execution": " ",
                "OrderType": "MKT-ATCH-LIMIT"
            }

            updated_tickers.add(symbol)

            if symbol in tickers_in_sheet:
                for col, val in update_data.items():
                    df.loc[df["Ticker"].str.upper() == symbol, col] = val
            else:
                df = pd.concat([df, pd.DataFrame([{"Ticker": symbol, **update_data}])], ignore_index=True)

        if Trading_Mode == "Paper":
            df = df[df["Ticker"].str.upper().isin(updated_tickers)]
        else:
            for symbol in tickers_in_sheet:
                if symbol not in updated_tickers:
                    for col, val in zip(["Amount", "Quantity", "TrailLimit%", "OrderType", "Status", "Execution"], [2000, " ", 5.0, "MKT-ATCH-LIMIT", "", ""]):
                        df.loc[df["Ticker"].str.upper() == symbol, col] = val

        update_sheet_in_excel(sheet_name, df)

    sync_sheet("BUY_Usual", "BUY")
    sync_sheet("SELL", "SELL")
