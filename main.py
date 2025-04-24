import logging
import pandas as pd
import math
from utils import (
    ib, get_market_price, qualify_contract, place_market_order,
    attach_trailing_limit, cancel_existing_orders, get_remaining_quantity,
    update_sheet_in_excel, append_to_log, place_limit_order,
    cancel_all_open_orders, add_trailing_limit_to_holdings,
    update_orders_page, init_ibkr_connection
)

# Configuration Flags
Trading_Mode = "Live"
CANCEL_ALL_FIRST = False
APPLY_TRAIL_TO_HOLDINGS = False
RUN_ORDER_PAGE_UPDATE = False

# Logging setup
logging.getLogger('ib_insync').setLevel(logging.WARNING)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Initialize IBKR connection based on flag
init_ibkr_connection(Trading_Mode)

# Excel file path is determined by init_ibkr_connection
from utils import excel_file


def handle_market_orders(index, df, ticker, amount, quantity, action, order_type, trail_limit_percent):
    try:
        market_price = get_market_price(ticker)
        quantity = math.ceil(amount / market_price) if quantity is None else quantity
        contract = qualify_contract(ticker)
        trade = place_market_order(contract, action, quantity)
        if trade.orderStatus.status in ["Filled", "Submitted"]:
            df.at[index, "Quantity"] = quantity
            df.at[index, "Amount"] = math.ceil(quantity * market_price)
            if order_type == "MKT-ATCH-LIMIT":
                trailing_trade = attach_trailing_limit(contract, action, quantity, market_price, trail_limit_percent)
                if trailing_trade.orderStatus.status in ["Filled", "Submitted"]:
                    df.at[index, "Status"] = "Order Placed with Limit Attached"
            else:
                df.at[index, "Status"] = "MKT Order Placed"
            df.at[index, "Execution"] = " "
            append_to_log(ticker, action, quantity, market_price)
            logger.info(f"[{order_type}] {action} {quantity} shares of {ticker} at ${market_price:.2f}")
    except Exception as e:
        logger.error(f"Error processing market order for {ticker}: {e}")

def handle_remove_limit_order(index, df, ticker, cancelled_tickers):
    try:
        logger.info(f"[REMOVE-LIMIT] Checking open orders for {ticker}:")
        if ticker not in cancelled_tickers:
            initial_count = len(ib.trades())
            cancel_existing_orders(ticker)
            final_count = len([t for t in ib.trades() if t.orderStatus.status != "Cancelled"])
            cancelled_tickers.add(ticker)
            if initial_count > final_count:
                df.at[index, "Status"] = "Limit Order Cancelled"
                df.at[index, "Execution"] = " "
                logger.info(f"[REMOVE-LIMIT] {ticker}: Limit order(s) cancelled.")
            else:
                logger.info(f"[REMOVE-LIMIT] {ticker}: No active limit orders found.")
    except Exception as e:
        logger.error(f"Error cancelling orders for {ticker}: {e}")

def handle_attach_limit(index, df, ticker, quantity, action, trail_limit_percent):
    try:
        contract = qualify_contract(ticker)
        cancel_existing_orders(ticker)
        market_price = get_market_price(ticker)
        if not quantity:
            quantity = get_remaining_quantity(ticker)
        trailing_trade = attach_trailing_limit(contract, action, quantity, market_price, trail_limit_percent)
        if trailing_trade.orderStatus.status in ["Filled", "Submitted"]:
            df.at[index, "Status"] = "Order Placed with Limit Attached"
            df.at[index, "Execution"] = " "
            logger.info(f"[ATCH-LMT] {ticker}: Trailing limit set for {quantity} shares at ${market_price:.2f}")
    except Exception as e:
        logger.error(f"Error attaching trailing limit for {ticker}: {e}")

def handle_lmt_attach_trail_limit(index, df, ticker, amount, quantity, action, trail_limit_percent):
    try:
        contract = qualify_contract(ticker)
        market_price = get_market_price(ticker)
        limit_price = round(market_price + 0.10, 2)

        logger.info(f"{ticker} Market Price = {market_price}, Limit Price = {limit_price}")

        quantity = math.ceil(amount / limit_price) if quantity is None else quantity

        trade = place_limit_order(contract, action, quantity, market_price)
        if trade.orderStatus.status in ["Filled", "Submitted"]:
            df.at[index, "Quantity"] = quantity
            df.at[index, "Amount"] = math.ceil(quantity * limit_price)
            trailing_trade = attach_trailing_limit(contract, action, quantity, limit_price, trail_limit_percent)
            if trailing_trade.orderStatus.status in ["Filled", "Submitted"]:
                df.at[index, "Status"] = "LMT with Trailing Limit Attached"
                df.at[index, "Execution"] = " "
                append_to_log(ticker, action, quantity, limit_price)
                logger.info(f"[LMT-ATTCH-TRAIL-LIMIT] {action} {quantity} shares of {ticker} at ${limit_price:.2f} with trailing limit")
        else:
            logger.warning(f"[LMT-ATTCH-TRAIL-LIMIT] {ticker}: Limit order not submitted (status: {trade.orderStatus.status})")
    except Exception as e:
        logger.error(f"Error in LMT-ATTCH-TRAIL-LIMIT for {ticker}: {e}")

def handle_close(index, df, ticker, quantity, action):
    try:
        contract = qualify_contract(ticker)
        cancel_existing_orders(ticker)
        if pd.isna(quantity) or quantity in ["", " "]:
            quantity = get_remaining_quantity(ticker)
        close_action = "SELL" if action == "BUY" else "BUY"
        trade = place_market_order(contract, close_action, quantity)
        if trade.orderStatus.status in ["Filled", "Submitted"]:
            remaining = get_remaining_quantity(ticker)
            df.at[index, "Quantity"] = pd.NA if remaining == 0 else remaining
            df.at[index, "Status"] = "Closed"
            df.at[index, "Execution"] = " "
            final_price = get_market_price(ticker)
            append_to_log(ticker, close_action, quantity, final_price)
            logger.info(f"[CLOSE] {ticker}: Closed {quantity} shares at ${final_price:.2f}")
    except Exception as e:
        logger.error(f"Error closing position for {ticker}: {e}")

def process_sheet(sheet_name, df):
    cancelled_tickers = set()
    for index, row in df.iterrows():
        try:
            execution = str(row.get("Execution", "")).strip().upper()
            if execution != "TRANSMIT":
                continue
            ticker = row["Ticker"]
            amount = row["Amount"]
            order_type = str(row.get("OrderType", "")).strip().upper()
            trail_limit_percent = row.get("TrailLimit%", 4.0)
            quantity = row.get("Quantity", None)
            quantity = int(quantity) if not pd.isna(quantity) else None
            action = "BUY" if sheet_name.startswith("BUY") else "SELL"
            match order_type:
                case "LMT-ATTCH-TRAIL-LIMIT":
                    handle_lmt_attach_trail_limit(index, df, ticker, amount, quantity, action, trail_limit_percent)
                case "MKT" | "MKT-ATCH-LIMIT":
                    handle_market_orders(index, df, ticker, amount, quantity, action, order_type, trail_limit_percent)
                case "REMOVE-LIMIT-ORDER":
                    if quantity in [None, "", " "]:
                        handle_remove_limit_order(index, df, ticker, cancelled_tickers)
                case "ATCH-LMT":
                    handle_attach_limit(index, df, ticker, quantity, action, trail_limit_percent)
                case "CLOSE":
                    handle_close(index, df, ticker, quantity, action)
        except Exception as e:
            logger.error(f"Unexpected error in row {index} of sheet {sheet_name}: {e}")
    update_sheet_in_excel(sheet_name, df)

# Set this flag to True if you want to cancel all open orders before running
CANCEL_ALL_FIRST = False

# Set this flag to True if you want to attach limit trail orders all open orders before running
APPLY_TRAIL_TO_HOLDINGS = False

# Set this flag to True if you want to attach limit trail orders all open orders before running
RUN_ORDER_PAGE_UPDATE = False

# Run logic
def run():
    try:
        if CANCEL_ALL_FIRST:
            cancel_all_open_orders()
            logger.info("Order cancellation requested. Retrying order execution after cleanup.")
            return

        if APPLY_TRAIL_TO_HOLDINGS:
            add_trailing_limit_to_holdings(trail_limit_percent=2.5, side="SELL", tickers=["IAU"])
            logger.info("Trailing limit orders applied to all holdings.")
            return

        if RUN_ORDER_PAGE_UPDATE:
            update_orders_page(Trading_Mode)
            logger.info("Updated Orders with holdings with Buy_Usual and SELL sheet.")
            return

        sheets = pd.read_excel(excel_file, sheet_name=None)
        for sheet_name, sheet_data in sheets.items():
            if sheet_name.startswith("BUY") or sheet_name.startswith("SELL"):
                sheet_data.columns = sheet_data.columns.str.strip()
                process_sheet(sheet_name, sheet_data)
    finally:
        ib.disconnect()

if __name__ == "__main__":
    run()

