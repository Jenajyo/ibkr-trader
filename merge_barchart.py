import pandas as pd
import os
from openpyxl import load_workbook

# Define file paths
base_path = r"C:\Users\jyoti\Downloads"
output_path = r"C:\Users\jyoti\Downloads\IBKR_Trader\PaperTrading"
output_file = os.path.join(output_path, "merged_intraday.csv")
orders_file = os.path.join(output_path, "Orders_PaperTrading.xlsx")

# List of input files to merge
input_files = [
    "top-1-signal-strength-intraday-04-24-2025.csv",
    "top-1-signal-direction-intraday-04-24-2025.csv",
    "top-stocks-to-own-intraday-04-24-2025.csv",
    "all-us-exchanges-price-volume-leaders-04-24-2025.csv"
]

# Read and concatenate all files
dfs = []
for file in input_files:
    full_path = os.path.join(base_path, file)
    df = pd.read_csv(full_path)
    dfs.append(df)

merged_df = pd.concat(dfs, ignore_index=True)

# Drop duplicates based on the Symbol column
if "Symbol" in merged_df.columns:
    merged_df = merged_df.drop_duplicates(subset=["Symbol"])
else:
    raise ValueError("❌ 'Symbol' column not found in input files.")

# Sort by Price Vol column if it exists
if "Price Vol" in merged_df.columns:
    merged_df = merged_df.sort_values(by="Price Vol", ascending=False)
else:
    raise ValueError("❌ 'Price Vol' column not found in input files.")

# Save the final output
merged_df.to_csv(output_file, index=False)
print(f"✅ Merged, deduplicated, and sorted file saved to: {output_file}")

# Filter symbols for buying recommendation
buy_candidates = merged_df[
    (merged_df["Short Term"] == "100% Buy") &
    (merged_df["Medium Term"] == "100% Buy") &
    (merged_df["Long Term"] == "100% Buy") &
    (merged_df["# Analysts"] > 5)
]
buy_candidate_symbols = set(buy_candidates["Symbol"].str.upper())

# Load existing BUY_Usual sheet
existing_df = pd.read_excel(orders_file, sheet_name="BUY_Usual")
existing_df["Ticker"] = existing_df["Ticker"].astype(str).str.upper()

# Get list of current holdings from BUY_Usual sheet
current_holdings = set(existing_df["Ticker"].unique())

# Case 1 & 2: Add new entries or skip existing
new_rows = []
for _, row in buy_candidates.iterrows():
    symbol = row["Symbol"].upper()
    if symbol not in existing_df["Ticker"].values:
        new_rows.append({
            "Ticker": symbol,
            "Amount": 1000,
            "Quantity": " ",
            "TrailLimit%": 7,
            "OrderType": "MKT-ATCH-LIMIT",
            "Status": " ",
            "Execution": "TRANSMIT"
        })

# Case 3: Set OrderType to "ATCH-LMT" for existing symbols in holdings but not in buy_candidates (and not already ATCH-LMT with 2.5)
for idx, row in existing_df.iterrows():
    ticker = row["Ticker"]
    if (
        ticker not in buy_candidate_symbols and
        row.get("OrderType") != "ATCH-LMT" or row.get("TrailLimit%") != 2.5
    ):
        existing_df.at[idx, "OrderType"] = "ATCH-LMT"
        existing_df.at[idx, "TrailLimit%"] = 2.5
        existing_df.at[idx, "Execution"] = "TRANSMIT"

# Append new rows
updated_df = pd.concat([existing_df, pd.DataFrame(new_rows)], ignore_index=True)

# Save to Excel
with pd.ExcelWriter(orders_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    updated_df.to_excel(writer, sheet_name="BUY_Usual", index=False)

print(f"✅ BUY_Usual sheet updated in Orders_PaperTrading.xlsx with new rows and status updates.")
