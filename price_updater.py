import datetime as dt
import os
from dataclasses import dataclass
from typing import Dict, List

import pandas as pd
import yfinance as yf

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "Funfamily_Stock_Tracker.xlsx")
STOCK_SHEET_NAME = "Stock Sheet"
HEADER_ROW_INDEX = 2  # zero-based row index in Excel where the real headers live

TICKER_MAP: Dict[str, str] = {
    "TSMC": "TSM",
}


@dataclass
class StockRow:
    idx: int
    purchaser: str
    ticker: str
    currency: str
    original_quant_any: float
    original_quant_sgd: float
    original_price_any: float


def load_stock_sheet() -> pd.DataFrame:
    """
    Load the stock sheet using the confirmed header row.
    """
    # Header is at Excel row 3 (0-based index 2), but pandas expects
    # the header argument as "row number containing column names"
    # counting from 0 within the sheet data. Given the file structure,
    # using header=1 correctly picks up the Purchaser/... header row.
    df = pd.read_excel(EXCEL_PATH, sheet_name=STOCK_SHEET_NAME, header=1)
    print(df.columns)  # debug: confirm columns read correctly
    # Drop completely empty rows if any
    df = df.dropna(how="all")
    return df


def extract_rows(df: pd.DataFrame) -> List[StockRow]:
    rows: List[StockRow] = []
    for i, r in df.iterrows():
        ticker = str(r.get("Ticker") or "").strip()
        if not ticker:
            continue
        rows.append(
            StockRow(
                idx=i,
                purchaser=str(r.get("Purchaser") or "").strip(),
                ticker=ticker,
                currency=str(r.get("Currency") or "").strip(),
                original_quant_any=float(r.get("Original Purchase Quantum(Any $)", 0.0) or 0.0),
                original_quant_sgd=float(r.get("Original Purchase Quantum(S$)", 0.0) or 0.0),
                original_price_any=float(r.get("Original Purchase Price (Any $)", 0.0) or 0.0),
            )
        )
    return rows


def fetch_latest_prices(tickers: List[str]) -> Dict[str, float]:
    """
    Fetch latest close prices for all tickers via yfinance.
    Returns a mapping ticker -> last close price.
    """
    unique = sorted({t for t in tickers if t})
    prices: Dict[str, float] = {}
    for symbol in unique:
        # Apply mapping for tickers like TSMC -> TSM
        yf_symbol = TICKER_MAP.get(symbol, symbol)
        try:
            # Special handling for 0P0001Q0TW.SI which may not have 1d data
            if yf_symbol == "0P0001Q0TW.SI":
                data = yf.download(yf_symbol, period="5d", interval="1d", progress=False)
            else:
                data = yf.download(yf_symbol, period="1d", interval="1d", progress=False)
            if not data.empty:
                close = float(data["Close"].dropna().iloc[-1])
                prices[symbol] = close
        except Exception:
            # If a ticker fails, skip it; caller can decide how to handle missing prices.
            continue
    return prices


def recompute_row_values(row: StockRow, new_price_any: float) -> Dict[str, float]:
    """
    Given a stock row and a new price in the original currency, compute:
    - Holding Price (Any $)
    - Gross Holding Value (S$)
    - Net Earning/Loss (S$)
    based on the observed data model in the Excel file.
    """
    if row.original_price_any <= 0 or row.original_quant_any <= 0 or row.original_quant_sgd <= 0:
        return {}

    # Approximate units purchased
    units = row.original_quant_any / row.original_price_any

    # Effective FX rate from original quantum any$ -> SGD
    fx_rate = row.original_quant_sgd / row.original_quant_any

    gross_value_sgd = units * new_price_any * fx_rate
    net_pnl_sgd = gross_value_sgd - row.original_quant_sgd

    return {
        "Holding Price (Any $)": new_price_any,
        "Gross Holding Value (S$)": gross_value_sgd,
        "Net Earning/Loss (S$)": net_pnl_sgd,
    }


def update_prices() -> None:
    """
    Main entry point: load the stock sheet, fetch latest prices from Yahoo Finance,
    update the holding price and related SGD fields, and write back to Excel.
    """
    df = load_stock_sheet()
    rows = extract_rows(df)
    if not rows:
        return

    prices = fetch_latest_prices([r.ticker for r in rows])

    for row in rows:
        new_price = prices.get(row.ticker)
        if new_price is None:
            continue
        updates = recompute_row_values(row, new_price)
        if not updates:
            continue
        for col, val in updates.items():
            if col in df.columns:
                df.at[row.idx, col] = val

    # Write back to the same Excel file, preserving other sheets
    book = pd.read_excel(EXCEL_PATH, sheet_name=None, header=None)
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="w") as writer:
        for sheet_name, sheet_df in book.items():
            if sheet_name == STOCK_SHEET_NAME:
                # We need to restore the original structure: first two header rows, then data.
                # Re-read stock sheet raw (no header) to get the top rows, then append updated data.
                raw = sheet_df
                header_rows = raw.iloc[:HEADER_ROW_INDEX, :]
                # align columns with raw: we overwrite from HEADER_ROW_INDEX onwards
                updated = df.reindex(columns=raw.columns)
                combined = pd.concat([header_rows, updated], ignore_index=True)
                combined.to_excel(writer, sheet_name=sheet_name, header=False, index=False)
            else:
                sheet_df.to_excel(writer, sheet_name=sheet_name, header=False, index=False)


if __name__ == "__main__":
    # Run a one-off update when the script is executed directly.
    print(f"[{dt.datetime.now()}] Running price updater...")
    update_prices()
    print(f"[{dt.datetime.now()}] Price update complete.")

