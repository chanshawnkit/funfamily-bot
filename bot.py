import logging
import os
from typing import Dict, List, Tuple

import pandas as pd
from dotenv import load_dotenv
from telegram import ParseMode, Update
from telegram.ext import (
    CallbackContext,
    CommandHandler,
    Filters,
    MessageHandler,
    Updater,
)

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "Funfamily_Stock_Tracker.xlsx")
STOCK_SHEET_NAME = "Stock Sheet"
DEPOSIT_SHEET_NAME = "Deposit Sheet"
HEADER_ROW_INDEX = 2  # zero-based row index where real headers start

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


def load_workbook() -> Dict[str, pd.DataFrame]:
    """
    Load both sheets.
    """
    # Stock sheet with proper headers
    stock = pd.read_excel(
        EXCEL_PATH,
        sheet_name=STOCK_SHEET_NAME,
        header=HEADER_ROW_INDEX,
    ).dropna(how="all")

    # Deposit sheet raw to capture both the summary and transaction areas
    deposit_raw = pd.read_excel(EXCEL_PATH, sheet_name=DEPOSIT_SHEET_NAME, header=None)

    return {"stock": stock, "deposit_raw": deposit_raw}


def get_members(stock_df: pd.DataFrame) -> List[str]:
    members = sorted(
        {str(x).strip() for x in stock_df["Purchaser"].dropna().unique() if str(x).strip()}
    )
    return members


def aggregate_member_pnl(stock_df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregate per-member invested, market value, and P&L using stock rows only.
    All values are in SGD, per .cursorrules.
    """
    grouped = (
        stock_df.groupby("Purchaser", dropna=True)
        .agg(
            invested_sgd=("Original Purchase Quantum(S$)", "sum"),
            market_value_sgd=("Gross Holding Value (S$)", "sum"),
            pnl_sgd=("Net Earning/Loss (S$)", "sum"),
        )
        .reset_index()
    )
    grouped["pnl_pct"] = grouped.apply(
        lambda r: (r["pnl_sgd"] / r["invested_sgd"] * 100.0) if r["invested_sgd"] else 0.0,
        axis=1,
    )
    return grouped


def format_currency(amount: float) -> str:
    return f"S${amount:,.2f}"


def format_pct(pct: float) -> str:
    sign = "+" if pct > 0 else ""
    return f"{sign}{pct:.2f}%"


def build_portfolio_text(stock_df: pd.DataFrame) -> str:
    summary = aggregate_member_pnl(stock_df)

    lines: List[str] = []
    lines.append("*Family Portfolio*")

    total_invested = summary["invested_sgd"].sum()
    total_value = summary["market_value_sgd"].sum()
    total_pnl = summary["pnl_sgd"].sum()
    total_pnl_pct = (total_pnl / total_invested * 100.0) if total_invested else 0.0

    for _, row in summary.iterrows():
        name = row["Purchaser"]
        lines.append(
            f"\n*{name}*\n"
            f"Invested: {format_currency(row['invested_sgd'])}\n"
            f"Value: {format_currency(row['market_value_sgd'])}\n"
            f"P&L: {format_currency(row['pnl_sgd'])} ({format_pct(row['pnl_pct'])})"
        )

    lines.append(
        f"\n*Total*\n"
        f"Invested: {format_currency(total_invested)}\n"
        f"Value: {format_currency(total_value)}\n"
        f"P&L: {format_currency(total_pnl)} ({format_pct(total_pnl_pct)})"
    )

    return "\n".join(lines)


def compute_units(row: pd.Series) -> float:
    """
    Approximate units from amount and price in original currency.
    """
    amt = float(row.get("Original Purchase Quantum(Any $)", 0.0) or 0.0)
    price = float(row.get("Original Purchase Price (Any $)", 0.0) or 0.0)
    if price <= 0:
        return 0.0
    return amt / price


def build_member_positions_text(stock_df: pd.DataFrame, member: str) -> str:
    f = stock_df[stock_df["Purchaser"].astype(str).str.strip().str.lower() == member.lower()]
    if f.empty:
        return f"No holdings found for *{member}*."

    lines: List[str] = []
    lines.append(f"*{member}'s Holdings*")

    for _, r in f.iterrows():
        ticker = str(r["Ticker"])
        platform = str(r["Platform"])
        product = str(r["Product"])
        currency = str(r["Currency"])
        units = compute_units(r)
        price = float(r.get("Holding Price (Any $)", 0.0) or 0.0)
        market_sgd = float(r.get("Gross Holding Value (S$)", 0.0) or 0.0)
        invested_sgd = float(r.get("Original Purchase Quantum(S$)", 0.0) or 0.0)
        pnl_sgd = float(r.get("Net Earning/Loss (S$)", 0.0) or 0.0)
        pnl_pct = (pnl_sgd / invested_sgd * 100.0) if invested_sgd else 0.0

        lines.append(
            f"\n*{ticker}* ({platform})\n"
            f"{product}\n"
            f"Units: {units:.3f} ({currency})\n"
            f"Price: {price:.4f} {currency}\n"
            f"Value: {format_currency(market_sgd)}\n"
            f"P&L: {format_currency(pnl_sgd)} ({format_pct(pnl_pct)})"
        )

    return "\n".join(lines)


def parse_update_args(args: List[str]) -> Tuple[str, float]:
    if len(args) != 2:
        raise ValueError("Usage: /update TICKER PRICE")
    ticker = args[0].strip()
    try:
        price = float(args[1])
    except ValueError as exc:
        raise ValueError("Price must be a number, e.g. 395.50") from exc
    return ticker, price


def apply_manual_price_update(stock_df: pd.DataFrame, ticker: str, price: float) -> pd.DataFrame:
    """
    Update holding price and recompute value and P&L for all rows matching ticker.
    """
    mask = stock_df["Ticker"].astype(str).str.strip().str.upper() == ticker.upper()
    if not mask.any():
        raise ValueError(f"No rows found for ticker {ticker}")

    df = stock_df.copy()
    for idx, row in df[mask].iterrows():
        amt_any = float(row.get("Original Purchase Quantum(Any $)", 0.0) or 0.0)
        amt_sgd = float(row.get("Original Purchase Quantum(S$)", 0.0) or 0.0)
        orig_price_any = float(row.get("Original Purchase Price (Any $)", 0.0) or 0.0)
        if amt_any <= 0 or amt_sgd <= 0 or orig_price_any <= 0:
            continue

        units = amt_any / orig_price_any
        fx = amt_sgd / amt_any
        gross_sgd = units * price * fx
        pnl_sgd = gross_sgd - amt_sgd

        df.at[idx, "Holding Price (Any $)"] = price
        df.at[idx, "Gross Holding Value (S$)"] = gross_sgd
        df.at[idx, "Net Earning/Loss (S$)"] = pnl_sgd

    return df


def save_stock_sheet(updated_stock: pd.DataFrame) -> None:
    """
    Persist updated stock sheet back into the workbook, preserving other sheets
    and header rows.
    """
    book = pd.read_excel(EXCEL_PATH, sheet_name=None, header=None)
    raw_stock = book[STOCK_SHEET_NAME]
    header_rows = raw_stock.iloc[:HEADER_ROW_INDEX, :]
    updated_aligned = updated_stock.reindex(columns=raw_stock.columns)

    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="w") as writer:
        for sheet_name, sheet_df in book.items():
            if sheet_name == STOCK_SHEET_NAME:
                combined = pd.concat([header_rows, updated_aligned], ignore_index=True)
                combined.to_excel(writer, sheet_name=sheet_name, header=False, index=False)
            else:
                sheet_df.to_excel(writer, sheet_name=sheet_name, header=False, index=False)


def help_text() -> str:
    return (
        "/portfolio - Show full family portfolio summary\n"
        "/mystocks <name> - Show one member's holdings (e.g. /mystocks Jasmine)\n"
        "/update <ticker> <price> - Manually update a holding price (e.g. /update MSFT 395.50)\n"
        "/pnl - Show P&L per member\n"
        "/help - Show this help message"
    )


def start_command(update: Update, context: CallbackContext) -> None:
    update.message.reply_text(help_text())


def help_command(update: Update, context: CallbackContext) -> None:
    update.message.reply_text(help_text())


def portfolio_command(update: Update, context: CallbackContext) -> None:
    book = load_workbook()
    stock = book["stock"]
    text = build_portfolio_text(stock)
    update.message.reply_text(text, parse_mode=ParseMode.MARKDOWN)


def mystocks_command(update: Update, context: CallbackContext) -> None:
    if not context.args:
        update.message.reply_text("Usage: /mystocks <member name>\nExample: /mystocks Jasmine")
        return
    member = " ".join(context.args).strip()
    book = load_workbook()
    stock = book["stock"]
    text = build_member_positions_text(stock, member)
    update.message.reply_text(text, parse_mode=ParseMode.MARKDOWN)


def update_command(update: Update, context: CallbackContext) -> None:
    try:
        ticker, price = parse_update_args(context.args)
    except ValueError as e:
        update.message.reply_text(str(e))
        return

    book = load_workbook()
    stock = book["stock"]
    try:
        updated = apply_manual_price_update(stock, ticker, price)
    except ValueError as e:
        update.message.reply_text(str(e))
        return

    save_stock_sheet(updated)

    # Rebuild summary for feedback
    text = build_member_positions_text(updated, member="")  # not used; show a short confirmation instead
    update.message.reply_text(
        f"Updated {ticker} price to {price:.4f}. Excel file saved.", parse_mode=ParseMode.MARKDOWN
    )


def pnl_command(update: Update, context: CallbackContext) -> None:
    book = load_workbook()
    stock = book["stock"]
    summary = aggregate_member_pnl(stock)

    lines: List[str] = []
    lines.append("*P&L by Member*")
    for _, row in summary.iterrows():
        lines.append(
            f"\n*{row['Purchaser']}*\n"
            f"P&L: {format_currency(row['pnl_sgd'])} ({format_pct(row['pnl_pct'])})"
        )

    total_pnl = summary["pnl_sgd"].sum()
    total_invested = summary["invested_sgd"].sum()
    total_pct = (total_pnl / total_invested * 100.0) if total_invested else 0.0
    lines.append(
        f"\n*Total*\n"
        f"P&L: {format_currency(total_pnl)} ({format_pct(total_pct)})"
    )

    update.message.reply_text("\n".join(lines), parse_mode=ParseMode.MARKDOWN)


def unknown_command(update: Update, context: CallbackContext) -> None:
    update.message.reply_text("Unknown command. Type /help for available commands.")


def main() -> None:
    load_dotenv()
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    if not token:
        raise RuntimeError("TELEGRAM_BOT_TOKEN is not set in .env")

    updater = Updater(token=token, use_context=True)
    dp = updater.dispatcher

    dp.add_handler(CommandHandler("start", start_command))
    dp.add_handler(CommandHandler("help", help_command))
    dp.add_handler(CommandHandler("portfolio", portfolio_command))
    dp.add_handler(CommandHandler("mystocks", mystocks_command))
    dp.add_handler(CommandHandler("update", update_command))
    dp.add_handler(CommandHandler("pnl", pnl_command))
    dp.add_handler(MessageHandler(Filters.command, unknown_command))

    updater.start_polling()
    logger.info("Bot started. Listening for commands...")
    updater.idle()


if __name__ == "__main__":
    main()

