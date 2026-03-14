import logging
import os
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
)
from price_updater import update_prices

import pandas as pd
from openpyxl import load_workbook as openpyxl_load

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "Funfamily_Stock_Tracker.xlsx")
STOCK_SHEET  = "Stock Sheet"
DEPOSIT_SHEET = "Deposit Sheet"

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


def load_stock() -> pd.DataFrame:
    df = pd.read_excel(EXCEL_PATH, sheet_name=STOCK_SHEET, header=1)
    df = df.dropna(how="all")
    df = df[df["Ticker"].notna()]
    return df


def fmt_sgd(val: float) -> str:
    return f"S${val:,.2f}"


def fmt_pct(val: float) -> str:
    sign = "+" if val >= 0 else ""
    return f"{sign}{val:.2f}%"


async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(help_text())


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(help_text())


async def portfolio_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    df = load_stock()
    grouped = df.groupby("Purchaser").agg(
        invested=("Original Purchase Quantum(S$)", "sum"),
        value=("Gross Holding Value (S$)", "sum"),
        pnl=("Net Earning/Loss (S$)", "sum"),
    ).reset_index()

    lines = ["*Family Portfolio*\n"]
    total_inv = total_val = total_pnl = 0.0

    for _, r in grouped.iterrows():
        pct = (r.pnl / r.invested * 100) if r.invested else 0
        lines.append(
            f"*{r.Purchaser}*\n"
            f"Invested: {fmt_sgd(r.invested)}\n"
            f"Value:    {fmt_sgd(r.value)}\n"
            f"P&L:      {fmt_sgd(r.pnl)} ({fmt_pct(pct)})\n"
        )
        total_inv += r.invested
        total_val += r.value
        total_pnl += r.pnl

    total_pct = (total_pnl / total_inv * 100) if total_inv else 0
    lines.append(
        f"*Total*\n"
        f"Invested: {fmt_sgd(total_inv)}\n"
        f"Value:    {fmt_sgd(total_val)}\n"
        f"P&L:      {fmt_sgd(total_pnl)} ({fmt_pct(total_pct)})"
    )

    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def mystocks_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not context.args:
        await update.message.reply_text("Usage: /mystocks <name>\nExample: /mystocks Jasmine")
        return

    member = " ".join(context.args).strip()
    df = load_stock()
    rows = df[df["Purchaser"].str.strip().str.lower() == member.lower()]

    if rows.empty:
        await update.message.reply_text(f"No holdings found for {member}.")
        return

    lines = [f"*{member}'s Holdings*\n"]
    for _, r in rows.iterrows():
        qty   = float(r.get("Original Purchase Quantum(Any $)", 0) or 0)
        price = float(r.get("Original Purchase Price (Any $)", 0) or 0)
        units = qty / price if price > 0 else 0
        hold  = float(r.get("Holding Price (Any $)", 0) or 0)
        val   = float(r.get("Gross Holding Value (S$)", 0) or 0)
        inv   = float(r.get("Original Purchase Quantum(S$)", 0) or 0)
        pnl   = float(r.get("Net Earning/Loss (S$)", 0) or 0)
        pct   = (pnl / inv * 100) if inv else 0

        lines.append(
            f"*{r['Ticker']}* ({r['Platform']})\n"
            f"Units: {units:.3f}\n"
            f"Price: {hold:.4f} {r['Currency']}\n"
            f"Value: {fmt_sgd(val)}\n"
            f"P&L:   {fmt_sgd(pnl)} ({fmt_pct(pct)})\n"
        )

    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def pnl_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    df = load_stock()
    grouped = df.groupby("Purchaser").agg(
        invested=("Original Purchase Quantum(S$)", "sum"),
        pnl=("Net Earning/Loss (S$)", "sum"),
    ).reset_index()

    lines = ["*P&L by Member*\n"]
    for _, r in grouped.iterrows():
        pct = (r.pnl / r.invested * 100) if r.invested else 0
        lines.append(
            f"*{r.Purchaser}*\n"
            f"P&L: {fmt_sgd(r.pnl)} ({fmt_pct(pct)})\n"
        )

    total_pnl = grouped["pnl"].sum()
    total_inv = grouped["invested"].sum()
    total_pct = (total_pnl / total_inv * 100) if total_inv else 0
    lines.append(f"*Total P&L*\n{fmt_sgd(total_pnl)} ({fmt_pct(total_pct)})")

    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def update_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if len(context.args) != 2:
        await update.message.reply_text("Usage: /update TICKER PRICE\nExample: /update MSFT 395.50")
        return
    ticker = context.args[0].upper()
    try:
        price = float(context.args[1])
    except ValueError:
        await update.message.reply_text("Price must be a number, e.g. 395.50")
        return

    wb = openpyxl_load(EXCEL_PATH)
    ws = wb[STOCK_SHEET]

    updated = 0
    row = 3
    while True:
        cell_ticker = ws.cell(row=row, column=8).value
        if not cell_ticker:
            break
        if str(cell_ticker).strip().upper() == ticker:
            orig_qty  = float(ws.cell(row=row, column=3).value or 0)
            orig_sgd  = float(ws.cell(row=row, column=5).value or 0)
            orig_price = float(ws.cell(row=row, column=10).value or 0)
            if orig_price > 0 and orig_qty > 0 and orig_sgd > 0:
                units = orig_qty / orig_price
                fx    = orig_sgd / orig_qty
                gross = units * price * fx
                pnl   = gross - orig_sgd
                ws.cell(row=row, column=11).value = round(price, 4)
                ws.cell(row=row, column=12).value = round(gross, 6)
                ws.cell(row=row, column=13).value = round(pnl, 6)
                updated += 1
        row += 1

    if updated == 0:
        await update.message.reply_text(f"Ticker {ticker} not found.")
        return

    wb.save(EXCEL_PATH)
    await update.message.reply_text(
        f"Updated {ticker} to {price:.4f}. {updated} row(s) saved."
    )


async def refresh_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Fetching latest prices, please wait...")
    try:
        update_prices()
        await update.message.reply_text("Prices updated successfully.")
    except Exception as e:
        await update.message.reply_text(f"Error updating prices: {e}")


def help_text() -> str:
    return (
        "/portfolio - Full family portfolio summary\n"
        "/mystocks <name> - One member's holdings\n"
        "/pnl - Profit & loss per member\n"
        "/update <ticker> <price> - Manually set a price\n"
        "/refresh - Fetch latest prices from Yahoo Finance\n"
        "/help - Show this message"
    )


def main():
    load_dotenv()
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    if not token:
        raise RuntimeError("TELEGRAM_BOT_TOKEN not set in .env")

    app = Application.builder().token(token).build()

    app.add_handler(CommandHandler("start",     start_command))
    app.add_handler(CommandHandler("help",      help_command))
    app.add_handler(CommandHandler("portfolio", portfolio_command))
    app.add_handler(CommandHandler("mystocks",  mystocks_command))
    app.add_handler(CommandHandler("pnl",       pnl_command))
    app.add_handler(CommandHandler("update",    update_command))
    app.add_handler(CommandHandler("refresh",   refresh_command))

    logger.info("Bot started. Listening for commands...")
    app.run_polling()


if __name__ == "__main__":
    main()
