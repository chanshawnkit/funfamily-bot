import logging
import os
from datetime import date
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

HEADER_ROW = 2  # Excel row containing column headers for Stock Sheet

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


def append_stock_position(
    purchaser: str,
    purchase_date: date | str,
    amount_any: float,
    currency: str,
    amount_sgd: float,
    platform: str,
    ticker: str,
    product: str | None = None,
    reference: str | None = None,
) -> None:
    """
    Append a new stock position as a row in the Stock Sheet without modifying
    existing data or headers.
    """
    wb = openpyxl_load(EXCEL_PATH)
    if STOCK_SHEET not in wb.sheetnames:
        raise RuntimeError(f"Sheet '{STOCK_SHEET}' not found in workbook")

    ws = wb[STOCK_SHEET]

    # Build a mapping from header name -> column index based on HEADER_ROW.
    header_map: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=HEADER_ROW, column=col).value
        if val is None:
            continue
        header_name = str(val).strip()
        if header_name:
            header_map[header_name] = col

    required_headers = [
        "Purchaser",
        "Date of Purchase",
        "Original Purchase Quantum(Any $)",
        "Currency",
        "Original Purchase Quantum(S$)",
        "Platform",
        "Product",
        "Ticker",
        "Reference",
        "Original Purchase Price (Any $)",
        "Holding Price (Any $)",
        "Gross Holding Value (S$)",
        "Net Earning/Loss (S$)",
    ]

    missing = [h for h in required_headers if h not in header_map]
    if missing:
        raise RuntimeError(f"Missing expected columns in Stock Sheet: {', '.join(missing)}")

    # Determine the next empty row after existing data.
    next_row = ws.max_row + 1

    def set_by_header(header: str, value):
        col_idx = header_map.get(header)
        if col_idx is not None:
            ws.cell(row=next_row, column=col_idx).value = value

    # Normalize date to a format consistent with existing rows.
    if isinstance(purchase_date, date):
        date_value = purchase_date
    else:
        date_value = str(purchase_date)

    set_by_header("Purchaser", purchaser)
    set_by_header("Date of Purchase", date_value)
    set_by_header("Original Purchase Quantum(Any $)", float(amount_any))
    set_by_header("Currency", currency)
    set_by_header("Original Purchase Quantum(S$)", float(amount_sgd))
    set_by_header("Platform", platform)
    set_by_header("Product", product or "")
    set_by_header("Ticker", ticker)
    set_by_header("Reference", reference or "")

    # Leave price/valuation columns blank initially; price updater will fill them.
    set_by_header("Original Purchase Price (Any $)", None)
    set_by_header("Holding Price (Any $)", None)
    set_by_header("Gross Holding Value (S$)", None)
    set_by_header("Net Earning/Loss (S$)", None)

    wb.save(EXCEL_PATH)


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


async def addstock_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    /addstock <member> <ticker> <date> <amount_any> <currency> <amount_sgd> <platform>
    Example:
      /addstock "Shawn Mun" 0P0001Q0TW.SI 2026-03-15 5000 SGD 5000 Endowus
    """
    args = context.args or []
    if len(args) < 7:
        await update.message.reply_text(
            "Usage: /addstock <member> <ticker> <date> <amount_any> <currency> <amount_sgd> <platform>\n"
            'Example: /addstock "Shawn Mun" 0P0001Q0TW.SI 2026-03-15 5000 SGD 5000 Endowus'
        )
        return

    # Support quoted member names with spaces, e.g. "Shawn Mun"
    member_end_idx = 0
    if args[0].startswith('"'):
        for i, token in enumerate(args):
            if token.endswith('"'):
                member_end_idx = i
                break
        else:
            await update.message.reply_text(
                "Could not parse member name. If it contains spaces, wrap it in quotes."
            )
            return
        member_tokens = args[0 : member_end_idx + 1]
        member = " ".join(member_tokens).strip('"').strip()
        remaining = args[member_end_idx + 1 :]
    else:
        member = args[0]
        remaining = args[1:]

    if len(remaining) < 6:
        await update.message.reply_text(
            "Usage: /addstock <member> <ticker> <date> <amount_any> <currency> <amount_sgd> <platform>\n"
            'Example: /addstock "Shawn Mun" 0P0001Q0TW.SI 2026-03-15 5000 SGD 5000 Endowus'
        )
        return

    ticker = remaining[0]
    date_str = remaining[1]
    amount_any_str = remaining[2]
    currency = remaining[3].upper()
    amount_sgd_str = remaining[4]
    platform = " ".join(remaining[5:]).strip()

    # Basic validation
    try:
        purchase_date = date.fromisoformat(date_str)
    except ValueError:
        await update.message.reply_text(
            "Date must be in YYYY-MM-DD format, e.g. 2026-03-15."
        )
        return

    try:
        amount_any = float(amount_any_str)
        amount_sgd = float(amount_sgd_str)
    except ValueError:
        await update.message.reply_text(
            "Amounts must be numbers, e.g. 5000 or 5000.50."
        )
        return

    allowed_currencies = {"SGD", "USD"}
    if currency not in allowed_currencies:
        await update.message.reply_text(
            f"Currency must be one of: {', '.join(sorted(allowed_currencies))}."
        )
        return

    try:
        append_stock_position(
            purchaser=member,
            purchase_date=purchase_date,
            amount_any=amount_any,
            currency=currency,
            amount_sgd=amount_sgd,
            platform=platform,
            ticker=ticker,
        )
    except Exception as e:
        logger.exception("Failed to append stock position")
        await update.message.reply_text(f"Error adding stock position: {e}")
        return

    await update.message.reply_text(
        f"Added stock for {member}: {ticker} on {purchase_date.isoformat()} "
        f"for {amount_any:.2f} {currency} ({amount_sgd:.2f} SGD) via {platform}."
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
        "/addstock <member> <ticker> <date> <amount_any> <currency> <amount_sgd> <platform> - Add a new stock position\n"
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
    app.add_handler(CommandHandler("addstock",  addstock_command))

    logger.info("Bot started. Listening for commands...")
    app.run_polling()


if __name__ == "__main__":
    main()
