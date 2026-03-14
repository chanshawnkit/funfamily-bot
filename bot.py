import logging
import os
from datetime import date
from dotenv import load_dotenv
from telegram import ReplyKeyboardMarkup, ReplyKeyboardRemove, Update
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)
from price_updater import update_prices

import pandas as pd
from openpyxl import load_workbook as openpyxl_load

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "Funfamily_Stock_Tracker.xlsx")
STOCK_SHEET  = "Stock Sheet"
DEPOSIT_SHEET = "Deposit Sheet"

HEADER_ROW = 2  # Excel row containing column headers for Stock Sheet

SELECT_MEMBER, ENTER_TICKER, ENTER_DATE, ENTER_AMOUNT_ANY, SELECT_CURRENCY, ENTER_AMOUNT_SGD, SELECT_PLATFORM = range(7)

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


async def addstock_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    df = load_stock()
    members = sorted(df["Purchaser"].dropna().unique().tolist())
    if not members:
        await update.message.reply_text(
            "No existing members found in the stock sheet. Please add at least one row in Excel first."
        )
        return ConversationHandler.END

    keyboard = [[m] for m in members]
    reply_markup = ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)

    context.user_data["addstock"] = {}
    await update.message.reply_text(
        "Please select a member:", reply_markup=reply_markup
    )
    return SELECT_MEMBER


async def addstock_member(update: Update, context: ContextTypes.DEFAULT_TYPE):
    member = (update.message.text or "").strip()
    if not member:
        await update.message.reply_text("Please choose a member from the list.")
        return SELECT_MEMBER

    context.user_data.setdefault("addstock", {})["member"] = member
    await update.message.reply_text(
        "Please enter the ticker symbol (e.g. AAPL, 0P0001Q0TW.SI):",
        reply_markup=ReplyKeyboardRemove(),
    )
    return ENTER_TICKER


async def addstock_ticker(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ticker = (update.message.text or "").strip()
    if not ticker:
        await update.message.reply_text("Ticker cannot be empty. Please enter a ticker:")
        return ENTER_TICKER

    context.user_data.setdefault("addstock", {})["ticker"] = ticker
    await update.message.reply_text(
        "Please enter the purchase date in YYYY-MM-DD format (e.g. 2026-03-15):"
    )
    return ENTER_DATE


async def addstock_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    date_text = (update.message.text or "").strip()
    try:
        purchase_date = date.fromisoformat(date_text)
    except ValueError:
        await update.message.reply_text(
            "Date must be in YYYY-MM-DD format, e.g. 2026-03-15. Please try again:"
        )
        return ENTER_DATE

    context.user_data.setdefault("addstock", {})["purchase_date"] = purchase_date
    await update.message.reply_text(
        "Please enter the original purchase quantum (Any $), e.g. 5000:"
    )
    return ENTER_AMOUNT_ANY


async def addstock_amount_any(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (update.message.text or "").replace(",", "").strip()
    try:
        amount_any = float(text)
        if amount_any <= 0:
            raise ValueError
    except ValueError:
        await update.message.reply_text(
            "Amount must be a positive number, e.g. 5000 or 5000.50. Please try again:"
        )
        return ENTER_AMOUNT_ANY

    context.user_data.setdefault("addstock", {})["amount_any"] = amount_any

    keyboard = [["SGD", "USD"]]
    reply_markup = ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text(
        "Please select the currency:", reply_markup=reply_markup
    )
    return SELECT_CURRENCY


async def addstock_currency(update: Update, context: ContextTypes.DEFAULT_TYPE):
    currency = (update.message.text or "").strip().upper()
    allowed = {"SGD", "USD"}
    if currency not in allowed:
        await update.message.reply_text(
            "Currency must be SGD or USD. Please tap one of the buttons:"
        )
        return SELECT_CURRENCY

    context.user_data.setdefault("addstock", {})["currency"] = currency
    await update.message.reply_text(
        "Please enter the original purchase quantum in SGD, e.g. 5000:",
        reply_markup=ReplyKeyboardRemove(),
    )
    return ENTER_AMOUNT_SGD


async def addstock_amount_sgd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (update.message.text or "").replace(",", "").strip()
    try:
        amount_sgd = float(text)
        if amount_sgd <= 0:
            raise ValueError
    except ValueError:
        await update.message.reply_text(
            "Amount in SGD must be a positive number, e.g. 5000 or 5000.50. Please try again:"
        )
        return ENTER_AMOUNT_SGD

    context.user_data.setdefault("addstock", {})["amount_sgd"] = amount_sgd

    df = load_stock()
    platforms = sorted(df["Platform"].dropna().unique().tolist())
    if platforms:
        keyboard = [[p] for p in platforms]
        reply_markup = ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)
        await update.message.reply_text(
            "Please select the platform (or type a new one):", reply_markup=reply_markup
        )
    else:
        await update.message.reply_text(
            "Please enter the platform (e.g. Endowus, IBKR):",
            reply_markup=ReplyKeyboardRemove(),
        )
    return SELECT_PLATFORM


async def addstock_platform(update: Update, context: ContextTypes.DEFAULT_TYPE):
    platform = (update.message.text or "").strip()
    if not platform:
        await update.message.reply_text(
            "Platform cannot be empty. Please enter or select a platform:"
        )
        return SELECT_PLATFORM

    data = context.user_data.setdefault("addstock", {})
    data["platform"] = platform

    try:
        append_stock_position(
            purchaser=data["member"],
            purchase_date=data["purchase_date"],
            amount_any=data["amount_any"],
            currency=data["currency"],
            amount_sgd=data["amount_sgd"],
            platform=data["platform"],
            ticker=data["ticker"],
        )
    except Exception as e:
        logger.exception("Failed to append stock position")
        await update.message.reply_text(
            f"Error adding stock position: {e}", reply_markup=ReplyKeyboardRemove()
        )
        return ConversationHandler.END

    await update.message.reply_text(
        (
            f"Added stock for {data['member']}: {data['ticker']} on "
            f"{data['purchase_date'].isoformat()} for "
            f"{data['amount_any']:.2f} {data['currency']} "
            f"({data['amount_sgd']:.2f} SGD) via {data['platform']}."
        ),
        reply_markup=ReplyKeyboardRemove(),
    )
    context.user_data.pop("addstock", None)
    return ConversationHandler.END


async def addstock_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("addstock", None)
    await update.message.reply_text(
        "Add stock operation cancelled.", reply_markup=ReplyKeyboardRemove()
    )
    return ConversationHandler.END


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
        "/addstock - Guided flow to add a new stock position\n"
        "/help - Show this message"
    )


def main():
    load_dotenv()
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    if not token:
        raise RuntimeError("TELEGRAM_BOT_TOKEN not set in .env")

    app = Application.builder().token(token).build()

    addstock_conv = ConversationHandler(
        entry_points=[CommandHandler("addstock", addstock_start)],
        states={
            SELECT_MEMBER: [MessageHandler(filters.TEXT & ~filters.COMMAND, addstock_member)],
            ENTER_TICKER: [MessageHandler(filters.TEXT & ~filters.COMMAND, addstock_ticker)],
            ENTER_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, addstock_date)],
            ENTER_AMOUNT_ANY: [MessageHandler(filters.TEXT & ~filters.COMMAND, addstock_amount_any)],
            SELECT_CURRENCY: [MessageHandler(filters.TEXT & ~filters.COMMAND, addstock_currency)],
            ENTER_AMOUNT_SGD: [MessageHandler(filters.TEXT & ~filters.COMMAND, addstock_amount_sgd)],
            SELECT_PLATFORM: [MessageHandler(filters.TEXT & ~filters.COMMAND, addstock_platform)],
        },
        fallbacks=[CommandHandler("cancel", addstock_cancel)],
    )

    app.add_handler(CommandHandler("start",     start_command))
    app.add_handler(CommandHandler("help",      help_command))
    app.add_handler(CommandHandler("portfolio", portfolio_command))
    app.add_handler(CommandHandler("mystocks",  mystocks_command))
    app.add_handler(CommandHandler("pnl",       pnl_command))
    app.add_handler(CommandHandler("update",    update_command))
    app.add_handler(CommandHandler("refresh",   refresh_command))
    app.add_handler(addstock_conv)

    logger.info("Bot started. Listening for commands...")
    app.run_polling()


if __name__ == "__main__":
    main()
