import asyncio
import hmac
import os
from contextlib import asynccontextmanager

import anthropic
import httpx
from fastapi import FastAPI, Header, HTTPException, Request

import portfolio_db as db
from config import validate_anthropic_env


@asynccontextmanager
async def lifespan(_app: FastAPI):
    validate_anthropic_env()
    yield


app = FastAPI(title="FunFamily Stock Bot", lifespan=lifespan)

SYSTEM_PROMPT = (
    "You are the assistant for a private family investment portfolio. "
    "Use tools for all portfolio facts and changes; never invent figures. "
    "Keep Telegram replies concise and mobile-friendly. All totals are SGD unless stated."
)

TOOLS = [
    {"name": "get_portfolio_summary", "description": "Get the family portfolio summary.",
     "input_schema": {"type": "object", "properties": {}, "required": []}},
    {"name": "get_member_holdings", "description": "Get one member's detailed holdings.",
     "input_schema": {"type": "object", "properties": {"member": {"type": "string"}}, "required": ["member"]}},
    {"name": "add_stock_position", "description": "Add a completed stock purchase.",
     "input_schema": {"type": "object", "properties": {
         "member": {"type": "string"}, "ticker": {"type": "string"},
         "purchase_date": {"type": "string", "description": "YYYY-MM-DD"},
         "amount_any": {"type": "number"}, "currency": {"type": "string"},
         "amount_sgd": {"type": "number"}, "price_any": {"type": "number"},
         "platform": {"type": "string"}, "product": {"type": "string"},
     }, "required": ["member", "ticker", "purchase_date", "amount_any", "currency", "amount_sgd", "price_any", "platform"]}},
    {"name": "update_stock_price", "description": "Set the current price for every position with a ticker.",
     "input_schema": {"type": "object", "properties": {"ticker": {"type": "string"}, "price": {"type": "number"}}, "required": ["ticker", "price"]}},
    {"name": "refresh_all_prices", "description": "Refresh all prices from Yahoo Finance.",
     "input_schema": {"type": "object", "properties": {}, "required": []}},
]


def required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"{name} is not configured")
    return value


def allowed_chat(chat_id: int) -> bool:
    raw = os.getenv("TELEGRAM_ALLOWED_CHAT_IDS") or os.getenv("TELEGRAM_DAILY_CHAT_IDS", "")
    allowed = {int(value.strip()) for value in raw.split(",") if value.strip()}
    return chat_id in allowed


async def telegram_request(method: str, payload: dict) -> dict:
    token = required_env("TELEGRAM_BOT_TOKEN")
    async with httpx.AsyncClient(timeout=30) as client:
        response = await client.post(f"https://api.telegram.org/bot{token}/{method}", json=payload)
        response.raise_for_status()
        result = response.json()
        if not result.get("ok"):
            raise RuntimeError(result.get("description", "Telegram API request failed"))
        return result


async def send_message(chat_id: int, text: str) -> None:
    await telegram_request("sendMessage", {"chat_id": chat_id, "text": text[:4096]})


def execute_tool(name: str, values: dict) -> str:
    if name == "get_portfolio_summary":
        return db.portfolio_summary()
    if name == "get_member_holdings":
        return db.portfolio_summary(values["member"], detailed=True)
    if name == "add_stock_position":
        position_id = db.add_position(values)
        return f"Added position #{position_id} for {values['member']}: {values['ticker'].upper()}."
    if name == "update_stock_price":
        count = db.update_ticker_price(values["ticker"], float(values["price"]))
        return f"Updated {count} position(s) for {values['ticker'].upper()}."
    if name == "refresh_all_prices":
        count, missing = db.refresh_prices()
        suffix = f" Missing: {', '.join(missing)}." if missing else ""
        return f"Refreshed {count} position(s).{suffix}"
    raise ValueError(f"Unknown tool: {name}")


async def natural_language_reply(text: str) -> str:
    api_key, model = validate_anthropic_env()
    client = anthropic.Anthropic(api_key=api_key)
    messages = [{"role": "user", "content": text}]
    response = await asyncio.to_thread(
        client.messages.create, model=model, max_tokens=1024,
        system=SYSTEM_PROMPT, tools=TOOLS, messages=messages,
    )
    while response.stop_reason == "tool_use":
        messages.append({"role": "assistant", "content": response.content})
        results = []
        for block in response.content:
            if block.type == "tool_use":
                result = await asyncio.to_thread(execute_tool, block.name, block.input)
                results.append({"type": "tool_result", "tool_use_id": block.id, "content": result})
        messages.append({"role": "user", "content": results})
        response = await asyncio.to_thread(
            client.messages.create, model=model, max_tokens=1024,
            system=SYSTEM_PROMPT, tools=TOOLS, messages=messages,
        )
    return "".join(block.text for block in response.content if block.type == "text") or "Done."


async def command_reply(text: str) -> str:
    command, *arguments = text.strip().split()
    command = command.split("@", 1)[0].lower()
    if command in {"/start", "/help"}:
        return (
            "/portfolio - family summary\n/mystocks <name> - member holdings\n"
            "/pnl - profit and loss\n/update <ticker> <price> - set a price\n"
            "/refresh - refresh Yahoo prices\n/remove <position-id> confirm - delete a position\n"
            "You can also ask portfolio questions in natural language."
        )
    if command in {"/portfolio", "/pnl"}:
        return await asyncio.to_thread(db.portfolio_summary)
    if command == "/mystocks":
        if not arguments:
            return "Usage: /mystocks <name>"
        return await asyncio.to_thread(db.portfolio_summary, " ".join(arguments), True)
    if command == "/update":
        if len(arguments) != 2:
            return "Usage: /update <ticker> <price>"
        count = await asyncio.to_thread(db.update_ticker_price, arguments[0], float(arguments[1]))
        return f"Updated {count} position(s) for {arguments[0].upper()}."
    if command == "/refresh":
        count, missing = await asyncio.to_thread(db.refresh_prices)
        return f"Refreshed {count} position(s)." + (f" Missing: {', '.join(missing)}." if missing else "")
    if command == "/remove":
        if len(arguments) != 2 or arguments[1].lower() != "confirm":
            return "Usage: /remove <position-id> confirm. Find IDs with /mystocks <name>."
        removed = await asyncio.to_thread(db.remove_position, int(arguments[0]))
        return "Position removed." if removed else "Position ID not found."
    return "Unknown command. Use /help."


async def handle_message(message: dict) -> None:
    chat_id = int(message["chat"]["id"])
    if not allowed_chat(chat_id):
        return
    text = message.get("text", "").strip()
    if not text:
        return
    try:
        reply = await command_reply(text) if text.startswith("/") else await natural_language_reply(text)
    except (TypeError, ValueError) as exc:
        reply = f"I couldn't understand that value: {exc}"
    await send_message(chat_id, reply)


@app.get("/api/health")
async def health():
    await asyncio.to_thread(db.ensure_schema)
    return {"ok": True}


@app.post("/api/telegram")
async def telegram_webhook(request: Request, x_telegram_bot_api_secret_token: str | None = Header(default=None)):
    expected = required_env("TELEGRAM_WEBHOOK_SECRET")
    if not x_telegram_bot_api_secret_token or not hmac.compare_digest(x_telegram_bot_api_secret_token, expected):
        raise HTTPException(status_code=401, detail="Invalid webhook secret")
    update = await request.json()
    await asyncio.to_thread(db.ensure_schema)
    update_id = int(update["update_id"])
    if not await asyncio.to_thread(db.claim_telegram_update, update_id):
        return {"ok": True, "duplicate": True}
    message = update.get("message") or update.get("edited_message")
    try:
        if message:
            await handle_message(message)
    except Exception:
        await asyncio.to_thread(db.release_telegram_update, update_id)
        raise
    return {"ok": True}


@app.get("/api/daily-update")
async def daily_update(authorization: str | None = Header(default=None)):
    expected = f"Bearer {required_env('CRON_SECRET')}"
    if not authorization or not hmac.compare_digest(authorization, expected):
        raise HTTPException(status_code=401, detail="Invalid cron secret")
    await asyncio.to_thread(db.ensure_schema)
    updated, missing = await asyncio.to_thread(db.refresh_prices)
    report = await asyncio.to_thread(db.daily_report)
    if missing:
        report += f"\n\nPrice refresh unavailable: {', '.join(missing)}"
    chat_ids = [int(value.strip()) for value in required_env("TELEGRAM_DAILY_CHAT_IDS").split(",") if value.strip()]
    await asyncio.gather(*(send_message(chat_id, report) for chat_id in chat_ids))
    return {"ok": True, "updated_positions": updated, "chats": len(chat_ids)}
