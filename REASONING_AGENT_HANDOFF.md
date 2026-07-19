# FunFamily Bot — Reasoning Agent Handoff

Last updated: 2026-07-19 (Asia/Singapore)

## Current outcome

The private family portfolio Telegram bot is live at `https://funfamily-bot.vercel.app`.
Production uses Telegram webhooks, a FastAPI Vercel Function, Supabase pooled
PostgreSQL, OpenRouter through the Anthropic Python SDK, and Vercel Cron.

Repository: private `chanshawnkit/funfamily-bot`, branch `main`.

Never put credentials, environment values, database passwords, or tokens in this
handoff, source code, logs, chat, or Git history.

## Production architecture

- `POST /api/telegram`
  - Validates `X-Telegram-Bot-Api-Secret-Token`.
  - Deduplicates Telegram update IDs in PostgreSQL.
  - Accepts messages only from `TELEGRAM_ALLOWED_CHAT_IDS`.
  - Uses the sender's positive Telegram user ID for trade authorization.
  - Leaves a persistent 👀 reaction as a best-effort acknowledgement.
  - Handles commands directly and natural language through Anthropic/OpenRouter.
- `GET /api/health`
  - Validates required startup configuration and checks/creates the DB schema.
- `GET /api/daily-update`
  - Requires `CRON_SECRET`, refreshes prices, and sends the report to
    `TELEGRAM_DAILY_CHAT_IDS`.

The stable Telegram webhook target is:

```text
https://funfamily-bot.vercel.app/api/telegram
```

## Environment variable names

Vercel Production and local ignored `.env` files use these names:

```text
ANTHROPIC_API_KEY
ANTHROPIC_MODEL
ANTHROPIC_BASE_URL
DATABASE_URL
TELEGRAM_BOT_TOKEN
TELEGRAM_WEBHOOK_SECRET
TELEGRAM_ALLOWED_CHAT_IDS
TELEGRAM_DAILY_CHAT_IDS
TELEGRAM_TRADE_ADMIN_USER_IDS
CRON_SECRET
PUBLIC_BASE_URL
DAILY_UPDATE_TIME
```

`DAILY_UPDATE_TIME` applies to local polling mode; production scheduling is in
`vercel.json` and evaluated in UTC. `portfolio_db.database_url()` accepts only
`DATABASE_URL` or `POSTGRES_URL` (exact case).

Do not inspect or print environment values. Vercel names may be checked with
`vercel env ls production`, which reports values as encrypted.

## Authorization and mutations

- Everyone in an allowed family chat can query portfolio information.
- Only sender IDs in `TELEGRAM_TRADE_ADMIN_USER_IDS` can add purchases or remove
  positions. Authorization uses Telegram's numeric `message.from.id`, never a
  spoofable display name.
- `/remove <position-id> confirm` is the deterministic deletion command.
- Natural-language deletion uses `remove_stock_position`; it requires a numeric
  position ID and explicit confirmation in the current message, and repeats the
  admin check at execution time.
- The application does not yet model a proper sale transaction with sale date,
  proceeds, and realized P&L. Removing a position is deletion, not sale accounting.

## Telegram presentation

- The bot calls `setMessageReaction` with 👀 and deliberately does not clear it.
- Reaction failures are non-fatal because group reaction settings may disallow it.
- Outbound text is sent with Telegram `parse_mode=HTML`.
- `telegram_html()` escapes arbitrary model output first and converts `**text**`
  to `<b>text</b>`, fixing literal Markdown markers without allowing HTML injection.

## Natural-language tools

The Vercel assistant exposes only the tools declared in `api/index.py`:

- portfolio summary
- member holdings
- add stock position (admin-only)
- remove stock position (admin-only and confirmation-required)
- update a ticker price
- refresh all Yahoo Finance prices

Tool authorization must remain enforced in Python, not only in the model prompt.

## Database

- Supabase PostgreSQL is the source of truth in production.
- `ensure_schema()` creates required tables and seeds an empty `positions` table
  from `Funfamily_Stock_Tracker.xlsx` using an advisory lock.
- `telegram_updates` provides webhook deduplication.
- Price refresh currently uses Yahoo Finance.

## Anthropic/OpenRouter

`config.validate_anthropic_env()` returns `(api_key, model, base_url)`.
`ANTHROPIC_BASE_URL` defaults to Anthropic and can point to
`https://openrouter.ai/api`. The model only sees tools supplied by this app; an MCP
connection in Claude Desktop or another Claude session is not inherited by API
calls made from Vercel.

## IBKR finding

There is currently no IBKR integration. Repository search found IBKR only as an
example value for the position `platform` field. There is no IBKR SDK/client, MCP
client, remote MCP URL, authentication configuration, or Vercel bridge.

Recommended future sequence:

1. Add read-only balances, positions, cash, and transaction history through a
   separately reachable and authenticated IBKR service.
2. Reconcile broker positions against the portfolio database with explicit source
   labeling and no automatic writes.
3. Consider trading only as a separate phase with admin authorization, an order
   preview, idempotency controls, and a second explicit confirmation.

An IBKR/TWS or Client Portal Gateway running only on a user's computer is not
reachable from Vercel without a secure, always-on bridge. Never expose a local
gateway directly to the public internet.

## Deployment and verification

Normal flow:

1. Run `python -m unittest discover -s tests -v`.
2. Run `git diff --check` and review the staged files.
3. Push `main`; Vercel Git integration creates a Production deployment.
4. Wait until the deployment is `Ready`.
5. Confirm `https://funfamily-bot.vercel.app` aliases the new deployment.
6. Confirm `/api/health` returns `{"ok":true}`.
7. Test a Telegram command, a natural-language query, an unauthorized mutation,
   and an authorized confirmed mutation.

Do not run local polling while the webhook is registered; Telegram supports only
one update-delivery mode at a time.

## Workspace hygiene

- `.env`, `.env.local`, `.env.vercel`, and `.vercel/` must remain ignored.
- `.env.example` contains placeholders only and remains tracked.
- Review `git status` carefully: pre-existing `.gitignore` changes may be unrelated,
  and no local environment file should ever be staged.
