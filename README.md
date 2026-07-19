# FunFamily Stock Bot

A private Telegram bot for tracking the family's stock portfolio. Production uses Telegram webhooks, Vercel Functions, and Postgres; local development can still use the original polling/Excel workflow.

## Features

- `/portfolio`, `/mystocks <name>`, and `/pnl` summaries
- guided `/addstock` and `/removestock` workflows
- `/refresh` for Yahoo Finance prices and `/update` for manual prices
- natural-language portfolio questions and actions
- daily family/member report at a configurable Singapore time
- Vercel webhook and secured Cron deployment
- automatic first-run import of the supplied Excel workbook into an empty database

## Setup

1. Create a Python 3.12 virtual environment.
2. Install dependencies with `pip install -r requirements.txt`.
3. Copy `.env.example` to `.env` and enter the Telegram and Anthropic credentials.
4. Add the private Telegram chat ID(s) to `TELEGRAM_DAILY_CHAT_IDS`.
5. Start the bot with `python bot.py`.

The daily job defaults to `08:00` Asia/Singapore. Change `DAILY_UPDATE_TIME` using 24-hour `HH:MM` format. If `TELEGRAM_DAILY_CHAT_IDS` is empty, the bot still runs but daily messages are disabled.

## Deploy on Vercel

1. Import this GitHub repository into Vercel.
2. Create a Postgres database (Vercel Marketplace/Neon or any reachable PostgreSQL service).
3. Add every variable from `.env.example` to the Vercel Production environment. Use long random values for `TELEGRAM_WEBHOOK_SECRET` and `CRON_SECRET`.

### Configure Anthropic in Vercel

In the Vercel dashboard, open **FunFamily Bot project → Settings → Environment Variables** and add:

| Name | Value | Environment |
| --- | --- | --- |
| `ANTHROPIC_API_KEY` | Your Anthropic Console API key | Production |
| `ANTHROPIC_MODEL` | `claude-sonnet-4-5-20250929` | Production |
| `ANTHROPIC_BASE_URL` | `https://api.anthropic.com` | Production |

Treat `ANTHROPIC_API_KEY` as sensitive and do not paste it into source files, `vercel.json`, GitHub, or `.env.example`. If you test Preview deployments, add separate values to **Preview** as well. Environment-variable changes only affect new deployments, so open **Deployments**, select the latest deployment, and choose **Redeploy** after saving them.

The FastAPI lifespan validates both variables before accepting traffic. Local polling mode performs the same validation after loading the untracked `.env` file.

To route requests through OpenRouter instead, set `ANTHROPIC_API_KEY` to your OpenRouter key, set `ANTHROPIC_BASE_URL=https://openrouter.ai/api`, and select an Anthropic model identifier supported by OpenRouter in `ANTHROPIC_MODEL`. The base URL is optional and continues to default to Anthropic's API.
4. Set `DATABASE_URL` to the provider's pooled Postgres connection string. On the first request, the schema is created and the bundled workbook seeds the database if it is empty.
5. Deploy the Production project and set `PUBLIC_BASE_URL` to its stable `https://...vercel.app` URL or custom domain.
6. Register the webhook from a trusted local machine:

```sh
python scripts/set_webhook.py
```

7. Visit `/api/health` and then check Telegram's `getWebhookInfo` API. The daily report runs at `00:00 UTC` (08:00 Singapore) through `vercel.json`. Vercel Hobby may execute it within that hour rather than at the exact minute.

Vercel installs production dependencies from `pyproject.toml`; keep its dependency list aligned with `requirements.txt` when packages change.

The production endpoints are:

- `POST /api/telegram`: validates Telegram's webhook secret, deduplicates update IDs, processes commands/natural language, and stores mutations in Postgres.
- `GET /api/daily-update`: validates Vercel's `CRON_SECRET`, refreshes prices, and sends the family report.
- `GET /api/health`: creates/checks the schema and confirms database connectivity.

Position deletion is deliberately explicit: use `/mystocks <name>` to see database IDs, then `/remove <id> confirm`.

## Local polling mode

`python bot.py` remains available for local development and uses `Funfamily_Stock_Tracker.xlsx`. Do not run polling while the Telegram webhook is active; Telegram supports only one update-delivery mode at a time.

## Security

Never commit `.env`. Keep the bot in a private family chat and rotate any credential that has previously been exposed.
