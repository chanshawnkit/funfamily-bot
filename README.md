# FunFamily Stock Bot

A private Telegram bot for tracking the family's stock portfolio in Excel. It supports Telegram commands, natural-language questions through Anthropic, Yahoo Finance price refreshes, and a scheduled daily performance message for each family member and the combined portfolio.

## Features

- `/portfolio`, `/mystocks <name>`, and `/pnl` summaries
- guided `/addstock` and `/removestock` workflows
- `/refresh` for Yahoo Finance prices and `/update` for manual prices
- natural-language portfolio questions and actions
- daily family/member report at a configurable Singapore time
- Docker-ready deployment

## Setup

1. Create a Python 3.12 virtual environment.
2. Install dependencies with `pip install -r requirements.txt`.
3. Copy `.env.example` to `.env` and enter the Telegram and Anthropic credentials.
4. Add the private Telegram chat ID(s) to `TELEGRAM_DAILY_CHAT_IDS`.
5. Start the bot with `python bot.py`.

The daily job defaults to `08:00` Asia/Singapore. Change `DAILY_UPDATE_TIME` using 24-hour `HH:MM` format. If `TELEGRAM_DAILY_CHAT_IDS` is empty, the bot still runs but daily messages are disabled.

## Deployment

Build and run the container from this repository:

```sh
docker build -t funfamily-bot .
docker run --env-file .env funfamily-bot
```

Use persistent storage for `Funfamily_Stock_Tracker.xlsx`; the bot updates that file as prices and positions change. GitHub stores the source, but a continuously running container host is still required for Telegram polling and daily jobs.

## Security

Never commit `.env`. Keep the bot in a private family chat and rotate any credential that has previously been exposed.
