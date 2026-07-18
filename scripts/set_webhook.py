import os
import httpx
from dotenv import load_dotenv


load_dotenv()
token = os.environ["TELEGRAM_BOT_TOKEN"]
base_url = os.environ["PUBLIC_BASE_URL"].rstrip("/")
webhook_secret = os.environ["TELEGRAM_WEBHOOK_SECRET"]

response = httpx.post(
    f"https://api.telegram.org/bot{token}/setWebhook",
    json={
        "url": f"{base_url}/api/telegram",
        "secret_token": webhook_secret,
        "drop_pending_updates": True,
        "allowed_updates": ["message", "edited_message"],
    },
    timeout=30,
)
response.raise_for_status()
result = response.json()
if not result.get("ok"):
    raise RuntimeError(result)
print(result["description"])
