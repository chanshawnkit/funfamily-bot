import os
import unittest
from unittest.mock import AsyncMock, patch

from fastapi.testclient import TestClient

from api.index import app, command_reply, execute_tool, set_processing_reaction, trade_admin
import portfolio_db
from config import validate_anthropic_env


class ApiSecurityTests(unittest.TestCase):
    def test_webhook_rejects_wrong_secret(self):
        with patch.dict(os.environ, {"TELEGRAM_WEBHOOK_SECRET": "correct"}):
            response = TestClient(app).post(
                "/api/telegram",
                headers={"X-Telegram-Bot-Api-Secret-Token": "wrong"},
                json={"update_id": 1},
            )
        self.assertEqual(response.status_code, 401)

    def test_cron_rejects_wrong_secret(self):
        with patch.dict(os.environ, {"CRON_SECRET": "correct"}):
            response = TestClient(app).get(
                "/api/daily-update", headers={"Authorization": "Bearer wrong"}
            )
        self.assertEqual(response.status_code, 401)


class AnthropicConfigTests(unittest.TestCase):
    def test_anthropic_environment_is_required(self):
        with patch.dict(os.environ, {}, clear=True):
            with self.assertRaisesRegex(RuntimeError, "ANTHROPIC_API_KEY, ANTHROPIC_MODEL"):
                validate_anthropic_env()

    def test_anthropic_environment_is_returned(self):
        values = {
            "ANTHROPIC_API_KEY": "test-key",
            "ANTHROPIC_MODEL": "claude-sonnet-5",
        }
        with patch.dict(os.environ, values, clear=True):
            self.assertEqual(
                validate_anthropic_env(),
                ("test-key", values["ANTHROPIC_MODEL"], "https://api.anthropic.com"),
            )

    def test_anthropic_environment_respects_custom_base_url(self):
        values = {
            "ANTHROPIC_API_KEY": "sk-or-test-key",
            "ANTHROPIC_MODEL": "claude-sonnet-5",
            "ANTHROPIC_BASE_URL": "https://openrouter.ai/api",
        }
        with patch.dict(os.environ, values, clear=True):
            self.assertEqual(
                validate_anthropic_env(),
                ("sk-or-test-key", "claude-sonnet-5", "https://openrouter.ai/api"),
            )

class CommandTests(unittest.IsolatedAsyncioTestCase):
    async def test_processing_reaction_sets_persistent_look_emoji(self):
        with patch("api.index.telegram_request", new=AsyncMock()) as telegram_request:
            await set_processing_reaction(-5235714051, 99)
        telegram_request.assert_awaited_once_with(
            "setMessageReaction",
            {
                "chat_id": -5235714051,
                "message_id": 99,
                "reaction": [{"type": "emoji", "emoji": "👀"}],
            },
        )

    async def test_portfolio_command_uses_database_summary(self):
        with patch("api.index.asyncio.to_thread", new=AsyncMock(return_value="summary")):
            self.assertEqual(await command_reply("/portfolio"), "summary")

    async def test_remove_requires_explicit_confirmation(self):
        reply = await command_reply("/remove 42", can_trade=True)
        self.assertIn("confirm", reply)

    async def test_remove_is_restricted_to_trade_admin(self):
        reply = await command_reply("/remove 42 confirm", can_trade=False)
        self.assertIn("Only Chan Shawn Kit", reply)

    def test_purchase_tool_is_restricted_to_trade_admin(self):
        with patch.object(portfolio_db, "add_position") as add_position:
            reply = execute_tool("add_stock_position", {}, can_trade=False)
        self.assertIn("Only Chan Shawn Kit", reply)
        add_position.assert_not_called()

    def test_natural_language_delete_is_restricted_to_trade_admin(self):
        with patch.object(portfolio_db, "remove_position") as remove_position:
            reply = execute_tool(
                "remove_stock_position",
                {"position_id": 7, "confirmed": True},
                can_trade=False,
            )
        self.assertIn("Only Chan Shawn Kit", reply)
        remove_position.assert_not_called()

    def test_natural_language_delete_requires_confirmation(self):
        with patch.object(portfolio_db, "remove_position") as remove_position:
            reply = execute_tool(
                "remove_stock_position",
                {"position_id": 7, "confirmed": False},
                can_trade=True,
            )
        self.assertIn("/remove 7 confirm", reply)
        remove_position.assert_not_called()

    def test_confirmed_natural_language_delete_removes_position(self):
        with patch.object(portfolio_db, "remove_position", return_value=True) as remove_position:
            reply = execute_tool(
                "remove_stock_position",
                {"position_id": 7, "confirmed": True},
                can_trade=True,
            )
        self.assertEqual(reply, "Position removed.")
        remove_position.assert_called_once_with(7)

    def test_trade_admin_uses_telegram_user_id(self):
        with patch.dict(os.environ, {"TELEGRAM_TRADE_ADMIN_USER_IDS": "123,456"}):
            self.assertTrue(trade_admin(456))
            self.assertFalse(trade_admin(789))


class SummaryTests(unittest.TestCase):
    def test_summary_aggregates_complete_positions(self):
        rows = [
            {"purchaser": "Jasmine", "amount_sgd": 1000.0,
             "gross_value_sgd": 1100.0, "net_pnl_sgd": 100.0},
            {"purchaser": "Jasmine", "amount_sgd": 500.0,
             "gross_value_sgd": None, "net_pnl_sgd": None},
        ]
        with patch.object(portfolio_db, "positions", return_value=rows):
            summary = portfolio_db.portfolio_summary()
        self.assertIn("Jasmine: S$1,100.00", summary)
        self.assertIn("P&L S$100.00 (+10.00%)", summary)


if __name__ == "__main__":
    unittest.main()
