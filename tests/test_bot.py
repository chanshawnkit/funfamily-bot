import os
import unittest
from unittest.mock import Mock, patch

import pandas as pd

import bot


class DailyReportTests(unittest.TestCase):
    def test_report_includes_members_total_and_missing_prices(self):
        rows = pd.DataFrame(
            [
                {
                    "Purchaser": "Jasmine",
                    "Ticker": "MSFT",
                    "Original Purchase Quantum(S$)": 1000.0,
                    "Holding Price (Any $)": 110.0,
                    "Gross Holding Value (S$)": 1100.0,
                    "Net Earning/Loss (S$)": 100.0,
                },
                {
                    "Purchaser": "Shawn Mun",
                    "Ticker": "AAPL",
                    "Original Purchase Quantum(S$)": 500.0,
                    "Holding Price (Any $)": None,
                    "Gross Holding Value (S$)": None,
                    "Net Earning/Loss (S$)": None,
                },
            ]
        )
        with patch.object(bot, "load_stock", return_value=rows):
            report = bot.build_daily_report()

        self.assertIn("Jasmine: S$1,100.00", report)
        self.assertIn("Total P&L: S$100.00 (+10.00%)", report)
        self.assertIn("Needs price data: AAPL", report)

    def test_scheduler_is_disabled_without_chat_ids(self):
        app = Mock()
        with patch.dict(os.environ, {"TELEGRAM_DAILY_CHAT_IDS": ""}):
            self.assertEqual(bot.schedule_daily_updates(app), 0)
        app.job_queue.run_daily.assert_not_called()


if __name__ == "__main__":
    unittest.main()
