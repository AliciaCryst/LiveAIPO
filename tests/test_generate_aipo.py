import unittest
from unittest.mock import patch
from io import BytesIO

from openpyxl import Workbook

try:
    import pandas as pd
except ModuleNotFoundError:  # Parser tests still run before optional runtime deps are installed.
    pd = None

from generate_aipo import _quote_from_frame, get_quotes, parse_holdings


def workbook_bytes(rows):
    workbook = Workbook()
    sheet = workbook.active
    for row in rows:
        sheet.append(row)
    output = BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class HoldingsTests(unittest.TestCase):
    def test_published_symbols_are_preserved_and_unquotable_rows_removed(self):
        content = workbook_bytes([
            ("AIPO", None, None, None, None, None),
            ("09/28/2026", None, None, None, None, None),
            ("% of Net Assets", "Name", "Identifier", "CUSIP", "Shares Held", "Market Value"),
            (None, None, None, None, None, None),
            ("35%", "Alphabet Inc Class A", "GOOGL", "X", 1, 35),
            (0.25, "Alphabet Inc Class C", "GOOG", "X", 1, 25),
            ("20%", "Berkshire Hathaway", "BRK.B", "X", 1, 20),
            (0.19999997, "Berkshire duplicate row", "BRK.B", "X", 1, 20),
            (0.00000003, "Unquotable", "2602335D", "X", 1, 1),
            ("0.2%", "Government Obligations Fund", "FGXXX", "X", 1, 1),
            ("0.2%", "Cash & Other", "Cash&Other", "X", 1, 1),
        ])
        holdings, date, excluded = parse_holdings(content, min_holdings=3)
        self.assertEqual(date, "2026-09-28")
        self.assertEqual([(h.ticker, h.name, h.weight) for h in holdings], [
            ("BRK-B", "Berkshire duplicate row", 39.999997),
            ("GOOGL", "Alphabet Inc Class A", 35.0),
            ("GOOG", "Alphabet Inc Class C", 25.0),
        ])
        self.assertEqual([symbol for _, symbol, _ in excluded], ["2602335D"])

    def test_missing_header_is_rejected(self):
        content = workbook_bytes([("AIPO",), ("09/28/2026",), ("wrong", "header")])
        with self.assertRaisesRegex(ValueError, "% of Net Assets"):
            parse_holdings(content)

    def test_wrong_fund_is_rejected(self):
        content = workbook_bytes([("NOT-AIPO",), ("09/28/2026",)])
        with self.assertRaisesRegex(ValueError, "identify itself as AIPO"):
            parse_holdings(content)

    @unittest.skipIf(pd is None, "pandas is not installed")
    def test_quotes_use_previous_valid_close_and_handle_missing(self):
        days = pd.to_datetime(["2026-09-23 15:59", "2026-09-24 09:30", "2026-09-24 09:34"]).tz_localize("America/New_York")
        columns = pd.MultiIndex.from_product([["NVDA", "BRK-B"], ["Close"]])
        frame = pd.DataFrame([[120, 310], [None, 309], [117, 305]], index=days, columns=columns)
        self.assertEqual(_quote_from_frame(frame, "NVDA"), {
            "price": 117.0, "change": -3.0, "changePct": -2.5, "priceDate": "2026-09-24 06:34 PDT"
        })
        self.assertEqual(_quote_from_frame(frame, "BRK-B")["changePct"], -1.6129)
        self.assertIsNone(_quote_from_frame(frame, "MISSING"))

    @unittest.skipIf(pd is None, "pandas is not installed")
    def test_failed_quote_batch_does_not_silently_publish(self):
        with patch("yfinance.download", return_value=pd.DataFrame()):
            with self.assertRaisesRegex(RuntimeError, "existing output was kept"):
                get_quotes([f"TEST{i}" for i in range(15)],
                           batch_size=10, delay=0, retry_delay=0, max_retries=1)


if __name__ == "__main__":
    unittest.main()
