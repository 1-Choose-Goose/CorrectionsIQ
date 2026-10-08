from __future__ import annotations

import unittest
from datetime import date
from pathlib import Path

from app_gui import CalculationResult, HeadcountApp
from average_headcount import Segment


class NavigationStateTests(unittest.TestCase):
    def test_calculation_result_is_stored_while_about_page_is_open(self) -> None:
        app = object.__new__(HeadcountApp)
        app.active_module = "about"
        app.result = None
        app.calculation_error = "previous error"
        segment = Segment(
            report_date=date(2026, 1, 1),
            value=100,
            days=1,
            weighted=100,
            path=Path("report.docx"),
            matches=(),
        )
        result = CalculationResult(
            segments=[segment],
            total_days=1,
            total_weighted=100,
            average=100.0,
            institution="ИК-1",
            start=date(2026, 1, 1),
            end=date(2026, 1, 1),
            category=None,
        )

        app._accept_calculation_result(result)

        self.assertIs(app.result, result)
        self.assertIsNone(app.calculation_error)


if __name__ == "__main__":
    unittest.main()
