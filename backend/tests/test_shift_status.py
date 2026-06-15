from __future__ import annotations

import unittest
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from main import parse_time_slots, summarize_slot_status


class ShiftStatusTest(unittest.TestCase):
    def assert_status(self, time_text: str, expected: str) -> None:
        self.assertEqual(summarize_slot_status(parse_time_slots(time_text)), expected)

    def test_requested_time_ranges(self) -> None:
        self.assert_status("9:00-18:00", "白天在")
        self.assert_status("13:00-18:00", "下午在")
        self.assert_status("16:00-24:00", "晚上在")
        self.assert_status("13:00-24:00", "下午在晚上也在")

    def test_existing_excel_time_ranges(self) -> None:
        self.assert_status("9:00-13:00", "上午在")
        self.assert_status("18:00-24：00", "晚上在")
        self.assert_status("16:00-24：00", "晚上在")
        self.assert_status("9:00-24：00", "全天在")


if __name__ == "__main__":
    unittest.main()
