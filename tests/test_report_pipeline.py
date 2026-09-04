import datetime as dt
import json
import re
import unittest
from pathlib import Path

import openpyxl

import process_data
import weekly_report


ROOT = Path(__file__).resolve().parents[1]


class ReportPipelineTests(unittest.TestCase):
    def test_filename_dates_and_as_of_selection(self):
        self.assertEqual(
            weekly_report.date_from_filename(Path("source_20260821.xlsx")),
            dt.date(2026, 8, 21),
        )
        self.assertEqual(
            weekly_report.date_from_filename(Path("source_update 210826.xlsx")),
            dt.date(2026, 8, 21),
        )
        self.assertIsNone(
            weekly_report.latest_file(
                ["gd tu doanh/Thong ke Tu doanh Co phieu_update *.xlsx"],
                dt.date(2026, 8, 21),
            )
        )

    def test_missing_source_path_is_blocking(self):
        errors, _ = weekly_report.validate_sources(
            {"proprietary": ROOT / "does-not-exist.xlsx"},
            dt.date(2026, 8, 21),
            {"critical_sources": ["proprietary"], "source_max_age_days": 7},
        )
        self.assertTrue(errors)

    def test_reference_commentary_matches_required_structure(self):
        payload = json.loads(
            (ROOT / "inputs" / "commentary_reference_2026-08-24.json").read_text(encoding="utf-8")
        )
        validated = weekly_report.validate_commentary(payload)
        self.assertEqual(validated["slide8"], validated["slide3"][:2])
        self.assertTrue(validated["slide3"][0].startswith("Trong tuần "))
        self.assertTrue(validated["slide4"][0].startswith("Xét về các nhóm ngành,"))

    def test_vingroup_schema_rejects_valuation_export(self):
        files = list((ROOT / "dong gop cua vingroup").glob("FiinProX_*Doanh_nghiep*20260821*.xlsx"))
        results = {path.name: process_data.is_vingroup_price_share_export(str(path)) for path in files}
        self.assertTrue(any("(1)" in name and valid for name, valid in results.items()))
        self.assertTrue(any("(2)" in name and not valid for name, valid in results.items()))

    def test_nn_schema_rejects_valuation_export(self):
        files = list((ROOT / "nn ban rong").glob("FiinProX_*20260821*.xlsx"))
        results = {path.name: process_data.is_nn_buy_sell_room_export(str(path)) for path in files}
        self.assertTrue(any(name.endswith("20260821.xlsx") and valid for name, valid in results.items()))
        self.assertTrue(any("(2)" in name and not valid for name, valid in results.items()))

    def test_vingroup_sanity_check_rejects_corrupted_processed_rows(self):
        path = ROOT / "dong gop cua vingroup" / "Dong gop cua vingroup_update 210826.xlsx"
        self.assertIsNone(weekly_report.compute_vingroup_contribution(path, dt.date(2026, 8, 21)))

    def test_vingroup_contribution_accepts_valid_price_share_rows(self):
        records = []
        for day, index, price in (
            (dt.date(2026, 8, 20), 1700.0, 100_000.0),
            (dt.date(2026, 8, 21), 1710.0, 101_000.0),
        ):
            stocks = []
            for _ in range(4):
                stocks.extend([price, 1_000_000_000.0])
            records.append({"date": day, "index": index, "market_cap": 7_000_000.0, "stocks": stocks})
        result = weekly_report.compute_vingroup_metrics(records, dt.date(2026, 8, 21))
        self.assertIsNotNone(result)
        self.assertGreater(result["week"]["group_points"], 0)
        self.assertLess(result["week"]["group_points"], 10)

    def test_output_code_has_no_calibri_or_aptos_defaults(self):
        files = [
            ROOT / "weekly_report.py",
            ROOT / "process_data.py",
            ROOT / "process_data_api.py",
            ROOT / "scripts" / "windows_office.ps1",
            ROOT / "scripts" / "build_proprietary_workbook.mjs",
        ]
        for path in files:
            text = path.read_text(encoding="utf-8")
            self.assertIsNone(
                re.search(r"(?:name\s*=\s*|name:\s*)[\"'](?:Calibri|Aptos)[\"']", text),
                path.name,
            )

    def test_font_guard_detects_old_reference_deck(self):
        path = ROOT / "output_reports" / "2026-08-21" / "MBS Dau Tu - BC Thi truong Tuan - 24.08.2026.pptx"
        with self.assertRaises(RuntimeError):
            weekly_report.assert_presentation_fonts(path)

    def test_pdf_font_guard_detects_old_reference_pdf(self):
        path = ROOT / "output_reports" / "2026-08-21" / "MBS Dau Tu - BC Thi truong Tuan - 24.08.2026.pdf"
        with self.assertRaises(RuntimeError):
            weekly_report.assert_pdf_fonts(path)


if __name__ == "__main__":
    unittest.main()
