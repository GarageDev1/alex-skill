from __future__ import annotations

import importlib.util
from pathlib import Path
import tempfile
import unittest
from unittest import mock
import zipfile


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "recalc.py"
SPEC = importlib.util.spec_from_file_location("valuation_recalc", SCRIPT)
assert SPEC and SPEC.loader
recalc = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(recalc)


class RecalcTests(unittest.TestCase):
    def test_auto_returns_all_available_backends_in_priority_order(self) -> None:
        with mock.patch.object(
            recalc,
            "available_backends",
            return_value={
                "excel-com": True,
                "excel-applescript": False,
                "libreoffice": True,
            },
        ):
            self.assertEqual(recalc.choose_backends("auto"), ["excel-com", "libreoffice"])

    def test_macos_excel_backend_is_selectable(self) -> None:
        with mock.patch.object(
            recalc,
            "available_backends",
            return_value={
                "excel-com": False,
                "excel-applescript": True,
                "libreoffice": False,
            },
        ):
            self.assertEqual(recalc.choose_backends("auto"), ["excel-applescript"])

    def test_darwin_detection_supports_excel_and_libreoffice_fallback(self) -> None:
        with (
            mock.patch.object(recalc.platform, "system", return_value="Darwin"),
            mock.patch.object(recalc, "_mac_excel_installed", return_value=True),
            mock.patch.object(recalc.shutil, "which", side_effect=lambda name: "/usr/bin/osascript" if name == "osascript" else None),
            mock.patch.object(recalc, "_libreoffice_command", return_value="/opt/homebrew/bin/soffice"),
        ):
            available = recalc.available_backends()

        self.assertFalse(available["excel-com"])
        self.assertTrue(available["excel-applescript"])
        self.assertTrue(available["libreoffice"])

    def test_explicit_unavailable_backend_fails(self) -> None:
        with mock.patch.object(
            recalc,
            "available_backends",
            return_value={
                "excel-com": False,
                "excel-applescript": False,
                "libreoffice": True,
            },
        ):
            with self.assertRaisesRegex(RuntimeError, "后端不可用"):
                recalc.choose_backends("excel-com")

    def test_failed_recalc_never_changes_source(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "model.xlsx"
            with zipfile.ZipFile(source, "w") as archive:
                archive.writestr("xl/workbook.xml", "<workbook/>")
            original = source.read_bytes()

            with mock.patch.object(recalc, "_recalc_excel_com", side_effect=RuntimeError("boom")):
                with self.assertRaisesRegex(RuntimeError, "boom"):
                    recalc.recalc_one(source, "excel-com")

            self.assertEqual(source.read_bytes(), original)
            self.assertFalse(source.with_name(f".{source.name}.recalc-ready").exists())


if __name__ == "__main__":
    unittest.main()
