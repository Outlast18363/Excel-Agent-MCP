"""Unit tests for the headless ``sheet_screenshot`` service behavior."""

from __future__ import annotations

import subprocess
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

from openpyxl import Workbook

from excel_mcp.helpers import ExcelServiceError
from excel_mcp.service import ExcelService, WorkbookSession


class SheetScreenshotServiceTests(unittest.TestCase):
    """Verify freshness, command construction, and export error handling."""

    def setUp(self) -> None:
        self.service = ExcelService()
        self.temp_dir = tempfile.TemporaryDirectory(prefix="excel-mcp-sheet-shot-test-")
        self.addCleanup(self.temp_dir.cleanup)
        self.workbook_path = Path(self.temp_dir.name) / "book.xlsx"

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Summary"
        workbook.create_sheet("Hidden")
        workbook.save(self.workbook_path)
        workbook.close()

    def test_get_sheet_page_index_uses_sheet_order(self) -> None:
        self.assertEqual(self.service._get_sheet_page_index(str(self.workbook_path), "Summary"), 1)
        self.assertEqual(self.service._get_sheet_page_index(str(self.workbook_path), "Hidden"), 2)

    def test_sync_workbook_for_fresh_render_recalculates_and_saves(self) -> None:
        app = MagicMock()
        workbook = MagicMock()
        session = WorkbookSession(
            workbook_id="wb_001",
            workbook=workbook,
            app=app,
            path=str(self.workbook_path.resolve()),
            read_only=False,
            visible=False,
        )
        self.service._workbooks[session.workbook_id] = session
        self.service._path_index[session.path] = session.workbook_id

        self.service._sync_workbook_for_fresh_render(session.path)

        app.calculate.assert_called_once_with()
        workbook.save.assert_called_once_with()

    def test_sync_workbook_for_fresh_render_rejects_read_only_session(self) -> None:
        session = WorkbookSession(
            workbook_id="wb_001",
            workbook=MagicMock(),
            app=MagicMock(),
            path=str(self.workbook_path.resolve()),
            read_only=True,
            visible=False,
        )
        self.service._workbooks[session.workbook_id] = session
        self.service._path_index[session.path] = session.workbook_id

        with self.assertRaises(ExcelServiceError):
            self.service._sync_workbook_for_fresh_render(session.path)

    def test_sync_workbook_for_fresh_render_raises_on_save_failure(self) -> None:
        app = MagicMock()
        workbook = MagicMock()
        workbook.save.side_effect = RuntimeError("boom")
        session = WorkbookSession(
            workbook_id="wb_001",
            workbook=workbook,
            app=app,
            path=str(self.workbook_path.resolve()),
            read_only=False,
            visible=False,
        )
        self.service._workbooks[session.workbook_id] = session
        self.service._path_index[session.path] = session.workbook_id

        with self.assertRaises(ExcelServiceError):
            self.service._sync_workbook_for_fresh_render(session.path)

    def test_export_sheet_pdf_with_libreoffice_builds_expected_command(self) -> None:
        export_dir = Path(self.temp_dir.name) / "export"
        profile_dir = Path(self.temp_dir.name) / "profile"
        export_dir.mkdir()
        profile_dir.mkdir()

        with patch("excel_mcp.service.subprocess.run", return_value=SimpleNamespace(returncode=0, stdout="", stderr="")) as mock_run:
            self.service._export_sheet_pdf_with_libreoffice(
                workbook_path=str(self.workbook_path.resolve()),
                export_dir=export_dir,
                profile_dir=profile_dir,
                page_index=2,
                soffice_path="/usr/bin/soffice",
                timeout_seconds=30,
            )

        command = mock_run.call_args.args[0]
        self.assertEqual(command[0], "/usr/bin/soffice")
        self.assertIn("--headless", command)
        self.assertTrue(any(part.startswith("-env:UserInstallation=file://") for part in command))
        self.assertIn("--convert-to", command)
        self.assertTrue(any("SinglePageSheets" in part for part in command))
        self.assertTrue(any('"value":"2"' in part for part in command))

    def test_export_sheet_pdf_with_libreoffice_raises_on_timeout(self) -> None:
        export_dir = Path(self.temp_dir.name) / "export-timeout"
        profile_dir = Path(self.temp_dir.name) / "profile-timeout"
        export_dir.mkdir()
        profile_dir.mkdir()

        with patch(
            "excel_mcp.service.subprocess.run",
            side_effect=subprocess.TimeoutExpired(cmd=["soffice"], timeout=1),
        ):
            with self.assertRaises(ExcelServiceError):
                self.service._export_sheet_pdf_with_libreoffice(
                    workbook_path=str(self.workbook_path.resolve()),
                    export_dir=export_dir,
                    profile_dir=profile_dir,
                    page_index=1,
                    soffice_path="/usr/bin/soffice",
                    timeout_seconds=1,
                )

    def test_sheet_screenshot_rejects_invalid_pixel_constraints(self) -> None:
        with self.assertRaises(ExcelServiceError):
            self.service.sheet_screenshot(
                path=str(self.workbook_path),
                sheet="Summary",
                max_width_px=0,
            )

    def test_sheet_screenshot_returns_minimal_payload(self) -> None:
        target_path = Path(self.temp_dir.name) / "sheet.png"

        def fake_export(**kwargs: object) -> None:
            pdf_path = Path(kwargs["export_dir"]) / f"{self.workbook_path.stem}.pdf"
            pdf_path.write_bytes(b"%PDF-1.4\n")

        with patch("excel_mcp.service.resolve_soffice_path", return_value="/usr/bin/soffice"), patch.object(
            self.service,
            "_export_sheet_pdf_with_libreoffice",
            side_effect=fake_export,
        ), patch.object(
            self.service,
            "_rasterize_pdf_first_page",
            side_effect=lambda **kwargs: Path(kwargs["output_path"]).write_bytes(b"png"),
        ):
            response = self.service.sheet_screenshot(
                path=str(self.workbook_path),
                sheet="Summary",
                output_path=str(target_path),
            )

        self.assertEqual(response, {"image_path": str(target_path)})
        self.assertTrue(target_path.exists())
