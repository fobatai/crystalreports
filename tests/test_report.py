"""Integration tests for the CrystalReport main class."""

import pytest
from pathlib import Path

from crystalreports import CrystalReport, ReportMetadata

SAMPLE_RPT = Path(__file__).resolve().parent.parent / "SafiPrint.rpt"


@pytest.fixture
def report():
    rpt = CrystalReport(SAMPLE_RPT)
    yield rpt
    rpt.close()


@pytest.fixture
def ole_only_report():
    """Report opened without SDK."""
    rpt = CrystalReport(SAMPLE_RPT, use_sdk=False)
    yield rpt
    rpt.close()


class TestOleLayer:
    """Tests that always work (pure Python)."""

    def test_metadata(self, ole_only_report):
        meta = ole_only_report.metadata
        assert isinstance(meta, ReportMetadata)

    def test_streams(self, ole_only_report):
        streams = ole_only_report.streams
        assert isinstance(streams, list)
        assert len(streams) > 0

    def test_embedded_images(self, ole_only_report):
        images = ole_only_report.embedded_images
        assert isinstance(images, list)

    def test_subreports(self, ole_only_report):
        subs = ole_only_report.subreports
        assert isinstance(subs, list)

    def test_has_sdk_false(self, ole_only_report):
        assert ole_only_report.has_sdk is False

    def test_repr(self, ole_only_report):
        r = repr(ole_only_report)
        assert "SafiPrint.rpt" in r
        assert "OLE-only" in r

    def test_context_manager(self):
        with CrystalReport(SAMPLE_RPT, use_sdk=False) as rpt:
            assert rpt.metadata is not None


class TestSDKLayer:
    """Tests that require crpe32.dll."""

    def test_has_sdk(self, report):
        # May or may not be True depending on environment
        if report.has_sdk:
            assert len(report.tables) > 0
            assert len(report.formulas) > 0

    def test_tables_require_sdk(self, ole_only_report):
        from crystalreports.exceptions import SDKNotAvailableError
        with pytest.raises(SDKNotAvailableError):
            _ = ole_only_report.tables

    def test_formulas_require_sdk(self, ole_only_report):
        from crystalreports.exceptions import SDKNotAvailableError
        with pytest.raises(SDKNotAvailableError):
            _ = ole_only_report.formulas

    def test_sql_query_require_sdk(self, ole_only_report):
        from crystalreports.exceptions import SDKNotAvailableError
        with pytest.raises(SDKNotAvailableError):
            _ = ole_only_report.sql_query


class TestSaveMetadata:
    def test_set_metadata_creates_copy(self, ole_only_report, tmp_path):
        dest = tmp_path / "copy.rpt"
        ole_only_report.set_metadata(output_path=dest, title="Test Title")
        assert dest.exists()
