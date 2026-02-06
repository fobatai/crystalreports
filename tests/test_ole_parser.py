"""Tests for the pure-Python OLE layer."""

import pytest
from pathlib import Path

from crystalreports import OleParser, ReportMetadata, EmbeddedImage, SubreportInfo
from crystalreports.exceptions import ReportOpenError, OleParseError

SAMPLE_RPT = Path(__file__).resolve().parent.parent / "SafiPrint.rpt"


@pytest.fixture
def parser():
    """Open the sample .rpt file with OleParser."""
    p = OleParser(SAMPLE_RPT)
    yield p
    p.close()


class TestOleParserOpen:
    def test_open_valid_file(self, parser):
        assert parser._ole is not None

    def test_open_nonexistent_raises(self):
        with pytest.raises(ReportOpenError):
            OleParser("nonexistent.rpt")

    def test_context_manager(self):
        with OleParser(SAMPLE_RPT) as p:
            assert p._ole is not None
        assert p._ole is None


class TestMetadata:
    def test_returns_report_metadata(self, parser):
        meta = parser.get_metadata()
        assert isinstance(meta, ReportMetadata)

    def test_creating_application_not_none(self, parser):
        meta = parser.get_metadata()
        assert meta.creating_application is not None


class TestStreams:
    def test_list_streams_returns_list(self, parser):
        streams = parser.list_streams()
        assert isinstance(streams, list)
        assert len(streams) > 0

    def test_all_streams_are_strings(self, parser):
        for s in parser.list_streams():
            assert isinstance(s, str)

    def test_get_stream_returns_bytes(self, parser):
        streams = parser.list_streams()
        data = parser.get_stream(streams[0])
        assert isinstance(data, bytes)

    def test_get_nonexistent_stream_raises(self, parser):
        with pytest.raises(OleParseError):
            parser.get_stream("nonexistent/stream")


class TestEmbeddedImages:
    def test_returns_list(self, parser):
        images = parser.get_embedded_images()
        assert isinstance(images, list)

    def test_images_are_embedded_image(self, parser):
        images = parser.get_embedded_images()
        for img in images:
            assert isinstance(img, EmbeddedImage)
            assert isinstance(img.data, bytes)
            assert img.size > 0


class TestSubreports:
    def test_returns_list(self, parser):
        subs = parser.list_subreports()
        assert isinstance(subs, list)

    def test_subreports_are_subreport_info(self, parser):
        for sub in parser.list_subreports():
            assert isinstance(sub, SubreportInfo)


class TestReportInfo:
    def test_returns_dict(self, parser):
        info = parser.get_report_info()
        assert isinstance(info, dict)
        assert "num_streams" in info
        assert "num_images" in info

    def test_parse_returns_all_keys(self, parser):
        result = parser.parse()
        assert "metadata" in result
        assert "streams" in result
        assert "embedded_images" in result
        assert "subreports" in result
