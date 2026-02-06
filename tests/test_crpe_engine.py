"""Tests for the native CRPE engine layer.

These tests require Crystal Reports to be installed (crpe32.dll available).
They are automatically skipped if the SDK is not found.
"""

import pytest
from pathlib import Path

from crystalreports.crpe_engine import CrpeEngine, is_sdk_available
from crystalreports.models import TableInfo, FormulaInfo, SectionInfo, ParameterInfo

SAMPLE_RPT = Path(__file__).resolve().parent.parent / "SafiPrint.rpt"

pytestmark = pytest.mark.skipif(
    not is_sdk_available(),
    reason="Crystal Reports SDK (crpe32.dll) not available",
)


@pytest.fixture(scope="module")
def engine():
    eng = CrpeEngine()
    yield eng
    eng.close()


@pytest.fixture
def job(engine):
    j = engine.open(SAMPLE_RPT)
    yield j
    j.close()


class TestEngine:
    def test_engine_opens(self, engine):
        # Engine opens lazily when first job is opened
        engine._ensure_engine()
        assert engine._opened

    def test_open_report(self, job):
        assert job._handle >= 0


class TestTables:
    def test_get_tables(self, job):
        tables = job.get_tables()
        assert isinstance(tables, list)
        assert len(tables) > 0

    def test_tables_are_table_info(self, job):
        for t in job.get_tables():
            assert isinstance(t, TableInfo)


class TestFormulas:
    def test_get_formulas(self, job):
        formulas = job.get_formulas()
        assert isinstance(formulas, list)
        assert len(formulas) > 0

    def test_formulas_are_formula_info(self, job):
        for f in job.get_formulas():
            assert isinstance(f, FormulaInfo)


class TestSQL:
    def test_get_sql_query(self, job):
        sql = job.get_sql_query()
        assert isinstance(sql, str)


class TestSections:
    def test_get_sections(self, job):
        sections = job.get_sections()
        assert isinstance(sections, list)
        assert len(sections) > 0


class TestParameters:
    def test_get_parameters(self, job):
        params = job.get_parameters()
        assert isinstance(params, list)


class TestSortFields:
    def test_get_sort_fields(self, job):
        fields = job.get_sort_fields()
        assert isinstance(fields, list)
