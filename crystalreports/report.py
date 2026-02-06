"""Unified interface for reading and modifying Crystal Reports .rpt files.

:class:`CrystalReport` combines the pure-Python OLE layer (always available)
with the optional native CRPE engine layer (requires crpe32.dll).
"""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Optional, Union

from .exceptions import SDKNotAvailableError
from .models import (
    EmbeddedImage,
    ExportFormat,
    FormulaInfo,
    ParameterInfo,
    ReportMetadata,
    SectionInfo,
    SortFieldInfo,
    SubreportInfo,
    TableInfo,
)
from .ole_parser import OleParser

# Lazy import to avoid hard dependency on crpe32.dll
_CrpeEngine = None
_CrpeJob = None


def _get_crpe_classes():
    global _CrpeEngine, _CrpeJob
    if _CrpeEngine is None:
        from .crpe_engine import CrpeEngine as _CE, CrpeJob as _CJ
        _CrpeEngine = _CE
        _CrpeJob = _CJ
    return _CrpeEngine, _CrpeJob


class CrystalReport:
    """High-level interface for Crystal Reports .rpt files.

    Parameters
    ----------
    path : str or Path
        Path to the .rpt file.
    use_sdk : bool
        Attempt to load the native CRPE engine.  If *False* or the DLL
        is not available, only the pure-Python OLE layer is used.

    Examples
    --------
    >>> with CrystalReport("report.rpt") as rpt:
    ...     print(rpt.metadata)
    ...     print(rpt.tables)
    """

    def __init__(self, path: Union[str, Path], use_sdk: bool = True):
        self._path = Path(path).resolve()
        self._ole = OleParser(self._path)
        self._engine = None
        self._job = None

        if use_sdk:
            try:
                CrpeEngine, _ = _get_crpe_classes()
                self._engine = CrpeEngine()
                self._job = self._engine.open(self._path)
            except (SDKNotAvailableError, Exception):
                self._engine = None
                self._job = None

    # ------------------------------------------------------------------
    # Context manager
    # ------------------------------------------------------------------

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        return False

    def close(self) -> None:
        """Release all resources."""
        if self._job is not None:
            self._job.close()
            self._job = None
        if self._engine is not None:
            self._engine.close()
            self._engine = None
        if self._ole is not None:
            self._ole.close()
            self._ole = None

    # ------------------------------------------------------------------
    # SDK availability
    # ------------------------------------------------------------------

    @property
    def has_sdk(self) -> bool:
        """True if the native CRPE engine is loaded and a job is open."""
        return self._job is not None

    def _require_sdk(self) -> None:
        if not self.has_sdk:
            raise SDKNotAvailableError(
                "This operation requires the Crystal Reports SDK (crpe32.dll)."
            )

    # ------------------------------------------------------------------
    # Properties — OLE layer (always available)
    # ------------------------------------------------------------------

    @property
    def metadata(self) -> ReportMetadata:
        """Report metadata from OLE2 SummaryInformation."""
        return self._ole.get_metadata()

    @property
    def embedded_images(self) -> list[EmbeddedImage]:
        """Embedded BMP images from the OLE container."""
        return self._ole.get_embedded_images()

    @property
    def subreports(self) -> list[SubreportInfo]:
        """Subreport entries found in the OLE structure."""
        return self._ole.list_subreports()

    @property
    def streams(self) -> list[str]:
        """All OLE stream paths."""
        return self._ole.list_streams()

    # ------------------------------------------------------------------
    # Properties — CRPE layer (requires SDK)
    # ------------------------------------------------------------------

    @property
    def tables(self) -> list[TableInfo]:
        """Database tables used by the report."""
        self._require_sdk()
        return self._job.get_tables()

    @property
    def formulas(self) -> list[FormulaInfo]:
        """Formula fields defined in the report."""
        self._require_sdk()
        return self._job.get_formulas()

    @property
    def sql_query(self) -> str:
        """The SQL query used by the report."""
        self._require_sdk()
        return self._job.get_sql_query()

    @property
    def parameters(self) -> list[ParameterInfo]:
        """Parameter fields."""
        self._require_sdk()
        return self._job.get_parameters()

    @property
    def sections(self) -> list[SectionInfo]:
        """Report sections."""
        self._require_sdk()
        return self._job.get_sections()

    @property
    def sort_fields(self) -> list[SortFieldInfo]:
        """Sort fields."""
        self._require_sdk()
        return self._job.get_sort_fields()

    @property
    def n_groups(self) -> int:
        """Number of groups."""
        self._require_sdk()
        return self._job.get_n_groups()

    # ------------------------------------------------------------------
    # Editing — OLE layer
    # ------------------------------------------------------------------

    def set_metadata(self, output_path: Optional[Union[str, Path]] = None,
                     **kwargs) -> None:
        """Update report metadata (title, author, comments, etc.).

        Parameters
        ----------
        output_path : str or Path, optional
            Where to save.  Defaults to overwriting the original file.
        **kwargs
            Metadata fields to update.  See :class:`ReportMetadata`.
        """
        dest = output_path or self._path
        self._ole.set_metadata(dest, **kwargs)

    def replace_image(self, index: int,
                      image: Union[str, Path, bytes],
                      output_path: Optional[Union[str, Path]] = None) -> None:
        """Replace an embedded image.

        Parameters
        ----------
        index : int
            Zero-based index of the image.
        image : str, Path, or bytes
            File path or raw image data.
        output_path : str or Path, optional
            Save destination.
        """
        if isinstance(image, (str, Path)):
            data = Path(image).read_bytes()
        else:
            data = image
        self._ole.replace_embedded_image(index, data, output_path)

    # ------------------------------------------------------------------
    # Editing — CRPE layer
    # ------------------------------------------------------------------

    def set_table_connection(self, table_index: int, **kwargs) -> None:
        """Update the connection info for a table.

        Parameters
        ----------
        table_index : int
            Zero-based table index.
        **kwargs
            Passed to :meth:`CrpeJob.set_table_location`:
            ``location``, ``sublocation``, ``connect_buffer``, ``dll_name``.
        """
        self._require_sdk()
        self._job.set_table_location(table_index, **kwargs)

    def set_formula(self, name: str, text: str) -> None:
        """Set formula text by name."""
        self._require_sdk()
        self._job.set_formula(name, text)

    def set_sql_query(self, sql: str) -> None:
        """Set the SQL query."""
        self._require_sdk()
        self._job.set_sql_query(sql)

    # ------------------------------------------------------------------
    # Export
    # ------------------------------------------------------------------

    def export(self, output_path: Union[str, Path],
               fmt: Union[str, ExportFormat] = "pdf") -> None:
        """Export the report to the given format.

        Parameters
        ----------
        output_path : str or Path
            Destination file.
        fmt : str or ExportFormat
            ``"pdf"``, ``"xls"``, ``"xlsx"``, ``"rtf"``, ``"csv"``,
            ``"xml"``, or ``"txt"``.
        """
        self._require_sdk()
        if isinstance(fmt, str):
            fmt = ExportFormat(fmt.lower())
        self._job.export(output_path, fmt)

    # ------------------------------------------------------------------
    # Save
    # ------------------------------------------------------------------

    def save(self, output_path: Optional[Union[str, Path]] = None) -> None:
        """Save modifications.

        If the CRPE engine is available, uses ``PESavePrintJob``.
        Otherwise, copies the OLE file (metadata edits are written
        directly by :meth:`set_metadata`).
        """
        dest = output_path or self._path
        if self.has_sdk:
            self._job.save(dest)
        else:
            self._ole.save(dest)

    # ------------------------------------------------------------------
    # Convenience
    # ------------------------------------------------------------------

    def get_stream(self, stream_path: str) -> bytes:
        """Read raw bytes from an OLE stream."""
        return self._ole.get_stream(stream_path)

    def __repr__(self) -> str:
        sdk_status = "SDK" if self.has_sdk else "OLE-only"
        return f"<CrystalReport({self._path.name!r}, {sdk_status})>"
