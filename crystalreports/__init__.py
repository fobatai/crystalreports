"""crystalreports - Python library for reading and modifying Crystal Reports .rpt files."""

from .exceptions import (
    CrystalReportsError,
    ExportError,
    OleParseError,
    ReportOpenError,
    SDKNotAvailableError,
)
from .models import (
    EmbeddedImage,
    ExportFormat,
    FormulaInfo,
    ParameterInfo,
    ReportMetadata,
    ReportObject,
    SectionInfo,
    SortFieldInfo,
    SubreportInfo,
    TableInfo,
)
from .ole_parser import OleParser
from .report import CrystalReport

try:
    from .crpe_engine import rgb_to_colorref
except Exception:
    def rgb_to_colorref(r: int, g: int, b: int) -> int:
        """Convert RGB values (0-255) to a Windows COLORREF (0x00BBGGRR)."""
        return (b << 16) | (g << 8) | r

__version__ = "0.1.0"

__all__ = [
    # Main class
    "CrystalReport",
    # OLE parser
    "OleParser",
    # Models
    "ReportMetadata",
    "TableInfo",
    "FormulaInfo",
    "SectionInfo",
    "ParameterInfo",
    "EmbeddedImage",
    "SubreportInfo",
    "SortFieldInfo",
    "ReportObject",
    "ExportFormat",
    # Helpers
    "rgb_to_colorref",
    # Exceptions
    "CrystalReportsError",
    "SDKNotAvailableError",
    "ReportOpenError",
    "ExportError",
    "OleParseError",
]
