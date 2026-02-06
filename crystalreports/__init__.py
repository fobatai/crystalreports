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
    SectionInfo,
    SortFieldInfo,
    SubreportInfo,
    TableInfo,
)
from .ole_parser import OleParser
from .report import CrystalReport

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
    "ExportFormat",
    # Exceptions
    "CrystalReportsError",
    "SDKNotAvailableError",
    "ReportOpenError",
    "ExportError",
    "OleParseError",
]
