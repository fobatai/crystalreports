"""Data models for the crystalreports package."""

from __future__ import annotations

import enum
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


class ExportFormat(enum.Enum):
    """Supported export formats for Crystal Reports."""
    PDF = "pdf"
    XLS = "xls"
    XLSX = "xlsx"
    RTF = "rtf"
    CSV = "csv"
    XML = "xml"
    TXT = "txt"


@dataclass
class ReportMetadata:
    """Metadata extracted from the OLE2 SummaryInformation stream."""
    title: Optional[str] = None
    subject: Optional[str] = None
    author: Optional[str] = None
    keywords: Optional[str] = None
    comments: Optional[str] = None
    creating_application: Optional[str] = None
    create_time: Optional[datetime] = None
    last_save_time: Optional[datetime] = None
    revision_number: Optional[str] = None


@dataclass
class TableInfo:
    """Information about a database table used in the report."""
    index: int = 0
    name: str = ""
    location: str = ""
    sublocation: str = ""
    connection_string: str = ""
    dll_name: str = ""


@dataclass
class FormulaInfo:
    """Information about a formula field in the report."""
    index: int = 0
    name: str = ""
    text: str = ""


@dataclass
class SectionInfo:
    """Information about a report section."""
    index: int = 0
    code: int = 0
    name: str = ""


@dataclass
class ParameterInfo:
    """Information about a parameter field in the report."""
    index: int = 0
    name: str = ""
    prompt: str = ""
    default_value: Optional[str] = None
    value_type: int = 0


@dataclass
class EmbeddedImage:
    """An embedded image extracted from the OLE2 container."""
    index: int = 0
    name: str = ""
    format: str = "bmp"
    data: bytes = field(default=b"", repr=False)

    @property
    def size(self) -> int:
        return len(self.data)


@dataclass
class SubreportInfo:
    """Information about a subreport embedded in the report."""
    index: int = 0
    name: str = ""
    stream_path: str = ""


@dataclass
class SortFieldInfo:
    """Information about a sort field."""
    index: int = 0
    field_name: str = ""
    direction: str = "ascending"


@dataclass
class ReportObject:
    """Information about an object on the report layout."""
    handle: int = 0
    name: str = ""
    object_type: str = ""
    section_code: int = 0
    left: int = 0
    top: int = 0
    right: int = 0
    bottom: int = 0
