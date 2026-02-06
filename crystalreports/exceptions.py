"""Custom exceptions for the crystalreports package."""


class CrystalReportsError(Exception):
    """Base exception for all Crystal Reports errors."""


class SDKNotAvailableError(CrystalReportsError):
    """Raised when the Crystal Reports SDK (crpe32.dll) is not available."""


class ReportOpenError(CrystalReportsError):
    """Raised when a report file cannot be opened."""


class ExportError(CrystalReportsError):
    """Raised when exporting a report fails."""


class OleParseError(CrystalReportsError):
    """Raised when OLE2 parsing fails."""
