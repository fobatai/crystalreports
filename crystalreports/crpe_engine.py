"""Native SDK layer for Crystal Reports via crpe32.dll (ctypes).

This module wraps the Crystal Reports Print Engine (CRPE) API using
:mod:`ctypes`.  It requires a local installation of Crystal Reports
(e.g. Crystal Reports 2020) that provides ``crpe32.dll``.

If the DLL is not available the module can still be imported — all
functions will raise :class:`~.exceptions.SDKNotAvailableError`.
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import os
from pathlib import Path
from typing import Optional, Union

from .exceptions import (
    CrystalReportsError,
    ExportError,
    ReportOpenError,
    SDKNotAvailableError,
)
from .models import (
    ExportFormat,
    FormulaInfo,
    ParameterInfo,
    SectionInfo,
    SortFieldInfo,
    TableInfo,
)

# ------------------------------------------------------------------
# DLL discovery
# ------------------------------------------------------------------

_DEFAULT_DLL_PATHS = [
    r"C:\Program Files (x86)\SAP BusinessObjects\SAP BusinessObjects Enterprise XI 4.0\win64_x64\crpe32.dll",
    r"C:\Program Files (x86)\SAP BusinessObjects\Crystal Reports for .NET Framework 4.0\Common\SAP BusinessObjects Enterprise XI 4.0\win64_x64\crpe32.dll",
    r"C:\Program Files\SAP BusinessObjects\SAP BusinessObjects Enterprise XI 4.0\win64_x64\crpe32.dll",
]

_dll: Optional[ctypes.WinDLL] = None
_dll_path: Optional[str] = None


def _find_dll() -> Optional[str]:
    """Locate crpe32.dll on the system."""
    env_path = os.environ.get("CRPE32_DLL_PATH")
    if env_path and Path(env_path).exists():
        return env_path
    for p in _DEFAULT_DLL_PATHS:
        if Path(p).exists():
            return p
    return None


def _load_dll() -> ctypes.WinDLL:
    """Load crpe32.dll, caching the result."""
    global _dll, _dll_path
    if _dll is not None:
        return _dll

    path = _find_dll()
    if path is None:
        raise SDKNotAvailableError(
            "crpe32.dll not found. Install Crystal Reports or set "
            "CRPE32_DLL_PATH environment variable."
        )

    dll_dir = str(Path(path).parent)
    os.add_dll_directory(dll_dir)
    os.environ["PATH"] = dll_dir + ";" + os.environ.get("PATH", "")

    try:
        _dll = ctypes.WinDLL(path)
    except OSError as exc:
        raise SDKNotAvailableError(f"Failed to load crpe32.dll: {exc}") from exc

    _dll_path = path
    _setup_prototypes(_dll)
    return _dll


def is_sdk_available() -> bool:
    """Return True if the Crystal Reports SDK can be loaded."""
    try:
        _load_dll()
        return True
    except SDKNotAvailableError:
        return False


# ------------------------------------------------------------------
# ctypes structure definitions
# ------------------------------------------------------------------

PE_WORD = ctypes.c_ushort
PE_LONG = ctypes.c_long
PE_HANDLE = ctypes.c_short


class PETableLocation(ctypes.Structure):
    _fields_ = [
        ("StructSize", PE_WORD),
        ("Location", ctypes.c_char * 256),
        ("SubLocation", ctypes.c_char * 256),
        ("ConnectBuffer", ctypes.c_char * 512),
        ("DLLName", ctypes.c_char * 64),
        ("DescriptiveName", ctypes.c_char * 256),
        ("DatabaseType", PE_WORD),
    ]


class PEExportOptions(ctypes.Structure):
    _fields_ = [
        ("StructSize", PE_WORD),
        ("formatDLLName", ctypes.c_char * 64),
        ("destinationDLLName", ctypes.c_char * 64),
        ("fileName", ctypes.c_char * 512),
        ("nFormatOptions", PE_LONG),
        ("formatOptions", ctypes.c_void_p),
        ("nDestinationOptions", PE_LONG),
        ("destinationOptions", ctypes.c_void_p),
    ]


# ------------------------------------------------------------------
# Export format/destination DLL mapping
# ------------------------------------------------------------------

_FORMAT_DLLS = {
    ExportFormat.PDF: b"u2fpdf.dll",
    ExportFormat.XLS: b"u2fxls.dll",
    ExportFormat.XLSX: b"u2fxls.dll",
    ExportFormat.RTF: b"u2frtf.dll",
    ExportFormat.CSV: b"u2fcsv.dll",
    ExportFormat.XML: b"u2fxml.dll",
    ExportFormat.TXT: b"u2ftxt.dll",
}

_DEST_DISK_DLL = b"u2ddisk.dll"


class DiskDestOptions(ctypes.Structure):
    _fields_ = [
        ("StructSize", PE_WORD),
        ("fileName", ctypes.c_char * 512),
    ]


# ------------------------------------------------------------------
# Function prototypes setup
# ------------------------------------------------------------------

def _setup_prototypes(dll: ctypes.WinDLL) -> None:
    """Declare argtypes/restype for the PE_ functions we use."""
    # Engine lifecycle
    dll.PEOpenEngine.argtypes = []
    dll.PEOpenEngine.restype = ctypes.c_bool
    dll.PECloseEngine.argtypes = []
    dll.PECloseEngine.restype = None

    # Print job
    dll.PEOpenPrintJob.argtypes = [ctypes.c_char_p]
    dll.PEOpenPrintJob.restype = PE_HANDLE
    dll.PEClosePrintJob.argtypes = [PE_HANDLE]
    dll.PEClosePrintJob.restype = None

    # Error handling
    dll.PEGetErrorCode.argtypes = [PE_HANDLE]
    dll.PEGetErrorCode.restype = PE_WORD
    dll.PEGetErrorText.argtypes = [PE_HANDLE, ctypes.POINTER(ctypes.c_int), ctypes.c_char_p]
    dll.PEGetErrorText.restype = ctypes.c_bool

    # Tables
    dll.PEGetNTables.argtypes = [PE_HANDLE]
    dll.PEGetNTables.restype = PE_WORD
    dll.PEGetNthTableLocation.argtypes = [PE_HANDLE, PE_WORD, ctypes.POINTER(PETableLocation)]
    dll.PEGetNthTableLocation.restype = ctypes.c_bool
    dll.PESetNthTableLocation.argtypes = [PE_HANDLE, PE_WORD, ctypes.POINTER(PETableLocation)]
    dll.PESetNthTableLocation.restype = ctypes.c_bool

    # Formula count
    dll.PEGetNFormulas.argtypes = [PE_HANDLE]
    dll.PEGetNFormulas.restype = PE_WORD

    # PEGetNthFormulaEx — returns 64-bit text handles
    # (job, index, *nameHandle, *nameLen, *textHandle, *textLen)
    dll.PEGetNthFormulaEx.argtypes = [
        PE_HANDLE, PE_WORD,
        ctypes.POINTER(ctypes.c_void_p), ctypes.POINTER(ctypes.c_int),
        ctypes.POINTER(ctypes.c_void_p), ctypes.POINTER(ctypes.c_int),
    ]
    dll.PEGetNthFormulaEx.restype = ctypes.c_bool

    # PEGetHandleStringEx — convert text handle to string
    # (handle, buffer, bufferSize) -> BOOL
    dll.PEGetHandleStringEx.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_int]
    dll.PEGetHandleStringEx.restype = ctypes.c_bool

    # PEGetFormulaW — get formula text by name (returns text handle)
    dll.PEGetFormulaW.argtypes = [
        PE_HANDLE, ctypes.c_wchar_p,
        ctypes.POINTER(ctypes.c_void_p), ctypes.POINTER(ctypes.c_int),
    ]
    dll.PEGetFormulaW.restype = ctypes.c_bool

    # PESetFormula / PESetFormulaW
    dll.PESetFormula.argtypes = [PE_HANDLE, ctypes.c_char_p, ctypes.c_char_p]
    dll.PESetFormula.restype = ctypes.c_bool

    # SQL query
    dll.PEGetSQLQueryEx.argtypes = [
        PE_HANDLE,
        ctypes.POINTER(ctypes.c_void_p), ctypes.POINTER(ctypes.c_int),
    ]
    dll.PEGetSQLQueryEx.restype = ctypes.c_bool

    dll.PESetSQLQuery.argtypes = [PE_HANDLE, ctypes.c_char_p]
    dll.PESetSQLQuery.restype = ctypes.c_bool

    # Sections
    dll.PEGetNSections.argtypes = [PE_HANDLE]
    dll.PEGetNSections.restype = PE_WORD

    # Groups & sorting
    dll.PEGetNGroups.argtypes = [PE_HANDLE]
    dll.PEGetNGroups.restype = PE_WORD
    dll.PEGetNSortFields.argtypes = [PE_HANDLE]
    dll.PEGetNSortFields.restype = PE_WORD

    # Parameters
    dll.PEGetNParameterFields.argtypes = [PE_HANDLE]
    dll.PEGetNParameterFields.restype = PE_WORD

    # Subreports
    dll.PEGetNSubreportsInSection.argtypes = [PE_HANDLE, PE_WORD]
    dll.PEGetNSubreportsInSection.restype = PE_WORD
    dll.PEOpenSubreport.argtypes = [PE_HANDLE, ctypes.c_char_p]
    dll.PEOpenSubreport.restype = PE_HANDLE

    # Export
    dll.PEExportTo.argtypes = [PE_HANDLE, ctypes.POINTER(PEExportOptions)]
    dll.PEExportTo.restype = ctypes.c_bool

    # Suppress dialogs (optional — not present in all builds)
    try:
        dll.PESetDialogParentWindowHandle.argtypes = [PE_HANDLE, wt.HWND]
        dll.PESetDialogParentWindowHandle.restype = ctypes.c_bool
    except AttributeError:
        pass

    # Save (optional — not present in all builds)
    try:
        dll.PESavePrintJob.argtypes = [PE_HANDLE, ctypes.c_char_p]
        dll.PESavePrintJob.restype = ctypes.c_bool
    except AttributeError:
        pass


# ------------------------------------------------------------------
# Error helpers
# ------------------------------------------------------------------

def _get_error_text(dll, job_handle: int) -> str:
    """Retrieve the last error message from the engine."""
    text_handle = ctypes.c_int(0)
    buf = ctypes.create_string_buffer(512)
    code = dll.PEGetErrorCode(job_handle)
    if code == 0:
        return ""
    try:
        dll.PEGetErrorText(job_handle, ctypes.byref(text_handle), buf)
        return f"PE error {code}: {buf.value.decode('latin-1', errors='replace')}"
    except Exception:
        return f"PE error {code}"


def _check(dll, job_handle: int, result: bool, action: str = "") -> None:
    """Raise on failure."""
    if not result:
        msg = _get_error_text(dll, job_handle)
        raise CrystalReportsError(f"{action} failed: {msg}" if action else msg)


# ------------------------------------------------------------------
# Handle string helper
# ------------------------------------------------------------------

_HANDLE_BUF_SIZE = 8192


def _handle_to_str(dll, handle_value: int) -> str:
    """Convert a CRPE text handle (from Ex functions) to a Python string."""
    if not handle_value:
        return ""
    buf = ctypes.create_string_buffer(_HANDLE_BUF_SIZE)
    ok = dll.PEGetHandleStringEx(handle_value, buf, _HANDLE_BUF_SIZE)
    if ok and buf.value:
        return buf.value.decode("latin-1", errors="replace")
    return ""


# ------------------------------------------------------------------
# CrpeJob — single report session
# ------------------------------------------------------------------

class CrpeJob:
    """A handle to an opened Crystal Reports print job.

    Use :func:`CrpeEngine.open` or the context manager to create one.
    """

    def __init__(self, dll: ctypes.WinDLL, handle: int, path: str):
        self._dll = dll
        self._handle = handle
        self._path = path
        self._closed = False

    def close(self) -> None:
        if not self._closed and self._handle >= 0:
            self._dll.PEClosePrintJob(self._handle)
            self._closed = True

    # -- Tables --

    def get_tables(self) -> list[TableInfo]:
        n = self._dll.PEGetNTables(self._handle)
        tables: list[TableInfo] = []
        for i in range(n):
            loc = PETableLocation()
            loc.StructSize = ctypes.sizeof(PETableLocation)
            ok = self._dll.PEGetNthTableLocation(self._handle, i, ctypes.byref(loc))
            if not ok:
                continue
            tables.append(TableInfo(
                index=i,
                name=loc.DescriptiveName.decode("latin-1", errors="replace").rstrip("\x00"),
                location=loc.Location.decode("latin-1", errors="replace").rstrip("\x00"),
                sublocation=loc.SubLocation.decode("latin-1", errors="replace").rstrip("\x00"),
                connection_string=loc.ConnectBuffer.decode("latin-1", errors="replace").rstrip("\x00"),
                dll_name=loc.DLLName.decode("latin-1", errors="replace").rstrip("\x00"),
            ))
        return tables

    def set_table_location(self, index: int, location: str = "",
                           sublocation: str = "", connect_buffer: str = "",
                           dll_name: str = "") -> None:
        """Update the connection info for the *index*-th table."""
        loc = PETableLocation()
        loc.StructSize = ctypes.sizeof(PETableLocation)
        self._dll.PEGetNthTableLocation(self._handle, index, ctypes.byref(loc))
        if location:
            loc.Location = location.encode("latin-1")
        if sublocation:
            loc.SubLocation = sublocation.encode("latin-1")
        if connect_buffer:
            loc.ConnectBuffer = connect_buffer.encode("latin-1")
        if dll_name:
            loc.DLLName = dll_name.encode("latin-1")
        ok = self._dll.PESetNthTableLocation(self._handle, index, ctypes.byref(loc))
        _check(self._dll, self._handle, ok, "SetNthTableLocation")

    # -- Formulas --

    def get_formulas(self) -> list[FormulaInfo]:
        n = self._dll.PEGetNFormulas(self._handle)
        formulas: list[FormulaInfo] = []
        for i in range(n):
            nh = ctypes.c_void_p(0)
            nl = ctypes.c_int(0)
            th = ctypes.c_void_p(0)
            tl = ctypes.c_int(0)
            ok = self._dll.PEGetNthFormulaEx(
                self._handle, i,
                ctypes.byref(nh), ctypes.byref(nl),
                ctypes.byref(th), ctypes.byref(tl),
            )
            name = f"Formula{i}"
            text = ""
            if ok:
                name = _handle_to_str(self._dll, nh.value) or name
                text = _handle_to_str(self._dll, th.value)
            formulas.append(FormulaInfo(index=i, name=name, text=text))
        return formulas

    def get_formula(self, name: str) -> str:
        """Get formula text by name (e.g. ``"FileName"`` or ``"{@FileName}"``)."""
        clean = name.strip("{}").lstrip("@")
        th = ctypes.c_void_p(0)
        tl = ctypes.c_int(0)
        ok = self._dll.PEGetFormulaW(
            self._handle, clean,
            ctypes.byref(th), ctypes.byref(tl),
        )
        if ok:
            return _handle_to_str(self._dll, th.value)
        return ""

    def set_formula(self, name: str, text: str) -> None:
        """Set formula text by name."""
        clean = name.strip("{}")
        if not clean.startswith("@"):
            clean = "@" + clean
        ok = self._dll.PESetFormula(
            self._handle,
            clean.encode("latin-1"),
            text.encode("latin-1"),
        )
        _check(self._dll, self._handle, ok, f"SetFormula({name})")

    # -- SQL --

    def get_sql_query(self) -> str:
        sh = ctypes.c_void_p(0)
        sl = ctypes.c_int(0)
        ok = self._dll.PEGetSQLQueryEx(self._handle, ctypes.byref(sh), ctypes.byref(sl))
        if ok:
            return _handle_to_str(self._dll, sh.value)
        return ""

    def set_sql_query(self, sql: str) -> None:
        ok = self._dll.PESetSQLQuery(self._handle, sql.encode("latin-1"))
        _check(self._dll, self._handle, ok, "SetSQLQuery")

    # -- Sections --

    def get_sections(self) -> list[SectionInfo]:
        n = self._dll.PEGetNSections(self._handle)
        return [SectionInfo(index=i, code=0) for i in range(n)]

    # -- Parameters --

    def get_parameters(self) -> list[ParameterInfo]:
        n = self._dll.PEGetNParameterFields(self._handle)
        return [ParameterInfo(index=i) for i in range(n)]

    # -- Sort fields --

    def get_sort_fields(self) -> list[SortFieldInfo]:
        n = self._dll.PEGetNSortFields(self._handle)
        return [SortFieldInfo(index=i) for i in range(n)]

    # -- Groups --

    def get_n_groups(self) -> int:
        return self._dll.PEGetNGroups(self._handle)

    # -- Export --

    def export(self, output_path: Union[str, Path],
               fmt: ExportFormat = ExportFormat.PDF) -> None:
        """Export the report to the given format."""
        output_path = str(Path(output_path).resolve())
        format_dll = _FORMAT_DLLS.get(fmt)
        if format_dll is None:
            raise ExportError(f"Unsupported export format: {fmt}")

        dest_opts = DiskDestOptions()
        dest_opts.StructSize = ctypes.sizeof(DiskDestOptions)
        dest_opts.fileName = output_path.encode("latin-1")

        opts = PEExportOptions()
        opts.StructSize = ctypes.sizeof(PEExportOptions)
        opts.formatDLLName = format_dll
        opts.destinationDLLName = _DEST_DISK_DLL
        opts.nFormatOptions = 0
        opts.formatOptions = None
        opts.nDestinationOptions = ctypes.sizeof(DiskDestOptions)
        opts.destinationOptions = ctypes.cast(ctypes.byref(dest_opts), ctypes.c_void_p)

        ok = self._dll.PEExportTo(self._handle, ctypes.byref(opts))
        if not ok:
            msg = _get_error_text(self._dll, self._handle)
            raise ExportError(f"Export to {fmt.value} failed: {msg}")

    # -- Save --

    def save(self, output_path: Optional[Union[str, Path]] = None) -> None:
        """Save modifications back to an .rpt file."""
        dest = str(Path(output_path).resolve()) if output_path else self._path
        try:
            ok = self._dll.PESavePrintJob(self._handle, dest.encode("latin-1"))
            _check(self._dll, self._handle, ok, "SavePrintJob")
        except AttributeError:
            raise CrystalReportsError(
                "PESavePrintJob not available in this crpe32.dll build"
            )


# ------------------------------------------------------------------
# CrpeEngine — manages DLL lifecycle
# ------------------------------------------------------------------

class CrpeEngine:
    """High-level wrapper around the CRPE print engine.

    Use as a context manager::

        with CrpeEngine() as engine:
            job = engine.open("report.rpt")
            print(job.get_tables())
    """

    def __init__(self):
        self._dll: Optional[ctypes.WinDLL] = None
        self._opened = False

    def _ensure_engine(self) -> ctypes.WinDLL:
        if self._dll is None:
            self._dll = _load_dll()
        if not self._opened:
            ok = self._dll.PEOpenEngine()
            if not ok:
                raise CrystalReportsError("PEOpenEngine failed")
            self._opened = True
        return self._dll

    def open(self, path: Union[str, Path], suppress_dialogs: bool = True) -> CrpeJob:
        """Open a Crystal Reports file and return a :class:`CrpeJob`."""
        dll = self._ensure_engine()
        rpt_path = str(Path(path).resolve())
        handle = dll.PEOpenPrintJob(rpt_path.encode("latin-1"))
        if handle < 0:
            raise ReportOpenError(f"PEOpenPrintJob failed for: {rpt_path}")

        if suppress_dialogs:
            try:
                dll.PESetDialogParentWindowHandle(handle, 0)
            except Exception:
                pass

        return CrpeJob(dll, handle, rpt_path)

    def close(self) -> None:
        """Close the CRPE engine."""
        if self._dll is not None and self._opened:
            try:
                self._dll.PECloseEngine()
            except Exception:
                pass
            self._opened = False

    def __enter__(self):
        self._ensure_engine()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        return False
