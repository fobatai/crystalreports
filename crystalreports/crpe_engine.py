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
import struct as _struct
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
    ReportObject,
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


class PEObjectInfo(ctypes.Structure):
    _fields_ = [
        ("StructSize", PE_WORD),
        ("ObjectType", PE_WORD),
        ("Left", PE_LONG),
        ("Top", PE_LONG),
        ("Right", PE_LONG),
        ("Bottom", PE_LONG),
        ("SectionN", PE_WORD),
    ]


_SECTION_AREA_NAMES = {
    1: "Report Header",
    2: "Page Header",
    3: "Group Header",
    4: "Detail",
    5: "Group Footer",
    6: "Report Footer",
    7: "Page Footer",
}

_OBJECT_TYPE_NAMES = {
    1: "Field", 2: "Text", 3: "Line", 4: "Box",
    5: "Subreport", 6: "OLE", 7: "Graph", 8: "CrossTab",
    9: "BLOB", 10: "Map",
}

# Section codes use: area * 6000 + sub_section * 50
_SECTION_AREA_MULTIPLIER = 6000
_SECTION_SUB_STEP = 50
# Scan format probe struct (minimal — only StructSize matters)
_SECTION_PROBE_SIZE = 28  # enough bytes for PEGetSectionFormat


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

    # Section format (for discovering valid section codes)
    dll.PEGetSectionFormat.argtypes = [
        PE_HANDLE, ctypes.c_short,
        ctypes.c_void_p,  # pointer to PESectionOptions (variable size)
    ]
    dll.PEGetSectionFormat.restype = ctypes.c_bool

    # Object enumeration in sections
    dll.PEGetNObjectsInSection.argtypes = [PE_HANDLE, ctypes.c_short]
    dll.PEGetNObjectsInSection.restype = ctypes.c_short
    dll.PEGetNthObjectInSection.argtypes = [PE_HANDLE, ctypes.c_short, ctypes.c_short]
    dll.PEGetNthObjectInSection.restype = ctypes.c_int  # DWORD handle

    # Object info
    dll.PEGetObjectInfo.argtypes = [
        PE_HANDLE, ctypes.c_int, ctypes.c_void_p,
    ]
    dll.PEGetObjectInfo.restype = ctypes.c_bool

    # Object name
    dll.PEGetObjectName.argtypes = [
        PE_HANDLE, ctypes.c_int,
        ctypes.POINTER(ctypes.c_void_p), ctypes.POINTER(ctypes.c_int),
    ]
    dll.PEGetObjectName.restype = ctypes.c_bool

    # Object modification — SetObjectInfo (move/resize)
    dll.PESetObjectInfo.argtypes = [PE_HANDLE, ctypes.c_int, ctypes.c_void_p]
    dll.PESetObjectInfo.restype = ctypes.c_bool

    # Field font — per-object font changes (works on Field objects)
    dll.PESetFieldFont.argtypes = [
        PE_HANDLE, ctypes.c_int,  # job, objectHandle
        ctypes.c_char_p,  # faceName
        ctypes.c_short,   # fontFamily (0=don't change)
        ctypes.c_short,   # fontPitch (0=don't change)
        ctypes.c_short,   # charSet (0=don't change)
        ctypes.c_short,   # pointSize (0=don't change)
        ctypes.c_short,   # isItalic (PE_UNCHANGED=3)
        ctypes.c_short,   # isUnderlined (PE_UNCHANGED=3)
        ctypes.c_short,   # isStrikeOut (PE_UNCHANGED=3)
        ctypes.c_short,   # fontWeight (0=don't change, 400=normal, 700=bold)
    ]
    dll.PESetFieldFont.restype = ctypes.c_bool

    # Section-wide font — scope: 1=fields, 2=text, 3=both
    dll.PESetFont.argtypes = [
        PE_HANDLE, ctypes.c_short, ctypes.c_short,  # job, sectionCode, scope
        ctypes.c_char_p,  # faceName
        ctypes.c_short,   # fontFamily
        ctypes.c_short,   # fontPitch
        ctypes.c_short,   # charSet
        ctypes.c_short,   # pointSize
        ctypes.c_short,   # isItalic
        ctypes.c_short,   # isUnderlined
        ctypes.c_short,   # isStrikeOut
        ctypes.c_short,   # fontWeight
    ]
    dll.PESetFont.restype = ctypes.c_bool

    # Font color — MUST pass pointer to COLORREF, not direct value
    dll.PESetObjectFontColor.argtypes = [
        PE_HANDLE, ctypes.c_int, ctypes.POINTER(ctypes.c_ulong),
    ]
    dll.PESetObjectFontColor.restype = ctypes.c_bool

    # Section height — Get uses out-param, Set uses direct value
    dll.PEGetSectionHeight.argtypes = [
        PE_HANDLE, ctypes.c_short, ctypes.POINTER(ctypes.c_long),
    ]
    dll.PEGetSectionHeight.restype = ctypes.c_bool
    dll.PESetSectionHeight.argtypes = [PE_HANDLE, ctypes.c_short, ctypes.c_long]
    dll.PESetSectionHeight.restype = ctypes.c_bool

    # Section format (read/write section properties)
    dll.PESetSectionFormat.argtypes = [
        PE_HANDLE, ctypes.c_short, ctypes.c_void_p,
    ]
    dll.PESetSectionFormat.restype = ctypes.c_bool

    # Box object info (move/resize boxes)
    dll.PEGetBoxObjectInfo.argtypes = [PE_HANDLE, ctypes.c_int, ctypes.c_void_p]
    dll.PEGetBoxObjectInfo.restype = ctypes.c_bool
    dll.PESetBoxObjectInfo.argtypes = [PE_HANDLE, ctypes.c_int, ctypes.c_void_p]
    dll.PESetBoxObjectInfo.restype = ctypes.c_bool

    # Line object info (move/resize lines)
    dll.PEGetLineObjectInfo.argtypes = [PE_HANDLE, ctypes.c_int, ctypes.c_void_p]
    dll.PEGetLineObjectInfo.restype = ctypes.c_bool
    dll.PESetLineObjectInfo.argtypes = [PE_HANDLE, ctypes.c_int, ctypes.c_void_p]
    dll.PESetLineObjectInfo.restype = ctypes.c_bool

    # Delete object
    dll.PEDeleteObject.argtypes = [PE_HANDLE, ctypes.c_int]
    dll.PEDeleteObject.restype = ctypes.c_bool

    # Margins
    dll.PEGetMargins.argtypes = [
        PE_HANDLE,
        ctypes.POINTER(ctypes.c_short), ctypes.POINTER(ctypes.c_short),
        ctypes.POINTER(ctypes.c_short), ctypes.POINTER(ctypes.c_short),
    ]
    dll.PEGetMargins.restype = ctypes.c_bool
    dll.PESetMargins.argtypes = [
        PE_HANDLE, ctypes.c_short, ctypes.c_short,
        ctypes.c_short, ctypes.c_short,
    ]
    dll.PESetMargins.restype = ctypes.c_bool

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


def rgb_to_colorref(r: int, g: int, b: int) -> int:
    """Convert RGB values (0-255) to a Windows COLORREF (0x00BBGGRR)."""
    return (b << 16) | (g << 8) | r


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
        """Set formula text by name.

        The name should be the bare formula name (e.g. ``"FileName"``),
        without the ``@`` or ``{@...}`` wrapper.
        """
        clean = name.strip("{}").lstrip("@")
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

    def get_section_codes(self) -> list[int]:
        """Discover all valid section codes by probing.

        Section codes follow the pattern ``area * 6000 + sub * 50``.
        """
        codes: list[int] = []
        probe = ctypes.create_string_buffer(_SECTION_PROBE_SIZE)
        struct_size = PE_WORD(_SECTION_PROBE_SIZE)
        ctypes.memmove(probe, ctypes.byref(struct_size), 2)  # set StructSize
        for area in range(1, 9):
            base = area * _SECTION_AREA_MULTIPLIER
            for sub in range(100):  # max 100 sub-sections per area
                code = base + sub * _SECTION_SUB_STEP
                ok = self._dll.PEGetSectionFormat(
                    self._handle, code, probe,
                )
                if ok:
                    codes.append(code)
                elif sub > 0:
                    break  # no more sub-sections in this area
        return codes

    # -- Objects --

    def get_objects_in_section(self, section_code: int) -> list[ReportObject]:
        """Enumerate all report objects in the given section."""
        n = self._dll.PEGetNObjectsInSection(self._handle, section_code)
        if n <= 0:
            return []
        objects: list[ReportObject] = []
        for i in range(n):
            oh = self._dll.PEGetNthObjectInSection(self._handle, section_code, i)
            info = PEObjectInfo()
            info.StructSize = ctypes.sizeof(PEObjectInfo)
            ok = self._dll.PEGetObjectInfo(
                self._handle, oh, ctypes.byref(info),
            )
            name = ""
            nh = ctypes.c_void_p(0)
            nl = ctypes.c_int(0)
            ok2 = self._dll.PEGetObjectName(
                self._handle, oh, ctypes.byref(nh), ctypes.byref(nl),
            )
            if ok2 and nh.value:
                name = _handle_to_str(self._dll, nh.value)
            obj_type = _OBJECT_TYPE_NAMES.get(info.ObjectType,
                                               f"Unknown({info.ObjectType})")
            objects.append(ReportObject(
                handle=oh,
                name=name,
                object_type=obj_type,
                section_code=section_code,
                left=info.Left,
                top=info.Top,
                right=info.Right,
                bottom=info.Bottom,
            ))
        return objects

    def get_all_objects(self) -> list[ReportObject]:
        """Enumerate all report objects across all sections."""
        all_objects: list[ReportObject] = []
        for code in self.get_section_codes():
            all_objects.extend(self.get_objects_in_section(code))
        return all_objects

    # -- Object modification --

    def move_object(self, handle: int, left: int, top: int,
                    right: int, bottom: int,
                    section_code: int = 0) -> None:
        """Move/resize an object by setting its bounds (in twips).

        Automatically detects the object type and uses the correct API:
        - Box objects use ``PESetBoxObjectInfo``
        - Line objects use ``PESetLineObjectInfo``
        - Other objects use ``PESetObjectInfo``

        Parameters
        ----------
        section_code : int, optional
            Required for Box/Line objects to auto-increase section height
            when the object bottom would exceed it.
        """
        # Read object info to determine type
        info = PEObjectInfo()
        info.StructSize = ctypes.sizeof(PEObjectInfo)
        ok = self._dll.PEGetObjectInfo(self._handle, handle, ctypes.byref(info))
        _check(self._dll, self._handle, ok, f"GetObjectInfo({handle})")

        obj_type = info.ObjectType
        if obj_type == 4:  # Box
            self._move_box(handle, left, top, right, bottom, section_code)
        elif obj_type == 3:  # Line
            self._move_line(handle, left, top, right, bottom, section_code)
        else:
            info.Left = left
            info.Top = top
            info.Right = right
            info.Bottom = bottom
            ok = self._dll.PESetObjectInfo(
                self._handle, handle, ctypes.byref(info),
            )
            _check(self._dll, self._handle, ok, f"SetObjectInfo({handle})")

    def _ensure_section_height(self, section_code: int, needed: int) -> None:
        """Increase section height if *needed* exceeds current height."""
        if section_code <= 0:
            return
        height = ctypes.c_long(0)
        ok = self._dll.PEGetSectionHeight(
            self._handle, section_code, ctypes.byref(height),
        )
        if ok and needed > height.value:
            self._dll.PESetSectionHeight(
                self._handle, section_code, needed + 20,
            )

    @staticmethod
    def _set_long_at(buf, offset: int, value: int) -> None:
        """Write a signed 32-bit int into *buf* at *offset*."""
        b = value.to_bytes(4, "little", signed=True)
        for i in range(4):
            buf[offset + i] = b[i]

    def _move_box(self, handle: int, left: int, top: int,
                  right: int, bottom: int, section_code: int) -> None:
        """Move/resize a Box using PESetBoxObjectInfo.

        BoxObjectInfo (44 bytes) stores coordinates as:
          offset 4:  left
          offset 8:  top
          offset 16: left + right  (x-extent sum)
          offset 20: top + bottom  (y-extent sum, clamped to section height)

        Cross-section boxes (spanning multiple sub-sections) have a
        different encoding for offset 12/20.  For these boxes, only
        left and top can be reliably changed; right and bottom are
        preserved as-is because PEObjectInfo does not propagate x-sum
        changes for cross-section boxes.
        """
        buf = (ctypes.c_byte * 44)()
        ctypes.cast(buf, ctypes.POINTER(ctypes.c_ushort))[0] = 44
        ok = self._dll.PEGetBoxObjectInfo(self._handle, handle, buf)
        _check(self._dll, self._handle, ok, f"GetBoxObjectInfo({handle})")

        # Detect cross-section box via offset 12
        is_cross = False
        if section_code > 0:
            off12 = _struct.unpack_from("<i", bytes(buf), 12)[0]
            is_cross = off12 != section_code + 65536

        if is_cross:
            # Cross-section: only left/top propagate to PEObjectInfo
            self._set_long_at(buf, 4, left)
            self._set_long_at(buf, 8, top)
            # Keep offset 16 (x-sum) and 20 (cross-section bottom) unchanged
        else:
            self._ensure_section_height(section_code, top + bottom)
            self._set_long_at(buf, 4, left)
            self._set_long_at(buf, 8, top)
            self._set_long_at(buf, 16, left + right)
            self._set_long_at(buf, 20, top + bottom)

        ok = self._dll.PESetBoxObjectInfo(self._handle, handle, buf)
        _check(self._dll, self._handle, ok, f"SetBoxObjectInfo({handle})")

    def _move_line(self, handle: int, left: int, top: int,
                   right: int, bottom: int, section_code: int) -> None:
        """Move/resize a Line using PESetLineObjectInfo.

        LineObjectInfo (36 bytes) stores coordinates as:
          offset 8:  left
          offset 12: top
          offset 20: left + right  (x-extent sum)
          offset 24: top + bottom  (y-extent sum, clamped to section height)

        Cross-section lines have offset 4 != section_code + 65536;
        for those only left and top are reliably modified.
        """
        buf = (ctypes.c_byte * 36)()
        ctypes.cast(buf, ctypes.POINTER(ctypes.c_ushort))[0] = 36
        ok = self._dll.PEGetLineObjectInfo(self._handle, handle, buf)
        _check(self._dll, self._handle, ok, f"GetLineObjectInfo({handle})")

        # Detect cross-section line via offset 4
        is_cross = False
        if section_code > 0:
            off4 = _struct.unpack_from("<i", bytes(buf), 4)[0]
            is_cross = off4 != section_code + 65536

        if is_cross:
            self._set_long_at(buf, 8, left)
            self._set_long_at(buf, 12, top)
        else:
            self._ensure_section_height(section_code, top + bottom)
            self._set_long_at(buf, 8, left)
            self._set_long_at(buf, 12, top)
            self._set_long_at(buf, 20, left + right)
            self._set_long_at(buf, 24, top + bottom)

        ok = self._dll.PESetLineObjectInfo(self._handle, handle, buf)
        _check(self._dll, self._handle, ok, f"SetLineObjectInfo({handle})")

    def set_field_font(self, handle: int, face_name: str = "",
                       point_size: int = 0, bold: Optional[bool] = None,
                       italic: Optional[bool] = None,
                       underline: Optional[bool] = None,
                       strikeout: Optional[bool] = None) -> None:
        """Change font for a Field object.

        Parameters
        ----------
        handle : int
            Object handle from :meth:`get_objects_in_section`.
        face_name : str
            Font face (e.g. ``"Arial"``).  Empty = no change.
        point_size : int
            Font size in points.  0 = no change.
        bold : bool or None
            ``True``/``False`` to set, ``None`` to leave unchanged.
        italic, underline, strikeout : bool or None
            Same behaviour as *bold*.
        """
        PE_UNCHANGED = 3
        name_bytes = face_name.encode("latin-1") if face_name else b""
        weight = 0  # don't change
        if bold is True:
            weight = 700
        elif bold is False:
            weight = 400
        ok = self._dll.PESetFieldFont(
            self._handle, handle, name_bytes,
            0, 0, 0,  # fontFamily, fontPitch, charSet — don't change
            point_size,
            (1 if italic is True else 0 if italic is False else PE_UNCHANGED),
            (1 if underline is True else 0 if underline is False else PE_UNCHANGED),
            (1 if strikeout is True else 0 if strikeout is False else PE_UNCHANGED),
            weight,
        )
        _check(self._dll, self._handle, ok, f"SetFieldFont({handle})")

    def set_section_font(self, section_code: int, face_name: str = "",
                         point_size: int = 0, bold: Optional[bool] = None,
                         italic: Optional[bool] = None,
                         scope: int = 1) -> None:
        """Change font for all objects in a section.

        Parameters
        ----------
        section_code : int
            Section code (e.g. 12000 for Page Header).
        scope : int
            1 = field objects, 2 = text objects, 3 = both.
        """
        # NOTE: PESetFont uses 0="off/don't change", 1="on" for booleans.
        # Unlike PESetFieldFont which uses PE_UNCHANGED=3.
        name_bytes = face_name.encode("latin-1") if face_name else b""
        weight = 0
        if bold is True:
            weight = 700
        elif bold is False:
            weight = 400
        ok = self._dll.PESetFont(
            self._handle, section_code, scope, name_bytes,
            0, 0, 0,  # fontFamily, fontPitch, charSet
            point_size,
            (1 if italic is True else 0),
            0,  # underline off
            0,  # strikeout off
            weight,
        )
        _check(self._dll, self._handle, ok,
               f"SetFont(section={section_code}, scope={scope})")

    def set_object_font_color(self, handle: int, color: int) -> None:
        """Set font color for an object.

        Parameters
        ----------
        handle : int
            Object handle.
        color : int
            COLORREF value (``0x00BBGGRR``).
            Use :func:`rgb_to_colorref` to convert from (r, g, b).
        """
        c = ctypes.c_ulong(color)
        ok = self._dll.PESetObjectFontColor(
            self._handle, handle, ctypes.byref(c),
        )
        _check(self._dll, self._handle, ok, f"SetObjectFontColor({handle})")

    def delete_object(self, handle: int) -> None:
        """Delete a report object."""
        ok = self._dll.PEDeleteObject(self._handle, handle)
        _check(self._dll, self._handle, ok, f"DeleteObject({handle})")

    def get_section_height(self, section_code: int) -> int:
        """Get section height in twips."""
        height = ctypes.c_long(0)
        ok = self._dll.PEGetSectionHeight(
            self._handle, section_code, ctypes.byref(height),
        )
        _check(self._dll, self._handle, ok,
               f"GetSectionHeight({section_code})")
        return height.value

    def set_section_height(self, section_code: int, height: int) -> None:
        """Set section height in twips."""
        ok = self._dll.PESetSectionHeight(self._handle, section_code, height)
        _check(self._dll, self._handle, ok,
               f"SetSectionHeight({section_code})")

    def get_margins(self) -> tuple[int, int, int, int]:
        """Get page margins as ``(left, right, top, bottom)`` in twips."""
        ml = ctypes.c_short(0)
        mr = ctypes.c_short(0)
        mt = ctypes.c_short(0)
        mb = ctypes.c_short(0)
        ok = self._dll.PEGetMargins(
            self._handle,
            ctypes.byref(ml), ctypes.byref(mr),
            ctypes.byref(mt), ctypes.byref(mb),
        )
        _check(self._dll, self._handle, ok, "GetMargins")
        return ml.value, mr.value, mt.value, mb.value

    def set_margins(self, left: int, right: int,
                    top: int, bottom: int) -> None:
        """Set page margins in twips."""
        ok = self._dll.PESetMargins(self._handle, left, right, top, bottom)
        _check(self._dll, self._handle, ok, "SetMargins")

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
        """Save modifications to an .rpt file.

        .. note::

           The open file cannot be overwritten — you **must** provide a
           different *output_path* (Save-As).  If *output_path* is
           ``None`` or equals the currently-open path, a
           :class:`CrystalReportsError` is raised.
        """
        if output_path is None:
            raise CrystalReportsError(
                "PESavePrintJob cannot overwrite the open file. "
                "Provide a different output_path."
            )
        dest = str(Path(output_path).resolve())
        if dest == self._path:
            raise CrystalReportsError(
                "PESavePrintJob cannot overwrite the open file. "
                "Provide a different output_path."
            )
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
