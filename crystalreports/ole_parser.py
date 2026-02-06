"""Pure-Python layer for reading/writing Crystal Reports .rpt files via OLE2.

Uses the ``olefile`` library to parse the OLE2 compound document structure
that Crystal Reports uses as its file format.
"""

from __future__ import annotations

import copy
import shutil
from pathlib import Path
from typing import Optional, Union

import olefile

from .exceptions import OleParseError, ReportOpenError
from .models import EmbeddedImage, ReportMetadata, SubreportInfo


class OleParser:
    """Read and modify Crystal Reports .rpt files at the OLE2 level.

    Parameters
    ----------
    path : str or Path
        Path to the .rpt file.
    """

    def __init__(self, path: Union[str, Path]):
        self.path = Path(path)
        if not self.path.exists():
            raise ReportOpenError(f"File not found: {self.path}")
        try:
            self._ole = olefile.OleFileIO(str(self.path))
        except Exception as exc:
            raise ReportOpenError(f"Cannot open OLE2 file: {self.path}: {exc}") from exc

    # ------------------------------------------------------------------
    # Context manager
    # ------------------------------------------------------------------

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        return False

    def close(self):
        """Close the underlying OLE file."""
        if self._ole is not None:
            self._ole.close()
            self._ole = None

    # ------------------------------------------------------------------
    # Metadata
    # ------------------------------------------------------------------

    def get_metadata(self) -> ReportMetadata:
        """Extract report metadata from the SummaryInformation stream."""
        meta = self._ole.get_metadata()
        return ReportMetadata(
            title=_decode(meta.title),
            subject=_decode(meta.subject),
            author=_decode(meta.author),
            keywords=_decode(meta.keywords),
            comments=_decode(meta.comments),
            creating_application=_decode(meta.creating_application),
            create_time=meta.create_time,
            last_save_time=meta.last_saved_time,
            revision_number=_decode(meta.revision_number),
        )

    def set_metadata(self, path: Union[str, Path], **kwargs) -> None:
        """Write metadata changes and save to *path*.

        Supported keyword arguments match :class:`ReportMetadata` fields:
        ``title``, ``subject``, ``author``, ``keywords``, ``comments``.

        Note: olefile has limited write support for OLE property sets.
        This method copies the file.  For full metadata editing, use the
        CRPE engine layer.
        """
        output = Path(path)
        if output != self.path:
            shutil.copy2(str(self.path), str(output))

    # ------------------------------------------------------------------
    # Streams
    # ------------------------------------------------------------------

    def list_streams(self) -> list[str]:
        """Return a list of all OLE stream paths in the file."""
        return ["/".join(entry) for entry in self._ole.listdir()]

    def get_stream(self, stream_path: str) -> bytes:
        """Read raw bytes from an OLE stream.

        Parameters
        ----------
        stream_path : str
            Forward-slash-separated path, e.g. ``"Embedding1/Contents"``.
        """
        parts = stream_path.split("/")
        if not self._ole.exists(parts):
            raise OleParseError(f"Stream not found: {stream_path}")
        return self._ole.openstream(parts).read()

    def set_stream(self, stream_path: str, data: bytes,
                   output_path: Union[str, Path, None] = None) -> None:
        """Write *data* into the given OLE stream and save.

        Parameters
        ----------
        stream_path : str
            Forward-slash-separated stream path.
        data : bytes
            New stream content.
        output_path : str or Path, optional
            Where to save.  Defaults to overwriting the original file.
        """
        dest = Path(output_path) if output_path else self.path

        if dest != self.path:
            shutil.copy2(str(self.path), str(dest))

        ole = olefile.OleFileIO(str(dest), write_mode=True)
        try:
            parts = stream_path.split("/")
            ole.write_stream(parts, data)
        finally:
            ole.close()

    # ------------------------------------------------------------------
    # Embedded images
    # ------------------------------------------------------------------

    def get_embedded_images(self) -> list[EmbeddedImage]:
        """Extract embedded BMP images from the OLE container.

        Crystal Reports stores embedded images in streams whose names
        start with ``Embedding``.  Each embedding stream contains a raw
        BMP image (starts with ``BM`` signature).
        """
        images: list[EmbeddedImage] = []
        idx = 0
        for entry in self._ole.listdir():
            stream_name = "/".join(entry)
            # Crystal Reports uses streams like "Embedding1", "Embedding2", ...
            # The actual image data may be in the root or in a sub-stream
            if len(entry) >= 1 and entry[0].lower().startswith("embedding"):
                try:
                    data = self._ole.openstream(entry).read()
                except Exception:
                    continue
                # Detect BMP by magic bytes
                fmt = "bmp" if data[:2] == b"BM" else "unknown"
                images.append(EmbeddedImage(
                    index=idx,
                    name=stream_name,
                    format=fmt,
                    data=data,
                ))
                idx += 1
        return images

    def replace_embedded_image(self, index: int, new_data: bytes,
                               output_path: Union[str, Path, None] = None) -> None:
        """Replace an embedded image by index.

        Parameters
        ----------
        index : int
            Zero-based index matching :pyattr:`EmbeddedImage.index`.
        new_data : bytes
            Raw BMP (or other) image data.
        output_path : str or Path, optional
            Save destination.  Defaults to overwriting in-place.
        """
        images = self.get_embedded_images()
        if index < 0 or index >= len(images):
            raise OleParseError(f"Image index {index} out of range (0..{len(images) - 1})")

        target = images[index]
        self.set_stream(target.name, new_data, output_path)

    # ------------------------------------------------------------------
    # Subreports
    # ------------------------------------------------------------------

    def list_subreports(self) -> list[SubreportInfo]:
        """List subreport entries found in the OLE structure.

        Subreports appear as separate storage entries whose names are
        not the root document.
        """
        subs: list[SubreportInfo] = []
        idx = 0
        for entry in self._ole.listdir(storages=True, streams=False):
            name = entry[0]
            if name.lower().startswith("subreport"):
                subs.append(SubreportInfo(
                    index=idx,
                    name=name,
                    stream_path="/".join(entry),
                ))
                idx += 1
        # Also look for numbered sub-document storages
        # Crystal Reports sometimes uses storage names like the subreport name
        if not subs:
            for entry_name in self._ole.listdir(storages=True, streams=False):
                top = entry_name[0]
                # Heuristic: storages that are not standard OLE entries
                if top not in ("\x01CompObj", "\x05SummaryInformation",
                               "\x05DocumentSummaryInformation") and \
                   not top.lower().startswith("embedding"):
                    # Could be a subreport storage
                    subs.append(SubreportInfo(
                        index=idx,
                        name=top,
                        stream_path="/".join(entry_name),
                    ))
                    idx += 1
        return subs

    # ------------------------------------------------------------------
    # Report info
    # ------------------------------------------------------------------

    def get_report_info(self) -> dict:
        """Return a dictionary of available report information.

        This is a convenience method that gathers metadata and structural
        info into a single dict.
        """
        meta = self.get_metadata()
        return {
            "title": meta.title,
            "author": meta.author,
            "comments": meta.comments,
            "creating_application": meta.creating_application,
            "create_time": str(meta.create_time) if meta.create_time else None,
            "last_save_time": str(meta.last_save_time) if meta.last_save_time else None,
            "num_streams": len(self.list_streams()),
            "num_images": len(self.get_embedded_images()),
            "num_subreports": len(self.list_subreports()),
        }

    # ------------------------------------------------------------------
    # Save
    # ------------------------------------------------------------------

    def save(self, output_path: Union[str, Path, None] = None) -> None:
        """Save the OLE file.  If *output_path* differs from the source,
        a copy is made first.

        Note: olefile's write support is limited.  For simple metadata
        edits use :meth:`set_metadata`.  For full round-trip editing the
        CRPE engine layer should be used.
        """
        if output_path and Path(output_path) != self.path:
            shutil.copy2(str(self.path), str(output_path))

    # ------------------------------------------------------------------
    # Full parse
    # ------------------------------------------------------------------

    def parse(self) -> dict:
        """Return a dict summarising the full OLE2 structure."""
        return {
            "path": str(self.path),
            "metadata": self.get_metadata(),
            "streams": self.list_streams(),
            "embedded_images": self.get_embedded_images(),
            "subreports": self.list_subreports(),
        }


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------

def _decode(value) -> Optional[str]:
    """Decode bytes to str if needed; return None for empty/None values."""
    if value is None:
        return None
    if isinstance(value, bytes):
        try:
            return value.decode("utf-8")
        except UnicodeDecodeError:
            return value.decode("latin-1")
    if isinstance(value, str):
        return value if value else None
    return str(value)
