# crystalreports

Python library for reading and modifying Crystal Reports `.rpt` files.

## Features

Two-layer architecture:

- **OLE layer** (pure Python, always available) — reads metadata, embedded images, subreports, and raw OLE streams using `olefile`
- **CRPE engine layer** (requires Crystal Reports installed) — reads/writes tables, formulas, SQL queries, parameters, sections via `crpe32.dll`

## Installation

```bash
pip install .
```

Or in development mode:

```bash
pip install -e ".[dev]"
```

## Quick Start

```python
from crystalreports import CrystalReport

with CrystalReport("report.rpt") as rpt:
    # Always available (OLE layer)
    print(rpt.metadata)
    print(rpt.embedded_images)
    print(rpt.subreports)
    print(rpt.streams)

    # Requires Crystal Reports SDK
    if rpt.has_sdk:
        print(rpt.tables)
        print(rpt.formulas)
        print(rpt.sql_query)
        print(rpt.parameters)
        print(rpt.sections)

        # Export to PDF
        rpt.export("output.pdf", fmt="pdf")
```

## OLE Layer (Pure Python)

Works on any system with Python and `olefile`:

```python
from crystalreports import OleParser

with OleParser("report.rpt") as ole:
    meta = ole.get_metadata()        # Title, author, dates
    streams = ole.list_streams()     # All OLE stream paths
    images = ole.get_embedded_images()  # Embedded BMP images
    subs = ole.list_subreports()     # Subreport entries
    data = ole.get_stream("path")    # Raw stream bytes
```

## CRPE Engine Layer (Native SDK)

Requires Crystal Reports 2020 (or compatible) installed. The library auto-detects `crpe32.dll`:

```python
from crystalreports.crpe_engine import CrpeEngine

with CrpeEngine() as engine:
    job = engine.open("report.rpt")

    tables = job.get_tables()
    formulas = job.get_formulas()
    sql = job.get_sql_query()

    # Modify
    job.set_formula("@MyFormula", "1 + 1")
    job.set_table_location(0, location="NewTable")

    # Export
    job.export("output.pdf")

    job.close()
```

You can also set the DLL path via environment variable:

```bash
set CRPE32_DLL_PATH=C:\path\to\crpe32.dll
```

## Requirements

- Python >= 3.10
- `olefile` >= 0.46
- Crystal Reports 2020 (optional, for CRPE engine layer)

## License

MIT
