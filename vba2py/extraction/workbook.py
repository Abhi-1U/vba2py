"""Load an Excel workbook and recover the original VBA module text."""

import re
from dataclasses import dataclass
from pathlib import Path
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from ..model import CompileError


class ExtractionError(CompileError):
    """The workbook could not be read as a VBA project."""


@dataclass(frozen=True)
class SourceModule:
    name: str
    source: str
    kind: str = "standard"


def _kinds(parser):
    result = {}
    for _, project_path, _ in parser.find_vba_projects() or ():
        with parser.ole_file.openstream(project_path) as stream:
            contents = stream.read()
        # PROJECT names use the VBA code page. VB_Name in decoded sources is
        # authoritative for names. Compare raw project names when Unicode applies.
        for line in contents.splitlines():
            category, _, value = line.partition(b"=")
            kind = {b"document": "document", b"class": "class",
                    b"module": "standard", b"baseclass": "form"}.get(category.lower())
            if kind and value:
                name = value.split(b"/", 1)[0]
                result[name.decode("cp1252", errors="replace").casefold()] = kind
    return result


def extract_modules(workbook) -> tuple[SourceModule, ...]:
    """Extract VBA from openpyxl's retained macro-enabled workbook archive."""
    from oletools.olevba import VBA_Parser

    path = Path(workbook)
    if path.suffix.lower() not in {".xlsm", ".xltm"}:
        raise ExtractionError("VBA extraction requires an .xlsm or .xltm workbook")
    book = None
    try:
        book = load_workbook(path, keep_vba=True, data_only=False)
        archive = book.vba_archive
        parts = [name for name in archive.namelist() if name.casefold().endswith("/vbaproject.bin")]
        if "xl/vbaProject.bin" in parts:
            part = "xl/vbaProject.bin"
        elif len(parts) == 1:
            part = parts[0]
        else:
            raise ExtractionError(f"{path} does not contain one VBA project part")
        parser = VBA_Parser(part, data=archive.read(part), relaxed=False)
        try:
            kinds = _kinds(parser)
            document_names = {str(book.code_name or "ThisWorkbook").casefold()}
            document_names.update(str(sheet.sheet_properties.codeName).casefold() for sheet in book.worksheets
                                  if sheet.sheet_properties.codeName)
            modules = []
            for _, stream, filename, source in parser.extract_macros():
                if not isinstance(source, str):
                    raise ExtractionError(f"Could not decode VBA module {filename!r}")
                attribute = re.search(r'^Attribute\s+VB_Name\s*=\s*"([^"]+)"', source, re.I | re.M)
                name = attribute[1] if attribute else Path(filename).stem
                extension = Path(filename).suffix.casefold()
                kind = kinds.get(name.casefold())
                if kind is None:
                    if name.casefold() in document_names or re.search(r'^Attribute\s+VB_Base\s*=\s*"0\{000208', source, re.I | re.M):
                        kind = "document"
                    else:
                        kind = {".bas": "standard", ".cls": "class", ".frm": "form"}.get(extension)
                if kind is None:
                    raise ExtractionError(f"Cannot determine module type for {filename!r} ({stream})")
                modules.append(SourceModule(name, source, kind))
            if not modules:
                raise ExtractionError(f"{path} contains no VBA source modules")
            return tuple(modules)
        finally:
            parser.close()
    except ExtractionError:
        raise
    except (OSError, BadZipFile, InvalidFileException, ValueError, KeyError) as error:
        raise ExtractionError(f"Could not extract VBA from {path}: {error}") from error
    finally:
        if book is not None:
            book.close()
            if book.vba_archive is not None:
                book.vba_archive.close()
