"""The openpyxl-backed subset of the Excel object model."""

import os
from copy import copy
from pathlib import Path
from tempfile import NamedTemporaryFile

from openpyxl import Workbook as OpenpyxlWorkbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
from openpyxl.utils.exceptions import InvalidFileException
from zipfile import BadZipFile

from .arrays import VBAArray
from .errors import VBAError
from .filesystem import FileSystem, file_error
from .values import EMPTY, MISSING, NULL, ErrorValue, integer, string, truth

MAX_ROW = 1048576
MAX_COLUMN = 16384
MAX_MATERIALIZED_CELLS = 2_000_000
EXCEL_ERRORS = {2000: "#NULL!", 2007: "#DIV/0!", 2015: "#VALUE!", 2023: "#REF!",
                2029: "#NAME?", 2036: "#NUM!", 2042: "#N/A", 2043: "#GETTING_DATA"}
EXCEL_ERROR_CODES = {text: code for code, text in EXCEL_ERRORS.items()}


def _sheet_value(cell):
    if cell is None or cell.value is None:
        return EMPTY
    if cell.data_type == "e" and cell.value in EXCEL_ERROR_CODES:
        return ErrorValue(EXCEL_ERROR_CODES[cell.value])
    return cell.value


class Indexed:
    def _vba_item(self, *args):
        return self.Item(*args)


class Workbooks(Indexed):
    def __init__(self, app):
        self.app = app

    @property
    def Count(self):
        return len(self.app._books)

    def __iter__(self):
        return iter(tuple(self.app._books))

    def Item(self, Index):
        if isinstance(Index, str):
            book = next((item for item in self if item.Name.casefold() == Index.casefold()
                         or item.FullName.casefold() == Index.casefold()), None)
            if book is not None:
                return book
        else:
            index = integer(Index)
            if 1 <= index <= self.Count:
                return self.app._books[index - 1]
        raise VBAError(9)

    def Open(self, Filename, UpdateLinks=MISSING, ReadOnly=False, **options):
        if options or UpdateLinks is not MISSING and integer(UpdateLinks) != 0:
            raise VBAError(1004, "Unsupported Workbooks.Open options")
        path = self.app.files.path(Filename)
        for book in self:
            if book._path == path:
                book.Activate()
                return book
        try:
            raw = load_workbook(path, keep_vba=path.suffix.lower() in {".xlsm", ".xltm"}, data_only=False)
        except OSError as error:
            raise file_error(error) from None
        except (BadZipFile, InvalidFileException, ValueError) as error:
            raise VBAError(1004, str(error)) from None
        return self.app._attach(raw, path, truth(ReadOnly))

    def Add(self):
        return self.app._attach(OpenpyxlWorkbook())


class Application:
    def __init__(self, directory=None):
        self.files = FileSystem(directory)
        self._books = []
        self._unsaved = 0
        self.Workbooks = Workbooks(self)
        self.ActiveWorkbook = None
        self.Selection = None
        self.ActiveCell = None
        self.ScreenUpdating = True
        self.DisplayAlerts = True
        self.EnableEvents = True
        self.Calculation = -4105
        self.StatusBar = False
        self.CutCopyMode = False
        self._clipboard = None
        self.WorksheetFunction = WorksheetFunction()

    def _attach(self, raw, path=None, read_only=False):
        if path is None:
            self._unsaved += 1
        book = Workbook(self, raw, path, read_only)
        self._books.append(book)
        book.Activate()
        return book

    @property
    def ActiveSheet(self):
        if self.ActiveWorkbook is None:
            raise VBAError(91)
        return self.ActiveWorkbook.ActiveSheet

    @property
    def Worksheets(self):
        if self.ActiveWorkbook is None:
            raise VBAError(91)
        return self.ActiveWorkbook.Worksheets

    Sheets = Worksheets

    def Range(self, Cell1, Cell2=MISSING):
        return self.ActiveSheet.Range(Cell1, Cell2)

    @property
    def Cells(self):
        return self.ActiveSheet.Cells

    @property
    def Rows(self):
        return self.ActiveSheet.Rows

    @property
    def Columns(self):
        return self.ActiveSheet.Columns

    def Goto(self, Reference, Scroll=False):
        reference = Reference if isinstance(Reference, Range) else self.Range(Reference)
        reference.Select()

    def Quit(self):
        for book in tuple(self._books):
            book.Close(False)
        self.files.close()


class Workbook:
    def __init__(self, app, raw, path, read_only):
        self.Application = app
        self._raw = raw
        self._path = path
        self._name = f"Book{app._unsaved}" if path is None else path.name
        self._closed = False
        self._saved = path is not None
        self._cache = None
        self._cache_loaded = False
        self._cache_path = path
        self._changed_formulas = set()
        self._wrappers = {}
        self.ReadOnly = read_only
        self.Worksheets = Worksheets(self)
        self.Sheets = self.Worksheets

    def _check(self):
        if self._closed:
            raise VBAError(1004, "Workbook is closed")

    @property
    def Name(self):
        return self._name

    @property
    def Path(self):
        return str(self._path.parent) if self._path else ""

    @property
    def FullName(self):
        return str(self._path) if self._path else self.Name

    @property
    def Saved(self):
        return self._saved

    @Saved.setter
    def Saved(self, value):
        self._saved = truth(value)

    @property
    def CodeName(self):
        return self._raw.code_name or "ThisWorkbook"

    @property
    def ActiveSheet(self):
        self._check()
        return self._sheet(self._raw.active)

    def _sheet(self, raw):
        if raw not in self._wrappers:
            self._wrappers[raw] = Worksheet(self, raw)
        return self._wrappers[raw]

    def _cached_value(self, sheet, row, column):
        if (sheet, row, column) in self._changed_formulas or self._cache_path is None:
            return EMPTY
        if not self._cache_loaded:
            self._cache = load_workbook(self._cache_path, data_only=True)
            self._cache_loaded = True
        name = self._sheet(sheet)._original_name
        return _sheet_value(self._cache[name]._cells.get((row, column))) if name in self._cache else EMPTY

    def Activate(self):
        self._check()
        app = self.Application
        app.ActiveWorkbook = self
        sheet = self.ActiveSheet
        app.Selection = sheet._selection()
        app.ActiveCell = app.Selection.Item(1, 1)

    def Save(self):
        if self.ReadOnly:
            raise VBAError(1004, "Read-only workbook requires SaveAs")
        if self._path is None:
            raise VBAError(1004, "New workbook requires SaveAs")
        self.SaveAs(self._path)

    def SaveAs(self, Filename, FileFormat=MISSING, **options):
        self._check()
        if options:
            raise VBAError(1004, "Unsupported Workbook.SaveAs options")
        path = Filename if isinstance(Filename, Path) else self.Application.files.path(Filename)
        if not path.suffix:
            path = path.with_suffix(".xlsm" if self._raw.vba_archive else ".xlsx")
        formats = {".xlsx": 51, ".xlsm": 52, ".xltx": 54, ".xltm": 53}
        if path.suffix.lower() not in formats or FileFormat is not MISSING and integer(FileFormat) != formats[path.suffix.lower()]:
            raise VBAError(1004, "Unsupported workbook file format")
        if path.suffix.lower() in {".xlsm", ".xltm"} and self._raw.vba_archive is None:
            raise VBAError(1004, "A macro-enabled output requires a retained VBA project")
        if path == self._path and self.ReadOnly:
            raise VBAError(1004, "Workbook is read-only")
        if any(book is not self and book._path == path for book in self.Application.Workbooks):
            raise VBAError(1004, "A workbook with that path is already open")
        archive = self._raw.vba_archive
        was_template = self._raw.template
        if path.suffix.lower() in {".xlsx", ".xltx"}:
            self._raw.vba_archive = None
        self._raw.template = path.suffix.lower() in {".xltx", ".xltm"}
        temporary = None
        try:
            if self._cache_path == path and not self._cache_loaded:
                self._cache = load_workbook(path, data_only=True)
                self._cache_loaded = True
            with NamedTemporaryFile(dir=path.parent, suffix=path.suffix, delete=False) as handle:
                temporary = Path(handle.name)
            self._raw.save(temporary)
            os.replace(temporary, path)
        except OSError as error:
            self._raw.vba_archive, self._raw.template = archive, was_template
            raise file_error(error) from None
        except BaseException:
            self._raw.vba_archive, self._raw.template = archive, was_template
            raise
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)
        if archive is not None and self._raw.vba_archive is None:
            archive.close()
        self._saved = True
        self._path, self._name, self.ReadOnly = path, path.name, False

    def Close(self, SaveChanges=MISSING, Filename=MISSING, RouteWorkbook=MISSING):
        self._check()
        if RouteWorkbook is not MISSING:
            raise VBAError(1004, "Workbook routing is unsupported")
        if SaveChanges is not MISSING and truth(SaveChanges):
            self.Save() if Filename is MISSING else self.SaveAs(Filename)
        self._raw.close()
        if self._raw.vba_archive is not None:
            self._raw.vba_archive.close()
        if self._cache is not None:
            self._cache.close()
        self._closed = True
        app = self.Application
        app._books.remove(self)
        if app.ActiveWorkbook is self:
            if app._books:
                app._books[-1].Activate()
            else:
                app.ActiveWorkbook = app.ActiveCell = app.Selection = None


class Worksheets(Indexed):
    def __init__(self, book):
        self.book = book

    @property
    def Count(self):
        self.book._check()
        return len(self.book._raw.worksheets)

    def __iter__(self):
        self.book._check()
        return (self.book._sheet(sheet) for sheet in self.book._raw.worksheets)

    def Item(self, Index):
        if isinstance(Index, str):
            for sheet in self:
                if sheet.Name.casefold() == Index.casefold():
                    return sheet
        else:
            index = integer(Index)
            if 1 <= index <= self.Count:
                return self.book._sheet(self.book._raw.worksheets[index - 1])
        raise VBAError(9)

    def Add(self, Before=MISSING, After=MISSING, Count=1, Type=MISSING):
        self.book._check()
        if (Before is not MISSING and After is not MISSING or Type is not MISSING
                and integer(Type) != -4167 or integer(Count) < 1):
            raise VBAError(1004)
        existing = Before if Before is not MISSING else After
        if existing is MISSING:
            existing = self.book.ActiveSheet
        if not isinstance(existing, Worksheet) or existing.Parent is not self.book:
            raise VBAError(1004)
        position = self.book._raw.index(existing._raw) + int(After is not MISSING)
        created = []
        for offset in range(integer(Count)):
            created.append(self.book._sheet(self.book._raw.create_sheet(index=position + offset)))
        self.book._saved = False
        created[0].Activate()
        return created[0]


class Worksheet:
    def __init__(self, book, raw):
        self.Parent = book
        self.Application = book.Application
        self._raw = raw
        self._original_name = raw.title
        self._selected = "A1"

    @property
    def Name(self):
        self.Parent._check()
        return self._raw.title

    @Name.setter
    def Name(self, value):
        self.Parent._check()
        name = string(value)
        if len(name) > 31 or any(sheet is not self and sheet.Name.casefold() == name.casefold()
                                 for sheet in self.Parent.Worksheets):
            raise VBAError(1004)
        self._raw.title = name
        self.Parent._saved = False

    @property
    def CodeName(self):
        return self._raw.sheet_properties.codeName or ""

    @property
    def Index(self):
        return self.Parent._raw.index(self._raw) + 1

    @property
    def Visible(self):
        return {"visible": -1, "hidden": 0, "veryHidden": 2}[self._raw.sheet_state]

    @Visible.setter
    def Visible(self, value):
        visibility = integer(value)
        if visibility not in {-1, 0, 2}:
            raise VBAError(1004)
        self._raw.sheet_state = {-1: "visible", 0: "hidden", 2: "veryHidden"}[visibility]
        self.Parent._saved = False

    @property
    def Cells(self):
        return Range(self, 1, 1, MAX_ROW, MAX_COLUMN)

    @property
    def Rows(self):
        return self.Cells.Rows

    @property
    def Columns(self):
        return self.Cells.Columns

    @property
    def UsedRange(self):
        raw = self._raw
        return Range(self, raw.min_row, raw.min_column, raw.max_row, raw.max_column)

    def Range(self, Cell1, Cell2=MISSING):
        self.Parent._check()
        if Cell2 is not MISSING:
            first, last = self.Range(Cell1), self.Range(Cell2)
            if first.Parent is not self or last.Parent is not self:
                raise VBAError(1004)
            return Range(self, min(first.Row, last.Row), min(first.Column, last.Column),
                         max(first._bottom, last._bottom), max(first._right, last._right))
        if isinstance(Cell1, Range):
            return Cell1
        address = string(Cell1)
        if "!" in address:
            sheet, address = address.rsplit("!", 1)
            return self.Parent.Worksheets.Item(sheet.strip("'").replace("''", "'")).Range(address)
        try:
            left, top, right, bottom = range_boundaries(address)
        except ValueError:
            definition = self._raw.defined_names.get(address) or self.Parent._raw.defined_names.get(address)
            if definition is None:
                raise VBAError(1004, f"Invalid range {address!r}") from None
            destinations = list(definition.destinations)
            if len(destinations) != 1:
                raise VBAError(1004, "Expected a named range with one destination")
            name, address = destinations[0]
            return self.Parent.Worksheets.Item(name).Range(address)
        return Range(self, top or 1, left or 1, bottom or MAX_ROW, right or MAX_COLUMN)

    def _selection(self):
        return self.Range(self._selected)

    def Activate(self):
        self.Parent._check()
        if self.Visible != -1:
            raise VBAError(1004, "Cannot activate a hidden worksheet")
        self.Parent._raw.active = self.Parent._raw.index(self._raw)
        self.Parent.Activate()

    Select = Activate

    def Delete(self):
        book = self.Parent
        book._check()
        if self.Visible == -1 and sum(sheet.Visible == -1 for sheet in book.Worksheets) <= 1:
            raise VBAError(1004, "A workbook requires at least one visible worksheet")
        active = book._raw.active
        book._raw.remove(self._raw)
        book._saved = False
        if active is self._raw:
            active = next(raw for raw in book._raw.worksheets if raw.sheet_state == "visible")
        book._raw.active = active
        if self.Application.ActiveWorkbook is book:
            book.Activate()


class Range(Indexed):
    def __init__(self, sheet, top, left, bottom, right):
        if not (1 <= top <= bottom <= MAX_ROW and 1 <= left <= right <= MAX_COLUMN):
            raise VBAError(1004, "Range coordinates are outside the worksheet")
        self.Parent, self.Row, self.Column = sheet, top, left
        self._bottom, self._right = bottom, right
        self.Application = sheet.Application

    @property
    def Cells(self):
        return self

    @property
    def Count(self):
        size = self.CountLarge
        if size > 2**31 - 1:
            raise VBAError(6)
        return size

    @property
    def CountLarge(self):
        return (self._bottom - self.Row + 1) * self._width

    @property
    def _width(self):
        return self._right - self.Column + 1

    @property
    def Rows(self):
        return RangeAxis(self, True)

    @property
    def Columns(self):
        return RangeAxis(self, False)

    @property
    def EntireRow(self):
        return Range(self.Parent, self.Row, 1, self._bottom, MAX_COLUMN)

    @property
    def EntireColumn(self):
        return Range(self.Parent, 1, self.Column, MAX_ROW, self._right)

    def Item(self, RowIndex, ColumnIndex=MISSING):
        row = integer(RowIndex)
        if ColumnIndex is MISSING:
            pos, offset = divmod(row - 1, self._width)
            return Range(self.Parent, self.Row + pos, self.Column + offset,
                         self.Row + pos, self.Column + offset)
        column = column_index_from_string(ColumnIndex) if isinstance(ColumnIndex, str) and ColumnIndex.isalpha() else integer(ColumnIndex)
        return self.Offset(row - 1, column - 1).Resize(1, 1) if self.CountLarge == 1 else Range(
            self.Parent, self.Row + row - 1, self.Column + column - 1,
            self.Row + row - 1, self.Column + column - 1)

    def __iter__(self):
        for row, column in self._positions():
            yield Range(self.Parent, row, column, row, column)

    def _positions(self):
        self.Parent.Parent._check()
        for row in range(self.Row, self._bottom + 1):
            for column in range(self.Column, self._right + 1):
                yield row, column

    def _existing_positions(self):
        for row, col in self.Parent._raw._cells:
            if self.Row <= row <= self._bottom and self.Column <= col <= self._right:
                yield row, col

    def _check_materialization(self):
        if self.CountLarge > MAX_MATERIALIZED_CELLS:
            raise VBAError(1004, "Range operation exceeds the supported cell count")

    def _cell(self, row, column):
        self.Parent.Parent._check()
        return self.Parent._raw.cell(row, column)

    def _read(self, row, col, formula=False):
        self.Parent.Parent._check()
        cell = self.Parent._raw._cells.get((row, col))
        if cell is not None and cell.data_type == "f" and not formula:
            return self.Parent.Parent._cached_value(self.Parent._raw, row, col)
        return _sheet_value(cell)

    def _values(self, formula):
        if self.CountLarge == 1:
            return self._read(self.Row, self.Column, formula)
        self._check_materialization()
        array = VBAArray(((1, self._bottom - self.Row + 1), (1, self._width)))
        for row, column in self._positions():
            value = self._read(row, column, formula)
            if value is not EMPTY:
                array.set((row - self.Row + 1, column - self.Column + 1), value)
        return array

    def _write(self, value):
        if isinstance(value, Range):
            value = value.Value
        array = value if isinstance(value, VBAArray) else None
        if array is not None and (array.bounds is None or len(array.bounds) not in {1, 2}):
            raise VBAError(1004, "Expected a one- or two-dimensional array")
        book = self.Parent.Parent
        book._check()
        if value is EMPTY or value is NULL:
            positions = self._existing_positions()
        else:
            self._check_materialization()
            positions = self._positions()
        for row, col in positions:
            cell = self._cell(row, col)
            item = value
            if array is not None:
                rows = row - self.Row
                columns = col - self.Column
                offsets = (columns,) if len(array.bounds) == 1 else (rows, columns)
                indices = tuple(lower + offset for offset, (lower, _) in zip(offsets, array.bounds))
                if any(index > bound[1] for index, bound in zip(indices, array.bounds)):
                    cell.value = "#N/A"
                    book._saved = False
                    book._changed_formulas.add((self.Parent._raw, row, col))
                    continue
                item = array.get(indices)
            if isinstance(item, ErrorValue):
                if item.code not in EXCEL_ERRORS:
                    raise VBAError(1004, f"Excel cannot store error code {item.code}")
                cell.value = EXCEL_ERRORS[item.code]
                cell.data_type = "e"
            else:
                cell.value = None if item is EMPTY or item is NULL else item
            book._changed_formulas.add((self.Parent._raw, row, col))
            book._saved = False

    @property
    def Value(self):
        return self._values(False)

    @Value.setter
    def Value(self, value):
        self._write(value)

    Value2 = Value

    @property
    def Formula(self):
        return self._values(True)

    @Formula.setter
    def Formula(self, value):
        self._write(value)

    def _vba_default_value(self):
        return self.Value

    def _vba_assign_default(self, value):
        self.Value = value

    @property
    def Text(self):
        if self.CountLarge != 1:
            return NULL
        value = self.Value
        return EXCEL_ERRORS.get(value.code, str(value)) if isinstance(value, ErrorValue) else string(value)

    @property
    def Font(self):
        return RangeFont(self)

    @property
    def Interior(self):
        return RangeInterior(self)

    @property
    def NumberFormat(self):
        self.Parent.Parent._check()
        formats = {self._cell(r, c).number_format for r, c in self._existing_positions()}
        if not formats:
            return "General"
        if len(formats) > 1:
            return NULL
        first = formats.pop()
        if first == "General" or sum(1 for _ in self._existing_positions()) == self.CountLarge:
            return first
        return NULL

    @NumberFormat.setter
    def NumberFormat(self, value):
        self._check_materialization()
        for row, col in self._positions():
            self._cell(row, col).number_format = string(value)
        self.Parent.Parent._saved = False

    @property
    def Address(self):
        return self.address()

    def address(self, RowAbsolute=True, ColumnAbsolute=True, ReferenceStyle=1, External=False,
                RelativeTo=MISSING):
        if integer(ReferenceStyle) != 1 or RelativeTo is not MISSING:
            raise VBAError(1004, "Only A1-style addresses are supported")

        def coordinate(row, col):
            return ("$" if truth(ColumnAbsolute) else "") + get_column_letter(col) + (
                "$" if truth(RowAbsolute) else "") + str(row)

        result = coordinate(self.Row, self.Column)
        if self.CountLarge != 1:
            result += ":" + coordinate(self._bottom, self._right)
        if truth(External):
            sheet = self.Parent.Name.replace("'", "''")
            result = f"'[{self.Parent.Parent.Name}]{sheet}'!{result}"
        return result

    def Offset(self, RowOffset=0, ColumnOffset=0):
        row, column = integer(RowOffset), integer(ColumnOffset)
        return Range(self.Parent, self.Row + row, self.Column + column,
                     self._bottom + row, self._right + column)

    def Resize(self, RowSize=MISSING, ColumnSize=MISSING):
        height = self._bottom - self.Row + 1 if RowSize is MISSING else integer(RowSize)
        width = self._width if ColumnSize is MISSING else integer(ColumnSize)
        return Range(self.Parent, self.Row, self.Column, self.Row + height - 1, self.Column + width - 1)

    def Range(self, Cell1, Cell2=MISSING):
        relative = self.Parent.Range(Cell1, Cell2)
        return relative.Offset(self.Row - 1, self.Column - 1)

    def End(self, Direction):
        steps = {-4162: (-1, 0), -4121: (1, 0), -4159: (0, -1), -4161: (0, 1)}
        if integer(Direction) not in steps:
            raise VBAError(1004)
        dy, dx = steps[integer(Direction)]
        row, col = self.Row, self.Column
        raw = self.Parent._raw
        self.Parent.Parent._check()

        def filled(r, c):
            cell = raw._cells.get((r, c))
            return cell is not None and cell.value is not None

        def inside(r, c):
            return 1 <= r <= MAX_ROW and 1 <= c <= MAX_COLUMN

        adjacent = row + dy, col + dx
        if inside(*adjacent) and filled(row, col) and filled(*adjacent):
            while inside(row + dy, col + dx) and filled(row + dy, col + dx):
                row, col = row + dy, col + dx
        elif inside(*adjacent):
            candidates = [((r - row) * dy + (c - col) * dx, r, c)
                          for (r, c), cell in raw._cells.items() if cell.value is not None
                          and (not dy and r == row or not dx and c == col)
                          and (r - row) * dy + (c - col) * dx > 0]
            if candidates:
                _, row, col = min(candidates)
            else:
                row = (MAX_ROW if dy > 0 else 1) if dy else row
                col = (MAX_COLUMN if dx > 0 else 1) if dx else col
        return Range(self.Parent, row, col, row, col)

    def Select(self):
        if self.Application.ActiveWorkbook is not self.Parent.Parent or self.Application.ActiveSheet is not self.Parent:
            raise VBAError(1004, "Selection requires an active worksheet")
        self.Parent._selected = self.Address
        self.Application.Selection = self
        self.Application.ActiveCell = self.Item(1, 1)

    def Activate(self):
        self.Select()

    def ClearContents(self):
        self.Value = EMPTY

    def Clear(self):
        self.ClearContents()
        for row, col in self._existing_positions():
            self._cell(row, col)._style = None

    def Copy(self, Destination=MISSING):
        self._check_materialization()
        styles = ("alignment", "border", "fill", "font", "number_format", "protection")
        entries = []
        for row, col in self._positions():
            cell = self.Parent._raw._cells.get((row, col))
            formatting = {name: copy(getattr(cell, name)) for name in styles} if cell is not None and cell.has_style else {}
            entries.append((row - self.Row, col - self.Column, None if cell is None else cell.value, formatting))
        if Destination is MISSING:
            self.Application._clipboard = entries
            self.Application.CutCopyMode = 1
        elif not isinstance(Destination, Range):
            raise VBAError(424)
        else:
            Destination._paste(entries)

    def _paste(self, entries):
        if entries and (self.Row + entries[-1][0] > MAX_ROW or self.Column + entries[-1][1] > MAX_COLUMN):
            raise VBAError(1004, "Paste area falls outside the worksheet")
        for y, x, value, styles in entries:
            target = self._cell(self.Row + y, self.Column + x)
            # Keep formula text, including references, intact during the copy.
            target.value = value
            target._style = None
            for name, style in styles.items():
                setattr(target, name, style)
            book = self.Parent.Parent
            book._changed_formulas.add((self.Parent._raw, target.row, target.column))
            book._saved = False

    def PasteSpecial(self, Paste=-4104, Operation=-4142, SkipBlanks=False, Transpose=False):
        if (integer(Paste), integer(Operation)) != (-4104, -4142) or truth(SkipBlanks) or truth(Transpose):
            raise VBAError(1004, "Unsupported PasteSpecial options")
        if self.Application._clipboard is None or not self.Application.CutCopyMode:
            raise VBAError(1004, "Clipboard is empty")
        self._paste(self.Application._clipboard)


def _excel_color(color):
    if color is None or color.type != "rgb":
        return 0
    rgb = color.rgb[-6:]
    red, green, blue = (int(rgb[index:index + 2], 16) for index in (0, 2, 4))
    return red + 256 * green + 65536 * blue


def _rgb(color):
    value = integer(color)
    if not 0 <= value <= 0xFFFFFF:
        raise VBAError(1004, "Expected a 24-bit Excel color")
    return "FF" + "".join(f"{(value >> shift) & 255:02X}" for shift in (0, 8, 16))


class RangeStyle:
    def __init__(self, cells):
        self._range = cells

    def _get(self, kind, transform):
        region = self._range
        region._check_materialization()
        sample = set()
        for row, col in region._positions():
            cell = region.Parent._raw._cells.get((row, col))
            style = getattr(cell, kind) if cell is not None else None
            sample.add(transform(style))
            if len(sample) > 1:
                return NULL
        return sample.pop()

    def _set(self, kind, transform):
        region = self._range
        region._check_materialization()
        for row, col in region._positions():
            cell = region._cell(row, col)
            setattr(cell, kind, transform(copy(getattr(cell, kind))))
        region.Parent.Parent._saved = False


class RangeFont(RangeStyle):
    def _field(self, name, default=None):
        return self._get("font", lambda font: getattr(font, name) if font else default)

    def _format(self, name, value):
        def apply(font):
            setattr(font, name, value)
            return font
        self._set("font", apply)

    @property
    def Bold(self):
        return self._field("bold", False)

    @Bold.setter
    def Bold(self, value):
        self._format("bold", truth(value))

    @property
    def Italic(self):
        return self._field("italic", False)

    @Italic.setter
    def Italic(self, value):
        self._format("italic", truth(value))

    @property
    def Size(self):
        return self._field("sz", 11)

    @Size.setter
    def Size(self, value):
        from .values import numeric
        size = float(numeric(value))
        if not 0 < size <= 409:
            raise VBAError(1004)
        self._format("sz", size)

    @property
    def Name(self):
        return self._field("name", "Calibri")

    @Name.setter
    def Name(self, value):
        self._format("name", string(value))

    @property
    def Color(self):
        return self._get("font", lambda font: _excel_color(font.color) if font else 0)

    @Color.setter
    def Color(self, value):
        self._format("color", _rgb(value))


class RangeInterior(RangeStyle):
    @property
    def Color(self):
        return self._get("fill", lambda fill: _excel_color(fill.fgColor) if fill and fill.patternType else 0)

    @Color.setter
    def Color(self, value):
        color = _rgb(value)
        self._set("fill", lambda _: PatternFill(patternType="solid", fgColor=color))


class RangeAxis(Indexed):
    def __init__(self, parent, rows):
        self.parent, self.rows = parent, rows

    @property
    def Count(self):
        parent = self.parent
        return parent._bottom - parent.Row + 1 if self.rows else parent._width

    @property
    def EntireRow(self):
        return self.parent.EntireRow

    @property
    def EntireColumn(self):
        return self.parent.EntireColumn

    def Item(self, Index):
        index = column_index_from_string(Index) if not self.rows and isinstance(Index, str) and Index.isalpha() else integer(Index)
        offset = index - 1
        p = self.parent
        if self.rows:
            return Range(p.Parent, p.Row + offset, p.Column, p.Row + offset, p._right)
        return Range(p.Parent, p.Row, p.Column + offset, p._bottom, p.Column + offset)

    def __iter__(self):
        return (self.Item(i) for i in range(1, self.Count + 1))


class WorksheetFunction:
    def _values(self, args):
        from .values import numeric
        for arg in args:
            if isinstance(arg, Range):
                for row, col in arg._positions():
                    value = arg._read(row, col)
                    if isinstance(value, (int, float)) and not isinstance(value, bool):
                        yield value
            elif isinstance(arg, VBAArray):
                yield from (value for value in arg if isinstance(value, (int, float)) and not isinstance(value, bool))
            else:
                yield numeric(arg)

    def Sum(self, *args):
        return sum(self._values(args))

    def Min(self, *args):
        return min(self._values(args), default=0)

    def Max(self, *args):
        return max(self._values(args), default=0)

    def Count(self, *args):
        return sum(1 for _ in self._values(args))

    def Average(self, *args):
        values = list(self._values(args))
        if not values:
            raise VBAError(1004)
        return sum(values) / len(values)
