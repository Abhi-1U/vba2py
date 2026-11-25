"""Selected VBA library functions and constants."""

import calendar
import math
import random
import sys
from datetime import datetime, timedelta
from functools import partial
from decimal import Decimal

from .arrays import VBAArray
from .errors import VBAError
from .formatting import format_currency, format_datetime, format_number, format_percent, format_value
from .library import LIBRARY_CONSTANTS, LIBRARY_FUNCTIONS, first_day_of_week
from .values import (
    EMPTY, EPOCH, MISSING, NULL, ErrorValue, coerce, default_value, desc,
    integer, invariant_number, is_object, numeric, string, truth, val,
)

CONSTANTS = {
    **LIBRARY_CONSTANTS,
    "vbgeneraldate": 0, "vblongdate": 1, "vbshortdate": 2, "vblongtime": 3, "vbshorttime": 4,
    "vbusedefault": -2, "vbtrue": -1, "vbfalse": 0,
    "vbcr": "\r", "vblf": "\n", "vbcrlf": "\r\n", "vbnewline": "\r\n", "vbtab": "\t",
    "vbnullstring": "", "vbnullchar": "\0", "vbback": "\b", "vbformfeed": "\f",
    "vbbinarycompare": 0, "vbtextcompare": 1, "vbusecompareoption": -1,
    "vbsunday": 1, "vbmonday": 2, "vbtuesday": 3, "vbwednesday": 4,
    "vbthursday": 5, "vbfriday": 6, "vbsaturday": 7,
    "vbblack": 0, "vbred": 255, "vbgreen": 65280, "vbyellow": 65535,
    "vbblue": 16711680, "vbmagenta": 16711935, "vbcyan": 16776960, "vbwhite": 16777215,
    "vbok": 1, "vbcancel": 2, "vbokonly": 0, "vbobjecterror": -2147221504,
    "vbempty": 0, "vbnull": 1, "vbinteger": 2, "vblong": 3, "vbsingle": 4,
    "vbdouble": 5, "vbcurrency": 6, "vbdate": 7, "vbstring": 8, "vbobject": 9,
    "vberror": 10, "vbboolean": 11, "vbvariant": 12, "vbdecimal": 14,
    "vbbyte": 17, "vblonglong": 20, "vbarray": 8192,
    "vbreadonly": 1, "vbhidden": 2, "vbsystem": 4, "vbdirectory": 16, "vbarchive": 32,
    "xlup": -4162, "xldown": -4121, "xltoleft": -4159, "xltoright": -4161,
    "xlcalculationautomatic": -4105, "xlcalculationmanual": -4135,
    "xlopenxmlworkbook": 51, "xlopenxmlworkbookmacroenabled": 52,
    "xlopenxmltemplatemacroenabled": 53, "xlopenxmltemplate": 54,
    "xlpasteall": -4104, "xlpastespecialoperationnone": -4142, "xla1": 1, "xlr1c1": -4150,
    "xlerrnull": 2000, "xlerrdiv0": 2007, "xlerrvalue": 2015, "xlerrref": 2023,
    "xlerrname": 2029, "xlerrnum": 2036, "xlerrna": 2042, "xlerrgettingdata": 2043,
    "xlworksheet": -4167, "xlsheetvisible": -1, "xlsheethidden": 0, "xlsheetveryhidden": 2,
}
EXCEL_GLOBALS = {
    "workbooks", "activeworkbook", "activesheet", "activecell", "worksheets", "sheets",
    "range", "cells", "rows", "columns", "selection", "worksheetfunction",
}


class Collection:
    def __init__(self):
        self._values = []

    @property
    def Count(self):
        return len(self._values)

    def _index(self, index):
        if isinstance(index, str):
            key = index.casefold()
            for position, (name, _) in enumerate(self._values):
                if name == key:
                    return position
        else:
            index = integer(index)
            if 1 <= index <= self.Count:
                return index - 1
        raise VBAError(5)

    def Add(self, Item, Key=MISSING, Before=MISSING, After=MISSING):
        key = None if Key is MISSING else string(Key).casefold()
        if key is not None and any(name == key for name, _ in self._values):
            raise VBAError(457)
        if Before is not MISSING and After is not MISSING:
            raise VBAError(5)
        pos = self.Count
        if Before is not MISSING:
            pos = self._index(Before)
        elif After is not MISSING:
            pos = self._index(After) + 1
        self._values.insert(pos, (key, Item))

    def Item(self, Index):
        return self._values[self._index(Index)][1]

    _vba_item = Item

    def Remove(self, Index):
        del self._values[self._index(Index)]

    def __iter__(self):
        return (value for _, value in self._values)


class Dictionary:
    def __init__(self):
        self._items = {}
        self._mode = 0

    @property
    def Count(self):
        return len(self._items)

    @property
    def CompareMode(self):
        return self._mode

    @CompareMode.setter
    def CompareMode(self, value):
        if self._items or integer(value) not in {0, 1}:
            raise VBAError(5)
        self._mode = integer(value)

    def _key(self, key):
        key = default_value(key)
        if isinstance(key, str):
            return key.casefold() if self._mode else key
        try:
            hash(key)
        except TypeError:
            raise VBAError(13) from None
        return key

    def Add(self, Key, Item):
        key = self._key(Key)
        if key in self._items:
            raise VBAError(457)
        self._items[key] = (Key, Item)

    def Exists(self, Key):
        return self._key(Key) in self._items

    def Item(self, Key):
        return self._items.setdefault(self._key(Key), (Key, EMPTY))[1]

    _vba_item = Item

    def _vba_set_item(self, indices, value):
        if len(indices) != 1:
            raise VBAError(450)
        key = self._key(indices[0])
        self._items[key] = (self._items.get(key, indices)[0], value)

    def Keys(self):
        return VBAArray.from_values(key for key, _ in self._items.values())

    def Items(self):
        return VBAArray.from_values(value for _, value in self._items.values())

    def Remove(self, Key):
        try:
            del self._items[self._key(Key)]
        except KeyError:
            raise VBAError(32811, "Element not found") from None

    def RemoveAll(self):
        self._items.clear()

    def __iter__(self):
        return (key for key, _ in self._items.values())


class DebugObject:
    def __init__(self, output):
        self.output = output

    def Print(self, *args):
        print(*(string(value) if value is not NULL else "Null" for value in args), file=self.output)

    def print_fields(self, fields):
        for value, separator in fields:
            self.output.write("Null" if value is NULL else string(value))
            if separator == ",":
                self.output.write(" " if len(fields) > 1 else "\t")
        if not fields or not fields[-1][1]:
            self.output.write("\n")

    def Assert(self, Expression):
        if not truth(Expression):
            raise VBAError(5, "Debug.Assert failed")


def create_object(Class, ServerName=MISSING):
    if ServerName is MISSING and string(Class).casefold() == "scripting.dictionary":
        return Dictionary()
    raise VBAError(429, f"Unsupported automation class {string(Class)!r}")


def _string_arg(function):
    def wrapper(String):
        return NULL if String is NULL else function(string(String))
    return wrapper


def left(String, Length):
    count = integer(Length)
    if count < 0:
        raise VBAError(5)
    return NULL if String is NULL else string(String)[:count]


def right(String, Length):
    count = integer(Length)
    if count < 0:
        raise VBAError(5)
    return NULL if String is NULL else string(String)[-count:] if count else ""


def mid(String, Start, Length=MISSING):
    offset = integer(Start)
    count = None if Length is MISSING else integer(Length)
    if offset < 1 or count is not None and count < 0:
        raise VBAError(5)
    return NULL if String is NULL else string(String)[offset - 1:None if count is None else offset - 1 + count]


def _compare_mode(value, default):
    mode = integer(value)
    if mode == -1:
        return default
    if mode not in {0, 1}:
        raise VBAError(5)
    return "text" if mode else "binary"


def _regex_flags(mode):
    import re
    return re.IGNORECASE if mode == "text" else 0


def instr(mode, *args):
    import re
    if not 2 <= len(args) <= 4:
        raise VBAError(450)
    if len(args) == 2:
        start, source, match, compare = 1, *args, -1
    elif len(args) == 3:
        start, source, match = args
        compare = -1
    else:
        start, source, match, compare = args
    if NULL in (start, source, match, compare):
        return NULL
    index = integer(start)
    text, needle = string(source), string(match)
    if index < 1:
        raise VBAError(5)
    if index > len(text):
        return 0
    if not needle:
        return index
    result = re.search(re.escape(needle), text[index - 1:], _regex_flags(_compare_mode(compare, mode)))
    return result.start() + index if result else 0


def replace(String, Find, Replace, Start=1, Count=-1, Compare=0, *, option_compare="binary"):
    import re
    start, count = integer(Start), integer(Count)
    if start < 1 or count < -1:
        raise VBAError(5)
    text, needle, replacement = string(String)[start - 1:], string(Find), string(Replace)
    if not needle or count == 0:
        return text
    return re.sub(re.escape(needle), lambda _: replacement, text, count=max(count, 0),
                  flags=_regex_flags(_compare_mode(Compare, option_compare)))


def split(Expression, Delimiter=" ", Limit=-1, Compare=0, *, option_compare="binary"):
    import re
    text, delimiter = string(Expression), string(Delimiter)
    limit = integer(Limit)
    if limit < -1:
        raise VBAError(5)
    if not text or limit == 0:
        return VBAArray.from_values([])
    if not delimiter or limit == 1:
        return VBAArray.from_values([text])
    return VBAArray.from_values(re.split(re.escape(delimiter), text, maxsplit=max(0, limit - 1),
                                         flags=_regex_flags(_compare_mode(Compare, option_compare))))


def join(SourceArray, Delimiter=" "):
    if not isinstance(SourceArray, VBAArray) or SourceArray.bounds is None or len(SourceArray.bounds) != 1:
        raise VBAError(13)
    return string(Delimiter).join(string(value) for value in SourceArray)


def _month(year, month):
    adjustment, month0 = divmod(month - 1, 12)
    return year + adjustment, month0 + 1


def date_value(value):
    return coerce(value, "Date")


def date_serial_function(Year, Month, Day):
    year, month, day = map(integer, (Year, Month, Day))
    if 0 <= year < 100:
        year += 2000 if year < 30 else 1900
    year, month = _month(year, month)
    try:
        return datetime(year, month, 1) + timedelta(days=day - 1)
    except (ValueError, OverflowError):
        raise VBAError(5) from None


def date_add(Interval, Number, Date):
    interval, count, current = string(Interval).casefold(), integer(Number), date_value(Date)
    if interval in {"yyyy", "q", "m"}:
        year, month = _month(current.year, current.month + count * {"yyyy": 12, "q": 3, "m": 1}[interval])
        try:
            return current.replace(year=year, month=month, day=min(current.day, calendar.monthrange(year, month)[1]))
        except (ValueError, OverflowError):
            raise VBAError(5) from None
    units = {"y": 86400, "d": 86400, "w": 86400, "ww": 604800, "h": 3600, "n": 60, "s": 1}
    if interval not in units:
        raise VBAError(5)
    try:
        return current + timedelta(seconds=count * units[interval])
    except OverflowError:
        raise VBAError(5) from None


def date_diff(Interval, Date1, Date2, FirstDayOfWeek=1, FirstWeekOfYear=1):
    interval, first, last = string(Interval).lower(), date_value(Date1), date_value(Date2)
    if interval == "yyyy":
        return last.year - first.year
    if interval == "q":
        return (last.year - first.year) * 4 + (last.month - 1) // 3 - (first.month - 1) // 3
    if interval == "m":
        return (last.year - first.year) * 12 + last.month - first.month
    days = (last.date() - first.date()).days
    if interval in {"y", "d"}:
        return days
    if interval == "w":
        return abs(days) // 7 * (-1 if days < 0 else 1)
    if interval == "ww":
        boundary = first_day_of_week(FirstDayOfWeek)
        return (days + (first.weekday() - boundary) % 7) // 7
    if interval == "h":
        return days * 24 + last.hour - first.hour
    if interval == "n":
        return date_diff("h", first, last) * 60 + last.minute - first.minute
    if interval == "s":
        return date_diff("n", first, last) * 60 + last.second - first.second
    raise VBAError(5)


def _int_fn(function):
    def wrapper(Number):
        if Number is NULL:
            return NULL
        value = numeric(Number)
        kind = desc(value)
        return coerce(function(value), kind)
    return wrapper


def _math_fn(function):
    def wrapper(Number):
        if Number is NULL:
            return NULL
        try:
            return function(numeric(Number))
        except ValueError:
            raise VBAError(5) from None
        except OverflowError:
            raise VBAError(6) from None
    return wrapper


def _length(value):
    return NULL if value is NULL else len(string(value))


def _char(value, wide=False):
    code = integer(value)
    maximum = 65535 if wide else 255
    if not (-32768 if wide else 0) <= code <= maximum:
        raise VBAError(5)
    return chr(code & maximum) if wide else bytes([code]).decode("cp1252", errors="replace")


def _asc(value, wide=False):
    text = string(value)
    if not text:
        raise VBAError(5)
    if wide:
        code = int.from_bytes(text[0].encode("utf-16-le", "surrogatepass")[:2], "little")
        return code - 65536 if code >= 32768 else code
    return text[0].encode("cp1252", errors="replace")[0]


def _bound(value, dimension, upper):
    if not isinstance(value, VBAArray):
        raise VBAError(13)
    return value.bound(dimension, upper)


def _var_type(value):
    names = {"Empty": 0, "Null": 1, "Integer": 2, "Long": 3, "Single": 4, "Double": 5,
             "Currency": 6, "Date": 7, "String": 8, "Error": 10, "Boolean": 11,
             "Variant": 12, "Decimal": 14, "Byte": 17, "LongLong": 20, "Object": 9}
    if isinstance(value, VBAArray):
        return 8192 + names.get(value.type_name.title(), 9)
    return names.get(desc(value), 9)


def _is_numeric(value):
    if value is NULL or isinstance(value, datetime):
        return False
    try:
        numeric(value)
        return True
    except VBAError:
        return False


def _is_date(value):
    if not isinstance(value, (str, datetime)):
        return False
    try:
        date_value(value)
        return True
    except VBAError:
        return False


def _repeat(count, value):
    size = integer(count)
    if size < 0:
        raise VBAError(5)
    if isinstance(value, str):
        if not value:
            raise VBAError(5)
        character = value[0]
    else:
        character = _char(integer(value) & 255)
    return character * size


def _str(Number):
    value = numeric(Number)
    return (" " if value >= 0 else "") + invariant_number(value)


def _rgb(Red, Green, Blue):
    channels = (min(integer(value), 255) for value in (Red, Green, Blue))
    red, green, blue = channels
    if min(red, green, blue) < 0:
        raise VBAError(5)
    return red + 256 * green + 65536 * blue


def _pmt(Rate, NPer, PV, FV=0, Due=0):
    rate, nper, pv, fv, due = map(numeric, (Rate, NPer, PV, FV, Due))
    rate, nper, pv, fv, due = map(float, (rate, nper, pv, fv, due))
    if due not in {0, 1} or nper == 0 or rate <= -1:
        raise VBAError(5)
    if rate == 0:
        return -(pv + fv) / nper
    growth = math.expm1(nper * math.log1p(rate))
    return -(pv * (1 + growth) + fv) * rate / ((1 + rate * due) * growth)


CONVERSIONS = {
    "cbool": "boolean", "cbyte": "byte", "ccur": "currency", "cdate": "date", "cvdate": "date",
    "cdbl": "double", "cdec": "decimal", "cint": "integer", "clng": "long",
    "clnglng": "longlong", "clngptr": "longptr", "csng": "single", "cstr": "string",
}

FUNCTIONS = {
    **LIBRARY_FUNCTIONS,
    "format": format_value, "formatnumber": format_number, "formatcurrency": format_currency,
    "formatpercent": format_percent, "formatdatetime": format_datetime,
    "len": _length, "left": left, "right": right, "mid": mid,
    "trim": _string_arg(lambda text: text.strip(" ")), "ltrim": _string_arg(lambda text: text.lstrip(" ")),
    "rtrim": _string_arg(lambda text: text.rstrip(" ")), "ucase": _string_arg(str.upper),
    "lcase": _string_arg(str.lower), "strreverse": _string_arg(lambda text: text[::-1]),
    "replace": replace, "split": split, "join": join, "string": _repeat,
    "space": lambda Number: _repeat(Number, " "),
    "chr": _char, "chrw": lambda CharCode: _char(CharCode, True),
    "asc": _asc, "ascw": lambda String: _asc(String, True),
    "str": _str,
    "hex": lambda Number: format(integer(Number) & 0xFFFFFFFF if integer(Number) < 0 else integer(Number), "X"),
    "oct": lambda Number: format(integer(Number) & 0xFFFFFFFF if integer(Number) < 0 else integer(Number), "o"),
    "val": val, "cvar": default_value, "abs": _int_fn(abs), "int": _int_fn(math.floor),
    "fix": _int_fn(math.trunc), "sgn": _int_fn(lambda value: (value > 0) - (value < 0)),
    "round": lambda Expression, NumDigitsAfterDecimal=0: round(numeric(Expression), integer(NumDigitsAfterDecimal)),
    "sqr": _math_fn(math.sqrt), "exp": _math_fn(math.exp), "log": _math_fn(math.log),
    "sin": _math_fn(math.sin), "cos": _math_fn(math.cos), "tan": _math_fn(math.tan), "atn": _math_fn(math.atan),
    "lbound": lambda ArrayName, Dimension=MISSING: _bound(ArrayName, Dimension, False),
    "ubound": lambda ArrayName, Dimension=MISSING: _bound(ArrayName, Dimension, True),
    "isarray": lambda VarName: isinstance(VarName, VBAArray), "isempty": lambda Expression: Expression is EMPTY,
    "isnull": lambda Expression: Expression is NULL, "isobject": is_object,
    "ismissing": lambda ArgName: ArgName is MISSING,
    "iserror": lambda Expression: Expression is MISSING or isinstance(Expression, ErrorValue),
    "cverr": lambda ErrorNumber: ErrorValue(integer(ErrorNumber)) if not isinstance(ErrorNumber, ErrorValue) else ErrorNumber,
    "isnumeric": _is_numeric, "isdate": _is_date, "typename": desc, "vartype": _var_type,
    "iif": lambda Expression, TruePart, FalsePart: TruePart if truth(Expression) else FalsePart,
    "choose": lambda Index, *Choice: Choice[integer(Index)-1] if 1 <= integer(Index) <= len(Choice) else NULL,
    "date": lambda: datetime.combine(datetime.now().date(), EPOCH.time()),
    "time": lambda: datetime.combine(EPOCH.date(), datetime.now().time()),
    "now": datetime.now, "timer": lambda: (datetime.now() - datetime.combine(datetime.now().date(), EPOCH.time())).total_seconds(),
    "datevalue": lambda Date: datetime.combine(date_value(Date).date(), EPOCH.time()),
    "timevalue": lambda Time: datetime.combine(EPOCH.date(), date_value(Time).time()),
    "dateserial": date_serial_function,
    "timeserial": lambda Hour, Minute, Second: EPOCH + timedelta(seconds=(integer(Hour)*3600 + integer(Minute)*60 + integer(Second)) % 86400),
    "dateadd": date_add, "datediff": date_diff,
    "year": lambda Date: date_value(Date).year, "month": lambda Date: date_value(Date).month,
    "day": lambda Date: date_value(Date).day, "hour": lambda Time: date_value(Time).hour,
    "minute": lambda Time: date_value(Time).minute, "second": lambda Time: date_value(Time).second,
    "rgb": _rgb, "pmt": _pmt, "createobject": create_object,
}


class Builtins:
    def __init__(self, program, output=None):
        self.program = program
        self.debug = DebugObject(output if output is not None else sys.stdout)
        self._random = random.Random()
        self._last = 0.0

    def rnd(self, Number=MISSING):
        if Number is not MISSING and numeric(Number) < 0:
            self._random.seed(float(numeric(Number)))
        elif Number is not MISSING and numeric(Number) == 0:
            return self._last
        self._last = self._random.random()
        return self._last

    def randomize(self, Number=MISSING):
        self._random.seed(None if Number is MISSING else float(numeric(Number)))

    def get(self, key, frame, qualified=False):
        fs = self.program.app.files
        functions = {
            "freefile": fs.free_file, "eof": fs.eof, "lof": fs.lof, "loc": fs.loc,
            "seek": fs.seek, "input": fs.input_chars, "dir": fs.dir, "chdir": fs.chdir,
            "curdir": lambda Drive=MISSING: str(fs.directory), "mkdir": fs.mkdir, "rmdir": fs.rmdir,
            "kill": fs.kill, "filecopy": fs.copy, "filelen": lambda PathName: fs.path(PathName).stat().st_size,
            "rnd": self.rnd, "randomize": self.randomize,
            "msgbox": lambda Prompt, Buttons=0, Title=MISSING: (self.debug.Print(Prompt), 1)[1],
            "array": lambda *args: VBAArray.from_values(args, 0 if qualified else frame.spec["base"]),
            "instr": lambda *args: instr(frame.spec["compare"], *args),
            "strcomp": lambda String1, String2, Compare=-1: self._strcomp(String1, String2, Compare, frame.spec["compare"]),
        }
        if key in CONSTANTS:
            return CONSTANTS[key]
        if key in CONVERSIONS:
            return lambda Expression: coerce(Expression, CONVERSIONS[key])
        if key in {"filter", "instrrev", "replace", "split"}:
            return partial(FUNCTIONS[key], option_compare=frame.spec["compare"])
        return functions.get(key, FUNCTIONS.get(key, MISSING))

    @staticmethod
    def _strcomp(String1, String2, compare, default):
        if NULL in (String1, String2):
            return NULL
        first, second = string(String1), string(String2)
        if _compare_mode(compare, default) == "text":
            first, second = first.casefold(), second.casefold()
        return (first > second) - (first < second)
