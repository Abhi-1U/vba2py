"""Additional VBA string, calendar and financial functions."""

import math
import re
from datetime import date
from locale import windows_locale

from babel.core import Locale, UnknownLocaleError
from babel.dates import get_day_names, get_month_names

from .arrays import VBAArray
from .errors import VBAError
from .locale import current_locale
from .values import NULL, Integer, Long, coerce, default_value, integer, string, truth


LIBRARY_CONSTANTS = {
    "vbusesystem": 0,
    "vbusesystemdayofweek": 0,
    "vbfirstjan1": 1,
    "vbfirstfourdays": 2,
    "vbfirstfullweek": 3,
    "vbuppercase": 1,
    "vblowercase": 2,
    "vbpropercase": 3,
    "vbwide": 4,
    "vbnarrow": 8,
    "vbkatakana": 16,
    "vbhiragana": 32,
    "vbunicode": 64,
    "vbfromunicode": 128,
}

_TURKIC_LOWER = str.maketrans({"I": "ı", "İ": "i"})
_TURKIC_UPPER = str.maketrans({"i": "İ", "ı": "I"})
_TURKIC_FOLD_UNITS = re.compile(r"I\u0307|.", re.S)
_WORD_BREAK = re.compile(r"([\x00\t\n\v\f\r \u3000]+)")


def _case(text, mode, selected_locale):
    turkic = selected_locale.language in {"az", "tr"}
    if mode == "upper":
        return (text.translate(_TURKIC_UPPER) if turkic else text).upper()
    if turkic:
        text = text.replace("I\u0307", "İ").translate(_TURKIC_LOWER)
    return text.lower() if mode == "lower" else text.casefold()


def _text_compare(compare, option_compare):
    mode = integer(compare)
    if mode not in {-1, 0, 1}:
        raise VBAError(5)
    return mode == 1 or (mode == -1 and option_compare == "text")


def _text_positions(source, needle):
    """Find case-insensitive matches without splitting a case-expanded character."""
    locale = current_locale()
    if locale.language in {"az", "tr"}:
        units = ((match.end(), match[0]) for match in _TURKIC_FOLD_UNITS.finditer(source))
    else:
        units = enumerate(source, start=1)
    chunks = []
    offsets = {0: 0}
    offset = 0
    for index, unit in units:
        part = _case(unit, "fold", locale)
        chunks.append(part)
        offset += len(part)
        offsets[offset] = index
    folded = "".join(chunks)
    key = _case(needle, "fold", locale)
    position = folded.find(key)
    while position != -1:
        if position in offsets and position + len(key) in offsets:
            yield offsets[position]
        position = folded.find(key, position + 1)


def _utf16_length(text):
    return len(text.encode("utf-16-le", errors="surrogatepass")) // 2


def instr_rev(StringCheck, StringMatch, Start=-1, Compare=0, *, option_compare="binary"):
    source, needle = map(default_value, (StringCheck, StringMatch))
    if source is NULL or needle is NULL:
        return NULL
    start = integer(Start)
    if start != -1 and start < 1:
        raise VBAError(5)
    text_compare = _text_compare(Compare, option_compare)
    source, needle = string(source), string(needle)
    encoded = source.encode("utf-16-le", errors="surrogatepass")
    length = len(encoded) // 2
    if start == -1:
        start = length
    if not source or start > length:
        return Long(0)
    if not needle:
        return Long(start)
    prefix = encoded[:start * 2].decode("utf-16-le", errors="surrogatepass")
    if text_compare:
        position = max(_text_positions(prefix, needle), default=-1)
    else:
        position = prefix.rfind(needle)
    return Long(_utf16_length(prefix[:position]) + 1 if position >= 0 else 0)


def _require_vector(value):
    if not isinstance(value, VBAArray) or value.bounds is None or len(value.bounds) != 1:
        raise VBAError(13)
    return value


def filter_strings(SourceArray, Match, Include=True, Compare=0, *, option_compare="binary"):
    source = _require_vector(SourceArray)
    needle = string(Match)
    include = coerce(Include, "Boolean")
    text_compare = _text_compare(Compare, option_compare)

    def includes(value):
        if not needle or not text_compare:
            return needle in value
        return next(_text_positions(value, needle), None) is not None

    strings = (string(value) for value in source)
    return VBAArray.from_values(
        (value for value in strings if includes(value) == include), type_name="String",
    )


def _casing_locale(lcid):
    if lcid in {0, 1024, 2048}:
        return current_locale()
    try:
        return Locale.parse(windows_locale[lcid])
    except (KeyError, ValueError, UnknownLocaleError):
        raise VBAError(5, f"Unsupported locale ID {lcid}") from None


def str_conv(String, Conversion, LCID=0):
    source = default_value(String)
    if source is NULL:
        return NULL
    conversion = integer(Conversion)
    if conversion not in {1, 2, 3}:
        raise VBAError(5, "StrConv supports upper, lower and proper case conversions")
    locale = _casing_locale(integer(LCID))
    text = string(source)
    if conversion == 1:
        return _case(text, "upper", locale)
    if conversion == 2:
        return _case(text, "lower", locale)

    def proper_word(word):
        for position, character in enumerate(word):
            if character.isalpha():
                return word[:position] + _case(character, "upper", locale) + word[position + 1:]
        return word

    return "".join(proper_word(word) for word in _WORD_BREAK.split(_case(text, "lower", locale)))


def switch(*pairs):
    if not pairs or len(pairs) % 2:
        raise VBAError(5, "Switch requires pairs of conditions and values")
    for condition, value in zip(pairs[::2], pairs[1::2]):
        if truth(condition):
            return value
    return NULL


def month_name(Month, Abbreviate=False):
    month = integer(Month)
    if not 1 <= month <= 12:
        raise VBAError(5)
    width = "abbreviated" if coerce(Abbreviate, "Boolean") else "wide"
    return get_month_names(width=width, context="stand-alone", locale=current_locale())[month]


def first_day_of_week(value):
    requested = integer(value)
    if requested == 0:
        return current_locale().first_week_day
    if not 1 <= requested <= 7:
        raise VBAError(5)
    return (requested - 2) % 7


def weekday_name(Weekday, Abbreviate=False, FirstDayOfWeek=0):
    position = integer(Weekday)
    if not 1 <= position <= 7:
        raise VBAError(5)
    start = first_day_of_week(FirstDayOfWeek)
    width = "abbreviated" if coerce(Abbreviate, "Boolean") else "wide"
    names = get_day_names(width=width, context="stand-alone", locale=current_locale())
    return names[(start + position - 1) % 7]


def weekday(Date, FirstDayOfWeek=1):
    value = default_value(Date)
    if value is NULL:
        return NULL
    return Integer((coerce(value, "Date").weekday() - first_day_of_week(FirstDayOfWeek)) % 7 + 1)


def _minimum_days_in_week(value):
    rule = integer(value)
    if rule == 0:
        return current_locale().min_week_days
    if rule in {1, 2, 3}:
        return {1: 1, 2: 4, 3: 7}[rule]
    raise VBAError(5)


def _first_week_start(year, first_day, minimum_days):
    january = date(year, 1, 1)
    preceding_days = (january.weekday() - first_day) % 7
    offset = -preceding_days if 7 - preceding_days >= minimum_days else 7 - preceding_days
    return january.toordinal() + offset


def _week_number(value, first_day, minimum_days):
    ordinal = value.toordinal()
    start = _first_week_start(value.year, first_day, minimum_days)
    if ordinal < start:
        start = _first_week_start(value.year - 1, first_day, minimum_days)
    elif value.year < 9999 and ordinal >= _first_week_start(value.year + 1, first_day, minimum_days):
        return 1
    return (ordinal - start) // 7 + 1


def date_part(Interval, Date, FirstDayOfWeek=1, FirstWeekOfYear=1):
    interval = string(Interval).casefold()
    source = default_value(Date)
    if source is NULL:
        return NULL
    value = coerce(source, "Date")
    first_day = first_day_of_week(FirstDayOfWeek)
    minimum_days = _minimum_days_in_week(FirstWeekOfYear)
    parts = {
        "yyyy": value.year, "q": (value.month - 1) // 3 + 1, "m": value.month,
        "y": value.timetuple().tm_yday, "d": value.day, "h": value.hour,
        "n": value.minute, "s": value.second,
    }
    if interval in parts:
        return Integer(parts[interval])
    if interval == "w":
        return Integer((value.weekday() - first_day) % 7 + 1)
    if interval == "ww":
        return Integer(_week_number(value, first_day, minimum_days))
    raise VBAError(5)


def _annuity(Rate, NPer, Pmt, balance, Type, *, future):
    rate, periods, payment, remaining, due = (
        coerce(value, "Double") for value in (Rate, NPer, Pmt, balance, Type)
    )
    if rate <= -1 or due not in {0, 1}:
        raise VBAError(5)
    if rate == 0:
        return coerce(-(remaining + payment * periods), "Double")
    exponent = periods * math.log1p(rate)
    if not future:
        exponent = -exponent
    try:
        factor = math.expm1(exponent) / rate
        if not future:
            factor = -factor
        result = -(remaining * math.exp(exponent) + payment * (1 + rate * due) * factor)
    except OverflowError:
        raise VBAError(6) from None
    return coerce(result, "Double")


def future_value(Rate, NPer, Pmt, PV=0, Type=0):
    return _annuity(Rate, NPer, Pmt, PV, Type, future=True)


def present_value(Rate, NPer, Pmt, FV=0, Type=0):
    return _annuity(Rate, NPer, Pmt, FV, Type, future=False)


def net_present_value(Rate, Values):
    rate = coerce(Rate, "Double")
    if rate <= -1:
        raise VBAError(5)
    flows = [coerce(value, "Double") for value in _require_vector(Values)]
    if not flows or min(flows) >= 0 or max(flows) <= 0:
        raise VBAError(5, "NPV requires positive and negative cash flows")
    total = 0.0
    for value in reversed(flows):
        total = (total + value) / (1 + rate)
    return coerce(total, "Double")


LIBRARY_FUNCTIONS = {
    "instrrev": instr_rev,
    "filter": filter_strings,
    "strconv": str_conv,
    "switch": switch,
    "monthname": month_name,
    "weekdayname": weekday_name,
    "weekday": weekday,
    "datepart": date_part,
    "fv": future_value,
    "pv": present_value,
    "npv": net_present_value,
}
