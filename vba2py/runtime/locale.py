"""Locale data and text conversion scoped to a running VBA project."""

import re
from contextlib import contextmanager
from contextvars import ContextVar
from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation
from functools import lru_cache

from babel.core import Locale, UnknownLocaleError
from babel.dates import format_date, format_time, get_date_format, get_month_names, get_period_names, tokenize_pattern
from babel.numbers import get_currency_symbol, get_decimal_symbol, get_group_symbol, get_territory_currencies

_DEFAULT_LOCALE = Locale.parse("en_US")
_ACTIVE_LOCALE = ContextVar("vba2py_locale", default=_DEFAULT_LOCALE)
OLE_EPOCH_DATE = date(1899, 12, 30)
_DATE_SEPARATORS = r"[\s/,\.\-年月日]+"
_BIDI_MARKS = str.maketrans("", "", "\u061c\u200e\u200f")
_EXPONENT = re.compile(r"\s*[eEdD]\s*([+\-−]?)\s*(\d+)$")
_SPACE_GROUPS = re.compile(r"[ \u00a0\u202f]")
_ISO_DATE = re.compile(r"\d{4}-\d\d-\d\d(?:[ T]\d\d:\d\d(?::\d\d(?:\.\d+)?)?)?")


def resolve_locale(locale):
    if isinstance(locale, Locale):
        return locale
    if locale is None:
        return _DEFAULT_LOCALE
    try:
        return Locale.parse(locale.replace("-", "_"))
    except (AttributeError, TypeError, ValueError, UnknownLocaleError):
        raise ValueError(f"Unknown VBA locale {locale!r}. Use a locale such as en_US or de_DE.") from None


def current_locale():
    return _ACTIVE_LOCALE.get()


@contextmanager
def use_locale(locale):
    token = _ACTIVE_LOCALE.set(resolve_locale(locale))
    try:
        yield
    finally:
        _ACTIVE_LOCALE.reset(token)


@lru_cache(maxsize=64)
def _currency_tokens(locale):
    if locale.territory is None:
        return ()
    tokens = set()
    for currency in get_territory_currencies(locale.territory):
        tokens.update((currency, get_currency_symbol(currency, locale)))
    return tuple(sorted((token.casefold() for token in tokens), key=len, reverse=True))


def parse_local_number(text):
    """Parse a localized number into an exact Decimal, or raise ValueError."""
    locale = current_locale()
    text = text.translate(_BIDI_MARKS).strip()
    negative = text.startswith("(") and text.endswith(")")
    signed = negative or text[:1] in {"+", "-", "−"}
    if negative:
        text = text[1:-1].strip()
    elif signed:
        negative = text[0] != "+"
        text = text[1:].lstrip()

    folded = text.casefold()
    for token in _currency_tokens(locale):
        if folded.startswith(token):
            text = text[len(token):].lstrip()
            if not signed and text[:1] in {"+", "-", "−"}:
                negative = text[0] != "+"
                text = text[1:].lstrip()
            break
        if folded.endswith(token):
            text = text[:-len(token)].rstrip()
            break

    exponent = ""
    if (match := _EXPONENT.search(text)) is not None:
        exponent = "E" + match[1].replace("−", "-") + match[2]
        text = text[:match.start()]

    decimal = get_decimal_symbol(locale)
    group = get_group_symbol(locale)
    # Windows settings often use space where CLDR specifies a non-breaking space.
    if group.isspace():
        text = _SPACE_GROUPS.sub(group, text)
    whole, point, fraction = text.partition(decimal)
    if group in whole:
        parts = whole.split(group)
        primary, secondary = locale.decimal_formats[None].grouping
        if (not parts[0].isdecimal() or not 1 <= len(parts[0]) <= secondary
                or not all(part.isdecimal() and len(part) == secondary for part in parts[1:-1])
                or not parts[-1].isdecimal() or len(parts[-1]) != primary):
            raise ValueError("Invalid digit grouping")
        whole = "".join(parts)

    if (not whole and not fraction or whole and not whole.isdecimal()
            or fraction and not fraction.isdecimal()):
        raise ValueError("Invalid localized number")
    canonical = ("-" if negative else "") + (whole or "0") + ("." + fraction if point else "") + exponent
    try:
        return Decimal(canonical)
    except InvalidOperation:
        raise ValueError("Invalid localized number") from None


@lru_cache(maxsize=64)
def _date_order(locale):
    fields = (value[0] for kind, value in tokenize_pattern(get_date_format("short", locale).pattern)
              if kind == "field")
    return tuple(field for field in fields if field in "yMd")


def _name_key(text):
    return " ".join(text.translate(_BIDI_MARKS).casefold().split()).rstrip(".")


@lru_cache(maxsize=64)
def _month_names(locale):
    return {_name_key(label): index for width in ("wide", "abbreviated")
            for context in ("format", "stand-alone")
            for index, label in get_month_names(width, context, locale).items()}


def locale_separator(pattern, fallback):
    tokens = tokenize_pattern(pattern.pattern)
    for index, (kind, value) in enumerate(tokens[1:-1], 1):
        if (kind == "chars" and tokens[index - 1][0] == tokens[index + 1][0] == "field"
                and value.strip() and not any(char.isalpha() for char in value)):
            return value
    return fallback


@lru_cache(maxsize=64)
def _time_pattern(locale, literal):
    periods = {"am": "am", "pm": "pm", "a.m.": "am", "p.m.": "pm"}
    if literal:
        periods.update({"a": "am", "p": "pm"})
    else:
        for width in ("wide", "abbreviated"):
            names = get_period_names(width, context="format", locale=locale)
            for key in ("am", "pm"):
                if key in names:
                    label = " ".join(names[key].translate(_BIDI_MARKS).casefold().split())
                    periods[label] = key
    alternatives = "|".join(re.escape(name) for name in sorted(periods, key=len, reverse=True))
    regional_separator = "." if literal else locale_separator(locale.time_formats["medium"], ":")
    separator = rf"(?:{re.escape(regional_separator)}|:)"
    pattern = re.compile(
        rf"(?<!\S)(?P<h>\d{{1,2}})(?:\s*{separator}\s*(?P<m>\d{{1,2}})"
        rf"(?:\s*{separator}\s*(?P<s>\d{{1,2}}))?)?\s*(?P<p>{alternatives})?$",
        re.I,
    )
    return pattern, periods


def _year(value):
    digits = int(value)
    if len(value) <= 2:
        return digits + (2000 if digits < 30 else 1900)
    return digits


def _make_date(year, month, day):
    if year < 100:
        raise ValueError("VBA dates start at year 100")
    return date(year, month, day)


def _literal_numeric_date(parts):
    left, middle, *right = parts
    year = date.today().year
    if not right:
        candidates = [(year, int(left), int(middle)), (year, int(middle), int(left)),
                      (_year(middle), int(left), 1), (_year(left), int(middle), 1)]
    else:
        (right,) = right
        candidates = [(_year(right), int(left), int(middle)),
                      (_year(left), int(middle), int(right)),
                      (_year(right), int(middle), int(left))]
    for candidate in candidates:
        try:
            return _make_date(*candidate)
        except ValueError:
            continue
    raise ValueError("Invalid VBA date literal")


def _calendar_date(text, locale, literal):
    parts = re.split(_DATE_SEPARATORS, text.strip(" \t年月日"))
    if 2 <= len(parts) <= 3 and all(part.isdecimal() and len(part) <= 4 for part in parts):
        if literal:
            return _literal_numeric_date(parts)
        order = _date_order(locale)
        if len(parts) == 3 and len(parts[0]) == 4:
            order = ("y", "M", "d")
        fields = dict(zip(order if len(parts) == 3 else (field for field in order if field != "y"), parts))
        year = _year(fields["y"]) if "y" in fields else date.today().year
        return _make_date(year, int(fields["M"]), int(fields["d"]))

    folded = _name_key(text)
    for name, month in sorted(_month_names(locale).items(), key=lambda item: -len(item[0])):
        match = re.search(rf"(?<!\w){re.escape(name)}\.?(?!\w)", folded)
        if match is None:
            continue
        remaining = folded[:match.start()] + " " + folded[match.end():]
        parts = re.split(_DATE_SEPARATORS, remaining.strip(" \t.,/-"))
        if not 1 <= len(parts) <= 2 or not all(part.isdecimal() for part in parts):
            break
        if len(parts) == 1:
            if literal and int(parts[0]) > 31:
                return _make_date(_year(parts[0]), month, 1)
            return _make_date(date.today().year, month, int(parts[0]))
        first, second = parts
        if len(first) >= 3 or int(first) > 31:
            first, second = second, first
        return _make_date(_year(second), month, int(first))
    raise ValueError("Invalid date text")


def parse_date_text(text, *, literal=False):
    """Parse a Gregorian date or time. Source date tokens use English names and VBA order."""
    locale = _DEFAULT_LOCALE if literal else current_locale()
    text = " ".join(text.translate(_BIDI_MARKS).split())
    if _ISO_DATE.fullmatch(text):
        result = datetime.fromisoformat(text)
        if result.year < 100:
            raise ValueError("VBA dates start at year 100")
        return result

    pattern, periods = _time_pattern(locale, literal)
    hours = time()
    if (match := pattern.search(text)) is not None and (match["m"] is not None or match["p"] is not None):
        hour = int(match["h"])
        if match["p"] is not None:
            if not literal and not 1 <= hour <= 12:
                raise ValueError("Invalid hour for a 12-hour clock")
            if hour <= 12:
                hour = hour % 12 + (12 if periods[match["p"].casefold()] == "pm" else 0)
        try:
            hours = time(hour, int(match["m"] or 0), int(match["s"] or 0))
        except ValueError:
            # A punctuation separator such as "." may also belong to a short date.
            if not match.start() and match["p"] is None and ":" not in text:
                return datetime.combine(_calendar_date(text, locale, literal), time())
            raise
        text = text[:match.start()].rstrip(" ,")
        if not text:
            return datetime.combine(OLE_EPOCH_DATE, hours)

    return datetime.combine(_calendar_date(text, locale, literal), hours)


def display_date(value):
    locale = current_locale()
    clock = format_time(value, "medium", locale=locale)
    if value.date() == OLE_EPOCH_DATE:
        return clock
    short_date = get_date_format("short", locale).pattern
    # VBA's general date display keeps the full year with the local date order.
    short_date = re.sub(r"(?<!y)yy(?!y)", "yyyy", short_date)
    day = format_date(value, short_date, locale=locale)
    return day if value.time() == time() else f"{day} {clock}"
