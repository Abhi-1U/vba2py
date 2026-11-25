"""VBA numeric, date, and string formatting with regional defaults from CLDR."""

import re
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP, localcontext

from babel import dates, numbers

from .errors import VBAError
from .locale import current_locale, locale_separator
from .values import EPOCH, MISSING, NULL, coerce, default_value, integer, numeric, string


_DATE_NAMES = {
    "general date": "c", "long date": "dddddd", "medium date": "medium date",
    "short date": "ddddd", "long time": "ttttt", "medium time": "medium time",
    "short time": "short time",
}
_DATE_CODES = ("c", "dddddd", "ddddd", "ttttt", "short time")
_DATE_TOKEN = re.compile(
    r"am/pm|ampm|a/p|d{1,6}|t{5}|yyyy|yy|[cqy]|w{1,2}|m{1,4}|(?:hh?|nn?|ss?)",
    re.IGNORECASE,
)
_NUMBER_MASK = re.compile(r"(?:[0#,]+(?:\.[0#]*)?|\.[0#]+)(?:[Ee][+-][0#]+)?")
_BOOLEAN_NAMES = {"yes/no": ("No", "Yes"), "true/false": ("False", "True"), "on/off": ("Off", "On")}
_LITERAL = "\x00"


def _sections(pattern):
    """Return format characters, distinguishing quoted and escaped characters."""
    sections = [[]]
    quoted = False
    index = 0
    while index < len(pattern):
        char = pattern[index]
        index += 1
        if char == '"':
            if quoted and pattern[index:index + 1] == '"':
                sections[-1].append(('"', True))
                index += 1
            else:
                quoted = not quoted
        elif char == "\\" and not quoted:
            if index == len(pattern):
                raise VBAError(5, "Incomplete format escape")
            sections[-1].append((pattern[index], True))
            index += 1
        elif char == ";" and not quoted:
            sections.append([])
        else:
            sections[-1].append((char, quoted))
    if quoted or len(sections) > 4:
        raise VBAError(5, "Invalid Format sections")
    return sections


def _unquoted(tokens):
    return "".join(_LITERAL if literal else char for char, literal in tokens)


def _text(tokens):
    return "".join(char for char, _ in tokens)


def _decimal(value):
    value = numeric(value, exact=True)
    result = value if isinstance(value, Decimal) else Decimal(str(value))
    if not result.is_finite() or result.adjusted() > 308:
        raise VBAError(6)
    return result


def _shift(value, places):
    sign, digits, exponent = value.as_tuple()
    return Decimal((sign, digits, exponent + places))


def _round(value, decimals):
    with localcontext() as context:
        context.prec = max(32, value.adjusted() + decimals + 2)
        return value.quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)


def _group(whole, widths, separator):
    if not widths:
        return whole
    primary, secondary = widths
    groups = []
    while len(whole) > primary:
        groups.append(whole[-primary:])
        whole = whole[:-primary]
        primary = secondary
    groups.append(whole)
    return separator.join(reversed(groups))


def _digits(value, whole_min, fraction_min, fraction_max, grouping, locale):
    whole, _, fraction = format(_round(value, fraction_max), "f").partition(".")
    whole = ("" if whole == "0" else whole).zfill(whole_min)
    fraction = fraction.rstrip("0").ljust(fraction_min, "0")
    whole = _group(whole, grouping, numbers.get_group_symbol(locale))
    return whole + (numbers.get_decimal_symbol(locale) + fraction if fraction else "")


def _tristate(value, default):
    option = integer(value)
    if option == -2:
        return default
    if option not in {0, -1}:
        raise VBAError(5, "Invalid formatting option")
    return option == -1


def _numeric_style(style, locale):
    if style == "currency":
        currencies = numbers.get_territory_currencies(locale.territory or "US")
        currency = currencies[0] if currencies else "XXX"
        return (locale.currency_formats["standard"], numbers.get_currency_precision(currency),
                numbers.get_currency_symbol(currency, locale))
    if style == "percent":
        return locale.percent_formats[None], 2, ""
    return locale.decimal_formats[None], 2, ""


def _format_regional(value, style, digits, leading, parens, grouping):
    value = default_value(value)
    if value is NULL:
        return ""
    locale = current_locale()
    regional, default_digits, currency_symbol = _numeric_style(style, locale)
    precision = integer(digits)
    precision = default_digits if precision == -1 else precision
    if not 0 <= precision <= 99:
        raise VBAError(5, "Invalid number of decimal places")
    widths = regional.grouping if _tristate(grouping, regional.grouping[0] < 1000) else None
    minimum = int(_tristate(leading, regional.int_prec[0] > 0) or not precision)
    amount = _decimal(value)
    negative = amount < 0
    magnitude = _shift(amount.copy_abs(), 2) if style == "percent" else amount.copy_abs()
    result = _digits(magnitude, minimum, precision, precision, widths, locale)
    prefix, suffix = regional.prefix[int(negative)], regional.suffix[int(negative)]
    if negative:
        defaults_to_parens = "(" in prefix and ")" in suffix
        if _tristate(parens, defaults_to_parens):
            prefix, suffix = "(" + regional.prefix[0], regional.suffix[0] + ")"
        elif defaults_to_parens:
            prefix, suffix = "-" + regional.prefix[0], regional.suffix[0]
    else:
        _tristate(parens, False)
    prefix = prefix.replace("-", numbers.get_minus_sign_symbol(locale))
    return (prefix + result + suffix).replace("¤", currency_symbol)


def format_number(Expression, NumDigitsAfterDecimal=-1, IncludeLeadingDigit=-2,
                  UseParensForNegativeNumbers=-2, GroupDigits=-2):
    return _format_regional(Expression, "number", NumDigitsAfterDecimal, IncludeLeadingDigit,
                            UseParensForNegativeNumbers, GroupDigits)


def format_currency(Expression, NumDigitsAfterDecimal=-1, IncludeLeadingDigit=-2,
                    UseParensForNegativeNumbers=-2, GroupDigits=-2):
    return _format_regional(Expression, "currency", NumDigitsAfterDecimal, IncludeLeadingDigit,
                            UseParensForNegativeNumbers, GroupDigits)


def format_percent(Expression, NumDigitsAfterDecimal=-1, IncludeLeadingDigit=-2,
                   UseParensForNegativeNumbers=-2, GroupDigits=-2):
    return _format_regional(Expression, "percent", NumDigitsAfterDecimal, IncludeLeadingDigit,
                            UseParensForNegativeNumbers, GroupDigits)


def _named_date(value, code, locale):
    date_styles = {"ddddd": "short", "dddddd": "full", "medium date": "medium"}
    if code in date_styles:
        return dates.format_date(value, date_styles[code], locale=locale)
    if code == "ttttt":
        return dates.format_time(value.time(), "medium", locale=locale)
    if code in {"medium time", "short time"}:
        sep = locale_separator(locale.time_formats["medium"], ":")
        pattern = f"h'{sep}'mm a" if code == "medium time" else f"HH'{sep}'mm"
        return dates.format_time(value.time(), pattern, locale=locale)
    if value.date() == EPOCH.date():
        return _named_date(value, "ttttt", locale)
    date_text = _named_date(value, "ddddd", locale)
    if value.time() == EPOCH.time():
        return date_text
    return date_text + " " + _named_date(value, "ttttt", locale)


def format_datetime(Date, NamedFormat=0):
    code = integer(NamedFormat)
    if not 0 <= code < len(_DATE_CODES):
        raise VBAError(5, "Unknown date format")
    value = default_value(Date)
    return "" if value is NULL else _named_date(coerce(value, "Date"), _DATE_CODES[code], current_locale())


def _date_pattern(value, tokens, locale, first_day, first_week):
    fields = list(_DATE_TOKEN.finditer(_unquoted(tokens)))
    hours12 = any(item[0].lower() in {"am/pm", "ampm", "a/p"} for item in fields)
    hour = (value.hour % 12 or 12) if hours12 else value.hour
    pm = value.hour >= 12
    separators = {
        "/": locale_separator(locale.date_formats["short"], "/"),
        ":": locale_separator(locale.time_formats["medium"], ":"),
    }

    def literal_text(chunk):
        return "".join(char if literal else separators.get(char, char) for char, literal in chunk)

    numbers_by_field = {"d": value.day, "m": value.month, "h": hour, "n": value.minute,
                        "s": value.second, "q": (value.month - 1) // 3 + 1}
    previous = None
    result = []
    pos = 0
    for match in fields:
        gap = tokens[pos:match.start()]
        result.append(literal_text(gap))
        symbol = match[0].lower()
        length = len(symbol)
        if symbol in _DATE_CODES:
            part = _named_date(value, symbol, locale)
        elif symbol in {"ampm", "am/pm", "a/p"}:
            if symbol == "ampm":
                part = dates.format_time(value.time(), "a", locale=locale)
            else:
                part = match[0].split("/")[int(pm)]
        elif symbol in {"ddd", "dddd", "mmm", "mmmm"}:
            locale_format = {"ddd": "EEE", "dddd": "EEEE", "mmm": "MMM", "mmmm": "MMMM"}[symbol]
            part = dates.format_date(value, locale_format, locale=locale)
        elif symbol in {"w", "ww"}:
            from .library import date_part
            part = str(date_part(symbol, value, first_day, first_week))
        elif symbol in {"y", "yy", "yyyy"}:
            if length == 1:
                year_part = value.timetuple().tm_yday
            else:
                year_part = value.year % 100 if length == 2 else value.year
            part = str(year_part).zfill(length)
        else:
            field = symbol[0]
            if field == "m" and previous in {"h", "hh"} and all(char in ": " for char, _ in gap):
                field = "n"
            part = str(numbers_by_field[field]).zfill(length)
        result.append(part)
        previous = symbol
        pos = match.end()
    result.append(literal_text(tokens[pos:]))
    return "".join(result)


def _string_pattern(value, tokens):
    raw = _unquoted(tokens)
    if ">" in raw:
        value = value.upper()
    if "<" in raw:
        value = value.lower()
    left = "!" in raw
    width = raw.count("@") + raw.count("&")
    if not width:
        return value + "".join(char for char, literal in tokens if literal or char not in "<>!")
    overflow = value[width:] if left else value[:max(0, len(value) - width)]
    value = value[:width] if left else value[-width:]
    padding = [None] * (width - len(value))
    characters = iter([*value, *padding] if left else [*padding, *value])
    result = []
    for char, literal in tokens:
        if not literal and char in "@&":
            character = next(characters)
            result.append(character if character is not None else " " if char == "@" else "")
        elif literal or char not in "<>!":
            result.append(char)
    formatted = "".join(result)
    return formatted + overflow if left else overflow + formatted


def _integer_pattern(value, tokens):
    digits = format(_round(value, 0), "f")
    remaining = list(digits) if digits != "0" else []
    pattern = []
    for char, literal in reversed(tokens):
        if not literal and char in "0#":
            char = remaining.pop() if remaining else "0" if char == "0" else ""
        pattern.append(char)
    # Extra digits belong just before the first placeholder, after any prefix.
    prefix_length = next(index for index, (char, literal) in enumerate(tokens)
                         if not literal and char in "0#")
    result = list(reversed(pattern))
    result[prefix_length:prefix_length] = remaining
    return "".join(result)


def _numeric_pattern(value, tokens, automatic_sign, locale):
    raw = _unquoted(tokens)
    masks = [match for match in _NUMBER_MASK.finditer(raw) if any(ch in "0#" for ch in match[0])]
    if not masks:
        return _text(tokens)
    sign = numbers.get_minus_sign_symbol(locale) if value < 0 and automatic_sign else ""
    if len(masks) > 1:
        if any(set(match[0]) - {"0", "#"} for match in masks):
            raise VBAError(5, "Unsupported numeric format mask")
        return sign + _integer_pattern(_shift(value.copy_abs(), 2 * raw.count("%")), tokens)
    match = masks[0]
    mantissa, *exponent = re.split(r"([eE][+-])", match[0], maxsplit=1)
    whole, _, fraction = mantissa.partition(".")
    scale = len(whole) - len(whole.rstrip(","))
    grouping = locale.decimal_formats[None].grouping if "," in whole.rstrip(",") else None
    magnitude = _shift(value.copy_abs(), 2 * raw.count("%") - 3 * scale)
    exponent_suffix = ""
    minimum = whole.count("0")
    fraction_min = len(fraction.rstrip("#"))
    if exponent:
        fraction_places = len(fraction)
        exponent_value = magnitude.adjusted() - max(0, minimum - 1) if magnitude else 0
        magnitude = _shift(magnitude, -exponent_value)
        if _round(magnitude, fraction_places).adjusted() >= max(1, minimum):
            magnitude = _shift(magnitude, -1)
            exponent_value += 1
        exponent_sign = ("-" if exponent_value < 0 else "+" if exponent[0][1] == "+" else "")
        exponent_suffix = exponent[0][0] + exponent_sign + str(abs(exponent_value)).zfill(len(exponent[1]))
    digit_text = _digits(magnitude, minimum, fraction_min, len(fraction), grouping, locale)
    if mantissa.endswith("."):
        digit_text += numbers.get_decimal_symbol(locale)
    return sign + _text(tokens[:match.start()]) + digit_text + exponent_suffix + _text(tokens[match.end():])


def format_value(Expression, Format=MISSING, FirstDayOfWeek=1, FirstWeekOfYear=1):
    value = default_value(Expression)
    pattern = string(Format)[:257] if Format is not MISSING else ""
    if not pattern:
        return "" if value is NULL else string(value)
    name = pattern.casefold()
    locale = current_locale()
    if name in _DATE_NAMES:
        return "" if value is NULL else _named_date(coerce(value, "Date"), _DATE_NAMES[name], locale)
    if name in _BOOLEAN_NAMES:
        return "" if value is NULL else _BOOLEAN_NAMES[name][numeric(value) != 0]
    if name in {"general number", "currency", "fixed", "standard", "percent", "scientific"}:
        if value is NULL:
            return ""
        if name == "general number":
            return string(numeric(value))
        if name == "currency":
            return format_currency(value, 2)
        if name == "percent":
            return format_percent(value, 2)
        if name == "scientific":
            return _numeric_pattern(_decimal(value), _sections("0.00E+00")[0], True, locale)
        return format_number(value, 2, -1, 0, -1 if name == "standard" else 0)

    sections = _sections(pattern)
    raw = _unquoted(sections[0])
    if any(char in raw for char in "@&<>!"):
        if value is NULL and len(sections) == 1:
            return ""
        text = "" if value is NULL else string(value)
        tokens = (sections[1] or sections[0]) if len(sections) > 1 and not text else sections[0]
        return _string_pattern(text, tokens)
    if value is NULL:
        return _text(sections[3]) if len(sections) == 4 else ""
    if re.search(r"[0#%]", raw) is None and _DATE_TOKEN.search(raw):
        return _date_pattern(coerce(value, "Date"), sections[0], locale, FirstDayOfWeek, FirstWeekOfYear)
    if len(sections) == 1 and re.search(r"[0#%]", raw) is None:
        return _text(sections[0])
    amount = _decimal(value)
    index = 1 if amount < 0 and len(sections) >= 2 else 2 if amount == 0 and len(sections) >= 3 else 0
    tokens = sections[index] or sections[0]
    return _numeric_pattern(amount, tokens, index != 1 or not sections[1], locale)
