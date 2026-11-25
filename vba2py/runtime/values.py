"""VBA values, conversions, and expression operators."""

import math
import operator
import re
import struct
from dataclasses import dataclass
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_EVEN, localcontext
from enum import Enum, auto

from babel.numbers import get_decimal_symbol

from .errors import VBAError
from .locale import current_locale, display_date, parse_date_text, parse_local_number


class Special(Enum):
    EMPTY = auto()
    NULL = auto()
    MISSING = auto()


EMPTY, NULL, MISSING = Special
EPOCH = datetime(1899, 12, 30)


class Integer(int):
    pass


class Long(int):
    pass


class LongLong(int):
    pass


class Byte(int):
    pass


class Single(float):
    pass


class Currency(Decimal):
    pass


@dataclass(frozen=True)
class ErrorValue:
    code: int
    _vba_value_type = True
    _vba_type_name = "Error"

    def __str__(self):
        return f"Error {self.code}"


_INT_TYPES = {"byte": (Byte, 0, 255), "integer": (Integer, -32768, 32767),
              "long": (Long, -2147483648, 2147483647),
              "longlong": (LongLong, -(2**63), 2**63 - 1)}
_NUMERIC = (int, float, Decimal)
_SCALAR = _NUMERIC + (str, datetime, Special)
_VAL_PREFIX = re.compile(r"^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eEdD][+-]?[0-9]+)?")
_BASED_NUMBER = re.compile(r"&(?:[hH][0-9a-fA-F]+|[oO]?[0-7]+)")
_DECIMAL_LIMIT = Decimal(2**96 - 1)
_DATE_MIN = (datetime(100, 1, 1) - EPOCH).days
_DATE_MAX = (datetime(9999, 12, 31) - EPOCH).days
_KNOWN_TYPES = {
    "variant", "object", "boolean", "string", "date", "currency", "decimal", "single",
    "double", "longptr", "integer", "long", "longlong", "byte",
}
_COMPARE = {"=": operator.eq, "<>": operator.ne, "<": operator.lt,
            ">": operator.gt, "<=": operator.le, ">=": operator.ge}


def sentinel(name):
    return Special[name]


def is_object(value):
    return value is None or not isinstance(value, _SCALAR) and not getattr(type(value), "_vba_value_type", False)


def default_value(value):
    if value is None:
        raise VBAError(91)
    getter = getattr(type(value), "_vba_default_value", None)
    return getter(value) if getter else value


def date(text):
    # Kept as the entry point for date literals in previously emitted scripts.
    try:
        return parse_date_text(text, literal=True)
    except ValueError:
        raise VBAError(13, f"Invalid date literal {text!r}") from None


def _date_from_text(text):
    try:
        return parse_date_text(text)
    except ValueError:
        try:
            return from_date_serial(numeric(text))
        except VBAError as error:
            if error.number not in {6, 13}:
                raise
            raise VBAError(13, f"Invalid date {text!r}") from None


def date_serial(value):
    delta = value - EPOCH
    fraction = (delta.seconds + delta.microseconds / 1e6) / 86400
    return delta.days - fraction if delta.days < 0 else delta.days + fraction


def from_date_serial(value):
    if not math.isfinite(value) or not _DATE_MIN - 1 < value < _DATE_MAX + 1:
        raise VBAError(6)
    whole = math.trunc(value)
    try:
        return EPOCH + timedelta(days=whole, seconds=float(abs(value - whole)) * 86400)
    except OverflowError:
        raise VBAError(6) from None


def numeric(value, *, exact=False):
    value = default_value(value)
    if value is EMPTY:
        return Integer(0)
    if value is NULL:
        raise VBAError(94)
    if value is MISSING:
        raise VBAError(449)
    if isinstance(value, bool):
        return Integer(-int(value))
    if isinstance(value, datetime):
        return date_serial(value)
    if isinstance(value, _NUMERIC):
        return value
    if isinstance(value, str):
        text = value.strip()
        if _BASED_NUMBER.fullmatch(text):
            return number(text)
        try:
            result = parse_local_number(text)
        except ValueError:
            raise VBAError(13) from None
        if exact:
            return result
        if result.as_tuple().exponent == 0 and -(2**31) <= result < 2**31:
            return number(str(result))
        return coerce(result, "double")
    raise VBAError(13)


def integer(value):
    try:
        parsed = numeric(value, exact=True)
        # Bound the generic integer conversion to VBA's largest signed integer.
        if not -(2**63) - 1 < parsed < 2**63:
            raise VBAError(6)
        result = round(parsed)
        if not -(2**63) <= result < 2**63:
            raise VBAError(6)
        return result
    except (OverflowError, ValueError, InvalidOperation):
        raise VBAError(6) from None


def invariant_number(value):
    """Numeric text using the decimal point, for Str and portable file output."""
    if value == 0:
        return "0"
    if isinstance(value, Decimal):
        text = format(value, "f")
        return text.rstrip("0").rstrip(".") if "." in text else text
    if isinstance(value, float):
        precision = 7 if isinstance(value, Single) else 15
        return format(value, f".{precision}g").upper()
    return str(value)


def string(value):
    value = default_value(value)
    if value is EMPTY:
        return ""
    if value is NULL:
        raise VBAError(94)
    if value is MISSING:
        raise VBAError(449)
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, datetime):
        return display_date(value)
    if isinstance(value, str):
        return value
    if isinstance(value, ErrorValue):
        return str(value)
    if isinstance(value, _NUMERIC):
        return invariant_number(value).replace(".", get_decimal_symbol(current_locale()))
    raise VBAError(13)


def truth(value):
    value = default_value(value)
    if value is NULL or value is EMPTY:
        return False
    if isinstance(value, str):
        if value in {"#TRUE#", "#FALSE#"}:
            return value == "#TRUE#"
        if value.casefold() in {"true", "false"}:
            return value.casefold() == "true"
        return bool(numeric(value))
    if isinstance(value, bool):
        return value
    return bool(numeric(value))


def coerce(value, type_name, fixed_length=None):
    kind = type_name.casefold()
    if kind == "variant":
        return value
    if kind == "object" or kind not in _KNOWN_TYPES:
        if is_object(value) or getattr(type(value), "_vba_value_type", False):
            return value
        raise VBAError(424)
    if kind == "string":
        result = string(value)
        return result[:fixed_length].ljust(fixed_length) if fixed_length is not None else result
    if value is NULL:
        raise VBAError(94)
    if kind == "boolean":
        return truth(value)
    if kind == "date":
        value = default_value(value)
        if isinstance(value, datetime):
            if value.year < 100:
                raise VBAError(6)
            return value
        return _date_from_text(value) if isinstance(value, str) else from_date_serial(numeric(value))
    if kind == "longptr":
        kind = "longlong"
    if kind in _INT_TYPES:
        cls, lower, upper = _INT_TYPES[kind]
        result = integer(value)
        if not lower <= result <= upper:
            raise VBAError(6)
        return cls(result)
    try:
        if kind in {"decimal", "currency"}:
            result = Decimal(str(numeric(value, exact=True)))
            if kind == "currency":
                result = result.quantize(Decimal("0.0001"), rounding=ROUND_HALF_EVEN)
                if not Decimal("-922337203685477.5808") <= result <= Decimal("922337203685477.5807"):
                    raise VBAError(6)
                return Currency(result)
            return _decimal96(result)
        result = float(numeric(value))
        if not math.isfinite(result):
            raise VBAError(6)
        if kind == "single":
            # Some platforms produce infinity when packing an out-of-range Single.
            result = struct.unpack("f", struct.pack("f", result))[0]
            if not math.isfinite(result):
                raise VBAError(6)
            return Single(result)
        return result
    except (OverflowError, InvalidOperation):
        raise VBAError(6) from None


def _decimal96(value):
    if not value.is_finite() or value.copy_abs() > _DECIMAL_LIMIT:
        raise VBAError(6)
    with localcontext() as ctx:
        ctx.prec = 60
        scale = min(28, max(0, -value.as_tuple().exponent))
        integral_digits = max(0, value.adjusted() + 1)
        scale = min(scale, max(0, 29 - integral_digits))
        rounded = value.quantize(Decimal((0, (1,), -scale)), rounding=ROUND_HALF_EVEN)
        if scale and rounded.copy_abs().scaleb(scale) > _DECIMAL_LIMIT:
            rounded = value.quantize(Decimal((0, (1,), 1 - scale)), rounding=ROUND_HALF_EVEN)
        return rounded


def default_for(type_name, fixed_length=None):
    kind = type_name.casefold()
    if kind == "variant":
        return EMPTY
    if kind == "object" or kind not in _KNOWN_TYPES:
        return None
    return coerce(EMPTY, type_name, fixed_length)


def number(text):
    suffixes = {"%": "integer", "&": "long", "^": "longlong", "!": "single", "#": "double", "@": "currency"}
    suffix = text[-1] if text[-1] in suffixes else ""
    core = text[:-1] if suffix else text
    if core.startswith("&"):
        base = 16 if core[1].lower() == "h" else 8
        offset = 2 if core[1].isalpha() else 1
        result = int(core[offset:], base)
        bits = 16 if suffix == "%" or not suffix and result <= 0xFFFF else 64 if suffix == "^" else 32
        if 2 ** (bits - 1) <= result < 2 ** bits:
            result -= 2 ** bits
        return coerce(result, {16: "integer", 32: "long", 64: "longlong"}[bits])
    if suffix == "@":
        return coerce(Decimal(core), "currency")
    result = float(core) if any(char in core for char in ".eE") else int(core)
    if suffix:
        return coerce(result, suffixes[suffix])
    if isinstance(result, float) or not -(2**31) <= result < 2**31:
        return coerce(result, "double")
    return Integer(result) if -32768 <= result <= 32767 else Long(result)


def desc(value):
    if value is EMPTY:
        return "Empty"
    if value is NULL:
        return "Null"
    if value is MISSING:
        return "Error"
    if value is None:
        return "Nothing"
    for key, cls in ((bool, "Boolean"), (Byte, "Byte"), (Integer, "Integer"), (Long, "Long"),
                     (LongLong, "LongLong"), (Single, "Single"), (int, "Long"), (float, "Double"),
                     (Currency, "Currency"), (Decimal, "Decimal"), (datetime, "Date"), (str, "String")):
        if isinstance(value, key):
            return cls
    label = getattr(value, "_vba_type_name", None)
    return label or type(value).__name__


def like(value, pattern, compare):
    if compare == "text":
        value, pattern = value.casefold(), pattern.casefold()
    result = []
    position = 0
    while position < len(pattern):
        char = pattern[position]
        if char == "[":
            end = pattern.find("]", position + 1)
            if end < 0:
                raise VBAError(93, "Invalid pattern string")
            chars = pattern[position + 1:end]
            position = end
            if chars:
                negate = chars.startswith("!") and len(chars) > 1
                chars = chars[1:] if negate else chars
                chars = "".join(c if c == "-" and i not in {0, len(chars) - 1} else re.escape(c)
                                for i, c in enumerate(chars))
                result.append("[" + ("^" if negate else "") + chars + "]")
        else:
            result.append({"*": ".*", "?": ".", "#": "[0-9]"}.get(char, re.escape(char)))
        position += 1
    try:
        return re.fullmatch("".join(result), value, re.S) is not None
    except re.error:
        raise VBAError(93, "Invalid pattern string") from None


def _compare(op, left, right, mode):
    if left is EMPTY:
        left = "" if isinstance(right, str) else 0
    if right is EMPTY:
        right = "" if isinstance(left, str) else 0
    if isinstance(left, str) and isinstance(right, str):
        if mode == "text":
            left, right = left.casefold(), right.casefold()
    elif isinstance(left, str) or isinstance(right, str):
        # A Variant string sorts after a Variant numeric value.
        return _COMPARE[op](isinstance(left, str), isinstance(right, str))
    else:
        left, right = numeric(left), numeric(right)
    return _COMPARE[op](left, right)


def _logical(op, left, right):
    logical = isinstance(left, bool) and isinstance(right, bool)
    if NULL in (left, right):
        other = right if left is NULL else left
        if other is NULL:
            return NULL
        if op == "and" and integer(other) == 0:
            return False
        if op == "or" and integer(other) == -1:
            return True
        if op == "imp" and (left is NULL and integer(right) == -1 or
                             right is NULL and integer(left) == 0):
            return True
        return NULL
    a, b = integer(left), integer(right)
    result = {"and": operator.and_, "or": operator.or_, "xor": operator.xor,
              "eqv": lambda x, y: ~(x ^ y), "imp": lambda x, y: ~x | y}[op](a, b)
    if logical:
        return bool(result)
    if isinstance(left, LongLong) or isinstance(right, LongLong):
        return LongLong(result)
    if isinstance(left, (Long, int)) and not isinstance(left, (Integer, bool, Byte)) or isinstance(right, (Long, int)) and not isinstance(right, (Integer, bool, Byte)):
        return coerce(result, "long")
    return coerce(result, "integer")


def _arithmetic_result(result, left, right):
    types = (type(left), type(right))
    if isinstance(result, Decimal):
        kind = "currency" if Currency in types and Decimal not in types else "decimal"
        return coerce(result, kind)
    if float in types:
        return coerce(result, "double")
    if Single in types:
        return coerce(result, "single")
    if LongLong in types:
        return coerce(result, "longlong")
    if Long in types or int in types:
        return coerce(result, "long")
    return coerce(result, "integer")


def binary(op, left, right, compare="binary"):
    if op == "is":
        if not is_object(left) or not is_object(right):
            raise VBAError(424)
        return left is right
    left, right = default_value(left), default_value(right)
    if op in {"and", "or", "xor", "eqv", "imp"}:
        return _logical(op, left, right)
    if op == "&":
        if left is NULL and right is NULL:
            return NULL
        return string(EMPTY if left is NULL else left) + string(EMPTY if right is NULL else right)
    if left is NULL or right is NULL:
        return NULL
    if op in _COMPARE:
        return _compare(op, left, right, compare)
    if op == "like":
        return like(string(left), string(right), compare)
    if op == "+":
        if isinstance(left, str) and (isinstance(right, str) or right is EMPTY):
            return left + string(right)
        if left is EMPTY and isinstance(right, str):
            return right
    a, b = numeric(left), numeric(right)
    if op in {"\\", "mod"}:
        a, b = integer(a), integer(b)
        if not b:
            raise VBAError(11)
        quotient = abs(a) // abs(b) * (-1 if (a < 0) != (b < 0) else 1)
        return _arithmetic_result(quotient if op == "\\" else a - quotient * b, numeric(left), numeric(right))
    if op in {"/", "^"}:
        try:
            result = float(a) / float(b) if op == "/" else math.pow(a, b)
        except ZeroDivisionError:
            raise VBAError(11) from None
        except ValueError:
            raise VBAError(5) from None
        except OverflowError:
            raise VBAError(6) from None
        return coerce(result, "double")
    left_number, right_number = a, b
    arithmetic = {"+": operator.add, "-": operator.sub, "*": operator.mul}[op]
    if isinstance(a, Decimal) or isinstance(b, Decimal):
        if isinstance(a, float) or isinstance(b, float):
            result = arithmetic(float(a), float(b))
        else:
            # A product of two 96-bit Decimal values can need 58 digits
            # before rounding to the VBA result subtype.
            with localcontext() as context:
                context.prec = 60
                result = arithmetic(Decimal(a), Decimal(b))
    else:
        result = arithmetic(a, b)
    if op in {"+", "-"} and (isinstance(left, datetime) != isinstance(right, datetime)):
        return from_date_serial(result)
    return _arithmetic_result(result, left_number, right_number)


def unary(op, value):
    value = default_value(value)
    if value is NULL:
        return NULL
    if op == "not":
        return not value if isinstance(value, bool) else LongLong(~integer(value)) if isinstance(value, LongLong) else Long(~integer(value))
    operand = numeric(value)
    if op == "+":
        return operand
    if isinstance(operand, Long):
        return coerce(-operand, "longlong") if operand == -(2**31) else coerce(-operand, "long")
    if isinstance(operand, Integer) and operand == -32768:
        return Long(-operand)
    negated = operand.copy_negate() if isinstance(operand, Decimal) else -operand
    return _arithmetic_result(negated, operand, operand)


def val(value):
    text = re.sub(r"[ \t\r\n]", "", string(value))
    match = _BASED_NUMBER.match(text)
    if match:
        literal = match[0]
        suffix = text[match.end():match.end() + 1]
        return number(literal + suffix) if suffix in {"&", "^"} else number(literal)

    match = _VAL_PREFIX.match(text)
    if match is None:
        return Integer(0)
    literal = match[0].replace("D", "E").replace("d", "e")
    suffix = text[match.end():match.end() + 1]
    if suffix == "$":
        raise VBAError(13)
    if suffix in {"%", "&", "^"}:
        value = Decimal(literal)
        if value != value.to_integral_value():
            raise VBAError(13)
        return coerce(value, {"%": "integer", "&": "long", "^": "longlong"}[suffix])
    return number(literal + suffix) if suffix in {"@", "!", "#"} else number(literal)
