"""Turn physical VBA source lines into statements without losing line locations."""

import re
from dataclasses import dataclass

from ..model import CompileError


@dataclass(frozen=True)
class SourceLine:
    number: int
    text: str


_STRING_OR_DATE = re.compile(r'"(?:[^"\n]|"")*"|(?<![\w.])#\s*(?:\d|[A-Za-z]{3,9}(?=[ /,\-]+\d))[^#"\n]*#(?!\d)')
_CONTINUATION = re.compile(r"[ \t]+_\s*$")


def masked(text):
    return _STRING_OR_DATE.sub(lambda match: " " * len(match[0]), text)


def split_top(text, separator=","):
    """Split on punctuation or a regex word pattern outside strings and parentheses."""
    pattern = re.compile(re.escape(separator)) if len(separator) == 1 else re.compile(separator, re.I)
    clean = masked(text)
    parts = []
    depth = 0
    last = 0
    position = 0
    while position < len(clean):
        char = clean[position]
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        elif depth == 0:
            match = pattern.match(clean, position)
            if match and not (separator == ":" and clean[position + 1:position + 2] == "="):
                parts.append(text[last:position].strip())
                position = last = match.end()
                continue
        position += 1
    parts.append(text[last:].strip())
    return parts


def _without_comment(text):
    clean = masked(text)
    apostrophe = clean.find("'")
    if apostrophe >= 0:
        text = text[:apostrophe]
        clean = clean[:apostrophe]
    remark = re.search(r"(?:^|(?<!:)\:(?!=))\s*Rem(?:\s|$)", clean, re.I)
    if remark:
        text = text[:remark.start()]
    return text.rstrip()


def logical_lines(source):
    pending = []
    start = 0
    for number, physical in enumerate(source.lstrip("\ufeff").splitlines(), 1):
        physical = _without_comment(physical)
        if not pending:
            start = number
        continuation = _CONTINUATION.search(physical)
        pending.append(physical[:continuation.start()] if continuation else physical)
        if continuation:
            continue
        text = " ".join(pending).strip()
        pending.clear()
        if text:
            yield SourceLine(start, text)
    if pending:
        raise CompileError(f"line {start}: unfinished line continuation")


def source_lines(source, constants=None):
    """Resolve conditional compilation using the caller's target constants."""
    from ..runtime.errors import VBAError
    from ..runtime.values import binary, number, sentinel, truth, unary
    from .expressions import parse

    predefined = {"vba6": False, "vba7": True, "win16": False,
                  "win32": False, "win64": False, "mac": False}
    predefined.update({key.casefold(): value for key, value in (constants or {}).items()})

    def evaluate(expression):
        kind, args = expression.kind, expression.args
        if kind == "name":
            return predefined.get(args[0].casefold(), 0)
        if kind == "literal":
            return args[0]
        if kind == "number":
            return number(args[0])
        if kind == "sentinel":
            return sentinel(args[0])
        if kind == "group":
            return evaluate(args[0])
        if kind == "unary":
            return unary(args[0], evaluate(args[1]))
        if kind == "binary":
            return binary(args[0], evaluate(args[1]), evaluate(args[2]))
        raise CompileError("unsupported conditional compilation expression")

    def parse_directive(expression, line, condition=False):
        try:
            value = evaluate(parse(expression))
            return truth(value) if condition else value
        except (CompileError, VBAError) as error:
            raise CompileError(f"line {line}: {error}") from None

    stack = []
    active = True
    for line in logical_lines(source):
        directive = re.fullmatch(r"#(If|ElseIf|Else|End\s+If|Const)\b\s*(.*)", line.text, re.I)
        if not directive:
            if active:
                yield line
            continue
        kind, argument = directive[1].casefold(), directive[2]
        if kind == "const":
            if active:
                name, delimiter, expression = argument.partition("=")
                if not delimiter or not name.strip().isidentifier():
                    raise CompileError(f"line {line.number}: invalid #Const declaration")
                predefined[name.strip().casefold()] = parse_directive(expression, line.number)
        elif kind == "if":
            expression = re.fullmatch(r"(.+)\s+Then", argument, re.I)
            if not expression:
                raise CompileError(f"line {line.number}: expected Then")
            taken = active and parse_directive(expression[1], line.number, condition=True)
            stack.append((active, taken, False))
            active = taken
        elif not stack:
            raise CompileError(f"line {line.number}: unmatched #{directive[1]}")
        elif kind.startswith("end"):
            if argument:
                raise CompileError(f"line {line.number}: unexpected text after #End If")
            active, _, _ = stack.pop()
        else:
            outer, taken, saw_else = stack[-1]
            if saw_else:
                raise CompileError(f"line {line.number}: branch after #Else")
            if kind == "elseif":
                expression = re.fullmatch(r"(.+)\s+Then", argument, re.I)
                if not expression:
                    raise CompileError(f"line {line.number}: expected Then")
                active = outer and not taken and parse_directive(expression[1], line.number, condition=True)
            else:
                if argument:
                    raise CompileError(f"line {line.number}: unexpected text after #Else")
                active = outer and not taken
            stack[-1] = outer, taken or active, kind == "else"
    if stack:
        raise CompileError("unterminated #If")
