"""Statement and block parser. Expression boundaries are parsed by Lark."""

import re
from collections import deque

from ..model import (
    CompileError, Declaration, Diagnostic, Expr, Module, Parameter, Procedure,
    ProjectError, Statement, TYPE_SUFFIXES, identifier,
)
from .expressions import parse
from .preprocess import SourceLine, masked, source_lines, split_top

_NAME = r"[^\W\d]\w*[%&^!#$@]?"
_DECLARATION = re.compile(
    rf"^({_NAME})(\s*\((.*)\))?(?:\s+As\s+(?:(New)\s+)?"
    r"([\w.]+)(?:\s*\*\s*(\d+))?)?$", re.I
)
_HEADER = re.compile(
    rf"^(?:(Public|Private|Friend)\s+)?(?:(Static)\s+)?"
    rf"(Sub|Function|Property\s+(?:Get|Let|Set))\s+({_NAME})(.*)$", re.I
)
_DEFINITION = re.compile(rf"^(?:(Public|Private)\s+)?(Type|Enum)\s+({_NAME})$", re.I)
_DECL_START = re.compile(r"^(?:(Public|Private|Global)\s+)?(?:(Dim|Static|Const)\b\s*)?", re.I)


def _split_once(text, separator):
    parts = split_top(text, separator)
    if len(parts) != 2:
        raise CompileError(f"expected {separator!r}")
    return parts


def _declaration(text, *, constant=False, static=False, private=False):
    segments = split_top(text, "=") if constant else [text]
    match = _DECLARATION.fullmatch(segments[0])
    if not match or constant and (len(segments) < 2 or not segments[1]):
        raise CompileError(f"invalid declaration {text!r}")
    name, array, dimensions, as_new, kind, length = match.groups()
    bounds = (parse(dimensions, "dimensions") if dimensions.strip() else ()) if array else None
    return Declaration(
        name, kind or TYPE_SUFFIXES.get(name[-1], "Variant"), bounds,
        int(length) if length else None, bool(as_new),
        parse("=".join(segments[1:])) if constant else None, constant, static, private,
    )


def _parameter(text):
    flags = set()
    while (match := re.match(r"^(Optional|ByRef|ByVal|ParamArray)\s+", text, re.I)):
        flag = match[1].casefold()
        if flag in flags or flag in {"byref", "byval"} and flags & {"byref", "byval"}:
            raise CompileError("conflicting parameter modifiers")
        flags.add(flag)
        text = text[match.end():]
    segments = split_top(text, "=")
    declaration = _declaration(segments[0])
    default = parse("=".join(segments[1:])) if len(segments) > 1 else None
    optional = "optional" in flags
    paramarray = "paramarray" in flags
    if default is not None and not optional:
        raise CompileError("a default value requires Optional")
    if paramarray and (flags != {"paramarray"} or declaration.bounds != ()
                       or declaration.type_name.casefold() != "variant"):
        raise CompileError("ParamArray requires an unqualified Variant array")
    if declaration.bounds not in (None, ()):
        raise CompileError("array parameters must have empty bounds")
    return Parameter(declaration, "byval" not in flags, optional, default, paramarray)


def _procedure(match, line):
    access, static, kind, name, tail = match.groups()
    tail = tail.strip()
    params = []
    if tail.startswith("("):
        depth = 0
        for index, char in enumerate(masked(tail)):
            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                if depth == 0:
                    arguments = tail[1:index].strip()
                    params = [_parameter(item) for item in split_top(arguments)] if arguments else []
                    tail = tail[index + 1:].strip()
                    break
        else:
            raise CompileError("unterminated parameter list")
    kind = " ".join(kind.casefold().split())
    type_match = re.fullmatch(r"As\s+([\w.]+)", tail, re.I) if tail else None
    if tail and (type_match is None or kind not in {"function", "property get"}):
        raise CompileError(f"unexpected procedure header text {tail!r}")
    return_type = type_match[1] if type_match else TYPE_SUFFIXES.get(name[-1], "Variant")
    if any(param.paramarray for param in params[:-1]):
        raise CompileError("ParamArray must be the final parameter")
    if any((param.optional or param.paramarray) for param in params):
        first_optional = next(i for i, param in enumerate(params) if param.optional or param.paramarray)
        if any(not (param.optional or param.paramarray) for param in params[first_optional:]):
            raise CompileError("required parameters must precede optional parameters")
    private = bool(access and access.casefold() == "private")
    return Procedure(name, kind, line, params, return_type, [], private, bool(static))


def _assignment(text):
    first = re.split(r"(?<![<>:=])=(?!=)", masked(text), maxsplit=1)
    if len(first) < 2:
        return None
    depth = first[0].count("(") - first[0].count(")")
    if depth != 0:
        return None
    index = len(first[0])
    try:
        lhs = parse(text[:index])
    except CompileError:
        return None
    if lhs.kind in {"name", "member", "invoke"}:
        return lhs, parse(text[index + 1:])
    return None


def _call(text, explicit=False):
    if explicit:
        result = parse(text)
        return result if result.kind == "invoke" else Expr("invoke", (result, ()))
    # In statement syntax F (x) passes a grouped expression ByVal, while F(x)
    # in expression syntax or Call F(x) has an ordinary argument list.
    for match in re.finditer(r"\s+", masked(text)):
        index = match.start()
        if not text[:index].strip() or not text[index:].strip():
            continue
        try:
            callee = parse(text[:index])
            if callee.kind in {"name", "member", "invoke"}:
                return Expr("invoke", (callee, parse(text[index:], "arguments")))
        except CompileError:
            continue
    result = parse(text)
    if result.kind not in {"invoke", "name", "member"}:
        raise CompileError(f"unsupported statement {text!r}")
    if result.kind == "invoke" and len(result.args[1]) == 1:
        name, argument = result.args[1][0]
        if name is None and argument is not None:
            return Expr("invoke", (result.args[0], ((None, Expr("group", (argument,))),)))
    return result if result.kind == "invoke" else Expr("invoke", (result, ()))


def _inline_else(text):
    # Else belongs to the closest preceding unmatched single-line If.
    depth = 0
    pending = 0
    at_start = True
    for token in re.finditer(r"\(|\)|:|[^\W\d]\w*|\S", masked(text)):
        word = token[0].casefold()
        if word == "(":
            depth += 1
        elif word == ")":
            depth -= 1
        if depth or word == ")":
            continue
        if at_start and word == "if":
            pending += 1
        if word == "else":
            if not pending:
                return text[:token.start()].strip(), text[token.end():].strip()
            pending -= 1
        at_start = word in {"then", "else", ":"}
    return text, ""


def _print_fields(payload):
    start = 0
    values = []
    for separator in re.finditer(r"[;,]", masked(payload)):
        fragment = masked(payload[start:separator.start()])
        if fragment.count("(") != fragment.count(")"):
            continue
        value = payload[start:separator.start()].strip()
        values.append((parse(value) if value else None, separator[0]))
        start = separator.end()
    last = payload[start:].strip()
    if last:
        values.append((parse(last), ""))
    return tuple(values)


def _file_statement(text, low):
    if low.startswith("name "):
        source, destination = _split_once(text[5:], r"\bAs\b")
        return "file_rename", (parse(source), parse(destination))
    if low.startswith("open "):
        match = re.fullmatch(r"Open\s+(.+?)\s+For\s+(Input|Output|Append|Binary|Random)"
                             r"(?:\s+Access\s+(Read Write|Read|Write))?\s+As\s+#?(.+)", text, re.I)
        if not match:
            raise CompileError("unsupported Open statement")
        return "file_open", (parse(match[1]), match[2].casefold(), parse(match[4]),
                             match[3].casefold() if match[3] else None)
    if low == "close" or low.startswith("close "):
        args = text[5:].strip()
        return "file_close", tuple(parse(item.lstrip("#")) for item in split_top(args)) if args else ()
    if (match := re.match(r"^(Print|Write)\s+#", text, re.I)):
        pieces = split_top(text[match.end():])
        channel = parse(pieces[0])
        payload = ",".join(pieces[1:])
        values = _print_fields(payload)
        return "file_print", (channel, tuple(values), match[1].casefold() == "write")
    if (match := re.match(r"^(Line\s+Input|Input)\s+#", text, re.I)):
        pieces = split_top(text[match.end():])
        if len(pieces) < 2:
            raise CompileError("expected input targets")
        return "file_input", (parse(pieces[0]), tuple(parse(target) for target in pieces[1:]),
                              match[1].casefold().startswith("line"))
    return None


class Parser:
    def __init__(self, module, lines):
        self.module = module
        self.lines = deque(lines)
        self.line = 1

    def take(self):
        source = self.lines.popleft()
        self.line = source.number
        return source.text

    def attribute(self, text):
        if (match := re.fullmatch(rf"Attribute\s+({_NAME})\.VB_UserMemId\s*=\s*0", text, re.I)):
            name = identifier(match[1])
            if self.module.default_member not in {None, name}:
                raise CompileError("Multiple default members")
            self.module.default_member = name

    def module_body(self):
        while self.lines:
            text = self.take()
            low = text.casefold()
            if low.startswith("attribute "):
                self.attribute(text)
                continue
            if low.startswith("version ") and self.module.kind in {"class", "document"}:
                if self.lines and self.lines[0].text.casefold() == "begin":
                    self.take()
                    while self.lines and self.take().casefold() != "end":
                        pass
                continue
            if low.startswith("option "):
                self.option(low)
                continue
            if (match := _DEFINITION.fullmatch(text)):
                self.definition(match)
                continue
            pieces = split_top(text, ":")
            if (match := _HEADER.fullmatch(pieces[0])):
                procedure = _procedure(match, self.line)
                if len(pieces) > 1:
                    self.lines.appendleft(SourceLine(self.line, ":".join(pieces[1:])))
                procedure.body = self.block(procedure, lambda end: end.startswith("end "))
                ending = self.take() if self.lines else ""
                if ending.casefold() != "end " + procedure.kind.split()[0]:
                    raise CompileError(f"expected End {procedure.kind.split()[0].title()}")
                self.module.procedures.append(procedure)
                continue
            for piece in pieces:
                decl = self.declarations(piece)
                if decl is None:
                    raise CompileError(f"unsupported module declaration {piece!r}")
                self.module.declarations.extend(decl)

    def option(self, text):
        if text == "option explicit":
            self.module.option_explicit = True
        elif text == "option private module":
            self.module.private = True
        elif (match := re.fullmatch(r"option base ([01])", text)):
            self.module.option_base = int(match[1])
        elif (match := re.fullmatch(r"option compare (text|binary)", text)):
            self.module.option_compare = match[1]
        else:
            raise CompileError(f"unsupported option {text!r}")

    def definition(self, match):
        access, kind, name = match.groups()
        private = bool(access and access.casefold() == "private")
        members = []
        prior = Expr("literal", (-1,))
        while self.lines:
            text = self.take()
            if text.casefold() == f"end {kind.casefold()}":
                if kind.casefold() == "type":
                    key = identifier(name)
                    if key in self.module.types:
                        raise CompileError(f"duplicate Type {name!r}")
                    self.module.types[key] = members
                else:
                    self.module.declarations.extend(members)
                return
            if kind.casefold() == "type":
                members.append(_declaration(text))
            else:
                values = split_top(text, "=")
                field = values[0]
                if not re.fullmatch(_NAME, field):
                    raise CompileError(f"invalid Enum member {field!r}")
                prior = parse("=".join(values[1:])) if len(values) > 1 else Expr(
                    "binary", ("+", prior, Expr("literal", (1,)))
                )
                members.append(Declaration(field, "Long", initial=prior, constant=True, private=private))
                prior = Expr("name", (field,))
        raise CompileError(f"expected End {kind}")

    def declarations(self, text):
        match = _DECL_START.match(text)
        if not match.end() or not any(match.groups()):
            return None
        access, mode = (item.casefold() if item else "" for item in match.groups())
        return [_declaration(part, constant=mode == "const", static=mode == "static",
                             private=access == "private" or not access and mode in {"dim", "const"})
                for part in split_top(text[match.end():])]

    def block(self, procedure, stop):
        result = []
        while self.lines:
            text = self.lines[0].text
            if stop(split_top(text, ":")[0].casefold()):
                break
            self.take()
            result.extend(self.line_statements(procedure, text, self.line))
        return result

    def _end(self, predicate, description):
        if not self.lines:
            raise CompileError(f"expected {description}")
        text = self.take()
        first, *rest = split_top(text, ":")
        if not predicate(first.casefold()):
            raise CompileError(f"expected {description}")
        if rest and any(rest):
            self.lines.appendleft(SourceLine(self.line, ":".join(rest)))
        return first

    def line_statements(self, procedure, text, line):
        if text.casefold().startswith("attribute "):
            self.attribute(text)
            return []
        # VBA accepts both Colon-terminated names and classic numeric line labels.
        parts = split_top(text, ":")
        result = []
        if len(parts) > 1 and re.fullmatch(rf"(?:{_NAME}|\d+)", parts[0]) and parts[0].casefold() not in {
            "stop", "return", "end", "close", "do", "else", "wend", "loop",
        }:
            result.append(Statement(line, "label", (parts.pop(0),)))
        elif (label := re.match(r"^(\d+)\s+(.*)", text)):
            result.append(Statement(line, "label", (label[1],)))
            parts = split_top(label[2], ":")
        while parts:
            piece = parts.pop(0)
            if not piece:
                continue
            low = piece.casefold()
            if low.startswith("if "):
                piece = ":".join([piece, *parts])
                parts.clear()
            elif (low in {"do", "else"} or low.startswith(("for ", "while ", "with ", "select case ",
                                                            "do while ", "do until "))) and parts:
                self.lines.appendleft(SourceLine(line, ":".join(parts)))
                parts.clear()
            result.append(self.statement(procedure, piece, line))
        return result

    def statement(self, procedure, text, line):
        low = text.casefold()
        op, args = self._statement(procedure, text, low, line)
        return Statement(line, op, args)

    def _statement(self, procedure, text, low, line):
        if low.startswith("if "):
            pieces = split_top(text[3:], r"\bThen\b")
            if len(pieces) < 2:
                raise CompileError("expected Then")
            condition = parse(pieces[0])
            tail = " Then ".join(pieces[1:]).strip()
            if tail and tail != ":":
                yes, no = _inline_else(tail)
                if yes.isdigit():
                    yes = "GoTo " + yes
                true_body = self.line_statements(procedure, yes, line)
                if no.isdigit():
                    no = "GoTo " + no
                false_body = self.line_statements(procedure, no, line) if no else []
                return "if", ([(condition, true_body)], false_body)
            stop = lambda part: part == "end if" or part == "else" or part.startswith("elseif ")
            arms = [(condition, self.block(procedure, stop))]
            while self.lines and self.lines[0].text.casefold().startswith("elseif "):
                branch = self.take()
                condition, tail = _split_once(branch[7:], r"\bThen\b")
                if tail and tail != ":":
                    raise CompileError("unexpected text after ElseIf Then")
                arms.append((parse(condition), self.block(procedure, stop)))
            other = []
            if self.lines and split_top(self.lines[0].text, ":")[0].casefold() == "else":
                self._end(lambda part: part == "else", "Else")
                other = self.block(procedure, lambda part: part == "end if")
            self._end(lambda part: part == "end if", "End If")
            return "if", (arms, other)
        if low.startswith("for "):
            if (each := re.fullmatch(r"Each\s+(.+?)\s+In\s+(.+)", text[4:], re.I)):
                op = "for_each"
                header = (parse(each[1]), parse(each[2]))
            else:
                variable, interval = _split_once(text[4:], "=")
                start, end = _split_once(interval, r"\bTo\b")
                steps = split_top(end, r"\bStep\b")
                if len(steps) > 2:
                    raise CompileError("invalid Step clause")
                op = "for"
                header = (parse(variable), parse(start), parse(steps[0]),
                          parse(steps[1]) if len(steps) == 2 else Expr("literal", (1,)))
            stop = lambda part: part == "next" or part.startswith("next ")
            body = self.block(procedure, stop)
            ending = self._end(stop, "Next")
            if ending[4:].strip():
                variables = split_top(ending[4:].strip())
                target = header[0]
                if target.kind != "name" or identifier(variables[0]) != identifier(target.args[0]):
                    raise CompileError("Next variable does not match For")
                if len(variables) > 1:
                    self.lines.appendleft(SourceLine(self.line, "Next " + ", ".join(variables[1:])))
            return op, (*header, body)
        if low == "do" or low.startswith(("do while ", "do until ", "while ")):
            wend = low.startswith("while ")
            pre = parse(text[6:]) if wend else parse(text[9:]) if low != "do" else None
            stop = (lambda part: part == "wend") if wend else (
                lambda part: part == "loop" or part.startswith(("loop while ", "loop until "))
            )
            body = self.block(procedure, stop)
            ending = self._end(stop, "Wend" if wend else "Loop")
            post = parse(ending[11:]) if ending.casefold().startswith("loop while ") else (
                parse(ending[11:]) if ending.casefold().startswith("loop until ") else None
            )
            if pre is not None and post is not None:
                raise CompileError("Do may only have one condition")
            return "loop", (pre, "until" in low[:8], post, ending.casefold().startswith("loop until "), body, wend)
        if low.startswith("with "):
            expression = parse(text[5:])
            stop = lambda part: part == "end with"
            body = self.block(procedure, stop)
            self._end(stop, "End With")
            return "with", (expression, body)
        if low.startswith("select case "):
            subject = parse(text[12:])
            branches = []
            seen_else = False
            stop = lambda part: part.startswith("case ") or part == "end select"
            while self.lines and self.lines[0].text.casefold().startswith("case "):
                branch = self.take()
                first, *rest = split_top(branch, ":")
                if rest and any(rest):
                    self.lines.appendleft(SourceLine(self.line, ":".join(rest)))
                if seen_else:
                    raise CompileError("Case after Case Else")
                patterns = []
                seen_else = first.casefold() == "case else"
                if not seen_else:
                    for pattern in split_top(first[5:]):
                        if (comparison := re.fullmatch(r"Is\s*(<=|>=|<>|[<>=])\s*(.+)", pattern, re.I)):
                            patterns.append((comparison[1], parse(comparison[2]), None))
                        else:
                            interval = split_top(pattern, r"\bTo\b")
                            if len(interval) > 2:
                                raise CompileError("invalid Case range")
                            patterns.append(("range", parse(interval[0]), parse(interval[1]))
                                            if len(interval) == 2 else ("=", parse(interval[0]), None))
                branches.append((patterns, self.block(procedure, stop)))
            self._end(lambda part: part == "end select", "End Select")
            return "select", (subject, branches)
        if (decl := self.declarations(text)) is not None:
            if any(item.static for item in decl) and procedure.static:
                raise CompileError("Static declaration inside a Static procedure")
            procedure.locals.extend(decl)
            return "nop", ()
        if low.startswith(("redim ", "erase ")):
            if low.startswith("erase "):
                return "erase", tuple(parse(part) for part in split_top(text[6:]))
            preserved = low.startswith("redim preserve ")
            declarations = tuple(_declaration(part) for part in split_top(text[15 if preserved else 6:]))
            if any(not declaration.bounds for declaration in declarations):
                raise CompileError("ReDim requires array bounds")
            return "redim", (preserved, declarations)
        if (match := re.fullmatch(r"(GoTo|GoSub|Resume|On\s+(?:Local\s+)?Error)\s+(.+)", text, re.I)):
            keyword, arg = match[1].casefold(), match[2].strip()
            if keyword.startswith("on"):
                if arg.casefold() == "resume next":
                    return "error_mode", ("next", None)
                label = re.fullmatch(rf"GoTo\s+({_NAME}|\d+|-1)", arg, re.I)
                if not label:
                    raise CompileError("invalid On Error target")
                return "error_mode", ("off" if label[1] in {"0", "-1"} else "goto", label[1])
            if not re.fullmatch(rf"{_NAME}|\d+", arg):
                raise CompileError(f"invalid {keyword} target")
            return keyword, (arg,)
        if low in {"return", "resume", "end", "stop"}:
            return low, ()
        if (match := re.fullmatch(r"Exit\s+(For|Do|Sub|Function|Property)", text, re.I)):
            return "exit", (match[1].casefold(),)
        if (file_statement := _file_statement(text, low)) is not None:
            return file_statement
        if low.startswith(("set ", "let ")):
            assignment = _assignment(text[4:])
            if assignment is None:
                raise CompileError("expected assignment")
            return "assign", (*assignment, low.startswith("set "))
        if (match := re.match(r"^Debug\s*\.\s*Print(?:\s+|$)", text, re.I)):
            payload = text[match.end():]
            return "debug_print", (_print_fields(payload),)
        if low.startswith("call "):
            return "call", (_call(text[5:], explicit=True),)
        if low == "error" or low.startswith("error "):
            return "raise_error", (parse(text[5:]),)
        if (match := re.match(r"^Mid\$?\s*\(", text, re.I)):
            assignment = _assignment(text)
            if assignment:
                _, arguments = assignment[0].args
                if len(arguments) not in {2, 3} or any(name is not None or arg is None for name, arg in arguments):
                    raise CompileError("invalid Mid assignment")
                return "mid_assign", (tuple(arg for _, arg in arguments), assignment[1])
        if (assignment := _assignment(text)) is not None:
            return "assign", (*assignment, False)
        if re.match(r"^(end|else|elseif|loop|next|wend|case|then|put|get|lset|rset|raiseevent)\b", low):
            raise CompileError(f"unexpected or unsupported statement {text!r}")
        return "call", (_call(text),)


def parse_module(name, source, kind="standard", conditional_constants=None):
    parser = None
    try:
        lines = list(source_lines(source, conditional_constants))
        parser = Parser(Module(name, kind), lines)
        parser.module_body()
        return parser.module
    except CompileError as error:
        match = re.match(r"line (\d+): (.*)", str(error))
        line, message = (int(match[1]), match[2]) if match else (
            parser.line if parser else 1, str(error)
        )
        raise ProjectError([Diagnostic(name, line, message)]) from None
