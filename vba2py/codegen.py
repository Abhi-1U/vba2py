"""Lower procedure blocks to Python with explicit VBA control-flow targets."""

from dataclasses import dataclass, field

from .model import Diagnostic, ProjectError, identifier


def _tuple(parts):
    return "(" + "".join(part + ", " for part in parts) + ")"


def _dict(parts):
    return "{" + ", ".join(f"{key!r}: {value}" for key, value in parts) + "}"


def _lambda(expression, with_key=None):
    return "None" if expression is None else "lambda f: " + ExpressionWriter(with_key).expr(expression)


class ExpressionWriter:
    def __init__(self, with_key=None):
        self.with_key = with_key

    def expr(self, node):
        kind, args = node.kind, node.args
        if kind == "literal":
            return repr(args[0])
        if kind in {"number", "date", "sentinel"}:
            return f"{kind}({args[0]!r})"
        if kind == "name":
            return f"f.get({args[0]!r})"
        if kind == "group":
            return self.expr(args[0])
        if kind == "with":
            if self.with_key is None:
                raise ValueError("Leading member access requires With")
            return f"f.with_object(f.temps.get({self.with_key}))"
        if kind == "new":
            return f"f.new({args[0]!r})"
        if kind == "type_of":
            return f"f.type_of({self.expr(args[0])}, {args[1]!r})"
        if kind == "binary":
            return f"f.binary({args[0]!r}, {self.expr(args[1])}, {self.expr(args[2])})"
        if kind == "unary":
            return f"unary({args[0]!r}, {self.expr(args[1])})"
        if kind == "member":
            return f"f.member({self.expr(args[0])}, {args[1]!r})"
        if kind == "invoke":
            return f"invoke({self.target(args[0])}, {self.arguments(args[1])})"
        raise ValueError(f"Unexpected expression {kind}")

    def target(self, node):
        kind, args = node.kind, node.args
        if kind == "member":
            return f"f.member({self.expr(args[0])}, {args[1]!r}, invoke=True)"
        if kind == "name":
            return f"f.target({args[0]!r})"
        return self.expr(node)

    def argument(self, node):
        if node is None:
            return "Argument(MISSING, omitted=True)"
        kind, args = node.kind, node.args
        if kind == "name":
            return f"f.arg_name({args[0]!r})"
        if kind == "member":
            return f"f.arg_member({self.expr(args[0])}, {args[1]!r})"
        if kind == "invoke":
            return f"f.arg_invoke({self.target(args[0])}, {self.arguments(args[1])})"
        return f"Argument({self.expr(node)})"

    def arguments(self, args):
        return _tuple(f"({name!r}, {self.argument(node)})" for name, node in args)

    def dimensions(self, dimensions):
        return "f.bounds(" + _tuple(
            f"({'MISSING' if lo is None else self.expr(lo)}, {self.expr(hi)})"
            for lo, hi in dimensions
        ) + ")"

    def assign(self, target, value, object_assignment):
        kind, args = target.kind, target.args
        if kind == "name":
            return f"f.assign_name({args[0]!r}, {value}, {object_assignment})"
        if kind == "member":
            return f"f.assign_member({self.expr(args[0])}, {args[1]!r}, {value}, object_assignment={object_assignment})"
        if kind == "invoke":
            head, arguments = args
            if head.kind == "member":
                obj, name = head.args
                return f"f.assign_member({self.expr(obj)}, {name!r}, {value}, {self.arguments(arguments)}, {object_assignment})"
            if head.kind == "name":
                source = f"f.index_target({head.args[0]!r})"
            else:
                source = self.target(head)
            return f"f.assign_index({source}, {self.arguments(arguments)}, {value}, {object_assignment})"
        raise ValueError("Expected an assignable expression")


@dataclass(eq=False)
class Target:
    index: int | None = None


@dataclass
class Instruction:
    line: int
    statements: list[str] = field(default_factory=list)
    next_label: Target | None = None
    branch: tuple[str, Target, Target] | None = None
    jump: Target | None = None
    dispatch: str | None = None
    returns: bool = False
    error_target: Target | None = None
    resume_target: Target | None = None


class ProcedureWriter:
    def __init__(self, module, procedure):
        self.module = module
        self.procedure = procedure
        self.instructions = []
        self.labels = {}
        self.references = {}
        self.loops = []
        self.with_keys = []
        self.line = procedure.line

    def error(self, message):
        return ProjectError([Diagnostic(self.module.name, self.line, message)])

    def mark(self, target):
        target.index = len(self.instructions)

    def target(self, name):
        key = identifier(name)
        self.references[key] = self.line
        return self.labels.setdefault(key, Target())

    def add(self, *statements, following=None, branch=None, jump=None, dispatch=None,
            returns=False, error_target=None, resume_target=None):
        self.instructions.append(Instruction(self.line, list(statements), following, branch,
                                             jump, dispatch, returns, error_target, resume_target))

    def block(self, body):
        for statement in body:
            self.line = statement.line
            try:
                self.statement(statement)
            except ValueError as error:
                if isinstance(error, ProjectError):
                    raise
                raise self.error(str(error)) from None

    def statement(self, statement):
        op, args = statement.op, statement.args
        line = statement.line
        writer = ExpressionWriter(self.with_keys[-1] if self.with_keys else None)
        expr = writer.expr
        if op == "nop":
            return
        if op == "label":
            name = identifier(args[0])
            target = self.labels.setdefault(name, Target())
            if target.index is not None:
                raise self.error(f"Duplicate label {args[0]!r}")
            self.mark(target)
            if name.isdigit():
                self.add(f"f.erl = {int(name)}")
            return
        if op == "assign":
            self.add(writer.assign(args[0], expr(args[1]), args[2]))
            return
        if op == "call":
            target, call_args = args[0].args
            self.add(f"invoke({writer.target(target)}, {writer.arguments(call_args)}, statement=True)")
            return
        if op == "if":
            end = Target()
            for condition, body in args[0]:
                yes, no = Target(), Target()
                self.line = line
                self.add(branch=(f"truth({expr(condition)})", yes, no), following=end)
                self.mark(yes)
                self.block(body)
                self.add(jump=end)
                self.mark(no)
            self.block(args[1])
            self.mark(end)
            return
        if op in {"for", "for_each"}:
            counter, *rest, body = args
            if counter.kind != "name":
                raise self.error("Expected a loop counter variable")
            name = counter.args[0]
            key = len(self.instructions)
            test, enter, end = Target(), Target(), Target()
            for_each = op == "for_each"
            if for_each:
                self.add(f"f.start_each({key}, {expr(rest[0])})", following=end)
                condition = f"f.next_each({key}, {name!r})"
            else:
                start, limit, step = map(expr, rest)
                self.add(f"f.start_for({key}, {start}, {limit}, {step}, {name!r})", following=end)
                condition = f"f.for_test({key}, {name!r})"
            self.mark(test)
            self.add(branch=(condition, enter, end), following=end)
            self.mark(enter)
            self.loops.append(("for", end))
            self.block(body)
            self.loops.pop()
            self.line = line
            step_code = [] if for_each else [f"f.for_step({key}, {name!r})"]
            self.add(*step_code, jump=test, following=end)
            self.mark(end)
            return
        if op == "loop":
            pre, invert_pre, post, invert_post, body, wend = args
            test, enter, end = Target(), Target(), Target()
            self.mark(test)
            if pre is not None:
                condition = ("not " if invert_pre else "") + f"truth({expr(pre)})"
                self.add(branch=(condition, enter, end), following=end)
            self.mark(enter)
            self.loops.append(("while" if wend else "do", end))
            self.block(body)
            self.loops.pop()
            self.line = line
            if post is None:
                self.add(jump=test)
            else:
                condition = ("not " if invert_post else "") + f"truth({expr(post)})"
                self.add(branch=(condition, enter, end), following=end)
            self.mark(end)
            return
        if op == "with":
            key = len(self.instructions)
            end = Target()
            self.add(f"f.temps[{key}] = f.with_object({expr(args[0])})", following=end)
            self.with_keys.append(key)
            self.block(args[1])
            self.with_keys.pop()
            self.mark(end)
            return
        if op == "select":
            key = len(self.instructions)
            end = Target()
            self.add(f"f.temps[{key}] = {expr(args[0])}", following=end)
            subject = f"f.select_value({key})"
            for alternatives, body in args[1]:
                if alternatives:
                    yes, no = Target(), Target()
                    checks = []
                    for operator, first, last in alternatives:
                        if operator == "range":
                            check = (f"f.case_range({subject}, {expr(first)}, {expr(last)})")
                        else:
                            check = f"truth(f.binary({operator!r}, {subject}, {expr(first)}))"
                        checks.append(check)
                    self.line = line
                    self.add(branch=(" or ".join(checks), yes, no), following=end)
                    self.mark(yes)
                self.block(body)
                self.add(jump=end)
                if alternatives:
                    self.mark(no)
            self.mark(end)
            return
        if op == "exit":
            kind = args[0]
            if kind in {"for", "do"}:
                for loop, end in reversed(self.loops):
                    if loop == kind:
                        self.add(jump=end)
                        return
                raise self.error(f"Exit {kind.title()} outside the matching loop")
            if kind != self.procedure.kind.split()[0]:
                raise self.error(f"Exit {kind.title()} outside the matching procedure")
            self.add(returns=True)
            return
        if op in {"goto", "gosub"}:
            target = self.target(args[0])
            if op == "gosub":
                self.add(f"f.gosubs.append({len(self.instructions) + 1})", jump=target)
            else:
                self.add(jump=target)
            return
        if op == "return":
            self.add(dispatch="f.go_return()")
            return
        if op == "error_mode":
            mode, label = args
            if mode == "goto":
                self.add(error_target=self.target(label))
            else:
                self.add(f"f.set_error_mode({mode!r}, None)")
            return
        if op == "resume":
            destination = "same" if not args or args[0] == "0" else "next" if args[0].casefold() == "next" else None
            if destination is None:
                self.add(resume_target=self.target(args[0]))
            else:
                self.add(dispatch=f"f.resume({destination!r})")
            return
        if op == "redim":
            preserve, declarations = args
            for item in declarations:
                self.add(f"f.redim({item.name!r}, {writer.dimensions(item.bounds)}, {preserve}, {item.type_name!r})")
            return
        if op == "erase":
            self.add(f"f.erase({_tuple(map(expr, args))})")
            return
        if op == "end":
            self.add("raise EndProgram()")
            return
        if op == "stop":
            self.add("raise VBAError(5, 'VBA Stop statement reached')")
            return
        if op == "raise_error":
            self.add(f"f.program.err.Raise({expr(args[0])})")
            return
        if op == "mid_assign":
            arguments, replacement = args
            original, start, *length = arguments
            length = expr(length[0]) if length else "MISSING"
            value = f"f.replace_mid({expr(original)}, {expr(start)}, {length}, {expr(replacement)})"
            self.add(writer.assign(original, value, False))
            return
        if op == "debug_print":
            values = _tuple(f"({'EMPTY' if field is None else expr(field)}, {separator!r})"
                            for field, separator in args[0])
            self.add(f"f.program.builtins.debug.print_fields({values})")
            return
        if op == "file_open":
            self.add(f"f.program.app.files.open({expr(args[0])}, {args[1]!r}, {expr(args[2])}, {args[3]!r})")
            return
        if op == "file_close":
            self.add(f"f.program.app.files.close(*{_tuple(map(expr, args))})")
            return
        if op == "file_print":
            channel, fields, write = args
            values = _tuple(f"({'EMPTY' if field is None else expr(field)}, {separator!r})" for field, separator in fields)
            self.add(f"f.program.app.files.print_line({expr(channel)}, {values}, {write})")
            return
        if op == "file_input":
            channel, targets, line_input = args
            read = f"_items = f.program.app.files.input({expr(channel)}, {len(targets)}, {line_input})"
            self.add(read, *(writer.assign(target, f"_items[{i}]", False) for i, target in enumerate(targets)))
            return
        if op == "file_rename":
            self.add(f"f.program.app.files.rename({expr(args[0])}, {expr(args[1])})")
            return
        raise self.error(f"Unsupported statement {op}")

    def source(self, python_name):
        self.block(self.procedure.body)
        for label, target in self.labels.items():
            if target.index is None:
                self.line = self.references[label]
                raise self.error(f"Undefined label {label!r}")
        lines = [f"def {python_name}(f):",
                 f"    # {self.module.name}.{self.procedure.name} (VBA line {self.procedure.line})",
                 "    pc = 0", f"    while pc < {len(self.instructions)}:", "        try:",
                 "            match pc:"]
        for index, instruction in enumerate(self.instructions):
            def inline(text):
                return "                    " + text

            lines.append(f"                case {index}:")
            lines.append(inline(f"f.line = {instruction.line}"))
            next_step = index + 1 if instruction.next_label is None else instruction.next_label.index
            lines.append(inline(f"following = {next_step}"))
            lines.extend(map(inline, instruction.statements))
            if instruction.error_target is not None:
                lines.append(inline(f"f.set_error_mode('goto', {instruction.error_target.index})"))
            if instruction.resume_target is not None:
                lines.append(inline(f"pc = f.resume({instruction.resume_target.index})"))
            elif instruction.returns:
                lines.extend([inline("f.program.err.Clear()"), inline("return f.result()")])
            elif instruction.branch is not None:
                condition, yes, no = instruction.branch
                lines.append(inline(f"pc = {yes.index} if {condition} else {no.index}"))
            elif instruction.jump is not None:
                lines.append(inline(f"pc = {instruction.jump.index}"))
            elif instruction.dispatch is not None:
                lines.append(inline(f"pc = {instruction.dispatch}"))
            else:
                lines.append(inline(f"pc = {index + 1}"))
        if self.instructions:
            lines.extend([
                "        except VBAError as error:",
                "            pc = f.handle_error(error, pc, following)",
                "        except OSError as error:",
                "            pc = f.handle_error(file_error(error), pc, following)",
                "        except (ValueError, TypeError, OverflowError, ZeroDivisionError) as error:",
                "            pc = f.handle_error(conversion_error(error), pc, following)",
                "    return f.result()",
            ])
        else:
            lines = lines[:3] + ["    return f.result()"]
        return "\n".join(lines)


def _declaration(decl):
    if decl.bounds is None:
        bounds = "None"
    else:
        bounds = "lambda f: " + ExpressionWriter().dimensions(decl.bounds) if decl.bounds else "()"
    return _dict([
        ("type", repr(decl.type_name)), ("length", repr(decl.fixed_length)), ("bounds", bounds),
        ("new", repr(decl.as_new)), ("initial", _lambda(decl.initial)),
        ("constant", repr(decl.constant)), ("static", repr(decl.static)), ("private", repr(decl.private)),
    ])


def _parameter(parameter):
    decl = parameter.declaration
    return _dict([
        ("key", repr(identifier(decl.name))), ("type", repr(decl.type_name)),
        ("length", repr(decl.fixed_length)), ("array", repr(decl.bounds is not None)),
        ("byref", repr(parameter.byref)), ("optional", repr(parameter.optional)),
        ("default", _lambda(parameter.default)), ("paramarray", repr(parameter.paramarray)),
    ])


def emit(modules):
    source = [
        '"""Generated from VBA by vba2py."""',
        "from vba2py.runtime.program import (",
        "    Argument, EMPTY, EndProgram, MISSING, Program, VBAError,",
        "    conversion_error, date, file_error, invoke, number, sentinel, truth, unary,",
        ")",
        "",
    ]
    metadata = []
    for module_index, module in enumerate(modules):
        procedures = []
        for proc_index, procedure in enumerate(module.procedures):
            python_name = f"_procedure_{module_index}_{proc_index}"
            source.extend(["", ProcedureWriter(module, procedure).source(python_name), ""])
            key = identifier(procedure.name)
            procedures.append(((key, procedure.kind), _dict([
                ("name", repr(procedure.name)), ("key", repr(key)), ("type", repr(procedure.return_type)),
                ("kind", repr(procedure.kind)), ("private", repr(procedure.private)), ("static", repr(procedure.static)),
                ("parameters", _tuple(map(_parameter, procedure.parameters))), ("body", python_name),
                ("locals", _dict((identifier(item.name), _declaration(item)) for item in procedure.locals)),
            ])))
        metadata.append((identifier(module.name), _dict([
            ("name", repr(module.name)), ("kind", repr(module.kind)), ("private", repr(module.private)),
            ("default", repr(module.default_member)), ("base", repr(module.option_base)), ("compare", repr(module.option_compare)),
            ("explicit", repr(module.option_explicit)), ("procedures", _dict(procedures)),
            ("declarations", _dict((identifier(item.name), _declaration(item)) for item in module.declarations)),
            ("types", _dict((name, _dict((identifier(item.name), _declaration(item)) for item in declarations))
                            for name, declarations in module.types.items())),
        ])))
    # Emit one module specification per line so generated diffs stay local.
    source.append("\nMODULES = {")
    for key, value in metadata:
        source.append(f"    {key!r}: {value},")
    source.extend([
        "}", "", "",
        "def create_program(workbook=None, **options):",
        "    return Program(MODULES, workbook, **options)",
        "", "",
        'if __name__ == "__main__":',
        "    from vba2py.cli import run_generated",
        "    raise SystemExit(run_generated(create_program))",
    ])
    return "\n".join(source) + "\n"
