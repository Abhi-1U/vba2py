"""Project-level VBA declaration checks and Option Explicit resolution."""

from dataclasses import fields

from .model import Declaration, Diagnostic, Expr, Statement, identifier
from .runtime.builtins import CONSTANTS, CONVERSIONS, EXCEL_GLOBALS, FUNCTIONS
from .runtime.excel import Workbook, Worksheet

_BUILTINS = set(CONSTANTS) | set(CONVERSIONS) | set(FUNCTIONS) | EXCEL_GLOBALS | {
    "application", "err", "debug", "vba", "thisworkbook", "erl", "freefile", "eof", "lof", "loc",
    "seek", "input", "dir", "chdir", "curdir", "mkdir", "rmdir", "kill", "filecopy", "filelen",
    "rnd", "randomize", "msgbox", "array", "instr", "strcomp",
}
_DOCUMENT_MEMBERS = {name.casefold() for cls in (Workbook, Worksheet)
                     for name in dir(cls) if not name.startswith("_")} | {
    "application", "worksheets", "sheets", "parent", "readonly",
}


def _duplicates(items):
    seen = set()
    for key, spelling, line in items:
        if key in seen:
            yield spelling, line
        seen.add(key)


def _body_items(obj):
    if isinstance(obj, Statement):
        yield obj
        for arg in obj.args:
            yield from _body_items(arg)
    elif isinstance(obj, (list, tuple)):
        for item in obj:
            yield from _body_items(item)


def _walk_expr(obj):
    if isinstance(obj, Expr):
        yield obj
        for arg in obj.args:
            yield from _walk_expr(arg)
    elif isinstance(obj, (list, tuple)):
        for item in obj:
            yield from _walk_expr(item)
    elif isinstance(obj, Declaration):
        for field in fields(obj):
            yield from _walk_expr(getattr(obj, field.name))


def validate(modules):
    diagnostics = []
    global_names = {identifier(module.name) for module in modules if module.kind != "class"} | _BUILTINS
    for module in modules:
        if module.kind == "standard":
            global_names.update(identifier(decl.name) for decl in module.declarations if not decl.private)
            global_names.update(identifier(proc.name) for proc in module.procedures if not proc.private)
    for spelling, line in _duplicates((identifier(module.name), module.name, 1) for module in modules):
        diagnostics.append(Diagnostic(spelling, line, "Duplicate module name"))

    for module in modules:
        def issue(line, message):
            diagnostics.append(Diagnostic(module.name, line, message))

        own = {identifier(decl.name) for decl in module.declarations}
        functions = {identifier(proc.name) for proc in module.procedures}
        for name, line in _duplicates((identifier(decl.name), decl.name, 1) for decl in module.declarations):
            issue(line, f"Duplicate module declaration {name!r}")
        for name, line in _duplicates(((identifier(proc.name), "procedure" if "property" not in proc.kind else proc.kind),
                                       proc.name, proc.line) for proc in module.procedures):
            issue(line, f"Duplicate procedure {name!r}")
        if module.default_member is not None and not any(
            identifier(proc.name) == module.default_member and not proc.private
            and proc.kind in {"function", "property get", "property let", "property set"}
            for proc in module.procedures
        ):
            issue(1, f"Invalid default member {module.default_member!r}")
        for key in own & functions:
            issue(1, f"Module declaration conflicts with procedure {key!r}")
        for procedure in module.procedures:
            key = identifier(procedure.name)
            local_decls = [item.declaration for item in procedure.parameters] + procedure.locals
            local_names = {identifier(decl.name) for decl in local_decls}
            for name, line in _duplicates((identifier(decl.name), decl.name, procedure.line)
                                          for decl in local_decls):
                issue(line, f"Duplicate local declaration {name!r}")
            if key in local_names and procedure.kind in {"function", "property get"}:
                issue(procedure.line, f"Local declaration conflicts with return value {procedure.name!r}")
            if not module.option_explicit:
                continue
            body = list(_body_items(procedure.body))
            implicit_arrays = {identifier(decl.name) for stmt in body if stmt.op == "redim"
                               for decl in stmt.args[1]}
            known = global_names | own | functions | local_names | implicit_arrays
            if module.kind != "standard":
                known |= {"me"}
            if module.kind == "document":
                known |= _DOCUMENT_MEMBERS
            diagnosed = set()

            def check(item, line):
                for expr in _walk_expr(item):
                    if expr.kind == "name":
                        name = expr.args[0]
                        check_key = line, identifier(name)
                        if check_key[1] not in known and check_key not in diagnosed:
                            issue(line, f"Name {name!r} is not declared (Option Explicit)")
                            diagnosed.add(check_key)

            for stmt in body:
                for item in stmt.args:
                    check(item, stmt.line)
            for decl in local_decls:
                check(decl, procedure.line)
            for param in procedure.parameters:
                check(param.default, procedure.line)
    return diagnostics
