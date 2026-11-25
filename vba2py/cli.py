"""Command line interface for extraction, translation, and execution."""

import argparse
import json
import re
import sys
from pathlib import Path

from .compiler import compile_modules, compile_workbook
from .extraction import SourceModule, extract_modules
from .model import CompileError
from .runtime.errors import VBAError
from .runtime.locale import resolve_locale
from .runtime.values import EMPTY


def _json(value):
    try:
        return json.loads(value)
    except json.JSONDecodeError as error:
        raise argparse.ArgumentTypeError(f"Expected a JSON value: {error.msg}") from None


def _binding(value):
    name, sep, data = value.partition("=")
    if not sep or not name.isidentifier():
        raise argparse.ArgumentTypeError("Expected NAME=JSON")
    return name, _json(data)


def _source(path):
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            text = path.read_text(encoding=encoding)
            break
        except UnicodeDecodeError:
            if encoding == "cp1252":
                raise
    attr = re.search(r'^Attribute\s+VB_Name\s*=\s*"([^"]+)"', text, re.I | re.M)
    kind = {".cls": "class", ".frm": "form"}.get(path.suffix.casefold(), "standard")
    if kind == "class" and re.search(r'^Attribute\s+VB_Base\s*=\s*"0\{000208', text, re.I | re.M):
        kind = "document"
    return SourceModule(attr[1] if attr else path.stem, text, kind)


def _locale(value):
    try:
        return resolve_locale(value)
    except ValueError as error:
        raise argparse.ArgumentTypeError(str(error)) from None


def _run_options(parser):
    parser.add_argument("entrypoint", help="Procedure name or Module.Procedure")
    parser.add_argument("--arg", type=_json, action="append", default=[], help="Positional JSON argument")
    parser.add_argument("--named", type=_binding, action="append", default=[], metavar="NAME=JSON")
    parser.add_argument("--save", type=Path, help="Save the macro workbook to this output path")
    parser.add_argument("--directory", type=Path, help="Initial VBA working directory")
    parser.add_argument("--locale", type=_locale, metavar="LOCALE",
                        help="VBA date and number locale, such as de-DE (default: en-US)")


def _execute(factory, workbook, args):
    workbook = workbook.resolve() if workbook else None
    program = factory(workbook, directory=args.directory, locale=args.locale)
    try:
        result = program.run(args.entrypoint, *args.arg, **dict(args.named))
        if args.save:
            if program.this_workbook is None:
                raise VBAError(91, "--save requires an input workbook")
            program.this_workbook.SaveAs(args.save.resolve())
        if result is not EMPTY:
            print(result)
    finally:
        program.close()
    return 0


def run_generated(factory, argv=None):
    parser = argparse.ArgumentParser(description="Run a transpiled VBA procedure")
    _run_options(parser)
    parser.add_argument("--workbook", type=Path, help="Workbook to open as ThisWorkbook")
    args = parser.parse_args(argv)
    try:
        return _execute(factory, args.workbook, args)
    except (VBAError, OSError) as error:
        print(error, file=sys.stderr)
        return 1


def main(argv=None):
    parser = argparse.ArgumentParser(prog="vba2py", description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    extract = subparsers.add_parser("extract", help="Extract source modules from a macro workbook")
    extract.add_argument("workbook", type=Path)
    extract.add_argument("-o", "--output", type=Path, required=True)
    translate = subparsers.add_parser("compile", help="Transpile a macro workbook or exported VBA sources")
    translate.add_argument("inputs", type=Path, nargs="+")
    translate.add_argument("-o", "--output", type=Path, required=True)
    translate.add_argument("-D", "--define", type=_binding, action="append", default=[], metavar="NAME=JSON")
    run = subparsers.add_parser("run", help="Compile and run a procedure from a macro workbook")
    run.add_argument("workbook", type=Path)
    run.add_argument("-D", "--define", type=_binding, action="append", default=[], metavar="NAME=JSON")
    _run_options(run)
    args = parser.parse_args(argv)
    try:
        if args.command == "extract":
            modules = extract_modules(args.workbook)
            args.output.mkdir(parents=True, exist_ok=True)
            for module in modules:
                suffix = {"standard": ".bas", "class": ".cls", "document": ".cls", "form": ".frm"}[module.kind]
                filename = re.sub(r"[^\w-]", "_", module.name) + suffix
                (args.output / filename).write_text(module.source, encoding="utf-8")
        elif args.command == "compile":
            constants = dict(args.define)
            workbook_types = {".xlsm", ".xltm"}
            if len(args.inputs) == 1 and args.inputs[0].suffix.casefold() in workbook_types:
                result = compile_workbook(args.inputs[0], conditional_constants=constants)
            elif any(path.suffix.casefold() not in {".bas", ".cls", ".frm"} for path in args.inputs):
                parser.error("Provide one macro workbook or one or more .bas/.cls/.frm source files")
            else:
                result = compile_modules(map(_source, args.inputs), conditional_constants=constants)
            args.output.write_text(result.python, encoding="utf-8")
        else:
            result = compile_workbook(args.workbook, conditional_constants=dict(args.define))
            return _execute(result.create_program, args.workbook, args)
    except (CompileError, VBAError, OSError, UnicodeError) as error:
        print(error, file=sys.stderr)
        return 1
    return 0
