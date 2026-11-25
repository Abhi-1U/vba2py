"""Public APIs for compiling source modules and macro workbooks."""

from dataclasses import dataclass
from functools import cached_property

from .codegen import emit
from .extraction import SourceModule, extract_modules
from .model import Diagnostic, ProjectError
from .parser.parser import parse_module
from .semantic import validate


@dataclass(frozen=True)
class Compilation:
    python: str
    modules: tuple[SourceModule, ...]

    @cached_property
    def _factory(self):
        namespace = {"__name__": "vba2py_generated"}
        exec(compile(self.python, "<vba2py-generated>", "exec"), namespace)
        return namespace["create_program"]

    def create_program(self, workbook=None, **options):
        """Create an independent instance of the transpiled project."""
        return self._factory(workbook, **options)


def compile_modules(sources, *, conditional_constants=None) -> Compilation:
    """Compile SourceModule objects into portable Python source."""
    sources = tuple(sources)
    parsed, diagnostics = [], []
    for source in sources:
        if source.kind not in {"standard", "class", "document"}:
            diagnostics.append(Diagnostic(source.name, 1, f"Unsupported module type {source.kind!r}"))
            continue
        try:
            parsed.append(parse_module(source.name, source.source, source.kind, conditional_constants))
        except ProjectError as error:
            diagnostics.extend(error.diagnostics)
    diagnostics.extend(validate(parsed))
    if diagnostics:
        raise ProjectError(diagnostics)
    return Compilation(emit(parsed), sources)


def compile_source(source, *, name="Module1", kind="standard", conditional_constants=None) -> Compilation:
    return compile_modules((SourceModule(name, source, kind),), conditional_constants=conditional_constants)


def compile_workbook(workbook, *, conditional_constants=None) -> Compilation:
    return compile_modules(extract_modules(workbook), conditional_constants=conditional_constants)
