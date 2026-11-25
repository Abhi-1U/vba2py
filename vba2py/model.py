"""Source model shared by the parser, validator, and Python emitter."""

from dataclasses import dataclass, field
from typing import Any


class CompileError(ValueError):
    """A source error or an unsupported language construct."""


@dataclass
class Diagnostic:
    module: str
    line: int
    message: str

    def __str__(self):
        return f"{self.module}:{self.line}: {self.message}"


class ProjectError(CompileError):
    def __init__(self, diagnostics):
        self.diagnostics = tuple(diagnostics)
        super().__init__("\n".join(map(str, self.diagnostics)))


@dataclass(frozen=True)
class Expr:
    kind: str
    args: tuple[Any, ...]


@dataclass
class Declaration:
    name: str
    type_name: str = "Variant"
    bounds: tuple[tuple[Expr | None, Expr], ...] | None = None
    fixed_length: int | None = None
    as_new: bool = False
    initial: Expr | None = None
    constant: bool = False
    static: bool = False
    private: bool = False


@dataclass
class Parameter:
    declaration: Declaration
    byref: bool = True
    optional: bool = False
    default: Expr | None = None
    paramarray: bool = False


@dataclass
class Statement:
    line: int
    op: str
    args: tuple[Any, ...] = ()


@dataclass
class Procedure:
    name: str
    kind: str
    line: int
    parameters: list[Parameter]
    return_type: str
    body: list[Statement]
    private: bool = False
    static: bool = False
    locals: list[Declaration] = field(default_factory=list)


@dataclass
class Module:
    name: str
    kind: str = "standard"
    procedures: list[Procedure] = field(default_factory=list)
    declarations: list[Declaration] = field(default_factory=list)
    types: dict[str, list[Declaration]] = field(default_factory=dict)
    default_member: str | None = None
    option_base: int = 0
    option_compare: str = "binary"
    option_explicit: bool = False
    private: bool = False


def identifier(name: str) -> str:
    return name.rstrip("%&^!#$@").casefold()


TYPE_SUFFIXES = {
    "%": "Integer", "&": "Long", "^": "LongLong", "!": "Single",
    "#": "Double", "$": "String", "@": "Currency",
}
