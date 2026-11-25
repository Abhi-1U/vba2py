"""Lark grammar and source-tree transformation for VBA expressions."""

from importlib.resources import files

from lark import Lark, Transformer, UnexpectedInput, v_args

from ..model import CompileError, Expr


@v_args(inline=True)
class ExpressionTransformer(Transformer):
    def binary(self, value, *tail):
        for operator, right in zip(tail[::2], tail[1::2]):
            value = Expr("binary", (str(operator).casefold(), value, right))
        return value

    def unary(self, operator, value):
        return Expr("unary", (str(operator).casefold(), value))

    def postfix(self, value, *trailers):
        for kind, data in trailers:
            value = Expr(kind, (value, data))
        return value

    def member(self, name):
        return "member", str(name)

    def invoke(self, arguments=()):
        return "invoke", arguments

    def name(self, name):
        return Expr("name", (str(name),))

    def with_member(self, name):
        return Expr("member", (Expr("with", ()), str(name)))

    def group(self, value):
        return Expr("group", (value,))

    def qualified_name(self, *parts):
        return ".".join(parts)

    def new(self, _keyword, name):
        return Expr("new", (str(name),))

    def type_of(self, _keyword, value, _is, name):
        return Expr("type_of", (value, str(name)))

    def string(self, value):
        return Expr("literal", (value[1:-1].replace('""', '"'),))

    def date(self, value):
        return Expr("date", (value[1:-1],))

    def number(self, value):
        return Expr("number", (str(value),))

    based = number

    def true(self, _):
        return Expr("literal", (True,))

    def false(self, _):
        return Expr("literal", (False,))

    def nothing(self, _):
        return Expr("literal", (None,))

    def null(self, _):
        return Expr("sentinel", ("NULL",))

    def empty(self, _):
        return Expr("sentinel", ("EMPTY",))

    def positional(self, value):
        return None, value

    def named(self, name, value):
        return str(name), value

    def missing(self):
        return None, None

    def arguments(self, *args):
        return tuple(arg if arg is not None else (None, None) for arg in args)

    def expression_list(self, *expressions):
        return expressions

    def dimensions(self, *dimensions):
        return dimensions

    def dimension(self, first, *rest):
        return (first, rest[-1]) if rest else (None, first)


_PARSER = Lark(
    files("vba2py").joinpath("grammar/expressions.lark").read_text(encoding="utf-8"),
    parser="lalr", lexer="contextual", transformer=ExpressionTransformer(),
    start=["expression", "arguments", "expression_list", "dimensions"],
)


def parse(text: str, start: str = "expression"):
    try:
        return _PARSER.parse(text.strip(), start=start)
    except UnexpectedInput as error:
        context = (error.get_context(text.strip()).strip().splitlines() or ["<end of input>"])[0]
        raise CompileError(f"invalid {start.replace('_', ' ')} near column {error.column}: {context}") from None
