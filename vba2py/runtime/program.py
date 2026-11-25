"""Procedure frames, project/module scopes, argument binding, and late dispatch."""

import inspect
from dataclasses import dataclass
from functools import lru_cache

from ..model import TYPE_SUFFIXES, identifier
from . import values as v
from .arrays import VBAArray
from .builtins import Builtins, Collection, Dictionary, EXCEL_GLOBALS
from .errors import EndProgram, ErrObject, VBAError, conversion_error
from .excel import Application, Indexed, Range
from .filesystem import file_error
from .locale import resolve_locale, use_locale


def copy_value(value):
    return value.copy() if isinstance(value, (VBAArray, Record)) else value


class Slot:
    def __init__(self, type_name="Variant", value=v.EMPTY, *, length=None, constant=False,
                 factory=None, array=False):
        self.type_name, self.value = type_name, value
        self.length, self.constant, self.factory, self.array = length, constant, factory, array

    def get(self):
        if self.factory is not None and self.value is None:
            self.value = self.factory()
        return self.value

    def set(self, value, object_assignment=False):
        if self.constant:
            raise VBAError(5, "Cannot change a constant")
        if object_assignment:
            if not v.is_object(value) or self.array:
                raise VBAError(424)
        else:
            if self.factory is not None and self.value is None:
                self.get()
            if self.type_name.casefold() != "variant" and self.value is not None:
                setter = getattr(type(self.value), "_vba_assign_default", None)
                if setter is not None:
                    setter(self.value, value)
                    return
            if v.is_object(value):
                value = v.default_value(value)
            if v.is_object(value):
                raise VBAError(438, "Object has no default value")
        self.store(value)

    def store(self, value):
        if self.array:
            if not isinstance(value, VBAArray):
                raise VBAError(13)
            if self.type_name.casefold() != value.type_name.casefold():
                raise VBAError(13)
            if isinstance(self.value, VBAArray) and self.value.fixed:
                raise VBAError(10)
            self.value = value.copy()
            self.value.fixed = False
        else:
            value = v.coerce(value, self.type_name, self.length)
            self.value = copy_value(value)
            if isinstance(self.value, VBAArray):
                self.value.fixed = False


class Element:
    constant = False

    def __init__(self, array, indices):
        self.array, self.indices = array, indices
        self.type_name = array.type_name
        self.array._index(indices)

    def get(self):
        return self.array.get(self.indices)

    def set(self, value, object_assignment=False):
        if object_assignment and not v.is_object(value):
            raise VBAError(424)
        if not object_assignment and v.is_object(value):
            value = v.default_value(value)
        self.array.set(self.indices, copy_value(value))

    store = set


@dataclass
class Argument:
    value: object
    reference: Slot | Element | None = None
    omitted: bool = False


@dataclass
class BoundProcedure:
    program: object
    context: object
    key: tuple[str, str]

    def __call__(self, args):
        return self.program._invoke(self.context, self.key, args)


class Context:
    def __init__(self, program, spec, host=None):
        self.program, self.spec, self.host = program, spec, host
        self.fields = {}
        self.initializing = set()
        self.statics = {}
        self.initialized = False
        self._vba_type_name = spec["name"]

    def _vba_default_value(self):
        name = self.spec["default"]
        if name is None:
            return self
        return evaluate(self.lookup(name))

    def _vba_assign_default(self, value):
        name = self.spec["default"]
        if name is None:
            raise VBAError(438, "Object has no default property")
        self.set(name, value)

    def lookup(self, name, internal=False, assignment=None, host_fallback=True):
        key = identifier(name)
        declarations = self.spec["declarations"]
        if key in declarations and (internal or not declarations[key]["private"]):
            return self.program._initialize_field(self, key)
        kinds = (assignment,) if assignment else ("function", "sub", "property get")
        for kind in kinds:
            procedure = self.spec["procedures"].get((key, kind))
            if procedure is not None and (internal or not procedure["private"]):
                return BoundProcedure(self.program, self, (key, kind))
        if self.host is not None and host_fallback:
            try:
                return get_member(self.host, name)
            except VBAError as error:
                if error.number != 438:
                    raise
        return v.MISSING

    def set(self, name, value, args=(), object_assignment=False, internal=False):
        target = self.lookup(name, internal, "property set" if object_assignment else "property let")
        if isinstance(target, Slot) and not args:
            target.set(value, object_assignment)
        elif isinstance(target, Slot):
            _set_index(target, args, value, object_assignment)
        elif isinstance(target, BoundProcedure):
            final = target.context.spec["procedures"][target.key]["parameters"][-1]["key"]
            target((*args, (final if any(name is not None for name, _ in args) else None, Argument(value))))
        elif self.host is not None:
            assign_member(self.host, name, value, args, object_assignment)
        else:
            raise VBAError(438)


class Record:
    _vba_value_type = True

    def __init__(self, type_name, fields):
        self._vba_type_name = type_name
        self.fields = fields

    def copy(self):
        return Record(self._vba_type_name,
                      {key: Slot(field.type_name, copy_value(field.value),
                                 length=field.length, array=field.array)
                       for key, field in self.fields.items()})


class Library:
    def __init__(self, frame):
        self.frame = frame


@lru_cache(maxsize=256)
def _attributes(cls):
    return {name.casefold(): name for name in dir(cls) if not name.startswith("_")}


def _attribute(obj, name):
    key = identifier(name)
    if obj is None:
        raise VBAError(91)
    attribute = _attributes(type(obj)).get(key)
    if attribute is None and hasattr(obj, "__dict__"):
        attribute = next((item for item in vars(obj) if not item.startswith("_") and item.casefold() == key), None)
    if attribute is None:
        raise VBAError(438, f"{type(obj).__name__} has no member {name!r}")
    return attribute


def get_member(obj, name, *, invoke=False, caller=None):
    if isinstance(obj, Library):
        if identifier(name) == "strings" or identifier(name) in {"interaction", "conversion", "math", "datetime", "information", "filesystem", "financial"}:
            return obj
        result = obj.frame.program.builtins.get(identifier(name), obj.frame, qualified=True)
        if result is v.MISSING:
            raise VBAError(438, f"VBA has no member {name!r}")
        return result
    if identifier(name) == "address" and isinstance(obj, Range) and invoke:
        return obj.address
    if isinstance(obj, Context):
        result = obj.lookup(name, internal=caller is obj)
        if result is not v.MISSING:
            return result
        raise VBAError(438, f"{obj._vba_type_name} has no member {name!r}")
    if caller is not None and (document := caller.program.documents.get(id(obj))) is not None:
        result = document.lookup(name, internal=caller is document, host_fallback=False)
        if result is not v.MISSING:
            return result
    if isinstance(obj, Record):
        try:
            return obj.fields[identifier(name)]
        except KeyError:
            raise VBAError(438) from None
    return getattr(obj, _attribute(obj, name))


def assign_member(obj, name, value, args, object_assignment, caller=None):
    if isinstance(obj, Context):
        obj.set(name, value, args, object_assignment, internal=caller is obj)
        return
    kind = "property set" if object_assignment else "property let"
    if caller is not None and (document := caller.program.documents.get(id(obj))) is not None:
        if document.lookup(name, internal=caller is document, assignment=kind, host_fallback=False) is not v.MISSING:
            document.set(name, value, args, object_assignment, internal=caller is document)
            return
    if isinstance(obj, Record):
        target = get_member(obj, name)
        if args:
            _set_index(target.get(), args, value, object_assignment)
        else:
            target.set(value, object_assignment)
        return
    if args:
        if identifier(name) == "item" and isinstance(obj, Dictionary):
            _set_index(obj, args, value, object_assignment)
            return
        _set_index(get_member(obj, name, invoke=True), args, value, object_assignment)
        return
    if not object_assignment and v.is_object(value):
        value = v.default_value(value)
    attribute = _attribute(obj, name)
    descriptor = getattr(type(obj), attribute, None)
    if callable(descriptor) or isinstance(descriptor, property) and descriptor.fset is None:
        raise VBAError(450, f"Member {name!r} is read-only")
    setattr(obj, attribute, value)


def _python_arg(value):
    return value.get() if isinstance(value, (Slot, Element)) else value


_OMITTED = object()


def _expand_args(args, distinguish_omitted=False):
    positions, names = [], {}
    named = False
    for name, arg in args:
        value = _OMITTED if distinguish_omitted and arg.omitted else arg.value
        if name is None:
            if named:
                raise VBAError(448)
            positions.append(value)
        else:
            named = True
            key = identifier(name)
            if key in names:
                raise VBAError(448)
            names[key] = value
    return positions, names


def _native_call(function, args):
    positions, names = _expand_args(args, distinguish_omitted=True)
    try:
        signature = inspect.signature(function)
    except (TypeError, ValueError):
        if names:
            raise VBAError(448) from None
        return function(*positions)
    parameters = signature.parameters
    normalized = {identifier(name): name for name in parameters}
    keywords = {}
    for key, value in names.items():
        if key not in normalized and not any(p.kind is p.VAR_KEYWORD for p in parameters.values()):
            raise VBAError(448, f"Named argument {key!r} not found")
        keywords[normalized.get(key, key)] = value
    try:
        bound = signature.bind(*positions, **keywords)
    except TypeError as error:
        raise VBAError(450, str(error)) from None
    for key, value in tuple(bound.arguments.items()):
        parameter = parameters[key]
        if value is _OMITTED:
            if parameter.default is inspect.Parameter.empty:
                raise VBAError(449)
            bound.arguments[key] = parameter.default
    return function(*bound.args, **bound.kwargs)


def invoke(target, args, *, statement=False):
    target = _python_arg(target)
    if isinstance(target, BoundProcedure):
        return target(args)
    if isinstance(target, Context):
        name = target.spec["default"]
        if name is None:
            raise VBAError(438, "Object has no default property")
        return invoke(target.lookup(name), args)
    if isinstance(target, VBAArray):
        positions, names = _expand_args(args)
        if names:
            raise VBAError(448)
        return target.get(positions)
    if callable(target):
        return _native_call(target, args)
    if isinstance(target, Indexed):
        return _native_call(target.Item, args)
    item = getattr(type(target), "_vba_item", None)
    if item:
        return _native_call(item.__get__(target), args)
    if target is None:
        raise VBAError(91)
    if target is v.EMPTY:
        raise VBAError(35)
    if statement or args:
        raise VBAError(438)
    return target


def evaluate(target):
    value = _python_arg(target)
    return invoke(value, ()) if isinstance(value, BoundProcedure) or callable(value) else value


def _set_index(target, args, rhs, object_assignment):
    target = _python_arg(target)
    if isinstance(target, Context):
        name = target.spec["default"]
        if name is None:
            raise VBAError(438, "Object has no default property")
        target.set(name, rhs, args, object_assignment)
        return
    if isinstance(target, VBAArray):
        positions, names = _expand_args(args)
        if names:
            raise VBAError(448)
        Element(target, positions).set(rhs, object_assignment)
        return
    if hasattr(type(target), "_vba_set_item"):
        positions, names = _expand_args(args)
        if names:
            raise VBAError(448)
        target._vba_set_item(positions, rhs if object_assignment or not v.is_object(rhs) else v.default_value(rhs))
        return
    result = invoke(target, args)
    put = getattr(type(result), "_vba_assign_default", None)
    if put is None or object_assignment:
        raise VBAError(450)
    put(result, rhs)


class Frame:
    def __init__(self, program, context, procedure=None):
        self.program, self.context, self.procedure = program, context, procedure
        self.spec = context.spec
        self.slots = {}
        self.temps = {}
        self.gosubs = []
        self.error_mode = "off"
        self.handler = None
        self.error_point = None
        self.handler_active = False
        self.line = 0
        self.erl = 0

    def member(self, target, name, invoke=False):
        result = get_member(target, name, invoke=invoke, caller=self.context)
        return result if invoke else evaluate(result)

    def lookup(self, name, *, call=False, writing=False, declare=False):
        key = identifier(name)
        if key in self.slots and not (call and self.procedure and self.procedure["key"] == key):
            return self.slots[key]
        result = self.context.lookup(name, internal=True)
        if result is not v.MISSING and (not writing or isinstance(result, Slot)):
            return result
        if writing:
            if self.spec["explicit"] and not declare:
                raise VBAError(35, f"Variable {name!r} is not defined")
            slot = self.slots[key] = Slot(TYPE_SUFFIXES.get(name[-1], "Variant"))
            slot.value = v.default_for(slot.type_name)
            return slot
        if key == "me" and self.spec["kind"] != "standard":
            return self.context.host if self.context.host is not None else self.context
        if key == "erl":
            return self.erl
        if key == "application":
            return self.program.app
        if key == "err":
            return self.program.err
        if key == "debug":
            return self.program.builtins.debug
        if key == "vba":
            return Library(self)
        if key == "thisworkbook":
            return self.program.this_workbook
        result = self.program.external(key, self.context)
        if result is not v.MISSING:
            return result
        if key in EXCEL_GLOBALS:
            return get_member(self.program.app, key)
        result = self.program.builtins.get(key, self)
        if result is not v.MISSING:
            return result
        if call:
            raise VBAError(35, f"Sub or Function {name!r} is not defined")
        return self.lookup(name, writing=True)

    def get(self, name):
        return evaluate(self.lookup(name))

    def target(self, name):
        return self.lookup(name, call=True)

    def arg_name(self, name):
        result = self.lookup(name)
        if isinstance(result, (Slot, Element)) and not result.constant:
            return Argument(result.get(), result)
        return Argument(evaluate(result))

    def arg_invoke(self, target, arguments):
        original = _python_arg(target)
        if isinstance(original, VBAArray):
            positions, names = _expand_args(arguments)
            if names:
                raise VBAError(448)
            ref = Element(original, positions)
            return Argument(ref.get(), ref)
        return Argument(invoke(original, arguments))

    def arg_member(self, obj, name):
        result = get_member(obj, name, caller=self.context)
        if isinstance(result, (Slot, Element)) and not result.constant:
            return Argument(result.get(), result)
        return Argument(evaluate(result))

    def assign_name(self, name, value, object_assignment=False):
        key = identifier(name)
        if key in self.slots:
            self.slots[key].set(value, object_assignment)
            return
        slot = self.context.fields.get(key)
        if slot is not None:
            slot.set(value, object_assignment)
            return
        kind = "property set" if object_assignment else "property let"
        if (key, kind) in self.spec["procedures"]:
            self.context.set(name, value, object_assignment=object_assignment, internal=True)
            return
        external = self.program.external(key, self.context)
        if isinstance(external, Slot):
            external.set(value, object_assignment)
            return
        if self.context.host is not None:
            try:
                assign_member(self.context.host, name, value, (), object_assignment)
                return
            except VBAError as error:
                if error.number != 438:
                    raise
        self.lookup(name, writing=True).set(value, object_assignment)

    def assign_member(self, obj, name, rhs, args=(), object_assignment=False):
        assign_member(obj, name, rhs, args, object_assignment, self.context)

    def index_target(self, name):
        return self.lookup(name) if identifier(name) in self.slots else self.target(name)

    def replace_mid(self, original, start, length, replacement):
        target, text, offset = v.string(original), v.string(replacement), v.integer(start) - 1
        count = len(text) if length is v.MISSING else v.integer(length)
        if offset < 0 or offset >= len(target) or count < 0:
            raise VBAError(5)
        write = min(count, len(text), len(target) - offset)
        return target[:offset] + text[:write] + target[offset + write:]

    def select_value(self, key):
        if key not in self.temps:
            raise VBAError(5, "Select Case is not initialized")
        return self.temps[key]

    def case_range(self, value, first, last):
        return v.truth(self.binary(">=", value, first)) and v.truth(self.binary("<=", value, last))

    def assign_index(self, target, args, value, object_assignment=False):
        _set_index(target, args, value, object_assignment)

    def redim(self, name, bounds, preserve, type_name):
        key = identifier(name)
        if key in self.slots:
            slot = self.slots[key]
        elif key in self.context.fields:
            slot = self.context.fields[key]
        elif isinstance(result := self.program.external(key, self.context), Slot):
            slot = result
        else:
            slot = self.lookup(name, writing=True, declare=True)
        value = slot.get()
        if isinstance(value, VBAArray):
            if type_name.casefold() == "variant":
                type_name = value.type_name
            if getattr(slot, "array", False) and type_name.casefold() != slot.type_name.casefold():
                raise VBAError(13)
        else:
            if slot.type_name.casefold() != "variant":
                raise VBAError(13)
            value = VBAArray(type_name=type_name)
        slot.store(value.redim(bounds, preserve, type_name))

    def erase(self, values):
        for value in values:
            if not isinstance(value, VBAArray):
                raise VBAError(13)
            value.erase()

    def bounds(self, dimensions):
        result = tuple((self.spec["base"] if lower is v.MISSING else v.integer(lower), v.integer(upper))
                       for lower, upper in dimensions)
        if any(hi < lo for lo, hi in result):
            raise VBAError(9)
        return result

    def result(self):
        if self.procedure["kind"] in {"function", "property get"}:
            return copy_value(self.slots[self.procedure["key"]].get())
        return v.EMPTY

    def binary(self, operator, left, right):
        return v.binary(operator, left, right, self.spec["compare"])

    def type_of(self, obj, type_name):
        if obj is None:
            return False
        if identifier(type_name) == "object":
            return v.is_object(obj)
        return identifier(v.desc(obj)) == identifier(type_name.rsplit(".", 1)[-1])

    def new(self, name):
        return self.program.new(name)

    def with_object(self, value):
        if value is None:
            raise VBAError(91)
        if not (v.is_object(value) or isinstance(value, Record)):
            raise VBAError(424)
        return value

    def start_for(self, key, start, end, step, name):
        numbers = tuple(map(v.numeric, (start, end, step)))
        self.assign_name(name, numbers[0])
        self.temps[key] = numbers[1:]

    def for_test(self, key, name):
        try:
            end, step = self.temps[key]
        except KeyError:
            raise VBAError(92, "For loop not initialized") from None
        value = v.numeric(self.get(name))
        return value <= end if step >= 0 else value >= end

    def for_step(self, key, name):
        try:
            step = self.temps[key][1]
        except KeyError:
            raise VBAError(92, "For loop not initialized") from None
        self.assign_name(name, self.binary("+", self.get(name), step))

    def start_each(self, key, iterable):
        try:
            self.temps[key] = iter(iterable)
        except TypeError:
            raise VBAError(424) from None

    def next_each(self, key, name):
        if key not in self.temps:
            raise VBAError(92, "For Each loop not initialized")
        try:
            element = next(self.temps[key])
        except StopIteration:
            return False
        self.assign_name(name, element, v.is_object(element))
        return True

    def set_error_mode(self, mode, target):
        self.program.err.Clear()
        self.error_mode, self.handler = mode, target
        if mode == "off":
            self.handler_active = False
            self.error_point = None

    def handle_error(self, error, current, following):
        name = self.procedure["name"]
        error.locations.append((self.spec["name"], name, self.line))
        if not error.source:
            error.source = f"{self.spec['name']}.{name}"
        self.program.err.capture(error, error.source)
        if self.handler_active or self.error_mode == "off":
            raise error
        if self.error_mode == "next":
            return following
        self.error_point, self.handler_active = (current, following), True
        return self.handler

    def resume(self, target):
        if self.error_point is None:
            raise VBAError(20)
        current, following = self.error_point
        self.error_point, self.handler_active = None, False
        self.program.err.Clear()
        return current if target == "same" else following if target == "next" else target

    def go_return(self):
        if not self.gosubs:
            raise VBAError(3)
        return self.gosubs.pop()


# Public surface imported by generated Python.
truth, unary, number, date, sentinel = v.truth, v.unary, v.number, v.date, v.sentinel
MISSING, EMPTY = v.MISSING, v.EMPTY


class Program:
    def __init__(self, modules, workbook=None, *, application=None, directory=None, output=None,
                 locale=None):
        self.locale = resolve_locale(locale)
        with use_locale(self.locale):
            self._load(modules, workbook, application, directory, output)

    def _load(self, modules, workbook, application, directory, output):
        self.specs = modules
        self.app = application if application is not None else Application(directory)
        self.err = ErrObject()
        self.builtins = Builtins(self, output)
        self.this_workbook = self.app.Workbooks.Open(str(workbook)) if workbook is not None else None
        self.contexts = {}
        self.documents = {}
        self._initializing = set()
        for key, spec in self.specs.items():
            if spec["kind"] == "class":
                continue
            host = None
            if spec["kind"] == "document" and self.this_workbook is not None:
                book = self.this_workbook
                if key == identifier(book.CodeName) or key == "thisworkbook":
                    host = book
                else:
                    host = next((sheet for sheet in book.Worksheets if sheet.CodeName.casefold() == key), None)
            self.contexts[key] = Context(self, spec, host)
            if host is not None:
                self.documents[id(host)] = self.contexts[key]
        for state in self.contexts.values():
            self._initialize(state)

    def _declare(self, frame, declaration):
        kind = declaration["type"]
        factory = None
        bounds = declaration["bounds"]
        if bounds is not None:
            dimensions = bounds(frame) if bounds else None
            value = VBAArray(dimensions, kind, fixed=bool(bounds), fixed_length=declaration["length"])
        elif declaration["new"]:
            factory = lambda: self.new(kind)
            value = None
        elif self._record_spec(kind, frame.context):
            value = self._record(kind, frame)
        else:
            value = v.default_for(kind, declaration["length"])
        return Slot(kind, value, length=declaration["length"], factory=factory, array=bounds is not None)

    def _initialize(self, ctx):
        if ctx.initialized:
            return
        ctx.initialized = True
        for key in ctx.spec["declarations"]:
            self._initialize_field(ctx, key)

    def _initialize_field(self, ctx, key):
        if key in ctx.fields:
            return ctx.fields[key]
        if key in ctx.initializing:
            raise VBAError(5, f"Circular declaration {ctx.spec['name']}.{key}")
        ctx.initializing.add(key)
        try:
            frame = Frame(self, ctx)
            declaration = ctx.spec["declarations"][key]
            slot = self._declare(frame, declaration)
            self._initialize_slot(slot, frame, declaration)
            ctx.fields[key] = slot
            return slot
        finally:
            ctx.initializing.remove(key)

    @staticmethod
    def _initialize_slot(slot, frame, declaration):
        if declaration["initial"] is not None:
            slot.store(declaration["initial"](frame))
        slot.constant = declaration["constant"]

    def _record_spec(self, name, context):
        key = identifier(name)
        local = context.spec["types"].get(key)
        return local if local is not None else next((spec["types"][key] for spec in self.specs.values()
                                                    if key in spec["types"]), None)

    def _record(self, name, frame):
        key = identifier(name)
        if key in self._initializing:
            raise VBAError(13, f"Recursive Type {name!r}")
        self._initializing.add(key)
        try:
            return Record(name, {field: self._declare(frame, decl)
                                 for field, decl in self._record_spec(name, frame.context).items()})
        finally:
            self._initializing.remove(key)

    def external(self, key, current):
        if key in self.contexts:
            context = self.contexts[key]
            return context.host if context.host is not None else context
        matches = []
        for ctx in self.contexts.values():
            if ctx is current or ctx.spec["kind"] != "standard":
                continue
            value = ctx.lookup(key)
            if value is not v.MISSING:
                matches.append(value)
        if len(matches) > 1:
            raise VBAError(35, f"Ambiguous public name {key!r}")
        return matches[0] if matches else v.MISSING

    def new(self, name):
        key = identifier(name)
        if key in {"collection", "vba.collection"}:
            return Collection()
        if key in {"dictionary", "scripting.dictionary"}:
            return Dictionary()
        spec = self.specs.get(key)
        if spec is None or spec["kind"] != "class":
            raise VBAError(429, f"Unsupported class {name!r}")
        ctx = Context(self, spec)
        self._initialize(ctx)
        initializer = ctx.lookup("Class_Initialize", internal=True)
        if isinstance(initializer, BoundProcedure):
            initializer(())
        return ctx

    def _invoke(self, context, key, arguments):
        procedure = context.spec["procedures"][key]
        frame = Frame(self, context, procedure)
        self._bind(frame, procedure["parameters"], arguments)
        if procedure["kind"] in {"function", "property get"}:
            kind = procedure["type"]
            frame.slots[key[0]] = Slot(kind, v.default_for(kind))
        statics = context.statics.setdefault(key, {})
        for name, declaration in procedure["locals"].items():
            persistent = procedure["static"] or declaration["static"]
            if persistent and name in statics:
                frame.slots[name] = statics[name]
                continue
            slot = frame.slots[name] = self._declare(frame, declaration)
            self._initialize_slot(slot, frame, declaration)
            if persistent:
                statics[name] = slot
        return procedure["body"](frame)

    def _bind(self, frame, parameters, arguments):
        named = {}
        positioned = []
        for name, arg in arguments:
            if name is None:
                if named:
                    raise VBAError(448)
                positioned.append(arg)
            else:
                key = identifier(name)
                if key in named:
                    raise VBAError(448)
                named[key] = arg
        index = 0
        for p in parameters:
            key = p["key"]
            if p["paramarray"]:
                frame.slots[key] = Slot("Variant", VBAArray.from_values(a.value for a in positioned[index:]), array=True)
                index = len(positioned)
                continue
            if index < len(positioned):
                arg = positioned[index]
                index += 1
                if key in named:
                    raise VBAError(448)
            else:
                arg = named.pop(key, Argument(v.MISSING))
            kind = p["type"]
            if arg.value is v.MISSING:
                if not p["optional"]:
                    raise VBAError(449, f"Argument {key!r} is not optional")
                value = p["default"](frame) if p["default"] else (
                    v.MISSING if kind.casefold() == "variant" else v.default_for(kind)
                )
            elif p["byref"] and arg.reference is not None:
                if kind.casefold() != "variant" and arg.reference.type_name.casefold() != kind.casefold():
                    raise VBAError(13, f"ByRef argument {key!r} has the wrong type")
                if p["array"] and not isinstance(arg.value, VBAArray):
                    raise VBAError(13)
                frame.slots[key] = arg.reference
                continue
            else:
                value = arg.value
                if p["array"] and not isinstance(value, VBAArray):
                    raise VBAError(13)
                value = copy_value(value if p["array"] else v.coerce(value, kind, p["length"]))
            frame.slots[key] = Slot(kind, value, length=p["length"], array=p["array"])
        if named:
            raise VBAError(448, f"Named argument {next(iter(named))!r} not found")
        if index < len(positioned):
            raise VBAError(450)

    def run(self, entrypoint, *args, **kwargs):
        parts = entrypoint.rsplit(".", 1)
        name = identifier(parts[-1])
        candidates = self.contexts.values() if len(parts) == 1 else (
            (ctx,) if (ctx := self.contexts.get(identifier(parts[0]))) is not None else ()
        )
        entries = []
        for context in candidates:
            for kind in ("sub", "function", "property get"):
                if (name, kind) in context.spec["procedures"]:
                    entries.append((context, (name, kind)))
        if len(entries) != 1:
            raise VBAError(35, f"{'Ambiguous' if entries else 'Unknown'} entrypoint {entrypoint!r}")
        arguments = tuple((None, Argument(arg)) for arg in args) + tuple(
            (key, Argument(value)) for key, value in kwargs.items())
        self.err.Clear()
        with use_locale(self.locale):
            try:
                return self._invoke(*entries[0], arguments)
            except RecursionError:
                raise VBAError(28) from None
            except EndProgram:
                self.err.Clear()
                for ctx in self.contexts.values():
                    ctx.fields.clear()
                    ctx.statics.clear()
                    ctx.initialized = False
                for ctx in self.contexts.values():
                    self._initialize(ctx)
                self.app.files.close()
                return v.EMPTY

    def close(self):
        self.app.Quit()
