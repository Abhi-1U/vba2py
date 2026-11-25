"""Runtime errors as exposed through VBA's Err object."""

_MESSAGES = {
    3: "Return without GoSub", 5: "Invalid procedure call or argument", 6: "Overflow",
    7: "Out of memory", 9: "Subscript out of range", 10: "This array is fixed or temporarily locked",
    11: "Division by zero", 13: "Type mismatch", 20: "Resume without error",
    28: "Out of stack space", 35: "Sub or Function not defined", 52: "Bad file name or number",
    53: "File not found", 54: "Bad file mode", 55: "File already open", 57: "Device I/O error",
    58: "File already exists", 59: "Bad record length", 62: "Input past end of file",
    67: "Too many files", 70: "Permission denied", 75: "Path/File access error",
    76: "Path not found", 91: "Object variable or With block variable not set",
    94: "Invalid use of Null", 424: "Object required", 429: "ActiveX component can't create object",
    438: "Object doesn't support this property or method",
    448: "Named argument not found", 449: "Argument not optional",
    450: "Wrong number of arguments or invalid property assignment",
    457: "This key is already associated with an element of this collection",
    1004: "Application-defined or object-defined error",
}


class VBAError(Exception):
    def __init__(self, number, description=None, source=None):
        self.number = int(number)
        self.description = _MESSAGES.get(self.number, "Application-defined or object-defined error") if description is None else str(description)
        self.source = str(source) if source else ""
        self.locations = []
        super().__init__(self.description)

    def __str__(self):
        where = "\n".join(f"  at {module}.{procedure} (line {line})"
                          for module, procedure, line in self.locations)
        return f"VBA error {self.number}: {self.description}" + (f"\n{where}" if where else "")


class EndProgram(BaseException):
    """An End statement unwinds the entire VBA project."""


def conversion_error(error):
    if isinstance(error, OverflowError):
        return VBAError(6)
    if isinstance(error, ZeroDivisionError):
        return VBAError(11)
    if isinstance(error, TypeError):
        return VBAError(13, str(error))
    return VBAError(5, str(error))


class ErrObject:
    def __init__(self):
        self.Clear()

    def _vba_default_value(self):
        return self.Number

    def Clear(self):
        self.Number = 0
        self.Description = ""
        self.Source = ""
        self.HelpFile = ""
        self.HelpContext = 0

    def capture(self, error, source):
        self.Number = error.number
        self.Description = error.description
        self.Source = error.source or source

    def Raise(self, Number, Source=None, Description=None, HelpFile=None, HelpContext=None):
        from .values import integer, string
        number = integer(Number)
        raise VBAError(number, string(Description) if Description is not None else None,
                       string(Source) if Source is not None else None)
