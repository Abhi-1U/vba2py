from io import StringIO
from textwrap import dedent

import pytest

from vba2py import ProjectError, SourceModule, VBAError, compile_modules, compile_source
from vba2py.runtime.values import EMPTY, NULL


def program(source, **options):
    return compile_source(dedent(source)).create_program(**options)


def test_arithmetic_options_and_variant():
    p = program("""
        Option Explicit
        Option Compare Text
        Function Check() As String
          Dim v, n As Long
          v = Empty
          n = -2^2 + 14 \\ 3 Mod 3 * 1 + CInt(2.5)
          Check = CStr(n) & ":" & CStr(v + 3) & ":" & CStr("Ab2" Like "a?#")
        End Function
        Function Maybe(Optional x)
          If IsMissing(x) Then Maybe = Null Else Maybe = x
        End Function
        Function Nulls()
          Nulls = ((Null And False) = False) And ((Null Or True) = True)
        End Function
    """)
    assert p.run("check") == "-1:3:True"
    assert p.run("maybe") is NULL
    assert p.run("maybe", EMPTY) is EMPTY
    assert p.run("nulls") is True


def test_byref_grouping_named_arguments_and_array_element():
    p = program("""
        Option Explicit
        Sub Bump(ByRef target As Long, Optional ByVal stepAmount As Long = 3)
            target = target + stepAmount
        End Sub
        Function Compute() As Long
            Dim x As Long, a(1 To 2) As Long
            x = 5
            Bump x
            Bump (x)
            Call Bump(stepAmount:=2, target:=x)
            Call Bump(a(2), 10)
            Compute = a(2) + x
        End Function
    """)
    assert p.run("compute") == 20


def test_option_base_preserve_and_arrays_copy_in_vba_order():
    p = program("""
        Option Base 1
        Option Explicit
        Function Build() As String
            Dim x() As Long, y, z, v
            ReDim x(-1 To 0, 2 To 2)
            x(-1, 2) = 7
            x(0, 2) = 9
            ReDim Preserve x(-1 To 0, 2 To 3)
            x(-1, 3) = 11
            y = x
            x(0, 2) = 44
            For Each v In y
                Build = Build & CStr(v) & ","
            Next
            z = Array(1, 2)
            Build = Build & CStr(LBound(z)) & ":" & CStr(LBound(VBA.Array(1)))
            Erase x
        End Function
    """)
    assert p.run("Build") == "7,9,11,0,1:0"


def test_recursive_calls_static_state_and_qualified_modules():
    a = SourceModule("Accounting", dedent("""
        Option Explicit
        Public Counter As Long
        Function Factorial(ByVal n As Long) As Long
            If n <= 1 Then Factorial = 1: Exit Function
            Factorial = n * Factorial(n - 1)
        End Function
    """))
    b = SourceModule("Entry", dedent("""
        Option Explicit
        Static Function RunIt(ParamArray values() As Variant) As Long
            Dim seen As Long
            Counter = Counter + 1
            seen = seen + UBound(values) + 1
            RunIt = Accounting.Factorial(4) + seen + Counter
        End Function
    """))
    c = compile_modules([a, b])
    p = c.create_program()
    assert p.run("Entry.RunIt", 5, 6) == 27
    assert p.run("runit", 7) == 29
    assert c.create_program().run("RunIt") == 25


def test_control_flow_labels_loops_and_error_handlers():
    p = program("""
        Option Explicit
        Function RunIt() As Long
            Dim i As Long, total As Long, errorLine As Long
            On Error GoTo handler
            For i = 5 To 0 Step -1
                Select Case i
                    Case 2 To 3, Is > 4
                        total = total + i
                    Case 4
                        GoSub addTen
                    Case Else
                        Exit For
                End Select
            Next i
            Do
                total = total + 1
            Loop Until total > 20
        110 total = total / 0
            RunIt = total + errorLine
            Exit Function
        addTen:
            total = total + 10
            Return
        handler:
            errorLine = Erl + Err.Number
            Resume Next
        End Function
        Function Caller() As Long
            On Error Resume Next
            Fails
            Caller = Err.Number
            On Error GoTo 0
        End Function
        Sub Fails()
            Err.Raise 207, "custom", "bad"
        End Sub
    """)
    assert p.run("RunIt") == 142
    assert p.run("Caller") == 207
    assert p.err.Number == 0


def test_classes_properties_collections_and_type():
    declarations = SourceModule("Counter", dedent("""
        Option Explicit
        Private n As Long
        Private Sub Class_Initialize()
            n = 1
        End Sub
        Public Property Get Value() As Long
            Value = n
        End Property
        Public Property Let Value(ByVal x As Long)
            n = x
        End Property
    """), "class")
    entry = SourceModule("Entry", dedent("""
        Option Explicit
        Private Type Person
            Name As String * 5
            Id As Long
        End Type
        Function RunIt() As String
            Dim p As Person, c As New Counter, stuff As New Collection, p2 As Person
            p.Name = "Sam"
            p.Id = 2
            p2 = p
            With c
                .Value = .Value + p.Id
            End With
            stuff.Add c.Value, Key:="sum"
            If TypeOf c Is Counter Then RunIt = p2.Name & CStr(stuff("SUM"))
        End Function
    """))
    assert compile_modules([declarations, entry]).create_program().run("RunIt") == "Sam  3"


def test_unhandled_errors_report_vba_locations():
    p = program("""
        Sub Outer()
            Inner
        End Sub
        Sub Inner()
            Err.Raise 200, Description:="failure"
        End Sub
    """)
    with pytest.raises(VBAError, match=r"VBA error 200: failure") as failure:
        p.run("Outer")
    assert failure.value.locations == [("Module1", "Inner", 6), ("Module1", "Outer", 3)]


@pytest.mark.parametrize(("line", "message"), [
    ("v = 1", "not declared"),
    ("GoTo missing", "Undefined label"),
    ("Exit For", "outside the matching loop"),
    ("With x\nEnd With", "not declared"),
])
def test_compile_diagnostics_include_source_location(line, message):
    with pytest.raises(ProjectError, match=f"Module1:3: .*{message}"):
        compile_source(f"Option Explicit\nSub Demo()\n{line}\nEnd Sub")


def test_conditional_compilation_continuations_and_comments():
    p = program("""
        #Const threshold = 4
        #If Win64 And threshold >= 4 Then
        Declare PtrSafe Function Native Lib "example" () As Long
        #Else
        Function RunIt() As String
            RunIt = "it's:" & _
                    "A:B"  ' ignored
            Rem ignored
        End Function
        #End If
    """)
    assert p.run("RunIt") == "it's:A:B"


def test_case_insensitive_procedure_and_debug_output():
    output = StringIO()
    p = program("""
        Sub Show()
            Debug.Print "Hi", True
        End Sub
    """, output=output)
    p.run("sHoW")
    assert output.getvalue() == "Hi True\n"


def test_single_line_nested_if_mid_assignment_and_raised_error():
    p = program("""
        Option Explicit
        Function UseIf(ByVal outer As Boolean, ByVal inner As Boolean) As String
            Dim s As String
            s = "123"
            If outer Then If inner Then Mid(s, 2, 1) = "A" Else Mid(s, 2, 1) = "B" Else Mid(s, 2, 1) = "C"
            UseIf = s
        End Function
        Function ErrorValue() As Boolean
            Dim e
            e = CVErr(xlErrNA)
            ErrorValue = IsError(e) And (VarType(e) = vbError)
        End Function
    """)
    assert p.run("UseIf", False, True) == "1C3"
    assert p.run("UseIf", True, False) == "1B3"
    assert p.run("UseIf", True, True) == "1A3"
    assert p.run("ErrorValue") is True


def test_extracted_class_default_property():
    cls = SourceModule("Reading", dedent("""
        Private actual As Long
        Property Get Value() As Long
        Attribute Value.VB_UserMemId = 0
            Value = actual
        End Property
        Property Let Value(ByVal valueIn As Long)
            actual = valueIn
        End Property
    """), "class")
    entry = SourceModule("Entry", dedent("""
        Option Explicit
        Function GetReading() As Long
            Dim item As New Reading
            item = 9
            GetReading = item + 1
        End Function
    """))
    assert compile_modules([cls, entry]).create_program().run("GetReading") == 10
