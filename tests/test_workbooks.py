import hashlib
import io
import os
import subprocess
import sys
from pathlib import Path
from zipfile import ZipFile

import pytest
from openpyxl import Workbook, load_workbook

from vba2py import ExtractionError, SourceModule, compile_modules, compile_source, compile_workbook, extract_modules

FIXTURE = Path(__file__).parent / "fixtures" / "sample.xlsm"


def test_real_xlsm_extraction_and_generated_python(tmp_path):
    modules = extract_modules(FIXTURE)
    assert ("Module1", "standard") in {(m.name, m.kind) for m in modules}
    assert ("ThisWorkbook", "document") in {(m.name, m.kind) for m in modules}
    assert any("Sub say_hello" in m.source for m in modules)
    result = compile_workbook(FIXTURE)
    output = io.StringIO()
    program = result.create_program(FIXTURE, output=output)
    try:
        program.run("Module1.say_hello")
    finally:
        program.close()
    assert "Hello from Python!" in output.getvalue()

    script = tmp_path / "generated.py"
    script.write_text(result.python, encoding="utf-8")
    root = str(Path(__file__).resolve().parents[1])
    environment = dict(os.environ, PYTHONPATH=os.pathsep.join([root, os.environ.get("PYTHONPATH", "")]))
    result_cli = subprocess.run([sys.executable, str(script), "say_hello", "--workbook", str(FIXTURE)],
                                check=True, capture_output=True, text=True, env=environment)
    assert result_cli.stdout == "Hello from Python!\n"


def test_open_edit_close_and_roundtrip_macros_and_formula_text(tmp_path):
    source = tmp_path / "secondary.xlsx"
    output = tmp_path / "result.xlsm"
    initial = Workbook()
    initial.active.title = "Lookup"
    initial.active["B2"] = 10
    initial.save(source)
    initial.close()
    compilation = compile_source("""
Option Explicit
Sub ImportValues()
    Dim w As Workbook, data As Variant
    Set w = Workbooks.Open(Filename:="secondary.xlsx", ReadOnly:=True)
    data = w.Worksheets("Lookup").Range("B2")
    With ThisWorkbook.Worksheets("data")
        .Range("A2").Value2 = data + 1
        .Cells(3, "B") = 12
        .Range("C3").Value = .Range("A1").Value
        .Range("A1").Copy .Range("D1")
    End With
    w.Close SaveChanges:=False
End Sub
""")
    p = compilation.create_program(FIXTURE, directory=tmp_path)
    try:
        p.run("ImportValues")
        assert p.app.Workbooks.Count == 1
        p.this_workbook.SaveAs(str(output))
    finally:
        p.close()
    with ZipFile(output) as saved, ZipFile(FIXTURE) as original:
        macro = "xl/vbaProject.bin"
        assert hashlib.sha256(saved.read(macro)).digest() == hashlib.sha256(original.read(macro)).digest()
    result = load_workbook(output)
    try:
        sheet = result["Data"]
        assert [sheet[key].value for key in ["A1", "A2", "B3", "C3", "D1"]] == [
            "=1+2", 11, 12, 3, "=1+2",
        ]
    finally:
        result.close()
    secondary = load_workbook(source)
    try:
        assert secondary["Lookup"]["B2"].value == 10
    finally:
        secondary.close()


def test_sheet_document_module_resolves_me_and_code_name(tmp_path):
    p = compile_modules([SourceModule("Sheet1", """
Option Explicit
Sub Paint()
    Me.Cells(2, 2).Value = 4
    Range("B3").Value = Sheet1.Range("B2") + 3
End Sub
""", "document")]).create_program(FIXTURE, directory=tmp_path)
    try:
        p.run("Sheet1.Paint")
        assert p.this_workbook.Worksheets.Item(1).Range("B3").Value == 7
    finally:
        p.close()


def test_file_channels_and_relative_paths(tmp_path):
    p = compile_source("""
Option Explicit
Function SaveRead() As String
    Dim n As Integer, line As String, item As String, count As Long
    n = FreeFile
    Open "text.txt" For Output As #n
    Print #n, "hello"
    Write #n, "line,value", 42
    Close #n
    Name "text.txt" As "renamed.txt"
    Open "renamed.txt" For Input As #n
    Line Input #n, line
    Input #n, item, count
    Close #n
    SaveRead = line & ":" & item & ":" & CStr(count)
End Function
""").create_program(directory=tmp_path)
    try:
        assert p.run("SaveRead") == "hello:line,value:42"
        assert (tmp_path / "renamed.txt").read_bytes() == b'hello\r\n"line,value",42\r\n'
    finally:
        p.close()


def test_invalid_macro_workbook_has_actionable_error(tmp_path):
    bad = tmp_path / "bad.xlsm"
    bad.write_bytes(b"invalid")
    with pytest.raises(ExtractionError, match="Could not extract VBA"):
        extract_modules(bad)


def test_excel_font_fill_error_values_and_range_identity(tmp_path):
    p = compile_modules([
        SourceModule("Sheet1", """
Public Function Caption() As String
    Caption = Me.Name
End Function
""", "document"),
        SourceModule("Entry", """
Option Explicit
Function Style() As String
    Dim r As Range
    With Sheet1.Range("B2")
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(20, 40, 80)
        .Value = CVErr(xlErrNA)
    End With
    Set r = Sheet1.Range("B3")
    r = "same workbook"
    If Sheet1 Is ThisWorkbook.Sheets(1) Then Style = Sheet1.Caption()
End Function
"""),
    ]).create_program(FIXTURE, directory=tmp_path)
    try:
        assert p.run("Style") == "Data"
        out = tmp_path / "style.xlsx"
        p.this_workbook.SaveAs(str(out))
    finally:
        p.close()
    book = load_workbook(out)
    try:
        cell = book.active["B2"]
        assert cell.value == "#N/A" and cell.data_type == "e"
        assert cell.font.bold and cell.font.sz == 14
        assert cell.fill.fgColor.rgb == "FF142850"
        assert book.active["B3"].value == "same workbook"
    finally:
        book.close()
