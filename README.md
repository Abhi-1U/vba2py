# vba2py

Transpile Excel VBA projects into runnable Python. The compiler reads the workbook with openpyxl, decodes its embedded VBA streams with oletools, parses expressions with Lark, and emits Python backed by a VBA runtime. Excel and Windows are optional for the supported workflows.

## Install

Requires Python 3.10 or later.

```sh
python -m pip install .
```

## Convert and run

```sh
vba2py extract Report.xlsm -o extracted
vba2py compile Report.xlsm -o report.py
python report.py Module1.Main --workbook Report.xlsm --save Result.xlsm
```

The CLI can also compile and run without writing a script.

```sh
vba2py run Report.xlsm Module1.Main --save Result.xlsm
vba2py run Report.xlsm Module1.Total --arg 10 --named 'currency="INR"'
```

You can compile edited or exported source modules with `vba2py compile Module1.bas Counter.cls -o macros.py`. Use `-D Win64=true` to select a conditional compilation target. The default target defines `VBA7` and leaves `Win16`, `Win32`, `Win64`, and `Mac` false.

The workbook supplied at run time becomes `ThisWorkbook`. Relative paths used by macros resolve against the working directory, which defaults to the current directory and can be set with `--directory`. Workbook changes are persisted when the macro explicitly saves them or when you pass `--save`. A macro can open, edit, and save additional workbooks with the same file access API. Use the execution commands with macros you trust, since a macro can read, write, rename, or delete files accessible to the Python process.

## Locale and formatting

The default runtime locale is `en-US`. Use `--locale` on either execution command, or set `locale` when creating a program from Python. Locale settings apply to that program and do not change the machine's regional settings.

```sh
vba2py run Report.xlsm Module1.Main --locale de-DE --save Result.xlsm
python report.py Module1.Main --workbook Report.xlsm --locale de-DE
```

```python
from vba2py import compile_workbook

program = compile_workbook("Report.xlsm").create_program("Report.xlsm", locale="de-DE")
try:
    program.run("Module1.Main")
finally:
    program.close()
```

In `de-DE`, `CDbl("1.234,50")` reads a grouped decimal, `CDate("04/05/2025")` reads 4 May, and `FormatNumber(1234.5, 2)` returns `"1.234,50"`. Typed assignments and argument conversions follow the same locale rules. `CDec` and `CCur` retain decimal precision through conversion, and integer conversions use round-half-to-even. Source literals, `Val`, `Str`, and numeric output from `Write #` use the VBA source or portable conventions. [Microsoft VBA string coercion rules](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/00113388-401b-41c2-8107-dc3fc0485554)

`Format` supports common named styles and custom numeric, date/time, and string masks. `FormatNumber`, `FormatCurrency`, `FormatPercent`, and `FormatDateTime` use regional defaults, including grouping widths, decimal separators, and currency symbols. The library also provides `InStrRev`, `Filter`, `StrConv` casing, `Switch`, `MonthName`, `WeekdayName`, `DatePart`, `FV`, `PV`, and `NPV`. [Microsoft VBA Format reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/format-function-visual-basic-for-applications)

Regional defaults come from Babel's CLDR data. Windows custom regional settings, the Hijri calendar, some mixed custom masks, and `StrConv` width, kana, and code-page conversions remain outside this implementation.

## Python API

```python
from vba2py import compile_source

source = """
Option Explicit

Function SumSquares(ByVal count As Long) As Long
    Dim i As Long
    Dim squares() As Long
    ReDim squares(1 To count)
    For i = 1 To count
        squares(i) = i * i
        SumSquares = SumSquares + squares(i)
    Next i
End Function
"""

compiled = compile_source(source)
program = compiled.create_program()
try:
    print(program.run("SumSquares", 10))  # 385
finally:
    program.close()
```

`compile_workbook(path)` returns the same `Compilation` object as `compile_source(source)`. Read `compilation.python` to save or inspect the emitted script, or call `compilation.create_program(path)` to run it in process. Pass a sequence of `SourceModule(name, source, kind)` to `compile_modules` to assemble a project with standard, class, and workbook or worksheet document modules.

## Compatibility

- **Language.** Sub, Function, Property Get/Let/Set, module and procedure scope, Static variables, ByVal and ByRef, optional and named arguments, ParamArray, conditional compilation, user-defined types, enums, class initialization, and default class properties. Control flow includes If, Select Case, For, For Each, Do, While, With, GoTo, GoSub, On Error, Resume, and Err.
- **Values.** Variant subtypes including Empty, Null, Missing, and error values, scalar coercion, fixed-length strings, and VBA arithmetic, comparison, and bitwise operators. Arrays support explicit and negative bounds, multiple dimensions, Option Base, ReDim Preserve, and Erase. The library includes Collection, an in-memory Scripting.Dictionary, and a subset of VBA functions.
- **Excel.** Workbooks.Open/Add/Save/SaveAs/Close, worksheet lookup and creation, and common Range and Cells operations. The runtime supports rectangular data, range navigation, named range lookup, selection, copy and paste, basic cell formatting, and a small set of worksheet functions.
- **Files.** Numbered text files with Open, Close, Print, Write, Input, and Line Input, plus directory and common file operations. Binary channels support byte-oriented positioning and string reads. Random-access records and the VBA Put/Get statements are outside the implemented subset.

Existing formula text is preserved and copied verbatim. Explicit assignments to `Range.Formula` are supported. `Range.Value` can read an existing formula result from the input workbook's cache when that result is present. The runtime does not recalculate formulas, and openpyxl can remove cached results when saving. Saving as `.xlsm` retains the VBA project from the input workbook. Workbook features outside openpyxl's coverage are subject to its round-trip limitations. [openpyxl workbook loading and saving](https://openpyxl.readthedocs.io/en/stable/tutorial.html#loading-from-a-file)

This is a compatibility implementation for a defined subset of VBA and Excel. COM automation, Windows API declarations, UserForms, and automatic Excel event dispatch are outside that subset. You can invoke a compiled document procedure such as `ThisWorkbook.Workbook_Open` explicitly. The runtime also has partial coverage of Excel properties, VBA library functions, and locale-dependent behavior. Compile errors identify the source module and line. Late-bound calls to unimplemented object members produce VBA runtime errors.

The compiler uses an intermediate set of control-flow instructions to generate Python for each procedure. The emitted functions use explicit jump targets to preserve labels and statement-level error handling. The generated script uses the installed `vba2py` package for the runtime. [Microsoft VBA language specification](https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/d5418146-0bd2-45eb-9c7a-fd9502722c74) · [Lark grammar reference](https://lark-parser.readthedocs.io/en/stable/grammar.html) · [oletools VBA extraction API](https://github.com/decalage2/oletools/blob/master/oletools/doc/olevba.md)

## Tests

```sh
python -m pip install '.[test]'
python -m pytest
```

The suite includes an actual macro-enabled workbook and covers VBA extraction, execution of emitted Python, workbook I/O, and preservation of embedded macros and formula text. The embedded sample VBA project comes from the XlsxWriter example and is distributed under the [BSD 2-Clause license](tests/fixtures/XlsxWriter-LICENSE.txt). [XlsxWriter macro example](https://xlsxwriter.readthedocs.io/example_macros.html)
