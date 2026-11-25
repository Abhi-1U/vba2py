from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from textwrap import dedent

import pytest

from vba2py import VBAError, compile_source
from vba2py.cli import run_generated
from vba2py.runtime.locale import current_locale


@pytest.mark.parametrize(("locale", "price", "quantity", "half", "expected"), [
    ("en-US", "$1,234.50", " 1,001 ", "2.5", "1234.5|1001|2|1235.5"),
    ("de-DE", "1.234,50 €", "1.001", "2,5", "1234,5|1001|2|1235,5"),
    ("fr_FR", "1 234,50 €", "1\u202f001", "2,5", "1234,5|1001|2|1235,5"),
    ("en-IN", "₹12,34,567.50", "1,001", "2.5", "1234567.5|1001|2|1234568.5"),
])
def test_coercion_uses_locale_in_assignments_arguments_and_arithmetic(locale, price, quantity, half, expected):
    source = dedent("""
        Option Explicit
        Function ReadTotals(ByVal priceText As String, ByVal quantity As Long, _
                            ByVal halfText As String) As String
            Dim price As Currency
            price = priceText
            ReadTotals = CStr(price) & "|" & CStr(quantity) & "|" & _
                         CStr(CInt(halfText)) & "|" & CStr(priceText + 1)
        End Function
    """)
    assert compile_source(source).create_program(locale=locale).run("ReadTotals", price, quantity, half) == expected


@pytest.mark.parametrize(("locale", "text", "expected"), [
    ("en-US", "04/05/2025 12:30 PM", datetime(2025, 4, 5, 12, 30)),
    ("en-GB", "04/05/2025 00:30", datetime(2025, 5, 4, 0, 30)),
    ("de-DE", "04.05.2025 12:30", datetime(2025, 5, 4, 12, 30)),
    ("fr-FR", "4 mai 2025 12:30", datetime(2025, 5, 4, 12, 30)),
    ("de-DE", "2025-04-05T12:30:00", datetime(2025, 4, 5, 12, 30)),
    ("de-DE", "13:45", datetime(1899, 12, 30, 13, 45)),
    ("de-DE", "0,5", datetime(1899, 12, 30, 12)),
])
def test_date_text_locale_and_invariant_literals(locale, text, expected):
    p = compile_source(dedent("""
        Option Explicit
        Function ReadDate(ByVal Stamp As String) As Date
            ReadDate = CVDate(Stamp)
        End Function
        Function DateOnly(ByVal Stamp As String) As Date
            DateOnly = DateValue(Stamp)
        End Function
        Function ReadTime(ByVal Stamp As String) As Date
            ReadTime = TimeValue(Stamp)
        End Function
        Function SourceLiteral() As Date
            SourceLiteral = #April 5, 2025 12:30 PM#: Exit Function
        End Function
        Function USOrder() As Date
            USOrder = #04/05/2025#
        End Function
    """)).create_program(locale=locale)
    assert p.run("ReadDate", text) == expected
    assert p.run("DateOnly", text) == expected.replace(hour=0, minute=0, second=0)
    assert p.run("ReadTime", text) == datetime.combine(datetime(1899, 12, 30).date(), expected.time())
    assert p.run("SourceLiteral") == datetime(2025, 4, 5, 12, 30)
    assert p.run("USOrder") == datetime(2025, 4, 5)


def test_full_precision_decimal_and_integer_coercion():
    p = compile_source(dedent("""
        Option Explicit
        Function AsDecimal(ByVal s As String)
            AsDecimal = CDec(s)
        End Function
        Function AsCurrency(ByVal s As String) As Currency
            AsCurrency = CCur(s)
        End Function
        Function BankersRounding(ByVal s As String) As Integer
            BankersRounding = CInt(s)
        End Function
        Function AsSingle(ByVal s As String) As Single
            AsSingle = CSng(s)
        End Function
        Function ConvertBool(ByVal s As String) As Boolean
            ConvertBool = CBool(s)
        End Function
    """)).create_program(locale="de-DE")
    maximum = "79228162514264337593543950335"
    assert p.run("AsDecimal", maximum) == Decimal(maximum)
    assert p.run("AsDecimal", "0,1234567890123456789012345678") == Decimal("0.1234567890123456789012345678")
    assert p.run("AsCurrency", "922.337.203.685.477,5807 €") == Decimal("922337203685477.5807")
    assert p.run("AsCurrency", "1,00005") == Decimal("1.0000")
    assert p.run("AsCurrency", "1,00015") == Decimal("1.0002")
    assert p.run("BankersRounding", "2,500000000000000000001") == 3
    assert p.run("BankersRounding", "2,5") == 2
    assert p.run("BankersRounding", "-3,5") == -4
    assert p.run("ConvertBool", "#TRUE#") is True
    assert p.run("ConvertBool", "#FALSE#") is False
    for entry, argument in [
        ("AsDecimal", "79228162514264337593543950336"),
        ("AsSingle", "1,2E+39"),
        ("AsCurrency", "922.337.203.685.477,5808"),
        ("BankersRounding", "32767,5"),
    ]:
        with pytest.raises(VBAError) as error:
            p.run(entry, argument)
        assert error.value.number == 6


def test_str_val_and_numeric_literals_keep_a_decimal_point():
    p = compile_source(dedent("""
        Option Explicit
        Function Checks() As String
            Checks = CStr(1.25) & "|" & Str(1.25) & "|" & CStr(Val("1,25")) & "|" & _
                     CStr(Val("1.25")) & "|" & CStr(CDbl("1,25 D - 2")) & "|" & _
                     CStr(IsNumeric("12.34,50"))
        End Function
        Function InvalidSuffix()
            InvalidSuffix = Val("50.5%")
        End Function
    """)).create_program(locale="de-DE")
    assert p.run("Checks") == "1,25| 1.25|1|1,25|0,0125|False"
    with pytest.raises(VBAError) as error:
        p.run("InvalidSuffix")
    assert error.value.number == 13


def test_programs_can_use_separate_locales_even_after_a_runtime_error():
    compiled = compile_source(dedent("""
        Function Price() As String
            Price = CStr(1234.5)
        End Function
        Function BadDate() As Date
            BadDate = CDate("31/02/2025")
        End Function
    """))
    german = compiled.create_program(locale="de-DE")
    english = compiled.create_program()
    with pytest.raises(VBAError) as error:
        german.run("BadDate")
    assert error.value.number == 13
    assert str(current_locale()) == "en_US"
    with ThreadPoolExecutor(max_workers=2) as executor:
        results = list(executor.map(lambda project: project.run("Price"), [german, english] * 20))
    assert results == ["1234,5", "1234.5"] * 20


def test_generated_cli_accepts_locale_and_reports_invalid_names(capsys):
    compiled = compile_source("Function Price() As String\nPrice = CStr(1234.5)\nEnd Function")
    assert run_generated(compiled.create_program, ["Price", "--locale", "de-DE"]) == 0
    assert capsys.readouterr().out == "1234,5\n"
    with pytest.raises(SystemExit) as error:
        run_generated(compiled.create_program, ["Price", "--locale", "zz-ZZ"])
    assert error.value.code == 2
    assert "Unknown VBA locale 'zz-ZZ'" in capsys.readouterr().err


def test_file_print_is_localized_while_write_keeps_portable_numbers(tmp_path: Path):
    p = compile_source(dedent("""
        Option Explicit
        Function FileOutput() As Double
            Dim n As Integer
            n = FreeFile
            Open "localized.txt" For Output As #n
            Print #n, 12.5
            Close #n
            Open "portable.txt" For Output As #n
            Write #n, 12.5, CCur(12.5)
            Close #n
            Open "portable.txt" For Input As #n
            Dim a As Double, b As Currency
            Input #n, a, b
            Close #n
            FileOutput = a + b
        End Function
    """)).create_program(locale="de-DE", directory=tmp_path)
    try:
        assert p.run("FileOutput") == 25
    finally:
        p.close()
    assert (tmp_path / "localized.txt").read_bytes() == b" 12,5 \r\n"
    assert (tmp_path / "portable.txt").read_bytes() == b"12.5,12.5\r\n"


def test_locale_time_separator_round_trips_and_invalid_meridiem_fails():
    p = compile_source(dedent("""
        Option Explicit
        Function RoundTrip() As Date
            RoundTrip = CDate(FormatDateTime(#15:07:05#, vbLongTime))
        End Function
        Function ValidTime(ByVal s As String) As Boolean
            ValidTime = IsDate(s)
        End Function
        Function WeekBoundaries() As Long
            WeekBoundaries = DateDiff("ww", #01/06/2024#, #01/07/2024#, vbUseSystem)
        End Function
    """)).create_program(locale="fi-FI")
    assert p.run("RoundTrip") == datetime(1899, 12, 30, 15, 7, 5)
    assert p.run("ValidTime", "22:00 PM") is False
    assert p.run("ValidTime", "0 PM") is False
    assert p.run("WeekBoundaries") == 0


def test_decimal_arithmetic_preserves_the_final_digit_at_96_bit_precision():
    p = compile_source(dedent("""
        Option Explicit
        Function StepDown()
            StepDown = CDec("79228162514264337593543950335") - CDec("1")
        End Function
        Function NegativeMax()
            NegativeMax = -CDec("79228162514264337593543950335")
        End Function
        Function TooLarge()
            TooLarge = CDec("79228162514264337593543950335") + CDec("1")
        End Function
    """)).create_program()
    assert p.run("StepDown") == Decimal("79228162514264337593543950334")
    assert p.run("NegativeMax") == Decimal("-79228162514264337593543950335")
    with pytest.raises(VBAError) as error:
        p.run("TooLarge")
    assert error.value.number == 6
