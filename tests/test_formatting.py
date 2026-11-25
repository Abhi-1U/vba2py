"""Exercise formatting functions through compiled VBA and regional settings."""

import pytest

from vba2py import VBAError, compile_source


@pytest.fixture(scope="module")
def macros():
    return compile_source('''
Option Explicit

Function Show(ByVal value As Variant, ByVal mask As String) As String
    Show = Format$(value, mask)
End Function

Function NumberReport() As String
    NumberReport = FormatNumber(1234.5, 2) & "|" & _
        VBA.Strings.FormatNumber(-0.25, 2, vbFalse, vbTrue, vbFalse)
End Function

Function CurrencyReport() As String
    CurrencyReport = FormatCurrency(1234.5, , GroupDigits:=vbTrue)
End Function

Function PercentReport() As String
    PercentReport = FormatPercent(0.125, NumDigitsAfterDecimal:=1)
End Function

Function DateReport() As String
    Dim value As Date
    value = #2024-07-15 15:07:08#
    DateReport = Format(value, "dddd d mmmm yyyy") & "|" & _
        Format(value, "hh:mm:ss AM/PM") & "|" & _
        Format(value, "dd/mm/yyyy") & "|" & _
        FormatDateTime(value, vbShortTime)
End Function

Function WeekReport() As String
    WeekReport = Format(#2021-01-01#, "w/ww", vbMonday, vbFirstFourDays)
End Function

Function NullReport() As String
    NullReport = Format(Null, "0;(0);""zero"";""missing""") & "|" & _
        Format(Null, "@;""no text""") & "|" & Format(Null, "0.00")
End Function

Function FormatError() As Long
    On Error Resume Next
    Dim result As String
    result = FormatNumber(1, -2)
    FormatError = Err.Number
End Function
''', name="Reports")


@pytest.mark.parametrize(("value", "mask", "expected"), [
    (1234.125, "0.00", "1234.13"),
    (12, "00000", "00012"),
    (1.25, "0.0##", "1.25"),
    (1.2, "0.0##", "1.2"),
    (-1234.5, '"$"#,##0.00;("$"#,##0.00)', "($1,234.50)"),
    (-1234, '"$"#,##0;;"zero"', "-$1,234"),
    (0, '0;(0);"zero; items"', "zero; items"),
    (12345678, '0,,.0 "million"', "12.3 million"),
    (0.0123, "0.00E-00", "1.23E-02"),
    (9996, "0.00E+00", "1.00E+04"),
    (1200, "0.0e-00", "1.2e03"),
    (0.125, "0.0%", "12.5%"),
    (0.25, "#.00", ".25"),
    (4155550177, "(000) 000-0000", "(415) 555-0177"),
    (1234, '0" / "00', "12 / 34"),
    (3.14, r"0.00\%", "3.14%"),
    (0, "Yes/No", "No"),
    (-1, "True/False", "True"),
    (1, "On/Off", "On"),
    (1234.5, "General Number", "1234.5"),
    (1234.5, "Fixed", "1234.50"),
    (1234.5, "Standard", "1,234.50"),
    ("ab", ">@@@!", "AB "),
    ("ab", "@@@", " ab"),
    ("ab", "!&&&", "ab"),
    ("ab", '">"@@', ">ab"),
    ("", '@;"empty"', "empty"),
])
def test_vba_format_masks(macros, value, mask, expected):
    assert macros.create_program().run("Show", value, mask) == expected


@pytest.mark.parametrize(("locale", "number", "currency", "percent"), [
    ("en-US", "1,234.50|(.25)", "$1,234.50", "12.5%"),
    ("de-DE", "1.234,50|(,25)", "1.234,50\xa0€", "12,5\xa0%"),
    ("hi-IN", "1,234.50|(.25)", "₹1,234.50", "12.5%"),
    ("ja-JP", "1,234.50|(.25)", "￥1,235", "12.5%"),
])
def test_regional_convenience_formatters(macros, locale, number, currency, percent):
    program = macros.create_program(locale=locale)
    assert program.run("NumberReport") == number
    assert program.run("CurrencyReport") == currency
    assert program.run("PercentReport") == percent


def test_indian_grouping_and_english_program_default(macros):
    india = macros.create_program(locale="hi-IN")
    english = macros.create_program()
    assert india.run("Show", 1234567.5, "#,##0.00") == "12,34,567.50"
    assert english.run("Show", 1234567.5, "#,##0.00") == "1,234,567.50"


def test_date_month_names_separators_and_time_formats(macros):
    english = macros.create_program()
    german = macros.create_program(locale="de-DE")
    assert english.run("DateReport") == "Monday 15 July 2024|03:07:08 PM|15/07/2024|15:07"
    assert german.run("DateReport") == "Montag 15 Juli 2024|03:07:08 PM|15.07.2024|15:07"
    assert english.run("WeekReport") == "5/53"


def test_null_sections_and_invalid_parameters(macros):
    program = macros.create_program()
    assert program.run("NullReport") == "missing|no text|"
    assert program.run("FormatError") == 5
    with pytest.raises(VBAError, match="Invalid Format sections") as caught:
        program.run("Show", 42, '0;"unclosed')
    assert caught.value.number == 5
