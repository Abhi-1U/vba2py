from textwrap import dedent

import pytest

from vba2py import VBAError, compile_source
from vba2py.runtime.values import NULL


def test_instr_rev_honors_start_compare_and_utf16_positions():
    compiled = compile_source(dedent("""
        Option Explicit
        Option Compare Text
        Function FindLast(ByVal source As String, ByVal needle As String, _
                          ByVal position As Long, ByVal mode As Long) As Long
            FindLast = InStrRev(source, needle, position, mode)
        End Function
    """))
    project = compiled.create_program()
    assert project.run("FindLast", "ababa", "ABA", -1, -1) == 3
    assert project.run("FindLast", "ababa", "ABA", 4, -1) == 1
    assert project.run("FindLast", "ababa", "ABA", -1, 0) == 0
    assert project.run("FindLast", "😀abcab", "ab", -1, 0) == 6
    assert project.run("FindLast", "😀abcab", "ab", 6, 0) == 3
    assert project.run("FindLast", "abc", "", 2, 0) == 2
    assert project.run("FindLast", "abc", "a", 4, 0) == 0
    turkish = compiled.create_program(locale="tr-TR")
    assert turkish.run("FindLast", "I\u0307 I", "i", -1, -1) == 1


def test_filter_returns_a_zero_based_string_array_and_can_exclude_matches():
    project = compile_source(dedent("""
        Option Explicit
        Option Compare Text
        Option Base 1
        Function Match(ByVal keep As Boolean) As String
            Dim result
            result = Filter(Array("Alpha", "beta", "ALPHABETA"), "alpha", _
                            Include:=keep, Compare:=vbUseCompareOption)
            Match = CStr(LBound(result)) & ":" & TypeName(result) & ":" & Join(result, "|")
        End Function
    """)).create_program()
    assert project.run("Match", True) == "0:String():Alpha|ALPHABETA"
    assert project.run("Match", False) == "0:String():beta"


def test_str_conv_proper_case_and_locale_override():
    project = compile_source(dedent("""
        Option Explicit
        Function Proper() As String
            Proper = StrConv("DOESN'T    STOP", vbProperCase)
        End Function
        Function Upper() As String
            Upper = StrConv("i", vbUpperCase)
        End Function
        Function UpperTurkish() As String
            UpperTurkish = StrConv("i", vbUpperCase, LCID:=1055)
        End Function
    """))
    assert project.create_program().run("Proper") == "Doesn't    Stop"
    assert project.create_program().run("Upper") == "I"
    assert project.create_program(locale="tr-TR").run("Upper") == "İ"
    assert project.create_program().run("UpperTurkish") == "İ"


def test_switch_selects_the_first_true_case_and_evaluates_all_values():
    project = compile_source(dedent("""
        Option Explicit
        Function Pick()
            Pick = Switch(True, "first", True, "second")
        End Function
        Function MissingMatch()
            MissingMatch = Switch(False, "unused")
        End Function
        Function Fails()
            Fails = Switch(True, "first", False, 1 / 0)
        End Function
    """)).create_program()
    assert project.run("Pick") == "first"
    assert project.run("MissingMatch") is NULL
    with pytest.raises(VBAError) as error:
        project.run("Fails")
    assert error.value.number == 11


def test_calendar_functions_use_the_program_locale_and_week_rules():
    compiled = compile_source(dedent("""
        Option Explicit
        Function CalendarNames() As String
            CalendarNames = MonthName(3) & "|" & WeekdayName(1) & "|" & _
                            CStr(Weekday(#2/12/2024#, vbUseSystem))
        End Function
        Function WeekNumber(ByVal day As Date) As Integer
            WeekNumber = DatePart("ww", day, vbMonday, vbFirstFourDays)
        End Function
        Function DatePieces() As Long
            DatePieces = DatePart("q", #4/26/2026#) * 100 + DatePart("y", #1/30/2026#)
        End Function
    """))
    american = compiled.create_program()
    german = compiled.create_program(locale="de-DE")
    assert american.run("CalendarNames") == "March|Sunday|2"
    assert german.run("CalendarNames") == "März|Montag|1"
    assert american.run("WeekNumber", "1/1/2021") == 53
    assert american.run("WeekNumber", "12/31/2018") == 1
    assert american.run("DatePieces") == 230


def test_financial_functions_apply_due_dates_and_first_period_discount():
    project = compile_source(dedent("""
        Option Explicit
        Function FutureWithEarlyPayments() As Double
            FutureWithEarlyPayments = FV(0.1, 2, -100, Type:=1)
        End Function
        Function PresentWithZeroInterest() As Double
            PresentWithZeroInterest = PV(0, 3, -10, FV:=-20)
        End Function
        Function CashFlows() As Double
            Dim flows(1 To 3) As Double
            flows(1) = -100
            flows(2) = 60
            flows(3) = 60
            CashFlows = NPV(0.1, flows)
        End Function
    """)).create_program()
    assert project.run("FutureWithEarlyPayments") == pytest.approx(231)
    assert project.run("PresentWithZeroInterest") == 50.0
    assert project.run("CashFlows") == pytest.approx(-100 / 1.1 + 60 / 1.1**2 + 60 / 1.1**3)
