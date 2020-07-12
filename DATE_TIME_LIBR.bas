Attribute VB_Name = "DATE_TIME_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : MINUTES_DIFFERENCE_FUNC
'DESCRIPTION   : This function computes the difference in minutes between
'two dates and times: D1 and D2 are dates, T1 and T2 are times As variant
'1130 = 11:30, 1840 = 18:40

'LIBRARY       : DATE
'GROUP         : TIME
'ID            : 001
'LAST UPDATE   : 11 / 02 / 2004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'**********************************************************************************
'**********************************************************************************

Function MINUTES_DIFFERENCE_FUNC(ByVal DATE1_VAL As Variant, _
ByVal TIME1_VAL As Variant, _
ByVal DATE2_VAL As Variant, _
ByVal TIME2_VAL As Variant)

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim ATEMP_MULT As Double
Dim BTEMP_MULT As Double

On Error GoTo ERROR_LABEL

ATEMP_VAL = Int(TIME1_VAL / 100) + (((TIME1_VAL / 100) - Int(TIME1_VAL / 100)) / 0.6) '23.50
ATEMP_VAL = ATEMP_VAL / 24

BTEMP_VAL = Int(TIME2_VAL / 100) + (((TIME2_VAL / 100) - Int(TIME2_VAL / 100)) / 0.6) '23.50
BTEMP_VAL = BTEMP_VAL / 24

ATEMP_MULT = DATE1_VAL + ATEMP_VAL
BTEMP_MULT = DATE2_VAL + BTEMP_VAL

MINUTES_DIFFERENCE_FUNC = DateDiff("n", ATEMP_MULT, BTEMP_MULT)

Exit Function
ERROR_LABEL:
MINUTES_DIFFERENCE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : MILITAR_TIME_FUNC
'DESCRIPTION   : Input TIME_VAL: 24-hour time As variant,e.g.,
'1130=11:30, 1650=16:50 Output, time as serial time e.g, 0.5 for noon.
' 12 noon = 0.5 of a day

'LIBRARY       : DATE
'GROUP         : MILITAR
'ID            : 001
'LAST UPDATE   : 11 / 02 / 2004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'**********************************************************************************
'**********************************************************************************

Function MILITAR_TIME_FUNC(ByVal TIME_VAL As Double)

Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

TEMP_VAL = Int(TIME_VAL / 100) + (((TIME_VAL / 100) - _
           Int(TIME_VAL / 100)) / 0.6) '23.50

TEMP_VAL = TEMP_VAL / 24

MILITAR_TIME_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
MILITAR_TIME_FUNC = Err.number
End Function


Function TIME_STAMP_TO_DATE_FUNC(Optional ByVal TIME_VAL As Double = 0) As Date
On Error GoTo ERROR_LABEL ' Jan 1 1970 = 25569
If TIME_VAL = 0 Then: TIME_VAL = Now
TIME_STAMP_TO_DATE_FUNC = (TIME_VAL / 60 / 60 / 24) + 25569
Exit Function
ERROR_LABEL:
TIME_STAMP_TO_DATE_FUNC = 0
End Function


'Adding Times

Function ADD_TIME_FUNC(ByVal TIME_VAL As Date, _
ByVal H_VAL As Integer, _
ByVal M_VAL As Integer, _
ByVal S_VAL As Integer)
'H:M:S

On Error GoTo ERROR_LABEL

ADD_TIME_FUNC = TIME_VAL + TimeSerial(H_VAL, M_VAL, S_VAL)

Exit Function
ERROR_LABEL:
ADD_TIME_FUNC = Err.number
End Function


'Subtracting Times

Function SUBTRACT_TIME_FUNC(ByVal TIME_VAL As Date, _
ByVal H_VAL As Integer, _
ByVal M_VAL As Integer, _
ByVal S_VAL As Integer)
'H:M:S

On Error GoTo ERROR_LABEL

SUBTRACT_TIME_FUNC = TIME_VAL - TimeSerial(H_VAL, M_VAL, S_VAL)

Exit Function
ERROR_LABEL:
SUBTRACT_TIME_FUNC = Err.number
End Function
Time Intervals

'You can determine the number of hours and minutes between two times by subtracting the two
'times.  However, since Excel cannot handle negative times, you must use an =IF statement to
'adjust the time accordingly.  If your times were entered without a date (e.g, 22:30), the
'following statement will compute the interval between two times in TIME1_VAL and TIME2_VAL .

Function HOURS_MINUTES_DIFFERENCE_FUNC(ByVal TIME1_VAL As Variant, _
ByVal TIME2_VAL As Variant, _
Optional ByVal VERSION As Integer = 0)

On Error GoTo ERROR_LABEL

Select Case VERSION
Case 0
    HOURS_MINUTES_DIFFERENCE_FUNC = IIf(TIME1_VAL > TIME2_VAL, TIME2_VAL + 1 - TIME1_VAL, _
                                    TIME2_VAL - TIME1_VAL)

Case Else
'The "+1" in the formula causes Excel to treat TIME2_VAL  as if it were in the next day, so
'02:30-22:00 will result in 4:30, four hours and thirty minutes, which is what we would
'expect.  To covert this to a decimal number, for example, 4.5, indicating how many hours,
'multiply the result by 24 and format the cell as General or Decimal, as in
    HOURS_MINUTES_DIFFERENCE_FUNC = 24 * (IIf(TIME1_VAL > TIME2_VAL, TIME2_VAL + 1 - TIME1_VAL, _
                                    TIME2_VAL - TIME1_VAL))
End Select

Exit Function
ERROR_LABEL:
HOURS_MINUTES_DIFFERENCE_FUNC = Err.number
End Function


'For many scheduling or payroll applications, it is useful to round times to the nearest hour,
'half-hour, or quarter-hour.  The MROUND function, which is part of the Analysis ToolPack
'add-in module, is very useful for this.  Suppose you have a time in cell TIME1_VAL.
'In TIME2_VAL , enter the number of minutes to which you want to round the time -- for example,
'enter 30 to round to the nearest half-hour.  The formula

Function TIME_ROUNDING_FUNC(ByVal TIME1_VAL As Variant, _
ByVal TIME2_VAL As Variant, _
Optional ByVal VERSION As Integer = 0)

On Error GoTo ERROR_LABEL

Select Case VERSION
Case 0
    TIME_ROUNDING_FUNC = TimeSerial(Hour(TIME1_VAL), WorksheetFunction.MRound(Minute(TIME1_VAL), TIME2_VAL), 0)

'Will return a time rounded to the nearest half-hour, either up or down, depending what is closest.
'For example, 12:14 is rounded to 12:00, and 12:15 is rounded to 12:30.

'To round either up or down to the nearest interval, enter the interval in TIME2_VAL, and use
'either of the following formulas:
Case 1
    TIME_ROUNDING_FUNC = Time(Hour(TIME1_VAL), WorksheetFunction.Floor(Minute(TIME1_VAL), TIME2_VAL), 0)
    'to round to the previous interval (always going earlier, or staying the same).
Case Else
    TIME_ROUNDING_FUNC = Time(Hour(TIME1_VAL), WorksheetFunction.Ceiling(Minute(TIME1_VAL), TIME2_VAL), 0)
    'to round to the next interval (always going later, or staying the same).
End Select

Exit Function
ERROR_LABEL:
TIME_ROUNDING_FUNC = Err.number
End Function

