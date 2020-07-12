Attribute VB_Name = "DATE_WEEK_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : DAYS_WEEK_TWO_DATES_FUNC
'DESCRIPTION   : This returns the number of Day Of Week days between two dates.
'For example, the number of Tuesdays between 15-Jan-2009 and 26-July-2010.
'LIBRARY       : DATE
'GROUP         : WEEK
'ID            : 00X
'LAST UPDATE   : 01/09/2009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************


Function DAYS_WEEK_TWO_DATES_FUNC( _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
ByVal WEEK_DAY_INT As VbDayOfWeek)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DaysOfWeekBetweenTwoDate
' This function returns the number of DaysOfWeek between START_DATE and
' END_DATE. START_DATE is the first date, END_DATE is the last date, and
' WEEK_DAY_INT is an long  between 1 and 7 (1 = Sunday, 2 = Monday, ...
' 7 = Saturday). If START_DATE is later than END_DATE, the result is #NUM!.
' If WEEK_DAY_INT is out of range, the result is #VALUE.
' Note that this function uses WS_MOD_FUNC to use Excel's worksheet function MOD
' rather than VBA's Mod operator.
'
' Worksheet function equivalent:
'
' =((END_DATE-MOD(WEEKDAY(END_DATE)-WEEK_DAY_INT,7)-START_DATE-
'   MOD(WEEK_DAY_INT-WEEKDAY(START_DATE)+7,7))/7)+1
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo ERROR_LABEL

If START_DATE > END_DATE Then: GoTo ERROR_LABEL
If (WEEK_DAY_INT < vbSunday) Or (WEEK_DAY_INT > vbSaturday) Then: GoTo ERROR_LABEL
If (START_DATE < 0) Or (END_DATE < 0) Then: GoTo ERROR_LABEL

DAYS_WEEK_TWO_DATES_FUNC = _
    ((END_DATE - WS_MOD_FUNC(Weekday(END_DATE) - WEEK_DAY_INT, 7) - START_DATE - _
    WS_MOD_FUNC(WEEK_DAY_INT - Weekday(START_DATE) + 7, 7)) / 7) + 1


Exit Function
ERROR_LABEL:
DAYS_WEEK_TWO_DATES_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : DAYS_WEEK_MONTH_FUNC
'DESCRIPTION   : This returns the number of a given Day Of Week in a given month
'and year. For example, the number of Tuesdays in April, 2009.
'LIBRARY       : DATE
'GROUP         : WEEK
'ID            : 00X
'LAST UPDATE   : 01/09/2009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************


Function DAYS_WEEK_MONTH_FUNC(ByVal MONTH_INT As Long, _
ByVal YEAR_INT As Long, _
ByVal WEEK_DAY_INT As VbDayOfWeek)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DAYS_WEEK_MONTH_FUNC
' This function returns the number of DaysOfWeek in the month MONTH_INT in
' year YEAR_INT.  If either the MONTH_INT or YEAR_INT value is out of range, the
' result is #VALUE.
' Note that this function uses WS_MOD_FUNC to use Excel's worksheet function MOD
' rather than VBA's Mod operator.
'
' Formula equivalent:
'   =((DATE(YEAR_INT,MMonth+1,0)-MOD(WEEKDAY(DATE(YEAR_INT,MMonth+1,0))-WEEK_DAY_INT,7)-
'       DATE(YEAR_INT,MONTH_INT,1)-MOD(WEEK_DAY_INT-WEEKDAY(DATE(YEAR_INT,MONTH_INT,1))+7,7))/7)+1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo ERROR_LABEL

If (MONTH_INT < 1) Or (MONTH_INT > 12) Then: GoTo ERROR_LABEL
If (YEAR_INT < 1900) Or (YEAR_INT > 9999) Then: GoTo ERROR_LABEL
If (WEEK_DAY_INT < vbSunday) Or (WEEK_DAY_INT > vbSaturday) Then: GoTo ERROR_LABEL

DAYS_WEEK_MONTH_FUNC = ((DateSerial(YEAR_INT, MONTH_INT + 1, 0) - _
    WS_MOD_FUNC(Weekday(DateSerial(YEAR_INT, MONTH_INT + 1, 0)) - WEEK_DAY_INT, 7) - _
    DateSerial(YEAR_INT, MONTH_INT, 1) - WS_MOD_FUNC(WEEK_DAY_INT - _
    Weekday(DateSerial(YEAR_INT, MONTH_INT, 1)) + 7, 7)) / 7) + 1

Exit Function
ERROR_LABEL:
DAYS_WEEK_MONTH_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : PREVIOUS_DAY_WEEK_FUNC
'DESCRIPTION   : This returns the date of the first Day Of Week before a given date.
'For example, the date of the Tuesday before 15-June-2009.
'LIBRARY       : DATE
'GROUP         : WEEK
'ID            : 00X
'LAST UPDATE   : 01/09/2009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function PREVIOUS_DAY_WEEK_FUNC(ByVal START_DATE As Date, _
ByVal WEEK_DAY_INT As VbDayOfWeek)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PREVIOUS_DAY_WEEK_FUNC
' This function returns the date of the WEEK_DAY_INT prior to START_DATE.
' Note that this function uses WS_MOD_FUNC to use Excel's worksheet function MOD
' rather than VBA's Mod operator.
' Formula equivalent:
'  =StartDate-MOD(WEEKDAY(START_DATE)-WEEK_DAY_INT,7)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo ERROR_LABEL

If (WEEK_DAY_INT < vbSunday) Or (WEEK_DAY_INT > vbSaturday) Then: GoTo ERROR_LABEL
If (START_DATE < 0) Then: GoTo ERROR_LABEL

PREVIOUS_DAY_WEEK_FUNC = START_DATE - WS_MOD_FUNC(Weekday(START_DATE) - WEEK_DAY_INT, 7)

Exit Function
ERROR_LABEL:
PREVIOUS_DAY_WEEK_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NEXT_DAY_WEEK_FUNC
'DESCRIPTION   : This returns the date of the first Day Of Week following a given
'date. For example, the date of the first Tuesday after 15-June-2009.
'LIBRARY       : DATE
'GROUP         : WEEK
'ID            : 00X
'LAST UPDATE   : 01/09/2009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function NEXT_DAY_WEEK_FUNC(ByVal START_DATE As Date, _
ByVal WEEK_DAY_INT As VbDayOfWeek)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NEXT_DAY_WEEK_FUNC
' This function returns the date of the WEEK_DAY_INT following START_DATE.
' Note that this function uses WS_MOD_FUNC to use Excel's worksheet function MOD
' rather than VBA's Mod operator.
' Formula equivalent:
' =StartDate+MOD(WEEK_DAY_INT-WEEKDAY(START_DATE),7)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo ERROR_LABEL

If (WEEK_DAY_INT < vbSunday) Or (WEEK_DAY_INT > vbSaturday) Then: GoTo ERROR_LABEL
If (START_DATE < 0) Then: GoTo ERROR_LABEL

NEXT_DAY_WEEK_FUNC = START_DATE + WS_MOD_FUNC(WEEK_DAY_INT - Weekday(START_DATE), 7)

Exit Function
ERROR_LABEL:
NEXT_DAY_WEEK_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : FIRST_DAY_WEEK_MONTH_FUNC
'DESCRIPTION   : This returns the date of the first Day Of Week day in a given month
'and year. For example, the date of the first Friday in March, 2010.
'LIBRARY       : DATE
'GROUP         : WEEK
'ID            : 00X
'LAST UPDATE   : 01/09/2009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function FIRST_DAY_WEEK_MONTH_FUNC(ByVal MONTH_INT As Long, _
ByVal YEAR_INT As Long, _
ByVal WEEK_DAY_INT As VbDayOfWeek)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This returns the date of the first WEEK_DAY_INT in month MM in year YYYY.
' Formula equivalent:
'   =DATE(YEAR_INT,MONTH_INT,1)+(MOD(WEEK_DAY_INT-WEEKDAY(DATE(YEAR_INT,MONTH_INT,1)),7))
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo ERROR_LABEL

If (WEEK_DAY_INT < vbSunday) Or (WEEK_DAY_INT > vbSaturday) Then: GoTo ERROR_LABEL
If (MONTH_INT < 1) Or (MONTH_INT > 12) Then: GoTo ERROR_LABEL
    
FIRST_DAY_WEEK_MONTH_FUNC = DateSerial(YEAR_INT, MONTH_INT, 1) + _
    WS_MOD_FUNC(WEEK_DAY_INT - Weekday(DateSerial(YEAR_INT, MONTH_INT, 1)), 7)


Exit Function
ERROR_LABEL:
FIRST_DAY_WEEK_MONTH_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : LAST_DAY_WEEK_MONTH_FUNC
'DESCRIPTION   : This returns the date of the last Day Of Week day in a given month
'and year. For exampe, the date of the last Friday in May, 2009.
'LIBRARY       : DATE
'GROUP         : WEEK
'ID            : 00X
'LAST UPDATE   : 01/09/2009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function LAST_DAY_WEEK_MONTH_FUNC(ByVal MONTH_INT As Long, _
ByVal YEAR_INT As Long, _
ByVal WEEK_DAY_INT As VbDayOfWeek)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LAST_DAY_WEEK_MONTH_FUNC
' This returns the date of the last WEEK_DAY_INT in month MM in year YYYY.
' Formula equivalent:
'       =DATE(YEAR_INT,MMonth+1,0)-ABS(WEEKDAY(DATE(YEAR_INT,MMonth+1,0))-WEEK_DAY_INT)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo ERROR_LABEL

If (WEEK_DAY_INT < vbSunday) Or (WEEK_DAY_INT > vbSaturday) Then: GoTo ERROR_LABEL

LAST_DAY_WEEK_MONTH_FUNC = DateSerial(YEAR_INT, MONTH_INT + 1, 0) - _
    Abs(Weekday(DateSerial(YEAR_INT, MONTH_INT + 1, 0)) - WEEK_DAY_INT)


Exit Function
ERROR_LABEL:
LAST_DAY_WEEK_MONTH_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NTH_DAY_WEEK_MONTH_FUNC
'DESCRIPTION   : This returns the date of the NTH_INT Day Of Week day in a given
'month and year. For example, the date of the third Friday in May, 2009.
'LIBRARY       : DATE
'GROUP         : WEEK
'ID            : 00X
'LAST UPDATE   : 01/09/2009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function NTH_DAY_WEEK_MONTH_FUNC(ByVal MONTH_INT As Long, _
ByVal YEAR_INT As Long, _
ByVal WEEK_DAY_INT As VbDayOfWeek, _
ByVal NTH_INT As Long)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NTH_DAY_WEEK_MONTH_FUNC
' This returns the NTH_INT Day Of Week in month MM in year YYYY.
' Formula equivalent:
'
'   =DATE(YEAR_INT,MONTH_INT,1)+(MOD(WEEK_DAY_INT-WEEKDAY(DATE(YEAR_INT,MONTH_INT,1)),7))+(7*(NTH_INT-1))
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo ERROR_LABEL

If (MONTH_INT < 1) Or (MONTH_INT > 12) Then: GoTo ERROR_LABEL
If (YEAR_INT < 1900) Or (YEAR_INT > 9999) Then: GoTo ERROR_LABEL
If (WEEK_DAY_INT < vbSunday) Or (WEEK_DAY_INT > vbSaturday) Then: GoTo ERROR_LABEL
If NTH_INT < 0 Then: GoTo ERROR_LABEL

NTH_DAY_WEEK_MONTH_FUNC = DateSerial(YEAR_INT, MONTH_INT, 1) + _
    (WS_MOD_FUNC(WEEK_DAY_INT - Weekday(DateSerial(YEAR_INT, MONTH_INT, 1)), 7)) + _
    (7 * (NTH_INT - 1))

Exit Function
ERROR_LABEL:
NTH_DAY_WEEK_MONTH_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : LAST_DAY_WEEK_YEAR_FUNC
'DESCRIPTION   :
'This returns the date of the last Day Of Week day of a given year. For example, the date of the
'last Monday in 2009. These functions, in both worksheet formula and VBA implementations, are
'described below. The WS_MOD_FUNC function, which is used in place of VBA's Mod operator, is as follows:

'LIBRARY       : DATE
'GROUP         : WEEK
'ID            : 00X
'LAST UPDATE   : 01/09/2009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function LAST_DAY_WEEK_YEAR_FUNC(ByVal YEAR_INT As Long, _
ByVal WEEK_DAY_INT As VbDayOfWeek)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LAST_DAY_WEEK_YEAR_FUNC
' This returns the last WEEK_DAY_INT of the year YEAR_INT.
' Formula equivalent:
'   =DATE(YEAR_INT,12,31)-MOD(WEEKDAY(DATE(YEAR_INT,12,31))-WEEK_DAY_INT,7)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo ERROR_LABEL

If (YEAR_INT < 1900) Or (YEAR_INT > 9999) Then: GoTo ERROR_LABEL
If (WEEK_DAY_INT < vbSunday) Or (WEEK_DAY_INT > vbSaturday) Then: GoTo ERROR_LABEL

LAST_DAY_WEEK_YEAR_FUNC = DateSerial(YEAR_INT, 12, 31) - _
    WS_MOD_FUNC(Weekday(DateSerial(YEAR_INT, 12, 31)) - WEEK_DAY_INT, 7)

Exit Function
ERROR_LABEL:
LAST_DAY_WEEK_YEAR_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : FIRST_DAY_WEEK_YEAR_FUNC
'DESCRIPTION   : This returns the date of the first Day Of Week day of a given
'year. For example, the date of the first Friday in 2009.
'LIBRARY       : DATE
'GROUP         : WEEK
'ID            : 00X
'LAST UPDATE   : 01/09/2009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function FIRST_DAY_WEEK_YEAR_FUNC(ByVal YEAR_INT As Long, _
ByVal WEEK_DAY_INT As VbDayOfWeek)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FIRST_DAY_WEEK_YEAR_FUNC
' This returns the date of the first WEEK_DAY_INT in the year YEAR_INT.
' Formula equivalent:
'   =DATE(YEAR_INT,1,1)+MOD(WEEK_DAY_INT-WEEKDAY(DATE(YEAR_INT,1,1)),7)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo ERROR_LABEL

If (YEAR_INT < 1900) Or (YEAR_INT > 9999) Then: GoTo ERROR_LABEL
If (WEEK_DAY_INT < vbSunday) Or (WEEK_DAY_INT > vbSaturday) Then: GoTo ERROR_LABEL

FIRST_DAY_WEEK_YEAR_FUNC = DateSerial(YEAR_INT, 1, 1) + _
    WS_MOD_FUNC(WEEK_DAY_INT - Weekday(DateSerial(YEAR_INT, 1, 1)), 7)

Exit Function
ERROR_LABEL:
FIRST_DAY_WEEK_YEAR_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : FIRST_MONDAY_WEEK_YEAR_FUNC
'DESCRIPTION   : Returns the date of the Monday of the week containing
'the first of which year.

'LIBRARY       : DATE
'GROUP         : WEEK
'ID            : 001
'LAST UPDATE   : 11 / 02 / 2004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'**********************************************************************************
'**********************************************************************************

Function FIRST_MONDAY_WEEK_YEAR_FUNC(ByVal YEAR_VAL As Integer)
    
Dim WEEK_DAY_VAL As Integer
Dim TEMP_DATE As Date 'NewYear

On Error GoTo ERROR_LABEL

TEMP_DATE = DateSerial(YEAR_VAL, 1, 1)
WEEK_DAY_VAL = (TEMP_DATE - 2) Mod 7
'Generate Week Day index where Monday = 0
If WEEK_DAY_VAL < 4 Then
    FIRST_MONDAY_WEEK_YEAR_FUNC = TEMP_DATE - WEEK_DAY_VAL
Else
    FIRST_MONDAY_WEEK_YEAR_FUNC = TEMP_DATE - WEEK_DAY_VAL + 7
End If

Exit Function
ERROR_LABEL:
FIRST_MONDAY_WEEK_YEAR_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NTH_DAY_WEEK_FUNC
'DESCRIPTION   : This function returns a date representing the Nth DayOfWeek
' for the supplied Month and Year (e.g, Third Tuesday of May, 1998)

'LIBRARY       : DATE
'GROUP         : WEEK
'ID            : 002
'LAST UPDATE   : 11 / 02 / 2004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function NTH_DAY_WEEK_FUNC(ByVal Y_VAL As Integer, _
ByVal M_VAL As Integer, _
ByVal N_VAL As Integer, _
ByVal D_VAL As Integer)

On Error GoTo ERROR_LABEL

NTH_DAY_WEEK_FUNC = DateSerial(Y_VAL, M_VAL, (8 - _
                Weekday(DateSerial(Y_VAL, M_VAL, 1), _
                (D_VAL + 1) Mod 8)) + ((N_VAL - 1) * 7))

Exit Function
ERROR_LABEL:
NTH_DAY_WEEK_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : WEEK_NUMBER_FUNC
'DESCRIPTION   : This function returns the number of the week in date's year
'LIBRARY       : DATE
'GROUP         : WEEK
'ID            : 003
'LAST UPDATE   : 11 / 02 / 2004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'**********************************************************************************
'**********************************************************************************


Function WEEK_NUMBER_FUNC(ByVal SETTLEMENT As Date)
    
Dim TEMP_DATE As Date

On Error GoTo ERROR_LABEL

SETTLEMENT = Int(SETTLEMENT)

TEMP_DATE = _
DateSerial(Year(SETTLEMENT + (8 - Weekday(SETTLEMENT)) Mod 7 - 3), 1, 1)

WEEK_NUMBER_FUNC = _
((SETTLEMENT - TEMP_DATE - 3 + (Weekday(TEMP_DATE) + 1) Mod 7)) \ 7 + 1
    
Exit Function
ERROR_LABEL:
WEEK_NUMBER_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : ISE_WEEK_NUMBER_FUNC
'DESCRIPTION   : VERSION: missing or <> 2 then returns week number,
'  = 2 then YYWW
'LIBRARY       : DATE
'GROUP         : WEEK
'ID            : 004

'LAST UPDATE   : 11 / 02 / 2004

'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'**********************************************************************************
'**********************************************************************************

Function ISE_WEEK_NUMBER_FUNC(ByVal ANY_DATE As Date, _
Optional ByVal VERSION As Integer = 0)
    
Dim i As Integer
Dim j As Integer 'YearNum
Dim k As Integer 'This Year

Dim PREV_START As Date 'PreviousYearStart
Dim THIS_START As Date 'This Year Start
Dim NEW_START As Date 'Next Year Start

On Error GoTo ERROR_LABEL

k = Year(ANY_DATE)
THIS_START = FIRST_MONDAY_WEEK_YEAR_FUNC(k)
PREV_START = FIRST_MONDAY_WEEK_YEAR_FUNC(k - 1)
NEW_START = FIRST_MONDAY_WEEK_YEAR_FUNC(k + 1)

Select Case ANY_DATE
    Case Is >= NEW_START
        i = (ANY_DATE - NEW_START) \ 7 + 1
        j = Year(ANY_DATE) + 1
    Case Is < THIS_START
        i = (ANY_DATE - PREV_START) \ 7 + 1
        j = Year(ANY_DATE) - 1
    Case Else
        i = (ANY_DATE - THIS_START) \ 7 + 1
        j = Year(ANY_DATE)
End Select

Select Case VERSION
    Case 0
        ISE_WEEK_NUMBER_FUNC = i
    Case Else
        ISE_WEEK_NUMBER_FUNC = CInt(Format(Right(j, 2), "00") & _
            Format(i, "00"))
End Select
    
Exit Function
ERROR_LABEL:
ISE_WEEK_NUMBER_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      :
'DESCRIPTION   :
'LIBRARY       : DATE
'GROUP         : WEEK
'ID            : 00X
'LAST UPDATE   : 01/09/2009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function WS_MOD_FUNC(ByVal NUMBER_VAL As Double, _
ByVal DIVISOR_VAL As Double) As Double
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WS_MOD_FUNC
' The Excel worksheet function MOD and the VBA Mod operator
' work differently and can return different results under
' certain circumstances. For continuity between the worksheet
' formulas and the VBA code, we use this WS_MOD_FUNC function, which
' produces the same result as the Excel MOD worksheet function,
' rather than the VBA Mod operator.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo ERROR_LABEL

WS_MOD_FUNC = NUMBER_VAL - DIVISOR_VAL * Int(NUMBER_VAL / DIVISOR_VAL)

Exit Function
ERROR_LABEL:
WS_MOD_FUNC = Err.number
End Function

