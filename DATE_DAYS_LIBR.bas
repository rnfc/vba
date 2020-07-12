Attribute VB_Name = "DATE_DAYS_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : COUNT_DAYS_FUNC
'DESCRIPTION   : CALCULATE THE NUMBER OF DAYS BETWEEN TWO DATES
'LIBRARY       : DATE
'GROUP         : DAYS
'ID            : 001
'LAST UPDATE   : 11 / 02 / 2004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'**********************************************************************************
'**********************************************************************************

Function COUNT_DAYS_FUNC(ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal COUNT_BASIS As Integer = 0)

Dim DAY1_VAL As Long
Dim DAY2_VAL As Long

Dim MONTH1_VAL As Long
Dim MONTH2_VAL As Long

Dim YEAR1_VAL As Long
Dim YEAR2_VAL As Long

On Error GoTo ERROR_LABEL

If START_DATE = 0 Then: GoTo ERROR_LABEL
If END_DATE = 0 Then: GoTo ERROR_LABEL

If END_DATE < START_DATE Then: GoTo ERROR_LABEL

If END_DATE = START_DATE Then
    COUNT_DAYS_FUNC = 0
    Exit Function
End If

If COUNT_BASIS = 1 Or COUNT_BASIS = 2 Or COUNT_BASIS = 3 Then 'Actual
   COUNT_DAYS_FUNC = DateDiff("d", START_DATE, END_DATE) 'END_DATE - START_DATE
   Exit Function
End If

DAY1_VAL = Day(START_DATE)
DAY2_VAL = Day(END_DATE)

MONTH1_VAL = Month(START_DATE)
MONTH2_VAL = Month(END_DATE)

YEAR1_VAL = Year(START_DATE)
YEAR2_VAL = Year(END_DATE)

Select Case COUNT_BASIS
Case 0 'us (nasd) 30/360
   If DAY1_VAL = 31 Then DAY1_VAL = 30
   If DAY2_VAL = 31 And DAY1_VAL = 30 Then DAY2_VAL = 30
Case Else '4 'Europe 30
   If DAY1_VAL = 31 Then DAY1_VAL = 30
   If DAY2_VAL = 31 Then DAY2_VAL = 30
End Select

COUNT_DAYS_FUNC = (YEAR2_VAL - YEAR1_VAL) * 360 + (MONTH2_VAL - MONTH1_VAL) * 30 + _
                  (DAY2_VAL - DAY1_VAL)

Exit Function
ERROR_LABEL:
COUNT_DAYS_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : EDATE_FUNC
'DESCRIPTION   : Returns the serial number that represents the date that
'is the indicated number of months before or after a specified date
'(the start_date). Use EDATE to calculate maturity dates or due dates
'that fall on the same day of the month as the date of issue.
'LIBRARY       : DATE
'GROUP         : COUNT
'ID            : 002
'LAST UPDATE   : 11 / 02 / 2004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function EDATE_FUNC(Optional ByVal DATE_VAL As Date = 0, _
Optional ByVal months As Variant = 1)

'DATE_VAL: is a date that represents the start date. Dates should
'be entered by using the DATE function, or as results of other formulas
'or functions. For example, use DATE(2008,5,23) for the 23rd day of
'May, 2008. Problems can occur if dates are entered as text.

'Months: is the number of months before or after DATE_VAL. A positive
'value for months yields a future date; a negative value yields a past date.

On Error GoTo ERROR_LABEL

If DATE_VAL = 0 Then
    DATE_VAL = DateSerial(Year(Now), Month(Now), Day(Now))
End If
    
EDATE_FUNC = DateAdd("m", months, DATE_VAL)

Exit Function
ERROR_LABEL:
EDATE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : YEARFRAC_FUNC
'DESCRIPTION   : RETURNS THE YEAR FRACTION
'LIBRARY       : DATE
'GROUP         : DAYS
'ID            : 003
'LAST UPDATE   : 11 / 02 / 2004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function YEARFRAC_FUNC(ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal COUNT_BASIS As Integer = 0)

On Error GoTo ERROR_LABEL

'-----------------------------------------------------------------------------

'Actual/actual (in periods), is used for U.S. Treasury bonds, 30/360 is
'used for U.S. corporate and municipal bonds, and actual / 360 is used for
'U.S. Treasury bills and other money market instruments.

'-----------------------------------------------------------------------------

If START_DATE = 0 Then: GoTo ERROR_LABEL
If END_DATE = 0 Then: GoTo ERROR_LABEL
If START_DATE > END_DATE Then: GoTo ERROR_LABEL

If END_DATE = START_DATE Then
  YEARFRAC_FUNC = 0
  Exit Function
End If

YEARFRAC_FUNC = COUNT_DAYS_FUNC(START_DATE, END_DATE, COUNT_BASIS) _
                / DAYS_PER_YEAR_FUNC(START_DATE, COUNT_BASIS)


Exit Function
ERROR_LABEL:
YEARFRAC_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : DAYS_PER_YEAR_FUNC
'DESCRIPTION   : Return the number of days in a year
'LIBRARY       : DATE
'GROUP         : COUNT
'ID            : 004
'LAST UPDATE   : 11 / 02 / 2004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function DAYS_PER_YEAR_FUNC(Optional ByVal DATE_VAL As Date, _
Optional ByVal COUNT_BASIS As Integer = 1)

On Error GoTo ERROR_LABEL
  
'--------------------------------------------------------------------------
Select Case COUNT_BASIS
'--------------------------------------------------------------------------
Case 0, 2, 4 '360
'--------------------------------------------------------------------------
    DAYS_PER_YEAR_FUNC = 360
'--------------------------------------------------------------------------
Case 1 'Actual --> Fix This one
'--------------------------------------------------------------------------
    If DATE_VAL = 0 Then
        DAYS_PER_YEAR_FUNC = 365
    Else
        DAYS_PER_YEAR_FUNC = IIf(IS_DATE_LEAP_YEAR_FUNC(DATE_VAL), 365.25, 365)
    End If
'--------------------------------------------------------------------------
Case 3
'--------------------------------------------------------------------------
    DAYS_PER_YEAR_FUNC = 365 '365
'--------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
DAYS_PER_YEAR_FUNC = Err.number
End Function

