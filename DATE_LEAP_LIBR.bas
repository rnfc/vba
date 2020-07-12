Attribute VB_Name = "DATE_LEAP_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : IS_DATE_LEAP_YEAR_FUNC
'DESCRIPTION   : Check if the specified Date is a leap year
'LIBRARY       : DATE
'GROUP         : LEAP
'ID            : 001
'LAST UPDATE   : 11 / 02 / 2004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function IS_DATE_LEAP_YEAR_FUNC(ByVal DATE_VAL As Date)
On Error GoTo ERROR_LABEL
    IS_DATE_LEAP_YEAR_FUNC = Month(DateSerial(Year(DATE_VAL), 2, 29)) = 2
Exit Function
ERROR_LABEL:
IS_DATE_LEAP_YEAR_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : LEAP_YEARS_PERIOD_FUNC
'DESCRIPTION   : Leap Years Period Function
'LIBRARY       : DATE
'GROUP         : LEAP
'ID            : 003
'LAST UPDATE   : 11 / 02 / 2004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Function LEAP_YEARS_PERIOD_FUNC(ByVal FIRST_DATE As Date, _
ByVal SECOND_DATE As Date)

Dim i As Integer
Dim j As Integer

On Error GoTo ERROR_LABEL

If FIRST_DATE > SECOND_DATE Then
    LEAP_YEARS_PERIOD_FUNC = -1
    Exit Function
End If

For j = Year(FIRST_DATE) To Year(SECOND_DATE)
    If IS_DATE_LEAP_YEAR_FUNC(j) = True Then
        i = i + 1
    End If
Next j

If FIRST_DATE > DateSerial(Year(FIRST_DATE), 2, 29) And _
    IS_DATE_LEAP_YEAR_FUNC(Year(FIRST_DATE)) = True Then
    i = i - 1
End If

If SECOND_DATE < DateSerial(Year(SECOND_DATE), 2, 29) And _
    IS_DATE_LEAP_YEAR_FUNC(Year(SECOND_DATE)) = True Then
    i = i - 1
End If

LEAP_YEARS_PERIOD_FUNC = i

Exit Function
ERROR_LABEL:
LEAP_YEARS_PERIOD_FUNC = Err.number
End Function
