Attribute VB_Name = "DATE_HOLIDAYS_LIBR"
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : HOLIDAYS_COUNT_FUNC
'DESCRIPTION   : Holidays COUNTER between 2 Dates
'LIBRARY       : DATE
'GROUP         : HOLIDAYS
'ID            : 001

'LAST UPDATE   : 11 / 02 / 2004

'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function HOLIDAYS_COUNT_FUNC(ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByRef HOLIDAYS_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim HOLIDAYS_VECTOR As Variant

'HOLIDAYS_RNG: is a list of one or more dates to
'exclude from the working calendar, such as state and federal
'holidays and floating holidays. The list can be either a range
'of cells that contain the dates or an array constant of the serial
'numbers that represent the dates.

On Error GoTo ERROR_LABEL

HOLIDAYS_VECTOR = HOLIDAYS_RNG
If UBound(HOLIDAYS_VECTOR, 1) = 1 Then
    HOLIDAYS_VECTOR = MATRIX_TRANSPOSE_FUNC(HOLIDAYS_VECTOR)
End If

NROWS = UBound(HOLIDAYS_VECTOR, 1)
j = 0
For i = 1 To NROWS
    If IsDate(HOLIDAYS_VECTOR(i, 1)) = True Then
        If (HOLIDAYS_VECTOR(i, 1) >= START_DATE) And _
           (HOLIDAYS_VECTOR(i, 1) <= END_DATE) _
            And (Weekday(HOLIDAYS_VECTOR(i, 1)) <> 1) And _
            (Weekday(HOLIDAYS_VECTOR(i, 1)) <> 7) Then
                j = j + 1
        End If
    End If
Next i

HOLIDAYS_COUNT_FUNC = j

Exit Function
ERROR_LABEL:
HOLIDAYS_COUNT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : THANKSGIVING_DATE_FUNC
'DESCRIPTION   : Returns the date of Thanksgiving for the given year
'LIBRARY       : DATE
'GROUP         : HOLIDAYS
'ID            : 004

'LAST UPDATE   : 11 / 02 / 2004

'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function THANKSGIVING_DATE_FUNC(ByVal YEAR_INT As Integer)

On Error GoTo ERROR_LABEL

THANKSGIVING_DATE_FUNC = _
DateSerial(YEAR_INT, 11, 29 - Weekday(DateSerial(YEAR_INT, 11, 1), vbFriday))

Exit Function
ERROR_LABEL:
THANKSGIVING_DATE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EASTER_DATE_FUNC
'DESCRIPTION   : Calculate the correct date for EASTER_DATE Sunday between the
'years 1900 and 2078
'LIBRARY       : DATE
'GROUP         : HOLIDAYS
'ID            : 003
'LAST UPDATE   : 11 / 02 / 2004

'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function EASTER_DATE_FUNC(ByVal YEAR_INT As Integer)

Dim k As Integer

On Error GoTo ERROR_LABEL

'DateSerial(YEAR_INT, 1, 1) 'New Year's day
'DateSerial(YEAR_INT, 5, 1) 'May day
'DateSerial(YEAR_INT, 12, 25) 'Xmas day
'DateSerial(YEAR_INT, 12, 26) ' Boxing day
'EASTER_DATE_FUNC(YEAR_INT) - 2 ' Good Friday
'HOLIDAYS_VECTOR(5, 1) + 3 ' EASTER_DATE Monday

k = (((255 - 11 * (YEAR_INT Mod 19)) - 21) Mod 30) + 21

EASTER_DATE_FUNC = DateSerial(YEAR_INT, 3, 1) + k + (k > 48) _
            + 6 - ((YEAR_INT + YEAR_INT \ 4 + _
            k + (k > 48) + 1) Mod 7)

Exit Function
ERROR_LABEL:
EASTER_DATE_FUNC = Err.number
End Function
