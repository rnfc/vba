Attribute VB_Name = "WEB_NUMBER_DECIMALS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : GET_DECIMALS_FUNC
'DESCRIPTION   : Extract Decimals from a real number
'LIBRARY       : EVALUATE
'GROUP         : DECIMAL
'ID            : 001
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function GET_DECIMALS_FUNC(ByVal DATA_VAL As Variant)

Dim i As Long
Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL

TEMP_VAL = DATA_VAL - Fix(DATA_VAL)
i = Int(Log(Abs(DATA_VAL)) / Log(10)) + 1 'integer digits
TEMP_VAL = CDec(Round(TEMP_VAL, 15 - i))

GET_DECIMALS_FUNC = TEMP_VAL

Exit Function
ERROR_LABEL:
GET_DECIMALS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : DECIMAL_SEPARATOR_FUNC
'DESCRIPTION   : Return the current environment setting for decimal separator
'about 2-3 us, that is 20 times faster than
'Excel.Application.International(xlDecimalSeparator)
'LIBRARY       : NUMBERS
'GROUP         : DECIMAL
'ID            : 002
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function DECIMAL_SEPARATOR_FUNC()
On Error GoTo ERROR_LABEL
DECIMAL_SEPARATOR_FUNC = Mid(CStr(1 / 2), 2, 1)
Exit Function
ERROR_LABEL:
DECIMAL_SEPARATOR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : REMOVE_DECIMALS_NUMBER_FUNC
'DESCRIPTION   : Substitute decimal point from symb1 to symb2 (1.8 us)
'the bulti-in function Replace takes about 5.5 (us)
'LIBRARY       : NUMBERS
'GROUP         : DECIMAL
'ID            : 003
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function REMOVE_DECIMALS_NUMBER_FUNC(ByVal DATA_VAL As Variant)

Dim i As Long
Dim DATA_STR As String

On Error GoTo ERROR_LABEL
DATA_STR = DATA_VAL
i = InStr(1, DATA_STR, DECIMAL_SEPARATOR_FUNC())

If i > 0 Then
    DATA_STR = Mid(DATA_STR, 1, i - 1)
    REMOVE_DECIMALS_NUMBER_FUNC = CDec(DATA_STR)
Else
    REMOVE_DECIMALS_NUMBER_FUNC = CDec(DATA_STR)
End If

Exit Function
ERROR_LABEL:
REMOVE_DECIMALS_NUMBER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COUNT_DECIMALS_NUMBER_FORMAT_FUNC
'DESCRIPTION   : 'Finds the number of decimals in a number format string
'NUMBER_FORMAT = Cell.NumberFormat; or number
'COUNT_DECIMALS_NUMBER_FORMAT_FUNC("#,##0.000;[Red]#,##0.000") = 3
'LIBRARY       : NUMBERS
'GROUP         : DECIMAL
'ID            : 004
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function COUNT_DECIMALS_NUMBER_FORMAT_FUNC(ByVal NUMBER_FORMAT As String)
    
Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP_DEC As String
Dim TEMP_CHR As String
Dim TEMP_LIKE As String

On Error GoTo ERROR_LABEL

TEMP_CHR = DECIMAL_SEPARATOR_FUNC()
TEMP_LIKE = "*" & TEMP_CHR & "*"

TEMP_DEC = NUMBER_FORMAT

If TEMP_DEC Like TEMP_LIKE Then
    k = Len(TEMP_DEC)
    For i = InStr(TEMP_DEC, TEMP_CHR) + 1 To k
        If Mid(TEMP_DEC, i, 1) Like "#" Then
        'If the characters following the decimal point are either 0 or the
            j = j + 1
            'number sign, "#", count them.
        Else
            COUNT_DECIMALS_NUMBER_FORMAT_FUNC = j
            Exit Function
        End If
    Next i
    COUNT_DECIMALS_NUMBER_FORMAT_FUNC = j
Else 'If there is no decimal point, the number of decimals is null.
    COUNT_DECIMALS_NUMBER_FORMAT_FUNC = 0
End If

Exit Function
ERROR_LABEL:
COUNT_DECIMALS_NUMBER_FORMAT_FUNC = Err.number
End Function

