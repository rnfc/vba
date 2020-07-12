Attribute VB_Name = "WEB_NUMBER_FORMAT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : FORMAT_NUMBER_STRING_FUNC
'DESCRIPTION   : Implementing a Number Format
'LIBRARY       : NUMBERS
'GROUP         : STRING
'ID            : 001
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function FORMAT_NUMBER_STRING_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal FORMAT_STR As String = "0.0")

Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP_STR As String

Dim TEMP_VAL As Double
Dim TEMP_FACTOR As Double

On Error GoTo ERROR_LABEL

If IS_NUMERIC_FUNC(DATA_VAL, DECIMAL_SEPARATOR_FUNC()) = False Then: _
GoTo ERROR_LABEL
k = 100000 'To Avoid Error Traps

TEMP_FACTOR = 1
'-----------------------------------------------------------------
If DATA_VAL > 0 Then
'-----------------------------------------------------------------
    j = 0
    Do While TEMP_FACTOR > DATA_VAL
        TEMP_FACTOR = TEMP_FACTOR / 10
        j = j + 1
        If j = k Then: Exit Do  'To Avoid Error Traps
    Loop

    TEMP_VAL = -(Log(TEMP_FACTOR) / Log(10#)) + 1

    TEMP_STR = FORMAT_STR
    
    For i = 1 To TEMP_VAL
        TEMP_STR = TEMP_STR & Right(FORMAT_STR, 1)
    Next i

    FORMAT_NUMBER_STRING_FUNC = (Format(DATA_VAL, TEMP_STR))

'-----------------------------------------------------------------
ElseIf DATA_VAL < 0 Then
'-----------------------------------------------------------------
    j = 0
    Do While TEMP_FACTOR < DATA_VAL
        TEMP_FACTOR = TEMP_FACTOR / 10
        j = j + 1
        If j = k Then: Exit Do 'To Avoid Error Traps
    Loop

    TEMP_VAL = -(Log(TEMP_FACTOR) / Log(10#)) + 1

    TEMP_STR = FORMAT_STR

    For i = 1 To TEMP_VAL
        TEMP_STR = TEMP_STR & Right(FORMAT_STR, 1)
    Next i

    FORMAT_NUMBER_STRING_FUNC = Format(DATA_VAL, TEMP_STR)
'-----------------------------------------------------------------
Else
'-----------------------------------------------------------------
    FORMAT_NUMBER_STRING_FUNC = Format(DATA_VAL, "0.0")
'-----------------------------------------------------------------
End If
'-----------------------------------------------------------------

Exit Function
ERROR_LABEL:
FORMAT_NUMBER_STRING_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : NUMBER_DECIMAL_FORMAT_FUNC
'DESCRIPTION   : Implementing a Number Decimal Format
'LIBRARY       : NUMBERS
'GROUP         : STRING
'ID            : 002
'LAST UPDATE   : 12/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function NUMBER_DECIMAL_FORMAT_FUNC(ByVal DATA_VAL As Double, _
Optional ByVal FORMAT_NUMBER_STR As String = "0.00000000")

Dim i As Integer
Dim TEMP_STR As String
Dim DEC_SEPAR_STR As String

On Error GoTo ERROR_LABEL
    
    DEC_SEPAR_STR = DECIMAL_SEPARATOR_FUNC()
    TEMP_STR = Format(DATA_VAL, FORMAT_NUMBER_STR)
    i = InStr(1, TEMP_STR, DEC_SEPAR_STR, vbBinaryCompare)
    If i = 0 Then: GoTo ERROR_LABEL
    If DATA_VAL < 1# And DATA_VAL > 0# Then
        TEMP_STR = Mid(TEMP_STR, 1, i) & Right(TEMP_STR, Len(TEMP_STR) - i)
    End If
    
    NUMBER_DECIMAL_FORMAT_FUNC = TEMP_STR
Exit Function
ERROR_LABEL:
NUMBER_DECIMAL_FORMAT_FUNC = Err.number
End Function


