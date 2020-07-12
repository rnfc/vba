Attribute VB_Name = "WEB_NUMBER_FACTORIAL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_FACTORIAL_FORMULA_FUNC
'DESCRIPTION   : Convert a FORMULA with ! symbol with Fact() function
'LIBRARY       : STRING
'GROUP         : FACTORIAL
'ID            : 001
'UPDATE        : 01/21/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function CONVERT_FACTORIAL_FORMULA_FUNC(ByVal DATA_STR As Variant, _
Optional ByVal FUNC_NAME_STR As String = "FACT_FUNC")

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim ATEMP_STR As String
Dim BTEMP_STR As String
Dim CTEMP_STR As String
Dim DTEMP_STR As String

On Error GoTo ERROR_LABEL

DTEMP_STR = DATA_STR
k = InStr(1, DTEMP_STR, "!")
If k = 0 Then: GoTo ERROR_LABEL

Do Until k = 0
    j = 0
    i = k - 1
    h = 0
    Do
        Select Case Mid(DTEMP_STR, i, 1)
            Case "(", "[", "{"
                j = j - 1
            Case ")", "]", "}"
                j = j + 1
            Case "*", "/", "+", "-", "^"
                If j = 0 Then
                    h = i + 1: Exit Do
                End If
            Case Else 'nothig to do
        End Select
        i = i - 1
        If i = 0 Then h = 1 '.end of the string
    Loop Until i = 0
    
    ATEMP_STR = Mid(DTEMP_STR, h, k - h)  'argument of factorial function
    BTEMP_STR = Left(DTEMP_STR, h - 1)
    CTEMP_STR = Right(DTEMP_STR, Len(DTEMP_STR) - k)
    DTEMP_STR = BTEMP_STR & FUNC_NAME_STR & "(" & ATEMP_STR & ")" & CTEMP_STR
    k = InStr(1, DTEMP_STR, "!", 0)
Loop

CONVERT_FACTORIAL_FORMULA_FUNC = DTEMP_STR

Exit Function
ERROR_LABEL:
CONVERT_FACTORIAL_FORMULA_FUNC = Err.number
End Function
