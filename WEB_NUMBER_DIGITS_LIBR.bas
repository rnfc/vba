Attribute VB_Name = "WEB_NUMBER_DIGITS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : INTEGER_COUNT_TOTAL_DIGITS_FUNC
'DESCRIPTION   : counts total digits of an integer number
'LIBRARY       : NUMBERS
'GROUP         : DIGITS
'ID            : 001
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function COUNT_TOTAL_DIGITS_FUNC(ByVal X_VAL As Variant)

On Error GoTo ERROR_LABEL

If X_VAL <> 0 Then
    COUNT_TOTAL_DIGITS_FUNC = (Int(Log(Abs(CDbl(X_VAL))) / Log(10)) + 1)
Else
    COUNT_TOTAL_DIGITS_FUNC = 0
End If

Exit Function
ERROR_LABEL:
COUNT_TOTAL_DIGITS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VALID_DIGITS_FUNC
'DESCRIPTION   : Check if all digits are different: DATA_VAL must be an integer
'LIBRARY       : NUMBERS
'GROUP         : DIGITS
'ID            : 002
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function VALID_DIGITS_FUNC(ByVal DATA_VAL As Variant, _
Optional ByVal NO_DIGITS As Integer = 20)

Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP_VAL As Variant
Dim TEMP_CHR As String
Dim TEMP_STR As String
Dim TEMP_ARR As Variant
    
On Error GoTo ERROR_LABEL

TEMP_VAL = DATA_VAL
ReDim TEMP_ARR(0 To NO_DIGITS)
TEMP_CHR = "*" & DECIMAL_SEPARATOR_FUNC() & "*"

If IS_NUMERIC_FUNC(TEMP_VAL, DECIMAL_SEPARATOR_FUNC()) Then
    If Int(TEMP_VAL) <> TEMP_VAL Then: GoTo ERROR_LABEL
    'TEMP_VAL must be an integer
    
    k = 0
    For i = 0 To NO_DIGITS
        'TEMP_ARR(i) = TEMP_VAL Mod 10
        'bug #VALUE for number=9876543210
        TEMP_ARR(i) = TEMP_VAL - Int(TEMP_VAL / 10) * 10
        TEMP_VAL = (TEMP_VAL - TEMP_ARR(i)) / 10
        If TEMP_VAL > 0 Then k = k + 1
    Next i

    For i = 0 To k - 1
        For j = (i + 1) To k
            If TEMP_ARR(i) = TEMP_ARR(j) Then
                VALID_DIGITS_FUNC = False
                Exit Function
            End If
        Next j
    Next i
    VALID_DIGITS_FUNC = True
Else
    TEMP_STR = TEMP_VAL
    If TEMP_STR Like TEMP_CHR Then: GoTo ERROR_LABEL
    'TEMP_VAL must be an integer
    
    TEMP_STR = REMOVE_CHARACTER_FUNC(TEMP_STR, ",")
    TEMP_STR = REMOVE_CHARACTER_FUNC(TEMP_STR, " ")
    TEMP_STR = REMOVE_CHARACTER_FUNC(TEMP_STR, "(")
    TEMP_STR = REMOVE_CHARACTER_FUNC(TEMP_STR, ")")
    TEMP_STR = REMOVE_CHARACTER_FUNC(TEMP_STR, "-")

    For i = 1 To Len(TEMP_STR) - 1
        For j = (i + 1) To Len(TEMP_STR)
            If Mid(TEMP_STR, i, 1) = Mid(TEMP_STR, j, 1) Then
                VALID_DIGITS_FUNC = False
                Exit Function
            End If
        Next j
    Next i
    VALID_DIGITS_FUNC = True
End If

Exit Function
ERROR_LABEL:
VALID_DIGITS_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SUM_DIGITS_FUNC
'DESCRIPTION   : SUM ALL THE DIGITS: DATA_VAL must be an integer
'LIBRARY       : NUMBERS
'GROUP         : DIGITS
'ID            : 003
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function SUM_DIGITS_FUNC(ByVal DATA_VAL As Variant)
    
Dim i As Long
Dim j As Long

Dim TEMP_CHR As String
Dim TEMP_STR As String

Dim TEMP_VAL As Variant
Dim TEMP_SUM As Double
Dim TEMP_SCALE As Double
Dim TEMP_DIGIT As Double

On Error GoTo ERROR_LABEL

TEMP_VAL = DATA_VAL
TEMP_CHR = "*" & DECIMAL_SEPARATOR_FUNC() & "*"

If IS_NUMERIC_FUNC(TEMP_VAL, DECIMAL_SEPARATOR_FUNC()) Then
    
    If Int(TEMP_VAL) <> TEMP_VAL Then: GoTo ERROR_LABEL
    'TEMP_VAL must be an integer
    
    TEMP_SCALE = Abs(TEMP_VAL)
    TEMP_SUM = 0
    Do Until TEMP_SCALE = 0
        TEMP_DIGIT = TEMP_SCALE Mod 10
        TEMP_SCALE = (TEMP_SCALE - TEMP_DIGIT) / 10
        TEMP_SUM = TEMP_SUM + TEMP_DIGIT
    Loop
Else
    TEMP_STR = TEMP_VAL
    If TEMP_STR Like TEMP_CHR Then: GoTo ERROR_LABEL
    'TEMP_VAL must be an integer
    
    TEMP_STR = REMOVE_CHARACTER_FUNC(TEMP_STR, ",")
    TEMP_STR = REMOVE_CHARACTER_FUNC(TEMP_STR, " ")
    TEMP_STR = REMOVE_CHARACTER_FUNC(TEMP_STR, "(")
    TEMP_STR = REMOVE_CHARACTER_FUNC(TEMP_STR, ")")
    TEMP_STR = REMOVE_CHARACTER_FUNC(TEMP_STR, "-")
    
    TEMP_SUM = 0
    j = Len(TEMP_STR)
    For i = 1 To j
        TEMP_SUM = TEMP_SUM + _
        CONVERT_STRING_LETTERS_NUMBER_FUNC(Mid(TEMP_STR, i, 1))
    Next i
End If

SUM_DIGITS_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
SUM_DIGITS_FUNC = Err.number
End Function
