Attribute VB_Name = "WEB_NUMBER_STRIP_LIBR"



Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : STRIP_NUMERICS_FUNC
'DESCRIPTION   :  Handle non-numeric observations (N/A and Null values).
'LIBRARY       : STATISTICS
'GROUP         : VALIDATION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************


Function STRIP_NUMERICS_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim DATA_VECTOR As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG

If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

NROWS = UBound(DATA_VECTOR, 1)
ReDim ATEMP_VECTOR(1 To NROWS, 1 To 1)
j = 0
For i = 1 To NROWS
    If IsNumeric(DATA_VECTOR(i, 1)) And Not IsEmpty(DATA_VECTOR(i, 1)) Then
        j = j + 1
        ATEMP_VECTOR(j, 1) = DATA_VECTOR(i, 1)
    End If
Next i
ReDim BTEMP_VECTOR(1 To j, 1 To 1)
For i = 1 To j
    BTEMP_VECTOR(i, 1) = ATEMP_VECTOR(i, 1)
Next i
    
STRIP_NUMERICS_FUNC = BTEMP_VECTOR

Exit Function
ERROR_LABEL:
STRIP_NUMERICS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : STRIP_NUMBER_STRING_FUNC
'DESCRIPTION   : Function to compare strings, either by stripping
'all non-numeric characters, or by stripping all non-alpha characters.
'Useful in queries when having to compare strings that are functionally
'the same, but typed differently, e.g., with hyphens or commas.
'LIBRARY       : STRING
'GROUP         : STRIP
'ID            : 001
'UPDATE        : 02/20/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function STRIP_NUMBER_STRING_FUNC(ByRef DATA_STR As String, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long

Dim TEMP_STR As String
Dim TEMP_CHR As String

Dim BLN_FLAG As Boolean

On Error Resume Next

TEMP_STR = ""
'------------------------------------------------------------------------------
Select Case VERSION
'------------------------------------------------------------------------------
Case 0 'Strip All Non Numeric Chars
    For j = Len(DATA_STR) To 1 Step -1
        BLN_FLAG = True
        TEMP_CHR = Mid(DATA_STR, j, 1)
        For i = Asc("0") To Asc("9")
            If TEMP_CHR = Chr(i) Then
                BLN_FLAG = False
                Exit For
            End If
        Next i
        If BLN_FLAG = False Then: TEMP_STR = TEMP_CHR & TEMP_STR
    Next j
'------------------------------------------------------------------------------
Case Else 'Strip All Non Alpha Numeric Chars
'------------------------------------------------------------------------------
    For j = Len(DATA_STR) To 1 Step -1
        BLN_FLAG = True
        TEMP_CHR = Mid(DATA_STR, j, 1)
        For i = Asc("A") To Asc("Z")
            If TEMP_CHR = Chr(i) Then
                BLN_FLAG = False
                Exit For
            End If
        Next i
        If BLN_FLAG = False Then: TEMP_STR = TEMP_CHR & TEMP_STR
    Next j
'------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------

STRIP_NUMBER_STRING_FUNC = TEMP_STR

Exit Function
ERROR_LABEL:
STRIP_NUMBER_STRING_FUNC = Err.number
End Function
