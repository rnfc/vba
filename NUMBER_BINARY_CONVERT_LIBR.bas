Attribute VB_Name = "NUMBER_BINARY_CONVERT_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    

'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_BINARY_INTEGER_FUNC
'DESCRIPTION   : Convert binary to integer
'LIBRARY       : NUMBER_BINARY
'GROUP         : CONVERT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function CONVERT_BINARY_INTEGER_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(DATA_RNG) = True Then
    DATA_MATRIX = DATA_RNG
    For i = 1 To UBound(DATA_MATRIX, 1)
        For j = 1 To UBound(DATA_MATRIX, 2)
            DATA_MATRIX(i, j) = CInt(DATA_MATRIX(i, j))
        Next j
    Next i
Else
    DATA_MATRIX = CInt(DATA_RNG)
End If

CONVERT_BINARY_INTEGER_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
CONVERT_BINARY_INTEGER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_BINARY_STRING_FUNC
'DESCRIPTION   : Converte a binary vector into a string binary
'LIBRARY       : NUMBER_BINARY
'GROUP         : CONVERT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function CONVERT_BINARY_STRING_FUNC(ByRef DATA_RNG As Variant)

Dim j As Long
Dim NCOLUMNS As Long
Dim TEMP_STR As String
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 2) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

NCOLUMNS = UBound(DATA_VECTOR, 2)
TEMP_STR = ""
For j = 1 To NCOLUMNS
    TEMP_STR = TEMP_STR & CStr(DATA_VECTOR(1, j))
Next j

CONVERT_BINARY_STRING_FUNC = TEMP_STR
Exit Function
ERROR_LABEL:
CONVERT_BINARY_STRING_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_BINARY_VECTOR_FUNC
'DESCRIPTION   : Convert a binary string into a binary vector
'LIBRARY       : NUMBER_BINARY
'GROUP         : CONVERT
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

'// PERFECT

Function CONVERT_BINARY_VECTOR_FUNC(ByVal BINARY_STR As String)

Dim i As Long
Dim j As Long
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

j = Len(BINARY_STR)
ReDim DATA_VECTOR(1 To 1, 1 To j)
For i = 1 To j
    DATA_VECTOR(1, i) = CInt(Mid(BINARY_STR, i, 1))
Next i

CONVERT_BINARY_VECTOR_FUNC = DATA_VECTOR

Exit Function
ERROR_LABEL:
CONVERT_BINARY_VECTOR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_BINARY_DECIMAL_FUNC
'DESCRIPTION   : Converts binary into decimal up to 15 digits
'LIBRARY       : NUMBER_BINARY
'GROUP         : CONVERT
'ID            : 004
'LAST UPDATE   : 12/08/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function CONVERT_BINARY_DECIMAL_FUNC(ByVal BINARY_STR As String)

Dim i As Long
Dim j As Long

Dim ATEMP_STR As String
Dim BTEMP_STR As String

Dim TEMP_SUM As Double
Dim TEMP_MULT As Double

On Error GoTo ERROR_LABEL

If BINARY_STR = "0" Then
    CONVERT_BINARY_DECIMAL_FUNC = 0
    Exit Function
End If

ATEMP_STR = BINARY_STR
TEMP_MULT = 1
TEMP_SUM = 0

j = Len(ATEMP_STR)

For i = j To 1 Step -1
    BTEMP_STR = Mid(ATEMP_STR, i, 1)
    If BTEMP_STR = "1" Then TEMP_SUM = TEMP_SUM + TEMP_MULT
    TEMP_MULT = 2 * TEMP_MULT
Next i

CONVERT_BINARY_DECIMAL_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
CONVERT_BINARY_DECIMAL_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_DECIMAL_BINARY_FUNC
'DESCRIPTION   : Converts decimal into binary up to 50 digits
'LIBRARY       : NUMBER_BINARY
'GROUP         : CONVERT
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

'// PERFECT

Function CONVERT_DECIMAL_BINARY_FUNC(ByVal X_VAL As Double, _
Optional ByVal DIGITS As Integer = 2)

Dim i As Long
Dim j As Long
Dim k As Long

Dim TEMP_MULT As Double
Dim TEMP_DELTA As Double

Dim ATEMP_STR As String
Dim BTEMP_STR As String

On Error GoTo ERROR_LABEL

If X_VAL = 0 Then
    ATEMP_STR = "0"
Else
    TEMP_DELTA = X_VAL
    ATEMP_STR = ""
    j = Int(Log(TEMP_DELTA) / Log(2))
    TEMP_MULT = 2 ^ j
    For i = j To 0 Step -1
        If TEMP_DELTA >= TEMP_MULT Then
            TEMP_DELTA = TEMP_DELTA - TEMP_MULT
            BTEMP_STR = "1"
        Else
            BTEMP_STR = "0"
        End If
        If BTEMP_STR = 1 Or Len(ATEMP_STR) > 0 Then
            ATEMP_STR = ATEMP_STR & BTEMP_STR
        End If
        TEMP_MULT = TEMP_MULT / 2
    Next
End If

If DIGITS <> 0 Then
    k = DIGITS - Len(ATEMP_STR)
    If k > 0 Then ATEMP_STR = String(k, "0") & ATEMP_STR
End If

CONVERT_DECIMAL_BINARY_FUNC = ATEMP_STR
Exit Function
ERROR_LABEL:
CONVERT_DECIMAL_BINARY_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CONVERT_VECTOR_BINARY_FUNC
'DESCRIPTION   : Converts a binary function from binary-vector form
'LIBRARY       : NUMBER_BINARY
'GROUP         : CONVERT
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function CONVERT_VECTOR_BINARY_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CHR_STR As String = "x")

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_STR As String

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NROWS, 1 To 2)
If NCOLUMNS = 2 Then
    'decimal format input
    If VarType(DATA_MATRIX(1, 1)) = vbString Then
        For i = 1 To NROWS
            TEMP_MATRIX(i, 1) = CONVERT_BINARY_DECIMAL_FUNC(DATA_MATRIX(i, 1))
            TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
        Next i
    Else
        For i = 1 To NROWS
            TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
            TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
        Next i
    End If
Else
    'vector-binary format input
    For i = 1 To NROWS
        TEMP_STR = ""
        For j = 1 To NCOLUMNS - 1
            TEMP_STR = TEMP_STR & CStr(DATA_MATRIX(i, j))
        Next j
        TEMP_MATRIX(i, 1) = CONVERT_BINARY_DECIMAL_FUNC(TEMP_STR)
        TEMP_MATRIX(i, 2) = DATA_MATRIX(i, NCOLUMNS)
    Next i
End If

TEMP_MATRIX = BINARY_SORT_FUNC(TEMP_MATRIX, 1, 1)

k = BINARY_VARIABLES_COUNTER_FUNC(TEMP_MATRIX(NROWS, 1))
NSIZE = 2 ^ k
ReDim TEMP_VECTOR(1 To NSIZE, 1 To 1)
For i = 1 To NSIZE
    TEMP_VECTOR(i, 1) = CHR_STR
Next i
For i = 1 To NROWS
    TEMP_VECTOR(i, 1) = TEMP_MATRIX(i, 2)
Next i

CONVERT_VECTOR_BINARY_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CONVERT_VECTOR_BINARY_FUNC = Err.number
End Function
