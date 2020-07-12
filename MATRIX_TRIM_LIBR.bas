Attribute VB_Name = "MATRIX_TRIM_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRIM_SMALL_VALUES_FUNC
'DESCRIPTION   : Eliminate small values in a matrix
'LIBRARY       : MATRIX
'GROUP         : TRIM
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TRIM_SMALL_VALUES_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal epsilon As Double = 10 ^ -14)

Dim i As Long
Dim j As Long
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

For i = 1 To UBound(DATA_MATRIX, 1)
    For j = 1 To UBound(DATA_MATRIX, 2)
        If IsNumeric(DATA_MATRIX(i, j)) Then
            If Abs(DATA_MATRIX(i, j)) < epsilon Then DATA_MATRIX(i, j) = 0
        End If
    Next j
Next i

MATRIX_TRIM_SMALL_VALUES_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_TRIM_SMALL_VALUES_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_TRIM_FUNC
'DESCRIPTION   : Eliminate empty cells or Z values from a vector and resize
'the vector
'LIBRARY       : MATRIX
'GROUP         : TRIM
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_TRIM_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal REF_VAL As Variant = 0)

Dim i As Long
Dim k As Long

Dim SROW As Long
Dim NROWS As Long

Dim DATA_VECTOR As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

SROW = LBound(DATA_VECTOR, 1)
NROWS = UBound(DATA_VECTOR, 1)

    k = 0
    
    ReDim TEMP_VECTOR(SROW To NROWS, 1 To 1)
    
    For i = SROW To NROWS
        If (IsEmpty(DATA_VECTOR(i, 1)) = True) Or _
        (Trim(DATA_VECTOR(i, 1)) = "") Or _
        (DATA_VECTOR(i, 1) = REF_VAL) _
        Then
          k = k + 1
        Else: TEMP_VECTOR(i - k, 1) = DATA_VECTOR(i, 1)
        End If
    Next i
     
    ReDim DATA_VECTOR(SROW To NROWS - k, 1 To 1)
    For i = SROW To NROWS - k
         DATA_VECTOR(i, 1) = TEMP_VECTOR(i, 1)
    Next i

VECTOR_TRIM_FUNC = DATA_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_TRIM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRIM_FUNC
'DESCRIPTION   : Eliminate empty cells or Z values from a matrix and resize the matrix
'(use column X as base vector)
'LIBRARY       : MATRIX
'GROUP         : TRIM
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TRIM_FUNC(ByRef DATA_MATRIX As Variant, _
Optional ByVal ACOLUMN As Long = 1, _
Optional ByVal REF_VAL As Variant = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

'Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'DATA_MATRIX = DATA_RNG
SROW = LBound(DATA_MATRIX, 1)
SCOLUMN = LBound(DATA_MATRIX, 2)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

k = 0
ReDim TEMP_MATRIX(SROW To NROWS, SCOLUMN To NCOLUMNS)
For i = SROW To NROWS
    If (IsEmpty(DATA_MATRIX(i, ACOLUMN)) = True) Or _
        (Trim(DATA_MATRIX(i, ACOLUMN)) = "") Or _
        (DATA_MATRIX(i, ACOLUMN) = REF_VAL) Then
        k = k + 1
    Else
        For j = SCOLUMN To NCOLUMNS
            TEMP_MATRIX(i - k, j) = DATA_MATRIX(i, j)
        Next j
    End If
Next i
     
ReDim DATA_MATRIX(SROW To NROWS - k, SCOLUMN To NCOLUMNS)
For i = SROW To NROWS - k
    For j = SCOLUMN To NCOLUMNS
        DATA_MATRIX(i, j) = TEMP_MATRIX(i, j)
    Next j
Next i

MATRIX_TRIM_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_TRIM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_TRIM_NULL_FUNC

'DESCRIPTION   : Function to fill one dimensional range (only column or row) into
'into an array and to SHORT_FLAG it if it contains spaces or error values (including
'N/A at the the end).

'LIBRARY       : MATRIX
'GROUP         : TRIM
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_TRIM_NULL_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim NROWS As Long

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

NROWS = UBound(DATA_VECTOR, 1)

NROWS = UBound(DATA_VECTOR, 1)
Do While IsEmpty(DATA_VECTOR(NROWS, 1)) _
    Or IsError(DATA_VECTOR(NROWS, 1)) _
    Or IsNull(DATA_VECTOR(NROWS, 1))
    NROWS = NROWS - 1
Loop

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    TEMP_VECTOR(i, 1) = DATA_VECTOR(i, 1)
Next i

VECTOR_TRIM_NULL_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_TRIM_NULL_FUNC = Err.number
End Function
