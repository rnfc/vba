Attribute VB_Name = "MATRIX_DIAGONAL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_DIAGONAL_VECTOR_FUNC
'DESCRIPTION   : This function extracts the diagonals from a matrix
'and returns a vector.

'LIBRARY       : MATRIX
'GROUP         : DIAGONAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_DIAGONAL_VECTOR_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)
    
Dim i As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NCOLUMNS = UBound(DATA_MATRIX, 2)
NROWS = UBound(DATA_MATRIX, 1)

If NCOLUMNS <> NROWS Then: GoTo ERROR_LABEL

Select Case VERSION
Case 0
    ReDim TEMP_VECTOR(1 To NCOLUMNS) 'SAME AS = (1 to 1, 1 to NCOLUMNS)
    For i = 1 To NCOLUMNS: TEMP_VECTOR(i) = DATA_MATRIX(i, i): Next i
Case Else
    ReDim TEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NCOLUMNS: TEMP_VECTOR(i, 1) = DATA_MATRIX(i, i): Next i
End Select

MATRIX_DIAGONAL_VECTOR_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
MATRIX_DIAGONAL_VECTOR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_DIAGONAL_MATRIX_FUNC
'DESCRIPTION   : Returns to standard (expanded) form the diagonal matrix having the
'vector "d" as its diagonal

'LIBRARY       : MATRIX
'GROUP         : DIAGONAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_DIAGONAL_MATRIX_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim NSIZE As Long

Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

NSIZE = UBound(DATA_VECTOR, 1)

ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
For i = 1 To NSIZE
    For j = 1 To NSIZE
        If i = j Then
            TEMP_MATRIX(i, i) = DATA_VECTOR(i, 1)
        Else
            TEMP_MATRIX(i, j) = 0
        End If
    Next j
Next i

VECTOR_DIAGONAL_MATRIX_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
VECTOR_DIAGONAL_MATRIX_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_DIAGONAL_CONVERT_FUNC
'DESCRIPTION   : This function extracts the diagonals from a matrix
'and returns a vector. Output vector can be a Matrix row
'or column as well

'LIBRARY       : MATRIX
'GROUP         : DIAGONAL
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_DIAGONAL_CONVERT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 1)

'Optional VERSION:0 first diagonal , VERSION:1 secondary diagonal

Dim i As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

If VERSION = 0 Then: VERSION = 1  'for empty cell
'choose the min between rows and columns COUNTER

NSIZE = NROWS
If NROWS > NCOLUMNS Then NSIZE = NCOLUMNS
    
ReDim TEMP_VECTOR(1 To NSIZE, 1 To 1)
    
If VERSION = 1 Then 'get the first diagonal
    For i = 1 To NSIZE
        TEMP_VECTOR(i, 1) = DATA_MATRIX(i, i)
    Next i
Else 'get the secondary diagonal
    For i = 1 To NSIZE
        TEMP_VECTOR(i, 1) = DATA_MATRIX(i, NCOLUMNS - i + 1)
    Next i
End If
    
MATRIX_DIAGONAL_CONVERT_FUNC = TEMP_VECTOR
  
Exit Function
ERROR_LABEL:
MATRIX_DIAGONAL_CONVERT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_DIAGONAL_ERROR_FUNC
'DESCRIPTION   : Returns the mean of all absolute values out of the first
'diagonal of a square matrix return the error for the diagonal matrix
'LIBRARY       : MATRIX
'GROUP         : DIAGONAL
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_DIAGONAL_ERROR_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim TEMP_SUM As Double
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)

TEMP_SUM = 0
For i = 1 To NROWS
    For j = 1 To NROWS
        If i <> j Then TEMP_SUM = _
        TEMP_SUM + Abs(DATA_MATRIX(i, j))
    Next j
Next i

MATRIX_DIAGONAL_ERROR_FUNC = TEMP_SUM / (NROWS ^ 2 - NROWS)

Exit Function
ERROR_LABEL:
MATRIX_DIAGONAL_ERROR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_TRACE_FUNC
'DESCRIPTION   : Returns the trace of a matrix (sum of elements on
'leading diagonal)

'LIBRARY       : MATRIX
'GROUP         : DIAGONAL
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_TRACE_FUNC(ByRef DATA_RNG As Variant)
    
    Dim i As Long
    Dim NCOLUMNS As Long
    Dim TEMP_SUM As Double
    
    Dim DATA_MATRIX As Variant
    
    On Error GoTo ERROR_LABEL
    
    DATA_MATRIX = DATA_RNG
    NCOLUMNS = UBound(DATA_MATRIX, 2)
    
    TEMP_SUM = 0
    For i = 1 To NCOLUMNS
        TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, i)
    Next i
    
    MATRIX_TRACE_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
MATRIX_TRACE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_HESSENBERG_FUNC
'DESCRIPTION   : As known, a matrix is in Hessenberg form if all zz values under
'the lower sub diagonal are zero.
'LIBRARY       : MATRIX
'GROUP         : DIAGONAL
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_HESSENBERG_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long

Dim XTEMP_VALUE As Double
Dim YTEMP_VALUE As Double
Dim ZTEMP_VALUE As Double

Dim ADATA_MATRIX As Variant
Dim BDATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

ADATA_MATRIX = DATA_RNG
NSIZE = UBound(ADATA_MATRIX, 1)

For k = 2 To NSIZE - 1
    i = k
    XTEMP_VALUE = 0
    For j = k To NSIZE
        If (Abs(ADATA_MATRIX(j, k - 1)) > Abs(XTEMP_VALUE)) Then
            XTEMP_VALUE = ADATA_MATRIX(j, k - 1)
            i = j
        End If
    Next j
    If (i <> k) Then
        For j = k - 1 To NSIZE
            ZTEMP_VALUE = ADATA_MATRIX(i, j)
            ADATA_MATRIX(i, j) = ADATA_MATRIX(k, j)
            ADATA_MATRIX(k, j) = ZTEMP_VALUE
        Next j
        For j = 1 To NSIZE
            ZTEMP_VALUE = ADATA_MATRIX(j, i)
            ADATA_MATRIX(j, i) = ADATA_MATRIX(j, k)
            ADATA_MATRIX(j, k) = ZTEMP_VALUE
        Next j
    End If
    
    If (XTEMP_VALUE <> 0) Then
        For i = k + 1 To NSIZE
            YTEMP_VALUE = ADATA_MATRIX(i, k - 1)
            If (YTEMP_VALUE <> 0) Then
                YTEMP_VALUE = YTEMP_VALUE / XTEMP_VALUE
                ADATA_MATRIX(i, k - 1) = YTEMP_VALUE
                For j = k To NSIZE
                    ADATA_MATRIX(i, j) = ADATA_MATRIX(i, j) - _
                    YTEMP_VALUE * ADATA_MATRIX(k, j)
                Next j
                For j = 1 To NSIZE
                     ADATA_MATRIX(j, k) = ADATA_MATRIX(j, k) + _
                     YTEMP_VALUE * ADATA_MATRIX(j, i)
                Next j
            End If
        Next i
    End If
Next k

ReDim BDATA_MATRIX(1 To NSIZE, 1 To NSIZE)
For i = 1 To NSIZE
    For j = 1 To NSIZE
        If j >= i - 1 Then BDATA_MATRIX(i, j) = ADATA_MATRIX(i, j)
    Next j
Next i

MATRIX_HESSENBERG_FUNC = BDATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_HESSENBERG_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CHARACTERISTIC_FUNC
'DESCRIPTION   : Returns the characteristic matrix at the value lambda
'LIBRARY       : MATRIX
'GROUP         : DIAGONAL
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CHARACTERISTIC_FUNC(ByRef DATA_RNG As Variant, _
ByVal LAMBDA As Double)

Dim i As Long
Dim NSIZE As Long
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NSIZE = MINIMUM_FUNC(UBound(DATA_MATRIX, 1), UBound(DATA_MATRIX, 2))

For i = 1 To NSIZE
    DATA_MATRIX(i, i) = DATA_MATRIX(i, i) - LAMBDA
Next i

MATRIX_CHARACTERISTIC_FUNC = DATA_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_CHARACTERISTIC_FUNC = Err.number
End Function
