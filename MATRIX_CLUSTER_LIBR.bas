Attribute VB_Name = "MATRIX_CLUSTER_LIBR"

'// PERFECT

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_DISSIMILARITY_FUNC

'DESCRIPTION   : Returns distance/dissimilarity matrix, using euclidian distances
'between objects. A row in a matrix represents an objects, a column a set of
'observations. NSIZE: specifies minimum number of observations needed in order
'to include object

'LIBRARY       : MATRIX
'GROUP         : CLUSTER
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009

'************************************************************************************
'************************************************************************************

Function MATRIX_DISSIMILARITY_FUNC(ByRef DATA_RNG As Variant, _
ByVal NSIZE As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NROWS)

For i = 1 To NROWS - 1
    For j = i + 1 To NROWS
        l = 0
        For k = 1 To NCOLUMNS
            If IsNumeric(DATA_MATRIX(i, k)) And IsNumeric(DATA_MATRIX(j, k)) Then
            ' only calculate distance if numbers
                TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) + (DATA_MATRIX(i, k) - DATA_MATRIX(j, k)) ^ 2
                l = l + 1
            End If
        Next k
        
        If l > NSIZE Then
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) ^ 0.5
        Else
            TEMP_MATRIX(i, j) = "N/A"
        End If
        TEMP_MATRIX(j, i) = TEMP_MATRIX(i, j)
        ' TEMP_MATRIX(j, i) = "N/A"
    Next j
    ' TEMP_MATRIX(i, i) = "N/A"
Next i
' TEMP_MATRIX(NROWS, NROWS) = "N/A"
MATRIX_DISSIMILARITY_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_DISSIMILARITY_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_DISTANCE_MATRIX_FUNC

'DESCRIPTION   : Expresses distance matrix as a vector, including
'only the upper diagonal part

'LIBRARY       : MATRIX
'GROUP         : CLUSTER
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009

'************************************************************************************
'************************************************************************************

Function VECTOR_DISTANCE_MATRIX_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
NSIZE = NROWS * NCOLUMNS

ReDim TEMP_VECTOR(1 To NSIZE, 1 To 1)
k = 1
For i = 1 To NROWS
    For j = i + 1 To NROWS
        If IsNumeric(DATA_MATRIX(i, j)) Then
            TEMP_VECTOR(k, 1) = DATA_MATRIX(i, j)
            k = k + 1
        End If
    Next j
Next i
VECTOR_DISTANCE_MATRIX_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_DISTANCE_MATRIX_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_AVERAGE_DISTANCE_FUNC
'DESCRIPTION   : Calculates vector with average distances per object
'LIBRARY       : MATRIX
'GROUP         : CLUSTER
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009

'************************************************************************************
'************************************************************************************

Function VECTOR_AVERAGE_DISTANCE_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
ReDim TEMP_VECTOR(1 To 1, 1 To NROWS)
For i = 1 To NROWS
    k = 0
    For j = 1 To NROWS
        If IsNumeric(DATA_MATRIX(i, j)) Then
            TEMP_VECTOR(1, i) = TEMP_VECTOR(1, i) + DATA_MATRIX(i, j)
            k = k + 1
        End If
    Next j
    TEMP_VECTOR(1, i) = TEMP_VECTOR(1, i) / k
Next i
VECTOR_AVERAGE_DISTANCE_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_AVERAGE_DISTANCE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_Z_SCORES_FUNC
'DESCRIPTION   : Computes z scores of a data vector
'LIBRARY       : MATRIX
'GROUP         : CLUSTER
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function VECTOR_Z_SCORES_FUNC(ByRef DATA_RNG As Variant)
    
Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MEAN_VAL  As Double
Dim SIGMA_VAL As Double

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) = 1 Then: _
    DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    
    TEMP_VECTOR = MATRIX_GET_COLUMN_FUNC(DATA_MATRIX, j, 1)
    MEAN_VAL = MATRIX_MEAN_FUNC(TEMP_VECTOR)(1, 1)
    SIGMA_VAL = MATRIX_STDEVP_FUNC(TEMP_VECTOR)(1, 1)

    For i = 1 To NROWS
         TEMP_MATRIX(i, j) = (TEMP_VECTOR(i, 1) - MEAN_VAL) / SIGMA_VAL
    Next i
Next j

VECTOR_Z_SCORES_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
VECTOR_Z_SCORES_FUNC = Err.number
End Function
