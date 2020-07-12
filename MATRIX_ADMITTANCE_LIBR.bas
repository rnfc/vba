Attribute VB_Name = "MATRIX_ADMITTANCE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_ADMITTANCE_FUNC

'DESCRIPTION   : Returns the Admittance Matrix of a Linear Passive Network Graph
'A complex admittance has a real part (conductance) and an imaginary part
'(susceptance). This is very useful for solving linear passive network, where
'V is the vector of nodal voltage, I is the vector of nodal current and
'[ Y ] is the admittance matrix.If N+1 is the number of
'nodes, then the matrix dimension will be (N x N). (usually the references
'nodes is set at V = 0)Usually V, I, [ Y ] are complex. The function returns
'an (N x 2*N) array. The first N columns contain the real part and the
'last N columns contain the immaginary part. If all branch-admittances are
'real, the matrix will be also real and the function return a square (N x N)
'array.

'LIBRARY       : MATRIX
'GROUP         : NETWORK
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009

'************************************************************************************
'************************************************************************************

Function MATRIX_ADMITTANCE_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim h As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

For k = 1 To NROWS 'search for max node
    If DATA_MATRIX(k, 1) > h Then h = DATA_MATRIX(k, 1)
    If DATA_MATRIX(k, 2) > h Then h = DATA_MATRIX(k, 2)
Next k

'build the admittance matrix
ReDim TEMP_MATRIX(1 To h, 1 To 2 * h)
For k = 1 To NROWS
    i = DATA_MATRIX(k, 1)
    j = DATA_MATRIX(k, 2)
    If i <> 0 Then
        TEMP_MATRIX(i, i) = TEMP_MATRIX(i, i) + DATA_MATRIX(k, 3)
        If NCOLUMNS = 4 Then TEMP_MATRIX(i, i + h) = _
            TEMP_MATRIX(i, i + h) + DATA_MATRIX(k, 4)
    End If
    If j <> 0 Then
        TEMP_MATRIX(j, j) = TEMP_MATRIX(j, j) + DATA_MATRIX(k, 3)
        If NCOLUMNS = 4 Then TEMP_MATRIX(j, j + h) = _
            TEMP_MATRIX(j, j + h) + DATA_MATRIX(k, 4)
    End If
    If i <> 0 And j <> 0 Then
        TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) - DATA_MATRIX(k, 3)
        TEMP_MATRIX(j, i) = TEMP_MATRIX(i, j)
        If NCOLUMNS = 4 Then
            TEMP_MATRIX(i, j + h) = TEMP_MATRIX(i, j + h) - DATA_MATRIX(k, 4)
            TEMP_MATRIX(j, i + h) = TEMP_MATRIX(i, j + h)
        End If
    End If
Next k

MATRIX_ADMITTANCE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_ADMITTANCE_FUNC = Err.number
End Function
