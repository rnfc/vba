Attribute VB_Name = "MATRIX_RANK_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : RANK_MATRIX_FUNC
'DESCRIPTION   : Returns the rank of a given matrix. It computes the
'sub-space of  Ax = 0, and counts the null column-vectors of the sub-space.

'LIBRARY       : MATRIX
'GROUP         : RANK
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function RANK_MATRIX_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal epsilon As Double = 10 ^ -15)

'VERSION = 0 --> diagonal
'VERSION = 1 --> triangle

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant
Dim TEMP3_MATRIX As Variant

Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2) 'get the dimension
'reduce DATA_RNG to DATA_MATRIX square matrix TEMP3_MATRIX with the
'lowest dimension

If NROWS <> NCOLUMNS Then
    TEMP1_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_RNG)
    If NROWS < NCOLUMNS Then
        TEMP2_MATRIX = MMULT_FUNC(DATA_MATRIX, TEMP1_MATRIX, 70)
        NSIZE = NROWS
    Else
        TEMP2_MATRIX = MMULT_FUNC(TEMP1_MATRIX, DATA_MATRIX, 70)
        NSIZE = NCOLUMNS
    End If
Else
    TEMP2_MATRIX = DATA_MATRIX  'nothing to do
    NSIZE = NROWS
End If

'compute the sub-space of Ax=0
TEMP3_MATRIX = MATRIX_GS_SINGULAR_LINEAR_SYSTEM_FUNC(TEMP2_MATRIX, , VERSION, epsilon, 0)
'count null column-vectors of sub-space
k = NSIZE
For j = 1 To NSIZE
    TEMP_SUM = 0
    For i = 1 To NSIZE
        TEMP_SUM = TEMP_SUM + Abs(TEMP3_MATRIX(i, j))
    Next i
    If TEMP_SUM > epsilon Then k = k - 1
Next j

RANK_MATRIX_FUNC = k
  
Exit Function
ERROR_LABEL:
RANK_MATRIX_FUNC = Err.number
End Function
