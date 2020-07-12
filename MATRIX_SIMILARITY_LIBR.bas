Attribute VB_Name = "MATRIX_SIMILARITY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SIMILARITY_TRANSFORM_FUNC

'DESCRIPTION   : This operation is also called the "similarity transform" of
'matrix A by matrix B --> B^(-1)*A*B
'Similarity transforms play a crucial role in the computation of
'eigenvalues, because they leave the eigenvalues of the matrix A
'unchanged. For real symmetric matrices, B is orthogonal. The
'similarity transformation is the also called "orthogonal transform.

'LIBRARY       : MATRIX
'GROUP         : SIMILARITY
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function MATRIX_SIMILARITY_TRANSFORM_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant)

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant

On Error GoTo ERROR_LABEL

TEMP1_MATRIX = DATA1_RNG: TEMP2_MATRIX = DATA2_RNG
TEMP1_MATRIX = MMULT_FUNC(TEMP1_MATRIX, TEMP2_MATRIX, 70)
TEMP2_MATRIX = MATRIX_LU_INVERSE_FUNC(TEMP2_MATRIX)
TEMP1_MATRIX = MMULT_FUNC(TEMP2_MATRIX, TEMP1_MATRIX, 70)

MATRIX_SIMILARITY_TRANSFORM_FUNC = TEMP1_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_SIMILARITY_TRANSFORM_FUNC = Err.number
End Function
