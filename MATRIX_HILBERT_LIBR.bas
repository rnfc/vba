Attribute VB_Name = "MATRIX_HILBERT_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_HILBERT_FUNC
'DESCRIPTION   : Returns the inverse of the (NSIZE x NSIZE) Hilbert's matrix.
'Note: this matrix is always integer. Hilbert 's  matrices are a strongly
'hill-conditioned and are useful for testing algorithms
'LIBRARY       : MATRIX
'GROUP         : HILBERT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_HILBERT_FUNC(ByVal NSIZE As Long)

Dim i As Long
Dim j As Long
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
For i = 1 To NSIZE
    For j = 1 To NSIZE
        TEMP_MATRIX(i, j) = 1 / (i + j - 1)
    Next j
Next i

MATRIX_HILBERT_FUNC = TEMP_MATRIX
Exit Function
ERROR_LABEL:
MATRIX_HILBERT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_HILBERT_INVERSE_FUNC
'DESCRIPTION   : Returns the Inverse of Hilbert's matrix
'LIBRARY       : MATRIX
'GROUP         : HILBERT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_HILBERT_INVERSE_FUNC(ByVal NSIZE As Single)

Dim i As Single
Dim j As Single

Dim ii As Single
Dim jj As Single

Dim TEMP_ARR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_ARR(1 To NSIZE)
ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)

'--------------------------------------------------------------------
TEMP_ARR(1) = NSIZE
'--------------------------------------------------------------------
For i = 2 To NSIZE
    ii = (NSIZE - i + 1) * (NSIZE + i - 1)
    jj = (i - 1) ^ 2
    TEMP_ARR(i) = TEMP_ARR(i - 1) * ii / jj
Next i
'--------------------------------------------------------------------
'compute the inverse of Hilbert's matrix
'--------------------------------------------------------------------
For i = 1 To NSIZE
    For j = 1 To NSIZE
        TEMP_MATRIX(i, j) = (-1) ^ (i + j) * TEMP_ARR(i) * _
                            TEMP_ARR(j) / (i + j - 1)
    Next j
Next i
'--------------------------------------------------------------------

MATRIX_HILBERT_INVERSE_FUNC = TEMP_MATRIX
Exit Function
ERROR_LABEL:
MATRIX_HILBERT_INVERSE_FUNC = Err.number
End Function

