Attribute VB_Name = "MATRIX_INVERSE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_INVERSE_FUNC
'DESCRIPTION   : INVERSE OF A MATRIX
'LIBRARY       : MATRIX
'GROUP         : INVERSE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_INVERSE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

On Error GoTo ERROR_LABEL

Select Case VERSION
    Case 0 'Lu Decomposition
        MATRIX_INVERSE_FUNC = MATRIX_LU_INVERSE_FUNC(DATA_RNG)
    Case 1 'SVD Decompsition
        MATRIX_INVERSE_FUNC = MATRIX_SVD_INVERSE_FUNC(DATA_RNG)
    Case Else 'Excel
        MATRIX_INVERSE_FUNC = MATRIX_CHOLESKY_INVERSE_FUNC(DATA_RNG, 0, True)
End Select

Exit Function
ERROR_LABEL:
MATRIX_INVERSE_FUNC = Err.number 'CVErr(xlErrValue)
End Function
