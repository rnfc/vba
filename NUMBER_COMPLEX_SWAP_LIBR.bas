Attribute VB_Name = "NUMBER_COMPLEX_SWAP_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : CPLX_MATRIX_SWAP_ROW_FUNC
'DESCRIPTION   : Swaps complex rows k and i
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : SWAP
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function CPLX_MATRIX_SWAP_ROW_FUNC(ByRef COMPLEX_OBJ() As Cplx, _
ByVal k As Long, _
ByVal i As Long)
    
Dim j As Long
Dim TEMP_OBJ As Cplx

On Error GoTo ERROR_LABEL

CPLX_MATRIX_SWAP_ROW_FUNC = False
For j = LBound(COMPLEX_OBJ, 2) To UBound(COMPLEX_OBJ, 2)
    TEMP_OBJ = COMPLEX_OBJ(i, j)
    COMPLEX_OBJ(i, j) = COMPLEX_OBJ(k, j)
    COMPLEX_OBJ(k, j) = TEMP_OBJ
Next j
CPLX_MATRIX_SWAP_ROW_FUNC = True
    
Exit Function
ERROR_LABEL:
CPLX_MATRIX_SWAP_ROW_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CPLX_MATRIX_SWAP_COLUMN_FUNC
'DESCRIPTION   : Swaps complex columns k and i
'LIBRARY       : NUMBER_COMPLEX
'GROUP         : SWAP
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function CPLX_MATRIX_SWAP_COLUMN_FUNC(ByRef COMPLEX_OBJ() As Cplx, _
ByVal k As Long, _
ByVal j As Long)
    
Dim i As Long
Dim TEMP_OBJ As Cplx

On Error GoTo ERROR_LABEL

CPLX_MATRIX_SWAP_COLUMN_FUNC = False
For i = LBound(COMPLEX_OBJ, 1) To UBound(COMPLEX_OBJ, 1)
    TEMP_OBJ = COMPLEX_OBJ(i, j)
    COMPLEX_OBJ(i, j) = COMPLEX_OBJ(i, k)
    COMPLEX_OBJ(i, k) = TEMP_OBJ
Next i
CPLX_MATRIX_SWAP_COLUMN_FUNC = True

Exit Function
ERROR_LABEL:
CPLX_MATRIX_SWAP_COLUMN_FUNC = False
End Function
