Attribute VB_Name = "MATRIX_RC_DIMENSION_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : IS_ARRAY_FUNC
'DESCRIPTION   : Check if argument is a 1 dimension (e.g., vector) array
'LIBRARY       : MATRIX
'GROUP         : RC_DIMENSION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function IS_1D_ARRAY_FUNC(ByRef DATA_RNG As Variant)
Dim NSIZE As Long
Dim DATA_VECTOR As Variant
On Error GoTo ERROR_LABEL
DATA_VECTOR = DATA_RNG
IS_1D_ARRAY_FUNC = False
NSIZE = UBound(DATA_VECTOR, 1)
IS_1D_ARRAY_FUNC = True
NSIZE = UBound(DATA_VECTOR, 2)
IS_1D_ARRAY_FUNC = False
Exit Function
ERROR_LABEL:
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : IS_2D_ARRAY_FUNC
'DESCRIPTION   : Check if argument is a 2 dimension (e.g., matrix) array
'LIBRARY       : MATRIX
'GROUP         : RC_DIMENSION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function IS_2D_ARRAY_FUNC(ByRef DATA_RNG As Variant)
Dim NSIZE As Long
Dim DATA_MATRIX As Variant
On Error GoTo ERROR_LABEL
DATA_MATRIX = DATA_RNG
IS_2D_ARRAY_FUNC = False
NSIZE = UBound(DATA_MATRIX, 1)
NSIZE = UBound(DATA_MATRIX, 2)
IS_2D_ARRAY_FUNC = True
Exit Function
ERROR_LABEL:
End Function
