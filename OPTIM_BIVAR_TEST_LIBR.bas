Attribute VB_Name = "OPTIM_BIVAR_TEST_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.

                            
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_BIVAR_OBJ_1_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_TEST
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_BIVAR_OBJ_1_FUNC(ByRef PARAM_RNG As Variant)

Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
CALL_BIVAR_OBJ_1_FUNC = (10 * PARAM_VECTOR(1, 1) ^ 2 + 1) ^ 2 * PARAM_VECTOR(2, 1) ^ 2 - 1

Exit Function
ERROR_LABEL:
CALL_BIVAR_OBJ_1_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_BIVAR_OBJ_2_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_TEST
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_BIVAR_OBJ_2_FUNC(ByRef PARAM_RNG As Variant)

Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
CALL_BIVAR_OBJ_2_FUNC = 2 * PARAM_VECTOR(1, 1) ^ 2 - PARAM_VECTOR(2, 1) ^ 2 - 2

Exit Function
ERROR_LABEL:
CALL_BIVAR_OBJ_2_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_BIVAR_OBJ_3_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_TEST
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************


Function CALL_BIVAR_OBJ_3_FUNC(ByRef PARAM_RNG As Variant)

Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
CALL_BIVAR_OBJ_3_FUNC = 2 * PARAM_VECTOR(1, 1) ^ 2 + 2 * PARAM_VECTOR(1, 1) * PARAM_VECTOR(2, 1) + 5 * PARAM_VECTOR(2, 1) ^ 2 - 2

Exit Function
ERROR_LABEL:
CALL_BIVAR_OBJ_3_FUNC = Err.number
End Function
