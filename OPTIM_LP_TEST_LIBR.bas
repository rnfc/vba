Attribute VB_Name = "OPTIM_LP_TEST_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLP_OBJ_1_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : LP_TEST
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function CALL_NLP_OBJ_1_FUNC(ByRef PARAM_RNG As Variant)

Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)


CALL_NLP_OBJ_1_FUNC = PARAM_VECTOR(1, 1) ^ 2 + _
                             2 * PARAM_VECTOR(2, 1) ^ 2 - _
                             2 * PARAM_VECTOR(1, 1) - _
                             8 * PARAM_VECTOR(2, 1) + 9

Exit Function
ERROR_LABEL:
CALL_NLP_OBJ_1_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLP_CONST_1_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : LP_TEST
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLP_CONST_1_FUNC(ByRef PARAM_RNG As Variant)

Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

CALL_NLP_CONST_1_FUNC = True

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

If (PARAM_VECTOR(1, 1) < 0) Or (PARAM_VECTOR(2, 1) < 0) Or _
   (PARAM_VECTOR(1, 1) > 2) Or (PARAM_VECTOR(2, 1) > 3) Or _
   ((PARAM_VECTOR(1, 1) * 3 + PARAM_VECTOR(2, 1) * 2) > 6) Or _
   ((PARAM_VECTOR(1, 1) * 2 + PARAM_VECTOR(2, 1) * -1) > 1) Then
   
   CALL_NLP_CONST_1_FUNC = False
End If

Exit Function
ERROR_LABEL:
CALL_NLP_CONST_1_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLP_OBJ_2_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : LP_TEST
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function CALL_NLP_OBJ_2_FUNC(ByRef PARAM_RNG As Variant)

Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)


CALL_NLP_OBJ_2_FUNC = 1 / (PARAM_VECTOR(1, 1) ^ 2 - PARAM_VECTOR(1, 1) * _
                        PARAM_VECTOR(2, 1) - PARAM_VECTOR(1, 1) + _
                        PARAM_VECTOR(2, 1) ^ 2 - PARAM_VECTOR(2, 1) + 2)

Exit Function
ERROR_LABEL:
CALL_NLP_OBJ_2_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLP_CONST_2_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : LP_TEST
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function CALL_NLP_CONST_2_FUNC(ByRef PARAM_RNG As Variant)

Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

CALL_NLP_CONST_2_FUNC = True

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

If (PARAM_VECTOR(1, 1) < 0) Or (PARAM_VECTOR(2, 1) < 0) Or _
   (PARAM_VECTOR(1, 1) > 3) Or (PARAM_VECTOR(2, 1) > 2) Or _
   ((PARAM_VECTOR(1, 1) * 1 + PARAM_VECTOR(2, 1) * -1) > 2) Or _
   ((PARAM_VECTOR(1, 1) * 1 + PARAM_VECTOR(2, 1) * -1) < -1) Then
   
   CALL_NLP_CONST_2_FUNC = False
End If

Exit Function
ERROR_LABEL:
CALL_NLP_CONST_2_FUNC = False
End Function
