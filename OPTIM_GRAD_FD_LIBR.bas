Attribute VB_Name = "OPTIM_GRAD_FD_LIBR"
'// PERFECT

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : UNIVAR_FD_GRAD_APPROX_FUNC
'DESCRIPTION   : Approximate the gradient with finite differences for
'univariate functions
'LIBRARY       : OPTIMIZATION
'GROUP         : GRAD_FD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function UNIVAR_FD_GRAD_APPROX_FUNC(ByVal FUNC_NAME_STR As String, _
ByVal X0_VAL As Double)

Dim j As Long

Dim DG_VAL As Double
Dim GN_VAL As Double
Dim DGN_VAL As Double
Dim GRAD_VAL As Double
Dim DELTA_VAL As Double

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP3_VAL As Double

Dim epsilon As Double
    
On Error GoTo ERROR_LABEL
    
epsilon = 4 * 10 ^ -4
j = 0
Do
    TEMP1_VAL = X0_VAL
    TEMP1_VAL = TEMP1_VAL + epsilon / 2
    TEMP3_VAL = Excel.Application.Run(FUNC_NAME_STR, TEMP1_VAL)
    
    TEMP1_VAL = X0_VAL
    TEMP1_VAL = TEMP1_VAL - epsilon / 2
    TEMP2_VAL = Excel.Application.Run(FUNC_NAME_STR, TEMP1_VAL)
    
    DELTA_VAL = (TEMP3_VAL - TEMP2_VAL) / epsilon
    If j > 0 Then 'difference norm criterion
        DG_VAL = DELTA_VAL - GRAD_VAL
        DGN_VAL = (DG_VAL)
        GN_VAL = (DELTA_VAL)
        If DGN_VAL < 0.2 * GN_VAL Then: Exit Do
    End If
    
    GRAD_VAL = DELTA_VAL  'save gradient
    epsilon = epsilon / 4
    j = j + 1
Loop Until epsilon < 10 ^ -9

UNIVAR_FD_GRAD_APPROX_FUNC = DELTA_VAL

Exit Function
ERROR_LABEL:
UNIVAR_FD_GRAD_APPROX_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTVAR_FD_GRAD_VALID_FUNC
'DESCRIPTION   : Finite-difference gradient validation function
'LIBRARY       : OPTIMIZATION
'GROUP         : GRAD_FD
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function MULTVAR_FD_GRAD_VALID_FUNC(ByVal FUNC_NAME_STR As String, _
ByVal GRAD_STR_NAME As String, _
ByRef PARAM_RNG As Variant, _
Optional ByRef SCALE_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True)

Dim i As Long
Dim NSIZE As Long

Dim ABS_VAL As Double
Dim NORM_VAL As Double
Dim ERROR_VAL As Double

Dim F_ARR As Variant
Dim G_ARR As Variant

Dim PARAM_VECTOR As Variant
Dim SCALE_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

If IsArray(SCALE_RNG) = True Then
    SCALE_VECTOR = SCALE_RNG
    If UBound(SCALE_VECTOR, 1) = 1 Then
        SCALE_VECTOR = MATRIX_TRANSPOSE_FUNC(SCALE_VECTOR)
    End If
Else
    ReDim SCALE_VECTOR(1 To NSIZE, 1 To 1)
    For i = 1 To NSIZE
        SCALE_VECTOR(i, 1) = 1
    Next i
End If

G_ARR = MULTVAR_CALL_GRAD_FUNC(GRAD_STR_NAME, PARAM_VECTOR, _
               SCALE_VECTOR, MIN_FLAG)
                                                    
F_ARR = MULTVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
               PARAM_VECTOR, SCALE_VECTOR, MIN_FLAG)

ERROR_VAL = 0
For i = 1 To NSIZE
    ERROR_VAL = ERROR_VAL + Abs(G_ARR(i, 1) - F_ARR(i, 1))
Next i

NORM_VAL = 0 'return the Euclidean norm of a vector
For i = 1 To UBound(F_ARR, 1)
    NORM_VAL = NORM_VAL + F_ARR(i, 1) ^ 2
Next i
ABS_VAL = (NORM_VAL) ^ 0.5
    
If ERROR_VAL > 0.001 * ABS_VAL Then
    MULTVAR_FD_GRAD_VALID_FUNC = False
Else
    MULTVAR_FD_GRAD_VALID_FUNC = True
End If

Exit Function
ERROR_LABEL:
MULTVAR_FD_GRAD_VALID_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTVAR_FD_GRAD_APPROX_FUNC
'DESCRIPTION   : Approximate the gradient with finite differences
'LIBRARY       : OPTIMIZATION
'GROUP         : GRAD_FD
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function MULTVAR_FD_GRAD_APPROX_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef PARAM_RNG As Variant, _
Optional ByRef SCALE_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim NSIZE As Long

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double

Dim MULT_VAL As Double
Dim NORM_VAL As Double
Dim DELTA_VAL As Double
Dim BOUND_VAL As Double
Dim FACTOR_VAL As Double

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim DELTA_VECTOR As Variant
Dim GRAD_VECTOR As Variant
Dim PARAM_VECTOR As Variant
Dim SCALE_VECTOR As Variant

On Error GoTo ERROR_LABEL

If MIN_FLAG = True Then
    FACTOR_VAL = 1
Else
    FACTOR_VAL = -1
End If

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
    
NSIZE = UBound(PARAM_VECTOR, 1)

If IsArray(SCALE_RNG) = True Then
    SCALE_VECTOR = SCALE_RNG
    If UBound(SCALE_VECTOR, 1) = 1 Then: SCALE_VECTOR = MATRIX_TRANSPOSE_FUNC(SCALE_VECTOR)
Else
    ReDim SCALE_VECTOR(1 To NSIZE, 1 To 1)
    For i = 1 To NSIZE
        SCALE_VECTOR(i, 1) = 1
    Next i
End If

For i = 1 To NSIZE
    PARAM_VECTOR(i, 1) = SCALE_VECTOR(i, 1) * PARAM_VECTOR(i, 1)
Next i

ReDim GRAD_VECTOR(1 To NSIZE, 1 To 1)
ReDim DELTA_VECTOR(1 To NSIZE, 1 To 1)

DELTA_VAL = 4 * 10 ^ -4
j = 0

'--------------------------------------------------------------------------------------
Do
'--------------------------------------------------------------------------------------
    
    For i = 1 To NSIZE
        TEMP2_VECTOR = PARAM_VECTOR
        TEMP2_VECTOR(i, 1) = TEMP2_VECTOR(i, 1) + DELTA_VAL / 2
        TEMP2_VAL = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, TEMP2_VECTOR, "", MIN_FLAG)
        
        TEMP2_VECTOR = PARAM_VECTOR
        TEMP2_VECTOR(i, 1) = TEMP2_VECTOR(i, 1) - DELTA_VAL / 2
        TEMP1_VAL = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, TEMP2_VECTOR, "", MIN_FLAG)
        
        GRAD_VECTOR(i, 1) = (TEMP2_VAL - TEMP1_VAL) / DELTA_VAL
    Next i
'--------------------------------------------------------------------------------------
    If j > 0 Then
'--------------------------------------------------------------------------------------
        'difference norm criterion
        For i = 1 To NSIZE
            DELTA_VECTOR(i, 1) = GRAD_VECTOR(i, 1) - TEMP1_VECTOR(i, 1)
        Next i
        
        NORM_VAL = 0 'return the Euclidean norm of a vector
        For i = 1 To NSIZE
            NORM_VAL = NORM_VAL + DELTA_VECTOR(i, 1) ^ 2
        Next i
        BOUND_VAL = (NORM_VAL) ^ 0.5
        
        NORM_VAL = 0 'return the Euclidean norm of a vector
        For i = 1 To NSIZE
            NORM_VAL = NORM_VAL + GRAD_VECTOR(i, 1) ^ 2
        Next i
        MULT_VAL = (NORM_VAL) ^ 0.5

        If BOUND_VAL < 0.2 * MULT_VAL Then Exit Do
'--------------------------------------------------------------------------------------
    End If
'--------------------------------------------------------------------------------------
    TEMP1_VECTOR = GRAD_VECTOR  'save gradient
    DELTA_VAL = DELTA_VAL / 4
    j = j + 1
'--------------------------------------------------------------------------------------
Loop Until DELTA_VAL < 10 ^ -9
'--------------------------------------------------------------------------------------

MULTVAR_FD_GRAD_APPROX_FUNC = GRAD_VECTOR

Exit Function
ERROR_LABEL:
MULTVAR_FD_GRAD_APPROX_FUNC = Err.number
End Function
