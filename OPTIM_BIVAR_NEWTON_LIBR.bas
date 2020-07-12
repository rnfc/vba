Attribute VB_Name = "OPTIM_BIVAR_NEWTON_LIBR"
'-------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : NEWTON_BIVAR_ZERO_FUNC
'DESCRIPTION   : This newton algorithm solves the implicit equation f(x, y) = 0
'returning a set of points (xi, yi) that satisfy the given equation

'LIBRARY       : OPTIMIZATION
'GROUP         : BIVAR_NEWTON
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function NEWTON_BIVAR_ZERO_FUNC(ByVal FUNC_NAME_STR As String, _
Optional ByVal X0_VAL As Double = 0, _
Optional ByVal Y0_VAL As Double = 0, _
Optional ByVal GRAD_STR_NAME As String = "", _
Optional ByRef CONVERGE_FLAG As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 800, _
Optional ByVal tolerance As Double = 0.0000000001, _
Optional ByVal epsilon As Double = 10 ^ -15)

Dim i As Long
Dim X_VAL As Double
Dim Y_VAL As Double

Dim D_VAL As Double
Dim N_VAL As Double
Dim M_VAL As Double
Dim V_VAL As Double

Dim FUNC_VECTOR As Variant
Dim GRADIENT_MATRIX As Variant

Dim RESULTS_VECTOR(1 To 2, 1 To 1) As Variant
Dim PARAM_VECTOR(1 To 2, 1 To 1) As Variant

On Error GoTo ERROR_LABEL

X_VAL = X0_VAL
Y_VAL = Y0_VAL
RESULTS_VECTOR(1, 1) = X_VAL
RESULTS_VECTOR(2, 1) = Y_VAL

V_VAL = 0.1

N_VAL = 10000

CONVERGE_FLAG = 1
i = 0

1982:

Do While i <= nLOOPS
    i = i + 1
    PARAM_VECTOR(1, 1) = X_VAL: PARAM_VECTOR(2, 1) = Y_VAL
    FUNC_VECTOR = Excel.Application.Run(FUNC_NAME_STR, PARAM_VECTOR)
    M_VAL = Abs(FUNC_VECTOR(1, 1)) + Abs(FUNC_VECTOR(2, 1))
    If (M_VAL <= tolerance) Then: GoTo 1983
    PARAM_VECTOR(1, 1) = X_VAL
    PARAM_VECTOR(2, 1) = Y_VAL
    If GRAD_STR_NAME <> "" Then
        GRADIENT_MATRIX = Excel.Application.Run(GRAD_STR_NAME, PARAM_VECTOR)
    Else
        GRADIENT_MATRIX = JACOBI_FORWARD_FUNC(FUNC_NAME_STR, PARAM_VECTOR, epsilon)
        'GRADIENT_MATRIX = JACOBI_CENTRAL_FUNC(FUNC_NAME_STR, PARAM_VECTOR, EPSILON)
    End If
    D_VAL = GRADIENT_MATRIX(1, 1) * GRADIENT_MATRIX(2, 2) - _
            GRADIENT_MATRIX(1, 2) * GRADIENT_MATRIX(2, 1)
    If D_VAL = 0 Then
     X_VAL = X_VAL * (2 * Rnd(1) - 1) + Rnd()
     Y_VAL = Y_VAL * (2 * Rnd(1) - 1) + Rnd()
     GoTo 1982
    End If
    X_VAL = X_VAL - (GRADIENT_MATRIX(2, 2) * FUNC_VECTOR(1, 1) - GRADIENT_MATRIX(1, 2) * FUNC_VECTOR(2, 1)) / D_VAL * V_VAL
    Y_VAL = Y_VAL - (-GRADIENT_MATRIX(2, 1) * FUNC_VECTOR(1, 1) + GRADIENT_MATRIX(1, 1) * FUNC_VECTOR(2, 1)) / D_VAL * V_VAL
    
    If Abs(X_VAL) > 10 Or Abs(Y_VAL) > 10 Then
        X_VAL = X0_VAL * (2 * Rnd(1) - 1) + Rnd() * 0.01
        Y_VAL = Y0_VAL * (2 * Rnd(1) - 1) + Rnd() * 0.01
        GoTo 1982
    End If
    
    If M_VAL < N_VAL Then
        RESULTS_VECTOR(1, 1) = X_VAL
        RESULTS_VECTOR(2, 1) = Y_VAL
        N_VAL = M_VAL
    End If
   
Loop
COUNTER = i
CONVERGE_FLAG = -1

1983:

If CONVERGE_FLAG = 1 Then
    RESULTS_VECTOR(1, 1) = X_VAL
    RESULTS_VECTOR(2, 1) = Y_VAL
Else
    GoTo ERROR_LABEL
End If
NEWTON_BIVAR_ZERO_FUNC = RESULTS_VECTOR

Exit Function
ERROR_LABEL:
NEWTON_BIVAR_ZERO_FUNC = Err.number
End Function
