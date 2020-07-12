Attribute VB_Name = "OPTIM_MULTVAR_RESIZE_LIBR"

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
'FUNCTION      : MULTVAR_RESIZE_OPTIM_FUNC
'DESCRIPTION   : Optimization routine with Montecarlo method with resizing
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_RESIZE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function MULTVAR_RESIZE_OPTIM_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef CONST_RNG As Variant, _
Optional ByVal GRAD_STR_NAME As String, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal nLOOPS As Long = 10000, _
Optional ByVal epsilon As Double = 0.000000000000001)

Dim i As Long

Dim NROWS As Long

Dim CONST_BOX As Variant

Dim SCALE_VECTOR As Variant
Dim PARAM_VECTOR As Variant

Dim XTEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

'-----------------------------------------------------------------------------------
On Error GoTo ERROR_LABEL
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
    CONST_BOX = CONST_RNG
    CONST_BOX = MULTVAR_LOAD_CONST_FUNC(CONST_BOX, 1)
'-----------------------------------------------------------------------------------
    NROWS = UBound(CONST_BOX)
    TEMP_MATRIX = MULTVAR_SCALE_CONST_FUNC(CONST_BOX)
    CONST_BOX = TEMP_MATRIX(LBound(TEMP_MATRIX))
    SCALE_VECTOR = TEMP_MATRIX(UBound(TEMP_MATRIX))
'-----------------------------------------------------------------------------------
    If VERSION = 0 Then: GoTo 1983
'-----------------------------------------------------------------------------------
    If IsArray(PARAM_RNG) = True Then
        PARAM_VECTOR = PARAM_RNG
        If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
    Else
        PARAM_VECTOR = MULTVAR_MC_OPTIM_FUNC(CONST_BOX, FUNC_NAME_STR, _
                        SCALE_VECTOR, MIN_FLAG, nLOOPS, 0, epsilon)
    End If
'-----------------------------------------------------------------------------------
    ReDim XTEMP_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        XTEMP_VECTOR(i, 1) = (1 / SCALE_VECTOR(i, 1)) * PARAM_VECTOR(i, 1)
    Next i
'-----------------------------------------------------------------------------------
    Select Case VERSION
'-----------------------------------------------------------------------------------
        Case 0
'-----------------------------------------------------------------------------------
1983:
            PARAM_VECTOR = MULTVAR_MC_OPTIM_FUNC(CONST_BOX, FUNC_NAME_STR, _
                           SCALE_VECTOR, MIN_FLAG, nLOOPS, 0, epsilon)
'-----------------------------------------------------------------------------------
        Case 1
'-----------------------------------------------------------------------------------
            PARAM_VECTOR = MULTVAR_GRAD_OPTIM_FUNC(XTEMP_VECTOR, CONST_BOX, _
                    FUNC_NAME_STR, GRAD_STR_NAME, _
                    SCALE_VECTOR, MIN_FLAG, nLOOPS, 0, 0, epsilon)
'-----------------------------------------------------------------------------------
        Case 2
'-----------------------------------------------------------------------------------
            PARAM_VECTOR = MULTVAR_GRAD_OPTIM_FUNC(XTEMP_VECTOR, CONST_BOX, _
                    FUNC_NAME_STR, GRAD_STR_NAME, _
                    SCALE_VECTOR, MIN_FLAG, nLOOPS, 1, 0, epsilon)
'-----------------------------------------------------------------------------------
        Case Else
'-----------------------------------------------------------------------------------
            PARAM_VECTOR = MULTVAR_DFP_OPTIM_FUNC(XTEMP_VECTOR, CONST_BOX, _
                    FUNC_NAME_STR, GRAD_STR_NAME, _
                    SCALE_VECTOR, MIN_FLAG, nLOOPS, 0, epsilon)
'-----------------------------------------------------------------------------------
    End Select
'-----------------------------------------------------------------------------------
    
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
If (GRAD_STR_NAME <> "") Then
    If Not MULTVAR_FD_GRAD_VALID_FUNC(FUNC_NAME_STR, _
        GRAD_STR_NAME, PARAM_VECTOR, "", MIN_FLAG) Then 'Gradient: dubious accuracy
        PARAM_VECTOR = MULTVAR_NR_OPTIM_FUNC(PARAM_VECTOR, _
                        FUNC_NAME_STR, GRAD_STR_NAME, _
                        MIN_FLAG, 3, 0, epsilon)
                        'COUNTER = COUNTER + (NROWS ^ 2 + NROWS)
    End If
End If
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
    
MULTVAR_RESIZE_OPTIM_FUNC = PARAM_VECTOR

Exit Function
ERROR_LABEL:
    MULTVAR_RESIZE_OPTIM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTVAR_GRAD_OPTIM_FUNC

'DESCRIPTION   : This is a very popular algorithm that cannot miss. It requires
'a gradient evaluation at each step which can be approximated internally by
'the finite difference method or supplied directly by the user as well. The exact
'gradient information improves the accuracy of the final result, but in many
'case these differences are not relevant to the extra effort. The starting point
'should be chosen sufficiently close to the optimized one.

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_RESIZE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************



Function MULTVAR_GRAD_OPTIM_FUNC(ByRef PARAM_RNG As Variant, _
ByRef CONST_RNG As Variant, _
ByVal FUNC_NAME_STR As String, _
Optional ByVal GRAD_STR_NAME As String, _
Optional ByRef SCALE_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal VERSION As Integer = 1, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal epsilon As Double = 10 ^ -13)

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim COUNTER As Long
Dim NO_POINTS As Long

Dim CONST_BOX As Variant

Dim XTEMP_ABS As Double
Dim XTEMP_ERR As Double
Dim INIT_ERR As Double

Dim TEMP_SCAL As Double
Dim TEMP_FACT As Double
Dim TEMP_NORM As Double

Dim FIRST_MULT As Double
Dim SECOND_MULT As Double

Dim MIN_FUNC_VAL As Double

Dim SCALE_VECTOR As Variant
Dim START_VECTOR As Variant
Dim VALUES_VECTOR As Variant

Dim FTEMP_VECTOR As Variant
Dim GTEMP_VECTOR As Variant
Dim PTEMP_VECTOR As Variant
Dim VTEMP_VECTOR As Variant
Dim XTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

XTEMP_VECTOR = PARAM_RNG
If UBound(XTEMP_VECTOR, 1) = 1 Then: XTEMP_VECTOR = MATRIX_TRANSPOSE_FUNC(XTEMP_VECTOR)

CONST_BOX = CONST_RNG

NROWS = UBound(XTEMP_VECTOR, 1)

If IsArray(SCALE_RNG) = True Then
    SCALE_VECTOR = SCALE_RNG
    If UBound(SCALE_VECTOR, 1) = 1 Then: SCALE_VECTOR = MATRIX_TRANSPOSE_FUNC(SCALE_VECTOR)
Else
    ReDim SCALE_VECTOR(1 To NROWS, 1 To 1)
    For ii = 1 To NROWS
        SCALE_VECTOR(ii, 1) = 1
    Next ii
End If


NO_POINTS = 10

ReDim START_VECTOR(1 To NROWS, 1 To 1)
ReDim VALUES_VECTOR(1 To NROWS, 1 To 1)

ReDim PTEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim VTEMP_VECTOR(1 To NROWS, 1 To 1)

For ii = 1 To NROWS
    START_VECTOR(ii, 1) = XTEMP_VECTOR(ii, 1)
Next ii

XTEMP_ERR = 1
COUNTER = 0

If GRAD_STR_NAME = "" Then
    FTEMP_VECTOR = MULTVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                    XTEMP_VECTOR, SCALE_VECTOR, MIN_FLAG)
    COUNTER = 2 * NROWS
Else
    FTEMP_VECTOR = MULTVAR_CALL_GRAD_FUNC(GRAD_STR_NAME, _
                XTEMP_VECTOR, SCALE_VECTOR, MIN_FLAG)
End If

For ii = 1 To NROWS
    PTEMP_VECTOR(ii, 1) = -FTEMP_VECTOR(ii, 1)
Next ii
kk = 0
Do
    VALUES_VECTOR = MULTVAR_CONST_BOUND_FUNC(CONST_BOX, _
                    PTEMP_VECTOR, START_VECTOR)
    
    XTEMP_VECTOR = SEGMENT_OPTIMIZATION_FUNC(START_VECTOR, VALUES_VECTOR, _
                    MIN_FUNC_VAL, jj, FUNC_NAME_STR, SCALE_VECTOR, MIN_FLAG, _
                    NO_POINTS, nLOOPS)
    
    COUNTER = COUNTER + jj
    If GRAD_STR_NAME = "" Then
        GTEMP_VECTOR = MULTVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        XTEMP_VECTOR, SCALE_VECTOR, MIN_FLAG)
        COUNTER = COUNTER + 2 * NROWS
    Else
        GTEMP_VECTOR = MULTVAR_CALL_GRAD_FUNC(GRAD_STR_NAME, _
                        XTEMP_VECTOR, SCALE_VECTOR, MIN_FLAG)
    End If
    
    Select Case VERSION
        Case 0              'Gradient
            For ii = 1 To NROWS
                PTEMP_VECTOR(ii, 1) = -GTEMP_VECTOR(ii, 1)
            Next ii
        Case 1              'Conjugate Gradient
            TEMP_NORM = 0 'return the Euclidean norm of a vector
            For ii = 1 To UBound(GTEMP_VECTOR)
                TEMP_NORM = TEMP_NORM + GTEMP_VECTOR(ii, 1) ^ 2
            Next ii
            SECOND_MULT = Sqr(TEMP_NORM)
            
            TEMP_NORM = 0 'return the Euclidean norm of a vector
            For ii = 1 To UBound(FTEMP_VECTOR)
                TEMP_NORM = TEMP_NORM + FTEMP_VECTOR(ii, 1) ^ 2
            Next ii
            FIRST_MULT = Sqr(TEMP_NORM)
            
            
            TEMP_FACT = 0
            For ii = 1 To NROWS
                TEMP_FACT = GTEMP_VECTOR(ii, 1) * FTEMP_VECTOR(ii, 1)
            Next ii
            
            If TEMP_FACT >= 0.2 * SECOND_MULT ^ 2 Then
                For ii = 1 To NROWS
                    PTEMP_VECTOR(ii, 1) = -GTEMP_VECTOR(ii, 1)
                Next ii
            Else
                TEMP_SCAL = (SECOND_MULT / FIRST_MULT) ^ 2
                For ii = 1 To NROWS
                    PTEMP_VECTOR(ii, 1) = -GTEMP_VECTOR(ii, 1) + _
                                        TEMP_SCAL * PTEMP_VECTOR(ii, 1)
                Next ii
            End If
    End Select
    
    'check surface constrains --------------
    For ii = 1 To NROWS
        If Abs(XTEMP_VECTOR(ii, 1) - CONST_BOX(ii, 1)) < epsilon Then
            VTEMP_VECTOR(ii, 1) = -1 'variable equal to lower bound
        ElseIf Abs(XTEMP_VECTOR(ii, 1) - CONST_BOX(ii, 2)) < epsilon Then
            VTEMP_VECTOR(ii, 1) = 1 'variable equal to upper bound
        ElseIf XTEMP_VECTOR(ii, 1) < CONST_BOX(ii, 1) - epsilon Then
            VTEMP_VECTOR(ii, 1) = -2 'variable less then lower bound
        ElseIf XTEMP_VECTOR(ii, 1) > CONST_BOX(ii, 2) + epsilon Then
            VTEMP_VECTOR(ii, 1) = 2 'variable higher then upper bound
        Else
            VTEMP_VECTOR(ii, 1) = 0 'variable internal to both bounds
        End If
    Next ii
    
    For ii = 1 To NROWS
        If VTEMP_VECTOR(ii, 1) * PTEMP_VECTOR(ii, 1) > 0 Then PTEMP_VECTOR(ii, 1) = 0
    Next ii
    
    'check stop
    XTEMP_ERR = 0
    For ii = 1 To NROWS
        XTEMP_ERR = XTEMP_ERR + (XTEMP_VECTOR(ii, 1) - START_VECTOR(ii, 1)) ^ 2
    Next ii
    XTEMP_ERR = Sqr(XTEMP_ERR)
            
    TEMP_NORM = 0 'return the Euclidean norm of a vector
    For ii = 1 To UBound(XTEMP_VECTOR)
        TEMP_NORM = TEMP_NORM + XTEMP_VECTOR(ii, 1) ^ 2
    Next ii
    XTEMP_ABS = Sqr(TEMP_NORM)
    
    If XTEMP_ABS > 1 Then XTEMP_ERR = XTEMP_ERR / XTEMP_ABS
    If kk > 4 Then
        If XTEMP_ERR > 0.3 * INIT_ERR And XTEMP_ERR < 10 ^ -5 Then
            'convergence slow, reduce the accuracy
            epsilon = epsilon * 10
        End If
    End If
    If XTEMP_ERR < epsilon Then Exit Do
    INIT_ERR = XTEMP_ERR
    '
    For ii = 1 To NROWS
        START_VECTOR(ii, 1) = XTEMP_VECTOR(ii, 1)
        FTEMP_VECTOR(ii, 1) = GTEMP_VECTOR(ii, 1)
    Next ii
    
    kk = kk + 1
Loop Until COUNTER > nLOOPS

Select Case OUTPUT
    Case 0
        For ii = 1 To UBound(XTEMP_VECTOR, 1)
            XTEMP_VECTOR(ii, 1) = XTEMP_VECTOR(ii, 1) * SCALE_VECTOR(ii, 1)
        Next ii
        MULTVAR_GRAD_OPTIM_FUNC = XTEMP_VECTOR
    Case 1
        MULTVAR_GRAD_OPTIM_FUNC = MIN_FUNC_VAL
    Case 2
        MULTVAR_GRAD_OPTIM_FUNC = COUNTER
    Case Else
        MULTVAR_GRAD_OPTIM_FUNC = XTEMP_ERR
End Select

Exit Function
ERROR_LABEL:
    MULTVAR_GRAD_OPTIM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTVAR_NR_OPTIM_FUNC

'DESCRIPTION   :
'The most popular algorithm for solving nonlinear equations.
'It needs the exact gradient for approximating the Hessian matrix. It is
'extremely fast and accurate but, because of its poor global convergence
'performance, it is used only for refining the final result from another
'algorithm.

'The function will attempt to refine the final result with 2-3 extra
'iterations of the Newton-Raphson algorithm.

'This option always requires the gradient function, for evaluating the
'Hessian matrix with sufficient accuracy to obtain a good optimum value.
'It is a numerical problem, inherent in the loss of accuracies of the
'differences obtained by numerical subtractions.

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_RESIZE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function MULTVAR_NR_OPTIM_FUNC(ByRef PARAM_RNG As Variant, _
ByVal FUNC_NAME_STR As String, _
ByVal GRAD_STR_NAME As String, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal epsilon As Double = 10 ^ -13)

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim NSIZE As Long
Dim COUNTER As Long

Dim TEMP_DET As Double
Dim XTEMP_ERR As Double

Dim MIN_FUNC_VAL As Double

Dim TEMP_ARR As Variant

Dim FIRST_VECTOR As Variant
Dim SECOND_VECTOR As Variant

Dim FUNC_VECTOR As Variant
Dim XTEMP_VECTOR As Variant

Dim GTEMP_VECTOR As Variant
Dim PTEMP_VECTOR As Variant
Dim VTEMP_VECTOR As Variant

Dim GRAD_VECTOR As Variant
Dim HESSIAN_MATRIX As Variant

Dim LAMBDA As Double

On Error GoTo ERROR_LABEL

LAMBDA = 10 ^ -5

XTEMP_VECTOR = PARAM_RNG
If UBound(XTEMP_VECTOR, 1) = 1 Then: XTEMP_VECTOR = MATRIX_TRANSPOSE_FUNC(XTEMP_VECTOR)
NROWS = UBound(XTEMP_VECTOR)

ReDim PTEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim FIRST_VECTOR(1 To NROWS, 1 To 1)
ReDim SECOND_VECTOR(1 To NROWS, 1 To 1)
ReDim GRAD_VECTOR(1 To NROWS, 1 To 1)
ReDim HESSIAN_MATRIX(1 To NROWS, 1 To NROWS + 1)

COUNTER = 0

Do
    COUNTER = COUNTER + 1
    GTEMP_VECTOR = MULTVAR_CALL_GRAD_FUNC(GRAD_STR_NAME, _
                    XTEMP_VECTOR, "", MIN_FLAG)

    For ii = 1 To NROWS
        HESSIAN_MATRIX(ii, NROWS + 1) = GTEMP_VECTOR(ii, 1)
    Next ii
    
    If COUNTER = 1 Then 'Approximate the Hessian matrix with the gradient
        For kk = 1 To NROWS
            NSIZE = NROWS - kk + 1
            ReDim VTEMP_VECTOR(1 To 2 * NSIZE, 1 To NROWS)
            For ii = 1 To NSIZE
                For jj = 1 To NROWS
                    VTEMP_VECTOR(ii, jj) = XTEMP_VECTOR(jj, 1)
                    VTEMP_VECTOR(ii + NSIZE, jj) = VTEMP_VECTOR(ii, jj)
                    If kk + ii - 1 = jj Then
                        VTEMP_VECTOR(ii, jj) = VTEMP_VECTOR(ii, jj) + LAMBDA / 2
                        VTEMP_VECTOR(ii + NSIZE, jj) = _
                            VTEMP_VECTOR(ii + NSIZE, jj) - LAMBDA / 2
                    End If
                Next jj
            Next ii
            FUNC_VECTOR = MULTVAR_CALL_POINT_GRAD_FUNC(FUNC_NAME_STR, _
                          GRAD_STR_NAME, VTEMP_VECTOR, kk, _
                          "", MIN_FLAG)
            
            For ii = 1 To NSIZE
                jj = kk + ii - 1
                    HESSIAN_MATRIX(kk, jj) = (FUNC_VECTOR(ii, 1) - _
                            FUNC_VECTOR(NSIZE + ii, 1)) / LAMBDA
                    If kk <> jj Then HESSIAN_MATRIX(jj, kk) = HESSIAN_MATRIX(kk, jj)
            Next ii
        Next kk
    End If
        
    TEMP_ARR = MATRIX_TRIANGULAR_LINEAR_SYSTEM_FUNC(HESSIAN_MATRIX)
    TEMP_DET = TEMP_ARR(UBound(TEMP_ARR))
    
    If TEMP_DET <= epsilon Then Exit Do 'nothing to do
    GRAD_VECTOR = TEMP_ARR(LBound(TEMP_ARR))
    For ii = 1 To NROWS
        XTEMP_VECTOR(ii, 1) = XTEMP_VECTOR(ii, 1) - GRAD_VECTOR(ii, 1)
    Next ii
    
    MIN_FUNC_VAL = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, _
                    XTEMP_VECTOR, "", MIN_FLAG)

Loop Until COUNTER > nLOOPS

Select Case OUTPUT
    Case 0
        MULTVAR_NR_OPTIM_FUNC = XTEMP_VECTOR
    Case 1
        MULTVAR_NR_OPTIM_FUNC = MIN_FUNC_VAL
    Case 2
        MULTVAR_NR_OPTIM_FUNC = COUNTER
    Case Else
        MULTVAR_NR_OPTIM_FUNC = XTEMP_ERR
End Select

Exit Function
ERROR_LABEL:
MULTVAR_NR_OPTIM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTVAR_DFP_OPTIM_FUNC

'DESCRIPTION   : Davidon-Fletcher-Powell minimization algorithm
'This is a sophisticated and efficient method for finding extremes of
'smooth-regular functions. It requires a gradient evaluation
'at each step which can be approximated internally by the finite difference
'method or supplied directly by the user as well. The exact gradient
'information improves the accuracy of the final result, but in many case these
'differences are not relevant to the extra effort. The starting point should be
'chosen sufficiently close to the optimized one, even if the region is larger
'than the allowable region for a CG solution.

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_RESIZE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function MULTVAR_DFP_OPTIM_FUNC(ByRef PARAM_RNG As Variant, _
ByRef CONST_RNG As Variant, _
ByVal FUNC_NAME_STR As String, _
Optional ByVal GRAD_STR_NAME As String, _
Optional ByRef SCALE_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal epsilon As Double = 10 ^ -13)

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim COUNTER As Long
Dim NO_POINTS As Long

Dim TEMP_PROD As Double
Dim TEMP_NORM As Double
Dim TEMP_SUM As Double

Dim INIT_ERR As Double
Dim XTEMP_ERR As Double
Dim XTEMP_ABS As Double
Dim MIN_FUNC_VAL As Double

Dim HTEMP_MATRIX As Variant
Dim WTEMP_MATRIX As Variant

Dim CONST_BOX As Variant

Dim DTEMP_VECTOR As Variant
Dim FTEMP_VECTOR As Variant
Dim GTEMP_VECTOR As Variant
Dim STEMP_VECTOR As Variant
Dim UTEMP_VECTOR As Variant
Dim VTEMP_VECTOR As Variant
Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

Dim SCALE_VECTOR As Variant
Dim VALUES_VECTOR As Variant
Dim PARAM_VECTOR As Variant

Dim tolerance As Double

Dim loop_count As Long

On Error GoTo ERROR_LABEL

tolerance = epsilon

NO_POINTS = 10

XTEMP_VECTOR = PARAM_RNG
If UBound(XTEMP_VECTOR, 1) = 1 Then: XTEMP_VECTOR = MATRIX_TRANSPOSE_FUNC(XTEMP_VECTOR)
NROWS = UBound(XTEMP_VECTOR, 1)

CONST_BOX = CONST_RNG

If IsArray(SCALE_RNG) = True Then
    SCALE_VECTOR = SCALE_RNG
    If UBound(SCALE_VECTOR, 1) = 1 Then: SCALE_VECTOR = MATRIX_TRANSPOSE_FUNC(SCALE_VECTOR)
Else
    ReDim SCALE_VECTOR(1 To NROWS, 1 To 1)
    For ii = 1 To NROWS
        SCALE_VECTOR(ii, 1) = 1
    Next ii
End If


ReDim HTEMP_MATRIX(1 To NROWS, 1 To NROWS)
ReDim WTEMP_MATRIX(1 To NROWS, 1 To NROWS)

ReDim PARAM_VECTOR(1 To NROWS, 1 To 1)
ReDim VALUES_VECTOR(1 To NROWS, 1 To 1)

ReDim DTEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim UTEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim VTEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim STEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim YTEMP_VECTOR(1 To NROWS, 1 To 1)

For ii = 1 To NROWS
    PARAM_VECTOR(ii, 1) = XTEMP_VECTOR(ii, 1)
Next ii

ReDim HTEMP_MATRIX(1 To NROWS, 1 To NROWS) 'identity matrix
For ii = 1 To NROWS
    For jj = 1 To NROWS
        If ii = jj Then
            HTEMP_MATRIX(ii, jj) = 1
        Else
            HTEMP_MATRIX(ii, jj) = 0
        End If
    Next jj
Next ii

XTEMP_ERR = 1
INIT_ERR = 1
If GRAD_STR_NAME = "" Then
    FTEMP_VECTOR = MULTVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                    PARAM_VECTOR, SCALE_VECTOR, MIN_FLAG)
    COUNTER = 2 * NROWS
Else
    FTEMP_VECTOR = MULTVAR_CALL_GRAD_FUNC(GRAD_STR_NAME, _
                    PARAM_VECTOR, SCALE_VECTOR, MIN_FLAG)
End If

'Fletcher-Powell algorithm begins
COUNTER = 0
Do
    'set starting direction
    For ii = 1 To NROWS
        DTEMP_VECTOR(ii, 1) = 0
        For jj = 1 To NROWS
            DTEMP_VECTOR(ii, 1) = DTEMP_VECTOR(ii, 1) - HTEMP_MATRIX(ii, jj) * _
                                 FTEMP_VECTOR(jj, 1)
        Next jj
    Next ii
    
    'check surface constrains --------------
    For ii = 1 To NROWS
        If Abs(XTEMP_VECTOR(ii, 1) - CONST_BOX(ii, 1)) < tolerance Then
            VTEMP_VECTOR(ii, 1) = -1 'variable equal to lower bound
        ElseIf Abs(XTEMP_VECTOR(ii, 1) - CONST_BOX(ii, 2)) < tolerance Then
            VTEMP_VECTOR(ii, 1) = 1 'variable equal to upper bound
        ElseIf XTEMP_VECTOR(ii, 1) < CONST_BOX(ii, 1) - tolerance Then
            VTEMP_VECTOR(ii, 1) = -2 'variable less then lower bound
        ElseIf XTEMP_VECTOR(ii, 1) > CONST_BOX(ii, 2) + tolerance Then
            VTEMP_VECTOR(ii, 1) = 2 'variable higher then upper bound
        Else
            VTEMP_VECTOR(ii, 1) = 0 'variable internal to both bounds
        End If
    Next ii
    
    For ii = 1 To NROWS
        If VTEMP_VECTOR(ii, 1) * _
           DTEMP_VECTOR(ii, 1) > 0 Then DTEMP_VECTOR(ii, 1) = 0
    Next ii
    '-------------------------------------
    
    VALUES_VECTOR = MULTVAR_CONST_BOUND_FUNC(CONST_BOX, _
                    DTEMP_VECTOR, PARAM_VECTOR)
                    
    
    XTEMP_VECTOR = SEGMENT_OPTIMIZATION_FUNC(PARAM_VECTOR, VALUES_VECTOR, _
                MIN_FUNC_VAL, kk, FUNC_NAME_STR, SCALE_VECTOR, _
                MIN_FLAG, NO_POINTS, nLOOPS)
    
    COUNTER = COUNTER + kk
    
    'new gradient
    If GRAD_STR_NAME = "" Then
        GTEMP_VECTOR = MULTVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        XTEMP_VECTOR, SCALE_VECTOR, MIN_FLAG)
        COUNTER = COUNTER + 2 * NROWS
    Else
        GTEMP_VECTOR = MULTVAR_CALL_GRAD_FUNC(GRAD_STR_NAME, _
                        XTEMP_VECTOR, SCALE_VECTOR, MIN_FLAG)
    End If
    
    For ii = 1 To NROWS
        STEMP_VECTOR(ii, 1) = XTEMP_VECTOR(ii, 1) - PARAM_VECTOR(ii, 1)
        YTEMP_VECTOR(ii, 1) = GTEMP_VECTOR(ii, 1) - FTEMP_VECTOR(ii, 1)
    Next ii
    
    'check stop
    
    TEMP_NORM = 0 'return the Euclidean norm of a vector
    For ii = 1 To UBound(XTEMP_VECTOR)
        TEMP_NORM = TEMP_NORM + XTEMP_VECTOR(ii, 1) ^ 2
    Next ii
    XTEMP_ABS = Sqr(TEMP_NORM)
    
    TEMP_NORM = 0 'return the Euclidean norm of a vector
    For ii = 1 To UBound(STEMP_VECTOR)
        TEMP_NORM = TEMP_NORM + STEMP_VECTOR(ii, 1) ^ 2
    Next ii
    XTEMP_ERR = Sqr(TEMP_NORM)


    If XTEMP_ABS > 1 Then XTEMP_ERR = XTEMP_ERR / XTEMP_ABS
    If loop_count > 4 Then
        If XTEMP_ERR > 0.3 * INIT_ERR And XTEMP_ERR < 10 ^ -5 Then
            'convergence slow, reduce the accuracy
            tolerance = tolerance * 10
        End If
    End If
    If loop_count > 0 And XTEMP_ERR <= tolerance Then
        Exit Do
    End If
    INIT_ERR = XTEMP_ERR
    '
    ReDim UTEMP_VECTOR(1 To NROWS, 1 To 1)
    For ii = 1 To NROWS
        UTEMP_VECTOR(ii, 1) = 0
        For jj = 1 To NROWS
            UTEMP_VECTOR(ii, 1) = UTEMP_VECTOR(ii, 1) + _
                HTEMP_MATRIX(ii, jj) * YTEMP_VECTOR(jj, 1)
        Next jj
    Next ii
     
    TEMP_SUM = 0
    For ii = 1 To NROWS
        TEMP_SUM = TEMP_SUM + STEMP_VECTOR(ii, 1) * YTEMP_VECTOR(ii, 1)
    Next ii
    TEMP_PROD = TEMP_SUM
    If TEMP_PROD = 0 Then: TEMP_PROD = 0.000000000000001
    
    For ii = 1 To NROWS
        For jj = 1 To NROWS
            WTEMP_MATRIX(ii, jj) = STEMP_VECTOR(ii, 1) * _
                                   STEMP_VECTOR(jj, 1) / TEMP_PROD
            If ii <> jj Then WTEMP_MATRIX(jj, ii) = WTEMP_MATRIX(ii, jj)
        Next jj
    Next ii
    
    For ii = 1 To NROWS
        For jj = 1 To NROWS
            HTEMP_MATRIX(ii, jj) = HTEMP_MATRIX(ii, jj) + WTEMP_MATRIX(ii, jj)
        Next jj
    Next ii
    
    TEMP_SUM = 0
    For ii = 1 To NROWS
        TEMP_SUM = TEMP_SUM + YTEMP_VECTOR(ii, 1) * UTEMP_VECTOR(ii, 1)
    Next ii
    TEMP_PROD = TEMP_SUM
    If TEMP_PROD = 0 Then: TEMP_PROD = 0.000000000000001
    
    For ii = 1 To NROWS
        For jj = 1 To NROWS
            WTEMP_MATRIX(ii, jj) = UTEMP_VECTOR(ii, 1) * _
                                   UTEMP_VECTOR(jj, 1) / TEMP_PROD
            If ii <> jj Then WTEMP_MATRIX(jj, ii) = WTEMP_MATRIX(ii, jj)
        Next jj
    Next ii
    
    For ii = 1 To NROWS
        For jj = 1 To NROWS
            HTEMP_MATRIX(ii, jj) = HTEMP_MATRIX(ii, jj) - WTEMP_MATRIX(ii, jj)
        Next jj
    Next ii
    
    For ii = 1 To NROWS
        FTEMP_VECTOR(ii, 1) = GTEMP_VECTOR(ii, 1)
        PARAM_VECTOR(ii, 1) = XTEMP_VECTOR(ii, 1)
    Next ii
    
    COUNTER = COUNTER + 1
    loop_count = loop_count + 1
Loop Until COUNTER > nLOOPS

Select Case OUTPUT
    Case 0
        For ii = 1 To UBound(XTEMP_VECTOR, 1)
            XTEMP_VECTOR(ii, 1) = XTEMP_VECTOR(ii, 1) * SCALE_VECTOR(ii, 1)
        Next ii
        
        MULTVAR_DFP_OPTIM_FUNC = XTEMP_VECTOR
    Case 1
        MULTVAR_DFP_OPTIM_FUNC = MIN_FUNC_VAL
    Case 2
        MULTVAR_DFP_OPTIM_FUNC = COUNTER
    Case Else
        MULTVAR_DFP_OPTIM_FUNC = XTEMP_ERR
End Select

Exit Function
ERROR_LABEL:
MULTVAR_DFP_OPTIM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MULTVAR_MC_OPTIM_FUNC

'DESCRIPTION   : This is another derivative-free algorithm. It simply "shoots" a
'set of random points and takes the best extreme value (max or min). Usually the
'accuracy is not comparable with the other algorithms (only about 5%), and it also
'requires a considerable extra effort and time. On the other hand, it's
'absolutely insensitive to the presence of unwanted local extremes, and works
'with smooth and discontinues functions as well. In this implementation, the
'random algorithm can increase the accuracy (0.01%) by a "resizing" strategy
'(under particular conditions of the objective function). On the contrary, this
'algorithm is not adaptable for functions that have a large "flat" region near the
'extreme, like what happens in the least squared optimization. Convergence
'problems do not exist because a starting point is not necessary.

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_RESIZE
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function MULTVAR_MC_OPTIM_FUNC(ByRef CONST_RNG As Variant, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef SCALE_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal epsilon As Double = 10 ^ -13)

Dim hh As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim COUNTER As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NTRIALS As Long

Dim DELTA_ERR As Double
Dim XTEMP_ERR As Double
Dim YTEMP_ERR As Double

Dim TEMP_FACT As Double
Dim TEMP_VALUE As Double

Dim MIN_FUNC_VAL As Double

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

Dim CONST_BOX As Variant

Dim XMEAN_VECTOR As Variant
Dim XDEAV_VECTOR As Variant

Dim PARAM_VECTOR As Variant
Dim XTEMP_VECTOR As Variant

Dim SCALE_VECTOR As Variant

On Error GoTo ERROR_LABEL

CONST_BOX = CONST_RNG

NROWS = UBound(CONST_BOX, 1)
ReDim XTEMP_VECTOR(1 To NROWS, 1 To 2)
If IsArray(SCALE_RNG) = True Then
    SCALE_VECTOR = SCALE_RNG
    If UBound(SCALE_VECTOR, 1) = 1 Then: SCALE_VECTOR = MATRIX_TRANSPOSE_FUNC(SCALE_VECTOR)
Else
    ReDim SCALE_VECTOR(1 To NROWS, 1 To 1)
    For ii = 1 To NROWS
        SCALE_VECTOR(ii, 1) = 1
    Next ii
End If


NSIZE = Int(17 * NROWS ^ 0.5)
NTRIALS = Int(0.02 * nLOOPS * NROWS ^ 0.5)
COUNTER = 0

For ii = 1 To NROWS
    XTEMP_VECTOR(ii, 1) = CONST_BOX(ii, 1)
    XTEMP_VECTOR(ii, 2) = CONST_BOX(ii, 2)
Next ii

DELTA_ERR = 10 ^ 300
hh = 0 'stack empty
' main loop begins

ReDim BTEMP_MATRIX(1 To NSIZE, 1 To NROWS + 1)
Do
    
    ReDim XMEAN_VECTOR(1 To NROWS, 1 To 1)
    ReDim XDEAV_VECTOR(1 To NROWS, 1 To 1) 'contains statistic for each variable
    ReDim ATEMP_MATRIX(1 To NTRIALS, 1 To NROWS + 1)
    
        ReDim PARAM_VECTOR(1 To UBound(XTEMP_VECTOR), 1 To 1)
        For jj = 1 To UBound(XTEMP_VECTOR)
            ATEMP_MATRIX(1, jj + 1) = (XTEMP_VECTOR(jj, 2) - _
                XTEMP_VECTOR(jj, 1)) * Rnd + XTEMP_VECTOR(jj, 1)
            PARAM_VECTOR(jj, 1) = ATEMP_MATRIX(1, jj + 1)
        Next jj
        ATEMP_MATRIX(1, 1) = _
            MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, PARAM_VECTOR, _
                                        SCALE_VECTOR, MIN_FLAG)
        For ii = 2 To UBound(ATEMP_MATRIX)
            For jj = 1 To UBound(XTEMP_VECTOR)
                ATEMP_MATRIX(ii, jj + 1) = ATEMP_MATRIX(ii - 1, jj + 1)
            Next jj
            jj = (ii Mod UBound(XTEMP_VECTOR)) + 1
            ATEMP_MATRIX(ii, jj + 1) = (XTEMP_VECTOR(jj, 2) - _
                XTEMP_VECTOR(jj, 1)) * Rnd + XTEMP_VECTOR(jj, 1)
            PARAM_VECTOR(jj, 1) = ATEMP_MATRIX(ii, jj + 1)
            
            ATEMP_MATRIX(ii, 1) = _
                MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, _
                    PARAM_VECTOR, SCALE_VECTOR, MIN_FLAG)
        Next ii

'    If Not MIN_FLAG Then
 '       For ii = 1 To UBound(ATEMP_MATRIX)
  '          ATEMP_MATRIX(ii, 1) = -ATEMP_MATRIX(ii, 1)
   '     Next ii
    'End If
    
    'Extracts the subset of the lowest ATEMP_MATRIX -----------------------
    For ii = 1 To NTRIALS
        TEMP_VALUE = ATEMP_MATRIX(ii, 1)
        If ii <= NSIZE And hh = 0 Then
            For jj = 1 To NROWS + 1
                BTEMP_MATRIX(ii, jj) = ATEMP_MATRIX(ii, jj)
            Next jj
            If ii = NSIZE Then: BTEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(BTEMP_MATRIX, 1, 1)
        Else
            hh = 1 'stack full
            If TEMP_VALUE < BTEMP_MATRIX(NSIZE, 1) Then
                For jj = 1 To NROWS + 1
                    BTEMP_MATRIX(NSIZE, jj) = ATEMP_MATRIX(ii, jj)
                Next jj
                BTEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(BTEMP_MATRIX, 1, 1)
            End If
        End If
    Next
    
    '--- computes the average and the standard deviation ----------
    For jj = 1 To NROWS
        For ii = 1 To NSIZE
            XMEAN_VECTOR(jj, 1) = XMEAN_VECTOR(jj, 1) + BTEMP_MATRIX(ii, jj + 1)
        Next ii
        XMEAN_VECTOR(jj, 1) = XMEAN_VECTOR(jj, 1) / NSIZE 'average
    Next jj
    For jj = 1 To NROWS
        For ii = 1 To NSIZE
            XDEAV_VECTOR(jj, 1) = XDEAV_VECTOR(jj, 1) + _
                    (BTEMP_MATRIX(ii, jj + 1) - XMEAN_VECTOR(jj, 1)) ^ 2
        Next ii
        XDEAV_VECTOR(jj, 1) = Sqr(XDEAV_VECTOR(jj, 1) / NSIZE) 'standard deviation
    Next jj
    MIN_FUNC_VAL = 0
    For ii = 1 To NSIZE
        MIN_FUNC_VAL = MIN_FUNC_VAL + BTEMP_MATRIX(ii, 1)
    Next ii
    MIN_FUNC_VAL = MIN_FUNC_VAL / NSIZE
    YTEMP_ERR = 0
    For ii = 1 To NSIZE
        YTEMP_ERR = YTEMP_ERR + (BTEMP_MATRIX(ii, 1) - MIN_FUNC_VAL) ^ 2
    Next ii
    YTEMP_ERR = Sqr(YTEMP_ERR / NSIZE)
    '
    XTEMP_ERR = 0
    For ii = 1 To NROWS
        'set constrains
        If XMEAN_VECTOR(ii, 1) < CONST_BOX(ii, 1) Then _
            XMEAN_VECTOR(ii, 1) = CONST_BOX(ii, 1)
        If XMEAN_VECTOR(ii, 1) > CONST_BOX(ii, 2) Then _
            XMEAN_VECTOR(ii, 1) = CONST_BOX(ii, 2)
        TEMP_FACT = XDEAV_VECTOR(ii, 1)
        If (CONST_BOX(ii, 2) - CONST_BOX(ii, 1)) > 0 Then _
            TEMP_FACT = TEMP_FACT / (CONST_BOX(ii, 2) - CONST_BOX(ii, 1))
        If TEMP_FACT < 1 / 16 Then
            XTEMP_VECTOR(ii, 1) = XMEAN_VECTOR(ii, 1) - _
                4 * XDEAV_VECTOR(ii, 1) 'Xmin
            XTEMP_VECTOR(ii, 2) = XMEAN_VECTOR(ii, 1) + _
                4 * XDEAV_VECTOR(ii, 1) 'Xmax
        ElseIf TEMP_FACT < 1 / 4 Then
            XTEMP_VECTOR(ii, 1) = XMEAN_VECTOR(ii, 1) - _
                2 * XDEAV_VECTOR(ii, 1) 'Xmin
            XTEMP_VECTOR(ii, 2) = XMEAN_VECTOR(ii, 1) + _
            2 * XDEAV_VECTOR(ii, 1) 'Xmax
        Else
            XTEMP_VECTOR(ii, 1) = XMEAN_VECTOR(ii, 1) - _
                0.5 * (CONST_BOX(ii, 2) - CONST_BOX(ii, 1))  'Xmin
            XTEMP_VECTOR(ii, 2) = XMEAN_VECTOR(ii, 1) + 0.5 * _
                (CONST_BOX(ii, 2) - CONST_BOX(ii, 1))  'Xmax
        End If
        XTEMP_ERR = XTEMP_ERR + TEMP_FACT  'relative error
        If XTEMP_VECTOR(ii, 1) < CONST_BOX(ii, 1) Then _
            XTEMP_VECTOR(ii, 1) = CONST_BOX(ii, 1)
        If XTEMP_VECTOR(ii, 2) > CONST_BOX(ii, 2) Then _
            XTEMP_VECTOR(ii, 2) = CONST_BOX(ii, 2)
    Next ii
    XTEMP_ERR = XTEMP_ERR / NROWS
    COUNTER = COUNTER + NTRIALS
    '--------------------------------------------------------
    For jj = 1 To NROWS
        XMEAN_VECTOR(jj, 1) = BTEMP_MATRIX(1, jj + 1)
    Next jj
    TEMP_VALUE = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, _
                XMEAN_VECTOR, SCALE_VECTOR, MIN_FLAG)
    '--------------------------------------------------
    If XTEMP_ERR < 0.5 * DELTA_ERR Then
        DELTA_ERR = XTEMP_ERR
        kk = 0
    Else
        NTRIALS = 2 * NTRIALS   'increment the points set
        If NTRIALS > 0.1 * nLOOPS Then NTRIALS = 0.1 * nLOOPS 'limit
        kk = kk + 1
        If kk > 3 Then Exit Do  'useless to contine
    End If
   
    If COUNTER > nLOOPS Then GoTo 1983
    'relative point error check
    If XTEMP_ERR < epsilon Then Exit Do
    '--------------------------------------------------------
Loop  'end of he main loop

1983:

Select Case OUTPUT
    Case 0
        BTEMP_MATRIX = XTEMP_VECTOR
        ReDim XTEMP_VECTOR(1 To UBound(BTEMP_MATRIX, 1), 1 To 1)
        For ii = 1 To UBound(XTEMP_VECTOR, 1)
            'XTEMP_VECTOR(ii, 1) = BTEMP_MATRIX(ii, 1) * SCALE_VECTOR(ii, 1)
            XTEMP_VECTOR(ii, 1) = BTEMP_MATRIX(ii, 2) * SCALE_VECTOR(ii, 1)
        Next ii
        
        MULTVAR_MC_OPTIM_FUNC = XTEMP_VECTOR
    Case 1
        For ii = 1 To UBound(XTEMP_VECTOR, 1)
            XTEMP_VECTOR(ii, 1) = XTEMP_VECTOR(ii, 1) * SCALE_VECTOR(ii, 1)
            XTEMP_VECTOR(ii, 2) = XTEMP_VECTOR(ii, 2) * SCALE_VECTOR(ii, 1)
        Next ii
        
        MULTVAR_MC_OPTIM_FUNC = XTEMP_VECTOR
    Case 2
        MULTVAR_MC_OPTIM_FUNC = MIN_FUNC_VAL
    Case 3
        MULTVAR_MC_OPTIM_FUNC = COUNTER
    Case 4
        MULTVAR_MC_OPTIM_FUNC = XTEMP_ERR
    Case Else
        MULTVAR_MC_OPTIM_FUNC = YTEMP_ERR
End Select

Exit Function
ERROR_LABEL:
MULTVAR_MC_OPTIM_FUNC = Err.number
End Function
