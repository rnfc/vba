Attribute VB_Name = "OPTIM_NLE_ROOTS_LIBR"


'// PERFECT

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : NLE_NEWTON_MAEHLY_FUNC

'DESCRIPTION   : Global rootfinder;
'This function attempts to finds all roots of a nonlinear system in a given
'space range using the random searching method + the Newton- Maehly formula
'for zeros suppression.

'This functions works inside a specific box range and does not need any
'starting point. It is quite time expensive and, like other rootfinder
'algorithms, there is no guarantee that the process succeeds. If the
'function takes too long try to reduce the searching area.

'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_ROOTS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function NLE_NEWTON_MAEHLY_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef CONST_RNG As Variant, _
Optional ByVal RELAX_FLAG As Boolean = False, _
Optional ByVal TRACE_FLAG As Boolean = True, _
Optional ByVal FAIL_PERC As Double = 0.1, _
Optional ByVal trials As Long = 100, _
Optional ByVal nLOOPS As Long = 40, _
Optional ByVal epsilon As Double = 0.000000000000001)

'--------------------------------------------------------------------
'INSIGHTFUL COMMENT:
'--------------------------------------------------------------------
'Solving equations with 3 or more variables may be
'very difficult because generally we cannot use the
'graph method. As we have seen it is necessary to
'locate the space region where the roots are with
'sufficiently precision. Or, at the least, have an idea
'of the limits where the variables can move. This is
'necessary for setting the constraints box.
'The variable bounding can be often discover by
'examining the equations of the system itself.
'--------------------------------------------------------------------


Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long
Dim NSIZE As Double
Dim NO_VAR As Long
Dim NO_ROOTS As Long

Dim TEMP_ERR As Double
Dim DELTA_ERR As Double

Dim TEMP_SUM As Double
Dim TEMP_DELTA As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim RELAX_VAL As Double
Dim BOUND_VAL As Double
Dim CONVERG_VAL As Double

Dim CONST_BOX As Variant
Dim CONST_DATA As Variant

Dim FIT_VECTOR As Variant
Dim PARAM_VECTOR As Variant
Dim XTEMP_VECTOR As Variant

'Dim TEMP_ARR As Variant
Dim TEMP_MATRIX As Variant
Dim TRACE_MATRIX As Variant
Dim JACOBI_MATRIX As Variant
Dim ROOTS_MATRIX As Variant

Dim LAMBDA As Double
Dim tolerance As Double

On Error GoTo ERROR_LABEL
CONVERG_VAL = -1
CONST_DATA = CONST_RNG
ReDim CONST_BOX(1 To UBound(CONST_DATA, 2), 1 To 2)
For i = 1 To UBound(CONST_DATA, 2)
    CONST_BOX(i, 1) = CONST_DATA(1, i)       'Xmin
    CONST_BOX(i, 2) = CONST_DATA(2, i)       'Xmax
Next i

NO_VAR = UBound(CONST_BOX, 1)

ReDim PARAM_VECTOR(1 To NO_VAR, 1 To 1)

NROWS = 10
ReDim ROOTS_MATRIX(1 To NROWS, 1 To NO_VAR)
NO_ROOTS = 0

NSIZE = trials * FAIL_PERC
If NSIZE < 5 Then NSIZE = 5
Randomize
For k = 1 To trials
    'generate random point in the CONST_BOX domain
    For i = 1 To NO_VAR
        PARAM_VECTOR(i, 1) = CONST_BOX(i, 1) + _
                            (CONST_BOX(i, 2) - CONST_BOX(i, 1)) * Rnd
    Next i
    
    NO_VAR = UBound(PARAM_VECTOR)
    ReDim GRAD_VECTOR(1 To NO_VAR, 1 To 1)
    ReDim DERIV_VECTOR(1 To NO_VAR, 1 To 1)

    jj = 0
    CONVERG_VAL = -1
    tolerance = 0.001
    LAMBDA = 10 ^ 99
    
    XTEMP_VECTOR = PARAM_VECTOR
    Do 'iteration Newton begins
        RELAX_VAL = 1#
        Do
            FIT_VECTOR = MODIFIED_ROOT_POLES_FUNC(FUNC_NAME_STR, PARAM_VECTOR, _
                            ROOTS_MATRIX, NO_ROOTS, 10 ^ -10)
            TEMP_SUM = 0
            For i = 1 To UBound(FIT_VECTOR, 1)
                If Abs(FIT_VECTOR(i, 1)) > TEMP_SUM Then _
                    TEMP_SUM = Abs(FIT_VECTOR(i, 1))
            Next i
            TEMP_ERR = TEMP_SUM
                
            If RELAX_FLAG = False Or (jj = 0) Or (TEMP_ERR <= 1.5 * DELTA_ERR) _
            Then Exit Do
            
            RELAX_VAL = RELAX_VAL / 2 'try with a short distance point
            For i = 1 To NO_VAR
                GRAD_VECTOR(i, 1) = RELAX_VAL * GRAD_VECTOR(i, 1)
                PARAM_VECTOR(i, 1) = XTEMP_VECTOR(i, 1) - GRAD_VECTOR(i, 1)
            Next i
        Loop Until RELAX_VAL < 0.1

'check break-off criterion
    
        TEMP_SUM = 0
        For i = 1 To UBound(PARAM_VECTOR, 1)
            If Abs(PARAM_VECTOR(i, 1)) > TEMP_SUM Then _
                TEMP_SUM = Abs(PARAM_VECTOR(i, 1))
        Next i
        ATEMP_VAL = TEMP_SUM
        
        TEMP_SUM = 0
        For i = 1 To UBound(GRAD_VECTOR, 1)
            If Abs(GRAD_VECTOR(i, 1)) > TEMP_SUM Then _
                TEMP_SUM = Abs(GRAD_VECTOR(i, 1))
        Next i
        BTEMP_VAL = TEMP_SUM
        
        If ATEMP_VAL <= 10 ^ 6 * epsilon Then
            TEMP_DELTA = BTEMP_VAL
        Else: TEMP_DELTA = BTEMP_VAL / ATEMP_VAL
        End If
        If (jj > 0) And TEMP_DELTA <= epsilon Then
            CONVERG_VAL = 0
            Exit Do
        End If
        If TEMP_ERR > LAMBDA Then
            CONVERG_VAL = -1
            GoTo 1982
        End If
    
        'save old values
        DELTA_ERR = TEMP_ERR
        XTEMP_VECTOR = PARAM_VECTOR
    '------------------------------------------------------------------------------
        JACOBI_MATRIX = JACOBI_EVALUATE_MODIFIED_FUNC(FUNC_NAME_STR, _
                    PARAM_VECTOR, ROOTS_MATRIX, NO_ROOTS, 10 ^ -5)
    '---------------------evaluate Jacobian of the modified functions
     
    '    ReDim TEMP_MATRIX(1 To UBound(JACOBI_MATRIX, 1), 1 To UBound(JACOBI_MATRIX, 1) + 1)
     '   'build the system matrix
     '   For i = 1 To UBound(JACOBI_MATRIX, 1)
      '      For j = 1 To UBound(JACOBI_MATRIX, 1)
       '         TEMP_MATRIX(i, j) = JACOBI_MATRIX(i, j)
        '    Next j
        'Next i
      '  For i = 1 To UBound(JACOBI_MATRIX, 1)
       '     TEMP_MATRIX(i, UBound(JACOBI_MATRIX, 1) + 1) = FIT_VECTOR(i, 1)
        'Next i
    
    '-----------------------------------------------------------------------
        'TEMP_ARR = MATRIX_TRIANGULAR_LINEAR_SYSTEM_FUNC(TEMP_MATRIX)
    '-----------------------------------------------------------------------
        TEMP_MATRIX = MATRIX_LU_LINEAR_SYSTEM_FUNC(JACOBI_MATRIX, FIT_VECTOR)
    '-----------------------------------------------------------------------
    
        'TEMP_DET = TEMP_ARR(2)
        'If TEMP_DET <> 0 Then
        If IsArray(TEMP_MATRIX) = True Then
            'TEMP_MATRIX = TEMP_ARR(1)
            For i = 1 To UBound(JACOBI_MATRIX, 1)
                GRAD_VECTOR(i, 1) = TEMP_MATRIX(i, 1)
            Next i
            
            'take the Newton increment
            For i = 1 To NO_VAR
                PARAM_VECTOR(i, 1) = PARAM_VECTOR(i, 1) - GRAD_VECTOR(i, 1)
            Next i
        Else
1981:
            'take a random increment
            
            For i = 1 To NO_VAR
                PARAM_VECTOR(i, 1) = PARAM_VECTOR(i, 1) - tolerance * _
                    (Rnd - 0.5) * (PARAM_VECTOR(i, 1) + 1)
            Next i
        End If
        'check constrain
        
        BOUND_VAL = 0
        For i = 1 To UBound(PARAM_VECTOR)
            If PARAM_VECTOR(i, 1) < CONST_BOX(i, 1) Or _
               PARAM_VECTOR(i, 1) > CONST_BOX(i, 2) Then
                BOUND_VAL = i
                Exit For
            End If
        Next i
        
        If BOUND_VAL > 0 Then
            CONVERG_VAL = -1
            GoTo 1982
        End If
        jj = jj + 1
    Loop Until jj > nLOOPS

1982:
    
    BOUND_VAL = 0
    For i = 1 To UBound(PARAM_VECTOR)
        If PARAM_VECTOR(i, 1) < CONST_BOX(i, 1) Or _
           PARAM_VECTOR(i, 1) > CONST_BOX(i, 2) Then
            BOUND_VAL = i
            Exit For
        End If
    Next i
 
    If BOUND_VAL > 0 Then CONVERG_VAL = -1

    If CONVERG_VAL = -1 Then 'convergence or out of bound;
    'try another point
        ii = ii + 1
    Else
        'root found; add to the array
        For i = 1 To UBound(PARAM_VECTOR)
            If Abs(PARAM_VECTOR(i, 1)) < 10 * epsilon Then _
            PARAM_VECTOR(i, 1) = 0
        Next i
        
        If MATRIX_CROSS_CHECK_VECTOR_FUNC(ROOTS_MATRIX, PARAM_VECTOR, _
            100 * Sqr(epsilon)) = 0 Then
            NO_ROOTS = NO_ROOTS + 1
            NROWS = UBound(ROOTS_MATRIX)
            If NO_ROOTS > NROWS Then
                NROWS = 2 * NROWS
                ROOTS_MATRIX = _
                MATRIX_RESIZE_FUNC(ROOTS_MATRIX, NROWS, NO_VAR)
            End If
            
            ROOTS_MATRIX = MATRIX_INSERT_VALUE_FUNC(ROOTS_MATRIX, _
                           PARAM_VECTOR, NO_ROOTS)
            
            If TRACE_FLAG = True Then
                If NO_ROOTS = 1 Then
                    ReDim TRACE_MATRIX(1 To 1, 1 To (1 + NO_VAR + 3))
                    TRACE_MATRIX(1, 1) = "SOLUTION"
                    For j = 1 To NO_VAR
                        TRACE_MATRIX(1, 1 + j) = "X" & j & ":"
                    Next j
                    TRACE_MATRIX(1, 1 + NO_VAR + 1) = "ERROR"
                    TRACE_MATRIX(1, 1 + NO_VAR + 2) = "ITERATION"
                    TRACE_MATRIX(1, 1 + NO_VAR + 3) = "TRIAL"
                End If
            
                TRACE_MATRIX = MATRIX_RESIZE_FUNC(TRACE_MATRIX, 1 + _
                                NO_ROOTS, 1 + NO_VAR + 3)
                TRACE_MATRIX(1 + NO_ROOTS, 1) = "ROOT " & NO_ROOTS & ":"
                For j = 1 To NO_VAR
                    If IsEmpty(ROOTS_MATRIX(NO_ROOTS, j)) = False Then
                        TRACE_MATRIX(1 + NO_ROOTS, 1 + j) = _
                            ROOTS_MATRIX(NO_ROOTS, j)
                    End If
                Next j
                TRACE_MATRIX(1 + NO_ROOTS, 1 + NO_VAR + 1) = TEMP_ERR
                TRACE_MATRIX(1 + NO_ROOTS, 1 + NO_VAR + 2) = jj
                TRACE_MATRIX(1 + NO_ROOTS, 1 + NO_VAR + 3) = k
            End If
    
            ii = 0  'reset trials jj
        Else
            ii = ii + 1
        End If
    End If
    If ii > NSIZE Then Exit For  'nothing to do
Next k
                        
'If NO_ROOTS = 0 Then GoTo ERROR_LABEL 'no solution found

If TRACE_FLAG = True Then
    NLE_NEWTON_MAEHLY_FUNC = TRACE_MATRIX
Else
    For i = 1 To UBound(PARAM_VECTOR, 1)
        PARAM_VECTOR(i, 1) = ROOTS_MATRIX(NO_ROOTS, i)
    Next i
    NLE_NEWTON_MAEHLY_FUNC = PARAM_VECTOR
End If

Exit Function
ERROR_LABEL:
'-------------------------------------------------------------------------------------
' CONVERG_VAL
'      1 convergence reached: abolute residual |f(x)| < tol
'      0 convergence reached: relative error |dx/x| < tol
'     -1 convergence failed.
'-------------------------------------------------------------------------------------
'The number of equations must be equal to the variables one.
'-------------------------------------------------------------------------------------
    NLE_NEWTON_MAEHLY_FUNC = CONVERG_VAL     'convergence not met
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NLE_NEWTON_RAPHSON_FUNC

'DESCRIPTION   : Newton-Rapshon algorithm with relaxation strategy
'This algorithm is the prototype for nearly all NLE solvers. The
'method works quite well. In the general case, the method
'converges rapidly (quadratically) towards a solution. The
'drawbacks of the method are that the Jacobian is expensive to
'calculate, and there is no guarantee that a root will ever be
'found unless your starting value is close to the root. This
'version adopts a relaxation strategy to improve the
'convergence stability

'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_ROOTS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function NLE_NEWTON_RAPHSON_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef PARAM_RNG As Variant, _
Optional ByVal RELAX_FLAG As Boolean = True, _
Optional ByVal TRACE_FLAG As Boolean = False, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal epsilon As Double = 0.000000000000001)

' PARAM_RNG = starting point; at the end PARAM_VECTOR contains the solution

'RELAX_FLAG: Switches on /off the relaxation strategy. If disabled the
'simple traditional Newton-Raphson algorithm is used. If enabled the macro
'exhibits a better global convergence behaviour. In any case this parameter
'does not affect the final accuracy.

Dim i As Long
Dim j As Long

Dim NO_VAR As Long
Dim COUNTER As Long

Dim TEMP_ERR As Double
Dim TEMP_RESID As Double

Dim RELAX_VAL As Double
Dim CONVERG_VAL As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double

Dim FIT_VECTOR As Variant
Dim GRAD_VECTOR As Variant
Dim XTEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim TRACE_MATRIX As Variant
Dim JACOBI_MATRIX As Variant

Dim GAMMA As Double
Dim LAMBDA As Double
Dim tolerance As Double

On Error GoTo ERROR_LABEL
CONVERG_VAL = -1

If IsArray(PARAM_RNG) = False Then
    ReDim PARAM_VECTOR(1 To 1, 1 To 1)
    PARAM_VECTOR(1, 1) = PARAM_RNG
Else
    PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If

NO_VAR = UBound(PARAM_VECTOR, 1)
ReDim GRAD_VECTOR(1 To NO_VAR, 1 To 1)
If TRACE_FLAG = True Then: ReDim TRACE_MATRIX(1 To NO_VAR + 1, 1 To nLOOPS)

COUNTER = 0
CONVERG_VAL = -1
tolerance = 0.001
GAMMA = 10 ^ 99

XTEMP_VECTOR = PARAM_VECTOR
TEMP_RESID = 0


Do
    
    If TRACE_FLAG = True Then
    'Switches on /off the trace of the root trajectory. If selected,
    'the macro opens an auxiliary input box requiring the cell where the
    'output will begin.
        For i = 1 To NO_VAR
            TRACE_MATRIX(i, COUNTER + 1) = PARAM_VECTOR(i, 1)
        Next i
        TRACE_MATRIX(NO_VAR + 1, COUNTER + 1) = TEMP_RESID
        'The input box sets the
        'error limit of the residual error defined as: max{|fi(x)|}.
    End If
    
    RELAX_VAL = 1#
    Do
        FIT_VECTOR = Excel.Application.Run(FUNC_NAME_STR, PARAM_VECTOR)
        'evaluate function values at point x
    CTEMP_VAL = 0
        For j = 1 To UBound(FIT_VECTOR, 2)
            For i = 1 To UBound(FIT_VECTOR, 1)
                If Abs(FIT_VECTOR(i, j)) > CTEMP_VAL Then _
                    CTEMP_VAL = Abs(FIT_VECTOR(i, j))
            Next i
        Next j
        TEMP_RESID = CTEMP_VAL
        
        If RELAX_FLAG = False Or (COUNTER = 0) Or (TEMP_RESID <= 1.5 * LAMBDA) _
           Then Exit Do
        RELAX_VAL = RELAX_VAL / 2
        'try with a short distance point
        For i = 1 To NO_VAR
            GRAD_VECTOR(i, 1) = RELAX_VAL * GRAD_VECTOR(i, 1)
            PARAM_VECTOR(i, 1) = XTEMP_VECTOR(i, 1) - GRAD_VECTOR(i, 1)
        Next i

    Loop Until RELAX_VAL < 0.1

    'check break-off criterion -----------------------
    CTEMP_VAL = 0
        For i = 1 To UBound(PARAM_VECTOR, 1)
                If Abs(PARAM_VECTOR(i, 1)) > CTEMP_VAL Then _
                    CTEMP_VAL = Abs(PARAM_VECTOR(i, 1))
        Next i
        ATEMP_VAL = CTEMP_VAL
    
    CTEMP_VAL = 0
        For i = 1 To UBound(GRAD_VECTOR, 1)
            If Abs(GRAD_VECTOR(i, 1)) > CTEMP_VAL Then _
                    CTEMP_VAL = Abs(GRAD_VECTOR(i, 1))
        Next i
        BTEMP_VAL = CTEMP_VAL
    
    If ATEMP_VAL <= 10 ^ 6 * epsilon Then
        TEMP_ERR = BTEMP_VAL
    Else: TEMP_ERR = BTEMP_VAL / ATEMP_VAL
    End If
    
    If (COUNTER > 0) And TEMP_ERR <= epsilon Then CONVERG_VAL = 0: Exit Do
    If TEMP_RESID > GAMMA Then GoTo ERROR_LABEL
    '----------------------------------------------------
    'save old values
    LAMBDA = TEMP_RESID
    XTEMP_VECTOR = PARAM_VECTOR
    JACOBI_MATRIX = JACOBI_CENTRAL_FUNC(FUNC_NAME_STR, XTEMP_VECTOR, 10 ^ -5)
    
'    ReDim TEMP_MATRIX(1 To NO_VAR, 1 To NO_VAR + 1)
    
    'build the system matrix
 '   For i = 1 To NO_VAR
  '      For j = 1 To NO_VAR
   '         TEMP_MATRIX(i, j) = JACOBI_MATRIX(i, j)
    '    Next j
    'Next i
    'For i = 1 To NO_VAR
     '   TEMP_MATRIX(i, NO_VAR + 1) = FIT_VECTOR(i, 1)
    'Next i
    
    'TEMP_ARR = MATRIX_TRIANGULAR_LINEAR_SYSTEM_FUNC(TEMP_MATRIX)
    'DETERM_VAL = TEMP_ARR(UBOUND(TEMP_ARR))
    
    TEMP_MATRIX = MATRIX_LU_LINEAR_SYSTEM_FUNC(JACOBI_MATRIX, FIT_VECTOR)
    
    
'    If DETERM_VAL <> 0 Then        'take the Newton increment
    If IsArray(TEMP_MATRIX) = True Then
'        TEMP_MATRIX = TEMP_ARR(LBOUND(TEMP_ARR))
        For i = 1 To NO_VAR
            GRAD_VECTOR(i, 1) = TEMP_MATRIX(i, 1)
            PARAM_VECTOR(i, 1) = PARAM_VECTOR(i, 1) - GRAD_VECTOR(i, 1)
        Next i
    Else
        'take a random increment
        For i = 1 To NO_VAR
            GRAD_VECTOR(i, 1) = tolerance * _
                            (Rnd - 0.5) * (PARAM_VECTOR(i, 1) + 1)
            PARAM_VECTOR(i, 1) = PARAM_VECTOR(i, 1) - GRAD_VECTOR(i, 1)
        Next i
    End If
    COUNTER = COUNTER + 1

Loop Until COUNTER > nLOOPS


If TRACE_FLAG = True Then
    NLE_NEWTON_RAPHSON_FUNC = MATRIX_TRIM_FUNC(TRACE_MATRIX, 1, "")
Else
    NLE_NEWTON_RAPHSON_FUNC = PARAM_VECTOR
End If

Exit Function
ERROR_LABEL:
'-------------------------------------------------------------------------------------
' CONVERG_VAL
'      1 convergence reached: abolute residual |f(x)| < tol
'      0 convergence reached: relative error |dx/x| < tol
'     -1 convergence failed.
'-------------------------------------------------------------------------------------
'The number of equations must be equal to the variables one.
'-------------------------------------------------------------------------------------

    NLE_NEWTON_RAPHSON_FUNC = CONVERG_VAL     'convergence not met
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NLE_BROYDEN_FUNC

'DESCRIPTION   : This is a so called Quasi-Newton (Variable Metric) method .It
'avoids the expense of calculating the Jacobian providing a
'more fast and cheap approximation by generalization of the
'one-dimensional secant approximation.

'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_ROOTS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function NLE_BROYDEN_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef PARAM_RNG As Variant, _
Optional ByVal TRACE_FLAG As Boolean = False, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal epsilon As Double = 0.000000000000001)

' PARAM_RNG = starting point; at the end PARAM_VECTOR contains the solution

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NO_VAR As Long
Dim COUNTER As Long

Dim CONVERG_VAL As Double

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double

Dim TEMP_ERR As Double
Dim TEMP_RESID As Double
Dim TEMP_VALUE As Double

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant
Dim CTEMP_MATRIX As Variant
Dim DTEMP_MATRIX As Variant
Dim ETEMP_MATRIX As Variant
Dim FTEMP_MATRIX As Variant

Dim TRACE_MATRIX As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim CTEMP_VECTOR As Variant
Dim DTEMP_VECTOR As Variant

Dim PARAM_VECTOR As Variant

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL
CONVERG_VAL = -1

If IsArray(PARAM_RNG) = False Then
    ReDim PARAM_VECTOR(1 To 1, 1 To 1)
    PARAM_VECTOR(1, 1) = PARAM_RNG
Else
    PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If

NO_VAR = UBound(PARAM_VECTOR, 1)

ReDim ATEMP_VECTOR(1 To NO_VAR, 1 To 1)

ReDim DTEMP_VECTOR(1 To NO_VAR, 1 To 1)
ReDim YTEMP_VECTOR(1 To NO_VAR, 1 To 1)
ReDim DTEMP_MATRIX(1 To NO_VAR, 1 To 1)

ReDim FTEMP_MATRIX(1 To 1, 1 To NO_VAR)
ReDim BTEMP_VECTOR(1 To NO_VAR, 1 To 1)
ReDim CTEMP_VECTOR(1 To 1, 1 To NO_VAR)

ReDim ATEMP_MATRIX(1 To NO_VAR, 1 To NO_VAR)
ReDim BTEMP_MATRIX(1 To NO_VAR, 1 To NO_VAR)

If TRACE_FLAG = True Then: ReDim TRACE_MATRIX(1 To NO_VAR + 1, 1 To nLOOPS)

CONVERG_VAL = -1
XTEMP_VECTOR = Excel.Application.Run(FUNC_NAME_STR, PARAM_VECTOR)
For i = 1 To NO_VAR
    ATEMP_VECTOR(i, 1) = XTEMP_VECTOR(i, 1)
Next i
COUNTER = 0
TEMP_RESID = 0

    If TRACE_FLAG = True Then
    'Switches on /off the trace of the root trajectory. If selected,
    'the macro opens an auxiliary input box requiring the cell where the
    'output will begin.
        For i = 1 To NO_VAR
            TRACE_MATRIX(i, COUNTER + 1) = PARAM_VECTOR(i, 1)
        Next i
        TRACE_MATRIX(NO_VAR + 1, COUNTER + 1) = TEMP_RESID 'The input box sets the
        'error limit of the residual error defined as: max{|fi(x)|}.
    End If
        
    ATEMP_MATRIX = JACOBI_INVERSE_FUNC(FUNC_NAME_STR, PARAM_VECTOR, 0.00001)
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

    ReDim CTEMP_MATRIX(1 To UBound(ATEMP_MATRIX, 1), 1 To UBound(ATEMP_VECTOR, 2))
    
    For jj = 1 To UBound(ATEMP_VECTOR, 2)
        For ii = 1 To UBound(ATEMP_MATRIX, 1)
            For kk = 1 To UBound(ATEMP_MATRIX, 2)
                CTEMP_MATRIX(ii, jj) = CTEMP_MATRIX(ii, jj) + _
                    ATEMP_MATRIX(ii, kk) * ATEMP_VECTOR(kk, jj)
            Next kk
        Next ii
    Next jj


For i = 1 To NO_VAR
    CTEMP_MATRIX(i, 1) = -CTEMP_MATRIX(i, 1)
    PARAM_VECTOR(i, 1) = PARAM_VECTOR(i, 1) + CTEMP_MATRIX(i, 1)
Next i

COUNTER = 1

Do
    For i = 1 To NO_VAR
        DTEMP_VECTOR(i, 1) = ATEMP_VECTOR(i, 1)
    Next i
    
    XTEMP_VECTOR = Excel.Application.Run(FUNC_NAME_STR, PARAM_VECTOR)
    
    For i = 1 To NO_VAR
        ATEMP_VECTOR(i, 1) = XTEMP_VECTOR(i, 1)
        YTEMP_VECTOR(i, 1) = ATEMP_VECTOR(i, 1) - DTEMP_VECTOR(i, 1)
    Next i
    
    ReDim DTEMP_MATRIX(1 To UBound(ATEMP_MATRIX, 1), 1 To UBound(YTEMP_VECTOR, 2))
    
    For jj = 1 To UBound(YTEMP_VECTOR, 2)
        For ii = 1 To UBound(ATEMP_MATRIX, 1)
            For kk = 1 To UBound(ATEMP_MATRIX, 2)
                DTEMP_MATRIX(ii, jj) = DTEMP_MATRIX(ii, jj) + _
                    ATEMP_MATRIX(ii, kk) * YTEMP_VECTOR(kk, jj)
            Next kk
        Next ii
    Next jj
    
    
    For i = 1 To NO_VAR
        DTEMP_MATRIX(i, 1) = -DTEMP_MATRIX(i, 1)
        CTEMP_VECTOR(1, i) = CTEMP_MATRIX(i, 1)
    Next i
    
    ReDim ETEMP_MATRIX(1 To UBound(CTEMP_VECTOR, 1), 1 To UBound(DTEMP_MATRIX, 2))
    
    For jj = 1 To UBound(DTEMP_MATRIX, 2)
        For ii = 1 To UBound(CTEMP_VECTOR, 1)
            For kk = 1 To UBound(CTEMP_VECTOR, 2)
                ETEMP_MATRIX(ii, jj) = ETEMP_MATRIX(ii, jj) + _
                    CTEMP_VECTOR(ii, kk) * DTEMP_MATRIX(kk, jj)
            Next kk
        Next ii
    Next jj
    
    ETEMP_MATRIX(1, 1) = -ETEMP_MATRIX(1, 1)
    
    ReDim FTEMP_MATRIX(1 To UBound(CTEMP_VECTOR, 1), 1 To UBound(ATEMP_MATRIX, 2))
    
    For jj = 1 To UBound(ATEMP_MATRIX, 2)
        For ii = 1 To UBound(CTEMP_VECTOR, 1)
            For kk = 1 To UBound(CTEMP_VECTOR, 2)
                FTEMP_MATRIX(ii, jj) = FTEMP_MATRIX(ii, jj) + _
                    CTEMP_VECTOR(ii, kk) * ATEMP_MATRIX(kk, jj)
            Next kk
        Next ii
    Next jj
    
    For i = 1 To NO_VAR
        BTEMP_VECTOR(i, 1) = (CTEMP_MATRIX(i, 1) + _
            DTEMP_MATRIX(i, 1)) / (ETEMP_MATRIX(1, 1) + 1E-16)
    Next i
    
    ReDim BTEMP_MATRIX(1 To UBound(BTEMP_VECTOR, 1), 1 To UBound(FTEMP_MATRIX, 2))
    
    For jj = 1 To UBound(FTEMP_MATRIX, 2)
        For ii = 1 To UBound(BTEMP_VECTOR, 1)
            For kk = 1 To UBound(BTEMP_VECTOR, 2)
                BTEMP_MATRIX(ii, jj) = BTEMP_MATRIX(ii, jj) + _
                    BTEMP_VECTOR(ii, kk) * FTEMP_MATRIX(kk, jj)
            Next kk
        Next ii
    Next jj
    
    For i = 1 To NO_VAR
        For j = 1 To NO_VAR
            ATEMP_MATRIX(i, j) = ATEMP_MATRIX(i, j) + BTEMP_MATRIX(i, j)
        Next j
    Next i
    
    ReDim CTEMP_MATRIX(1 To UBound(ATEMP_MATRIX, 1), 1 To UBound(ATEMP_VECTOR, 2))
    
    For jj = 1 To UBound(ATEMP_VECTOR, 2)
        For ii = 1 To UBound(ATEMP_MATRIX, 1)
            For kk = 1 To UBound(ATEMP_MATRIX, 2)
                CTEMP_MATRIX(ii, jj) = CTEMP_MATRIX(ii, jj) + _
                    ATEMP_MATRIX(ii, kk) * ATEMP_VECTOR(kk, jj)
            Next kk
        Next ii
    Next jj
    
    For i = 1 To NO_VAR
        CTEMP_MATRIX(i, 1) = -CTEMP_MATRIX(i, 1)
        PARAM_VECTOR(i, 1) = PARAM_VECTOR(i, 1) + CTEMP_MATRIX(i, 1)
    Next i
    
    XTEMP_VECTOR = Excel.Application.Run(FUNC_NAME_STR, PARAM_VECTOR)
    
    'reinitialize the Broyden matrix after 25 steps
    If COUNTER Mod 25 = 0 Then
        ATEMP_MATRIX = JACOBI_INVERSE_FUNC(FUNC_NAME_STR, PARAM_VECTOR, 0.00001)
    End If
    
    '--- break-off criterions ------------------------------------
        TEMP_VALUE = 0
        For j = 1 To UBound(XTEMP_VECTOR, 2)
            For i = 1 To UBound(XTEMP_VECTOR, 1)
                If Abs(XTEMP_VECTOR(i, j)) > TEMP_VALUE Then _
                    TEMP_VALUE = Abs(XTEMP_VECTOR(i, j))
            Next i
        Next j
        TEMP_RESID = TEMP_VALUE
    
        TEMP_VALUE = 0
        For j = 1 To UBound(CTEMP_MATRIX, 2)
            For i = 1 To UBound(CTEMP_MATRIX, 1)
                If Abs(CTEMP_MATRIX(i, j)) > TEMP_VALUE Then _
                    TEMP_VALUE = Abs(CTEMP_MATRIX(i, j))
            Next i
        Next j
        ATEMP_SUM = TEMP_VALUE
    
        TEMP_VALUE = 0
        For j = 1 To UBound(PARAM_VECTOR, 2)
            For i = 1 To UBound(PARAM_VECTOR, 1)
                If Abs(PARAM_VECTOR(i, j)) > TEMP_VALUE Then _
                    TEMP_VALUE = Abs(PARAM_VECTOR(i, j))
            Next i
        Next j
        BTEMP_SUM = TEMP_VALUE
    
    
    If TRACE_FLAG = True Then
    'Switches on /off the trace of the root trajectory. If selected,
    'the macro opens an auxiliary input box requiring the cell where the
    'output will begin.
        For i = 1 To NO_VAR
            TRACE_MATRIX(i, COUNTER + 1) = PARAM_VECTOR(i, 1)
        Next i
        TRACE_MATRIX(NO_VAR + 1, COUNTER + 1) = TEMP_RESID
        'The input box sets the
        'error limit of the residual error defined as: max{|fi(x)|}.
    End If
    
    
    If BTEMP_SUM > epsilon Then
          TEMP_ERR = ATEMP_SUM / BTEMP_SUM
    Else: TEMP_ERR = ATEMP_SUM
    End If
    
    COUNTER = COUNTER + 1
    If TEMP_ERR < epsilon Then
        CONVERG_VAL = 0
        Exit Do 'convergence OK
    End If
    If TEMP_RESID < epsilon Then
        CONVERG_VAL = 1
        Exit Do 'convergence OK
    End If
    If COUNTER > nLOOPS Then
        CONVERG_VAL = -1
        GoTo ERROR_LABEL 'convergence fails
    End If
    '-------------------------------------------------------
Loop

If TRACE_FLAG = True Then
    NLE_BROYDEN_FUNC = MATRIX_TRIM_FUNC(TRACE_MATRIX, 1, "")
Else
    NLE_BROYDEN_FUNC = PARAM_VECTOR
End If

Exit Function
ERROR_LABEL:
'-------------------------------------------------------------------------------------
' CONVERG_VAL
'      1 convergence reached: abolute residual |f(x)| < tol
'      0 convergence reached: relative error |dx/x| < tol
'     -1 convergence failed.
'-------------------------------------------------------------------------------------
'The number of equations must be equal to the variables one.
'-------------------------------------------------------------------------------------

    NLE_BROYDEN_FUNC = CONVERG_VAL     'convergence not met
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NLE_BROWN_FUNC

'DESCRIPTION   : This algorithm is based on an iterative method which is a
'variation of Newton's method using Gaussian elimination in a
'manner similar to the Gauss-Seidel process. Convergence is
'roughly quadratic. All partial derivatives required by the
'algorithm are approximated by first difference quotients.
'The convergence behavior is affected by the ordering of the
'equations, and it is advantageous to place linear and mildly
'nonlinear equations first in the ordering.

' Brown algorithm
'  sources : Brown K.kkk.: A quadratically convergent Newton-like
'            method based upon Gaussian elimination,
'            Siam J.Numer.Anal.vol 6 (1969), 560 - 569

'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_ROOTS
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function NLE_BROWN_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef PARAM_RNG As Variant, _
Optional ByVal TRACE_FLAG As Boolean = False, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal epsilon As Double = 0.000000000000001)

' PARAM_RNG = starting point; at the end PARAM_VECTOR contains the solution

Dim i As Long
Dim j As Long
Dim k As Long

Dim hh As Long
Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim hhh As Long
Dim iii As Long
Dim jjj As Long
Dim kkk As Long

Dim NO_VAR As Long
Dim COUNTER As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double

Dim ATEMP_DELTA As Double
Dim BTEMP_DELTA As Double

Dim TEMP_MAX As Double
Dim TEMP_SUB As Double
Dim TEMP_HOLD As Double
Dim TEMP_TEST As Double
Dim TEMP_PLUS As Double
Dim TEMP_RESID As Double
Dim TEMP_FACTOR As Double

Dim TEMP_VAL As Double
Dim SING_FLAG As Boolean

Dim CONVERG_VAL As Double

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

Dim FIT_VECTOR As Variant
Dim PARAM_VECTOR As Variant
Dim DELTA_VECTOR As Variant
Dim XTEMP_VECTOR As Variant

Dim LAMBDA As Double
Dim tolerance As Double

On Error GoTo ERROR_LABEL
CONVERG_VAL = -1

If IsArray(PARAM_RNG) = False Then
    ReDim PARAM_VECTOR(1 To 1, 1 To 1)
    PARAM_VECTOR(1, 1) = PARAM_RNG
Else
    PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
End If

NO_VAR = UBound(PARAM_VECTOR, 1)

ReDim ATEMP_MATRIX(1 To NO_VAR, 1 To NO_VAR + 3)
ReDim BTEMP_MATRIX(1 To NO_VAR, 1 To NO_VAR + 1)

ReDim FIT_VECTOR(1 To NO_VAR, 1 To 1)
ReDim XTEMP_VECTOR(1 To NO_VAR, 1 To 1)

COUNTER = 0
SING_FLAG = False
TEMP_VAL = 0
ATEMP_DELTA = 0.01
tolerance = 2 * 10 ^ -16

If TRACE_FLAG = True Then: _
    ReDim TRACE_MATRIX(1 To NO_VAR + 1, 1 To nLOOPS)

    If TRACE_FLAG = True Then
    'Switches on /off the trace of the root trajectory. If selected,
    'the macro opens an auxiliary input box requiring the cell where the
    'output will begin.
        For j = 1 To NO_VAR
            TRACE_MATRIX(j, COUNTER + 1) = PARAM_VECTOR(j, 1)
        Next j
        TRACE_MATRIX(NO_VAR + 1, COUNTER + 1) = TEMP_RESID 'The input box sets the
        'error limit of the residual error defined as: max{|fi(x)|}.
    End If

    For j = 1 To NO_VAR
        ATEMP_MATRIX(j, NO_VAR + 2) = PARAM_VECTOR(j, 1)
    Next j

'   iteration begins
For kkk = 1 To nLOOPS
    
'------------------FIRST PASS: computes an approximation by Brown'algorithm

    ReDim DELTA_VECTOR(1 To NO_VAR, 1 To 1)
    
    For jjj = 1 To NO_VAR
       BTEMP_MATRIX(1, jjj) = jjj
       XTEMP_VECTOR(jjj, 1) = ATEMP_MATRIX(jjj, NO_VAR + 2)
    Next jjj
'   linearization of the K-th coordinate function
    
    For k = 1 To NO_VAR
       hhh = 0
       TEMP_FACTOR = 0.001
       For jjj = 1 To 3
          If (k > 1) Then
                For ii = k To 2 Step -1
                   hh = BTEMP_MATRIX(ii - 1, NO_VAR + 1)
                   XTEMP_VECTOR(hh, 1) = 0#
                   For jj = ii To NO_VAR
                      kk = BTEMP_MATRIX(ii, jj)
                      XTEMP_VECTOR(hh, 1) = XTEMP_VECTOR(hh, 1) + _
                                            ATEMP_MATRIX(ii - 1, kk) * _
                                            XTEMP_VECTOR(kk, 1)
                   Next jj
                   XTEMP_VECTOR(hh, 1) = XTEMP_VECTOR(hh, 1) + _
                                ATEMP_MATRIX(ii - 1, NO_VAR + 1)
                Next ii
          End If
          
          DELTA_VECTOR = Excel.Application.Run(FUNC_NAME_STR, XTEMP_VECTOR)
          DTEMP_VAL = DELTA_VECTOR(k, 1)

'   determining the iii-th discretization stepsize and the iii-th difference quotient
          
          For iii = k To NO_VAR
             ATEMP_VAL = BTEMP_MATRIX(k, iii)
             TEMP_HOLD = XTEMP_VECTOR(ATEMP_VAL, 1)
             LAMBDA = TEMP_FACTOR * TEMP_HOLD
             If (Abs(LAMBDA) <= tolerance) Then LAMBDA = 0.001
             XTEMP_VECTOR(ATEMP_VAL, 1) = TEMP_HOLD + LAMBDA
             If (k > 1) Then
                For ii = k To 2 Step -1
                   hh = BTEMP_MATRIX(ii - 1, NO_VAR + 1)
                   XTEMP_VECTOR(hh, 1) = 0#
                   For jj = ii To NO_VAR
                      kk = BTEMP_MATRIX(ii, jj)
                      XTEMP_VECTOR(hh, 1) = XTEMP_VECTOR(hh, 1) + _
                        ATEMP_MATRIX(ii - 1, kk) * XTEMP_VECTOR(kk, 1)
                   Next jj
                   XTEMP_VECTOR(hh, 1) = XTEMP_VECTOR(hh, 1) + _
                        ATEMP_MATRIX(ii - 1, NO_VAR + 1)
                Next ii
             
             End If
             
             DELTA_VECTOR = Excel.Application.Run(FUNC_NAME_STR, XTEMP_VECTOR)
             CTEMP_VAL = DELTA_VECTOR(k, 1)
             XTEMP_VECTOR(ATEMP_VAL, 1) = TEMP_HOLD
             ATEMP_MATRIX(ATEMP_VAL, NO_VAR + 3) = _
                (CTEMP_VAL - DTEMP_VAL) / LAMBDA
             
             If (Abs(ATEMP_MATRIX(ATEMP_VAL, NO_VAR + 3)) <= tolerance) Then
                hhh = hhh + 1
             Else
                If (Abs(DTEMP_VAL / ATEMP_MATRIX(ATEMP_VAL, NO_VAR + 3)) >= 1E+20) _
                    Then hhh = hhh + 1
             End If
          Next iii
          If (hhh <= NO_VAR - k) Then
             SING_FLAG = False
             Exit For
          Else
             SING_FLAG = True
             TEMP_FACTOR = TEMP_FACTOR * 10#
             hhh = 0
          End If
       Next jjj
'
       If (Not SING_FLAG) Then
          If (k < NO_VAR) Then
             BTEMP_VAL = BTEMP_MATRIX(k, k)

'   determining the difference quotient of largest magnitude
             
             TEMP_MAX = Abs(ATEMP_MATRIX(BTEMP_VAL, NO_VAR + 3))
             TEMP_PLUS = k + 1
             For iii = TEMP_PLUS To NO_VAR
                TEMP_SUB = BTEMP_MATRIX(k, iii)
                TEMP_TEST = Abs(ATEMP_MATRIX(TEMP_SUB, NO_VAR + 3))
                If (TEMP_TEST < TEMP_MAX) Then
                   BTEMP_MATRIX(TEMP_PLUS, iii) = TEMP_SUB
                Else
                   BTEMP_MATRIX(TEMP_PLUS, iii) = BTEMP_VAL
                   BTEMP_VAL = TEMP_SUB
                End If
             Next iii
             
             If (Abs(ATEMP_MATRIX(BTEMP_VAL, NO_VAR + 3)) <= tolerance) _
                Then SING_FLAG = True
                    BTEMP_MATRIX(k, NO_VAR + 1) = BTEMP_VAL
             If (Not SING_FLAG) Then
                ATEMP_MATRIX(k, NO_VAR + 1) = 0#
'
'   solving the K-th equation for XMAX
'
                For jjj = TEMP_PLUS To NO_VAR
                   TEMP_SUB = BTEMP_MATRIX(TEMP_PLUS, jjj)
                   ATEMP_MATRIX(k, TEMP_SUB) = -ATEMP_MATRIX(TEMP_SUB, NO_VAR + 3) / _
                                             ATEMP_MATRIX(BTEMP_VAL, NO_VAR + 3)
                   ATEMP_MATRIX(k, NO_VAR + 1) = ATEMP_MATRIX(k, NO_VAR + 1) + _
                                             ATEMP_MATRIX(TEMP_SUB, NO_VAR + 3) * _
                                             XTEMP_VECTOR(TEMP_SUB, 1)
                Next jjj
                ATEMP_MATRIX(k, NO_VAR + 1) = (ATEMP_MATRIX(k, NO_VAR + 1) - DTEMP_VAL) / _
                                           ATEMP_MATRIX(BTEMP_VAL, NO_VAR + 3) + _
                                           XTEMP_VECTOR(BTEMP_VAL, 1)
             Else
                GoTo 1983
             End If
          Else
'
'   solving the NO_VAR-th coordinate function by use of the
'   discrete Newton-method for one variable
'
             If (Abs(ATEMP_MATRIX(ATEMP_VAL, NO_VAR + 3)) <= tolerance) Then
                SING_FLAG = True
             Else
                ATEMP_MATRIX(k, NO_VAR + 1) = 0#
                BTEMP_VAL = ATEMP_VAL
                ATEMP_MATRIX(k, NO_VAR + 1) = (ATEMP_MATRIX(k, NO_VAR + 1) - _
                                           DTEMP_VAL) / ATEMP_MATRIX(BTEMP_VAL, _
                                           NO_VAR + 3) + XTEMP_VECTOR(BTEMP_VAL, 1)
             End If
          End If
       Else
          GoTo 1983
       End If
    Next k

1983:


'----------------Determining of the approximate solution by backsubstitution-------------

    If (Not SING_FLAG) Then

       XTEMP_VECTOR(BTEMP_VAL, 1) = ATEMP_MATRIX(NO_VAR, NO_VAR + 1)
       If (NO_VAR > 1) Then
                For ii = NO_VAR To 2 Step -1
                   hh = BTEMP_MATRIX(ii - 1, NO_VAR + 1)
                   XTEMP_VECTOR(hh, 1) = 0#
                   For jj = ii To NO_VAR
                      kk = BTEMP_MATRIX(ii, jj)
                      XTEMP_VECTOR(hh, 1) = XTEMP_VECTOR(hh, 1) + _
                        ATEMP_MATRIX(ii - 1, kk) * XTEMP_VECTOR(kk, 1)
                   Next jj
                   XTEMP_VECTOR(hh, 1) = XTEMP_VECTOR(hh, 1) + _
                                         ATEMP_MATRIX(ii - 1, NO_VAR + 1)
                Next ii
       End If
    End If
    
    COUNTER = kkk
    
    If (Not SING_FLAG) Then

'   Test of  break-off criterion
'   test the relative change of the iterates
       For i = 1 To NO_VAR
          TEMP_RESID = (XTEMP_VECTOR(i, 1) - ATEMP_MATRIX(i, NO_VAR + 2)) _
            / (ATEMP_MATRIX(i, NO_VAR + 2) + epsilon)
          If (Abs(TEMP_RESID) >= epsilon) Then Exit For
       Next i

    If TRACE_FLAG = True Then
    'Switches on /off the trace of the root trajectory. If selected,
    'the macro opens an auxiliary input box requiring the cell where the
    'output will begin.
        For j = 1 To NO_VAR
            TRACE_MATRIX(j, COUNTER + 1) = XTEMP_VECTOR(j, 1)
        Next j
        TRACE_MATRIX(NO_VAR + 1, COUNTER + 1) = TEMP_RESID 'The input box sets the
        'error limit of the residual error defined as: max{|fi(x)|}.
    End If
       
       If i > NO_VAR Then TEMP_VAL = 1: GoTo 1984

'   test the functional value
       FIT_VECTOR = Excel.Application.Run(FUNC_NAME_STR, XTEMP_VECTOR)
       For i = 1 To NO_VAR
          ETEMP_VAL = FIT_VECTOR(i, 1)
          If (Abs(ETEMP_VAL) > tolerance) Then Exit For
       Next i
       
       If i > NO_VAR Then TEMP_VAL = 2: GoTo 1984
'   test the limiting accuracy
       BTEMP_DELTA = Abs(XTEMP_VECTOR(1, 1) - ATEMP_MATRIX(1, NO_VAR + 2))
       For i = 2 To NO_VAR
          BTEMP_DELTA = IIf(BTEMP_DELTA > Abs(XTEMP_VECTOR(i, 1) - _
                        ATEMP_MATRIX(i, NO_VAR + 2)), _
                        BTEMP_DELTA, Abs(XTEMP_VECTOR(i, 1) - _
                        ATEMP_MATRIX(i, NO_VAR + 2)))
       Next i
       If (BTEMP_DELTA <= 0.001) Then
          If (ATEMP_DELTA <= BTEMP_DELTA) Then
             TEMP_VAL = 3
             GoTo 1984
          End If
       End If
       ATEMP_DELTA = BTEMP_DELTA
          For i = 1 To NO_VAR
             ATEMP_MATRIX(i, NO_VAR + 2) = XTEMP_VECTOR(i, 1)
          Next i
    Else
       GoTo 1984
    End If
Next kkk

1984:
    If SING_FLAG Or TEMP_VAL = 3 Then
        CONVERG_VAL = -1
    Else
        For i = 1 To NO_VAR
           PARAM_VECTOR(i, 1) = XTEMP_VECTOR(i, 1)
        Next i
        If (TEMP_VAL = 1) Then
            CONVERG_VAL = 0
        ElseIf (TEMP_VAL = 2) Then
            CONVERG_VAL = 1
        End If
        FIT_VECTOR = Excel.Application.Run(FUNC_NAME_STR, PARAM_VECTOR)
        'update function value
    End If

If TRACE_FLAG = True Then
    NLE_BROWN_FUNC = MATRIX_TRIM_FUNC(TRACE_MATRIX, 1, "")
Else
    NLE_BROWN_FUNC = PARAM_VECTOR
End If

Exit Function
ERROR_LABEL:
'-------------------------------------------------------------------------------------
' CONVERG_VAL
'      1 convergence reached: abolute residual |f(x)| < tol
'      0 convergence reached: relative error |dx/x| < tol
'     -1 convergence failed.
'-------------------------------------------------------------------------------------
'The number of equations must be equal to the variables one.
'-------------------------------------------------------------------------------------

    NLE_BROWN_FUNC = CONVERG_VAL     'convergence not met
End Function
