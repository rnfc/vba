Attribute VB_Name = "OPTIM_MULTVAR_NELDER_LIBR"
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
'FUNCTION      : NELDER_MEAD_OPTIMIZATION_FRAME_FUNC

'DESCRIPTION   : This function is commonly used nonlinear optimization algorithm.
'It is due to Nelder & Mead (1965) and is a numerical method for minimising an
'objective function in a many-dimensional space.

'The method uses the concept of a simplex, which is a polytope of N + 1
'vertices in N dimensions; a line segment on a line, a triangle on a
'plane, a tetrahedron in three-dimensional space and so forth.

'The method approximately finds a locally optimal solution to a problem
'with N variables when the objective function varies smoothly. For example,
'a suspension bridge engineer has to choose how thick each strut, cable
'and pier must be. Clearly these all link together, but it is not easy to
'visualise the impact of changing any specific element. The engineer can
'use the Nelder-Mead method to generate trial designs which are then tested
'on a large computer model. As each run of the simulation is expensive it is
'important to make good decisions about where to look. Nelder-Mead generates
'a new test position by extrapolating the behaviour of the objective function
'measured at each test point arranged as a simplex. The algorithm then chooses
'to Replace one of these test points with the new test point and so the
'algorithm progresses.

'The simplest step is to Replace the worst point with a point reflected
'through the centroid of the remaining N points. If this point is better
'than the best current point, then we can try stretching exponentially out
'along this line. On the other hand, if this new point isn't much better
'than the previous value then we are stepping across a valley, so we shrink
'the simplex towards the best point.

'Like all general purpose multidimensional optimisation algorithms, Nelder-Mead
'occasionally gets stuck in a rut. The standard approach to handle this is to
'restart the algorithm with a new simplex starting at the current best value.
'This can be extended in a similar way to simulated annealing to try and
'escape small local minima.

'Many variations exist depending on the actual nature of problem being solved.
'The most common, perhaps, is to use a constant size small simplex that climbs
'local gradients to local maximums. Visualize a small triangle on an elevation
'map flip flopping its way up a hill to a local peak. This, however, tends to
'perform poorly against the method described in the paper as it makes lots
'of small unnecessary steps in areas of little interest.

'REFERENCE: http://en.wikipedia.org/wiki/Nelder-Mead_method

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_NELDER
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function NELDER_MEAD_OPTIMIZATION_FRAME_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef PARAM_RNG As Variant, _
Optional ByRef CONST_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal epsilon As Double = 0.000000000000001)

Dim i As Long
Dim j As Long

Dim TEMP_ARR As Variant

Dim NO_VAR As Long
Dim CONST_FLAG As Boolean

Dim CONST_DATA As Variant
Dim CONST_BOX As Variant

Dim YTEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant
Dim PARAM_MATRIX As Variant
Dim SCALE_VECTOR As Variant

Dim tolerance As Double

On Error GoTo ERROR_LABEL

tolerance = epsilon

If IsArray(CONST_RNG) = False Then
    CONST_FLAG = False
Else
    CONST_FLAG = True
    CONST_DATA = CONST_RNG
End If

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

'------------------------------------------------------------------------------------
Select Case CONST_FLAG
'------------------------------------------------------------------------------------
Case True 'choose a starting point with constrains
'------------------------------------------------------------------------------------
    NO_VAR = UBound(CONST_DATA, 2) ' How Many Variables
    CONST_BOX = MULTVAR_LOAD_CONST_FUNC(CONST_DATA, 1)
    TEMP_ARR = MULTVAR_SCALE_CONST_FUNC(CONST_BOX)
    CONST_BOX = TEMP_ARR(LBound(TEMP_ARR)) 'rescaling variables
    SCALE_VECTOR = TEMP_ARR(UBound(TEMP_ARR))
    For i = 1 To NO_VAR
         PARAM_VECTOR(i, 1) = PARAM_VECTOR(i, 1) / SCALE_VECTOR(i, 1)
    Next i
    
    ReDim PARAM_MATRIX(1 To NO_VAR + 1, 1 To NO_VAR)  'simplex
    For i = 1 To NO_VAR
        If CONST_BOX(i, 1) > PARAM_VECTOR(i, 1) Then PARAM_VECTOR(i, 1) = CONST_BOX(i, 1)
        If CONST_BOX(i, 2) < PARAM_VECTOR(i, 1) Then PARAM_VECTOR(i, 1) = CONST_BOX(i, 2)
    Next i
    For i = 1 To NO_VAR + 1
        For j = 1 To NO_VAR
             PARAM_MATRIX(i, j) = (PARAM_VECTOR(j, 1) + CONST_BOX(j, 1) + CONST_BOX(j, 2)) / 3
             If i = j Then PARAM_MATRIX(i, j) = PARAM_MATRIX(i, j) + 0.1 * (Rnd - 0.5)
        Next j
    Next i
    YTEMP_VECTOR = NELDER_MEAD_OPTIMIZATION1_FUNC(FUNC_NAME_STR, PARAM_MATRIX, SCALE_VECTOR, CONST_BOX, MIN_FLAG, OUTPUT, nLOOPS, tolerance)
'------------------------------------------------------------------------------------
Case Else 'choose a starting point with no constraint
'------------------------------------------------------------------------------------
    TEMP_ARR = MULTVAR_SCALE_CONST_FUNC(PARAM_VECTOR)
    PARAM_VECTOR = TEMP_ARR(LBound(TEMP_ARR))
    SCALE_VECTOR = TEMP_ARR(UBound(TEMP_ARR))
    NO_VAR = UBound(PARAM_VECTOR, 1)
    ReDim PARAM_MATRIX(1 To NO_VAR + 1, 1 To NO_VAR)  'simplex
    For i = 1 To NO_VAR + 1
        For j = 1 To NO_VAR
             PARAM_MATRIX(i, j) = PARAM_VECTOR(j, 1)
             If i = j Then PARAM_MATRIX(i, j) = PARAM_MATRIX(i, j) + 0.1
        Next j
    Next i
    YTEMP_VECTOR = NELDER_MEAD_OPTIMIZATION1_FUNC(FUNC_NAME_STR, PARAM_MATRIX, SCALE_VECTOR, "", MIN_FLAG, OUTPUT, nLOOPS, tolerance)
'------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------

NELDER_MEAD_OPTIMIZATION_FRAME_FUNC = YTEMP_VECTOR

Exit Function
ERROR_LABEL:
NELDER_MEAD_OPTIMIZATION_FRAME_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : NELDER_MEAD_OPTIMIZATION1_FUNC

'DESCRIPTION   : Optimization of multidimensional nonlinear functions

' Numerical Recipies in Fortran77; W.H. Press, et al.; Cambridge U. Press
' E. Chelouan, et al.; Genetic and Nelder-Mead...;
' EJOR 148(2003) 335-348
' J.C. Lagarias, et al.; Convergence properties...;
' SIAM J Optim. 9(1,1), 112-147

'The Nelder–Mead downhill simplex algorithm is a popular derivative-free
'optimization method. It is based on the idea of function comparisons among a
'simplex of N + 1 points. Depending on the function values, the simplex is
'reflected or shrunk away from the maximum point. Although there are no
'theoretical results on the convergence of the algorithm, it works very well on
'a wide range of practical problems. It is a good choice when a one-optimum
'solution is wanted with minimum programming effort. It can also be used to
'minimize functions that are not differentiable, or that we cannot differentiate.

'It shows a very robust behavior and converges over a very large set of
'starting points. In our experience it is the best general purpose algorithm;
'solid as a rock. It's a jack of all trades.

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_NELDER
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function NELDER_MEAD_OPTIMIZATION1_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef PARAM_RNG As Variant, _
Optional ByRef SCALE_RNG As Variant, _
Optional ByRef CONST_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal tolerance As Double = 0.000000000000001)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long
Dim COUNTER As Long

Dim FIRST_ROOT As Double
Dim SECOND_ROOT As Double
Dim THIRD_ROOT As Double
Dim FORTH_ROOT As Double
Dim TEMP_ROOT As Double

Dim TEMP_DELTA As Double
Dim TEMP_VALUE As Double

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim CTEMP_VECTOR As Variant
Dim DTEMP_VECTOR As Variant
Dim ETEMP_VECTOR As Variant
Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant
Dim ZTEMP_VECTOR As Variant

Dim CONST_BOX As Variant
Dim SCALE_VECTOR As Variant
Dim PARAM_MATRIX As Variant

Dim EXIT_FLAG As Boolean
Dim CONST_FLAG As Boolean

Dim epsilon As Double

On Error GoTo ERROR_LABEL

epsilon = 0.0000000001
CONST_FLAG = False

PARAM_MATRIX = PARAM_RNG
CONST_BOX = CONST_RNG
If IsArray(CONST_BOX) = True Then: CONST_FLAG = True
    
NROWS = UBound(PARAM_MATRIX, 1)
NCOLUMNS = UBound(PARAM_MATRIX, 2) 'NO VARIABLES

If IsArray(SCALE_RNG) = True Then
    SCALE_VECTOR = SCALE_RNG
    If UBound(SCALE_VECTOR, 1) = 1 Then: SCALE_VECTOR = MATRIX_TRANSPOSE_FUNC(SCALE_VECTOR)
Else
    ReDim SCALE_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NCOLUMNS
        SCALE_VECTOR(i, 1) = 1
    Next i
End If

ReDim ATEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
ReDim BTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
ReDim ETEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
ReDim CTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
ReDim DTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)

ReDim ZTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
ReDim YTEMP_VECTOR(1 To NCOLUMNS + 1, 1 To 1)

For i = 1 To NCOLUMNS + 1 'Evaluamos la funcion en los vertices
'del simplex
    For j = 1 To NCOLUMNS
        ZTEMP_VECTOR(j, 1) = PARAM_MATRIX(i, j)
    Next j
    
    If CONST_FLAG = True Then
        For k = 1 To NCOLUMNS
            If CONST_BOX(k, 1) > ZTEMP_VECTOR(k, 1) Then ZTEMP_VECTOR(k, 1) = CONST_BOX(k, 1)
            If CONST_BOX(k, 2) < ZTEMP_VECTOR(k, 1) Then ZTEMP_VECTOR(k, 1) = CONST_BOX(k, 2)
        Next k
    End If
    
    TEMP_ROOT = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, ZTEMP_VECTOR, SCALE_VECTOR, MIN_FLAG)
    YTEMP_VECTOR(i, 1) = TEMP_ROOT
Next i

ReDim ORD_VECTOR(1 To NCOLUMNS, 1 To 1)
For j = 2 To NROWS
    EXIT_FLAG = False
    i = j - 1
    Do While EXIT_FLAG = False
        If YTEMP_VECTOR(i + 1, 1) < YTEMP_VECTOR(i, 1) Then
            TEMP_VALUE = YTEMP_VECTOR(i, 1)
            For k = 1 To NCOLUMNS
                ORD_VECTOR(k, 1) = PARAM_MATRIX(i, k)
            Next k
            YTEMP_VECTOR(i, 1) = YTEMP_VECTOR(i + 1, 1)
            For k = 1 To NCOLUMNS
                PARAM_MATRIX(i, k) = PARAM_MATRIX(i + 1, k)
            Next k
            YTEMP_VECTOR(i + 1, 1) = TEMP_VALUE
            For k = 1 To NCOLUMNS
                PARAM_MATRIX(i + 1, k) = ORD_VECTOR(k, 1)
            Next k
            i = i - 1
            If i = 0 Then: EXIT_FLAG = True
        Else
            EXIT_FLAG = True
        End If
    Loop
Next j

COUNTER = 0
TEMP_DELTA = 1
Do While TEMP_DELTA >= tolerance And COUNTER < nLOOPS
    For i = 1 To NCOLUMNS
        ATEMP_VECTOR(i, 1) = 0
        For j = 1 To NCOLUMNS
            ATEMP_VECTOR(i, 1) = ATEMP_VECTOR(i, 1) + PARAM_MATRIX(j, i)
        Next j
        ATEMP_VECTOR(i, 1) = ATEMP_VECTOR(i, 1) / NCOLUMNS
    Next i

    For i = 1 To NCOLUMNS
        BTEMP_VECTOR(i, 1) = 2 * ATEMP_VECTOR(i, 1) - PARAM_MATRIX(NCOLUMNS + 1, i)
    Next i
    If CONST_FLAG = True Then
        For i = 1 To NCOLUMNS
            If CONST_BOX(i, 1) > BTEMP_VECTOR(i, 1) Then BTEMP_VECTOR(i, 1) = CONST_BOX(i, 1)
            If CONST_BOX(i, 2) < BTEMP_VECTOR(i, 1) Then BTEMP_VECTOR(i, 1) = CONST_BOX(i, 2)
        Next i
    End If
    
    TEMP_ROOT = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, BTEMP_VECTOR, SCALE_VECTOR, MIN_FLAG)
    FIRST_ROOT = TEMP_ROOT

    If FIRST_ROOT < YTEMP_VECTOR(NCOLUMNS, 1) Then
        If YTEMP_VECTOR(1, 1) <= FIRST_ROOT Then
            For i = 1 To NCOLUMNS
                PARAM_MATRIX(NCOLUMNS + 1, i) = BTEMP_VECTOR(i, 1)
            Next i
            YTEMP_VECTOR(NCOLUMNS + 1, 1) = FIRST_ROOT
        Else
        
            For i = 1 To NCOLUMNS
                ETEMP_VECTOR(i, 1) = 3 * ATEMP_VECTOR(i, 1) - 2 * PARAM_MATRIX(NCOLUMNS + 1, i)
            Next i
            If CONST_FLAG = True Then
                For i = 1 To NCOLUMNS
                    If CONST_BOX(i, 1) > ETEMP_VECTOR(i, 1) Then ETEMP_VECTOR(i, 1) = CONST_BOX(i, 1)
                    If CONST_BOX(i, 2) < ETEMP_VECTOR(i, 1) Then ETEMP_VECTOR(i, 1) = CONST_BOX(i, 2)
                Next i
            End If
            
            TEMP_ROOT = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, ETEMP_VECTOR, SCALE_VECTOR, MIN_FLAG)
            SECOND_ROOT = TEMP_ROOT
                       
            If SECOND_ROOT < FIRST_ROOT Then
                For i = 1 To NCOLUMNS
                    PARAM_MATRIX(NCOLUMNS + 1, i) = ETEMP_VECTOR(i, 1)
                Next i
                YTEMP_VECTOR(NCOLUMNS + 1, 1) = SECOND_ROOT
            Else
                For i = 1 To NCOLUMNS
                    PARAM_MATRIX(NCOLUMNS + 1, i) = BTEMP_VECTOR(i, 1)
                Next i
                YTEMP_VECTOR(NCOLUMNS + 1, 1) = FIRST_ROOT
            End If
        End If
    Else
        If FIRST_ROOT < YTEMP_VECTOR(NCOLUMNS + 1, 1) Then
            For i = 1 To NCOLUMNS
                CTEMP_VECTOR(i, 1) = (3 / 2) * ATEMP_VECTOR(i, 1) - (1 / 2) * PARAM_MATRIX(NCOLUMNS + 1, i)
            Next i
            
            If CONST_FLAG = True Then
                For i = 1 To NCOLUMNS
                    If CONST_BOX(i, 1) > CTEMP_VECTOR(i, 1) Then CTEMP_VECTOR(i, 1) = CONST_BOX(i, 1)
                    If CONST_BOX(i, 2) < CTEMP_VECTOR(i, 1) Then CTEMP_VECTOR(i, 1) = CONST_BOX(i, 2)
                Next i
            End If
            
            TEMP_ROOT = _
            MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, _
                CTEMP_VECTOR, SCALE_VECTOR, MIN_FLAG)
            THIRD_ROOT = TEMP_ROOT
        
            If THIRD_ROOT <= FIRST_ROOT Then
                For i = 1 To NCOLUMNS
                    PARAM_MATRIX(NCOLUMNS + 1, i) = CTEMP_VECTOR(i, 1)
                Next i
                YTEMP_VECTOR(NCOLUMNS + 1, 1) = THIRD_ROOT
            Else
                For i = 2 To NCOLUMNS + 1
                    For k = 1 To NCOLUMNS
                        PARAM_MATRIX(i, k) = (1 / 2) * (PARAM_MATRIX(1, k) + _
                            PARAM_MATRIX(i, k))
                        ZTEMP_VECTOR(k, 1) = PARAM_MATRIX(i, k)
                    Next k
                    If CONST_FLAG = True Then
                        For j = 1 To NCOLUMNS
                            If CONST_BOX(j, 1) > ZTEMP_VECTOR(j, 1) Then ZTEMP_VECTOR(j, 1) = CONST_BOX(j, 1)
                            If CONST_BOX(j, 2) < ZTEMP_VECTOR(j, 1) Then ZTEMP_VECTOR(j, 1) = CONST_BOX(j, 2)
                        Next j
                    End If
                    TEMP_ROOT = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, ZTEMP_VECTOR, SCALE_VECTOR, MIN_FLAG)
                    YTEMP_VECTOR(i, 1) = TEMP_ROOT
                Next i
            End If
        Else
            For i = 1 To NCOLUMNS
                DTEMP_VECTOR(i, 1) = (1 / 2) * (ATEMP_VECTOR(i, 1) + PARAM_MATRIX(NCOLUMNS + 1, i))
            Next i
            If CONST_FLAG = True Then
                For i = 1 To NCOLUMNS
                    If CONST_BOX(i, 1) > DTEMP_VECTOR(i, 1) Then DTEMP_VECTOR(i, 1) = CONST_BOX(i, 1)
                    If CONST_BOX(i, 2) < DTEMP_VECTOR(i, 1) Then DTEMP_VECTOR(i, 1) = CONST_BOX(i, 2)
                Next i
            End If

            TEMP_ROOT = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, DTEMP_VECTOR, SCALE_VECTOR, MIN_FLAG)
            FORTH_ROOT = TEMP_ROOT
            
            If FORTH_ROOT < YTEMP_VECTOR(NCOLUMNS + 1, 1) Then
                For i = 1 To NCOLUMNS
                    PARAM_MATRIX(NCOLUMNS + 1, i) = DTEMP_VECTOR(i, 1)
                Next i
                YTEMP_VECTOR(NCOLUMNS + 1, 1) = FORTH_ROOT
            Else
                
                For i = 2 To NCOLUMNS + 1
                    For k = 1 To NCOLUMNS
                        PARAM_MATRIX(i, k) = (1 / 2) * (PARAM_MATRIX(1, k) + PARAM_MATRIX(i, k))
                        ZTEMP_VECTOR(k, 1) = PARAM_MATRIX(i, k)
                    Next k
                    If CONST_FLAG = True Then
                        For k = 1 To NCOLUMNS
                            If CONST_BOX(k, 1) > ZTEMP_VECTOR(k, 1) Then ZTEMP_VECTOR(k, 1) = CONST_BOX(k, 1)
                            If CONST_BOX(k, 2) < ZTEMP_VECTOR(k, 1) Then ZTEMP_VECTOR(k, 1) = CONST_BOX(k, 2)
                        Next k
                    End If
                    TEMP_ROOT = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, ZTEMP_VECTOR, SCALE_VECTOR, MIN_FLAG)
                    YTEMP_VECTOR(i, 1) = TEMP_ROOT
                Next i
            End If
        End If
    End If
    
    ReDim ORD_VECTOR(1 To NCOLUMNS, 1 To 1)
    For j = 2 To NROWS
        EXIT_FLAG = False
        i = j - 1
        Do While EXIT_FLAG = False
            If YTEMP_VECTOR(i + 1, 1) < YTEMP_VECTOR(i, 1) Then
                TEMP_VALUE = YTEMP_VECTOR(i, 1)
                For k = 1 To NCOLUMNS
                    ORD_VECTOR(k, 1) = PARAM_MATRIX(i, k)
                Next k
                YTEMP_VECTOR(i, 1) = YTEMP_VECTOR(i + 1, 1)
                For k = 1 To NCOLUMNS
                    PARAM_MATRIX(i, k) = PARAM_MATRIX(i + 1, k)
                Next k
                YTEMP_VECTOR(i + 1, 1) = TEMP_VALUE
                For k = 1 To NCOLUMNS
                    PARAM_MATRIX(i + 1, k) = ORD_VECTOR(k, 1)
                Next k
                i = i - 1
                If i = 0 Then: EXIT_FLAG = True
            Else
                EXIT_FLAG = True
            End If
        Loop
    Next j

    TEMP_DELTA = 2 * Abs(YTEMP_VECTOR(1, 1) - YTEMP_VECTOR(NCOLUMNS + 1, 1)) / (Abs(YTEMP_VECTOR(1, 1)) + Abs(YTEMP_VECTOR(NCOLUMNS + 1, 1)) + epsilon)
    COUNTER = COUNTER + 1
Loop

'If COUNTER >= nLOOPS Then: GoTo ERROR_LABEL 'iterations overflow
For j = 1 To NROWS
    For i = 1 To NCOLUMNS
        PARAM_MATRIX(j, i) = PARAM_MATRIX(j, i) * SCALE_VECTOR(i, 1)
    Next i
Next j

Select Case OUTPUT
    Case 0 'Best Parameters
        ReDim XTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
        For i = 1 To NCOLUMNS
            XTEMP_VECTOR(i, 1) = PARAM_MATRIX(NROWS, i)
        Next i
        NELDER_MEAD_OPTIMIZATION1_FUNC = XTEMP_VECTOR
    Case 1 'Parameters
        NELDER_MEAD_OPTIMIZATION1_FUNC = PARAM_MATRIX
    Case 2 'Function Value
        NELDER_MEAD_OPTIMIZATION1_FUNC = YTEMP_VECTOR
    Case Else
        NELDER_MEAD_OPTIMIZATION1_FUNC = COUNTER
End Select

Exit Function
ERROR_LABEL:
NELDER_MEAD_OPTIMIZATION1_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : NELDER_MEAD_OPTIMIZATION2_FUNC
'DESCRIPTION   : Nelder-Mead modified
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_NELDER
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function NELDER_MEAD_OPTIMIZATION2_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef DATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long

Dim F1_VAL As Double
Dim FC_VAL As Double
Dim FE_VAL As Double
Dim FN_VAL As Double
Dim FW_VAL As Double
Dim FR_VAL As Double
Dim FCC_VAL As Double

Dim NEW_VAL As Double
Dim NEW_VECTOR As Variant
Dim SHRINK_VAL As Double

Dim RHO_VAL As Double
Dim XI_VAL As Double
Dim GAMMA_VAL As Double
Dim SIGMA_VAL As Double

Dim TEMP_MATRIX As Variant
Dim X1_VECTOR As Variant
'Dim XN_VECTOR As Variant
Dim XW_VECTOR As Variant
Dim XBAR_VECTOR As Variant
Dim XR_VECTOR As Variant
Dim XE_VECTOR As Variant
Dim XC_VECTOR As Variant
Dim XCC_VECTOR As Variant
Dim YDATA_VECTOR As Variant
Dim TEMP_VECTOR As Variant

Dim PARAM_VECTOR As Variant

Dim nLOOPS As Long
Dim tolerance As Double

On Error GoTo ERROR_LABEL

nLOOPS = 5000
tolerance = 0.0000001
RHO_VAL = 1
XI_VAL = 2
GAMMA_VAL = 0.5
SIGMA_VAL = 0.5

PARAM_VECTOR = (PARAM_RNG)
If UBound(PARAM_VECTOR, 1) = 1 Then: _
    PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

NROWS = UBound(PARAM_VECTOR, 1)

ReDim TEMP_MATRIX(1 To NROWS + 1, 1 To NROWS + 1)
ReDim X1_VECTOR(1 To NROWS, 1 To 1)
'ReDim XN_VECTOR(1 To NROWS, 1 To 1)
ReDim XW_VECTOR(1 To NROWS, 1 To 1)
ReDim XBAR_VECTOR(1 To NROWS, 1 To 1)

ReDim XR_VECTOR(1 To NROWS, 1 To 1)
ReDim XE_VECTOR(1 To NROWS, 1 To 1)
ReDim XC_VECTOR(1 To NROWS, 1 To 1)
ReDim XCC_VECTOR(1 To NROWS, 1 To 1)
ReDim YDATA_VECTOR(1 To NROWS, 1 To 1)
ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    TEMP_MATRIX(1, i + 1) = PARAM_VECTOR(i, 1)
Next i
TEMP_MATRIX(1, 1) = Excel.Application.Run(FUNC_NAME_STR, DATA_RNG, PARAM_VECTOR)
For j = 1 To NROWS
    For i = 1 To NROWS
        If (i = j) Then
            If (PARAM_VECTOR(i, 1) = 0) Then
                TEMP_MATRIX(j + 1, i + 1) = 0.05
            Else
                TEMP_MATRIX(j + 1, i + 1) = PARAM_VECTOR(i, 1) * 1.05
            End If
        Else
            TEMP_MATRIX(j + 1, i + 1) = PARAM_VECTOR(i, 1)
        End If
        TEMP_VECTOR(i, 1) = TEMP_MATRIX(j + 1, i + 1)
    Next i
    TEMP_MATRIX(j + 1, 1) = Excel.Application.Run(FUNC_NAME_STR, DATA_RNG, TEMP_VECTOR)
Next j

For j = 1 To NROWS
    For i = 1 To NROWS
        If (i = j) Then
            TEMP_MATRIX(j + 1, i + 1) = PARAM_VECTOR(i, 1) * 1.05
        Else
            TEMP_MATRIX(j + 1, i + 1) = PARAM_VECTOR(i, 1)
        End If
        TEMP_VECTOR(i, 1) = TEMP_MATRIX(j + 1, i + 1)
    Next i
    TEMP_MATRIX(j + 1, 1) = Excel.Application.Run(FUNC_NAME_STR, DATA_RNG, TEMP_VECTOR)
Next j

For l = 1 To nLOOPS
    TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 1, 1)
    If (Abs(TEMP_MATRIX(1, 1) - TEMP_MATRIX(NROWS + 1, 1)) < tolerance) Then
        Exit For
    End If
    F1_VAL = TEMP_MATRIX(1, 1)
    For i = 1 To NROWS
        X1_VECTOR(i, 1) = TEMP_MATRIX(1, i + 1)
    Next i
    FN_VAL = TEMP_MATRIX(NROWS, 1)
    'For i = 1 To NROWS
     '   XN_VECTOR(i, 1) = TEMP_MATRIX(NROWS, i + 1)
    'Next i
    FW_VAL = TEMP_MATRIX(NROWS + 1, 1)
    For i = 1 To NROWS
        XW_VECTOR(i, 1) = TEMP_MATRIX(NROWS + 1, i + 1)
    Next i
    For i = 1 To NROWS
        XBAR_VECTOR(i, 1) = 0
        For j = 1 To NROWS
            XBAR_VECTOR(i, 1) = XBAR_VECTOR(i, 1) + TEMP_MATRIX(j, i + 1)
        Next j
        XBAR_VECTOR(i, 1) = XBAR_VECTOR(i, 1) / NROWS
    Next i
    For i = 1 To NROWS
        XR_VECTOR(i, 1) = XBAR_VECTOR(i, 1) + RHO_VAL * _
                         (XBAR_VECTOR(i, 1) - XW_VECTOR(i, 1))
    Next i
    FR_VAL = Excel.Application.Run(FUNC_NAME_STR, DATA_RNG, (XR_VECTOR))
    SHRINK_VAL = 0
    If ((FR_VAL >= F1_VAL) And (FR_VAL < FN_VAL)) Then
        NEW_VECTOR = XR_VECTOR
        NEW_VAL = FR_VAL
    ElseIf (FR_VAL < F1_VAL) Then
        For i = 1 To NROWS
            XE_VECTOR(i, 1) = XBAR_VECTOR(i, 1) + XI_VAL * _
                             (XR_VECTOR(i, 1) - XBAR_VECTOR(i, 1))
        Next i
        FE_VAL = Excel.Application.Run(FUNC_NAME_STR, DATA_RNG, (XE_VECTOR))
        If (FE_VAL < FR_VAL) Then
            NEW_VECTOR = XE_VECTOR
            NEW_VAL = FE_VAL
        Else
            NEW_VECTOR = XR_VECTOR
            NEW_VAL = FR_VAL
        End If
    ElseIf (FR_VAL >= FN_VAL) Then
        If ((FR_VAL >= FN_VAL) And (FR_VAL < FW_VAL)) Then
            For i = 1 To NROWS
                XC_VECTOR(i, 1) = XBAR_VECTOR(i, 1) + GAMMA_VAL * _
                                 (XR_VECTOR(i, 1) - XBAR_VECTOR(i, 1))
            Next i
            FC_VAL = Excel.Application.Run(FUNC_NAME_STR, DATA_RNG, (XC_VECTOR))
            If (FC_VAL <= FR_VAL) Then
                NEW_VECTOR = XC_VECTOR
                NEW_VAL = FC_VAL
            Else
                SHRINK_VAL = 1
            End If
        Else
            For i = 1 To NROWS
                XCC_VECTOR(i, 1) = XBAR_VECTOR(i, 1) - GAMMA_VAL * _
                                  (XBAR_VECTOR(i, 1) - XW_VECTOR(i, 1))
            Next i
            FCC_VAL = Excel.Application.Run(FUNC_NAME_STR, DATA_RNG, (XCC_VECTOR))
            If (FCC_VAL < FW_VAL) Then
                NEW_VECTOR = XCC_VECTOR
                NEW_VAL = FCC_VAL
            Else
                SHRINK_VAL = 1
            End If
        End If
    End If
    If (SHRINK_VAL = 1) Then
        For k = 2 To NROWS + 1
            For i = 1 To NROWS
                TEMP_MATRIX(k, i + 1) = X1_VECTOR(i, 1) + SIGMA_VAL * _
                                     (TEMP_MATRIX(k, i + 1) - X1_VECTOR(1, 1))
                TEMP_VECTOR(i, 1) = TEMP_MATRIX(k, i + 1)
            Next i
            TEMP_MATRIX(k, 1) = Excel.Application.Run(FUNC_NAME_STR, DATA_RNG, TEMP_VECTOR)
        Next k
    Else
        For i = 1 To NROWS
            TEMP_MATRIX(NROWS + 1, i + 1) = NEW_VECTOR(i, 1)
        Next i
        TEMP_MATRIX(NROWS + 1, 1) = NEW_VAL
    End If
Next l
If (l = nLOOPS + 1) Then: GoTo ERROR_LABEL
'    NELDER_MEAD_OPTIMIZATION2_FUNC = _
'        "Maximum Iteration (" & nLOOPS & ") exceeded"
'    Exit Function
'End If
TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP_MATRIX, 1, 1)
For i = 2 To NROWS + 1
    YDATA_VECTOR(i - 1, 1) = TEMP_MATRIX(1, i)
Next i
NELDER_MEAD_OPTIMIZATION2_FUNC = (YDATA_VECTOR)

Exit Function
ERROR_LABEL:
NELDER_MEAD_OPTIMIZATION2_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : NELDER_MEAD_OPTIMIZATION3_FUNC

'DESCRIPTION   : this routine shows the evolution of the simplex method
'while solving for minimum of a function.

'Nelder mead simplex method is an algorithm for finding minimum point
'for a function, and is based on a convex hull concept. The simplex
'is a structure of dimensions = dimensions of input parameters+1

'For example, if we were minimizing a function of 2 parameters, the
'resulting figure is a 3 point structure or a triangle.

'The initial simplex figure is set to an arbitrary value. After each
'iteration, function values are calculated at the vertex points of the
'triangles, and four operations are done in either combinations:

'-Reflection
'-Expansion
'-Contraction
'-Shrink

'The decision is made by reflecting the worst point and checking how the
'function value at new point compares to the previous points.

'Note that after each Iteration, the simplex structure
'moves towards the function minimum and the size also gets smaller. Finally
'the method coverges when simplex size gets below tolerance, and this returns
'the function minima.

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_NELDER
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function NELDER_MEAD_OPTIMIZATION3_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef PARAM_RNG As Variant, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 0.000000000000001)

'tolerance determines when to converge

Dim h As Long
Dim i As Long
Dim j As Long

Dim NSIZE As Long

Dim SIZE_VAL As Double
Dim DENOM_VAL As Double
Dim SCALING_VAL As Double

Dim TEMP_SUM As Double

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim PARAM_VECTOR As Variant
Dim PARAM_MATRIX As Variant

Dim CENTROID_VECTOR As Variant
Dim REFLECTED_VECTOR As Variant
Dim EXPANDED_VECTOR As Variant
Dim CONTRACTED_VECTOR As Variant
Dim PARAM_BEST_VECTOR As Variant
Dim PARAM_WORST_VECTOR As Variant

Dim ACCEPTED_VECTOR As Variant
Dim F_REFLECTED_VAL As Double
Dim F_2ND_WORST_VAL As Double
Dim F_BEST_VAL As Double
Dim F_WORST_VAL As Double
Dim F_CONTRACTED_VAL As Double
Dim F_EXPANDED_VAL As Double

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR) = 1 Then: _
  PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)

'returns initial matrix with simplex coordinates

'first column of this structure will have function values
'rest of columns will have coordinates
ReDim PARAM_MATRIX(1 To NSIZE + 1, 1 To NSIZE + 1)
'set first vector simply to initial params
PARAM_MATRIX(1, 1) = Excel.Application.Run(FUNC_NAME_STR, PARAM_VECTOR)
For i = 2 To NSIZE + 1
  PARAM_MATRIX(1, i) = PARAM_VECTOR(i - 1, 1)
Next i

'calc scaling factor by taking hightest value of input param
ReDim TEMP_VECTOR(1 To NSIZE, 1 To 1)
For i = 1 To NSIZE
  TEMP_VECTOR(i, 1) = Abs(PARAM_VECTOR(i, 1))
Next i
TEMP_VECTOR = MATRIX_SORT_COLUMNS_FUNC(TEMP_VECTOR, 1)
SCALING_VAL = TEMP_VECTOR(NSIZE, 1)
If SCALING_VAL < 1 Then: SCALING_VAL = 1

'set the remaining vectors to unit vectors
For i = 2 To NSIZE + 1 'loop over each row
  For j = 2 To NSIZE + 1 'loop over cells in a row
    PARAM_MATRIX(i, j) = PARAM_VECTOR(j - 1, 1)
  Next j
  PARAM_MATRIX(i, i) = PARAM_MATRIX(i, i) + SCALING_VAL
  ReDim TEMP_VECTOR(1 To NSIZE, 1 To 1)
  For j = 2 To NSIZE + 1
    TEMP_VECTOR(j - 1, 1) = PARAM_MATRIX(i, j)
  Next j
  PARAM_MATRIX(i, 1) = _
      Excel.Application.Run(FUNC_NAME_STR, TEMP_VECTOR)
Next i
PARAM_MATRIX = MATRIX_SORT_COLUMNS_FUNC(PARAM_MATRIX, 1)
    
ReDim TEMP_VECTOR(1 To NSIZE, 1 To 1)

For h = 1 To nLOOPS

  'check for convergence
  ReDim TEMP_MATRIX(1 To NSIZE, 1 To 1)
  For i = 2 To NSIZE + 1 'looping over points
    For j = 2 To NSIZE + 1 'looping over coordinates of a point
      TEMP_MATRIX(i - 1, 1) = _
          TEMP_MATRIX(i - 1, 1) + Abs(PARAM_MATRIX(i, j) - PARAM_MATRIX(1, j))
    Next j
  Next i
  TEMP_MATRIX = MATRIX_SORT_COLUMNS_FUNC(TEMP_MATRIX, 1)
  DENOM_VAL = 0
  For i = 1 To NSIZE
    DENOM_VAL = DENOM_VAL + Abs(PARAM_MATRIX(1, i + 1))
  Next i
  If DENOM_VAL < 1 Then: DENOM_VAL = 1
  SIZE_VAL = TEMP_MATRIX(NSIZE, 1) / DENOM_VAL
  
  If SIZE_VAL < tolerance Then
    For i = 1 To NSIZE
      TEMP_VECTOR(i, 1) = PARAM_MATRIX(1, i + 1)
    Next i
    NELDER_MEAD_OPTIMIZATION3_FUNC = TEMP_VECTOR
    Exit Function
  End If
  
  'best point of PARAM_MATRIX is the first row and worst is the last row
  'so lets reflect the worst point to go farthest away from it
  
  'calculate CENTROID_VECTOR of the point excluding the worst point
  ReDim CENTROID_VECTOR(1 To NSIZE, 1 To 1)
  For i = 2 To NSIZE + 1 'columns
    TEMP_SUM = 0
    For j = 1 To NSIZE 'rows
      TEMP_SUM = TEMP_SUM + PARAM_MATRIX(j, i)
    Next j
    CENTROID_VECTOR(i - 1, 1) = TEMP_SUM / NSIZE
  Next i
  
  ReDim REFLECTED_VECTOR(1 To NSIZE, 1 To 1)
  ReDim EXPANDED_VECTOR(1 To NSIZE, 1 To 1)
  ReDim CONTRACTED_VECTOR(1 To NSIZE, 1 To 1)
  ReDim PARAM_BEST_VECTOR(1 To NSIZE, 1 To 1)
  ReDim PARAM_WORST_VECTOR(1 To NSIZE, 1 To 1)
  
  For i = 1 To NSIZE
    REFLECTED_VECTOR(i, 1) = 2 * CENTROID_VECTOR(i, 1) - PARAM_MATRIX(NSIZE + 1, i + 1)
    PARAM_WORST_VECTOR(i, 1) = PARAM_MATRIX(NSIZE + 1, i + 1)
    PARAM_BEST_VECTOR(i, 1) = PARAM_MATRIX(1, i + 1)
  Next i
  
  ACCEPTED_VECTOR = REFLECTED_VECTOR
  F_REFLECTED_VAL = Excel.Application.Run(FUNC_NAME_STR, REFLECTED_VECTOR)
  F_2ND_WORST_VAL = PARAM_MATRIX(NSIZE, 1)
  F_BEST_VAL = PARAM_MATRIX(1, 1)
  F_WORST_VAL = PARAM_MATRIX(NSIZE + 1, 1)
  
  If F_REFLECTED_VAL < F_2ND_WORST_VAL Then
    'we are doing good in moving towards this direction
    'let us see if this new point outperforms our best point
    If F_REFLECTED_VAL < F_BEST_VAL Then
      'let us go more and expand in this direction
      For i = 1 To NSIZE
        EXPANDED_VECTOR(i, 1) = 2 * REFLECTED_VECTOR(i, 1) - CENTROID_VECTOR(i, 1)
      Next i
      F_EXPANDED_VAL = Excel.Application.Run(FUNC_NAME_STR, EXPANDED_VECTOR)
      If F_EXPANDED_VAL < F_BEST_VAL Then: ACCEPTED_VECTOR = EXPANDED_VECTOR
    End If
  Else
  
    If F_REFLECTED_VAL < F_WORST_VAL Then
      TEMP_VECTOR = REFLECTED_VECTOR
    Else
      TEMP_VECTOR = PARAM_WORST_VECTOR
    End If
    For i = 1 To NSIZE
      CONTRACTED_VECTOR(i, 1) = 0.5 * TEMP_VECTOR(i, 1) + 0.5 * CENTROID_VECTOR(i, 1)
    Next i
    F_CONTRACTED_VAL = _
      Excel.Application.Run(FUNC_NAME_STR, CONTRACTED_VECTOR)
    If F_CONTRACTED_VAL < F_2ND_WORST_VAL Then
      ACCEPTED_VECTOR = CONTRACTED_VECTOR
    Else
    
      'shrink all coordinates
      For i = 2 To NSIZE
        For j = 2 To NSIZE + 1
          PARAM_MATRIX(i, j) = (PARAM_BEST_VECTOR(j - 1, 1) + PARAM_MATRIX(i, j)) / 2
          TEMP_VECTOR(j - 1, 1) = PARAM_MATRIX(i, j)
        Next j
        PARAM_MATRIX(i, 1) = Excel.Application.Run(FUNC_NAME_STR, TEMP_VECTOR)
      Next i
      For i = 1 To NSIZE
        TEMP_VECTOR(i, 1) = (PARAM_MATRIX(1, i + 1) + PARAM_MATRIX(NSIZE + 1, i + 1)) / 2
      Next i
      ACCEPTED_VECTOR = TEMP_VECTOR
    End If
   
  End If
  
  'Replace worst parameters with new choice
  For i = 1 To NSIZE
    PARAM_MATRIX(NSIZE + 1, i + 1) = ACCEPTED_VECTOR(i, 1)
  Next i
  PARAM_MATRIX(NSIZE + 1, 1) = Excel.Application.Run(FUNC_NAME_STR, ACCEPTED_VECTOR)
  
  PARAM_MATRIX = MATRIX_SORT_COLUMNS_FUNC(PARAM_MATRIX, 1)
Next h

  GoTo ERROR_LABEL 'iterations did not converge
Exit Function
ERROR_LABEL:
NELDER_MEAD_OPTIMIZATION3_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : NELDER_MEAD_OPTIMIZATION4_FUNC
'DESCRIPTION   : Nelder-Mead modified to handle complex function parameter
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_NELDER
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function NELDER_MEAD_OPTIMIZATION4_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef XDATA_RNG As Variant, _
ByRef YDATA_RNG As Variant, _
ByRef CONST_RNG As Variant, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal nLOOPS As Long = 5000, _
Optional ByVal tolerance As Double = 0.0000000001)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long

Dim NSIZE As Long

Dim RHO_VAL As Double
Dim XI_VAL As Double
Dim GAMMA_VAL As Double
Dim SIGMA_VAL As Double

Dim FC_VAL As Double
Dim FD_VAL As Double
Dim F1_VAL As Double
Dim FN_VAL As Double
Dim FW_VAL As Double
Dim FR_VAL As Double
Dim FE_VAL As Double

Dim SHRINK_VAL As Double
Dim NEW_FUNC_VAL As Double

Dim X1_VECTOR As Variant
Dim XB_VECTOR As Variant
Dim XC_VECTOR As Variant
Dim XD_VECTOR As Variant
Dim XE_VECTOR As Variant
Dim XF_VECTOR As Variant
Dim XN_VECTOR As Variant
Dim XR_VECTOR As Variant
Dim XW_VECTOR As Variant

Dim XP_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim INITIAL_VECTOR As Variant

Dim YDATA_VECTOR As Variant
Dim XDATA_MATRIX As Variant
Dim CONST_VECTOR As Variant
Dim PARAM_VECTOR As Variant

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

RHO_VAL = 1
XI_VAL = 2
GAMMA_VAL = 0.5
SIGMA_VAL = 0.5

XDATA_MATRIX = XDATA_RNG
YDATA_VECTOR = YDATA_RNG
If UBound(YDATA_VECTOR, 1) = 1 Then
    YDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(YDATA_VECTOR)
End If
If UBound(XDATA_MATRIX, 1) <> UBound(YDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

'----------------------------------------------------------------
If IsArray(CONST_RNG) Then
'----------------------------------------------------------------
    CONST_VECTOR = CONST_RNG
    If UBound(CONST_VECTOR, 1) = 1 Then
        CONST_VECTOR = MATRIX_TRANSPOSE_FUNC(CONST_VECTOR)
    End If
'----------------------------------------------------------------
Else
'----------------------------------------------------------------
    ReDim CONST_VECTOR(1 To 1, 1 To 1)
    CONST_VECTOR(1, 1) = CONST_RNG
'----------------------------------------------------------------
End If
'----------------------------------------------------------------
If IsArray(PARAM_RNG) Then
'----------------------------------------------------------------
    PARAM_VECTOR = PARAM_RNG
    If UBound(PARAM_VECTOR, 1) = 1 Then
        PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
    End If
'----------------------------------------------------------------
Else
'----------------------------------------------------------------
    ReDim PARAM_VECTOR(1 To 1, 1 To 1)
    PARAM_VECTOR(1, 1) = PARAM_RNG
'----------------------------------------------------------------
End If
'----------------------------------------------------------------
NSIZE = UBound(PARAM_VECTOR, 1)
'----------------------------------------------------------------

ReDim INITIAL_VECTOR(1 To NSIZE, 1 To 1)
For i = 1 To NSIZE
    INITIAL_VECTOR(i, 1) = PARAM_VECTOR(i, 1)
Next i

ReDim X1_VECTOR(1 To NSIZE, 1 To 1)
ReDim XN_VECTOR(1 To NSIZE, 1 To 1)
ReDim XW_VECTOR(1 To NSIZE, 1 To 1)
ReDim XB_VECTOR(1 To NSIZE, 1 To 1)
ReDim XR_VECTOR(1 To NSIZE, 1 To 1)
ReDim XE_VECTOR(1 To NSIZE, 1 To 1)
ReDim XC_VECTOR(1 To NSIZE, 1 To 1)
ReDim XD_VECTOR(1 To NSIZE, 1 To 1)
ReDim XF_VECTOR(1 To NSIZE, 1 To 1)
ReDim DATA_MATRIX(1 To NSIZE + 1, 1 To NSIZE + 1)

'---------------------------------------------------------------------------------------------------------------
For i = 1 To NSIZE: DATA_MATRIX(1, i + 1) = INITIAL_VECTOR(i, 1): Next i
'---------------------------------------------------------------------------------------------------------------
DATA_MATRIX(1, 1) = Excel.Application.Run(FUNC_NAME_STR, XDATA_MATRIX, YDATA_VECTOR, CONST_VECTOR, INITIAL_VECTOR)
'---------------------------------------------------------------------------------------------------------------
For j = 1 To NSIZE
    For i = 1 To NSIZE
        If (i = j) Then
            If (INITIAL_VECTOR(i, 1) = 0) Then
                DATA_MATRIX(j + 1, i + 1) = 0.05
            Else
                DATA_MATRIX(j + 1, i + 1) = INITIAL_VECTOR(i, 1) * 1.05
            End If
        Else
            DATA_MATRIX(j + 1, i + 1) = INITIAL_VECTOR(i, 1)
        End If
        XF_VECTOR(i, 1) = DATA_MATRIX(j + 1, i + 1)
    Next i
    DATA_MATRIX(j + 1, 1) = Excel.Application.Run(FUNC_NAME_STR, XDATA_MATRIX, YDATA_VECTOR, CONST_VECTOR, XF_VECTOR)
Next j
'---------------------------------------------------------------------------------------------------------------
For j = 1 To NSIZE
    For i = 1 To NSIZE
        If (i = j) Then
            DATA_MATRIX(j + 1, i + 1) = INITIAL_VECTOR(i, 1) * 1.05
        Else
            DATA_MATRIX(j + 1, i + 1) = INITIAL_VECTOR(i, 1)
        End If
        XF_VECTOR(i, 1) = DATA_MATRIX(j + 1, i + 1)
    Next i
    DATA_MATRIX(j + 1, 1) = Excel.Application.Run(FUNC_NAME_STR, XDATA_MATRIX, YDATA_VECTOR, CONST_VECTOR, XF_VECTOR)
Next j
'---------------------------------------------------------------------------------------------------------------
For k = 1 To nLOOPS
'---------------------------------------------------------------------------------------------------------------
    GoSub SORT_LINE
    If (Abs(DATA_MATRIX(1, 1) - DATA_MATRIX(NSIZE + 1, 1)) < tolerance) Then
        Exit For
    End If
'---------------------------------------------------------------------------------------------------------------
    F1_VAL = DATA_MATRIX(1, 1)
    For i = 1 To NSIZE
        X1_VECTOR(i, 1) = DATA_MATRIX(1, i + 1)
    Next i
'---------------------------------------------------------------------------------------------------------------
    FN_VAL = DATA_MATRIX(NSIZE, 1)
    For i = 1 To NSIZE
        XN_VECTOR(i, 1) = DATA_MATRIX(NSIZE, i + 1)
    Next i
'---------------------------------------------------------------------------------------------------------------
    FW_VAL = DATA_MATRIX(NSIZE + 1, 1)
    For i = 1 To NSIZE
        XW_VECTOR(i, 1) = DATA_MATRIX(NSIZE + 1, i + 1)
    Next i
'---------------------------------------------------------------------------------------------------------------
    For i = 1 To NSIZE
        XB_VECTOR(i, 1) = 0
        For j = 1 To NSIZE
            XB_VECTOR(i, 1) = XB_VECTOR(i, 1) + DATA_MATRIX(j, i + 1)
        Next j
        XB_VECTOR(i, 1) = XB_VECTOR(i, 1) / NSIZE
    Next i
'---------------------------------------------------------------------------------------------------------------
    For i = 1 To NSIZE
        XR_VECTOR(i, 1) = XB_VECTOR(i, 1) + RHO_VAL * (XB_VECTOR(i, 1) - XW_VECTOR(i, 1))
    Next i
    FR_VAL = Excel.Application.Run(FUNC_NAME_STR, XDATA_MATRIX, YDATA_VECTOR, CONST_VECTOR, XR_VECTOR)
'---------------------------------------------------------------------------------------------------------------
    SHRINK_VAL = 0
    If ((FR_VAL >= F1_VAL) And (FR_VAL < FN_VAL)) Then
'---------------------------------------------------------------------------------------------------------------
        XP_VECTOR = XR_VECTOR
        NEW_FUNC_VAL = FR_VAL
'---------------------------------------------------------------------------------------------------------------
    ElseIf (FR_VAL < F1_VAL) Then
'---------------------------------------------------------------------------------------------------------------
        For i = 1 To NSIZE
            XE_VECTOR(i, 1) = XB_VECTOR(i, 1) + XI_VAL * (XR_VECTOR(i, 1) - XB_VECTOR(i, 1))
        Next i
        FE_VAL = Excel.Application.Run(FUNC_NAME_STR, XDATA_MATRIX, YDATA_VECTOR, CONST_VECTOR, XE_VECTOR)
        If (FE_VAL < FR_VAL) Then
            XP_VECTOR = XE_VECTOR
            NEW_FUNC_VAL = FE_VAL
        Else
            XP_VECTOR = XR_VECTOR
            NEW_FUNC_VAL = FR_VAL
        End If
'---------------------------------------------------------------------------------------------------------------
    ElseIf (FR_VAL >= FN_VAL) Then
'---------------------------------------------------------------------------------------------------------------
        If ((FR_VAL >= FN_VAL) And (FR_VAL < FW_VAL)) Then
            For i = 1 To NSIZE
                XC_VECTOR(i, 1) = XB_VECTOR(i, 1) + GAMMA_VAL * (XR_VECTOR(i, 1) - XB_VECTOR(i, 1))
            Next i
            FC_VAL = Excel.Application.Run(FUNC_NAME_STR, XDATA_MATRIX, YDATA_VECTOR, CONST_VECTOR, XC_VECTOR)
            If (FC_VAL <= FR_VAL) Then
                XP_VECTOR = XC_VECTOR
                NEW_FUNC_VAL = FC_VAL
            Else
                SHRINK_VAL = 1
            End If
'---------------------------------------------------------------------------------------------------------------
        Else
'---------------------------------------------------------------------------------------------------------------
            For i = 1 To NSIZE
                XD_VECTOR(i, 1) = XB_VECTOR(i, 1) - GAMMA_VAL * (XB_VECTOR(i, 1) - XW_VECTOR(i, 1))
            Next i
            FD_VAL = Excel.Application.Run(FUNC_NAME_STR, XDATA_MATRIX, YDATA_VECTOR, CONST_VECTOR, XD_VECTOR)
            If (FD_VAL < FW_VAL) Then
                XP_VECTOR = XD_VECTOR
                NEW_FUNC_VAL = FD_VAL
            Else
                SHRINK_VAL = 1
            End If
        End If
'---------------------------------------------------------------------------------------------------------------
    End If
'---------------------------------------------------------------------------------------------------------------
    If (SHRINK_VAL = 1) Then
'---------------------------------------------------------------------------------------------------------------
        For j = 2 To NSIZE + 1
            For i = 1 To NSIZE
                DATA_MATRIX(j, i + 1) = X1_VECTOR(i, 1) + SIGMA_VAL * (DATA_MATRIX(j, i + 1) - X1_VECTOR(1, 1))
                XF_VECTOR(i, 1) = DATA_MATRIX(j, i + 1)
            Next i
            DATA_MATRIX(j, 1) = Excel.Application.Run(FUNC_NAME_STR, XDATA_MATRIX, YDATA_VECTOR, CONST_VECTOR, XF_VECTOR)
        Next j
'---------------------------------------------------------------------------------------------------------------
    Else
'---------------------------------------------------------------------------------------------------------------
        For i = 1 To NSIZE
            DATA_MATRIX(NSIZE + 1, i + 1) = XP_VECTOR(i, 1)
        Next i
        DATA_MATRIX(NSIZE + 1, 1) = NEW_FUNC_VAL
'---------------------------------------------------------------------------------------------------------------
    End If
'---------------------------------------------------------------------------------------------------------------
Next k
'---------------------------------------------------------------------------------------------------------------
'If (k = nLOOPS + 1) Then
 '   GoTo ERROR_LABEL 'Maximum Iteration (" & nLOOPS & ") exceeded
'End If
'---------------------------------------------------------------------------------------------------------------
GoSub SORT_LINE
'---------------------------------------------------------------------------------------------------------------
For i = 1 To NSIZE: PARAM_VECTOR(i, 1) = DATA_MATRIX(1, i + 1): Next i
'---------------------------------------------------------------------------------------------------------------
NELDER_MEAD_OPTIMIZATION4_FUNC = PARAM_VECTOR

Exit Function
'---------------------------------------------------------------------------------------------------------------
SORT_LINE:
'---------------------------------------------------------------------------------------------------------------
    ii = UBound(DATA_MATRIX, 1)
    jj = UBound(DATA_MATRIX, 2)
    ReDim TEMP_VECTOR(1 To jj, 1 To 1)
    ReDim TEMP_MATRIX(1 To ii, 1 To jj)
    For i = ii - 1 To 1 Step -1
        For j = 1 To i
            If (DATA_MATRIX(j, 1) > DATA_MATRIX(j + 1, 1)) Then
                For l = 1 To jj
                    TEMP_VECTOR(l, 1) = DATA_MATRIX(j + 1, l)
                    DATA_MATRIX(j + 1, l) = DATA_MATRIX(j, l)
                    DATA_MATRIX(j, l) = TEMP_VECTOR(l, 1)
                Next l
            End If
        Next j
    Next i
    Erase TEMP_VECTOR
    Erase TEMP_MATRIX
'---------------------------------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
NELDER_MEAD_OPTIMIZATION4_FUNC = Err.number
End Function

