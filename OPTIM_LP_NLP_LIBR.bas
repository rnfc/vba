Attribute VB_Name = "OPTIM_LP_NLP_LIBR"

'-------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'Non-Linear Programming with constrained optimization is much difficult than
'unconstrained optimization: we have to find the best point of the function
'respecting all the constraints that may be equalities or inequalities. The
'solution (the optimum point), in fact, may not occur at the top of a peak
'or at the bottom of a valley.

'The main elements of any constrained optimization problem are: the objective
'function, the variables, the constraints and sometime the variable bounds.
'When the objective function is not linear (example a quadratic function) and the
'constraints are linear we have a so called NLP with linear constraints.
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : NLP_OPTIMIZATION_FUNC

'DESCRIPTION   : This function solves a non-linear programming problem having
'linear constraints. It uses the CG+MC algorithm. This algorithm works fine
'with quadratic functions but it can also work with other non linear smooth
'functions. It needs information about
'1. The range containing the the function to optimize (objective function)
'2. The range containing the variables to be changed (max 9
'variables)
'3. The range containing the variable bounds (minimum and maximum
'limits)
'4. The range containing the linear constraints coefficients .
'The constraint accepted are "<", ">", "<=", ">="
'Note that, for this macro, the symbols "<" and "<=" or ">" and ">=" are equivalent

'LIBRARY       : OPTIMIZATION
'GROUP         : LP_NLP
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function NLP_OPTIMIZATION_FUNC(ByRef PARAM_RNG As Variant, _
ByRef CONST_RNG As Variant, _
ByRef COEF_RNG As Variant, _
ByVal FUNC_NAME_STR As String, _
Optional ByVal GRAD_STR_NAME As String = "", _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByVal RND_FLAG As Boolean = True, _
Optional ByVal METHOD As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 10 ^ -12)

'CONST_RNG : Constraint box for parameters
'COEF_RNG  : Coefficients of the constrain equations

'IF METHOD = 0 Then: Gradient Else: Conjugate Gradient

'RND_FLAG = 1: activates/deactivates the random starting algorithm. If selected,
'the starting point is chosen randomly inside the given constraints box.
'Otherwise the algorithm starts with the initial variables value
'This feature may be useful when we already know a sufficiently close solution
'or when there are many local optima points.

Dim i As Long
Dim j As Long

Dim NSIZE As Long
Dim NO_VAR As Long

Dim TEMP_SUM As Double
Dim TEMP_DELTA As Double
Dim FUNC_VALUE As Double

Dim SYMB_STR As String

Dim CONST_BOX As Variant
Dim COEF_MATRIX As Variant

Dim XTEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant
Dim SCALE_VECTOR As Variant

Dim TEMP_ARR As Variant
Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

Dim MIN_FUNC_VAL As Double

Dim CONVERG_VAL As Integer

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)
'-----------------------------------------------------------------------------------
    CONST_BOX = CONST_RNG
    CONST_BOX = MULTVAR_LOAD_CONST_FUNC(CONST_BOX, 1)
'-----------------------------------------------------------------------------------
    TEMP_ARR = MULTVAR_SCALE_CONST_FUNC(CONST_BOX)
    CONST_BOX = TEMP_ARR(LBound(TEMP_ARR))
    SCALE_VECTOR = TEMP_ARR(UBound(TEMP_ARR))
'-----------------------------------------------------------------------------------

COEF_MATRIX = COEF_RNG

ReDim XTEMP_VECTOR(1 To NSIZE, 1 To 1)

For i = 1 To NSIZE
     XTEMP_VECTOR(i, 1) = (1 / SCALE_VECTOR(i, 1)) * PARAM_VECTOR(i, 1)
Next i

'check if the starting point belongs to the box
For i = 1 To NSIZE 'fix the constrain coordinate XTEMP_VECTOR
    If CONST_BOX(i, 1) > XTEMP_VECTOR(i, 1) Then XTEMP_VECTOR(i, 1) = CONST_BOX(i, 1)
    If CONST_BOX(i, 2) < XTEMP_VECTOR(i, 1) Then XTEMP_VECTOR(i, 1) = CONST_BOX(i, 2)
Next i

ReDim PARAM_VECTOR(1 To NSIZE, 1 To 1)
For i = 1 To NSIZE
     PARAM_VECTOR(i, 1) = XTEMP_VECTOR(i, 1)
Next i

ReDim ATEMP_MATRIX(1 To (2 * NSIZE + UBound(COEF_MATRIX)), 1 To NSIZE + 2)
For j = 1 To NSIZE 'first, add the box constraints
    i = 2 * j
    ATEMP_MATRIX(i - 1, j) = -1
    ATEMP_MATRIX(i - 1, NSIZE + 1) = "<="
    ATEMP_MATRIX(i - 1, NSIZE + 2) = -CONST_BOX(j, 1)
    ATEMP_MATRIX(i, j) = 1
    ATEMP_MATRIX(i, NSIZE + 1) = "<="
    ATEMP_MATRIX(i, NSIZE + 2) = CONST_BOX(j, 2)
Next j

For i = 1 To UBound(COEF_MATRIX) 'last, add the linear constraints
    For j = 1 To NSIZE
        ATEMP_MATRIX(i + 2 * NSIZE, j) = SCALE_VECTOR(j, 1) * COEF_MATRIX(i, j)
    Next j
    ATEMP_MATRIX(i + 2 * NSIZE, NSIZE + 1) = COEF_MATRIX(i, NSIZE + 1)
    ATEMP_MATRIX(i + 2 * NSIZE, NSIZE + 2) = COEF_MATRIX(i, NSIZE + 2)
Next i

NO_VAR = (UBound(ATEMP_MATRIX, 2) - 2)

ReDim BTEMP_MATRIX(1 To UBound(ATEMP_MATRIX, 1), 1 To NO_VAR + 1)
'constraints coefficients matrix

For i = 1 To UBound(ATEMP_MATRIX, 1) 'converts and returns the matrix
'BTEMP_MATRIX from a range of linear constraints ai1*x1+ai2*x2+...aim*xm <= bi
'for i=1...; all the constraints are converted in "<="

    SYMB_STR = ATEMP_MATRIX(i, NO_VAR + 1)
    
    If InStr(1, SYMB_STR, ">") > 0 Or Asc(SYMB_STR) = 179 Then
        SYMB_STR = ">"
    ElseIf InStr(1, SYMB_STR, "<") > 0 Or Asc(SYMB_STR) = 163 Then
        SYMB_STR = "<"
    ElseIf InStr(1, SYMB_STR, "=") Then
        SYMB_STR = "="
    Else
        GoTo ERROR_LABEL 'constraints syntax error
    End If
    
    For j = 1 To NO_VAR
        BTEMP_MATRIX(i, j) = ATEMP_MATRIX(i, j)
    Next j
    BTEMP_MATRIX(i, NO_VAR + 1) = _
        ATEMP_MATRIX(i, NO_VAR + 2)   'right constant term
    If SYMB_STR = ">" Then
        For j = 1 To UBound(BTEMP_MATRIX, 2)
            BTEMP_MATRIX(i, j) = -BTEMP_MATRIX(i, j)
        Next j
    End If
Next i

'-----------------------------------------------------------------------------
If RND_FLAG = True Then
'-----------------------------------------------------------------------------
    TEMP_DELTA = nLOOPS / 5
    If TEMP_DELTA < 50 Then TEMP_DELTA = 50
    FUNC_VALUE = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, XTEMP_VECTOR, SCALE_VECTOR, MIN_FLAG)
    PARAM_VECTOR = NLP_MC_RND_OPTIM_FUNC(CONST_BOX, BTEMP_MATRIX, FUNC_NAME_STR, SCALE_VECTOR, MIN_FLAG, TEMP_DELTA, 0)
'-----------------------------------------------------------------------------
    If IsArray(PARAM_VECTOR) = False Then: GoTo 1983 'no starting point found"
'-----------------------------------------------------------------------------

    MIN_FUNC_VAL = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, PARAM_VECTOR, SCALE_VECTOR, MIN_FLAG)
    If (MIN_FLAG And MIN_FUNC_VAL < FUNC_VALUE) Or (Not MIN_FLAG And MIN_FUNC_VAL > FUNC_VALUE) Then
        'the new starting point is better
        For i = 1 To NSIZE
            XTEMP_VECTOR(i, 1) = PARAM_VECTOR(i, 1)
        Next i
    Else
        CONVERG_VAL = 0 'check if the point XTEMP_VECTOR belongs
        'to the allowed region
        For i = 1 To UBound(BTEMP_MATRIX)
            'compute the contraint i-th at the point XTEMP_VECTOR
            TEMP_SUM = 0
            For j = 1 To UBound(XTEMP_VECTOR)
                TEMP_SUM = TEMP_SUM + BTEMP_MATRIX(i, j) * XTEMP_VECTOR(j, 1)
            Next j
                TEMP_SUM = TEMP_SUM - BTEMP_MATRIX(i, UBound(XTEMP_VECTOR) + 1)
            If TEMP_SUM > tolerance Then 'constraint i-th is violated
                CONVERG_VAL = i
                Exit For
            End If
        Next i
        
        If CONVERG_VAL <> 0 Then
            For i = 1 To NSIZE
                XTEMP_VECTOR(i, 1) = PARAM_VECTOR(i, 1)
            Next i
        End If
    End If
End If

1983:
XTEMP_VECTOR = NLP_CG_OPTIM_FUNC(XTEMP_VECTOR, BTEMP_MATRIX, _
               FUNC_NAME_STR, GRAD_STR_NAME, SCALE_VECTOR, _
               MIN_FLAG, nLOOPS, METHOD, 0)

'---------------------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------------------
    Case 0 'Min/Max Parameters
'---------------------------------------------------------------------------------------
        NLP_OPTIMIZATION_FUNC = XTEMP_VECTOR
'---------------------------------------------------------------------------------------
    Case 1
'---------------------------------------------------------------------------------------
        COEF_MATRIX = COEF_RNG
        COEF_MATRIX = MATRIX_TRANSPOSE_FUNC(MATRIX_REMOVE_COLUMNS_FUNC(COEF_MATRIX, UBound(COEF_MATRIX, 2) - 1, 2))
        ReDim PARAM_VECTOR(1 To UBound(COEF_MATRIX, 2), 1 To 1)
        For i = 1 To UBound(PARAM_VECTOR, 1)
            PARAM_VECTOR(i, 1) = MATRIX_SUM_PRODUCT_FUNC(MATRIX_GET_COLUMN_FUNC(COEF_MATRIX, i, 1), XTEMP_VECTOR)
        Next i
        NLP_OPTIMIZATION_FUNC = PARAM_VECTOR
'---------------------------------------------------------------------------------------
    Case Else
'---------------------------------------------------------------------------------------
        NLP_OPTIMIZATION_FUNC = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, XTEMP_VECTOR, "", MIN_FLAG)
'---------------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
NLP_OPTIMIZATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NLP_MC_RND_OPTIM_FUNC
'DESCRIPTION   : Approaches for the approximate global minimum with the Montecarlo
'method using linear constrains
'LIBRARY       : OPTIMIZATION
'GROUP         : LP_NLP
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function NLP_MC_RND_OPTIM_FUNC(ByRef CONST_RNG As Variant, _
ByRef COEF_RNG As Variant, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef SCALE_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByVal nLOOPS As Variant = 1000, _
Optional ByVal OUTPUT As Integer = 0)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long

Dim TEMP_MIN As Double
Dim TEMP_SUM As Double

Dim MIN_FUNC_VAL As Double

Dim CONST_BOX As Variant

Dim TEMP_MATRIX As Variant
Dim COEF_MATRIX As Variant
Dim PARAM_VECTOR As Variant
Dim SCALE_VECTOR As Variant

Dim DTEMP_VECTOR As Variant

Dim CONVERG_VAL As Long

Dim tolerance As Double

On Error GoTo ERROR_LABEL

tolerance = 10 ^ -12

CONST_BOX = CONST_RNG
COEF_MATRIX = COEF_RNG

NSIZE = UBound(CONST_BOX, 1)
If IsArray(SCALE_RNG) = True Then
    SCALE_VECTOR = SCALE_RNG
    If UBound(SCALE_VECTOR, 1) = 1 Then: SCALE_VECTOR = MATRIX_TRANSPOSE_FUNC(SCALE_VECTOR)
Else
    ReDim SCALE_VECTOR(1 To NSIZE, 1 To 1)
    For i = 1 To NSIZE
        SCALE_VECTOR(i, 1) = 1
    Next i
End If

ReDim DTEMP_VECTOR(1 To NSIZE, 1 To 1)
ReDim TEMP_MATRIX(1 To nLOOPS, 1 To NSIZE + 1)

ReDim PARAM_VECTOR(1 To UBound(CONST_BOX), 1 To 1)
For k = 1 To UBound(CONST_BOX)
    TEMP_MATRIX(1, k + 1) = (CONST_BOX(k, 2) - CONST_BOX(k, 1)) * Rnd + CONST_BOX(k, 1)
    PARAM_VECTOR(k, 1) = TEMP_MATRIX(1, k + 1)
Next k
TEMP_MATRIX(1, 1) = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, PARAM_VECTOR, SCALE_VECTOR, MIN_FLAG)

For h = 2 To UBound(TEMP_MATRIX)
    For k = 1 To UBound(CONST_BOX)
        TEMP_MATRIX(h, k + 1) = TEMP_MATRIX(h - 1, k + 1)
    Next k
    k = (h Mod UBound(CONST_BOX)) + 1
    TEMP_MATRIX(h, k + 1) = (CONST_BOX(k, 2) - CONST_BOX(k, 1)) * Rnd + CONST_BOX(k, 1)
    PARAM_VECTOR(k, 1) = TEMP_MATRIX(h, k + 1)
    TEMP_MATRIX(h, 1) = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, PARAM_VECTOR, SCALE_VECTOR, MIN_FLAG)
Next h

'    If Not MIN_FLAG Then
'       For h = 1 To UBound(TEMP_MATRIX)
'          TEMP_MATRIX(h, 1) = -TEMP_MATRIX(h, 1)
'     Next h
'End If

'search for the minimum
For h = 2 To UBound(TEMP_MATRIX)
    For k = 1 To NSIZE
        DTEMP_VECTOR(k, 1) = TEMP_MATRIX(h, k + 1)
    Next k
    CONVERG_VAL = 0 'check if the point DTEMP_VECTOR belongs
    'to the allowed region
    For i = 1 To UBound(COEF_MATRIX)
        'compute the contraint i-th at the point DTEMP_VECTOR
        TEMP_SUM = 0
        For j = 1 To UBound(DTEMP_VECTOR)
            TEMP_SUM = TEMP_SUM + COEF_MATRIX(i, j) * DTEMP_VECTOR(j, 1)
        Next j
            TEMP_SUM = TEMP_SUM - COEF_MATRIX(i, UBound(DTEMP_VECTOR) + 1)
        If TEMP_SUM > tolerance Then
            'constraint i-th is violated
            CONVERG_VAL = i
            Exit For
        End If
    Next i
    If CONVERG_VAL = 0 Then
        If MIN_FUNC_VAL > TEMP_MATRIX(h, 1) Or TEMP_MIN = 0 Then
            MIN_FUNC_VAL = TEMP_MATRIX(h, 1)
            TEMP_MIN = h
        End If
    End If
Next h

If TEMP_MIN > 0 Then
    Select Case OUTPUT
        Case 0
            ReDim PARAM_VECTOR(1 To NSIZE, 1 To 1)
            For k = 1 To NSIZE
                PARAM_VECTOR(k, 1) = TEMP_MATRIX(TEMP_MIN, k + 1)
            Next k
            NLP_MC_RND_OPTIM_FUNC = PARAM_VECTOR
        Case Else
            NLP_MC_RND_OPTIM_FUNC = MIN_FUNC_VAL
    End Select
Else
    NLP_MC_RND_OPTIM_FUNC = -1 'No Point Found
End If

Exit Function
ERROR_LABEL:
NLP_MC_RND_OPTIM_FUNC = -1
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NLP_CG_OPTIM_FUNC
'DESCRIPTION   : NLP CG Algorithm
'LIBRARY       : OPTIMIZATION
'GROUP         : LP_NLP
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function NLP_CG_OPTIM_FUNC(ByRef PARAM_RNG As Variant, _
ByRef CONST_RNG As Variant, _
ByVal FUNC_NAME_STR As String, _
Optional ByVal GRAD_STR_NAME As String = "", _
Optional ByRef SCALE_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal METHOD As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim COUNTER As Long

Dim NTRIALS As Long
Dim NO_POINTS As Long

Dim TEMP_VAL As Double
Dim TEMP_MAX As Double
Dim TEMP_MIN As Double

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double
Dim CTEMP_SUM As Double
Dim ETEMP_SUM As Double

Dim XTEMP_ERR As Double
Dim XTEMP_ABS As Double
Dim XINIT_ERR As Double

Dim TEMP_NORM As Double
Dim TEMP_MULT As Double
Dim TEMP_DELTA As Double

Dim FIRST_MULT As Double
Dim SECOND_MULT As Double
Dim THIRD_MULT As Double

Dim MIN_FUNC_VAL As Double

Dim GRAD_VECTOR As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant
Dim CTEMP_VECTOR As Variant
Dim DTEMP_VECTOR As Variant
Dim VTEMP_VECTOR As Variant
Dim ZTEMP_VECTOR As Variant

Dim CONST_BOX As Variant
Dim PARAM_VECTOR As Variant
Dim SCALE_VECTOR As Variant

Dim CONVERG_VAL As Integer
Dim CHECK_FLAG As Boolean

Dim LAMBDA As Double
Dim epsilon As Double
Dim tolerance As Double

On Error GoTo ERROR_LABEL

tolerance = 10 ^ -13
LAMBDA = tolerance

CONST_BOX = CONST_RNG
PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NROWS = UBound(PARAM_VECTOR, 1)

If IsArray(SCALE_RNG) = True Then
    SCALE_VECTOR = SCALE_RNG
    If UBound(SCALE_VECTOR, 1) = 1 Then: SCALE_VECTOR = MATRIX_TRANSPOSE_FUNC(SCALE_VECTOR)
Else
    ReDim SCALE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        SCALE_VECTOR(i, 1) = 1
    Next i
End If

NO_POINTS = 10

ReDim CTEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim DTEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim VTEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim ZTEMP_VECTOR(1 To NROWS, 1 To 1)

ReDim GRAD_VECTOR(1 To NROWS, 1 To 1)
ReDim BTEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim ATEMP_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    ATEMP_VECTOR(i, 1) = PARAM_VECTOR(i, 1)
Next i
XTEMP_ERR = 1
NTRIALS = 0

If GRAD_STR_NAME = "" Then
    ZTEMP_VECTOR = MULTVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, PARAM_VECTOR, SCALE_VECTOR, MIN_FLAG)
    NTRIALS = 2 * NROWS
Else
    ZTEMP_VECTOR = MULTVAR_CALL_GRAD_FUNC(GRAD_STR_NAME, PARAM_VECTOR, SCALE_VECTOR, MIN_FLAG)
End If

For i = 1 To NROWS
    DTEMP_VECTOR(i, 1) = -ZTEMP_VECTOR(i, 1)
Next i

COUNTER = 0

Do
    
    epsilon = tolerance
    k = 1
    For j = 1 To UBound(PARAM_VECTOR)
        CTEMP_VECTOR(j, 1) = DTEMP_VECTOR(j, 1)
    Next j
            
    TEMP_NORM = 0
    For j = 1 To UBound(PARAM_VECTOR) 'return the Euclidean norm of a vector
        TEMP_NORM = TEMP_NORM + PARAM_VECTOR(j, 1) ^ 2
    Next j
    TEMP_MULT = Sqr(TEMP_NORM)
    Do 'find a new direction CTEMP_VECTOR compatible with all constrains
        'CONVERG_VAL =  0 direction found
        '     -1 direction not found

        'compute the contraint k-th at the point PARAM_VECTOR
        ATEMP_SUM = 0
        For j = 1 To UBound(PARAM_VECTOR)
            ATEMP_SUM = ATEMP_SUM + CONST_BOX(k, j) * PARAM_VECTOR(j, 1)
        Next j
        ATEMP_SUM = ATEMP_SUM - CONST_BOX(k, UBound(PARAM_VECTOR) + 1)
        
        If Abs(ATEMP_SUM) <= (epsilon * (1 + TEMP_MULT)) Then
            'the point PARAM_VECTOR belongs to the contraint. Check if the
            'constraint is active compute the scalar product between the
            'current direction DTEMP_VECTOR and the gradient of the contraint
            BTEMP_SUM = 0
            ETEMP_SUM = 0
            For j = 1 To UBound(PARAM_VECTOR)
                BTEMP_SUM = BTEMP_SUM + CONST_BOX(k, j) * CTEMP_VECTOR(j, 1)
                ETEMP_SUM = ETEMP_SUM + CONST_BOX(k, j) ^ 2
            Next j
            BTEMP_SUM = BTEMP_SUM / ETEMP_SUM
            If BTEMP_SUM > epsilon Then
                'the constraint is active. Eliminate the normal component of
                'the direction
                For j = 1 To UBound(PARAM_VECTOR)
                    CTEMP_VECTOR(j, 1) = CTEMP_VECTOR(j, 1) - BTEMP_SUM * CONST_BOX(k, j)
                Next j
            End If
        End If
        k = k + 1
    Loop Until k > UBound(CONST_BOX)

    CTEMP_SUM = 0
    For j = 1 To UBound(CTEMP_VECTOR)
        CTEMP_SUM = CTEMP_SUM + CTEMP_VECTOR(j, 1) ^ 2
    Next j
    CTEMP_SUM = Sqr(CTEMP_SUM)
    If CTEMP_SUM > epsilon Then
        For j = 1 To UBound(CTEMP_VECTOR)
            CTEMP_VECTOR(j, 1) = CTEMP_VECTOR(j, 1) / CTEMP_SUM
        Next j
    End If

'check if the new direction and the old one have the same sense
' compute the scalar product between the 2 directions
    BTEMP_SUM = 0
    For j = 1 To UBound(PARAM_VECTOR)
        BTEMP_SUM = BTEMP_SUM + DTEMP_VECTOR(j, 1) * CTEMP_VECTOR(j, 1)
    Next j
    If BTEMP_SUM <= 0 Then
          CONVERG_VAL = -1
    Else
          CONVERG_VAL = 0
    End If
    
    If CONVERG_VAL < 0 Then: Exit Do
    'new direction
    For i = 1 To UBound(CTEMP_VECTOR)
        DTEMP_VECTOR(i, 1) = CTEMP_VECTOR(i, 1)
    Next i
    'find the bound
    epsilon = 10 ^ -12
    TEMP_MAX = 10 ^ 200
    TEMP_MIN = 0
    CHECK_FLAG = True
    k = 1 'Find a segement [a, TEMP_DELTA] belongs to the allowed
    'region defined by
    'the linear constraints, protection of then point "PARAM_VECTOR",
    'along direction "d"

    Do
        'compute the contraint k-th at the point PARAM_VECTOR
        ATEMP_SUM = 0
        For j = 1 To UBound(PARAM_VECTOR)
            ATEMP_SUM = ATEMP_SUM + CONST_BOX(k, j) * PARAM_VECTOR(j, 1)
        Next j
        ATEMP_SUM = ATEMP_SUM - CONST_BOX(k, UBound(PARAM_VECTOR) + 1)
        'compute the scalar product between the direction DTEMP_VECTOR and the
        'gradient of the contraint
        BTEMP_SUM = 0
        For j = 1 To UBound(PARAM_VECTOR)
            BTEMP_SUM = BTEMP_SUM + CONST_BOX(k, j) * DTEMP_VECTOR(j, 1)
        Next j
        'tree decision
        If Abs(ATEMP_SUM) <= epsilon Then  ' belongs to the constraint
            If BTEMP_SUM > epsilon Then  'direction not allowed
                  TEMP_MAX = 0
                  Exit Do
            Else: TEMP_VAL = -1  '(discharge)
            End If
        Else
            If Abs(BTEMP_SUM) <= epsilon Then  'direction parallel
                  TEMP_VAL = -1  '(discharge)
            Else: TEMP_VAL = -ATEMP_SUM / BTEMP_SUM
            End If
        End If
        'resizes the interval if necessary
        If ATEMP_SUM > epsilon Then
            'the point PARAM_VECTOR does not belong to the allowed region
            CHECK_FLAG = False '
            If TEMP_VAL < 0 Then TEMP_MIN = 10 ^ 200: Exit Do  'illimited region
            If TEMP_VAL > TEMP_MIN Then TEMP_MIN = TEMP_VAL
        Else
            'the point PARAM_VECTOR belongs to the allowed region
            If TEMP_VAL >= 0 And TEMP_VAL < TEMP_MAX Then TEMP_MAX = TEMP_VAL
        End If
        k = k + 1
    Loop Until k > UBound(CONST_BOX)

    'check the results
    '     -1 illimited segment
    '     -2 segment not found but PARAM_VECTOR belongs to the allowed region
    '     -3 segment not found and PARAM_VECTOR does not belongs to the
    '     allowed region

    If TEMP_MAX = 10 ^ 200 Then 'illimited
        If CHECK_FLAG Then CONVERG_VAL = -1 Else CONVERG_VAL = -3
        GoTo 1983
    End If
    If TEMP_MAX < TEMP_MIN Then  'unreal
        If CHECK_FLAG Then CONVERG_VAL = -1 Else CONVERG_VAL = -3
        GoTo 1983
    End If
    If TEMP_MAX = TEMP_MIN Then  'lenght zero
        If CHECK_FLAG Then CONVERG_VAL = -2 Else CONVERG_VAL = -3
        GoTo 1983
    End If
'compute the segment extreme [ATEMP_VECTOR, BTEMP_VECTOR]
    ReDim ATEMP_VECTOR(1 To UBound(PARAM_VECTOR), 1 To 1)
    ReDim BTEMP_VECTOR(1 To UBound(PARAM_VECTOR), 1 To 1)
    For j = 1 To UBound(PARAM_VECTOR)
        ATEMP_VECTOR(j, 1) = PARAM_VECTOR(j, 1) + DTEMP_VECTOR(j, 1) * TEMP_MIN
        BTEMP_VECTOR(j, 1) = PARAM_VECTOR(j, 1) + DTEMP_VECTOR(j, 1) * TEMP_MAX
    Next j
    CONVERG_VAL = 0 'segment found
    
1983:
    If CONVERG_VAL < 0 Then: Exit Do
    'minimize along the given direction
    PARAM_VECTOR = SEGMENT_OPTIMIZATION_FUNC(ATEMP_VECTOR, BTEMP_VECTOR, _
                    MIN_FUNC_VAL, h, FUNC_NAME_STR, SCALE_VECTOR, _
                    MIN_FLAG, NO_POINTS, nLOOPS)
    
    NTRIALS = NTRIALS + h

    If GRAD_STR_NAME = "" Then
        GRAD_VECTOR = MULTVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                    PARAM_VECTOR, SCALE_VECTOR, MIN_FLAG)
        NTRIALS = NTRIALS + 2 * NROWS
    Else
        GRAD_VECTOR = MULTVAR_CALL_GRAD_FUNC(GRAD_STR_NAME, _
                    PARAM_VECTOR, SCALE_VECTOR, MIN_FLAG)
    End If
   
    Select Case METHOD
        Case 0              'Gradient
            For i = 1 To NROWS
                DTEMP_VECTOR(i, 1) = -GRAD_VECTOR(i, 1)
            Next i
        Case 1              'Conjugate Gradient
            TEMP_NORM = 0
            For j = 1 To UBound(GRAD_VECTOR) 'return the Euclidean norm of a vector
                TEMP_NORM = TEMP_NORM + GRAD_VECTOR(j, 1) ^ 2
            Next j
            FIRST_MULT = Sqr(TEMP_NORM)
            TEMP_NORM = 0
            For j = 1 To UBound(ZTEMP_VECTOR) 'return the Euclidean norm of a vector
                TEMP_NORM = TEMP_NORM + ZTEMP_VECTOR(j, 1) ^ 2
            Next j
            SECOND_MULT = Sqr(TEMP_NORM)
            
            THIRD_MULT = 0
            For i = 1 To NROWS
                THIRD_MULT = GRAD_VECTOR(i, 1) * ZTEMP_VECTOR(i, 1)
            Next i
            
            If THIRD_MULT >= 0.2 * FIRST_MULT ^ 2 Then
                For i = 1 To NROWS
                    DTEMP_VECTOR(i, 1) = -GRAD_VECTOR(i, 1)
                Next i
            Else
                TEMP_DELTA = (FIRST_MULT / SECOND_MULT) ^ 2
                For i = 1 To NROWS
                    DTEMP_VECTOR(i, 1) = -GRAD_VECTOR(i, 1) + TEMP_DELTA * _
                                        DTEMP_VECTOR(i, 1)
                Next i
            End If
    End Select
   
    'check stop
    XTEMP_ERR = 0
    For i = 1 To NROWS
        XTEMP_ERR = XTEMP_ERR + (PARAM_VECTOR(i, 1) - ATEMP_VECTOR(i, 1)) ^ 2
    Next i
    XTEMP_ERR = Sqr(XTEMP_ERR)
    TEMP_NORM = 0
    For j = 1 To UBound(PARAM_VECTOR) 'return the Euclidean norm of a vector
        TEMP_NORM = TEMP_NORM + PARAM_VECTOR(j, 1) ^ 2
    Next j
    XTEMP_ABS = Sqr(TEMP_NORM)
    
    If XTEMP_ABS > 1 Then XTEMP_ERR = XTEMP_ERR / XTEMP_ABS
    If COUNTER > 4 Then
        If XTEMP_ERR > 0.3 * XINIT_ERR And XTEMP_ERR < 10 ^ -5 Then
            'convergence slow, reduce the accuracy
            LAMBDA = LAMBDA * 10
        End If
    End If
    If XTEMP_ERR < LAMBDA Then: Exit Do
    XINIT_ERR = XTEMP_ERR

    For i = 1 To NROWS
        ATEMP_VECTOR(i, 1) = PARAM_VECTOR(i, 1)
        ZTEMP_VECTOR(i, 1) = GRAD_VECTOR(i, 1)
    Next i
    
    COUNTER = COUNTER + 1
Loop Until NTRIALS > nLOOPS

Select Case OUTPUT
Case 0
    For i = 1 To UBound(PARAM_VECTOR, 1)
        PARAM_VECTOR(i, 1) = PARAM_VECTOR(i, 1) * SCALE_VECTOR(i, 1)
    Next i
    NLP_CG_OPTIM_FUNC = PARAM_VECTOR
Case 1
    NLP_CG_OPTIM_FUNC = MIN_FUNC_VAL
Case 2
    NLP_CG_OPTIM_FUNC = CONVERG_VAL
Case 3
    NLP_CG_OPTIM_FUNC = NTRIALS
Case Else
    NLP_CG_OPTIM_FUNC = XTEMP_ERR
End Select

Exit Function
ERROR_LABEL:
NLP_CG_OPTIM_FUNC = Err.number
End Function
