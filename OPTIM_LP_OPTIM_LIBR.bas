Attribute VB_Name = "OPTIM_LP_OPTIM_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : LP_OPTIMIZATION_FUNC

'DESCRIPTION   : This function solves a linear programming problem. A Linear
'program represent a problem in which we have to find the optima value
'(maximum or minimum) of a linear function of certain variables
'(objective function) subject to linear constraints on them.

'LP - Linear programming

'This function solves a linear programming problem by the Simplex algorithm
'Its input parameters are:
'• The coefficients vector of the linear objective function to optimize
'• The coefficients matrix of the linear constraints

'LIBRARY       : OPTIMIZATION
'GROUP         : LP_OPTIM
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LP_OPTIMIZATION_FUNC(ByRef COEF_RNG As Variant, _
ByRef CONST_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByVal OUTPUT As Integer = 2, _
Optional ByVal epsilon As Double = 10 ^ -13)

Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim ii As Integer
Dim jj As Integer

Dim iii As Integer
Dim jjj As Integer
Dim kkk As Integer

Dim NROWS As Integer
Dim NCOLUMNS As Integer

Dim ADATA_ARR() As Integer
Dim BDATA_ARR() As Integer

Dim MIN_VAL As Integer
Dim CONVERG_VAL As Integer

Dim COEF_VECTOR As Variant
Dim CONST_MATRIX As Variant
Dim PARAM_VECTOR As Variant
Dim DATA_MATRIX() As Double

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'epsilon is the absolute precision, which should be adjusted to the scale of
'your variables.
'CONVERG_VAL = (  ' 0 OK ,solution find,
            '-1 none solution, objective function is unbounded,
            '-2 none solution, objective function is bounded )

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

'COEF_RNG: The coefficients vector of the linear objective function to
'optimize: x1 + x2 + 3x3 - 0.5x4
'[1   1   3   -0.5]

'CONST_RNG: The coefficients matrix of the linear constraints the
'symbols "<" and "<=" or ">" and ">=" are numerically equivalent for
'this function.
'[1   0   2   0   <  10  ]
'[0   2   0  -7  <   0   ]
'[0   1  -1  2   >   0.5 ]
'[1   1   1   1   =  9   ]

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

If MIN_FLAG = False Then: MIN_VAL = 1 'maximization
COEF_VECTOR = COEF_RNG
If UBound(COEF_VECTOR, 2) = 1 Then: COEF_VECTOR = MATRIX_TRANSPOSE_FUNC(COEF_VECTOR)
NCOLUMNS = UBound(COEF_VECTOR, 2)  'maximum number of variables expected

CONST_MATRIX = CONST_RNG
NROWS = UBound(CONST_MATRIX, 1) 'maximum number of constraints expected

If UBound(CONST_MATRIX, 2) <> NCOLUMNS + 2 Then: GoTo ERROR_LABEL
'may be wrong selection

'rearrange Constraints symbols
For i = 1 To NROWS
    If InStr(1, CONST_MATRIX(i, NCOLUMNS + 1), ">") > 0 Or _
        Asc(CONST_MATRIX(i, NCOLUMNS + 1)) = 179 Then
        CONST_MATRIX(i, NCOLUMNS + 1) = ">"
    ElseIf InStr(1, CONST_MATRIX(i, NCOLUMNS + 1), "<") > 0 Or _
        Asc(CONST_MATRIX(i, NCOLUMNS + 1)) = 163 Then
        CONST_MATRIX(i, NCOLUMNS + 1) = "<"
    Else
        CONST_MATRIX(i, NCOLUMNS + 1) = "="
    End If
    If CONST_MATRIX(i, NCOLUMNS + 2) < 0 Then
        CONST_MATRIX(i, NCOLUMNS + 2) = -CONST_MATRIX(i, NCOLUMNS + 2)
        For j = 1 To NCOLUMNS
            CONST_MATRIX(i, j) = -CONST_MATRIX(i, j)
        Next j
        If CONST_MATRIX(i, NCOLUMNS + 1) = ">" Then
            CONST_MATRIX(i, NCOLUMNS + 1) = "<"
        ElseIf CONST_MATRIX(i, NCOLUMNS + 1) = "<" Then
            CONST_MATRIX(i, NCOLUMNS + 1) = ">"
        End If
    End If
Next i

jj = NCOLUMNS + 1
ii = NROWS + 2

ReDim DATA_MATRIX(1 To ii, 1 To jj)

ReDim ADATA_ARR(1 To NCOLUMNS)
ReDim BDATA_ARR(1 To NROWS)

ReDim PARAM_VECTOR(1 To NCOLUMNS + 1, 1 To 1)

'load function coefficients
For j = 1 To NCOLUMNS
    DATA_MATRIX(1, j + 1) = COEF_VECTOR(1, j)
    If MIN_VAL = 0 Then _
        DATA_MATRIX(1, j + 1) = -DATA_MATRIX(1, j + 1) 'minimization
Next j
''load Constraint coefficients
k = 1
For i = 1 To NROWS
    If CONST_MATRIX(i, NCOLUMNS + 1) = "<" Then
        iii = iii + 1
        GoSub 1983
    End If
Next i
For i = 1 To NROWS
    If CONST_MATRIX(i, NCOLUMNS + 1) = ">" Then
        jjj = jjj + 1
        GoSub 1983
    End If
Next i
For i = 1 To NROWS
    If CONST_MATRIX(i, NCOLUMNS + 1) = "=" Then
        kkk = kkk + 1
        GoSub 1983
    End If
Next i

Call LP_SIMPLEX_OPTIM_FUNC(DATA_MATRIX, NROWS, NCOLUMNS, ii, _
    jj, iii, jjj, kkk, CONVERG_VAL, _
    ADATA_ARR, BDATA_ARR, epsilon)


'---------------------------------------------------------------------------------------
Select Case CONVERG_VAL
'---------------------------------------------------------------------------------------
Case 0 'solution found
'---------------------------------------------------------------------------------------
    For i = 1 To NROWS
        If BDATA_ARR(i) <= NCOLUMNS Then
            j = BDATA_ARR(i)
            PARAM_VECTOR(BDATA_ARR(i), 1) = DATA_MATRIX(i + 1, 1)
        End If
    Next
    PARAM_VECTOR(NCOLUMNS + 1, 1) = DATA_MATRIX(1, 1)
    If MIN_VAL = 0 Then PARAM_VECTOR(NCOLUMNS + 1, 1) = _
                       -PARAM_VECTOR(NCOLUMNS + 1, 1) 'minimization

'---------------------------------------------------------------------------------------
    Select Case OUTPUT
'---------------------------------------------------------------------------------------
    Case 0 'Coefficients + Min/Max Func Value
'---------------------------------------------------------------------------------------
        LP_OPTIMIZATION_FUNC = PARAM_VECTOR
        Exit Function
'---------------------------------------------------------------------------------------
    Case 1 'Constraints Values
'---------------------------------------------------------------------------------------
        CONST_MATRIX = CONST_RNG
        CONST_MATRIX = MATRIX_TRANSPOSE_FUNC(MATRIX_REMOVE_COLUMNS_FUNC(CONST_MATRIX, UBound(CONST_MATRIX, 2) - 1, 2))
        PARAM_VECTOR = MATRIX_REMOVE_ROWS_FUNC(PARAM_VECTOR, UBound(PARAM_VECTOR, 1), 1)
        ReDim COEF_VECTOR(1 To UBound(CONST_MATRIX, 2), 1 To 1)
        For i = 1 To UBound(COEF_VECTOR, 1)
            COEF_VECTOR(i, 1) = _
                    MATRIX_SUM_PRODUCT_FUNC(MATRIX_GET_COLUMN_FUNC(CONST_MATRIX, i, 1), PARAM_VECTOR)
        Next i
        LP_OPTIMIZATION_FUNC = COEF_VECTOR
        Exit Function
'---------------------------------------------------------------------------------------
    Case Else 'Min/Max Func Value
'---------------------------------------------------------------------------------------
        PARAM_VECTOR = MATRIX_REMOVE_ROWS_FUNC(PARAM_VECTOR, UBound(PARAM_VECTOR, 1), 1)
        LP_OPTIMIZATION_FUNC = MATRIX_SUM_PRODUCT_FUNC(PARAM_VECTOR, MATRIX_TRANSPOSE_FUNC(COEF_VECTOR))
        Exit Function
'---------------------------------------------------------------------------------------
    End Select
'---------------------------------------------------------------------------------------
    LP_OPTIMIZATION_FUNC = PARAM_VECTOR
'---------------------------------------------------------------------------------------
Case -1
'---------------------------------------------------------------------------------------
'unbounded constrains region, so no feasible solution exists
    LP_OPTIMIZATION_FUNC = "Unbounded constraints region"
'---------------------------------------------------------------------------------------
Case -2
'---------------------------------------------------------------------------------------
'bounded constrains region, but no solution exists
    LP_OPTIMIZATION_FUNC = "No solution exists"
'---------------------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------------------
    GoTo ERROR_LABEL
'---------------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------------


Exit Function

1983:
    k = k + 1
    DATA_MATRIX(k, 1) = CONST_MATRIX(i, NCOLUMNS + 2)
    For j = 1 To NCOLUMNS
        DATA_MATRIX(k, j + 1) = -CONST_MATRIX(i, j)
    Next j
Return

ERROR_LABEL:
LP_OPTIMIZATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LP_CHECK_CONST_FUNC
'DESCRIPTION   : Symbols accepted ("<", ">", "=" "<=", ">=") by the Linear
'Programming solver function.
'LIBRARY       : OPTIMIZATION
'GROUP         : LP_OPTIM
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LP_CHECK_CONST_FUNC(ByVal SYMB_STR As String)
On Error GoTo ERROR_LABEL
    LP_CHECK_CONST_FUNC = False
    SYMB_STR = Trim(SYMB_STR)
    If SYMB_STR = "" Then
        LP_CHECK_CONST_FUNC = False
    ElseIf InStr(1, SYMB_STR, ">") > 0 Or Asc(SYMB_STR) = 179 Then
        LP_CHECK_CONST_FUNC = True
    ElseIf InStr(1, SYMB_STR, "<") > 0 Or Asc(SYMB_STR) = 163 Then
        LP_CHECK_CONST_FUNC = True
    ElseIf SYMB_STR = "=" Then
        LP_CHECK_CONST_FUNC = True
    Else
        LP_CHECK_CONST_FUNC = False
    End If
Exit Function
ERROR_LABEL:
    LP_CHECK_CONST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LP_CHECK_CONST_FUNC
'DESCRIPTION   : Symbols accepted ("<", ">", "=" "<=", ">=") by the Linear
'Programming solver function.
'LIBRARY       : OPTIMIZATION
'GROUP         : LP_OPTIM
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LP_SIMPLEX_OPTIM_FUNC(ByRef DATA_MATRIX() As Double, _
ByRef NROWS As Integer, _
ByRef NCOLUMNS As Integer, _
ByRef ii As Integer, _
ByRef jj As Integer, _
ByRef iii As Integer, _
ByRef jjj As Integer, _
ByRef kkk As Integer, _
ByRef CONVERG_VAL As Integer, _
ByRef ADATA_ARR() As Integer, _
ByRef BDATA_ARR() As Integer, _
ByRef epsilon As Double)

Dim i As Integer
Dim k As Integer
Dim l As Integer

Dim ll As Integer
Dim lll As Integer

Dim hh As Integer
Dim hhh As Integer

Dim MAX_VAL As Double
Dim TEMP_SUM As Double

Dim ATEMP_ARR() As Integer
Dim BTEMP_ARR() As Integer

On Error GoTo ERROR_LABEL

ReDim ATEMP_ARR(1 To NCOLUMNS)
ReDim BTEMP_ARR(1 To NROWS)

If (NROWS <> iii + jjj + kkk) Then Exit Function
'bad input constraint counts in LP_SIMPLEX_OPTIM_FUNC'
l = NCOLUMNS
For k = 1 To NCOLUMNS
    ATEMP_ARR(k) = k   'Initialize index list of columns admissible for exchange.
    ADATA_ARR(k) = k 'Initially make all variables right-hand.
Next k
For i = 1 To NROWS
    If (DATA_MATRIX(i + 1, 1) < 0) Then Exit Function
    BDATA_ARR(i) = NCOLUMNS + i
Next i
If (jjj + kkk = 0) Then GoTo 1983
For i = 1 To jjj
    BTEMP_ARR(i) = 1
Next i
For k = 1 To NCOLUMNS + 1
    TEMP_SUM = 0#
    For i = iii + 1 To NROWS
        TEMP_SUM = TEMP_SUM + DATA_MATRIX(i + 1, k)
    Next i
    DATA_MATRIX(NROWS + 2, k) = -TEMP_SUM
Next k
1981:  '<<<<<<<<<<<<< LP_OPTIMIZATION_FUNC algorithm begin
Call LP_SIMPLEX_MAX_FUNC(DATA_MATRIX, ii, jj, _
        NROWS + 1, ATEMP_ARR, l, 0, hhh, MAX_VAL)
If (MAX_VAL <= epsilon And DATA_MATRIX(NROWS + 2, 1) < -epsilon) Then
    CONVERG_VAL = -2
    Exit Function
ElseIf (MAX_VAL <= epsilon And DATA_MATRIX(NROWS + 2, 1) <= epsilon) Then
    For ll = iii + jjj + 1 To NROWS
        If (BDATA_ARR(ll) = ll + NCOLUMNS) Then
            Call LP_SIMPLEX_MAX_FUNC(DATA_MATRIX, _
                ii, jj, ll, ATEMP_ARR, l, 1, hhh, MAX_VAL)
            If (MAX_VAL > epsilon) Then GoTo 1982
        End If
    Next ll
    For i = iii + 1 To iii + jjj
        If (BTEMP_ARR(i - iii) = 1) Then
            For k = 1 To NCOLUMNS + 1
                DATA_MATRIX(i + 1, k) = -DATA_MATRIX(i + 1, k)
            Next k
        End If
    Next i
    GoTo 1983
End If
Call LP_SIMPLEX_PIVOT_FUNC(DATA_MATRIX, NROWS, _
    NCOLUMNS, ii, jj, ll, hhh, epsilon)
If (ll = 0) Then
    CONVERG_VAL = -1
    Exit Function
End If
1982:
Call LP_SIMPLEX_SWAP_FUNC(DATA_MATRIX, ii, jj, NROWS + 1, NCOLUMNS, ll, hhh)
'Exchange a left- and a right-hand variable (phase one), then update lists.
If (BDATA_ARR(ll) >= NCOLUMNS + iii + jjj + 1) Then
    For k = 1 To l
        If (ATEMP_ARR(k) = hhh) Then Exit For
    Next k
    l = l - 1
    For lll = k To l
        ATEMP_ARR(lll) = ATEMP_ARR(lll + 1)
    Next lll
Else
    hh = BDATA_ARR(ll) - iii - NCOLUMNS
    If (hh >= 1) Then 'Exchanged out an jjj type constraint.
        If (BTEMP_ARR(hh) <> 0) Then
            BTEMP_ARR(hh) = 0
            DATA_MATRIX(NROWS + 2, hhh + 1) = DATA_MATRIX(NROWS + 2, hhh + 1) + 1#
            For i = 1 To NROWS + 2
                DATA_MATRIX(i, hhh + 1) = -DATA_MATRIX(i, hhh + 1)
            Next i
        End If
    End If
End If
lll = ADATA_ARR(hhh) 'Update lists of left- and right-hand variables.
ADATA_ARR(hhh) = BDATA_ARR(ll)
BDATA_ARR(ll) = lll
GoTo 1981 'Still in phase one, go back to begin.
'End of phase one code for fnding an initial feasible solution.
'Now, in phase two, optimize it.
1983:
Call LP_SIMPLEX_MAX_FUNC(DATA_MATRIX, ii, jj, 0, ATEMP_ARR, l, 0, hhh, MAX_VAL)
If (MAX_VAL <= epsilon) Then 'Done. Solution found. Return .
    CONVERG_VAL = 0
    Exit Function
End If

Call LP_SIMPLEX_PIVOT_FUNC(DATA_MATRIX, NROWS, _
    NCOLUMNS, ii, jj, ll, hhh, epsilon)
'Locate a pivot element (phase two).
If (ll = 0) Then 'Objective function is unbounded. return.
    CONVERG_VAL = -1
    Exit Function
End If

Call LP_SIMPLEX_SWAP_FUNC(DATA_MATRIX, ii, jj, NROWS, NCOLUMNS, ll, hhh)
'Exchange a left- and a right-hand variable (phase two),
lll = ADATA_ARR(hhh) 'update lists of left- and right-hand variables,
ADATA_ARR(hhh) = BDATA_ARR(ll)
BDATA_ARR(ll) = lll
GoTo 1983 'return for another iteration.

Exit Function
ERROR_LABEL:
LP_SIMPLEX_OPTIM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LP_CHECK_CONST_FUNC
'DESCRIPTION   : Symbols accepted ("<", ">", "=" "<=", ">=") by the Linear
'Programming solver function.
'LIBRARY       : OPTIMIZATION
'GROUP         : LP_OPTIM
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LP_SIMPLEX_MAX_FUNC(ByRef DATA_MATRIX() As Double, _
ByRef ii As Integer, _
ByRef jj As Integer, _
ByRef ll As Integer, _
ByRef ADATA_ARR() As Integer, _
ByRef l As Integer, _
ByRef METHOD As Integer, _
ByRef hhh As Integer, _
ByRef MAX_VAL As Double)

'Determines the maximum of those elements whose index is
'contained in the supplied list
Dim k As Integer
Dim TEST_VAL As Double

On Error GoTo ERROR_LABEL

If (l <= 0) Then 'No eligible columns.
    MAX_VAL = 0#
Else
    hhh = ADATA_ARR(1)
    MAX_VAL = DATA_MATRIX(ll + 1, hhh + 1)
    For k = 2 To l
        If (METHOD = 0) Then
            TEST_VAL = DATA_MATRIX(ll + 1, ADATA_ARR(k) + 1) - MAX_VAL
        Else
            TEST_VAL = Abs(DATA_MATRIX(ll + 1, ADATA_ARR(k) + 1)) - Abs(MAX_VAL)
        End If
        If (TEST_VAL > 0) Then
            MAX_VAL = DATA_MATRIX(ll + 1, ADATA_ARR(k) + 1)
            hhh = ADATA_ARR(k)
        End If
    Next k
End If

Exit Function
ERROR_LABEL:
LP_SIMPLEX_MAX_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LP_CHECK_CONST_FUNC
'DESCRIPTION   : Symbols accepted ("<", ">", "=" "<=", ">=") by the Linear
'Programming solver function.
'LIBRARY       : OPTIMIZATION
'GROUP         : LP_OPTIM
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function LP_SIMPLEX_PIVOT_FUNC(ByRef DATA_MATRIX() As Double, _
ByRef NROWS As Integer, _
ByRef NCOLUMNS As Integer, _
ByRef ii As Integer, _
ByRef jj As Integer, _
ByRef ll As Integer, _
ByRef hhh As Integer, _
ByRef tolerance As Double)

'Locate a pivot element, taking degeneracy into account.
Dim i As Integer
Dim k As Integer

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double

Dim epsilon As Double

On Error GoTo ERROR_LABEL

epsilon = tolerance

ll = 0
For i = 1 To NROWS
    If (DATA_MATRIX(i + 1, hhh + 1) < -epsilon) Then GoTo 1982
Next i
Exit Function 'No possible pivots. Return with message.
1982:
CTEMP_VAL = -DATA_MATRIX(i + 1, 1) / DATA_MATRIX(i + 1, hhh + 1)
ll = i
For i = ll + 1 To NROWS
    If (DATA_MATRIX(i + 1, hhh + 1) < -epsilon) Then
        ATEMP_VAL = -DATA_MATRIX(i + 1, 1) / DATA_MATRIX(i + 1, hhh + 1)
        If (ATEMP_VAL < CTEMP_VAL) Then
            ll = i
            CTEMP_VAL = ATEMP_VAL
        ElseIf (ATEMP_VAL = CTEMP_VAL) Then 'We have a degeneracy.
            For k = 1 To NCOLUMNS
                DTEMP_VAL = -DATA_MATRIX(ll + 1, k + 1) / DATA_MATRIX(ll + 1, hhh + 1)
                BTEMP_VAL = -DATA_MATRIX(i + 1, k + 1) / DATA_MATRIX(i + 1, hhh + 1)
                If (BTEMP_VAL <> DTEMP_VAL) Then GoTo 1983
            Next k
1983:
            If (BTEMP_VAL < DTEMP_VAL) Then ll = i
        End If
    End If
Next i

Exit Function
ERROR_LABEL:
LP_SIMPLEX_PIVOT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LP_SIMPLEX_SWAP_FUNC
'DESCRIPTION   : Matrix operations to exchange a left-hand and right-hand variable.
'LIBRARY       : LINEAR PROGRAMMING
'GROUP         : OPTIMIZATION
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function LP_SIMPLEX_SWAP_FUNC(ByRef DATA_MATRIX() As Double, _
ByRef ii As Integer, _
ByRef jj As Integer, _
ByRef NROWS As Integer, _
ByRef NCOLUMNS As Integer, _
ByRef ll As Integer, _
ByRef hhh As Integer)

Dim i As Integer
Dim k As Integer

Dim PIVOT_VAL As Double

On Error GoTo ERROR_LABEL

PIVOT_VAL = 1# / DATA_MATRIX(ll + 1, hhh + 1)
For i = 1 To NROWS + 1
    If (i - 1 <> ll) Then
        DATA_MATRIX(i, hhh + 1) = DATA_MATRIX(i, hhh + 1) * PIVOT_VAL
        For k = 1 To NCOLUMNS + 1
            If (k - 1 <> hhh) Then
                DATA_MATRIX(i, k) = DATA_MATRIX(i, k) - _
                        DATA_MATRIX(ll + 1, k) * DATA_MATRIX(i, hhh + 1)
            End If
        Next k
    End If
Next i

For k = 1 To NCOLUMNS + 1
    If (k - 1 <> hhh) Then _
        DATA_MATRIX(ll + 1, k) = -DATA_MATRIX(ll + 1, k) * PIVOT_VAL
Next k
DATA_MATRIX(ll + 1, hhh + 1) = PIVOT_VAL

Exit Function
ERROR_LABEL:
LP_SIMPLEX_SWAP_FUNC = Err.number
End Function

