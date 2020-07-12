Attribute VB_Name = "INTEGRATION_ROMBERG_LIBR"

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------
' ROMBERG_INTEGRATION_FUNC - Romberg Integration of f(X_VAL).
'-----------------------------------------------------------------------------------------------------
'Using this function is very simple. For example, try
'=ROMBERG_INTEGRATION_FUNC("sin(2*pi*x)+cos(2*pi*x)", 0, 0.5)
'It returns the value:
'0.318309886183791 better approximate of 1E-16
'-----------------------------------------------------------------------------------------------------

'// PERFECT

Function ROMBERG_INTEGRATION_FUNC(ByVal FORMULA_STR As String, _
ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
Optional ByVal RANK_VAL As Double = 16, _
Optional ByVal tolerance As Double = 10 ^ -15, _
Optional ByVal ERROR_STR As String = "--", _
Optional ByRef PARSER_OBJ As clsMathParser)

'---------------------------------------------------------------------------------------------------------------
'This is an example to show how to build an Excel function to calculate a definite integral.
'The following is complete code for integration with Romberg method
'---------------------------------------------------------------------------------------------------------------
' ROMBERG_INTEGRATION_FUNC - Romberg Integration of f(x).
'---------------------------------------------------------------------------------------------------------------
'FORMULA_STR = function f(x) to integrate
'LOWER_VAL = lower integration limit
'UPPER_VAL = upper integration limit
'RANK_VAL = (optional default=16) Sets the max samples = 2^R
'tolerance = (optional default=10^-15) Sets the max error allowed
'---------------------------------------------------------------------------------------------------------------
'Algorithm exits when one of the two above limits is reached
'---------------------------------------------------------------------------------------------------------------

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long ' Nodes

Dim OK_FLAG As Boolean

Dim X_VAL As Double
Dim D_VAL As Double

Dim Y1_VAL As Double
Dim Y2_VAL As Double

Dim TEMP_SUM As Double

Dim R_MAT() As Double
Dim Y_ARR() As Double
Dim U_ARR(1 To 3) As Variant

Dim ERROR_VAL As Double

On Error GoTo ERROR_LABEL

'---------------------------------------------------------------------------------------------------
If PARSER_OBJ Is Nothing Then: Set PARSER_OBJ = New clsMathParser
'---------------------------------------------------------------------------------------------------
If VarType(FORMULA_STR) <> vbString Then
    ERROR_STR = "string missing" 'Err.Raise 1001
End If
'---------------------------------------------------------------------------------------------------
'Define expression, perform syntax check and get its handle
OK_FLAG = PARSER_OBJ.StoreExpression(FORMULA_STR)
If Not OK_FLAG Then
    ERROR_STR = PARSER_OBJ.ErrorDescription 'Err.Raise 1001
End If
'---------------------------------------------------------------------------------------------------
k = 0: l = 1
ReDim R_MAT(0 To RANK_VAL, 0 To RANK_VAL)
ReDim Y_ARR(0 To l)
'---------------------------------------------------------------------------------------------------
Y_ARR(0) = PARSER_OBJ.Eval1(LOWER_VAL)
If Err.number <> 0 Then GoTo ERROR_LABEL
Y_ARR(1) = PARSER_OBJ.Eval1(UPPER_VAL)
If Err.number <> 0 Then GoTo ERROR_LABEL
'---------------------------------------------------------------------------------------------------
D_VAL = UPPER_VAL - LOWER_VAL
R_MAT(k, k) = D_VAL * (Y_ARR(0) + Y_ARR(1)) / 2
'---------------------------------------------------------------------------------------------------
'start loop
Do
    k = k + 1
    l = 2 * l
    D_VAL = D_VAL / 2
    'compute e reorganize the vector of function values
    ReDim Preserve Y_ARR(0 To l)
    For i = l To 1 Step -1
        If i Mod 2 = 0 Then
            Y_ARR(i) = Y_ARR(i / 2)
        Else
            X_VAL = LOWER_VAL + i * D_VAL
            Y_ARR(i) = PARSER_OBJ.Eval1(X_VAL)
            If Err.number <> 0 Then GoTo ERROR_LABEL
        End If
    Next i
        'now begin with Romberg method
    TEMP_SUM = 0
    For i = 1 To l
        TEMP_SUM = TEMP_SUM + Y_ARR(i) + Y_ARR(i - 1) 'trapezoidal formula
    Next i
    R_MAT(k, 0) = D_VAL * TEMP_SUM / 2
    For j = 1 To k
        Y1_VAL = R_MAT(k - 1, j - 1)
        Y2_VAL = R_MAT(k, j - 1)
        R_MAT(k, j) = Y2_VAL + (Y2_VAL - Y1_VAL) / (4 ^ j - 1) 'Richardson's extrapolation
    Next j 'check error
    ERROR_VAL = Abs(R_MAT(k, k) - R_MAT(k, k - 1))
    If Abs(R_MAT(k, k)) > 10 Then
        ERROR_VAL = ERROR_VAL / Abs(R_MAT(k, k))
    End If
Loop Until ERROR_VAL < tolerance Or k >= RANK_VAL
'---------------------------------------------------------------------------------------------------
U_ARR(1) = R_MAT(k, k)
U_ARR(2) = 2 ^ k
U_ARR(3) = ERROR_VAL

'---------------------------------------------------------------------------------------------------
ROMBERG_INTEGRATION_FUNC = U_ARR
'---------------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ROMBERG_INTEGRATION_FUNC = ERROR_STR
End Function


'-------------------------------------------------------------------------------------------------------
'INTEGRATION routine for bidimensional functions f(x, y) in normal domains
'Method used is 2D-Romberg algorithm

'Input:
'Funct = f(x, y) integration function
'XMin(1) and XMax(1) are the boundary for the x-axis
'YMin(2) and YMax(2) are the boundary for the y-axis
'n: set the max iterations for Romberg method; points used are about 4^n
'ACCURACY_VAL set the relative error

'Output:
'Approximated integral
'Estimated error
'Counter of points evaluated
'Returns any error detected
'http://digilander.libero.it/foxes/integr/double_integrals.htm
'-------------------------------------------------------------------------------------------------------

Function ROMBERG_2D_INTEGRATION_FUNC(ByRef FORMULAS_RNG As Variant, _
ByRef XMIN_RNG As Variant, _
ByRef XMAX_RNG As Variant, _
ByRef YMIN_RNG As Variant, _
ByRef YMAX_RNG As Variant, _
Optional ByRef KMAX_RNG As Variant = 9, _
Optional ByRef ERROR_RNG As Variant = 10 ^ -7, _
Optional ByRef POLAR_FLAG_RNG As Variant = False)

'left bound function  h1(y)
'right bound function  h2(y)
'lower bound function  g1(x)
'upper bound function  g2(x)
'integration function  f(x,y)
'Polar --> polar coordinates TRUE/FALSE

'Remember to compare with the exact value (if known). This will be
'useful to compute the true error

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long '%
Dim jj As Long '%
Dim kk As Long
Dim iii As Long
Dim jjj As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim T0_VAL As Double
Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double

Dim Y0_VAL As Double
Dim Y1_VAL As Double
Dim Y2_VAL As Double
Dim TEMP_SUM As Double

Dim H_ARR(1 To 2) As Double
Dim P_ARR(1 To 2) As Double

Dim PMAX_ARR(1 To 2) As Double
Dim PMIN_ARR(1 To 2) As Double

Dim TEMP0_VAL As Double
Dim GRID_ARR() As Double
Dim ERROR_VAL As Double
Dim VARID_ARR() As Long '%()
Dim EXPR_STR As String
Dim INTORD_ARR() As Long '%()
Dim VARMAX_ARR() As Long '%()

Dim OK_FLAG As Boolean
Dim ERROR_FLAG As Boolean

Dim TEMP_MATRIX() As Double

Dim NSIZE As Variant
Dim ACCURACY_VAL As Variant
Dim FORMULAS_STR As String
Dim ERROR_STR As String
Dim POLAR_FLAG As Variant
Dim LOWER_BOUND_ARR() As Variant
Dim UPPER_BOUND_ARR() As Variant

Dim HEADINGS_STR As String
Dim INTEGR_4POINTS_VAL As Double

Dim FORMULAS_VECTOR As Variant
Dim XMIN_VECTOR As Variant
Dim XMAX_VECTOR As Variant
Dim YMIN_VECTOR As Variant
Dim YMAX_VECTOR As Variant
Dim KMAX_VECTOR As Variant
Dim ERROR_VECTOR As Variant
Dim POLAR_FLAG_VECTOR As Variant

Dim DATA_MATRIX As Variant

Dim nLOOPS As Long
Dim epsilon As Double
Dim tolerance As Double

On Error GoTo ERROR_LABEL

nLOOPS = 9
epsilon = 10 ^ -7

If IsArray(FORMULAS_RNG) Then
    FORMULAS_VECTOR = FORMULAS_RNG
    If UBound(FORMULAS_VECTOR, 1) = 1 Then
        FORMULAS_VECTOR = MATRIX_TRANSPOSE_FUNC(FORMULAS_VECTOR)
    End If
Else
    ReDim FORMULAS_VECTOR(1 To 1, 1 To 1)
    FORMULAS_VECTOR(1, 1) = FORMULAS_RNG
End If
NCOLUMNS = UBound(FORMULAS_VECTOR, 1)

If IsArray(XMIN_RNG) Then
    XMIN_VECTOR = XMIN_RNG
    If UBound(XMIN_VECTOR, 1) = 1 Then
        XMIN_VECTOR = MATRIX_TRANSPOSE_FUNC(XMIN_VECTOR)
    End If
Else
    ReDim XMIN_VECTOR(1 To 1, 1 To 1)
    XMIN_VECTOR(1, 1) = XMIN_RNG
End If
If NCOLUMNS <> UBound(XMIN_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(XMAX_RNG) Then
    XMAX_VECTOR = XMAX_RNG
    If UBound(XMAX_VECTOR, 1) = 1 Then
        XMAX_VECTOR = MATRIX_TRANSPOSE_FUNC(XMAX_VECTOR)
    End If
Else
    ReDim XMAX_VECTOR(1 To 1, 1 To 1)
    XMAX_VECTOR(1, 1) = XMAX_RNG
End If
If NCOLUMNS <> UBound(XMAX_VECTOR, 1) Then: GoTo ERROR_LABEL


If IsArray(YMIN_RNG) Then
    YMIN_VECTOR = YMIN_RNG
    If UBound(YMIN_VECTOR, 1) = 1 Then
        YMIN_VECTOR = MATRIX_TRANSPOSE_FUNC(YMIN_VECTOR)
    End If
Else
    ReDim YMIN_VECTOR(1 To 1, 1 To 1)
    YMIN_VECTOR(1, 1) = YMIN_RNG
End If
If NCOLUMNS <> UBound(YMIN_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(YMAX_RNG) Then
    YMAX_VECTOR = YMAX_RNG
    If UBound(YMAX_VECTOR, 1) = 1 Then
        YMAX_VECTOR = MATRIX_TRANSPOSE_FUNC(YMAX_VECTOR)
    End If
Else
    ReDim YMAX_VECTOR(1 To 1, 1 To 1)
    YMAX_VECTOR(1, 1) = YMAX_RNG
End If
If NCOLUMNS <> UBound(YMAX_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(KMAX_RNG) Then
    KMAX_VECTOR = KMAX_RNG
    If UBound(KMAX_VECTOR, 1) = 1 Then
        KMAX_VECTOR = MATRIX_TRANSPOSE_FUNC(KMAX_VECTOR)
    End If
Else
    ReDim KMAX_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NCOLUMNS
        KMAX_VECTOR(i, 1) = IIf(IsNumeric(KMAX_RNG), KMAX_RNG, nLOOPS)
    Next i
End If
If NCOLUMNS <> UBound(KMAX_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(ERROR_RNG) Then
    ERROR_VECTOR = ERROR_RNG
    If UBound(ERROR_VECTOR, 1) = 1 Then
        ERROR_VECTOR = MATRIX_TRANSPOSE_FUNC(ERROR_VECTOR)
    End If
Else
    ReDim ERROR_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NCOLUMNS
        ERROR_VECTOR(i, 1) = IIf(IsNumeric(ERROR_RNG), ERROR_RNG, epsilon)
    Next i
End If
If NCOLUMNS <> UBound(ERROR_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(POLAR_FLAG_RNG) Then
    POLAR_FLAG_VECTOR = POLAR_FLAG_RNG
    If UBound(POLAR_FLAG_VECTOR, 1) = 1 Then
        POLAR_FLAG_VECTOR = MATRIX_TRANSPOSE_FUNC(POLAR_FLAG_VECTOR)
    End If
Else
    ReDim POLAR_FLAG_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NCOLUMNS
        POLAR_FLAG_VECTOR(i, 1) = IIf(POLAR_FLAG_RNG <> "", POLAR_FLAG_RNG, False)
    Next i
End If
If NCOLUMNS <> UBound(POLAR_FLAG_VECTOR, 1) Then: GoTo ERROR_LABEL

HEADINGS_STR = "F(xy),Xmin,Xmax,Ymin,Ymax,K max,ErrRel Max,Polar," & _
"Integral,Time,Points,Estim. Error rel,Error Msg,"

ReDim DATA_MATRIX(0 To NCOLUMNS, 1 To 13)
i = 1
For k = 1 To 13
    j = InStr(i, HEADINGS_STR, ",")
    DATA_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
    i = j + 1
Next k
'----------------------------------------------------------------------------------
For kk = 1 To NCOLUMNS
'----------------------------------------------------------------------------------
    DATA_MATRIX(kk, 1) = FORMULAS_VECTOR(kk, 1)
    If DATA_MATRIX(kk, 1) = "" Then
        ERROR_STR = "Integration function missing"
        GoTo 1983
    End If
    DATA_MATRIX(kk, 2) = XMIN_VECTOR(kk, 1)
    DATA_MATRIX(kk, 3) = XMAX_VECTOR(kk, 1)
    DATA_MATRIX(kk, 4) = YMIN_VECTOR(kk, 1)
    DATA_MATRIX(kk, 5) = YMAX_VECTOR(kk, 1)
    If DATA_MATRIX(kk, 2) = "" Or DATA_MATRIX(kk, 3) = "" Or _
       DATA_MATRIX(kk, 4) = "" Or DATA_MATRIX(kk, 5) = "" Then
        ERROR_STR = "boundary limits missing"
        GoTo 1983
    End If
    
    DATA_MATRIX(kk, 6) = KMAX_VECTOR(kk, 1)
    If DATA_MATRIX(kk, 6) = "" Then: DATA_MATRIX(kk, 6) = nLOOPS
    
    DATA_MATRIX(kk, 7) = ERROR_VECTOR(kk, 1)
    If DATA_MATRIX(kk, 7) = "" Then: DATA_MATRIX(kk, 7) = epsilon
    
    DATA_MATRIX(kk, 8) = POLAR_FLAG_VECTOR(kk, 1)
    If DATA_MATRIX(kk, 8) = "" Then: DATA_MATRIX(kk, 8) = False
    
    ReDim LOWER_BOUND_ARR(1 To 2)
    ReDim UPPER_BOUND_ARR(1 To 2)
    
    FORMULAS_STR = DATA_MATRIX(kk, 1)
    LOWER_BOUND_ARR(1) = DATA_MATRIX(kk, 2)
    UPPER_BOUND_ARR(1) = DATA_MATRIX(kk, 3)
    
    LOWER_BOUND_ARR(2) = DATA_MATRIX(kk, 4)
    UPPER_BOUND_ARR(2) = DATA_MATRIX(kk, 5)
    NSIZE = DATA_MATRIX(kk, 6)
    ACCURACY_VAL = DATA_MATRIX(kk, 7) 'ErrRel max
    POLAR_FLAG = DATA_MATRIX(kk, 8)
    T0_VAL = Timer

    GoSub INTEGRATION_LINE
    DATA_MATRIX(kk, 10) = Timer - T0_VAL

    DATA_MATRIX(kk, 9) = TEMP_MATRIX(k, k)  'approximated integral
    If jj = 1 Then DATA_MATRIX(kk, 9) = -DATA_MATRIX(kk, 9) 'change sign for Y-Domain
    DATA_MATRIX(kk, 11) = (NROWS + 1) * (NROWS + 1)  'points evaluated
    DATA_MATRIX(kk, 12) = ERROR_VAL  'estimate error
1983:
    DATA_MATRIX(kk, 13) = ERROR_STR
'----------------------------------------------------------------------------------
Next kk
'----------------------------------------------------------------------------------

ROMBERG_2D_INTEGRATION_FUNC = DATA_MATRIX

'---------------------------------------------------------------------------------------
Exit Function
'---------------------------------------------------------------------------------------
INTEGRATION_LINE:
'---------------------------------------------------------------------------------------
    ReDim RESULTS_ARR(1 To 3)
    ReDim VARID_ARR(1 To 2)
    ReDim INTORD_ARR(1 To 2)
    ReDim VARMAX_ARR(1 To 2)
    
    ERROR_STR = "" 'initialize return message variable
    'parse and syntax check of f(x,y)
    
    Dim FUN_OBJ As New clsMathParser
    EXPR_STR = FORMULAS_STR
    OK_FLAG = FUN_OBJ.StoreExpression(EXPR_STR)
    If Not OK_FLAG Then
        ERROR_STR = FUN_OBJ.Expression + " : " + FUN_OBJ.ErrorDescription
        GoTo 1984
    End If
    'parse and syntax check of bound function -----------------
    Dim FMIN_OBJ(1 To 2) As New clsMathParser
    Dim FMAX_OBJ(1 To 2) As New clsMathParser
    For i = 1 To 2
        EXPR_STR = LOWER_BOUND_ARR(i)
        OK_FLAG = FMIN_OBJ(i).StoreExpression(EXPR_STR)
        If Not OK_FLAG Then
            ERROR_STR = FMIN_OBJ(i).Expression & " : " & FMIN_OBJ(i).ErrorDescription
            GoTo 1984
        End If
        EXPR_STR = UPPER_BOUND_ARR(i)
        OK_FLAG = FMAX_OBJ(i).StoreExpression(EXPR_STR)
        If Not OK_FLAG Then
            ERROR_STR = FMAX_OBJ(i).Expression & " : " & FMAX_OBJ(i).ErrorDescription
            GoTo 1984
        End If
    Next i
        
    'Check and build the variables Index
    If FUN_OBJ.VarTop > 2 Then
        ERROR_STR = "too many variables for f(x,y) function"
        GoTo 1984
    ElseIf FUN_OBJ.VarTop = 0 Then 'detect if the integration function is constant
        Y0_VAL = FUN_OBJ.Eval  'store the value for future elaboration
    End If
    For i = 1 To FUN_OBJ.VarTop
        Select Case LCase(FUN_OBJ.VarName(i))
        Case "x"
            VARID_ARR(1) = i
        Case "y"
            VARID_ARR(2) = i
        Case Else
            ERROR_STR = "Sorry, variables must be x, y"
            GoTo 1984
        End Select
    Next i
    
    'check if the domain is normal to a specific axes and build the integration order
    For i = 1 To 2
        VARMAX_ARR(i) = MAXIMUM_FUNC(FMIN_OBJ(i).VarTop, FMAX_OBJ(i).VarTop)
        INTORD_ARR(i) = i
    Next i
    If VARMAX_ARR(INTORD_ARR(1)) > VARMAX_ARR(INTORD_ARR(2)) Then
        TEMP0_VAL = INTORD_ARR(1)
        INTORD_ARR(1) = INTORD_ARR(2)
        INTORD_ARR(2) = TEMP0_VAL
    End If
    If VARMAX_ARR(INTORD_ARR(1)) > 0 Then
        ERROR_STR = "Normal domain not found for any axes"
        GoTo 1984
    End If
    
    For i = 1 To 2
        If VARMAX_ARR(INTORD_ARR(i)) >= i Then
            ERROR_STR = "Bounding error"
            GoTo 1984
        End If
    Next i
    'calculate the initial box -----------
    ii = INTORD_ARR(1)
    jj = INTORD_ARR(2)
    PMIN_ARR(ii) = FMIN_OBJ(ii).Eval
    PMAX_ARR(ii) = FMAX_OBJ(ii).Eval
    PMIN_ARR(jj) = FMIN_OBJ(jj).Eval1(PMIN_ARR(ii))
    PMAX_ARR(jj) = PMIN_ARR(jj)
    TEMP0_VAL = FMIN_OBJ(jj).Eval1(PMAX_ARR(ii))
    PMAX_ARR(jj) = MAXIMUM_FUNC(PMAX_ARR(jj), TEMP0_VAL)
    PMIN_ARR(jj) = MINIMUM_FUNC(PMIN_ARR(jj), TEMP0_VAL)
    TEMP0_VAL = FMAX_OBJ(jj).Eval1(PMIN_ARR(ii))
    PMAX_ARR(jj) = MAXIMUM_FUNC(PMAX_ARR(jj), TEMP0_VAL)
    PMIN_ARR(jj) = MINIMUM_FUNC(PMIN_ARR(jj), TEMP0_VAL)
    TEMP0_VAL = FMAX_OBJ(jj).Eval1(PMAX_ARR(ii))
    PMAX_ARR(jj) = MAXIMUM_FUNC(PMAX_ARR(jj), TEMP0_VAL)
    PMIN_ARR(jj) = MINIMUM_FUNC(PMIN_ARR(jj), TEMP0_VAL)
    '-------------------------------------
    
    k = 0
    NROWS = 2 ^ k
    For i = 1 To 2: H_ARR(i) = (PMAX_ARR(i) - PMIN_ARR(i)) / NROWS: Next
    ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)
    ERROR_VAL = 1#
    tolerance = ACCURACY_VAL  'set relative error limit
    Do Until k >= NSIZE Or ERROR_VAL <= tolerance
        k = k + 1
        'build the mesh
        ReDim GRID_ARR(0 To NROWS, 0 To NROWS, 1 To 3)
        For i = 0 To NROWS
            P_ARR(ii) = PMIN_ARR(ii) + H_ARR(ii) * i
            PMIN_ARR(jj) = FMIN_OBJ(jj).Eval1(P_ARR(ii))
            PMAX_ARR(jj) = FMAX_OBJ(jj).Eval1(P_ARR(ii))
            
            If ERROR_FLAG Then
                ERROR_STR = "Wrong domain bound"
                GoTo 1984
            End If
            H_ARR(jj) = (PMAX_ARR(jj) - PMIN_ARR(jj)) / NROWS
            For j = 0 To NROWS
                P_ARR(jj) = PMIN_ARR(jj) + H_ARR(jj) * j
                GRID_ARR(i, j, 1) = P_ARR(1)
                GRID_ARR(i, j, 2) = P_ARR(2)
                If FUN_OBJ.VarTop > 0 Then
                    If POLAR_FLAG Then
                        FUN_OBJ.VarValue(VARID_ARR(1)) = P_ARR(1) * Cos(P_ARR(2))
                        FUN_OBJ.VarValue(VARID_ARR(2)) = P_ARR(1) * Sin(P_ARR(2))
                    Else
                        FUN_OBJ.VarValue(VARID_ARR(1)) = P_ARR(1)
                        FUN_OBJ.VarValue(VARID_ARR(2)) = P_ARR(2)
                    End If
                    GRID_ARR(i, j, 3) = FUN_OBJ.Eval
                Else
                    GRID_ARR(i, j, 3) = Y0_VAL
                End If
                If ERROR_FLAG Then ERROR_STR = "Singularity: dubious accuracy" 'Evaluation error
                If POLAR_FLAG Then GRID_ARR(i, j, 3) = GRID_ARR(i, j, 3) * P_ARR(1)
            Next j
        Next i
        'integral computing with 2D-trapezoidal formula
        TEMP_SUM = 0
        For i = 0 To NROWS - 1
            For j = 0 To NROWS - 1 'Integrate 4 Points
                iii = i + 1
                jjj = j + 1
                TEMP1_VAL = (GRID_ARR(iii, j, 1) - GRID_ARR(i, j, 1)) * (GRID_ARR(iii, jjj, 2) - GRID_ARR(iii, j, 2)) - (GRID_ARR(iii, jjj, 1) - GRID_ARR(iii, j, 1)) * (GRID_ARR(iii, j, 2) - GRID_ARR(i, j, 2))
                TEMP2_VAL = (GRID_ARR(i, jjj, 1) - GRID_ARR(iii, jjj, 1)) * (GRID_ARR(i, j, 2) - GRID_ARR(i, jjj, 2)) - (GRID_ARR(i, j, 1) - GRID_ARR(i, jjj, 1)) * (GRID_ARR(i, jjj, 2) - GRID_ARR(iii, jjj, 2))
                INTEGR_4POINTS_VAL = (TEMP1_VAL * (GRID_ARR(i, j, 3) + GRID_ARR(iii, j, 3) + GRID_ARR(iii, jjj, 3)) + TEMP2_VAL * (GRID_ARR(i, j, 3) + GRID_ARR(i, jjj, 3) + GRID_ARR(iii, jjj, 3))) / 6
                TEMP_SUM = TEMP_SUM + INTEGR_4POINTS_VAL
            Next j
        Next i
        TEMP_MATRIX(k, 1) = TEMP_SUM
        NROWS = 2 * NROWS
        For i = 1 To 2
            H_ARR(i) = H_ARR(i) / 2
        Next i
        For j = 2 To k 'Richardson's extrapolation
            Y1_VAL = TEMP_MATRIX(k - 1, j - 1)
            Y2_VAL = TEMP_MATRIX(k, j - 1)
            TEMP_MATRIX(k, j) = Y2_VAL + (Y2_VAL - Y1_VAL) / (4 ^ (j - 1) - 1)
        Next j
        'error loop evaluation
        If k > 1 Then
            ERROR_VAL = Abs((TEMP_MATRIX(k, k) - TEMP_MATRIX(k, k - 1)))
            If Abs(TEMP_MATRIX(k, k)) > 1 Then ERROR_VAL = ERROR_VAL / Abs(TEMP_MATRIX(k, k))
        End If
    Loop
1984:
'------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------
ERROR_LABEL:
ROMBERG_2D_INTEGRATION_FUNC = Err.number
End Function
