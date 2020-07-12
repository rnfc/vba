Attribute VB_Name = "FINAN_PORT_WEIGHTS_SHARPE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_SHARPE_WEIGHTS_OPTIMIZER_FUNC

'DESCRIPTION   : The Gradient Method: algorithm implementation of W. Sharpe's
'algorithm to solve the standard asset allocation problem.

'Algorithm
'The routine uses the gradient quadratic programming method of Sharpe ["An Algorithm for Portfolio Improvement," ,
'in Advances in Mathematical Programming and Financial Planning, K.D.Lawrence, J.B. Guerard, Jr., and Gary D. Reeves,
'Editors, JAI Press, Inc., 1987, pp. 155-170] to find the feasible portfolio with the maximum possible Utility:

'Up = ep - ((sdp ^ 2) / T)
'where:
'Up = the utility of the portfolio
'ep = the portfolio's expected return
'sdp = the portfolio's standard deviation of return
't = the investor's risk tolerance

'Optimization algorithm for a standard portfolio allocation problem
'REFERENCE: http://www.stanford.edu/~wfsharpe/mia/opt/mia_opt1.htm

'LIBRARY       : PORTFOLIO
'GROUP         : GRADIENT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_SHARPE_WEIGHTS_OPTIMIZER_FUNC(ByVal BUDGET_VAL As Double, _
ByVal RISK_TOLERANCE As Double, _
ByRef EXPECTED_RNG As Variant, _
ByRef COVAR_RNG As Variant, _
Optional ByRef LOWER_RNG As Variant = 0, _
Optional ByRef UPPER_RNG As Variant = 1)

'BUDGET: Investor Exposure (e.g., 100%)
'RISK_TOLERANCE: Investor risk tolerance

'Some Comments:
'The inputs are resampled exactly once.
'These calculations illustrate the instability of the classical efficient
'frontier due to statistical uncertainty in the inputs.
'The further we move away from the minimum variance portfolios (bottom left),
'the more pronounced the instability of the efficient frontier.
'therefore, the minimum variance portfolio is a "hedge" against
'"model risk"

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim BUY_VAL As Double
Dim SELL_VAL As Double

Dim DELTA_VAL As Double
Dim FACTOR_VAL As Double

Dim MEAN_VECTOR As Variant
Dim SIGMA_VECTOR As Variant

Dim COVAR_MATRIX As Variant
Dim LOWER_VECTOR As Variant
Dim UPPER_VECTOR As Variant
Dim EXPECTED_VECTOR As Variant
Dim WEIGHTS_VECTOR As Variant

Dim nLOOPS As Long
Dim tolerance As Double
Dim epsilon As Double

On Error GoTo ERROR_LABEL

EXPECTED_VECTOR = EXPECTED_RNG
If UBound(EXPECTED_VECTOR, 1) = 1 Then
    EXPECTED_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPECTED_VECTOR)
End If
NROWS = UBound(EXPECTED_VECTOR, 1)

COVAR_MATRIX = COVAR_RNG
If UBound(COVAR_MATRIX, 1) <> UBound(COVAR_MATRIX, 2) Then: GoTo ERROR_LABEL
If UBound(COVAR_MATRIX, 1) <> UBound(EXPECTED_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(LOWER_RNG) = True Then
    LOWER_VECTOR = LOWER_RNG
    If UBound(LOWER_VECTOR, 1) = 1 Then: _
        LOWER_VECTOR = MATRIX_TRANSPOSE_FUNC(LOWER_VECTOR)
    If UBound(COVAR_MATRIX, 1) <> UBound(LOWER_VECTOR, 1) Then: GoTo ERROR_LABEL
Else
    ReDim LOWER_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        LOWER_VECTOR(i, 1) = LOWER_RNG
    Next i
End If

If IsArray(UPPER_RNG) = True Then
    UPPER_VECTOR = UPPER_RNG
    If UBound(UPPER_VECTOR, 1) = 1 Then: _
        UPPER_VECTOR = MATRIX_TRANSPOSE_FUNC(UPPER_VECTOR)
    If UBound(COVAR_MATRIX, 1) <> UBound(UPPER_VECTOR, 1) Then: GoTo ERROR_LABEL
Else
    ReDim UPPER_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        UPPER_VECTOR(i, 1) = UPPER_RNG
    Next i
End If

' set maximum number of iterations and minimum marginal utility change
nLOOPS = 10000
epsilon = 0.0001
tolerance = 1E+200

' allocate memory for swap vector, marginal utility vector
ReDim SIGMA_VECTOR(1 To NROWS, 1 To 1)
ReDim MEAN_VECTOR(1 To NROWS, 1 To 1)

' set an initial feasible allocation
ReDim WEIGHTS_VECTOR(1 To NROWS, 1 To 1)

TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 1 To NROWS
    TEMP1_SUM = TEMP1_SUM + LOWER_VECTOR(i, 1)
    TEMP2_SUM = TEMP2_SUM + UPPER_VECTOR(i, 1) - LOWER_VECTOR(i, 1)
Next i
DELTA_VAL = (BUDGET_VAL - TEMP1_SUM) / TEMP2_SUM
For i = 1 To NROWS
    WEIGHTS_VECTOR(i, 1) = LOWER_VECTOR(i, 1) + DELTA_VAL * (UPPER_VECTOR(i, 1) - LOWER_VECTOR(i, 1))
Next i

' set interations k
k = 0

Do While 1 = 1
    ' compute marginal utilities
    For i = 1 To NROWS
        TEMP1_SUM = 0
        For j = 1 To NROWS
            TEMP1_SUM = TEMP1_SUM + WEIGHTS_VECTOR(j, 1) * COVAR_MATRIX(i, j)
        Next j
        MEAN_VECTOR(i, 1) = RISK_TOLERANCE * EXPECTED_VECTOR(i, 1) - 2 * TEMP1_SUM
    Next i
    
    ' fine best asset to buy and sell
    ii = 0: BUY_VAL = tolerance * -1
    jj = 0: SELL_VAL = tolerance
    For i = 1 To NROWS
        If WEIGHTS_VECTOR(i, 1) < UPPER_VECTOR(i, 1) Then ' possible buy
            If MEAN_VECTOR(i, 1) > BUY_VAL Then
                BUY_VAL = MEAN_VECTOR(i, 1)
                ii = i
            End If
        End If
        If WEIGHTS_VECTOR(i, 1) > LOWER_VECTOR(i, 1) Then ' possible sell
            If MEAN_VECTOR(i, 1) < SELL_VAL Then
                SELL_VAL = MEAN_VECTOR(i, 1)
                jj = i
            End If
        End If
    Next i
    ' terminate if change in mean is less than threshold value
    If (BUY_VAL - SELL_VAL) <= epsilon Then: Exit Do
    ' set up swap vector
    For i = 1 To NROWS: SIGMA_VECTOR(i, 1) = 0: Next i
    SIGMA_VECTOR(ii, 1) = 1: SIGMA_VECTOR(jj, 1) = -1
    ' compute optimal amount of swap without regard to asset bounds
    TEMP2_SUM = 0
    For i = 1 To NROWS
        TEMP1_SUM = 0
        For j = 1 To NROWS
            TEMP1_SUM = TEMP1_SUM + WEIGHTS_VECTOR(j, 1) * COVAR_MATRIX(i, j)
        Next j
        TEMP2_SUM = TEMP2_SUM + SIGMA_VECTOR(i, 1) * (RISK_TOLERANCE * EXPECTED_VECTOR(i, 1) - 2 * TEMP1_SUM)
    Next i
    TEMP3_SUM = 0
    For i = 1 To NROWS
        For j = 1 To NROWS
            TEMP3_SUM = TEMP3_SUM + SIGMA_VECTOR(i, 1) * SIGMA_VECTOR(j, 1) * COVAR_MATRIX(i, j)
        Next j
    Next i
    FACTOR_VAL = TEMP2_SUM / (2 * TEMP3_SUM)
    ' reduce amount if required to keep ii from exceeding its upper bound
    If FACTOR_VAL > (UPPER_VECTOR(ii, 1) - WEIGHTS_VECTOR(ii, 1)) Then
        FACTOR_VAL = UPPER_VECTOR(ii, 1) - WEIGHTS_VECTOR(ii, 1)
    End If
    ' reduce amount if required to keep jj from falling below lower bound
    If FACTOR_VAL > (WEIGHTS_VECTOR(jj, 1) - LOWER_VECTOR(jj, 1)) Then
        FACTOR_VAL = WEIGHTS_VECTOR(jj, 1) - LOWER_VECTOR(jj, 1)
    End If
    ' terminate if amount is zero
    If FACTOR_VAL = 0 Then: Exit Do
    ' count iteration and terminate if maximum k exceeded
    k = k + 1
    If k > nLOOPS Then: Exit Do
    ' change mix
    For i = 1 To NROWS: WEIGHTS_VECTOR(i, 1) = WEIGHTS_VECTOR(i, 1) + SIGMA_VECTOR(i, 1) * FACTOR_VAL: Next i
Loop

PORT_SHARPE_WEIGHTS_OPTIMIZER_FUNC = WEIGHTS_VECTOR

Exit Function
ERROR_LABEL:
PORT_SHARPE_WEIGHTS_OPTIMIZER_FUNC = Err.number
End Function
