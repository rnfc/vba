Attribute VB_Name = "FINAN_PORT_FRONTIER_BAYES_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_BAYES_STEIN_FRONTIER_FUNC

'DESCRIPTION   : The observed international home bias has traditionally been
'viewed as an anomaly. We provide statistical evidence contrary to this view
'within a mean-variance framework. We investigate two methods of estimating
'the expected return and covariance parameters: (i) the Bayes-Stein
'"shrinkage" algorithm, and (ii) the traditional Markowitz approach.

'In in-sample tests, neither the Bayes-Stein tangency allocation vector, nor
'the Markowitz tangency allocation vectors are significantly different from a
'100% domestic allocation (i.e. extreme home bias). The result is robust to
'the shorting of equity, and across foreign exchange hedge strategies.

'We also conduct out-of-sample tests with a view toward investment performance.
'We find that a 100% domestic allocation typically outperforms both the
'Bayes-Stein and Markowitz tangency portfolios. Overall, the theorized gains to
'international diversification appear difficult to capture in practice, and
'hence investors exhibiting a strong home bias are not necessarily acting
'irrationally.

'---------------------------------------------------------------------------------
'Domestic versus International Portfolio Selection: A Statistical
'Examination of the Home Bias by Larry R. Gorman / Bjorn N. Jorgensen
'published in Multinational Finance Journal, 2002, vol. 6, no. 3&4, pp. 131–166
'---------------------------------------------------------------------------------

'LIBRARY       : PORTFOLIO
'GROUP         : FRONTIER
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'UPDATE        : 04/01/2008
'************************************************************************************
'************************************************************************************

Function PORT_BAYES_STEIN_FRONTIER_FUNC(ByRef DATA_RNG As Variant, _
ByRef RISK_TOLERANCE_RNG As Variant, _
Optional ByRef FACTOR_RNG As Variant = 1, _
Optional ByVal BUDGET_VAL As Double = 1, _
Optional ByRef LOWER_RNG As Variant = 0, _
Optional ByRef UPPER_RNG As Variant = 1, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal COUNT_BASIS As Double = 250, _
Optional ByVal OUTPUT As Integer = 4)

Dim i As Long
Dim j As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double

Dim DATA_MATRIX As Variant
Dim LOWER_VECTOR As Variant
Dim UPPER_VECTOR As Variant
Dim FACTOR_VECTOR As Variant
Dim TRANSPOSE_FACTOR_VECTOR As Variant
Dim TOLERANCE_VECTOR As Variant

Dim TEMP_GROUP As Variant
Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant

Dim PSI_VAL As Double
Dim LAMBDA_VAL As Double
Dim MVP_RETURN_VAL As Double
Dim MVP_VOLATILITY_VAL As Double
Dim INVERSE_COVAR_MATRIX As Variant

Dim HIST_COVAR_MATRIX As Variant
Dim HIST_RETURNS_VECTOR As Variant
Dim HIST_MVP_WEIGHTS_VECTOR As Variant
 
Dim BAYES_STEIN_COVAR_MATRIX As Variant
Dim BAYES_STEIN_EXPECTED_RETURNS_VECTOR As Variant

Dim MARKOWITZ_FRONTIER_MATRIX As Variant
Dim BAYES_STEIN_FRONTIER_MATRIX As Variant
 
On Error GoTo ERROR_LABEL
'---------------------------------------------------------------------------------

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

TOLERANCE_VECTOR = RISK_TOLERANCE_RNG
If UBound(TOLERANCE_VECTOR, 1) = 1 Then: _
    TOLERANCE_VECTOR = MATRIX_TRANSPOSE_FUNC(TOLERANCE_VECTOR)
NSIZE = UBound(TOLERANCE_VECTOR, 1)
'---------------------------------------------------------------------------------

If IsArray(LOWER_RNG) = True Then
    LOWER_VECTOR = LOWER_RNG
    If UBound(LOWER_VECTOR, 1) = 1 Then: _
        LOWER_VECTOR = MATRIX_TRANSPOSE_FUNC(LOWER_VECTOR)
    If NCOLUMNS <> UBound(LOWER_VECTOR, 1) Then: GoTo ERROR_LABEL
Else
    ReDim LOWER_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NCOLUMNS
        LOWER_VECTOR(i, 1) = LOWER_RNG
    Next i
End If
'---------------------------------------------------------------------------------

If IsArray(UPPER_RNG) = True Then
    UPPER_VECTOR = UPPER_RNG
    If UBound(UPPER_VECTOR, 1) = 1 Then: _
        UPPER_VECTOR = MATRIX_TRANSPOSE_FUNC(UPPER_VECTOR)
    If NCOLUMNS <> UBound(UPPER_VECTOR, 1) Then: GoTo ERROR_LABEL
Else
    ReDim UPPER_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NCOLUMNS
        UPPER_VECTOR(i, 1) = UPPER_RNG
    Next i
End If

'---------------------------------------------------------------------------------
If IsArray(FACTOR_RNG) = True Then
    FACTOR_VECTOR = FACTOR_RNG
    If UBound(FACTOR_VECTOR, 1) = 1 Then: _
        FACTOR_VECTOR = MATRIX_TRANSPOSE_FUNC(FACTOR_VECTOR)
    If NCOLUMNS <> UBound(FACTOR_VECTOR, 1) Then: GoTo ERROR_LABEL
Else
    ReDim FACTOR_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NCOLUMNS
        FACTOR_VECTOR(i, 1) = FACTOR_RNG
    Next i
End If
TRANSPOSE_FACTOR_VECTOR = MATRIX_TRANSPOSE_FUNC(FACTOR_VECTOR)
'---------------------------------------------------------------------------------
HIST_COVAR_MATRIX = MATRIX_COVARIANCE_FRAME3_FUNC(DATA_MATRIX, 0, 0)
If OUTPUT = 1 Then
    PORT_BAYES_STEIN_FRONTIER_FUNC = HIST_COVAR_MATRIX
    Exit Function
End If
'---------------------------------------------------------------------------------
INVERSE_COVAR_MATRIX = MATRIX_INVERSE_FUNC(HIST_COVAR_MATRIX, 2)
'---------------------------------------------------------------------------------
MVP_VOLATILITY_VAL = 1 / MMULT_FUNC(MMULT_FUNC(TRANSPOSE_FACTOR_VECTOR, INVERSE_COVAR_MATRIX, 70), FACTOR_VECTOR, 70)(1, 1)
If OUTPUT = 4 Then
    PORT_BAYES_STEIN_FRONTIER_FUNC = MVP_VOLATILITY_VAL
    Exit Function
End If

HIST_RETURNS_VECTOR = MATRIX_AVERAGE_RETURNS_FUNC(DATA_MATRIX, 0, 0)
If OUTPUT = 2 Then
    PORT_BAYES_STEIN_FRONTIER_FUNC = HIST_RETURNS_VECTOR
    Exit Function
End If

'---------------------------------------------------------------------------------
ReDim TEMP1_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
'---------------------------------------------------------------------------------
For j = 1 To NCOLUMNS
'---------------------------------------------------------------------------------
    For i = 1 To NCOLUMNS
        TEMP1_MATRIX(i, j) = INVERSE_COVAR_MATRIX(i, j) * MVP_VOLATILITY_VAL
    Next i
'---------------------------------------------------------------------------------
Next j
'---------------------------------------------------------------------------------
HIST_MVP_WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(MMULT_FUNC(TEMP1_MATRIX, FACTOR_VECTOR, 70))

If OUTPUT = 3 Then
    PORT_BAYES_STEIN_FRONTIER_FUNC = HIST_MVP_WEIGHTS_VECTOR
    Exit Function
End If
MVP_RETURN_VAL = MMULT_FUNC(HIST_MVP_WEIGHTS_VECTOR, HIST_RETURNS_VECTOR, 70)(1, 1)
    
If OUTPUT = 5 Then
    PORT_BAYES_STEIN_FRONTIER_FUNC = MVP_RETURN_VAL
    Exit Function
End If

'---------------------------------------------------------------------------------
ReDim TEMP1_MATRIX(1 To 1, 1 To NCOLUMNS)
'---------------------------------------------------------------------------------
For j = 1 To NCOLUMNS
'---------------------------------------------------------------------------------
    TEMP1_MATRIX(1, j) = HIST_RETURNS_VECTOR(j, 1) - MVP_RETURN_VAL * FACTOR_VECTOR(j, 1)
'---------------------------------------------------------------------------------
Next j
'---------------------------------------------------------------------------------

TEMP1_VAL = MMULT_FUNC(MMULT_FUNC(TEMP1_MATRIX, INVERSE_COVAR_MATRIX, 70), MATRIX_TRANSPOSE_FUNC(TEMP1_MATRIX), 70)(1, 1)

LAMBDA_VAL = (NCOLUMNS + 2) * (NROWS + 1) / (TEMP1_VAL * (NROWS - NCOLUMNS - 2))
PSI_VAL = LAMBDA_VAL / (NROWS + LAMBDA_VAL)

'---------------------------------------------------------------------------------
ReDim BAYES_STEIN_EXPECTED_RETURNS_VECTOR(1 To NCOLUMNS, 1 To 1) 'Per Asset
ReDim BAYES_STEIN_COVAR_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
'---------------------------------------------------------------------------------
TEMP1_VAL = MMULT_FUNC(FACTOR_VECTOR, TRANSPOSE_FACTOR_VECTOR, 70)(1, 1) / MVP_VOLATILITY_VAL ^ -1
TEMP2_VAL = LAMBDA_VAL / (NROWS * (NROWS + 1 + LAMBDA_VAL))
'---------------------------------------------------------------------------------
For j = 1 To NCOLUMNS
'---------------------------------------------------------------------------------
    BAYES_STEIN_EXPECTED_RETURNS_VECTOR(j, 1) = (1 - PSI_VAL) * HIST_RETURNS_VECTOR(j, 1) + PSI_VAL * MVP_RETURN_VAL * FACTOR_VECTOR(j, 1)
'---------------------------------------------------------------------------------
    For i = 1 To NCOLUMNS
'---------------------------------------------------------------------------------
        BAYES_STEIN_COVAR_MATRIX(i, j) = HIST_COVAR_MATRIX(i, j) * (1 + 1 / (NROWS + LAMBDA_VAL)) + TEMP1_VAL * TEMP2_VAL
'---------------------------------------------------------------------------------
    Next i
'---------------------------------------------------------------------------------
Next j
'---------------------------------------------------------------------------------

If OUTPUT = 6 Then
    PORT_BAYES_STEIN_FRONTIER_FUNC = BAYES_STEIN_EXPECTED_RETURNS_VECTOR
    Exit Function
End If

If OUTPUT = 7 Then
    PORT_BAYES_STEIN_FRONTIER_FUNC = BAYES_STEIN_COVAR_MATRIX
    Exit Function
End If

'---------------------------------------------------------------------------------
ReDim TEMP1_VECTOR(1 To 1, 1 To NCOLUMNS)
ReDim TEMP2_VECTOR(1 To 1, 1 To NCOLUMNS)
'---------------------------------------------------------------------------------
ReDim BAYES_STEIN_FRONTIER_MATRIX(1 To NCOLUMNS + 3, 1 To NSIZE)
ReDim MARKOWITZ_FRONTIER_MATRIX(1 To NCOLUMNS + 3, 1 To NSIZE)
'---------------------------------------------------------------------------------
For j = 1 To NSIZE
'---------------------------------------------------------------------------------
    TEMP1_MATRIX = PORT_SHARPE_WEIGHTS_OPTIMIZER_FUNC(BUDGET_VAL, _
                   TOLERANCE_VECTOR(j, 1), _
                   BAYES_STEIN_EXPECTED_RETURNS_VECTOR, _
                   BAYES_STEIN_COVAR_MATRIX, LOWER_VECTOR, UPPER_VECTOR)
                   'Bayes-Stein Frontier Portfolios
'---------------------------------------------------------------------------------
    TEMP2_MATRIX = PORT_SHARPE_WEIGHTS_OPTIMIZER_FUNC(BUDGET_VAL, _
                   TOLERANCE_VECTOR(j, 1), HIST_RETURNS_VECTOR, _
                   HIST_COVAR_MATRIX, _
                   LOWER_VECTOR, UPPER_VECTOR)
                   'Markowitz Frontier Portfolios
'---------------------------------------------------------------------------------
    TEMP1_VAL = 0
    TEMP2_VAL = 0
'---------------------------------------------------------------------------------
    For i = 1 To NCOLUMNS
'---------------------------------------------------------------------------------
        BAYES_STEIN_FRONTIER_MATRIX(i + 2, j) = TEMP1_MATRIX(i, 1)
        TEMP1_VECTOR(1, i) = BAYES_STEIN_FRONTIER_MATRIX(i + 2, j)
        TEMP1_VAL = TEMP1_VAL + TEMP1_VECTOR(1, i)
'---------------------------------------------------------------------------------
        MARKOWITZ_FRONTIER_MATRIX(i + 2, j) = TEMP2_MATRIX(i, 1)
        TEMP2_VECTOR(1, i) = MARKOWITZ_FRONTIER_MATRIX(i + 2, j)
        TEMP2_VAL = TEMP2_VAL + TEMP2_VECTOR(1, i)
'---------------------------------------------------------------------------------
    Next i
'---------------------------------------------------------------------------------
    BAYES_STEIN_FRONTIER_MATRIX(1, j) = MMULT_FUNC(TEMP1_VECTOR, _
            BAYES_STEIN_EXPECTED_RETURNS_VECTOR, _
            70)(1, 1) * COUNT_BASIS 'Expected Return BS
'---------------------------------------------------------------------------------
    BAYES_STEIN_FRONTIER_MATRIX(2, j) = MMULT_FUNC(MMULT_FUNC(TEMP1_VECTOR, _
                                        BAYES_STEIN_COVAR_MATRIX, 70), _
                                        MATRIX_TRANSPOSE_FUNC(TEMP1_VECTOR), 70)(1, 1) _
                                        * Sqr(COUNT_BASIS) 'Volatility BS
'---------------------------------------------------------------------------------
    BAYES_STEIN_FRONTIER_MATRIX(NCOLUMNS + 3, j) = TEMP1_VAL 'Budget
'---------------------------------------------------------------------------------
    MARKOWITZ_FRONTIER_MATRIX(1, j) = _
        MMULT_FUNC(TEMP2_VECTOR, HIST_RETURNS_VECTOR, 70)(1, 1) * COUNT_BASIS
        'Expected Return
'---------------------------------------------------------------------------------
    MARKOWITZ_FRONTIER_MATRIX(2, j) = _
        MMULT_FUNC(MMULT_FUNC(TEMP2_VECTOR, HIST_COVAR_MATRIX, 70), _
        MATRIX_TRANSPOSE_FUNC(TEMP2_VECTOR), 70)(1, 1) * Sqr(COUNT_BASIS) 'Volatility
'---------------------------------------------------------------------------------
    MARKOWITZ_FRONTIER_MATRIX(NCOLUMNS + 3, j) = TEMP2_VAL 'Budget
'---------------------------------------------------------------------------------
Next j
'---------------------------------------------------------------------------------
    
If OUTPUT = 8 Then
    PORT_BAYES_STEIN_FRONTIER_FUNC = BAYES_STEIN_FRONTIER_MATRIX
    Exit Function
End If

If OUTPUT = 9 Then
    PORT_BAYES_STEIN_FRONTIER_FUNC = MARKOWITZ_FRONTIER_MATRIX
    Exit Function
End If
    
ReDim TEMP_GROUP(1 To 9)
TEMP_GROUP(1) = HIST_COVAR_MATRIX
TEMP_GROUP(2) = HIST_RETURNS_VECTOR
TEMP_GROUP(3) = HIST_MVP_WEIGHTS_VECTOR
TEMP_GROUP(4) = MVP_VOLATILITY_VAL
TEMP_GROUP(5) = MVP_RETURN_VAL
TEMP_GROUP(6) = BAYES_STEIN_EXPECTED_RETURNS_VECTOR
TEMP_GROUP(7) = BAYES_STEIN_COVAR_MATRIX
TEMP_GROUP(8) = BAYES_STEIN_FRONTIER_MATRIX
TEMP_GROUP(9) = MARKOWITZ_FRONTIER_MATRIX

PORT_BAYES_STEIN_FRONTIER_FUNC = TEMP_GROUP

Exit Function
ERROR_LABEL:
PORT_BAYES_STEIN_FRONTIER_FUNC = Err.number
End Function
