Attribute VB_Name = "FINAN_PORT_MOMENTS_SHARPE_LIBR"

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_SHARPE_RATIO_FUNC
'DESCRIPTION   : RETURNS A VECTOR WITH SHARPE RATIOS
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_SHARPE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_SHARPE_RATIO_FUNC(ByRef BENCH_RNG As Variant, _
ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double

Dim ASSET_VECTOR As Variant
Dim BENCH_VECTOR As Variant

Dim MEAN_ASSET_VECTOR As Variant
Dim MEAN_BENCH_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ASSET_VECTOR = DATA_RNG
If UBound(ASSET_VECTOR, 1) = 1 Then: _
ASSET_VECTOR = MATRIX_TRANSPOSE_FUNC(ASSET_VECTOR)

BENCH_VECTOR = BENCH_RNG
If UBound(BENCH_VECTOR, 1) = 1 Then: _
BENCH_VECTOR = MATRIX_TRANSPOSE_FUNC(BENCH_VECTOR)

If UBound(ASSET_VECTOR, 1) <> UBound(BENCH_VECTOR, 1) Then: GoTo ERROR_LABEL

If DATA_TYPE <> 0 Then
    ASSET_VECTOR = MATRIX_PERCENT_FUNC(ASSET_VECTOR, LOG_SCALE)
    BENCH_VECTOR = MATRIX_PERCENT_FUNC(BENCH_VECTOR, LOG_SCALE)
End If

NROWS = UBound(ASSET_VECTOR, 1)
NCOLUMNS = UBound(ASSET_VECTOR, 2)

MEAN_ASSET_VECTOR = MATRIX_MEAN_FUNC(ASSET_VECTOR)
MEAN_BENCH_VECTOR = MATRIX_MEAN_FUNC(BENCH_VECTOR)
        
ReDim TEMP_MATRIX(1 To 1, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    TEMP_SUM = 0
    For i = 1 To NROWS
        TEMP_SUM = TEMP_SUM + ((ASSET_VECTOR(i, j) - MEAN_ASSET_VECTOR(1, j)) ^ 2)
    Next i
    TEMP_SUM = TEMP_SUM / (NROWS) ^ 2
    TEMP_MATRIX(1, j) = (MEAN_ASSET_VECTOR(1, j) - MEAN_BENCH_VECTOR(1, 1)) / TEMP_SUM
Next j

PORT_SHARPE_RATIO_FUNC = TEMP_MATRIX
    
Exit Function
ERROR_LABEL:
PORT_SHARPE_RATIO_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_SHARPE_RATIO_AFTER_FEES_FUNC
'DESCRIPTION   : Help the manager to set up an x-leverage version of the base
'program, at the request of a client
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_SHARPE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_SHARPE_RATIO_AFTER_FEES_FUNC(ByVal MGMT_FEES_RNG As Variant, _
ByVal PERFORMANCE_FEES_RNG As Variant, _
ByVal CASH_RATE_RNG As Variant, _
ByVal EXPECTED_RETURN_RNG As Variant, _
ByVal VOLATILITY_RNG As Variant, _
Optional ByVal LEVERAGE_MULT_RNG As Variant = 1)

'MARKET_NEUTRAL_PORTFOLIO STRATEGY
'EXPECTED_RETURN --> AFTER FEES

Dim j As Long
Dim NCOLUMNS As Long

Dim MGMT_FEE As Double
Dim PERFORMANCE_FEE As Double
Dim CASH_RATE As Double
Dim EXPECTED_RETURN As Double
Dim VOLATILITY As Double
Dim LEVERAGE_MULT As Double

Dim TEMP_MATRIX As Variant

Dim MGMT_FEES_VECTOR As Variant
Dim PERFORMANCE_FEES_VECTOR As Variant
Dim CASH_RATE_VECTOR As Variant
Dim EXPECTED_RETURN_VECTOR As Variant
Dim VOLATILITY_VECTOR As Variant
Dim LEVERAGE_MULT_VECTOR As Variant

'--------------------------------------------------------------------------
On Error GoTo ERROR_LABEL
'--------------------------------------------------------------------------

MGMT_FEES_VECTOR = MGMT_FEES_RNG
If UBound(MGMT_FEES_VECTOR, 2) = 1 Then
    MGMT_FEES_VECTOR = MATRIX_TRANSPOSE_FUNC(MGMT_FEES_VECTOR)
End If
NCOLUMNS = UBound(MGMT_FEES_VECTOR, 2)

PERFORMANCE_FEES_VECTOR = PERFORMANCE_FEES_RNG
If UBound(PERFORMANCE_FEES_VECTOR, 2) = 1 Then
    PERFORMANCE_FEES_VECTOR = MATRIX_TRANSPOSE_FUNC(PERFORMANCE_FEES_VECTOR)
End If

CASH_RATE_VECTOR = CASH_RATE_RNG
If UBound(CASH_RATE_VECTOR, 2) = 1 Then
    CASH_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(CASH_RATE_VECTOR)
End If

EXPECTED_RETURN_VECTOR = EXPECTED_RETURN_RNG
If UBound(EXPECTED_RETURN_VECTOR, 2) = 1 Then
    EXPECTED_RETURN_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPECTED_RETURN_VECTOR)
End If

VOLATILITY_VECTOR = VOLATILITY_RNG
If UBound(VOLATILITY_VECTOR, 2) = 1 Then
    VOLATILITY_VECTOR = MATRIX_TRANSPOSE_FUNC(VOLATILITY_VECTOR)
End If

'------------------------------------------------------------------------
If IsArray(LEVERAGE_MULT_RNG) = True Then
'------------------------------------------------------------------------
    LEVERAGE_MULT_VECTOR = LEVERAGE_MULT_RNG
    If UBound(LEVERAGE_MULT_VECTOR, 2) = 1 Then
        LEVERAGE_MULT_VECTOR = MATRIX_TRANSPOSE_FUNC(LEVERAGE_MULT_VECTOR)
    End If
'------------------------------------------------------------------------
Else
'------------------------------------------------------------------------
    ReDim LEVERAGE_MULT_VECTOR(1 To 1, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS
        LEVERAGE_MULT_VECTOR(1, j) = LEVERAGE_MULT_RNG
    Next j
'------------------------------------------------------------------------
End If
'------------------------------------------------------------------------

'------------------------------------------------------------------------
ReDim TEMP_MATRIX(1 To 8, 1 To NCOLUMNS + 1)
'------------------------------------------------------------------------

TEMP_MATRIX(1, 1) = "SHARPE RATIO FOR THE END-INVESTOR"
TEMP_MATRIX(2, 1) = "CURRENT PORTFOLIO - RETURNS BEFORE FEES"
TEMP_MATRIX(3, 1) = "CURRENT PORTFOLIO - ACTIVE RETURN BEFORE FEES"
TEMP_MATRIX(4, 1) = "CURRENT PORTFOLIO - SHARPE RATIO BEFORE FEES"
TEMP_MATRIX(5, 1) = "RETURN BEFORE FEES - LEVERAGE PORTFOLIO"
TEMP_MATRIX(6, 1) = "RETURN AFTER FEES - LEVERAGE PORTFOLIO"
TEMP_MATRIX(7, 1) = "VOLATILITY - LEVERAGE PORTFOLIO"
TEMP_MATRIX(8, 1) = "SHARPE RATIO AFTER FEES - LEVERAGE PORTFOLIO"

'-------------------------------------------------------------------------
For j = 1 To NCOLUMNS
'-------------------------------------------------------------------------
    MGMT_FEE = MGMT_FEES_VECTOR(1, j)
    PERFORMANCE_FEE = PERFORMANCE_FEES_VECTOR(1, j)
    CASH_RATE = CASH_RATE_VECTOR(1, j)
    EXPECTED_RETURN = EXPECTED_RETURN_VECTOR(1, j)
    VOLATILITY = VOLATILITY_VECTOR(1, j)
    LEVERAGE_MULT = LEVERAGE_MULT_VECTOR(1, j)
'-------------------------------------------------------------------------
    TEMP_MATRIX(1, j + 1) = (EXPECTED_RETURN - CASH_RATE) / VOLATILITY
    TEMP_MATRIX(2, j + 1) = EXPECTED_RETURN / (1 - PERFORMANCE_FEE) + MGMT_FEE
    TEMP_MATRIX(3, j + 1) = TEMP_MATRIX(2, j + 1) - CASH_RATE
    TEMP_MATRIX(4, j + 1) = TEMP_MATRIX(3, j + 1) / VOLATILITY
    TEMP_MATRIX(5, j + 1) = CASH_RATE + LEVERAGE_MULT * TEMP_MATRIX(3, j + 1)
    TEMP_MATRIX(6, j + 1) = (TEMP_MATRIX(5, j + 1) - MGMT_FEE) * (1 - PERFORMANCE_FEE)
    TEMP_MATRIX(7, j + 1) = LEVERAGE_MULT * VOLATILITY
    TEMP_MATRIX(8, j + 1) = (TEMP_MATRIX(6, j + 1) - CASH_RATE) / TEMP_MATRIX(7, j + 1)
'--------------------------------------------------------------------------
Next j
'--------------------------------------------------------------------------
PORT_SHARPE_RATIO_AFTER_FEES_FUNC = TEMP_MATRIX
'--------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
PORT_SHARPE_RATIO_AFTER_FEES_FUNC = Err.number
End Function
