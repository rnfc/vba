Attribute VB_Name = "FINAN_FI_MBS_SIMUL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_MBS_VALUATION_FUNC
'DESCRIPTION   : Hull-White MBS Simmulation Valuation
'LIBRARY       : MBS
'GROUP         : SIMULATION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_MBS_VALUATION_FUNC(ByVal SHORT_RATE As Double, _
ByVal KAPPA As Double, _
ByVal VOLATILITY As Double, _
ByRef MEAN_YIELD_RNG As Variant, _
ByRef RANDOM_RATE_RNG As Variant, _
ByVal INITIAL_ASSET_VALUE As Double, _
ByVal ASSET_VOLATILITY As Double, _
ByVal ASSET_RECOVERY As Double, _
ByRef CDR_RNG As Variant, _
ByRef AMORT_RNG As Variant, _
ByRef RANDOM_ASSET_RNG As Variant, _
Optional ByVal DELTA_TENOR As Double = 0.5)

'INITIAL_ASSET_PRICE = AFTER INITIAL DEPRECIATION
'SHORT_RATE = LOG & ANNUALIZED
'MEAN_YIELD_RNG = SIMULATED_AVG_YIELD_PERIOD (Excluding First
    'Period [DELTA_TENOR])
'SHORT_RATE = (Log-Linear & Annualized)

Dim i As Long
Dim NROWS As Long
Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double

Dim TEMP_VECTOR As Variant

Dim YIELD_PERIOD_VECTOR As Variant
Dim RATE_RND_VECTOR As Variant
Dim ASSET_RND_VECTOR As Variant
Dim AMORT_VECTOR As Variant 'Depreciation Vector
Dim CDR_VECTOR As Variant

On Error GoTo ERROR_LABEL

'---------------------------------------------------------------
YIELD_PERIOD_VECTOR = MEAN_YIELD_RNG
If NROWS = 1 Then
    YIELD_PERIOD_VECTOR = MATRIX_TRANSPOSE_FUNC(YIELD_PERIOD_VECTOR)
End If
'---------------------------------------------------------------
RATE_RND_VECTOR = RANDOM_RATE_RNG
If UBound(RATE_RND_VECTOR, 1) = 1 Then
    RATE_RND_VECTOR = MATRIX_TRANSPOSE_FUNC(RATE_RND_VECTOR)
End If
'---------------------------------------------------------------
ASSET_RND_VECTOR = RANDOM_ASSET_RNG
If UBound(ASSET_RND_VECTOR, 1) = 1 Then
    ASSET_RND_VECTOR = MATRIX_TRANSPOSE_FUNC(ASSET_RND_VECTOR)
End If
'---------------------------------------------------------------
AMORT_VECTOR = AMORT_RNG
If UBound(AMORT_VECTOR, 1) = 1 Then
    AMORT_VECTOR = MATRIX_TRANSPOSE_FUNC(AMORT_VECTOR)
End If
'---------------------------------------------------------------
CDR_VECTOR = CDR_RNG
If UBound(CDR_VECTOR, 1) = 1 Then
    CDR_VECTOR = MATRIX_TRANSPOSE_FUNC(CDR_VECTOR)
End If
'---------------------------------------------------------------

If UBound(ASSET_RND_VECTOR, 1) <> UBound(AMORT_VECTOR, 1) Then: _
GoTo ERROR_LABEL

If UBound(ASSET_RND_VECTOR, 1) <> UBound(CDR_VECTOR, 1) Then: _
GoTo ERROR_LABEL

If (UBound(ASSET_RND_VECTOR, 1) - 1) <> UBound(RATE_RND_VECTOR, 1) Then: _
GoTo ERROR_LABEL

NROWS = UBound(YIELD_PERIOD_VECTOR, 1)

ReDim TEMP_VECTOR(0 To NROWS + 1, 1 To 13)

ATEMP_SUM = DELTA_TENOR
TEMP_VECTOR(0, 1) = "TENOR"
TEMP_VECTOR(1, 1) = ATEMP_SUM _
'-----> MATURITY OF THE SHORT_RATE ANNUALIZED

TEMP_VECTOR(0, 2) = "SIM YIELD PERIOD"
TEMP_VECTOR(1, 2) = ""

TEMP_VECTOR(0, 3) = "MEAN YIELD"
TEMP_VECTOR(1, 3) = ""

TEMP_VECTOR(0, 4) = "VARIANCE YIELD"
TEMP_VECTOR(1, 4) = ""

TEMP_VECTOR(0, 5) = "HW_YIELD_PERIOD"
TEMP_VECTOR(1, 5) = SHORT_RATE * DELTA_TENOR

TEMP_VECTOR(0, 6) = "ESTIMATED ZERO"
TEMP_VECTOR(1, 6) = Exp(-1 * DELTA_TENOR * SHORT_RATE)

TEMP_VECTOR(0, 7) = "ASSET DRIFT"
TEMP_VECTOR(1, 7) = Log(1 + SHORT_RATE) _
- Log(1 + AMORT_VECTOR(1, 1))

TEMP_VECTOR(0, 8) = "RISK FREE FACTOR"
TEMP_VECTOR(1, 8) = (TEMP_VECTOR(1, 7) - 1 / 2 * _
(ASSET_VOLATILITY ^ 2)) * DELTA_TENOR

TEMP_VECTOR(0, 9) = "ASSET VOLATILITY"
TEMP_VECTOR(1, 9) = (ASSET_VOLATILITY * DELTA_TENOR ^ 0.5 * ASSET_RND_VECTOR(1, 1))

'-------------------------------------------------------

TEMP_VECTOR(0, 10) = "ASSET DISC. FACTOR"
TEMP_VECTOR(1, 10) = Exp(TEMP_VECTOR(1, 8) + TEMP_VECTOR(1, 9))

TEMP_VECTOR(0, 11) = "ASSET PRICE"
TEMP_VECTOR(1, 11) = INITIAL_ASSET_VALUE * TEMP_VECTOR(1, 10)

TEMP_VECTOR(0, 12) = "UNRECOVERED ASSET"
TEMP_VECTOR(1, 12) = TEMP_VECTOR(1, 11)

TEMP_VECTOR(0, 13) = "RECOVERY"
TEMP_VECTOR(1, 13) = TEMP_VECTOR(1, 12) * CDR_VECTOR(1, 1) * ASSET_RECOVERY

BTEMP_SUM = TEMP_VECTOR(1, 5)

For i = 2 To NROWS + 1
    
    ATEMP_SUM = ATEMP_SUM + DELTA_TENOR
    TEMP_VECTOR(i, 1) = ATEMP_SUM
    TEMP_VECTOR(i, 2) = YIELD_PERIOD_VECTOR(i - 1, 1)

    TEMP_VECTOR(i, 3) = HW_MBS_SPOT_FUNC(TEMP_VECTOR(i - 1, 5), _
    TEMP_VECTOR(i, 2), KAPPA, DELTA_TENOR)

    TEMP_VECTOR(i, 4) = HW_MBS_VAR_FUNC(VOLATILITY, KAPPA, DELTA_TENOR)
    
    TEMP_VECTOR(i, 5) = HW_MBS_DISC_FUNC(TEMP_VECTOR(i - 1, 5), _
    TEMP_VECTOR(i, 2), KAPPA, VOLATILITY, DELTA_TENOR, _
    RATE_RND_VECTOR(i - 1, 1))
    
    BTEMP_SUM = BTEMP_SUM + TEMP_VECTOR(i, 5)
    
    TEMP_VECTOR(i, 6) = Exp(-1 * BTEMP_SUM)
    
    TEMP_VECTOR(i, 7) = Log(1 + (TEMP_VECTOR(i, 5) / DELTA_TENOR)) _
    - Log(1 + AMORT_VECTOR(i, 1))

    TEMP_VECTOR(i, 8) = (TEMP_VECTOR(i, 7) - 1 / 2 * _
    (ASSET_VOLATILITY ^ 2)) * DELTA_TENOR

    TEMP_VECTOR(i, 9) = (ASSET_VOLATILITY * DELTA_TENOR ^ 0.5 * _
    ASSET_RND_VECTOR(i, 1))

    TEMP_VECTOR(i, 10) = Exp(TEMP_VECTOR(i, 8) + TEMP_VECTOR(i, 9))
    
    TEMP_VECTOR(i, 11) = TEMP_VECTOR(i - 1, 11) * TEMP_VECTOR(i, 10)
    
    TEMP_VECTOR(i, 12) = (TEMP_VECTOR(i - 1, 12) - (TEMP_VECTOR(i - 1, 13) / _
    ASSET_RECOVERY)) * TEMP_VECTOR(i, 10)
    
    TEMP_VECTOR(i, 13) = TEMP_VECTOR(i, 12) * CDR_VECTOR(i, 1) * ASSET_RECOVERY

Next i


HW_MBS_VALUATION_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
HW_MBS_VALUATION_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_MBS_SIMULATION_FUNC
'DESCRIPTION   : Fitting the Term-Structure of Zero Bond Prices in the Hull-White Model

'For the purpose of valuating MBSs and CMOs, any arbitrage-free model of the term
'structure of interest rates can be used. Equilibrium interest rate models are
'based on the assumption that bond prices, and yields, are determined by the
'market’s assessment of the evolution of the short-term interest rate. For the
'following models, the short rate is assumed to follow a diffusion (a continuous
'time stochastic) process.

'The Hull-White model have some form of reversion, reverting the generated
'interest rate paths to some "normal" level.

'Without reversion, interest rates can obtain unreasonably high and low levels.
'Volatility, over time, would theoretically approach infinity. Similarly, a large
'percentage assumption of volatility will result in greater fluctuations in yield.
'This results in a greater probability of the opportunity to refinance. The
'outcome of refinancing is a greater value attributed to the implied call option,
'and a higher resulting option cost. Finally, the speed of reversion, ultimately
'affects the shape of the yield curve.  In fact, if it is high the yield curve
'quickly tends toward the long-run yield rate.

'LIBRARY       : MBS
'GROUP         : SIMULATION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_MBS_SIMULATION_FUNC(ByVal nLOOPS As Variant, _
ByVal SETTLEMENT As Date, _
ByVal YIELD_RNG As Variant, _
ByVal KAPPA As Double, _
ByVal VOLATILITY As Double, _
ByVal SHIFT_YIELD As Double, _
Optional ByVal DELTA_TENOR As Double = 0.5, _
Optional ByVal MIN_RATE As Double = 0.01, _
Optional ByVal MAX_RATE As Double = 0.2, _
Optional ByVal MOMENT_FLAG As Boolean = False, _
Optional ByVal CORREL_RNG As Variant, _
Optional ByVal COUNT_BASIS As Integer = 0, _
Optional ByVal HOLIDAYS_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 1)

'KAPPA: The larger the positive number the stronger the mean reversion.

Dim i As Long
Dim k As Long
Dim NCOLUMNS As Long

Dim DELTA_RATE As Double
Dim SHORT_RATE As Double
Dim MID_POINT As Double

Dim FIRST_VAL As Double '
Dim SECOND_VAL As Double

Dim TEMP_SUM As Double
Dim TEMP_VALUE As Double

Dim TEMP_MATRIX As Variant
Dim SIMULATION_VECTOR As Variant

Dim TEMP_VECTOR As Variant
Dim DISCOUNT_VECTOR As Variant
Dim NORMAL_RANDOM_MATRIX As Variant

On Error GoTo ERROR_LABEL


TEMP_MATRIX = HW_MBS_YIELD_TABLE_FUNC(SETTLEMENT, YIELD_RNG, SHIFT_YIELD, _
              DELTA_TENOR, COUNT_BASIS, HOLIDAYS_RNG)

DISCOUNT_VECTOR = MATRIX_TRANSPOSE_FUNC(MATRIX_GET_COLUMN_FUNC(TEMP_MATRIX, 7, 1))
NCOLUMNS = UBound(DISCOUNT_VECTOR, 2)

SHORT_RATE = Log(1 / (DISCOUNT_VECTOR(1, 1))) / DELTA_TENOR
'Short Rate (Log-Linear & Annualized)

ReDim RATE_MATRIX(1 To nLOOPS, 1 To NCOLUMNS)
ReDim ZERO_MATRIX(1 To nLOOPS, 1 To NCOLUMNS)
ReDim SIMULATION_VECTOR(1 To 2, 1 To NCOLUMNS)

DELTA_RATE = 0.00000000000001
MID_POINT = (MAX_RATE + MIN_RATE) / 2
NORMAL_RANDOM_MATRIX = MULTI_NORMAL_RANDOM_MATRIX_FUNC(1, nLOOPS, NCOLUMNS - 1, 0, 1, True, MOMENT_FLAG, CORREL_RNG, 0)

For i = 1 To NCOLUMNS
    If i = 1 Then
        Call HW_MBS_SIMULATION_OBJ_FUNC(TEMP_VALUE, nLOOPS, i, DELTA_RATE, _
        SHORT_RATE, DELTA_TENOR, KAPPA, VOLATILITY, RATE_MATRIX, ZERO_MATRIX, NORMAL_RANDOM_MATRIX)
        
        DELTA_RATE = TEMP_VALUE
        
        SIMULATION_VECTOR(1, i) = Exp(-i * DELTA_TENOR * SHORT_RATE)
        SIMULATION_VECTOR(2, i) = ""
    Else

        k = 0
        DELTA_RATE = 0
        
        Do Until k = NCOLUMNS - 1 'Bisec searching algorithm
    '---------------------------------------------------------------------------------
    'Bisec method, unlike the Newton-Ralphson method, does not require the
    'main function in question. However, it does requires an initial min and max value.
    'These values can determine the way the search is converged. The major challenge
    'to using this method is that the first differential (first derivative)
    'of the equation is required as an input for the search procedure.
    'Sometimes, it may be difficult or impossible to derive that.
    '---------------------------------------------------------------------------------
             k = k + 1
             
             Call HW_MBS_SIMULATION_OBJ_FUNC(FIRST_VAL, nLOOPS, i, _
             DELTA_RATE, SHORT_RATE, DELTA_TENOR, KAPPA, VOLATILITY, RATE_MATRIX, _
             ZERO_MATRIX, NORMAL_RANDOM_MATRIX)
             
             Call HW_MBS_SIMULATION_OBJ_FUNC(SECOND_VAL, nLOOPS, i, _
             DELTA_RATE + MID_POINT, SHORT_RATE, DELTA_TENOR, KAPPA, _
             VOLATILITY, RATE_MATRIX, ZERO_MATRIX, NORMAL_RANDOM_MATRIX)
             
             TEMP_VALUE = DISCOUNT_VECTOR(1, i)
             
             DELTA_RATE = DELTA_RATE + (TEMP_VALUE - FIRST_VAL) / _
             ((SECOND_VAL - FIRST_VAL) / MID_POINT)
             
             MID_POINT = MID_POINT / 2
             
             Call HW_MBS_SIMULATION_OBJ_FUNC(TEMP_VALUE, nLOOPS, i, DELTA_RATE, _
                    SHORT_RATE, DELTA_TENOR, KAPPA, VOLATILITY, RATE_MATRIX, ZERO_MATRIX, _
                    NORMAL_RANDOM_MATRIX)
             If (TEMP_VALUE - DISCOUNT_VECTOR(1, i)) ^ 2 < 0.00000000000001 Then Exit Do
        Loop
        SIMULATION_VECTOR(1, i) = TEMP_VALUE 'Calculated Average Zeros
        SIMULATION_VECTOR(2, i) = DELTA_RATE 'Simulated Mean Rate
    End If
    
Next i

Select Case OUTPUT
Case 0
    ReDim TEMP_VECTOR(0 To UBound(SIMULATION_VECTOR, 2), 1 To 4)
    TEMP_SUM = DELTA_TENOR
    TEMP_VECTOR(0, 1) = "MATURITY"
    TEMP_VECTOR(1, 1) = TEMP_SUM
    TEMP_VECTOR(0, 2) = "OBSERVED ZERO PRICES"
    TEMP_VECTOR(1, 2) = DISCOUNT_VECTOR(1, 1)
    TEMP_VECTOR(0, 3) = "SIMULATED AVERAGE ZEROS"
    TEMP_VECTOR(1, 3) = ""
    TEMP_VECTOR(0, 4) = "DIFFERENCE SQUARED"
    TEMP_VECTOR(1, 4) = ""
    For i = 2 To UBound(SIMULATION_VECTOR, 2)
        TEMP_SUM = TEMP_SUM + DELTA_TENOR
        TEMP_VECTOR(i, 1) = TEMP_SUM 'Maturity (in date
        TEMP_VECTOR(i, 2) = DISCOUNT_VECTOR(1, i) 'Observed Zero Prices
        TEMP_VECTOR(i, 3) = SIMULATION_VECTOR(1, i) 'Calculated Average Zeros
        TEMP_VECTOR(i, 4) = Abs(TEMP_VECTOR(i, 2) - _
        TEMP_VECTOR(i, 3)) ^ 2 'Difference Squared
    Next i
    HW_MBS_SIMULATION_FUNC = TEMP_VECTOR
Case Else
    HW_MBS_SIMULATION_FUNC = MATRIX_TRANSPOSE_FUNC(SIMULATION_VECTOR)
End Select

Exit Function
ERROR_LABEL:
HW_MBS_SIMULATION_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_MBS_SIMULATION_OBJ_FUNC
'DESCRIPTION   : Hull-White MBS Simmulation Calibration
'LIBRARY       : MBS
'GROUP         : SIMULATION
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_MBS_SIMULATION_OBJ_FUNC(ByRef RESULT_VALUE As Double, _
ByVal nLOOPS As Long, _
ByVal i As Long, _
ByVal DELTA_RATE As Double, _
ByVal SHORT_RATE As Double, _
ByVal DELTA_TENOR As Double, _
ByVal KAPPA As Double, _
ByVal VOLATILITY As Double, _
ByRef RATE_MATRIX As Variant, _
ByRef ZERO_MATRIX As Variant, _
ByRef NORMAL_RANDOM_MATRIX As Variant)

Dim j As Long
Dim k As Long

Dim TEMP_SUM As Double
Dim TEMP_RATE As Double
Dim CUMUL_ZERO As Double

On Error GoTo ERROR_LABEL

HW_MBS_SIMULATION_OBJ_FUNC = False

TEMP_RATE = 0
TEMP_SUM = 0
CUMUL_ZERO = 0

If i = 1 Then
    For j = 1 To nLOOPS
        RATE_MATRIX(j, 1) = SHORT_RATE * DELTA_TENOR
        ZERO_MATRIX(j, 1) = Exp(-i * DELTA_TENOR * SHORT_RATE)
    Next j
Else
    For j = 1 To nLOOPS
        TEMP_RATE = HW_MBS_DISC_FUNC(RATE_MATRIX(j, i - 1), _
            DELTA_RATE, KAPPA, VOLATILITY, DELTA_TENOR, _
            NORMAL_RANDOM_MATRIX(j, i - 1))
                      
        TEMP_SUM = 0
        For k = 2 To (i - 1) 'Previous Zero Rates
            TEMP_SUM = TEMP_SUM + RATE_MATRIX(j, k)
        Next k
                       
        RATE_MATRIX(j, i) = TEMP_RATE
        ZERO_MATRIX(j, i) = Exp(-(RATE_MATRIX(j, 1) + TEMP_RATE + TEMP_SUM))
                       
        CUMUL_ZERO = CUMUL_ZERO + ZERO_MATRIX(j, i)
    Next j
End If
        
RESULT_VALUE = CUMUL_ZERO / nLOOPS

HW_MBS_SIMULATION_OBJ_FUNC = True

Exit Function
ERROR_LABEL:
HW_MBS_SIMULATION_OBJ_FUNC = False
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_MBS_YIELD_TABLE_FUNC
'DESCRIPTION   : Linear Rates to Log-Linear Rates Table
'LIBRARY       : MBS
'GROUP         : SIMULATION
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************


Function HW_MBS_YIELD_TABLE_FUNC(ByVal SETTLEMENT As Date, _
ByRef YIELD_RNG As Variant, _
Optional ByVal SHIFT_YIELD As Double = 0, _
Optional ByVal DELTA_TENOR As Double = 0.5, _
Optional ByVal COUNT_BASIS As Integer = 0, _
Optional ByRef HOLIDAYS_RNG As Variant)

Dim i As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim YIELD_VECTOR As Variant 'Linear Yields
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

YIELD_VECTOR = YIELD_RNG
If UBound(YIELD_VECTOR, 1) = 1 Then
    YIELD_VECTOR = MATRIX_TRANSPOSE_FUNC(YIELD_VECTOR)
End If
NROWS = UBound(YIELD_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 7)

TEMP_MATRIX(0, 1) = "PERIODS"
TEMP_MATRIX(0, 2) = "TENOR"
TEMP_MATRIX(0, 3) = "MATURITY"
TEMP_MATRIX(0, 4) = "ACTUAL_YIELD"
TEMP_MATRIX(0, 5) = "LOG_YIELD"
TEMP_MATRIX(0, 6) = "YIELD_PERIOD"
TEMP_MATRIX(0, 7) = "ZERO_PRICE"

TEMP_SUM = DELTA_TENOR

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = TEMP_SUM
    TEMP_MATRIX(i, 3) = _
    WORKMONTH_FUNC(SETTLEMENT, 12 * TEMP_MATRIX(i, 2), HOLIDAYS_RNG)
    TEMP_MATRIX(i, 4) = YIELD_VECTOR(i, 1)
    TEMP_MATRIX(i, 5) = Log(1 + TEMP_MATRIX(i, 4) + SHIFT_YIELD)
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 5) * DELTA_TENOR
    TEMP_MATRIX(i, 7) = Exp(-1 * TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 1))
    
    TEMP_SUM = TEMP_SUM + DELTA_TENOR
Next i

HW_MBS_YIELD_TABLE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
HW_MBS_YIELD_TABLE_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_MBS_RATE_FUNC
'DESCRIPTION   : Convert from Linear Rates to Log Rates
'LIBRARY       : MBS
'GROUP         : SIMULATION
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_MBS_RATE_FUNC(ByVal RATE As Double, _
Optional ByVal TENOR As Double = 1, _
Optional ByVal OUTPUT As Integer = 0)

On Error GoTo ERROR_LABEL

Select Case OUTPUT
    Case 0
       HW_MBS_RATE_FUNC = Log(1 + RATE) * TENOR 'TENOR Could Also be
       'DELTA_STEP FOR EACH PERIOD :)
    Case 1 'From Log Rates to Log Discount Factors: Observed Zero Prices
       HW_MBS_RATE_FUNC = Exp(-RATE * TENOR) _
        'TENOR --> COULD ALSO REPRESENT ANY TYPE OF
        'PERIODS :)
End Select

Exit Function
ERROR_LABEL:
HW_MBS_RATE_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_MBS_DISC_FUNC
'DESCRIPTION   : Discrete Hull White Calculator
'LIBRARY       : MBS
'GROUP         : SIMULATION
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_MBS_DISC_FUNC(ByVal SPOT_RATE As Double, _
ByVal DELTA_RATE As Double, _
ByVal KAPPA As Double, _
ByVal VOLATILITY As Double, _
ByVal DELTA_TENOR As Double, _
ByVal NORMAL_RANDOM_VAL As Double)

On Error GoTo ERROR_LABEL

HW_MBS_DISC_FUNC = HW_MBS_SPOT_FUNC(SPOT_RATE, DELTA_RATE, KAPPA, DELTA_TENOR) _
+ ((HW_MBS_VAR_FUNC(VOLATILITY, KAPPA, DELTA_TENOR) ^ 0.5) * NORMAL_RANDOM_VAL)

Exit Function
ERROR_LABEL:
HW_MBS_DISC_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_MBS_SPOT_FUNC
'DESCRIPTION   : Discrete Hull White Spot Rate
'LIBRARY       : MBS
'GROUP         : SIMULATION
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_MBS_SPOT_FUNC(ByVal SPOT_RATE As Double, _
ByVal DELTA_RATE As Double, _
ByVal KAPPA As Double, _
ByVal DELTA_TENOR As Double)

On Error GoTo ERROR_LABEL

    HW_MBS_SPOT_FUNC = DELTA_RATE * (1 - Exp(-KAPPA * DELTA_TENOR)) + (SPOT_RATE) * _
    Exp(-KAPPA * DELTA_TENOR)

Exit Function
ERROR_LABEL:
HW_MBS_SPOT_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : HW_MBS_VAR_FUNC
'DESCRIPTION   : Discrete Hull White Spot Rate Variance
'LIBRARY       : MBS
'GROUP         : SIMULATION
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function HW_MBS_VAR_FUNC(ByVal VOLATILITY As Double, _
ByVal KAPPA As Double, _
ByVal DELTA_TENOR As Double)

On Error GoTo ERROR_LABEL

HW_MBS_VAR_FUNC = (VOLATILITY ^ 2 / (2 * KAPPA)) * (1 - Exp(-2 * KAPPA * DELTA_TENOR))

Exit Function
ERROR_LABEL:
HW_MBS_VAR_FUNC = Err.number
End Function
