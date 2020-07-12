Attribute VB_Name = "FINAN_FI_BOND_TERM_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : TERM_STRUCTURE_PARTIAL_ADJUSTMENT_FUNC
'DESCRIPTION   : Short-run and long-run impacts of a change in an exogenous
'variable on an endogenous variable when lagged endogenous
'variable is included.
'LIBRARY       : FIXED_INCOME
'GROUP         : BOND_TERM
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function TERM_STRUCTURE_PARTIAL_ADJUSTMENT_FUNC(ByRef TENOR_RNG As Variant, _
ByRef YIELD_RNG As Variant, _
ByRef CHG_RNG As Variant, _
ByVal GAMMA1_VAL As Double, _
ByVal GAMMA2_VAL As Double, _
ByVal GAMMA3_VAL As Double, _
Optional ByVal OUTPUT As Integer = 0)

'CHG_RNG --> One Time Change in Spot Rate (0 No-Chg, 1 Chg)
'TENOR_RNG --> Time to Maturity


'Normally, as the time to maturity increases, the yield increases. Short
'term rates are set, primarily, by the U.S. Federal reserve (i.e. Greenspan)
'whereas long term rates are set by market sentiment. Economists (and others!)
'are interested in the difference between long term (usually 30-year Treasuries)
'and short term (3-month T bills) rates. Anyway, there's little difference between
'Fed Rates and 3-month T-bill rates and the predictive prowess of the Yield Curve
'isn't so precise that.

'If the long term rate is (heaven forbid) less than the short term rate, it's an
'indication of (perhaps) a coming recession - meaning two successive quarters
'with negative GDP growth. Hence, in the past, been a reasonable indicator of a
'coming recession. Let's look more closely at the chart for the early 1990s.
'Here you can see that the inversion of the yield curve did predict the recession.
'In fact the inversion lasted as long as the subsequent recession .

Dim i As Long
Dim NROWS As Long

Dim LONG_RUN_RATE As Double 'In Equilibrium
Dim RATE_ADJUSTMENT As Double '---> LAMBDA
Dim ACTUAL_LEVEL_SPOT As Double
Dim DESIRED_LEVEL_SPOT As Double

'----------------------------------------------------------------------------------------
'From a macroeconomic perspective, the short-term interest rate is a policy instrument
'under the direct control of the central bank. From a finance perspective, long rates
'are risk-adjusted averages of expected future short rates.
'----------------------------------------------------------------------------------------

Dim TEMP_SUM As Double

Dim CHG_VECTOR As Variant
Dim YIELD_VECTOR As Variant
Dim TEMP_VECTOR As Variant
Dim TENOR_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

TENOR_VECTOR = TENOR_RNG
If UBound(TENOR_VECTOR, 1) = 1 Then
    TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(TENOR_VECTOR)
End If

YIELD_VECTOR = YIELD_RNG
If UBound(YIELD_VECTOR, 1) = 1 Then
    YIELD_VECTOR = MATRIX_TRANSPOSE_FUNC(YIELD_VECTOR)
End If

CHG_VECTOR = CHG_RNG
If UBound(CHG_VECTOR, 1) = 1 Then
    CHG_VECTOR = MATRIX_TRANSPOSE_FUNC(CHG_VECTOR)
End If

If UBound(CHG_VECTOR, 1) <> UBound(TENOR_VECTOR) Then: GoTo ERROR_LABEL

RATE_ADJUSTMENT = 1 - GAMMA2_VAL
ACTUAL_LEVEL_SPOT = GAMMA1_VAL / RATE_ADJUSTMENT
DESIRED_LEVEL_SPOT = GAMMA2_VAL / RATE_ADJUSTMENT
LONG_RUN_RATE = GAMMA1_VAL / (1 - GAMMA2_VAL)

NROWS = UBound(CHG_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 7)

TEMP_MATRIX(0, 1) = "TENOR"
TEMP_MATRIX(0, 2) = "CHG. DUMMIES"
TEMP_MATRIX(0, 3) = "R(t-1)"
TEMP_MATRIX(0, 4) = "R*"
TEMP_MATRIX(0, 5) = "ACTUAL YIELD"
TEMP_MATRIX(0, 6) = "ESTIMATED YIELD"
TEMP_MATRIX(0, 7) = "MSE"

TEMP_MATRIX(1, 1) = TENOR_VECTOR(1, 1)
TEMP_MATRIX(1, 2) = CHG_VECTOR(1, 1)
TEMP_MATRIX(1, 3) = ""
TEMP_MATRIX(1, 4) = ""
TEMP_MATRIX(1, 5) = YIELD_VECTOR(1, 1)
TEMP_MATRIX(1, 6) = ""

TEMP_SUM = 0
For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = TENOR_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = CHG_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = YIELD_VECTOR(i - 1, 1)
    TEMP_MATRIX(i, 4) = ACTUAL_LEVEL_SPOT + DESIRED_LEVEL_SPOT * (TEMP_MATRIX(i, 2))
    TEMP_MATRIX(i, 5) = YIELD_VECTOR(i, 1)
    TEMP_MATRIX(i, 6) = GAMMA1_VAL + GAMMA2_VAL * TEMP_MATRIX(i, 3) + GAMMA3_VAL * TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 7) = Abs(TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 6)) ^ 2
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 7)
Next i

TEMP_MATRIX(1, 7) = TEMP_SUM '--> Use solver to minimize this sum by
'changing First, Second and Third Gamma.

'------------------------------------------------------------------------------------
Select Case OUTPUT
'------------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------------
    TERM_STRUCTURE_PARTIAL_ADJUSTMENT_FUNC = TEMP_MATRIX
'------------------------------------------------------------------------------------
Case Else
'------------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 4, 1 To 2)
    TEMP_VECTOR(1, 1) = "RATE_ADJUSTMENT [LAMBDA]"
    TEMP_VECTOR(1, 2) = RATE_ADJUSTMENT
        
    TEMP_VECTOR(2, 1) = "ACTUAL_LEVEL_SPOT"
    TEMP_VECTOR(2, 2) = ACTUAL_LEVEL_SPOT
        
    TEMP_VECTOR(3, 1) = "DESIRED_LEVEL_SPOT"
    TEMP_VECTOR(3, 2) = DESIRED_LEVEL_SPOT
        
    TEMP_VECTOR(4, 1) = "LONG_RUN_RATE"
    TEMP_VECTOR(4, 2) = LONG_RUN_RATE
    
    TERM_STRUCTURE_PARTIAL_ADJUSTMENT_FUNC = TEMP_VECTOR
'------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
TERM_STRUCTURE_PARTIAL_ADJUSTMENT_FUNC = Err.number
End Function

