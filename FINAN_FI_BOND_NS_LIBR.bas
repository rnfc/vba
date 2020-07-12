Attribute VB_Name = "FINAN_FI_BOND_NS_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
Private PUB_COUPON_VEC As Variant
Private PUB_PRICE_VEC As Variant
Private PUB_MATURITY_VEC As Variant

Private PUB_SHORT_RATE_VAL As Double

Private PUB_SETTLEMENT_VAL As Date
Private PUB_FREQUENCY_VAL As Integer
Private PUB_REDEMPTION_VAL As Double
Private PUB_COUNT_BASIS_VAL As Integer
Private PUB_GUESS_YIELD_VAL As Double

Private PUB_DELTA_TENOR_VAL As Double
Private PUB_START_TENOR_VAL As Double
Private PUB_END_TENOR_VAL As Double

Private Const PUB_EPSILON As Double = 2 ^ 52 '1E-100

'Nelson-Siegel Yield Curve Model
'The idea of the Nelson-Siegel (N&S) approach is to fit the empirical form of the yield
'curve with a pre-specified functional form of the spot rates which is a function of the
'time to maturity of the bonds.

'The algorithm fits a term structure of spot rates to the universe of Government bonds.
'This can in turn be used to detect over-, respectively underpriced bonds.

'The N&S parameters are found by minimizing the sum of squared differences between model
'and market prices. A more sophisticated version weighs these differences with the inverse
'of the duration as proposed in Bliss (1998).

'References:
'Nelson, C. R. & Siegel, A. F. (1987). Parsimonious modeling of yield curves, Journal of
'Business 60(4): 473-489.

'Bliss, R. R. (1997). Testing Term Structure Estimation Methods. Advances in Futures and
'Options Research(9), 197-231.

'Formulas from Van Landschoot, Astrid. The Term Structure of Credit Spreads on Euro Corporate
'Bonds. Working Paper, CentER, Tilburg University. April 2003. Downloaded from
'http://ideas.repec.org/p/dgr/kubcen/200346.html February 2004.

'Note that the formula in this paper has a small typo and just shows the basic Nelson &
'Siegel where t2 = t1.

'Svensson (1994) extend Nelson-Siegel (1987) model described above with b3 parameter term.

'References:
'Svensson, L. (1994). Estimating and interpreting forward interest rates: Sweden 1992-4.
'Discussion paper, Centre for Economic Policy Research(1051).

'Anderson, N., Breedon, F., Deacon, M., Derry, A., & Murphy, G. (1996). Estimating and
'interpreting the yield curve.

'Chichester: John Wiley Series in Financial Economics and Quantitative Analysis. Chapter 2.4.6, pgs. 36-41.

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : NS_BOND_FITTING_FUNC
'DESCRIPTION   : This Function implements fitting of Yield curve
'using Nelson Siegel parameterrization
'LIBRARY       : BOND
'GROUP         : NS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************
 
Function NS_BOND_FITTING_FUNC(ByVal SHORT_RATE As Double, _
ByVal SETTLEMENT As Date, _
ByRef MATURITY_RNG As Variant, _
ByRef COUPON_RNG As Variant, _
ByRef CLEAN_PRICE_RNG As Variant, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal REDEMPTION As Double = 100, _
Optional ByVal COUNT_BASIS As Integer = 0, _
Optional ByVal GUESS_YIELD As Double = 0.2, _
Optional ByVal nLOOPS As Long = 500, _
Optional ByVal epsilon As Double = 10 ^ -15)

'BETA0_VAL --> Long-run levels of interest rates --> determines
'magnitude and the direction of the hump
'BETA1_VAL --> Short-run component --> determines magnitude and the
'direction of the hump
'BETA2_VAL --> Medium-term component --> determines magnitude and
'the direction of the hump
'BETA3_VAL --> You can think about BETA1_VAL + BETA2_VAL = Short_Rate
'TAU1_VAL  --> Decay parameter 1 --> determines magnitude and the
'direction of the hump
'TAU2_VAL  --> Decay parameter 2 --> determines magnitude and the
'direction of the hump

Dim CLEAN_VECTOR As Variant
Dim PARAM_VECTOR As Variant
Dim COUPON_VECTOR As Variant
Dim MATURITY_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

MATURITY_VECTOR = MATURITY_RNG
If UBound(MATURITY_VECTOR, 1) = 1 Then: MATURITY_VECTOR = MATRIX_TRANSPOSE_FUNC(MATURITY_VECTOR)

COUPON_VECTOR = COUPON_RNG
If UBound(COUPON_VECTOR, 1) = 1 Then: COUPON_VECTOR = MATRIX_TRANSPOSE_FUNC(COUPON_VECTOR)
If UBound(MATURITY_VECTOR, 1) <> UBound(COUPON_VECTOR, 1) Then: GoTo ERROR_LABEL

CLEAN_VECTOR = CLEAN_PRICE_RNG
If UBound(CLEAN_VECTOR, 1) = 1 Then: CLEAN_VECTOR = MATRIX_TRANSPOSE_FUNC(CLEAN_VECTOR)
If UBound(MATURITY_VECTOR, 1) <> UBound(CLEAN_VECTOR, 1) Then: GoTo ERROR_LABEL

PUB_COUPON_VEC = COUPON_VECTOR
PUB_PRICE_VEC = CLEAN_VECTOR
PUB_MATURITY_VEC = MATURITY_VECTOR
PUB_SHORT_RATE_VAL = SHORT_RATE
PUB_SETTLEMENT_VAL = SETTLEMENT
PUB_FREQUENCY_VAL = FREQUENCY
PUB_REDEMPTION_VAL = REDEMPTION
PUB_COUNT_BASIS_VAL = COUNT_BASIS
PUB_GUESS_YIELD_VAL = GUESS_YIELD
PUB_DELTA_TENOR_VAL = 0.2 '--> Fix This
PUB_START_TENOR_VAL = 0.00001 '--> Fix This
PUB_END_TENOR_VAL = 10 '--> Fix This

PARAM_VECTOR = NELDER_MEAD_OPTIMIZATION_FRAME_FUNC("NS_BOND_OBJ1_FUNC", PARAM_VECTOR, "", True, 0, nLOOPS, epsilon)
NS_BOND_FITTING_FUNC = PARAM_VECTOR

Exit Function
ERROR_LABEL:
NS_BOND_FITTING_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NS_BOND_ERROR_FUNC
'DESCRIPTION   : Nelson Siegel parameterrization error
'LIBRARY       : BOND
'GROUP         : NS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************
 

Function NS_BOND_ERROR_FUNC(ByVal SHORT_RATE As Double, _
ByVal SETTLEMENT As Date, _
ByRef MATURITY_RNG As Variant, _
ByRef COUPON_RNG As Variant, _
ByRef CLEAN_PRICE_RNG As Variant, _
ByRef PARAM_RNG As Variant, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal REDEMPTION As Double = 100, _
Optional ByVal COUNT_BASIS As Integer = 0, _
Optional ByVal GUESS_YIELD As Double = 0.2)

Dim PARAM_VECTOR As Variant

Dim CLEAN_VECTOR As Variant
Dim COUPON_VECTOR As Variant

Dim MATURITY_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
'BETA0_VAL --> Long-run levels of interest rates --> determines
'magnitude and the direction of the hump
'BETA1_VAL --> Short-run component --> determines magnitude and the
'direction of the hump
'BETA2_VAL --> Medium-term component --> determines magnitude and
'the direction of the hump
'BETA3_VAL --> You can think about BETA1_VAL + BETA2_VAL = Short_Rate
'TAU1_VAL  --> Decay parameter 1 --> determines magnitude and the
'direction of the hump
'TAU2_VAL  --> Decay parameter 2 --> determines magnitude and the
'direction of the hump

MATURITY_VECTOR = MATURITY_RNG
If UBound(MATURITY_VECTOR, 1) = 1 Then: MATURITY_VECTOR = MATRIX_TRANSPOSE_FUNC(MATURITY_VECTOR)

COUPON_VECTOR = COUPON_RNG
If UBound(COUPON_VECTOR, 1) = 1 Then: COUPON_VECTOR = MATRIX_TRANSPOSE_FUNC(COUPON_VECTOR)

CLEAN_VECTOR = CLEAN_PRICE_RNG
If UBound(CLEAN_VECTOR, 1) = 1 Then: CLEAN_VECTOR = MATRIX_TRANSPOSE_FUNC(CLEAN_VECTOR)

If UBound(MATURITY_VECTOR, 1) <> UBound(COUPON_VECTOR, 1) Then: GoTo ERROR_LABEL
If UBound(MATURITY_VECTOR, 1) <> UBound(CLEAN_VECTOR, 1) Then: GoTo ERROR_LABEL

PUB_COUPON_VEC = COUPON_VECTOR
PUB_PRICE_VEC = CLEAN_VECTOR
PUB_MATURITY_VEC = MATURITY_VECTOR
PUB_SHORT_RATE_VAL = SHORT_RATE
PUB_SETTLEMENT_VAL = SETTLEMENT
PUB_FREQUENCY_VAL = FREQUENCY
PUB_REDEMPTION_VAL = REDEMPTION
PUB_COUNT_BASIS_VAL = COUNT_BASIS
PUB_GUESS_YIELD_VAL = GUESS_YIELD
PUB_DELTA_TENOR_VAL = 0.2 '--> Fix This
PUB_START_TENOR_VAL = 0.00001 '--> Fix This
PUB_END_TENOR_VAL = 10 '--> Fix This

NS_BOND_ERROR_FUNC = Array(NS_BOND_OBJ1_FUNC(PARAM_VECTOR), NS_BOND_CONSTRAINTS_FUNC(PARAM_VECTOR))

Exit Function
ERROR_LABEL:
NS_BOND_ERROR_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NS_BOND_ARBITRAGE_FUNC
'DESCRIPTION   : The idea of the Nelson-Siegel (N&S) arbitrage approach is to
'fit the empirical form of the yield curve with a pre-specified.
'LIBRARY       : BOND
'GROUP         : NS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function NS_BOND_ARBITRAGE_FUNC(ByVal SETTLEMENT As Date, _
ByRef MATURITY_RNG As Variant, _
ByRef COUPON_RNG As Variant, _
ByRef CLEAN_PRICE_RNG As Variant, _
ByVal BETA0_VAL As Double, _
ByVal BETA1_VAL As Double, _
ByVal BETA2_VAL As Double, _
ByVal BETA3_VAL As Double, _
ByVal TAU1_VAL As Double, _
ByVal TAU2_VAL As Double, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal REDEMPTION As Double = 100, _
Optional ByVal COUNT_BASIS As Integer = 0, _
Optional ByVal GUESS_YIELD As Double = 0.2)

'BETA0_VAL --> Long-run levels of interest rates --> determines
'magnitude and the direction of the hump
'BETA1_VAL --> Short-run component --> determines magnitude and the
'direction of the hump
'BETA2_VAL --> Medium-term component --> determines magnitude and
'the direction of the hump
'BETA3_VAL --> You can think about BETA1_VAL + BETA2_VAL = Short_Rate
'TAU1_VAL  --> Decay parameter 1 --> determines magnitude and the
'direction of the hump
'TAU2_VAL  --> Decay parameter 2 --> determines magnitude and the
'direction of the hump


Dim i As Long

Dim NROWS As Long

Dim TEMP_VECTOR As Variant
Dim CLEAN_VECTOR As Variant
Dim COUPON_VECTOR As Variant
Dim MATURITY_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

MATURITY_VECTOR = MATURITY_RNG
If UBound(MATURITY_VECTOR, 1) = 1 Then: MATURITY_VECTOR = MATRIX_TRANSPOSE_FUNC(MATURITY_VECTOR)

COUPON_VECTOR = COUPON_RNG
If UBound(COUPON_VECTOR, 1) = 1 Then: COUPON_VECTOR = MATRIX_TRANSPOSE_FUNC(COUPON_VECTOR)

CLEAN_VECTOR = CLEAN_PRICE_RNG
If UBound(CLEAN_VECTOR, 1) = 1 Then: CLEAN_VECTOR = MATRIX_TRANSPOSE_FUNC(CLEAN_VECTOR)

If UBound(MATURITY_VECTOR, 1) <> UBound(COUPON_VECTOR, 1) Then: GoTo ERROR_LABEL
If UBound(MATURITY_VECTOR, 1) <> UBound(CLEAN_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(MATURITY_VECTOR, 1)
ReDim TEMP_MATRIX(0 To NROWS, 1 To 7)

TEMP_MATRIX(0, 1) = ("MID CLEAN")
TEMP_MATRIX(0, 2) = ("MID DIRTY")
TEMP_MATRIX(0, 3) = ("BOND YIELD")
TEMP_MATRIX(0, 4) = ("NS CASH PRICE")
TEMP_MATRIX(0, 5) = ("BOND DURATION")
TEMP_MATRIX(0, 6) = ("WEIGHTS")
TEMP_MATRIX(0, 7) = ("(-CHEAP) / (+RICH) ")

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = CLEAN_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = CLEAN_VECTOR(i, 1) + ACCRINT_FUNC(SETTLEMENT, MATURITY_VECTOR(i, 1), COUPON_VECTOR(i, 1), FREQUENCY, COUNT_BASIS)
    TEMP_MATRIX(i, 3) = BOND_YIELD_FUNC(CLEAN_VECTOR(i, 1), SETTLEMENT, MATURITY_VECTOR(i, 1), COUPON_VECTOR(i, 1), FREQUENCY, REDEMPTION, COUNT_BASIS, GUESS_YIELD)
    TEMP_MATRIX(i, 4) = NS_BOND_CASH_PRICE_FUNC(SETTLEMENT, MATURITY_VECTOR(i, 1), COUPON_VECTOR(i, 1), BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL, FREQUENCY, REDEMPTION, COUNT_BASIS, 0)
    TEMP_MATRIX(i, 5) = BOND_CONVEXITY_DURATION_FUNC(SETTLEMENT, MATURITY_VECTOR(i, 1), COUPON_VECTOR(i, 1), TEMP_MATRIX(i, 3), FREQUENCY, REDEMPTION, COUNT_BASIS)(3)
Next i

TEMP_VECTOR = NS_BOND_INVERSE_DURATION_FUNC(MATRIX_GET_COLUMN_FUNC(TEMP_MATRIX, 2, 1), MATRIX_GET_COLUMN_FUNC(TEMP_MATRIX, 4, 1), MATRIX_GET_COLUMN_FUNC(TEMP_MATRIX, 5, 1), REDEMPTION, 3)
If IsArray(TEMP_VECTOR) = False Then: GoTo ERROR_LABEL
For i = 1 To NROWS
    TEMP_MATRIX(i, 6) = TEMP_VECTOR(i, 1)
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 4)
Next i
NS_BOND_ARBITRAGE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
NS_BOND_ARBITRAGE_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NS_BOND_CASH_PRICE_FUNC
'DESCRIPTION   : Function that returns the present value of a bond discounted on the
'extended Nelson & Siegel (87) term structure as % of face value
'LIBRARY       : BOND
'GROUP         : NS
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function NS_BOND_CASH_PRICE_FUNC(ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
ByVal COUPON As Double, _
ByVal BETA0_VAL As Double, _
ByVal BETA1_VAL As Double, _
ByVal BETA2_VAL As Double, _
ByVal BETA3_VAL As Double, _
ByVal TAU1_VAL As Double, _
ByVal TAU2_VAL As Double, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal REDEMPTION As Double = 100, _
Optional ByVal COUNT_BASIS As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim PERIODS As Double
'Dim TENOR As Double
Dim TENOR As Double
Dim TEMP_SUM As Double ' PV of Cash Flow
Dim NS_YIELD_VAL As Double
Dim MATURITY_VECTOR As Variant


On Error GoTo ERROR_LABEL

If TAU2_VAL = 0 Then TAU2_VAL = TAU1_VAL  'for basic Nelson Siegel TAU2_VAL = TAU1_VAL

MATURITY_VECTOR = _
BOND_DATES_BOND_TENOR_FUNC(SETTLEMENT, MATURITY, FREQUENCY, COUNT_BASIS)

PERIODS = UBound(MATURITY_VECTOR, 1)
TENOR = MATURITY_VECTOR(PERIODS, 1)

NS_YIELD_VAL = NS_SVENSSON_YIELD_FUNC(TENOR, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL)
TEMP_SUM = (COUPON / FREQUENCY) * 100 * NS_DISCOUNT_FACTOR_FUNC(TENOR, NS_YIELD_VAL, 1, 0) 'PV of Last Payment
TEMP_SUM = TEMP_SUM + REDEMPTION * NS_DISCOUNT_FACTOR_FUNC(TENOR, NS_YIELD_VAL, 1, 0) 'PV Of Face

'TENOR = TENOR - (1 / FREQUENCY)
'While TENOR > 0   ' Loop though all coupons backward
For i = (PERIODS - 1) To 1 Step -1
    TENOR = MATURITY_VECTOR(i, 1)
    NS_YIELD_VAL = NS_SVENSSON_YIELD_FUNC(TENOR, BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL)
    TEMP_SUM = TEMP_SUM + 100 * COUPON / FREQUENCY * NS_DISCOUNT_FACTOR_FUNC(TENOR, NS_YIELD_VAL, 1, 0) 'PV Of Coupons
 '   TENOR = TENOR - (1 / FREQUENCY)
Next i
'Wend

Select Case OUTPUT
Case 0 'Cash Price
    NS_BOND_CASH_PRICE_FUNC = TEMP_SUM
Case Else 'Clean Price
    NS_BOND_CASH_PRICE_FUNC = TEMP_SUM - ACCRINT_FUNC(SETTLEMENT, MATURITY, COUPON, FREQUENCY, COUNT_BASIS)
End Select

Exit Function
ERROR_LABEL:
NS_BOND_CASH_PRICE_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : NS_SPOT_RATE_TABLE_FUNC
'DESCRIPTION   : Nelson & Siegel Spot Rate Table
'LIBRARY       : BOND
'GROUP         : NS
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'**********************************************************************************
'**********************************************************************************

Function NS_SPOT_RATE_TABLE_FUNC(ByVal BETA0_VAL As Double, _
ByVal BETA1_VAL As Double, _
ByVal BETA2_VAL As Double, _
ByVal BETA3_VAL As Double, _
ByVal TAU1_VAL As Double, _
ByVal TAU2_VAL As Double, _
Optional ByVal DELTA_TENOR As Double = 0.2, _
Optional ByVal START_TENOR As Double = 0.00001, _
Optional ByVal END_TENOR As Double = 10)

Dim i As Long
Dim PERIODS As Long
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL
    
If START_TENOR < 0.00001 Then: START_TENOR = 0.00001
If DELTA_TENOR < 0.1 Then: DELTA_TENOR = 0.1
PERIODS = CInt((END_TENOR - START_TENOR) / DELTA_TENOR) + 1

ReDim TEMP_MATRIX(0 To PERIODS, 1 To 3)
TEMP_MATRIX(0, 1) = "TENOR"
TEMP_MATRIX(0, 2) = "DISCOUNT FACTOR"
TEMP_MATRIX(0, 3) = "ZERO RATE"

TEMP_MATRIX(1, 1) = START_TENOR
TEMP_MATRIX(1, 3) = NS_SVENSSON_YIELD_FUNC(TEMP_MATRIX(1, 1), BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL)
TEMP_MATRIX(1, 2) = Exp(-1 * TEMP_MATRIX(1, 3) * TEMP_MATRIX(1, 1))

For i = 2 To PERIODS
    TEMP_MATRIX(i, 1) = TEMP_MATRIX(i - 1, 1) + DELTA_TENOR
    TEMP_MATRIX(i, 3) = NS_SVENSSON_YIELD_FUNC(TEMP_MATRIX(i, 1), BETA0_VAL, BETA1_VAL, BETA2_VAL, BETA3_VAL, TAU1_VAL, TAU2_VAL)
    TEMP_MATRIX(i, 2) = Exp(-1 * TEMP_MATRIX(i, 3) * TEMP_MATRIX(i, 1))
Next i

NS_SPOT_RATE_TABLE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
NS_SPOT_RATE_TABLE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NS_BOND_OBJ1_FUNC
'DESCRIPTION   : Nelson Siegel Objective Function
'LIBRARY       : BOND
'GROUP         : NS
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************
 
Private Function NS_BOND_OBJ1_FUNC(ByRef PARAM_VECTOR As Variant)

Dim i As Long
Dim NROWS As Long

Dim YIELD_VAL As Double

Dim YTEMP_VAL As Variant
Dim XTEMP_VAL As Double

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant
Dim TEMP3_VECTOR As Variant

On Error GoTo ERROR_LABEL

NROWS = UBound(PUB_MATURITY_VEC, 1)
ReDim TEMP1_VECTOR(1 To NROWS, 1 To 1)
ReDim TEMP2_VECTOR(1 To NROWS, 1 To 1)
ReDim TEMP3_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    TEMP1_VECTOR(i, 1) = PUB_PRICE_VEC(i, 1) + ACCRINT_FUNC(PUB_SETTLEMENT_VAL, PUB_MATURITY_VEC(i, 1), PUB_COUPON_VEC(i, 1), PUB_FREQUENCY_VAL, PUB_COUNT_BASIS_VAL)
    YIELD_VAL = BOND_YIELD_FUNC(PUB_PRICE_VEC(i, 1), PUB_SETTLEMENT_VAL, PUB_MATURITY_VEC(i, 1), PUB_COUPON_VEC(i, 1), PUB_FREQUENCY_VAL, PUB_REDEMPTION_VAL, PUB_COUNT_BASIS_VAL, PUB_GUESS_YIELD_VAL)
    TEMP2_VECTOR(i, 1) = NS_BOND_CASH_PRICE_FUNC(PUB_SETTLEMENT_VAL, PUB_MATURITY_VEC(i, 1), PUB_COUPON_VEC(i, 1), PARAM_VECTOR(1, 1), PARAM_VECTOR(2, 1), PARAM_VECTOR(3, 1), PARAM_VECTOR(4, 1), PARAM_VECTOR(5, 1), PARAM_VECTOR(6, 1), PUB_FREQUENCY_VAL, PUB_REDEMPTION_VAL, PUB_COUNT_BASIS_VAL, 0)
    TEMP3_VECTOR(i, 1) = BOND_CONVEXITY_DURATION_FUNC(PUB_SETTLEMENT_VAL, PUB_MATURITY_VEC(i, 1), PUB_COUPON_VEC(i, 1), YIELD_VAL, PUB_FREQUENCY_VAL, PUB_REDEMPTION_VAL, PUB_COUNT_BASIS_VAL)(3)
Next i

YTEMP_VAL = NS_BOND_INVERSE_DURATION_FUNC(TEMP1_VECTOR, TEMP2_VECTOR, TEMP3_VECTOR, PUB_REDEMPTION_VAL, 0)
If YTEMP_VAL <> PUB_EPSILON Then
    If NS_BOND_CONSTRAINTS_FUNC(PARAM_VECTOR) = False Then
        'GoTo ERROR_LABEL
        XTEMP_VAL = PUB_EPSILON
    Else
        XTEMP_VAL = 1
    End If
    NS_BOND_OBJ1_FUNC = Abs(YTEMP_VAL * XTEMP_VAL) ^ 2
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
NS_BOND_OBJ1_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NS_BOND_CONSTRAINTS_FUNC
'DESCRIPTION   : Nelson Siegel parameterrization constraint function
'LIBRARY       : BOND
'GROUP         : NS
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************
 
Function NS_BOND_CONSTRAINTS_FUNC(ByRef PARAM_VECTOR As Variant)

Dim i As Long
Dim NROWS As Long

Dim PUB_DELTA_TENOR_VAL As Double
Dim START_TENOR As Double
Dim END_TENOR As Double

Dim TEMP_MATRIX As Variant

Const tolerance0 As Double = 10 ^ -8
'Const tolerance1 As Double = 10 ^ -10
On Error GoTo ERROR_LABEL

NS_BOND_CONSTRAINTS_FUNC = True

'-----------------------------------------------------------------------------------
If Abs((PARAM_VECTOR(1, 1) + PARAM_VECTOR(2, 1)) - PUB_SHORT_RATE_VAL) ^ 2 > tolerance0 Then
    NS_BOND_CONSTRAINTS_FUNC = False
    Exit Function
End If
'Debug.Print "Pass 1"
'-----------------------------------------------------------------------------------
If PUB_START_TENOR_VAL < 0.00001 Then: PUB_START_TENOR_VAL = 0.00001
If PUB_DELTA_TENOR_VAL < 0.1 Then: PUB_DELTA_TENOR_VAL = 0.1
NROWS = (PUB_END_TENOR_VAL - PUB_START_TENOR_VAL) / PUB_DELTA_TENOR_VAL
ReDim TEMP_MATRIX(0 To NROWS, 1 To 4)
TEMP_MATRIX(0, 1) = PUB_START_TENOR_VAL
TEMP_MATRIX(0, 2) = NS_SVENSSON_YIELD_FUNC(TEMP_MATRIX(0, 1), PARAM_VECTOR(1, 1), PARAM_VECTOR(2, 1), PARAM_VECTOR(3, 1), PARAM_VECTOR(4, 1), PARAM_VECTOR(5, 1), PARAM_VECTOR(6, 1))
TEMP_MATRIX(0, 3) = Exp(-1 * TEMP_MATRIX(0, 2) * TEMP_MATRIX(0, 1))
TEMP_MATRIX(0, 4) = 0

If TEMP_MATRIX(0, 2) < 0 Then
   NS_BOND_CONSTRAINTS_FUNC = False
   Exit Function
End If
'Debug.Print "Pass 2"
'-----------------------------------------------------------------------------------

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = TEMP_MATRIX(i - 1, 1) + PUB_DELTA_TENOR_VAL
    TEMP_MATRIX(i, 2) = NS_SVENSSON_YIELD_FUNC(TEMP_MATRIX(i, 1), PARAM_VECTOR(1, 1), PARAM_VECTOR(2, 1), PARAM_VECTOR(3, 1), PARAM_VECTOR(4, 1), PARAM_VECTOR(5, 1), PARAM_VECTOR(6, 1))
    TEMP_MATRIX(i, 3) = Exp(-1 * TEMP_MATRIX(i, 2) * TEMP_MATRIX(i, 1))
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i - 1, 4) - TEMP_MATRIX(i, 4)
    If TEMP_MATRIX(i, 4) < 0 Then
       NS_BOND_CONSTRAINTS_FUNC = False
       Exit Function
    End If
Next i
'Debug.Print "Pass 3"
'-----------------------------------------------------------------------------------

If TEMP_MATRIX(NROWS, 2) < 0 Then
    NS_BOND_CONSTRAINTS_FUNC = False
    Exit Function
End If
'Debug.Print "Pass 4"
'-----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
NS_BOND_CONSTRAINTS_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : NS_BOND_INVERSE_DURATION_FUNC
'DESCRIPTION   : Min objective function for testing Term Structure Estimation
'Methods. Advances in Inverse duration weight.
'LIBRARY       : BOND
'GROUP         : NS
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function NS_BOND_INVERSE_DURATION_FUNC(ByRef CASH_PRICE_VECTOR As Variant, _
ByRef FAIR_VALUE_VECTOR As Variant, _
ByRef DURATION_VECTOR As Variant, _
Optional ByVal PAR_VALUE As Double = 100, _
Optional ByVal OUTPUT As Integer = 0)

'DURATION_RNG MUST BE BASED ON THE CLEAN_PRICE_RNG --> HOLDING THE YTM CONSTANT
'FAIR_VALUE_RNG MUST BE BASED ON THE MODEL CASH PRICE

Dim i As Long
Dim NROWS As Long

Dim TEMP_VAL As Double
Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

NROWS = UBound(DURATION_VECTOR, 1)

TEMP1_SUM = 0: TEMP2_SUM = 0: TEMP3_SUM = 0
For i = 1 To NROWS: TEMP1_SUM = TEMP1_SUM + 1 / DURATION_VECTOR(i, 1): Next i
If OUTPUT > 1 Then: ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    TEMP_VAL = 1 / DURATION_VECTOR(i, 1) / TEMP1_SUM
    If OUTPUT <= 1 Then
        TEMP2_SUM = TEMP2_SUM + Abs(FAIR_VALUE_VECTOR(i, 1) / PAR_VALUE - CASH_PRICE_VECTOR(i, 1) / PAR_VALUE) ^ 2
        TEMP3_SUM = TEMP3_SUM + (Abs(FAIR_VALUE_VECTOR(i, 1) / PAR_VALUE - CASH_PRICE_VECTOR(i, 1) / PAR_VALUE) * TEMP_VAL) ^ 2
    Else
        TEMP_VECTOR(i, 1) = TEMP_VAL
    End If
Next i

'THE OBJECTIVE FUNCTION IS EITHER TO MINIMIZE TEMP2_SUM OR TEMP3_SUM BY CHANGING
'THE PARAMETERS OF THE MODEL:

'SOME OF THE CONSTRAINTS MUST INCLUDE:
'---> Rate r at time 0 must remain positive (mmin is a value just slightly larger than 0)
'---> Rate at the end of the estimation horizon must remain positive
'---> Discount functions must be non-increasing

Select Case OUTPUT
Case 0
    NS_BOND_INVERSE_DURATION_FUNC = TEMP2_SUM * 1000
    'Inverse duration weighted function x 10^3
Case 1
    NS_BOND_INVERSE_DURATION_FUNC = TEMP3_SUM * 10000
    'Inverse duration weighted function x 10^5
Case Else
    NS_BOND_INVERSE_DURATION_FUNC = TEMP_VECTOR 'Weights (wi)
End Select

Exit Function
ERROR_LABEL:
NS_BOND_INVERSE_DURATION_FUNC = PUB_EPSILON
End Function
