Attribute VB_Name = "FINAN_FI_BOND_YIELD_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private PUB_BOND_ARR As Variant
Private Const PUB_EPSILON As Double = 2 ^ 52

'************************************************************************************
'************************************************************************************
'FUNCTION      : BOND_CASH_PRICE_FUNC
'DESCRIPTION   : Returns the cash or clean price of a security that pays
'periodic interest.
'LIBRARY       : BOND
'GROUP         : YIELD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function BOND_CASH_PRICE_FUNC(ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
ByVal COUPON As Double, _
ByVal YIELD As Double, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal REDEMPTION As Double = 100, _
Optional ByVal COUNT_BASIS As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim h As Long
Dim i As Long 'PERIODS
Dim j As Long 'COUPONS
Dim k As Long

Dim PDAYS_VAL As Double
Dim NDAYS_VAL As Double

Dim TEMP_SUM As Double

On Error GoTo ERROR_LABEL
    
If SETTLEMENT > MATURITY Then
    BOND_CASH_PRICE_FUNC = 0
    Exit Function
End If
If SETTLEMENT = MATURITY Then
    BOND_CASH_PRICE_FUNC = REDEMPTION
    Exit Function
End If

k = FREQUENCY
If k = 0 Then: k = k + 1

i = COUPNUM_FUNC(SETTLEMENT, MATURITY, k)
PDAYS_VAL = COUPDAYBS_FUNC(SETTLEMENT, MATURITY, k, COUNT_BASIS)
NDAYS_VAL = COUPDAYSNC_FUNC(SETTLEMENT, MATURITY, k, COUNT_BASIS)
j = PDAYS_VAL + NDAYS_VAL

'----------------------------------------------------------------------
Select Case COUNT_BASIS
'----------------------------------------------------------------------
Case 0, 1, 4 'US (NASD) 30/360 ; Actual/Actual; European 30/360
'----------------------------------------------------------------------
Case 2 'Actual / 360 --> PERFECT
'----------------------------------------------------------------------
    PDAYS_VAL = ((j / (360 / k)) * PDAYS_VAL)
    NDAYS_VAL = j - (PDAYS_VAL)
'----------------------------------------------------------------------
Case 3 'Actual / 365 --> PERFECT
'----------------------------------------------------------------------
    PDAYS_VAL = ((j / (365 / k)) * PDAYS_VAL)
    NDAYS_VAL = j - (PDAYS_VAL)
'----------------------------------------------------------------------
End Select
'----------------------------------------------------------------------
TEMP_SUM = 0
For h = 1 To i: TEMP_SUM = TEMP_SUM + (100 * (COUPON / k) / (1 + (YIELD / k)) ^ (h - 1 + (NDAYS_VAL / j))): Next h
TEMP_SUM = REDEMPTION / ((1 + YIELD / k) ^ (i - 1 + (NDAYS_VAL / j))) + TEMP_SUM - 100 * (COUPON / k) * (PDAYS_VAL / j)

Select Case OUTPUT
Case 0 'Cash Price
    BOND_CASH_PRICE_FUNC = TEMP_SUM + ACCRINT_FUNC(SETTLEMENT, MATURITY, COUPON, k, COUNT_BASIS)
Case Else 'Clean Price
    BOND_CASH_PRICE_FUNC = TEMP_SUM
End Select

Exit Function
ERROR_LABEL:
BOND_CASH_PRICE_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ACCRINT_FUNC
'DESCRIPTION   : Returns the accrued interest of a security that pays
'periodic interest.
'LIBRARY       : BOND
'GROUP         : PRICE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************


Function ACCRINT_FUNC(ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
ByVal COUPON As Double, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal COUNT_BASIS As Integer = 0)
    
Dim j As Long

Dim PDAYS_VAL As Double
Dim NDAYS_VAL As Double

Dim FACTOR_VAL As Double

On Error GoTo ERROR_LABEL
    
If (MATURITY <= SETTLEMENT) Or (FREQUENCY < 1) Or (COUPON <= 0) Then
    ACCRINT_FUNC = 0
    Exit Function
End If
 
PDAYS_VAL = COUPDAYBS_FUNC(SETTLEMENT, MATURITY, FREQUENCY, COUNT_BASIS)
NDAYS_VAL = COUPDAYSNC_FUNC(SETTLEMENT, MATURITY, FREQUENCY, COUNT_BASIS)
j = PDAYS_VAL + NDAYS_VAL

If j = 0 Then
    ACCRINT_FUNC = 0
    Exit Function
End If
 
Select Case COUNT_BASIS
Case 0, 4 'US (NASD) 30/360 ; European 30/360
    FACTOR_VAL = PDAYS_VAL / j
Case 1 'Actual / Actual
    FACTOR_VAL = PDAYS_VAL / j
Case 2 'Actual / 360
    FACTOR_VAL = PDAYS_VAL / (360 / FREQUENCY)
Case 3 'Actual / 365
    FACTOR_VAL = PDAYS_VAL / (365 / FREQUENCY)
End Select
 
ACCRINT_FUNC = (COUPON / FREQUENCY) * 100 * FACTOR_VAL
          
Exit Function
ERROR_LABEL:
ACCRINT_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BOND_YIELD_FUNC
'DESCRIPTION   : Returns the yield of a security that pays periodic interest.
'LIBRARY       : BOND
'GROUP         : YIELD
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 15-06-2010
'************************************************************************************
'************************************************************************************

Function BOND_YIELD_FUNC(ByVal CLEAN_PRICE As Double, _
ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
ByVal COUPON As Double, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal REDEMPTION As Double = 100, _
Optional ByVal COUNT_BASIS As Integer = 0, _
Optional ByVal GUESS_YIELD As Double = 0.3)

Dim nLOOPS As Long
Dim CONVERG_VAL As Integer

Dim COUNTER As Long

Dim Y_VAL As Double
Dim tolerance As Double

On Error GoTo ERROR_LABEL

If (MATURITY <= SETTLEMENT) Then
    BOND_YIELD_FUNC = 0
    Exit Function
End If

ReDim PUB_BOND_ARR(1 To 7)
PUB_BOND_ARR(1) = COUPON: PUB_BOND_ARR(2) = FREQUENCY
PUB_BOND_ARR(3) = REDEMPTION: PUB_BOND_ARR(4) = COUNT_BASIS
PUB_BOND_ARR(5) = CLEAN_PRICE: PUB_BOND_ARR(6) = MATURITY
PUB_BOND_ARR(7) = SETTLEMENT

CONVERG_VAL = 0: COUNTER = 0: nLOOPS = 600: tolerance = 10 ^ -15

Y_VAL = PARAB_ZERO_FUNC(-GUESS_YIELD, GUESS_YIELD, "CALL_BOND_YIELD_OBJ_FUNC", CONVERG_VAL, COUNTER, nLOOPS, tolerance)
'Y_VAL = SECANT_ZERO_FUNC(-GUESS_YIELD, GUESS_YIELD, "CALL_BOND_YIELD_OBJ_FUNC", CONVERG_VAL, COUNTER, nLOOPS, tolerance)
'Y_VAL = NEWTON_ZERO_FUNC(GUESS_YIELD, "CALL_BOND_YIELD_OBJ_FUNC", "", CONVERG_VAL, COUNTER, nLOOPS, tolerance)
'Y_VAL = MULLER_ZERO_FUNC(-GUESS_YIELD, GUESS_YIELD, "CALL_BOND_YIELD_OBJ_FUNC", CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL <> 0 Or Y_VAL = 2 ^ 52 Then
    BOND_YIELD_FUNC = GUESS_YIELD
Else
    BOND_YIELD_FUNC = Y_VAL
End If
'BOND_YIELD_FUNC = CALL_TEST_ZERO_FRAME_FUNC(-GUESS_YIELD, GUESS_YIELD, "CALL_BOND_YIELD_OBJ_FUNC", nLOOPS, tolerance)

Exit Function
ERROR_LABEL:
BOND_YIELD_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_BOND_YIELD_OBJ_FUNC
'DESCRIPTION   : Bond Yield Function for the Root finding algorithm
'LIBRARY       : BOND
'GROUP         : YIELD
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 15-06-2010
'************************************************************************************
'************************************************************************************

Function CALL_BOND_YIELD_OBJ_FUNC(ByRef X_VAL As Double)

Dim Y_VAL As Double

On Error GoTo ERROR_LABEL

Y_VAL = BOND_CASH_PRICE_FUNC(PUB_BOND_ARR(7), _
        PUB_BOND_ARR(6), _
        PUB_BOND_ARR(1), _
        X_VAL, _
        PUB_BOND_ARR(2), _
        PUB_BOND_ARR(3), _
        PUB_BOND_ARR(4), 1)

CALL_BOND_YIELD_OBJ_FUNC = Abs(Y_VAL - PUB_BOND_ARR(5)) ^ 2

Exit Function
ERROR_LABEL:
CALL_BOND_YIELD_OBJ_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CURRENT_YIELD_FUNC
'DESCRIPTION   : Returns the yield of a security that pays periodic interest
'LIBRARY       : BOND
'GROUP         : YIELD
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 15-06-2010
'************************************************************************************
'************************************************************************************


Function CURRENT_YIELD_FUNC(ByVal CLEAN_PRICE As Double, _
ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
ByVal COUPON As Double, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal COUNT_BASIS As Integer = 0)

On Error GoTo ERROR_LABEL

If (MATURITY <= SETTLEMENT) Or (COUPON = 0) Or (FREQUENCY < 1) Then
    CURRENT_YIELD_FUNC = 0
    Exit Function
End If

CURRENT_YIELD_FUNC = COUPON * 100 / (CLEAN_PRICE - ACCRINT_FUNC(SETTLEMENT, _
                     MATURITY, COUPON, FREQUENCY, COUNT_BASIS))

Exit Function
ERROR_LABEL:
CURRENT_YIELD_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EFFECTIVE_YIELD_FUNC
'DESCRIPTION   : Returns the effective yield of a security that pays periodic
'interest
'LIBRARY       : BOND
'GROUP         : YIELD
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 15-06-2010
'************************************************************************************
'************************************************************************************


Function EFFECTIVE_YIELD_FUNC(ByVal CLEAN_PRICE As Double, _
ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
ByVal COUPON As Double, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal REDEMPTION As Double = 100, _
Optional ByVal COUNT_BASIS As Integer = 0, _
Optional ByVal GUESS_YIELD As Double = 0.1)

On Error GoTo ERROR_LABEL


If (MATURITY <= SETTLEMENT) Or (FREQUENCY < 1) Then
    EFFECTIVE_YIELD_FUNC = 0
    Exit Function
End If

EFFECTIVE_YIELD_FUNC = (1 + BOND_YIELD_FUNC(CLEAN_PRICE, SETTLEMENT, _
    MATURITY, COUPON, FREQUENCY, REDEMPTION, COUNT_BASIS, _
    GUESS_YIELD) / FREQUENCY) ^ FREQUENCY - 1

Exit Function
ERROR_LABEL:
EFFECTIVE_YIELD_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_YIELD_FUNC
'DESCRIPTION   : Returns the yield to call of a security that pays
'periodic interest
'LIBRARY       : BOND
'GROUP         : YIELD
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 15-06-2010
'************************************************************************************
'************************************************************************************

Function CALL_YIELD_FUNC(ByVal SETTLEMENT As Date, _
ByVal CALL_DATE As Date, _
ByVal CLEAN_PRICE As Double, _
ByVal CALL_PRICE As Double, _
ByVal COUPON As Double, _
ByVal FREQUENCY As Integer, _
Optional ByVal COUNT_BASIS As Integer = 0, _
Optional ByVal GUESS_YIELD As Double = 0.1)

On Error GoTo ERROR_LABEL

If (CALL_DATE <= SETTLEMENT) Then
    CALL_YIELD_FUNC = 0
    Exit Function
End If

CALL_YIELD_FUNC = BOND_YIELD_FUNC(CLEAN_PRICE, SETTLEMENT, CALL_DATE, COUPON, _
                  FREQUENCY, CALL_PRICE, COUNT_BASIS, GUESS_YIELD)

Exit Function
ERROR_LABEL:
CALL_YIELD_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : REAL_YIELD_FUNC

'DESCRIPTION   : Returns the effective annual interest rate, given the
'reinvestment rate of the coupons

'LIBRARY       : BOND
'GROUP         : YIELD
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 15-06-2010
'************************************************************************************
'************************************************************************************

Function REAL_YIELD_FUNC(ByVal SETTLEMENT As Date, _
ByVal MATURITY As Date, _
ByVal SALE_DATE As Date, _
ByVal REINV_RATE As Double, _
ByVal COUPON As Double, _
ByVal TARGET_YIELD As Double, _
ByVal ENTRY_PRICE As Double, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal REDEMPTION As Double = 100, _
Optional ByVal COUNT_BASIS As Integer = 0)

'A bond swap involving a switch from a low-coupon bond to a higher-coupon bond
'of similar quality and maturity in order to pick up higher current yield and
'a better yield to maturity.  The investor is not trying to predict interest
'rate changes, and the swap is not based on any imbalance in yield spread.
'The object is simply to seek higher realized compounded yield.

'REINVESTMENT RATE MUST BE ENTERED with m = FREQUENCY compounding
'YIELD MUST BE ENTERED with m = FREQUENCY compounding
'COUPON RATE MUST BE ENTERED with m = FREQUENCY compounding

Dim TENOR_VAL As Double
Dim ACCRUED_VAL As Double
Dim TARGET_PRICE As Double 'AT SALES DATE Assuming YIELD
Dim REAL_YIELD_VAL As Double

On Error GoTo ERROR_LABEL

If (MATURITY <= SETTLEMENT) Or (FREQUENCY < 1) Then
    REAL_YIELD_FUNC = 0
    Exit Function
End If

TARGET_PRICE = BOND_CASH_PRICE_FUNC(SALE_DATE, MATURITY, COUPON, TARGET_YIELD, FREQUENCY, REDEMPTION, COUNT_BASIS, 0) 'This is the estimated cash price
TENOR_VAL = YEARFRAC_FUNC(SETTLEMENT, SALE_DATE, COUNT_BASIS)
ACCRUED_VAL = ((1 + REINV_RATE / FREQUENCY) ^ (FREQUENCY * TENOR_VAL) - 1) / (REINV_RATE / FREQUENCY) * COUPON / FREQUENCY
ACCRUED_VAL = ACCRUED_VAL * 100 + TARGET_PRICE 'SAME AS FV
'FORMULA WITH CONSTANT CASH FLOWS
REAL_YIELD_VAL = (((ACCRUED_VAL / ENTRY_PRICE) ^ (1 / (TENOR_VAL * FREQUENCY))) - 1) * FREQUENCY 'nominal compound YIELD
REAL_YIELD_VAL = (1 + REAL_YIELD_VAL / FREQUENCY) ^ FREQUENCY - 1
'Realized Compound YIELD
REAL_YIELD_FUNC = REAL_YIELD_VAL

Exit Function
ERROR_LABEL:
REAL_YIELD_FUNC = PUB_EPSILON
End Function

'------------------------------------------------------------------------------
'---------------------------Active bond stragegies-----------------------------
'------------------------------------------------------------------------------


'BOND STRATEGY A: "Rate Anticipation Swap"

'For a rate anticipation swap, you rebalance your bond portfolio as necessary
'to take advantage of an anticipated change in yields.  If you expect yields
'to decline, you would swap into bonds that would have the largest gain in value
'when that happens (longer maturity, lower coupon, etc.)

'BOND STRATEGY B: "Substitution swap"

'Substitution swap:  The substitution swap is done to exploit an
'apparent short-term mispricing between two bond issues that are
'identical with respect to coupon rate, credit rating, and time
'to maturity.  It is subject to considerably more risk than the
'pure ESTIMATED_YIELD pickup swap, as the apparent mispricing may persist
'because of quality differences the market detects before the
'bond-rating agencies do.

'BOND STRATEGY C: "Intermarket Spread Swap"

'Intermarket Spread Swap:  An intermarket spread swap is done to exploit
'an apparent short-term change in the spreads between bonds of different
'risk.  For example, if the spread between Treasuries and Corporate bonds
'is wider than normal an appropriate strategy would be to shift from
'government to corporate bonds.  When the spread returns to normal, the
'values of Treasuries will fall and/or the values of corporate bonds will rise.
'In a nutshell, if the spread does narrow then corporates will outperform
'treasuries during the correction.

'The risk under this strategy is that you are betting that the increase in
'yields is an anomaly and will only exist temporarily.  If, in fact, the
'increased spread is justified based on justified increases in the default
'yield for corporates, then the superior performance you expect will not be
'realized.  The larger yield spread may be just enough to cover the addition
'risk.  If the spread does not narrow, the performance of your corporate
'investment will be much lower, as outlined below.

'Example:  Currently hold:  20-year T-bond, 7.0% coupon priced at $1,000.00
'to yield 7.0%.  Swap candidate:  20-year., Baa 10% coupon priced at $1,000
'to yield 10.0%.  The spread is 300 basis points but suppose the historical
'spread is 200 basis points.  Suppose you anticipate the spread to narrow to
'200 basis points via a 100 basis point decrease in corporate yields.  (Note,
'we could come up with just about any combination that results in the spread
'decreasing by 100 basis points.  The outcome would be similar, namely corporates
'would outperform Treasuries)  The analysis for a one-year time frame is provided
'below.  For simplicity, assume the next semi-annual coupon payment occurs in
'exactly six months.

'BOND STRATEGY D: "pure yield pickup swap"

'The rewards for a pure yield pickup swap are automatic and instantaneous
'in that both a higher-coupon yield and a higher yield to maturity are
'realized from the swap.

'Additional advantages:
'1.  No specific work-out period needed because the investor is assumed
'to hold the new bond to maturity.
'2.  No need to predict interest rates.
'3.  No need to examine bond values for mispricing.

'Other risks involved in the pure yield pickup swap include:
'1.  Increased risk of call in the event of a decline in interest rates.
'2.  Reinvestment risk is larger with higher-coupon bonds.
'3.  The two bonds may have diferent yields because the market anticipates a
'credit rating change for one of them.

'Suppose the market is anticipating the P bond will be downgraded to an A
'rating.

'Suppose this occurs, and by the end of the one-year holding period the P
'bond has an A rating and is selling to yield 12.5%.  The return calculations
'for the P bond are given below.

'-------------------------------------------------------------------------------
'Last but not least, another way to perform a pure-yield pickup swap is to
'move to a longer-maturity bond.  With an upward sloping yield curve, longer
'maturity bonds will have higher yields than shorter maturity bonds.  The risk
'here is that if yields increase the longer term bond will have a larger price
'decline than the shorter maturity bond.  This strategy is sometimes called
'"riding the yield curve" and is more typically done with shorter term
'securities.  This strategy involves forecasting what the yield curve will look
'like at the end of the investment horizon.  The bet is that the yield curve will
'not change during the investment period.  Again, the risk is that rates will
'increase, which is what we might expect with an upward-sloping yield curve.

'Finally, for tax swaps you may wish to offset capital gains in
'one security by selling bonds with a capital loss (selling at a discount to
'your purchase price).  You would then swap into a bond with as nearly identical
'features as the sold bond as possible.  You use the capital loss on the sale
'for tax purposes and still maintain your current position in the market.
'-------------------------------------------------------------------------------

