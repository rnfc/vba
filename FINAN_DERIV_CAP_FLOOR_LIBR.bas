Attribute VB_Name = "FINAN_DERIV_CAP_FLOOR_LIBR"

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : SWAPTION_PREMIUM_FUNC
'DESCRIPTION   : Calculate swaption premium using BS Model
'value in home currency
'LIBRARY       : FIXED_INCOME
'GROUP         : CAP_FLOOR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'**********************************************************************************
'**********************************************************************************

Function SWAPTION_PREMIUM_FUNC(ByVal TENOR As Double, _
ByVal COMPOUNDING As Integer, _
ByVal FORWARD As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal SIGMA_SWAP As Double, _
Optional ByVal OPTION_FLAG As Integer = 1)

Dim D1_VAL As Double
Dim D2_VAL As Double
    
'TENOR: Tenor of swap in years
'COMPOUNDING: Compoundings per year
'FORWARD: Underlying swap rate
'STRIKE: STRIKE price
'EXPIRATION: Time to maturity
'RATE: Risk-free rate
'SIGMA: Swap rate volatility

On Error GoTo ERROR_LABEL

D1_VAL = (Log(FORWARD / STRIKE) + SIGMA_SWAP ^ 2 / 2 * EXPIRATION) _
/ (SIGMA_SWAP * Sqr(EXPIRATION))

D2_VAL = D1_VAL - SIGMA_SWAP * Sqr(EXPIRATION)

Select Case OPTION_FLAG
    Case 1 ', "Call", "c" 'Payer SWAPTION_PREMIUM_FUNC
        
        SWAPTION_PREMIUM_FUNC = ((1 - 1 / (1 + FORWARD / COMPOUNDING) ^ _
        (TENOR * COMPOUNDING)) / FORWARD) * Exp(-RATE * EXPIRATION) * _
        (FORWARD * CND_FUNC(D1_VAL) - STRIKE * CND_FUNC(D2_VAL))
    
    Case -1 ', "Put", "p" 'Receiver SWAPTION_PREMIUM_FUNC
        
        SWAPTION_PREMIUM_FUNC = ((1 - 1 / (1 + FORWARD / COMPOUNDING) ^ _
        (TENOR * COMPOUNDING)) / FORWARD) * Exp(-RATE * EXPIRATION) * _
        (STRIKE * CND_FUNC(-D2_VAL) - FORWARD * CND_FUNC(-D1_VAL))
    
    Case Else
        GoTo ERROR_LABEL
End Select

Exit Function
ERROR_LABEL:
SWAPTION_PREMIUM_FUNC = Err.number
End Function
     
'************************************************************************************
'************************************************************************************
'FUNCTION      : CAP_CASH_FLOW_FUNC
'-----------------------------------------------------------------------------------
'DESCRIPTION   : Caps and floors trade over the counter, and are tools that
'companies can use to manage interest rate exposure, and that fixed income
'managers can use to make trading profits and manage their portfolio

'A cap can be associated with any underlying interest rate.  For the
'most part, they are linked to LIBOR. A cap must specify a cap rate--which
'is the strike price of the option, a notional principle, used to compute
'the cash flows, and a tenor, which indicates the term of the rate as
'well as the frequency of resets and payments.

'Example:
'    Cap on three-month LIBOR with three-month tenor
'    Notional Principal of $10 million
'    Cap Rate 4.78821%
'    Current Date 12/14/05
'    Valuation Date 12/16/05
'    Maturity: 12/16/10
'-----------------------------------------------------------------------------------
'LIBRARY       : FIXED_INCOME
'GROUP         : CAP_FLOOR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function CAP_CASH_FLOW_FUNC(ByVal RESET_DATE As Date, _
ByVal PAY_DATE As Date, _
ByVal NOTIONAL As Double, _
ByVal CAP_RATE As Double, _
ByVal SPOT_LIBOR As Double, _
Optional ByVal NUMER_COUNT_BASIS As Integer = 1, _
Optional ByVal DENOM_COUNT_BASIS As Integer = 2)

'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
'Suppose that on 9/16/06, spot 3-month LIBOR is 5.2%.
'  In this case, the cash flow on 12/16/06 will be calculated as follows:
'                Reset Date: 16/09/2006
'                Pay Date:    16/12/2006
'                Not Princ:   10000000
'                Cap Rate:    0.0478821
'                Days between:   91
'                Year Fraction: 0.252777778
'                Spot LIBOR:  0.052
'                Cash Flow:   10,409.14
'If on any of the reset dates, (Spot) 3-Mo LIBOR is less than the
'cap rate, the cash flow 3 months later is 0.
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------

On Error GoTo ERROR_LABEL

CAP_CASH_FLOW_FUNC = MAXIMUM_FUNC((SPOT_LIBOR - CAP_RATE), 0) * _
                            (COUNT_DAYS_FUNC(RESET_DATE, PAY_DATE, _
                            NUMER_COUNT_BASIS) _
                            / DAYS_PER_YEAR_FUNC(RESET_DATE, DENOM_COUNT_BASIS)) * NOTIONAL

'To value the cash flows that occur on future date.  We would use
'the discount factor from the Bloomberg "Curves" screen

Exit Function
ERROR_LABEL:
CAP_CASH_FLOW_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CAPLET_RESET_RATE_FUNC
'DESCRIPTION   : Note that this function corresponds to the Reset Rate
'used to value the caplet that resets on START_DATE --> SAME as BLOOMBERG
'LIBRARY       : FIXED_INCOME
'GROUP         : CAP_FLOOR
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function CAPLET_RESET_RATE_FUNC(ByVal SETTLEMENT As Date, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
ByVal START_RATE As Double, _
ByVal END_RATE As Double, _
Optional ByVal NUMER_COUNT_BASIS As Integer = 1, _
Optional ByVal DENOM_COUNT_BASIS As Integer = 2)

Dim D1_VAL As Double
Dim D2_VAL As Double
Dim FV_VAL As Double

'D1_VAL number of days from settlement date until the start date of
'   the forward period (I.e., the Reset Date).
'D2_VAL number of days from settlement date until the end date of the
'   forward period (I.e., the next Reset Date).
'r1 spot rate for D1_VAL days.
'r2 spot rate for D2_VAL days.
'FV = future value: FV = (1 + [(r2*D2_VAL)/360]) / (1 + [(r1*D1_VAL) / 360]))
'f = forward rate: [(FV - 1) / (D2_VAL-D1_VAL)] * 360

On Error GoTo ERROR_LABEL

D1_VAL = COUNT_DAYS_FUNC(SETTLEMENT, START_DATE, NUMER_COUNT_BASIS)
D2_VAL = COUNT_DAYS_FUNC(SETTLEMENT, END_DATE, NUMER_COUNT_BASIS)

FV_VAL = (1 + (END_RATE * D2_VAL / _
        DAYS_PER_YEAR_FUNC(SETTLEMENT, DENOM_COUNT_BASIS))) / _
          (1 + (START_RATE * D1_VAL / _
        DAYS_PER_YEAR_FUNC(SETTLEMENT, DENOM_COUNT_BASIS)))

CAPLET_RESET_RATE_FUNC = ((FV_VAL - 1) / (D2_VAL - D1_VAL)) * _
                     DAYS_PER_YEAR_FUNC(SETTLEMENT, DENOM_COUNT_BASIS)

Exit Function
ERROR_LABEL:
CAPLET_RESET_RATE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : BLACK_SCHOLES_CAP_FLOOR_FUNC
'-----------------------------------------------------------------------------------
'DESCRIPTION   :
'Black 's Model is widely used by Fixed Income Traders to characterize the
'value of caps. Black 's model was originally designed to ascertain the
'value of an option on a futures contract.

'In the Black-Scholes model, (or any finance model where there is no arbitrage),
'we simply take the expected value of the contract at expiration and discount
'this back to today.

'For convenience sake, we do this in the "Equivalent risk-neutral World" since
'only in this (equivalent) world do we know the discount factor without doing a
'lot of work (and that is the risk-free rate).

'In Black-Scholes, the numeraire is $1 today.
'To use Black's model, the numeraire is $1 at the time of delivery.  This
'trick allows us to handle the problem of having the bond price be a random
'variable at a future date, but assuming that the risk-free rate between now and
'then is constant.  Thus we not only do the valuation in our equivalent
'risk-neutral world, but we move to an equivalent risk-neutral forward world.

'In this world, the expected future spot rate is the forward rate, and its
'standard deviation is the same as in the "physical world."
'Black 's Model assumes that the future spot rate is lognormally distributed.
'-----------------------------------------------------------------------------------
'LIBRARY       : FIXED_INCOME
'GROUP         : CAP_FLOOR
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function BLACK_SCHOLES_CAP_FLOOR_FUNC(ByVal SETTLEMENT As Date, _
ByVal RESET_DATE As Date, _
ByVal NEXT_DATE As Date, _
ByVal FORW_RATE As Double, _
ByVal STRIKE As Double, _
ByVal DISCOUNT As Double, _
ByVal SIGMA As Double, _
Optional ByVal NOTIONAL As Double = 10000000, _
Optional ByVal NUMER_COUNT_BASIS As Integer = 1, _
Optional ByVal DENOM_COUNT_BASIS As Integer = 4, _
Optional ByVal VERSION As Integer = 2, _
Optional ByVal CND_TYPE As Integer = 0)

Dim D1_VAL As Double
Dim D2_VAL As Double

Dim RESET_TENOR As Double
Dim NEXT_TENOR As Double

Dim FIRST_FACTOR As Double
Dim SECOND_FACTOR As Double

On Error GoTo ERROR_LABEL

RESET_TENOR = COUNT_DAYS_FUNC(SETTLEMENT, RESET_DATE, NUMER_COUNT_BASIS) / _
              DAYS_PER_YEAR_FUNC(SETTLEMENT, DENOM_COUNT_BASIS)

NEXT_TENOR = COUNT_DAYS_FUNC(RESET_DATE, NEXT_DATE, NUMER_COUNT_BASIS) / _
              DAYS_PER_YEAR_FUNC(RESET_DATE, DENOM_COUNT_BASIS)

FIRST_FACTOR = Log(FORW_RATE / STRIKE) + (SIGMA ^ 2 / 2) * RESET_TENOR
SECOND_FACTOR = SIGMA * Sqr(RESET_TENOR)
    
D1_VAL = FIRST_FACTOR / SECOND_FACTOR
D2_VAL = (Log(FORW_RATE / STRIKE) + (-SIGMA ^ 2 / 2) * RESET_TENOR) / _
         SECOND_FACTOR
    
Select Case VERSION
'--------------------------------------------------------------------------------------
Case 0 'Cap Function
'--------------------------------------------------------------------------------------
'FORWARD_RATE: Mean Rate
'DISCOUNT: Discount Rate
'STRIKE: Cap Rate

'SIGMA: this means that for valuing this caplet, the probability description
'of the 3-month (for example) LIBOR spot rate, x (reset day) months from
'settlement is lognormal, with mean and standard deviation
'computed as follows:

'a) Mean of log rate: Ln(FORWARD RATE)-0.5 *
'                    (SIGMA^2 * (RESET-SETTLEMENT)/YEAR FACTOR)
'b) std dev of log rate: (SIGMA^2 * (RESET-SETTLEMENT)/YEAR FACTOR)
'c) cumul prob distribution (CDF): LOGNORMDIST(dRATE,MEAN,STD DEV)
'd) PDF (y-axis): delta CDF (X2-X1)

'Example:
'Cap Rate: 0.0478821
'Discount: 0.95307
'Mean rate (= forward rate): 0.0483516
'Std dev (1-Yr): 0.118
'Term (Years): 0.75 = (RESET - SETTLEMENT) / 360

'today:       16 /12/2005
'reset date:  14/09/2006
'pay date:    18/12/2006
'next expiry: 14/12/2006


'Let me stress that this is not the market's distribution of the 90-Day (for example)
'spot LIBOR in 9 months (x reset days from settlement).  Instead it is that as it
'relates to valuing an option (the forward risk neutral world).

'Black 's Model values each caplet in a distinct forward risk-neutral world.
'As is the case in the Black-Scholes Model, we take the discounted expected cash
'flows from the option.  The difference is that Black-Scholes uses the equivalent
'risk-neutral world, whereas Black 's model uses the forward risk-neutral world
'corresponding to each caplet.

'The intuition behind this value is as follows:
'1) What is the equivalent risk-neutral probability that we would exercise
'   this caplet?

'   The strike rate is 0.0478821 (cap rate) which is less than the expected
'   3-month rate of 0.0483516 (in the forward risk-neutral world). We look at
'   the pdf graph and ask for the probability that 90-LIBOR will exceed
'   the strike rate CND_FUNC(D2_VAL). So in this case, there's a 51.8%  probability
'   that this caplet will be in the money when it expires.

'2) What is the expected value under the truncated distribution--where we only
'   look at the portion above the exercise price. 0.558267911 is the probability
'   associated with this truncated mean CND_FUNC(D1_VAL).

'   So what we expect to get if we exercise: 0.026993147 =(FORWARD_RATE * CND_FUNC(D1_VAL)
'   what this costs: 0.024788684 =(CAP RATe * CND_FUNC(D2_VAL)
'   Expected $ value (in 1 Year) 5511.157906 =(0.026993147-0.024788684)*10000000/4

'3) Bring this expected value back to today.
'   5252.519266  (5511.157906*DISCOUNT--> ignoring day-count issues)

'For the most part, the market does not use Black's model to value caps and
'floors, so much as it uses this model as a tool to characterize the prices.
'In particular, prices are often expressed as implied volatilities.  Notice
'on the Bloomberg valuation screen that the implied volatilities are getting
'higher as we move out in time.

'As in the valuation screen there is a (potentially) different implied vol for
'each caplet.  This is a method sometimes called spot volatilities.
'We could, of course solve for a cap value by forcing all caplet vols to be the
'same.  Such a situation is called flat volatilities.

'-----------------------------------------------------------------------------------
        BLACK_SCHOLES_CAP_FLOOR_FUNC = NOTIONAL * NEXT_TENOR * DISCOUNT * _
                        ((FORW_RATE * CND_FUNC(D1_VAL)) - (STRIKE * CND_FUNC(D2_VAL)))
'-----------------------------------------------------------------------------------
Case 1 'Floor Function
'Floors work analogously to caps.  We can think of a caplet as a call option
'on the future spot rate, the floorlet is an analogous put option
'-----------------------------------------------------------------------------------
        BLACK_SCHOLES_CAP_FLOOR_FUNC = NOTIONAL * NEXT_TENOR * DISCOUNT * _
                    ((STRIKE * CND_FUNC(-D2_VAL)) - (FORW_RATE * CND_FUNC(-D1_VAL)))
'The delta of a floorlet is -N(-D1_VAL): -CND_FUNC(-1 * (FIRST_FACTOR / SECOND_FACTOR))
'-----------------------------------------------------------------------------------
Case Else 'Value of the Swap.
        BLACK_SCHOLES_CAP_FLOOR_FUNC = (NOTIONAL * NEXT_TENOR * DISCOUNT * _
                        ((FORW_RATE * CND_FUNC(D1_VAL)) - _
                        (STRIKE * CND_FUNC(D2_VAL)))) - (NOTIONAL * _
                        NEXT_TENOR * DISCOUNT * _
                        ((STRIKE * CND_FUNC(-D2_VAL)) - (FORW_RATE * CND_FUNC(-D1_VAL))))
'Remember that the  "Swap Rate" is that rate which
'such a swap holder pays fixed makes the value of the swap 0
'As such one way to solve for the swap rate would be to use
'Solver to identify such a strike price--> WHICH OF COURSE IS THE FORWARD RATE

'-----------------------------------------------------------------------------------
'An interest rate swap is a portfolio of forward rate agreements.
'A forward rate agreement (FRA) entails the exchange of a pre-determined
'interest rate for the market rate at a pre-sepcified date, on a pre-specified
'notional principal.

'A FRA is valued by assuming that the current market forward rate will be
'realized at at swap date.

'    FRAs and swaps are priced to have 0 value at origination.  The "price"
'    is the pre-specified fixed rate that will be exchanged for the market
'    rate in the future.

'    In the previous example, we saw that the one-period swap (I.e., a FRA)
'    will have a price equal to the relevant forward rate.
'Thus, swap valuation entails the following steps:
'  1) Ascertain the relevant LIBOR forward rates,
'  2) Calculate the cash flows that will accrue -- assuming that
'    the realized future floating rates equal relevant the forward rates.
'  3) The value of the swap is the PV of these cash flows.  (So at
'    origination, the fixed "swap rate" will be the rate that makes
'    the PV of these flows equal to 0.)
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
BLACK_SCHOLES_CAP_FLOOR_FUNC = Err.number
End Function



'************************************************************************************
'************************************************************************************
'FUNCTION      : CAPLET_FORWARD_RATE_VOLATILITIES_FUNC
'-----------------------------------------------------------------------------------
'DESCRIPTION   : This function demonstrates calculation of forward rate
'volatilities from spot rate volatilities of caplets
'implied from black model

'Reference : Example 24.1 in Options.Futures and Other
'derivatives - John C. Hull
'-----------------------------------------------------------------------------------
'LIBRARY       : FIXED_INCOME
'GROUP         : CAP_FLOOR
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function CAPLET_FORWARD_RATE_VOLATILITIES_FUNC(ByRef TENOR_RNG As Variant, _
ByRef CAPLET_RNG As Variant)

Dim i As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim TEMP_MULT As Double

Dim TEMP_VECTOR As Variant
Dim TENOR_VECTOR As Variant
Dim CAPLET_VECTOR As Variant

'CAPLET_RNG: Caplet volatiity

On Error GoTo ERROR_LABEL

TENOR_VECTOR = TENOR_RNG
If UBound(TENOR_VECTOR, 1) = 1 Then: _
    TENOR_VECTOR = MATRIX_TRANSPOSE_FUNC(TENOR_VECTOR)

CAPLET_VECTOR = CAPLET_RNG
If UBound(CAPLET_VECTOR, 1) = 1 Then: _
    CAPLET_VECTOR = MATRIX_TRANSPOSE_FUNC(CAPLET_VECTOR)

If UBound(TENOR_VECTOR, 1) <> UBound(CAPLET_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(TENOR_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)

TEMP_SUM = 0
TEMP_MULT = 0

TEMP_VECTOR(1, 1) = CAPLET_VECTOR(1, 1)

For i = 2 To NROWS
'---------------------------------------------------------------------------------
    TEMP_MULT = TEMP_VECTOR(i - 1, 1) * TEMP_VECTOR(i - 1, 1)
    TEMP_SUM = TEMP_SUM + TEMP_MULT
'---------------------------------------------------------------------------------
    TEMP_VECTOR(i, 1) = (TENOR_VECTOR(i, 1) * CAPLET_VECTOR(i, 1) * _
                        CAPLET_VECTOR(i, 1) - TEMP_SUM) ^ 0.5
'---------------------------------------------------------------------------------
Next i

CAPLET_FORWARD_RATE_VOLATILITIES_FUNC = TEMP_VECTOR 'Forward rate volatity

Exit Function
ERROR_LABEL:
CAPLET_FORWARD_RATE_VOLATILITIES_FUNC = Err.number
End Function
