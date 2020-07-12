Attribute VB_Name = "FINAN_DERIV_ARITHMETIC_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : GEOMETRIC_AVERAGE_RATE_OPTION_FUNC
'DESCRIPTION   : Geometric average rate option
'LIBRARY       : DERIVATIVES
'GROUP         : ARITHMETIC
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function GEOMETRIC_AVERAGE_RATE_OPTION_FUNC(ByVal SPOT As Double, _
ByVal AVG_SPOT As Double, _
ByVal STRIKE As Double, _
ByVal ORIGINAL_TENOR As Double, _
ByVal REMAINING_TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

'---------------------------------
'SPOT price
'Average price
'STRIKE price
'Original time to maturity
'Remaining time to maturity
'Risk-free rate
'Cost of carry
'Volatility
'---------------------------------
    
Dim TEMP_TENOR As Double 'Observed or realized time period
Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

On Error GoTo ERROR_LABEL

ATEMP_VAL = 1 / 2 * (CARRY_COST - SIGMA ^ 2 / 6)
BTEMP_VAL = SIGMA / Sqr(3)

TEMP_TENOR = ORIGINAL_TENOR - REMAINING_TENOR

If TEMP_TENOR > 0 Then
    STRIKE = (TEMP_TENOR + REMAINING_TENOR) / REMAINING_TENOR * STRIKE - _
    TEMP_TENOR / REMAINING_TENOR * AVG_SPOT
    
    GEOMETRIC_AVERAGE_RATE_OPTION_FUNC = GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, STRIKE, _
        REMAINING_TENOR, RATE, ATEMP_VAL, BTEMP_VAL, OPTION_FLAG, CND_TYPE) * _
        REMAINING_TENOR / (TEMP_TENOR + REMAINING_TENOR)

ElseIf TEMP_TENOR = 0 Then
    GEOMETRIC_AVERAGE_RATE_OPTION_FUNC = GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, STRIKE, _
        ORIGINAL_TENOR, RATE, ATEMP_VAL, BTEMP_VAL, OPTION_FLAG, CND_TYPE)
End If

Exit Function
ERROR_LABEL:
GEOMETRIC_AVERAGE_RATE_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : TW_ARITHMETIC_AVERAGE_APPROX_FUNC

'DESCRIPTION   : Turnbull and Wakeman arithmetic average approximation
'Extensive testing has shown that the unconditional
'Asian Monte Carlo simulation produces results equal to
'the Turnbull - Wakeman approximation within the random
'variation inherent in the simulation.

'LIBRARY       : DERIVATIVES
'GROUP         : ARITHMETIC
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function TW_ARITHMETIC_AVERAGE_APPROX_FUNC(ByVal SPOT As Double, _
ByVal AVG_SPOT As Double, _
ByVal STRIKE As Double, _
ByVal ORIGINAL_TENOR As Double, _
ByVal REMAINING_TENOR As Double, _
ByVal START_TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

'---------------------------------
' SPOT price
' Average price
' STRIKE price
' Original time to maturity
' Remaining time to maturity
' Time to start of average period
' Risk-free rate
' Cost of carry
' Volatility
'---------------------------------

Dim M1_VAL As Double
Dim M2_VAL As Double
    
Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

Dim TEMP_TENOR As Double

On Error GoTo ERROR_LABEL

M1_VAL = (Exp(CARRY_COST * ORIGINAL_TENOR) - Exp(CARRY_COST * START_TENOR)) / _
    (CARRY_COST * (ORIGINAL_TENOR - START_TENOR))

M2_VAL = 2 * Exp((2 * CARRY_COST + SIGMA ^ 2) * ORIGINAL_TENOR) / _
    ((CARRY_COST + SIGMA ^ 2) * (2 * CARRY_COST + SIGMA ^ 2) * _
    (ORIGINAL_TENOR - START_TENOR) ^ 2) + 2 * Exp((2 * CARRY_COST + _
    SIGMA ^ 2) * START_TENOR) / (CARRY_COST * (ORIGINAL_TENOR - _
    START_TENOR) ^ 2) * (1 / (2 * CARRY_COST + SIGMA ^ 2) - _
    Exp(CARRY_COST * (ORIGINAL_TENOR - START_TENOR)) / _
    (CARRY_COST + SIGMA ^ 2))

ATEMP_VAL = Log(M1_VAL) / ORIGINAL_TENOR
BTEMP_VAL = Sqr(Log(M2_VAL) / ORIGINAL_TENOR - 2 * ATEMP_VAL)
TEMP_TENOR = ORIGINAL_TENOR - REMAINING_TENOR

If TEMP_TENOR > 0 Then
    STRIKE = ORIGINAL_TENOR / REMAINING_TENOR * STRIKE - TEMP_TENOR / _
    REMAINING_TENOR * AVG_SPOT
    
    TW_ARITHMETIC_AVERAGE_APPROX_FUNC = GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, STRIKE, _
        REMAINING_TENOR, RATE, ATEMP_VAL, BTEMP_VAL, OPTION_FLAG, CND_TYPE) * _
        REMAINING_TENOR / ORIGINAL_TENOR
Else
    TW_ARITHMETIC_AVERAGE_APPROX_FUNC = GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, STRIKE, _
        REMAINING_TENOR, RATE, ATEMP_VAL, BTEMP_VAL, OPTION_FLAG, CND_TYPE)
End If

Exit Function
ERROR_LABEL:
TW_ARITHMETIC_AVERAGE_APPROX_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LEVY_ARITHMETIC_AVERAGE_APPROX_FUNC
'DESCRIPTION   : Levy's arithmetic average approximation
'LIBRARY       : DERIVATIVES
'GROUP         : ARITHMETIC
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function LEVY_ARITHMETIC_AVERAGE_APPROX_FUNC(ByVal SPOT As Double, _
ByVal AVG_SPOT As Double, _
ByVal STRIKE As Double, _
ByVal ORIGINAL_TENOR As Double, _
ByVal REMAINING_TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

    Dim TEMP_SPOT As Double
    Dim TEMP_SIGMA As Double
    Dim TEMP_STRIKE As Double
    
    Dim D1_VAL As Double
    Dim D2_VAL As Double
    Dim ATEMP_VAL As Double
    Dim BTEMP_VAL As Double


'---------------------------------
'SPOT price
'Average price
'STRIKE price
'Original time to maturity
'Remaining time to maturity
'Risk-free rate
'Cost of carry
'Volatility
'---------------------------------

    On Error GoTo ERROR_LABEL

    TEMP_SPOT = SPOT / (ORIGINAL_TENOR * CARRY_COST) * (Exp((CARRY_COST - _
        RATE) * REMAINING_TENOR) - Exp(-RATE * REMAINING_TENOR))
    
    ATEMP_VAL = 2 * SPOT ^ 2 / (CARRY_COST + SIGMA ^ 2) * ((Exp((2 * CARRY_COST + _
        SIGMA ^ 2) * REMAINING_TENOR) - 1) / (2 * CARRY_COST + _
        SIGMA ^ 2) - (Exp(CARRY_COST * REMAINING_TENOR) - 1) / CARRY_COST)

    BTEMP_VAL = ATEMP_VAL / (ORIGINAL_TENOR ^ 2)
    
    TEMP_SIGMA = Log(BTEMP_VAL) - 2 * (RATE * REMAINING_TENOR + Log(TEMP_SPOT))
    
    TEMP_STRIKE = STRIKE - (ORIGINAL_TENOR - REMAINING_TENOR) / _
        ORIGINAL_TENOR * AVG_SPOT
    
    D1_VAL = 1 / Sqr(TEMP_SIGMA) * (Log(BTEMP_VAL) / 2 - Log(TEMP_STRIKE))
    D2_VAL = D1_VAL - Sqr(TEMP_SIGMA)
    
        Select Case OPTION_FLAG
            Case 1 ', "CALL", "C"
        LEVY_ARITHMETIC_AVERAGE_APPROX_FUNC = TEMP_SPOT * CND_FUNC(D1_VAL, CND_TYPE) - _
            TEMP_STRIKE * Exp(-RATE * REMAINING_TENOR) * CND_FUNC(D2_VAL, CND_TYPE)
            
            Case Else '-1 ', "PUT", "P"
        
        LEVY_ARITHMETIC_AVERAGE_APPROX_FUNC = (TEMP_SPOT * CND_FUNC(D1_VAL, CND_TYPE) - _
            TEMP_STRIKE * Exp(-RATE * REMAINING_TENOR) * CND_FUNC(D2_VAL, _
            CND_TYPE)) - TEMP_SPOT + TEMP_STRIKE * Exp(-RATE * REMAINING_TENOR)
        End Select
    
Exit Function
ERROR_LABEL:
LEVY_ARITHMETIC_AVERAGE_APPROX_FUNC = Err.number
End Function

                    
'//////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////

'The unique characteristic of an average price option is that the underlying
'asset prices are averaged over some predefined time interval. This tends to
'dampen the volatility and therefore average price options are less expensive
'than standard options.

'Average price options are path-dependent The price path followed by the
'underlying asset is crucial to the pricing of the option.

'Types of average price options:
'1) Average Price European options using geometric averaging (Rubinstein).
'2) Average Price European options using arithmetic averaging (Levi).
'3) Average Strike European options using geometric averaging (Rubinstein).

'Average Price options are commonly used in the foreign exchange, interest rate
'and commodity markets.

'Intrinsic Value Formula

'The payoff for Average Price options is:
'Call: max {O, Ua E}
'Put: max {0, E Ua}

'Where Ua is the average underlying price and E is the exercise or strike price.

'The payoff for Average Strike options is:
'Call: max {O, U Ea}
'Put: max {0, Ea U}

'where U is the underlying price and E is the average underlying price
'applied as the exercise or strike price.

'Benefits
'A benefit of average price options is that they reduce incentives for
'manipulation of the underlying price at expiration.

'Average strike options are often used by a seller to place a floor on the
'selling price of a sequence of sales of an asset over some time horizon.

'Average strike options are cheaper than standard options.
'Average price options are useful in situations where the trader/hedger is
'concerned only about the average price of a commodity which they regularly
'purchase.

'Variations

'The averaging period can span the whole life of an option or some shorter
'period; options with an averaging period less than the whole life are called
'partial average options.

'The average is typically based on daily prices but could be based on weekly
'or monthly data. This monitoring frequency is defined in the contract.

'The average may be arithmetic mean (standard average), a weighted average,
'or a geometric mean.

'The conventional assumptions that are made in the theoretical models are:
'Frictionless markets, constant volatility, continuous interest rates and
'the underlier follow a diffusion process.

'Uses

'1) Suppose a nine-month European average price contract calls for a
'payoff equal to the difference between the average price of a barrel
'of crude oil and a fixed exercise price of USD18. The averaging period
'is the last two months of the contract. The impact of this contract
'relative to a standard option contract is that the volatility is
'dampened by the averaging of the crude oil price, and therefore the
'option price is lower. The holder gains protection from potential price
'manipulation or sudden price spikes.

'2) A Canadian exporting firm doing business in the U.S. is exposed to
'Can$/US$ foreign exchange risk every week. For budgeting purposes the
'treasurer must pick some average exchange rate in which to quote Can$
'cash flows (derived from US$ revenue) for the current quarter. Suppose
'the treasurer chooses an average FX rate of Can$1.29/US$1.00. If the
'US$ strengthens, the cash flows will be greater than estimated, but if
'it weakens, the company's Can$ cash flows are decreased.

'Solution:

'Time 0 mo.: FX = Can$1.33/US$1.00. Buy a put option for US$ using
'weekly average rates with a strike of Can$1.29/US$1.00.

'Time 3 mo.:

'Case 1: FX = Can$1.30/US$1.00., avg. = Can$1.2850/US$1.00.
'Option is in-the money; payoff = $1.290 - $1.2850 P*4,PA
'P*4,PA = premium for option purchase, future-valued 3 months

'Case 2: FX = Can$1.30/US$1.00. avg. = Can$1.2950/US$1.00.
'Option is out-of-the-money, but the company's Can$ cash flow is higher as well.

'Important note:
'Foreign exchange rates used must be consistent in units.
'avg.(1.28, 1.30, 1.32) Can$/US$ avg.(1/1.28, 1/1.30, 1/1.32) US$/Can$

'//////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////

