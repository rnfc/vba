Attribute VB_Name = "FINAN_DERIV_SPREAD_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : SPREAD_OPTION_FUNC
'DESCRIPTION   : Spread option approximation
'LIBRARY       : DERIVATIVES
'GROUP         : SPREAD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'**********************************************************************************
'**********************************************************************************

Function SPREAD_OPTION_FUNC(ByVal FUTURES_A As Double, _
ByVal FUTURES_B As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal SIGMA_A As Double, _
ByVal SIGMA_B As Double, _
ByVal RHO_VAL As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

Dim TEMP_SIGMA As Double
Dim TEMP_FUTURES As Double

On Error GoTo ERROR_LABEL

TEMP_SIGMA = Sqr(SIGMA_A ^ 2 + (SIGMA_B * FUTURES_B / (FUTURES_B + STRIKE)) _
^ 2 - 2 * RHO_VAL * SIGMA_A * SIGMA_B * FUTURES_B / (FUTURES_B + STRIKE))
TEMP_FUTURES = FUTURES_A / (FUTURES_B + STRIKE)

SPREAD_OPTION_FUNC = GENERALIZED_BLACK_SCHOLES_FUNC(TEMP_FUTURES, 1, _
EXPIRATION, RATE, 0, TEMP_SIGMA, OPTION_FLAG, CND_TYPE) * (FUTURES_B + STRIKE)
    
Exit Function
ERROR_LABEL:
SPREAD_OPTION_FUNC = Err.number
End Function

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : EXTREME_SPREAD_OPTION_FUNC
'DESCRIPTION   : Extreme spread options
'LIBRARY       : DERIVATIVES
'GROUP         : SPREAD-OPTIONS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'**********************************************************************************
'**********************************************************************************

Function EXTREME_SPREAD_OPTION_FUNC(ByVal SPOT As Double, _
ByVal MIN_SPOT As Double, _
ByVal MAX_SPOT As Double, _
ByVal FIRST_TENOR As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

'------------------------------------------------------
'MIN_SPOT: Observed minimum
'MAX_SPOT: Observed Maximum
'FIRST_TENOR: First time period
'EXPIRATION: Time to maturity
'------------------------------------------------------
'[OPTION_FLAG: 1] Extreme spread call
'[OPTION_FLAG: 2] Extreme spread put
'[OPTION_FLAG: 3] Reverse extreme spread call
'[OPTION_FLAG: 4] Reverse extreme spread put
'------------------------------------------------------

        Dim kk As Long
        Dim ll As Long
        
        Dim ATEMP_VAL As Double
        Dim BTEMP_VAL As Double
        Dim CTEMP_VAL As Double
        Dim TEMP_SPOT As Double
        
        On Error GoTo ERROR_LABEL
        
        Select Case OPTION_FLAG
        Case 1, 3
            kk = 1
        Case Else
            kk = -1
        End Select
            
        Select Case OPTION_FLAG
        Case 1, 2
            ll = 1
        Case Else
            ll = -1
        End Select
            
        If (ll * kk) = 1 Then
            TEMP_SPOT = MAX_SPOT
        ElseIf (ll * kk) = -1 Then
            TEMP_SPOT = MIN_SPOT
        End If
        
        ATEMP_VAL = CARRY_COST - SIGMA ^ 2 / 2
        BTEMP_VAL = ATEMP_VAL + SIGMA ^ 2
        CTEMP_VAL = Log(TEMP_SPOT / SPOT)
        
'-----------------------------------------------------------------------------------------
'----------------------------------Extreme Spread Option
'-----------------------------------------------------------------------------------------
        Select Case ll
        Case 1
            EXTREME_SPREAD_OPTION_FUNC = kk * (SPOT * Exp((CARRY_COST - _
            RATE) * EXPIRATION) * (1 + SIGMA ^ 2 / (2 * CARRY_COST)) * _
            CND_FUNC(kk * (-CTEMP_VAL + BTEMP_VAL * EXPIRATION) / (SIGMA * _
            Sqr(EXPIRATION)), CND_TYPE) _
            - Exp(-RATE * (EXPIRATION - FIRST_TENOR)) * SPOT * Exp((CARRY_COST _
            - RATE) * EXPIRATION) * (1 + SIGMA ^ 2 / (2 * CARRY_COST)) * _
            CND_FUNC(kk * (-CTEMP_VAL + BTEMP_VAL * FIRST_TENOR) / (SIGMA * _
            Sqr(FIRST_TENOR)), CND_TYPE) _
            + Exp(-RATE * EXPIRATION) * TEMP_SPOT * CND_FUNC(kk * (CTEMP_VAL - ATEMP_VAL * _
            EXPIRATION) / (SIGMA * Sqr(EXPIRATION)), CND_TYPE) - _
            Exp(-RATE * EXPIRATION) * _
            TEMP_SPOT * SIGMA ^ 2 / (2 * CARRY_COST) * Exp(2 * ATEMP_VAL * CTEMP_VAL / SIGMA ^ 2) * _
            CND_FUNC(kk * (-CTEMP_VAL - ATEMP_VAL * EXPIRATION) / _
            (SIGMA * Sqr(EXPIRATION)), CND_TYPE) _
            - Exp(-RATE * EXPIRATION) * TEMP_SPOT * CND_FUNC(kk * (CTEMP_VAL - ATEMP_VAL * _
            FIRST_TENOR) / (SIGMA * Sqr(FIRST_TENOR)), CND_TYPE) + _
            Exp(-RATE * EXPIRATION) _
            * TEMP_SPOT * SIGMA ^ 2 / (2 * CARRY_COST) * Exp(2 * _
            ATEMP_VAL * CTEMP_VAL / SIGMA ^ 2) * _
            CND_FUNC(kk * (-CTEMP_VAL - ATEMP_VAL * FIRST_TENOR) / _
            (SIGMA * Sqr(FIRST_TENOR)), CND_TYPE))

'-----------------------------------------------------------------------------------------
'----------------------------------Reverse Extreme Spread Option
'-----------------------------------------------------------------------------------------
        Case Else '-1
            EXTREME_SPREAD_OPTION_FUNC = -kk * (SPOT * Exp((CARRY_COST - RATE) * _
            EXPIRATION) * (1 + SIGMA ^ 2 / (2 * CARRY_COST)) * CND_FUNC(kk * _
            (CTEMP_VAL - BTEMP_VAL * EXPIRATION) / (SIGMA * Sqr(EXPIRATION)), _
            CND_TYPE) + Exp(-RATE * _
            EXPIRATION) * TEMP_SPOT * CND_FUNC(kk * (-CTEMP_VAL + ATEMP_VAL * EXPIRATION) / _
            (SIGMA * Sqr(EXPIRATION)), CND_TYPE) - Exp(-RATE * EXPIRATION) * _
            TEMP_SPOT * SIGMA ^ 2 _
            / (2 * CARRY_COST) * Exp(2 * ATEMP_VAL * CTEMP_VAL / SIGMA ^ 2) * CND_FUNC(kk * _
            (CTEMP_VAL + ATEMP_VAL * EXPIRATION) / (SIGMA * Sqr(EXPIRATION)), _
            CND_TYPE) - SPOT * _
            Exp((CARRY_COST - RATE) * EXPIRATION) * (1 + SIGMA ^ 2 / (2 * _
            CARRY_COST)) * CND_FUNC(kk * (-BTEMP_VAL * (EXPIRATION - FIRST_TENOR)) / _
            (SIGMA * Sqr(EXPIRATION - FIRST_TENOR)), CND_TYPE) - Exp(-RATE * _
            (EXPIRATION - _
            FIRST_TENOR)) * SPOT * Exp((CARRY_COST - RATE) * EXPIRATION) * (1 - _
            SIGMA ^ 2 / (2 * CARRY_COST)) * CND_FUNC(kk * (ATEMP_VAL * (EXPIRATION - _
            FIRST_TENOR)) / (SIGMA * Sqr(EXPIRATION - FIRST_TENOR)), CND_TYPE))
        End Select

Exit Function
ERROR_LABEL:
EXTREME_SPREAD_OPTION_FUNC = Err.number
End Function

'///////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////

'About Spread Options

'Description
'A spread option will have a payoff equal to the difference between the
'prices of two assets and a fixed exercise (strike) price.

'Basis risk results in the need for spread options. Basis risk is defined
'as the differential in pricing between closely-linked commodities.

'The salient Rates that must be considered when valuing spread options are
'the volatility of each asset and their price correlation.

'Benefits

'Crack spread options are excellent tools for oil refiners for managing risks:
'The risk of buying crude efficiently at prices which sustain margins.
'The risk of not being able to control the margin between crude purchases
'and product sales.
'The risk of selling gasoline and heating oil in competitive wholesale and retail
'distribution markets where demand is highly variable, and driven by seasonal
'Rates such as weather.

'Features
'Correlation is a measure of how the change in one asset price is reflected in the
'other asset price. Gasoline and crude oil prices, for example, should be more
'correlated than gasoline and, say, silver.

'The implied correlation in the market price of the spread option is more important
'than the individual implied volatility of either asset.

'A trader will want to buy high correlation and sell weak correlation. As the
'correlation decreases, the price of the spread option increases.

'Intrinsic Value Formula
'The contract payoff is:
'Max [ (U2 - U1) - E, 0 ]
'Where U1 and U2 are the underlying asset prices and E is the exercise or strike price.

'Aliases
'Spread options are also known as outperformance options.

'Uses
'NYMEX introduced two spread option contracts in October 1994. The crack spread
'contracts are on the spread between gasoline and crude oil, and heating oil and crude oil.

'Some of the contract specs are:
'Quotes:
'Quoted price in $X.XX / bbl (same as underlying asset)
'Minimum tick $0.01 / bbl or $10.00 / point
'Strikes:
'At-the-money strike is product/crude spread rounded to nearest $0.25/bbl
'Strikes in 5 $0.25 increments above and below
'One additional strike price at next whole $1.00 increment above
'Two additional strikes at $2.00 increments above

'No negative strikes

'Listings: Listed for six consecutive months + two quarterly months
'Expiry: Friday before futures expiration (or 2nd Friday before if the
'first Friday is less than 3 business days prior to futures expiration)
'On 10/7/94 (first trading day):
'Heating oil: $0.5150/gal.; 42 gal/bbl
'Crude oil: $17.61
'Underlying spread: [0.5150*42-17.61] = $4.02
'At-the-money: $4.00
'Out-of-the-money: $4.25, $4.50, $4.75, $5.00, $5.25
'In-the-money: $3.75, $3.50, $3.25, $3.00, $2.75
'Deep out-of-the-money: $6.00, $8.00, $10.00
'Listed months: December 94 - May 95, June 95, September 95
'Exercise:
'STRIKE 4#
'Crude settlement 17.61
'Settlement price 21.61
'/ 42 gal $0.51452
'Round to nearest $0.005 $0.515
'Product futures leg $0.515
'X 42 gal $21.63
'Less: strike ($4.00)
'$17.63/bbl
'The Heat Crack Spread contract:

'Call: The right to hold a long position in the underlying Heating Oil
'Futures contract with a short position in the underlying Crude Oil Futures contract.

'Put: The right to hold a short position in the underlying Heating Oil Futures
'contract with a long position in the underlying Crude Oil Futures contract.

'Put and call options on the 1:1 futures price differential between NY Harbor
'heating oil and WTI crude.

'Spread calculated by heating oil futures price (42 gallons per bbl) - crude
'oil for same trading month.

'The Gasoline Crack Spread contract:

'Call: The right to hold a long position in the underlying Gasoline Futures
'contract with a short position in the underlying Crude Oil Futures contract.

'Put: The right to hold a short position in the underlying Gasoline Futures
'contract with a long position in the underlying Crude Oil Futures contract.

'Put and call options on the 1:1 futures price differential between NY Harbor
'unleaded gasoline (RFG) and WTI crude.

'Spread calculated by gasoline futures price (42 gallons per bbl) - crude oil
'for same trading month.

'Refer to the NYMEX Crack Spread Option Workbook for excellent examples of the
'use of crack spread options for:

'Independent Refiners

'Activity (position) in cash market:
'Long cracks: short crude/long products

'Hedging applications with crack options:
'Buy puts, sell call; collars, synthetic puts
'Product Traders and Bulk Marketers

'Activity (position) in cash market:
'Short cracks: short products to customers in pipeline markets and using WTI crude to hedge

'Hedging applications with crack options:
'Buy calls to offset short crack risk
'Unbounded Gasoline Distributors (downstream)

'Activity (position) in cash market:
'Short gasoline purchases vs. major retail companies

'Hedging applications with crack options:
'Buy out-of-the money calls to protect retail gasoline margins in the driving season.

'///////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////

