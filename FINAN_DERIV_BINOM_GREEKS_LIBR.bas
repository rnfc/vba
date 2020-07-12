Attribute VB_Name = "FINAN_DERIV_BINOM_PRICING_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1      'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function OPTION_BINOMIAL_TRINOMIAL_GREEKS_FUNC( _
ByRef TICKERS_RNG As Variant, _
ByRef PREMIUM_RNG As Variant, _
ByRef ASSET_PRICE_RNG As Variant, _
ByRef STRIKE_PRICE_RNG As Variant, _
ByRef RISK_FREE_RATE_RNG As Variant, _
ByRef DIVIDEND_YIELD_RNG As Variant, _
ByRef VOLATILITY_RNG As Variant, _
ByRef EXERCISE_DATE_RNG As Variant, _
Optional ByRef VALUATION_DATE_RNG As Variant = 0, _
Optional ByRef DTHETA_DATE_RNG As Variant, _
Optional ByRef DVEGA_VOLAT_RNG As Variant, _
Optional ByRef DRHO_RATE_RNG As Variant, _
Optional ByRef CONTRACTS_RNG As Variant = 100, _
Optional ByRef FEES_PAID_RNG As Variant = 1, _
Optional ByRef OPTION_TYPE_RNG As Variant = 1, _
Optional ByRef POSITION_TYPE_RNG As Variant = -1, _
Optional ByRef EXERCISE_TYPE_RNG As Variant = 0, _
Optional ByRef VALUATION_TYPE_RNG As Variant = 0, _
Optional ByRef NSTEPS_RNG As Variant = 150, _
Optional ByRef TDAYS_PER_YEAR_RNG As Variant = 252, _
Optional ByVal SCALE_FACTOR_RNG As Variant = 0.01, _
Optional ByRef CND_TYPE_RNG As Variant = 0)

'This function lets you enter combinations of options positions so you can
'view the payoff/risk graphs.

'If you know what the option is trading at, you can enter the market price in
'the "override premium" row to calculate the P&L of the position after taking
'into account the actual price of the option.

'This function will show the current Greek position values of your option
'strategy while referencing the a range of underlying prices.

'---------------------------------------------------------------------------------
'------------------Key Notes for Trading Strategies Involving Options
'---------------------------------------------------------------------------------

'"   High volatility associated with stock-market bottoms offers option traders
'tremendous profit potential if the correct option trading setups are deployed;
'however, many traders are familiar with only option buying (long) strategies,
'which unfortunately do not work very well in an environment of high volatility.
 
'"   Buying strategies - even those using bull and bear debit spreads - are
'generally poorly priced when there is high implied volatility. When a bottom
'is finally achieved, the collapse in high-priced options following a sharp
'drop in implied volatility strips away much of the profit potential. So even
'if you are correct in timing a market bottom, there may be little to no gain
'from a big reversal move following a capitulation sell-off.

'"   Through a net options selling approach, there is a way around this problem.
'Below i present a simple strategy that profits from falling volatility, offers
'a potential for profit regardless of market direction and requires little up-front
'capital if used with options on futures, for which there is much better margin
'treatment for option writing strategies.

'"   Trying to pick a bottom is hard enough even for savvy market technicians. Oversold
'indicators can remain oversold for a long time, and the market can continue to trade
'lower than expected. The decline in the broad equity-market measures in the summer of
'2002 offers a case in point. Many momentum indicators and some sentiment indicators
'were flashing buy signals well before we pivoted off July's lows. The correct option
'selling strategy, however, can make trading a market bottom considerably easier.

'The strategy i present below has little or no downside risk, thus eliminating the
'bottom-picking dilemma. This strategy also offers plenty of upside profit potential
'if the market experiences a solid rally once you are in your trade. More important,
'though, is the added benefit that comes with a sharp drop in implied volatility,
'which typically accompanies a capitulation reversal day and a follow-through
'multi-week rally. By getting short volatility, or short vega, the strategy thus
'offers an additional dimension for profit.

'---------------------------------------------------------------------------------
'----------------------------------Shorting Vega
'---------------------------------------------------------------------------------

'During the decline of summer 2002, the VIX, a measure of implied volatility of
'S&P 100 options, reached well over 50, which had not been seen since the crash
'of 1987. A high VIX means that options have become extremely expensive because
'of increased expected volatility, which gets priced into options. This presents
'a dilemma for buyers of options - whether of puts or calls - because the price
'of an option is so affected by implied volatility that it leaves traders long
'vega just when they should be short vega.

'Vega is a measure of how much an option price changes with a change in implied
'volatility. If, for instance, implied volatility drops to normal levels from
'extremes and the trader is long options (hence long vega), an option's price can
'decline even if the underlying moves in the intended direction.

'When there are high levels of implied volatility, selling options is, therefore,
'the preferred strategy, particularly since it can leave you short vega and thus
'able to profit from an imminent drop in implied volatility; however, it is possible
'for implied volatility to go higher (especially if the market goes lower), which
'leads to potential losses from still higher volatility. By deploying a selling
'strategy when implied volatility is at extremes compared to past levels, we can at
'least attempt to minimize this risk.

'"   Reverse Calendar Spreads

'To capture the profit potential created by wild market reversals to the upside and
'the accompanying collapse in implied volatility from extreme highs, the one strategy
'that works the best is called a 'reverse call calendar spread'.

'Normal calendar spreads are neutral strategies, involving selling a near-term option
'and buying a longer-term option, usually at the same strike price. The idea here is
'to have the market stay confined to a range so that the near-term option, which has
'a higher theta (the rate of time-value decay), will lose value more quickly than the
'long-term option. Typically, the spread is written for a debit (maximum risk). But
'another way to use calendar spreads is to reverse them - buying the near-term and
'selling the long-term, which works best when volatility is very high.

'The reverse calendar spread is not neutral and can generate a profit if the underlying
'makes a huge move in either direction. The risk lies in the possibility of the
'underlying going nowhere, whereby the short-term option loses time-value more quickly
'than the long-term option, which leads to a widening of the spread, exactly what is
'desired by the neutral calendar spreader. Having covered the concept of a normal and
'reverse calendar spread, let's apply the latter to S&P call options.

'At volatile market bottoms, the underlying is least likely to remain stationary over
'the near-term, which is an environment in which i like to use reverse calendar
'spreads; furthermore, there is a lot of implied volatility to sell, which, as
'mentioned above, adds profit potential.

'A reverse calendar spreads offers an excellent low-risk (provided you close the
'position before expiration of the shorter-term option) trading setup that has profit
'potential in both directions. This strategy, however, profits most from a market
'that is moving fast to the upside associated with collapsing implied volatility.
'The ideal time for deploying reverse call calendar spreads is, therefore, at or just
'following stock market capitulation, when huge moves of the underlying often occur
'rather quickly. Finally, the strategy requires very little upfront capital, which
'makes it attractive to traders with smaller accounts.

'-------------------------------------------------------------------------------
'-----------------------------------REFERENCES----------------------------------
'-------------------------------------------------------------------------------

'Barone-Adesi, G. & Whaley, R. (1987), `Efficient analytic approximation of
'American option values', Journal of Finance 42(2), 301-320.

'Bjerksund, P. & Stensland, G. (1993), `Closed form approximation of American
'options', Scandinavian Journal of Management 9, 87-99.

'Bjerksund, P. & Stensland, G. (2002), `Closed form valuation of American
'options'. http://www.nhh.no

'Bjork, T. (1998), Arbitrage Theory in Continuous Time, Oxford University Press.

'Black, F. & Scholes, M. (1973), `The pricing of options and corporate
'liabilities', Journal of Political Economy 73(May-June), 637-659.

'Bos, M. & Vandermark, S. (September 2002), `Finessing fixed dividends',
'Risk 15(9), 157-158.

'Cox, J., Ross, S. & Rubinstein, M. (1979), `Option pricing: a simplified
'approach', Journal of Financial Economics 7, 229-264.

'Frishling, V. (2002), `A discrete question', Risk January(1).

'Geske, R. (1979), `A note on an analytic formula for unprotected American
'call options on stocks with known dividends', Journal of Financial
'Economics 7(A), 375-380.

'Gray Stephen F. and Robert E. Whaley. “Reset put options: Valuation, Risk
'Characteristics and Excel.Application.” Australian Journal of Management 24 1
'(June 1999): 1-20.

'Gray Stephen F. and Robert E. Whaley, “Valuing S&P 500 Bear Market Warrants with a
'Periodic Reset.” Journal of Derivatives 5, 1 (Fall 1997): 99-106.

'Haug, E. G., Haug, J. & Lewis, A. (2003), `Back to basics: a new approach to
'the discrete dividend problem', WILMOTT Magazine September, 37-47.

'Roll, R. (1977), `An analytic valuation formula for unprotected American
'call options on stocks with known dividends', Journal of Financial
'Economics 5, 251-258.

'Whaley, R. (1981), `On the valuation of American call options on stocks
'with known dividends', Journal of Financial Economics 9, 207-211


'-------------------------------------------------------------------------------
'----------------------------RECOMMENDED BOOKS----------------------------------
'-------------------------------------------------------------------------------


'A Quick algorithm for Pricing European Average Options, Turnbull, S.M. and
'Wakeman, L.M., Journal of Financial and Quantitative Analysis 26, 377-389.

'Complex Derivatives, Erik Banks, Probus Publishing, Chicago, 1994.

'From Black-Scholes to Black Holes: New Frontiers in Options, Risk/Finex, London, 1992.

'New York Mercantile Exchange Crack Spread Options Workbook, 1994.

'Options, Futures and Other Derivative Securities, John Hull, Prentice-Hall,
'Englewood Cliffs, NJ, 1993.

'Options Markets, John Cox and Mark Rubinstein, Prentice-Hall, Englewood Cliffs,
'NJ, 1985.

'Options on Foreign Exchange, David F. DeRosa, Probus Publishing Company,
'Chicago, 1992.

'The Handbook of Exotic Options, Israel Nelken, ed., Irwin Professional
'Publishing, Chicago, 1996.

'Theory of Rational Option Pricing, Robert Merton, Bell Journal of Economics,
'4 (Spring 1973), 141-83.

'---------------------------------------------------------------------------------------
'http://www.optiontradingtips.com/
'---------------------------------------------------------------------------------------

'OPTION_TYPE = 1 --> CALL_OPTION
'OPTION_TYPE = -1 --> PUT_OPTION

'EXERCISE_TYPE = 0 --> Euro
'EXERCISE_TYPE = 1 --> Amer

'VALUATION_TYPE = 0 --> Black Scholes
'VALUATION_TYPE = 1 --> Binomial
'VALUATION_TYPE = 2 --> Trinomial

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim HEADINGS_STR As String
Dim PRICING_MODEL_STR As String

Dim TF_VAL As Double
Dim S_VAL As Double
Dim K_VAL As Double
Dim T_VAL As Double
Dim RF_VAL As Double
Dim DY_VAL As Double
Dim V_VAL As Double
Dim N_VAL As Long
Dim Q_VAL As Integer
Dim ET_VAL As Integer
Dim PT_VAL As Integer
Dim OT_VAL As Integer
Dim VT_VAL As Long
Dim SF_VAL As Double
Dim CND_VAL As Integer
Dim DTHETA_VAL As Double
Dim DVEGA_VAL As Double
Dim DRHO_VAL As Double
Dim P_VAL As Double

Dim D_VAL As Double
Dim EDATE_VAL As Date
Dim VDATE_VAL As Date

Dim TICKER_STR As String

Dim TICKERS_VECTOR As Variant
Dim PREMIUM_VECTOR As Variant 'Option Market Price
Dim ASSET_PRICE_VECTOR As Variant
Dim STRIKE_PRICE_VECTOR As Variant
Dim RISK_FREE_RATE_VECTOR As Variant
Dim DIVIDEND_YIELD_VECTOR As Variant
Dim VOLATILITY_VECTOR As Variant

Dim CONTRACTS_VECTOR As Variant
Dim FEES_PAID_VECTOR As Variant
Dim DTHETA_DATE_VECTOR As Variant
Dim DVEGA_VOLAT_VECTOR As Variant
Dim DRHO_RATE_VECTOR As Variant
Dim EXERCISE_DATE_VECTOR As Variant
Dim VALUATION_DATE_VECTOR As Variant

Dim EXERCISE_TYPE_VECTOR As Variant
Dim VALUATION_TYPE_VECTOR As Variant
Dim OPTION_TYPE_VECTOR As Variant 'Call/Put
Dim POSITION_TYPE_VECTOR As Variant 'Long/Short
Dim NSTEPS_VECTOR As Variant
Dim TDAYS_PER_YEAR_VECTOR As Variant
Dim CND_TYPE_VECTOR As Variant
Dim SCALE_FACTOR_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'----------------------------------------------------------------------------
If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If
NROWS = UBound(TICKERS_VECTOR, 1)
'----------------------------------------------------------------------------
If IsArray(PREMIUM_RNG) = True Then
    PREMIUM_VECTOR = PREMIUM_RNG
    If UBound(PREMIUM_VECTOR, 1) = 1 Then
        PREMIUM_VECTOR = MATRIX_TRANSPOSE_FUNC(PREMIUM_VECTOR)
    End If
Else
    ReDim PREMIUM_VECTOR(1 To 1, 1 To 1)
    PREMIUM_VECTOR(1, 1) = PREMIUM_RNG
End If
If NROWS <> UBound(PREMIUM_VECTOR, 1) Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(ASSET_PRICE_RNG) = True Then
    ASSET_PRICE_VECTOR = ASSET_PRICE_RNG
    If UBound(ASSET_PRICE_VECTOR, 1) = 1 Then
        ASSET_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(ASSET_PRICE_VECTOR)
    End If
Else
    ReDim ASSET_PRICE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        ASSET_PRICE_VECTOR(i, 1) = ASSET_PRICE_RNG
    Next i
End If
If UBound(ASSET_PRICE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(STRIKE_PRICE_RNG) = True Then
    STRIKE_PRICE_VECTOR = STRIKE_PRICE_RNG
    If UBound(STRIKE_PRICE_VECTOR, 1) = 1 Then
        STRIKE_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(STRIKE_PRICE_VECTOR)
    End If
Else
    ReDim STRIKE_PRICE_VECTOR(1 To 1, 1 To 1)
    STRIKE_PRICE_VECTOR(1, 1) = STRIKE_PRICE_RNG
End If
If UBound(STRIKE_PRICE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(RISK_FREE_RATE_RNG) = True Then
    RISK_FREE_RATE_VECTOR = RISK_FREE_RATE_RNG
    If UBound(RISK_FREE_RATE_VECTOR, 1) = 1 Then
        RISK_FREE_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(RISK_FREE_RATE_VECTOR)
    End If
Else
    ReDim RISK_FREE_RATE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        RISK_FREE_RATE_VECTOR(i, 1) = RISK_FREE_RATE_RNG
    Next i
End If
If UBound(RISK_FREE_RATE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(DIVIDEND_YIELD_RNG) = True Then
    DIVIDEND_YIELD_VECTOR = DIVIDEND_YIELD_RNG
    If UBound(DIVIDEND_YIELD_VECTOR, 1) = 1 Then
        DIVIDEND_YIELD_VECTOR = MATRIX_TRANSPOSE_FUNC(DIVIDEND_YIELD_VECTOR)
    End If
Else
    ReDim DIVIDEND_YIELD_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        DIVIDEND_YIELD_VECTOR(i, 1) = DIVIDEND_YIELD_RNG
    Next i
End If
If UBound(DIVIDEND_YIELD_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(VOLATILITY_RNG) = True Then
    VOLATILITY_VECTOR = VOLATILITY_RNG
    If UBound(VOLATILITY_VECTOR, 1) = 1 Then
        VOLATILITY_VECTOR = MATRIX_TRANSPOSE_FUNC(VOLATILITY_VECTOR)
    End If
Else
    ReDim VOLATILITY_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        VOLATILITY_VECTOR(i, 1) = VOLATILITY_RNG
    Next i
End If
If UBound(VOLATILITY_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(EXERCISE_DATE_RNG) = True Then
    EXERCISE_DATE_VECTOR = EXERCISE_DATE_RNG
    If UBound(EXERCISE_DATE_VECTOR, 1) = 1 Then
        EXERCISE_DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(EXERCISE_DATE_VECTOR)
    End If
Else
    ReDim EXERCISE_DATE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        EXERCISE_DATE_VECTOR(i, 1) = EXERCISE_DATE_RNG
    Next i
End If
If UBound(EXERCISE_DATE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(VALUATION_DATE_RNG) = True Then
    VALUATION_DATE_VECTOR = VALUATION_DATE_RNG
    If UBound(VALUATION_DATE_VECTOR, 1) = 1 Then
        VALUATION_DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(VALUATION_DATE_VECTOR)
    End If
Else
    ReDim VALUATION_DATE_VECTOR(1 To NROWS, 1 To 1)
    If VALUATION_DATE_RNG = 0 Then
        VALUATION_DATE_RNG = Now
        VALUATION_DATE_RNG = DateSerial(Year(VALUATION_DATE_RNG), Month(VALUATION_DATE_RNG), Day(VALUATION_DATE_RNG))
    End If
    For i = 1 To NROWS
        VALUATION_DATE_VECTOR(i, 1) = VALUATION_DATE_RNG
    Next i
End If
If UBound(VALUATION_DATE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(DTHETA_DATE_RNG) = True Then
    DTHETA_DATE_VECTOR = DTHETA_DATE_RNG
    If UBound(DTHETA_DATE_VECTOR, 1) = 1 Then
        DTHETA_DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DTHETA_DATE_VECTOR)
    End If
Else
    ReDim DTHETA_DATE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        DTHETA_DATE_VECTOR(i, 1) = DTHETA_DATE_RNG
    Next i
End If
If UBound(DTHETA_DATE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(DVEGA_VOLAT_RNG) = True Then
    DVEGA_VOLAT_VECTOR = DVEGA_VOLAT_RNG
    If UBound(DVEGA_VOLAT_VECTOR, 1) = 1 Then
        DVEGA_VOLAT_VECTOR = MATRIX_TRANSPOSE_FUNC(DVEGA_VOLAT_VECTOR)
    End If
Else
    ReDim DVEGA_VOLAT_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        DVEGA_VOLAT_VECTOR(i, 1) = DVEGA_VOLAT_RNG
    Next i
End If
If UBound(DVEGA_VOLAT_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(DRHO_RATE_RNG) = True Then
    DRHO_RATE_VECTOR = DRHO_RATE_RNG
    If UBound(DRHO_RATE_VECTOR, 1) = 1 Then
        DRHO_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DRHO_RATE_VECTOR)
    End If
Else
    ReDim DRHO_RATE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        DRHO_RATE_VECTOR(i, 1) = DRHO_RATE_RNG
    Next i
End If
If UBound(DRHO_RATE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(CONTRACTS_RNG) = True Then
    CONTRACTS_VECTOR = CONTRACTS_RNG
    If UBound(CONTRACTS_VECTOR, 1) = 1 Then
        CONTRACTS_VECTOR = MATRIX_TRANSPOSE_FUNC(CONTRACTS_VECTOR)
    End If
Else
    ReDim CONTRACTS_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        CONTRACTS_VECTOR(i, 1) = CONTRACTS_RNG
    Next i
End If
If UBound(CONTRACTS_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(FEES_PAID_RNG) = True Then
    FEES_PAID_VECTOR = FEES_PAID_RNG
    If UBound(FEES_PAID_VECTOR, 1) = 1 Then
        FEES_PAID_VECTOR = MATRIX_TRANSPOSE_FUNC(FEES_PAID_VECTOR)
    End If
Else
    ReDim FEES_PAID_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS: FEES_PAID_VECTOR(i, 1) = 0: Next i
End If
If UBound(FEES_PAID_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(OPTION_TYPE_RNG) = True Then
    OPTION_TYPE_VECTOR = OPTION_TYPE_RNG
    If UBound(OPTION_TYPE_VECTOR, 1) = 1 Then
        OPTION_TYPE_VECTOR = MATRIX_TRANSPOSE_FUNC(OPTION_TYPE_VECTOR)
    End If
Else
    ReDim OPTION_TYPE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        OPTION_TYPE_VECTOR(i, 1) = OPTION_TYPE_RNG
    Next i
End If
If UBound(OPTION_TYPE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(POSITION_TYPE_RNG) = True Then
    POSITION_TYPE_VECTOR = POSITION_TYPE_RNG
    If UBound(POSITION_TYPE_VECTOR, 1) = 1 Then
        POSITION_TYPE_VECTOR = MATRIX_TRANSPOSE_FUNC(POSITION_TYPE_VECTOR)
    End If
Else
    ReDim POSITION_TYPE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        POSITION_TYPE_VECTOR(i, 1) = POSITION_TYPE_RNG
    Next i
End If
If UBound(POSITION_TYPE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(EXERCISE_TYPE_RNG) = True Then
    EXERCISE_TYPE_VECTOR = EXERCISE_TYPE_RNG
    If UBound(EXERCISE_TYPE_VECTOR, 1) = 1 Then
        EXERCISE_TYPE_VECTOR = MATRIX_TRANSPOSE_FUNC(EXERCISE_TYPE_VECTOR)
    End If
Else
    ReDim EXERCISE_TYPE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        EXERCISE_TYPE_VECTOR(i, 1) = IIf(EXERCISE_TYPE_RNG <> 0, 1, 0) '0 --> Euro Else --> Amer
    Next i
End If
If UBound(EXERCISE_TYPE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(VALUATION_TYPE_RNG) = True Then
    VALUATION_TYPE_VECTOR = VALUATION_TYPE_RNG
    If UBound(VALUATION_TYPE_VECTOR, 1) = 1 Then
        VALUATION_TYPE_VECTOR = MATRIX_TRANSPOSE_FUNC(VALUATION_TYPE_VECTOR)
    End If
Else
    ReDim VALUATION_TYPE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        VALUATION_TYPE_VECTOR(i, 1) = VALUATION_TYPE_RNG
    Next i
End If
If UBound(VALUATION_TYPE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(NSTEPS_RNG) = True Then
    NSTEPS_VECTOR = NSTEPS_RNG
    If UBound(NSTEPS_VECTOR, 1) = 1 Then
        NSTEPS_VECTOR = MATRIX_TRANSPOSE_FUNC(NSTEPS_VECTOR)
    End If
Else
    ReDim NSTEPS_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        NSTEPS_VECTOR(i, 1) = NSTEPS_RNG
    Next i
End If
If UBound(NSTEPS_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(TDAYS_PER_YEAR_RNG) = True Then
    TDAYS_PER_YEAR_VECTOR = TDAYS_PER_YEAR_RNG
    If UBound(TDAYS_PER_YEAR_VECTOR, 1) = 1 Then
        TDAYS_PER_YEAR_VECTOR = MATRIX_TRANSPOSE_FUNC(TDAYS_PER_YEAR_VECTOR)
    End If
Else
    ReDim TDAYS_PER_YEAR_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        TDAYS_PER_YEAR_VECTOR(i, 1) = TDAYS_PER_YEAR_RNG
    Next i
End If
If UBound(TDAYS_PER_YEAR_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(SCALE_FACTOR_RNG) = True Then
    SCALE_FACTOR_VECTOR = SCALE_FACTOR_RNG
    If UBound(SCALE_FACTOR_VECTOR, 1) = 1 Then
        SCALE_FACTOR_VECTOR = MATRIX_TRANSPOSE_FUNC(SCALE_FACTOR_VECTOR)
    End If
Else
    ReDim SCALE_FACTOR_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        SCALE_FACTOR_VECTOR(i, 1) = SCALE_FACTOR_RNG
    Next i
End If
If UBound(SCALE_FACTOR_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------
If IsArray(CND_TYPE_RNG) = True Then
    CND_TYPE_VECTOR = CND_TYPE_RNG
    If UBound(CND_TYPE_VECTOR, 1) = 1 Then
        CND_TYPE_VECTOR = MATRIX_TRANSPOSE_FUNC(CND_TYPE_VECTOR)
    End If
Else
    ReDim CND_TYPE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        CND_TYPE_VECTOR(i, 1) = CND_TYPE_RNG
    Next i
End If
If UBound(CND_TYPE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL
'----------------------------------------------------------------------------

GoSub HEADERS_LINE
For i = 1 To NROWS
    GoSub CALCS_LINE
1983:
Next i

OPTION_BINOMIAL_TRINOMIAL_GREEKS_FUNC = TEMP_MATRIX

Exit Function
'----------------------------------------------------------------------------------------------------------------------
HEADERS_LINE:
'----------------------------------------------------------------------------------------------------------------------
    HEADINGS_STR = _
    "TICKER,CONTRACTS,PREMIUM,ASSET PRICE,STRIKE PRICE,RISK FREE RATE,DIVIDEND YIELD,VOLATILITY," & _
    "EXERCISE DATE,VALUATION DATE,DTHETA DATE,DVEGA VOLATILITY,DRHO RATE,OPTION TYPE,POSITION TYPE,VALUATION MODEL,NSTEPS,TDAYS PER YEAR,CND TYPE," & _
    "TICKER,POSITION SIZE,ASSET PRICE,STRIKE PRICE,ATM/ITM/OTM,MODEL PRICE,MARKET PRICE,INTRINSIC VALUE,TIME VALUE,IMPLIED VOLATILITY,DELTA,GAMMA,VEGA,THETA,RHO,SCALE FACTOR,TRANSACTION FEES,P&L,"
    
    j = Len(HEADINGS_STR)
    NCOLUMNS = 0
    For i = 1 To j
        If Mid(HEADINGS_STR, i, 1) = "," Then: NCOLUMNS = NCOLUMNS + 1
    Next i
    ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
    i = 1
    For k = 1 To NCOLUMNS
        j = InStr(i, HEADINGS_STR, ",")
        TEMP_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
        i = j + 1
    Next k
'----------------------------------------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------------------------------------------------------
CALCS_LINE:
'---------------------------------------------------------------------------------------------------------------------------------------
    OT_VAL = CInt(OPTION_TYPE_VECTOR(i, 1))
    '-----------------------------------------------------------------------------------------------------------------------------------
    If OT_VAL = 0 Then 'Position in the Underlying Asset
    '------------------------------------------------------------------------------------------------------------------
        TICKER_STR = TICKERS_VECTOR(i, 1)
        TEMP_MATRIX(i, 1) = TICKER_STR
        If ASSET_PRICE_VECTOR(i, 1) = "" Or STRIKE_PRICE_VECTOR(i, 1) = "" Then
            For j = 2 To NCOLUMNS: TEMP_MATRIX(i, j) = "": Next j
            GoTo 1983
        End If
        
        S_VAL = ASSET_PRICE_VECTOR(i, 1)
        K_VAL = STRIKE_PRICE_VECTOR(i, 1)
        RF_VAL = RISK_FREE_RATE_VECTOR(i, 1)
        DY_VAL = DIVIDEND_YIELD_VECTOR(i, 1)
        V_VAL = VOLATILITY_VECTOR(i, 1)
        Q_VAL = CInt(CONTRACTS_VECTOR(i, 1))
        TF_VAL = FEES_PAID_VECTOR(i, 1)
        PT_VAL = CInt(POSITION_TYPE_VECTOR(i, 1))
        If PT_VAL <> 1 Then: PT_VAL = -1 'Put / Short
        EDATE_VAL = EXERCISE_DATE_VECTOR(i, 1)
        VDATE_VAL = VALUATION_DATE_VECTOR(i, 1)
    
        TEMP_MATRIX(i, 2) = Q_VAL
        TEMP_MATRIX(i, 3) = ""
        TEMP_MATRIX(i, 4) = S_VAL
        TEMP_MATRIX(i, 5) = K_VAL 'Price Paid/Sold for the Stock
    
        TEMP_MATRIX(i, 6) = RF_VAL
        TEMP_MATRIX(i, 7) = DY_VAL
    
        TEMP_MATRIX(i, 8) = V_VAL
        TEMP_MATRIX(i, 9) = EDATE_VAL
        TEMP_MATRIX(i, 10) = VDATE_VAL
            
        For j = 11 To 14: TEMP_MATRIX(i, j) = "": Next j
    
        Select Case PT_VAL
        Case 1
            TEMP_MATRIX(i, 15) = "LONG"
            TEMP_MATRIX(i, 24) = IIf((K_VAL < S_VAL), "ITM", IIf(K_VAL = S_VAL, "ATM", "OTM"))
        Case -1
            TEMP_MATRIX(i, 15) = "SHORT"
            TEMP_MATRIX(i, 24) = IIf((K_VAL > S_VAL), "ITM", IIf(K_VAL = S_VAL, "ATM", "OTM"))
        End Select
        TEMP_MATRIX(i, 16) = "ASSET"
        For j = 17 To 19: TEMP_MATRIX(i, j) = "": Next j
        TEMP_MATRIX(i, 20) = TICKER_STR
        TEMP_MATRIX(i, 21) = Q_VAL '* S_VAL
        TEMP_MATRIX(i, 22) = S_VAL
        
        TEMP_MATRIX(i, 23) = K_VAL
                
        For j = 25 To NCOLUMNS - 2: TEMP_MATRIX(i, j) = "": Next j
        TEMP_MATRIX(i, 36) = TF_VAL
        TEMP_MATRIX(i, 37) = PT_VAL * (S_VAL - K_VAL) * Q_VAL - TF_VAL
    '------------------------------------------------------------------------------------------------------------------
    Else
    '------------------------------------------------------------------------------------------------------------------
    
    'Function option_e(Put_Call As String, S As Double, e As Double, Tmt As Double, r As Double, q As Double, SIGMA As Double, Command As String) As Double

        TICKER_STR = TICKERS_VECTOR(i, 1)
        TEMP_MATRIX(i, 1) = TICKER_STR
        If PREMIUM_VECTOR(i, 1) = "" Or ASSET_PRICE_VECTOR(i, 1) = "" Or STRIKE_PRICE_VECTOR(i, 1) = "" Then
            For j = 2 To NCOLUMNS: TEMP_MATRIX(i, j) = "": Next j
            GoTo 1983
        End If
        P_VAL = PREMIUM_VECTOR(i, 1)
        S_VAL = ASSET_PRICE_VECTOR(i, 1)
        K_VAL = STRIKE_PRICE_VECTOR(i, 1)
        RF_VAL = RISK_FREE_RATE_VECTOR(i, 1)
        DY_VAL = DIVIDEND_YIELD_VECTOR(i, 1)
        V_VAL = VOLATILITY_VECTOR(i, 1)
        Q_VAL = CInt(CONTRACTS_VECTOR(i, 1))
        TF_VAL = FEES_PAID_VECTOR(i, 1)
        DTHETA_VAL = (EXERCISE_DATE_VECTOR(i, 1) - DTHETA_DATE_VECTOR(i, 1)) / TDAYS_PER_YEAR_VECTOR(i, 1)
        DVEGA_VAL = DVEGA_VOLAT_VECTOR(i, 1)
        DRHO_VAL = DRHO_RATE_VECTOR(i, 1)
        EDATE_VAL = EXERCISE_DATE_VECTOR(i, 1)
        VDATE_VAL = VALUATION_DATE_VECTOR(i, 1)
        ET_VAL = CInt(EXERCISE_TYPE_VECTOR(i, 1))
        PT_VAL = CInt(POSITION_TYPE_VECTOR(i, 1))
        If PT_VAL <> 1 Then: PT_VAL = -1 'Short
        D_VAL = TDAYS_PER_YEAR_VECTOR(i, 1)
        T_VAL = (EDATE_VAL - VDATE_VAL) / D_VAL
        SF_VAL = SCALE_FACTOR_VECTOR(i, 1)
        
        TEMP_MATRIX(i, 2) = Q_VAL
        TEMP_MATRIX(i, 3) = P_VAL
        TEMP_MATRIX(i, 4) = S_VAL
        TEMP_MATRIX(i, 5) = K_VAL
        
        TEMP_MATRIX(i, 6) = RF_VAL
        TEMP_MATRIX(i, 7) = DY_VAL
        TEMP_MATRIX(i, 8) = V_VAL
        TEMP_MATRIX(i, 9) = EDATE_VAL
        TEMP_MATRIX(i, 10) = VDATE_VAL
        
        TEMP_MATRIX(i, 11) = DTHETA_DATE_VECTOR(i, 1)
        TEMP_MATRIX(i, 12) = DVEGA_VAL
        TEMP_MATRIX(i, 13) = DRHO_VAL
        
        Select Case OT_VAL
        Case 1
            TEMP_MATRIX(i, 14) = "CALL"
            TEMP_MATRIX(i, 24) = IIf((K_VAL < S_VAL), "ITM", IIf(K_VAL = S_VAL, "ATM", "OTM"))
        Case -1 '-1
            TEMP_MATRIX(i, 14) = "PUT"
            TEMP_MATRIX(i, 24) = IIf((K_VAL > S_VAL), "ITM", IIf(K_VAL = S_VAL, "ATM", "OTM"))
        Case Else
            For j = 14 To NCOLUMNS: TEMP_MATRIX(i, j) = "": Next j
            GoTo ERROR_LABEL
        End Select
        Select Case PT_VAL
        Case 1
            TEMP_MATRIX(i, 15) = "LONG"
        Case Else '-1
            TEMP_MATRIX(i, 15) = "SHORT"
        End Select
        
        TEMP_MATRIX(i, 18) = D_VAL
        
        TEMP_MATRIX(i, 20) = TICKER_STR
        TEMP_MATRIX(i, 21) = Q_VAL '* P_VAL
        TEMP_MATRIX(i, 22) = S_VAL
        TEMP_MATRIX(i, 23) = K_VAL
        
        TEMP_MATRIX(i, 26) = P_VAL
        TEMP_MATRIX(i, 27) = OT_VAL * (S_VAL - K_VAL) 'Intrinsic Value
        If TEMP_MATRIX(i, 27) < 0 Then: TEMP_MATRIX(i, 27) = 0
        TEMP_MATRIX(i, 28) = P_VAL - TEMP_MATRIX(i, 27) 'Time Value
        If TEMP_MATRIX(i, 28) < 0 Then: TEMP_MATRIX(i, 28) = 0
        
        '-----------------------------------------------------------------------------------------------------------------------------------
        'Delta --> The amount that the theoretical price will change if the market moves up/down 1 point
        'Gamma --> The amount that the Delta will change if the market moves up/down 1 point
        'Theta --> The amount that the theoretical price will change when x days passes.
        'Vega --> The amount that the theoretical price will change if the volatility of the asset moves up/down by 1 percentage point
        'Rho --> The amount that the theoretical price will change if interest rates move up/down by 1 percentage point
        VT_VAL = CInt(VALUATION_TYPE_VECTOR(i, 1))
        '-----------------------------------------------------------------------------------------------------------------------------------
        If VT_VAL = 0 Then 'BS
        '-----------------------------------------------------------------------------------------------------------------------------------
            CND_VAL = CND_TYPE_VECTOR(i, 1)
            TEMP_MATRIX(i, 16) = "BLACK - EURO"
            TEMP_MATRIX(i, 17) = ""
            TEMP_MATRIX(i, 19) = CND_VAL

            If OT_VAL = 1 Then PRICING_MODEL_STR = "CALL" Else PRICING_MODEL_STR = "PUT"
            With Excel.Application
    
                TEMP_MATRIX(i, 25) = .Run("EUROPEAN_" & PRICING_MODEL_STR & "_OPTION_FUNC", S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL, CND_VAL) 'Option Fair Value
      
                TEMP_MATRIX(i, 29) = BLACK_SCHOLES_IMPLIED_VOLATILITY_FUNC(P_VAL, S_VAL, K_VAL, T_VAL, RF_VAL, RF_VAL - DY_VAL, OT_VAL, 0, 1, CND_VAL)
                If TEMP_MATRIX(i, 29) = 2 ^ 52 Then: TEMP_MATRIX(i, 29) = CVErr(xlErrNA)
                '((P_VAL * (V_VAL - DVEGA_VAL)) + (TEMP_MATRIX(i,25) * DVEGA_VAL - TEMP_MATRIX(i,29) * V_VAL)) / (TEMP_MATRIX(i,25) - TEMP_MATRIX(i,29)) 'Implied Volatility guess
                
                TEMP_MATRIX(i, 30) = .Run(PRICING_MODEL_STR & "_OPTION_DELTA_FUNC", S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL, CND_VAL) 'Delta
                TEMP_MATRIX(i, 31) = .Run("CALL_PUT_OPTION_GAMMA_FUNC", S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL) 'Gamma
                
                TEMP_MATRIX(i, 32) = .Run("CALL_PUT_OPTION_VEGA_FUNC", S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL) 'Vega
                TEMP_MATRIX(i, 33) = .Run(PRICING_MODEL_STR & "_OPTION_THETA_FUNC", S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL, CND_VAL) 'Theta
                TEMP_MATRIX(i, 34) = .Run(PRICING_MODEL_STR & "_OPTION_RHO_FUNC", S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL, CND_VAL) 'Rho
            End With
            TEMP_MATRIX(i, 32) = TEMP_MATRIX(i, 32) * SF_VAL 'Vega
            TEMP_MATRIX(i, 33) = TEMP_MATRIX(i, 33) * 1 / D_VAL '1-Day Theta
            TEMP_MATRIX(i, 34) = TEMP_MATRIX(i, 34) * SF_VAL 'Rho
            TEMP_MATRIX(i, 35) = SF_VAL
        '-----------------------------------------------------------------------------------------------------------------------------------
        Else 'Binom/Trinom
        '-----------------------------------------------------------------------------------------------------------------------------------
            N_VAL = CInt(NSTEPS_VECTOR(i, 1))
            TEMP_MATRIX(i, 17) = N_VAL
            TEMP_MATRIX(i, 19) = ""
            If VT_VAL = 1 Then
                PRICING_MODEL_STR = "OPTION_BINOMIAL_TREE1_FUNC"
                TEMP_MATRIX(i, 16) = "BINOMIAL - " & IIf(ET_VAL = 0, "EURO", "AMER")
            Else
                PRICING_MODEL_STR = "OPTION_TRINOMIAL_TREE1_FUNC"
                TEMP_MATRIX(i, 16) = "TRINOMIAL - " & IIf(ET_VAL = 0, "EURO", "AMER")
            End If
            With Excel.Application
                TEMP_MATRIX(i, 25) = .Run(PRICING_MODEL_STR, S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, 0, 0, DTHETA_VAL, DVEGA_VAL, DRHO_VAL, P_VAL) 'Option Fair Value
                TEMP_MATRIX(i, 29) = .Run(PRICING_MODEL_STR, S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, 0, 6, DTHETA_VAL, DVEGA_VAL, DRHO_VAL, P_VAL) 'Implied Volatility Guess - Vega Volat
                
                TEMP_MATRIX(i, 30) = .Run(PRICING_MODEL_STR, S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, 0, 1, DTHETA_VAL, DVEGA_VAL, DRHO_VAL, P_VAL) 'Delta
                TEMP_MATRIX(i, 31) = .Run(PRICING_MODEL_STR, S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, 0, 2, DTHETA_VAL, DVEGA_VAL, DRHO_VAL, P_VAL) 'Gamma
                
                TEMP_MATRIX(i, 32) = .Run(PRICING_MODEL_STR, S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, 0, 4, DTHETA_VAL, DVEGA_VAL, DRHO_VAL, P_VAL) 'Vega
                TEMP_MATRIX(i, 33) = .Run(PRICING_MODEL_STR, S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, 0, 3, DTHETA_VAL, DVEGA_VAL, DRHO_VAL, P_VAL) 'Theta
                TEMP_MATRIX(i, 34) = .Run(PRICING_MODEL_STR, S_VAL, K_VAL, T_VAL, RF_VAL, DY_VAL, V_VAL, N_VAL, OT_VAL, 0, 5, DTHETA_VAL, DVEGA_VAL, DRHO_VAL, P_VAL) 'Rho
            End With
            TEMP_MATRIX(i, 32) = (TEMP_MATRIX(i, 32) - TEMP_MATRIX(i, 25)) * SF_VAL 'Vega
            If VT_VAL = 1 Then 'Binomial
                TEMP_MATRIX(i, 33) = (TEMP_MATRIX(i, 33) - TEMP_MATRIX(i, 25)) * -1 '1-Day Theta
            Else 'Trinomial
                TEMP_MATRIX(i, 33) = TEMP_MATRIX(i, 33) * 1 / D_VAL '1-Day Theta
            End If
            TEMP_MATRIX(i, 34) = (TEMP_MATRIX(i, 34) - TEMP_MATRIX(i, 25)) * SF_VAL 'Rho
            TEMP_MATRIX(i, 35) = SF_VAL
        '-----------------------------------------------------------------------------------------------------------------------------------
        End If
        '-----------------------------------------------------------------------------------------------------------------------------------
        For j = 30 To 34: TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) * Q_VAL: Next j
        TEMP_MATRIX(i, 36) = TF_VAL
        Select Case PT_VAL
        Case 1
            TEMP_MATRIX(i, 37) = TEMP_MATRIX(i, 27) * Q_VAL - TF_VAL
        Case Else
            TEMP_MATRIX(i, 37) = (P_VAL - TEMP_MATRIX(i, 27)) * Q_VAL - TF_VAL 'Premium Received for Writing the Calls/Puts - Intrinsic Value
        End Select
    '-----------------------------------------------------------------------------------------------------------------------------------
    End If
    '-----------------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
OPTION_BINOMIAL_TRINOMIAL_GREEKS_FUNC = Err.number
End Function