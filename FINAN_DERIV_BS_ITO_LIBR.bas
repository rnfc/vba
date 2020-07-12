Attribute VB_Name = "FINAN_DERIV_BS_ITO_LIBR"

'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------



'Do Sensitivity by changing Future Time Periods

'Reference: http://www.gummy-stuff.org/Ito-options-2.htm

Function ITO_OPTIONS_DISTRIBUTION1_FUNC(ByVal STRIKE_PRICE As Double, _
ByVal CURRENT_STOCK_PRICE As Double, _
ByVal ANNUAL_RETURN As Double, _
ByVal ANNUAL_VOLATILITY As Double, _
ByVal ANNUAL_CASH_RATE As Double, _
ByVal EXPIRATION As Double, _
ByVal FUTURE_TIME_PERIODS As Double, _
Optional ByVal REQUIRED_DELTA As Double = 0.8, _
Optional ByVal START_STOCK_PRICE As Double = 25.82, _
Optional ByVal DELTA_STOCK_PRICE As Double = 0.5, _
Optional ByVal NBINS As Long = 41, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

'EXPIRATION and FUTURE_TIME_PERIODS in years!!!!
'CASH_RATE --> Risk Free Rate
'FUTURE_TIME_PERIODS --> used to calculate the Expected Price and Option Premium

Dim i As Long
Dim j As Long
Dim TENOR As Double
Dim PI2_VAL As Double
Dim RTN_FACTOR As Double

Dim TEMP_STR As String
Dim TEMP_SUM As Double
Dim PRICE_VAL As Double
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

TENOR = EXPIRATION - FUTURE_TIME_PERIODS
RTN_FACTOR = ANNUAL_RETURN - 0.5 * ANNUAL_VOLATILITY ^ 2
PI2_VAL = 6.28318530717959
If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1 'Put

ReDim TEMP_MATRIX(0 To NBINS, 1 To 9)

'Ito Price Distribution
TEMP_MATRIX(0, 1) = "Price: P"
TEMP_MATRIX(0, 2) = "Ito: f(P)" '
TEMP_MATRIX(0, 3) = "Ito: F(P))" '
' B-S Option Distribution
TEMP_MATRIX(0, 4) = "Q(p)" 'Black Scholes Price
TEMP_MATRIX(0, 5) = "B-S: f(Q)"
TEMP_MATRIX(0, 6) = "B-S: F(Q)"
TEMP_MATRIX(0, 7) = "Gain @ T"
TEMP_MATRIX(0, 8) = "P f(P)"
TEMP_MATRIX(0, 9) = "Gain%"

i = 1
PRICE_VAL = START_STOCK_PRICE
TEMP_MATRIX(i, 1) = PRICE_VAL
GoSub PROBABILITY_LINE
TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 2)
TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 4) * TEMP_MATRIX(i, 2)
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 5)

TEMP_SUM = TEMP_MATRIX(i, 5)
PRICE_VAL = PRICE_VAL + DELTA_STOCK_PRICE
For i = 2 To NBINS
    TEMP_MATRIX(i, 1) = PRICE_VAL
    If TEMP_MATRIX(i, 1) > CURRENT_STOCK_PRICE And TEMP_MATRIX(i - 1, 1) <= CURRENT_STOCK_PRICE Then: j = i
    GoSub PROBABILITY_LINE
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i - 1, 3) + TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 4) * TEMP_MATRIX(i, 2)
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 5)
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i - 1, 6) + TEMP_MATRIX(i, 5)
    PRICE_VAL = PRICE_VAL + DELTA_STOCK_PRICE
Next i

PRICE_VAL = BLACK_SCHOLES_OPTION_FUNC(CURRENT_STOCK_PRICE, STRIKE_PRICE, EXPIRATION, ANNUAL_CASH_RATE, ANNUAL_VOLATILITY, OPTION_FLAG, CND_TYPE)
'Initial Option Price

i = 1
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / TEMP_SUM
TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 4) - PRICE_VAL
TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 1) * TEMP_MATRIX(i, 2)
TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 7) / PRICE_VAL
For i = 2 To NBINS
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / TEMP_SUM
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 4) - PRICE_VAL
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 1) * TEMP_MATRIX(i, 2)
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 7) / PRICE_VAL
Next i

If OUTPUT > 1 Then
    If j <> 0 Then
        ITO_OPTIONS_DISTRIBUTION1_FUNC = TEMP_MATRIX(j, 9)
    Else
        ITO_OPTIONS_DISTRIBUTION1_FUNC = "--"
    End If
    Exit Function
End If

Select Case OUTPUT
Case 0
    ITO_OPTIONS_DISTRIBUTION1_FUNC = TEMP_MATRIX
Case 1
    ReDim TEMP_VECTOR(1 To 10, 1 To 2)
    TEMP_VECTOR(1, 1) = "Expected Price at T = " & Format(FUTURE_TIME_PERIODS, "0.0000")
    TEMP_VECTOR(2, 1) = "Probability of attaining price " & Format(TEMP_MATRIX(j, 1), "$0.00")
    TEMP_VECTOR(3, 1) = "Expected Premium at " & Format(FUTURE_TIME_PERIODS, "0.0000")
    TEMP_VECTOR(4, 1) = "Expected Gain at " & Format(FUTURE_TIME_PERIODS, "0.0000")
    TEMP_VECTOR(5, 1) = "Expected Percentage Gain"
    
    TEMP_VECTOR(6, 1) = "Initial Option Price: Co"
    TEMP_VECTOR(7, 1) = "Initial Delta: D"
    TEMP_VECTOR(8, 1) = "Required Delta: Do"
    TEMP_VECTOR(9, 1) = "Required Strike: Ko"
    TEMP_VECTOR(10, 1) = "Ratio: Po/Ko"

    If j = 0 Then
        TEMP_STR = "Current Stock Price of " & Format(CURRENT_STOCK_PRICE, "0.00") & _
                                " is not within the range of the histogram"
        For i = 1 To 5: TEMP_VECTOR(i, 2) = TEMP_STR: Next i
    Else
        TEMP_VECTOR(1, 2) = TEMP_MATRIX(j, 1)
        TEMP_VECTOR(2, 2) = 1 - TEMP_MATRIX(j, 3)
        TEMP_VECTOR(3, 2) = TEMP_MATRIX(j, 4)
        TEMP_VECTOR(4, 2) = TEMP_MATRIX(j, 7)
        TEMP_VECTOR(5, 2) = TEMP_MATRIX(j, 9)
    End If
    TEMP_VECTOR(6, 2) = PRICE_VAL
    TEMP_VECTOR(7, 2) = CND_FUNC((Log(CURRENT_STOCK_PRICE / STRIKE_PRICE) + (ANNUAL_CASH_RATE + ANNUAL_VOLATILITY ^ 2 / 2) * EXPIRATION) / (ANNUAL_VOLATILITY * Sqr(EXPIRATION)), CND_TYPE)
    TEMP_VECTOR(8, 2) = REQUIRED_DELTA
    TEMP_VECTOR(9, 2) = CURRENT_STOCK_PRICE / Exp(NORMSINV_FUNC(TEMP_VECTOR(8, 2), 0, 1, 0) * ANNUAL_VOLATILITY * Sqr(EXPIRATION) - (ANNUAL_CASH_RATE + ANNUAL_VOLATILITY ^ 2 / 2) * EXPIRATION)
    TEMP_VECTOR(10, 2) = Exp(NORMSINV_FUNC(TEMP_VECTOR(8, 2), 0, 1, 0) * ANNUAL_VOLATILITY * Sqr(EXPIRATION) - (ANNUAL_CASH_RATE + ANNUAL_VOLATILITY ^ 2 / 2) * EXPIRATION)
    ITO_OPTIONS_DISTRIBUTION1_FUNC = TEMP_VECTOR
End Select

Exit Function
'-------------------------------------------------------------------------------------------
PROBABILITY_LINE:
'-------------------------------------------------------------------------------------------
    TEMP_MATRIX(i, 2) = 1 / (ANNUAL_VOLATILITY * Sqr(PI2_VAL * FUTURE_TIME_PERIODS)) / TEMP_MATRIX(i, 1) * Exp(-(1 / (2 * FUTURE_TIME_PERIODS * ANNUAL_VOLATILITY ^ 2)) * (Log(TEMP_MATRIX(i, 1) / CURRENT_STOCK_PRICE) - (ANNUAL_RETURN - 0.5 * ANNUAL_VOLATILITY ^ 2) * FUTURE_TIME_PERIODS) ^ 2) * DELTA_STOCK_PRICE
    TEMP_MATRIX(i, 4) = BLACK_SCHOLES_OPTION_FUNC(TEMP_MATRIX(i, 1), STRIKE_PRICE, TENOR, ANNUAL_CASH_RATE, ANNUAL_VOLATILITY, OPTION_FLAG, CND_TYPE)
'-------------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------------
ERROR_LABEL:
ITO_OPTIONS_DISTRIBUTION1_FUNC = Err.number
End Function

'Investors always hope to estimate the price of a stock at some time in the future and ...
'Having some estimate of the future price or, better still, a distribution of prices at some
'time in the future ...Let's say we download four years of weekly prices for GE stock,
'calculated the Mean and Standard Deviation of weekly returns (them's r and s), pick out the
'current price (that's Po), pick a time period (that's T) and use the magic Ito's formula to
'estimate the Probability of achieving a price P at time T = 10 weeks (FUTURE_TIME_PERIODS).
'f(P)=EXP(-((LN(P/Po)-Rtn*T)^2/(2*T*V^2)))/(V*P*SQRT(2Pi*T))
'where V = s is the Standard Deviation.

Function ITO_OPTIONS_DISTRIBUTION2_FUNC(ByVal STRIKE_PRICE As Double, _
ByVal CURRENT_STOCK_PRICE As Double, _
ByVal MEAN_RETURN_PER_PERIOD As Double, _
ByVal VOLATILITY_RETURN_PER_PERIOD As Double, _
ByVal CASH_RATE_PER_PERIOD As Double, _
Optional ByVal EXPIRATION As Double = 50, _
Optional ByVal FUTURE_TIME_PERIODS As Double = 10, _
Optional ByVal START_STOCK_PRICE As Double = 22, _
Optional ByVal DELTA_STOCK_PRICE As Double = 0.28, _
Optional ByVal NBINS As Long = 100, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

'PERIOD --> In this example is weeks
'CASH_RATE --> Risk Free Rate
'FUTURE_TIME_PERIODS --> used to calculate the Expected Price and Option Premium

Dim i As Long
Dim TENOR As Double
Dim PI2_VAL As Double
Dim RTN_FACTOR As Double

Dim PRICE_VAL As Double
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

TENOR = EXPIRATION - FUTURE_TIME_PERIODS
RTN_FACTOR = MEAN_RETURN_PER_PERIOD - 0.5 * VOLATILITY_RETURN_PER_PERIOD ^ 2
PI2_VAL = 6.28318530717959
If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1 'Put

ReDim TEMP_MATRIX(0 To NBINS, 1 To 8)

TEMP_MATRIX(0, 1) = "P"
TEMP_MATRIX(0, 2) = "f(P)" 'Price probability f(P) at FUTURE_TIME_PERIODS
TEMP_MATRIX(0, 3) = "F(p)" 'Cumulative probability f(P) at FUTURE_TIME_PERIODS
TEMP_MATRIX(0, 4) = "Q(p)" 'Black-Scholes Option Premium Q(P)
TEMP_MATRIX(0, 5) = "Q(P)*f(P)*dP"
TEMP_MATRIX(0, 6) = "SUM[Q(P)*f(P)*dP]"
TEMP_MATRIX(0, 7) = "f(P)dP"
TEMP_MATRIX(0, 8) = "SUM[Pf(P)dP]"

i = 1
PRICE_VAL = START_STOCK_PRICE
TEMP_MATRIX(i, 1) = PRICE_VAL
GoSub PROBABILITY_LINE
TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 2) * DELTA_STOCK_PRICE
TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 7)
TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 4) * TEMP_MATRIX(i, 7)
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 5)
TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 1) * TEMP_MATRIX(i, 7)
PRICE_VAL = PRICE_VAL + DELTA_STOCK_PRICE

For i = 2 To NBINS
    TEMP_MATRIX(i, 1) = PRICE_VAL
    GoSub PROBABILITY_LINE
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 2) * DELTA_STOCK_PRICE
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i - 1, 3) + TEMP_MATRIX(i, 7)
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 4) * TEMP_MATRIX(i, 7)
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i - 1, 6) + TEMP_MATRIX(i, 5)
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i - 1, 8) + TEMP_MATRIX(i, 1) * TEMP_MATRIX(i, 7)
    PRICE_VAL = PRICE_VAL + DELTA_STOCK_PRICE
Next i

Select Case OUTPUT
Case 0
    ITO_OPTIONS_DISTRIBUTION2_FUNC = TEMP_MATRIX
Case Else
    ReDim TEMP_VECTOR(1 To 3, 1 To 1)
    TEMP_VECTOR(1, 1) = "Expected Price at t = " & Format(FUTURE_TIME_PERIODS, "0.0") & " is " & Format(TEMP_MATRIX(NBINS, 8), "$0.00")
    PRICE_VAL = BLACK_SCHOLES_OPTION_FUNC(CURRENT_STOCK_PRICE, STRIKE_PRICE, TENOR, CASH_RATE_PER_PERIOD, VOLATILITY_RETURN_PER_PERIOD, OPTION_FLAG, CND_TYPE)
    TEMP_VECTOR(2, 1) = "Current Option premium should be (about) " & Format(PRICE_VAL, "$0.00")
    TEMP_VECTOR(3, 1) = "Option Premium at t = " & Format(FUTURE_TIME_PERIODS, "0.0") & " should be (about) " & Format(TEMP_MATRIX(NBINS, 6), "$0.00")
    ITO_OPTIONS_DISTRIBUTION2_FUNC = TEMP_VECTOR
End Select

'Exactly P ... to the nearest penny?
'It's the probability of having the price P lie in small intervals. For example, I have
'conclude that there's a 2.5% probability that the price will lie between $36.00 and $36.28.

'I'm assuming that somehow we've managed to come up with a price distribution.
'Okay, suppose we have the distribution. Let's call it f(P).
'Then f(P)dP gives the fraction (or percentage, if we multiply by 100) of stock prices that
'lie in an interval of width dP, about the price P.

'Now suppose that we have some other quantity, Q(P), that depends upon P (like a gain or loss
'or the value of an option etc.). We can then calculate the Expected Value of this other quantity
'by using: SUM[Q(P)f(P)d(P)].

'For each of a jillion values of P, we calculate Q(P)f(P)d(P) and then we add them all up.
'For the math-types, that's integrating ... like so: where "a" and "b" are the minimum and
'maximum prices that occur (or just use 0 and ?). It gives the Expected Value of that other thing, Q(P).
'For example, for each P we plot Q(P)f(P)d(P) using the distribution.

'But what's Q(P)?
'Well, it happens to be the option premium for GE stock at a time when the premium is about
'$5.00 ... but that's irrelevant. I just wanted to illustrate that, in this case (for example),
'the expected value is about $5.94.

'Not at all. You type in a stock symbol and click the Download button to get four years worth of weekly pirices.
'Then you adjust the plot range with Hi and Lo prices so you can see the plots.
'Then you enter all the stuff, like Strike Price, Risk-free rate, etc. and pick some time T, in the future.

'With the option parameters you've entered in the red boxes, the routine calculates the premium
'using Black-Scholes.

'Reference: http://gummy-stuff.org/price-probabilities.htm


Exit Function
'-------------------------------------------------------------------------------------------
PROBABILITY_LINE:
'-------------------------------------------------------------------------------------------
    TEMP_MATRIX(i, 2) = Exp(-((Log(TEMP_MATRIX(i, 1) / CURRENT_STOCK_PRICE) - RTN_FACTOR * FUTURE_TIME_PERIODS) ^ 2 / (2 * FUTURE_TIME_PERIODS * VOLATILITY_RETURN_PER_PERIOD ^ 2))) / (VOLATILITY_RETURN_PER_PERIOD * TEMP_MATRIX(i, 1) * Sqr(PI2_VAL * FUTURE_TIME_PERIODS))
    TEMP_MATRIX(i, 4) = BLACK_SCHOLES_OPTION_FUNC(TEMP_MATRIX(i, 1), STRIKE_PRICE, TENOR, CASH_RATE_PER_PERIOD, VOLATILITY_RETURN_PER_PERIOD, OPTION_FLAG, CND_TYPE)
'-------------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------------
ERROR_LABEL:
ITO_OPTIONS_DISTRIBUTION2_FUNC = Err.number
End Function

Function ITO_OPTIONS_DISTRIBUTION3_FUNC( _
Optional ByVal P0_VAL As Double = 34.87, _
Optional ByVal STRIKE As Double = 32, _
Optional ByVal CASH_RATE As Double = 5.68600096428673E-02 / 100, _
Optional ByVal DIVID_RATE As Double = 0, _
Optional ByVal MEAN_VAL As Double = 4.15798524074258E-02 / 100, _
Optional ByVal SIGMA_VAL As Double = 2.89706894888478 / 100, _
Optional ByVal EXPIRATION As Double = 30, _
Optional ByVal NO_PERIODS As Long = 5, _
Optional ByVal MIN_PRICE As Double = 24.87, _
Optional ByVal DELTA_PRICE As Double = 0.833333333333333, _
Optional ByVal NBINS As Long = 25, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

'P0 --> Current Price
'CASH RATE, DIVID_RATE, SIGMA and MEAN per period

'-------------------------------------------------------------------------------
'Kiyosi Ito studied mathematics in the Faculty of Science of the Imperial
'University of Tokyo, graduating in 1938. In the 1940s he wrote several
'papers on Stochastic Processes and, in particular, developed what is now
'called Ito Calculus.
'---------------------------------------------------------------------------------

'Now we 'll rewrite Ito's equation like so: dP = m(t,P)dt+ s(t,P) df(t)
'where f(t), which gives rise to the term df(t), is a Brownian Motion.

'where P(t) is the Price of a stock at time t and dP is the change in
'Price over some small time interval, dt.

'This change, dP, comes in two parts called DRIFT and DIFFUSION:

'm(t,P)dt is the deterministic part
's(t,P)df is the stochastic part, with df a random Brownian Motion
'with Mean = 0 and Standard Deviation = 1.

'References go here:

'Ito and Black-Scholes: http://orion.it.luc.edu/~tmallia/downloads/ito.pdf
'Black-Scholes Equation: http://srikant.org/thesis/node8.html '--> BEST
'Random Walks in Physics: http://physics.hkbu.edu.hk/~sci2240/manuals/project.pdf


Dim i As Long
Dim PI_VAL As Double

Dim A_VAL As Double
Dim B_VAL As Double
Dim K_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim PRICE_VAL As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1 'Put
PI_VAL = 3.14159265358979

TEMP1_SUM = 0

A_VAL = 1 / (SIGMA_VAL * Sqr(2 * PI_VAL * NO_PERIODS))
B_VAL = (MEAN_VAL - 0.5 * SIGMA_VAL ^ 2) * NO_PERIODS
K_VAL = 1 / (2 * NO_PERIODS * SIGMA_VAL ^ 2)

ReDim TEMP_MATRIX(0 To NBINS, 1 To 10)
TEMP_MATRIX(0, 1) = "Price"
TEMP_MATRIX(0, 2) = "f(P)"
TEMP_MATRIX(0, 3) = "F(P)"
TEMP_MATRIX(0, 4) = "f(Po)"
TEMP_MATRIX(0, 5) = "F(Po)"
'Option  f(Option)   F(Option)

TEMP_MATRIX(0, 6) = "Premium at Time Left" 'EXPIRATION - NO PERIODS
TEMP_MATRIX(0, 7) = "f(O)"
TEMP_MATRIX(0, 8) = "F(O)"
TEMP_MATRIX(0, 9) = "Delta Expiration"
TEMP_MATRIX(0, 10) = "Premium Expiration"

'---------------------------------------------------------------------------------
'For those interviewing for trading derivative desk,
'remember to consider the Riemann Integral to go through the logic
'behind Itos Calculus <by the way, Louis de Branges de Bourcia claimed
'to have proved the Riemann hypothesis - there's a $1M prize for this achievement>
'---------------------------------------------------------------------------------

PRICE_VAL = MIN_PRICE
TEMP1_SUM = 0
TEMP2_SUM = 0
For i = 1 To NBINS
    TEMP_MATRIX(i, 1) = PRICE_VAL
    TEMP_MATRIX(i, 2) = K_VAL * A_VAL * TEMP_MATRIX(i, 1) * Exp(-K_VAL * (Log(TEMP_MATRIX(i, 1) / P0_VAL) - B_VAL) ^ 2)
                        
    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 2)
    
    TEMP_MATRIX(i, 3) = TEMP1_SUM
    PRICE_VAL = PRICE_VAL + DELTA_PRICE
Next i

For i = 1 To NBINS
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 3) / TEMP1_SUM
Next i

For i = 1 To NBINS
    TEMP_MATRIX(i, 4) = ""
    TEMP_MATRIX(i, 5) = ""
    TEMP_MATRIX(i, 9) = ""
    TEMP_MATRIX(i, 10) = ""
    If i > 1 Then
        If TEMP_MATRIX(i - 1, 1) <= P0_VAL And _
           TEMP_MATRIX(i, 1) > P0_VAL Then
            TEMP_MATRIX(i - 1, 4) = TEMP_MATRIX(i - 1, 2)
            TEMP_MATRIX(i - 1, 5) = TEMP_MATRIX(i - 1, 3)
            'Probability that stock Price will be less than
            'Current Price when T = NO_PERIOS is
            
            If OPTION_FLAG = 1 Then
                TEMP_MATRIX(i - 1, 9) = CALL_OPTION_DELTA_FUNC(TEMP_MATRIX(i - 1, 1), STRIKE, EXPIRATION, CASH_RATE, DIVID_RATE, SIGMA_VAL, CND_TYPE)
                TEMP_MATRIX(i - 1, 10) = EUROPEAN_CALL_OPTION_FUNC(TEMP_MATRIX(i - 1, 1), STRIKE, EXPIRATION, CASH_RATE, DIVID_RATE, SIGMA_VAL, CND_TYPE)
            Else
                TEMP_MATRIX(i - 1, 9) = PUT_OPTION_DELTA_FUNC(TEMP_MATRIX(i - 1, 1), STRIKE, EXPIRATION, CASH_RATE, DIVID_RATE, SIGMA_VAL, CND_TYPE)
                TEMP_MATRIX(i - 1, 10) = EUROPEAN_PUT_OPTION_FUNC(TEMP_MATRIX(i - 1, 1), STRIKE, EXPIRATION, CASH_RATE, DIVID_RATE, SIGMA_VAL, CND_TYPE)
            End If
        End If
    End If
    
    If OPTION_FLAG = 1 Then
        TEMP_MATRIX(i, 6) = EUROPEAN_CALL_OPTION_FUNC(TEMP_MATRIX(i, 1), STRIKE, EXPIRATION - NO_PERIODS, CASH_RATE, DIVID_RATE, SIGMA_VAL, CND_TYPE)
    Else
        TEMP_MATRIX(i, 6) = EUROPEAN_PUT_OPTION_FUNC(TEMP_MATRIX(i, 1), STRIKE, EXPIRATION - NO_PERIODS, CASH_RATE, DIVID_RATE, SIGMA_VAL, CND_TYPE)
    End If
    
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 2)
    TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 7)
    TEMP_MATRIX(i, 8) = TEMP2_SUM
    
Next i
For i = 1 To NBINS: TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 8) / TEMP2_SUM: Next i

ITO_OPTIONS_DISTRIBUTION3_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ITO_OPTIONS_DISTRIBUTION3_FUNC = Err.number
End Function


'Beautiful Algo for initial screening!!!! --> Target: Highest Gain%


'We pick a stock we're interested in and look at historical data, extractin g a Mean Annual Return,
'Rt and Volatility V. We note the current stock price, Po.

'Then, for any time T in the future, we use Ito's magic formula to give the Price Distribution:
'f(P) = 1 / (V * P * SQRT(2 * Pi() * T)) * Exp(-(1 / (2 * T * V ^ 2)) * (Ln(P / Po) -
'(Rt - 0.5 * V ^ 2) * T) ^ 2), where it is the cumulative distribution at time T in the future.
   
'Now we calculate the Expected Stock Price (at time T), using this distribution.
'To do this, we dividethe range of (future) stock prices into subintervals of length dP and
'pick equally spaced prices: P1, P2, ... Pn. (See the above picture.)

'Then we calculate:
'E[T] = { P1 f(P1) + P2 f(P2) + ... + Pn f(Pn) } dP
'The prices P1, P2 etc. run from left to right and include Po ... somewhere in the middle
   
'Now we consider a Call Option with Strike Price K and expiry time Te.
'At time T in the future, the time to expiry is: TT = Te - T.
'The value of the Option will depend upon P (the stock price at time T) and TT (the time left to
'expiry) and some Risk-free Rate Rf ... and other stuff

'We assume it's given by the magic Black-Scholes formula:
'C(P) = P * NormSDist((Ln(P / K) + (Rf + V ^ 2 / 2) * TT) / (V * SQRT(TT))) -
'       K * Exp(-Rf * TT) * NormSDist((Ln(Po / K) + (Rf + V ^ 2 / 2) * TT) / (V * SQRT(TT)) - V * SQRT(TT))

'For example, putting TT = Te and P = Po will give the (initial) option premium. We'll call that:
'Co = C(Po) = Po * NormSDist((Ln(Po / K) + (Rf + V ^ 2 / 2) * Te) / (V * SQRT(Te))) -
'     K * Exp(-Rf * Te) * NormSDist((Ln(Po / K) + (Rf + V ^ 2 / 2) * Te) / (V * SQRT(Te)) - V * SQRT(Te))

'Our intention is to keep track of the value of our Call as Time T progresses, and sell it when the
'Gain is acceptable.

'Indeed, at each future time T, we can calculate an Expected stock price and the corresponding
'Call value, C(E). We look intently at C(E) - Co, the Gain in option premium, and sell when it's
'a maximum.

'A maximum? When's that ... and what should one choose as the initial Strike price and ...?
'Aha! That 's our problem, eh?
'For a given stock, say GE, we have no control over the price Po or the annual return Rt or
'Volatility V or risk-free rate Rf,. but we have lots of Strikes to choose from and ....
'And lots of Time to expiry, right? Yes.That 's Te. So we look over the available choices for
'K and Te and follow each option for 0 < T < Te, with the help of Ito and Black-Scholes and ...

'And pick the best. That 's our intention.
'And you believe in all this stuff, right? The option premium, doesn't even match Black-Scholes!
'Uh ... well ... it's close enough. For example, GE is now at Po = $34.33 and if we use Rt and V
'generated by historical data over the last 5 years and Rf = 4% and Te = 14 weeks (from TODAY to
'June 15, 2007) and K = $32.50, we'd get Co = $2.78, so ...

'But Figure 1 says 2.70. Besides, you took the option that had the "best" fit to Black-Scholes !!
'Okay, look at these Gains as T progresses from 1 week to just before expiry, at Te - 1 = 13 weeks.

'There are charts of probability distributions f, and the cumulative distribution F, for both
'the stock price P and option premium C at some time T in the future.

'I mentioned the Expected Stock Price (at time T). There 'll also be an Expected Call premium as
'well ... and that'll change with T.

'See? As T increases to the Expiry time Te, you can get neato charts
'By "Gain", I mean the Expected % change in Call premium.
