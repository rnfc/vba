Attribute VB_Name = "FINAN_ASSET_SIMUL_ITO_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_ITO_PRICES_DISTRIBUTION1_FUNC
'DESCRIPTION   : Distribution of stock prices (and Expected Prices)
'LIBRARY       : FINAN_ASSET_SIMUL
'GROUP         : ITO
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASSET_ITO_PRICES_DISTRIBUTION1_FUNC(ByVal CURRENT_STOCK_PRICE As Double, _
ByVal STRIKE_PRICE As Double, _
ByVal CASH_RATE As Double, _
ByVal MEAN_VAL As Double, _
ByVal VOLATILITY_VAL As Double, _
Optional ByVal MIN_PERIOD As Double = 1, _
Optional ByVal DELTA_PERIOD As Double = 1, _
Optional ByVal NBINS_PERIOD As Long = 3, _
Optional ByVal MIN_PRICE As Double = 5, _
Optional ByVal DELTA_PRICE As Double = 2.29166666666667, _
Optional ByVal NBINS_PRICE As Long = 25, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long

Dim PI_VAL As Double

Dim TEMP_PRICE As Double
Dim TEMP_PERIOD As Double 'Periods into the future

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979
If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1
'--------------------------------------------------------------------------------------
Select Case OUTPUT
'--------------------------------------------------------------------------------------
Case 0
'--------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To NBINS_PRICE, 1 To NBINS_PERIOD + 1)
    TEMP_MATRIX(0, 1) = "PRICE"
    
    TEMP_PERIOD = MIN_PERIOD
    For j = 1 To NBINS_PERIOD
        TEMP_MATRIX(0, j + 1) = "DISTRIBUTION: PERIODS FORWARD = " & TEMP_PERIOD
        TEMP_PERIOD = TEMP_PERIOD + DELTA_PERIOD
    Next j
    
    TEMP_PRICE = MIN_PRICE
    For i = 1 To NBINS_PRICE
        TEMP_MATRIX(i, 1) = TEMP_PRICE
        TEMP_PERIOD = MIN_PERIOD
        For j = 1 To NBINS_PERIOD
            TEMP_MATRIX(i, j + 1) = 1 / (VOLATILITY_VAL * TEMP_PRICE * Sqr(2 * PI_VAL * TEMP_PERIOD)) * _
                                    Exp(-((Log(TEMP_PRICE / CURRENT_STOCK_PRICE) - _
                                    (MEAN_VAL - VOLATILITY_VAL ^ 2 / 2) * TEMP_PERIOD) ^ 2) / _
                                    (2 * TEMP_PERIOD * VOLATILITY_VAL ^ 2))
            TEMP_PERIOD = TEMP_PERIOD + DELTA_PERIOD
        Next j
        TEMP_PRICE = TEMP_PRICE + DELTA_PRICE
    Next i
'--------------------------------------------------------------------------------------
Case Else
'--------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To 3, 1 To NBINS_PERIOD + 1)
    TEMP_MATRIX(0, 1) = "DISTRIBUTION"
    TEMP_MATRIX(1, 1) = "OPTION PREMIUM"
    TEMP_MATRIX(2, 1) = "FUTURE PRICE"
    TEMP_MATRIX(3, 1) = "DISTRIBUTION"
    
    TEMP_PERIOD = MIN_PERIOD
    For j = 1 To NBINS_PERIOD
        TEMP_MATRIX(0, j + 1) = "PERIODS FORWARD = " & TEMP_PERIOD
        TEMP_MATRIX(1, j + 1) = BLACK_SCHOLES_OPTION_FUNC(CURRENT_STOCK_PRICE, STRIKE_PRICE, _
                                TEMP_PERIOD, CASH_RATE, VOLATILITY_VAL, OPTION_FLAG, CND_TYPE)
                                
        TEMP_PRICE = CURRENT_STOCK_PRICE * Exp((MEAN_VAL - VOLATILITY_VAL ^ 2 / 2) * TEMP_PERIOD)
        'Median of the distribution
        
        'TEMP_PRICE = CURRENT_STOCK_PRICE * Exp((MEAN_VAL - 0) * TEMP_PERIOD)
        'Mean of the distribution
        
        TEMP_MATRIX(2, j + 1) = TEMP_PRICE
        TEMP_MATRIX(3, j + 1) = 1 / (VOLATILITY_VAL * TEMP_PRICE * Sqr(2 * PI_VAL * TEMP_PERIOD)) * _
                                    Exp(-((Log(TEMP_PRICE / CURRENT_STOCK_PRICE) - _
                                    (MEAN_VAL - VOLATILITY_VAL ^ 2 / 2) * TEMP_PERIOD) ^ 2) / _
                                    (2 * TEMP_PERIOD * VOLATILITY_VAL ^ 2))
        TEMP_PERIOD = TEMP_PERIOD + DELTA_PERIOD
    Next j
'--------------------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------------------

ASSET_ITO_PRICES_DISTRIBUTION1_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_ITO_PRICES_DISTRIBUTION1_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_ITO_PRICES_DISTRIBUTION2_FUNC
'DESCRIPTION   : Probability that stock Price will be
'less than $P. We 're talking about Ito's stochastic
'equation: dP = m(t,P)dt + s(t,P)df
'LIBRARY       : FINAN_ASSET_SIMUL
'GROUP         : ITO
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASSET_ITO_PRICES_DISTRIBUTION2_FUNC( _
Optional ByVal P0_VAL As Double = 36.5, _
Optional ByVal P1_VAL As Double = 30.5, _
Optional ByVal P2_VAL As Double = 40.5, _
Optional ByVal MEAN_VAL As Double = 1.55676587633033E-03, _
Optional ByVal SIGMA_VAL As Double = 8.67851826879331E-03, _
Optional ByVal NO_PERIODS As Long = 5, _
Optional ByVal MIN_PRICE As Double = 33, _
Optional ByVal DELTA_PRICE As Double = 0.291666666666667, _
Optional ByVal NBINS As Long = 25)

'P0 --> Current Price
'P1 --> Upper Price
'P2 --> Lower Price
'SIGMA and MEAN ARE NOT ANNUALIZED

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

Dim TEMP_SUM As Double
Dim TEMP_PRICE As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979

TEMP_SUM = 0

A_VAL = 1 / (SIGMA_VAL * Sqr(2 * PI_VAL * NO_PERIODS))
B_VAL = (MEAN_VAL - 0.5 * SIGMA_VAL ^ 2) * NO_PERIODS
K_VAL = 1 / (2 * NO_PERIODS * SIGMA_VAL ^ 2)

ReDim TEMP_MATRIX(0 To NBINS, 1 To 5)
TEMP_MATRIX(0, 1) = "Price"
TEMP_MATRIX(0, 2) = "f(P)"
TEMP_MATRIX(0, 3) = "F(P)"
TEMP_MATRIX(0, 4) = "f(Po)"
TEMP_MATRIX(0, 5) = "F(Po)"

'---------------------------------------------------------------------------------
'For those interviewing for trading derivative desk,
'remember to consider the Riemann Integral to go through the logic
'behind Itos Calculus <by the way, Louis de Branges de Bourcia claimed
'to have proved the Riemann hypothesis - there's a $1M prize for this achievement>
'---------------------------------------------------------------------------------

TEMP_PRICE = MIN_PRICE
TEMP_SUM = 0
For i = 1 To NBINS
    TEMP_MATRIX(i, 1) = TEMP_PRICE
    TEMP_MATRIX(i, 2) = A_VAL / TEMP_MATRIX(i, 1) * Exp(-K_VAL * _
                        (Log(TEMP_MATRIX(i, 1) / P0_VAL) - B_VAL) ^ 2)
                        'PROBABILITY_DENSITY
                        
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 2)
    
    TEMP_MATRIX(i, 3) = TEMP_SUM * DELTA_PRICE
    TEMP_MATRIX(i, 4) = ""
    TEMP_MATRIX(i, 5) = ""
    If i > 1 Then
        If TEMP_MATRIX(i - 1, 1) <= P0_VAL And _
           TEMP_MATRIX(i, 1) > P0_VAL Then
            TEMP_MATRIX(i - 1, 4) = TEMP_MATRIX(i - 1, 2)
            TEMP_MATRIX(i - 1, 5) = TEMP_MATRIX(i - 1, 3)
            'Probability that stock Price will be less than
            'Current Price when T = NO_PERIOS is
        End If
        If TEMP_MATRIX(i - 1, 1) <= P1_VAL And _
           TEMP_MATRIX(i, 1) > P1_VAL Then
            TEMP_MATRIX(i - 1, 4) = TEMP_MATRIX(i - 1, 2)
            TEMP_MATRIX(i - 1, 5) = TEMP_MATRIX(i - 1, 3)
            'Probability that stock Price will be less than
            'Current Price when T = NO_PERIOS is
        End If
        If TEMP_MATRIX(i - 1, 1) <= P2_VAL And _
           TEMP_MATRIX(i, 1) > P2_VAL Then
            TEMP_MATRIX(i - 1, 4) = TEMP_MATRIX(i - 1, 2)
            TEMP_MATRIX(i - 1, 5) = TEMP_MATRIX(i - 1, 3)
            'Probability that stock Price will be less than
            'Current Price when T = NO_PERIOS is
        End If
    End If
    
    TEMP_PRICE = TEMP_PRICE + DELTA_PRICE
Next i


ASSET_ITO_PRICES_DISTRIBUTION2_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_ITO_PRICES_DISTRIBUTION2_FUNC = Err.number
End Function

Function ITO_PROBABILITY_DENSITY_FUNC(ByVal P0_VAL As Double, _
ByVal P1_VAL As Double, _
ByVal MEAN_VAL As Double, _
ByVal SIGMA_VAL As Double, _
ByVal NO_PERIODS As Long)

Dim PI_VAL As Double

On Error GoTo ERROR_LABEL

PI_VAL = 2 * 3.14159265358979

ITO_PROBABILITY_DENSITY_FUNC = _
    Exp(-((Log(P1_VAL / P0_VAL) - _
    (MEAN_VAL - SIGMA_VAL ^ 2 / 2) * NO_PERIODS) ^ 2 / _
    (2 * NO_PERIODS * SIGMA_VAL ^ 2))) / _
    (SIGMA_VAL * P1_VAL * Sqr(PI_VAL * NO_PERIODS))

Exit Function
ERROR_LABEL:
ITO_PROBABILITY_DENSITY_FUNC = Err.number
End Function
