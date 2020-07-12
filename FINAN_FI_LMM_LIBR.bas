Attribute VB_Name = "FINAN_FI_LMM_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : LMM_CAPLET_MC_FUNC

'DESCRIPTION   : This function implements 1 factor LMM and calculates caplet
'price using monte carlo simulation as described in Equation
'24.17 Chap 24, Options, futures and derivates by John Hull
'Further the following things are additionally computed to
'compare the value obtained using Black formula:

'1.Caplet volatility for input in black formula is obtained by
'integrating the forward rate volatilities

'2.Discount factor for input in black formula is obtained by
'compounding the forward rates till expiry of the caplet
'Finally the 2 values are compared and MC simulation values
'seem to converge to analytical price of caplet

'LIBRARY       : FIXED_INCOME
'GROUP         : LMM
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function LMM_CAPLET_MC_FUNC(ByRef RATE_RNG As Variant, _
ByRef SIGMA_RNG As Variant, _
Optional ByVal PRINCIPAL As Double = 10000, _
Optional ByVal ACCR_PERIOD As Double = 1, _
Optional ByVal STEP_SIZE As Double = 1, _
Optional ByVal STRIKE_RATE As Double = 0.05, _
Optional ByVal START_TIME As Double = 4, _
Optional ByVal nLOOPS As Long = 100000)

'-------------------------------------------------------
'Zero Curve(RATE_RNG) / Forward Volatilities (SIGMA_RNG)
'-------------------------------------------------------
'Time Forward rate        T-t    Vol
' 0   0.05                  1   0.155
' 1   0.05                  2   0.2064
' 2   0.05                  3   0.1721
' 3   0.05                  4   0.1722
' 4   0.05                  5   0.1525
' 5   0.05                  6   0.1415
' 6   0.05                  7   0.1298
' 7   0.05                  8   0.1381
' 8   0.05                  9   0.136
' 9   0.05                 10   0.134
'10   0.05                 11   0.2
'-------------------------------------------------------

Dim hh As Long
Dim ii As Long
Dim jj As Long
Dim kk As Long
Dim ll As Long

Dim NSIZE As Long

Dim DF_VAL As Double
Dim PEG_VAL As Double
Dim DRIFT_VAL As Double
Dim SIGMA_VAL As Double
Dim SCHOCK_VAL As Double

Dim SUM_VAL As Double
Dim MEAN_VAL As Double

Dim FORWARD_ARR As Variant
Dim FORWARD_VECTOR As Variant

Dim SIGMA_ARR As Variant
Dim SIGMA_VECTOR As Variant

Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant
Dim CTEMP_ARR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'ACCR_PERIOD: accrual period in Yrs
'STEP_SIZE: incremental step size for advancing current time
'STRIKE_RATE: strike rate for caplet
'START_TIME: Start Time for caplet (Yrs)

PEG_VAL = FLOOR_FUNC(START_TIME / ACCR_PERIOD, 1)
'start peg for caplet
'eg., startpeg=4 implies caplet begins from 4th time period
'and ends in 5th period discretized by ACCR_PERIOD
NSIZE = PEG_VAL + 1

Randomize

FORWARD_VECTOR = RATE_RNG
SIGMA_VECTOR = SIGMA_RNG

ReDim FORWARD_ARR(0 To NSIZE)
ReDim SIGMA_ARR(0 To NSIZE)

jj = 1
For ii = 0 To NSIZE - 1
  FORWARD_ARR(ii) = FORWARD_VECTOR(jj, 1)
  SIGMA_ARR(ii) = SIGMA_VECTOR(jj, 1)
  jj = jj + 1
Next ii

ATEMP_ARR = FORWARD_ARR
BTEMP_ARR = FORWARD_ARR
ReDim CTEMP_ARR(0 To nLOOPS - 1)
'Fschockvec = VECTOR_RANDOM_BOX_MULLER_FUNC(n - 1)
'counter = 1

'I am simulating 1 factor Libor market model as described in
'equations in Hull's book "Options,Futures and Derivatives" .
'I am calculating price of a caplet using MC simulation and
'compare it with price obtained from Black formula by doing
'the following :
    '1.Take input volatility for Black formula as weighted sum
    'of forward volatities till start period of caplet
    '2.Use zero coupon curve for discounting till expiry of caplet

For ll = 0 To nLOOPS - 1
  FORWARD_ARR = BTEMP_ARR
  ATEMP_ARR = FORWARD_ARR
  hh = 0
  jj = 0
  Do While jj < NSIZE
    jj = FLOOR_FUNC(hh / ACCR_PERIOD, 1)
    
    'If NSIZE - jj - 1 > 0 Then
    '  Fschockvec = VECTOR_RANDOM_BOX_MULLER_FUNC(NSIZE - jj - 1)
    'End If
    
    
    SCHOCK_VAL = NORMSINV_FUNC(Rnd(), 0, 1, 0)
    For kk = jj + 1 To NSIZE - 1
      'Fschock = Fschockvec(counter)
      'Fschock = Fschockvec(1)
      'counter = counter + 1
       
'One thing I have noted during the experiment is that the first
'part of drift term tends to make a mismatch and caplet price
'from MC ends up higher as compared to caplet price from analytic
'formula. However if I remove the 1st part of drift term, the 2
'prices match to some extent.

'It looks to me that I have a problem somewhere in discounting.
'I would not discount the payoff using the original zero curve.
'It should be the actual simulated values during that time.

      DRIFT_VAL = 0
      For ii = jj + 1 To kk
         DRIFT_VAL = DRIFT_VAL + ACCR_PERIOD * _
                BTEMP_ARR(ii) * SIGMA_ARR(ii - jj - 1) * _
                SIGMA_ARR(kk - jj - 1) / (1 + ACCR_PERIOD * BTEMP_ARR(ii))
      Next ii
      'DRIFT_VAL = 0
      
      FORWARD_ARR(kk) = ATEMP_ARR(kk) * Exp((DRIFT_VAL - 0.5 * _
            SIGMA_ARR(kk - jj - 1) * SIGMA_ARR(kk - jj - 1)) * _
            STEP_SIZE + SIGMA_ARR(kk - jj - 1) * SCHOCK_VAL * STEP_SIZE ^ 0.5)
    Next kk
    
    ATEMP_ARR = FORWARD_ARR
    
    hh = hh + STEP_SIZE
  Loop
  DF_VAL = 1
  For ii = 0 To PEG_VAL
    DF_VAL = DF_VAL / (1 + ACCR_PERIOD * FORWARD_ARR(ii))
  Next ii
 
  CTEMP_ARR(ll) = PRINCIPAL * DF_VAL * ACCR_PERIOD * _
                MAXIMUM_FUNC(FORWARD_ARR(PEG_VAL) - STRIKE_RATE, 0)

Next ll

SUM_VAL = 0
DF_VAL = 1
For ii = 0 To PEG_VAL
  DF_VAL = DF_VAL / (1 + ACCR_PERIOD * BTEMP_ARR(ii))
Next ii
For ii = 0 To nLOOPS - 1
  SUM_VAL = SUM_VAL + CTEMP_ARR(ii)
Next ii
MEAN_VAL = SUM_VAL / nLOOPS

SIGMA_VAL = 0 'Black Volat Value
For ii = 0 To PEG_VAL - 1
    SIGMA_VAL = SIGMA_VAL + SIGMA_ARR(ii) * SIGMA_ARR(ii)
Next ii
SIGMA_VAL = (SIGMA_VAL / (PEG_VAL * ACCR_PERIOD)) ^ 0.5

ReDim TEMP_MATRIX(1 To 7, 1 To 2)
TEMP_MATRIX(1, 1) = "Forward rate applicable from T to T+1"
TEMP_MATRIX(2, 1) = "Black volatility"
TEMP_MATRIX(3, 1) = "Discount factor"
TEMP_MATRIX(4, 1) = "d1"
TEMP_MATRIX(5, 1) = "d2"
TEMP_MATRIX(6, 1) = "Caplet Price using Black formula"
TEMP_MATRIX(7, 1) = "Caplet Price using 1F simulation"

TEMP_MATRIX(1, 2) = BTEMP_ARR(PEG_VAL)
TEMP_MATRIX(2, 2) = SIGMA_VAL
TEMP_MATRIX(3, 2) = DF_VAL
TEMP_MATRIX(4, 2) = ((Log(BTEMP_ARR(PEG_VAL) / STRIKE_RATE) / _
                    Log(Exp(1))) + SIGMA_VAL * SIGMA_VAL * 0.5) / _
                    (SIGMA_VAL * START_TIME ^ 0.5)
TEMP_MATRIX(5, 2) = TEMP_MATRIX(4, 2) - SIGMA_VAL * START_TIME ^ 0.5
TEMP_MATRIX(6, 2) = PRINCIPAL * ACCR_PERIOD * DF_VAL * (BTEMP_ARR(PEG_VAL) * _
                            NORMSDIST_FUNC(TEMP_MATRIX(4, 2), 0, 1, 0) - _
                    STRIKE_RATE * NORMSDIST_FUNC(TEMP_MATRIX(5, 2), 0, 1, 0))
TEMP_MATRIX(7, 2) = MEAN_VAL

LMM_CAPLET_MC_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
LMM_CAPLET_MC_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LMM_FORWARD_SPOT_MC_FUNC
'DESCRIPTION   : Foward measure vs Spot measure vs Black formula
'for pricing in Libor market model
'LIBRARY       : FIXED_INCOME
'GROUP         : LMM
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function LMM_FORWARD_SPOT_MC_FUNC(ByRef RATE_RNG As Variant, _
ByRef SIGMA_RNG As Variant, _
ByVal PRINCIPAL As Double, _
ByVal ACCRUAL_TENOR As Double, _
ByVal STRIKE_RATE As Double, _
ByVal START_TENOR As Double, _
Optional ByVal nLOOPS As Long = 10000, _
Optional ByVal CND_TYPE As Integer = 0)
'Caplet Specifications
'ACCRUAL_TENOR:Accrual period (Yrs)
'START_TENOR: Start Time (yrs)

'This function calculates price of a caplet in Libor market model.
'Inputs are forward rate and forward voltilties term structure and
'caplet specifications.

'It produces an ouput of caplet prices using the following methods:

'1. Black formula
'Caplet price is given by
'Cn(0)=BlackFormula(Ln(0),sigmaN(N),Tn,K,ACCRUAL_TENOR*B_{n+1}(0))
'where
'n : index of caplet start time
'sigmaN(N) : the integral of forward rate volatilities from 0 to n
'Tn : Caplet start time in years
'K:  Strike
'ACCRUAL_TENOR : accrual period in years
'B_n(0) : price today of zero coupon bond expiring at time index n
'Ln(0) : Forward libor starting from Tn and ending at T{n+1) as seen today

'2.Forward measure
'Libor rates are evolved according to (3.114) as mentioned in Glasserman's
'book - Monte Carlo Methods in Financial Engineering Lm is a martingale under
'forward measure for maturity T_{m+1}. We assume sigmaM as deterministic and
'hence Lm(t) is lognormally distributed according to LN(-sigmaM*sigmaM/2,
'sigmaM*sigmaM)
'where sigmaM is as defined above

'3.Spot measure
'Forward rates are evolved from t=0 to Tn according to (22) as metioned in
'paper FORWARD RATE VOLATILITIES, SWAP RATE VOLATILITIES, AND THE IMPLEMENTATION
'OF THE LIBOR MARKET MODEL (John Hull and Alan White)

'The time step size taken is equal to the accrual
'step size based on assumption that DRIFT_VAL is constant within each accrual
'period. It works fine even for period of 1 year as mentioned in Hull

'RATE_RNG & SIGMA_RNG: Input Term Structure --> Forward rate / Forward Vol


Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim NSIZE As Long
Dim COUNTER As Long

Dim D1_VAL As Double
Dim D2_VAL As Double

Dim DRIFT_VAL As Double
Dim SIGMA_VAL As Double
Dim DELTA_TENOR As Double
Dim START_PEG_VAL As Double
Dim FORWARD_STRIKE_VAL As Double

Dim MEAN_FORWARD_VAL As Double
Dim MEAN_SPOT_VAL As Double
Dim BLACK_PRICE_VAL As Double
Dim FORWARD_VAL As Double

Dim DRIFT_SUM_VAL As Double
Dim BLACK_SIGMA_VAL As Double
Dim FIRST_DISCOUNT_FACTOR As Double
Dim SECOND_DISCOUNT_FACTOR As Double

Dim SPOT_SHOCK_VAL As Double
Dim FORWARD_SHOCK_VAL As Double
Dim OPTION_MATURITY_VAL As Double
Dim PAYOFF_SUM_SPOT_VAL As Double
Dim PAYOFF_SUM_FORWARD_VAL As Double

Dim RATE_VECTOR As Variant
Dim SIGMA_VECTOR As Variant
Dim OLD_FORWARD_ARR As Variant
Dim ORIGINAL_FORWARD_ARR As Variant
Dim FORWARD_SHOCK_ARR As Variant
Dim FORWARD_SPOT_SHOCK_ARR As Variant

On Error GoTo ERROR_LABEL
  
RATE_VECTOR = RATE_RNG
If UBound(RATE_VECTOR, 1) = 1 Then: _
  RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(RATE_VECTOR)
SIGMA_VECTOR = SIGMA_RNG
If UBound(SIGMA_VECTOR, 1) = 1 Then: _
  SIGMA_VECTOR = MATRIX_TRANSPOSE_FUNC(SIGMA_VECTOR)
  
START_PEG_VAL = FLOOR_FUNC(START_TENOR / _
ACCRUAL_TENOR, 1) 'start peg for caplet
'eg., startpeg=4 implies caplet begins from 4th time period
'and ends in 5th period discretized by ACCRUAL_TENOR

Randomize

NSIZE = START_PEG_VAL + 1
ReDim FORWARD_ARR(0 To NSIZE)
ReDim SIGMA_ARR(0 To NSIZE)
j = 1
For i = 0 To NSIZE
  FORWARD_ARR(i) = RATE_VECTOR(j, 1)
  SIGMA_ARR(i) = SIGMA_VECTOR(j + 1, 1)
  j = j + 1
Next i

OLD_FORWARD_ARR = FORWARD_ARR
ORIGINAL_FORWARD_ARR = FORWARD_ARR
BLACK_SIGMA_VAL = 0
For i = 0 To START_PEG_VAL - 1
  BLACK_SIGMA_VAL = BLACK_SIGMA_VAL + SIGMA_ARR(i) * SIGMA_ARR(i)
Next i
BLACK_SIGMA_VAL = (BLACK_SIGMA_VAL / START_PEG_VAL) ^ 0.5

FIRST_DISCOUNT_FACTOR = 1
For i = 0 To START_PEG_VAL - 1
  FIRST_DISCOUNT_FACTOR = FIRST_DISCOUNT_FACTOR / (1 + ACCRUAL_TENOR * _
              ORIGINAL_FORWARD_ARR(i))
Next i
 
DELTA_TENOR = START_PEG_VAL * ACCRUAL_TENOR
DRIFT_VAL = -0.5 * BLACK_SIGMA_VAL * BLACK_SIGMA_VAL * DELTA_TENOR
SIGMA_VAL = BLACK_SIGMA_VAL * DELTA_TENOR ^ 0.5
PAYOFF_SUM_FORWARD_VAL = 0
FORWARD_SHOCK_ARR = VECTOR_RANDOM_BOX_MULLER_FUNC(nLOOPS / 2)

FORWARD_SPOT_SHOCK_ARR = VECTOR_RANDOM_BOX_MULLER_FUNC(nLOOPS * START_PEG_VAL / 2)

SECOND_DISCOUNT_FACTOR = FIRST_DISCOUNT_FACTOR / (1 + ACCRUAL_TENOR * _
      ORIGINAL_FORWARD_ARR(START_PEG_VAL))
PAYOFF_SUM_SPOT_VAL = 0
COUNTER = 1

For h = 1 To nLOOPS / 2
  
  FORWARD_SHOCK_VAL = FORWARD_SHOCK_ARR(h)
  FORWARD_VAL = ORIGINAL_FORWARD_ARR(START_PEG_VAL) * _
      Exp(DRIFT_VAL + SIGMA_VAL * FORWARD_SHOCK_VAL)
  PAYOFF_SUM_FORWARD_VAL = PAYOFF_SUM_FORWARD_VAL + _
      PRINCIPAL * ACCRUAL_TENOR * SECOND_DISCOUNT_FACTOR * _
      MAXIMUM_FUNC(FORWARD_VAL - STRIKE_RATE, 0)
  
  FORWARD_ARR = ORIGINAL_FORWARD_ARR
  OLD_FORWARD_ARR = ORIGINAL_FORWARD_ARR
  For j = 0 To START_PEG_VAL - 1 'loop for time evolution
    SPOT_SHOCK_VAL = FORWARD_SPOT_SHOCK_ARR(COUNTER)
    'all forward rates should have same shock for 1F
    COUNTER = COUNTER + 1
    For k = j + 1 To START_PEG_VAL 'loop for each forward rate
      DRIFT_SUM_VAL = 0
      For i = j + 1 To k
        DRIFT_SUM_VAL = DRIFT_SUM_VAL + ACCRUAL_TENOR * _
                      OLD_FORWARD_ARR(i) * SIGMA_ARR(i - j - 1) * _
                      SIGMA_ARR(k - j - 1) / (1 + ACCRUAL_TENOR * _
                      OLD_FORWARD_ARR(i))
      Next i
      FORWARD_ARR(k) = OLD_FORWARD_ARR(k) * _
      Exp((DRIFT_SUM_VAL - 0.5 * SIGMA_ARR(k - j - 1) * _
      SIGMA_ARR(k - j - 1)) * ACCRUAL_TENOR + _
      SIGMA_ARR(k - j - 1) * SPOT_SHOCK_VAL * ACCRUAL_TENOR ^ 0.5)
    Next k
    OLD_FORWARD_ARR = FORWARD_ARR
  Next j
  FIRST_DISCOUNT_FACTOR = 1
  For i = 0 To START_PEG_VAL
    FIRST_DISCOUNT_FACTOR = FIRST_DISCOUNT_FACTOR / (1 + _
    ACCRUAL_TENOR * FORWARD_ARR(i))
  Next i
  PAYOFF_SUM_SPOT_VAL = PAYOFF_SUM_SPOT_VAL + PRINCIPAL * _
          FIRST_DISCOUNT_FACTOR * ACCRUAL_TENOR * _
          MAXIMUM_FUNC(FORWARD_ARR(START_PEG_VAL) - STRIKE_RATE, 0)


  COUNTER = COUNTER - START_PEG_VAL
  FORWARD_SHOCK_VAL = -FORWARD_SHOCK_VAL
  FORWARD_VAL = ORIGINAL_FORWARD_ARR(START_PEG_VAL) * _
              Exp(DRIFT_VAL + SIGMA_VAL * FORWARD_SHOCK_VAL)
  PAYOFF_SUM_FORWARD_VAL = PAYOFF_SUM_FORWARD_VAL + PRINCIPAL * _
              ACCRUAL_TENOR * SECOND_DISCOUNT_FACTOR * _
              MAXIMUM_FUNC(FORWARD_VAL - STRIKE_RATE, 0)
  
  FORWARD_ARR = ORIGINAL_FORWARD_ARR
  OLD_FORWARD_ARR = ORIGINAL_FORWARD_ARR
  For j = 0 To START_PEG_VAL - 1
    SPOT_SHOCK_VAL = -FORWARD_SPOT_SHOCK_ARR(COUNTER)
    'all forward rates should have same shock for 1F
    COUNTER = COUNTER + 1
    For k = j + 1 To START_PEG_VAL
      DRIFT_SUM_VAL = 0
      For i = j + 1 To k
        DRIFT_SUM_VAL = DRIFT_SUM_VAL + ACCRUAL_TENOR * _
                          OLD_FORWARD_ARR(i) * SIGMA_ARR(i - j - 1) * _
                          SIGMA_ARR(k - j - 1) / (1 + ACCRUAL_TENOR * _
                          OLD_FORWARD_ARR(i))
      Next i
      FORWARD_ARR(k) = OLD_FORWARD_ARR(k) * Exp((DRIFT_SUM_VAL - _
                      0.5 * SIGMA_ARR(k - j - 1) * SIGMA_ARR(k - j - 1)) * _
                      ACCRUAL_TENOR + SIGMA_ARR(k - j - 1) * _
                      SPOT_SHOCK_VAL * ACCRUAL_TENOR ^ 0.5)
    Next k
    OLD_FORWARD_ARR = FORWARD_ARR
  Next j
  FIRST_DISCOUNT_FACTOR = 1
  For i = 0 To START_PEG_VAL
    FIRST_DISCOUNT_FACTOR = FIRST_DISCOUNT_FACTOR / (1 + _
                            ACCRUAL_TENOR * FORWARD_ARR(i))
  Next i
  PAYOFF_SUM_SPOT_VAL = PAYOFF_SUM_SPOT_VAL + PRINCIPAL * _
                      FIRST_DISCOUNT_FACTOR * ACCRUAL_TENOR * _
                      MAXIMUM_FUNC(FORWARD_ARR(START_PEG_VAL) - STRIKE_RATE, 0)
Next h

FIRST_DISCOUNT_FACTOR = 1
For i = 0 To START_PEG_VAL
  FIRST_DISCOUNT_FACTOR = FIRST_DISCOUNT_FACTOR / (1 + _
                  ACCRUAL_TENOR * ORIGINAL_FORWARD_ARR(i))
Next i

MEAN_FORWARD_VAL = PAYOFF_SUM_FORWARD_VAL / nLOOPS
MEAN_SPOT_VAL = PAYOFF_SUM_SPOT_VAL / nLOOPS
FORWARD_STRIKE_VAL = ORIGINAL_FORWARD_ARR(START_PEG_VAL)
OPTION_MATURITY_VAL = START_PEG_VAL * ACCRUAL_TENOR
D1_VAL = Log(FORWARD_STRIKE_VAL / STRIKE_RATE) + _
  BLACK_SIGMA_VAL * BLACK_SIGMA_VAL * _
  OPTION_MATURITY_VAL * 0.5
D2_VAL = D1_VAL - BLACK_SIGMA_VAL * OPTION_MATURITY_VAL ^ 0.5
BLACK_PRICE_VAL = PRINCIPAL * ACCRUAL_TENOR * FIRST_DISCOUNT_FACTOR * _
          (FORWARD_STRIKE_VAL * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE_RATE * _
          CND_FUNC(D2_VAL, CND_TYPE))

'BLACK Formula / SPOT Measure / FORWARD Measure
LMM_FORWARD_SPOT_MC_FUNC = Array(BLACK_PRICE_VAL, MEAN_SPOT_VAL, _
                        MEAN_FORWARD_VAL)
'Caplet price output

Exit Function
ERROR_LABEL:
LMM_FORWARD_SPOT_MC_FUNC = Err.number
End Function
