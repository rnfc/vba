Attribute VB_Name = "FINAN_DERIV_HESTON_LIBR"
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
Public PUB_HEST_TENOR As Double 'Maturity of volatility swap
Public PUB_HEST_KAPPA As Double 'Volatility mean reversion rate
Public PUB_HEST_LAMBDA As Double 'Volatility of volatility
Public PUB_HEST_INIT_SIGMA As Double 'Initial volatility
Public PUB_HEST_LT_SIGMA As Double 'Long term mean volatility
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : HESTON_VOLATILITY_SWAP_FUNC
'DESCRIPTION   : Volatilty swap pricing in Heston Model
'REFERENCE: http://www.math.nyu.edu/fellows_fin_math/gatheral/lecture7_2005.pdf
'LIBRARY       : DERIVATIVES
'GROUP         : HESTON
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Function HESTON_VOLATILITY_SWAP_FUNC(ByVal INIT_SIGMA As Double, _
ByVal LT_SIGMA As Double, _
ByVal KAPPA As Double, _
ByVal LAMBDA As Double, _
ByVal TENOR As Double, _
Optional ByVal LOWER_BOUND As Double = 0.0001, _
Optional ByVal UPPER_BOUND As Double = 100000, _
Optional ByVal nLOOPS As Long = 10000, _
Optional ByVal tolerance As Double = 0.00001)

Dim PI_VAL As Double

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979
PUB_HEST_INIT_SIGMA = INIT_SIGMA 'Initial volatility
PUB_HEST_LT_SIGMA = LT_SIGMA 'Long term mean volatility
PUB_HEST_KAPPA = KAPPA 'Volatility mean reversion rate
PUB_HEST_LAMBDA = LAMBDA 'Volatility of volatility
PUB_HEST_TENOR = TENOR 'Maturity of volatility swap

HESTON_VOLATILITY_SWAP_FUNC = 1 / (2 * Sqr(PI_VAL * TENOR)) * GAULEG7_INTEGRATION_FUNC("HESTON_VOLATILITY_SWAP_INTEGRAND_FUNC", LOWER_BOUND, UPPER_BOUND, tolerance, nLOOPS)
  
Exit Function
ERROR_LABEL:
HESTON_VOLATILITY_SWAP_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HESTON_VOLATILITY_SWAP_INTEGRAND_FUNC
'DESCRIPTION   : Laplace Integrand
'LIBRARY       : DERIVATIVES
'GROUP         : HESTON
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Private Function HESTON_VOLATILITY_SWAP_INTEGRAND_FUNC(ByVal X_VAL As Double)
  
Dim PHI_VAL As Double

Dim A_VAL As Double
Dim B_VAL As Double
Dim C_VAL As Double

On Error GoTo ERROR_LABEL

PHI_VAL = Sqr(PUB_HEST_KAPPA ^ 2 + 2 * X_VAL * PUB_HEST_LAMBDA ^ 2)

A_VAL = ((2 * PHI_VAL * Exp((PHI_VAL + PUB_HEST_KAPPA) * PUB_HEST_TENOR * 0.5)) / ((PHI_VAL + PUB_HEST_KAPPA) * (Exp(PHI_VAL * PUB_HEST_TENOR) - 1) + 2 * PHI_VAL)) ^ ((2 * PUB_HEST_KAPPA * PUB_HEST_LT_SIGMA) / PUB_HEST_LAMBDA ^ 2)
B_VAL = (2 * (Exp(PHI_VAL * PUB_HEST_TENOR) - 1)) / ((PHI_VAL + PUB_HEST_KAPPA) * (Exp(PHI_VAL * PUB_HEST_TENOR) - 1) + 2 * PHI_VAL)
C_VAL = A_VAL * Exp(-(B_VAL * (X_VAL) * PUB_HEST_INIT_SIGMA))

HESTON_VOLATILITY_SWAP_INTEGRAND_FUNC = (1 - C_VAL) / (X_VAL ^ 1.5)

Exit Function
ERROR_LABEL:
HESTON_VOLATILITY_SWAP_INTEGRAND_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HESTON_EUROPEAN_MC_CALL_OPTION_FUNC
'DESCRIPTION   : This Function calculates euoropean call option prices in Heston
'Model using Monte carlo simulation.
'Resulting values are compared to table 1 in
'http://www.wilmott.com/pdfs/051111_mikh.pdf

'LIBRARY       : DERIVATIVES
'GROUP         : HESTON
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Function HESTON_EUROPEAN_MC_CALL_OPTION_FUNC(ByVal SPOT As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal INIT_SIGMA As Double, _
ByVal LT_SIGMA As Double, _
ByVal KAPPA As Double, _
ByVal LAMBDA As Double, _
ByVal RHO As Double, _
Optional ByVal MIN_STRIKE As Double = 0.5, _
Optional ByVal MAX_STRIKE As Double = 1.5, _
Optional ByVal NBINS_STRIKE As Long = 5, _
Optional ByVal nSTEPS As Long = 150, _
Optional ByVal nLOOPS As Long = 200, _
Optional ByVal MIN_SIGMA As Double = 0.00000000000001, _
Optional ByVal MAX_SIGMA As Double = 10, _
Optional ByVal CND_TYPE As Integer = 0)


'Initial Volatility Level(INIT_SIGMA)
'Long term volatility (LT_SIGMA)
'Mean reversion speed of volatility (kappa)
'Volatility of volatility (lambda)
'Correlation coefficient (rho)

'MIN_STRIKE: Min STRIKE
'MAX_STRIKE: Max STRIKE
'NBINS_STRIKE: No of TEMP_STRIKE steps
'nSTEPS: No of time steps
'nLOOPS: No of MC paths

Dim i As Long
Dim j As Long

Dim V1_VAL As Double
Dim V2_VAL As Double

Dim S1_VAL As Double
Dim S2_VAL As Double

Dim DW1_VAL As Double
Dim DW2_VAL As Double

Dim TEMP_SUM As Double
Dim TEMP_MULT As Double
Dim TEMP_STRIKE As Double

Dim DELTA_TENOR As Double
Dim SQR_RHO_VAL As Double

Dim ATEMP_ARR As Variant
Dim BTEMP_ARR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

Randomize
nLOOPS = FLOOR_FUNC(nLOOPS / 2, 1) * 2
    
DELTA_TENOR = TENOR / nSTEPS
  
SQR_RHO_VAL = (1 - RHO * RHO) ^ 0.5

TEMP_SUM = 0
ReDim TEMP_MATRIX(0 To NBINS_STRIKE, 1 To 3)

TEMP_MULT = (MAX_STRIKE - MIN_STRIKE) / (NBINS_STRIKE - 1)
TEMP_STRIKE = MIN_STRIKE

TEMP_MATRIX(0, 1) = "STRIKE"
TEMP_MATRIX(0, 2) = "OPT_PRICE"
TEMP_MATRIX(0, 3) = "IMPLIED_SIGMA"

For i = 1 To NBINS_STRIKE
    TEMP_MATRIX(i, 1) = TEMP_STRIKE
    TEMP_MATRIX(i, 2) = 0
    TEMP_MATRIX(i, 3) = 0
    TEMP_STRIKE = TEMP_STRIKE + TEMP_MULT
Next i

For i = 1 To nLOOPS / 2
  V1_VAL = INIT_SIGMA
  S1_VAL = SPOT
  V2_VAL = INIT_SIGMA
  S2_VAL = SPOT
  
  ATEMP_ARR = VECTOR_RANDOM_BOX_MULLER_FUNC(nSTEPS)
  BTEMP_ARR = VECTOR_RANDOM_BOX_MULLER_FUNC(nSTEPS)
  
  For j = 1 To nSTEPS
    DW1_VAL = ATEMP_ARR(j)
    DW2_VAL = RHO * DW1_VAL + SQR_RHO_VAL * BTEMP_ARR(j)
    S1_VAL = S1_VAL + RATE * S1_VAL * DELTA_TENOR + _
            (V1_VAL * DELTA_TENOR) ^ 0.5 * S1_VAL * DW1_VAL
    V1_VAL = V1_VAL + KAPPA * (LT_SIGMA - V1_VAL) * _
            DELTA_TENOR + LAMBDA * (V1_VAL * DELTA_TENOR) ^ 0.5 * DW2_VAL
    If V1_VAL < 0 Then: V1_VAL = -V1_VAL  'reflect if 0 is reached
    DW1_VAL = -DW1_VAL
    DW2_VAL = -DW2_VAL
    S2_VAL = S2_VAL + RATE * S2_VAL * DELTA_TENOR + _
            (V2_VAL * DELTA_TENOR) ^ 0.5 * S2_VAL * DW1_VAL
    V2_VAL = V2_VAL + KAPPA * (LT_SIGMA - V2_VAL) * _
            DELTA_TENOR + LAMBDA * (V2_VAL * DELTA_TENOR) ^ 0.5 * DW2_VAL
    If V2_VAL < 0 Then: V2_VAL = -V2_VAL  'reflect if 0 is reached
  Next j
  
  For j = 1 To NBINS_STRIKE
    TEMP_STRIKE = TEMP_MATRIX(j, 1)
    TEMP_MATRIX(j, 2) = TEMP_MATRIX(j, 2) + MAXIMUM_FUNC(S1_VAL - TEMP_STRIKE, 0) + _
                        MAXIMUM_FUNC(S2_VAL - TEMP_STRIKE, 0)
  Next j
  
Next i

For i = 1 To NBINS_STRIKE
  TEMP_MATRIX(i, 2) = TEMP_MATRIX(i, 2) / nLOOPS
  TEMP_MATRIX(i, 3) = BLACK_SCHOLES_IMPLIED_VOLATILITY_FUNC(TEMP_MATRIX(i, 2), SPOT, TEMP_MATRIX(i, 1), _
        TENOR, RATE, RATE - 0, 1, MIN_SIGMA, MAX_SIGMA, CND_TYPE)
Next i

HESTON_EUROPEAN_MC_CALL_OPTION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
HESTON_EUROPEAN_MC_CALL_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HESTON_OPTION_PRICE_FUNC
'DESCRIPTION   : Heston Option Pricing Function
'LIBRARY       : DERIVATIVES
'GROUP         : HESTON
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Function HESTON_OPTION_PRICE_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal LONG_VAR As Double, _
ByVal CURR_VAR As Double, _
ByVal KAPPA As Double, _
ByVal LAMBDA As Double, _
ByVal RHO As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal OUTPUT As Integer = 0)

' OPTION_FLAG = 1 call, = -1 put

'LONG_VAR --> long run variance
'CURR_VAR --> current variance
'KAPPA --> mean reversion
'LAMBDA --> VolVol , LAMBDA
'RHO --> correlation

Dim ii As Long
Dim jj As Long       ' how many of subintervals
Dim kk As Long   ' counter

Dim PI_VAL As Double
Dim TEMP_SUM As Double       ' summing over subintervals
Dim TEMP_FACT As Double
Dim TEMP_MULT As Double      ' integration variable

Dim TEMP_ALPHA As Double     ' bounds to give +- 1 by change of variables
Dim TEMP_BETA As Double
Dim TEMP_INTEGRAT As Double  ' to sum up

Dim TEMP_DEGREE As Double ' degree for Gauss
Dim TEMP_LENGTH As Double ' overall length of integration
Dim TEMP_WIDTH As Double     ' length of subintervals
Dim TEMP_INTERVAL As Double
Dim TEMP_MONEYNESS As Double ' - log moneyness

Dim GAUSS_ARR As Variant 'pre-computed data for Gauss-Legendre integration
Dim ERROR_STR As String

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1
ERROR_STR = ""

PI_VAL = 3.14159265358979

TEMP_DEGREE = 16 ' degree for Gauss
TEMP_LENGTH = 128 ' overall length of integration
TEMP_INTERVAL = 8

If OUTPUT <> 0 Then 'Info about constants
    ERROR_STR = " Integral ( Function(k),k=0.." & TEMP_LENGTH & _
    " ) with " & Round(TEMP_LENGTH / TEMP_INTERVAL) & _
    " subintervals of length = " & TEMP_INTERVAL
    GoTo ERROR_LABEL
End If

ReDim GAUSS_ARR(1 To 16, 1 To 2)
TEMP_MONEYNESS = Log(Exp(RATE * TENOR) * SPOT / STRIKE)
  
'COLUMN 1 --> zeros:
'------------------------------------------------------------------------------------
'One thing would be FFT (but Excel has not a reasonable one and it would be
'a mess with exactness). So i thought about adaptive quadrature but after some
'hack around i decided to equal spaced slicing in subintervalls and to apply a
'fast and simple Gauss-Legendre method. Due to Excel's restricted excactness of
'~ 15 digits a choose of degree=16 works (the higher degree above is nonsense,
'it introduces error). Playing with exacter tools a cut off for integration at
'128 gives good results.
'------------------------------------------------------------------------------------

GAUSS_ARR(1, 1) = -0.98940093499165
GAUSS_ARR(2, 1) = -0.944575023073233
GAUSS_ARR(3, 1) = -0.865631202387832
GAUSS_ARR(4, 1) = -0.755404408355003
GAUSS_ARR(5, 1) = -0.617876244402644
GAUSS_ARR(6, 1) = -0.458016777657227
GAUSS_ARR(7, 1) = -0.281603550779259
GAUSS_ARR(8, 1) = -9.50125098376374E-02
GAUSS_ARR(9, 1) = 9.50125098376374E-02
GAUSS_ARR(10, 1) = 0.281603550779259
GAUSS_ARR(11, 1) = 0.458016777657227
GAUSS_ARR(12, 1) = 0.617876244402644
GAUSS_ARR(13, 1) = 0.755404408355003
GAUSS_ARR(14, 1) = 0.865631202387832
GAUSS_ARR(15, 1) = 0.944575023073233
GAUSS_ARR(16, 1) = 0.98940093499165

'COLUMN 2 -->  weights:
GAUSS_ARR(1, 2) = 2.71524594117541E-02
GAUSS_ARR(2, 2) = 6.22535239386479E-02
GAUSS_ARR(3, 2) = 9.51585116824928E-02
GAUSS_ARR(4, 2) = 0.124628971255534
GAUSS_ARR(5, 2) = 0.149595988816577
GAUSS_ARR(6, 2) = 0.169156519395003
GAUSS_ARR(7, 2) = 0.182603415044924
GAUSS_ARR(8, 2) = 0.189450610455069
GAUSS_ARR(9, 2) = 0.189450610455069
GAUSS_ARR(10, 2) = 0.182603415044924
GAUSS_ARR(11, 2) = 0.169156519395003
GAUSS_ARR(12, 2) = 0.149595988816577
GAUSS_ARR(13, 2) = 0.124628971255534
GAUSS_ARR(14, 2) = 9.51585116824928E-02
GAUSS_ARR(15, 2) = 6.22535239386479E-02
GAUSS_ARR(16, 2) = 2.71524594117541E-02
'------------------------------------------------------------------------------------

' default values: TEMP_LENGTH = 128 = 16 jj of length 8
' with gDegree=16 evaluations for each subinterval, ie: 256
TEMP_WIDTH = TEMP_INTERVAL
jj = Round(TEMP_LENGTH / TEMP_INTERVAL)

If TENOR < 1 Then               ' take some care for periodics
  If Abs(TEMP_MONEYNESS) <= 0.0001 Then
    TEMP_WIDTH = 8
  Else
    TEMP_WIDTH = Abs(PI_VAL / TEMP_MONEYNESS)
    If 512 < TEMP_WIDTH Then
    Else
      TEMP_WIDTH = 512
    End If
  End If
  Do While (TEMP_INTERVAL + 1) < TEMP_WIDTH
    TEMP_WIDTH = TEMP_WIDTH / 2
  Loop
  jj = Round(1.5 * TEMP_LENGTH / TEMP_WIDTH) + 1
End If


TEMP_SUM = 0
For ii = 1 To jj ' add over the partition
'-----------------------------------------------------------------------------
  ' Gauss Legendre integration,  Abramowitz Stegun, p.887, 25.4.30
    TEMP_BETA = ((TEMP_WIDTH * CDbl(ii)) + (TEMP_WIDTH * CDbl(ii - 1))) / 2
    TEMP_ALPHA = ((TEMP_WIDTH * CDbl(ii)) - (TEMP_WIDTH * CDbl(ii - 1))) / 2

    TEMP_INTEGRAT = 0
    For kk = 1 To TEMP_DEGREE
      TEMP_MULT = TEMP_ALPHA * GAUSS_ARR(kk, 1) + TEMP_BETA
      ' change of variables for integration bounds
      TEMP_INTEGRAT = TEMP_INTEGRAT + GAUSS_ARR(kk, 2) * HESTON_OPTION_PRICE_INTEGRAND_FUNC(TEMP_MONEYNESS, KAPPA, TENOR, LONG_VAR, CURR_VAR, LAMBDA, RHO, TEMP_MULT)
      ' TEMP_MONEYNESS is the fourier variable, not the integration variable
    Next kk
    TEMP_FACT = TEMP_INTEGRAT * TEMP_ALPHA ' change of variables for
    'integration bounds
    TEMP_SUM = TEMP_SUM + TEMP_FACT
'-----------------------------------------------------------------------------
Next ii

Select Case OPTION_FLAG
Case 1 'Call Price
    HESTON_OPTION_PRICE_FUNC = (0.5 * (Exp(RATE * TENOR) * SPOT - STRIKE) + STRIKE * TEMP_SUM / PI_VAL) * Exp(-RATE * TENOR)
Case Else ' Put Price - use P/C parity
    HESTON_OPTION_PRICE_FUNC = (0.5 * (Exp(RATE * TENOR) * SPOT - STRIKE) + STRIKE * TEMP_SUM / PI_VAL + (STRIKE * Exp(-RATE * TENOR) - SPOT * Exp(-DIVD * TENOR))) * Exp(-RATE * TENOR)
End Select

Exit Function
ERROR_LABEL:
HESTON_OPTION_PRICE_FUNC = ERROR_STR
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HESTON_OPTION_PRICE_INTEGRAND_FUNC
'DESCRIPTION   : Heston Integrand
'LIBRARY       : DERIVATIVES
'GROUP         : HESTON
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/31/2009
'************************************************************************************
'************************************************************************************

Private Function HESTON_OPTION_PRICE_INTEGRAND_FUNC(ByVal MONEYNESS As Double, _
ByVal KAPPA As Double, _
ByVal TENOR As Double, _
ByVal LONG_VAR As Double, _
ByVal CURR_VAR As Double, _
ByVal LAMBDA As Double, _
ByVal RHO As Double, _
ByVal Z_MULT As Double)

Dim t277 As Double
Dim t308 As Double
Dim t254 As Double
Dim t270 As Double
Dim t276 As Double
Dim t272 As Double
Dim t256 As Double
Dim t273 As Double
Dim t275 As Double
Dim t278 As Double
Dim t253 As Double
Dim t246 As Double
Dim t279 As Double
Dim t287 As Double
Dim t288 As Double
Dim t289 As Double
Dim t291 As Double
Dim t290 As Double
Dim t284 As Double
Dim t280 As Double
Dim t292 As Double
Dim t249 As Double
Dim t293 As Double
Dim t261 As Double
Dim t295 As Double
Dim t296 As Double
Dim t262 As Double
Dim t297 As Double
Dim t298 As Double
Dim t301 As Double
Dim t251 As Double
Dim t302 As Double
Dim t303 As Double
Dim t264 As Double
Dim t250 As Double
Dim t304 As Double
Dim t265 As Double
Dim t305 As Double
Dim t266 As Double
Dim t267 As Double
Dim t307 As Double
Dim t268 As Double
Dim t252 As Double

Dim t341 As Double
Dim t342 As Double
Dim t344 As Double
Dim t343 As Double
Dim t350 As Double
Dim t353 As Double
Dim t319 As Double
Dim t320 As Double
Dim t354 As Double
Dim t331 As Double
Dim t355 As Double
Dim t314 As Double
Dim t356 As Double
Dim t357 As Double
Dim t332 As Double
Dim t358 As Double
Dim t359 As Double
Dim t360 As Double
Dim t361 As Double
Dim t316 As Double
Dim t322 As Double
Dim t333 As Double
Dim t363 As Double
Dim t347 As Double
Dim t364 As Double
Dim t365 As Double
Dim t334 As Double
Dim t366 As Double
Dim t367 As Double
Dim t368 As Double
Dim t369 As Double
Dim t370 As Double
Dim t371 As Double
Dim t380 As Double
Dim t372 As Double
Dim t312 As Double
Dim t376 As Double
Dim t377 As Double
Dim t336 As Double
Dim t317 As Double
Dim t379 As Double
Dim t381 As Double
Dim t337 As Double
Dim t318 As Double
Dim t327 As Double
Dim t382 As Double
Dim t339 As Double
Dim t328 As Double
Dim t330 As Double

Dim A_VAL As Double
Dim B_VAL As Double

On Error GoTo ERROR_LABEL

t308 = RHO * LAMBDA
t278 = LAMBDA * LAMBDA 't278 = LAMBDA ^ 2
t284 = RHO * RHO 't284 = rho ^ 2
t280 = KAPPA * KAPPA 't280 = KAPPA ^ 2
t279 = Z_MULT * Z_MULT 't279 = Z_MULT ^ 2
t304 = t278 * t279
t292 = Sqr(t280 * t280 + (-4 * KAPPA * _
        t308 + (2 * t284 + 2) * t280 + t278 + _
        (1 + (-2 + t284) * t284) * t304) * t304)

t293 = -t304 - t280 + t284 * t304
t291 = Sqr(2 * t292 - 2 * t293) 't291 = (2 * t292 - 2 * t293) ^ (1 / 2)
t289 = -t291 / 2 't289 = -1 / 2 * t291
t272 = KAPPA + t289
t273 = KAPPA + t291 / 2 't273 = KAPPA + 1 / 2 * t291
t276 = Z_MULT * t308

If (Sgn(2 * KAPPA * t276 - t278 * Z_MULT) = 0) Then
  t275 = Sgn((t280 - t278 * t279 * t284 + t278 * t279))
Else
  t275 = Sgn(2 * KAPPA * t276 - t278 * Z_MULT)
End If
t275 = t275 * -1

' The Sgn function is used to determine in which half-plane
' ("left" or "right") the complex-valued expression or number
' x lies x = x1 + x2*i

t290 = Sqr(2 * t292 + 2 * t293) 't290 = (2 * t292 + 2 * t293) ^ (1 / 2)
t288 = t290 / 2 't288 = 1 / 2 * t290
t287 = t275 * t288
t267 = -t276 + t287
t264 = 1 / (t273 * t273 + t267 * t267) 't264 = 1 / (t273 ^ 2 + t267 ^ 2)
t266 = -t276 - t275 * t290 / 2 't266 = -t276 - 1 / 2 * t275 * t290
t295 = (-t266 * t273 + t272 * t267) * t264
t307 = (t272 * t273 + t266 * t267) * t264
t270 = Exp(TENOR * t289)
t303 = Sin(TENOR * t288) * t275
t298 = t270 * t303
t268 = Cos(TENOR * t287)
t301 = t270 * t268
t249 = -t295 * t298 - 1 + t307 * t301
t297 = t307 * t303
t252 = (t295 * t268 + t297) * t270
t305 = CURR_VAR / (t249 * t249 + t252 * t252)

t256 = 1 - t307
t253 = 1 / (t256 * t256 + t295 * t295) 't253 = 1 / (t256 ^ 2 + t295 ^ 2)
t302 = t295 * t253
t277 = 1 / (LAMBDA * LAMBDA) 't277 = 1 / (LAMBDA ^ 2)
t296 = KAPPA * LONG_VAR * t277
t265 = 1 - t301
t262 = (t272 * t265 - t266 * t298) * t277
t261 = (t266 * t265 + t272 * t298) * t277
t254 = t295 * t301
t251 = t254 + t270 * t297
t250 = t254 + t307 * t298
t246 = t249 * t256 * t253


A_VAL = 1 / Z_MULT * Exp((t272 * TENOR - Log((-t246 + t250 * t302) ^ 2 _
        + (t250 * t256 - t295 + (t307 * t268 - t303 * t295) * _
        t295 * t270) ^ 2 * t253 ^ 2)) * t296 + (-t262 * t249 + _
        t261 * t252) * t305) * Sin((t266 * TENOR - 2 * _
        ATAN2_FUNC((t251 * t256 + t249 * t295) * t253, -t246 + _
        t251 * t302)) * t296 + (-t261 * t249 - t262 * t252) * _
        t305 + MONEYNESS * Z_MULT) ' For real arguments
        'x, y the two-argument
        ' function arctan(y, x)
        ' computes the principal value of the argument of the complex number
        ' x+I*y, so -Pi < arctan(y, x) <= Pi.


t344 = KAPPA * KAPPA 't344 = KAPPA ^ 2
t350 = RHO * RHO 't350 = rho ^ 2
t353 = Z_MULT * Z_MULT 't353 = Z_MULT ^ 2
t347 = LAMBDA * LAMBDA 't347 = LAMBDA ^ 2
t342 = RHO * LAMBDA
t364 = KAPPA * t342

If (Sgn(-2 * t364 * Z_MULT - t347 * Z_MULT + 2 * t347 * t350 * Z_MULT) = 0) Then
  t341 = Sgn((-t344 + 2 * t364 - t347 * t353 - _
        t347 * t350 + t347 * t350 * t353))
Else
  t341 = Sgn(-2 * t364 * Z_MULT - t347 * Z_MULT + 2 * t347 * t350 * Z_MULT)
End If

' The Sgn function is used to determine in which half-plane
' ("left" or "right") the complex-valued expression or number
' x lies x = x1 + x2*i

t354 = t353 * t353 't354 = t353 ^ 2
t370 = -t354 - t353
t380 = t350 * t347
t379 = 2 * t353
t382 = -4 * t364
t360 = 2 * Sqr((-t370 + (2 * t370 + (t354 + 1 + _
        t379) * t350) * t350) * t347 * t347 + _
        (-4 - 4 * t353) * t380 * t364 + (t382 + _
        (t379 + (6 + t379) * t350) * t347 + t344) * _
        t344) 't360 = 2 * ((-t370 + (2 * t370 + (t354 _
        + 1 + t379) * t350) * t350) * t347 ^ 2 + _
        (-4 - 4 * t353) * t380 * t364 + (t382 + _
        (t379 + (6 + t379) * t350) * t347 + t344) * _
        t344) ^ (1 / 2)
t361 = t382 + 2 * t344 + t347 * t379 + (2 - 2 * t353) * t380
t357 = Sqr(t360 - t361) / 2 't357 = 1 / 2 * (t360 - t361) ^ (1 / 2)
t356 = t341 * t357
t368 = Z_MULT * t342
t332 = -t368 + t356
t333 = t368 + t356
t359 = Sqr(t360 + t361) 't359 = (t360 + t361) ^ (1 / 2)
t358 = -t359 / 2 't358 = -1 / 2 * t359
t369 = KAPPA - t342
t336 = t358 + t369
t337 = t359 / 2 + t369 't337 = 1 / 2 * t359 + t369
t330 = 1 / (t337 * t337 + t332 * t332) 't330 = 1 / (t337 ^ 2 + t332 ^ 2)
t363 = (t333 * t337 + t336 * t332) * t330
t381 = (t336 * t337 - t333 * t332) * t330
t355 = TENOR * t357
t372 = t341 * Sin(t355)
t339 = Exp(TENOR * t358)
t367 = t339 * t372
t334 = Cos(t341 * t355)
t371 = t334 * t339
t314 = t363 * t367 + 1 - t381 * t371
t366 = t381 * t372
t317 = (t363 * t334 + t366) * t339
t377 = CURR_VAR / (t314 * t314 + t317 * t317)
t322 = 1 - t381
t319 = 1 / (t322 * t322 + t363 * t363) 't319 = 1 / (t322 ^ 2 + t363 ^ 2)
t376 = t319 * t363
t343 = 1 / t347
t365 = KAPPA * LONG_VAR * t343
t331 = 1 - t371
t328 = (-t333 * t331 + t336 * t367) * t343
t327 = (t336 * t331 + t333 * t367) * t343
t320 = t363 * t371
t318 = t320 + t339 * t366
t316 = t320 + t381 * t367
t312 = t314 * t322 * t319


B_VAL = -1 / Z_MULT * Exp((t327 * t314 + t328 * t317) * t377 + _
        (t336 * TENOR - Log((t312 + t316 * t376) ^ 2 + (t316 * _
        t322 - t363 + (t381 * t334 - t372 * t363) * t363 * _
        t339) ^ 2 * t319 ^ 2)) * t365) * Sin(-(-t333 * _
        TENOR - 2 * ATAN2_FUNC((t318 * t322 - t314 * t363) * _
        t319, t312 + t318 * t376)) * t365 - (t328 * t314 - _
        t327 * t317) * t377 - MONEYNESS * Z_MULT) ' For real arguments
        'x, y the two-argument
        ' function arctan(y, x)
        ' computes the principal value of the argument of the complex number
        ' x+I*y, so -Pi < arctan(y, x) <= Pi.

HESTON_OPTION_PRICE_INTEGRAND_FUNC = Exp(MONEYNESS) * B_VAL - A_VAL

Exit Function
ERROR_LABEL:
HESTON_OPTION_PRICE_INTEGRAND_FUNC = Err.number
End Function
