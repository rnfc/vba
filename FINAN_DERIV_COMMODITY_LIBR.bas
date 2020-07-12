Attribute VB_Name = "FINAN_DERIV_COMMODITY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MILTERSEN_SCHWARTZ_COMMODITY_OPTION_FUNC
'DESCRIPTION   : The Miltersen and Schwartz (1997) commodity option model
'(Gaussian case)
'LIBRARY       : DERIVATIVES
'GROUP         : COMMODITY
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function MILTERSEN_SCHWARTZ_COMMODITY_OPTION_FUNC( _
ByVal ZERO_COUPON As Double, _
ByVal FUTURES As Double, _
ByVal STRIKE As Double, _
ByVal OPT_EXPIRATION As Double, _
ByVal FUT_EXPIRATION As Double, _
ByVal SIGMA_SPOT As Double, _
ByVal SIGMA_YIELD As Double, _
ByVal SIGMA_FORWARD As Double, _
ByVal RHO_SPOT_YIELD As Double, _
ByVal RHO_SPOT_FORWARD As Double, _
ByVal RHO_YIELD_FORWARD As Double, _
ByVal MEAN_REVERSION_YIELD As Double, _
ByVal MEAN_REVERSION_FORWARD As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

'ZERO_COUPON: Price of zero coupon bond
'FUTURES: FUTURES price
'STRIKE: STRIKE price
'OPT_EXPIRATION: Time to option maturity
'FUT_EXPIRATION: Time to future contract maturity
'SIGMA_SPOT: Volatility of the spot commodity price
'SIGMA_YIELD: Volatility of future convenience yield
'SIGMA_FORWARD: Volatility of the forward interest rate
'RHO_SPOT_YIELD: Correlation commodity price and convenience yield
'RHO_SPOT_FORWARD: Correlation commodity price and forward rate
'RHO_YIELD_FORWARD: Correlation convenience yield and forward rate
'MEAN_REVERSION_YIELD: Speed of mean reversion convenience yield
'MEAN_REVERSION_FORWARD: Speed of mean reversion forward rates


Dim D1_VAL As Double
Dim D2_VAL As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1

ATEMP_VAL = SIGMA_SPOT ^ 2 * OPT_EXPIRATION + 2 * SIGMA_SPOT * (SIGMA_FORWARD * _
            RHO_SPOT_FORWARD * 1 / MEAN_REVERSION_FORWARD * (OPT_EXPIRATION - 1 / _
            MEAN_REVERSION_FORWARD * Exp(-MEAN_REVERSION_FORWARD * FUT_EXPIRATION) * _
            (Exp(MEAN_REVERSION_FORWARD * OPT_EXPIRATION) - 1)) _
            - SIGMA_YIELD * RHO_SPOT_YIELD * 1 / MEAN_REVERSION_YIELD * (OPT_EXPIRATION - _
            1 / MEAN_REVERSION_YIELD * Exp(-MEAN_REVERSION_YIELD * FUT_EXPIRATION) * _
            (Exp(MEAN_REVERSION_YIELD * OPT_EXPIRATION) - 1))) + SIGMA_YIELD ^ 2 * 1 / _
            MEAN_REVERSION_YIELD ^ 2 * (OPT_EXPIRATION + 1 / (2 * _
            MEAN_REVERSION_YIELD) * Exp(-2 * MEAN_REVERSION_YIELD * FUT_EXPIRATION) * _
            (Exp(2 * MEAN_REVERSION_YIELD * OPT_EXPIRATION) - 1) - 2 * 1 / _
            MEAN_REVERSION_YIELD * Exp(-MEAN_REVERSION_YIELD * FUT_EXPIRATION) * _
            (Exp(MEAN_REVERSION_YIELD * OPT_EXPIRATION) - 1)) + SIGMA_FORWARD ^ 2 * 1 / _
            MEAN_REVERSION_FORWARD ^ 2 * (OPT_EXPIRATION + 1 / (2 * MEAN_REVERSION_FORWARD) _
            * Exp(-2 * MEAN_REVERSION_FORWARD * FUT_EXPIRATION) * (Exp(2 * _
            MEAN_REVERSION_FORWARD * OPT_EXPIRATION) - 1) - 2 * 1 / MEAN_REVERSION_FORWARD _
            * Exp(-MEAN_REVERSION_FORWARD * FUT_EXPIRATION) * (Exp(MEAN_REVERSION_FORWARD * _
            OPT_EXPIRATION) - 1)) - 2 * SIGMA_YIELD * SIGMA_FORWARD * RHO_YIELD_FORWARD * 1 _
            / MEAN_REVERSION_YIELD * 1 / MEAN_REVERSION_FORWARD * (OPT_EXPIRATION - 1 / _
            MEAN_REVERSION_YIELD * Exp(-MEAN_REVERSION_YIELD * FUT_EXPIRATION) * _
            (Exp(MEAN_REVERSION_YIELD * OPT_EXPIRATION) - 1) - 1 / MEAN_REVERSION_FORWARD * _
            Exp(-MEAN_REVERSION_FORWARD * FUT_EXPIRATION) * (Exp(MEAN_REVERSION_FORWARD * _
            OPT_EXPIRATION) - 1) + 1 / (MEAN_REVERSION_YIELD + MEAN_REVERSION_FORWARD) * _
            Exp(-(MEAN_REVERSION_YIELD + MEAN_REVERSION_FORWARD) * FUT_EXPIRATION) * _
            (Exp((MEAN_REVERSION_YIELD + MEAN_REVERSION_FORWARD) * OPT_EXPIRATION) - 1))
                
BTEMP_VAL = SIGMA_FORWARD * 1 / MEAN_REVERSION_FORWARD * (SIGMA_SPOT * RHO_SPOT_FORWARD _
            * (OPT_EXPIRATION - 1 / MEAN_REVERSION_FORWARD * (1 - Exp(-MEAN_REVERSION_FORWARD _
            * OPT_EXPIRATION))) + SIGMA_FORWARD * 1 / MEAN_REVERSION_FORWARD * _
            (OPT_EXPIRATION - 1 / MEAN_REVERSION_FORWARD * Exp(-MEAN_REVERSION_FORWARD _
            * FUT_EXPIRATION) * (Exp(MEAN_REVERSION_FORWARD * OPT_EXPIRATION) - 1) - 1 / _
            MEAN_REVERSION_FORWARD * (1 - Exp(-MEAN_REVERSION_FORWARD * OPT_EXPIRATION)) _
            + 1 / (2 * MEAN_REVERSION_FORWARD) * Exp(-MEAN_REVERSION_FORWARD * FUT_EXPIRATION) _
            * (Exp(MEAN_REVERSION_FORWARD * OPT_EXPIRATION) - Exp(-MEAN_REVERSION_FORWARD * _
            OPT_EXPIRATION))) - SIGMA_YIELD * RHO_YIELD_FORWARD * 1 / MEAN_REVERSION_YIELD * _
            (OPT_EXPIRATION - 1 / MEAN_REVERSION_YIELD * Exp(-MEAN_REVERSION_YIELD * _
            FUT_EXPIRATION) * (Exp(MEAN_REVERSION_YIELD * OPT_EXPIRATION) - 1) - 1 / _
            MEAN_REVERSION_FORWARD * (1 - Exp(-MEAN_REVERSION_FORWARD * _
            OPT_EXPIRATION)) + 1 / (MEAN_REVERSION_YIELD + MEAN_REVERSION_FORWARD) * _
            Exp(-MEAN_REVERSION_YIELD * FUT_EXPIRATION) * (Exp(MEAN_REVERSION_YIELD * _
            OPT_EXPIRATION) - Exp(-MEAN_REVERSION_FORWARD * OPT_EXPIRATION))))
                
ATEMP_VAL = Sqr(ATEMP_VAL)
               
D1_VAL = (Log(FUTURES / STRIKE) - BTEMP_VAL + ATEMP_VAL ^ 2 / 2) / ATEMP_VAL
D2_VAL = (Log(FUTURES / STRIKE) - BTEMP_VAL - ATEMP_VAL ^ 2 / 2) / ATEMP_VAL
                
Select Case OPTION_FLAG
Case 1 ', "CALL", "C"
    MILTERSEN_SCHWARTZ_COMMODITY_OPTION_FUNC = ZERO_COUPON * (FUTURES * Exp(-BTEMP_VAL) * _
    CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * CND_FUNC(D2_VAL, CND_TYPE))
Case Else '-1 ', "PUT", "P"
    MILTERSEN_SCHWARTZ_COMMODITY_OPTION_FUNC = ZERO_COUPON * (STRIKE * CND_FUNC(-D2_VAL, CND_TYPE) - _
    FUTURES * Exp(-BTEMP_VAL) * CND_FUNC(-D1_VAL, CND_TYPE))
End Select
                
Exit Function
ERROR_LABEL:
MILTERSEN_SCHWARTZ_COMMODITY_OPTION_FUNC = Err.number
End Function
