Attribute VB_Name = "FINAN_DERIV_BS_PRICING_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : BLACK_SCHOLES_OPTION_FUNC
'DESCRIPTION   : European option on a stock with cash dividends
'LIBRARY       : DERIVATIVES
'GROUP         : BS_VALUATION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function BLACK_SCHOLES_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal VOLATILITY As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

'--> You can adjust the Spot Price for Dividends
    
Dim D1_VAL As Double
Dim D2_VAL As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1

D1_VAL = (Log(SPOT / STRIKE) + (RATE + VOLATILITY ^ 2 / 2) * EXPIRATION) / (VOLATILITY * Sqr(EXPIRATION))

D2_VAL = D1_VAL - VOLATILITY * Sqr(EXPIRATION)

Select Case OPTION_FLAG
Case 1 ', "CALL", "C"
    BLACK_SCHOLES_OPTION_FUNC = _
            SPOT * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * _
            Exp(-RATE * EXPIRATION) * CND_FUNC(D2_VAL, CND_TYPE)
Case Else '-1 ', "PUT", "P"
    BLACK_SCHOLES_OPTION_FUNC = STRIKE * Exp(-RATE * EXPIRATION) * _
        CND_FUNC(-D2_VAL, CND_TYPE) - SPOT * CND_FUNC(-D1_VAL, CND_TYPE)
End Select
        
Exit Function
ERROR_LABEL:
BLACK_SCHOLES_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EUROPEAN_CALL_OPTION_FUNC
'DESCRIPTION   : European Call Option Value
'LIBRARY       : DERIVATIVES
'GROUP         : BS_VALUATION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function EUROPEAN_CALL_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal VOLATILITY As Double, _
Optional ByVal CND_TYPE As Integer = 0)
 
On Error GoTo ERROR_LABEL
 
EUROPEAN_CALL_OPTION_FUNC = _
    SPOT * Exp(-DIVD * EXPIRATION) * _
    CND_FUNC(EUROPEAN_D1_DENSITY_FUNC(SPOT, _
    STRIKE, EXPIRATION, RATE, DIVD, VOLATILITY), CND_TYPE) - _
    STRIKE * Exp(EXPIRATION * -RATE) * _
    CND_FUNC(EUROPEAN_D2_DENSITY_FUNC(SPOT, _
    STRIKE, EXPIRATION, RATE, _
    DIVD, VOLATILITY), CND_TYPE)

Exit Function
ERROR_LABEL:
EUROPEAN_CALL_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EUROPEAN_PUT_OPTION_FUNC
'DESCRIPTION   : European Put Option Value
'LIBRARY       : DERIVATIVES
'GROUP         : BS_VALUATION
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function EUROPEAN_PUT_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal VOLATILITY As Double, _
Optional ByVal CND_TYPE As Integer = 0)
 
On Error GoTo ERROR_LABEL
 
EUROPEAN_PUT_OPTION_FUNC = _
    STRIKE * Exp(-RATE * EXPIRATION) * _
    CND_FUNC(-EUROPEAN_D2_DENSITY_FUNC(SPOT, _
    STRIKE, EXPIRATION, RATE, DIVD, VOLATILITY), CND_TYPE) - SPOT _
    * Exp(-DIVD * EXPIRATION) * _
    CND_FUNC(-EUROPEAN_D1_DENSITY_FUNC(SPOT, _
    STRIKE, EXPIRATION, RATE, DIVD, VOLATILITY), CND_TYPE)

Exit Function
ERROR_LABEL:
EUROPEAN_PUT_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GENERALIZED_BLACK_SCHOLES_FUNC
'DESCRIPTION   : The generalized Black and Scholes formula
'LIBRARY       : DERIVATIVES
'GROUP         : BS_VALUATION
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function GENERALIZED_BLACK_SCHOLES_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal VOLATILITY As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

'CARRY = COST_OF_CARRY = (RISK_FREE - DIVD_YIELD)

Dim D1_VAL As Double
Dim D2_VAL As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1

D1_VAL = (Log(SPOT / STRIKE) + (CARRY_COST + VOLATILITY ^ 2 / 2) * EXPIRATION) / _
(VOLATILITY * Sqr(EXPIRATION))

D2_VAL = D1_VAL - VOLATILITY * Sqr(EXPIRATION)

Select Case OPTION_FLAG
Case 1 ', "CALL", "C"
    GENERALIZED_BLACK_SCHOLES_FUNC = _
            SPOT * Exp((CARRY_COST - RATE) * EXPIRATION) * _
            CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * _
            Exp(-RATE * EXPIRATION) * _
            CND_FUNC(D2_VAL, CND_TYPE)
Case Else '-1 ', "PUT", "P"
    GENERALIZED_BLACK_SCHOLES_FUNC = _
            STRIKE * Exp(-RATE * EXPIRATION) * _
            CND_FUNC(-D2_VAL, CND_TYPE) - SPOT * _
            Exp((CARRY_COST - RATE) * _
            EXPIRATION) * CND_FUNC(-D1_VAL, CND_TYPE)
End Select
    
    
Exit Function
ERROR_LABEL:
GENERALIZED_BLACK_SCHOLES_FUNC = Err.number
End Function


Function BLACK_SCHOLES_FAIR_PRICE_FUNC(ByVal CURRENT_PRICE As Double, _
ByVal STRIKE_PRICE As Double, _
ByVal RISK_FREE_RATE As Double, _
ByVal VOLATILITY As Double, _
ByVal TENOR As Double, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim D1_VAL As Double
Dim D2_VAL As Double
Dim PRICE_VAL As Double 'Expected Price in t years
Dim PROFIT_VAL As Double 'Expected Profit in t years

On Error GoTo ERROR_LABEL

D1_VAL = Log(CURRENT_PRICE / STRIKE_PRICE) + ((RISK_FREE_RATE + VOLATILITY ^ 2 / 2) _
        * TENOR) / VOLATILITY / Sqr(TENOR)
D2_VAL = D1_VAL - VOLATILITY * Sqr(TENOR)

D1_VAL = CND_FUNC(D1_VAL, CND_TYPE)
D2_VAL = CND_FUNC(D2_VAL, CND_TYPE)

PROFIT_VAL = CURRENT_PRICE * D1_VAL
PRICE_VAL = PROFIT_VAL + CURRENT_PRICE

Select Case OUTPUT
Case 0
    BLACK_SCHOLES_FAIR_PRICE_FUNC = PRICE_VAL * Exp(-RISK_FREE_RATE * TENOR)
    ' / (1 + RISK_FREE_RATE) ^ TENOR
Case Else
    BLACK_SCHOLES_FAIR_PRICE_FUNC = CURRENT_PRICE * D1_VAL - _
                                STRIKE_PRICE * Exp(-RISK_FREE_RATE * TENOR) * D2_VAL
End Select

Exit Function
ERROR_LABEL:
BLACK_SCHOLES_FAIR_PRICE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EUROPEAN_D1_DENSITY_FUNC
'DESCRIPTION   : BS D1_VAL probability function
'LIBRARY       : DERIVATIVES
'GROUP         :
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function EUROPEAN_D1_DENSITY_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal VOLATILITY As Double)

On Error GoTo ERROR_LABEL

EUROPEAN_D1_DENSITY_FUNC = _
(Log(SPOT / STRIKE) + (RATE - DIVD + VOLATILITY ^ 2 / 2) * EXPIRATION) / _
(VOLATILITY * Sqr(EXPIRATION))

Exit Function
ERROR_LABEL:
EUROPEAN_D1_DENSITY_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EUROPEAN_D2_DENSITY_FUNC
'DESCRIPTION   : BS D2_VAL probability function
'LIBRARY       : DERIVATIVES
'GROUP         :
'ID            : 00X
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function EUROPEAN_D2_DENSITY_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal VOLATILITY As Double)

On Error GoTo ERROR_LABEL

EUROPEAN_D2_DENSITY_FUNC = _
    EUROPEAN_D1_DENSITY_FUNC(SPOT, STRIKE, _
    EXPIRATION, RATE, DIVD, VOLATILITY) - VOLATILITY * Sqr(EXPIRATION)

Exit Function
ERROR_LABEL:
EUROPEAN_D2_DENSITY_FUNC = Err.number
End Function

'------------------------------------------------------------------------
'Put-Call Parity and Capital Structure
'------------------------------------------------------------------------

'The pioneers of option pricing were Fischer Black, Myron Scholes,
'and Robert Merton. In the early 1970s they showed that options can
'be used to characterize the capital structure of a company. Today
'this model is widely used by financial institutions to assess a
'company's credit risk.

'To illustrate the model, consider a company that has assets that are
'financed with zero-coupon bonds and equity. Suppose that the bonds mature
'in five years at which time a principal payment of k is required. The
'company pays no dividends. If the assets are worth more than k in five
'years, the equity holders choose to repay the bond holders. If the assets
'are worth less than k, the equity holders choose to declare bankruptcy and
'the bond holders end up owning the company.

'The value of the equity in five years is therefore max(At - k,0) where At
'is the value of the company's assets at that time. This shows that the
'equity holders have a five-year European call option on the assets of the
'company with a strike price of k. What about the bondholders? The bondholders
'have given the equity holders the right to sell the company's assets to them
'for k in five years. The bonds are therefore worth the present value of k minus
'the value of a five-year European put option on the assets with a strike price of k.

'To summarize, if c and p are the values of the call and put options,
'respectively then:

'Value of equity = c
'Value of debt = PV(k) - p

'Denote the value of the assets of the company today by Ao. The value of
'the assets must equal the total value of the instruments used to finance
'the assets. This means that it must equal the sum of the value of the equity
'and the value of the debt, so that:

'Ao = C + [PV(k) - p]

'Rearranging this equation, we have:

'C PV(k) = p + Ao

'This is the put-call parity result for call and put options on the assets
'of the company.

'------------------------------------------------------------------------
