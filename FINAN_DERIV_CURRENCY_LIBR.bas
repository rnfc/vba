Attribute VB_Name = "FINAN_DERIV_CURRENCY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
'************************************************************************************
'************************************************************************************
'FUNCTION      : GARMAN_KOHLHAGEN_CURRENCY_OPTION_FUNC
'DESCRIPTION   : Garman and Kohlhagen (1983) Currency options
'LIBRARY       : DERIVATIVES
'GROUP         : CURRENCY
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function GARMAN_KOHLHAGEN_CURRENCY_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal HOME_RATE As Double, _
ByVal FOREIGN_RATE As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

'SPOT: Current price of underlying common stock
'STRIKE: Exercise Price
'HOME_RATE: Domestic risk-free rate of interest
'FOREIGN_RATE: Foreign risk-free rate of interest
'EXPIRATION: Time to expiration
'SIGMA: VOLATILITY
'FTd/f = S0d/f * EXP(rFd * T) * EXP(-rFf * T)

Dim D1_VAL As Double
Dim D2_VAL As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then OPTION_FLAG = -1

D1_VAL = (Log(SPOT / STRIKE) + (HOME_RATE - FOREIGN_RATE + SIGMA ^ 2 / 2) * EXPIRATION) / (SIGMA * Sqr(EXPIRATION))
D2_VAL = D1_VAL - SIGMA * Sqr(EXPIRATION)

Select Case OPTION_FLAG
Case 1 ', "CALL", "C"
    GARMAN_KOHLHAGEN_CURRENCY_OPTION_FUNC = SPOT * Exp(-FOREIGN_RATE * EXPIRATION) * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * Exp(-HOME_RATE * EXPIRATION) * CND_FUNC(D2_VAL, CND_TYPE)
Case Else '-1 ', "PUT", "P"
    GARMAN_KOHLHAGEN_CURRENCY_OPTION_FUNC = STRIKE * Exp(-HOME_RATE * EXPIRATION) * CND_FUNC(-D2_VAL, CND_TYPE) - SPOT * Exp(-FOREIGN_RATE * EXPIRATION) * CND_FUNC(-D1_VAL, CND_TYPE)
End Select

Exit Function
ERROR_LABEL:
GARMAN_KOHLHAGEN_CURRENCY_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : FX_DIGITAL_OPTION_PREMIUM_FUNC
'DESCRIPTION   : Digital option premium for fx option
'LIBRARY       : DERIVATIVES
'GROUP         : CURRENCY
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function FX_DIGITAL_OPTION_PREMIUM_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal HOME_RATE As Double, _
ByVal FOREIGN_RATE As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

Dim D1_VAL As Double
Dim D2_VAL As Double

On Error GoTo ERROR_LABEL

D1_VAL = (Log(SPOT / STRIKE) + (HOME_RATE - FOREIGN_RATE - 0.5 * SIGMA ^ 2) * EXPIRATION) / (SIGMA * Sqr(EXPIRATION))
D2_VAL = D1_VAL - SIGMA * Sqr(EXPIRATION)

Select Case OPTION_FLAG
Case 1 ', "Call", "c" 'FX Call Value
    FX_DIGITAL_OPTION_PREMIUM_FUNC = SPOT * Exp(-FOREIGN_RATE * EXPIRATION) * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * Exp(-HOME_RATE * EXPIRATION) * CND_FUNC(D2_VAL, CND_TYPE)
Case -1 ', "Put", "p" 'FX Option Value
    FX_DIGITAL_OPTION_PREMIUM_FUNC = STRIKE * Exp(-HOME_RATE * EXPIRATION) * CND_FUNC(-D2_VAL, CND_TYPE) - SPOT * Exp(-FOREIGN_RATE * EXPIRATION) * CND_FUNC(-D1_VAL, CND_TYPE)
Case Else 'digital option premium for fx option
    FX_DIGITAL_OPTION_PREMIUM_FUNC = Exp(-HOME_RATE * EXPIRATION) * CND_FUNC(D1_VAL, CND_TYPE)
End Select
    
Exit Function
ERROR_LABEL:
FX_DIGITAL_OPTION_PREMIUM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EQUITY_LINKED_FX_OPTION_FUNC
'DESCRIPTION   : Equity linked foreign exchange option
'Value in home currency
'LIBRARY       : DERIVATIVES
'GROUP         : CURRENCY
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function EQUITY_LINKED_FX_OPTION_FUNC(ByVal EXCHANGE As Double, _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal HOME_RATE As Double, _
ByVal FOREIGN_RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA_SPOT As Double, _
ByVal SIGMA_EXCHANGE As Double, _
ByVal RHO_VAL As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)


'EXCHANGE --> Exchange Rate
'SPOT --> Asset price
'STRIKE --> Strike price
'EXPIRATION --> Time to maturity
'HOME_RATE --> Domestic HOME_RATE
'FOREIGN_RATE --> Foreign HOME_RATE
'DIVD --> Dividend yield
'SIGMA_SPOT --> Volatility stock
'SIGMA_EXCHANGE --> Volatility currency
'RHO --> Correlation

Dim D1_VAL As Double
Dim D2_VAL As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then OPTION_FLAG = -1

D1_VAL = (Log(EXCHANGE / STRIKE) + (HOME_RATE - FOREIGN_RATE + RHO_VAL * SIGMA_SPOT * SIGMA_EXCHANGE + SIGMA_EXCHANGE ^ 2 / 2) * EXPIRATION) / (SIGMA_EXCHANGE * Sqr(EXPIRATION))
D2_VAL = D1_VAL - SIGMA_EXCHANGE * Sqr(EXPIRATION)

Select Case OPTION_FLAG
Case 1 ', "Call", "c"
    EQUITY_LINKED_FX_OPTION_FUNC = EXCHANGE * SPOT * Exp(-DIVD * EXPIRATION) * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * SPOT * Exp((FOREIGN_RATE - HOME_RATE - DIVD - RHO_VAL * SIGMA_SPOT * SIGMA_EXCHANGE) * EXPIRATION) * CND_FUNC(D2_VAL, CND_TYPE)
Case Else '-1 ', "Put", "p"
    EQUITY_LINKED_FX_OPTION_FUNC = STRIKE * SPOT * Exp((FOREIGN_RATE - HOME_RATE - DIVD - RHO_VAL * SIGMA_SPOT * SIGMA_EXCHANGE) * EXPIRATION) * CND_FUNC(-D2_VAL, CND_TYPE) - EXCHANGE * SPOT * Exp(-DIVD * EXPIRATION) * CND_FUNC(-D1_VAL, CND_TYPE)
End Select

Exit Function
ERROR_LABEL:
EQUITY_LINKED_FX_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FOREIGN_EQUITY_OPTION_STRUCK_FUNC
'DESCRIPTION   : Foreign equity option struck in home currency: Value in
'home currency
'LIBRARY       : DERIVATIVES
'GROUP         : CURRENCY
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function FOREIGN_EQUITY_OPTION_STRUCK_FUNC(ByVal EXCHANGE As Double, _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA_SPOT As Double, _
ByVal SIGMA_EXCHANGE As Double, _
ByVal RHO_VAL As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

'EXCHANGE --> EXCHANGE RATE
'SPOT --> Asset price
'STRIKE --> STRIKE price
'EXPIRATION --> Time to maturity
'RATE --> Domestic RATE
'DIVD --> Dividend yield
'SIGMA_SPOT --> Volatility stock
'SIGMA_EXCHANGE --> Volatility currency
'RHO --> Correlation (RATE)
    
Dim D1_VAL As Double
Dim D2_VAL As Double
Dim V_VAL As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then OPTION_FLAG = -1

V_VAL = Sqr(SIGMA_EXCHANGE ^ 2 + SIGMA_SPOT ^ 2 + 2 * _
RHO_VAL * SIGMA_EXCHANGE * SIGMA_SPOT)
D1_VAL = (Log(EXCHANGE * SPOT / STRIKE) + (RATE - DIVD + V_VAL ^ 2 / 2) * EXPIRATION) / (V_VAL * Sqr(EXPIRATION))
D2_VAL = D1_VAL - V_VAL * Sqr(EXPIRATION)

Select Case OPTION_FLAG
Case 1 ', "Call", "c"
    FOREIGN_EQUITY_OPTION_STRUCK_FUNC = EXCHANGE * SPOT * Exp(-DIVD * EXPIRATION) * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * Exp(-RATE * EXPIRATION) * CND_FUNC(D2_VAL, CND_TYPE)
Case Else '-1 ', "Put", "p"
    FOREIGN_EQUITY_OPTION_STRUCK_FUNC = STRIKE * Exp(-RATE * EXPIRATION) * CND_FUNC(-D2_VAL, CND_TYPE) - EXCHANGE * SPOT * Exp(-DIVD * EXPIRATION) * CND_FUNC(-D1_VAL, CND_TYPE)
End Select

Exit Function
ERROR_LABEL:
FOREIGN_EQUITY_OPTION_STRUCK_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : QUANTOS_OPTION_FUNC
'DESCRIPTION   : Fixed exchange rate foreign equity options-- Quantos
'Value in home currency. Calculates Quanto Option Price for an
'option on foreign stock
'LIBRARY       : DERIVATIVES
'GROUP         : CURRENCY
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function QUANTOS_OPTION_FUNC(ByVal FIXED_EXCHANGE As Double, _
ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal HOME_RATE As Double, _
ByVal FOREIGN_RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA_SPOT As Double, _
ByVal SIGMA_EXCHANGE As Double, _
ByVal RHO_VAL As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

'FIXED_EXCHANGE --> Fixed exchange Rate
'SPOT --> Asset price
'STRIKE --> Strike price
'EXPIRATION --> Time to maturity
'HOME_RATE --> Domestic HOME_RATE
'FOREIGN_RATE --> Foreign HOME_RATE
'DIVD --> Dividend yield
'SIGMA_SPOT --> Volatility stock
'SIGMA_EXCHANGE --> Volatility currency
'RHO --> Correlation

Dim D1_VAL As Double
Dim D2_VAL As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then OPTION_FLAG = -1

D1_VAL = (Log(SPOT / STRIKE) + (FOREIGN_RATE - DIVD - RHO_VAL * SIGMA_SPOT * SIGMA_EXCHANGE + SIGMA_SPOT ^ 2 / 2) * EXPIRATION) / (SIGMA_SPOT * Sqr(EXPIRATION))
D2_VAL = D1_VAL - SIGMA_SPOT * Sqr(EXPIRATION)

Select Case OPTION_FLAG
Case 1 ', "Call", "c"
    QUANTOS_OPTION_FUNC = FIXED_EXCHANGE * (SPOT * Exp((FOREIGN_RATE - HOME_RATE - DIVD - RHO_VAL * SIGMA_SPOT * SIGMA_EXCHANGE) * EXPIRATION) * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * Exp(-HOME_RATE * EXPIRATION) * CND_FUNC(D2_VAL, CND_TYPE))
Case Else '-1 ', "Put", "p"
    QUANTOS_OPTION_FUNC = FIXED_EXCHANGE * (STRIKE * Exp(-HOME_RATE * EXPIRATION) * CND_FUNC(-D2_VAL, CND_TYPE) - SPOT * Exp((FOREIGN_RATE - HOME_RATE - DIVD - RHO_VAL * SIGMA_SPOT * SIGMA_EXCHANGE) * EXPIRATION) * CND_FUNC(-D1_VAL, CND_TYPE))
End Select

Exit Function
ERROR_LABEL:
QUANTOS_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : TAKEOVER_FX_OPTION_FUNC
'DESCRIPTION   : Takeover foreign exchange options Value in home currency
'LIBRARY       : DERIVATIVES
'GROUP         : CURRENCY
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function TAKEOVER_FX_OPTION_FUNC(ByVal VALUE_FIRM As Double, _
ByVal QUANTITY As Double, _
ByVal EXCHANGE As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal HOME_RATE As Double, _
ByVal FOREIGN_RATE As Double, _
ByVal SIGMA_STOCK As Double, _
ByVal SIGMA_EXCHANGE As Double, _
ByVal RHO_VAL As Double, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)
    
'VALUE_FIRM --> Value of foreign firm
'QUANTITY --> Number of currency units
'EXCHANGE -- > EXCHANGE Rate
'STRIKE --> STRIKE price
'Expiration --> Time to maturity
'HOME_RATe --> Home risk-free Rate
'FOREIGN_RATE --> Foreign risk-free rate
'SIGMA_STOCK --> Stock price volatility
'SIGMA_EXCHANGE --> Exchange Rate volatility
'RHO --> Correlation coefficient

Dim A1_VAL As Double
Dim A2_VAL As Double

On Error GoTo ERROR_LABEL

A1_VAL = (Log(VALUE_FIRM / QUANTITY) + (FOREIGN_RATE - RHO_VAL * SIGMA_EXCHANGE * SIGMA_STOCK - SIGMA_STOCK ^ 2 / 2) * EXPIRATION) / (SIGMA_STOCK * Sqr(EXPIRATION))
A2_VAL = (Log(EXCHANGE / STRIKE) + (HOME_RATE - FOREIGN_RATE - SIGMA_EXCHANGE ^ 2 / 2) * EXPIRATION) / (SIGMA_EXCHANGE * Sqr(EXPIRATION))

TAKEOVER_FX_OPTION_FUNC = QUANTITY * (EXCHANGE * Exp(-FOREIGN_RATE * EXPIRATION) * CBND_FUNC(A2_VAL + SIGMA_EXCHANGE * Sqr(EXPIRATION), -A1_VAL - RHO_VAL * SIGMA_EXCHANGE * Sqr(EXPIRATION), -RHO_VAL, CND_TYPE, CBND_TYPE) - STRIKE * Exp(-HOME_RATE * EXPIRATION) * CBND_FUNC(-A1_VAL, A2_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE))
    
Exit Function
ERROR_LABEL:
TAKEOVER_FX_OPTION_FUNC = Err.number
End Function

'http://pages.stern.nyu.edu/~igiddy/options.htm
'http://www.m-x.ca/f_publications_en/currency_options.pdf
'http://www.anz.com/australia/support/general/PDS-WEB%20Foreign%20Currency%20Options.pdf
'http://www.cmegroup.com/trading/fx/
'http://www.msu.edu/~butler/Toolkit.xls
