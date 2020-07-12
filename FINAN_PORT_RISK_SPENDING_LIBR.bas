Attribute VB_Name = "FINAN_PORT_RISK_SPENDING_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_SUSTAINABLE_SPENDING_FUNC

'DESCRIPTION   : example routine illustrating investment return & performance in a
'"liability context" (probability of ruin).

'Reference: "A Sustainable Spending Rate without Simulation" by
'M. A. Milevsky & Ch. Robinson, FAJ Vol 61 No 6, 2005

'LIBRARY       : PORTFOLIO
'GROUP         : RISK_SPENDING
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_SUSTAINABLE_SPENDING_FUNC(ByVal ANNUAL_CONSUMPTION As Double, _
ByVal INITIAL_INVESTMENT As Double, _
ByVal EXPECTED_RETURN As Double, _
ByVal VOLATILITY As Double, _
ByVal IMPLIED_MORTALITY_RATE As Double)

'Mortality Approximation = Log(2)/MEDIAN_LIFE_SPAN

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To 10, 1 To 2)

TEMP_MATRIX(1, 1) = "FINITE LIFE"
TEMP_MATRIX(2, 1) = "ALPHA"
TEMP_MATRIX(3, 1) = "BETA"
TEMP_MATRIX(4, 1) = "SPV"
TEMP_MATRIX(5, 1) = "FINITE PROBABILITY RUIN"
TEMP_MATRIX(6, 1) = "INFINITE LIFE"
TEMP_MATRIX(7, 1) = "ALPHA"
TEMP_MATRIX(8, 1) = "BETA"
TEMP_MATRIX(9, 1) = "SPV"
TEMP_MATRIX(10, 1) = "INFINITE PROBABILITY RUIN"

TEMP_MATRIX(1, 2) = ""
TEMP_MATRIX(2, 2) = (2 * EXPECTED_RETURN + 4 * IMPLIED_MORTALITY_RATE) / _
(IMPLIED_MORTALITY_RATE + VOLATILITY ^ 2) - 1
TEMP_MATRIX(3, 2) = 0.5 * (IMPLIED_MORTALITY_RATE + VOLATILITY ^ 2)
TEMP_MATRIX(4, 2) = 1 / (IMPLIED_MORTALITY_RATE + _
            EXPECTED_RETURN - VOLATILITY ^ 2)
TEMP_MATRIX(5, 2) = _
GAMMA_DIST_FUNC(ANNUAL_CONSUMPTION / INITIAL_INVESTMENT, TEMP_MATRIX(2, 2), _
TEMP_MATRIX(3, 2), True, True)
TEMP_MATRIX(6, 2) = ""
TEMP_MATRIX(7, 2) = 2 * EXPECTED_RETURN / VOLATILITY ^ 2 - 1
TEMP_MATRIX(8, 2) = 0.5 * VOLATILITY ^ 2
TEMP_MATRIX(9, 2) = 1 / (EXPECTED_RETURN - VOLATILITY ^ 2)
TEMP_MATRIX(10, 2) = _
GAMMA_DIST_FUNC(ANNUAL_CONSUMPTION / INITIAL_INVESTMENT, TEMP_MATRIX(7, 2), _
TEMP_MATRIX(8, 2), True, True)


PORT_SUSTAINABLE_SPENDING_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_SUSTAINABLE_SPENDING_FUNC = Err.number
End Function
