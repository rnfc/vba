Attribute VB_Name = "FINAN_DERIV_OPTION_OPTION_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'**********************************************************************************
'**********************************************************************************
'FUNCTION      : OPTION_ON_OPTION_FUNC
'DESCRIPTION   : Options on options
'LIBRARY       : DERIVATIVES
'GROUP         : OPTION_OPTION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'**********************************************************************************
'**********************************************************************************

Function OPTION_ON_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE_OPT As Double, _
ByVal STRIKE_OPT_OPT As Double, _
ByVal EXPIRATION_OPT_OPT As Double, _
ByVal EXPIRATION_OPT As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)

'-----------------------------------------------
'[OPTION_FLAG =   1]   Call-on-call, "cc"
'[OPTION_FLAG = -11]   Call-on-put   "cp"
'[OPTION_FLAG =  11]   Put-on-call   "pc"
'[OPTION_FLAG =  -1]   Put-on-put    "pp"
'-----------------------------------------------
Dim i As Integer
Dim j As Integer

Dim Y1_VAL As Double
Dim Y2_VAL As Double

Dim Z1_VAL As Double
Dim Z2_VAL As Double
    
Dim RHO_VAL As Double

Dim TEMP_VALUE As Double
Dim TEMP_STRIKE As Double
Dim TEMP_SPOT As Double
Dim TEMP_DELTA As Double

Dim tolerance As Double

On Error GoTo ERROR_LABEL

Select Case OPTION_FLAG
Case 1, 2 ', "cc", "pc"
    j = 1
Case Else '4, 3 ', "pp", "cp"
    j = -1
    i = -1
End Select

'-----------------First-Pass calculation of critical price options on options

TEMP_SPOT = STRIKE_OPT
TEMP_VALUE = GENERALIZED_BLACK_SCHOLES_FUNC(TEMP_SPOT, STRIKE_OPT, _
            EXPIRATION_OPT - EXPIRATION_OPT_OPT, RATE, _
            CARRY_COST, SIGMA, j, CND_TYPE)

TEMP_DELTA = Exp((CARRY_COST - RATE) * (EXPIRATION_OPT - _
    EXPIRATION_OPT_OPT)) * (CND_FUNC((Log(TEMP_SPOT / STRIKE_OPT) + _
    (CARRY_COST + SIGMA ^ 2 / 2) * (EXPIRATION_OPT - _
    EXPIRATION_OPT_OPT)) / (SIGMA * Sqr((EXPIRATION_OPT - _
    EXPIRATION_OPT_OPT))), CND_TYPE) + i)

tolerance = 0.000001

Do While Abs(TEMP_VALUE - STRIKE_OPT_OPT) > tolerance
    'Newton-Raphson algorithm
    
    TEMP_SPOT = TEMP_SPOT - (TEMP_VALUE - STRIKE_OPT_OPT) / TEMP_DELTA
    
    TEMP_VALUE = GENERALIZED_BLACK_SCHOLES_FUNC(TEMP_SPOT, STRIKE_OPT, _
    EXPIRATION_OPT - EXPIRATION_OPT_OPT, RATE, _
    CARRY_COST, SIGMA, j, CND_TYPE)
    
    TEMP_DELTA = Exp((CARRY_COST - RATE) * (EXPIRATION_OPT - _
        EXPIRATION_OPT_OPT)) * (CND_FUNC((Log(TEMP_SPOT / STRIKE_OPT) + _
        (CARRY_COST + SIGMA ^ 2 / 2) * (EXPIRATION_OPT - _
        EXPIRATION_OPT_OPT)) / (SIGMA * Sqr((EXPIRATION_OPT - _
        EXPIRATION_OPT_OPT))), CND_TYPE) + i)
Loop

TEMP_STRIKE = TEMP_SPOT

'---------------------------------------------------------------------------------

RHO_VAL = Sqr(EXPIRATION_OPT_OPT / EXPIRATION_OPT)

Y1_VAL = (Log(SPOT / TEMP_STRIKE) + (CARRY_COST + SIGMA ^ 2 / 2) * _
EXPIRATION_OPT_OPT) / (SIGMA * Sqr(EXPIRATION_OPT_OPT))

Y2_VAL = Y1_VAL - SIGMA * Sqr(EXPIRATION_OPT_OPT)

Z1_VAL = (Log(SPOT / STRIKE_OPT) + (CARRY_COST + SIGMA ^ 2 / 2) * _
EXPIRATION_OPT) / (SIGMA * Sqr(EXPIRATION_OPT))

Z2_VAL = Z1_VAL - SIGMA * Sqr(EXPIRATION_OPT)

Select Case OPTION_FLAG
Case 1 ', "cc"
    OPTION_ON_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * _
        EXPIRATION_OPT) * CBND_FUNC(Z1_VAL, Y1_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) _
        - STRIKE_OPT * Exp(-RATE * EXPIRATION_OPT) * _
        CBND_FUNC(Z2_VAL, Y2_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - _
        STRIKE_OPT_OPT * Exp(-RATE * EXPIRATION_OPT_OPT) _
        * CND_FUNC(Y2_VAL, CND_TYPE)

Case 2 ', "pc"
    OPTION_ON_OPTION_FUNC = STRIKE_OPT * Exp(-RATE * EXPIRATION_OPT) * _
        CBND_FUNC(Z2_VAL, -Y2_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE) - SPOT * _
        Exp((CARRY_COST - RATE) * _
        EXPIRATION_OPT) * CBND_FUNC(Z1_VAL, -Y1_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE) + _
        STRIKE_OPT_OPT * _
        Exp(-RATE * EXPIRATION_OPT_OPT) * CND_FUNC(-Y2_VAL, CND_TYPE)

Case 3 ', "cp"
    OPTION_ON_OPTION_FUNC = STRIKE_OPT * Exp(-RATE * EXPIRATION_OPT) * _
        CBND_FUNC(-Z2_VAL, -Y2_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - _
        SPOT * Exp((CARRY_COST - RATE) * _
        EXPIRATION_OPT) * CBND_FUNC(-Z1_VAL, -Y1_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - _
        STRIKE_OPT_OPT * _
        Exp(-RATE * EXPIRATION_OPT_OPT) * CND_FUNC(-Y2_VAL, CND_TYPE)

Case Else '4, "pp"
    OPTION_ON_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * _
        EXPIRATION_OPT) * CBND_FUNC(-Z1_VAL, Y1_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE) _
        - STRIKE_OPT * _
        Exp(-RATE * EXPIRATION_OPT) * _
        CBND_FUNC(-Z2_VAL, Y2_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE) + _
        Exp(-RATE * EXPIRATION_OPT_OPT) * _
        STRIKE_OPT_OPT * CND_FUNC(Y2_VAL, CND_TYPE)
End Select
    
Exit Function
ERROR_LABEL:
OPTION_ON_OPTION_FUNC = Err.number
End Function
