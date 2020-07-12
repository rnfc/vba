Attribute VB_Name = "FINAN_DERIV_JUMP_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : JUMP_DIFFUSION_OPTION_FUNC
'DESCRIPTION   : Merton's (1976) jump diffusion model
'LIBRARY       : DERIVATIVES
'GROUP         : JUMP
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'**********************************************************************************
'**********************************************************************************

Function JUMP_DIFFUSION_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal SIGMA As Double, _
ByVal LAMBDA As Double, _
ByVal GAMMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal nLOOPS As Long = 10, _
Optional ByVal CND_TYPE As Integer = 0)

'LAMBDA = Jumps per year
'GAMMA = Percent of total volatility

Dim i As Long
Dim TEMP_VAL As Double
Dim TEMP_SUM As Double
Dim TEMP_SIGMA As Double
Dim DELTA_VAL As Double

On Error GoTo ERROR_LABEL

DELTA_VAL = Sqr(GAMMA * SIGMA ^ 2 / LAMBDA)
TEMP_VAL = Sqr(SIGMA ^ 2 - LAMBDA * DELTA_VAL ^ 2)

TEMP_SUM = 0
For i = 0 To nLOOPS
    TEMP_SIGMA = Sqr(TEMP_VAL ^ 2 + DELTA_VAL ^ 2 * (i / EXPIRATION))
    TEMP_SUM = TEMP_SUM + Exp(-LAMBDA * EXPIRATION) * (LAMBDA * _
        EXPIRATION) ^ i / FACTORIAL_FUNC(i) * BLACK_SCHOLES_OPTION_FUNC(SPOT, _
            STRIKE, EXPIRATION, RATE, TEMP_SIGMA, OPTION_FLAG, CND_TYPE)
Next i

JUMP_DIFFUSION_OPTION_FUNC = TEMP_SUM
    
Exit Function
ERROR_LABEL:
JUMP_DIFFUSION_OPTION_FUNC = Err.number
End Function


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : FRENCH_OPTION_FUNC
'DESCRIPTION   : The Black-Scholes model adjusted for trading day
'volatility (FRENCH)
'LIBRARY       : DERIVATIVES
'GROUP         : JUMP
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'**********************************************************************************
'**********************************************************************************

Function FRENCH_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal CALENDAR_TENOR As Double, _
ByVal TRADING_TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

Dim D1_VAL As Double
Dim D2_VAL As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1
D1_VAL = (Log(SPOT / STRIKE) + CARRY_COST * CALENDAR_TENOR + SIGMA ^ 2 / 2 * _
TRADING_TENOR) / (SIGMA * Sqr(TRADING_TENOR))
D2_VAL = D1_VAL - SIGMA * Sqr(TRADING_TENOR)

Select Case OPTION_FLAG
    Case 1 ', "CALL", "C"
        FRENCH_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * _
            CALENDAR_TENOR) * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * _
            Exp(-RATE * CALENDAR_TENOR) * CND_FUNC(D2_VAL, CND_TYPE)
    Case Else '-1 ', "PUT", "P"
        FRENCH_OPTION_FUNC = STRIKE * Exp(-RATE * CALENDAR_TENOR) * _
            CND_FUNC(-D2_VAL, CND_TYPE) - SPOT * Exp((CARRY_COST - RATE) * _
            CALENDAR_TENOR) * CND_FUNC(-D1_VAL, CND_TYPE)
End Select
    
Exit Function
ERROR_LABEL:
FRENCH_OPTION_FUNC = Err.number
End Function
