Attribute VB_Name = "FINAN_DERIV_CORRELATION_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : TWO_ASSET_CORRELATION_OPTION_FUNC
'DESCRIPTION   : Two asset correlation options; Exchange Options--> Digital
'Correlation
'LIBRARY       : DERIVATIVES
'GROUP         : CORRELATION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function TWO_ASSET_CORRELATION_OPTION_FUNC(ByVal SPOT_A As Double, _
ByVal SPOT_B As Double, _
ByVal STRIKE_A As Double, _
ByVal STRIKE_B As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST_A As Double, _
ByVal CARRY_COST_B As Double, _
ByVal SIGMA_A As Double, _
ByVal SIGMA_B As Double, _
ByVal RHO_VAL As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)

'RHO = Correlation(A,B)

Dim Y1_VAL As Double
Dim Y2_VAL As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1

Y1_VAL = (Log(SPOT_A / STRIKE_A) + (CARRY_COST_A - SIGMA_A ^ 2 / 2) _
* EXPIRATION) / (SIGMA_A * Sqr(EXPIRATION))

Y2_VAL = (Log(SPOT_B / STRIKE_B) + (CARRY_COST_B - SIGMA_B ^ 2 / 2) _
* EXPIRATION) / (SIGMA_B * Sqr(EXPIRATION))

Select Case OPTION_FLAG
Case 1 ', "CALL", "C"
    TWO_ASSET_CORRELATION_OPTION_FUNC = SPOT_B * Exp((CARRY_COST_B - RATE) * EXPIRATION) * _
    CBND_FUNC(Y2_VAL + SIGMA_B * Sqr(EXPIRATION), Y1_VAL + RHO_VAL * SIGMA_B * Sqr(EXPIRATION), _
    RHO_VAL, CND_TYPE, CBND_TYPE) - STRIKE_B * Exp(-RATE * EXPIRATION) * _
    CBND_FUNC(Y2_VAL, Y1_VAL, RHO_VAL, CND_TYPE, CBND_TYPE)
Case Else '-1 ', "PUT", "P"
    TWO_ASSET_CORRELATION_OPTION_FUNC = STRIKE_B * Exp(-RATE * EXPIRATION) * _
    CBND_FUNC(-Y2_VAL, -Y1_VAL, RHO_VAL, CND_TYPE, CBND_TYPE) - SPOT_B * _
    Exp((CARRY_COST_B - RATE) * EXPIRATION) _
    * CBND_FUNC(-Y2_VAL - SIGMA_B * Sqr(EXPIRATION), -Y1_VAL - RHO_VAL * SIGMA_B * _
    Sqr(EXPIRATION), RHO_VAL, CND_TYPE, CBND_TYPE)
End Select
    
Exit Function
ERROR_LABEL:
TWO_ASSET_CORRELATION_OPTION_FUNC = Err.number
End Function
