Attribute VB_Name = "FINAN_DERIV_EXTENDIBLE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : EXTENDIBLE_OPTION_WRITER_FUNC
'DESCRIPTION   : Writer extendible options
'LIBRARY       : DERIVATIVES
'GROUP         : EXTENDIBLE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'**********************************************************************************
'**********************************************************************************

Function EXTENDIBLE_OPTION_WRITER_FUNC(ByVal SPOT As Double, _
ByVal INITIAL_STRIKE As Double, _
ByVal EXTENDED_STRIKE As Double, _
ByVal INITIAL_EXPIRATION As Double, _
ByVal EXTENDED_EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Double = 0)

Dim Z1_VAL As Double
Dim Z2_VAL As Double
Dim RHO_VAL As Double

On Error GoTo ERROR_LABEL

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1

RHO_VAL = Sqr(INITIAL_EXPIRATION / EXTENDED_EXPIRATION)

Z1_VAL = (Log(SPOT / EXTENDED_STRIKE) + (CARRY_COST + SIGMA ^ 2 / 2) * _
    EXTENDED_EXPIRATION) / (SIGMA * Sqr(EXTENDED_EXPIRATION))

Z2_VAL = (Log(SPOT / INITIAL_STRIKE) + (CARRY_COST + SIGMA ^ 2 / 2) * _
    INITIAL_EXPIRATION) / (SIGMA * Sqr(INITIAL_EXPIRATION))

'--------------------------------------------------------------------------------
 Select Case OPTION_FLAG
'--------------------------------------------------------------------------------
 Case 1 ', "CALL", "C"
    EXTENDIBLE_OPTION_WRITER_FUNC = GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, INITIAL_STRIKE, _
        INITIAL_EXPIRATION, RATE, CARRY_COST, SIGMA, 1, CND_TYPE) + _
        SPOT * Exp((CARRY_COST - RATE) * EXTENDED_EXPIRATION) * _
        CBND_FUNC(Z1_VAL, -Z2_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE) - EXTENDED_STRIKE * _
        Exp(-RATE * EXTENDED_EXPIRATION) * CBND_FUNC(Z1_VAL - Sqr(SIGMA ^ 2 * _
        EXTENDED_EXPIRATION), -Z2_VAL + Sqr(SIGMA ^ 2 * _
        INITIAL_EXPIRATION), -RHO_VAL, CND_TYPE, CBND_TYPE)
'--------------------------------------------------------------------------------
 Case Else '-1 ', "PUT", "P"
'--------------------------------------------------------------------------------
    EXTENDIBLE_OPTION_WRITER_FUNC = GENERALIZED_BLACK_SCHOLES_FUNC(SPOT, INITIAL_STRIKE, _
       INITIAL_EXPIRATION, RATE, CARRY_COST, SIGMA, -1, CND_TYPE) + _
        EXTENDED_STRIKE * Exp(-RATE * EXTENDED_EXPIRATION) * _
        CBND_FUNC(-Z1_VAL + Sqr(SIGMA ^ 2 * EXTENDED_EXPIRATION), _
        Z2_VAL - Sqr(SIGMA ^ 2 * INITIAL_EXPIRATION), -RHO_VAL, CND_TYPE, _
        CBND_TYPE) - SPOT * Exp((CARRY_COST - RATE) * _
        EXTENDED_EXPIRATION) * CBND_FUNC(-Z1_VAL, Z2_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)
'--------------------------------------------------------------------------------
 End Select
'--------------------------------------------------------------------------------
        
Exit Function
ERROR_LABEL:
EXTENDIBLE_OPTION_WRITER_FUNC = Err.number
End Function
