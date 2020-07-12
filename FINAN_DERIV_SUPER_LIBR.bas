Attribute VB_Name = "FINAN_DERIV_SUPER_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : SUPER_SHARES_OPTION_FUNC
'DESCRIPTION   : Super Shares Options
'LIBRARY       : DERIVATIVES
'GROUP         : SUPER
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'**********************************************************************************
'**********************************************************************************

Function SUPER_SHARES_OPTION_FUNC(ByVal SPOT As Double, _
ByVal LOWER_STRIKE As Double, _
ByVal UPPER_STRIKE As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal CND_TYPE As Integer = 0)
 
Dim D1_VAL As Double
Dim D2_VAL As Double
    
On Error GoTo ERROR_LABEL
    
D1_VAL = (Log(SPOT / LOWER_STRIKE) + (CARRY_COST + SIGMA ^ 2 / 2) * TENOR) / _
        (SIGMA * Sqr(TENOR))

D2_VAL = (Log(SPOT / UPPER_STRIKE) + (CARRY_COST + SIGMA ^ 2 / 2) * TENOR) / _
        (SIGMA * Sqr(TENOR))

SUPER_SHARES_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * TENOR) / LOWER_STRIKE * _
    (CND_FUNC(D1_VAL, CND_TYPE) - CND_FUNC(D2_VAL, CND_TYPE))
    
Exit Function
ERROR_LABEL:
SUPER_SHARES_OPTION_FUNC = Err.number
End Function



