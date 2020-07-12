Attribute VB_Name = "FINAN_DERIV_FORWARD_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : BLACK_FUTURE_FORWARD_OPTION_FUNC
'DESCRIPTION   : Black (1977) Options on futures/forwards
'LIBRARY       : DERIVATIVES
'GROUP         : FORWARD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function BLACK_FUTURE_FORWARD_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal EXPIRATION As Double, _
ByVal RATE As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

    Dim D1_VAL As Double
    Dim D2_VAL As Double
    
    On Error GoTo ERROR_LABEL
    
    If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1
    
    D1_VAL = (Log(SPOT / STRIKE) + (SIGMA ^ 2 / 2) * EXPIRATION) / _
        (SIGMA * Sqr(EXPIRATION))
    D2_VAL = D1_VAL - SIGMA * Sqr(EXPIRATION)
    
    Select Case OPTION_FLAG
        Case 1 ', "CALL", "C"
            BLACK_FUTURE_FORWARD_OPTION_FUNC = Exp(-RATE * EXPIRATION) * _
                (SPOT * CND_FUNC(D1_VAL, CND_TYPE) - STRIKE * _
                CND_FUNC(D2_VAL, CND_TYPE))
        Case Else '-1 ', "PUT", "P"
            BLACK_FUTURE_FORWARD_OPTION_FUNC = Exp(-RATE * EXPIRATION) * _
                (STRIKE * CND_FUNC(-D2_VAL, CND_TYPE) - SPOT * _
                CND_FUNC(-D1_VAL, CND_TYPE))
    End Select

Exit Function
ERROR_LABEL:
BLACK_FUTURE_FORWARD_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FORWARD_START_OPTION_FUNC
'DESCRIPTION   : Options on forwards start options
'LIBRARY       : DERIVATIVES
'GROUP         : FORWARD
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function FORWARD_START_OPTION_FUNC(ByVal SPOT As Double, _
ByVal ALPHA As Double, _
ByVal START_TENOR As Double, _
ByVal END_TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)
        
    On Error GoTo ERROR_LABEL
    
    FORWARD_START_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * START_TENOR) * _
        GENERALIZED_BLACK_SCHOLES_FUNC(1, ALPHA, END_TENOR - START_TENOR, RATE, _
        CARRY_COST, SIGMA, OPTION_FLAG, CND_TYPE)
        
Exit Function
ERROR_LABEL:
FORWARD_START_OPTION_FUNC = Err.number
End Function
