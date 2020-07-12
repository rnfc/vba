Attribute VB_Name = "FINAN_DERIV_CASH_NOTHING_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : CASH_NOTHING_OPTION_FUNC
'DESCRIPTION   : Cash-or-nothing options
'LIBRARY       : DERIVATIVES
'GROUP         : CASH-NOTHING
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function CASH_NOTHING_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal cash As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

    Dim D_VAL As Double
    
    On Error GoTo ERROR_LABEL
    
    If OPTION_FLAG <> 1 Then OPTION_FLAG = -1
    
    D_VAL = (Log(SPOT / STRIKE) + (CARRY_COST - SIGMA ^ 2 / 2) * TENOR) / _
    (SIGMA * Sqr(TENOR))

    Select Case OPTION_FLAG
        Case 1 ', "c", "call"
            CASH_NOTHING_OPTION_FUNC = cash * Exp(-RATE * TENOR) * CND_FUNC(D_VAL, CND_TYPE)
        Case Else '-1 ', "p", "put"
            CASH_NOTHING_OPTION_FUNC = cash * Exp(-RATE * TENOR) * CND_FUNC(-D_VAL, CND_TYPE)
    End Select

Exit Function
ERROR_LABEL:
CASH_NOTHING_OPTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : TWO_ASSET_CASH_NOTHING_OPTION_FUNC
'DESCRIPTION   : Two asset cash-or-nothing options
'LIBRARY       : DERIVATIVES
'GROUP         : CASH-NOTHING
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function TWO_ASSET_CASH_NOTHING_OPTION_FUNC(ByVal SPOT_A As Double, _
ByVal SPOT_B As Double, _
ByVal STRIKE_A As Double, _
ByVal STRIKE_B As Double, _
ByVal cash As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST_A As Double, _
ByVal CARRY_COST_B As Double, _
ByVal SIGMA_A As Double, _
ByVal SIGMA_B As Double, _
ByVal RHO_VAL As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0, _
Optional ByVal CBND_TYPE As Integer = 0)
    
    Dim D1_VAL As Double
    Dim D2_VAL As Double
                                   
'OPTION_FLAG: [1] Cash-or-nothing call
'OPTION_FLAG: [2] Cash-or-nothing put
'OPTION_FLAG: [3] Cash-or-nothing up-down
'OPTION_FLAG: [4] Cash-or-nothing down-up

    On Error GoTo ERROR_LABEL

    D1_VAL = (Log(SPOT_A / STRIKE_A) + (CARRY_COST_A - _
            SIGMA_A ^ 2 / 2) * TENOR) / (SIGMA_A * Sqr(TENOR))
    
    D2_VAL = (Log(SPOT_B / STRIKE_B) + (CARRY_COST_B - _
            SIGMA_B ^ 2 / 2) * TENOR) / (SIGMA_B * Sqr(TENOR))

    Select Case OPTION_FLAG
        Case 1 'Cash-or-nothing call
            TWO_ASSET_CASH_NOTHING_OPTION_FUNC = cash * Exp(-RATE * TENOR) * _
                                CBND_FUNC(D1_VAL, D2_VAL, RHO_VAL, CND_TYPE, CBND_TYPE)
        Case 2 'Cash-or-nothing put
            TWO_ASSET_CASH_NOTHING_OPTION_FUNC = cash * Exp(-RATE * TENOR) * _
                                CBND_FUNC(-D1_VAL, -D2_VAL, RHO_VAL, CND_TYPE, CBND_TYPE)
        Case 3 'Cash-or-nothing up-down
            TWO_ASSET_CASH_NOTHING_OPTION_FUNC = cash * Exp(-RATE * TENOR) * _
                                CBND_FUNC(D1_VAL, -D2_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)
        Case Else '4 'Cash-or-nothing down-up
            TWO_ASSET_CASH_NOTHING_OPTION_FUNC = cash * Exp(-RATE * TENOR) * _
                                CBND_FUNC(-D1_VAL, D2_VAL, -RHO_VAL, CND_TYPE, CBND_TYPE)
    End Select

Exit Function
ERROR_LABEL:
TWO_ASSET_CASH_NOTHING_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_NOTHING_OPTION_FUNC
'DESCRIPTION   : Asset-or-nothing options
'LIBRARY       : DERIVATIVES
'GROUP         : CASH-NOTHING
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function ASSET_NOTHING_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal CARRY_COST As Double, _
ByVal SIGMA As Double, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal CND_TYPE As Integer = 0)

    Dim D_VAL As Double

    On Error GoTo ERROR_LABEL
    
    If OPTION_FLAG <> 1 Then OPTION_FLAG = -1
    
    D_VAL = (Log(SPOT / STRIKE) + (CARRY_COST + SIGMA ^ 2 / 2) * _
        TENOR) / (SIGMA * Sqr(TENOR))
    
    Select Case OPTION_FLAG
        Case 1 ', "c", "call"
            ASSET_NOTHING_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * _
                                TENOR) * CND_FUNC(D_VAL, CND_TYPE)
        Case Else '-1 ', "p", "put"
            ASSET_NOTHING_OPTION_FUNC = SPOT * Exp((CARRY_COST - RATE) * _
                                TENOR) * CND_FUNC(-D_VAL, CND_TYPE)
    End Select

Exit Function
ERROR_LABEL:
ASSET_NOTHING_OPTION_FUNC = Err.number
End Function
