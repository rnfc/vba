Attribute VB_Name = "FINAN_FI_BOND_RETURN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : BOND_HOLDING_PERIOD_RETURN_FUNC
'DESCRIPTION   : HOLDING_PERIOD_RETURN
'LIBRARY       : BOND
'GROUP         : RETURN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function BOND_HOLDING_PERIOD_RETURN_FUNC(ByVal CASH_RATE As Double, _
ByVal START_PRICE As Double, _
ByVal END_PRICE As Double, _
ByVal COUPON As Double, _
Optional ByVal PAR_VALUE As Double = 100, _
Optional ByVal OUTPUT As Integer = 0)

On Error GoTo ERROR_LABEL

'------------------------------------------------------------------------------
Select Case OUTPUT
Case 0 'Long Position
    BOND_HOLDING_PERIOD_RETURN_FUNC = (END_PRICE - START_PRICE) / START_PRICE + COUPON * PAR_VALUE / START_PRICE
Case Else 'Short Position
    BOND_HOLDING_PERIOD_RETURN_FUNC = (-END_PRICE + START_PRICE) / START_PRICE - COUPON * PAR_VALUE / START_PRICE + CASH_RATE
End Select

Exit Function
ERROR_LABEL:
BOND_HOLDING_PERIOD_RETURN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : BOND_YEAR_TRADE_BASIS_FUNC
'DESCRIPTION   : FI_BASIS_TRADE
'LIBRARY       : BOND
'GROUP         : RETURN
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function BOND_YEAR_TRADE_BASIS_FUNC(ByVal FAIR_PRICE As Double, _
ByVal ACTUAL_PRICE As Double, _
ByVal DURATION As Double, _
Optional ByVal EXPECTED_RETURN As Double, _
Optional ByVal FREQUENCY As Integer = 3)

'If FAIR > ACTUAL Then: Long Futures and Short Bonds

Dim PROFIT_VAL As Double
Dim PERIOD_RETURN As Double
Dim LEVERAGE_VAL As Double

On Error GoTo ERROR_LABEL

PROFIT_VAL = (DURATION * (FAIR_PRICE - ACTUAL_PRICE))
PERIOD_RETURN = (1 + EXPECTED_RETURN) ^ (1 / (FREQUENCY)) - 1
LEVERAGE_VAL = PERIOD_RETURN / PROFIT_VAL

BOND_YEAR_TRADE_BASIS_FUNC = Array(PROFIT_VAL, PERIOD_RETURN, LEVERAGE_VAL)

Exit Function
ERROR_LABEL:
BOND_YEAR_TRADE_BASIS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : YIELD_SWAP_TRADE_FUNC
'DESCRIPTION   : FI_ASSET_SWAP_TRADE
'LIBRARY       : BOND
'GROUP         : RETURN
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 10-06-2008
'************************************************************************************
'************************************************************************************

Function YIELD_SWAP_TRADE_FUNC(ByVal LONG_YIELD As Double, _
ByVal LONG_DURATION As Double, _
ByVal LONG_LENDING_MARGING As Double, _
ByVal SHORT_YIELD As Double, _
ByVal SHORT_DURATION As Double, _
ByVal SHORT_BORROW_MARGING As Double, _
ByVal EXPECTED_CARRY As Double, _
Optional ByVal FREQUENCY As Integer = 2)


Dim CARRY_VAL As Double
Dim LEVERAGE_VAL As Double 'Required
Dim PERIOD_RETURN As Double 'Holding Period REturn

On Error GoTo ERROR_LABEL

CARRY_VAL = LONG_YIELD + LONG_LENDING_MARGING - SHORT_YIELD - SHORT_BORROW_MARGING
LEVERAGE_VAL = EXPECTED_CARRY / CARRY_VAL
PERIOD_RETURN = EXPECTED_CARRY / FREQUENCY + LEVERAGE_VAL * (LONG_YIELD - SHORT_YIELD) / FREQUENCY * SHORT_DURATION

'If over a 6 month the manager expects the yield differential
'between the two bonds to halve, the holding period return is
'PERIOD_RETURN

YIELD_SWAP_TRADE_FUNC = Array(CARRY_VAL, LEVERAGE_VAL, PERIOD_RETURN)

Exit Function
ERROR_LABEL:
YIELD_SWAP_TRADE_FUNC = Err.number
End Function
