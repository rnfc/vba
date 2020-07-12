Attribute VB_Name = "FINAN_DERIV_BS_PAYOFF_LIBR"

'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------

'Portfolio payoff at option expiration
'x-axis: Stock price at expiration ($s)
'y-axis: total P&L ($s)

Function PORT_PAYOFF_OPTION_EXPIRATION_FUNC(ByVal MIN_PRICE_VAL As Double, _
ByVal DELTA_PRICE_VAL As Double, _
ByVal NBINS As Long, _
ByRef EXERCISE_PRICE_RNG As Variant, _
ByRef CALL_PUT_FEE_STOCK_RNG As Variant, _
ByRef NUMBER_CALLS_PUTS_STOCK_RNG As Variant, _
Optional ByRef CALL_PUT_STOCK_PURCHASE_RNG As Variant = 0, _
Optional ByRef LONG_SHORT_RNG As Variant = 1)

'Min & Delta --> Stock Price at Expiration

Dim i As Long
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_VECTOR(1 To NBINS, 1 To 2)
i = 1
TEMP_VECTOR(i, 1) = MIN_PRICE_VAL
TEMP_VECTOR(i, 2) = OPTION_PAYOFF_TABLE_FUNC(EXERCISE_PRICE_RNG, CALL_PUT_FEE_STOCK_RNG, TEMP_VECTOR(i, 1), NUMBER_CALLS_PUTS_STOCK_RNG, CALL_PUT_STOCK_PURCHASE_RNG, LONG_SHORT_RNG, 1)
For i = 2 To NBINS
    TEMP_VECTOR(i, 1) = TEMP_VECTOR(i - 1, 1) + DELTA_PRICE_VAL
    TEMP_VECTOR(i, 2) = OPTION_PAYOFF_TABLE_FUNC(EXERCISE_PRICE_RNG, CALL_PUT_FEE_STOCK_RNG, TEMP_VECTOR(i, 1), NUMBER_CALLS_PUTS_STOCK_RNG, CALL_PUT_STOCK_PURCHASE_RNG, LONG_SHORT_RNG, 1)
Next i

PORT_PAYOFF_OPTION_EXPIRATION_FUNC = TEMP_VECTOR
Exit Function
ERROR_LABEL:
PORT_PAYOFF_OPTION_EXPIRATION_FUNC = Err.number
End Function

'------------------------------------------------------------------------------------------------------------------------------------
'Enter the positions in the portfolio; enter a reasonable range of prices for the stock price on expiration
'------------------------------------------------------------------------------------------------------------------------------------

Function OPTION_PAYOFF_TABLE_FUNC( _
ByRef EXERCISE_PRICE_RNG As Variant, _
ByRef CALL_PUT_FEE_STOCK_RNG As Variant, _
ByRef STOCK_PRICE_MATURITY_RNG As Variant, _
ByRef NUMBER_CALLS_PUTS_STOCK_RNG As Variant, _
Optional ByRef CALL_PUT_STOCK_PURCHASE_RNG As Variant = 0, _
Optional ByRef LONG_SHORT_RNG As Variant = 1, _
Optional ByVal OUTPUT As Integer = 0)

'EXERCISE_PRICE_RNG: Exercise Price -> Strike Prices
'CALL_PUT_FEE_STOCK_RNG: Call/put fee, or price if stock
'STOCK_PRICE_MATURITY_RNG: Stock Price at option Maturity
'NUMBER_CALLS_PUTS_STOCK_RNG: Number of calls/puts/stock
'CALL_PUT_STOCK_PURCHASE_RNG: Call/Put/Stock purchase --> Stock / Call or Put
'LONG_SHORT_RNG: long/short --> Long / Short

Dim i As Long
Dim k As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant

Dim EXERCISE_PRICE_VECTOR As Variant
Dim CALL_PUT_FEE_STOCK_VECTOR As Variant
Dim STOCK_PRICE_MATURITY_VECTOR As Variant
Dim NUMBER_CALLS_PUTS_STOCK_VECTOR As Variant
Dim CALL_PUT_STOCK_PURCHASE_VECTOR As Variant
Dim LONG_SHORT_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(EXERCISE_PRICE_RNG) = True Then
    EXERCISE_PRICE_VECTOR = EXERCISE_PRICE_RNG
    If UBound(EXERCISE_PRICE_VECTOR, 1) = 1 Then
        EXERCISE_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(EXERCISE_PRICE_VECTOR)
    End If
Else
    ReDim EXERCISE_PRICE_VECTOR(1 To 1, 1 To 1)
    EXERCISE_PRICE_VECTOR(1, 1) = EXERCISE_PRICE_RNG
End If
NROWS = UBound(EXERCISE_PRICE_VECTOR, 1)

If IsArray(CALL_PUT_FEE_STOCK_RNG) = True Then
    CALL_PUT_FEE_STOCK_VECTOR = CALL_PUT_FEE_STOCK_RNG
    If UBound(CALL_PUT_FEE_STOCK_VECTOR, 1) = 1 Then
        CALL_PUT_FEE_STOCK_VECTOR = MATRIX_TRANSPOSE_FUNC(CALL_PUT_FEE_STOCK_VECTOR)
    End If
Else
    ReDim CALL_PUT_FEE_STOCK_VECTOR(1 To 1, 1 To 1)
    CALL_PUT_FEE_STOCK_VECTOR(1, 1) = CALL_PUT_FEE_STOCK_RNG
End If
If NROWS <> UBound(CALL_PUT_FEE_STOCK_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(STOCK_PRICE_MATURITY_RNG) = True Then
    STOCK_PRICE_MATURITY_VECTOR = STOCK_PRICE_MATURITY_RNG
    If UBound(STOCK_PRICE_MATURITY_VECTOR, 1) = 1 Then
        STOCK_PRICE_MATURITY_VECTOR = MATRIX_TRANSPOSE_FUNC(STOCK_PRICE_MATURITY_VECTOR)
    End If
Else
    ReDim STOCK_PRICE_MATURITY_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        STOCK_PRICE_MATURITY_VECTOR(i, 1) = STOCK_PRICE_MATURITY_RNG
    Next i
End If
If NROWS <> UBound(STOCK_PRICE_MATURITY_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(NUMBER_CALLS_PUTS_STOCK_RNG) = True Then
    NUMBER_CALLS_PUTS_STOCK_VECTOR = NUMBER_CALLS_PUTS_STOCK_RNG
    If UBound(NUMBER_CALLS_PUTS_STOCK_VECTOR, 1) = 1 Then
        NUMBER_CALLS_PUTS_STOCK_VECTOR = MATRIX_TRANSPOSE_FUNC(NUMBER_CALLS_PUTS_STOCK_VECTOR)
    End If
Else
    ReDim NUMBER_CALLS_PUTS_STOCK_VECTOR(1 To 1, 1 To 1)
    NUMBER_CALLS_PUTS_STOCK_VECTOR(1, 1) = NUMBER_CALLS_PUTS_STOCK_RNG
End If
If NROWS <> UBound(NUMBER_CALLS_PUTS_STOCK_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(CALL_PUT_STOCK_PURCHASE_RNG) = True Then
    CALL_PUT_STOCK_PURCHASE_VECTOR = CALL_PUT_STOCK_PURCHASE_RNG
    If UBound(CALL_PUT_STOCK_PURCHASE_VECTOR, 1) = 1 Then
        CALL_PUT_STOCK_PURCHASE_VECTOR = MATRIX_TRANSPOSE_FUNC(CALL_PUT_STOCK_PURCHASE_VECTOR)
    End If
Else
    ReDim CALL_PUT_STOCK_PURCHASE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        CALL_PUT_STOCK_PURCHASE_VECTOR(i, 1) = CALL_PUT_STOCK_PURCHASE_RNG
    Next i
End If
If NROWS <> UBound(CALL_PUT_STOCK_PURCHASE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(LONG_SHORT_RNG) = True Then
    LONG_SHORT_VECTOR = LONG_SHORT_RNG
    If UBound(LONG_SHORT_VECTOR, 1) = 1 Then
        LONG_SHORT_VECTOR = MATRIX_TRANSPOSE_FUNC(LONG_SHORT_VECTOR)
    End If
Else
    ReDim LONG_SHORT_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        LONG_SHORT_VECTOR(i, 1) = LONG_SHORT_RNG
    Next i
End If
If NROWS <> UBound(LONG_SHORT_VECTOR, 1) Then: GoTo ERROR_LABEL


ReDim TEMP_MATRIX(0 To NROWS, 1 To 9)

TEMP_MATRIX(0, 1) = "CALL/PUT/UNDERLYING PURCHASE"
TEMP_MATRIX(0, 2) = "LONG/SHORT"
TEMP_MATRIX(0, 3) = "EXERCISE PRICE"
TEMP_MATRIX(0, 4) = "CALL/PUT FEE, OR PRICE IF UNDERLYING"
TEMP_MATRIX(0, 5) = "UNDERLYING PRICE AT OPTION MATURITY"
TEMP_MATRIX(0, 6) = "NUMBER OF CALLS/PUTS/UNDERLYING"

TEMP_MATRIX(0, 7) = "PAYOFF"
TEMP_MATRIX(0, 8) = "PROFIT AND LOSS"
TEMP_MATRIX(0, 9) = "P&L"

TEMP_SUM = 0
For i = 1 To NROWS
    k = 0
    Select Case LCase(LONG_SHORT_VECTOR(i, 1))
    Case "1", "long"
        TEMP_MATRIX(i, 2) = "Long"
        k = 1
    Case "-1", "short"
        TEMP_MATRIX(i, 2) = "Short"
        k = -1
    End Select
    If k = 0 Then: GoTo 1983
    
    TEMP_MATRIX(i, 3) = CDbl(EXERCISE_PRICE_VECTOR(i, 1))
    TEMP_MATRIX(i, 4) = CDbl(CALL_PUT_FEE_STOCK_VECTOR(i, 1))
    TEMP_MATRIX(i, 5) = CDbl(STOCK_PRICE_MATURITY_VECTOR(i, 1))
    TEMP_MATRIX(i, 6) = CDbl(NUMBER_CALLS_PUTS_STOCK_VECTOR(i, 1))

    TEMP_MATRIX(i, 7) = 0: TEMP_MATRIX(i, 8) = 0
    Select Case LCase(CALL_PUT_STOCK_PURCHASE_VECTOR(i, 1))
    Case "1", "call", "c"
        TEMP_MATRIX(i, 1) = "Call"
        TEMP_MATRIX(i, 7) = MAXIMUM_FUNC(TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 3), 0)
        TEMP_MATRIX(i, 8) = (MAXIMUM_FUNC(TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 3), 0)) - TEMP_MATRIX(i, 4)
    Case "-1", "put", "p"
        TEMP_MATRIX(i, 1) = "Put"
        TEMP_MATRIX(i, 7) = MAXIMUM_FUNC(TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 5), 0)
        TEMP_MATRIX(i, 8) = MAXIMUM_FUNC(TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 5), 0) - TEMP_MATRIX(i, 4)
    Case Else
        TEMP_MATRIX(i, 1) = "Underlying"
        TEMP_MATRIX(i, 7) = MAXIMUM_FUNC(TEMP_MATRIX(i, 5), 0)
        TEMP_MATRIX(i, 8) = MAXIMUM_FUNC(TEMP_MATRIX(i, 5), 0) - TEMP_MATRIX(i, 4)
    End Select

    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 7) * k
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 8) * k
    
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 8)
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 9)
1983:
Next i

Select Case OUTPUT
Case 0
    OPTION_PAYOFF_TABLE_FUNC = TEMP_MATRIX
Case Else
    OPTION_PAYOFF_TABLE_FUNC = TEMP_SUM
End Select

Exit Function
ERROR_LABEL:
OPTION_PAYOFF_TABLE_FUNC = Err.number
End Function
