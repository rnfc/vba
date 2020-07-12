Attribute VB_Name = "FINAN_CURRENCIES_FORWARD_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : FORWARD_EXCHANGE_VALUATION_FUNC
'DESCRIPTION   : Forward/Futures Valuation on a currency
'LIBRARY       : CURRENCIES
'GROUP         : FORWARD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/05/2011
'************************************************************************************
'************************************************************************************

Function FORWARD_EXCHANGE_VALUATION_FUNC( _
ByVal BASE_NAME_FX_RNG As Variant, _
ByVal QUOTE_NAME_FX_RNG As Variant, _
ByVal SPOT_EXCHANGE_RATE_RNG As Variant, _
ByVal FORWARD_EXCHANGE_RATE_RNG As Variant, _
ByVal MATURITY_RNG As Variant, _
ByVal BASE_TBILL_RATE_RNG As Variant, _
ByVal QUOTE_TBILL_RATE_RNG As Variant)

'SPOT_EXCHANGE_RATE & FORWARD_EXCHANGE_RATE: Must be quoted as Base / Quote
'BASE_TBILL_RATE & QUOTE_TBILL_RATE: MUST BE CONTINUOUS RATE

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim HEADING_STR As String

Dim BASE_NAME_FX_VECTOR As Variant
Dim QUOTE_NAME_FX_VECTOR As Variant

Dim MATURITY_VECTOR As Variant
Dim BASE_TBILL_RATE_VECTOR As Variant
Dim QUOTE_TBILL_RATE_VECTOR As Variant
Dim SPOT_EXCHANGE_RATE_VECTOR As Variant
Dim FORWARD_EXCHANGE_RATE_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(BASE_NAME_FX_RNG) = True Then
    BASE_NAME_FX_VECTOR = BASE_NAME_FX_RNG
    If UBound(BASE_NAME_FX_VECTOR) = 1 Then
        BASE_NAME_FX_VECTOR = MATRIX_TRANSPOSE_FUNC(BASE_NAME_FX_VECTOR)
    End If
Else
    ReDim BASE_NAME_FX_VECTOR(1 To 1, 1 To 1)
    BASE_NAME_FX_VECTOR(1, 1) = BASE_NAME_FX_RNG
End If
NROWS = UBound(BASE_NAME_FX_VECTOR, 1)

If IsArray(QUOTE_NAME_FX_RNG) = True Then
    QUOTE_NAME_FX_VECTOR = QUOTE_NAME_FX_RNG
    If UBound(QUOTE_NAME_FX_VECTOR) = 1 Then
        QUOTE_NAME_FX_VECTOR = MATRIX_TRANSPOSE_FUNC(QUOTE_NAME_FX_VECTOR)
    End If
Else
    ReDim QUOTE_NAME_FX_VECTOR(1 To 1, 1 To 1)
    QUOTE_NAME_FX_VECTOR(1, 1) = QUOTE_NAME_FX_RNG
End If
If NROWS <> UBound(QUOTE_NAME_FX_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(SPOT_EXCHANGE_RATE_RNG) = True Then
    SPOT_EXCHANGE_RATE_VECTOR = SPOT_EXCHANGE_RATE_RNG
    If UBound(SPOT_EXCHANGE_RATE_VECTOR) = 1 Then
        SPOT_EXCHANGE_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(SPOT_EXCHANGE_RATE_VECTOR)
    End If
Else
    ReDim SPOT_EXCHANGE_RATE_VECTOR(1 To 1, 1 To 1)
    SPOT_EXCHANGE_RATE_VECTOR(1, 1) = SPOT_EXCHANGE_RATE_RNG
End If
If NROWS <> UBound(SPOT_EXCHANGE_RATE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(FORWARD_EXCHANGE_RATE_RNG) = True Then
    FORWARD_EXCHANGE_RATE_VECTOR = FORWARD_EXCHANGE_RATE_RNG
    If UBound(FORWARD_EXCHANGE_RATE_VECTOR) = 1 Then
        FORWARD_EXCHANGE_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(FORWARD_EXCHANGE_RATE_VECTOR)
    End If
Else
    ReDim FORWARD_EXCHANGE_RATE_VECTOR(1 To 1, 1 To 1)
    FORWARD_EXCHANGE_RATE_VECTOR(1, 1) = FORWARD_EXCHANGE_RATE_RNG
End If
If NROWS <> UBound(FORWARD_EXCHANGE_RATE_VECTOR) Then: GoTo ERROR_LABEL

If IsArray(MATURITY_RNG) = True Then
    MATURITY_VECTOR = MATURITY_RNG
    If UBound(MATURITY_VECTOR) = 1 Then
        MATURITY_VECTOR = MATRIX_TRANSPOSE_FUNC(MATURITY_VECTOR)
    End If
Else
    ReDim MATURITY_VECTOR(1 To 1, 1 To 1)
    MATURITY_VECTOR(1, 1) = MATURITY_RNG
End If
If NROWS <> UBound(MATURITY_VECTOR) Then: GoTo ERROR_LABEL

If IsArray(QUOTE_TBILL_RATE_RNG) = True Then
    QUOTE_TBILL_RATE_VECTOR = QUOTE_TBILL_RATE_RNG
    If UBound(QUOTE_TBILL_RATE_VECTOR) = 1 Then
        QUOTE_TBILL_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(QUOTE_TBILL_RATE_VECTOR)
    End If
Else
    ReDim QUOTE_TBILL_RATE_VECTOR(1 To 1, 1 To 1)
    QUOTE_TBILL_RATE_VECTOR(1, 1) = QUOTE_TBILL_RATE_RNG
End If
If NROWS <> UBound(QUOTE_TBILL_RATE_VECTOR) Then: GoTo ERROR_LABEL

If IsArray(BASE_TBILL_RATE_RNG) = True Then
    BASE_TBILL_RATE_VECTOR = BASE_TBILL_RATE_RNG
    If UBound(BASE_TBILL_RATE_VECTOR) = 1 Then
        BASE_TBILL_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(BASE_TBILL_RATE_VECTOR)
    End If
Else
    ReDim BASE_TBILL_RATE_VECTOR(1 To 1, 1 To 1)
    BASE_TBILL_RATE_VECTOR(1, 1) = BASE_TBILL_RATE_RNG
End If
If NROWS <> UBound(BASE_TBILL_RATE_VECTOR) Then: GoTo ERROR_LABEL

NCOLUMNS = 13
HEADING_STR = "BASE RATE, QUOTE RATE, MATURITY,SPOT EXCHANGE RATE,FORWARD EXCHANGE RATE,QUOTE TBILL RATE," & _
"BASE TBILL RATE,VALUE OF UNDERLYING ASSET,PV(EXERCISE PRICE),VALUE OF FORWARD CONTRACT," & _
"NEW FORWARD EXCH.RATE F1,F1-F0,PV (F1 - F0),"
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)

k = 1
For j = 1 To NCOLUMNS
    l = InStr(k, HEADING_STR, ",")
    TEMP_MATRIX(0, j) = Mid(HEADING_STR, k, l - k)
    k = l + 1
Next j

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = BASE_NAME_FX_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = QUOTE_NAME_FX_VECTOR(i, 1)
    
    TEMP_MATRIX(i, 3) = MATURITY_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = SPOT_EXCHANGE_RATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 5) = FORWARD_EXCHANGE_RATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 6) = QUOTE_TBILL_RATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 7) = BASE_TBILL_RATE_VECTOR(i, 1)
    
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 4) * Exp(-TEMP_MATRIX(i, 6) * TEMP_MATRIX(i, 3))
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 5) * Exp(-TEMP_MATRIX(i, 7) * TEMP_MATRIX(i, 3))
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 8) - TEMP_MATRIX(i, 9)
    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 4) * Exp((TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 6)) * TEMP_MATRIX(i, 3))
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 11) - TEMP_MATRIX(i, 5)
    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 12) * Exp(-TEMP_MATRIX(i, 7) * TEMP_MATRIX(i, 3))
Next i

FORWARD_EXCHANGE_VALUATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
FORWARD_EXCHANGE_VALUATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : FORWARD_EXCHANGE_ARBITRAGE_FUNC
'DESCRIPTION   : Forward/Futures Arbitrage on a currency
'LIBRARY       : CURRENCIES
'GROUP         : FORWARD
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/05/2011
'************************************************************************************
'************************************************************************************

Function FORWARD_EXCHANGE_ARBITRAGE_FUNC( _
ByVal BASE_NAME_FX_RNG As Variant, _
ByVal QUOTE_NAME_FX_RNG As Variant, _
ByVal SPOT_EXCHANGE_RATE_RNG As Variant, _
ByVal DELIVERY_PRICE_RNG As Variant, _
ByVal MATURITY_RNG As Variant, _
ByVal BASE_TBILL_RATE_RNG As Variant, _
ByVal QUOTE_TBILL_RATE_RNG As Variant, _
Optional ByVal FORWARD_EXCHANGE_RATE_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim HEADING_STR As String

Dim BASE_NAME_FX_VECTOR As Variant
Dim QUOTE_NAME_FX_VECTOR As Variant

Dim FORWARD_EXCHANGE_RATE_VECTOR As Variant
Dim DELIVERY_PRICE_VECTOR As Variant
Dim SPOT_EXCHANGE_RATE_VECTOR As Variant
Dim MATURITY_VECTOR As Variant
Dim QUOTE_TBILL_RATE_VECTOR As Variant
Dim BASE_TBILL_RATE_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(BASE_NAME_FX_RNG) = True Then
    BASE_NAME_FX_VECTOR = BASE_NAME_FX_RNG
    If UBound(BASE_NAME_FX_VECTOR) = 1 Then
        BASE_NAME_FX_VECTOR = MATRIX_TRANSPOSE_FUNC(BASE_NAME_FX_VECTOR)
    End If
Else
    ReDim BASE_NAME_FX_VECTOR(1 To 1, 1 To 1)
    BASE_NAME_FX_VECTOR(1, 1) = BASE_NAME_FX_RNG
End If
NROWS = UBound(BASE_NAME_FX_VECTOR, 1)

If IsArray(QUOTE_NAME_FX_RNG) = True Then
    QUOTE_NAME_FX_VECTOR = QUOTE_NAME_FX_RNG
    If UBound(QUOTE_NAME_FX_VECTOR) = 1 Then
        QUOTE_NAME_FX_VECTOR = MATRIX_TRANSPOSE_FUNC(QUOTE_NAME_FX_VECTOR)
    End If
Else
    ReDim QUOTE_NAME_FX_VECTOR(1 To 1, 1 To 1)
    QUOTE_NAME_FX_VECTOR(1, 1) = QUOTE_NAME_FX_RNG
End If
If NROWS <> UBound(QUOTE_NAME_FX_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(DELIVERY_PRICE_RNG) = True Then
    DELIVERY_PRICE_VECTOR = DELIVERY_PRICE_RNG
    If UBound(DELIVERY_PRICE_VECTOR) = 1 Then
        DELIVERY_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(DELIVERY_PRICE_VECTOR)
    End If
Else
    ReDim DELIVERY_PRICE_VECTOR(1 To 1, 1 To 1)
    DELIVERY_PRICE_VECTOR(1, 1) = DELIVERY_PRICE_RNG
End If
If NROWS <> UBound(DELIVERY_PRICE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(SPOT_EXCHANGE_RATE_RNG) = True Then
    SPOT_EXCHANGE_RATE_VECTOR = SPOT_EXCHANGE_RATE_RNG
    If UBound(SPOT_EXCHANGE_RATE_VECTOR) = 1 Then
        SPOT_EXCHANGE_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(SPOT_EXCHANGE_RATE_VECTOR)
    End If
Else
    ReDim SPOT_EXCHANGE_RATE_VECTOR(1 To 1, 1 To 1)
    SPOT_EXCHANGE_RATE_VECTOR(1, 1) = SPOT_EXCHANGE_RATE_RNG
End If
If NROWS <> UBound(SPOT_EXCHANGE_RATE_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(MATURITY_RNG) = True Then
    MATURITY_VECTOR = MATURITY_RNG
    If UBound(MATURITY_VECTOR) = 1 Then
        MATURITY_VECTOR = MATRIX_TRANSPOSE_FUNC(MATURITY_VECTOR)
    End If
Else
    ReDim MATURITY_VECTOR(1 To 1, 1 To 1)
    MATURITY_VECTOR(1, 1) = MATURITY_RNG
End If
If NROWS <> UBound(MATURITY_VECTOR) Then: GoTo ERROR_LABEL

If IsArray(QUOTE_TBILL_RATE_RNG) = True Then
    QUOTE_TBILL_RATE_VECTOR = QUOTE_TBILL_RATE_RNG
    If UBound(QUOTE_TBILL_RATE_VECTOR) = 1 Then
        QUOTE_TBILL_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(QUOTE_TBILL_RATE_VECTOR)
    End If
Else
    ReDim QUOTE_TBILL_RATE_VECTOR(1 To 1, 1 To 1)
    QUOTE_TBILL_RATE_VECTOR(1, 1) = QUOTE_TBILL_RATE_RNG
End If
If NROWS <> UBound(QUOTE_TBILL_RATE_VECTOR) Then: GoTo ERROR_LABEL

If IsArray(BASE_TBILL_RATE_RNG) = True Then
    BASE_TBILL_RATE_VECTOR = BASE_TBILL_RATE_RNG
    If UBound(BASE_TBILL_RATE_VECTOR) = 1 Then
        BASE_TBILL_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(BASE_TBILL_RATE_VECTOR)
    End If
Else
    ReDim BASE_TBILL_RATE_VECTOR(1 To 1, 1 To 1)
    BASE_TBILL_RATE_VECTOR(1, 1) = BASE_TBILL_RATE_RNG
End If
If NROWS <> UBound(BASE_TBILL_RATE_VECTOR) Then: GoTo ERROR_LABEL

If IsArray(FORWARD_EXCHANGE_RATE_RNG) = True Then
    FORWARD_EXCHANGE_RATE_VECTOR = FORWARD_EXCHANGE_RATE_RNG
    If UBound(FORWARD_EXCHANGE_RATE_VECTOR) = 1 Then
        FORWARD_EXCHANGE_RATE_VECTOR = MATRIX_TRANSPOSE_FUNC(FORWARD_EXCHANGE_RATE_VECTOR)
    End If
Else
    ReDim FORWARD_EXCHANGE_RATE_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        FORWARD_EXCHANGE_RATE_VECTOR(i, 1) = _
        FORWARD_EXCHANGE_RATE_FUNC(SPOT_EXCHANGE_RATE_VECTOR(i, 1), QUOTE_TBILL_RATE_VECTOR(i, 1), _
        BASE_TBILL_RATE_VECTOR(i, 1), MATURITY_VECTOR(i, 1))
    Next i
End If
If NROWS <> UBound(FORWARD_EXCHANGE_RATE_VECTOR) Then: GoTo ERROR_LABEL


NCOLUMNS = 20
HEADING_STR = "BASE RATE, QUOTE RATE, MATURITY,SPOT EXCHANGE RATE,QUOTE TBILL RATE,QUOTE TBILL DISCOUNT," & _
"BASE TBILL RATE,BASE TBILL DISCOUNT,DELIVERY PRICE,FORWARD RATE,TYPE OF ARBITRAGE," & _
"HOME TBILL: CC/RCC,HOME TBILL: NOW,HOME TBILL: MATURITY,BORROWING/LENDING BASE: CC/RCC," & _
"BORROWING/LENDING BASE: NOW,BORROWING/LENDING BASE: MATURITY,FORWARD: CC/RCC," & _
"FORWARD: MATURITY,TOTAL POSITION,"
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)

k = 1
For j = 1 To NCOLUMNS
    l = InStr(k, HEADING_STR, ",")
    TEMP_MATRIX(0, j) = Mid(HEADING_STR, k, l - k)
    k = l + 1
Next j
 
For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = BASE_NAME_FX_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = QUOTE_NAME_FX_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = MATURITY_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = SPOT_EXCHANGE_RATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 5) = QUOTE_TBILL_RATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 6) = 1 / (Exp(1) ^ (TEMP_MATRIX(i, 5) * TEMP_MATRIX(i, 3)))
    TEMP_MATRIX(i, 7) = BASE_TBILL_RATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 8) = 1 / (Exp(1) ^ (TEMP_MATRIX(i, 7) * TEMP_MATRIX(i, 3)))
    
    TEMP_MATRIX(i, 9) = DELIVERY_PRICE_VECTOR(i, 1)
    'In finance, a forward contract or simply a forward is a non-standardized contract between two parties to
    'buy or sell an asset at a specified future time at a price agreed today. This is in contrast to a spot
    'contract, which is an agreement to buy or sell an asset today. It costs nothing to enter a forward contract.
    'The party agreeing to buy the underlying asset in the future assumes a long position, and the party agreeing
    'to sell the asset in the future assumes a short position. The price agreed upon is called the delivery price,
    'which is equal to the forward price at the time the contract is entered into.
    
    TEMP_MATRIX(i, 10) = FORWARD_EXCHANGE_RATE_VECTOR(i, 1)

    If TEMP_MATRIX(i, 10) > TEMP_MATRIX(i, 9) Then
        TEMP_MATRIX(i, 11) = "Reverse cash and carry"
        TEMP_MATRIX(i, 12) = "SELL"
        TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 4) * TEMP_MATRIX(i, 6)
        TEMP_MATRIX(i, 14) = ("-S(T)")
        
        TEMP_MATRIX(i, 15) = ("INVEST")
        TEMP_MATRIX(i, 16) = -1 * TEMP_MATRIX(i, 13)
        TEMP_MATRIX(i, 17) = -TEMP_MATRIX(i, 16) / TEMP_MATRIX(i, 8)
        
        TEMP_MATRIX(i, 18) = "BUY"
        TEMP_MATRIX(i, 19) = ("S(T)") & " - " & Format(DELIVERY_PRICE_VECTOR(i, 1), "0.00")
        TEMP_MATRIX(i, 20) = 1 * (TEMP_MATRIX(i, 10) - TEMP_MATRIX(i, 9))
    Else
        TEMP_MATRIX(i, 11) = "Cash and carry"
        TEMP_MATRIX(i, 12) = "BUY"
        TEMP_MATRIX(i, 13) = -TEMP_MATRIX(i, 4) * TEMP_MATRIX(i, 6)
        TEMP_MATRIX(i, 14) = ("S(T)")
        TEMP_MATRIX(i, 15) = ("BORROW")
        TEMP_MATRIX(i, 16) = -1 * TEMP_MATRIX(i, 13)
        TEMP_MATRIX(i, 17) = -TEMP_MATRIX(i, 16) / TEMP_MATRIX(i, 8)
        TEMP_MATRIX(i, 18) = "SELL"
        TEMP_MATRIX(i, 19) = Format(DELIVERY_PRICE_VECTOR(i, 1), "0.00") & " -S(T)"
        TEMP_MATRIX(i, 20) = -1 * (TEMP_MATRIX(i, 10) - TEMP_MATRIX(i, 9))
    End If
Next i

FORWARD_EXCHANGE_ARBITRAGE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
FORWARD_EXCHANGE_ARBITRAGE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : FORWARD_EXCHANGE_RATE_FUNC
'DESCRIPTION   : Forward/Futures contract on a currency
'LIBRARY       : CURRENCIES
'GROUP         : FORWARD
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/05/2011
'************************************************************************************
'************************************************************************************

Function FORWARD_EXCHANGE_RATE_FUNC( _
ByVal SPOT_EXCHANGE_RATE As Double, _
ByVal QUOTE_TBILL_RATE As Double, _
ByVal BASE_TBILL_RATE As Double, _
ByVal MATURITY As Double)

'SPOT_EXCHANGE_RATE Must be Quoted as BASE / HOME
'TBILL_RATES MUST BE CONTINUOUS

On Error GoTo ERROR_LABEL

FORWARD_EXCHANGE_RATE_FUNC = SPOT_EXCHANGE_RATE * Exp((BASE_TBILL_RATE - _
QUOTE_TBILL_RATE) * MATURITY)

Exit Function
ERROR_LABEL:
FORWARD_EXCHANGE_RATE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : EXPECTED_EXCHANGE_RATE_FUNC
'DESCRIPTION   : Expected Exchange Rate
'LIBRARY       : CURRENCIES
'GROUP         : FORWARD
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/05/2011
'************************************************************************************
'************************************************************************************

Function EXPECTED_EXCHANGE_RATE_FUNC( _
ByVal SPOT_RATE As Double, _
ByVal EXP_INFLATION_RATE_HOME As Double, _
ByVal EXP_INFLATION_RATE_BASE As Double, _
ByVal MATURITY As Double)

On Error GoTo ERROR_LABEL

'SPOT_RATE = Number of units of BASE currency
'per unit of domestic currency

'MATURITY = Year

EXPECTED_EXCHANGE_RATE_FUNC = SPOT_RATE * (1 + EXP_INFLATION_RATE_BASE) ^ _
MATURITY / (1 + EXP_INFLATION_RATE_HOME) ^ MATURITY

Exit Function
ERROR_LABEL:
EXPECTED_EXCHANGE_RATE_FUNC = Err.number
End Function
