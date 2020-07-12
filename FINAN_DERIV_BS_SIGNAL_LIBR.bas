Attribute VB_Name = "FINAN_DERIV_BS_SIGNAL_LIBR"

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : PUT_CALL_RATIO_EMA_SIGNAL_FUNC
'DESCRIPTION   : Put/Call Ratio Buy-Sell Signal
'LIBRARY       : DERIVATIVES
'GROUP         : BS_PC_RATIO
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function PUT_CALL_RATIO_EMA_SIGNAL_FUNC(ByRef DATE_RNG As Variant, _
ByRef STOCK_PRICE_RNG As Variant, _
ByRef CALL_VOLUME_RNG As Variant, _
ByRef PUT_VOLUME_RNG As Variant, _
Optional ByVal BUY_SIGNAL As Double = 0.8, _
Optional ByVal SELL_SIGNAL As Double = 0.5, _
Optional ByVal EMA_PERIODS As Double = 10)

Dim i As Long
Dim NROWS As Long

Dim EMA_MULTIPLIER As Double

Dim DATE_VECTOR As Variant
Dim STOCK_PRICE_VECTOR As Variant
Dim CALL_VOLUME_VECTOR As Variant
Dim PUT_VOLUME_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 1) = 1 Then
    DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
End If
NROWS = UBound(DATE_VECTOR, 1)

STOCK_PRICE_VECTOR = STOCK_PRICE_RNG
If UBound(STOCK_PRICE_VECTOR, 1) = 1 Then
    STOCK_PRICE_VECTOR = MATRIX_TRANSPOSE_FUNC(STOCK_PRICE_VECTOR)
End If
If UBound(STOCK_PRICE_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

CALL_VOLUME_VECTOR = CALL_VOLUME_RNG
If UBound(CALL_VOLUME_VECTOR, 1) = 1 Then
    CALL_VOLUME_VECTOR = MATRIX_TRANSPOSE_FUNC(CALL_VOLUME_VECTOR)
End If
If UBound(CALL_VOLUME_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

PUT_VOLUME_VECTOR = PUT_VOLUME_RNG
If UBound(PUT_VOLUME_VECTOR, 1) = 1 Then
    PUT_VOLUME_VECTOR = MATRIX_TRANSPOSE_FUNC(PUT_VOLUME_VECTOR)
End If
If UBound(PUT_VOLUME_VECTOR, 1) <> NROWS Then: GoTo ERROR_LABEL

EMA_MULTIPLIER = 1 - 2 / (EMA_PERIODS + 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 9)
TEMP_MATRIX(0, 1) = "Date"
TEMP_MATRIX(0, 2) = "Stock Price"
TEMP_MATRIX(0, 3) = "Equity Call Volume"
TEMP_MATRIX(0, 4) = "Equity Put Volume"
TEMP_MATRIX(0, 5) = "Equity Total Volume"
TEMP_MATRIX(0, 6) = "Equity P/C Ratio"
TEMP_MATRIX(0, 7) = "P/C Ratio: " & "EMA - " & CStr(Format(EMA_PERIODS, 0))
TEMP_MATRIX(0, 8) = "Sell@"
TEMP_MATRIX(0, 9) = "Buy@"

i = 1
TEMP_MATRIX(i, 1) = DATE_VECTOR(i, 1)
TEMP_MATRIX(i, 2) = STOCK_PRICE_VECTOR(i, 1)
TEMP_MATRIX(i, 3) = CALL_VOLUME_VECTOR(i, 1)
TEMP_MATRIX(i, 4) = PUT_VOLUME_VECTOR(i, 1)
TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 3) + TEMP_MATRIX(i, 4)
If TEMP_MATRIX(i, 4) <> 0 Then
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 4) / TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 6)
    TEMP_MATRIX(i, 8) = IIf(TEMP_MATRIX(i, 7) <= SELL_SIGNAL, TEMP_MATRIX(i, 2), 0)
    TEMP_MATRIX(i, 9) = IIf(TEMP_MATRIX(i, 7) >= BUY_SIGNAL, TEMP_MATRIX(i, 2), 0)
Else
    TEMP_MATRIX(i, 6) = 0
    TEMP_MATRIX(i, 7) = 0
    TEMP_MATRIX(i, 8) = 0
    TEMP_MATRIX(i, 9) = 0
End If

For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = DATE_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = STOCK_PRICE_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = CALL_VOLUME_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = PUT_VOLUME_VECTOR(i, 1)
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 3) + TEMP_MATRIX(i, 4)
    
    If TEMP_MATRIX(i, 4) <> 0 Then
        TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 4) / TEMP_MATRIX(i, 3)
        TEMP_MATRIX(i, 7) = EMA_MULTIPLIER * TEMP_MATRIX(i - 1, 7) + (1 - EMA_MULTIPLIER) * TEMP_MATRIX(i, 6)
        TEMP_MATRIX(i, 8) = IIf(TEMP_MATRIX(i, 7) <= SELL_SIGNAL, TEMP_MATRIX(i, 2), 0)
        TEMP_MATRIX(i, 9) = IIf(TEMP_MATRIX(i, 7) >= BUY_SIGNAL, TEMP_MATRIX(i, 2), 0)
    Else
        TEMP_MATRIX(i, 6) = 0
        TEMP_MATRIX(i, 7) = 0
        TEMP_MATRIX(i, 8) = 0
        TEMP_MATRIX(i, 9) = 0
    End If
Next i

PUT_CALL_RATIO_EMA_SIGNAL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PUT_CALL_RATIO_EMA_SIGNAL_FUNC = Err.number
End Function
