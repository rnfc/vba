Attribute VB_Name = "FINAN_ASSET_SYSTEM_MACD_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function ASSET_TA_MACD_BUY_SIGNAL_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MACD_SLOW_PERIOD As Long = 5, _
Optional ByVal MACD_FAST_PERIOD As Long = 20, _
Optional ByVal MACD_TRIGGER_PERIOD As Long = 9, _
Optional ByVal HARD_STOP As Double = 0.03, _
Optional ByVal LONG_FLAG As Boolean = True)

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim PRICE_VAL As Double
Dim ALPHA1_VAL As Double
Dim ALPHA2_VAL As Double
Dim ALPHA3_VAL As Double

Dim IN_MARKET_FLAG As Boolean

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If LONG_FLAG = True Then k = 1 Else k = -1 'Long/Short
If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ALPHA1_VAL = 1 - 2 / (MACD_SLOW_PERIOD + 1)
ALPHA2_VAL = 1 - 2 / (MACD_FAST_PERIOD + 1)
ALPHA3_VAL = 1 - 2 / (MACD_TRIGGER_PERIOD + 1)
'-----------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 13)
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"
TEMP_MATRIX(0, 8) = "EMA: " & MACD_SLOW_PERIOD ' React Faster
TEMP_MATRIX(0, 9) = "EMA: " & MACD_FAST_PERIOD 'React Slower
TEMP_MATRIX(0, 10) = "MACD"
TEMP_MATRIX(0, 11) = "TRIGGER: " & MACD_TRIGGER_PERIOD
TEMP_MATRIX(0, 12) = IIf(LONG_FLAG = True, "LONG ", "SHORT ") & "SIGNAL"
TEMP_MATRIX(0, 13) = "STOP LOSS"

i = 1
For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000

TEMP_MATRIX(i, 8) = ALPHA1_VAL * TEMP_MATRIX(i, 7) + (1 - ALPHA1_VAL) * TEMP_MATRIX(i, 7)
TEMP_MATRIX(i, 9) = ALPHA2_VAL * TEMP_MATRIX(i, 7) + (1 - ALPHA2_VAL) * TEMP_MATRIX(i, 7)
TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 8) - TEMP_MATRIX(i, 9)
TEMP_MATRIX(i, 11) = ALPHA3_VAL * TEMP_MATRIX(i, 10) + (1 - ALPHA3_VAL) * TEMP_MATRIX(i, 10)

For i = 2 To NROWS
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 8) = ALPHA1_VAL * TEMP_MATRIX(i - 1, 8) + (1 - ALPHA1_VAL) * TEMP_MATRIX(i, 7)
    TEMP_MATRIX(i, 9) = ALPHA2_VAL * TEMP_MATRIX(i - 1, 9) + (1 - ALPHA2_VAL) * TEMP_MATRIX(i, 7)
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 8) - TEMP_MATRIX(i, 9)
    TEMP_MATRIX(i, 11) = ALPHA3_VAL * TEMP_MATRIX(i - 1, 11) + (1 - ALPHA3_VAL) * TEMP_MATRIX(i, 10)
    
    If TEMP_MATRIX(i - 1, 10) * k < TEMP_MATRIX(i - 1, 11) * k And TEMP_MATRIX(i, 10) * k > TEMP_MATRIX(i, 11) * k Then
        TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 7)
        PRICE_VAL = TEMP_MATRIX(i, 7)
        IN_MARKET_FLAG = True
    Else
        TEMP_MATRIX(i, 12) = ""
    End If
    
    If IN_MARKET_FLAG = True Then
        If (TEMP_MATRIX(i, 7) / PRICE_VAL - 1) * k < -HARD_STOP Then
            TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 7)
            IN_MARKET_FLAG = False
        End If
    Else
        TEMP_MATRIX(i, 13) = ""
    End If
Next i

ASSET_TA_MACD_BUY_SIGNAL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_MACD_BUY_SIGNAL_FUNC = Err.number
End Function

Function ASSETS_MACD_FAST_SHORT_SIGNAL_FUNC(ByVal TICKERS_RNG As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MACD_SLOW_PERIOD As Long = 5, _
Optional ByVal MACD_FAST_PERIOD As Long = 20, _
Optional ByVal MACD_TRIGGER_PERIOD As Long = 9)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim A0_VAL As Double
Dim A1_VAL As Double

Dim B0_VAL As Double
Dim B1_VAL As Double

Dim C0_VAL As Double
Dim C1_VAL As Double

Dim D0_VAL As Double
Dim D1_VAL As Double

Dim ALPHA1_VAL As Double
Dim ALPHA2_VAL As Double
Dim ALPHA3_VAL As Double

Dim LONG_FLAG As Boolean
Dim SHORT_FLAG As Boolean
Dim TICKER_STR As String
Dim DATA_VECTOR As Variant
Dim TICKERS_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKERS_RNG) Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If
NCOLUMNS = UBound(TICKERS_VECTOR, 1)

ALPHA1_VAL = 1 - 2 / (MACD_SLOW_PERIOD + 1)
ALPHA2_VAL = 1 - 2 / (MACD_FAST_PERIOD + 1)
ALPHA3_VAL = 1 - 2 / (MACD_TRIGGER_PERIOD + 1)

ReDim TEMP_MATRIX(0 To NCOLUMNS, 1 To 3)
TEMP_MATRIX(0, 1) = "TICKER"
TEMP_MATRIX(0, 2) = "LONG SIGNAL: MACD(" & MACD_SLOW_PERIOD & "," & MACD_FAST_PERIOD & "," & MACD_TRIGGER_PERIOD & ")"
TEMP_MATRIX(0, 3) = "SHORT SIGNAL: MACD(" & MACD_SLOW_PERIOD & "," & MACD_FAST_PERIOD & "," & MACD_TRIGGER_PERIOD & ")"

For j = 1 To NCOLUMNS
    TICKER_STR = TICKERS_VECTOR(j, 1)
    TEMP_MATRIX(j, 1) = TICKER_STR
    TEMP_MATRIX(j, 2) = ""
    DATA_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "A", False, False, True)
    If IsArray(DATA_VECTOR) = False Then: GoTo 1983
    NROWS = UBound(DATA_VECTOR, 1)

    i = 1
    A0_VAL = 0: A1_VAL = 0
    B0_VAL = 0: B1_VAL = 0
    C0_VAL = 0: C1_VAL = 0
    D0_VAL = 0: D1_VAL = 0
    
    A0_VAL = (ALPHA1_VAL * DATA_VECTOR(i, 1) + (1 - ALPHA1_VAL) * DATA_VECTOR(i, 1))
    B0_VAL = (ALPHA2_VAL * DATA_VECTOR(i, 1) + (1 - ALPHA2_VAL) * DATA_VECTOR(i, 1))
    C0_VAL = A0_VAL - B0_VAL
    D0_VAL = ALPHA3_VAL * C0_VAL + (1 - ALPHA3_VAL) * C0_VAL
    
    For i = 2 To NROWS
        A1_VAL = ALPHA1_VAL * A0_VAL + (1 - ALPHA1_VAL) * DATA_VECTOR(i, 1)
        B1_VAL = ALPHA2_VAL * B0_VAL + (1 - ALPHA2_VAL) * DATA_VECTOR(i, 1)
        C1_VAL = A1_VAL - B1_VAL
        D1_VAL = ALPHA3_VAL * D0_VAL + (1 - ALPHA3_VAL) * C1_VAL
        LONG_FLAG = False
        If C0_VAL < D0_VAL And C1_VAL > D1_VAL Then: LONG_FLAG = True
        SHORT_FLAG = False
        If C0_VAL > D0_VAL And C1_VAL < D1_VAL Then: SHORT_FLAG = True
        A0_VAL = A1_VAL: B0_VAL = B1_VAL
        C0_VAL = C1_VAL: D0_VAL = D1_VAL
    Next i
    TEMP_MATRIX(j, 2) = LONG_FLAG
    TEMP_MATRIX(j, 3) = SHORT_FLAG
1983:
Next j

ASSETS_MACD_FAST_SHORT_SIGNAL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_MACD_FAST_SHORT_SIGNAL_FUNC = Err.number
End Function
