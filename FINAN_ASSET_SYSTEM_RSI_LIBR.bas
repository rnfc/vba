Attribute VB_Name = "FINAN_ASSET_SYSTEM_RSI_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function RSI_TARGET_PRICES_FUNC(ByRef TICKERS_RNG As Variant, _
Optional ByVal LOW_TRIGGER_RNG As Variant = 20, _
Optional ByVal HIGH_TRIGGER_RNG As Variant = 80, _
Optional ByVal RSI_FACTOR_RNG As Variant = 2, _
Optional ByVal DAYS_WINDOWS_VAL As Long = 34, _
Optional ByVal HEADER_FLAG As Boolean = True)

'-----------------------------------------------------------------------------------------------------------*
' Function to return RSI indicator buy and sell target prices
'-----------------------------------------------------------------------------------------------------------*
' Samples of use = RSI_TARGET_PRICES_FUNC("MMM",20,80)
'-----------------------------------------------------------------------------------------------------------*

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TICKER_STR As String
Dim HEADINGS_STR As String

Dim RSI_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant
Dim LOW_TRIGGER_VECTOR As Variant
Dim HIGH_TRIGGER_VECTOR As Variant
Dim RSI_FACTOR_VECTOR As Variant

Dim QUOTES_MATRIX As Variant
Dim HISTORICAL_MATRIX As Variant

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
NROWS = UBound(TICKERS_VECTOR, 1): NCOLUMNS = 13

If IsArray(LOW_TRIGGER_RNG) Then
    LOW_TRIGGER_VECTOR = LOW_TRIGGER_RNG
    If UBound(LOW_TRIGGER_VECTOR, 1) = 1 Then
        LOW_TRIGGER_VECTOR = MATRIX_TRANSPOSE_FUNC(LOW_TRIGGER_VECTOR)
    End If
Else
    ReDim LOW_TRIGGER_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS: LOW_TRIGGER_VECTOR(i, 1) = LOW_TRIGGER_RNG: Next i
End If
If NROWS <> UBound(LOW_TRIGGER_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(HIGH_TRIGGER_RNG) Then
    HIGH_TRIGGER_VECTOR = HIGH_TRIGGER_RNG
    If UBound(HIGH_TRIGGER_VECTOR, 1) = 1 Then
        HIGH_TRIGGER_VECTOR = MATRIX_TRANSPOSE_FUNC(HIGH_TRIGGER_VECTOR)
    End If
Else
    ReDim HIGH_TRIGGER_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS: HIGH_TRIGGER_VECTOR(i, 1) = HIGH_TRIGGER_RNG: Next i
End If
If NROWS <> UBound(HIGH_TRIGGER_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(RSI_FACTOR_RNG) Then
    RSI_FACTOR_VECTOR = RSI_FACTOR_RNG
    If UBound(RSI_FACTOR_VECTOR, 1) = 1 Then
        RSI_FACTOR_VECTOR = MATRIX_TRANSPOSE_FUNC(RSI_FACTOR_VECTOR)
    End If
Else
    ReDim RSI_FACTOR_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS: RSI_FACTOR_VECTOR(i, 1) = RSI_FACTOR_RNG: Next i
End If
If NROWS <> UBound(RSI_FACTOR_VECTOR, 1) Then: GoTo ERROR_LABEL

If HEADER_FLAG = True Then
    ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
    HEADINGS_STR = "Symbol,Current RSI,Buy Target Price,Sell Target Price,Last Traded Price,Bid Price,Ask Price,Open Price," & _
    "Low Price,High Price,Volume,Previous Close,Previous RSI,"
    i = 1
    For k = 1 To NCOLUMNS
        j = InStr(i, HEADINGS_STR, ",")
        TEMP_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
        i = j + 1
    Next k
Else
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
End If
ReDim QUOTES_MATRIX(1 To 7, 1 To 1)
QUOTES_MATRIX(1, 1) = "last trade": QUOTES_MATRIX(2, 1) = "bid"
QUOTES_MATRIX(3, 1) = "ask": QUOTES_MATRIX(4, 1) = "open"
QUOTES_MATRIX(5, 1) = "low": QUOTES_MATRIX(6, 1) = "high"
QUOTES_MATRIX(7, 1) = "volume"

QUOTES_MATRIX = YAHOO_QUOTES_FUNC(TICKERS_VECTOR, QUOTES_MATRIX, 0, False, "")
HISTORICAL_MATRIX = YAHOO_HISTORICAL_DATA_SERIES1_FUNC(TICKERS_VECTOR, DateSerial(Year(Date), Month(Date), Day(Date) - DAYS_WINDOWS_VAL), Date, 6, "d", False, True)
j = UBound(HISTORICAL_MATRIX, 1)

For i = 1 To NROWS
    RSI_MATRIX = ASSET_TA_RSI_FUNC(HISTORICAL_MATRIX, RSI_FACTOR_VECTOR(i, 1), i + 1)
    k = UBound(RSI_MATRIX, 1)
    TEMP_MATRIX(i, 1) = TICKERS_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = 100 - 100 / (1 + (RSI_MATRIX(k, 5) + MAXIMUM_FUNC(QUOTES_MATRIX(i, 1) - HISTORICAL_MATRIX(j, i + 1), 0)) / (RSI_MATRIX(k, 6) + MAXIMUM_FUNC(0, HISTORICAL_MATRIX(j, i + 1) - QUOTES_MATRIX(i, 1))))
    TEMP_MATRIX(i, 3) = IIf(LOW_TRIGGER_VECTOR(i, 1) > RSI_MATRIX(k, 1), "--", HISTORICAL_MATRIX(j, i + 1) + RSI_MATRIX(k, 6) - RSI_MATRIX(k, 5) * (100 - LOW_TRIGGER_VECTOR(i, 1)) / LOW_TRIGGER_VECTOR(i, 1))
    TEMP_MATRIX(i, 4) = IIf(HIGH_TRIGGER_VECTOR(i, 1) < RSI_MATRIX(k, 1), "--", HISTORICAL_MATRIX(j, i + 1) - RSI_MATRIX(k, 5) + RSI_MATRIX(k, 6) * HIGH_TRIGGER_VECTOR(i, 1) / (100 - HIGH_TRIGGER_VECTOR(i, 1)))
    TEMP_MATRIX(i, 5) = QUOTES_MATRIX(i, 1)
    TEMP_MATRIX(i, 6) = QUOTES_MATRIX(i, 2)
    TEMP_MATRIX(i, 7) = QUOTES_MATRIX(i, 3)
    TEMP_MATRIX(i, 8) = QUOTES_MATRIX(i, 4)
    TEMP_MATRIX(i, 9) = QUOTES_MATRIX(i, 5)
    TEMP_MATRIX(i, 10) = QUOTES_MATRIX(i, 6)
    TEMP_MATRIX(i, 11) = QUOTES_MATRIX(i, 7)
    TEMP_MATRIX(i, 12) = HISTORICAL_MATRIX(j, i + 1)
    TEMP_MATRIX(i, 13) = RSI_MATRIX(k, 1)
Next i

RSI_TARGET_PRICES_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
RSI_TARGET_PRICES_FUNC = Err.number
End Function
