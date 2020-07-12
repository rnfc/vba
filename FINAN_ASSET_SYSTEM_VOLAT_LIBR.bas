Attribute VB_Name = "FINAN_ASSET_SYSTEM_VOLAT_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


Function ASSET_VOLATILITY_PRICES_SIGNAL_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal BUY_PERCENT As Double = 0.008, _
Optional ByVal SELL_PERCENT As Double = 0.01, _
Optional ByVal INITIAL_CASH As Double = 100, _
Optional ByVal MA_PERIODS As Double = 20, _
Optional ByVal OUTPUT As Integer = 2)

'Volatility and the Market
'After writing the stuff on September and the Market, I was looking carefully at other
'correlations and found interesting comparisons between stock market prices and volatility
'(or Standard Deviation).

'For example, the crash of '87 shows HUGE volatility in October/87 (see Figure 1) and ...
'But you'd expect that, right?
'Yes, I guess so, but I thought I'd play with the Pearson correlation between volatility
'and stock prices or stock returns to see whether one could anticipate market changes and ...

'And I have a spreadsheet which allows you to download two years worth of daily stock prices
'and you'd get a chart of Prices vs Standard Deviation and you could choose which one-month
'window you'd want to look at and whether you want the correlation between SD & Prices or
'between SD & Returns and ...

'You can move that one month window and choose which correlation you'd like to know and you
'even get a regression line for that one-month window and ...

'You can, for example, choose to Buy when the volatility exceeds 1.0% and Sell when it drops
'below 0.8% (as in the picture, above).

'In the example shown, although GE stock gained 5.8% over the period Sept/04 to Sept/06, your
'Portfolio would have gained 13.5%
'And you'd recommend that strategy?
'Huh? Me? Recommend? I just provide the toys ...

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim NROWS As Long

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double

Dim MEAN_VAL As Double
Dim VOLAT_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, _
                  START_DATE, END_DATE, "DAILY", "DOHLCVA", True, _
                  False, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 12)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ. CLOSE"
TEMP_MATRIX(0, 8) = "RETURNS"
TEMP_MATRIX(0, 9) = "VOLATILITY"
TEMP_MATRIX(0, 10) = "EQUITY"
TEMP_MATRIX(0, 11) = "CASH"
TEMP_MATRIX(0, 12) = "SYSTEM"

k = 3
i = k - 2
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
TEMP_MATRIX(i, 8) = 0
TEMP_MATRIX(i, 9) = 0
TEMP_MATRIX(i, 10) = 0
TEMP_MATRIX(i, 11) = INITIAL_CASH
TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 10) + TEMP_MATRIX(i, 11)

ATEMP_SUM = 0
For i = k - 1 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7) - 1
    
    ATEMP_SUM = ATEMP_SUM + TEMP_MATRIX(i, 8)
    
    If i <= (MA_PERIODS + k) Then
        BTEMP_SUM = 0
        For j = i To k - 1 Step -1
            BTEMP_SUM = BTEMP_SUM + (TEMP_MATRIX(j, 8) - (ATEMP_SUM / (i - 1))) ^ 2
        Next j
        BTEMP_SUM = (BTEMP_SUM / (i - 1)) ^ 0.5
    Else
        l = i - (MA_PERIODS + k - 1)
        ATEMP_SUM = ATEMP_SUM - TEMP_MATRIX(l, 8)
        BTEMP_SUM = 0
        For j = i To (l + 1) Step -1
            BTEMP_SUM = BTEMP_SUM + (TEMP_MATRIX(j, 8) - (ATEMP_SUM / (i - l))) ^ 2
        Next j
        BTEMP_SUM = (BTEMP_SUM / (i - l)) ^ 0.5
    End If
    TEMP_MATRIX(i, 9) = BTEMP_SUM

    If TEMP_MATRIX(i, 9) >= SELL_PERCENT Then
        TEMP_MATRIX(i, 10) = 0
    Else
        If (TEMP_MATRIX(i, 9) <= BUY_PERCENT And TEMP_MATRIX(i - 1, 10) = 0) Then
            TEMP_MATRIX(i, 10) = TEMP_MATRIX(i - 1, 11)
        Else
            TEMP_MATRIX(i, 10) = TEMP_MATRIX(i - 1, 10) * (1 + TEMP_MATRIX(i, 8))
        End If
    End If
    
    If (TEMP_MATRIX(i - 1, 10) > 0 And TEMP_MATRIX(i, 10) = 0) Then
        TEMP_MATRIX(i, 11) = TEMP_MATRIX(i - 1, 10) * (1 + TEMP_MATRIX(i, 8))
    Else
        If (TEMP_MATRIX(i - 1, 10) = 0 And TEMP_MATRIX(i, 10) > 0) Then
            TEMP_MATRIX(i, 11) = 0
        Else
            TEMP_MATRIX(i, 11) = TEMP_MATRIX(i - 1, 11)
        End If
    End If
    
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 10) + TEMP_MATRIX(i, 11)
        
    MEAN_VAL = MEAN_VAL + TEMP_MATRIX(i, 12) / TEMP_MATRIX(i - 1, 12) - 1

Next i

If OUTPUT = 0 Then
    ASSET_VOLATILITY_PRICES_SIGNAL_FUNC = TEMP_MATRIX
    Exit Function
End If

MEAN_VAL = MEAN_VAL / (NROWS - (k - 2))

For i = (k - 1) To NROWS
    VOLAT_VAL = VOLAT_VAL + ((TEMP_MATRIX(i, 12) / TEMP_MATRIX(i - 1, 12) - 1) - MEAN_VAL) ^ 2
Next i
VOLAT_VAL = (VOLAT_VAL / (NROWS - (k - 2))) ^ 0.5

If OUTPUT = 1 Then
    ASSET_VOLATILITY_PRICES_SIGNAL_FUNC = MEAN_VAL / VOLAT_VAL
Else
    ASSET_VOLATILITY_PRICES_SIGNAL_FUNC = Array(MEAN_VAL / VOLAT_VAL, MEAN_VAL, VOLAT_VAL)
End If

Exit Function
ERROR_LABEL:
ASSET_VOLATILITY_PRICES_SIGNAL_FUNC = "--"
End Function
