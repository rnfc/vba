Attribute VB_Name = "FINAN_ASSET_SYSTEM_GDX_LIBR"

'////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////

'Where one compares the changes in daily highs and lows and determines whether
'investors feel that your stock is going up (when the daily highs increase more
'than the lows decrease).

'Bullish Sign!

'When the highs increase more than the lows decrease? The ADX considers an exponential
'moving average, and that'd give a Buy signal.

'Stock price going up? Get on the bandwagon! The bulls are out! ADX is high! Buy! Buy!

'I assume that if ADX is low ... you sell, right? Yes , That 's the traditional interpretation,
'so when one of my favourite stocks was going up the past few days and ADX was high I thought
'SELL! You mean BUY! No, my reaction was to sell.

'Then I looked at the ADX for the stock using the function shown here and got something like
'Figure 1 which said Buy. That 's when I discovered that Buy signals were, for me at least,
'more like Sell signals so ...

'You mean: Stock prices are going up? Sell!
'Well, not the stock price but the ADX.

'>I'd say Buy when ADX is low, right?
'Yes , I 'd say so. That's just the opposite of what is suggested by those who play with the
'Directional Movement Indicator (and ADX which is derived therefrom).

'Uh ... yes, but it ignores highs and lows and just uses the closing prices. It 's sorta cheating,
'but then the proof is in the pudding, eh? Besides, it works for mutual funds which have a single
'price each day. Anyway , it 's interesting to consider the math behind using just the closing price
'(and no volume weighting as I did with VDX) ... and that's what I'd like to do now.

'Reference: http://www.gummy-stuff.org/ADX-stuff.htm

Function ASSET_GDX_SIGNAL_FUNC(ByRef TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal REFERENCE_DATE As Date, _
Optional ByVal SELL_PERCENT As Double = 0.3, _
Optional ByVal BUY_PERCENT As Double = -0.3, _
Optional ByVal EMA_PERIODS As Long = 14, _
Optional ByVal INITIAL_CASH As Double = 1000, _
Optional ByVal epsilon As Double = 0.001, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ALPHA_VAL As Double
Dim TEMP_SUM As Double
Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    DATA_MATRIX = _
    YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
    "d", "DOHLCVA", False, True, True)
End If
NROWS = UBound(DATA_MATRIX, 1)


If REFERENCE_DATE = 0 Or _
   REFERENCE_DATE < DATA_MATRIX(3, 1) Then
    REFERENCE_DATE = DATA_MATRIX(3, 1)
End If

NCOLUMNS = 20
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
For i = 1 To NROWS: For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = "": Next j: Next i

ALPHA_VAL = 1 - 2 / (EMA_PERIODS + 1)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"
TEMP_MATRIX(0, 8) = "CHANGE"
TEMP_MATRIX(0, 9) = "U"
TEMP_MATRIX(0, 10) = "L"
TEMP_MATRIX(0, 11) = "EMA(U)"
TEMP_MATRIX(0, 12) = "EMA(L)"
TEMP_MATRIX(0, 13) = "GDX"
TEMP_MATRIX(0, 14) = "SELL"
TEMP_MATRIX(0, 15) = "BUY"
TEMP_MATRIX(0, 16) = "INVESTED"
TEMP_MATRIX(0, 17) = "CASH"
TEMP_MATRIX(0, 18) = "PORTFOLIO"
TEMP_MATRIX(0, 19) = "SELL TRIGGER"
TEMP_MATRIX(0, 20) = "BUY TRIGGER"

i = 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j

k = 0
TEMP_SUM = 0
'-----------------------------------------------------------------------------------
For i = 2 To NROWS
'-----------------------------------------------------------------------------------
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) - TEMP_MATRIX(i - 1, 7)
    TEMP_MATRIX(i, 9) = IIf(TEMP_MATRIX(i, 8) > epsilon, TEMP_MATRIX(i, 8), epsilon)
    TEMP_MATRIX(i, 10) = IIf(TEMP_MATRIX(i, 8) < -epsilon, -TEMP_MATRIX(i, 8), epsilon)
    
    If i <> 2 Then
        TEMP_MATRIX(i, 11) = ALPHA_VAL * TEMP_MATRIX(i - 1, 11) + (1 - ALPHA_VAL) * TEMP_MATRIX(i, 9)
        TEMP_MATRIX(i, 12) = ALPHA_VAL * TEMP_MATRIX(i - 1, 12) + (1 - ALPHA_VAL) * TEMP_MATRIX(i, 10)
    Else
        TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 9)
        TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 10)
    End If
    TEMP_MATRIX(i, 13) = (TEMP_MATRIX(i, 11) - TEMP_MATRIX(i, 12)) / _
                         (TEMP_MATRIX(i, 11) + TEMP_MATRIX(i, 12))
                         
        
    If TEMP_MATRIX(i, 1) >= REFERENCE_DATE Then
        
        If TEMP_MATRIX(i, 1) = REFERENCE_DATE Then
            TEMP_MATRIX(i - 1, 16) = 0
            TEMP_MATRIX(i - 1, 17) = INITIAL_CASH
            TEMP_MATRIX(i - 1, 18) = TEMP_MATRIX(i - 1, 16) + TEMP_MATRIX(i - 1, 17)
            l = i + 1
        End If
        
        TEMP_MATRIX(i, 14) = IIf(TEMP_MATRIX(i, 13) > SELL_PERCENT, TEMP_MATRIX(i, 7), -1)
        TEMP_MATRIX(i, 15) = IIf(TEMP_MATRIX(i, 13) < BUY_PERCENT, TEMP_MATRIX(i, 7), -1)
        
        If TEMP_MATRIX(i, 14) > 0 Then
            TEMP_MATRIX(i, 16) = 0
        Else
            If TEMP_MATRIX(i, 15) > 0 And TEMP_MATRIX(i - 1, 17) > 0 Then
                TEMP_MATRIX(i, 16) = TEMP_MATRIX(i - 1, 17)
            Else
                TEMP_MATRIX(i, 16) = TEMP_MATRIX(i - 1, 16) * TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7)
            End If
        End If
        
        If TEMP_MATRIX(i, 14) > 0 And TEMP_MATRIX(i - 1, 16) > 0 Then
            TEMP_MATRIX(i, 17) = TEMP_MATRIX(i - 1, 16) * TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7)
        Else
            If TEMP_MATRIX(i, 15) > 0 Then
                TEMP_MATRIX(i, 17) = 0
            Else
                TEMP_MATRIX(i, 17) = TEMP_MATRIX(i - 1, 17)
            End If
        End If
        
        TEMP_MATRIX(i, 18) = TEMP_MATRIX(i, 16) + TEMP_MATRIX(i, 17)
        If TEMP_MATRIX(i, 1) > REFERENCE_DATE Then
            TEMP_SUM = TEMP_SUM + (TEMP_MATRIX(i, 18) / TEMP_MATRIX(i - 1, 18) - 1)
            k = k + 1
        End If
        TEMP_MATRIX(i, 19) = SELL_PERCENT
        TEMP_MATRIX(i, 20) = BUY_PERCENT
    End If
'-----------------------------------------------------------------------------------
Next i
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------
    ASSET_GDX_SIGNAL_FUNC = TEMP_MATRIX
'-----------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------
    If k = 0 Then: GoTo ERROR_LABEL
    MEAN_VAL = TEMP_SUM / k
    SIGMA_VAL = 0
    For i = l To NROWS
        SIGMA_VAL = SIGMA_VAL + ((TEMP_MATRIX(i, 18) / TEMP_MATRIX(i - 1, 18) - 1) - MEAN_VAL) ^ 2
    Next i
    SIGMA_VAL = (SIGMA_VAL / k) ^ 0.5
    If OUTPUT = 1 Then
        ASSET_GDX_SIGNAL_FUNC = MEAN_VAL / SIGMA_VAL
    Else
        ASSET_GDX_SIGNAL_FUNC = Array(MEAN_VAL / SIGMA_VAL, MEAN_VAL, SIGMA_VAL)
    End If
'-----------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_GDX_SIGNAL_FUNC = "--"
End Function
