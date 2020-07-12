Attribute VB_Name = "FINAN_ASSET_TA_STOCH_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'http://www.gummy-stuff.org/williams.htm
'http://www.gummy-stuff.org/Bollinger.htm#OSCILLATOR

'Long when %R rises above UPPER_BOUND
'Short when %R falls below LOWER_BOUND

Function ASSET_TA_STOCHASTIC_SIGNAL_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal STOCHASTIC_PERIODS As Long = 20, _
Optional ByVal UPPER_BOUND As Double = 0.8, _
Optional ByVal LOWER_BOUND As Double = 0.2, _
Optional ByVal MA1_PERIOD As Long = 20, _
Optional ByVal MA2_PERIOD As Long = 100, _
Optional ByVal PERIODS_BACK As Long = 50)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ATEMP_SUM As Double
Dim BTEMP_SUM As Double

Dim MAX_VAL As Double
Dim MIN_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

'-----------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 21)
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"
TEMP_MATRIX(0, 8) = MA1_PERIOD & " - PERIOD AVG"
TEMP_MATRIX(0, 9) = MA2_PERIOD & " - PERIOD AVG"
TEMP_MATRIX(0, 10) = "HIGH"
TEMP_MATRIX(0, 11) = "LOW"
TEMP_MATRIX(0, 12) = "STOCHASTIC"
TEMP_MATRIX(0, 13) = "WILLIAMS %R"
TEMP_MATRIX(0, 14) = "STOCHASTIC>"
TEMP_MATRIX(0, 15) = "STOCHASTIC<"
TEMP_MATRIX(0, 16) = "WILLIAMS %R>"
TEMP_MATRIX(0, 17) = "WILLIAMS %R<"
TEMP_MATRIX(0, 18) = "UPPER BOUND"
TEMP_MATRIX(0, 19) = "LOWER BOUND"
TEMP_MATRIX(0, 20) = "LONG SIGNAL: %R > " & Format(UPPER_BOUND, "0.0%")
TEMP_MATRIX(0, 21) = "SHORT SIGNAL: %R < " & Format(LOWER_BOUND, "0.0%")

i = 1
For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7)
TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 7)
For j = 10 To 21: TEMP_MATRIX(i, j) = "": Next j

ATEMP_SUM = TEMP_MATRIX(i, 7)
BTEMP_SUM = TEMP_MATRIX(i, 7)
For i = 2 To NROWS
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    
    ATEMP_SUM = ATEMP_SUM + TEMP_MATRIX(i, 7)
    If i < MA1_PERIOD Then
        TEMP_MATRIX(i, 8) = ATEMP_SUM / i
    Else
        TEMP_MATRIX(i, 8) = ATEMP_SUM / MA1_PERIOD
        ATEMP_SUM = ATEMP_SUM - TEMP_MATRIX(i - MA1_PERIOD + 1, 7)
    End If
    
    BTEMP_SUM = BTEMP_SUM + TEMP_MATRIX(i, 7)
    If i < MA2_PERIOD Then
        TEMP_MATRIX(i, 9) = BTEMP_SUM / i
    Else
        TEMP_MATRIX(i, 9) = BTEMP_SUM / MA2_PERIOD
        BTEMP_SUM = BTEMP_SUM - TEMP_MATRIX(i - MA2_PERIOD + 1, 7)
    End If
    
    If i >= NROWS - PERIODS_BACK Then
        MIN_VAL = 2 ^ 52: MAX_VAL = 2 ^ -52
        For j = i To (i - STOCHASTIC_PERIODS) Step -1
            If TEMP_MATRIX(j, 7) > MAX_VAL Then: MAX_VAL = TEMP_MATRIX(j, 7)
            If TEMP_MATRIX(j, 7) < MIN_VAL Then: MIN_VAL = TEMP_MATRIX(j, 7)
        Next j
        TEMP_MATRIX(i, 10) = MAX_VAL
        TEMP_MATRIX(i, 11) = MIN_VAL
        TEMP_MATRIX(i, 12) = (TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 11)) / (TEMP_MATRIX(i, 10) - TEMP_MATRIX(i, 11))
        TEMP_MATRIX(i, 13) = (TEMP_MATRIX(i, 10) - TEMP_MATRIX(i, 7)) / (TEMP_MATRIX(i, 10) - TEMP_MATRIX(i, 11))
                             
        TEMP_MATRIX(i, 14) = IIf(TEMP_MATRIX(i, 12) > UPPER_BOUND, TEMP_MATRIX(i, 12), "")
        TEMP_MATRIX(i, 15) = IIf(TEMP_MATRIX(i, 12) < LOWER_BOUND, TEMP_MATRIX(i, 12), "")
        TEMP_MATRIX(i, 16) = IIf(TEMP_MATRIX(i, 13) > UPPER_BOUND, TEMP_MATRIX(i, 13), "")
        TEMP_MATRIX(i, 17) = IIf(TEMP_MATRIX(i, 13) < LOWER_BOUND, TEMP_MATRIX(i, 13), "")
        
        TEMP_MATRIX(i, 18) = UPPER_BOUND 'Long Bound
        TEMP_MATRIX(i, 19) = LOWER_BOUND 'Short Bound
        
        TEMP_MATRIX(i, 20) = IIf((TEMP_MATRIX(i, 13) > UPPER_BOUND And TEMP_MATRIX(i - 1, 13) < UPPER_BOUND), TEMP_MATRIX(i, 7), "")
        TEMP_MATRIX(i, 21) = IIf((TEMP_MATRIX(i, 13) < LOWER_BOUND And TEMP_MATRIX(i - 1, 13) > LOWER_BOUND), TEMP_MATRIX(i, 7), "")

    Else
        For j = 10 To 21: TEMP_MATRIX(i, j) = "": Next j
    End If
Next i

ASSET_TA_STOCHASTIC_SIGNAL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_STOCHASTIC_SIGNAL_FUNC = Err.number
End Function

'Stochastic oscillator

'Here we compare the current stock price, P, with the smallest and largest stock prices over
'the past N days: its smallest daily Low and its largest daily High. If these are L and H
'respectively, then we determine how far up the range from L to H that the current price lies.
'That percentage is: %K = 100 (P - L) / (H - L)

'(You get %R = 100% when the Price equals the High over the past N days.)
'As you might imagine, we plot the values of %K (which lie between 0% and 100%) ... either with
'or without smoothing. (Smoothing involves taking a 2- or 3-day average of the %K values). No
'smoothing? It's a fast stochastic. In addition, we calculate the M day moving average of %K ...
'and call it %D. This might be a weighted average, such as described above.

'In any case, we watch to see when %K falls below or above some magic number (like below 20% or
'above 80%) ... or when it crosses %D. The chart below shows a few months of %K (with N=10 days)
'and a 2-day (smoothed) version and an M=5-day, simple moving average (that's %D) and some red
'and green arrows at the 20% and 80% values ... meaning SELL ... or maybe BUY ... or maybe ...

'Whereas the Stochastic Oscillator compares the current closing price with the Lowest price over
'the previous N days, the Williams %R determines how far down the range from H to L that the
'current price lies. %R = 100 (H - P) / (H - L)
'(You get %R = 100% when the Price equals the Low over the past N days.)
'Note that %K + %R = 100%

'Williams %R

'When buying and selling stock the principle rule is to Buy Low, Sell High.
'The questions become:
'1. Low compared to what?
'2. High compared to what?

'There are a jillion suggested answers to these questions, like Bollinger bands,
'moving averages, MACD, stochastics, Fibonacci, etc. (See this.)

'One that I never heard of is Williams %R (named after Larry Williams). I was, however,
'surprised to find that it's just the stochastic oscillator ... but upside down?

'Williams %R says:
'1. Low compared to the lowest price over the past N days.
'2. High compared to the highest price over the past N days.

'Williams %R
'Williams measures the current price on a scale running down from the highest to the lowest price.
'%R = 100(High - Price) / (High - Low)
'1. When %R is close to 100%, the price is close to the Low over the past N days.
'2. When %R is close to 0%, the price is close to the High over the past N days.

'The answers to the above questions?
'1. Buy when %R is close to 100%
'2. Sell when %R is close to 0%

'That puts the current price on a scale running up from the lowest to the highest price (over the
'past N days). %K = 100(Price - Low) / (High - Low)
    

'Stochastic oscillator
'Notice that Buy and Sell signals are the same. Only the name has been changed.

'>Buy and Sell signals?
'Yes. When %R is less than, say 20%, then Buy and ...
'>When it's above, say 80%, then Sell. Right?
'You got it.
'Here 's an example (using daily General Motors prices over the past few months)

'There are:
'20% and 80% levels.
'A plot of %R and the stock Price.
'Whenever %R drops below 20% there's a Sell signal.
'When it rises above 80% ...
'>Sell! NO.it 's Buy.
