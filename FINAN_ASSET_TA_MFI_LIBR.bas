Attribute VB_Name = "FINAN_ASSET_TA_MFI_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


Function ASSET_TA_MFI_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MA_PERIOD As Long = 14)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim FACTOR_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

FACTOR_VAL = 10000
'-----------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 14)
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"

'Flow- Total Flow+ Total Flow- Total Flow  MFI

TEMP_MATRIX(0, 8) = "AVG PRICE"
TEMP_MATRIX(0, 9) = "+ DAILY FLOW: " & MA_PERIOD
TEMP_MATRIX(0, 10) = "- DAILY FLOW: " & MA_PERIOD

TEMP_MATRIX(0, 11) = "+ FLOW TOTAL"
TEMP_MATRIX(0, 12) = "- FLOW TOTAL"
TEMP_MATRIX(0, 13) = "TOTAL FLOW"
'Total Flow = (X-day Total Flow+) + (X-day Total Flow-)


i = 1
For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / FACTOR_VAL
TEMP_MATRIX(i, 8) = (TEMP_MATRIX(i, 3) + TEMP_MATRIX(i, 4) + TEMP_MATRIX(i, 5)) / 3
TEMP_MATRIX(i, 9) = 0
TEMP_MATRIX(i, 10) = 0
TEMP_MATRIX(i, 11) = "": TEMP_MATRIX(i, 12) = ""
TEMP_MATRIX(i, 13) = "": TEMP_MATRIX(i, 14) = ""

TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 2 To NROWS
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / FACTOR_VAL
    TEMP_MATRIX(i, 8) = (TEMP_MATRIX(i, 3) + TEMP_MATRIX(i, 4) + TEMP_MATRIX(i, 5)) / 3
    
    If TEMP_MATRIX(i, 8) > TEMP_MATRIX(i - 1, 8) Then
        TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8) * TEMP_MATRIX(i, 6)
    Else
        TEMP_MATRIX(i, 9) = 0
    End If
    
    If TEMP_MATRIX(i, 8) < TEMP_MATRIX(i - 1, 8) Then
        TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 8) * TEMP_MATRIX(i, 6)
    Else
        TEMP_MATRIX(i, 10) = 0
    End If
    
    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 9)
    TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 10)
    If i >= MA_PERIOD Then
        TEMP_MATRIX(i, 11) = TEMP1_SUM
        TEMP_MATRIX(i, 12) = TEMP2_SUM
        TEMP_MATRIX(i, 13) = TEMP1_SUM + TEMP2_SUM
        If TEMP_MATRIX(i, 13) <> 0 Then
            TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 11) / TEMP_MATRIX(i, 13)
        Else
            TEMP_MATRIX(i, 14) = 0
        End If
        TEMP1_SUM = TEMP1_SUM - TEMP_MATRIX(i - MA_PERIOD + 1, 9)
        TEMP2_SUM = TEMP2_SUM - TEMP_MATRIX(i - MA_PERIOD + 1, 10)
    
    Else
        For j = 11 To 14: TEMP_MATRIX(i, j) = "": Next j
    End If
Next i

TEMP_MATRIX(0, 14) = "CURRENT MFI: " & Format(TEMP_MATRIX(NROWS, 14), "0.0%")
'MFI = (X-day Total Flow+) / (Total Flow)

ASSET_TA_MFI_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_MFI_FUNC = Err.number
End Function


'I always wondered how people measured momentum. When a stock price is on the move, upwards,
'maybe it's time to buy and go along for the ride and ... Or on the move down, eh?
'Yes, of course. But what does it mean ... this "on the move" stuff?
'That 's what it's called, but it's measured in so many ways and may involve so many factors and ...
'Strong upward trend in the price chart and greater than expected earnings growth.
'Strengrh, relative to other stocks (or "the market" as a whole), measured (perhaps) by the
'Relative Strength Index. (See RSI, here.)
'The market goes down, but your stock doesn't ... so it may explode when the market starts up again.
'Don't hold on to a falling stock ... and look for one that stays above its 50-day moving average.
'Small cap, low volume, less liquid stocks can make explosive moves.
'Look for new highs before the stock goes even higher.
'Keep your eye on the Directional Movement Indicator and ADX and ...

'In particular, I was interested in weighting the stock price according to the volume of trades
'associated with that price.

'A day when a million shares traded is more significant than a day where a hundred thousand traded,
'so I volume-weighted the prices and called the "modified" value of ADX ...

'Anyway, I now discover that the ol' standby RSI is being volume-weighted.
'Remember that, if RSI = 0.75, then you'd say: "The price has increased 75% of the time, over the
'past umpteen days". Now, RSI, when volume-weighted, is called ...

'Money Flow Index?
'Very good! You've read the title, eh?

'The idea Is this:
'1. Consider some representative daily price, say P = (High + Low + Close)/3     ... a daily
'"average" price
'2. Multiply P by the volume for that day giving P*V     ... representing the value of all
'the day's trades, hence the daily "Money Flow"
'3. Keep track of the Money Flows on those days when the price went UP, calling them "Positive"
'Money Flows
'4. Add ALL the volume-weighted prices for the past umpteen days     ... giving the total "Money Flow"
'over this time period
'5. Add up all those "Positive" Money Flows     ... summing the "positive" volume-weighted prices
'over the past umpteen days
'6. Determine the ratio: MFI = (SUM of positive Money Flows) / (SUM of ALL Money Flows)

'So if MFI = 0.75, then I'd say: "The money flow was positive 75% of the time, over the past umpteen days".
'You can say that if you like. It don't hardly matter to me.

'I 'd suggest playing with the spreadsheet, for various stocks, and seeing if, for example, large
'values of MFI (like 80% or better) indicate that you've reached a maximum stock price or maybe
'small value (like 30%) means you're at a minimum or maybe ...

'Reference: http://www.gummy-stuff.org/MFI.htm

