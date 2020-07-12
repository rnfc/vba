Attribute VB_Name = "FINAN_ASSET_TA_GANN_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Gann Charts

'Once upon a time we talked about Templeton's investment philosophy, and we
'concentrated on the one that said: Stock price fluctuations are proportional
'to the square root of the price.

'At the time I assumed that "fluctuations" referred to volatility or standard
'deviation. However, I now find that there are other, more interesting
'interpretations including Square Root of Time theory and Gann Charts and ...

'Yes. William D. Gann (1878-1955) was a trader who designed several unique
'techniques for analyzing price charts. we 'll talk about that shortly, but
'consider this:
'* Suppose the stock price varied with time like so: P(t) = t2
'* At time t+1 the price would be: P(t+1) = (t+1)^2 = t2 + 2t + 1
'* The "variation" or "fluctuation" in price is then: P(t+1) - P(t) =
'2t +1 = 2(t+1) - 1 = 2 {P(t+1)}1/2 - 1
'* In a similar manner, we can get: [P(t+n) - P(t)] / n = 2 {P(t+n)}1/2 - n

'Aha! The square root of price pops up, eh?
'Yes, so that'd lead us to believe that, by "fluctuations", we're talking about
'price changes and ... And do prices change like t2 ?

'That's a parabolic change, right?
'Yes and we've talked about that before, here.
'Let's take a peek at the Gann Wheel, sometimes called the Square of Nine:
'See the circles of numbers, in red, green and blue?
'Pick any number, like 6 in the blue circle.
'The square root of 6 is 2.4 to the nearest tenth..
'To this square root we do this:
'1. Add 0.5 to 2.4 and square to get: 8.7 rounded to 8.
'2. Add 1.0 to 2.4 and square to get: 11.9 rounded to 11.
'3. Add 1.5 to 2.4 and square to get: 15.6 rounded to 15.
'4. Add 2.0 to 2.4 and square to get: 19.8 rounded to 19.

    
'31  32  33  34  35  36  37
'30  13  14  15  16  17  38
'29  12  3   4   5   18  39
'28  11  2   1   6   19  40
'27  10  9   8   7   20  41
'26  25  24  23  22  21  42
'49  48  47  46  45  44  43

'Do you recognize those numbers?

'Okay, let's pick 23 in the green circle.
'The square root of 23 is 4.8 to the nearest tenth..
'To this square root we do this:
'1. Add 0.5 to 4.8 and square to get: 28.
'2. Add 1.0 to 4.8 and square to get: 34.
'3. Add 1.5 to 4.8 and square to get: 40.
'4. Add 2.0 to 4.8 and square to get: 46.

'You 're doing some rounding, right?
'Uh ... yes. Either up or down so I can reproduce Gann's Chart. Remember, it ain't my chart.
'So, do you recognize those numbers?
'Yes, they're at random locations on that chart.
'Random? No, they're the numbers we get when we travel clockwise about the circles, moving to
'the next outer circle when we get to the bottom of the chart.

'In that last example, starting at the number 23, we got the numbers shown here
'And that'll predict stock prices?
'we 'll see.

'Notice that the numbers go from 1 to 40, so if you have a stock with a price whose square root
'was outside this range, you'd rescale it.
'That is, if the price is, say $10,000 (like the DOW index), its square root is 100 which is
'outside the range. So, consider 1/10 the price, namely 1000 with a square root of 31.6, so
'we'd start at the number 32. Or consider 1/100 the price, namely 100 with a square root of
'10, so we'd start there.

'Every 0.5 added to the square root means a 90 degree clockwise rotation.
'The numbers you end up with are "special" ... in some sense which we'll investigate.

'You're gonna use that wheel?
'No. Remember that Gann was born in the 19th century so a wheel was convenient, but we'll
'just ... So why did you mention it? Because it 's interesting!

'Anyway , we 'll do the following:
'[1] Take the square root of the current price.
'[2] Add something to that square root (like 0.5 or 1.0 etc.)
'[3] Square the results to get "special" prices to watch out for.

'Then Sell?
'Or Buy. I should point out that you could subract something from the square root in step #2, to
'identify decreasing prices.

'Example?
'Okay, let's look at recent prices for the S&P500, here
'There was a Low on Aug 13, 2004 of 1060.72, so we do this:
'[1] Take the square root of 1060.72, giving 32.57
'[2] Add something to that 32.57 (like 0.5 or 1.0 etc.)
'[3] Square the results to get "special" prices to watch out for.

'For example, (32.57+2.5)2 = 1229.81 which is very nearly the S&P price on March 7, 2005, namely 1225.81.

'And all the other "special prices? What's so special about them?
'Ask Gann!
    
'Here are a few more:
'increments = 0.25
'increments = -0.25

Function ASSET_TA_GANN_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal DELTA1_VAL As Double = 0.25, _
Optional ByVal DELTA2_VAL As Double = 0.5)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

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

TEMP_MATRIX(0, 8) = "LOW"
TEMP_MATRIX(0, 9) = "( LOW ^ 0.5 + " & Format(DELTA1_VAL, "0.00") & " ) ^ 2"
TEMP_MATRIX(0, 10) = "( LOW ^ 0.5 + " & Format(DELTA2_VAL, "0.00") & " ) ^ 2"

TEMP_MATRIX(0, 11) = "HIGH"
TEMP_MATRIX(0, 12) = "( HIGH ^ 0.5 - " & Format(DELTA1_VAL, "0.00") & " ) ^ 2"
TEMP_MATRIX(0, 13) = "( HIGH ^ 0.5 - " & Format(DELTA2_VAL, "0.00") & " ) ^ 2"

MIN_VAL = 2 ^ 52: MAX_VAL = -2 ^ 52
For i = 1 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    If TEMP_MATRIX(i, 5) > MAX_VAL Then: MAX_VAL = TEMP_MATRIX(i, 5)
    If TEMP_MATRIX(i, 5) < MIN_VAL Then: MIN_VAL = TEMP_MATRIX(i, 5)
Next i

i = 1
TEMP_MATRIX(i, 8) = MIN_VAL
TEMP_MATRIX(i, 9) = (MIN_VAL ^ 0.5 + DELTA1_VAL) ^ 2
TEMP_MATRIX(i, 10) = (MIN_VAL ^ 0.5 + DELTA2_VAL) ^ 2

TEMP_MATRIX(i, 11) = MAX_VAL
TEMP_MATRIX(i, 12) = (MAX_VAL ^ 0.5 - DELTA1_VAL) ^ 2
TEMP_MATRIX(i, 13) = (MAX_VAL ^ 0.5 - DELTA2_VAL) ^ 2

For i = 2 To NROWS
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i - 1, 8)
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i - 1, 9)
    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i - 1, 10)
    TEMP_MATRIX(i, 11) = TEMP_MATRIX(i - 1, 11)
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i - 1, 12)
    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i - 1, 13)
Next i

ASSET_TA_GANN_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_GANN_FUNC = Err.number
End Function
