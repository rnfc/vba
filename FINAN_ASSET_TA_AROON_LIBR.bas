Attribute VB_Name = "FINAN_ASSET_TA_AROON_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'It seems there's a new indicator that'll indicate when a stock is trending up or down, and ...
'It was developed by Tushar Chande in 1995 to determine whether a stock is trending or not and
'how strong the trend is.

'Well ... new to me. Anyway, it's called the Aroon indicator. ("Aroon" means "Dawn's Early Light"
'in Sanskrit.) It goes Like this:
'Pick a time period, for example 10 days.
'Pick out the MAXimum and MINimum closing price over the past 10 days.
'If the MAX occurred M days ago, calculate:
'Aroon(UP) = 100 (1 - M/10)   ... it will be between 0 and 100
'If the MIN occurred N days ago, calculate:
'Aroon(DOWN) = 100 (1 - N/10)   ... it will also be between 0 and 100

'So, if Aroon(UP) is greater than, say, 70, then you might expect an upward trend.
'If Aroon(DOWN) is greater than 70, then you might expect an downward trend.


'That's okay for 2-Feb or thereabouts. It does go DOWN. But what about ...?
'What about 3-Mar, where it says UP?

'Anyway , there 's also an Aroon Oscillator, namely:
'Aroon(Oscillator) = Aroon(DOWN) - Aroon(UP)
'It 'll be between -100 and 100.
 
'Then you pick the number of periods and the UP and DOWN levels (like "70", in cells D1 and G1)
'and you get a bunch of charts.

'The chart of the closing price also indicates the prices where the UP and DOWN levels are exceeded.
'It 's in the lower chart. Look for it.

'Here are a few more examples, using a 10 day time period and aroon levels at 70, over a time
'interval ending in June, 2005.


'http://gummy-stuff.org/aroon.htm
'http://stockcharts.com/education/IndicatorAnalysis/indic-Aroon.htm

Function ASSET_TA_AROON_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal NSIZE As Long = 50, _
Optional ByVal NO_PERIODS As Long = 25, _
Optional ByVal UP_LEVEL As Double = 70, _
Optional ByVal DN_LEVEL As Double = 70)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long

Dim MAX_VAL As Double
Dim MIN_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, _
                  START_DATE, END_DATE, "DAILY", "DOHLCVA", False, _
                  True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 16)

'------------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"

TEMP_MATRIX(0, 8) = NO_PERIODS & " - PERIOD MAX"
TEMP_MATRIX(0, 9) = "PERIODS TO MAX"
TEMP_MATRIX(0, 10) = "AROON (UP)"

TEMP_MATRIX(0, 11) = NO_PERIODS & " - PERIOD MIN"
TEMP_MATRIX(0, 12) = "PERIODS TO MIN"
TEMP_MATRIX(0, 13) = "AROON (DN)"

TEMP_MATRIX(0, 14) = "UP TREND"
TEMP_MATRIX(0, 15) = "DOWN TREND"
TEMP_MATRIX(0, 16) = "OSCILLATOR"

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

l = (NROWS - NSIZE + 1)
For i = 1 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    
    If i >= l Then
        MIN_VAL = 2 ^ 52: MAX_VAL = 2 ^ -52
        k = i - NO_PERIODS
        kk = NO_PERIODS '+ 1
        For j = i To k Step -1
            If TEMP_MATRIX(j, 7) > MAX_VAL Then
                ii = kk
                MAX_VAL = TEMP_MATRIX(j, 7)
            End If
            
            If TEMP_MATRIX(j, 7) < MIN_VAL Then
                jj = kk
                MIN_VAL = TEMP_MATRIX(j, 7)
            End If
            kk = kk - 1
        Next j
        TEMP_MATRIX(i, 8) = MAX_VAL
        TEMP_MATRIX(i, 9) = NO_PERIODS - ii '+ 1
        TEMP_MATRIX(i, 10) = 100 * (NO_PERIODS - TEMP_MATRIX(i, 9)) / NO_PERIODS
        TEMP_MATRIX(i, 11) = MIN_VAL
        TEMP_MATRIX(i, 12) = NO_PERIODS - jj '+ 1
        TEMP_MATRIX(i, 13) = 100 * (NO_PERIODS - TEMP_MATRIX(i, 12)) / NO_PERIODS
        TEMP_MATRIX(i, 14) = IIf(TEMP_MATRIX(i, 10) >= UP_LEVEL, TEMP_MATRIX(i, 7), "")
        TEMP_MATRIX(i, 15) = IIf(TEMP_MATRIX(i, 13) >= DN_LEVEL, TEMP_MATRIX(i, 7), "")
        TEMP_MATRIX(i, 16) = TEMP_MATRIX(i, 13) - TEMP_MATRIX(i, 10)
    Else
        For j = 8 To 16: TEMP_MATRIX(i, j) = "": Next j
    End If
    
Next i

ASSET_TA_AROON_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_AROON_FUNC = Err.number
End Function

'http://gummy-stuff.org/Magic%20Box.doc

'The reason that I named this setup the Magic Box was so that you could relate to it and remember
'it faster. Without relationship to something, it is nothing to us.  It must be something memorable
'to you before you can set it into your subconscious mind and see it always.  If I said Snow White,
'or Peter Pan, or Moby Dick you could immediately draw it into your memory because it has relation
'to your past.  The Magic Box is named to bring good feelings and hope.  Just as Fairy Square and
'Knight's Crossing which will become memorable when we do their special presentation later.

'(A) The Magic Box is a pure two-day pattern.  Basically the "close" of a red candle printing on or
'near lower Bollinger Band with the next day printing an "open white candle" (A1) On day one the
'Aroon Down is setting at 100 on the indicator and price is near the lower Bollinger Band.
'The Candlestick is dark-shadowed.  (A2) On day two, to form the Magic Box, the Aroon Down must
'drop to 87.50, and price to form a white candle, to complete the box.  No other indicator is
'watched until this occurs for this setup, just the Bollinger Band and the Aroon Down.

'If you see a Magic Box develop, that being, on the first day the Aroon Down going from 100.00
'while the close is either on or just above the lower Bollinger Band and the second day the Aroon
'Down is now 87.50, you have a Magic Box.

'(B) Now, the buy opportunity occurs when the confirmation of two things happens: on day three or
'few days later the Aroon Down goes to 75.00 and the Williams%R comes above the -50%.  That is the
'buy.  The Williams must come through the -50%.  It may take more than three days but you must wait
'for the confirmation of the William%R. I have seen again and again, the Aroon Down come down from
'100.00 to zero without the Williams ever crossing the -50%, those are the ones you pass on.
'There will be no growth when this happens.  At most there is consolidation or a small drop.  During
'this process the Aroon Up can rise slowly but without the Williams you have nothing.

'There are other Magic Boxes that develop in a stock's life cycle, too, and they, too, can give great
'growth.  These are the supported median, the floating, and the rising Magic Boxes.  The buy-in
'criteria, though, will remain the same for all of them.  The Aroon Down comes to the 75.00 and the
'Williams crosses the -50%.  June 7th, 8th, and 9th accomplish this in the HGR chart, though the
'William%R took until June 13th to confirm. You must wait.
 

'Trading: It is highly recommended that you paper trade the system to familiarize yourself with
'it thoroughly.  As with anything in life nothing is guaranteed, so, always use appropriate stop
'loss according to your risk tolerance.

