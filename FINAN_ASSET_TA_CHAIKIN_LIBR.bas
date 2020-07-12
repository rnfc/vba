Attribute VB_Name = "FINAN_ASSET_TA_CHAIKIN_LIBR"

'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------


'Volume Flow Indicators: the Chaikin Oscillator

'Joe Granville published a book in 1963, Granville’s New Key to Stock Market Profits, in which
'he popularized (and named) a method for measuirng "volume flow".

'He called it On Balance Volume. We'll call it OBV.

'One starts with some OBV value (it might as well be 0) and do the following:
'* If today's closing price is larger than yesterday's, we add today's volume to the OBV.
'* If today's closing price is smaller than yesterday's, we subtract today's volume to the OBV.
'* If today's closing price equals yesterday's, we keep yesterday's OBV.

'The idea (enunciated by Granville) is that large volume changes anticipate large price changes.
'Hence we calculate the cumulative volume changes (as indicated above) and use it to anticipate
'price changes.

'Stare intently at the left chart below to see if you can identify that predictive
'characteristic.

'For example, if there's a large increase in volume (indicated by a rising OBV), then (presumably)
'there's a buying pressure that'll cause the price to increase.

'The OBV is like a spring, compressed and ready to generate price changes ... so they say.

'After other reincarnations of the OBV, Marc Chaikin developed the Chaikin Oscillator.
'That's the chart on the right, eh?
'Chaikin goes Like this:
'* Each day we calculate the Close Location Value:
'CLV = ( ( (Close - Low) - (High - Close) ) / (High - Low) )
'* We multiply CLV by todays' volume.
'* We keep track of the cumulative total of CLV*Volume, always adding today's value to yesterday's.
'That running total is called the Accumulation/Distribution Line.
'* Then, just as one does in calculating the MACD, we take the difference between the 3-day Exponential
'Moving Average (EMA) and the 10-day EMA. that 's the Chaikin Oscillator.


'Note that (High - Low) is the range of prices and it's split into two fractions:
'x = (High - Close)/(High - Low)   and   y = (Close - Low)/(High - Low)
'and x + y = 1 and ...

'And CLV = x - y, right?
'Actually, CLV = y - x.
    
'So CLV measures whether the Close is closer to the daily High or the daily Low. When the Close
'is closer to the High, CLV is positive. When it's closer to the Low, CLV is negative.

'When CLV is positive, it indicates a positive outlook on the stock and if that's associated
'with large volume, that suggests an upward pressure on the stock price.
'On the other hand, it CLV is negative ...

'Then the stock price will drop.
'I never said that!
'A negative CLV indicates that the Close is closer to the Low and if there's a large volume of
'trading, that suggests a downward pressure on the stock price.

'Because positive and negative CLV values are indicative of price movements when associated with
'large daily volume, that's the reason for multiplying CLV by the volume and ...

'First you say OBV is like a coiled spring. Now you talk about upward and downward pressures. Are
'we talking physics here? Hmmm ... interesting analogy. It's like Newton's F = ma.
'There 's a force F and a mass m which provides resistance to motion and the acceleration a indicates
'movement and ...

'Can we get back to Chaikin? I assume that CLV is the Chaikin Oscillator.
'I never said that! You haven't been listening. We gotta do the MACD thing!

'Remember:
'The sum of past CLVs is the Accumulation/Distribution Line which we'll call the ADL.
'It measures (much like OBV) positive and negative pressure on the price.

'In the words of Marc Chaikin:
'The closer a stock or average closes to its high, the more accumulation there was.
'Conversely, if a stock closes below its midpoint for the day, there was distribution on that day.
'The closer a stock closes to its low, the more distribution there was.

'Now the price could be moving South and, if a positive ADL pressure stays around for a while, the
'price motion will reverse and head North and ...

'You're kidding, right? Are we talking physical pressure in a Northerly direction ... on an object
'moving South? Uh ... why not? So we look to see if that pressure is increasing, recently, compared
'to what it was in the past. That'd indicate a change in the object's momentum.
'That suggests that we compare some recent average ADL to a longer historical average.
'That suggests moving averages and, if we want to weight recent ADL-values more heavily, that means ...

'That means we'd use an Exponential Moving Average, eh?
'You got it ... so we take the difference between the 3-day MACD (that's the recent average) and the
'10-day MACD.

'Aah, finally we got us the Chaikin Oscillator. But why "accumulation" and "distribution"? Why call them that?
'If the price is high, compared to the mid-point of the daily range (from Low to High), one considers this
'as money flowing in. If the price is low, it's flowing out.

'So the ADL is like a money flow indicator, right?
'Yes. Though it attempts to measure volume flow, in and out, if we were to multiply by the stock price,
'that'd be "money flow" ... in and out.
'The Accumulation/Distribution Line is also one of Chaikin's babies.

'Whether the flow is in or out depends upon the price being high or low compared to the
'midpoint of the daily range, right?
'Yes, according to Chaikin ... but, of course, you can have any criterion you like.
'Maybe you 'd like the price to be high or low compared to, say, the 10-day moving average or
'maybe you'd like to compare the closing price to (Open + High + Low + Close)/4 or maybe ...

'Maybe I don't like that comparison ... to the midpint. Maybe I'd like to compare the Close to, say ...
'Yes.IF you 'd like to generate some prescription you can do that, too.
'See the formula: ((CL - LO) - (HI - CL)) / (HI - LO)

'You can change it to, say: CL - (HI + LO + OP + CL)/4 if you like. Then volume would be added
'if the CLose were greater than (HIgh + LOw + OPen + CLose)/4.
'Type in your favourite formula and click a button and see what happens. It's fun!!

'But won't that change the ... uh ... the ...?
'It 'll change the CLV, hence the cumulative sum (which is the ADL) hence the Exponential Averages
'of the ADL, namely the Chaikin Oscillator.
'Then they'll be your indicator's.

'Well ... maybe I don't like the 3-day and 10-day EMA. Maybe I like the 15-day and 20-day or maybe ...
'Well, you can change that in the spreadsheet. See the sliders near cells R3 and S3?
'When you 're finished playing you get yourself Buy & Sell signals that'll make you a fortune.

'Okay, but I kinda like that OBV volume flow. While Chaikin sounds sexier, it only looks at today's
'prices: High, Low, etc. and makes a decision based on a single snapshot.

'Yes, I was thinking the same thing. With the OBV, it looks to see if the closing price is going up
'or down. It looks at today and yesterday.

'Going up? Investors are buying, they're hitting the asking price, maybe it's the beginning of an
'uptrend .. so we add the volume.

'Going down? Investors are selling at the bid, maybe it's the beginning of an downtrend .. so we
'subtract the volume.

'Well, I think I could invent something ...
'i 've added a piece to the spreadsheet displayed above. It now looks like this:

'Now you can change the formula for calculating the OBV.
'There are a number of variables you can use, namely:
'* OP = today's open, HI = today's high, LO = today's low, CL = today's close
'* Vol = today's volume
'* YC = yesterday's close
'* YO = yesterday's OBV
'* YH = yesterday's high
'* YL = yesterday's low

'Yesterday's OBV?
'If you want to add the OBVs, you need to add today's to yesterday's.
'It says to add, to yesterday's OBV(that's YO), the Vol muliplied by either
'1 (if the today's close is greater than yesterday's)
'-1 ( if the today's close is less than yesterday's)
'0 (if the close hasn't changed)

'You might change that magic OBV formula to, say:   YO + Vol*IF( CL>1.01*YC, 1 , IF( CL<0.99*YC, -1, 0) )
'Then you'd get excited about a possible uptrend (and add the Volume) if today's close were 1%
'greater than yesterday's.

'And you'd subtract the volume if the close fell by 1%, right?

'you 'd get flat tops to your modified OBV because nothing is either added or subtracted ... like this:
'Uh ... looks like you also get flat bottoms.
'Thanks for pointing that out ...

'If you REALLY wanted to identify when people were buying or selling you might change the magic
'OBV formula to: YO + Vol*IF( CL > YH, 1 , IF( CL < YL, -1, 0) )

'which says we add the volume if today's close is greater than yesterday's high and subtract
'if it's smaller than yesterday's low.

'If YH and YL stand for yesterday's high and low, why doesn't YO stand for yesterday's open?

'Some comments:
'* The Chaikin Oscillator uses moving averages hence needs to "get started" before the averages
'are worthy of displaying.

'For that reason, you need a year's worth of data, starts calculating the averages with the first
'downloaded values ... but displays them on the charts starting (about) 3 months later.

'* Changing the days in EMA1 and EMA2 only changes those guys that depend upon
'the EMAs. ... like the Chaikin Oscillator chart.

'* Since only the direction of the OBV is important (is it going up or down?), we always start
'  with OBV = 0.

'* Since the charts end "today" and start (about) 8 months earlier (to get Chaikin's EMAs
'"started"), the OBV chart starts at that time as well.

'That is, all the earlier downloaded data is ignored and OBV is starts at 0 (about) 8 months ago.
'* When you insert a new formula for the OBV, the insertion starts 8 months ago.
'* There's lots of stuff on Chaikin to look at. (I used to provide specific links, but they
'get broke ... so I don't do that no more ^#$%@!)


'REFERENCE: http://www.gummy-stuff.org/volume-flow.htm

Function ASSET_TA_CHAIKIN_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal EMA1_PERIOD As Long = 3, _
Optional ByVal EMA2_PERIOD As Long = 10, _
Optional ByVal START_PERIOD As Long = 94)

'-----------------------------------------------------------------------------
'CLV formula: ((CL-LO)-(HI-CL))/(HI-LO)
'OBV Formula: YO+Vol*IF(CL>YH,1,IF(CL<YL,-1,0))
'-----------------------------------------------------------------------------

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long

Dim ALPHA1_VAL As Double
Dim ALPHA2_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

ALPHA1_VAL = 1 - 2 / (EMA1_PERIOD + 1)
ALPHA2_VAL = 1 - 2 / (EMA2_PERIOD + 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 17)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ_CLOSE"

TEMP_MATRIX(0, 8) = "Y0: LAST OBV"
TEMP_MATRIX(0, 9) = "YC: LAST CLOSE"

TEMP_MATRIX(0, 10) = "OBV: YO + Vol*IF(CL>YH,1,IF(CL<YL,-1,0))"
TEMP_MATRIX(0, 11) = "CLV: ((CL-LO)-(HI-CL))/(HI-LO)"
TEMP_MATRIX(0, 12) = "ADL"

TEMP_MATRIX(0, 13) = EMA1_PERIOD & "_DAY_" & "EMA"
TEMP_MATRIX(0, 14) = EMA2_PERIOD & "_DAY_" & "EMA"
TEMP_MATRIX(0, 15) = "CHAIKIN"

TEMP_MATRIX(0, 16) = "YH: LAST HIGH"
TEMP_MATRIX(0, 17) = "YL: LAST LOW"

i = 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 10000
TEMP_MATRIX(i, 8) = ""
TEMP_MATRIX(i, 9) = ""
TEMP_MATRIX(i, 10) = "" ' 0

TEMP_MATRIX(i, 11) = ((TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 4)) - _
                      (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 5))) / _
                      (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4))

TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 11) * TEMP_MATRIX(i, 6)
TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 12)
TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 12)
TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 13) - TEMP_MATRIX(i, 14)

TEMP_MATRIX(i, 16) = ""
TEMP_MATRIX(i, 17) = ""


For i = 2 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 10000
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i - 1, 10)
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i - 1, 5)
    
    TEMP_MATRIX(i, 16) = TEMP_MATRIX(i - 1, 3)
    TEMP_MATRIX(i, 17) = TEMP_MATRIX(i - 1, 4)

    'YO + Vol*IF( CL > YH, 1,  IF( CL < YL, -1, 0) )
    If i >= START_PERIOD Then
        If TEMP_MATRIX(i, 5) > TEMP_MATRIX(i, 16) Then
            k = 1
        Else
            If TEMP_MATRIX(i, 5) < TEMP_MATRIX(i, 17) Then
                k = -1
            Else
                k = 0
            End If
        End If
        
        If TEMP_MATRIX(i, 8) <> "" Then
            TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 8) + TEMP_MATRIX(i, 6) * k
        Else
            TEMP_MATRIX(i, 10) = 0 + TEMP_MATRIX(i, 6) * k
        End If
    Else
        TEMP_MATRIX(i, 10) = "" '0
    End If
    
    '((CL - LO) - (HI - CL)) / (HI - LO)
    TEMP_MATRIX(i, 11) = ((TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 4)) - _
                          (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 5))) / _
                          (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4))
    'You can change it to, say: CL - (HI + LO + OP + CL)/4
    
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i - 1, 12) + TEMP_MATRIX(i, 11) * TEMP_MATRIX(i, 6)
    
    TEMP_MATRIX(i, 13) = ALPHA1_VAL * TEMP_MATRIX(i - 1, 13) + (1 - ALPHA1_VAL) * TEMP_MATRIX(i, 12)
    TEMP_MATRIX(i, 14) = ALPHA2_VAL * TEMP_MATRIX(i - 1, 14) + (1 - ALPHA2_VAL) * TEMP_MATRIX(i, 12)
    TEMP_MATRIX(i, 15) = TEMP_MATRIX(i, 13) - TEMP_MATRIX(i, 14)
    
Next i

ASSET_TA_CHAIKIN_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_CHAIKIN_FUNC = Err.number
End Function
