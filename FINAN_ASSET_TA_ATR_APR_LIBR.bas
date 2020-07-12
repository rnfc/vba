Attribute VB_Name = "FINAN_ASSET_TA_ATR_APR_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'Average True Range
'http://www.marketmasters.com.au/86.0.html
'http://www.gummy-stuff.org/ATR.htm

'James Welles Wilder introduced a number of "Technical Indicators" such as:
'Directional Movement Indicator, Relative Strength Index and Parabolic SAR and now
'I find there's yet another that ...
'Well , yes.It 's called Average True Range and it goes like this:
'Each day you calculate the largest of the following numbers:
'1. (today's High) - (today's Low)
'2. | (today's High) - (yesterday's Close) |
'3. | (today's Low) - (yesterday's Close) |
'The number is called the True Range (or TR).
'You then average these "True Ranges" over the past N days, calling the average the Average
'True Range (or ATR).

'Note that TR is the largest price movement in the last 24 hours, from yesterday's close to
'today's close. Note, too, that ATR is a smoooothing ritual.

'And what does ATR tell you?
'Apparently large or small values of ATR may indicate a dramatic price movement or trend reversal.
'For example, Figure 1 shows the prices of Exxon stock over the period June, 2005 to June, 2006.
'The largest value of ATR occurs where there's a significant change in the direction of price movement.

'So XOM stops tanking, right?
'That 's one way to put it, but ...
'And what about the other large values of ATR? What happens there?
'Your guess is as good as mine.
'Why 14 day average? Why not 23 day or 32 day or ...
'I think Wilder suggested 14-day, but you can pick any number you like.
    
'For example, here are some other averages for the same stock over the same period:
'I have an even better way to determine whether the stock prices will change direction.
'Yeah, what is it?
'Chart Price versus Date and watch for a change in direction.
'Very funny. Of course, we'd have to see if a "large" ATR value will actually predict a change and ...

'And what constitutes a "large" value?
'Exactly! If we're talking about $100-dollar stocks there will be plenty of "large" ATR values compared
'to a $2-dollar stock. Maybe we should scale the ATR values, eh?
'However, before we do that, it seems that using ATR as a dollar value is used by some as a stop loss
'mechanism.

'You 'd buy a stock at a price $P and calculate P - 2*ATR and sell if the price fell to that value.
'For example, if P = $10 and ATR = $0.65 then P - 2*ATR = 10.00 - 1.30 = 8.70 and that'd be your stop loss.

'My stop loss? I'd say it was yours. Anyway, what's a "large" ATR value?
'Okay , Here 's what we'll do ... divide everything by (yesterday's Close):

'Each day you calculate the largest of the following numbers:
'1. [ (today's High)- (today's Low)] / (yesterday's Close)         ... expressed as a percentage
'2. | (today's High) / (yesterday's Close) - 1 |     ... expressed as a percentage
'3. | (today's Low) / (yesterday's Close) - 1|     ... expressed as a percentage
'Called these numbers Percentage Ranges (or PR).
'You then average these "Percentage Ranges" over the past N days, calling the average the Average
'Percentage Range (or APR).

'So We 're expressing today's High and Low in terms of percentage changes from yesterday's Close.

'Here are some examples for the period June/05 to June/06:
'Wrong! But it seems you'd look for APRs of at least 3%.

'And if I want the reg'lar ATR instead of APR?
'There 's a button which looks like so you click the one you want.
'I might also mention that Wilder regarded ATR as a proxy for stock Volatility so we could ...
'And is it anything like Volatility?
'Yeah, it seems so. For example, here's a couple of charts comparing 14-day APR and Volatility:
'What about 14-day ATR?
'Similar. After all, APR is just a scaled version of ATR. For example, here are two charts
'comparing ATR and APR :

'By the way, there's this stock with an ATR of $35 and ...
'Whooeee! Now that is large!
'Ah, but it's BRK-B, currently selling for about $3,000 a share.
'If'n we look at the APR, it isn't that large ... about 1%

'Oh, I forgot to mention that, although ATR (or, perhaps more appropriately, APR) might stand in
'for volatility (that is, standard deviation), if we were to look at the distribution of daily
'returns, say for that XOM stock from June, 2005 to June, 2006, and calculate the Mean and Standard
'Deviation of those returns, then plot a Normal distribution with that same Mean and Standard
'Deviation we'd get an approximation to the return distribution as in Figure 3.

'Now we do the same, but use the Standard Deviation of the PR values instead (using that Standard
'Deviation as a stand-in for the actual Standard Deviation).


'Okay, but is that common, garden-variety average appropriate? I mean ...
'I know exactly what you mean, Indeed, I understand that some people (including Wilder?) use EMA, the
'Exponential Moving Average, instead, so I ...

'So you have a choice of ATR or APR and, for each, a regular Average or an Exponential Average, like so:


Function ASSET_TA_ATR_APR_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal APR_PERIOD As Long = 14, _
Optional ByVal VERSION As Integer = 0)

'If VERSION = 0 Then: ATR
'If VERSION > 0 Then: APR

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MAX_VAL As Double
Dim TEMP_SUM As Double
Dim ALPHA_VAL As Double
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

ALPHA_VAL = 1 - 2 / (APR_PERIOD + 1)
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
TEMP_MATRIX(0, 8) = "H - L"
TEMP_MATRIX(0, 9) = "H - pC"
TEMP_MATRIX(0, 10) = "L - pC"
TEMP_MATRIX(0, 11) = "PR"
TEMP_MATRIX(0, 12) = Format(APR_PERIOD, "0") & " - PERIODS EMA"
TEMP_MATRIX(0, 13) = Format(APR_PERIOD, "0") & " - PERIODS AVG"

i = 1
For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000

For j = 8 To 13: TEMP_MATRIX(i, j) = "": Next j

For i = 2 To NROWS
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    If VERSION <> 0 Then
        TEMP_MATRIX(i, 8) = (TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4)) / TEMP_MATRIX(i - 1, 5)
        TEMP_MATRIX(i, 9) = Abs(TEMP_MATRIX(i, 3) / TEMP_MATRIX(i - 1, 5) - 1)
        TEMP_MATRIX(i, 10) = Abs(TEMP_MATRIX(i, 4) / TEMP_MATRIX(i - 1, 5) - 1)
    Else
        TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4)
        TEMP_MATRIX(i, 9) = Abs(TEMP_MATRIX(i, 3) - TEMP_MATRIX(i - 1, 5))
        TEMP_MATRIX(i, 10) = Abs(TEMP_MATRIX(i, 4) - TEMP_MATRIX(i - 1, 5))
    End If
    MAX_VAL = 2 ^ -52
    If TEMP_MATRIX(i, 8) > MAX_VAL Then: MAX_VAL = TEMP_MATRIX(i, 8)
    If TEMP_MATRIX(i, 9) > MAX_VAL Then: MAX_VAL = TEMP_MATRIX(i, 9)
    If TEMP_MATRIX(i, 10) > MAX_VAL Then: MAX_VAL = TEMP_MATRIX(i, 10)
    TEMP_MATRIX(i, 11) = MAX_VAL
    If i <> 2 Then
        TEMP_MATRIX(i, 12) = ALPHA_VAL * TEMP_MATRIX(i - 1, 12) + (1 - ALPHA_VAL) * TEMP_MATRIX(i, 11)
    Else
        TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 11)
    End If
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 11)
    If i <= APR_PERIOD + 1 Then
        TEMP_MATRIX(i, 13) = TEMP_SUM / (i - 1)
    Else
        TEMP_MATRIX(i, 13) = TEMP_SUM / (APR_PERIOD + 1)
        TEMP_SUM = TEMP_SUM - TEMP_MATRIX(i - APR_PERIOD, 11)
    End If
Next i

ASSET_TA_ATR_APR_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_ATR_APR_FUNC = Err.number
End Function
