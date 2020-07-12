Attribute VB_Name = "FINAN_ASSET_PAIR_SISTER_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Sister Stocks where we got us two stocks that tend to move together.

'Remember that when there is dramatic deviation from their historical
'relationship, then we consider buying or maybe selling and ...

Function ASSETS_SISTER_STOCK_PREDICTORS_FUNC(ByRef TICKER1_STR As Variant, _
ByRef TICKER2_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MAX_MIN_PERIODS As Long = 20)

'MAX_MIN_PERIODS: Move slider to identify Max and Min over +/ days.

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim MAX1_VAL As Double
Dim MAX2_VAL As Double

Dim MIN1_VAL As Double
Dim MIN2_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA1_MATRIX As Variant
Dim DATA2_MATRIX As Variant

On Error GoTo ERROR_LABEL

'Once upon a time we talked about Sister Stocks where we got us two stocks that tend
'to move together. When there 's a dramatic deviation from their historical relationship,
'then we consider buying or maybe selling and ...

'Assuming they'll soon return to their historical association, right?
'You got it. Often, you're comparing a stock to some benchmark.

'Anyway, here we want to talk about ... uh, cousins. Stocks that aren't that close, but
'separated by a few days or weeks. We 'd like to have two stocks, Stock#1 and Stock#2, where
'Stock#1 tends to make its moves in advance of Stock#2.


'You mean #1 predicts the movement of #2?
'Hey! You 're gettin' schmarter! That's exactly what we're looking for!
'Well, I mean .. not exactly the day-to-day movements.

'We 'd just like to have the maxima and minima for Stock #1 precede the extrema of Stock
'#2. That way, we'd know in advance that, pretty soon, we should buy or sell Stock #2 and ...

'You're kidding, right? Stock #1 tells you when to buy or sell #2?
'That 's our hope and ... You 're dreamin'. If I knew how to predict like that I wouldn't be
'listening to you. I'd be ... Sitting in your mansion in the Caribbean sipping Piña Coladas ...
'yeah, I know, but I'm not asking that this advance warning be 100% accurate. Just consistent
'enough to make a few bucks. I assume you're investing in #2 and looking for #1 to tell you
'that a buy or sell is coming up ... soon. Yeah, exactly. In most pairs of stocks, there'll
'be no such predictive qualities. Indeed, it's easy to find pairs of stocks that behave almost
'exactly the same, day-by-day.
    
'Clearly, simultaneous movements up and down ain't useful. For example, the DOW and the S&P500:
'Then there are pairs that have little or no relationship:
'And then there's your invention ... in Figure 1, eh?
'Uh ... actually it's the past few months of GM and DRYS.
'And you've actually made a bundle on the predictive characteristics of GM, vis-a-vis DRYS??
'If I'd made a bundle I wouldn't be talking to you. I'd be ...
'Sitting in your mansion in the Caribbean sipping Piña Coladas ... yeah, I know.

'Pick a stock that we're interested in (Example: CBQ.TO in cell Q6)
'Move a slider to pick a range of days over which we identify Maximum and Minimum (Example:
'we're look at Max & Min over 20 days)

'Wait! What does that mean ... looking at Max & Min over 20 days?
'We run through all the days and identify those days that have the Maximum price over a
'40-day range, from -20 days to +20 days. For example, on July 23, 2007, GE closed at $40.82
'which was the largest closing price from (July 23 - 20 days) to (July 23 + 20 days).
'Then we stick the date on the charts.


If IsArray(TICKER1_STR) = True Then
    DATA1_MATRIX = TICKER1_STR
    TICKER1_STR = "STOCK A"
Else
    DATA1_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER1_STR, _
    START_DATE, END_DATE, "d", "DA", False, True, True)
End If

If IsArray(TICKER2_STR) = True Then
    DATA2_MATRIX = TICKER2_STR
    TICKER2_STR = "STOCK B"
Else
    DATA2_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER2_STR, _
    START_DATE, END_DATE, "d", "DA", False, True, True)
End If

If UBound(DATA1_MATRIX, 1) <> UBound(DATA2_MATRIX, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(DATA1_MATRIX, 1)

'------------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 7)
'------------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = TICKER1_STR & "_CLOSING_PRICES"
TEMP_MATRIX(0, 3) = TICKER2_STR & "_CLOSING_PRICES"

TEMP_MATRIX(0, 4) = TICKER1_STR & "_" & MAX_MIN_PERIODS & "_DAYS_MAXIMA"
TEMP_MATRIX(0, 5) = TICKER2_STR & "_" & MAX_MIN_PERIODS & "_DAYS_MAXIMA"

TEMP_MATRIX(0, 6) = TICKER1_STR & "_" & MAX_MIN_PERIODS & "_DAYS_MINIMA"
TEMP_MATRIX(0, 7) = TICKER2_STR & "_" & MAX_MIN_PERIODS & "_DAYS_MINIMA"
'------------------------------------------------------------------------------

TEMP_MATRIX(1, 1) = DATA1_MATRIX(1, 1)
TEMP_MATRIX(1, 2) = DATA1_MATRIX(1, 2)
TEMP_MATRIX(1, 3) = DATA2_MATRIX(1, 2)
TEMP_MATRIX(1, 4) = ""
TEMP_MATRIX(1, 5) = ""
TEMP_MATRIX(1, 6) = ""
TEMP_MATRIX(1, 7) = ""

'------------------------------------------------------------------------------
For i = 2 To NROWS
'------------------------------------------------------------------------------
    TEMP_MATRIX(i, 1) = DATA1_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA1_MATRIX(i, 2)
    TEMP_MATRIX(i, 3) = DATA2_MATRIX(i, 2)
    
    MAX1_VAL = -2 ^ 52: MAX2_VAL = -2 ^ 52
    MIN1_VAL = 2 ^ 52: MIN2_VAL = 2 ^ 52
    
'------------------------------------------------------------------------------
    If (i - 1 - MAX_MIN_PERIODS) >= 1 Then
'------------------------------------------------------------------------------
        If (i + 1 + MAX_MIN_PERIODS) <= NROWS Then
            For j = (i - 1 - MAX_MIN_PERIODS) To (i + 1 + MAX_MIN_PERIODS)
                If DATA1_MATRIX(j, 2) > MAX1_VAL Then: MAX1_VAL = DATA1_MATRIX(j, 2)
                If DATA1_MATRIX(j, 2) < MIN1_VAL Then: MIN1_VAL = DATA1_MATRIX(j, 2)
                
                If DATA2_MATRIX(j, 2) > MAX2_VAL Then: MAX2_VAL = DATA2_MATRIX(j, 2)
                If DATA2_MATRIX(j, 2) < MIN2_VAL Then: MIN2_VAL = DATA2_MATRIX(j, 2)
            Next j
'------------------------------------------------------------------------------
        Else
'------------------------------------------------------------------------------
            For j = (i - 1 - MAX_MIN_PERIODS) To NROWS
                If DATA1_MATRIX(j, 2) > MAX1_VAL Then: MAX1_VAL = DATA1_MATRIX(j, 2)
                If DATA1_MATRIX(j, 2) < MIN1_VAL Then: MIN1_VAL = DATA1_MATRIX(j, 2)
                
                If DATA2_MATRIX(j, 2) > MAX2_VAL Then: MAX2_VAL = DATA2_MATRIX(j, 2)
                If DATA2_MATRIX(j, 2) < MIN2_VAL Then: MIN2_VAL = DATA2_MATRIX(j, 2)
            Next j
'------------------------------------------------------------------------------
        End If
'------------------------------------------------------------------------------
    Else
'------------------------------------------------------------------------------
        If (i + 1 + MAX_MIN_PERIODS) <= NROWS Then
'------------------------------------------------------------------------------
            For j = 1 To (i + 1 + MAX_MIN_PERIODS)
                If DATA1_MATRIX(j, 2) > MAX1_VAL Then: MAX1_VAL = DATA1_MATRIX(j, 2)
                If DATA1_MATRIX(j, 2) < MIN1_VAL Then: MIN1_VAL = DATA1_MATRIX(j, 2)
                
                If DATA2_MATRIX(j, 2) > MAX2_VAL Then: MAX2_VAL = DATA2_MATRIX(j, 2)
                If DATA2_MATRIX(j, 2) < MIN2_VAL Then: MIN2_VAL = DATA2_MATRIX(j, 2)
            Next j
'------------------------------------------------------------------------------
        Else
'------------------------------------------------------------------------------
            For j = 1 To NROWS
                If DATA1_MATRIX(j, 2) > MAX1_VAL Then: MAX1_VAL = DATA1_MATRIX(j, 2)
                If DATA1_MATRIX(j, 2) < MIN1_VAL Then: MIN1_VAL = DATA1_MATRIX(j, 2)
                
                If DATA2_MATRIX(j, 2) > MAX2_VAL Then: MAX2_VAL = DATA2_MATRIX(j, 2)
                If DATA2_MATRIX(j, 2) < MIN2_VAL Then: MIN2_VAL = DATA2_MATRIX(j, 2)
            Next j
'------------------------------------------------------------------------------
        End If
'------------------------------------------------------------------------------
    End If
'------------------------------------------------------------------------------
    TEMP_MATRIX(i, 4) = IIf(DATA1_MATRIX(i, 2) = MAX1_VAL, DATA1_MATRIX(i, 2), "")
    TEMP_MATRIX(i, 5) = IIf(DATA2_MATRIX(i, 2) = MAX2_VAL, DATA2_MATRIX(i, 2), "")
    TEMP_MATRIX(i, 6) = IIf(DATA1_MATRIX(i, 2) = MIN1_VAL, DATA1_MATRIX(i, 2), "")
    TEMP_MATRIX(i, 7) = IIf(DATA2_MATRIX(i, 2) = MIN2_VAL, DATA2_MATRIX(i, 2), "")
'------------------------------------------------------------------------------
Next i
'------------------------------------------------------------------------------

ASSETS_SISTER_STOCK_PREDICTORS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_SISTER_STOCK_PREDICTORS_FUNC = Err.number
End Function
