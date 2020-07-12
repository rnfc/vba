Attribute VB_Name = "FINAN_ASSET_PAIR_DIREC_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function ASSETS_PAIR_DIRECTION_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date)

'You type in a couple of Yahoo stock symbols, pick an End Date then magic.
'the daily returns of the first with the daily returns of the second one day later.
'So you can see if the first influences the second, right? Something like that, and
'If you try it with a recent End Date and do it ag'in with an End Date a year ago,
'you often find a significant difference.

Dim k As Long
Dim i As Long
Dim j As Long
Dim h As Long

Dim jj As Long
Dim ii As Long
Dim kk As Long

Dim NCOLUMNS As Long

Dim DATA1_VAL As Double
Dim DATA2_VAL As Double

Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

TICKERS_VECTOR = TICKERS_RNG
If UBound(TICKERS_VECTOR, 1) = 1 Then
    TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
End If
jj = UBound(TICKERS_VECTOR, 1)

ReDim DATA_GROUP(1 To jj)
For j = 1 To jj: DATA_GROUP(j) = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKERS_VECTOR(j, 1), _
START_DATE, END_DATE, "DAILY", "OC", False, True, True): Next j

kk = jj * (jj - 1) / 2
NCOLUMNS = 7
ReDim TEMP_MATRIX(0 To kk, 1 To NCOLUMNS)
TEMP_MATRIX(0, 1) = "TICKER1"
TEMP_MATRIX(0, 2) = "TICKER2"
TEMP_MATRIX(0, 3) = "PREVIOUS DAY"
TEMP_MATRIX(0, 4) = "SAME DAY"
TEMP_MATRIX(0, 5) = "NEXT DAY"
TEMP_MATRIX(0, 6) = "OPEN"
TEMP_MATRIX(0, 7) = "UP/DOWN"

h = 1
For j = 1 To jj
    For i = j + 1 To jj
        TEMP_MATRIX(h, 1) = TICKERS_VECTOR(j, 1)
        TEMP_MATRIX(h, 2) = TICKERS_VECTOR(i, 1)
        If IsArray(DATA_GROUP(j)) = False Or IsArray(DATA_GROUP(i)) = False Then: GoTo 1983
        ii = UBound(DATA_GROUP(j), 1)
        If ii <> UBound(DATA_GROUP(i), 1) Then: GoTo 1983
        For k = 3 To NCOLUMNS: TEMP_MATRIX(h, k) = 0: Next k
        For k = 1 To ii
            If k > 2 Then
                DATA1_VAL = DATA_GROUP(j)(k, 2) / DATA_GROUP(j)(k - 1, 2) - 1
                DATA2_VAL = DATA_GROUP(i)(k - 1, 2) / DATA_GROUP(i)(k - 2, 2) - 1
                If (DATA1_VAL * DATA2_VAL) > 0 Then: TEMP_MATRIX(h, 3) = TEMP_MATRIX(h, 3) + 1
                'Daily changes in TICKER1 are in the same direction as the previous day's changes in TICKER2, x% of the time.
            End If
            If k > 1 Then
                DATA1_VAL = DATA_GROUP(j)(k, 2) / DATA_GROUP(j)(k - 1, 2) - 1
                DATA2_VAL = DATA_GROUP(i)(k, 2) / DATA_GROUP(i)(k - 1, 2) - 1
                If (DATA1_VAL * DATA2_VAL) > 0 Then: TEMP_MATRIX(h, 4) = TEMP_MATRIX(h, 4) + 1
                'Daily changes in TICKER1 are in the same direction as the same day's changes in TICKER2, x% of the time.
            End If
            If k > 1 And k < ii Then
                DATA1_VAL = DATA_GROUP(j)(k, 2) / DATA_GROUP(j)(k - 1, 2) - 1
                DATA2_VAL = DATA_GROUP(i)(k + 1, 2) / DATA_GROUP(i)(k, 2) - 1
                If (DATA1_VAL * DATA2_VAL) > 0 Then: TEMP_MATRIX(h, 5) = TEMP_MATRIX(h, 5) + 1
                'Daily changes in TICKER1 are in the same direction as the subsequent day's changes in TICKER2, x% of the time.
            End If
            DATA1_VAL = DATA_GROUP(j)(k, 2) / DATA_GROUP(j)(k, 1) - 1
            DATA2_VAL = DATA_GROUP(i)(k, 2) / DATA_GROUP(i)(k, 1) - 1
            If (DATA1_VAL * DATA2_VAL) > 0 Then: TEMP_MATRIX(h, 6) = TEMP_MATRIX(h, 6) + 1
            'The Day's Open-to-Close change in TICKER2 is in the same direction as the Opening change in TICKER1, x% of the time.
        
            If k > 1 Then
                If DATA_GROUP(j)(k, 1) < DATA_GROUP(j)(k - 1, 2) And DATA_GROUP(i)(k, 1) < DATA_GROUP(i)(k, 2) Then
                    TEMP_MATRIX(h, 7) = TEMP_MATRIX(h, 7) + 1
                End If
                'When TICKER1 opens DOWN, then TICKER2 will Close UP during that day, x% of the time.
            End If
        
'Portfolio: Buy TICKER2 at the Open and sell at the Close, whenever TICKER1 Opens DOWN.
        
        Next k
        TEMP_MATRIX(h, 3) = TEMP_MATRIX(h, 3) / (ii - 2)
        TEMP_MATRIX(h, 4) = TEMP_MATRIX(h, 4) / (ii - 1)
        TEMP_MATRIX(h, 5) = TEMP_MATRIX(h, 5) / (ii - 2)
        TEMP_MATRIX(h, 6) = TEMP_MATRIX(h, 6) / (ii - 0)
        TEMP_MATRIX(h, 7) = TEMP_MATRIX(h, 7) / (ii - 1)
        h = h + 1
    Next i
1983:
Next j
Erase DATA_GROUP
ASSETS_PAIR_DIRECTION_FUNC = TEMP_MATRIX

'------------------------------------------------------------------------------------------------
Exit Function
ERROR_LABEL:
ASSETS_PAIR_DIRECTION_FUNC = Err.number
End Function

'Some time ago I played with so-called sister stocks.
'Such pairs tend to move in the same direction
'... either UP or DOWN.

'What about pairs that move in the opposite direction?
'How about this pair?


'Interesting, eh?
'It raises some very important questions:
'[1] Where do we find such pairs?
'[2] Can we use them to make a $killing?

'... and most importantly:
'[3] What should we call 'em?

'http://www.google.com/search?source=ig&hl=en&rlz=1G1GGLQ_ENCA292&=&q=%22sister+stocks%22

'Remember when we talked about Darvas boxes?
'It was a scheme that relied upon the occurrence of some event or situation.
'When it happened, you made your move and X% of the time (historically speaking), you'd make money.
'If X% were in the neighbourhood of 50%, it ain't a good scheme.
'(Half the time you lose, eh?)
'However, if x% > 60% then (eventually!) you'd make money
'... or so the theory goes.

'Anyway, there seem to be lots of such magic events lying around.
'We talked about them what-are-they-called stocks,
'It was pretty easy to find a pair of stocks, A and B such that:

'On days when A opened DOWN, B would go UP that day
'--- most of the time.

'That suggests a trading strategy for certain magic pairs::
'Buy B at the Open and sell at the Close
'... whenever A opened DOWN.

'here 's another, simpler scheme based upon the fact that (historically speaking), certain magic stocks (much of the time) increase from Open to Close, even when they open DOWN. That was true of several DOW stocks, over the past year.

'For example, the IBM opened DOWN 150 times (over the past year).
'In 87 of those times, it nevertheless increased from Open to Close.
'that 's almost 60%, right?
'So here 's the strategy (for that magic stock):
'Buy at the Open and sell at the Close
'... whenever the stock opened DOWN.

'Alas, that 60% was just 46% the year before.
'http://ponzoblog.blogspot.com/search/label/Darvas%20Boxes

