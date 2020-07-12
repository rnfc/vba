Attribute VB_Name = "FINAN_ASSET_TA_OC_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


Function ASSETS_TA_OC_PROB_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date)

Dim i As Long
Dim j As Long

Dim k As Long
Dim kk As Long

Dim l As Long
Dim ll As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If
NCOLUMNS = UBound(TICKERS_VECTOR)

ReDim TEMP_VECTOR(0 To NCOLUMNS, 1 To 5)
TEMP_VECTOR(0, 1) = "SYMBOL"
TEMP_VECTOR(0, 2) = "(O>pC And C<O)"
TEMP_VECTOR(0, 3) = "(O<pC And C>O)"
TEMP_VECTOR(0, 4) = "(O>pC And C>O)"
TEMP_VECTOR(0, 5) = "(O<pC And C<O)"

For j = 1 To NCOLUMNS
    k = 0: kk = 0
    l = 0: ll = 0
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKERS_VECTOR(j, 1), START_DATE, END_DATE, "DAILY", "DOHLCVA", False, False, True)
    If IsArray(DATA_MATRIX) = False Then: GoTo 1983
    NROWS = UBound(DATA_MATRIX, 1)
    For i = 2 To NROWS
        'Open higher than Previous Close, but closes down, or Opens lower than previous close, but closes up
        If ((DATA_MATRIX(i, 2) > DATA_MATRIX(i - 1, 5) And DATA_MATRIX(i, 5) < DATA_MATRIX(i, 2))) Then
             k = k + 1
        End If
        If ((DATA_MATRIX(i, 2) < DATA_MATRIX(i - 1, 5) And DATA_MATRIX(i, 5) > DATA_MATRIX(i, 2))) Then
             kk = kk + 1
        End If
        'Open higher than previous close, then closes up, or Opens lower than previous close, then closes down.
        If ((DATA_MATRIX(i, 2) > DATA_MATRIX(i - 1, 5) And DATA_MATRIX(i, 5) > DATA_MATRIX(i, 2))) Then
             l = l + 1
        End If
        If ((DATA_MATRIX(i, 2) < DATA_MATRIX(i - 1, 5) And DATA_MATRIX(i, 5) < DATA_MATRIX(i, 2))) Then
             ll = ll + 1
        End If
    Next i
    TEMP_VECTOR(j, 2) = k / NROWS
    TEMP_VECTOR(j, 3) = kk / NROWS
    TEMP_VECTOR(j, 4) = l / NROWS
    TEMP_VECTOR(j, 5) = ll / NROWS
1983:
    TEMP_VECTOR(j, 1) = TICKERS_VECTOR(j, 1)
Next j

ASSETS_TA_OC_PROB_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
ASSETS_TA_OC_PROB_FUNC = Err.number
End Function

'The other day I was staring at the 5-day chart of a stock and noticed that,
'when today's Open is higher than yesterday's Close, it often dropped, ending
'the day down from the Open.

'There were a couple of days when Open was greater than previous day's Close and
'Close was less than the Open.

'What's pOpen and ...Them 's the previous day's Open and Close.
    
'Now, if the daily values were completely random, and yesterday's values didn't influence
'today's values, you'd expect to have:
'Open > pClose AND Open > Close
'occur about 25% of the time since there are two inequalities and they could be:
'> > or > < or < > or < <.
'If we expect they're equally likely, then each should occcur about 25% of the time.

'Note the following:

'You stick some condition like AND(Open > pClose,Open > Close)
'For each stock, the percentage of times the condition is satisfied is noted.

'Looks like that 25% is bang on, eh? Uh ... yes, for that particular condition.
'However, you can type in other conditions, for example:
'Condition   the Meaning of the Condition
'AND(Open < pOpen, Close < pClose)   Today's Open is less than the previous Open AND
'today's Close is less than the previous Close

'OR(Open < Close, Close < pOpen) Today's Open is less than the today's Close OR today's
'Close is less than the previous Open

'So I can use four values for today, namely Open, High, Low and Close and I can use ...
'You can use the previous Open, High, Low and Close. You can also use AND or OR.

'Come to think of it, you can also use Volume and pVolume, just in case that's of interest.
'You might try, for example:

'Open*Volume > pOpen*pVolume AND Close > pClose.

'In fact, I tried this one ... thinking that momentum may be at work for some stocks:
'OR( AND(Close > Open, pClose > pOpen), AND(Close < Open, pClose < pOpen) )

'It looks for stocks where, if the price goes up (or down) one day, it goes up (or down)
'the following day.

