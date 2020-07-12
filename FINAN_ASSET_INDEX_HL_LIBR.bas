Attribute VB_Name = "FINAN_ASSET_INDEX_HL_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : INDEX_HIGH_LOW_OSCILLATOR_FUNC
'DESCRIPTION   :
'There 's this thing called (something like): new Highs Lows Oscillator.
'You consider a basket of stocks (like the S&P500 collection) and do this:
'You count the number of stocks where (Today's Price) = (52-week High).
'You count the number of stocks where (Today's Price) = (52-week Low).
'You subtract the latter from the former.
'High/Low Oscillator = (Number of new 52-week Highs) - (Number of new 52-week Lows)

'Each day you repeat steps 1 to 3 so you get a chart of the variation in this number.
'That's the "oscillator".

'If you plot these numbers for the last 30 days or so, you'll get a graph that ...
'uh ... oscillates. Unfortunately , it 's often difficult (expecially these days) to
'find stocks that are at their 52-week High. Indeed , it 's likely that there are NONE.
'In fact, it's possible that ALL stocks are at their 52-week low.

'That makes the High/Low Oscillator sorta useless.
'so i 'd like to fiddle with that to see how many stocks have their current price within
'X% of the 52-week High. For example, if X% = 2%, then we'd look for stocks where (Today's
'Price) = 0.98*(52-week High)

'I'd also like to fiddle to see how many stocks have their current price within X% of the
'52-week Low. For example, if X% = 3%, then we'd look for stocks where (Today's Price) =
'1.03*(52-week Low). And you can play with X%, right?

'Right ... so we can get some idea of how many stocks are "in the neighbourhood" of their
'Highs and Lows (where X defines the neighbourhood). Further, suppose I told you that there
'were 5 more stocks "in the neighbourhood" of their 52-week High then were near their 52-week Low.
'In other words: HIGH_VAL/LOW_VAL Oscillator = (Number of new 52-week Highs) - (Number of new
'52-week Lows) = 5.

'Would you say that was a large number?
'Is 5 a large number? I wouldn't think so.

'But if I had just 20 stocks in my basket of stocks, then 5 would be pretty large, but if ...
'But if you had 500 stocks, then it'd be peanuts, eh?

'Exactly, so I'd like to fiddle with the numbers so that the number generated is a percentage
'of the number of stocks in the basket. To that end, we define:

'For a basket of NSIZE stocks:
'gHL(X) Oscillator = [ (Number of stocks within X% of their 52-week High) -
'(Number of stocks within X% of their 52-week Low) ] / NSIZE

'Then if I say gHL = 12%, it'd mean that the percentage of stocks in the basket which were near
'their High is 12% higher than the percentage near their Low.

'But gHL could be negative, eh? Yes, indeed.

'In this function you need a gaggle of stock (up to 500). Then a bunch of data is downloaded and
'the daily gHL is calculated for the past month (roughly). In fact, gHL(X) is calculated for
'X = 0%, 1% and 2%.

'Remember gHL stand for? --> great High Low.
'(Pct. within X% of High) - (Pct. within X% of Low)

'LIBRARY       : FINAN_ASSET
'GROUP         : INDEX
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function INDEX_HIGH_LOW_OSCILLATOR_FUNC(ByVal INDEX_STR As String, _
Optional ByVal NO_DAYS As Long = 410, _
Optional ByVal PERCENT As Double = 0.03, _
Optional ByVal WINDOWS_FACTOR As Long = 30, _
Optional ByVal REFRESH_CALLER As Variant, _
Optional ByVal SERVER_STR As String = "UNITED STATES")

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long

Dim NSIZE As Long
Dim NROWS As Long

Dim LOW_VAL As Long
Dim HIGH_VAL As Long

Dim END_DATE As Date
Dim START_DATE As Date

Dim TEMP_STR As String

Dim MIN_VAL As Double
Dim MAX_VAL As Double
Dim TEMP_VAL As Double

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim DATA_MATRIX(1 To 1, 1 To 4)
DATA_MATRIX(1, 1) = "Symbol"
DATA_MATRIX(1, 2) = "Last Trade"
DATA_MATRIX(1, 3) = "52-week high"
DATA_MATRIX(1, 4) = "52-week low"

DATA_MATRIX = YAHOO_INDEX_QUOTES_FUNC(INDEX_STR, DATA_MATRIX, "", _
              False, REFRESH_CALLER, SERVER_STR)
If IsArray(DATA_MATRIX) = False Then: GoTo 1983
'    INDEX_HIGH_LOW_OSCILLATOR_FUNC = DATA_MATRIX
'    Exit Function

NSIZE = UBound(DATA_MATRIX, 1)
m = NSIZE

END_DATE = Now
START_DATE = DateSerial(Year(END_DATE), _
             Month(END_DATE), Day(END_DATE) - NO_DAYS)

HIGH_VAL = 0
LOW_VAL = 0

ReDim TEMP_MATRIX(0 To WINDOWS_FACTOR, 1 To 2)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PERCENT"

h = 4
For i = 1 To NSIZE
    TEMP_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(DATA_MATRIX(i, 2), START_DATE, END_DATE, "d", "DOHLCVA", False, False, True)
    If IsArray(TEMP_VECTOR) = False Then
        m = m - 1
        GoTo 1983
    End If
    NROWS = UBound(TEMP_VECTOR, 1)
    
    MIN_VAL = 2 ^ 52: MAX_VAL = 2 ^ -52
    For k = WINDOWS_FACTOR + h To NROWS
        If TEMP_VECTOR(k, 7) > MAX_VAL Then: MAX_VAL = TEMP_VECTOR(k, 7)
        If TEMP_VECTOR(k, 7) < MIN_VAL Then: MIN_VAL = TEMP_VECTOR(k, 7)
    Next k
    
    HIGH_VAL = HIGH_VAL + IIf(TEMP_VECTOR(NROWS, 7) = MAX_VAL, 1, 0)
    LOW_VAL = LOW_VAL + IIf(TEMP_VECTOR(NROWS, 7) = MIN_VAL, 1, 0)
    
    For k = 1 To WINDOWS_FACTOR
        l = NROWS - WINDOWS_FACTOR + k
        If i = 1 Then: TEMP_MATRIX(k, 1) = TEMP_VECTOR(l, 1) 'Dates
        MIN_VAL = 2 ^ 52: MAX_VAL = 2 ^ -52
        For j = h + k To l
            If TEMP_VECTOR(j, 7) > MAX_VAL Then: MAX_VAL = TEMP_VECTOR(j, 7)
            If TEMP_VECTOR(j, 7) < MIN_VAL Then: MIN_VAL = TEMP_VECTOR(j, 7)
        Next j
        TEMP_VAL = 0
        If TEMP_VECTOR(l, 7) >= (1 - PERCENT) * MAX_VAL Then
            TEMP_VAL = 1
        Else
            If TEMP_VECTOR(l, 7) <= (1 + PERCENT) * MIN_VAL Then
                TEMP_VAL = -1
            Else
                TEMP_VAL = 0
            End If
        End If
        TEMP_MATRIX(k, 2) = TEMP_MATRIX(k, 2) + TEMP_VAL / m
    Next k

1983:
Next i

TEMP_STR = "No. at High = " & Format(HIGH_VAL, "0") & ", No. at Low = " & Format(LOW_VAL, "0")

INDEX_HIGH_LOW_OSCILLATOR_FUNC = Array(TEMP_MATRIX, TEMP_STR)

Exit Function
ERROR_LABEL:
INDEX_HIGH_LOW_OSCILLATOR_FUNC = Err.number
End Function
