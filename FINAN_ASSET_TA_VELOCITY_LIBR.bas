Attribute VB_Name = "FINAN_ASSET_TA_VELOCITY_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    

'References:
'http://www.gummy-stuff.org/velocity-of-trade.htm
'http://www.investopedia.com/terms/t/theta.asp

Function ASSET_TA_VELOCITY_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal TARGET_PRICE As Double = 60)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim FIRST_PRICE As Double

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

'-----------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 11)
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"
TEMP_MATRIX(0, 8) = "PRICE TRAVEL: TP = " & Format(TARGET_PRICE, "0.0")
TEMP_MATRIX(0, 9) = "TRADE VELOCITY"
TEMP_MATRIX(0, 10) = "UPS"
TEMP_MATRIX(0, 11) = "DOWNS"

i = 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
FIRST_PRICE = TEMP_MATRIX(i, 5)

TEMP_MATRIX(i, 8) = (TEMP_MATRIX(i, 5) - FIRST_PRICE) / (TARGET_PRICE - FIRST_PRICE)
For j = 9 To 11: TEMP_MATRIX(i, j) = "": Next j

MIN_VAL = 2 ^ 52: MAX_VAL = -2 ^ 52
For i = 2 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 8) = (TEMP_MATRIX(i, 5) - FIRST_PRICE) / (TARGET_PRICE - FIRST_PRICE)
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8) - TEMP_MATRIX(i - 1, 8)
Next i

For i = NROWS To 2 Step -1
    If TEMP_MATRIX(i, 9) > MAX_VAL Then: MAX_VAL = TEMP_MATRIX(i, 9)
    TEMP_MATRIX(i, 10) = MAX_VAL
    
    If TEMP_MATRIX(i, 9) < MIN_VAL Then: MIN_VAL = TEMP_MATRIX(i, 9)
    TEMP_MATRIX(i, 11) = MIN_VAL
Next i

ASSET_TA_VELOCITY_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_VELOCITY_FUNC = Err.number
End Function

'Velocity of Trade
'Suppose we have some stock or option or mutual fund and we want to see how quickly it
'achieves some target price.

'Assume the current price of the asset is Po.   Example: Po = $57.60
'we 'd like to (eventually) reach a target price of TP.   Example: TP = $60.00
'The price would then have to travel a distance of TP - Po.   Example: TP - Po = $2.40
'Suppose that, after t days, the price is P(t).   Example: P(3) = 58.45
'Then the price has travelled a distance of P(t) - Po.   Example: P(3) - Po = $0.85
'As a fraction of the distance required to achieve our target price, that's
'X(t) = (P(t) -Po) / (TP-Po)   Example: X(3) = (0.85) / (2.40) = 0.35 or 35%

'Now , we 'd like to see how rapidly our asset price is moving, in the direction of the target price, so ...
'So that's the velocity: V = dX/dt, right?
'Well, our t-values don't change continously. They're t = 0, 1, 2, etc., so we consider:
'V(t) = (X(t) - X(t-1) ) / (t - (t-1)) = X(t) - X(t-1).

'Now the graph of X(t) = (P(t) -Po) / (TP-Po) looks just like the price graph P(t) (except for some
'rescaling) as shown in Figure 1. So V(t) = X(t) - X(t-1) is just (P(t) -P(t-1)) / (TP-Po) is just
'the velocity associated with the price-graph (except for some rescaling).

'Upper velocity and lower velocity?
'Them 's just some level lines to see how you're doin'.
'So how come that velocity is sometimes negative, sometimes positive?
'When the price increases the velocity is positive and when it decreases the velocity is negative.
'if we were talking options, then this "velocity" would be related to theta.

