Attribute VB_Name = "FINAN_ASSET_TA_DONCHIAN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_DONCHIAN_CHANNEL_FUNC

'DESCRIPTION   : Buy when Close rises above yesterday's Donchian-Hi
'Sell when Close falls below yesterday's Donchian-Lo

'It's a trading system, originally designed by Richard Donchian (1905-1993).

'* You pick a nice number like 20 and look at the highest high and lowest low over the past 20 days.
'* You get very excited when the current stock price rises above the highest high, anticipating an
'  uptrend in the price.
'* You also expect a downtrend when the price falls below the lowest low over the past 20 days.
'* When the ...

'It shows the highest high and lowest low as well as the closing price over the past 20 days.
'Note that yesterday's upper Donchian is usually larger than today's closing price, but when
'the close rises above the upper Donchian ...

'That's regarded as the signal for the start of an uptrend, right? Apparently.

'Donchian channel for GE stock. Note that you can change the number of days in each of the upper-
'and lower-Donchian. Who knows? Maybe the Uptrend and Downtrend signals are better with
'different days.

'Reference:
'http://en.wikipedia.org/wiki/Richard_Donchian
'http://www.gummy-stuff.org/Donchian.htm

'LIBRARY       : FINAN_ASSET
'GROUP         : TA_DONCHIAN
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA

'************************************************************************************
'************************************************************************************

Function ASSET_TA_DONCHIAN_CHANNEL_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal HIGH_CHANNEL_PERIOD As Long = 20, _
Optional ByVal LOW_CHANNEL_PERIOD As Long = 15)

'Days in Hi-Channel = 20
'Days in Lo-Channel = 15
'Buy when Close rises above yesterday's Donchian-Hi
'Sell when Close falls below yesterday's Donchian-Lo

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim MIN_VAL As Double
Dim MAX_VAL As Double
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, _
                  START_DATE, END_DATE, "DAILY", "DOHLCVA", False, _
                  False, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 12)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME/1000"
TEMP_MATRIX(0, 7) = "ADJ_CLOSE"
TEMP_MATRIX(0, 8) = "RETURNS"

TEMP_MATRIX(0, 9) = "DON-HIGH"
TEMP_MATRIX(0, 10) = "DON-LOW"

TEMP_MATRIX(0, 11) = "UP TREND: " & Format(HIGH_CHANNEL_PERIOD, "0") & " PERIOD"
TEMP_MATRIX(0, 12) = "LOW TREND: " & Format(LOW_CHANNEL_PERIOD, "0") & " PERIOD"

i = 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 2) - 1

MAX_VAL = TEMP_MATRIX(i, 3)
MIN_VAL = TEMP_MATRIX(i, 4)

TEMP_MATRIX(i, 9) = MAX_VAL
TEMP_MATRIX(i, 10) = MIN_VAL
TEMP_MATRIX(i, 11) = ""
TEMP_MATRIX(i, 12) = ""

'---------------------------------------------------------------------------------
For i = 2 To NROWS
'---------------------------------------------------------------------------------
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 8) = DATA_MATRIX(i, 7) / DATA_MATRIX(i - 1, 7) - 1
    
    If i <= HIGH_CHANNEL_PERIOD Then
        If TEMP_MATRIX(i, 3) > MAX_VAL Then: MAX_VAL = TEMP_MATRIX(i, 3)
    Else
        MAX_VAL = TEMP_MATRIX(i, 3)
        For j = i To (i - HIGH_CHANNEL_PERIOD) Step -1
            If TEMP_MATRIX(j, 3) >= MAX_VAL Then: MAX_VAL = TEMP_MATRIX(j, 3)
        Next j
    End If
    
    If i <= LOW_CHANNEL_PERIOD Then
        If TEMP_MATRIX(i, 4) < MIN_VAL Then: MIN_VAL = TEMP_MATRIX(i, 4)
    Else
        MIN_VAL = TEMP_MATRIX(i, 4)
        For j = i To (i - LOW_CHANNEL_PERIOD) Step -1
            If TEMP_MATRIX(j, 4) <= MIN_VAL Then: MIN_VAL = TEMP_MATRIX(j, 4)
        Next j
    End If
    TEMP_MATRIX(i, 9) = MAX_VAL
    TEMP_MATRIX(i, 10) = MIN_VAL
    
    TEMP_MATRIX(i, 11) = IIf(TEMP_MATRIX(i, 5) > TEMP_MATRIX(i - 1, 9), TEMP_MATRIX(i, 5), "")
    TEMP_MATRIX(i, 12) = IIf(TEMP_MATRIX(i, 5) < TEMP_MATRIX(i - 1, 10), TEMP_MATRIX(i, 5), "")
'---------------------------------------------------------------------------------
Next i
'---------------------------------------------------------------------------------

ASSET_TA_DONCHIAN_CHANNEL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_DONCHIAN_CHANNEL_FUNC = Err.number
End Function
