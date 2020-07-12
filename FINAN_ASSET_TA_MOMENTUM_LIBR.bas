Attribute VB_Name = "FINAN_ASSET_TA_MOMENTUM_LIBR"

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------


'Opening Momentum
'http://www.gummy-stuff.org/opening-momentum.htm

Function ASSETS_OPENING_MOMENTUM_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long
Dim TICKER_STR As String

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant
Dim DATA_MATRIX As Variant

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

NCOLUMNS = UBound(TICKERS_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NCOLUMNS, 1 To 6)
TEMP_MATRIX(0, 1) = "TICKER"
TEMP_MATRIX(0, 2) = "START DATE"
TEMP_MATRIX(0, 3) = "END DATE"
TEMP_MATRIX(0, 4) = "NOBS"
TEMP_MATRIX(0, 5) = "INTERCEPT"
TEMP_MATRIX(0, 6) = "SLOPE"

For j = 1 To NCOLUMNS
    TICKER_STR = TICKERS_VECTOR(j, 1)
    TEMP_MATRIX(j, 1) = TICKER_STR
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DOHLC", False, False, True)
    If IsArray(DATA_MATRIX) = False Then: GoTo 1983
    NROWS = UBound(DATA_MATRIX, 1)
    If NROWS <= 2 Then: GoTo 1983
    TEMP_MATRIX(j, 2) = DATA_MATRIX(1, 1)
    TEMP_MATRIX(j, 3) = DATA_MATRIX(NROWS, 1)
    TEMP_MATRIX(j, 4) = NROWS
    ReDim XDATA_VECTOR(1 To NROWS - 1, 1 To 1)
    ReDim YDATA_VECTOR(1 To NROWS - 1, 1 To 1)
    For i = 2 To NROWS
        'Close to next-day-Open
        XDATA_VECTOR(i - 1, 1) = DATA_MATRIX(i, 2) / DATA_MATRIX(i - 1, 5) - 1
        'Close to next-day-High.
        YDATA_VECTOR(i - 1, 1) = DATA_MATRIX(i, 3) / DATA_MATRIX(i - 1, 5) - 1
    Next i
    If OUTPUT = j Then
        ASSETS_OPENING_MOMENTUM_FUNC = REGRESSION_SIMPLE_COEF_FUNC(XDATA_VECTOR, YDATA_VECTOR)
        Exit Function
    End If
'The High for the day is significantly higher than the Open??
'On those days when the Open is higher than yesterday's Close --> only those days when the stock opens UP.
'Then I decide to plot the Close to next-day-High vs the Close to next-day-Open.
'On those days when the stock opens UP?
    DATA_MATRIX = REGRESSION_SIMPLE_COEF_FUNC(XDATA_VECTOR, YDATA_VECTOR)
    If IsArray(DATA_MATRIX) = False Then: GoTo 1983
'What happens to the High when the stock Opens UP or DOWN.
    TEMP_MATRIX(j, 5) = DATA_MATRIX(2, 1)
    TEMP_MATRIX(j, 6) = DATA_MATRIX(1, 1)
1983:
Next j
ASSETS_OPENING_MOMENTUM_FUNC = TEMP_MATRIX

'------------------------------------------------------------------------------------------------------
Exit Function
'------------------------------------------------------------------------------------------------------
ERROR_LABEL:
ASSETS_OPENING_MOMENTUM_FUNC = Err.number
End Function

'HERE 's something that's intrigued me for some time:
'At stock opens UP by, say 2%. (That's like a 200 point gain in the DOW.)
'Do investors think: "WOW! I better sell to lock in my gains" (That'll drive the price down, eh?)
'Or do they think: "This stock is on a roll -- I better buy some!" (That'll drive the price up, eh?)

'So, what happens after the Open?
'that 's the question I asked myself. I've asked it before, http://www.gummy-stuff.org/open-to-high.htm

'I'm convinced that there's something, call it Opening Momentum, that'll drive the stock higher
'during the day. In particular, I wondered about my coal stocks, like Grande Cache
'http://finance.yahoo.com/q?s=GCE.TO
'what 's intriguing is trying to (somehow) display this "momentum".
