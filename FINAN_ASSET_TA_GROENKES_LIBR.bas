Attribute VB_Name = "FINAN_ASSET_TA_GROENKES_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    

'///////////////////////////////////////////////////////////////////////////////////////////////////
'Groenke's Visions Theory
'Reference: http://www.rongroenke.com/vision_v_theory.pdf
'///////////////////////////////////////////////////////////////////////////////////////////////////

Function GROENKES_VISIONS_SCREENER_FUNC(ByRef TICKERS_RNG As Variant, _
Optional ByVal REFRESH_CALLER As Variant, _
Optional ByVal SERVER_STR As String = "UNITED STATES", _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
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

ReDim TEMP_VECTOR(1 To 1, 1 To 6)

TEMP_VECTOR(1, 1) = "name"
TEMP_VECTOR(1, 2) = "time of last trade"
TEMP_VECTOR(1, 3) = "last trade"
TEMP_VECTOR(1, 4) = "50-day moving avg"
TEMP_VECTOR(1, 5) = "52-week high"
TEMP_VECTOR(1, 6) = "52-week low"

TEMP_VECTOR = YAHOO_QUOTES_FUNC(TICKERS_VECTOR, TEMP_VECTOR, REFRESH_CALLER, False, SERVER_STR)
NROWS = UBound(TEMP_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 13)

TEMP_MATRIX(0, 1) = "SYMBOL"
TEMP_MATRIX(0, 2) = "NAME"
TEMP_MATRIX(0, 3) = "TIME LAST TRADE"

TEMP_MATRIX(0, 4) = "PRICE LAST TRADE"
TEMP_MATRIX(0, 5) = "50 DAY MA"
TEMP_MATRIX(0, 6) = "52W LOW"
TEMP_MATRIX(0, 7) = "52W HIGH"
TEMP_MATRIX(0, 8) = "BUY LIMIT"
TEMP_MATRIX(0, 9) = "BUY RANK"
TEMP_MATRIX(0, 10) = "TAI"
TEMP_MATRIX(0, 11) = "TAI VALUE"
TEMP_MATRIX(0, 12) = "LOWER RANGE"
TEMP_MATRIX(0, 13) = "UPPER RANGE"

For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = TICKERS_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = TEMP_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = TEMP_VECTOR(i, 2)
    TEMP_MATRIX(i, 4) = TEMP_VECTOR(i, 3)
    If IsNumeric(TEMP_MATRIX(i, 4)) = False Then: GoTo 1983
    If TEMP_MATRIX(i, 4) = 0 Then: GoTo 1983
    TEMP_MATRIX(i, 5) = TEMP_VECTOR(i, 4)
    If TEMP_MATRIX(i, 5) = 0 Then: GoTo 1983
    TEMP_MATRIX(i, 6) = TEMP_VECTOR(i, 6)
    TEMP_MATRIX(i, 7) = TEMP_VECTOR(i, 5)
    TEMP_MATRIX(i, 8) = V_BUY_LIMIT_FUNC(TEMP_MATRIX(i, 6), TEMP_MATRIX(i, 7), VERSION)
    TEMP_MATRIX(i, 9) = V_BUY_RANK_FUNC(TEMP_MATRIX(i, 8), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 6), TEMP_MATRIX(i, 7), VERSION)
    If VERSION = 0 Then
        TEMP_MATRIX(i, 10) = V_TAI2_FUNC(TEMP_MATRIX(i, 8), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 6), TEMP_MATRIX(i, 7), TEMP_MATRIX(i, 5), 1)
        TEMP_MATRIX(i, 11) = V_TAI2_FUNC(TEMP_MATRIX(i, 8), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 6), TEMP_MATRIX(i, 7), TEMP_MATRIX(i, 5), 0)
    Else
        TEMP_MATRIX(i, 10) = V_TAI1_FUNC(TEMP_MATRIX(i, 9), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 5), 1)
        TEMP_MATRIX(i, 11) = V_TAI1_FUNC(TEMP_MATRIX(i, 9), TEMP_MATRIX(i, 4), TEMP_MATRIX(i, 5), 0)
    End If
    TEMP_MATRIX(i, 12) = V_LOWER_RANGE_FUNC(TEMP_MATRIX(i, 6), TEMP_MATRIX(i, 7))
    TEMP_MATRIX(i, 13) = V_UPPER_RANGE_FUNC(TEMP_MATRIX(i, 8), TEMP_MATRIX(i, 6), TEMP_MATRIX(i, 7))
1983:
Next i

GROENKES_VISIONS_SCREENER_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
GROENKES_VISIONS_SCREENER_FUNC = Err.number
End Function


'Buy Limit: maximum price to pay for the stock

Private Function V_BUY_LIMIT_FUNC(ByVal LOW_PRICE_VAL As Double, _
ByVal HIGH_PRICE_VAL As Double, _
Optional ByVal VERSION As Integer = 0)

' LOW_PRICE_VAL - 52 week low
' HIGH_PRICE_VAL - 52 week high

On Error GoTo ERROR_LABEL

If VERSION = 0 Then
    V_BUY_LIMIT_FUNC = LOW_PRICE_VAL + ((HIGH_PRICE_VAL - LOW_PRICE_VAL) * 0.25)
Else
    V_BUY_LIMIT_FUNC = (3 * LOW_PRICE_VAL + HIGH_PRICE_VAL) / 4
End If

Exit Function
ERROR_LABEL:
V_BUY_LIMIT_FUNC = Err.number
End Function

'stock buy rank

Private Function V_BUY_RANK_FUNC(ByVal BUY_LIMIT As Double, _
ByVal LAST_PRICE_VAL As Double, _
ByVal LOW_PRICE_VAL As Double, _
ByVal HIGH_PRICE_VAL As Double, _
Optional ByVal VERSION As Integer = 0)

' BUY_LIMIT - maximum price to pay for the stock
' LAST_PRICE_VAL - current stock price
' LOW_PRICE_VAL - 52 week LOW_PRICE_VAL
' HIGH_PRICE_VAL - 52 week HIGH_PRICE_VAL
                 
On Error GoTo ERROR_LABEL
        
If VERSION = 0 Then
    V_BUY_RANK_FUNC = (10 * (BUY_LIMIT - LAST_PRICE_VAL)) / ((HIGH_PRICE_VAL - LOW_PRICE_VAL) * 0.25)
Else
    V_BUY_RANK_FUNC = 40 * (BUY_LIMIT - LAST_PRICE_VAL) / (HIGH_PRICE_VAL - LOW_PRICE_VAL)
End If

Exit Function
ERROR_LABEL:
V_BUY_RANK_FUNC = Err.number
End Function

'take action indicator

Private Function V_TAI1_FUNC(ByVal BUY_RANK As Double, _
ByVal LAST_PRICE_VAL As Double, _
ByVal MA_VAL As Double, _
Optional ByVal OUTPUT As Integer = 1)

' BUY_RANK - stock buy rank
' LAST_PRICE_VAL - current stock price
' MA_VAL - 50 day moving average

On Error GoTo ERROR_LABEL

V_TAI1_FUNC = BUY_RANK * ((1 + (MA_VAL) / ((2 * MA_VAL) - LAST_PRICE_VAL)))
If OUTPUT = 0 Then Exit Function

If V_TAI1_FUNC >= 10 Then
    V_TAI1_FUNC = "2-GR" '2-Get Ready
Else
    If V_TAI1_FUNC > -5 Then
        V_TAI1_FUNC = "1-TA" '1-Time to Act
    Else
        If V_TAI1_FUNC > -10 Then
            V_TAI1_FUNC = "3-WT" '3-Wait
        Else
            V_TAI1_FUNC = "4-BI" '4-Bad Idea
        End If
    End If
End If

Exit Function
ERROR_LABEL:
V_TAI1_FUNC = Err.number
End Function


'take action indicator 2

Private Function V_TAI2_FUNC(ByVal BUY_LIMIT As Double, _
ByVal LAST_PRICE_VAL As Double, _
ByVal LOW_PRICE_VAL As Double, _
ByVal HIGH_PRICE_VAL As Double, _
ByVal MA_VAL As Double, _
Optional ByVal OUTPUT As Integer = 1)

' BUY_LIMIT - maximum price to pay for the stock
' LAST_PRICE_VAL - current stock price
' LOW_PRICE_VAL - 52 week LOW_PRICE_VAL
' HIGH_PRICE_VAL - 52 week HIGH_PRICE_VAL

On Error GoTo ERROR_LABEL

V_TAI2_FUNC = (40 * (((3 * LOW_PRICE_VAL + HIGH_PRICE_VAL) / 4) - LAST_PRICE_VAL) / (HIGH_PRICE_VAL - LOW_PRICE_VAL)) * (1 + MA_VAL / (2 * MA_VAL - LAST_PRICE_VAL))
If OUTPUT = 0 Then Exit Function

If V_TAI2_FUNC >= 10 Then
    V_TAI2_FUNC = "2-GR" '2-Get Ready
Else
    If V_TAI2_FUNC > -5 Then
        V_TAI2_FUNC = "1-TA" '1-Time to Act
    Else
        If V_TAI2_FUNC > -10 Then
            V_TAI2_FUNC = "3-WT" '3-Wait
        Else
            V_TAI2_FUNC = "4-BI" '4-Bad Idea
        End If
    End If
End If

Exit Function
ERROR_LABEL:
V_TAI2_FUNC = Err.number
End Function

'V Indicator Upper Leg

Private Function V_UPPER_RANGE_FUNC(ByVal BUY_LIMIT As Double, _
ByVal LOW_PRICE_VAL As Double, _
ByVal HIGH_PRICE_VAL As Double)

' BUY_LIMIT - maximum price to pay for the stock
' LOW_PRICE_VAL - 52 week LOW_PRICE_VAL
' HIGH_PRICE_VAL - 52 week HIGH_PRICE_VAL

On Error GoTo ERROR_LABEL

V_UPPER_RANGE_FUNC = BUY_LIMIT + ((HIGH_PRICE_VAL - LOW_PRICE_VAL) * 0.125)

Exit Function
ERROR_LABEL:
V_UPPER_RANGE_FUNC = Err.number
End Function

'V Indicator Lower Leg

Private Function V_LOWER_RANGE_FUNC(ByVal LOW_PRICE_VAL As Double, _
ByVal HIGH_PRICE_VAL As Double)

' LOW_PRICE_VAL - 52 week LOW_PRICE_VAL
' HIGH_PRICE_VAL - 52 week HIGH_PRICE_VAL

On Error GoTo ERROR_LABEL

V_LOWER_RANGE_FUNC = LOW_PRICE_VAL + ((HIGH_PRICE_VAL - LOW_PRICE_VAL) * 0.125)

Exit Function
ERROR_LABEL:
V_LOWER_RANGE_FUNC = Err.number
End Function

'backtest the Groenke Visions theory on an individual equity.
'Reference: http://www.rongroenke.com/

Function ASSETS_GROENKE_VISIONS_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal MA_DAYS As Long = 50, _
Optional ByVal DAYS_PER_YEAR As Long = 252, _
Optional ByVal tolerance As Double = 0.001, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long
Dim ll As Long

Dim ii_VAL As Double
Dim jj_VAL As Double
Dim kk_VAL As Double
Dim ll_VAL As Double

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim TEMP_SUM As Double

Dim TICKER_STR As String
Dim TICKERS_VECTOR As Variant

Dim TEMP_GROUP As Variant
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 2) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If
NSIZE = UBound(TICKERS_VECTOR, 2)

If OUTPUT > 0 Then
    ReDim TEMP_VECTOR(1 To 16, 1 To NSIZE + 1)
    TEMP_VECTOR(1, 1) = "SYMBOL"

    TEMP_VECTOR(2, 1) = "1-Time to Act {-5.0 <= TAI < 10.0}"
    TEMP_VECTOR(3, 1) = "2-Get Ready {10.0 <= TAI < 99.0}"
    TEMP_VECTOR(4, 1) = "3-Wait {-10.0 <= TAI < -5.0}"
    TEMP_VECTOR(5, 1) = "4-Bad Idea {-99.0 <= TAI < -10.0}"
    TEMP_VECTOR(6, 1) = "Grand Total"
    
    TEMP_VECTOR(7, 1) = "1-Time to Act Return"
    TEMP_VECTOR(8, 1) = "2-Get Ready Return"
    TEMP_VECTOR(9, 1) = "3-Wait Return"
    TEMP_VECTOR(10, 1) = "4-Bad Idea Return"
    TEMP_VECTOR(11, 1) = "Grand Total Return"
    
    TEMP_VECTOR(12, 1) = "1-Time to Act Average"
    TEMP_VECTOR(13, 1) = "2-Get Ready Average"
    TEMP_VECTOR(14, 1) = "3-Wait Average"
    TEMP_VECTOR(15, 1) = "4-Bad Idea Average"
    TEMP_VECTOR(16, 1) = "Grand Total Average"
End If

If NSIZE > 1 And OUTPUT <> 1 Then
    ReDim TEMP_GROUP(1 To NSIZE)
End If

For j = 1 To NSIZE
    TICKER_STR = TICKERS_VECTOR(1, j)
    TEMP_MATRIX = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR, Year(START_DATE), Month(START_DATE), Day(START_DATE), Year(END_DATE), Month(END_DATE), Day(END_DATE), "d", "DOHLCV", True, True, False, 0, 0)
    If IsArray(TEMP_MATRIX) = False Then: GoTo 1983
    NROWS = UBound(TEMP_MATRIX, 1)
    If NROWS < MA_DAYS Or NROWS < DAYS_PER_YEAR Then: GoTo 1983
    NCOLUMNS = UBound(TEMP_MATRIX, 2)
    
    ReDim Preserve TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS + 7)
    TEMP_MATRIX(0, 1) = TICKER_STR
    TEMP_MATRIX(0, NCOLUMNS + 1) = MA_DAYS & "-day SMA"
    TEMP_MATRIX(0, NCOLUMNS + 2) = DAYS_PER_YEAR & "-day High"
    TEMP_MATRIX(0, NCOLUMNS + 3) = DAYS_PER_YEAR & "-day Low"
    TEMP_MATRIX(0, NCOLUMNS + 4) = "NDO Factor"
    TEMP_MATRIX(0, NCOLUMNS + 5) = "TAI Value"
    TEMP_MATRIX(0, NCOLUMNS + 6) = "TAI Action"
    TEMP_MATRIX(0, NCOLUMNS + 7) = "Hi-Lo Dec"
    
    ii = 0: jj = 0: kk = 0: ll = 0
    ii_VAL = 1: jj_VAL = 1: kk_VAL = 1: ll_VAL = 1
    
    i = 1: TEMP_SUM = 0
    For k = 1 To MA_DAYS
        TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i + k - 1, 5)
    Next k
    TEMP_MATRIX(i, NCOLUMNS + 1) = TEMP_SUM / MA_DAYS
    
    MIN_VAL = 2 ^ 52: MAX_VAL = 2 ^ -52
    For k = 1 To DAYS_PER_YEAR
        If MAX_VAL < TEMP_MATRIX(k, 3) Then: MAX_VAL = TEMP_MATRIX(k, 3)
        If MIN_VAL > TEMP_MATRIX(k, 4) Then: MIN_VAL = TEMP_MATRIX(k, 4)
    Next k
    TEMP_MATRIX(i, NCOLUMNS + 2) = MAX_VAL
    TEMP_MATRIX(i, NCOLUMNS + 3) = MIN_VAL
    TEMP_MATRIX(i, NCOLUMNS + 4) = 1
    TEMP_MATRIX(i + 1, NCOLUMNS + 4) = 1
    TEMP_MATRIX(i, NCOLUMNS + 5) = GROENKE_VISIONS_TAI_VALUE_FUNC(TEMP_MATRIX(i, NCOLUMNS + 3), TEMP_MATRIX(i, NCOLUMNS + 2), TEMP_MATRIX(i, 5), TEMP_MATRIX(i, NCOLUMNS + 1))
    TEMP_MATRIX(i, NCOLUMNS + 6) = GROENKE_VISIONS_TAI_ACTION_FUNC(TEMP_MATRIX(i, NCOLUMNS + 5), TEMP_MATRIX(i, NCOLUMNS + 4), ii, ii_VAL, jj, jj_VAL, kk, kk_VAL, ll, ll_VAL)
    TEMP_MATRIX(i, NCOLUMNS + 7) = GROENKE_VISIONS_HIGH_LOW_DEC_FUNC(TEMP_MATRIX(i, NCOLUMNS + 3), TEMP_MATRIX(i, NCOLUMNS + 2), TEMP_MATRIX(i, 5), tolerance)
    For i = 2 To NROWS - MA_DAYS + 1
        TEMP_SUM = TEMP_SUM - TEMP_MATRIX(i - 1, 5) + TEMP_MATRIX(i + MA_DAYS - 1, 5)
        TEMP_MATRIX(i, NCOLUMNS + 1) = TEMP_SUM / MA_DAYS
        If i <= (NROWS - DAYS_PER_YEAR + 1) Then
            MIN_VAL = 2 ^ 52: MAX_VAL = 2 ^ -52
            For k = 1 To DAYS_PER_YEAR
                If MAX_VAL < TEMP_MATRIX(i + k - 1, 3) Then: MAX_VAL = TEMP_MATRIX(i + k - 1, 3)
                If MIN_VAL > TEMP_MATRIX(i + k - 1, 4) Then: MIN_VAL = TEMP_MATRIX(i + k - 1, 4)
            Next k
            TEMP_MATRIX(i, NCOLUMNS + 2) = MAX_VAL
            TEMP_MATRIX(i, NCOLUMNS + 3) = MIN_VAL
            If i > 2 Then
                TEMP_MATRIX(i, NCOLUMNS + 4) = TEMP_MATRIX(i - 2, 2) / TEMP_MATRIX(i - 1, 2)
            End If
            TEMP_MATRIX(i, NCOLUMNS + 5) = GROENKE_VISIONS_TAI_VALUE_FUNC(TEMP_MATRIX(i, NCOLUMNS + 3), TEMP_MATRIX(i, NCOLUMNS + 2), TEMP_MATRIX(i, 5), TEMP_MATRIX(i, NCOLUMNS + 1))
            TEMP_MATRIX(i, NCOLUMNS + 6) = GROENKE_VISIONS_TAI_ACTION_FUNC(TEMP_MATRIX(i, NCOLUMNS + 5), TEMP_MATRIX(i, NCOLUMNS + 4), ii, ii_VAL, jj, jj_VAL, kk, kk_VAL, ll, ll_VAL)
            TEMP_MATRIX(i, NCOLUMNS + 7) = GROENKE_VISIONS_HIGH_LOW_DEC_FUNC(TEMP_MATRIX(i, NCOLUMNS + 3), TEMP_MATRIX(i, NCOLUMNS + 2), TEMP_MATRIX(i, 5), tolerance)
        End If
    Next i
    
    If OUTPUT <> 0 Then: GoSub SUMMARY_LINE
    If NSIZE > 1 And OUTPUT <> 1 Then: TEMP_GROUP(j) = TEMP_MATRIX
    
1983:
Next j
If NSIZE = 1 Then: TEMP_GROUP = TEMP_MATRIX
Erase TEMP_MATRIX

Select Case OUTPUT
Case 0
    ASSETS_GROENKE_VISIONS_FUNC = TEMP_GROUP
Case 1
    ASSETS_GROENKE_VISIONS_FUNC = TEMP_VECTOR
Case Else
    ASSETS_GROENKE_VISIONS_FUNC = Array(TEMP_VECTOR, TEMP_GROUP)
End Select

Exit Function
'----------------------------------------------------------------------
SUMMARY_LINE:
'----------------------------------------------------------------------
    TEMP_VECTOR(6, j + 1) = NROWS - DAYS_PER_YEAR + 1
    TEMP_VECTOR(1, j + 1) = TICKER_STR
    
    TEMP_VECTOR(2, j + 1) = ii / TEMP_VECTOR(6, j + 1)
    TEMP_VECTOR(3, j + 1) = jj / TEMP_VECTOR(6, j + 1)
    TEMP_VECTOR(4, j + 1) = kk / TEMP_VECTOR(6, j + 1)
    TEMP_VECTOR(5, j + 1) = ll / TEMP_VECTOR(6, j + 1)
    
    TEMP_VECTOR(7, j + 1) = ii_VAL - 1
    TEMP_VECTOR(8, j + 1) = jj_VAL - 1
    TEMP_VECTOR(9, j + 1) = kk_VAL - 1
    TEMP_VECTOR(10, j + 1) = ll_VAL - 1
    
    TEMP_VECTOR(11, j + 1) = (TEMP_VECTOR(7, j + 1) + 1) * (TEMP_VECTOR(8, j + 1) + 1) * (TEMP_VECTOR(9, j + 1) + 1) * (TEMP_VECTOR(10, j + 1) + 1) - 1
    If ii <> 0 Then
        TEMP_VECTOR(12, j + 1) = (1 + TEMP_VECTOR(7, j + 1)) ^ (1 / ii) - 1
    Else
        TEMP_VECTOR(12, j + 1) = 0
    End If
    
    If jj <> 0 Then
        TEMP_VECTOR(13, j + 1) = (1 + TEMP_VECTOR(8, j + 1)) ^ (1 / jj) - 1
    Else
        TEMP_VECTOR(13, j + 1) = 0
    End If
    
    If kk <> 0 Then
        TEMP_VECTOR(14, j + 1) = (1 + TEMP_VECTOR(9, j + 1)) ^ (1 / kk) - 1
    Else
        TEMP_VECTOR(14, j + 1) = 0
    End If
    
    If ll <> 0 Then
        TEMP_VECTOR(15, j + 1) = (1 + TEMP_VECTOR(10, j + 1)) ^ (1 / ll) - 1
    Else
        TEMP_VECTOR(15, j + 1) = 0
    End If
    
    If (NROWS - DAYS_PER_YEAR + 1) <> 0 Then
        TEMP_VECTOR(16, j + 1) = (1 + TEMP_VECTOR(11, j + 1)) ^ (1 / (NROWS - DAYS_PER_YEAR + 1)) - 1
    Else
        TEMP_VECTOR(16, j + 1) = 0
    End If
'----------------------------------------------------------------------
Return
'----------------------------------------------------------------------
ERROR_LABEL:
ASSETS_GROENKE_VISIONS_FUNC = Err.number
End Function

Function ASSET_GROENKE_VISIONS_DECILE_FUNC(ByVal TICKER_STR As String, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal MA_DAYS As Long = 50, _
Optional ByVal DAYS_PER_YEAR As Long = 252, _
Optional ByVal tolerance As Double = 0.001)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DEC_VAL As Double 'High/Low Decile
Dim TAI_VAL As Double
Dim TAI_STR As String
Dim NDO_FACTOR As Double
Dim DATA_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'--------------------------------------------------------------------------------------------
DATA_MATRIX = ASSETS_GROENKE_VISIONS_FUNC(TICKER_STR, START_DATE, END_DATE, MA_DAYS, DAYS_PER_YEAR, tolerance, 0)
If IsArray(DATA_MATRIX) = False Then: GoTo ERROR_LABEL
'--------------------------------------------------------------------------------------------

NROWS = UBound(DATA_MATRIX, 1)
NROWS = (NROWS - DAYS_PER_YEAR + 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim DATA_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    DATA_VECTOR(i, 1) = DATA_MATRIX(i, NCOLUMNS - 2)
Next i
DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)

'--------------------------------------------------------------------------------------------
ReDim TEMP_MATRIX(1 To 24, 1 To 8)
For j = 1 To 8: For i = 1 To 24: TEMP_MATRIX(i, j) = "": Next i: Next j
'--------------------------------------------------------------------------------------------
' -5.0 <= TAI <  10.0 - Time To Act
' 10.0 <= TAI <  99.0 - Get Ready
'-10.0 <= TAI <  -5.0 - Wait
'-99.0 <= TAI < -10.0 - Bad Idea

TEMP_MATRIX(2, 1) = "0.9 <= TAI < 99.0"
TEMP_MATRIX(3, 1) = "-11.6 <= TAI < 0.9"
TEMP_MATRIX(4, 1) = "-25.6 <= TAI < -11.6"
TEMP_MATRIX(5, 1) = "-38.0 <= TAI < -25.6"
TEMP_MATRIX(6, 1) = "-46.3 <= TAI < -38.0"
TEMP_MATRIX(7, 1) = "-51.0 <= TAI < -46.3"
TEMP_MATRIX(8, 1) = "-54.8 <= TAI < -51.0"
TEMP_MATRIX(9, 1) = "-57.6 <= TAI < -54.8"
TEMP_MATRIX(10, 1) = "-59.9 <= TAI < -57.6"
TEMP_MATRIX(11, 1) = "-99.0 <= TAI < -59.9"
TEMP_MATRIX(12, 1) = "Grand Total"
'--------------------------------------------------------------------------------------------
TEMP_MATRIX(1, 1) = TICKER_STR & ": TAI Value Range"
TEMP_MATRIX(1, 2) = "Count"
TEMP_MATRIX(1, 3) = "Return"
TEMP_MATRIX(1, 4) = "Average"

'--------------------------------------------------------------------------------------------
TEMP_MATRIX(1, 5) = "Index"
TEMP_MATRIX(1, 6) = "Cumulative"
TEMP_MATRIX(1, 7) = "Lower Bound"
TEMP_MATRIX(1, 8) = "Upper Bound"

TEMP_MATRIX(2, 5) = 9
TEMP_MATRIX(2, 6) = Round(NROWS * TEMP_MATRIX(2, 5) / 10, 0)
j = TEMP_MATRIX(2, 6)
TEMP_MATRIX(2, 7) = DATA_VECTOR(j, 1)
TEMP_MATRIX(2, 8) = 99

For i = 1 To 8
    TEMP_MATRIX(2 + i, 5) = TEMP_MATRIX(1 + i, 5) - 1
    TEMP_MATRIX(2 + i, 6) = Round(NROWS * TEMP_MATRIX(2 + i, 5) / 10, 0)
    j = TEMP_MATRIX(2 + i, 6)
    TEMP_MATRIX(2 + i, 7) = DATA_VECTOR(j, 1)
    TEMP_MATRIX(2 + i, 8) = TEMP_MATRIX(1 + i, 7)
Next i

TEMP_MATRIX(11, 5) = 0
TEMP_MATRIX(11, 6) = Round(NROWS * TEMP_MATRIX(11, 5) / 10, 0)
TEMP_MATRIX(11, 7) = -99
TEMP_MATRIX(11, 8) = TEMP_MATRIX(10, 7)

'--------------------------------------------------------------------------------------------
TEMP_MATRIX(13, 1) = "High-Low Decile"
TEMP_MATRIX(13, 2) = "Count"
TEMP_MATRIX(13, 3) = "Return"
TEMP_MATRIX(13, 4) = "Average"
'--------------------------------------------------------------------------------------------
TEMP_MATRIX(13, 5) = "1-Time to Act"
TEMP_MATRIX(13, 6) = "2-Get Ready"
TEMP_MATRIX(13, 7) = "3-Wait"
TEMP_MATRIX(13, 8) = "4-Bad Idea"
'--------------------------------------------------------------------------------------------
j = 0
For i = 14 To 23
    TEMP_MATRIX(i, 1) = j
    j = j + 1
Next i

TEMP_MATRIX(24, 1) = "Grand Total"
'--------------------------------------------------------------------------------------------

For i = 2 To 11
    TEMP_MATRIX(i, 2) = 0
    TEMP_MATRIX(i, 3) = 1
    
    TEMP_MATRIX(i + 12, 2) = 0
    TEMP_MATRIX(i + 12, 3) = 1
    TEMP_MATRIX(i + 12, 4) = 0
    
    TEMP_MATRIX(i + 12, 5) = 0
    TEMP_MATRIX(i + 12, 6) = 0
    TEMP_MATRIX(i + 12, 7) = 0
    TEMP_MATRIX(i + 12, 8) = 0
Next i

For i = 1 To NROWS
    
    DEC_VAL = DATA_MATRIX(i, NCOLUMNS)
    TAI_STR = DATA_MATRIX(i, NCOLUMNS - 1)
    TAI_VAL = DATA_MATRIX(i, NCOLUMNS - 2)
    NDO_FACTOR = DATA_MATRIX(i, NCOLUMNS - 3)
    
    If 0.9 <= TAI_VAL And TAI_VAL < 99# Then
        TEMP_MATRIX(2, 2) = TEMP_MATRIX(2, 2) + 1
        TEMP_MATRIX(2, 3) = TEMP_MATRIX(2, 3) * NDO_FACTOR
    ElseIf -11.6 <= TAI_VAL And TAI_VAL < 0.9 Then
        TEMP_MATRIX(3, 2) = TEMP_MATRIX(3, 2) + 1
        TEMP_MATRIX(3, 3) = TEMP_MATRIX(3, 3) * NDO_FACTOR
    ElseIf -25.6 <= TAI_VAL And TAI_VAL < -11.6 Then
        TEMP_MATRIX(4, 2) = TEMP_MATRIX(4, 2) + 1
        TEMP_MATRIX(4, 3) = TEMP_MATRIX(4, 3) * NDO_FACTOR
    ElseIf -38# <= TAI_VAL And TAI_VAL < -25.6 Then
        TEMP_MATRIX(5, 2) = TEMP_MATRIX(5, 2) + 1
        TEMP_MATRIX(5, 3) = TEMP_MATRIX(5, 3) * NDO_FACTOR
    ElseIf -46.3 <= TAI_VAL And TAI_VAL < -38# Then
        TEMP_MATRIX(6, 2) = TEMP_MATRIX(6, 2) + 1
        TEMP_MATRIX(6, 3) = TEMP_MATRIX(6, 3) * NDO_FACTOR
    ElseIf -51# <= TAI_VAL And TAI_VAL < -46.3 Then
        TEMP_MATRIX(7, 2) = TEMP_MATRIX(7, 2) + 1
        TEMP_MATRIX(7, 3) = TEMP_MATRIX(7, 3) * NDO_FACTOR
    ElseIf -54.8 <= TAI_VAL And TAI_VAL < -51# Then
        TEMP_MATRIX(8, 2) = TEMP_MATRIX(8, 2) + 1
        TEMP_MATRIX(8, 3) = TEMP_MATRIX(8, 3) * NDO_FACTOR
    ElseIf -57.6 <= TAI_VAL And TAI_VAL < -54.8 Then
        TEMP_MATRIX(9, 2) = TEMP_MATRIX(9, 2) + 1
        TEMP_MATRIX(9, 3) = TEMP_MATRIX(9, 3) * NDO_FACTOR
    ElseIf -59.9 <= TAI_VAL And TAI_VAL < -57.6 Then
        TEMP_MATRIX(10, 2) = TEMP_MATRIX(10, 2) + 1
        TEMP_MATRIX(10, 3) = TEMP_MATRIX(10, 3) * NDO_FACTOR
    ElseIf -99# <= TAI_VAL And TAI_VAL < -59.9 Then
        TEMP_MATRIX(11, 2) = TEMP_MATRIX(11, 2) + 1
        TEMP_MATRIX(11, 3) = TEMP_MATRIX(11, 3) * NDO_FACTOR
    End If
        
    For j = 14 To 23
        If DEC_VAL = TEMP_MATRIX(j, 1) Then
            TEMP_MATRIX(j, 2) = TEMP_MATRIX(j, 2) + 1
            TEMP_MATRIX(j, 3) = TEMP_MATRIX(j, 3) * NDO_FACTOR
            If TAI_STR = TEMP_MATRIX(13, 5) Then
                TEMP_MATRIX(j, 5) = TEMP_MATRIX(j, 5) + 1
            ElseIf TAI_STR = TEMP_MATRIX(13, 6) Then
                TEMP_MATRIX(j, 6) = TEMP_MATRIX(j, 6) + 1
            ElseIf TAI_STR = TEMP_MATRIX(13, 7) Then
                TEMP_MATRIX(j, 7) = TEMP_MATRIX(j, 7) + 1
            ElseIf TAI_STR = TEMP_MATRIX(13, 8) Then
                TEMP_MATRIX(j, 8) = TEMP_MATRIX(j, 8) + 1
            End If
            Exit For
        End If
    Next j
Next i

TEMP_MATRIX(12, 2) = 0: TEMP_MATRIX(12, 3) = 1
TEMP_MATRIX(24, 2) = 0: TEMP_MATRIX(24, 3) = 1

TEMP_MATRIX(24, 5) = 0: TEMP_MATRIX(24, 6) = 0
TEMP_MATRIX(24, 7) = 0: TEMP_MATRIX(24, 8) = 0

For i = 2 To 11
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 3) - 1
    If TEMP_MATRIX(i, 2) <> 0 Then
        TEMP_MATRIX(i, 4) = (1 + TEMP_MATRIX(i, 3)) ^ (1 / TEMP_MATRIX(i, 2)) - 1
    Else
        TEMP_MATRIX(i, 4) = 0
    End If
    TEMP_MATRIX(12, 2) = TEMP_MATRIX(12, 2) + TEMP_MATRIX(i, 2)
    TEMP_MATRIX(12, 3) = TEMP_MATRIX(12, 3) * (1 + TEMP_MATRIX(i, 3))

    TEMP_MATRIX(i + 12, 3) = TEMP_MATRIX(i + 12, 3) - 1
    If TEMP_MATRIX(i + 12, 2) <> 0 Then
        TEMP_MATRIX(i + 12, 4) = (1 + TEMP_MATRIX(i + 12, 3)) ^ (1 / TEMP_MATRIX(i + 12, 2)) - 1
    Else
        TEMP_MATRIX(i + 12, 4) = 0
    End If
    TEMP_MATRIX(24, 2) = TEMP_MATRIX(24, 2) + TEMP_MATRIX(i + 12, 2)
    TEMP_MATRIX(24, 3) = TEMP_MATRIX(24, 3) * (1 + TEMP_MATRIX(i + 12, 3))

    TEMP_MATRIX(24, 5) = TEMP_MATRIX(24, 5) + TEMP_MATRIX(i + 12, 5)
    TEMP_MATRIX(24, 6) = TEMP_MATRIX(24, 6) + TEMP_MATRIX(i + 12, 6)
    TEMP_MATRIX(24, 7) = TEMP_MATRIX(24, 7) + TEMP_MATRIX(i + 12, 7)
    TEMP_MATRIX(24, 8) = TEMP_MATRIX(24, 8) + TEMP_MATRIX(i + 12, 8)
Next i

TEMP_MATRIX(12, 3) = TEMP_MATRIX(12, 3) - 1

If TEMP_MATRIX(12, 2) <> 0 Then
    TEMP_MATRIX(12, 4) = (1 + TEMP_MATRIX(12, 3)) ^ (1 / TEMP_MATRIX(12, 2)) - 1
Else
    TEMP_MATRIX(12, 4) = 0
End If
TEMP_MATRIX(24, 3) = TEMP_MATRIX(24, 3) - 1

If TEMP_MATRIX(24, 2) <> 0 Then
    TEMP_MATRIX(24, 4) = (1 + TEMP_MATRIX(24, 3)) ^ (1 / TEMP_MATRIX(24, 2)) - 1
Else
    TEMP_MATRIX(24, 4) = 0
End If
ASSET_GROENKE_VISIONS_DECILE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_GROENKE_VISIONS_DECILE_FUNC = Err.number
End Function


Private Function GROENKE_VISIONS_TAI_ACTION_FUNC(ByVal TAI_VAL As Double, _
ByVal NDO_FACTOR As Double, _
ByRef ii As Long, _
ByRef ii_VAL As Double, _
ByRef jj As Long, _
ByRef jj_VAL As Double, _
ByRef kk As Long, _
ByRef kk_VAL As Double, _
ByRef ll As Long, _
ByRef ll_VAL As Double)

On Error GoTo ERROR_LABEL

If 10 <= TAI_VAL And TAI_VAL < 2 ^ 52 Then
' 10.0 <= TAI <  99.0 - Get Ready
    GROENKE_VISIONS_TAI_ACTION_FUNC = "2-Get Ready"
    jj = jj + 1
    jj_VAL = jj_VAL * NDO_FACTOR
ElseIf -5 <= TAI_VAL And TAI_VAL < 10 Then
' -5.0 <= TAI <  10.0 - Time To Act
    GROENKE_VISIONS_TAI_ACTION_FUNC = "1-Time to Act"
    ii = ii + 1
    ii_VAL = ii_VAL * NDO_FACTOR
ElseIf -10 <= TAI_VAL And TAI_VAL < -5 Then
'-10.0 <= TAI <  -5.0 - Wait
    GROENKE_VISIONS_TAI_ACTION_FUNC = "3-Wait"
    kk = kk + 1
    kk_VAL = kk_VAL * NDO_FACTOR
ElseIf -2 ^ 52 <= TAI_VAL And TAI_VAL < -10 Then
'-99.0 <= TAI < -10.0 - Bad Idea
    GROENKE_VISIONS_TAI_ACTION_FUNC = "4-Bad Idea"
    ll = ll + 1
    ll_VAL = ll_VAL * NDO_FACTOR
End If

Exit Function
ERROR_LABEL:
GROENKE_VISIONS_TAI_ACTION_FUNC = Err.number
End Function

Private Function GROENKE_VISIONS_TAI_VALUE_FUNC(ByVal LOW52_VAL As Double, _
ByVal HIGH52_VAL As Double, _
ByVal LAST_PRICE_VAL As Double, _
ByVal MA_VAL As Double)

On Error GoTo ERROR_LABEL

GROENKE_VISIONS_TAI_VALUE_FUNC = (40 * (((3 * LOW52_VAL + HIGH52_VAL) / 4) - LAST_PRICE_VAL) / (HIGH52_VAL - LOW52_VAL)) * (1 + MA_VAL / (2 * MA_VAL - LAST_PRICE_VAL))

Exit Function
ERROR_LABEL:
GROENKE_VISIONS_TAI_VALUE_FUNC = Err.number
End Function

Private Function GROENKE_VISIONS_HIGH_LOW_DEC_FUNC(ByVal LOW52_VAL As Double, _
ByVal HIGH52_VAL As Double, _
ByVal LAST_PRICE_VAL As Double, _
Optional ByVal tolerance As Double = 0.001)

On Error GoTo ERROR_LABEL

GROENKE_VISIONS_HIGH_LOW_DEC_FUNC = Int(10 * (LAST_PRICE_VAL - LOW52_VAL) / (HIGH52_VAL - LOW52_VAL) - tolerance)
If GROENKE_VISIONS_HIGH_LOW_DEC_FUNC < 0 Then: GROENKE_VISIONS_HIGH_LOW_DEC_FUNC = 0

Exit Function
ERROR_LABEL:
GROENKE_VISIONS_HIGH_LOW_DEC_FUNC = Err.number
End Function

Sub PRINT_GROENKES_SCREENER()

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_RNG As Excel.Range
Dim DST_RNG As Excel.Range
Dim TEMP_RNG As Excel.Range
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

Set DATA_RNG = Excel.Application.InputBox("Symbols", "Groenkes Visions", , , , , , 8)
If DATA_RNG Is Nothing Then: Exit Sub

Call EXCEL_TURN_OFF_EVENTS_FUNC

Set DST_RNG = _
WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC("_"), ActiveWorkbook).Cells(3, 3)

TEMP_MATRIX = GROENKES_VISIONS_SCREENER_FUNC(DATA_RNG, , , 0)
If IsArray(TEMP_MATRIX) = False Then: GoTo 1983
        
SROW = LBound(TEMP_MATRIX, 1)
NROWS = UBound(TEMP_MATRIX, 1)
            
SCOLUMN = LBound(TEMP_MATRIX, 2)
NCOLUMNS = UBound(TEMP_MATRIX, 2)
            
Set TEMP_RNG = Range(DST_RNG.Cells(SROW, SCOLUMN), _
DST_RNG.Cells(NROWS, NCOLUMNS))
TEMP_RNG.value = TEMP_MATRIX
GoSub FORMAT_LINE

1983:
Call EXCEL_TURN_ON_EVENTS_FUNC

Exit Sub
'-----------------------------------------------------------------------------
FORMAT_LINE:
'-----------------------------------------------------------------------------
    With TEMP_RNG
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .WEIGHT = xlThin
        End With
        With .Rows(1)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        .ColumnWidth = 15
        .RowHeight = 15
    End With
    Return
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
ERROR_LABEL:
Call EXCEL_TURN_ON_EVENTS_FUNC
End Sub


