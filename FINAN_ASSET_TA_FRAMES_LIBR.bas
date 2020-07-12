Attribute VB_Name = "FINAN_ASSET_TA_FRAMES_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_FRAME1_FUNC

'DESCRIPTION   : Stock: Technical Analysis
'Many stock trading strategies involve staring at charts of Stock Price,
'Moving Averages, Exponential Moving Averages, MACD, RSI, Stochastics ...

'LIBRARY       : FINAN_ASSET_TA
'GROUP         : FRAME
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 09/06/2009
'************************************************************************************
'************************************************************************************

Function ASSET_TA_FRAME1_FUNC(ByVal TICKER_STR As Variant, _
ByVal INDEX_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MA1_PERIOD As Long = 27, _
Optional ByVal MA2_PERIOD As Long = 50, _
Optional ByVal EMA1_PERIOD As Long = 12, _
Optional ByVal EMA2_PERIOD As Long = 26, _
Optional ByVal RSI_PERIOD As Long = 14, _
Optional ByVal STO_PERIOD As Long = 14, _
Optional ByVal SMOOTH_PERIOD As Long = 3, _
Optional ByVal TRIGGER_PERIOD As Long = 9)

'------------------------------------------------------------------------------
'All that stuff is here, but let's recap a wee bit:
'------------------------------------------------------------------------------
'    * MA (Moving Average) of prices P:
'------------------------------------------------------------------------------
'         1. Pick a number, like N = 20.
'         2. Generate MA = Average of Stock Prices over the last N days.
'------------------------------------------------------------------------------
'    * EMA (Exponential Moving Average) of prices P:
'------------------------------------------------------------------------------
'         1. Pick a number, like N = 20 and calculate A = 1 - 2/(N+1).
'         2. Start with EMA = Stock Price.
'         3. Generate today's 20-day EMA from yesterday's like so:
'            EMA[today] = A*EMA[yesterday] + (1-A)*P[today].
'------------------------------------------------------------------------------
'    * MACD (Moving Average Convergence Divergence):
'------------------------------------------------------------------------------
'         1. Pick two numbers, like N = 12 and M = 26.
'         2. Generate:   MACD = (12-day EMA) - (26-day EMA).
'         3. Generate, in addition, the MACD "trigger" = 9-day EMA of MACD.
'            (It smooooths out the MACD curve.)
'------------------------------------------------------------------------------
'    * RSI (Relative Strength Index):
'------------------------------------------------------------------------------
'         1. Pick a number like N = 14.
'         2. Calculate U = the Average of stock Price Increases over the
'            past N days. (If it's not an Increase, ignore it.)
'         3. Calculate D = the Average of stock Price Decreases over the
'            past N days. (If it's not an Decrease, ignore it.)
'         4. Generate:   RSI = 100*{ 1 - 1/(1+U/D) }. (U and D are both
'            non-negative. If D = 0 you're in BIG trouble.)
'------------------------------------------------------------------------------
'    * RSC (Relative Strength Comparison):
'------------------------------------------------------------------------------
'         1. Pick a number like N = 14 and a "Benchmark" or Index.
'            (Like the DOW.)
'         2. Calculate P = the Stock Price Gain over the past N days.
'         3. Calculate Q = the Index Gain over the past N days.
'         4. Generate:   RSC = 100*P/Q. (If it's greater than 100 you're
'            laughin' ... cause the Stock is doin' better than the Index, eh?)
'------------------------------------------------------------------------------
'    * Stochastics (Fast and Slow):
'------------------------------------------------------------------------------
'         1. Pick a number like N = 14.
'         2. Note today's price, P.
'         3. Determine H = the largest stock Price over the past N days.
'         4. Determine L = the smallest stock Price over the past N days.
'         5. Generate:   %K Stochastic = 100*(P - L) / (H - L).
'         6. Generate, in addition, %D Stochastic = 3-day MA of %K. (It
'            smooooths out the "fast" %K stochastic curve.)
'------------------------------------------------------------------------------

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double
Dim TEMP4_SUM As Double
Dim TEMP5_SUM As Double

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim PERIOD1_VAL As Double
Dim PERIOD2_VAL As Double
Dim PERIOD3_VAL As Double

Dim DATA_MATRIX As Variant
Dim INDEX_MATRIX As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

'-----------------------------------------------------------------------------
If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If

If IsArray(INDEX_STR) = False Then
    INDEX_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(INDEX_STR, START_DATE, END_DATE, "d", "DOHLCVA", False, True, True)
Else
    INDEX_MATRIX = TICKER_STR
End If

If UBound(DATA_MATRIX, 1) <> UBound(INDEX_MATRIX, 1) Then: GoTo ERROR_LABEL
If UBound(DATA_MATRIX, 2) <> UBound(INDEX_MATRIX, 2) Then: GoTo ERROR_LABEL
'-----------------------------------------------------------------------------

PERIOD1_VAL = 1 - 2 / (EMA1_PERIOD + 1)
PERIOD2_VAL = 1 - 2 / (EMA2_PERIOD + 1)
PERIOD3_VAL = 1 - 2 / (TRIGGER_PERIOD + 1)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

'-----------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 20)
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "Dates"
TEMP_MATRIX(0, 2) = "Open"
TEMP_MATRIX(0, 3) = "High"
TEMP_MATRIX(0, 4) = "Low"
TEMP_MATRIX(0, 5) = "Close"
TEMP_MATRIX(0, 6) = "Volume"
TEMP_MATRIX(0, 7) = "Adj. Close"

TEMP_MATRIX(0, 8) = MA1_PERIOD & "-day MA"
TEMP_MATRIX(0, 9) = MA2_PERIOD & "-day MA"
TEMP_MATRIX(0, 10) = EMA1_PERIOD & "-day MA"
TEMP_MATRIX(0, 11) = EMA2_PERIOD & "-day MA"

TEMP_MATRIX(0, 12) = EMA1_PERIOD & "/" & EMA2_PERIOD & "-day MACD"
TEMP_MATRIX(0, 13) = "UpDays"
TEMP_MATRIX(0, 14) = "DownDays"
TEMP_MATRIX(0, 15) = RSI_PERIOD & "-day RSI"

TEMP_MATRIX(0, 16) = "%K: " & STO_PERIOD & "-day Fast Stochastic"
TEMP_MATRIX(0, 17) = "%D: " & SMOOTH_PERIOD & "-day Averaged  Stochastic"

TEMP_MATRIX(0, 18) = INDEX_STR & ": Adj. Close"
TEMP_MATRIX(0, 19) = _
    RSI_PERIOD & "-day % Rel. Strength (compared to " & INDEX_STR & ")"
TEMP_MATRIX(0, 20) = TRIGGER_PERIOD & "-day MACD trigger"
'-----------------------------------------------------------------------------

j = 0: k = 0: l = 0: m = 0: n = 0
TEMP1_SUM = 0: TEMP2_SUM = 0: TEMP3_SUM = 0
TEMP4_SUM = 0: TEMP5_SUM = 0

MIN_VAL = 2 ^ 52: MAX_VAL = -2 ^ 52
'-----------------------------------------------------------------------------
For i = 1 To NROWS
'-----------------------------------------------------------------------------
    
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
    TEMP_MATRIX(i, 3) = DATA_MATRIX(i, 3)
    TEMP_MATRIX(i, 4) = DATA_MATRIX(i, 4)
    TEMP_MATRIX(i, 5) = DATA_MATRIX(i, 5)
    TEMP_MATRIX(i, 6) = DATA_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 7) = DATA_MATRIX(i, 7)
    TEMP_MATRIX(i, 18) = INDEX_MATRIX(i, 7)
    
    If i <= MA1_PERIOD Then
        TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 7)
        TEMP_MATRIX(i, 8) = TEMP1_SUM / i
    Else
        If j > 0 Then: TEMP1_SUM = TEMP1_SUM - TEMP_MATRIX(j, 7)
        TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 7)
        TEMP_MATRIX(i, 8) = TEMP1_SUM / (MA1_PERIOD + 1)
        j = j + 1
    End If
    
    If i <= MA2_PERIOD Then
        TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 7)
        TEMP_MATRIX(i, 9) = TEMP2_SUM / i
    Else
        If k > 0 Then: TEMP2_SUM = TEMP2_SUM - TEMP_MATRIX(k, 7)
        TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 7)
        TEMP_MATRIX(i, 9) = TEMP2_SUM / (MA2_PERIOD + 1)
        k = k + 1
    End If
    
    If i = 1 Then
        TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 7)
        TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 7)
        TEMP_MATRIX(i, 12) = 0
        TEMP_MATRIX(i, 13) = 0
        TEMP_MATRIX(i, 14) = 0

        TEMP_MATRIX(i, 20) = 0
    Else
        
        TEMP_MATRIX(i, 10) = PERIOD1_VAL * TEMP_MATRIX(i - 1, 10) + (1 - PERIOD1_VAL) * TEMP_MATRIX(i, 7)
        TEMP_MATRIX(i, 11) = PERIOD2_VAL * TEMP_MATRIX(i - 1, 11) + (1 - PERIOD2_VAL) * TEMP_MATRIX(i, 7)
        TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 10) - TEMP_MATRIX(i, 11)
        TEMP_MATRIX(i, 13) = IIf(TEMP_MATRIX(i, 5) > TEMP_MATRIX(i - 1, 5), TEMP_MATRIX(i, 5) - TEMP_MATRIX(i - 1, 5), 0)
        TEMP_MATRIX(i, 14) = IIf(TEMP_MATRIX(i, 5) < TEMP_MATRIX(i - 1, 5), TEMP_MATRIX(i - 1, 5) - TEMP_MATRIX(i, 5), 0)
        TEMP_MATRIX(i, 20) = (PERIOD3_VAL * TEMP_MATRIX(i - 1, 20)) + ((1 - PERIOD2_VAL) * TEMP_MATRIX(i, 12))
    End If
    TEMP_MATRIX(i, 19) = 100 * (TEMP_MATRIX(i, 7) / TEMP_MATRIX(1, 7)) / (TEMP_MATRIX(i, 18) / TEMP_MATRIX(1, 18))
    '-----------------------------------------------------------------------------
    If i <= RSI_PERIOD Then
        TEMP3_SUM = TEMP3_SUM + TEMP_MATRIX(i, 13)
        TEMP4_SUM = TEMP4_SUM + TEMP_MATRIX(i, 14)
        If i = 1 Then
            TEMP_MATRIX(i, 15) = 0
        Else
            If TEMP4_SUM <> 0 Then
                TEMP_MATRIX(i, 15) = _
                100 - 100 / (1 + (TEMP3_SUM / i) / (TEMP4_SUM / i))
            Else
                TEMP_MATRIX(i, 15) = 0
            End If
        End If
    '-----------------------------------------------------------------------------
    Else
    '-----------------------------------------------------------------------------
        If l > 0 Then
            TEMP3_SUM = TEMP3_SUM - TEMP_MATRIX(l, 13)
            TEMP4_SUM = TEMP4_SUM - TEMP_MATRIX(l, 14)
        End If
        TEMP3_SUM = TEMP3_SUM + TEMP_MATRIX(i, 13)
        TEMP4_SUM = TEMP4_SUM + TEMP_MATRIX(i, 14)
        If TEMP4_SUM <> 0 Then
            TEMP_MATRIX(i, 15) = 100 - 100 / (1 + (TEMP3_SUM / (1 + RSI_PERIOD)) / (TEMP4_SUM / (1 + RSI_PERIOD)))
        Else
            TEMP_MATRIX(i, 15) = 0
        End If
        l = l + 1
    '-----------------------------------------------------------------------------
    End If
    '-----------------------------------------------------------------------------
    
    '-----------------------------------------------------------------------------
    If i <= STO_PERIOD Then
    '-----------------------------------------------------------------------------
        If TEMP_MATRIX(i, 7) < MIN_VAL Then: MIN_VAL = TEMP_MATRIX(i, 7)
        If TEMP_MATRIX(i, 7) > MAX_VAL Then: MAX_VAL = TEMP_MATRIX(i, 7)
    '-----------------------------------------------------------------------------
    Else
    '-----------------------------------------------------------------------------
        MIN_VAL = 2 ^ 52: MAX_VAL = -2 ^ 52
        For m = i To (i - STO_PERIOD) Step -1
            If TEMP_MATRIX(m, 7) < MIN_VAL Then: MIN_VAL = TEMP_MATRIX(m, 7)
            If TEMP_MATRIX(m, 7) > MAX_VAL Then: MAX_VAL = TEMP_MATRIX(m, 7)
        Next m
    '-----------------------------------------------------------------------------
    End If
    '-----------------------------------------------------------------------------
    
    If i = 1 Or Abs(MAX_VAL - MIN_VAL) <= 10 ^ -10 Then
        TEMP_MATRIX(i, 16) = 0
    Else
        TEMP_MATRIX(i, 16) = 100 * (TEMP_MATRIX(i, 7) - MIN_VAL) / (MAX_VAL - MIN_VAL)
    End If

    '-----------------------------------------------------------------------------
    If i <= SMOOTH_PERIOD - 1 Then
    '-----------------------------------------------------------------------------
        TEMP5_SUM = TEMP5_SUM + TEMP_MATRIX(i, 16)
        If i = 1 Then
            TEMP_MATRIX(i, 17) = 0
        Else
            TEMP_MATRIX(i, 17) = TEMP5_SUM / i
        End If
    '-----------------------------------------------------------------------------
    Else
    '-----------------------------------------------------------------------------
        If n > 0 Then: TEMP5_SUM = TEMP5_SUM - TEMP_MATRIX(n, 16)
        TEMP5_SUM = TEMP5_SUM + TEMP_MATRIX(i, 16)
        TEMP_MATRIX(i, 17) = TEMP5_SUM / (SMOOTH_PERIOD + 0)
        n = n + 1
    '-----------------------------------------------------------------------------
    End If
    '-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
Next i
'-----------------------------------------------------------------------------

ASSET_TA_FRAME1_FUNC = TEMP_MATRIX


Exit Function
ERROR_LABEL:
ASSET_TA_FRAME1_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_FRAME2_FUNC

'DESCRIPTION   : Stock: Technical Analysis
'Many stock trading strategies involve staring at charts of Stock Price,
'Moving Averages, Exponential Moving Averages, MACD, RSI, Stochastics ...

'LIBRARY       : FINAN_ASSET_TA
'GROUP         : FRAME
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 09/06/2009
'************************************************************************************
'************************************************************************************

Function ASSET_TA_FRAME2_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal ATR_PERIOD As Long = 14, _
Optional ByVal CCI_PERIOD As Long = 20, _
Optional ByVal EMA_PERIOD As Long = 10, _
Optional ByVal MACD1_PERIOD As Long = 5, _
Optional ByVal MACD2_PERIOD As Long = 20, _
Optional ByVal MACD3_PERIOD As Long = 9, _
Optional ByVal RSI_PERIOD As Long = 2, _
Optional ByVal SMA_PERIOD As Long = 10, _
Optional ByVal PERIOD1_VAL As Long = 14, _
Optional ByVal PERIOD2_VAL As Long = 3, _
Optional ByVal PERIOD3_VAL As Long = 1)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "d", "DOHLCV", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = 6

ReDim TEMP2_MATRIX(0 To NROWS, 1 To NCOLUMNS + 28)
TEMP2_MATRIX(0, 1) = "DATE"
TEMP2_MATRIX(0, 2) = "OPEN"
TEMP2_MATRIX(0, 3) = "HIGH"
TEMP2_MATRIX(0, 4) = "LOW"
TEMP2_MATRIX(0, 5) = "CLOSE"
TEMP2_MATRIX(0, 6) = "VOLUME"

For i = 1 To NROWS: For j = 1 To NCOLUMNS: TEMP2_MATRIX(i, j) = DATA_MATRIX(i, j): Next j: Next i
k = NCOLUMNS + 1
For h = 1 To 10

    Select Case h
    Case 1 'Perfect
        TEMP1_MATRIX = ASSET_TA_ADL_FUNC(DATA_MATRIX)
    Case 2
        TEMP1_MATRIX = ASSET_TA_ATR_FUNC(DATA_MATRIX, ATR_PERIOD)
    Case 3
        TEMP1_MATRIX = ASSET_TA_CCI_FUNC(DATA_MATRIX, CCI_PERIOD)
    Case 4
        TEMP1_MATRIX = ASSET_TA_EMA_FUNC(DATA_MATRIX, EMA_PERIOD)
    Case 5
        TEMP1_MATRIX = ASSET_TA_MACD_FUNC(DATA_MATRIX, MACD1_PERIOD, MACD2_PERIOD, MACD3_PERIOD)
    Case 6
        TEMP1_MATRIX = ASSET_TA_OBV_FUNC(DATA_MATRIX)
    Case 7
        TEMP1_MATRIX = ASSET_TA_ROC_FUNC(DATA_MATRIX)
    Case 8
        TEMP1_MATRIX = ASSET_TA_RSI_FUNC(DATA_MATRIX, RSI_PERIOD, 5)
    Case 9
        TEMP1_MATRIX = ASSET_TA_SMA_FUNC(DATA_MATRIX, SMA_PERIOD)
    Case 10
        TEMP1_MATRIX = ASSET_TA_STO_FUNC(DATA_MATRIX, PERIOD1_VAL, PERIOD2_VAL, PERIOD3_VAL)
    End Select
    
    For j = LBound(TEMP1_MATRIX, 2) To UBound(TEMP1_MATRIX, 2)
        For i = LBound(TEMP1_MATRIX, 1) To UBound(TEMP1_MATRIX, 1)
            TEMP2_MATRIX(i, k) = TEMP1_MATRIX(i, j)
        Next i
        k = k + 1
    Next j
Next h

ASSET_TA_FRAME2_FUNC = TEMP2_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_FRAME2_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_ADL_FUNC

'DESCRIPTION   : Technical analysis indicators / Accumulation/Distribution Line
'Note: The DATA_RNG parameter is assumed to be a range of historical quotes
'data (e.g. from Yahoo!) where the columns are Date/Open/High/Low/Close/Volume,
'the first row contain column names, and the rows are in ascending date sequence.

'LIBRARY       : FINAN_ASSET_TA
'GROUP         : FRAME
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 09/06/2009
'************************************************************************************
'************************************************************************************

Private Function ASSET_TA_ADL_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim NROWS As Long

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG 'Initialize return array
NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 2)

TEMP_MATRIX(0, 1) = "ADL"
TEMP_MATRIX(0, 2) = "CLV"
TEMP1_VAL = 0
TEMP2_VAL = 0
For i = 1 To NROWS
    If DATA_MATRIX(i, 3) > DATA_MATRIX(i, 4) Then
       TEMP2_VAL = (DATA_MATRIX(i, 5) - DATA_MATRIX(i, 4) + DATA_MATRIX(i, 5) - DATA_MATRIX(i, 3)) / (DATA_MATRIX(i, 3) - DATA_MATRIX(i, 4))
       TEMP1_VAL = TEMP1_VAL + TEMP2_VAL * DATA_MATRIX(i, 6)
    End If
    TEMP_MATRIX(i, 1) = TEMP1_VAL
    TEMP_MATRIX(i, 2) = TEMP2_VAL
Next i

ASSET_TA_ADL_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_ADL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_ATR_FUNC

'DESCRIPTION   : Average True Range
'Note: The DATA_RNG parameter is assumed to be a range of historical quotes
'data (e.g. from Yahoo!) where the columns are Date/Open/High/Low/Close/Volume,
'the first row contain column names, and the rows are in ascending date sequence.

'LIBRARY       : FINAN_ASSET_TA
'GROUP         : FRAME
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 09/06/2009
'************************************************************************************
'************************************************************************************

Private Function ASSET_TA_ATR_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal PERIOD As Long = 20)

Dim i As Long
Dim NROWS As Long

Dim TEMP_SUM As Double

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP3_VAL As Double
Dim TEMP4_VAL As Double
Dim TEMP5_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG 'Initialize return array
NROWS = UBound(DATA_MATRIX, 1)

If PERIOD = 0 Then PERIOD = 20
    
ReDim TEMP_MATRIX(0 To NROWS, 1 To 5)

TEMP_MATRIX(0, 1) = "ATR_" & PERIOD
TEMP_MATRIX(0, 2) = "H-L"
TEMP_MATRIX(0, 3) = "Abs(H-C1)"
TEMP_MATRIX(0, 4) = "Abs(L-C1)"
TEMP_MATRIX(0, 5) = "TR"
    
TEMP_SUM = 0

For i = 1 To NROWS
    TEMP2_VAL = DATA_MATRIX(i, 3) - DATA_MATRIX(i, 4)
    If i = 1 Then
       TEMP3_VAL = 0
       TEMP4_VAL = 0
       TEMP5_VAL = TEMP2_VAL
    Else
       TEMP3_VAL = Abs(DATA_MATRIX(i, 3) - DATA_MATRIX(i - 1, 5))
       TEMP4_VAL = Abs(DATA_MATRIX(i, 4) - DATA_MATRIX(i - 1, 5))
       TEMP5_VAL = IIf(TEMP4_VAL > TEMP3_VAL, TEMP4_VAL, TEMP3_VAL)
       TEMP5_VAL = IIf(TEMP5_VAL > TEMP2_VAL, TEMP5_VAL, TEMP2_VAL)
    End If
    If i > PERIOD Then
       TEMP1_VAL = (TEMP5_VAL + (PERIOD - 1) * _
                    TEMP_MATRIX(i - 1, 1)) / PERIOD
    Else
       TEMP_SUM = TEMP_SUM + TEMP5_VAL
       TEMP1_VAL = TEMP_SUM / i
    End If
    TEMP_MATRIX(i, 1) = TEMP1_VAL
    TEMP_MATRIX(i, 2) = TEMP2_VAL
    TEMP_MATRIX(i, 3) = TEMP3_VAL
    TEMP_MATRIX(i, 4) = TEMP4_VAL
    TEMP_MATRIX(i, 5) = TEMP5_VAL
Next i

ASSET_TA_ATR_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_ATR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_CCI_FUNC

'DESCRIPTION   : Commodity Channel Index
'Note: The DATA_RNG parameter is assumed to be a range of historical quotes
'data (e.g. from Yahoo!) where the columns are Date/Open/High/Low/Close/Volume,
'the first row contain column names, and the rows are in ascending date sequence.

'LIBRARY       : FINAN_ASSET_TA
'GROUP         : FRAME
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 09/06/2009
'************************************************************************************
'************************************************************************************

Private Function ASSET_TA_CCI_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal PERIOD As Long = 20)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP3_VAL As Double
Dim TEMP4_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG 'Initialize return array
NROWS = UBound(DATA_MATRIX, 1)

If PERIOD = 0 Then PERIOD = 20
    
ReDim TEMP_MATRIX(0 To NROWS, 1 To 4)

TEMP_MATRIX(0, 1) = "CCI_" & PERIOD
TEMP_MATRIX(0, 2) = "TP"
TEMP_MATRIX(0, 3) = "TPMA"
TEMP_MATRIX(0, 4) = "MD"

TEMP1_SUM = 0
For i = 1 To NROWS
    TEMP2_VAL = (DATA_MATRIX(i, 5) + _
                 DATA_MATRIX(i, 3) + DATA_MATRIX(i, 4)) / 3
'-------------------------------------------------------------------------------
    If i > PERIOD Then
'-------------------------------------------------------------------------------
       TEMP1_SUM = TEMP1_SUM + TEMP2_VAL - TEMP_MATRIX(i - PERIOD, 2)
       TEMP3_VAL = TEMP1_SUM / PERIOD
       TEMP2_SUM = Abs(TEMP3_VAL - TEMP2_VAL)
       For j = i - PERIOD + 1 To i - 1
           TEMP2_SUM = TEMP2_SUM + Abs(TEMP3_VAL - TEMP_MATRIX(j, 2))
       Next j
       TEMP4_VAL = TEMP2_SUM / PERIOD
       TEMP1_VAL = (TEMP2_VAL - TEMP3_VAL) / (0.015 * TEMP4_VAL)
'-------------------------------------------------------------------------------
    Else
'-------------------------------------------------------------------------------
       TEMP1_SUM = TEMP1_SUM + TEMP2_VAL
       TEMP3_VAL = TEMP1_SUM / i
       TEMP1_VAL = 0
       TEMP4_VAL = 0
'-------------------------------------------------------------------------------
    End If
'-------------------------------------------------------------------------------
    TEMP_MATRIX(i, 1) = TEMP1_VAL
    TEMP_MATRIX(i, 2) = TEMP2_VAL 'PERFECT
    TEMP_MATRIX(i, 3) = TEMP3_VAL 'PERFECT
    TEMP_MATRIX(i, 4) = TEMP4_VAL
Next i
ASSET_TA_CCI_FUNC = TEMP_MATRIX
    
Exit Function
ERROR_LABEL:
ASSET_TA_CCI_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_EMA_FUNC

'DESCRIPTION   : Exponential Moving Average
'Note: The DATA_RNG parameter is assumed to be a range of historical quotes
'data (e.g. from Yahoo!) where the columns are Date/Open/High/Low/Close/Volume,
'the first row contain column names, and the rows are in ascending date sequence.

'EMA (Exponential Moving Average) of prices P:
'1. Pick a number, like N = 20 and calculate A = 1 - 2/(N+1).
'2. Start with EMA = Stock Price.
'3. Generate today's 20-day EMA from yesterday's like so:
'EMA[today] = A*EMA[yesterday] + (1-A)*P[today].

'LIBRARY       : FINAN_ASSET_TA
'GROUP         : FRAME
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 09/06/2009
'************************************************************************************
'************************************************************************************

Private Function ASSET_TA_EMA_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal PERIOD As Long = 20)

Dim i As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG 'Initialize return array
NROWS = UBound(DATA_MATRIX, 1)

If PERIOD = 0 Then PERIOD = 50

ReDim TEMP_MATRIX(0 To NROWS, 1 To 1)
TEMP_MATRIX(0, 1) = "EMA_" & PERIOD
TEMP2_VAL = 2 / (PERIOD + 1)

TEMP_SUM = 0
For i = 1 To NROWS
    If i > PERIOD Then
       TEMP1_VAL = TEMP2_VAL * (DATA_MATRIX(i, 5) - TEMP1_VAL) + TEMP1_VAL
       TEMP_MATRIX(i, 1) = TEMP1_VAL
    Else
       TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, 5)
       TEMP1_VAL = TEMP_SUM / i
       TEMP_MATRIX(i, 1) = IIf(i < PERIOD, 0, TEMP1_VAL)
    End If
Next i

ASSET_TA_EMA_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_EMA_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_EMA_FUNC

'DESCRIPTION   : 'Moving Average Convergence Divergence
'Note: The DATA_RNG parameter is assumed to be a range of historical quotes
'data (e.g. from Yahoo!) where the columns are Date/Open/High/Low/Close/Volume,
'the first row contain column names, and the rows are in ascending date sequence.

'MACD (Moving Average Convergence Divergence):
'1. Pick two numbers, like N = 12 and M = 26.
'2. Generate:   MACD = (12-day EMA) - (26-day EMA).
'3. Generate, in addition, the MACD "trigger" = 9-day EMA of
'MACD. (It smooooths out the MACD curve.)

'LIBRARY       : FINAN_ASSET_TA
'GROUP         : FRAME
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 09/06/2009
'************************************************************************************
'************************************************************************************

Private Function ASSET_TA_MACD_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal MACD1_PERIOD As Long = 5, _
Optional ByVal MACD2_PERIOD As Long = 20, _
Optional ByVal MACD3_PERIOD As Long = 9)

Dim i As Long
Dim NROWS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP3_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG 'Initialize return array
NROWS = UBound(DATA_MATRIX, 1)

If MACD1_PERIOD = 0 Then MACD1_PERIOD = 12
If MACD2_PERIOD = 0 Then MACD2_PERIOD = 26
If MACD3_PERIOD = 0 Then MACD3_PERIOD = 9

ReDim TEMP_MATRIX(0 To NROWS, 1 To 3)

TEMP_MATRIX(0, 1) = "MACD: " & MACD1_PERIOD & "-" & MACD2_PERIOD & "-" & MACD3_PERIOD
TEMP_MATRIX(0, 2) = "SMA_" & MACD1_PERIOD
TEMP_MATRIX(0, 3) = "SMA_" & MACD2_PERIOD

TEMP1_SUM = 0
TEMP2_SUM = 0
TEMP3_SUM = 0
For i = 1 To NROWS
    If i > MACD1_PERIOD Then
       TEMP2_SUM = TEMP2_SUM + DATA_MATRIX(i, 5) - DATA_MATRIX(i - MACD1_PERIOD, 5)
       TEMP2_VAL = TEMP2_SUM / MACD1_PERIOD
    Else
       TEMP2_SUM = TEMP2_SUM + DATA_MATRIX(i, 5)
       TEMP2_VAL = TEMP2_SUM / i
    End If
    
    If i > MACD2_PERIOD Then
       TEMP3_SUM = TEMP3_SUM + DATA_MATRIX(i, 5) - DATA_MATRIX(i - MACD2_PERIOD, 5)
       TEMP3_VAL = TEMP3_SUM / MACD2_PERIOD
    Else
       TEMP3_SUM = TEMP3_SUM + DATA_MATRIX(i, 5)
       TEMP3_VAL = TEMP3_SUM / i
    End If
    
    If i > MACD3_PERIOD Then
       TEMP1_SUM = TEMP1_SUM + (TEMP2_VAL - TEMP3_VAL) - _
       (TEMP_MATRIX(i - MACD3_PERIOD, 2) - _
       TEMP_MATRIX(i - MACD3_PERIOD, 3))
       TEMP1_VAL = TEMP1_SUM / MACD3_PERIOD
    Else
       TEMP1_SUM = TEMP1_SUM + (TEMP2_VAL - TEMP3_VAL)
       TEMP1_VAL = TEMP1_SUM / i
    End If
    TEMP_MATRIX(i, 1) = TEMP1_VAL
    TEMP_MATRIX(i, 2) = TEMP2_VAL
    TEMP_MATRIX(i, 3) = TEMP3_VAL
Next i


ASSET_TA_MACD_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_MACD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_RSI_FUNC

'DESCRIPTION   : Relative Strength Index
'Note: The DATA_RNG parameter is assumed to be a range of historical quotes
'data (e.g. from Yahoo!) where the columns are Date/Open/High/Low/Close/Volume,
'the first row contain column names, and the rows are in ascending date sequence.

'RSI (Relative Strength Index):
'1. Pick a number like N = 14.
'2. Calculate U = the Average of stock Price Increases over the past N days.
'(If it's not an Increase, ignore it.)
'3. Calculate D = the Average of stock Price Decreases over the past N days.
'(If it's not an Decrease, ignore it.)
'4. Generate:   RSI = 100*{ 1 - 1/(1+U/D) }. (U and D are both
'non-negative. If D = 0 you're in BIG trouble.)

'LIBRARY       : FINAN_ASSET_TA
'GROUP         : FRAME
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 09/06/2009
'************************************************************************************
'************************************************************************************

Function ASSET_TA_RSI_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal PERIOD As Long = 20, _
Optional ByVal SCOLUMN As Long = 5)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP3_VAL As Double
Dim TEMP4_VAL As Double
Dim TEMP5_VAL As Double
Dim TEMP6_VAL As Double
Dim TEMP7_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG 'Initialize return array
NROWS = UBound(DATA_MATRIX, 1)

If PERIOD = 0 Then PERIOD = 20

NCOLUMNS = 7
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS: For i = 1 To NROWS: TEMP_MATRIX(i, j) = "": Next i: Next j

TEMP_MATRIX(0, 1) = "RSI_" & PERIOD
TEMP_MATRIX(0, 2) = "CHG"
TEMP_MATRIX(0, 3) = "ADVA"
TEMP_MATRIX(0, 4) = "DECL"
TEMP_MATRIX(0, 5) = "AVG_GAIN"
TEMP_MATRIX(0, 6) = "AVG_LOSS"
TEMP_MATRIX(0, 7) = "RS"

TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 3 To NROWS 'i = 2 to NROWS
    TEMP2_VAL = DATA_MATRIX(i, SCOLUMN) - DATA_MATRIX(i - 1, SCOLUMN)
    If i > PERIOD + 2 Then ' i > PERIOD + 1
       If TEMP2_VAL > 0 Then
          TEMP3_VAL = TEMP2_VAL
          TEMP4_VAL = 0
          TEMP5_VAL = ((PERIOD - 1) * TEMP5_VAL + TEMP2_VAL) / PERIOD
          TEMP6_VAL = ((PERIOD - 1) * TEMP6_VAL) / PERIOD
       Else
          TEMP3_VAL = 0
          TEMP4_VAL = -TEMP2_VAL
          TEMP5_VAL = ((PERIOD - 1) * TEMP5_VAL) / PERIOD
          TEMP6_VAL = ((PERIOD - 1) * TEMP6_VAL - TEMP2_VAL) / PERIOD
       End If
'----------------------------------------------------------------------------------
    Else
'----------------------------------------------------------------------------------
       If TEMP2_VAL > 0 Then
          TEMP3_VAL = TEMP2_VAL
          TEMP1_SUM = TEMP1_SUM + TEMP2_VAL
          TEMP4_VAL = 0
       Else
          TEMP3_VAL = 0
          TEMP4_VAL = -TEMP2_VAL
          TEMP2_SUM = TEMP2_SUM - TEMP2_VAL
       End If
       If i = PERIOD + 2 Then 'i = PERIOD + 1
          TEMP5_VAL = TEMP1_SUM / PERIOD
          TEMP6_VAL = TEMP2_SUM / PERIOD
       Else
          TEMP5_VAL = 0
          TEMP6_VAL = 0
       End If
'----------------------------------------------------------------------------------
    End If
'----------------------------------------------------------------------------------
    If TEMP5_VAL = 0 Then
       TEMP1_VAL = 0
       TEMP7_VAL = 0
'----------------------------------------------------------------------------------
    Else
'----------------------------------------------------------------------------------
       If TEMP6_VAL = 0 Then
          TEMP7_VAL = 0
          TEMP1_VAL = 100
       Else
          TEMP7_VAL = TEMP5_VAL / TEMP6_VAL
          TEMP1_VAL = 100 - (100 / (1 + TEMP7_VAL))
       End If
'----------------------------------------------------------------------------------
    End If
'----------------------------------------------------------------------------------
    If TEMP1_VAL <> 0 Then: TEMP_MATRIX(i, 1) = TEMP1_VAL
    If TEMP2_VAL <> 0 Then: TEMP_MATRIX(i, 2) = TEMP2_VAL
    If TEMP3_VAL <> 0 Then: TEMP_MATRIX(i, 3) = TEMP3_VAL
    If TEMP4_VAL <> 0 Then: TEMP_MATRIX(i, 4) = TEMP4_VAL
    If TEMP5_VAL <> 0 Then: TEMP_MATRIX(i, 5) = TEMP5_VAL
    If TEMP6_VAL <> 0 Then: TEMP_MATRIX(i, 6) = TEMP6_VAL
    If TEMP7_VAL <> 0 Then: TEMP_MATRIX(i, 7) = TEMP7_VAL
Next i

        
ASSET_TA_RSI_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_RSI_FUNC = Err.number
End Function


'On Balance Volume
'Note: The DATA_RNG parameter is assumed to be a range of historical quotes
'data (e.g. from Yahoo!) where the columns are Date/Open/High/Low/Close/Volume,
'the first row contain column names, and the rows are in ascending date sequence.

Private Function ASSET_TA_OBV_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim NROWS As Long

Dim TEMP_VAL As Double
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG 'Initialize return array
NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 1)
TEMP_MATRIX(0, 1) = "OBV"
TEMP_MATRIX(1, 1) = 0

TEMP_VAL = 0
For i = 2 To NROWS
    Select Case (DATA_MATRIX(i, 5) - DATA_MATRIX(i - 1, 5))
       Case Is > 0: TEMP_VAL = TEMP_VAL + DATA_MATRIX(i, 6)
       Case Is < 0: TEMP_VAL = TEMP_VAL - DATA_MATRIX(i, 6)
    End Select
    TEMP_MATRIX(i, 1) = TEMP_VAL
Next i

ASSET_TA_OBV_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_OBV_FUNC = Err.number
End Function

'Rate of Change
'Note: The DATA_RNG parameter is assumed to be a range of historical quotes
'data (e.g. from Yahoo!) where the columns are Date/Open/High/Low/Close/Volume,
'the first row contain column names, and the rows are in ascending date sequence.

'    * RSC (Relative Strength Comparison):
'         1. Pick a number like N = 14 and a "Benchmark" or Index. (Like the DOW.)
'         2. Calculate P = the Stock Price Gain over the past N days.
'         3. Calculate Q = the Index Gain over the past N days.
'         4. Generate:   RSC = 100*P/Q. (If it's greater than 100 you're
'         laughin' ... cause the Stock is doin' better than the Index, eh?)

Private Function ASSET_TA_ROC_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal PERIOD As Long = 21)

Dim i As Long
Dim NROWS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG 'Initialize return array
NROWS = UBound(DATA_MATRIX, 1)

    If PERIOD = 0 Then PERIOD = 21
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 1)
    TEMP_MATRIX(0, 1) = "ROC_" & PERIOD
    
    For i = 1 To NROWS
        If i < PERIOD + 1 Then
           TEMP_MATRIX(i, 1) = 0
        Else
           TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 5) / DATA_MATRIX(i - PERIOD, 5) - 1
        End If
    Next i
    
    ASSET_TA_ROC_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_ROC_FUNC = Err.number
End Function

'Simple Moving Average
'Note: The DATA_RNG parameter is assumed to be a range of historical quotes
'data (e.g. from Yahoo!) where the columns are Date/Open/High/Low/Close/Volume,
'the first row contain column names, and the rows are in ascending date sequence.

'    * MA (Moving Average) of prices P:
'         1. Pick a number, like N = 20.
'         2. Generate MA = Average of Stock Prices over the last N days.

Private Function ASSET_TA_SMA_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal PERIOD As Long = 50)

Dim i As Long
Dim NROWS As Long

Dim TEMP1_SUM As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG 'Initialize return array
NROWS = UBound(DATA_MATRIX, 1)

If PERIOD = 0 Then PERIOD = 50

ReDim TEMP_MATRIX(0 To NROWS, 1 To 1)
TEMP_MATRIX(0, 1) = "SMA_" & PERIOD

TEMP1_SUM = 0
For i = 1 To NROWS
    If i > PERIOD Then
       TEMP1_SUM = TEMP1_SUM + DATA_MATRIX(i, 5) - DATA_MATRIX(i - PERIOD, 5)
       TEMP_MATRIX(i, 1) = TEMP1_SUM / PERIOD
    Else
       TEMP1_SUM = TEMP1_SUM + DATA_MATRIX(i, 5)
       TEMP_MATRIX(i, 1) = TEMP1_SUM / i
    End If
Next i

ASSET_TA_SMA_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_SMA_FUNC = Err.number
End Function



'Stochastics
'Note: The DATA_RNG parameter is assumed to be a range of historical quotes
'data (e.g. from Yahoo!) where the columns are Date/Open/High/Low/Close/Volume,
'the first row contain column names, and the rows are in ascending date sequence.

'    * Stochastics (Fast and Slow):
'         1. Pick a number like N = 14.
'         2. Note today's price, P.
'         3. Determine H = the largest stock Price over the past N days.
'         4. Determine L = the smallest stock Price over the past N days.
'         5. Generate:   %K Stochastic = 100*(P - H) / (H - L).
'         6. Generate, in addition, %D Stochastic = 3-day MA of %K. (It
'         smooooths out the "fast" %K stochastic curve.)

Private Function ASSET_TA_STO_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal PERIOD1_VAL As Long = 5, _
Optional ByVal PERIOD2_VAL As Long = 20, _
Optional ByVal PERIOD3_VAL As Long = 9)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim LOW_VAL As Double
Dim HIGH_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim TEMP3_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG 'Initialize return array
NROWS = UBound(DATA_MATRIX, 1)

If PERIOD1_VAL = 0 Then PERIOD1_VAL = 14
If PERIOD2_VAL = 0 Then PERIOD2_VAL = 5
If PERIOD3_VAL = 0 Then PERIOD3_VAL = 1

ReDim TEMP_MATRIX(0 To NROWS, 1 To 3)

TEMP_MATRIX(0, 1) = "STOCH: " & PERIOD1_VAL & "-" & PERIOD2_VAL & "-" & PERIOD3_VAL
TEMP_MATRIX(0, 2) = "%K"
TEMP_MATRIX(0, 3) = "%D"

TEMP1_SUM = 0
TEMP2_SUM = 0

For i = 1 To NROWS
    HIGH_VAL = DATA_MATRIX(i, 3)
    LOW_VAL = DATA_MATRIX(i, 4)
    
    For j = IIf(i - PERIOD1_VAL + 1 > 1, i - PERIOD1_VAL + 1, 1) To i - 1
        If DATA_MATRIX(j, 3) > HIGH_VAL Then HIGH_VAL = DATA_MATRIX(j, 3)
        If DATA_MATRIX(j, 4) < LOW_VAL Then LOW_VAL = DATA_MATRIX(j, 4)
    Next j
    TEMP2_VAL = 100 * (DATA_MATRIX(i, 5) - LOW_VAL) / (HIGH_VAL - LOW_VAL)
    
    If i > PERIOD2_VAL Then
       TEMP2_SUM = TEMP2_SUM + TEMP2_VAL - TEMP_MATRIX(i - PERIOD2_VAL, 2)
       TEMP3_VAL = TEMP2_SUM / PERIOD2_VAL
    Else
       TEMP2_SUM = TEMP2_SUM + TEMP2_VAL
       TEMP3_VAL = TEMP2_SUM / i
    End If
    If i > PERIOD3_VAL Then
       TEMP1_SUM = TEMP1_SUM + TEMP3_VAL - TEMP_MATRIX(i - PERIOD3_VAL, 3)
       TEMP1_VAL = TEMP1_SUM / PERIOD3_VAL
    Else
       TEMP1_SUM = TEMP1_SUM + TEMP3_VAL
       TEMP1_VAL = TEMP1_SUM / i
    End If
    TEMP_MATRIX(i, 1) = TEMP1_VAL
    TEMP_MATRIX(i, 2) = TEMP2_VAL
    TEMP_MATRIX(i, 3) = TEMP3_VAL
Next i

ASSET_TA_STO_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_STO_FUNC = Err.number
End Function
