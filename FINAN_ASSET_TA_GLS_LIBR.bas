Attribute VB_Name = "FINAN_ASSET_TA_GLS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_TA_GLS_GLR_FUNC

'DESCRIPTION   : A measure of how good or bad the stock has been is the Ratio:
'Average Gain/Average Loss.
'G/L Ratio = Average[P / Min - 1] / Average[1 - P / Max]
'And if that's a big number, you buy the stock, right?
'Of course! Everybody knows the future is a replica of the past!

'LIBRARY       : FINAN_ASSET
'GROUP         : GLS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE    : 27/02/2008
'************************************************************************************
'************************************************************************************

Function ASSET_TA_GLS_GLR_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal FREQUENCY As Integer = 0)

'Gain/Loss Ratio
'Remember when we talked about Drawdown?
'We looked at the maximum stock price over the past umpteen years and
'compared it to the most recent price to see ...

'We now look at the minimum stock price over the past umpteen years and
'compare it to the most recent price to see how much we would have gained!

'If Max is the maximum and Min is the minimum price over the past umpteen
'years, and P is the current price, we look at:
'LOSS = 1 - p / Max
'and
'GAIN = p / Min - 1

'For example, if the price dropped from a maximum of $50 to the current price
'of $30, then LOSS = 1- 30/50 = 0.4 or a loss of 40%.

'If, over the same time period, the price increased from a minimum of $20 to
'the current price of $30, then GAIN = 30/20 - 1 = 0.5 or a gain of 50%.

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim PERIOD_STR As String
Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim LOSS_VAL As Double
Dim GAIN_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = True Then
    DATA_MATRIX = TICKER_STR
Else
    Select Case FREQUENCY
    Case 0
        PERIOD_STR = "d"
    Case 1
        PERIOD_STR = "w"
    Case Else
        PERIOD_STR = "m"
    End Select
    DATA_MATRIX = _
    YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
    PERIOD_STR, "DOHLCVA", False, True, True)
End If

NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 12)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "ADJ CLOSE"

TEMP_MATRIX(0, 9) = "PREV MAX" '9
TEMP_MATRIX(0, 11) = "PREV MIN" '11

i = 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000

TEMP_MATRIX(i, 8) = 1000
TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 8)
TEMP_MATRIX(i, 10) = 0
TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 9)
TEMP_MATRIX(i, 12) = 0

MAX_VAL = TEMP_MATRIX(i, 8)
MIN_VAL = TEMP_MATRIX(i, 8)

LOSS_VAL = 0: GAIN_VAL = 0
For i = 2 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000

    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i - 1, 8) * (1 + (DATA_MATRIX(i, 7) / DATA_MATRIX(i - 1, 7) - 1))
    
    If TEMP_MATRIX(i, 8) > MAX_VAL Then: MAX_VAL = TEMP_MATRIX(i, 8)
    If TEMP_MATRIX(i, 8) < MIN_VAL Then: MIN_VAL = TEMP_MATRIX(i, 8)
    TEMP_MATRIX(i, 9) = MAX_VAL
    TEMP_MATRIX(i, 11) = MIN_VAL
    
    TEMP_MATRIX(i, 10) = 1 - TEMP_MATRIX(i, 8) / TEMP_MATRIX(i, 9)
    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 8) / TEMP_MATRIX(i, 11) - 1
    
    LOSS_VAL = LOSS_VAL + TEMP_MATRIX(i, 10) 'Avg. Loss
    GAIN_VAL = GAIN_VAL + TEMP_MATRIX(i, 12) 'Avg. Gain
Next i
If LOSS_VAL = 0 Then: LOSS_VAL = 10 ^ -15
TEMP_MATRIX(0, 8) = "GROWTH OF $1K / GL RATIO = " & Format(GAIN_VAL / LOSS_VAL, "0.0")
TEMP_MATRIX(0, 10) = "AVG LOSS = " & Format(LOSS_VAL / NROWS, "0.0%") '10
TEMP_MATRIX(0, 12) = "AVG GAIN = " & Format(GAIN_VAL / NROWS, "0.0%") '12
   
ASSET_TA_GLS_GLR_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_TA_GLS_GLR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSETS_TA_GLS_GLR_FUNC

'DESCRIPTION   :
'You look at a bunch of monthly returns. (Example: The last 10 years worth.)
'You see what fraction p, are non-negative (Example: p = 52.9%)
'You calculate the Mean of those non-negative returns: G% (Example: G = 3.13%)
'You calculate the "Expected" non-negative return: p G (Example: p G = (52.9%)(3.13%)= 1.67%
'Then you repeat for the other, negative returns:
'q% negative returns with a Mean of L% hence an Expected negative return of (q%)(L%)
'(Example: q = 47.1% negative returns with a Mean of L = -3.33% giving an Expected negative
'return of (q%)(L%) = (47.1%)(-3.33%%) = -1.57%
'Don't tell me! You add them together, right?
'Wrong. You want the difference between the expected Gain and the expected Loss and that's
'their difference: GLS = pG - qL
'In our example, that'd be: GSL = (1.67%) - (-1.57%) = 3.24%.

'Reference:
'http://web.iese.edu/jestrada/PDF/Research/Others/GLS.pdf

'LIBRARY       : FINAN_ASSET
'GROUP         : GLS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE    : 27/02/2008
'************************************************************************************
'************************************************************************************

Function ASSETS_TA_GLS_GLR_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal NBINS As Long = 22, _
Optional ByVal FREQUENCY As Integer = 2, _
Optional ByVal OUTPUT As Integer = 0)

'IF FREQUENCY = 0 Then: Daily Data
'IF FREQUENCY = 1 Then: Weekly Data
'IF FREQUENCY => 2 Then: Monthly Data

Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim RETURN_VAL As Double
Dim AVG_GAIN_VAL As Double
Dim PROB_GAIN_VAL As Double

Dim AVG_LOSS_VAL As Double
Dim PROB_LOSS_VAL As Double

Dim TEMP_GROUP() As Variant
Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

Dim PERIOD_STR As String
Dim TICKER_STR As String

On Error GoTo ERROR_LABEL

If FREQUENCY = 0 Then
    PERIOD_STR = "d"
ElseIf FREQUENCY = 1 Then
    PERIOD_STR = "w"
Else
    PERIOD_STR = "m"
End If

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then: _
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If

NSIZE = UBound(TICKERS_VECTOR, 1)

'-------------------------------------------------------------------------------------
If OUTPUT <> 0 Then
'-------------------------------------------------------------------------------------
    ReDim TEMP_GROUP(1 To NSIZE)
'-------------------------------------------------------------------------------------
Else
'-------------------------------------------------------------------------------------
    ReDim TEMP_GROUP(0 To NSIZE, 1 To 12)
    TEMP_GROUP(0, 1) = "SYMBOL"
    TEMP_GROUP(0, 2) = "START DATE"
    TEMP_GROUP(0, 3) = "END DATE"
    TEMP_GROUP(0, 4) = "NO OBS"
    
    TEMP_GROUP(0, 5) = "AVG GAIN"
    TEMP_GROUP(0, 6) = "PROB GAIN"
    TEMP_GROUP(0, 7) = "#WINS"
    
    TEMP_GROUP(0, 8) = "AVG LOSS"
    TEMP_GROUP(0, 9) = "PROB LOSS"
    TEMP_GROUP(0, 10) = "#LOSSES"
    
    TEMP_GROUP(0, 11) = "GAIN LOSS SPREAD"
    TEMP_GROUP(0, 12) = "GAIN LOSS RATIO"
'-------------------------------------------------------------------------------------
End If
'-------------------------------------------------------------------------------------

'The standard deviation, arguably the most widely-used measure of risk, suffers from at
'least two limitations. First, the number itself offers little insight; after all, what
'is the intuition behind the square root of the average quadratic deviation from the
'arithmetic mean return? Second, investors tend to associate risk more with bad outcomes
'than with volatility. To overcome these limitations, this function introduces a new measure
'of risk, the gainloss spread (GLS), which is both intuitive and based on magnitudes that
'investors consider relevant when assessing risk. The evidence reported below shows that the
'GLS is highly correlated with the standard deviation, thus providing basically the same
'information about risk; and more correlated to mean returns than both the
'standard deviation and BETA, thus providing a tighter link between risk and return.

    
'-------------------------------------------------------------------------------------
If OUTPUT = 0 Then
'-------------------------------------------------------------------------------------
    For j = 1 To NSIZE
        TICKER_STR = TICKERS_VECTOR(j, 1)
        TEMP_GROUP(j, 1) = TICKER_STR
        DATA_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, PERIOD_STR, "DA", False, False, True)
        If IsArray(DATA_VECTOR) = False Then: GoTo 1983
        ii = 0: jj = 0
        AVG_GAIN_VAL = 0: PROB_GAIN_VAL = 0
        AVG_LOSS_VAL = 0: PROB_LOSS_VAL = 0
        
        NROWS = UBound(DATA_VECTOR, 1)
        For i = 2 To NROWS
            RETURN_VAL = DATA_VECTOR(i, 2) / DATA_VECTOR(i - 1, 2) - 1
            If RETURN_VAL > 0 Then
                ii = ii + 1
                AVG_GAIN_VAL = AVG_GAIN_VAL + RETURN_VAL
            ElseIf RETURN_VAL < 0 Then
                jj = jj + 1
                AVG_LOSS_VAL = AVG_LOSS_VAL + RETURN_VAL
            End If
        Next i
        
        If ii <> 0 Then
            AVG_GAIN_VAL = AVG_GAIN_VAL / ii
        Else
            AVG_GAIN_VAL = 0
        End If
        
        If jj <> 0 Then
            AVG_LOSS_VAL = AVG_LOSS_VAL / jj
        Else
            AVG_LOSS_VAL = 0
        End If
        If (jj + ii) <> 0 Then
            PROB_LOSS_VAL = jj / (jj + ii)
        Else
            PROB_LOSS_VAL = 0 'jj / 10 ^ -15
        End If
        PROB_GAIN_VAL = 1 - PROB_LOSS_VAL
        
        TEMP_GROUP(j, 2) = DATA_VECTOR(2, 1)
        TEMP_GROUP(j, 3) = DATA_VECTOR(NROWS, 1)
        TEMP_GROUP(j, 4) = NROWS - 1
        
        TEMP_GROUP(j, 5) = AVG_GAIN_VAL
        TEMP_GROUP(j, 6) = PROB_GAIN_VAL
        TEMP_GROUP(j, 7) = ii
        
        TEMP_GROUP(j, 8) = AVG_LOSS_VAL
        TEMP_GROUP(j, 9) = PROB_LOSS_VAL
        TEMP_GROUP(j, 10) = jj

        TEMP_GROUP(j, 11) = AVG_GAIN_VAL * PROB_GAIN_VAL - AVG_LOSS_VAL * PROB_LOSS_VAL
        'Gain Loss Spread
        
        If (AVG_LOSS_VAL * PROB_LOSS_VAL) <> 0 Then
            TEMP_GROUP(j, 12) = -1 * (AVG_GAIN_VAL * PROB_GAIN_VAL) / (AVG_LOSS_VAL * PROB_LOSS_VAL)
            'Gain Loss Ratio
        Else
            TEMP_GROUP(j, 12) = AVG_GAIN_VAL * PROB_GAIN_VAL '-1 * (AVG_GAIN_VAL * PROB_GAIN_VAL) / (10 ^ -15)
        End If
1983:
    Next j
'-------------------------------------------------------------------------------------
Else
'-------------------------------------------------------------------------------------
    For j = 1 To NSIZE
        TICKER_STR = TICKERS_VECTOR(j, 1)
        DATA_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, _
                      START_DATE, END_DATE, PERIOD_STR, "A", False, False, True)
        If IsArray(DATA_VECTOR) = False Then: GoTo 1984
    
        DATA_VECTOR = HISTOGRAM_DYNAMIC_FREQUENCY_FUNC(DATA_VECTOR, NBINS, 1, 0)
        NROWS = UBound(DATA_VECTOR, 1)
        ReDim TEMP_MATRIX(0 To NROWS, 1 To 4)
        TEMP_MATRIX(0, 1) = TICKER_STR & ": " & "BINS"
        TEMP_MATRIX(0, 2) = TICKER_STR & ": " & "FREQ"
        TEMP_MATRIX(0, 3) = TICKER_STR & ": " & "GAINS"
        TEMP_MATRIX(0, 4) = TICKER_STR & ": " & "LOSSES"
        
        For i = 1 To NROWS
            TEMP_MATRIX(i, 1) = DATA_VECTOR(i, 1)
            TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 2)
            'If TEMP_MATRIX(i, 2) = "" Then: TEMP_MATRIX(i, 2) = 0
            TEMP_MATRIX(i, 3) = IIf(TEMP_MATRIX(i, 1) > 0, TEMP_MATRIX(i, 2), "")
            TEMP_MATRIX(i, 4) = IIf(TEMP_MATRIX(i, 1) < 0, TEMP_MATRIX(i, 2), "")
        Next i
        
        TEMP_GROUP(j) = TEMP_MATRIX
1984:
    Next j
'-------------------------------------------------------------------------------------
End If
'-------------------------------------------------------------------------------------
    
ASSETS_TA_GLS_GLR_FUNC = TEMP_GROUP

Exit Function
ERROR_LABEL:
ASSETS_TA_GLS_GLR_FUNC = Err.number
End Function
