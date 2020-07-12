Attribute VB_Name = "FINAN_ASSET_MOMENTS_VOLAT_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'Historical Volatility Calculation

Function ASSETS_HISTORICAL_VOLATILITY_FUNC(ByVal TICKERS_RNG As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal PERIODS_RNG As Variant = 10, _
Optional ByVal TDAYS_PER_YEAR As Double = 252, _
Optional ByVal OUTPUT As Integer = 0)

'References:
'http://www.neuralmarkettrends.com/2007/05/29/calculating-historical-volatility/
'http://www.quantonline.co.za/Articles/article_volatility.htm
'http://www.neuralmarkettrends.com/wp-content/uploads/2007/05/volatility_calculation.pdf

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

Dim TICKERS_VECTOR As Variant
Dim PERIODS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(PERIODS_RNG) Then
    PERIODS_VECTOR = PERIODS_RNG
    If UBound(PERIODS_VECTOR, 2) = 1 Then
        PERIODS_VECTOR = MATRIX_TRANSPOSE_FUNC(PERIODS_VECTOR)
    End If
Else
    ReDim PERIODS_VECTOR(1 To 1, 1 To 1)
    PERIODS_VECTOR(1, 1) = PERIODS_RNG
End If
NCOLUMNS = UBound(PERIODS_VECTOR, 2)

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
    NROWS = UBound(TICKERS_VECTOR, 1)
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 7)
    TEMP_MATRIX(0, 1) = "HV d" & PERIODS_VECTOR(1, 1)
    TEMP_MATRIX(0, 2) = "MINIMUM"
    TEMP_MATRIX(0, 3) = "25TH PERCENTILE"
    TEMP_MATRIX(0, 4) = "50TH PERCENTILE"
    TEMP_MATRIX(0, 5) = "MEAN"
    TEMP_MATRIX(0, 6) = "75TH PERCENTILE"
    TEMP_MATRIX(0, 7) = "MAXIMUM"
    For i = 1 To NROWS
        DATA_MATRIX = ASSETS_HISTORICAL_VOLATILITY_FUNC(TICKERS_VECTOR(i, 1), START_DATE, END_DATE, PERIODS_VECTOR(1, 1), TDAYS_PER_YEAR, 1)
        TEMP_MATRIX(i, 1) = TICKERS_VECTOR(i, 1)
        For j = 1 To 6: TEMP_MATRIX(i, j + 1) = DATA_MATRIX(j + 1, 2): Next j
    Next i
Else
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKERS_RNG, START_DATE, END_DATE, "d", "DC", False, True, True)
    NROWS = UBound(DATA_MATRIX, 1)
    GoSub LOAD_LINE: GoSub VOLAT_LINE
    If OUTPUT <> 0 Then
        TEMP_MATRIX = MATRIX_GET_SUB_MATRIX_FUNC(TEMP_MATRIX, 2, NROWS, 4, NCOLUMNS + 3)
        TEMP_MATRIX = HISTOGRAM_PERCENTILE_TABLE_FUNC(TEMP_MATRIX)
        TEMP_MATRIX(1, 1) = TICKERS_RNG
        For j = 1 To NCOLUMNS: TEMP_MATRIX(1, j + 1) = "HV d" & PERIODS_VECTOR(1, j): Next j
    End If
End If

ASSETS_HISTORICAL_VOLATILITY_FUNC = TEMP_MATRIX

'----------------------------------------------------------------------------------------------------------
Exit Function
'----------------------------------------------------------------------------------------------------------
LOAD_LINE:
'----------------------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS + 3)
    TEMP_MATRIX(0, 1) = "DATE"
    TEMP_MATRIX(0, 2) = "PRICE"
    TEMP_MATRIX(0, 3) = "LOG-RETURN"
    For j = 1 To NCOLUMNS: TEMP_MATRIX(0, 3 + j) = "HV d" & PERIODS_VECTOR(1, j): Next j
    i = 1
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
    TEMP_MATRIX(i, 3) = ""
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, 3 + j) = "": Next j
    For i = 2 To NROWS
        TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
        TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
        TEMP_MATRIX(i, 3) = Log(TEMP_MATRIX(i, 2) / TEMP_MATRIX(i - 1, 2))
    Next i
'----------------------------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------------------------
VOLAT_LINE:
'----------------------------------------------------------------------------------------------------------
    For j = 1 To NCOLUMNS
        h = PERIODS_VECTOR(1, j)
        TEMP1_SUM = 0
        For i = 2 To NROWS
            TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 3)
            If i >= h + 1 Then
                TEMP2_SUM = 0
                For k = i To i - h + 1 Step -1
                    TEMP2_SUM = TEMP2_SUM + (TEMP_MATRIX(k, 3) - (TEMP1_SUM / h)) ^ 2
                Next k
                TEMP_MATRIX(i, 3 + j) = TDAYS_PER_YEAR ^ 0.5 * (TEMP2_SUM / (h - 1)) ^ 0.5
                TEMP1_SUM = TEMP1_SUM - TEMP_MATRIX(i - h + 1, 3)
            Else
                TEMP_MATRIX(i, 3 + j) = ""
            End If
        Next i
    Next j
'----------------------------------------------------------------------------------------------------------
Return
'----------------------------------------------------------------------------------------------------------
ERROR_LABEL:
ASSETS_HISTORICAL_VOLATILITY_FUNC = Err.Number
End Function

Function ASSET_CONDITIONAL_VOLATILITY_FUNC(ByRef DATES_RNG As Variant, _
ByRef DATA_RNG As Variant)

'www.oanda.com x-per-dollar exchange rates
'RiskMetrics volatility for monthly data: Vt+1 = [0.97 Vt^2 + 0.03 st^2]^0.5

Dim i As Long
Dim NROWS As Long

Dim MEAN_VAL As Double

Dim DATA_VECTOR As Variant
Dim DATES_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATES_VECTOR = DATES_RNG
If UBound(DATES_VECTOR, 1) = 1 Then
    DATES_VECTOR = MATRIX_TRANSPOSE_FUNC(DATES_VECTOR)
End If
NROWS = UBound(DATES_VECTOR, 1)

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If NROWS <> UBound(DATA_VECTOR, 1) Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NROWS, 1 To 4)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "DATA" 'St´/$
TEMP_MATRIX(0, 3) = "LOG-RETURN" 'Ln[s1(´/$)/s0(´/$)]
TEMP_MATRIX(0, 4) = "CHV" 'Conditional Historical Volatility

i = 1
TEMP_MATRIX(i, 1) = DATES_VECTOR(i, 1)
TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
TEMP_MATRIX(i, 3) = ""
TEMP_MATRIX(i, 4) = ""

MEAN_VAL = 0
For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = DATES_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = Log(TEMP_MATRIX(i, 2) / TEMP_MATRIX(i - 1, 2))
    MEAN_VAL = MEAN_VAL + TEMP_MATRIX(i, 3)
Next i

MEAN_VAL = (MEAN_VAL / (NROWS - 1))
TEMP_MATRIX(1, 4) = 0
For i = 2 To NROWS: TEMP_MATRIX(1, 4) = TEMP_MATRIX(1, 4) + (TEMP_MATRIX(i, 3) - MEAN_VAL) ^ 2: Next i
TEMP_MATRIX(1, 4) = (TEMP_MATRIX(1, 4) / (NROWS - 2)) ^ 0.5 ' set the starting point to the unconditional s
For i = 2 To NROWS: TEMP_MATRIX(i, 4) = (0.97 * TEMP_MATRIX(i - 1, 4) ^ 2 + 0.03 * TEMP_MATRIX(i, 3) ^ 2) ^ 0.5: Next i
ASSET_CONDITIONAL_VOLATILITY_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_CONDITIONAL_VOLATILITY_FUNC = Err.Number
End Function


Function ASSET_EXTREME_VOLATILITY_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date)

Const OC As Long = 2 'Open Column
Const HC As Long = 3 'High Column
Const LC As Long = 4 'Low Column
Const CC As Long = 5 'Close Column
Const BETA_VAL As Double = 0.601

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long
Dim HL_RETURN As Double
Dim CO_RETURN As Double
Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL


If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DOHLCV", False, False, True)
Else
    DATA_MATRIX = TICKER_STR
End If

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = 6

ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS + 2)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "OPEN"
TEMP_MATRIX(0, 3) = "HIGH"
TEMP_MATRIX(0, 4) = "LOW"
TEMP_MATRIX(0, 5) = "CLOSE"
TEMP_MATRIX(0, 6) = "VOLUME"
TEMP_MATRIX(0, 7) = "HL VOLATILITY" 'Parkinson
TEMP_MATRIX(0, 8) = "OHLC VOLATILITY" 'Garman Klass

For i = 1 To NROWS
    For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    CO_RETURN = Log(TEMP_MATRIX(i, CC) / TEMP_MATRIX(i, OC))
    HL_RETURN = Log(TEMP_MATRIX(i, HC) / TEMP_MATRIX(i, LC))
    TEMP_MATRIX(i, NCOLUMNS + 1) = BETA_VAL * (HL_RETURN ^ 2) ^ 0.5
    TEMP_MATRIX(i, NCOLUMNS + 2) = ((0.5 * (HL_RETURN) ^ 2) - ((2 * Log(2) - 1) * (CO_RETURN) ^ 2)) ^ 0.5
Next i

ASSET_EXTREME_VOLATILITY_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_EXTREME_VOLATILITY_FUNC = Err.Number
End Function



'UCITS IV SYNTHETIC RISK & REWARD INDICATOR (SRRI)
'UCIT IV Synthetic Risk & Reward (SRRI) Template: Volatility measurement and risk class assignment for a given return time series.

Function ASSET_SRRI_FUNC(ByRef DATES_RNG As Variant, _
ByRef RETURNS_RNG As Variant, _
ByRef CLASS_RNG As Variant, _
Optional ByVal MODE_INT As Long = 4, _
Optional ByVal MA_PERIODS As Long = 60, _
Optional ByVal PERIODS_PER_YEAR As Long = 12)
'For detailed explanations, see http://ec.europa.eu/internal_market/investment/docs/legal_texts/framework/091221-methodilogies-1_en.pdf

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim n As Long

Dim NROWS As Long
Dim NCOLUMNS As Long
Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant
Dim DATES_VECTOR  As Variant
Dim RETURNS_VECTOR As Variant
Dim CLASS_VECTOR As Variant 'DEFINITION OF RISK CLASSES
'Risk Class   1     2       3       4       5       6       7
'Lower Bound  0%    0.50%   2%      5%      10%     15%     25%

On Error GoTo ERROR_LABEL

DATES_VECTOR = DATES_RNG
If UBound(DATES_VECTOR, 1) = 1 Then
    DATES_VECTOR = MATRIX_TRANSPOSE_FUNC(DATES_VECTOR)
End If
NROWS = UBound(DATES_VECTOR, 1)

RETURNS_VECTOR = RETURNS_RNG
If UBound(RETURNS_VECTOR, 1) = 1 Then
    RETURNS_VECTOR = MATRIX_TRANSPOSE_FUNC(RETURNS_VECTOR)
End If
If NROWS <> UBound(RETURNS_VECTOR, 1) Then: GoTo ERROR_LABEL

CLASS_VECTOR = CLASS_RNG
If UBound(CLASS_VECTOR, 1) = 1 Then
    CLASS_VECTOR = MATRIX_TRANSPOSE_FUNC(CLASS_VECTOR)
End If
NCOLUMNS = UBound(CLASS_VECTOR, 1)
k = MODE_INT

ReDim TEMP_MATRIX(0 To NROWS, 1 To 6)
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "DATA"
TEMP_MATRIX(0, 3) = "ROLLING VOLATILITY" 'ANNUALIZED VOLATILITES (ROLLING X PERIODS)
TEMP_MATRIX(0, 4) = "MEASURED RISK CLASS"
TEMP_MATRIX(0, 5) = "ASSIGNED RISK CLASS"
TEMP_MATRIX(0, 6) = "CUMULATIVE # CLASS SWITCHES" '(rhs)

TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = DATES_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = RETURNS_VECTOR(i, 1)
    
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 2)
    If i >= MA_PERIODS Then
        h = i - MA_PERIODS + 1
        TEMP_MATRIX(i, 3) = 0
        For j = i To h Step -1
            TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 3) + (TEMP_MATRIX(j, 2) - (TEMP_SUM / MA_PERIODS)) ^ 2
        Next j
        TEMP_MATRIX(i, 3) = PERIODS_PER_YEAR ^ 0.5 * (TEMP_MATRIX(i, 3) / (MA_PERIODS - 1)) ^ 0.5
        TEMP_SUM = TEMP_SUM - TEMP_MATRIX(h, 2)
        
        'RISK CLASS MEASUREMENT & ASSIGNMENT
        TEMP_MATRIX(i, 4) = NCOLUMNS 'No Classes
        For j = 1 To NCOLUMNS - 1
            If TEMP_MATRIX(i, 3) >= CLASS_VECTOR(j, 1) And TEMP_MATRIX(i, 3) < CLASS_VECTOR(j + 1, 1) Then
                TEMP_MATRIX(i, 4) = j
            End If
        Next j
        
        If i < MA_PERIODS + k Then
            If i <> MA_PERIODS Then
                TEMP_MATRIX(i, 5) = TEMP_MATRIX(i - 1, 5)
            Else
                TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 4)
            End If
        Else
            TEMP_MATRIX(i, 5) = ""
            For j = 1 To k
                If TEMP_MATRIX(i - j, 4) = TEMP_MATRIX(i - 1, 5) Then: TEMP_MATRIX(i, 5) = TEMP_MATRIX(i - 1, 5)
            Next j
            If TEMP_MATRIX(i, 5) = "" Then: GoSub MODE_LINE
        End If
        
        If i > MA_PERIODS Then
            TEMP_MATRIX(i, 6) = IIf(TEMP_MATRIX(i, 5) <> TEMP_MATRIX(i - 1, 5), TEMP_MATRIX(i - 1, 6) + 1, TEMP_MATRIX(i - 1, 6))
        Else
            TEMP_MATRIX(i, 6) = 0
        End If
    Else
        TEMP_MATRIX(i, 3) = CVErr(xlErrNA) '""
        TEMP_MATRIX(i, 4) = CVErr(xlErrNA) '""
        TEMP_MATRIX(i, 5) = CVErr(xlErrNA) '""
        TEMP_MATRIX(i, 6) = CVErr(xlErrNA) '""
    End If
Next i

ASSET_SRRI_FUNC = TEMP_MATRIX

Exit Function
'--------------------------------------------------------------------------
MODE_LINE:
'--------------------------------------------------------------------------
    m = 0
    For l = k To 1 Step -1
        n = 0
        For j = i - k To i - 1
            If TEMP_MATRIX(j, 4) = TEMP_MATRIX(i - l, 4) Then: n = n + 1
        Next j
        If n > m Then
            m = n
            TEMP_MATRIX(i, 5) = TEMP_MATRIX(i - l, 4) 'mode
        End If
    Next l
'--------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------
ERROR_LABEL:
ASSET_SRRI_FUNC = Err.Number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_DOWNSIDE_VOLATILITY_FUNC

'DESCRIPTION   : Calculates downside volatilty based on a CASH_RATE

'Semi volatility

'Semi volatility is defined as the volatility of returns below the mean return...
'vs = sqrt[ {1 / N} * sum[ {rs(t) - avr}^2 ] ]
'vs... semi-volatility
'rs(t)... Portfolio returns r(t), for which r(t) < avr
'N... Number of portfolio returns below avr
'avr... average return: avr = sum[rs(t)]/N

'Calculating semi volatility isn't difficult, but a little bit cubersome.
'Semi-volatility cannot be expressed as a Lower Partial Moment. Semi
'volatility is the second moment of a specific part (= the part below
'the mean) of the return distribution.
 
'Downside volatility

'Downside volatility is a generalization of the semi volatility as is
'defined as the volatility of returns below a certain CASH_RATE return...
'vd = sqrt[ {1 / N} * sum[ {rd(t) - avr}^2 ] ]
'vd... downside volatility
'N... number of portfolio returns below a certain CASH_RATE return rt
'rd(t)... portfolio returns r(t), for which r(t) < rt
'avr... average rd: avr = mean[rd(t)] = sum[rd(t)]/N

'Downside volatility cannot be expressed as a lower Lower Partial Moment.
'Semi-volatility is the second moment of a specific part (= the part below
'the CASH_RATE) of the return distribution

'Downside volatility is not the same as downside deviation

'LIBRARY       : ASSET_MOMENTS
'GROUP         : VOLATILITY
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function ASSET_DOWNSIDE_VOLATILITY_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CASH_RATE As Double = 0, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim DATA_VECTOR As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

'---------------------------------------------------------------------------
'-------------------------Lower Partional Moments (LPM)---------------------
'---------------------------------------------------------------------------

'The Lower Partial Moments for discrete data can be defined as...

'LPM(m) = 1/n * sum[ d(t)*{ r(t)-L}^m ]
'LPM(m)... Lower Partial Moment of order m
'n... number of returns
'd(t)... indicator function: d(t) = 0 if r(t) > L, d(t) = 1 if r (t) <= 1
'L... some CASH_RATE
'r(t)... portfolio returns
'm... coefficient determining the shape of the penality function

'Note that m does not have to be an integer.

'LPM is a very general type of risk measure mainly used in academic
'research. A lot of other risk and return measures can be expressed
'as "special cases".

'LPM have the major advantage that they can represent different types
'of utility functions and their characteristics (risk aversion, marginal
'utility of wealth etc.)

'Downside deviation

'One of the better known LPM measures is downside deviation, defined as...

'dd^2 = 1/n * sum[ d(t)*{ r(t) - mar}^2 ]
'dd... downside deviation
'mar.. minimal acceptable return (= a certain CASH_RATE return)
'd(t)... indicator function: d(t) = 0 if r(t) >= mar, d(t) = 1 if r (t) < mar
'n... number of returns
'r(t)... portfolio returns r(t)

'Downside deviation is typically used in the context of the Sortino Ratio.
'Downside deviation can be expressed as a Lower Partial Moment with m = 2
'and L = mar.

'So which is the "correct" downside risk formula? I'm afraid that the one
'and correct solution does not exist. This implies that one should be very
'carful when comparing downside risk figures and other stats based on downside
'risk figures from different sources. Conceptually, the differences between
'the various formulas are rather minor. Personally, I call any downside risk
'measure "correct", as long as it can be expressed as a lower partial moment
'with parameters m and L.

'---------------------------------------------------------------------------

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)

NROWS = UBound(DATA_VECTOR, 1)
If CASH_RATE = 0 Then: CASH_RATE = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)

'-----------------------------------------------------------------------------------------
ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
'-----------------------------------------------------------------------------------------
Select Case VERSION
'-----------------------------------------------------------------------------------------
Case 0 'Calculates semi-downside-volatility
'-----------------------------------------------------------------------------------------
    For i = 1 To NROWS 'Returns Below CASH RATE / MEAN
        If DATA_VECTOR(i, 1) < CASH_RATE Then: TEMP_VECTOR(i, 1) = DATA_VECTOR(i, 1)
    Next i
    If OUTPUT = 0 Then
        ASSET_DOWNSIDE_VOLATILITY_FUNC = MATRIX_STDEV_FUNC(VECTOR_TRIM_FUNC(TEMP_VECTOR, 0))(1, 1)
    Else
        ASSET_DOWNSIDE_VOLATILITY_FUNC = TEMP_VECTOR
    End If
'-----------------------------------------------------------------------------------------
Case Else 'Calculates downside deviation
'-----------------------------------------------------------------------------------------
    For i = 1 To NROWS 'Squared Deviation of Returns Below CASH_RATE from CASH_RATE/Mean
        If DATA_VECTOR(i, 1) < CASH_RATE Then: TEMP_VECTOR(i, 1) = (DATA_VECTOR(i, 1) - CASH_RATE) ^ 2
    Next i
    If OUTPUT = 0 Then
        ASSET_DOWNSIDE_VOLATILITY_FUNC = MATRIX_MEAN_FUNC(VECTOR_TRIM_FUNC(TEMP_VECTOR, 0))(1, 1) ^ 0.5
    ElseIf OUTPUT = 1 Then
        ASSET_DOWNSIDE_VOLATILITY_FUNC = MATRIX_MEAN_FUNC(TEMP_VECTOR)(1, 1) ^ 0.5
    Else
        ASSET_DOWNSIDE_VOLATILITY_FUNC = TEMP_VECTOR
    End If
'-----------------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_DOWNSIDE_VOLATILITY_FUNC = Err.Number
End Function

'Sortino Ratios ... and Sharpe
'First, we talk about the Sharpe Ratio which is an attemt to measure "Risk" and "Reward".
'1. We look at a the average annual return of some asset (say, R = 9%) and compare it to some
'risk-free return (say Rf = 4%, so R - Rf = 5%). Our "Reward" is the excess return: R - Rf.
'2. That excess return may look pretty good, but the asset has volatility (or Standard Deviation,
'say SD = 25%). That 's our "Risk". It measures how far our annual returns deviate from their average.
'In our example, we might expect most returns to vary between R - SD = 9 - 25 = -16% and
'R + SD = 9 + 25 = 34%.
'3. We consider the Sharpe Ratio: (R - Rf) / SD ( In our example, it's: 5/25 = 0.20

'Big reward, small risk ... that's good, eh?
'Yes, so we might look for assets with large Sharpe Ratios.
'However, if the returns get really large compared to Rf, we should be happy.
'But that means the deviations from the average can be large, and that'd make SD larger and that'd
'make the Sharpe Ratio smaller. Then don't buy the asset! but you 'd want lots of returns larger
'than Rf, wouldn't you? Then buy it! Ay , there 's the rub.

'If we really wanted to use deviations from the average as a measure of "risk", then why include
'returns that are larger than Rf? We could just measure the deviation of those returns that are
'smaller than Rf, eh? That means we could take the standard deviation of those less-than-Rf
'returns. SDd2 = (1/n) ? (r - Rf)2 where SDd is the "downside volatility" and the sum includes
'only returns for which r < Rf. If we take SDd as our "risk", then we change Sharpe into Sortino:
'Sortino Ratio = (R - Rf) / SDd where R is the Mean annual return,
'Rf  is some risk-free return (or Minimum Acceptable Return: MAR)
'and SDd is the downside risk: SDd2 = (1/n) ? (r - Rf)2
'where the sum (or average) includes only returns for which r < Rf.
'And Sortino is better than Sharpe?
'One thing is for sure: Sortino doesn't penalize an asset for "upside" volatility ... which is a
'"good" volatility. However , there 's a problem with Sortino. It may happen that no returns are
'less than Rf  in which case SDd = 0 and Sortino = ?

'But how do they compare, Sortino and Sharpe?
'First off, notice that the denominator in the Sortino Ratio is smaller than Sharpe's denominator
'... since SDd contains fewer terms.
'That makes the Sortino Ratio larger, as in Figure 1, and ...
'i 'll bet fund managers like that! Smaller "risk", eh?
'Yes, and I understand that Sortino is (usually) used by hedge funds.
    
'Well ... I used ten years worth of monthly returns and used (as Rf) 1/12 of the annual risk-free
'rate (or MAR).

'Remember when I said that Sortino's "downside" volatility is smaller than the regular,
'garden-variety standard deviation used by Sharpe? You get a glimpse of that in Figure 2. See?

'calculating Sortino Ratios
'Okay, so here's what I did with the spreadsheet:
'I used monthly returns instead of annual returns ... since there are so many more of those available
'I multiplied the monthly returns by 12 to simulate annual returns.
'Then I calculated the Sortino and Sharpe ratios, using these "multiplied-by-12" returns.
'Then I noticed that the ratios didn't change by this "multiply-by-12" ritual.

'uppose we multiply all our returns by some positive parameter ? (for example, ? = 12).
'Note that SDd2, which is (1/n) ? (r - Rf)2, becomes (1/n) ? (?r -?Rf)2 = ?2(1/n) ? (r -Rf)2.
'Hence the downside volatility SDd gets multiplied by ? as well.
'(Indeed, any standard deviation gets multiplied by that parameter ... which is one reason I don't
'like standard deviation as a measure of "risk".)

'That 'd change (R - Rf) / SDd   into   (?R - ?Rf) / ?SDd = (R - Rf) / SDd.
'In other words, the Sortino Ratio don't change at all !

'Okay, so I just forgot about any effort to change monthly to annual, and just used monthly returns.
'In fact, I wanted to see the variations in ratios in a moving 2-year window ... that's 24 returns
'Note:
'The charts in Fig. 1 and 2 actually use a 1-year window ... but I've changed that.
'That makes it less likely that SDd = 0 as in Fig. 2.
'In fact, Figure 2 now looks like Figure 2a

'One other thing.
'The spreadsheet downloads 10 years worth of monthly returns.
'In my moving 2-year window I need an average for 24 monthly returns. That means I use the first two
'of the 10 years just to get a single 2-year average. That explains why, in the latest spreadsheet,
'the Sortino and Sharpe ratios (in the 2-year moving window) start two years later ... as in Figure 2a.

'And what about years when all returns were greater than Rf? You'd have SDd = 0 and ...
'Yeah, and Sortino = infinity. I know! I know!
'In fact, if you have just a single return less than Rf, then SDd = 0. That's a problem, eh?
'However, in order that SDd = 0 in a 2-year window, you'd need to have 24 consecutive returns all
'less than Rf Since That 's unlikely ...
'But I could choose Rf = 50% then they'd all be less ...
'Yeah, sure ... so pick a more reasonable Rf.

'Almost forgot this, too.
'If you actually downloaded a bunch of monthly returns and wanted to estimate the annual standard
'deviation, it's more complicated than simply multipliying all returns by 12 and getting a standard
'deviation multiplied by 12. You'd multiply the standard deviation by SQRT(12) !!

'That means (R - Rf) / SD will have the numerator multiplied by 12, but the denominator multiplied by SQRT(12).
'That means the Sortino and Sharpe ratios get multiplied by SQRT(12) ... since 12/SQRT(12) = SQRT(12).

'The spreadsheet will do that if'n you ask nice. Just say y:
'It will give ratios which are more like the ones you might see in the literature or on the internet
'(even tho' the increased values don't change the look of the charts).

'Didn't you say you did not like standard deviation as a measure of risk?
'Yes, because I can add 10% to all returns and the standard deviation doesn't change which makes it a
'lousy measure ...

'Then why don't you invent your own Sortino ratio, with your own measure of risk? You could call it ...
'uh, the Ponzorino Ratio. Hmmm ... not a bad idea:

'http://www.gummy-stuff.org/sortino.htm
'http://www.sortino.com/htm/biograph.htm
'http://www.gummy-stuff.org/VaR.htm

Function ASSET_SORTINO_SHARPE_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal MA_PERIOD As Long = 24, _
Optional ByVal CASH_RATE As Double = 0.04, _
Optional ByVal COUNT_BASIS As Double = 12, _
Optional ByVal ANNUALIZED_FLAG As Boolean = False)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long

Dim TEMP_VAL As Variant
Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double
Dim TEMP4_SUM As Double

Dim FACTOR_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, _
                  "d", "DA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

CASH_RATE = CASH_RATE / COUNT_BASIS
If ANNUALIZED_FLAG = True Then
    FACTOR_VAL = Sqr(COUNT_BASIS)
Else
    FACTOR_VAL = 1
End If

'-----------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To NROWS, 1 To 11)
'-----------------------------------------------------------------------------
TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PRICE"
TEMP_MATRIX(0, 3) = "RETURNS"
TEMP_MATRIX(0, 4) = "RTN < " & Format(CASH_RATE, "0.00%")

TEMP_MATRIX(0, 5) = "AVG RETURN: " & Format(MA_PERIOD, "0") & " MA PERIOD"

TEMP_MATRIX(0, 6) = "SORTINO RATIO*"
TEMP_MATRIX(0, 7) = "SHARPE RATIO*"
TEMP_MATRIX(0, 8) = "PONZORINO*"

TEMP_MATRIX(0, 9) = "DOWNSIDE SD: " & Format(MA_PERIOD, "0") & " MA PERIOD"
TEMP_MATRIX(0, 10) = "REGULAR SD: " & Format(MA_PERIOD, "0") & " MA PERIOD"
TEMP_MATRIX(0, 11) = "NORMALIZED SD: " & Format(MA_PERIOD, "0") & " MA PERIOD"

i = 1
TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
For j = 3 To 11: TEMP_MATRIX(i, j) = "": Next j

l = 0
TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 2) / TEMP_MATRIX(i - 1, 2) - 1
    TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(i, 3)
    
    If TEMP_MATRIX(i, 3) < CASH_RATE Then
        TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 3) - CASH_RATE
        TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 4)
        l = l + 1
    Else
        TEMP_MATRIX(i, 4) = ""
    End If
    
    If i >= MA_PERIOD + 1 Then
        k = i - MA_PERIOD + 1
        
        MEAN_VAL = TEMP1_SUM / MA_PERIOD
        TEMP_MATRIX(i, 5) = MEAN_VAL
        TEMP1_SUM = TEMP1_SUM - TEMP_MATRIX(k, 3)
        
        TEMP3_SUM = 0
        For j = i To k Step -1
            TEMP3_SUM = TEMP3_SUM + (TEMP_MATRIX(j, 3) - MEAN_VAL) ^ 2
        Next j
        SIGMA_VAL = (TEMP3_SUM / MA_PERIOD) ^ 0.5
        TEMP_MATRIX(i, 10) = SIGMA_VAL
        TEMP_MATRIX(i, 11) = NORMSDIST_FUNC(CASH_RATE, MEAN_VAL, SIGMA_VAL, 0)
                
        If TEMP_MATRIX(i, 10) <> 0 Then
            TEMP_MATRIX(i, 7) = FACTOR_VAL * (TEMP_MATRIX(i, 5) - CASH_RATE) / TEMP_MATRIX(i, 10)
        Else
            TEMP_MATRIX(i, 7) = 0
        End If
        
        If TEMP_MATRIX(i, 11) <> 0 Then
            TEMP_MATRIX(i, 8) = FACTOR_VAL * (TEMP_MATRIX(i, 5) - CASH_RATE) / TEMP_MATRIX(i, 11)
        Else
            TEMP_MATRIX(i, 8) = 0
        End If
        
        If l > 0 Then
            MEAN_VAL = TEMP2_SUM / l
            
            TEMP4_SUM = 0
            For j = i To k Step -1
                TEMP_VAL = TEMP_MATRIX(j, 4)
                If TEMP_VAL <> "" Then
                    TEMP4_SUM = TEMP4_SUM + (TEMP_VAL - MEAN_VAL) ^ 2
                End If
            Next j

            SIGMA_VAL = (TEMP4_SUM / l) ^ 0.5
            TEMP_MATRIX(i, 9) = SIGMA_VAL
            If TEMP_MATRIX(i, 9) <> 0 Then
                TEMP_MATRIX(i, 6) = FACTOR_VAL * (TEMP_MATRIX(i, 5) - CASH_RATE) / TEMP_MATRIX(i, 9)
            Else
                TEMP_MATRIX(i, 6) = 0
            End If
            If TEMP_MATRIX(k, 4) <> "" Then
                TEMP2_SUM = TEMP2_SUM - TEMP_MATRIX(k, 4)
                l = l - 1
            End If
        Else
            TEMP_MATRIX(i, 9) = 0
            TEMP_MATRIX(i, 6) = 0
        End If
    Else
        For j = 5 To 11: TEMP_MATRIX(i, j) = "": Next j
    End If
        
Next i

ASSET_SORTINO_SHARPE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_SORTINO_SHARPE_FUNC = Err.Number
End Function


'-----------------------------------------------------------------------------------------------------------
'Moving Return, Volatility and CAGR
'Reference: http://www.gummy-stuff.org/moving-CAGR.htm
'-----------------------------------------------------------------------------------------------------------

'Here is an interesting notion. You take a gander at stock prices over, say, the last 10 years.
'You notice that they went from P0 (10 years ago) to P1 (today). If the Compound Annual Growth Rate
'is CAGR, then that means: CAGR = (P1 / P0) ^ (1/10) -1
'That 's because: P0 (1 + CAGR)^ 10 = P1.

'But suppose you'd like to see how that CAGR has changed over the years.
'In fact, let's look at 3-year periods, where the stock price went from P0 to P1 over those 3 years.
'Then the CAGR would be: (P1 / P0) ^ (1/3) -1

'The following function downloads N years worth of data
'Stare in awe at the variation in the CAGR (calculated over a moving N-months period (36 = 3 years).

'The volatility (e.g., Standard Deviation) is based upon N years worth of monthly returns, multiplied by
'SQRT(12) to annualize. (See square root stuff.)

'Pay attention to the Volatility, it may go in a direction opposite to the CAGR (Check the Correlation)
'Since "Volatility" is a measure of stock return deviations from their average return, we might expect that
'unusually large deviations (either up or down) would result in a large Standard Deviation (or Volatility).

'Example, during times when the returns didn't stray much from their 3-year (36 months) average, the
'volatility may NOT vary a great deal. But when the returns starts to drop like a rock, deviations wil
'skyrock and volatility will from about X% to over X++++++%.

'-----------------------------------------------------------------------------------------------------------

Function ASSET_MOVING_RETURN_VOLATILITY_CAGR_FUNC(ByVal TICKER_STR As Variant, _
Optional ByVal START_DATE As Date, _
Optional ByVal END_DATE As Date, _
Optional ByVal NO_MONTHS As Long = 36)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MEAN_SUM As Double
Dim STDEVP_SUM As Double

Dim TEMP2_SUM As Double

Dim HEADINGS_STR As String

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKER_STR) = False Then
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "MONTHLY", "DOHLCVA", False, True, True)
Else
    DATA_MATRIX = TICKER_STR
End If
NROWS = UBound(DATA_MATRIX, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 11)

'-----------------------------------------------------------------------------------------------------
HEADINGS_STR = "DATE,OPEN,HIGH,LOW,CLOSE,VOLUME,ADJ.CLOSE,RETURN,MEAN*,VOLATILITY*,CAGR,"
j = Len(HEADINGS_STR)
NCOLUMNS = 0
For i = 1 To j
    If Mid(HEADINGS_STR, i, 1) = "," Then: NCOLUMNS = NCOLUMNS + 1
Next i
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
i = 1
For k = 1 To NCOLUMNS
    j = InStr(i, HEADINGS_STR, ",")
    TEMP_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
    i = j + 1
Next k
'-----------------------------------------------------------------------------------------------------

l = NO_MONTHS + 1: m = (NO_MONTHS / 12)
i = 1
For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 5) / TEMP_MATRIX(i, 2) - 1
MEAN_SUM = TEMP_MATRIX(i, 8)
STDEVP_SUM = MEAN_SUM * MEAN_SUM

TEMP_MATRIX(i, 9) = "": TEMP_MATRIX(i, 10) = "": TEMP_MATRIX(i, 11) = ""
For i = 2 To NROWS
    For j = 1 To 7: TEMP_MATRIX(i, j) = DATA_MATRIX(i, j): Next j
    TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 6) / 1000
    TEMP_MATRIX(i, 8) = TEMP_MATRIX(i, 7) / TEMP_MATRIX(i - 1, 7) - 1
    MEAN_SUM = MEAN_SUM + TEMP_MATRIX(i, 8)
    STDEVP_SUM = STDEVP_SUM + TEMP_MATRIX(i, 8) * TEMP_MATRIX(i, 8)
    If i <= l Then
        TEMP_MATRIX(i, 9) = MEAN_SUM / i
        'Some people asked if variance could actually be calculated in a single pass.
        'Yes, see the links below for one pass algorithms:
        'http://en.wikipedia.org/wiki/Algorithms_for_calculating_variance
        TEMP_MATRIX(i, 10) = ((STDEVP_SUM - ((MEAN_SUM * MEAN_SUM) / l)) / l) ^ 0.5 '(l-1)
        k = 1
    Else
        MEAN_SUM = MEAN_SUM - TEMP_MATRIX(i - l, 8)
        STDEVP_SUM = STDEVP_SUM - TEMP_MATRIX(i - l, 8) * TEMP_MATRIX(i - l, 8)
        TEMP_MATRIX(i, 9) = MEAN_SUM / l
        TEMP_MATRIX(i, 10) = ((STDEVP_SUM - ((MEAN_SUM * MEAN_SUM) / l)) / l) ^ 0.5 '(l-1)
        k = k + 1
    End If
    TEMP_MATRIX(i, 11) = (TEMP_MATRIX(i, 7) / TEMP_MATRIX(k, 7)) ^ (1 / m) - 1
 '   TEMP2_SUM = 0
  '  For j = i To k Step -1
   '     TEMP2_SUM = TEMP2_SUM + (TEMP_MATRIX(j, 8) - TEMP_MATRIX(i, 9)) ^ 2
    'Next j
'    TEMP_MATRIX(i, 10) = 12 ^ 0.5 * (TEMP2_SUM / (i - k + 1)) ^ 0.5
    TEMP_MATRIX(i, 9) = TEMP_MATRIX(i, 9) * 12
    TEMP_MATRIX(i, 10) = 12 ^ 0.5 * TEMP_MATRIX(i, 10)
Next i

ASSET_MOVING_RETURN_VOLATILITY_CAGR_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_MOVING_RETURN_VOLATILITY_CAGR_FUNC = Err.Number
End Function