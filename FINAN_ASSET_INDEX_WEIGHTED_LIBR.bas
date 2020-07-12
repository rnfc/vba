Attribute VB_Name = "FINAN_ASSET_INDEX_WEIGHTED_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'/////////////////////////////////////////////////////////////////////////////////////////////
'rough and ready "financial velocity" index, which is the ratio of an index of financial
'markets (the SPX, 10 year Treasury futures, EUR/USD, gold, and oil, all equally weighted)
'/////////////////////////////////////////////////////////////////////////////////////////////
'Well, it'd be neat to be able to create my own index which covers, say, gold stocks
'or car companies or retail stores etc. etc. So I figured I could download a gaggle of
'stock prices, weight them according to market caps and generate an index.

'No, the DOW is proportional to the average price of 30 stocks. Bigger stock prices
'have more weight. It 'd be more like the S&P where bigger companies (measured by the
'total value of all stocks) have more weight.

'I picked a dozen which have mkt caps over (about) $1B. In fact, them "weights" are the
'mkt caps, in billions - BTU being the largest. Note that BTU has a very high weighting.
'It turns out that the Index acts much like BTU.

'A Year 's worth of prices are downloaded and plotted and an Index is calculated based
'upon the dozen market capitalizations. There 's a separate sheet that'll do that. It
'downloads the market caps and generates the various weights.

'Judging from the coal index chart, compared to the DOW, coal stocks don't look good.
'That 's over the past year, but check out the past six months:

'http://www.gummy-stuff.org/coal-index.htm
'/////////////////////////////////////////////////////////////////////////////////////////////

Function ASSETS_WEIGHTED_CAP_INDEX_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal SROW As Long = 1, _
Optional ByVal MARKET_CAP_DENOM As Double = 1000000, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim WEIGHTED_PRICE As Double
Dim WEIGHTS_SUM As Double
Dim TICKERS_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim WEIGHTS_VECTOR As Variant

On Error GoTo ERROR_LABEL

TICKERS_VECTOR = TICKERS_RNG
If UBound(TICKERS_VECTOR, 1) = 1 Then
    TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
End If

DATA_MATRIX = _
YAHOO_HISTORICAL_DATA_SERIES1_FUNC(TICKERS_VECTOR, START_DATE, END_DATE, 6, "d", True, True)
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2) - 1 'Exclude Dates Vector

WEIGHTS_VECTOR = YAHOO_QUOTES_FUNC(TICKERS_VECTOR, "market capitalization", 0, False, "")
For j = 1 To NCOLUMNS: WEIGHTS_VECTOR(j, 1) = WEIGHTS_VECTOR(j, 1) / MARKET_CAP_DENOM: Next j

'---------------------------------------------------------------------------------------------------
Select Case OUTPUT
'---------------------------------------------------------------------------------------------------
Case 0 'Index Return Vector
'---------------------------------------------------------------------------------------------------
    WEIGHTS_SUM = 0
    For j = 1 To NCOLUMNS: WEIGHTS_SUM = WEIGHTS_SUM + WEIGHTS_VECTOR(j, 1): Next j
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 2)
    TEMP_MATRIX(0, 1) = "DATE"
    TEMP_MATRIX(0, 2) = "INDEX"
    For i = 1 To NROWS
        WEIGHTED_PRICE = 0
        For j = 1 To NCOLUMNS
            WEIGHTED_PRICE = WEIGHTED_PRICE + (WEIGHTS_VECTOR(j, 1) * _
                             DATA_MATRIX(i, j + 1) / DATA_MATRIX(SROW, j + 1))
        Next j
        TEMP_MATRIX(i, 2) = WEIGHTED_PRICE / WEIGHTS_SUM - 1
        TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    Next i
'---------------------------------------------------------------------------------------------------
Case Else 'Weighted Prices
'---------------------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS + 1)
    For j = 1 To NCOLUMNS + 1: TEMP_MATRIX(0, j) = DATA_MATRIX(0, j): Next j
    For i = 1 To NROWS
        TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
        For j = 1 To NCOLUMNS
            TEMP_MATRIX(i, j + 1) = (WEIGHTS_VECTOR(j, 1) * DATA_MATRIX(i, j + 1) / _
                                     DATA_MATRIX(SROW, j + 1))
        Next j
    Next i
'---------------------------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------------------------

ASSETS_WEIGHTED_CAP_INDEX_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_WEIGHTED_CAP_INDEX_FUNC = Err.number
End Function

'The GIndex
'Okay, so we'd like to buy the 30 DOW stocks.
'Not knowing which stocks should be overweighted, surely we'd invest equal amounts in each, right?
'No. He figures: the smaller the mkt cap, the more he'd invest.
'So I interpret that to mean the amount invested is inversely proportional to the mkt cap.
'So I whip out a ... Exactly ... and, using the DOW stocks over the past year, I get this
'It 's better than your g-Index !!

'Note that, for the past year, the largest weights are assigned to Alcoa (AA), DuPont (DD) and
'Travelers (TRV). Surprise! They were all up about 20% over the past year.
'On the other hand, some of the big guys, like Exxon (XOM) and Walmart (WMT), had negative
'returns. So, we may think allocations inversely proportional to the Mkt Cap is too drastic.
'This routine has an allocation proportional to MktCapP   and you get to pick P.
'It can also search for the best P in a range. Turns out that P = -1.24 is "best". That is, allocation
'should be proportional to 1 / MktCap1.24.

Function ASSETS_WEIGHTED_CAP_GINDEX_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByVal POWER_RNG As Variant = -2, _
Optional ByVal MARKET_CAP_DENOM As Double = 1000000)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim POWER_VAL As Double
Dim WEIGHTS_SUM As Double
Dim WEIGHTED_RETURN As Double

Dim TEMP_MATRIX As Variant
Dim POWER_VECTOR As Variant
Dim TICKERS_VECTOR As Variant
Dim WEIGHTS_VECTOR As Variant

Dim DATA_MATRIX As Variant
Dim MCAP_VECTOR As Variant

On Error GoTo ERROR_LABEL

TICKERS_VECTOR = TICKERS_RNG
If UBound(TICKERS_VECTOR, 1) = 1 Then
    TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
End If

DATA_MATRIX = _
YAHOO_HISTORICAL_DATA_SERIES1_FUNC(TICKERS_VECTOR, START_DATE, END_DATE, 6, "d", True, True)
NROWS = UBound(DATA_MATRIX, 1)

MCAP_VECTOR = YAHOO_QUOTES_FUNC(TICKERS_VECTOR, "market capitalization", 0, False, "")
For j = 1 To UBound(MCAP_VECTOR, 1) - 1: MCAP_VECTOR(j, 1) = MCAP_VECTOR(j + 1, 1) / MARKET_CAP_DENOM: Next j

'------------------------------------------------------------------------------------------------------------
If IsArray(POWER_RNG) = True Then
'------------------------------------------------------------------------------------------------------------
    POWER_VECTOR = POWER_RNG
    If UBound(POWER_VECTOR, 1) = 1 Then
        POWER_VECTOR = MATRIX_TRANSPOSE_FUNC(POWER_VECTOR)
    End If
    l = UBound(POWER_VECTOR, 1)
    ReDim TEMP_MATRIX(0 To l, 1 To 2)
    TEMP_MATRIX(0, 1) = "POWER VALUE"
    TEMP_MATRIX(0, 2) = "(GINDEX)-(" & UCase(TICKERS_VECTOR(1, 1)) & ")"
    For k = 1 To l
        POWER_VAL = POWER_VECTOR(k, 1)
        GoSub WEIGHTS_LINE
        For i = 2 To NROWS
            WEIGHTED_RETURN = 0
            For j = 1 To NCOLUMNS
                WEIGHTED_RETURN = WEIGHTED_RETURN + (WEIGHTS_VECTOR(j, 1) * (DATA_MATRIX(i, j + 2) / DATA_MATRIX(1, j + 2) - 1))
            Next j
        Next i
        TEMP_MATRIX(k, 1) = POWER_VAL
        TEMP_MATRIX(k, 2) = WEIGHTED_RETURN - (DATA_MATRIX(NROWS, 2) / DATA_MATRIX(1, 2) - 1)
    Next k
'------------------------------------------------------------------------------------------------------------
Else
'------------------------------------------------------------------------------------------------------------
    POWER_VAL = POWER_RNG
    GoSub WEIGHTS_LINE
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 3)
    TEMP_MATRIX(0, 1) = "DATE"
    TEMP_MATRIX(0, 2) = UCase(TICKERS_VECTOR(1, 1))
    TEMP_MATRIX(0, 3) = "GINDEX"
    
    i = 1
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
    TEMP_MATRIX(i, 2) = 0
    TEMP_MATRIX(i, 3) = 0
    For i = 2 To NROWS
        TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
        TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2) / DATA_MATRIX(1, 2) - 1
        For j = 1 To NCOLUMNS
            TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 3) + (WEIGHTS_VECTOR(j, 1) * (DATA_MATRIX(i, j + 2) / DATA_MATRIX(1, j + 2) - 1))
        Next j
    Next i
'------------------------------------------------------------------------------------------------------------
End If
'------------------------------------------------------------------------------------------------------------
ASSETS_WEIGHTED_CAP_GINDEX_FUNC = TEMP_MATRIX

'------------------------------------------------------------------------------------------------------
Exit Function
'------------------------------------------------------------------------------------------------------
WEIGHTS_LINE:
'------------------------------------------------------------------------------------------------------
    NCOLUMNS = UBound(MCAP_VECTOR, 1) - 1 'Exclude Dates & Index Vector
    ReDim WEIGHTS_VECTOR(1 To NCOLUMNS, 1 To 1)
    WEIGHTS_SUM = 0
    For j = 1 To NCOLUMNS
        WEIGHTS_VECTOR(j, 1) = MCAP_VECTOR(j, 1) ^ POWER_VAL
        WEIGHTS_SUM = WEIGHTS_SUM + WEIGHTS_VECTOR(j, 1)
    Next j
    For j = 1 To NCOLUMNS
        WEIGHTS_VECTOR(j, 1) = WEIGHTS_VECTOR(j, 1) / WEIGHTS_SUM
    Next j
'------------------------------------------------------------------------------------------------------
Return
'------------------------------------------------------------------------------------------------------
ERROR_LABEL:
ASSETS_WEIGHTED_CAP_GINDEX_FUNC = Err.number
End Function

'The geometric Index for n stocks is:

'REFERENCES:
'http://www.gummy-stuff.org/Indexes.htm
'http://gummy-stuff.org/Indexes2.htm
'http://www.gummy-stuff.org/Mkt_Indexes.htm
'http://www.gummy-stuff.org/theDOW.htm
'http://www.google.ca/search?hl=en&q="Market+Cap+Weighted+Index"&btnG=Search&meta=

Function ASSETS_WEIGHTED_PRICES_GINDEX_FUNC(ByRef TICKERS_RNG As Variant, _
ByVal START_DATE As Date, _
ByVal END_DATE As Date, _
Optional ByRef WEIGHTS_RNG As Variant, _
Optional ByVal SROW As Long = 1)
'First Entry TICKERS_RNG = Index Symbol

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long
Dim NCOLUMNS As Long
Dim INDEX_ARR() As Long

Dim K_VAL As Double
Dim WEIGHTS_SUM As Double
Dim WEIGHTS_VECTOR As Variant

Dim TICKERS_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

TICKERS_VECTOR = TICKERS_RNG
If UBound(TICKERS_VECTOR, 1) = 1 Then
    TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
End If

'---------------------------------------------------------------------------------------------------------------
If IsArray(WEIGHTS_RNG) = True Then
'---------------------------------------------------------------------------------------------------------------
    NROWS = UBound(TICKERS_VECTOR, 1) - 1 'Exclude Index Symbol
    WEIGHTS_VECTOR = WEIGHTS_RNG
    If UBound(WEIGHTS_VECTOR, 1) = 1 Then
        WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
    End If
    If UBound(WEIGHTS_VECTOR, 1) <> NROWS Then
        ReDim WEIGHTS_VECTOR(1 To NROWS, 1 To 1)
        For i = 1 To NROWS
            WEIGHTS_VECTOR(i, 1) = 1 / NROWS
        Next i
    End If
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIES1_FUNC(TICKERS_VECTOR, START_DATE, END_DATE, 6, "m", True, True)
    NROWS = UBound(DATA_MATRIX, 1)
    NCOLUMNS = UBound(DATA_MATRIX, 2)
    
    WEIGHTS_SUM = 0
    For j = 3 To NCOLUMNS 'exclude dates/index
        WEIGHTS_SUM = WEIGHTS_SUM + Log(DATA_MATRIX(SROW, j)) * WEIGHTS_VECTOR(j - 2, 1)
    Next j
    WEIGHTS_SUM = Exp(WEIGHTS_SUM)
    K_VAL = DATA_MATRIX(SROW, 2) / WEIGHTS_SUM
    
    ReDim TEMP_MATRIX(0 To NROWS, 1 To 3)
    TEMP_MATRIX(0, 1) = "DATES"
    TEMP_MATRIX(0, 2) = UCase(TICKERS_VECTOR(1, 1))
    TEMP_MATRIX(0, 3) = "g(" & TEMP_MATRIX(0, 2) & ")"
    For i = 1 To NROWS
        WEIGHTS_SUM = 0
        For j = 3 To NCOLUMNS 'exclude dates/index
            WEIGHTS_SUM = WEIGHTS_SUM + Log(DATA_MATRIX(i, j)) * WEIGHTS_VECTOR(j - 2, 1)
        Next j
        WEIGHTS_SUM = Exp(WEIGHTS_SUM)
        TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
        TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
        TEMP_MATRIX(i, 3) = WEIGHTS_SUM * K_VAL
    Next i
    ASSETS_WEIGHTED_PRICES_GINDEX_FUNC = TEMP_MATRIX
'---------------------------------------------------------------------------------------------------------------
Else
'Market Indexes
'If we compare the growth (decay?) of the DOW and S&P (over the past 10 years) and the g-Index,
'then ... Well ... uh, that's what I'm calling the Index where you just average the gains (beginning
'at some convenient point in time).

'Like 10 years ago? Yes.Anyway , you 'd get this
'Why would anyone use that g-Index instead of a Price-weighted or MktCap-weighted Index?
'I dunno. Suppose you wanted to invest in the DOW stocks and didn't know which would be the best performer.
'How much would you invest in each stock? Would you be guided by the prices or the mkt caps?

'Knowing nothing about their future behaviour? I think I'd invest equal amounts in each stock.
'Exactly ... and that'd be the g-Index.
'---------------------------------------------------------------------------------------------------------------
    DATA_MATRIX = YAHOO_HISTORICAL_DATA_SERIES1_FUNC(TICKERS_VECTOR, START_DATE, END_DATE, 6, "d", True, True)
    NROWS = UBound(DATA_MATRIX, 1)
    NCOLUMNS = UBound(DATA_MATRIX, 2)
    
    ReDim TEMP_VECTOR(0 To NROWS, 1 To 2) 'G-Index
    TEMP_VECTOR(0, 1) = "DATE"
    TEMP_VECTOR(0, 2) = "G-INDEX"
    
    ReDim INDEX_ARR(1 To NCOLUMNS - 1)
    For j = 2 To NCOLUMNS
        i = 1
        Do While DATA_MATRIX(i, j) = "" Or DATA_MATRIX(i, j) = 0
            i = i + 1
        Loop
        INDEX_ARR(j - 1) = i
    Next j
    
    For i = NROWS To 2 Step -1
        TEMP_VECTOR(i, 1) = DATA_MATRIX(i, 1)
        TEMP_VECTOR(i, 2) = 0
        k = 0
        For j = 2 To NCOLUMNS
            h = INDEX_ARR(j - 1)
            If i <= h Then: GoTo 1983
            If DATA_MATRIX(i, j) <> "" And DATA_MATRIX(i, j) <> 0 Then
                DATA_MATRIX(i, j) = DATA_MATRIX(i, j) / DATA_MATRIX(h, j) - 1
                TEMP_VECTOR(i, 2) = TEMP_VECTOR(i, 2) + DATA_MATRIX(i, j)
                k = k + 1
            End If
1983:
        Next j
        TEMP_VECTOR(i, 2) = TEMP_VECTOR(i, 2) / k 'Exclude Dates Vector
    Next i
    i = 1
    TEMP_VECTOR(i, 1) = DATA_MATRIX(i, 1)
    TEMP_VECTOR(i, 2) = 0
    For j = 2 To NCOLUMNS: DATA_MATRIX(i, j) = 0: Next j
    ASSETS_WEIGHTED_PRICES_GINDEX_FUNC = TEMP_VECTOR 'Array(TEMP_VECTOR, DATA_MATRIX)
'---------------------------------------------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSETS_WEIGHTED_PRICES_GINDEX_FUNC = Err.number
End Function
