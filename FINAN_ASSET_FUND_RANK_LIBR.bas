Attribute VB_Name = "FINAN_ASSET_FUND_RANK_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Written for the Ben Graham Centre for Value Investing

Function MUTUAL_FUNDS_RANKING_SCORES_FUNC(ByRef TICKERS_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long

Const ERROR_STR As String = "--"

Dim LOAD_STR As String
Dim TICKER_STR As String

Dim RETURN_RATING_STR As String
Dim RISK_RATING_STR As String
Dim MORNINGSTAR_RATING_VAL As Integer
Dim YEAR3_SP_VAL As Double
Dim YEAR5_SP_VAL As Double
Dim VOLATILITY_VAL As Double
Dim PE_VAL As Double
Dim TURNOVER_VAL As Double
Dim EXPENSE_RATIO_VAL As Double
Dim BEAR_MARKET_FUND_PERCENT_VAL As Double
Dim BEAR_MARKET_INDEX_PERCENT_VAL As Double

Dim TEMP_VAL As Variant
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
NROWS = UBound(TICKERS_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 17)

TEMP_MATRIX(0, 1) = "SYMBOL"
TEMP_MATRIX(0, 2) = "NAME"
TEMP_MATRIX(0, 3) = "INCEPTION DATE"

TEMP_MATRIX(0, 4) = "MORNINGSTAR OVERALL RATING"
TEMP_MATRIX(0, 5) = "LOAD"
TEMP_MATRIX(0, 6) = "RETURN RATING"
TEMP_MATRIX(0, 7) = "RISK RATING"

TEMP_MATRIX(0, 8) = "3 YEAR-S&P"
TEMP_MATRIX(0, 9) = "5 YEAR-S&P"
TEMP_MATRIX(0, 10) = "STANDARD DEV"
TEMP_MATRIX(0, 11) = "P/E"
TEMP_MATRIX(0, 12) = "TURNOVER"
TEMP_MATRIX(0, 13) = "TOTAL EXP"
TEMP_MATRIX(0, 14) = "BEAR MKT FUND"
TEMP_MATRIX(0, 15) = "BEAR MKT INDEX"
TEMP_MATRIX(0, 16) = "SCORE-D"
TEMP_MATRIX(0, 17) = "SCORE-%"

For i = 1 To NROWS
    k = 1: j = 1
    TICKER_STR = TICKERS_VECTOR(i, 1)
    If TICKER_STR = "" Then: GoTo 1983
    TEMP_MATRIX(i, j) = TICKER_STR
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 4811, ERROR_STR) 'Name
    j = j + 1
    If TEMP_VAL <> ERROR_STR Then
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = ERROR_STR
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 5065, ERROR_STR) 'Inception Date
    j = j + 1
    If TEMP_VAL <> ERROR_STR Then
        TEMP_MATRIX(i, j) = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = ERROR_STR
    End If
    
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 5064, ERROR_STR) 'MORNINGSTAR_RATING_VAL
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
        MORNINGSTAR_RATING_VAL = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = ERROR_STR
        k = -1
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 4939, ERROR_STR) 'LOAD_STR
    j = j + 1
    If TEMP_VAL <> ERROR_STR Then
        TEMP_MATRIX(i, j) = TEMP_VAL
        LOAD_STR = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = ERROR_STR
        k = -1
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 4930, ERROR_STR) 'RETURN_RATING_STR
    j = j + 1
    If TEMP_VAL <> ERROR_STR Then
        TEMP_MATRIX(i, j) = TEMP_VAL
        RETURN_RATING_STR = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = ERROR_STR
        k = -1
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 5093, ERROR_STR) 'RISK_RATING_STR
    j = j + 1
    If TEMP_VAL <> ERROR_STR Then
        TEMP_MATRIX(i, j) = TEMP_VAL
        RISK_RATING_STR = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = ERROR_STR
        k = -1
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 4979, ERROR_STR) 'YEAR3_SP_VAL
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
        YEAR3_SP_VAL = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = ERROR_STR
        k = -1
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 4980, ERROR_STR) 'YEAR5_SP_VAL
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
        YEAR5_SP_VAL = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = ERROR_STR
        k = -1
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 4671, ERROR_STR) 'VOLATILITY_VAL
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
        VOLATILITY_VAL = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = ERROR_STR
        k = -1
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 4681, ERROR_STR) 'PE_VAL
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
        PE_VAL = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = ERROR_STR
        k = -1
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 5077, ERROR_STR) 'TURNOVER_VAL
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
        TURNOVER_VAL = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = ERROR_STR
        k = -1
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 5086, ERROR_STR) 'EXPENSE_RATIO_VAL
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
        EXPENSE_RATIO_VAL = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = ERROR_STR
        k = -1
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 4950, ERROR_STR) 'BEAR_MARKET_FUND_PERCENT_VAL
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
        BEAR_MARKET_FUND_PERCENT_VAL = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = ERROR_STR
        k = -1
    End If
    
    TEMP_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, 4974, ERROR_STR) 'BEAR_MARKET_INDEX_PERCENT_VAL
    j = j + 1
    If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
        TEMP_MATRIX(i, j) = TEMP_VAL
        BEAR_MARKET_INDEX_PERCENT_VAL = TEMP_VAL
    Else
        TEMP_MATRIX(i, j) = ERROR_STR
        k = -1
    End If
    
    If k <> -1 Then
        TEMP_VAL = MUTUAL_FUND_RANKING_SCORE_FUNC(MORNINGSTAR_RATING_VAL, _
                   LOAD_STR, _
                   RETURN_RATING_STR, _
                   RISK_RATING_STR, _
                   YEAR3_SP_VAL, _
                   YEAR5_SP_VAL, _
                   VOLATILITY_VAL, _
                   PE_VAL, _
                   TURNOVER_VAL, _
                   EXPENSE_RATIO_VAL, _
                   BEAR_MARKET_FUND_PERCENT_VAL, _
                   BEAR_MARKET_INDEX_PERCENT_VAL, _
                   ERROR_STR) * 15
        If TEMP_VAL <> ERROR_STR And IsNumeric(TEMP_VAL) Then
            TEMP_MATRIX(i, 16) = TEMP_VAL
            TEMP_MATRIX(i, 17) = TEMP_MATRIX(i, 16) / 15
        Else
            TEMP_MATRIX(i, 16) = ERROR_STR
            TEMP_MATRIX(i, 17) = ERROR_STR
        End If
    Else
        TEMP_MATRIX(i, 16) = ERROR_STR
        TEMP_MATRIX(i, 17) = ERROR_STR
    End If
1983:
Next i

MUTUAL_FUNDS_RANKING_SCORES_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MUTUAL_FUNDS_RANKING_SCORES_FUNC = Err.number
End Function

Function MUTUAL_FUND_RANKING_SCORE_FUNC( _
ByVal MORNINGSTAR_RATING_VAL As Integer, _
ByVal LOAD_STR As String, _
ByVal RETURN_RATING_STR As String, _
ByVal RISK_RATING_STR As String, _
ByVal YEAR3_SP_VAL As Double, _
ByVal YEAR5_SP_VAL As Double, _
ByVal VOLATILITY_VAL As Double, _
ByVal PE_VAL As Double, _
ByVal TURNOVER_VAL As Double, _
ByVal EXPENSE_RATIO_VAL As Double, _
ByVal BEAR_MARKET_FUND_PERCENT_VAL As Double, _
ByVal BEAR_MARKET_INDEX_PERCENT_VAL As Double, _
Optional ByVal ERROR_STR As String = "--")

Dim i As Long
Dim j As Double
Dim TEMP_VAL As Variant

On Error GoTo ERROR_LABEL

'--------------------------------------------------------------
'Morningstar Rating: >=4
TEMP_VAL = 4
i = IIf(MORNINGSTAR_RATING_VAL >= TEMP_VAL, 1, 0)
'--------------------------------------------------------------
'Load: "No load"
TEMP_VAL = "No load"
i = i + IIf(LOAD_STR = TEMP_VAL, 1, 0)
'--------------------------------------------------------------
'Risk/Return Rating: Average, Above Average or High
GoSub RATING_LINE
'--------------------------------------------------------------
'3Year to S&P: ">=0.9
TEMP_VAL = 0.9
i = i + IIf(YEAR3_SP_VAL >= TEMP_VAL, 1, 0)
'--------------------------------------------------------------
'5Year to S&P: ">=0.9
TEMP_VAL = 0.9
i = i + IIf(YEAR5_SP_VAL >= TEMP_VAL, 1, 0)
'--------------------------------------------------------------
'Standard Deviation: < 20
TEMP_VAL = 20
i = i + IIf(VOLATILITY_VAL < TEMP_VAL, 1, 0)
'--------------------------------------------------------------
'Price/Earnings: < 25
TEMP_VAL = 25
i = i + IIf(PE_VAL < TEMP_VAL, 1, 0)
'--------------------------------------------------------------
'Turnover: < 50%
TEMP_VAL = 0.5
j = IIf(TURNOVER_VAL > 0, TURNOVER_VAL, 0)
i = i + IIf(j < TEMP_VAL, 1, 0)
'--------------------------------------------------------------
'Total Expense ratio: < 1.75%
TEMP_VAL = 0.0175
i = i + IIf(EXPENSE_RATIO_VAL < TEMP_VAL, 1, 0)
'--------------------------------------------------------------
'Bear Market %: < Index
i = i + IIf(BEAR_MARKET_FUND_PERCENT_VAL > BEAR_MARKET_INDEX_PERCENT_VAL, 1, 0)
'--------------------------------------------------------------

MUTUAL_FUND_RANKING_SCORE_FUNC = i / 15 'Max Score 9 + 3x2 = 15

Exit Function
'----------------------------------------------------------
RATING_LINE:
'----------------------------------------------------------
    Select Case LCase(Trim(RISK_RATING_STR))
    Case "high": i = i + -1
    Case "above average": i = i + 0
    Case "average": i = i + 1
    Case "below average": i = i + 2
    Case "low": i = i + 3
    End Select
    
    Select Case LCase(Trim(RETURN_RATING_STR))
    Case "high": i = i + 3
    Case "above average": i = i + 2
    Case "average": i = i + 1
    Case "below average": i = i + 0
    Case "low": i = i + -1
    End Select
'----------------------------------------------------------
Return
'----------------------------------------------------------
ERROR_LABEL:
MUTUAL_FUND_RANKING_SCORE_FUNC = ERROR_STR
End Function

'----------------------------------------------------------------------------------
'Return:
'----------------------------------------------------------------------------------
'3-Year Return -->
'http://quicktake.morningstar.com/FundNet/RatingsAndRisk.aspx?symbol=mchfx
'>3-Year

'5-Year Return -->
'http://quicktake.morningstar.com/FundNet/RatingsAndRisk.aspx?symbol=mchfx
'>5-Year

'Overall Return -->
'http://quicktake.morningstar.com/FundNet/RatingsAndRisk.aspx?symbol=mchfx
'>Overall
'----------------------------------------------------------------------------------
'Risk:
'----------------------------------------------------------------------------------
'3-Year Risk -->
'http://quicktake.morningstar.com/FundNet/RatingsAndRisk.aspx?symbol=mchfx
'>3-Year

'5-Year Risk -->
'http://quicktake.morningstar.com/FundNet/RatingsAndRisk.aspx?symbol=mchfx
'>5-Year

'Overall Risk -->
'http://quicktake.morningstar.com/FundNet/RatingsAndRisk.aspx?symbol=mchfx
'>Overall
'----------------------------------------------------------------------------------
