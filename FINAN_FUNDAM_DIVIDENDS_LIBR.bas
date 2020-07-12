Attribute VB_Name = "FINAN_FUNDAM_DIVIDENDS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : DDM_EPS_GROWTH_FUNC
'DESCRIPTION   : DIVIDEND DISCOUNT & EPS GROWTH FUNCTION
'LIBRARY       : FUNDAMENTAL
'GROUP         : DIVIDENDS
'ID            : 001
'LAST UPDATE   : 27/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function DDM_EPS_GROWTH_FUNC(ByVal CURRENT_DIVD As Double, _
ByVal CURRENT_EPS As Double, _
ByVal TN_PE_RATIO As Double, _
ByVal TENOR As Double, _
ByVal MIN_EPS_GROWTH As Double, _
ByVal MAX_EPS_GROWTH As Double, _
ByVal DELTA_EPS_GROWTH As Double, _
ByVal MIN_DISCOUNT_RATE As Double, _
ByVal MAX_DISCOUNT_RATE As Double, _
ByVal DELTA_DISCOUNT_RATE As Double)

'TENOR = YEARS or PERIODS
'TN_PE_RATIO = an assumed P/E Ratio, after N = TENOR

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim G_VAL As Double
Dim D_VAL As Double

Dim TEMP_VAL As Double
Dim FACTOR_VAL As Double

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

NROWS = (MAX_DISCOUNT_RATE - MIN_DISCOUNT_RATE) / DELTA_DISCOUNT_RATE + 1
NROWS = NROWS * 4
NCOLUMNS = (MAX_EPS_GROWTH - MIN_EPS_GROWTH) / DELTA_EPS_GROWTH + 2

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

D_VAL = MIN_DISCOUNT_RATE
For i = 1 To NROWS Step 4
    TEMP_MATRIX(i + 0, 1) = "EPS GROWTH:"
    TEMP_MATRIX(i + 1, 1) = "DIVD GROWTH:"
    TEMP_MATRIX(i + 2, 1) = "SHARE PRICE:"
    TEMP_MATRIX(i + 3, 1) = "R = " & Format(D_VAL, "0.00%")
    G_VAL = MIN_EPS_GROWTH
    For j = 2 To NCOLUMNS
        TEMP_VAL = (1 + G_VAL) / (1 + D_VAL)
        If TEMP_VAL <> 1 Then
            FACTOR_VAL = (1 - (TEMP_VAL) ^ TENOR) / (1 - ((1 + G_VAL) / (1 + D_VAL)))
        Else
            FACTOR_VAL = TENOR
        End If
        TEMP_MATRIX(i + 0, j) = G_VAL 'GROWTH_EPS_FACTOR
        TEMP_MATRIX(i + 1, j) = FACTOR_VAL 'MULTIPLE_FACTOR
        TEMP_MATRIX(i + 2, j) = CURRENT_DIVD * TEMP_MATRIX(i + 1, j) + TN_PE_RATIO * CURRENT_EPS * ((1 + G_VAL) / (1 + D_VAL)) ^ TENOR 'SHARE_PRICE
        TEMP_MATRIX(i + 3, j) = "-------------"
        G_VAL = G_VAL + DELTA_EPS_GROWTH
    Next j
    D_VAL = D_VAL + DELTA_DISCOUNT_RATE
Next i

DDM_EPS_GROWTH_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
DDM_EPS_GROWTH_FUNC = Err.number
End Function



'************************************************************************************
'************************************************************************************
'FUNCTION      : DIVIDEND_POLICY_ANALYSIS_FUNC

'DESCRIPTION   : Analysis of Dividend Policy
'--------------------------------------------------------------------------------------------------------------------------------------------
'Inputs for historical analysis
'--------------------------------------------------------------------------------------------------------------------------------------------
'1. To compare how much a firm has returned to its stockholder historically with how much it could have returned.
'2. To provide an assessment of project quality (ROE compared to cost of equity) and stock price performance
'over the period.
'a.Net Income
'b. Depreciation, amortization and other non-cash charges
'c. Capital expenditures: Please include acquisitions as part of capital expenditures
'd. Non-cash working capital changes
'In entering these numbers, please make sure that you get the signs right (check the comment box on each of these inputs)
'e. Dividends: Only cash dividends should be shown here (ignore stock dividends)
'f. Stock Buybacks: Include the cash flow associated with stock buybacks.
'For project assessment and stock price performance analysis
'a. Beta: You should really use an average beta over the historical period, but go ahead and use your current beta if you do not have this.
'b. Book Value of Equity: To compute return on equity.
'c. Return on the stock: This is the total return you would have made as an investor: It includes price appreciation + dividend yield each year
'd. Riskfree rate: The one-year government security rate at the start of each year (use the T.Bill rate)
'e. Return on Stock Market: This is the total return on the stock market each year
'(You can get the last two from the worksheet that is part of this spreadsheet that reports historical data on both)

'--------------------------------------------------------------------------------------------------------------------------------------------
'Output Historical Analysis
'--------------------------------------------------------------------------------------------------------------------------------------------
'1. FCFE and Cash Returned each year for the historical period
'2. Returns on equity, the stock and your required return each year for the historical period
'3. Averages of both over the entire period
'--------------------------------------------------------------------------------------------------------------------------------------------

'LIBRARY       : FUNDAMENTAL
'GROUP         : DIVIDENDS
'ID            : 002
'LAST UPDATE   : 27/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function DIVIDEND_POLICY_ANALYSIS_FUNC( _
ByRef PERIODS_RNG As Variant, _
ByRef NET_INCOME_RNG As Variant, _
ByRef DEPRECIATION_AMORTIZATION_RNG As Variant, _
ByRef CAPITAL_SPENDING_RNG As Variant, _
ByRef CHG_NON_CASH_WC_RNG As Variant, _
ByRef NET_DEBT_ISSUED_RNG As Variant, _
ByRef DIVIDENDS_RNG As Variant, _
ByRef EQUITY_REPURCHASES_RNG As Variant, _
ByRef STOCK_BETA_RNG As Variant, _
ByRef BV_EQUITY_RNG As Variant, _
ByRef STOCK_RETURN_RNG As Variant, _
ByRef TBILL_RATES_RNG As Variant, _
ByRef MARKET_RETURN_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

'--------------------------------------------------------------------------------------------------------------------------------------------
'Section 1 - Enter the following data for the years for which you have data (starting with earliest year and ending with the most recent year)
'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
'0) Year: Enter the following data relating to performance (starting with earliest year)
'1) Net Income: Net income before exptraordinary charges in each year.
'2) Depreciation & Amort: Enter depreciation, amortization and other non-cash charges.
'3) Capital Spending: Capital expenditures from each year, including acquisitions of other firms. (Enter as a positive number)
'4) Chg in Non-Cash WC: Enter changes in non-cash working capital. Using Balance Sheet:
                    '= Non-cash Working capital this year - Non-cash Working capital last year
                    'Using statement of cash flows
                    '= - (Change in non-cash working capital)
                    'Please remember to reverse the sign. An increase (decrease) in non-cash working capital should be entered a positive (negative) number.
'5) Net Debt Issued: Enter the change in interest-bearing debt from the previous year. Using balance sheets:
                'Interest-bearing debt this year - interest-bearing debt last year
                'Using statement of cash flow
                '= (Increase in LT Borrowing - Decrease in LT Borrowing + Increase in ST Borrowing - Decrease in ST Borrowing)

'--------------------------------------------------------------------------------------------------------------------------------------------
'Section 2 - Enter the dollar dividends paid and equity repurchases for each year of historical data (staring with earliest year): (Equity repurchases are in the statement of cash flows)
'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
'6) Dividends (in $): Total dollar dividends paid to common stockholders (not preferred) during the period
'7) Equity Repurchases (in $): Enter the stock buybacks  in statement of cashflows. Do not net out stock issues.

'--------------------------------------------------------------------------------------------------------------------------------------------
'Section 3 - Assessing Investment Quality and Stock Performance
'The following section of the dividend policy analysis looks at the quality of the firm's investments and the risk-adjusted,
'market-adjusted performance of your stock over the period. If you have a measure of excess returns (ROIC- Cost of capital)
'or  a Jensen's alpha computed for your stock, you can use those measures instead of the ones computed here.
'The last set of inputs to this analysis relate to project choice and performance.
'Make sure that you update the riskfree rate and return on the market to reflect the time period for your data.
'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
'8) Beta for the equity of this firm: Enter the beta of the company during the period. If you have a bottom up beta, use it. If not, use a regression beta.
'9) BV Equity: Enter the book value of common equity for each year (preferably at start) in the sample. If you want beginning of year equity, use the previous year's equity; for instance, enter the book value of equity for T and the book value of equity for T+1. You can also compute the average book equity each year...
'10) Returns on stock: This is the return you would have made on investing in the stock, inclusive of dividends and price appreciation.
'11) T.Bill rates: Enter the treasury bill rate at the start of each period.
'12) Returns on market: This is the return on the stock index, each year of the analysis.
'--------------------------------------------------------------------------------------------------------------------------------------------

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NCOLUMNS As Long
Dim HEADINGS_STR As String

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double
Dim TEMP4_SUM As Double

Dim PERIODS_VECTOR As Variant
Dim NET_INCOME_VECTOR As Variant
Dim DEPRECIATION_AMORTIZATION_VECTOR As Variant
Dim CAPITAL_SPENDING_VECTOR As Variant
Dim CHG_NON_CASH_WC_VECTOR As Variant
Dim NET_DEBT_ISSUED_VECTOR As Variant
Dim DIVIDENDS_VECTOR As Variant
Dim EQUITY_REPURCHASES_VECTOR As Variant
Dim STOCK_BETA_VECTOR As Variant
Dim BV_EQUITY_VECTOR As Variant
Dim STOCK_RETURN_VECTOR As Variant
Dim TBILL_RATES_VECTOR As Variant
Dim MARKET_RETURN_VECTOR As Variant

Dim TEMP0_MATRIX As Variant
Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant

On Error GoTo ERROR_LABEL

PERIODS_VECTOR = PERIODS_RNG
If UBound(PERIODS_VECTOR, 1) = 1 Then
    PERIODS_VECTOR = MATRIX_TRANSPOSE_FUNC(PERIODS_VECTOR)
End If
NROWS = UBound(PERIODS_VECTOR, 1)

NET_INCOME_VECTOR = NET_INCOME_RNG
If UBound(NET_INCOME_VECTOR, 1) = 1 Then
    NET_INCOME_VECTOR = MATRIX_TRANSPOSE_FUNC(NET_INCOME_VECTOR)
End If
If NROWS <> UBound(NET_INCOME_VECTOR, 1) Then: GoTo ERROR_LABEL

DEPRECIATION_AMORTIZATION_VECTOR = DEPRECIATION_AMORTIZATION_RNG
If UBound(DEPRECIATION_AMORTIZATION_VECTOR, 1) = 1 Then
    DEPRECIATION_AMORTIZATION_VECTOR = MATRIX_TRANSPOSE_FUNC(DEPRECIATION_AMORTIZATION_VECTOR)
End If
If NROWS <> UBound(DEPRECIATION_AMORTIZATION_VECTOR, 1) Then: GoTo ERROR_LABEL

CAPITAL_SPENDING_VECTOR = CAPITAL_SPENDING_RNG
If UBound(CAPITAL_SPENDING_VECTOR, 1) = 1 Then
    CAPITAL_SPENDING_VECTOR = MATRIX_TRANSPOSE_FUNC(CAPITAL_SPENDING_VECTOR)
End If
If NROWS <> UBound(CAPITAL_SPENDING_VECTOR, 1) Then: GoTo ERROR_LABEL

CHG_NON_CASH_WC_VECTOR = CHG_NON_CASH_WC_RNG
If UBound(CHG_NON_CASH_WC_VECTOR, 1) = 1 Then
    CHG_NON_CASH_WC_VECTOR = MATRIX_TRANSPOSE_FUNC(CHG_NON_CASH_WC_VECTOR)
End If
If NROWS <> UBound(CHG_NON_CASH_WC_VECTOR, 1) Then: GoTo ERROR_LABEL

NET_DEBT_ISSUED_VECTOR = NET_DEBT_ISSUED_RNG
If UBound(NET_DEBT_ISSUED_VECTOR, 1) = 1 Then
    NET_DEBT_ISSUED_VECTOR = MATRIX_TRANSPOSE_FUNC(NET_DEBT_ISSUED_VECTOR)
End If
If NROWS <> UBound(NET_DEBT_ISSUED_VECTOR, 1) Then: GoTo ERROR_LABEL

DIVIDENDS_VECTOR = DIVIDENDS_RNG
If UBound(DIVIDENDS_VECTOR, 1) = 1 Then
    DIVIDENDS_VECTOR = MATRIX_TRANSPOSE_FUNC(DIVIDENDS_VECTOR)
End If
If NROWS <> UBound(DIVIDENDS_VECTOR, 1) Then: GoTo ERROR_LABEL

EQUITY_REPURCHASES_VECTOR = EQUITY_REPURCHASES_RNG
If UBound(EQUITY_REPURCHASES_VECTOR, 1) = 1 Then
    EQUITY_REPURCHASES_VECTOR = MATRIX_TRANSPOSE_FUNC(EQUITY_REPURCHASES_VECTOR)
End If
If NROWS <> UBound(EQUITY_REPURCHASES_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(STOCK_BETA_RNG) = True Then
    STOCK_BETA_VECTOR = STOCK_BETA_RNG
    If UBound(STOCK_BETA_VECTOR, 1) = 1 Then
        STOCK_BETA_VECTOR = MATRIX_TRANSPOSE_FUNC(STOCK_BETA_VECTOR)
    End If
Else
    ReDim STOCK_BETA_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS: STOCK_BETA_VECTOR(i, 1) = STOCK_BETA_RNG: Next i
End If
If NROWS <> UBound(STOCK_BETA_VECTOR, 1) Then: GoTo ERROR_LABEL

BV_EQUITY_VECTOR = BV_EQUITY_RNG
If UBound(BV_EQUITY_VECTOR, 1) = 1 Then
    BV_EQUITY_VECTOR = MATRIX_TRANSPOSE_FUNC(BV_EQUITY_VECTOR)
End If
If NROWS <> UBound(BV_EQUITY_VECTOR, 1) Then: GoTo ERROR_LABEL

STOCK_RETURN_VECTOR = STOCK_RETURN_RNG
If UBound(STOCK_RETURN_VECTOR, 1) = 1 Then
    STOCK_RETURN_VECTOR = MATRIX_TRANSPOSE_FUNC(STOCK_RETURN_VECTOR)
End If
If NROWS <> UBound(STOCK_RETURN_VECTOR, 1) Then: GoTo ERROR_LABEL

TBILL_RATES_VECTOR = TBILL_RATES_RNG
If UBound(TBILL_RATES_VECTOR, 1) = 1 Then
    TBILL_RATES_VECTOR = MATRIX_TRANSPOSE_FUNC(TBILL_RATES_VECTOR)
End If
If NROWS <> UBound(TBILL_RATES_VECTOR, 1) Then: GoTo ERROR_LABEL

MARKET_RETURN_VECTOR = MARKET_RETURN_RNG
If UBound(MARKET_RETURN_VECTOR, 1) = 1 Then
    MARKET_RETURN_VECTOR = MATRIX_TRANSPOSE_FUNC(MARKET_RETURN_VECTOR)
End If
If NROWS <> UBound(MARKET_RETURN_VECTOR, 1) Then: GoTo ERROR_LABEL

'----------------------------------------------------------------------------------------------------------------------------------------------------
HEADINGS_STR = "Year,Net Income, - (Cap. Exp - Depr), - Delta Working Capital, + Net Debt Issued, = Free CF to Equity,Dividends, + Equity Repurchases, = Cash to Stockholders,Payout Ratio,Cash Paid as % of FCFE,ROE,Required rate of return,ROE - Cost of Equity,Returns on stock,Required rate of return,Jensen's alpha,"
'----------------------------------------------------------------------------------------------------------------------------------------------------
GoSub NCOLUMNS_LINE
ReDim TEMP0_MATRIX(0 To NROWS, 1 To NCOLUMNS)
i = 1
For k = 1 To NCOLUMNS
    j = InStr(i, HEADINGS_STR, ",")
    TEMP0_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
    i = j + 1
Next k
For i = 1 To NROWS
    TEMP0_MATRIX(i, 1) = PERIODS_VECTOR(i, 1)
    TEMP0_MATRIX(i, 2) = NET_INCOME_VECTOR(i, 1)
    TEMP0_MATRIX(i, 3) = CAPITAL_SPENDING_VECTOR(i, 1) - DEPRECIATION_AMORTIZATION_VECTOR(i, 1)
    TEMP0_MATRIX(i, 4) = CHG_NON_CASH_WC_VECTOR(i, 1)
    TEMP0_MATRIX(i, 5) = NET_DEBT_ISSUED_VECTOR(i, 1)
    TEMP0_MATRIX(i, 6) = TEMP0_MATRIX(i, 2) - TEMP0_MATRIX(i, 3) - TEMP0_MATRIX(i, 4) + TEMP0_MATRIX(i, 5)
    TEMP0_MATRIX(i, 7) = DIVIDENDS_VECTOR(i, 1)
    TEMP0_MATRIX(i, 8) = EQUITY_REPURCHASES_VECTOR(i, 1)
    TEMP0_MATRIX(i, 9) = TEMP0_MATRIX(i, 7) + TEMP0_MATRIX(i, 8)
    
    'Dividend Ratios
    If TEMP0_MATRIX(i, 2) <> 0 Then
        TEMP0_MATRIX(i, 10) = TEMP0_MATRIX(i, 7) / TEMP0_MATRIX(i, 2)
    Else
        TEMP0_MATRIX(i, 10) = CVErr(xlErrNA)
    End If
    
    If TEMP0_MATRIX(i, 6) <> 0 Then
        TEMP0_MATRIX(i, 11) = TEMP0_MATRIX(i, 9) / TEMP0_MATRIX(i, 6)
    Else
        TEMP0_MATRIX(i, 11) = CVErr(xlErrNA)
    End If
    
    'Performance Ratios - Accounting Measure
    If BV_EQUITY_VECTOR(i, 1) <> 0 Then
        TEMP0_MATRIX(i, 12) = NET_INCOME_VECTOR(i, 1) / BV_EQUITY_VECTOR(i, 1)
    Else
        TEMP0_MATRIX(i, 12) = CVErr(xlErrNA)
    End If
    
    If STOCK_BETA_VECTOR(i, 1) <> 0 Then
        TEMP0_MATRIX(i, 13) = TBILL_RATES_VECTOR(i, 1) + STOCK_BETA_VECTOR(i, 1) * (MARKET_RETURN_VECTOR(i, 1) - TBILL_RATES_VECTOR(i, 1))
        TEMP0_MATRIX(i, 16) = TEMP0_MATRIX(i, 13)
    Else
        TEMP0_MATRIX(i, 13) = CVErr(xlErrNA)
        TEMP0_MATRIX(i, 16) = CVErr(xlErrNA)
    End If
    
    If IsNumeric(TEMP0_MATRIX(i, 12)) And IsNumeric(TEMP0_MATRIX(i, 13)) Then
        TEMP0_MATRIX(i, 14) = TEMP0_MATRIX(i, 12) - TEMP0_MATRIX(i, 13)
    Else
        TEMP0_MATRIX(i, 14) = CVErr(xlErrNA)
    End If
    
    'Stock Performance Measure
    TEMP0_MATRIX(i, 15) = STOCK_RETURN_VECTOR(i, 1)
    If IsNumeric(TEMP0_MATRIX(i, 15)) And IsNumeric(TEMP0_MATRIX(i, 16)) Then
        TEMP0_MATRIX(i, 17) = TEMP0_MATRIX(i, 15) - TEMP0_MATRIX(i, 16)
    Else
        TEMP0_MATRIX(i, 17) = CVErr(xlErrNA)
    End If
Next i

If OUTPUT = 0 Then 'Analysis of Past Dividends
    DIVIDEND_POLICY_ANALYSIS_FUNC = TEMP0_MATRIX
    Exit Function
End If

'----------------------------------------------------------------------------------------------------------------------------------------------------
'Summary0 of calculations
'----------------------------------------------------------------------------------------------------------------------------------------------------
HEADINGS_STR = "Free CF to Equity,Dividends,Dividends+Repurchases,Average,Standard Deviation,Maximum,Minimum,"
GoSub NCOLUMNS_LINE
ReDim TEMP1_MATRIX(1 To 4, 1 To 5)
TEMP1_MATRIX(1, 1) = "Summary"
i = 1: l = 2
For k = 1 To 3
    j = InStr(i, HEADINGS_STR, ",")
    TEMP1_MATRIX(l, 1) = Mid(HEADINGS_STR, i, j - i)
    i = j + 1
    l = l + 1
Next k
l = 2
For k = 4 To 7
    j = InStr(i, HEADINGS_STR, ",")
    TEMP1_MATRIX(1, l) = Mid(HEADINGS_STR, i, j - i)
    i = j + 1
    l = l + 1
Next k
For i = 2 To 4
    TEMP1_MATRIX(i, 2) = 0: TEMP1_MATRIX(i, 3) = 0
    TEMP1_MATRIX(i, 4) = -2 ^ 52: TEMP1_MATRIX(i, 5) = 2 ^ 52
Next i
For i = 1 To NROWS
    TEMP1_MATRIX(2, 2) = TEMP1_MATRIX(2, 2) + TEMP0_MATRIX(i, 6)
    If TEMP0_MATRIX(i, 6) > TEMP1_MATRIX(2, 4) Then: TEMP1_MATRIX(2, 4) = TEMP0_MATRIX(i, 6)
    If TEMP0_MATRIX(i, 6) < TEMP1_MATRIX(2, 5) Then: TEMP1_MATRIX(2, 5) = TEMP0_MATRIX(i, 6)

    TEMP1_MATRIX(3, 2) = TEMP1_MATRIX(3, 2) + TEMP0_MATRIX(i, 7)
    If TEMP0_MATRIX(i, 7) > TEMP1_MATRIX(3, 4) Then: TEMP1_MATRIX(3, 4) = TEMP0_MATRIX(i, 7)
    If TEMP0_MATRIX(i, 7) < TEMP1_MATRIX(3, 5) Then: TEMP1_MATRIX(3, 5) = TEMP0_MATRIX(i, 7)

    TEMP1_MATRIX(4, 2) = TEMP1_MATRIX(4, 2) + TEMP0_MATRIX(i, 9)
    If TEMP0_MATRIX(i, 9) > TEMP1_MATRIX(4, 4) Then: TEMP1_MATRIX(4, 4) = TEMP0_MATRIX(i, 9)
    If TEMP0_MATRIX(i, 9) < TEMP1_MATRIX(4, 5) Then: TEMP1_MATRIX(4, 5) = TEMP0_MATRIX(i, 9)
Next i
TEMP1_MATRIX(2, 2) = TEMP1_MATRIX(2, 2) / NROWS
TEMP1_MATRIX(3, 2) = TEMP1_MATRIX(3, 2) / NROWS
TEMP1_MATRIX(4, 2) = TEMP1_MATRIX(4, 2) / NROWS
For i = 1 To NROWS
    TEMP1_MATRIX(2, 3) = TEMP1_MATRIX(2, 3) + (TEMP0_MATRIX(i, 6) - TEMP1_MATRIX(2, 2)) ^ 2
    TEMP1_MATRIX(3, 3) = TEMP1_MATRIX(3, 3) + (TEMP0_MATRIX(i, 7) - TEMP1_MATRIX(3, 2)) ^ 2
    TEMP1_MATRIX(4, 3) = TEMP1_MATRIX(4, 3) + (TEMP0_MATRIX(i, 9) - TEMP1_MATRIX(4, 2)) ^ 2
Next i
TEMP1_MATRIX(2, 3) = (TEMP1_MATRIX(2, 3) / (NROWS - 1)) ^ 0.5
TEMP1_MATRIX(3, 3) = (TEMP1_MATRIX(3, 3) / (NROWS - 1)) ^ 0.5
TEMP1_MATRIX(4, 3) = (TEMP1_MATRIX(4, 3) / (NROWS - 1)) ^ 0.5

If OUTPUT = 1 Then 'Summary1 of calculations
    Erase TEMP0_MATRIX
    DIVIDEND_POLICY_ANALYSIS_FUNC = TEMP1_MATRIX
    Exit Function
End If

'----------------------------------------------------------------------------------------------------------------------------------------------------
HEADINGS_STR = "Dividend Payout Ratio,Cash Paid as % of FCFE,ROE,Return on Stock,Required Return,ROE - Required return,Actual - Required Return,"
'----------------------------------------------------------------------------------------------------------------------------------------------------
GoSub NCOLUMNS_LINE
ReDim TEMP2_MATRIX(1 To 7, 1 To 2)
i = 1: l = 1
For k = 1 To 7
    j = InStr(i, HEADINGS_STR, ",")
    TEMP2_MATRIX(l, 1) = Mid(HEADINGS_STR, i, j - i)
    i = j + 1
    l = l + 1
Next k

TEMP1_SUM = 0: TEMP2_SUM = 0
TEMP3_SUM = 0: TEMP4_SUM = 0
For i = 1 To NROWS
    TEMP1_SUM = TEMP1_SUM + NET_INCOME_VECTOR(i, 1): TEMP2_SUM = TEMP2_SUM + BV_EQUITY_VECTOR(i, 1)
    TEMP3_SUM = TEMP3_SUM + TEMP0_MATRIX(i, 15): TEMP4_SUM = TEMP4_SUM + TEMP0_MATRIX(i, 16)
Next i
TEMP1_SUM = TEMP1_SUM / NROWS: TEMP2_SUM = TEMP2_SUM / NROWS
TEMP3_SUM = TEMP3_SUM / NROWS: TEMP4_SUM = TEMP4_SUM / NROWS

If TEMP1_SUM <> 0 Then
    TEMP2_MATRIX(1, 2) = TEMP1_MATRIX(3, 2) / TEMP1_SUM
Else
    TEMP2_MATRIX(1, 2) = CVErr(xlErrNA)
End If

If TEMP1_MATRIX(2, 2) <> 0 Then
    TEMP2_MATRIX(2, 2) = TEMP1_MATRIX(4, 2) / TEMP1_MATRIX(2, 2)
Else
    TEMP2_MATRIX(2, 2) = CVErr(xlErrNA)
End If

If TEMP2_SUM <> 0 Then
    TEMP2_MATRIX(3, 2) = TEMP1_SUM / TEMP2_SUM
Else
    TEMP2_MATRIX(3, 2) = CVErr(xlErrNA)
End If
TEMP2_MATRIX(4, 2) = TEMP3_SUM
TEMP2_MATRIX(5, 2) = TEMP4_SUM
TEMP2_MATRIX(6, 2) = TEMP2_MATRIX(3, 2) - TEMP2_MATRIX(5, 2)
TEMP2_MATRIX(7, 2) = TEMP2_MATRIX(4, 2) - TEMP2_MATRIX(5, 2)

If OUTPUT = 2 Then 'Average of calculations
    Erase TEMP0_MATRIX: Erase TEMP1_MATRIX
    DIVIDEND_POLICY_ANALYSIS_FUNC = TEMP2_MATRIX
Else
    DIVIDEND_POLICY_ANALYSIS_FUNC = Array(TEMP0_MATRIX, TEMP1_MATRIX, TEMP2_MATRIX)
End If

'--------------------------------------------------------------------------------------------------------------------------
Exit Function
'--------------------------------------------------------------------------------------------------------------------------
NCOLUMNS_LINE:
'--------------------------------------------------------------------------------------------------------------------------
    NCOLUMNS = 0
    i = 1
    Do
        j = InStr(i, HEADINGS_STR, ",")
        NCOLUMNS = NCOLUMNS + 1
        i = j + 1
    Loop Until i = 1
    NCOLUMNS = NCOLUMNS - 1
'--------------------------------------------------------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
DIVIDEND_POLICY_ANALYSIS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : FORECASTED_FCFE_DIVIDENDS_FUNC

'DESCRIPTION   : To provide forecasts of how much cash the firm will have available
'for stock buybacks in the future

'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
'INPUTS: For forecasts
'--------------------------------------------------------------------------------------------------------------------------------------------
'a. Expected growth rates in net income, dividends, depreciation, capital expenditures and revenues
'b. Working capital as a percent of revenues
'c. Debt as a percent of reinvestment, looking forward. As a default, you can use your historical average.
'--------------------------------------------------------------------------------------------------------------------------------------------
'OUTPUT: Forecasts
'--------------------------------------------------------------------------------------------------------------------------------------------
'1. Forecasted FCFE for next Z years
'2. Forecasted dividends for next Z years
'3. Cash available each year for stock buybacks for next Z years.
'--------------------------------------------------------------------------------------------------------------------------------------------

'LIBRARY       : FUNDAMENTAL
'GROUP         : DIVIDENDS
'ID            : 003
'LAST UPDATE   : 27/02/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function FORECASTED_FCFE_DIVIDENDS_FUNC( _
ByVal CURRENT_REVENUE_VAL As Double, _
ByVal CURRENT_NET_INCOME_VAL As Double, _
ByVal CURRENT_CAPITAL_EXPENDITURES_VAL As Double, _
ByVal CURRENT_DEPRECIATION_VAL As Double, _
ByVal CURRENT_DIVIDENDS_VAL As Double, _
ByVal NON_CASH_WORKING_CAPITAL_REVENUE_PERCENT_VAL As Double, _
ByVal EXPECTED_GROWTH_REVENUES_VAL As Double, _
ByVal EXPECTED_GROWTH_NET_INCOME_VAL As Double, _
ByVal EXPECTED_GROWTH_CAPITAL_EXPENDITURES_VAL As Double, _
ByVal EXPECTED_GROWTH_DEPRECIATION_VAL As Double, _
ByVal EXPECTED_GROWTH_DIVIDENDS_VAL As Double, _
ByVal EXPECTED_DEBT_CAPITAL_RATIO_VAL As Double, _
Optional ByVal NO_PERIODS_VAL As Long = 5)

'Expected growth in Revenues over next Z years. Projected growth rate in sales; if unavailable, look at historical growth.
'Expected growth in Net Income over next Z years. Enter projected growth in EPS, estimated by analysts for next Z years.
'Expected growth in capital expenditures in next Z years. Set equal to growth in revenues, if you do not have any information on growth.
'Expected growth in depreciation in next Z years. Generally, set equal to the growth rate in cap ex. If cap ex is significantly higher than depreciation, this growth rate can be set higher than cap ex growth.
'Enter non-cash working capital as a percent of revenues. You can use the estimate from the most recent year, or use the industry average.
'Enter revenues from most recent year. Current year's revenues
'Enter net income from most recent year or normalized value. If your net income in the current year is not normal or negative, this is your change to replace it with a normalized value.
'Enter capital expenditures from most recent year or normalized value. If capital expenditures are volatile, you can average it out over time and use a normalized value in here.
'Enter depreciation in most recent year
'Enter dividends paid in the most recent year
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Enter the debt to capital ratio to use in forecast (debt ratio used). This can be the current debt ratio or a predicted debt ratio, depending upon whether you think the firm's leverage will change over the period.
'Do you want to use the debt ratio from your historical analysis?. If yes, I will use the ratio of net debt issued to total reinvestment from your historical analysis.
'If yes, the debt ratio used will be. This number is computed using the reinvestment over the period and your net debt issues over the period. Consequently, it can be a strange number - greater than 100%. if you issued a lot of debt during the period or less than 0%. if you paid off debt. Feel free to override it.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Enter expected growth in dividends. This growth rate can reflect expected growth in earnings and management targets for dividends.

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long
Dim NCOLUMNS As Long
Dim HEADINGS_STR As String
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

NCOLUMNS = 9
HEADINGS_STR = "Period,Net Income, - (Cap Ex - Deprec'n) (1 - DR), -  Change in Working Capital (1 - DR),FCFE,Expected Dividends,Cash available for stock buybacks,Revenues,Non-cash WC,"
NROWS = NO_PERIODS_VAL
ReDim TEMP_MATRIX(0 To NROWS, 1 To NCOLUMNS)
i = 1
For k = 1 To NCOLUMNS
    j = InStr(i, HEADINGS_STR, ",")
    TEMP_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
    i = j + 1
Next k
For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = i
    TEMP_MATRIX(i, 2) = (1 + EXPECTED_GROWTH_NET_INCOME_VAL) ^ TEMP_MATRIX(i, 1) * CURRENT_NET_INCOME_VAL
    TEMP_MATRIX(i, 3) = ((1 + EXPECTED_GROWTH_CAPITAL_EXPENDITURES_VAL) ^ TEMP_MATRIX(i, 1) * CURRENT_CAPITAL_EXPENDITURES_VAL - (1 + EXPECTED_GROWTH_DEPRECIATION_VAL) ^ TEMP_MATRIX(i, 1) * CURRENT_DEPRECIATION_VAL) * (1 - EXPECTED_DEBT_CAPITAL_RATIO_VAL)
    TEMP_MATRIX(i, 6) = (1 + EXPECTED_GROWTH_DIVIDENDS_VAL) ^ TEMP_MATRIX(i, 1) * CURRENT_DIVIDENDS_VAL
    TEMP_MATRIX(i, 8) = CURRENT_REVENUE_VAL * (1 + EXPECTED_GROWTH_REVENUES_VAL) ^ TEMP_MATRIX(i, 1)
    If i > 1 Then TEMP_MATRIX(i, 9) = (TEMP_MATRIX(i, 8) - TEMP_MATRIX(i - 1, 8)) * NON_CASH_WORKING_CAPITAL_REVENUE_PERCENT_VAL Else TEMP_MATRIX(i, 9) = (TEMP_MATRIX(i, 8) - CURRENT_REVENUE_VAL) * NON_CASH_WORKING_CAPITAL_REVENUE_PERCENT_VAL
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 9) * (1 - EXPECTED_DEBT_CAPITAL_RATIO_VAL)
    TEMP_MATRIX(i, 5) = TEMP_MATRIX(i, 2) - TEMP_MATRIX(i, 3) - TEMP_MATRIX(i, 4)
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 6)
Next i
FORECASTED_FCFE_DIVIDENDS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
FORECASTED_FCFE_DIVIDENDS_FUNC = Err.number
End Function
