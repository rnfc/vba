Attribute VB_Name = "WEB_SERVICE_ADVFN_LIBR"
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
Option Compare Text  'Uppercase letters to be equivalent to lowercase letters.
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'It is of paramount importance to remember that, any resulting companies
'from the stock screening process constitute only the starting point of
'more meticulous and in-depth research.
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
Private PUB_ADVFN_HASH_OBJ As clsTypeHash
'-----------------------------------------------------------------------------------------------------------
Private Const PUB_ADVFN_SERVER_STR As String = "ca"
'-----------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : ADVFN_KEY_STATISTICS_FUNC
'DESCRIPTION   : ADVFN Key Statistics Wrapper
'LIBRARY       : HTML
'GROUP         : ADVFN
'ID            : 001
'LAST UPDATE   : 2013.10.07
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ADVFN_KEY_STATISTICS_FUNC(ByRef TICKERS_RNG As Variant, _
Optional ByVal VERSION As Long = 0)

Dim h As Long
Dim i As Long
Dim j As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_STR As String
Dim DATA_STR As String
Dim REFER_STR As String
Dim HEADER_STR As String
Dim TICKER_STR As String
Dim KEY_STR As String
Dim SRC_URL_STR As String

Dim TEMP_ARR As Variant
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

Call INITIALIZE_ADVFN_HASH_TABLE_FUNC(False) '(10000)

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then: _
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_RNG)
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If
NCOLUMNS = UBound(TICKERS_VECTOR, 1)

KEY_STR = VERSION
For i = 1 To NCOLUMNS
    TICKERS_VECTOR(i, 1) = CONVERT_YAHOO_TICKER_FUNC(TICKERS_VECTOR(i, 1), "ADVFN")
    KEY_STR = KEY_STR & "|" & TICKERS_VECTOR(i, 1)
Next i

If PUB_ADVFN_HASH_OBJ.Exists(KEY_STR) = True Then
    TEMP_MATRIX = PUB_ADVFN_HASH_OBJ(KEY_STR)
    If IsArray(TEMP_MATRIX) = False Then: GoTo ERROR_LABEL
    ADVFN_KEY_STATISTICS_FUNC = CONVERT_STRING_NUMBER_FUNC(TEMP_MATRIX)
    Exit Function
End If

GoSub HEADERS_LINE
TEMP_ARR = Split(TEMP_ARR, "|", -1)
NROWS = UBound(TEMP_ARR) - LBound(TEMP_ARR) + 1
    
If VERSION = 0 Then
    ReDim TEMP_MATRIX(1 To NROWS + 1, 0 To NCOLUMNS) As String
Else
    ReDim TEMP_MATRIX(1 To NROWS + 1, 0 To NCOLUMNS) As Variant
End If

TEMP_MATRIX(1, 0) = Replace(Replace(UCase(HEADER_STR), ">", ""), ":", "")
If TEMP_MATRIX(1, 0) = "COMPANY NAME" Then
    TEMP_MATRIX(1, 0) = "COMPANY INFO"
ElseIf TEMP_MATRIX(1, 0) = "Industry Information" Then
    TEMP_MATRIX(1, 0) = "KEY STATS"
End If

ii = 2
For kk = LBound(TEMP_ARR) To UBound(TEMP_ARR)
    TEMP_MATRIX(ii, 0) = Replace(TEMP_ARR(kk), ">", "")
    ii = ii + 1
Next kk

'-----------------------------------------------------------------------------
Select Case VERSION
'-----------------------------------------------------------------------------
Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
'-----------------------------------------------------------------------------
    
    For jj = 1 To NCOLUMNS
        GoSub TICKER_URL_LINE
        If DATA_STR = "0" Or DATA_STR = PUB_WEB_DATA_SYSTEM_ERROR_STR Then: GoTo 1990
        DATA_STR = Replace(DATA_STR, "&nbsp;", "")
        ii = InStr(1, DATA_STR, HEADER_STR)
        If ii = 0 Then: GoTo 1990
        DATA_STR = Mid(DATA_STR, ii, Len(DATA_STR) - ii)
        ii = 2
        
        h = 1
        For kk = LBound(TEMP_ARR) To UBound(TEMP_ARR)
            REFER_STR = CStr(TEMP_ARR(kk))
            If REFER_STR = "WWW Address: " Then
                i = InStr(h, DATA_STR, REFER_STR)
                If i = 0 Then
                    i = h
                    GoTo 1983
                End If
                i = InStr(i, DATA_STR, "href='") + Len("href='")
                j = InStr(i, DATA_STR, "'")
                TEMP_STR = Mid(DATA_STR, i, j - i)
            ElseIf REFER_STR = "Industry Information:" Then
                i = InStr(h, DATA_STR, REFER_STR)
                If i = 0 Then
                    i = h
                    GoTo 1983
                End If
                i = InStr(i, DATA_STR, "<b>") + Len("<b>")
                j = InStr(i, DATA_STR, "<a")
                TEMP_STR = Replace(Mid(DATA_STR, i, j - i), "</b>", "")
            ElseIf REFER_STR = "More Like This:" Then
                i = InStr(h, DATA_STR, "Industry Information:")
                If i = 0 Then
                    i = h
                    GoTo 1983
                End If
                i = InStr(i, DATA_STR, "<b>") + Len("<b>")
                i = InStr(i, DATA_STR, "<a href=") + Len("<a href=") + 1
                j = InStr(i, DATA_STR, ">") - 1
                TEMP_STR = "http://www.advfn.com" & Mid(DATA_STR, i, j - i)
            ElseIf REFER_STR = "Yesterday's Close" Then
                i = InStr(h, DATA_STR, REFER_STR)
                If i = 0 Then
                    i = h
                    GoTo 1983
                End If
                i = InStr(i, DATA_STR, "<span") + Len("<span")
                i = InStr(i, DATA_STR, ">") + Len(">")
                j = InStr(i, DATA_STR, "<")
                TEMP_STR = Mid(DATA_STR, i, j - i)
            Else
                i = InStr(h, DATA_STR, REFER_STR)
                If i = 0 Then
                    i = h
                    GoTo 1983
                End If
                i = InStr(i, DATA_STR, "</td>") + Len("</td>")
                i = InStr(i, DATA_STR, ">") + Len(">")
                j = InStr(i, DATA_STR, "<")
                TEMP_STR = Mid(DATA_STR, i, j - i)
            End If
            
            TEMP_MATRIX(ii, jj) = IIf(TEMP_STR = "NA", 0, TEMP_STR)
1983:
            ii = ii + 1
        Next kk
1990:
    Next jj

'-----------------------------------------------------------------------------
Case 10
'-----------------------------------------------------------------------------
    For jj = 1 To NCOLUMNS
        GoSub TICKER_URL_LINE
        If DATA_STR = "0" Or DATA_STR = PUB_WEB_DATA_SYSTEM_ERROR_STR Then: GoTo 1991
        DATA_STR = Replace(DATA_STR, "&nbsp;", "")
        
        ii = InStr(1, DATA_STR, HEADER_STR)
        If ii = 0 Then: GoTo 1991
        DATA_STR = Mid(DATA_STR, ii, Len(DATA_STR) - ii)
        
        ii = 2
        
        For kk = LBound(TEMP_ARR) To UBound(TEMP_ARR)
            If ii <= 9 Then
                REFER_STR = "Volume"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 2 Then: i = h + Len(REFER_STR)
            Else
                REFER_STR = "52-Wks-Range"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 10 Then: i = h + Len(REFER_STR)
            End If
            If ii = 9 Or ii = 17 Then
                j = InStr(i, DATA_STR, "</td></tr>")
            Else
                j = InStr(i, DATA_STR, "</td><td")
            End If
            i = j
            Do While Mid(DATA_STR, i, 1) <> ">": i = i - 1: Loop
            i = i + 1
            TEMP_STR = Trim(Mid(DATA_STR, i, j - i))
            i = j + 1
            TEMP_MATRIX(ii, jj) = IIf(TEMP_STR = "NA", 0, TEMP_STR)
            ii = ii + 1
        Next kk
1991:
    Next jj
    
'-----------------------------------------------------------------------------
Case 11 ' HISTORICAL PRICE/VOLUME --> PERFECT
'-----------------------------------------------------------------------------
    For jj = 1 To NCOLUMNS
        GoSub TICKER_URL_LINE
        If DATA_STR = "0" Or DATA_STR = PUB_WEB_DATA_SYSTEM_ERROR_STR Then: GoTo 1992
        DATA_STR = Replace(DATA_STR, "&nbsp;", "")
        
        ii = InStr(1, DATA_STR, HEADER_STR)
        If ii = 0 Then: GoTo 1992
        DATA_STR = Mid(DATA_STR, ii, Len(DATA_STR) - ii)
        
        ii = 2
        
        For kk = LBound(TEMP_ARR) To UBound(TEMP_ARR)
            If ii <= 8 Then
                REFER_STR = "1 Week"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 2 Then: i = h + Len(REFER_STR) + Len("</td><td")
            ElseIf ii > 8 And ii <= 15 Then
                REFER_STR = "4 Weeks"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 9 Then: i = h + Len(REFER_STR) + Len("</td><td")
            ElseIf ii > 15 And ii <= 22 Then
                REFER_STR = "13 Weeks"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 16 Then: i = h + Len(REFER_STR) + Len("</td><td")
            ElseIf ii > 22 And ii <= 29 Then
                REFER_STR = "26 Weeks"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 23 Then: i = h + Len(REFER_STR) + Len("</td><td")
            ElseIf ii > 29 And ii <= 36 Then
                REFER_STR = "52 Weeks"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 30 Then: i = h + Len(REFER_STR) + Len("</td><td")
            ElseIf ii > 36 And ii <= 43 Then
                REFER_STR = "YTD"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 37 Then: i = h + Len(REFER_STR) + Len("</td><td")
            ElseIf ii > 43 And ii <= 50 Then
                REFER_STR = "Beta (36-Mnth)"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 44 Then
                    i = h + Len(REFER_STR)
                    i = InStr(i, DATA_STR, "</td><td") + Len("</td><td")
                End If
            End If
            
            If ii = 8 Or ii = 15 Or ii = 22 Or ii = 29 Or _
               ii = 36 Or ii = 43 Or ii = 50 Then
                j = InStr(i, DATA_STR, "</td></tr>")
            Else
                j = InStr(i, DATA_STR, "</td><td")
            End If
            i = j
            Do While Mid(DATA_STR, i, 1) <> ">": i = i - 1: Loop
            i = i + 1
            TEMP_STR = Trim(Mid(DATA_STR, i, j - i))
            i = j + 1
            TEMP_MATRIX(ii, jj) = IIf(TEMP_STR = "NA", 0, TEMP_STR)
            ii = ii + 1
        Next kk
1992:
    Next jj
    
'-----------------------------------------------------------------------------
Case 12 ' GROWTH RATES --> PERFECT
'-----------------------------------------------------------------------------
    For jj = 1 To NCOLUMNS
        GoSub TICKER_URL_LINE
        If DATA_STR = "0" Or DATA_STR = PUB_WEB_DATA_SYSTEM_ERROR_STR Then: GoTo 1993
        DATA_STR = Replace(DATA_STR, "&nbsp;", "")
        
        ii = InStr(1, DATA_STR, HEADER_STR)
        If ii = 0 Then: GoTo 1993
        DATA_STR = Mid(DATA_STR, ii, Len(DATA_STR) - ii)
        
        ii = 2
        For kk = LBound(TEMP_ARR) To UBound(TEMP_ARR)
            If ii <= 4 Then
                REFER_STR = "Revenue"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 2 Then: i = h + Len(REFER_STR) + Len("</td><td")
            ElseIf ii > 4 And ii <= 7 Then
                REFER_STR = "Income"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 5 Then: i = h + Len(REFER_STR) + Len("</td><td")
            ElseIf ii > 7 And ii <= 10 Then
                REFER_STR = "Dividend"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 8 Then: i = h + Len(REFER_STR) + Len("</td><td")
            ElseIf ii > 10 And ii <= 13 Then
                REFER_STR = "Capital Spending"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 11 Then: i = h + Len(REFER_STR) + Len("</td><td")
            ElseIf ii > 13 And ii <= 16 Then
                REFER_STR = "R&D"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 14 Then: i = h + Len(REFER_STR) + Len("</td><td")
            ElseIf ii > 16 And ii <= 19 Then
                REFER_STR = "Normalized Inc."
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 17 Then: i = h + Len(REFER_STR) + Len("</td><td")
            End If
            
            If ii = 4 Or ii = 7 Or ii = 10 Or ii = 13 Or ii = 16 Or ii = 19 Then
                j = InStr(i, DATA_STR, "</td></tr>")
            Else
                j = InStr(i, DATA_STR, "</td><td")
            End If
            i = j
            Do While Mid(DATA_STR, i, 1) <> ">": i = i - 1: Loop
            i = i + 1
            TEMP_STR = Trim(Mid(DATA_STR, i, j - i))
            i = j + 1
            TEMP_MATRIX(ii, jj) = IIf(TEMP_STR = "NA", 0, TEMP_STR)
            ii = ii + 1
        Next kk
1993:
    Next jj
    
'-----------------------------------------------------------------------------
Case Else 'CHANGES IN GROWTH RATES --> PERFECT

'-----------------------------------------------------------------------------
    For jj = 1 To NCOLUMNS
        GoSub TICKER_URL_LINE
        If DATA_STR = "0" Or DATA_STR = PUB_WEB_DATA_SYSTEM_ERROR_STR Then: GoTo 1994
        DATA_STR = Replace(DATA_STR, "&nbsp;", "")

        ii = InStr(1, DATA_STR, HEADER_STR)
        If ii = 0 Then: GoTo 1994
        DATA_STR = Mid(DATA_STR, ii, Len(DATA_STR) - ii)
        
        ii = 2
        
        For kk = LBound(TEMP_ARR) To UBound(TEMP_ARR)
            If ii <= 4 Then
                REFER_STR = "Revenue %"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 2 Then: i = h + Len(REFER_STR) + Len("</td><td")
            ElseIf ii > 4 And ii <= 7 Then
                REFER_STR = "Earnings %"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 5 Then: i = h + Len(REFER_STR) + Len("</td><td")
            ElseIf ii > 7 And ii <= 10 Then
                REFER_STR = "EPS %"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 8 Then: i = h + Len(REFER_STR) + Len("</td><td")
            ElseIf ii > 10 And ii <= 13 Then
                REFER_STR = "EPS $"
                h = InStr(1, DATA_STR, REFER_STR)
                If h = 0 Then: Exit For
                If ii = 11 Then: i = h + Len(REFER_STR) + Len("</td><td")
            End If
            
            If ii = 4 Or ii = 7 Or ii = 10 Or ii = 13 Then
                j = InStr(i, DATA_STR, "</td></tr>")
            Else
                j = InStr(i, DATA_STR, "</td><td")
            End If
            i = j
            Do While Mid(DATA_STR, i, 1) <> ">": i = i - 1: Loop
            i = i + 1
            TEMP_STR = Trim(Mid(DATA_STR, i, j - i))
            i = j + 1
            TEMP_MATRIX(ii, jj) = IIf(TEMP_STR = "NA", 0, TEMP_STR)
            ii = ii + 1
        Next kk
1994:
    Next jj
    
'-----------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------

Call PUB_ADVFN_HASH_OBJ.Add(KEY_STR, TEMP_MATRIX)
ADVFN_KEY_STATISTICS_FUNC = CONVERT_STRING_NUMBER_FUNC(TEMP_MATRIX)

'-----------------------------------------------------------------------------
Exit Function
'-----------------------------------------------------------------------------
HEADERS_LINE:
'-----------------------------------------------------------------------------
    '---------------------------------------------------------------------------------
    If VERSION = 0 Then '--> PERFECT
    '---------------------------------------------------------------------------------
        HEADER_STR = ">Company Name"
        TEMP_ARR = ">Company Name: |>Ticker Symbol: |WWW Address: |>CEO: |>" & _
                   "No. of Employees: |>Common Issue Type: |>Business Description:|" & _
                   "Industry Information:|More Like This:"
    '---------------------------------------------------------------------------------
    ElseIf VERSION = 1 Then '--> PERFECT
    '---------------------------------------------------------------------------------
        HEADER_STR = ">KEY FIGURES (Latest Twelve Months - LTM)"
        TEMP_ARR = "Yesterday's Close|>PE Ratio - LTM|>Market Capitalisation|>" & _
                   "Latest Shares Outstanding|>Earnings pS (EPS)|>Dividend pS (DPS)|>" & _
                   "Dividend Yield|>Dividend Payout Ratio|>Revenue per Employee|>" & _
                   "Effective Tax Rate|>Float|>Float as % of Shares Outstanding|>" & _
                   "Foreign Sales|>Domestic Sales|>" & _
                   "Selling, General & Adm/tive (SG&A) as % of Revenue|>" & _
                   "Research & Devlopment (R&D) as % of Revenue|>Gross Profit Margin|>" & _
                   "EBITDA Margin|>Pre-Tax Profit Margin|>Assets Turnover|>" & _
                   "Return on Assets (ROA)|>Return on Equity (ROE)|>" & _
                   "Return on Capital Invested (ROCI)|>Current Ratio|>" & _
                   "Leverage Ratio (Assets/Equity)|>Interest Cover|>" & _
                   "Total Debt/Equity (Gearing Ratio)|>LT Debt/Total Capital|>" & _
                   "Working Capital pS|>Cash pS|>Book-Value pS|>Tangible Book-Value pS|>" & _
                   "Cash Flow pS|>Free Cash Flow pS"
    '---------------------------------------------------------------------------------
    ElseIf VERSION = 2 Then ' --> PERFECT
    '---------------------------------------------------------------------------------
        HEADER_STR = ">KEY FIGURES (LTM): Price info"
        TEMP_ARR = ">Price/Book Ratio|>Price/Tangible Book Ratio|>" & _
                   "Price/Cash Flow|>Price/Free Cash Flow|>P/E as % of Industry Group|>" & _
                   "P/E as % of Sector Segment"
    '---------------------------------------------------------------------------------
    ElseIf VERSION = 3 Then '--> PERFECT
    '---------------------------------------------------------------------------------
        HEADER_STR = ">DIVIDEND INFO"
        TEMP_ARR = ">Dividend Declared Date|>Dividend Ex-Date|>Dividend Record Date|>" & _
                   "Dividend Pay Date|>Dividend Amount|>Type of Payment|>Dividend Rate|>" & _
                   "Current Dividend Yield|>5-Y Average Dividend Yield|>Payout Ratio|>" & _
                   "5-Y Average Payout Ratio"
    '---------------------------------------------------------------------------------
    ElseIf VERSION = 4 Then 'SOLVENCY RATIOS --> PERFECT
    '---------------------------------------------------------------------------------
        HEADER_STR = ">SHORT-TERM SOLVENCY RATIOS (LIQUIDITY)"
        TEMP_ARR = ">Net Working Capital Ratio|>Current Ratio|>Quick Ratio (Acid Test)|>" & _
                   "Liquidity Ratio (Cash)|>Receivables Turnover|>Average Collection Period|>" & _
                   "Working Capital/Equity|>Working Capital pS|>Free Cash Flow Margin|>" & _
                   "Free Cash Flow Margin 5YEAR AVG|>Cash-Flow pS|>Free Cash-Flow pS"
    '---------------------------------------------------------------------------------
    ElseIf VERSION = 5 Then 'SOLVENCY RATIOS --> PERFECT
    '---------------------------------------------------------------------------------
        HEADER_STR = ">FINANCIAL STRUCTURE RATIOS"
        TEMP_ARR = ">Altman's Z-Score Ratio|>Financial Leverage Ratio (Assets/Equity)|>" & _
                    "Debt Ratio|>Total Debt/Equity (Gearing Ratio)|>LT Debt/Equity|>" & _
                    "LT Debt/Capital Invested|>LT Debt/Total Liabilities|>Interest Cover|>" & _
                    "Interest/Capital Invested"
    
    '---------------------------------------------------------------------------------
    ElseIf VERSION = 6 Then 'VALUATION RATIOS --> PERFECT
    '---------------------------------------------------------------------------------
        HEADER_STR = ">MULTIPLES"
        TEMP_ARR = ">PQ Ratio|>Tobin's Q Ratio|>Current P/E Ratio - LTM|>" & _
                    "Enterprise Value (EV)/EBITDA|>Enterprise Value (EV)/Free Cash Flow|>" & _
                    "Dividend Yield|>Price/Tangible Book Ratio - LTM|>Price/Book Ratio - LTM|>" & _
                    "Price/Cash Flow Ratio|>Price/Free Cash Flow Ratio - LTM|>Price/Sales Ratio|>" & _
                    "P/E Ratio (1 month ago) - LTM|>P/E Ratio (26 weeks ago) - LTM|>" & _
                    "P/E Ratio (52 weeks ago) - LTM|>5-Y High P/E Ratio|>5-Y Low P/E Ratio|>" & _
                    "5-Y Average P/E Ratio|>Current P/E Ratio as % of 5-Y Average P/E|>" & _
                    "P/E as % of Industry Group|>P/E as % of Sector Segment|>" & _
                    "Current 12 Month Normalized P/E Ratio - LTM"
    
    '---------------------------------------------------------------------------------
    ElseIf VERSION = 7 Then 'VALUATION RATIOS --> PERFECT
    '---------------------------------------------------------------------------------
        HEADER_STR = ">PER SHARE FIGURES"
        TEMP_ARR = ">LT Debt pS|>Current Liabilities pS|>Tangible Book Value pS - LTM|>" & _
                   "Book Value pS - LTM|>Capital Invested pS|>Cash pS - LTM|>Cash Flow pS - LTM|>" & _
                   "Free Cash Flow pS - LTM|>Earnings pS (EPS)"
    
    '---------------------------------------------------------------------------------
    ElseIf VERSION = 8 Then 'OPERATING RATIOS
    '---------------------------------------------------------------------------------
        HEADER_STR = ">PROFITABILITY"
        TEMP_ARR = ">Free Cash Flow Margin|>Free Cash Flow Margin 5YEAR AVG|>" & _
                    "Net Profit Margin|>Net Profit Margin - 5YEAR AVRG.|>" & _
                    "Equity Productivity|>Return on Equity (ROE)|>" & _
                    "Return on Equity (ROE) - 5YEAR AVRG.|>Capital Invested Productivity|>" & _
                    "Return on Capital Invested (ROCI)|>" & _
                    "Return on Capital Invested (ROCI) - 5YEAR AVRG.|>" & _
                    "Assets Productivity|>Return on Assets (ROA)|>" & _
                    "Return on Assets (ROA) - 5YEAR AVRG.|>Gross Profit Margin|>" & _
                    "Gross Profit Margin - 5YEAR AVRG.|>EBITDA Margin - LTM|>" & _
                    "EBIT Margin - LTM|>Pre-Tax Profit Margin|>" & _
                    "Pre-Tax Profit Margin - 5YEAR AVRG.|>Effective Tax Rate|>" & _
                    "Effective Tax Rate - 5YEAR AVRG."
    
    '---------------------------------------------------------------------------------
    ElseIf VERSION = 9 Then 'OPERATING RATIOS
    '---------------------------------------------------------------------------------
        HEADER_STR = ">EFFICIENCY RATIOS"
        TEMP_ARR = ">Cash Conversion Cycle|>Revenue per Employee|>" & _
                    "Net Income per Employee|>Average Collection Period|>" & _
                    "Receivables Turnover|>Day's Inventory Turnover Ratio|>" & _
                    "Inventory Turnover|>Inventory/Sales|>Accounts Payble/Sales|>" & _
                    "Assets/Revenue|>Assets/Revenue - 5YEAR AVRG.|>" & _
                    "Net Working Capital Turnover|>Fixed Assets Turnover|>" & _
                    "Total Assets Turnover|>Revenue per $ Cash|>Revenue per $ Plant|>" & _
                    "Revenue per $ Common Equity|>Revenue per $ Capital Invested|>" & _
                    "Selling, General & Adm/tive (SG&A) as % of Revenue|>" & _
                    "SG&A Expense as % of Revenue - 5YEAR AVRG.|>" & _
                    "Research & Devlopment (R&D) as % of Revenue|>" & _
                    "R&D Expense as % of Revenue - 5YEAR AVRG."
    
    '---------------------------------------------------------------------------------
    ElseIf VERSION = 10 Then ' SUMMARY (QUOTES) --> PERFECT
    '---------------------------------------------------------------------------------
        HEADER_STR = "Industry Information:"
        TEMP_ARR = ">Price|>Day Change|>Bid|>Ask|>Open|>High|>Low|>Volume|>" & _
                   "Market Cap (mil)|>Shares Outstanding (mil)|>Beta|>EPS|>DPS|>P/E|>" & _
                   "Yield|>52-Wks-Range"
    '---------------------------------------------------------------------------------
    ElseIf VERSION = 11 Then ' HISTORICAL PRICE/VOLUME --> PERFECT
    '---------------------------------------------------------------------------------
        HEADER_STR = ">PRICE/VOLUME"
        TEMP_ARR = ">1 Week: High|>1 Week: Low|>1 Week: Close|>1 Week: % Price Chg|>" & _
        "1 Week: % Price Chg vs. Mkt.|>1 Week: Avg. Daily Vol|>1 Week: Total Vol|>" & _
        "4 Weeks: High|>4 Weeks: Low|>4 Weeks: Close|>4 Weeks: % Price Chg|>" & _
        "4 Weeks: % Price Chg vs. Mkt.|>4 Weeks: Avg. Daily Vol|>4 Weeks: Total Vol|>" & _
        "13 Weeks: High|>13 Weeks: Low|>13 Weeks: Close|>13 Weeks: % Price Chg|>" & _
        "13 Weeks: % Price Chg vs. Mkt.|>13 Weeks: Avg. Daily Vol|>" & _
        "13 Weeks: Total Vol|>26 Weeks: High|>26 Weeks: Low|>26 Weeks: Close|>" & _
        "26 Weeks: % Price Chg|>26 Weeks: % Price Chg vs. Mkt.|>" & _
        "26 Weeks: Avg. Daily Vol|>26 Weeks: Total Vol|>52 Weeks: High|>" & _
        "52 Weeks: Low|>52 Weeks: Close|>52 Weeks: % Price Chg|>" & _
        "52 Weeks: % Price Chg vs. Mkt.|>52 Weeks: Avg. Daily Vol|>" & _
        "52 Weeks: Total Vol|>YTD: High|>YTD: Low|>YTD: Close|>YTD: % Price Chg|>" & _
        "YTD: % Price Chg vs. Mkt.|>YTD: Avg. Daily Vol|>YTD: Total Vol|>" & _
        "Moving Average: 5-Days|>Moving Average: 10-Days|>" & _
        "Moving Average: 10-Weeks|>Moving Average: 30-Weeks|>Moving Average: 200-Days|>" & _
        "Moving Average: Beta (60-Mnth)|>Moving Average: Beta (36-Mnth)"
    
    '---------------------------------------------------------------------------------
    ElseIf VERSION = 12 Then ' GROWTH RATES --> PERFECT
    '---------------------------------------------------------------------------------
        HEADER_STR = ">GROWTH RATES"
        TEMP_ARR = ">Revenue: 5-Year Growh|>Revenue: R= of 5-Year Growth|>" & _
                   "Revenue: 3-Year Growth|>Income: 5-Year Growh|>" & _
                   "Income: R= of 5-Year Growth|>Income: 3-Year Growth|>" & _
                   "Dividend: 5-Year Growh|>Dividend: R= of 5-Year Growth|>" & _
                   "Dividend: 3-Year Growth|>Capital Spending: 5-Year Growh|>" & _
                   "Capital Spending: R= of 5-Year Growth|>" & _
                   "Capital Spending: 3-Year Growth|>R&D: 5-Year Growh|>" & _
                   "R&D: R= of 5-Year Growth|>R&D: 3-Year Growth|>" & _
                   "Normalized Inc.: 5-Year Growh|>Normalized Inc.: R= of 5-Year Growth|>" & _
                   "Normalized Inc.: 3-Year Growth"
    '---------------------------------------------------------------------------------
    Else 'CHANGES IN GROWTH RATES --> PERFECT
    '---------------------------------------------------------------------------------
        HEADER_STR = ">CHANGES"
        TEMP_ARR = ">Revenue %: YTD vs. Last YTD|>Revenue %: Curr Qtr vs. Qtr 1-Yr ago|>" & _
                    "Revenue %: Annual vs. Last Annual|>Earnings %: YTD vs. Last YTD|>" & _
                    "Earnings %: Curr Qtr vs. Qtr 1-Yr ago|>Earnings %: Annual vs. Last Annual|>" & _
                    "EPS %: YTD vs. Last YTD|>EPS %: Curr Qtr vs. Qtr 1-Yr ago|>" & _
                    "EPS %: Annual vs. Last Annual|>EPS $: YTD vs. Last YTD|>" & _
                    "EPS $: Curr Qtr vs. Qtr 1-Yr ago|>EPS $: Annual vs. Last Annual"
    '---------------------------------------------------------------------------------
    End If
    '---------------------------------------------------------------------------------
'-----------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------
TICKER_URL_LINE:
'-----------------------------------------------------------------------------
    TICKER_STR = TICKERS_VECTOR(jj, 1)
    SRC_URL_STR = ADVFN_URL_STRING_FUNC(TICKER_STR, 0, 0, False)
    TEMP_MATRIX(1, jj) = DECODE_URL_FUNC(TICKER_STR)
    DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
'-----------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------
ERROR_LABEL:
Call INITIALIZE_ADVFN_HASH_TABLE_FUNC(True)
ADVFN_KEY_STATISTICS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ADVFN_STATEMENTS_FUNC
'DESCRIPTION   : Retrieve ADVFN Statements for US & CAD Public Companies
'LIBRARY       : HTML
'GROUP         : ADVFN
'ID            : 002
'LAST UPDATE   : 2013.10.08
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function ADVFN_STATEMENTS_FUNC(ByVal TICKER_STR As String, _
Optional ByRef ELEMENTS_RNG As Variant, _
Optional ByVal QUATERLY_FLAG As Boolean = False, _
Optional ByVal NO_PERIODS As Long = 10)

'NO_PERIODS--> 'Maximum data periods
'Bound Array --> 0-55; (56/4) = 14 years of Financial Statements

'Sub TEST_ADVFN_STATEMENTS_FUNC()
'Call INITIALIZE_ADVFN_HASH_TABLE_FUNC(True) '(10000)
'Call ADVFN_STATEMENTS_FUNC("NASDAQ:MSFT", ADVFN_STATEMENT_ELEMENTS_FUNC(0, 0, False, 1), False)

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim LOOK_STR As String

Dim SRC_URL_STR As String

Dim TEMP_STR As String
Dim KEY_STR As String

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim ELEMENT_VECTOR As Variant

On Error GoTo ERROR_LABEL

'-----------------------------------------------------------------------
Call INITIALIZE_ADVFN_HASH_TABLE_FUNC(False) '(10000)
TICKER_STR = CONVERT_YAHOO_TICKER_FUNC(TICKER_STR, "ADVFN")
KEY_STR = TICKER_STR & "|" & QUATERLY_FLAG & "|" & NO_PERIODS
If IsMissing(ELEMENTS_RNG) = False Then
    If IsArray(ELEMENTS_RNG) = True Then
        ELEMENT_VECTOR = ELEMENTS_RNG
        If UBound(ELEMENT_VECTOR, 1) = 1 Then
            ELEMENT_VECTOR = MATRIX_TRANSPOSE_FUNC(ELEMENTS_RNG)
        End If
    Else
        ReDim ELEMENT_VECTOR(1 To 1, 1 To 1)
        ELEMENT_VECTOR(1, 1) = ELEMENTS_RNG
    End If
    For i = LBound(ELEMENT_VECTOR, 1) To UBound(ELEMENT_VECTOR, 1)
        KEY_STR = KEY_STR & "|" & ELEMENT_VECTOR(i, 1)
    Next i
End If
'-----------------------------------------------------------------------
If PUB_ADVFN_HASH_OBJ.Exists(KEY_STR) = True Then
    TEMP_MATRIX = PUB_ADVFN_HASH_OBJ(KEY_STR)
    If IsArray(TEMP_MATRIX) = False Then: GoTo ERROR_LABEL
    ADVFN_STATEMENTS_FUNC = CONVERT_STRING_NUMBER_FUNC(TEMP_MATRIX)
    Exit Function
End If

SRC_URL_STR = ADVFN_URL_STRING_FUNC(TICKER_STR, 1, 0, QUATERLY_FLAG)
'---------------------------------------------------------------------------------------
If NO_PERIODS = 0 Then 'First 5 periods
'---------------------------------------------------------------------------------------
    Call PARSE_ADVFN_STATEMENTS_FUNC(SRC_URL_STR, DATA_MATRIX, "")
'---------------------------------------------------------------------------------------
Else
'---------------------------------------------------------------------------------------
    If QUATERLY_FLAG = False Then LOOK_STR = "'start_date'" Else LOOK_STR = "'istart_date'"
    kk = PARSE_ADVFN_STATEMENTS_FUNC(SRC_URL_STR, TEMP_MATRIX, LOOK_STR) 'Base 0
    If IsArray(TEMP_MATRIX) = False Then: GoTo ERROR_LABEL
    kk = kk + 1 'Most Recent Period
    
    NROWS = UBound(TEMP_MATRIX, 1)
    If NO_PERIODS > kk Then: NO_PERIODS = kk
    ReDim DATA_MATRIX(1 To NROWS, 1 To NO_PERIODS + 1)
    For i = 1 To NROWS: DATA_MATRIX(i, 1) = TEMP_MATRIX(i, 1): Next i
    k = NO_PERIODS + 1
    kk = kk - 4
    Do
        SRC_URL_STR = ADVFN_URL_STRING_FUNC(TICKER_STR, 1, CStr(kk), QUATERLY_FLAG)
        Call PARSE_ADVFN_STATEMENTS_FUNC(SRC_URL_STR, TEMP_MATRIX, "")
        If IsArray(TEMP_MATRIX) = False Then: Exit Do
        For jj = UBound(TEMP_MATRIX, 2) To 2 Step -1
            For ii = 1 To UBound(TEMP_MATRIX, 1)
                DATA_MATRIX(ii, k) = TEMP_MATRIX(ii, jj)
            Next ii
            k = k - 1
            If k = 1 Then: Exit For
        Next jj
        kk = kk - 5
    Loop While k > 1
'---------------------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------------------

If IsArray(ELEMENTS_RNG) = True Then: GoSub ELEMENTS_LINE
Call PUB_ADVFN_HASH_OBJ.Add(KEY_STR, DATA_MATRIX)
ADVFN_STATEMENTS_FUNC = CONVERT_STRING_NUMBER_FUNC(DATA_MATRIX)

'---------------------------------------------------------------------------------------
Exit Function
'---------------------------------------------------------------------------------------
ELEMENTS_LINE:
'---------------------------------------------------------------------------------------
    Dim ELEMENTS_HASH_OBJ As New clsTypeHash
    ELEMENTS_HASH_OBJ.SetSize UBound(ELEMENT_VECTOR, 1)
    ELEMENTS_HASH_OBJ.IgnoreCase = False

    NROWS = UBound(DATA_MATRIX, 1)
    For i = 1 To NROWS 'Indexing Entries
        If Trim(DATA_MATRIX(i, 1)) = "" Then: GoTo 1982
        TEMP_STR = LCase(CStr(DATA_MATRIX(i, 1)))
        If ELEMENTS_HASH_OBJ.Exists(TEMP_STR) = False Then
            Call ELEMENTS_HASH_OBJ.Add(TEMP_STR, CStr(i))
        Else 'Check for duplicates such as:
            'depreciation  22|153|
            'amortization  24|154|
            'amortization of intangibles 25|155|
            'deferred income taxes       121|156|
            k = CLng(ELEMENTS_HASH_OBJ(TEMP_STR))
            ELEMENTS_HASH_OBJ(TEMP_STR) = CStr(k) & "|" & CStr(i) & "|"
            'Debug.Print DATA_MATRIX(i, 1), ELEMENTS_HASH_OBJ(TEMP_STR)
        End If
1982:
    Next i

    NROWS = UBound(ELEMENT_VECTOR, 1)
    NCOLUMNS = UBound(DATA_MATRIX, 2) - 1 'Skip Headers
    
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    For i = 1 To NROWS
        TEMP_STR = LCase(CStr(ELEMENT_VECTOR(i, 1)))
        If ELEMENTS_HASH_OBJ.Exists(TEMP_STR) = True Then
            TEMP_STR = ELEMENTS_HASH_OBJ(TEMP_STR)
            If InStr(1, TEMP_STR, "|") > 0 Then
                ii = InStr(1, TEMP_STR, "|")
                jj = InStr(ii + 1, TEMP_STR, "|")
                
                k = CLng(Mid(TEMP_STR, 1, ii - 1)) 'Used 1st
                kk = CLng(Mid(TEMP_STR, ii + 1, jj - (ii + 1))) 'Replace for the 2nd
                
                TEMP_STR = LCase(CStr(ELEMENT_VECTOR(i, 1)))
                ELEMENTS_HASH_OBJ(TEMP_STR) = kk
            Else
                k = CLng(TEMP_STR)
            End If
            For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = DATA_MATRIX(k, j + 1): Next j
        End If
    Next i
    DATA_MATRIX = TEMP_MATRIX: Erase TEMP_MATRIX: Set ELEMENTS_HASH_OBJ = Nothing
'---------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------
ERROR_LABEL:
ADVFN_STATEMENTS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PARSE_ADVFN_STATEMENTS_FUNC
'DESCRIPTION   : Parse ADVFN Statements for US & CAD Public Companies
'LIBRARY       : HTML
'GROUP         : ADVFN
'ID            : 003
'LAST UPDATE   : 29/08/2008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Private Function PARSE_ADVFN_STATEMENTS_FUNC(ByVal SRC_URL_STR As String, _
ByRef DATA_MATRIX As Variant, _
Optional ByVal LOOK_STR As String = "")

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim DATA_STR As String
Dim TEMP_STR As String
Dim REFER_STR As String
Dim TEMP_ARR() As String

On Error GoTo ERROR_LABEL

DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
If DATA_STR = "0" Or DATA_STR = PUB_WEB_DATA_SYSTEM_ERROR_STR Then: _
GoTo ERROR_LABEL

If LOOK_STR = "" Then: GoTo 1982

j = IIf(IsNumeric(LOOK_STR), LOOK_STR, InStr(DATA_STR, LOOK_STR))
i = IIf(j + 32767 <= Len(DATA_STR), 32767, Len(DATA_STR) - j + 1)
TEMP_STR = Mid(DATA_STR, j, i) 'Looking for the last period!!!
        
k = InStr(1, TEMP_STR, "</select") 'Looking for the last period!!!
If k = 0 Then: GoTo ERROR_LABEL
i = k
Do While Mid(TEMP_STR, i, 2) <> "='"
    i = i - 1
Loop
i = i + 2
j = InStr(i, TEMP_STR, "'")

TEMP_STR = Trim(Mid(TEMP_STR, i, j - i))
If Val(TEMP_STR) <> 0 Then 'Find the last period
    PARSE_ADVFN_STATEMENTS_FUNC = CLng(TEMP_STR)
Else
1982:
    PARSE_ADVFN_STATEMENTS_FUNC = 0
End If

DATA_STR = Replace(DATA_STR, Chr(10), "")
DATA_STR = Replace(DATA_STR, Chr(13), "")

ReDim TEMP_ARR(1 To 1)

k = 0
i = InStr(1, DATA_STR, "INDICATORS")
If i = 0 Then: GoTo ERROR_LABEL
i = i + Len("INDICATORS")
j = InStr(i, DATA_STR, "</table>")
If j = 0 Then: GoTo ERROR_LABEL
j = j + Len("</table>")

DATA_STR = Mid(DATA_STR, i, j - i)

i = 1
l = 1
Do
    TEMP_STR = ""
    Do
        i = InStr(i, DATA_STR, "'>")
        If i = 0 Then GoTo 1983
        i = i + Len("'>")
        j = InStr(i, DATA_STR, "<")
        If j = 0 Then GoTo 1983
        REFER_STR = Trim(Mid(DATA_STR, i, j - i))
        If REFER_STR <> "" And REFER_STR <> "*" Then
            TEMP_STR = TEMP_STR & REFER_STR & "|"
        End If
    Loop Until Mid(DATA_STR, j, Len("</td></tr>")) = "</td></tr>"
    
    k = k + 1
    ReDim Preserve TEMP_ARR(1 To k)
    TEMP_ARR(k) = TEMP_STR
1983:
    l = l + 1
    If l > 500 Then: GoTo ERROR_LABEL
Loop Until Mid(DATA_STR, j, Len("</td></tr></table>")) = "</td></tr></table>"

DATA_MATRIX = SPLIT_ARRAY_FUNC(TEMP_ARR, "|", "", 1)
Erase TEMP_ARR

Exit Function
ERROR_LABEL:
Call INITIALIZE_ADVFN_HASH_TABLE_FUNC(True)
PARSE_ADVFN_STATEMENTS_FUNC = Err.number
End Function

Function ADVFN_STATEMENT_ELEMENTS_FUNC(Optional ByVal CTYPE As Long = 0, _
Optional ByVal OUTPUT As Long = 0, _
Optional ByVal QUATERLY_FLAG As Boolean = False, _
Optional ByVal NO_PERIODS As Long = 1) 'As of Oct 8th, 2013

Dim i As Long
Dim j As Long

Dim k As Long
Dim l As Long
Dim m As Long
Dim NSIZE As Long
Dim NROWS As Long

Dim HEADER_STR As String
Dim TEMP_GROUP As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
Select Case CTYPE
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
Case 0 'NON-FINANCIAL INSTITUTION
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
    Select Case OUTPUT
    Case 0
        NSIZE = 18: GoSub LOAD_LINE
    Case Else
        If OUTPUT = 1 Then
            'INDICATORS
            HEADER_STR = IIf(QUATERLY_FLAG = True, "quarter", "year") & _
            " end date|date preliminary data loaded|" & _
            "earnings period indicator|quarterly indicator|basic earnings indicator|" & _
            "template indicator|preliminary full context ind|projected fiscal year date|" & _
            "number of months last report period|"
        ElseIf OUTPUT = 2 Then
            'INCOME STATEMENT
            HEADER_STR = "operating revenue|total revenue|adjustments to revenue|" & _
            "cost of sales|cost of sales with depreciation|gross margin|" & _
            "gross operating profit|Research & Development (R&D) Expense|" & _
            "Selling, General & Administrative (SG&A) Expense|advertising|" & _
            "operating income|EBITDA|depreciation|depreciation (unrecognized)|" & _
            "amortization|amortization of intangibles|operating profit after depreciation|" & _
            "interest income|earnings from equity interest|other income net|" & _
            "income, acquired in process r&a|Income, Restructuring and M&A|" & _
            "other special charges|special income charges|EBIT|interest expense|" & _
            "pre-tax income|income taxes|minority interest|" & _
            "pref. securities of subsid. trust|income before income taxes|" & _
            "net income (continuing operations)|net income (discontinued operations)|" & _
            "net income (total operations)|extraordinary income/losses|" & _
            "income from cum. effect of acct. change|income from tax loss carryforward|" & _
            "other gains/losses|total net income|normalized income|" & _
            "net income available for common|preferred dividends|excise taxes|" & _
            "Basic EPS (Continuing)|Basic EPS (Discontinued)|" & _
            "Basic EPS from Total Operations|Basic EPS (Extraordinary Items)|" & _
            "Basic EPS (Cum. Effect of Acct. Change)|Basic EPS (Tax Loss Carry Forward)|" & _
            "Basic EPS (Other Gains/Losses)|Basic EPS - Total|Basic EPS - Normalized|" & _
            "Diluted EPS (Continuing)|Diluted EPS (Discontinued)|" & _
            "Diluted EPS from Total Operations|Diluted EPS (Extraordinary)|" & _
            "Diluted EPS (Cum. Effect of Acct. Change)|Diluted EPS (Tax Loss Carry Forward)|" & _
            "Diluted EPS (Other Gains/Losses)|Diluted EPS - Total|Diluted EPS - Normalized|" & _
            "Dividends Paid Per Share (DPS)|"
        ElseIf OUTPUT = 3 Then
            'INCOME STATEMENT (YEAR-TO-DATE)
            HEADER_STR = "Revenue (YTD)|Net Income from Total Operations (YTD)|" & _
            "EPS from Total Operations (YTD)|Dividends Paid Per Share (YTD)|"
        ElseIf OUTPUT = 4 Then
            'BALANCE SHEET --> ASSETS
            HEADER_STR = "cash & equivalents|restricted cash|marketable securities|" & _
            "accounts receivable|loans receivable|other receivable|receivables|" & _
            "inventories, raw materials|inventories, work in progress|" & _
            "inventories, purchased components|inventories, finished goods|" & _
            "inventories, other|inventories, adjustments & allowances|" & _
            "inventories|prepaid expenses|current defered income taxes|" & _
            "other current assets|total current assets|land and improvements|" & _
            "building and improvements|machinery, furniture & equipment|" & _
            "construction in progress|other fixed assets|total fixed assets|" & _
            "gross fixed assets|accumulated depreciation|net fixed assets|" & _
            "intangibles|cost in excess|non-current deferred income taxes|" & _
            "other non-current assets|total non-current assets|total assets|" & _
            "inventory valuation method|"
        ElseIf OUTPUT = 5 Then
            'BALANCE SHEET --> EQUITY & LIABILITIES
            HEADER_STR = "accounts payable|notes payable|short-term debt|" & _
            "accrued expenses|accrued liabilities|deferred revenues|" & _
            "current deferred income taxes|other current liabilities|" & _
            "total current liabilities|long-term debt|capital lease obligations|" & _
            "deferred income taxes|other non-current liabilities|" & _
            "minority interest liability|preferred secur. of subsid. trust|" & _
            "preferred equity outside stock equity|total non-current liabilities|" & _
            "total liabilities|preferred stock equity|common stock equity|" & _
            "common par|additional paid-in capital|cumulative translation adjustments|" & _
            "retained earnings|treasury stock|other equity adjustments|" & _
            "total capitalization|total equity|total liabilities & stock equity|" & _
            "cash flow|working capital|free cash flow|invested capital|" & _
            "shares out (common class only)|preferred shares|total ordinary shares|" & _
            "total common shares out|treasury shares|basic weighted shares|" & _
            "diluted weighted shares|number of employees|number of part-time employees|"
        ElseIf OUTPUT = 6 Then
            'CASH-FLOW STATEMENT --> OPERATING ACTIVITIES
            HEADER_STR = "net income/loss|depreciation|amortization|" & _
            "amortization of intangibles|deferred income taxes|operating gains|" & _
            "extraordinary gains|(increase) decrease in receivables|" & _
            "(increase) decrease in inventories|(increase) decrease in prepaid expenses|" & _
            "(increase) decrease in other current assets|decrease (increase) in payables|" & _
            "decrease (increase) in other current liabilities|" & _
            "decrease (increase) in other working capital|other non-cash items|" & _
            "net cash from continuing operations|net cash from discontinued operations|" & _
            "net cash from total operating activities|"
        ElseIf OUTPUT = 7 Then
            'CASH-FLOW STATEMENT --> INVESTING ACTIVITIES
            HEADER_STR = "sale of property, plant & equipment|sale of long-term investments|" & _
            "sale of short-term investments|purchase of property, plant & equipment|" & _
            "acquisitions|purchase of long-term investments|" & _
            "purchase of short-term investments|other investing changes, net|" & _
            "cash from discontinued investing activities|" & _
            "net cash from investing activities|"
        ElseIf OUTPUT = 8 Then
            'CASH-FLOW STATEMENT --> FINANCING ACTIVITIES
            HEADER_STR = "issuance of debt|issuance of capital stock|" & _
            "repayment of long-term debt|repurchase of capital stock|" & _
            "payment of cash dividends|other financing charges, net|" & _
            "cash from discontinued financing activities|" & _
            "net cash from financing activities|"
        ElseIf OUTPUT = 9 Then
            'CASH-FLOW STATEMENT --> NET CASH FLOW
            HEADER_STR = "effect exchange rate changes|net change in cash & equivalents|" & _
            "cash at beginning of period|cash end of period|foreign sales|" & _
            "domestic sales|auditor name|auditor report|"
        ElseIf OUTPUT = 10 Then
            'RATIOS CALCULATIONS --> PROFIT MARGINS
            HEADER_STR = "Close PE Ratio|High PE Ratio|Low PE Ratio|" & _
            "gross profit margin|pre-tax profit margin|post-tax profit margin|" & _
            "net profit margin|interest coverage (cont. operations)|" & _
            "interest as % of invested capital|effective tax rate|income per employee|"
        ElseIf OUTPUT = 11 Then
            'RATIOS CALCULATIONS --> NORMALIZED RATIOS
            HEADER_STR = "Normalized Close PE Ratio|Normalized High PE Ratio|" & _
            "Normalized Low PE Ratio|normalized net profit margin|Normalized ROE|" & _
            "Normalized ROA|Normalized ROCI|normalized income per employee|"
        ElseIf OUTPUT = 12 Then
            'RATIOS CALCULATIONS --> SOLVENCY RATIOS
            HEADER_STR = "quick ratio|current ratio|payout ratio|" & _
            "total debt/equity ratio|long-term debt/total capital|"
        ElseIf OUTPUT = 13 Then
            'RATIOS CALCULATIONS --> EFFICIENCY RATIOS
            HEADER_STR = "leverage ratio|asset turnover|cash as % of revenue|" & _
            "receivables as % of revenue|SG&A as % of Revenue|R&D as % of Revenue|"
        ElseIf OUTPUT = 14 Then
            'RATIOS CALCULATIONS --> ACTIVITY RATIOS
            HEADER_STR = "revenue per $ cash|revenue per $  plant (net)|" & _
            "revenue per $ common equity|revenue per $ invested capital|"
        ElseIf OUTPUT = 15 Then
            'RATIOS CALCULATIONS --> LIQUIDITY RATIOS
            HEADER_STR = "receivables turnover|inventory turnover|" & _
            "receivables per day sales|sales per $ receivables|" & _
            "sales per $ inventory|revenue/assets|" & _
            "number of days cost of goods in inventory|current assets per share|" & _
            "total assets per share|intangibles as % of book-value|" & _
            "inventory as % of revenue|"
        ElseIf OUTPUT = 16 Then
            'RATIOS CALCULATIONS --> CAPITAL STRUCTURE RATIOS
            HEADER_STR = "long-term debt per share|current liabilities per share|" & _
            "cash per share|LT-Debt to Equity Ratio|LT-Debt as % of Invested Capital|" & _
            "LT-Debt as % of Total Debt|total debt as % total assets|" & _
            "working captial as % of equity|revenue per share|book value per share|" & _
            "tangible book value per share|price/revenue ratio|price/equity ratio|" & _
            "price/tangible book ratio|working capital as % of price|"
        ElseIf OUTPUT = 17 Then
            'RATIOS CALCULATIONS --> PROFITABILITY
            HEADER_STR = "working capital per share|cash flow per share|" & _
            "free cash flow per share|Return on Stock Equity (ROE)|" & _
            "Return on Capital Invested (ROCI)|Return on Assets (ROA)|" & _
            "price/cash flow ratio|price/free cash flow ratio|sales per employee|"
        Else 'If OUTPUT = 18 Then
            'RATIOS CALCULATIONS --> AGAINST THE INDUSTRY RATIOS
            HEADER_STR = "% of sales-to-industry|% of earnings-to-industry|" & _
            "% of EPS-to-Industry|% of price-to-industry|% of PE-to-Industry|" & _
            "% of price/book-to-industry|% of price/sales-to-industry|" & _
            "% of price/cashflow-to-industry|% of pric/free cashlow-to-industry|" & _
            "% of debt/equity-to-industry|% of current ratio-to-industry|" & _
            "% of gross profit margin-to-industry|% of pre-tax profit margin-to-industry|" & _
            "% of post-tax profit margin-to-industry|% of net profit margin-to-industry|" & _
            "% of ROE-to-Industry|% of leverage-to-industry|"
        End If
        GoSub SPLIT_LINE
    End Select
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
Case Else 'FINANCIAL INSTITUTION
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
    Select Case OUTPUT
    Case 0
        NSIZE = 9: GoSub LOAD_LINE
    Case Else
        If OUTPUT = 1 Then
            'INDICATORS
            HEADER_STR = IIf(QUATERLY_FLAG = True, "quarter", "year") & _
            " end date|date preliminary data loaded|" & _
            "earnings period indicator|quarterly indicator|" & _
            "basic earnings indicator|format indicator|preliminary full context indicator|" & _
            "projected fiscal year end date|number of months since last reporting period|"
        ElseIf OUTPUT = 2 Then
            'INCOME STATEMENT
            HEADER_STR = "loans|investment securities|lease financing income|other interest income|" & _
            "federal funds sold (purchased)|interest bearing deposits|loans held for resale|" & _
            "trading account securities|time deposits placed|other money market investments|" & _
            "total money market investments|total interest income|deposits|short-term deposits|" & _
            "long-term deposits|federal funds purchased (securities sold)|capitalized lease obligations|" & _
            "other interest expense|total interest expense|net interest income (expense)|provision for loan loss|" & _
            "trust fees by commissions|service charge on deposit accounts|other service charges|security transactions|" & _
            "premiums earned|net realized capital gains|investment banking profit|other non-interest income|" & _
            "total non-interest income|salaries and employee benefits|net occupancy expense|promotions and advertising|" & _
            "property liability insurance claims|policy acquisition costs|amortization deferred policy acquisition cost|" & _
            "current and future benefits|other non-interest expense|total non-interest expense|" & _
            "premium tax credit|minority interest|income taxes|Income, Aquired in Process R&D|" & _
            "Income,Restructuring and M&A|other special charges|special income (charges)|" & _
            "net income from continuing operations|net income from discontinued operations|" & _
            "net income from total operations|extraordinary income losses|" & _
            "income from cumumulative effect of accounting change|income from tax loss carryforward|other gains (losses)|" & _
            "total net income|normalized income|net income available for common|" & _
            "preferred dividends|Basic EPS (Continuing)|Basic EPS (Discontinued)|Basic EPS from Total Operations)|" & _
            "Basic EPS (Extraordinary)|Basic EPS (Cum. Effect of Acc. Change)|Basic EPS (Tax Loss Carry Forward)|" & _
            "Basic EPS (Other Gains/Losses)|Basic EPS - Total|Basic EPS - Normalized|" & _
            "Diluted EPS (Continuing)|Diluted EPS (Discontinued)|Diluted EPS from Total Operations|" & _
            "Diluted EPS (Extraordinary)|Diluted EPS (Cum. Effect of Acc. Change)|" & _
            "Diluted EPS (Tax Loss Carry Forward)|Diluted EPS (Other Gains/Losses)|" & _
            "Diluted EPS - Total|Diluted EPS - Normalized|dividends paid per share|"
        ElseIf OUTPUT = 3 Then
            'INCOME STATEMENT (YEAR-TO-DATE)
            HEADER_STR = "Revenues (YTD)|Net income from Total Operations (YTD)|" & _
            "Diluted EPS from Total Operations (YTD)|Dividends Paid Per Share (YTD)|"
        ElseIf OUTPUT = 4 Then
            'BALANCE SHEET --> ASSETS
            HEADER_STR = "cash and due from banks|restricted cash|" & _
            "federal funds sold (securities purchased)|" & _
            "interest bearing deposits at other banks|investment securities, net|" & _
            "loans|unearnedp remiums|allowance for loans and lease losses|net loans|" & _
            "premises & equipment|due from customers acceptance|" & _
            "trading account securities|other receivables|accrued interest|" & _
            "deferred acquisition cost|accrued investment income|" & _
            "separate account business|time deposits placed|intangible assets|" & _
            "other assets|total assets|"
        ElseIf OUTPUT = 5 Then
            'BALANCE SHEET --> EQUITY & LIABILITIES
            HEADER_STR = "non-interest bearing deposits|interest bearing deposits|short -term debt|" & _
            "other liabilities|bankers acceptance outstanding|federal funds purchased (securities sold)|" & _
            "accrued taxes|accrued interest payables|other payables|capital lease obligations|" & _
            "claims and claim expense|future policy benefits|unearned premiums|" & _
            "policy holder funds|participating policy holder equity|separate accounts business|" & _
            "minority interest|long-term debt|preferred stock equity|common stock equity|" & _
            "common par|additional paid in capital|cumulative translation adjustment|" & _
            "retained earnings|treasury stock|other equity adjustments|foreign currency adjustments|" & _
            "net unrealized loss (gain) on investments|net unrealized loss (gain) on foreign currency|" & _
            "net other unearned losses (gains)|total equity|total liabilities|shares outstanding common class only|" & _
            "preferred shares|total ordinary shares|total common shares outstanding|treasury shares|" & _
            "basic weighted shares outstanding|diluted weighted shares outstanding|number of employees|" & _
            "number of part-time employees|"
        ElseIf OUTPUT = 6 Then
            'CASH FLOW STATEMENT --> OPERATING ACTIVITIES
            HEADER_STR = "net income earnings|provision for loan losses|" & _
            "depreciation and amortization|deferred income taxes|" & _
            "change in assets - receivables|change in liabilities - payables|" & _
            "investment securities gain|net policy acquisition costs|" & _
            "realized investment gains|net premiums receivables|change in income taxes|" & _
            "other non-cash items|net cash from operating activities|"
        ElseIf OUTPUT = 7 Then
            'CASH FLOW STATEMENT --> INVESTING ACTIVITIES
            HEADER_STR = "proceeds from sale - material investment|" & _
            "purchase of investment securities|net increase federal funds sold|" & _
            "purchase of property & equipment|acquisitions|other investing changes net|" & _
            "net cash from investing activities|"
        ElseIf OUTPUT = 8 Then
            'CASH FLOW STATEMENT --> FINANCING ACTIVITIES
            HEADER_STR = "net change in deposits|cash dividends paid|" & _
            "repayment of long-term debt|change of short-term debt|" & _
            "issuance of long-term debt|issuance of preferred stock|" & _
            "issuance of common stock|purchase of treasury stock|" & _
            "other financing activities|net cash from financing activities|"
        Else 'If OUTPUT = 9 Then
            'CASH FLOW STATEMENT --> NET CASH FLOW
            HEADER_STR = "effect of exchange rate changes|" & _
            "net change in cash & equivalents|cash at beginning of period|" & _
            "cash at end of period|total risk-based capital ratio|" & _
            "auditors name|auditors report|"
        End If
        GoSub SPLIT_LINE
    End Select
'--------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------

ADVFN_STATEMENT_ELEMENTS_FUNC = TEMP_VECTOR

'--------------------------------------------------------------------------
Exit Function
'--------------------------------------------------------------------------
SPLIT_LINE:
'--------------------------------------------------------------------------
k = 0
For l = 1 To Len(HEADER_STR)
    If Mid(HEADER_STR, l, 1) = "|" Then: k = k + 1
Next l
ReDim TEMP_VECTOR(1 To k, 1 To 1) As String
i = 1
For l = 1 To k
    j = InStr(i, HEADER_STR, "|")
    TEMP_VECTOR(l, 1) = Mid(HEADER_STR, i, j - i)
    i = j + 1
Next l
'--------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------
LOAD_LINE:
'--------------------------------------------------------------------------
ReDim TEMP_GROUP(1 To NSIZE, 1 To 2)
m = 0
For l = 1 To NSIZE
    TEMP_GROUP(l, 1) = ADVFN_STATEMENT_ELEMENTS_FUNC(0, l, QUATERLY_FLAG)
    TEMP_GROUP(l, 2) = UBound(TEMP_GROUP(l, 1))
    m = m + TEMP_GROUP(l, 2)
Next l
If NO_PERIODS <= 0 Then: NO_PERIODS = 1
NROWS = m * NO_PERIODS
'--------------------------------------------------------------------------
If NROWS = m Then '1 Period
'--------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
    l = 1
    For i = 1 To NSIZE
        m = TEMP_GROUP(i, 2)
        For j = 1 To m
            TEMP_VECTOR(l, 1) = TEMP_GROUP(i, 1)(j, 1)
            l = l + 1
        Next j
    Next i
'--------------------------------------------------------------------------
Else
'--------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To NROWS, 1 To 2)
    l = 1
    For i = 1 To NSIZE
        m = TEMP_GROUP(i, 2)
        For j = 1 To m
            For k = 0 To NO_PERIODS - 1
                TEMP_VECTOR(l, 1) = k 'Period
                TEMP_VECTOR(l, 2) = TEMP_GROUP(i, 1)(j, 1) 'Account
                l = l + 1
            Next k
        Next j
    Next i
'--------------------------------------------------------------------------
End If
'--------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------
ERROR_LABEL:
ADVFN_STATEMENT_ELEMENTS_FUNC = Err.number
End Function


Function ADVFN_URL_STRING_FUNC(ByVal TICKER_STR As String, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal NO_PERIODS As Integer = 0, _
Optional ByVal QUATERLY_FLAG As Boolean = False)

Dim i As Long
Dim j As Long

Dim SRC_URL_STR As String
Dim EXCHANGE_STR As String
Dim SYMBOL_STR As String

On Error GoTo ERROR_LABEL

i = InStr(1, TICKER_STR, ":")
If i = 0 Then: GoTo ERROR_LABEL
EXCHANGE_STR = Mid(TICKER_STR, 1, i - 1)
i = i + 1
j = Len(TICKER_STR)
SYMBOL_STR = Mid(TICKER_STR, i, j - i + 1)
SRC_URL_STR = "http://" & PUB_ADVFN_SERVER_STR & ".advfn.com/"
SRC_URL_STR = SRC_URL_STR & "exchanges/" & EXCHANGE_STR & "/" & SYMBOL_STR & "/financials"

'-----------------------------------------------------------------------------------------
Select Case VERSION
'-----------------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------------
    ADVFN_URL_STRING_FUNC = SRC_URL_STR
'-----------------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------------
    If NO_PERIODS = 0 Then
        If QUATERLY_FLAG = False Then
            ADVFN_URL_STRING_FUNC = SRC_URL_STR & "?btn=annual_reports&mode=company_data"
        Else
            ADVFN_URL_STRING_FUNC = SRC_URL_STR & "?btn=quarterly_reports&mode=company_data"
        End If
    Else
        If QUATERLY_FLAG = False Then
            ADVFN_URL_STRING_FUNC = SRC_URL_STR & "?btn=start_date&start_date=" & CStr(NO_PERIODS - 1) & "&mode=annual_reports"
        Else
            ADVFN_URL_STRING_FUNC = SRC_URL_STR & "?btn=istart_date&istart_date=" & CStr(NO_PERIODS - 1) & "&mode=quarterly_reports"
        End If
    End If
'-----------------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ADVFN_URL_STRING_FUNC = ""
End Function

Sub INITIALIZE_ADVFN_HASH_TABLE_FUNC(Optional ByVal RESET_FLAG As Boolean = False) '(Optional ByVal SIZE As Long = 10000)
On Error Resume Next
If RESET_FLAG = True Or PUB_ADVFN_HASH_OBJ Is Nothing Then
    Set PUB_ADVFN_HASH_OBJ = New clsTypeHash
    PUB_ADVFN_HASH_OBJ.SetSize 10000 'SIZE
    PUB_ADVFN_HASH_OBJ.IgnoreCase = False
End If
End Sub

Sub PRINT_ADVFN_STATISTICS()

Dim i As Long

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_RNG As Excel.Range
Dim DST_RNG As Excel.Range
Dim TEMP_RNG As Excel.Range
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

Set DATA_RNG = Excel.Application.InputBox("Symbols", _
                   "ADVFN", , , , , , 8)
If DATA_RNG Is Nothing Then: Exit Sub
Call EXCEL_TURN_OFF_EVENTS_FUNC

Set DST_RNG = _
WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC("_"), _
ActiveWorkbook).Cells(3, 3)

For i = 0 To 13
    
    TEMP_MATRIX = ADVFN_KEY_STATISTICS_FUNC(DATA_RNG, i)
    If IsArray(TEMP_MATRIX) = False Then: GoTo 1983

    SROW = LBound(TEMP_MATRIX, 1)
    NROWS = UBound(TEMP_MATRIX, 1)
    
    SCOLUMN = LBound(TEMP_MATRIX, 2)
    NCOLUMNS = UBound(TEMP_MATRIX, 2)
    
    Set TEMP_RNG = Range(DST_RNG.Cells(SROW, SCOLUMN), _
      DST_RNG.Cells(NROWS, NCOLUMNS))
    TEMP_RNG.value = TEMP_MATRIX
    GoSub FORMAT_LINE
    Set DST_RNG = DST_RNG.Offset(NROWS - SROW + 2, 0)
1983:
Next i

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
        With .Columns(1)
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


