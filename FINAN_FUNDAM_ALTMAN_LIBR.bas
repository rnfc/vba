Attribute VB_Name = "FINAN_FUNDAM_ALTMAN_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Sub PRINT_HISTORICAL_ALTMAN_Z_SCORE_FUNC()
'ByRef SRC_RNG As Excel.Range, _
ByVal NO_ASSETS As Long, _
ByVal SCOLUMN As Long, _
ByRef FINDEX_ARR As Variant)

Dim i As Long
Dim j As Long

Dim SROW As Long
Dim SCOLUMN As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim NO_ASSETS As Long

Dim TEMP_RNG As Excel.Range
Dim DATA_RNG As Excel.Range
Dim FINDEX_ARR As Variant
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

FINDEX_ARR = Array(4, 4, 6, 2, 3, _
                2, 2, 2, 2, 5, _
                5, 1, 3, 3, 1, _
                3, 1, 1, 5, 1, _
                1, 1, 1, 1, 1, _
                1, 1, 1, 1, 1, _
                1, 1, 1, 3, 3, _
                1, 3, 1, 1, 1, _
                1, 1, 1, 1, 1, _
                1, 1, 1, 3, 3, _
                5, 1, 5, 2)

TICKERS_VECTOR = Excel.Application.InputBox("Symbol(s)", "Altman Z-Score", , , , , , 8)
'If TICKERS_VECTOR = False Then: Exit Sub
Call EXCEL_TURN_OFF_EVENTS_FUNC
TEMP_MATRIX = TICKERS_VECTOR
If IsArray(TEMP_MATRIX) = False Then
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TEMP_MATRIX
Else
    Erase TEMP_MATRIX
    If UBound(TICKERS_VECTOR, 1) = 1 Then: TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
End If
NO_ASSETS = UBound(TICKERS_VECTOR, 1)

TEMP_MATRIX = HISTORICAL_ALTMAN_Z_SCORE_FUNC(TICKERS_VECTOR)
If IsArray(TEMP_MATRIX) = False Then: Exit Sub
        
SROW = LBound(TEMP_MATRIX, 1)
NROWS = UBound(TEMP_MATRIX, 1)
            
SCOLUMN = LBound(TEMP_MATRIX, 2)
NCOLUMNS = UBound(TEMP_MATRIX, 2)
            
Set DATA_RNG = WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC("_"), ActiveWorkbook).Cells(1, 1)
Set DATA_RNG = Range(DATA_RNG.Cells(SROW, SCOLUMN), DATA_RNG.Cells(NROWS, NCOLUMNS))
With DATA_RNG
    .RowHeight = 15
    .ColumnWidth = 10
    .VerticalAlignment = xlCenter
End With
DATA_RNG = TEMP_MATRIX
Set TEMP_RNG = DATA_RNG.Columns(10 + 1) 'Quaterly Data 10
With TEMP_RNG.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .WEIGHT = xlMedium
End With

NCOLUMNS = DATA_RNG.Columns.COUNT
NROWS = DATA_RNG.Rows.COUNT

j = LBound(FINDEX_ARR)
For i = 1 To NROWS Step NO_ASSETS + 1
    Set TEMP_RNG = Range(DATA_RNG.Cells(i + 1, 1), DATA_RNG.Cells(i + 1, NCOLUMNS)).Offset(-1, 0)
    With TEMP_RNG
        .Cells(1, 1).InsertIndent 1
        With .Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With .Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .Bold = True
        End With
    End With
    Set TEMP_RNG = Range(DATA_RNG.Cells(i + 1, 1), DATA_RNG.Cells(i + NO_ASSETS, NCOLUMNS))
    With TEMP_RNG
        .Columns(1).InsertIndent 2
        .Rows.Group
    End With
    
    If (UBound(FINDEX_ARR) - LBound(FINDEX_ARR) + 1) >= j Then
        Set TEMP_RNG = Range(DATA_RNG.Cells(i + 1, 2), DATA_RNG.Cells(i + NO_ASSETS, NCOLUMNS))
        With TEMP_RNG
            .HorizontalAlignment = xlRight
            Select Case FINDEX_ARR(j)
            Case 1: .NumberFormat = "#,##0"
            Case 2: .NumberFormat = "#,##0.00"
            Case 3: .NumberFormat = "0.00%"
            Case 4: .NumberFormat = "mmm-yy"
            Case 5: .NumberFormat = "0.00""x"""
            Case Else
                .HorizontalAlignment = xlCenter
                .NumberFormat = "0"
            End Select
            j = j + 1
        End With
    End If
Next i

Call EXCEL_TURN_ON_EVENTS_FUNC

Exit Sub
ERROR_LABEL:
End Sub

Private Function HISTORICAL_ALTMAN_Z_SCORE_FUNC(ByRef TICKERS_RNG As Variant)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim m As Long
Dim n As Long
Dim o As Long
Dim p As Long 'Position Row

Dim ii As Long
Dim jj As Long

Dim r(1 To 7) As Long 'Row #

Dim V1 As Integer
Dim V2 As Integer

Dim Y_VAL As Long
Dim Q_VAL As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ERROR_STR As String
Dim TICKER_STR As String
Dim HEADING_STR As String
Dim HEADINGS_STR As String

Dim SKIP_FLAG As Boolean

Dim DATE_VAL As Variant
Dim DATA_ARR(1 To 7) As Variant
Dim INDEX_OBJ As New Collection
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

ERROR_STR = "-"
TICKERS_VECTOR = TICKERS_RNG
If UBound(TICKERS_VECTOR, 1) = 1 Then
    TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
End If
NROWS = UBound(TICKERS_VECTOR, 1)
NCOLUMNS = 10 + 20 + 1

V2 = 0: SKIP_FLAG = False
NSIZE = (NROWS + 1) * 54
ReDim TEMP_MATRIX(1 To NSIZE, 1 To NCOLUMNS)
For i = 1 To NSIZE: For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = "": Next j: Next i

i = 1: HEADING_STR = "End Period": Y_VAL = 5196: Q_VAL = 8006: GoSub LOAD2_LINE: GoSub INDEX_LINE

i = i + NROWS + 1: HEADING_STR = "Data Loaded": Y_VAL = 5206: Q_VAL = 8026: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Ending Quarter": Y_VAL = 0: Q_VAL = 8066: GoSub LOAD2_LINE: GoSub INDEX_LINE


i = i + NROWS + 1: HEADING_STR = "Basic EPS": Y_VAL = 5786: Q_VAL = 9186: GoSub LOAD2_LINE: GoSub INDEX_LINE
h = 4: V1 = 6: r(1) = CLng(INDEX_OBJ("Basic EPS")): r(2) = r(1): i = i + NROWS + 1: HEADING_STR = "Basic EPS Growth": SKIP_FLAG = True: GoSub LOAD3_LINE: GoSub INDEX_LINE

r(1) = CLng(INDEX_OBJ("Data Loaded")): i = i + NROWS + 1: HEADING_STR = "Closing Stock Price": GoSub LOAD1_LINE: GoSub INDEX_LINE

i = i + NROWS + 1: HEADING_STR = "PE Close": Y_VAL = 7146: Q_VAL = 11906: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "PE High": Y_VAL = 7156: Q_VAL = 11926: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "PE Low": Y_VAL = 7166: Q_VAL = 11946: GoSub LOAD2_LINE: GoSub INDEX_LINE

i = i + NROWS + 1: HEADING_STR = "xPE-EV": GoSub INDEX_LINE 'x PE-EV (EV / NET INCOME)
i = i + NROWS + 1: HEADING_STR = "xCFFO": GoSub INDEX_LINE 'x CFFO (Market Cap / Net Cash From Total Operating Activities)

i = i + NROWS + 1: HEADING_STR = "Revenue": Y_VAL = 5296: Q_VAL = 8206: GoSub LOAD2_LINE: GoSub INDEX_LINE
h = 4: V1 = 6: r(1) = CLng(INDEX_OBJ("Revenue")): r(2) = r(1): i = i + NROWS + 1: HEADING_STR = "Revenue Growth YoY": SKIP_FLAG = True: GoSub LOAD3_LINE: GoSub INDEX_LINE 'Revenue Growth YoY (Qt1 / Qt5)
h = 1: V1 = 6: r(1) = CLng(INDEX_OBJ("Revenue")): r(2) = r(1): i = i + NROWS + 1: HEADING_STR = "Revenue Growth QoQ": SKIP_FLAG = True: GoSub LOAD3_LINE: GoSub INDEX_LINE

i = i + NROWS + 1: HEADING_STR = "Working Capital": Y_VAL = 6586: Q_VAL = 10786: GoSub LOAD2_LINE: GoSub INDEX_LINE
h = 4: V1 = 6: r(1) = CLng(INDEX_OBJ("Working Capital")): r(2) = r(1): i = i + NROWS + 1: HEADING_STR = "WC Growth": SKIP_FLAG = True: GoSub LOAD3_LINE: GoSub INDEX_LINE

i = i + NROWS + 1: HEADING_STR = "CAPEX": Y_VAL = 6916: Q_VAL = 11446: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Depreciation": Y_VAL = 6716: Q_VAL = 11046: GoSub LOAD2_LINE: GoSub INDEX_LINE


h = 0: V1 = 8: r(2) = CLng(INDEX_OBJ("CAPEX")): r(1) = CLng(INDEX_OBJ("Depreciation")): i = i + NROWS + 1: HEADING_STR = "Capex/Depreciation": SKIP_FLAG = False: GoSub LOAD3_LINE: GoSub INDEX_LINE

i = i + NROWS + 1: HEADING_STR = "Receivables": Y_VAL = 6006: Q_VAL = 9626: GoSub LOAD2_LINE: GoSub INDEX_LINE

i = i + NROWS + 1: HEADING_STR = "Inventory Raw Materials": Y_VAL = 6016: Q_VAL = 9646: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Inventory Work in Process": Y_VAL = 6026: Q_VAL = 9666: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Inventory Purchased Components ": Y_VAL = 6036: Q_VAL = 9686: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Inventory Finished Goods": Y_VAL = 6046: Q_VAL = 9706: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Inventory Other": Y_VAL = 6056: Q_VAL = 9726: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Inventory Ajustments and Allowances": Y_VAL = 6066: Q_VAL = 9746: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Inventory": Y_VAL = 6076: Q_VAL = 9766: GoSub LOAD2_LINE: GoSub INDEX_LINE


i = i + NROWS + 1: HEADING_STR = "Net Fixed Assets": Y_VAL = 6206: Q_VAL = 10026: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Total Assets": Y_VAL = 6266: Q_VAL = 10146: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Total Liabilities": Y_VAL = 6456: Q_VAL = 10526: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Retained Earnings": Y_VAL = 6516: Q_VAL = 10646: GoSub LOAD2_LINE: GoSub INDEX_LINE

i = i + NROWS + 1: HEADING_STR = "EBITDA": Y_VAL = 5396: Q_VAL = 8406: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "EBIT": Y_VAL = 5526: Q_VAL = 8666: GoSub LOAD2_LINE: GoSub INDEX_LINE

h = 4: V1 = 6: r(1) = CLng(INDEX_OBJ("EBIT")): r(2) = r(1): i = i + NROWS + 1: HEADING_STR = "EBIT Growth": SKIP_FLAG = True: GoSub LOAD3_LINE: GoSub INDEX_LINE 'EBIT Growth (Qt1 / Qt5)

h = 0: V1 = 4: r(2) = CLng(INDEX_OBJ("EBIT")): r(1) = CLng(INDEX_OBJ("Revenue")): i = i + NROWS + 1: HEADING_STR = "EBIT Margin": SKIP_FLAG = False: GoSub LOAD3_LINE: GoSub INDEX_LINE 'EBIT / Revenue
i = i + NROWS + 1: HEADING_STR = "Net Income": Y_VAL = 5666: Q_VAL = 8946: GoSub LOAD2_LINE: GoSub INDEX_LINE

'
h = 0: V1 = 4: r(2) = CLng(INDEX_OBJ("Net Income")): r(1) = CLng(INDEX_OBJ("Revenue")): i = i + NROWS + 1: HEADING_STR = "Net Margin": SKIP_FLAG = False: GoSub LOAD3_LINE: GoSub INDEX_LINE 'Net Income / Revenue

i = i + NROWS + 1: HEADING_STR = "Net Cash From Total Operating Activities (CFFO)": Y_VAL = 6876: Q_VAL = 11366: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Weight Common Shares": Y_VAL = 6666: Q_VAL = 10946: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Long Term Debt": Y_VAL = 6376: Q_VAL = 10366: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Short Term Debt": Y_VAL = 6306: Q_VAL = 10226: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Cash": Y_VAL = 5946: Q_VAL = 9506: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Restricted Cash": Y_VAL = 6316: Q_VAL = 10246: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Marketable Securities": Y_VAL = 6326: Q_VAL = 10266: GoSub LOAD2_LINE: GoSub INDEX_LINE

 '
h = 0: V1 = 9: r(2) = CLng(INDEX_OBJ("Weight Common Shares")): r(1) = CLng(INDEX_OBJ("Closing Stock Price")): i = i + NROWS + 1: HEADING_STR = "Market Cap": SKIP_FLAG = False: GoSub LOAD3_LINE: GoSub INDEX_LINE 'Weight Common Shares x Stock Price Data Loaded

p = i: h = 0: V1 = 4: r(2) = CLng(INDEX_OBJ("Market Cap")): r(1) = CLng(INDEX_OBJ("Net Cash From Total Operating Activities (CFFO)")): i = CLng(INDEX_OBJ("xCFFO")): HEADING_STR = "xCFFO": SKIP_FLAG = False: GoSub LOAD3_LINE 'Market Cap / Net Cash From Total Operating Activities (CFFO)
i = p

h = 0: V1 = 11: r(2) = CLng(INDEX_OBJ("Marketable Securities")): r(1) = CLng(INDEX_OBJ("Cash")): i = i + NROWS + 1: HEADING_STR = "Cash on BS": SKIP_FLAG = False: GoSub LOAD3_LINE: GoSub INDEX_LINE 'Marketable Securities + Cash
h = 0: V1 = 11: r(2) = CLng(INDEX_OBJ("Short Term Debt")): r(1) = CLng(INDEX_OBJ("Long Term Debt")): i = i + NROWS + 1: HEADING_STR = "Total Debt": SKIP_FLAG = False: GoSub LOAD3_LINE: GoSub INDEX_LINE 'Long Term Debt + Short Term Debt

V1 = 0: r(1) = CLng(INDEX_OBJ("Market Cap")): r(2) = CLng(INDEX_OBJ("Cash on BS")): r(3) = CLng(INDEX_OBJ("Total Debt")): i = i + NROWS + 1: HEADING_STR = "Enterprise Value (EV)": GoSub LOAD4_LINE: GoSub INDEX_LINE 'Market Cap + Total Debt - Cash on BS

p = i: h = 0: V1 = 4: r(2) = CLng(INDEX_OBJ("Enterprise Value (EV)")): r(1) = CLng(INDEX_OBJ("Net Income")): i = CLng(INDEX_OBJ("xPE-EV")): HEADING_STR = "xPE-EV": SKIP_FLAG = False: GoSub LOAD3_LINE '(EV / NET INCOME)
i = p

V2 = 1: h = 0: V1 = 4: r(2) = CLng(INDEX_OBJ("EBIT")): r(1) = CLng(INDEX_OBJ("Enterprise Value (EV)")): i = i + NROWS + 1: HEADING_STR = "EBIT Yield (ttm)": SKIP_FLAG = False: GoSub LOAD3_LINE: GoSub INDEX_LINE 'EBIT / EV:V2 = 0

i = i + NROWS + 1: HEADING_STR = "ROIC (ttm)": GoSub INDEX_LINE 'ROIC

V2 = 2: h = 0: V1 = 4: r(2) = CLng(INDEX_OBJ("Total Debt")): r(1) = CLng(INDEX_OBJ("EBITDA")): i = i + NROWS + 1: HEADING_STR = "Debt/EBITDA(ttm)": SKIP_FLAG = False: GoSub LOAD3_LINE: GoSub INDEX_LINE 'Total Debt / EBITDA:V2 = 0
    
'
h = 0: V1 = 10: r(2) = CLng(INDEX_OBJ("Total Assets")): r(1) = CLng(INDEX_OBJ("Cash on BS")): i = i + NROWS + 1: HEADING_STR = "Invested Capital (IC)": SKIP_FLAG = False: GoSub LOAD3_LINE: GoSub INDEX_LINE 'Total Assets - Cash on BS

p = i: V2 = 1: h = 0: V1 = 4: r(2) = CLng(INDEX_OBJ("EBIT")): r(1) = CLng(INDEX_OBJ("Invested Capital (IC)")): i = CLng(INDEX_OBJ("ROIC (ttm)")): HEADING_STR = "ROIC (ttm)": SKIP_FLAG = False: GoSub LOAD3_LINE 'EBIT / IC:V2 = 0
i = p

V2 = 2: h = 0: V1 = 4: r(2) = CLng(INDEX_OBJ("Enterprise Value (EV)")): r(1) = CLng(INDEX_OBJ("EBITDA")): i = i + NROWS + 1: HEADING_STR = "xEBITDA(ttm)-EV": SKIP_FLAG = False: GoSub LOAD3_LINE: GoSub INDEX_LINE 'EV / EBITDA:V2 = 0

V1 = 1: r(7) = CLng(INDEX_OBJ("Revenue")): r(1) = CLng(INDEX_OBJ("Working Capital")): r(2) = CLng(INDEX_OBJ("Total Assets")): r(6) = CLng(INDEX_OBJ("Total Liabilities")): r(3) = CLng(INDEX_OBJ("Retained Earnings")): r(4) = CLng(INDEX_OBJ("EBIT")): r(5) = CLng(INDEX_OBJ("Market Cap")): i = i + NROWS + 1: HEADING_STR = "Altman Z-Score": GoSub LOAD4_LINE: GoSub INDEX_LINE
'Altman Z-Score for stock in the site http://www.grahaminvestor.com/articles/quantitative-tools/the-altman-z-score/
'n1 = FQ1, Working Capital
'n2 = FQ1, Total Assets
'n3 = FQ1, Retained Earnings
'n4 = FQ1, EBIT
'n5 = FQ2, EBIT
'n6 = FQ3, EBIT
'n7 = FQ4, EBIT
'n8 = n4 + n5 + n6 + n7
'n9 = Market Capitalization
'n10 = Total Liabilities
'n11 = FQ1, Operating Revenue
'n12 = FQ2, Operating Revenue
'n13 = FQ3, Operating Revenue
'n14 = FQ4, Operating Revenue
'n15 = n11 + n12 + n13 + n14
'SpecialExtractio n = 1.2 * (n1 / n2) + 1.4 * (n3 / n2)
'+ 3.3 * (n8 / n2) + 0.6 * (n9 / n10 / 1000) + (n15 / n2)
    
HISTORICAL_ALTMAN_Z_SCORE_FUNC = TEMP_MATRIX

'---------------------------------------------------------------------------------------------------------------------------------
Exit Function
'---------------------------------------------------------------------------------------------------------------------------------
LOAD1_LINE:
'---------------------------------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(i, 1) = HEADING_STR
    For l = 1 To NROWS
        TICKER_STR = TICKERS_VECTOR(l, 1)
        TEMP_MATRIX(i + l, 1) = TICKER_STR
        j = 2
        ii = l + r(1)
        For k = 0 To 9
            DATA_ARR(2) = TEMP_MATRIX(ii, j)
            If ((DATA_ARR(2) <> "") And (IsDate(DATA_ARR(2)))) Then
                DATE_VAL = DATA_ARR(2) 'Using Date of Data Loaded
                DATA_ARR(2) = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR, Year(DATE_VAL), Month(DATE_VAL), Day(DATE_VAL), Year(DATE_VAL), Month(DATE_VAL), Day(DATE_VAL), "d", "A", 0)
                If IsArray(DATA_ARR(2)) = True Then: TEMP_MATRIX(i + l, j) = DATA_ARR(2)(1, 1)
            End If
            j = j + 1
        Next k
        For k = 0 To 19
            DATA_ARR(2) = TEMP_MATRIX(ii, j)
            If ((DATA_ARR(2) <> "") And (IsDate(DATA_ARR(2)))) Then
                DATE_VAL = DATA_ARR(2) 'Using Date of Data Loaded
                DATA_ARR(2) = YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR, Year(DATE_VAL), Month(DATE_VAL), Day(DATE_VAL), Year(DATE_VAL), Month(DATE_VAL), Day(DATE_VAL), "d", "A", 0)
                If IsArray(DATA_ARR(2)) = True Then: TEMP_MATRIX(i + l, j) = DATA_ARR(2)(1, 1)
            End If
            j = j + 1
        Next k
        'MATRIX_STDEV_FUNC(MATRIX_PERCENT_FUNC(YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR,YEAR(DATE_VAL-365),MONTH(DATE_VAL-365),DAY(DATE_VAL-365),YEAR(DATE_VAL),MONTH(DATE_VAL),DAY(DATE_VAL),"d","a",FALSE,FALSE,TRUE),1))*252^0.5
    Next l
'---------------------------------------------------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------------------------------------------------
LOAD2_LINE:
'---------------------------------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(i, 1) = HEADING_STR
    For l = 1 To NROWS
        TICKER_STR = TICKERS_VECTOR(l, 1)
        TEMP_MATRIX(i + l, 1) = TICKER_STR
        If Y_VAL > 0 Then
            j = 2
            For k = 0 To 9
                DATA_ARR(2) = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, Y_VAL + k, ERROR_STR)
                If DATA_ARR(2) <> ERROR_STR Then: TEMP_MATRIX(i + l, j) = DATA_ARR(2)
                j = j + 1
            Next k
        Else
            j = 12
        End If
        If Q_VAL > 0 Then
            For k = 0 To 19
                DATA_ARR(2) = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER_STR, Q_VAL + k, ERROR_STR)
                If DATA_ARR(2) <> ERROR_STR Then: TEMP_MATRIX(i + l, j) = DATA_ARR(2)
                j = j + 1
            Next k
        End If
    Next l
'---------------------------------------------------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------------------------------------------------
LOAD3_LINE:
'---------------------------------------------------------------------------------------------------------------------------------

    TEMP_MATRIX(i, 1) = HEADING_STR
    For k = 1 To NROWS
        TICKER_STR = TICKERS_VECTOR(k, 1): TEMP_MATRIX(i + k, 1) = TICKER_STR
        jj = k + r(2): ii = k + r(1)
        If SKIP_FLAG = False Then
            j = 2
            For l = 0 To 9
                If TEMP_MATRIX(jj, j) = "" Then: GoTo 1980
                If TEMP_MATRIX(ii, j + h) = "" Then: GoTo 1980
                
                DATA_ARR(2) = TEMP_MATRIX(jj, j)
                DATA_ARR(1) = TEMP_MATRIX(ii, j + h)
                DATA_ARR(3) = "": GoSub CALC_LINE
                If DATA_ARR(3) <> "" Then: TEMP_MATRIX(i + k, j) = DATA_ARR(3)
                j = j + 1
            Next l
        End If
1980:
        j = 12
        Select Case V2
        Case 0
            For l = 0 To 19 - h
                If TEMP_MATRIX(jj, j) = "" Then: GoTo 1981
                If TEMP_MATRIX(ii, j + h) = "" Then: GoTo 1981
                
                DATA_ARR(2) = TEMP_MATRIX(jj, j)
                DATA_ARR(1) = TEMP_MATRIX(ii, j + h)
                DATA_ARR(3) = "": GoSub CALC_LINE
                If DATA_ARR(3) <> "" Then: TEMP_MATRIX(i + k, j) = DATA_ARR(3)
                j = j + 1
            Next l
        Case 1
            For l = 0 To 19 - 4
                DATA_ARR(2) = 0
                For m = 1 To 4
                    If TEMP_MATRIX(jj, j + m - 1) = "" Then: GoTo 1981
                    DATA_ARR(2) = DATA_ARR(2) + TEMP_MATRIX(jj, j + m - 1)
                Next m
                If TEMP_MATRIX(ii, j) = "" Then: GoTo 1981
                DATA_ARR(1) = TEMP_MATRIX(ii, j)
                DATA_ARR(3) = "": GoSub CALC_LINE
                If DATA_ARR(3) <> "" Then: TEMP_MATRIX(i + k, j) = DATA_ARR(3)
                j = j + 1
            Next l
        Case 2
            For l = 0 To 19 - 4
                DATA_ARR(1) = 0
                For m = 1 To 4
                    If TEMP_MATRIX(ii, j + m - 1) = "" Then: GoTo 1981
                    DATA_ARR(1) = DATA_ARR(1) + TEMP_MATRIX(ii, j + m - 1)
                Next m
                If TEMP_MATRIX(jj, j) = "" Then: GoTo 1981
                DATA_ARR(2) = TEMP_MATRIX(jj, j)
                DATA_ARR(3) = "": GoSub CALC_LINE
                If DATA_ARR(3) <> "" Then: TEMP_MATRIX(i + k, j) = DATA_ARR(3)
                j = j + 1
            Next l
        Case Else
            For l = 0 To 19 - 4
                DATA_ARR(1) = 0: DATA_ARR(2) = 0
                For m = 1 To 4
                    If TEMP_MATRIX(ii, j + m - 1) = "" Then: GoTo 1981
                    If TEMP_MATRIX(jj, j + m - 1) = "" Then: GoTo 1981
                    
                    DATA_ARR(1) = DATA_ARR(1) + TEMP_MATRIX(ii, j + m - 1)
                    DATA_ARR(2) = DATA_ARR(2) + TEMP_MATRIX(jj, j + m - 1)
                Next m
                DATA_ARR(3) = "": GoSub CALC_LINE
                If DATA_ARR(3) <> "" Then: TEMP_MATRIX(i + k, j) = DATA_ARR(3)
                j = j + 1
            Next l
        End Select
1981:
    Next k
'---------------------------------------------------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------------------------------------------------
LOAD4_LINE:
'---------------------------------------------------------------------------------------------------------------------------------
    TEMP_MATRIX(i, 1) = HEADING_STR
    For k = 1 To NROWS
        TICKER_STR = TICKERS_VECTOR(k, 1)
        TEMP_MATRIX(i + k, 1) = TICKER_STR
        '----------------------------------------------------------------------------------------------------
        Select Case V1
        '----------------------------------------------------------------------------------------------------
        Case 0 'Enterprise Value
            j = 2
            For m = 1 To 2
                If m = 1 Then n = 9 Else n = 19
                For l = 0 To n
                    For o = 1 To 3
                        DATA_ARR(o) = 0
                        If TEMP_MATRIX(k + r(o), j) = "" Then: GoTo 1982
                        DATA_ARR(o) = TEMP_MATRIX(k + r(o), j)
                    Next o
                    TEMP_MATRIX(i + k, j) = DATA_ARR(1) - DATA_ARR(2) + DATA_ARR(3)
1982:
                    j = j + 1
                Next l
            Next m
            
        '----------------------------------------------------------------------------------------------------
        Case Else 'Altman Z-Score
        '----------------------------------------------------------------------------------------------------
            j = 2
            For m = 1 To 2
                If m = 1 Then n = 9 Else n = 19 - 4
                For l = 0 To n
                    For o = 1 To 7
                        DATA_ARR(o) = 0
                        If TEMP_MATRIX(k + r(o), j) = "" Then: GoTo 1983
                        DATA_ARR(o) = TEMP_MATRIX(k + r(o), j)
                    Next o
                    If n = 19 - 4 Then
                        For o = 1 To 3
                        
                            If (TEMP_MATRIX(k + r(4), j + o) = "") Or (TEMP_MATRIX(k + r(7), j + o) = "") Then: GoTo 1983
                            DATA_ARR(4) = DATA_ARR(4) + TEMP_MATRIX(k + r(4), j + o)
                            DATA_ARR(7) = DATA_ARR(7) + TEMP_MATRIX(k + r(7), j + o)
                        Next o
                    End If
                    If DATA_ARR(2) <> 0 And DATA_ARR(6) <> 0 Then
                        TEMP_MATRIX(i + k, j) = _
                            1.2 * (DATA_ARR(1) / DATA_ARR(2)) + _
                            1.4 * (DATA_ARR(3) / DATA_ARR(2)) + _
                            3.3 * (DATA_ARR(4) / DATA_ARR(2)) + _
                            0.6 * (DATA_ARR(5) / DATA_ARR(6)) + _
                            (DATA_ARR(7) / DATA_ARR(2))
                            'The Interpretation of Altman Z-Score:
                            'Z-SCORE ABOVE 3.0 –The company is considered ‘Safe’ based on the financial figures only.
                            'Z-SCORE BETWEEN 2.7 and 2.99 – ‘On Alert’. This zone is an area where one should ‘Exercise Caution’.
                            'Z-SCORE BETWEEN 1.8 and 2.7 – Good chance of the company going bankrupt within 2 years of operations from the date of financial figures given.
                            'Z-SCORE BELOW 1.80- Probability of Financial Catastrophe is Very High.
                            'If the Altman Z-Score is close to or below 3, then it would be as well to do some serious due diligence
                            'on the company in question before even considering investing.
                    End If
1983:
                    j = j + 1
                Next l
            Next m
        '----------------------------------------------------------------------------------------------------
        End Select
        '----------------------------------------------------------------------------------------------------
    Next k
'---------------------------------------------------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------------------------------------------------
CALC_LINE:
'---------------------------------------------------------------------------------------------------------------------------------
    Select Case V1
    Case 1
        If (DATA_ARR(1) <> 0) Then: DATA_ARR(3) = 1 / DATA_ARR(1)
    Case 2
        If (DATA_ARR(2) <> 0) Then: DATA_ARR(3) = 1 / DATA_ARR(2)
    Case 3
        If (DATA_ARR(2) <> 0) Then: DATA_ARR(3) = DATA_ARR(1) / DATA_ARR(2)
    Case 4
        If (DATA_ARR(1) <> 0) Then: DATA_ARR(3) = DATA_ARR(2) / DATA_ARR(1)
    Case 5
        If (DATA_ARR(2) <> 0) Then: DATA_ARR(3) = DATA_ARR(1) / DATA_ARR(2) - 1
    Case 6
        If (DATA_ARR(1) <> 0) Then: DATA_ARR(3) = DATA_ARR(2) / DATA_ARR(1) - 1
    Case 7
        If (DATA_ARR(2) <> 0) Then: DATA_ARR(3) = Abs(DATA_ARR(1)) / DATA_ARR(2)
    Case 8
        If (DATA_ARR(1) <> 0) Then: DATA_ARR(3) = Abs(DATA_ARR(2)) / DATA_ARR(1)
    Case 9
        DATA_ARR(3) = DATA_ARR(2) * DATA_ARR(1)
    Case 10
        DATA_ARR(3) = DATA_ARR(2) - DATA_ARR(1)
    Case 11
        DATA_ARR(3) = DATA_ARR(2) + DATA_ARR(1)
    End Select
'---------------------------------------------------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------------------------------------------------
INDEX_LINE:
'---------------------------------------------------------------------------------------------------------------------------------
    Call INDEX_OBJ.Add(CStr(i), HEADING_STR)
'---------------------------------------------------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
HISTORICAL_ALTMAN_Z_SCORE_FUNC = Err.number
End Function


