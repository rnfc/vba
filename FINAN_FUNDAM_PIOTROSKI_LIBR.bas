Attribute VB_Name = "FINAN_FUNDAM_PIOTROSKI_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Sub PRINT_HISTORICAL_PIOTROSKI_FSCORE_FUNC()
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

FINDEX_ARR = Array(4, 4, 1, 1, 1, _
                    1, 1, 1, 1, 1, _
                    1, 6, 6, 6, 6, _
                    6, 6, 6, 6, 6, _
                    6, 6, 2)

TICKERS_VECTOR = Excel.Application.InputBox("Symbol(s)", "Piotroski F-Score", , , , , , 8)
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

TEMP_MATRIX = HISTORICAL_PIOTROSKI_FSCORE_FUNC(TICKERS_VECTOR)
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

'Piotroski (www.chicagobooth.edu/faculty/selectedpapers/sp84.pdf)

Function HISTORICAL_PIOTROSKI_FSCORE_FUNC(ByRef TICKERS_RNG As Variant)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim m As Long
Dim o As Long
Dim p As Long 'Position Row

Dim ii As Long
Dim jj As Long

Dim r(1 To 7) As Long 'Row #

Dim Y_VAL As Long
Dim Q_VAL As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ERROR_STR As String
Dim TICKER_STR As String
Dim HEADING_STR As String
Dim HEADINGS_STR As String

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

NSIZE = (NROWS + 1) * 23
ReDim TEMP_MATRIX(1 To NSIZE, 1 To NCOLUMNS)
For i = 1 To NSIZE: For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = "": Next j: Next i

i = 1: HEADING_STR = "End Period": Y_VAL = 5196: Q_VAL = 8006: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Data Loaded": Y_VAL = 5206: Q_VAL = 8026: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Operating Revenue": Y_VAL = 5286: Q_VAL = 8186: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Gross Operating Profit": Y_VAL = 5346: Q_VAL = 8306: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Net Income (Continuing Operations)": Y_VAL = 5596: Q_VAL = 8806: GoSub LOAD2_LINE: GoSub INDEX_LINE

i = i + NROWS + 1: HEADING_STR = "Current Assets": Y_VAL = 6116: Q_VAL = 9846: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Total Assets": Y_VAL = 6266: Q_VAL = 10146: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Current Liabilities": Y_VAL = 6366: Q_VAL = 10346: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Long Term Debt": Y_VAL = 6376: Q_VAL = 10366: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Total Common Shares Outstanding": Y_VAL = 6646: Q_VAL = 10906: GoSub LOAD2_LINE: GoSub INDEX_LINE

i = i + NROWS + 1: HEADING_STR = "Net Cash From Continuing Operations": Y_VAL = 6856: Q_VAL = 11326: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Ending Quarter": Y_VAL = 0: Q_VAL = 8066: GoSub LOAD2_LINE: GoSub INDEX_LINE

m = 10
HEADINGS_STR = "Positive Net Income,Positive Operating Cash Flow,Increasing Net Income,Operating Cash flow exceeds Net Income," & _
               "Decreasing ratio of long-term debt to assets,Increasing Current Ratio,No increase in outstanding shares," & _
               "Increasing Gross Margins,Increasing Asset Turnover,Overall Score,"

h = i: j = 1: l = i: p = (NROWS + 1)
For o = 1 To m
    k = InStr(j, HEADINGS_STR, ","): HEADING_STR = Mid(HEADINGS_STR, j, k - j): i = l + p * o
    TEMP_MATRIX(i, 1) = HEADING_STR: GoSub INDEX_LINE
    j = k + 1
Next o
i = h + m * (NROWS + 1): r(1) = CLng(INDEX_OBJ("Data Loaded")): i = i + NROWS + 1: HEADING_STR = "Closing Stock Price": GoSub LOAD1_LINE: GoSub INDEX_LINE
i = h
'-------------------------------------------------------------------------------------------------------------------------------------
For l = 1 To NROWS
'-------------------------------------------------------------------------------------------------------------------------------------
    TICKER_STR = TICKERS_VECTOR(l, 1)
    For k = 1 To m
        jj = k * p + i + l
        TEMP_MATRIX(jj, 1) = TICKER_STR
    Next k
    
    ii = CLng(INDEX_OBJ("Overall Score")) + l
    j = 2
    '----------------------------------------------------------------------------------------------------------------------------
    For k = 1 To 9 'Perfect
    '----------------------------------------------------------------------------------------------------------------------------
        TEMP_MATRIX(ii, j) = 0
        jj = CLng(INDEX_OBJ("Positive Net Income")) + l
        DATA_ARR(1) = "": DATA_ARR(3) = ""
        r(1) = CLng(INDEX_OBJ("Net Income (Continuing Operations)")) + l
        If (TEMP_MATRIX(r(1), j + 0) <> "") Then 'Positive Net Income
            DATA_ARR(1) = TEMP_MATRIX(r(1), j + 0)
            TEMP_MATRIX(jj, j) = IIf(DATA_ARR(1) > 0, 1, 0)
            TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            If (TEMP_MATRIX(r(1), j + 1) <> "") Then 'Increasing Net Income
                jj = CLng(INDEX_OBJ("Increasing Net Income")) + l
                DATA_ARR(3) = TEMP_MATRIX(r(1), j + 1)
                TEMP_MATRIX(jj, j) = IIf(DATA_ARR(1) > DATA_ARR(3), 1, 0)
                TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            End If
        End If
        
        jj = CLng(INDEX_OBJ("Positive Operating Cash Flow")) + l
        DATA_ARR(2) = ""
        r(1) = CLng(INDEX_OBJ("Net Cash From Continuing Operations")) + l
        If (TEMP_MATRIX(r(1), j + 0) <> "") Then 'Positive Operating Cash Flow
            DATA_ARR(2) = TEMP_MATRIX(r(1), j + 0)
            TEMP_MATRIX(jj, j) = IIf(DATA_ARR(2) > 0, 1, 0)
            TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            If (DATA_ARR(1) <> "") Then 'Operating Cash flow exceeds Net Income
                jj = CLng(INDEX_OBJ("Operating Cash flow exceeds Net Income")) + l
                TEMP_MATRIX(jj, j) = IIf(DATA_ARR(2) > DATA_ARR(1), 1, 0)
                TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            End If
        End If
        
        jj = CLng(INDEX_OBJ("Decreasing ratio of long-term debt to assets")) + l
        DATA_ARR(1) = "": DATA_ARR(2) = ""
        DATA_ARR(3) = "": DATA_ARR(4) = ""
        r(1) = CLng(INDEX_OBJ("Long Term Debt")) + l
        r(2) = CLng(INDEX_OBJ("Total Assets")) + l
        If ((TEMP_MATRIX(r(1), j + 0) <> "") And (TEMP_MATRIX(r(1), j + 1) <> "") And _
            (TEMP_MATRIX(r(2), j + 0) <> "") And (TEMP_MATRIX(r(2), j + 1) <> "")) Then 'Decreasing ratio of long-term debt to assets
            DATA_ARR(1) = TEMP_MATRIX(r(1), j + 0): DATA_ARR(2) = TEMP_MATRIX(r(1), j + 1)
            DATA_ARR(3) = TEMP_MATRIX(r(2), j + 0): DATA_ARR(4) = TEMP_MATRIX(r(2), j + 1)
            If ((DATA_ARR(3) <> 0) And (DATA_ARR(4) <> 0)) Then
                TEMP_MATRIX(jj, j) = IIf((DATA_ARR(1) / DATA_ARR(3)) < (DATA_ARR(2) / DATA_ARR(4)), 1, 0)
                If DATA_ARR(1) = 0 Then: TEMP_MATRIX(jj, j) = 1
                TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            End If
        End If
                    
        jj = CLng(INDEX_OBJ("Increasing Current Ratio")) + l
        DATA_ARR(1) = "": DATA_ARR(2) = ""
        DATA_ARR(3) = "": DATA_ARR(4) = ""
        r(1) = CLng(INDEX_OBJ("Current Assets")) + l
        r(2) = CLng(INDEX_OBJ("Current Liabilities")) + l
        If ((TEMP_MATRIX(r(1), j + 0) <> "") And (TEMP_MATRIX(r(1), j + 1) <> "") And _
            (TEMP_MATRIX(r(2), j + 0) <> "") And (TEMP_MATRIX(r(2), j + 1) <> "")) Then 'Increasing Current Ratio
            DATA_ARR(1) = TEMP_MATRIX(r(1), j + 0): DATA_ARR(2) = TEMP_MATRIX(r(1), j + 1)
            DATA_ARR(3) = TEMP_MATRIX(r(2), j + 0): DATA_ARR(4) = TEMP_MATRIX(r(2), j + 1)
            If ((DATA_ARR(3) <> 0) And (DATA_ARR(4) <> 0)) Then
                TEMP_MATRIX(jj, j) = IIf((DATA_ARR(1) / DATA_ARR(3)) > (DATA_ARR(2) / DATA_ARR(4)), 1, 0)
                TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            End If
        End If
        
        jj = CLng(INDEX_OBJ("No increase in outstanding shares")) + l
        DATA_ARR(1) = "": DATA_ARR(2) = ""
        r(1) = CLng(INDEX_OBJ("Total Common Shares Outstanding")) + l
        If ((TEMP_MATRIX(r(1), j + 0) <> "") And (TEMP_MATRIX(r(1), j + 1) <> "")) Then 'No increase in outstanding shares
            DATA_ARR(1) = TEMP_MATRIX(r(1), j + 0): DATA_ARR(2) = TEMP_MATRIX(r(1), j + 1)
            TEMP_MATRIX(jj, j) = IIf(DATA_ARR(1) > DATA_ARR(2), 0, 1)
            TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
        End If
        
        jj = CLng(INDEX_OBJ("Increasing Gross Margins")) + l
        DATA_ARR(1) = "": DATA_ARR(2) = ""
        DATA_ARR(3) = "": DATA_ARR(4) = ""
        r(1) = CLng(INDEX_OBJ("Gross Operating Profit")) + l
        r(2) = CLng(INDEX_OBJ("Operating Revenue")) + l
        If ((TEMP_MATRIX(r(1), j + 0) <> "") And (TEMP_MATRIX(r(1), j + 1) <> "") And _
            (TEMP_MATRIX(r(2), j + 0) <> "") And (TEMP_MATRIX(r(2), j + 1) <> "")) Then 'Increasing Gross Margins
            DATA_ARR(1) = TEMP_MATRIX(r(1), j + 0): DATA_ARR(2) = TEMP_MATRIX(r(1), j + 1)
            DATA_ARR(3) = TEMP_MATRIX(r(2), j + 0): DATA_ARR(4) = TEMP_MATRIX(r(2), j + 1)
            If ((DATA_ARR(3) <> 0) And (DATA_ARR(4) <> 0)) Then
                TEMP_MATRIX(jj, j) = IIf((DATA_ARR(1) / DATA_ARR(3)) > (DATA_ARR(2) / DATA_ARR(4)), 1, 0)
                TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            End If
        End If
        
        jj = CLng(INDEX_OBJ("Increasing Asset Turnover")) + l
        DATA_ARR(1) = "": DATA_ARR(2) = ""
        DATA_ARR(3) = "": DATA_ARR(4) = ""
        r(1) = CLng(INDEX_OBJ("Operating Revenue")) + l
        r(2) = CLng(INDEX_OBJ("Total Assets")) + l
        If ((TEMP_MATRIX(r(1), j + 0) <> "") And (TEMP_MATRIX(r(1), j + 1) <> "") And _
            (TEMP_MATRIX(r(2), j + 0) <> "") And (TEMP_MATRIX(r(2), j + 1) <> "")) Then 'Increasing Asset Turnover
            DATA_ARR(1) = TEMP_MATRIX(r(1), j + 0): DATA_ARR(2) = TEMP_MATRIX(r(1), j + 1)
            DATA_ARR(3) = TEMP_MATRIX(r(2), j + 0): DATA_ARR(4) = TEMP_MATRIX(r(2), j + 1)
            If ((DATA_ARR(3) <> 0) And (DATA_ARR(4) <> 0)) Then
                TEMP_MATRIX(jj, j) = IIf((DATA_ARR(1) / DATA_ARR(3)) > (DATA_ARR(2) / DATA_ARR(4)), 1, 0)
                TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            End If
        End If
        j = j + 1
    '----------------------------------------------------------------------------------------------------------------------------
    Next k
    '----------------------------------------------------------------------------------------------------------------------------
    
    j = j + 1
    h = 2
    '----------------------------------------------------------------------------------------------------------------------------
    For k = 0 To 19 - 4
    '----------------------------------------------------------------------------------------------------------------------------
        TEMP_MATRIX(ii, j) = 0
        r(1) = CLng(INDEX_OBJ("Ending Quarter")) + l
        If (TEMP_MATRIX(r(1), j + 0) = 4) Then: h = h + 1
        
        jj = CLng(INDEX_OBJ("Positive Net Income")) + l
        DATA_ARR(1) = "": DATA_ARR(3) = ""
        r(1) = CLng(INDEX_OBJ("Net Income (Continuing Operations)")) + l
        If ((TEMP_MATRIX(r(1), j + 0) <> "") And (TEMP_MATRIX(r(1), j + 1) <> "") And (TEMP_MATRIX(r(1), j + 2) <> "") And (TEMP_MATRIX(r(1), j + 3) <> "")) Then 'Positive Net Income
            DATA_ARR(1) = TEMP_MATRIX(r(1), j + 0) + TEMP_MATRIX(r(1), j + 1) + TEMP_MATRIX(r(1), j + 2) + TEMP_MATRIX(r(1), j + 3)
            TEMP_MATRIX(jj, j) = IIf(DATA_ARR(1) > 0, 1, 0)
            TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            If (TEMP_MATRIX(r(1), h) <> "") Then 'Increasing Net Income
                jj = CLng(INDEX_OBJ("Increasing Net Income")) + l
                DATA_ARR(3) = TEMP_MATRIX(r(1), h)
                TEMP_MATRIX(jj, j) = IIf(DATA_ARR(1) > DATA_ARR(3), 1, 0)
                TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            End If
        End If
        
        jj = CLng(INDEX_OBJ("Positive Operating Cash Flow")) + l
        DATA_ARR(2) = ""
        r(1) = CLng(INDEX_OBJ("Net Cash From Continuing Operations")) + l
        If ((TEMP_MATRIX(r(1), j + 0) <> "") And (TEMP_MATRIX(r(1), j + 4) <> "") And (TEMP_MATRIX(r(1), h) <> "")) Then  'Positive Operating Cash Flow
            DATA_ARR(2) = TEMP_MATRIX(r(1), j + 0) - TEMP_MATRIX(r(1), j + 4) + TEMP_MATRIX(r(1), h)
            TEMP_MATRIX(jj, j) = IIf(DATA_ARR(2) > 0, 1, 0)
            TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            If (DATA_ARR(1) <> "") Then 'Operating Cash flow exceeds Net Income
                jj = CLng(INDEX_OBJ("Operating Cash flow exceeds Net Income")) + l
                TEMP_MATRIX(jj, j) = IIf(DATA_ARR(2) > DATA_ARR(1), 1, 0)
                TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            End If
        End If
        
        jj = CLng(INDEX_OBJ("Decreasing ratio of long-term debt to assets")) + l
        DATA_ARR(1) = "": DATA_ARR(2) = ""
        DATA_ARR(3) = "": DATA_ARR(4) = ""
        r(1) = CLng(INDEX_OBJ("Long Term Debt")) + l
        r(2) = CLng(INDEX_OBJ("Total Assets")) + l
        If ((TEMP_MATRIX(r(1), j + 0) <> "") And (TEMP_MATRIX(r(1), h) <> "") And _
            (TEMP_MATRIX(r(2), j + 0) <> "") And (TEMP_MATRIX(r(2), h) <> "")) Then 'Decreasing ratio of long-term debt to assets
            DATA_ARR(1) = TEMP_MATRIX(r(1), j + 0): DATA_ARR(2) = TEMP_MATRIX(r(1), h)
            DATA_ARR(3) = TEMP_MATRIX(r(2), j + 0): DATA_ARR(4) = TEMP_MATRIX(r(2), h)
            If ((DATA_ARR(3) <> 0) And (DATA_ARR(4) <> 0)) Then
                TEMP_MATRIX(jj, j) = IIf((DATA_ARR(1) / DATA_ARR(3)) < (DATA_ARR(2) / DATA_ARR(4)), 1, 0)
                If DATA_ARR(1) = 0 Then: TEMP_MATRIX(jj, j) = 1
                TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            End If
        End If
        
        jj = CLng(INDEX_OBJ("Increasing Current Ratio")) + l
        DATA_ARR(1) = "": DATA_ARR(2) = ""
        DATA_ARR(3) = "": DATA_ARR(4) = ""
        r(1) = CLng(INDEX_OBJ("Current Assets")) + l
        r(2) = CLng(INDEX_OBJ("Current Liabilities")) + l
        If ((TEMP_MATRIX(r(1), j + 0) <> "") And (TEMP_MATRIX(r(1), h) <> "") And _
            (TEMP_MATRIX(r(2), j + 0) <> "") And (TEMP_MATRIX(r(2), h) <> "")) Then 'Increasing Current Ratio
            DATA_ARR(1) = TEMP_MATRIX(r(1), j + 0): DATA_ARR(2) = TEMP_MATRIX(r(1), h)
            DATA_ARR(3) = TEMP_MATRIX(r(2), j + 0): DATA_ARR(4) = TEMP_MATRIX(r(2), h)
            If ((DATA_ARR(3) <> 0) And (DATA_ARR(4) <> 0)) Then
                TEMP_MATRIX(jj, j) = IIf((DATA_ARR(1) / DATA_ARR(3)) > (DATA_ARR(2) / DATA_ARR(4)), 1, 0)
                TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            End If
        End If
        
        jj = CLng(INDEX_OBJ("No increase in outstanding shares")) + l
        DATA_ARR(1) = "": DATA_ARR(2) = ""
        r(1) = CLng(INDEX_OBJ("Total Common Shares Outstanding")) + l
        If ((TEMP_MATRIX(r(1), j + 0) <> "") And (TEMP_MATRIX(r(1), h) <> "")) Then 'No increase in outstanding shares
            DATA_ARR(1) = TEMP_MATRIX(r(1), j + 0): DATA_ARR(2) = TEMP_MATRIX(r(1), h)
            TEMP_MATRIX(jj, j) = IIf(DATA_ARR(1) > DATA_ARR(2), 0, 1)
            TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
        End If
        
        jj = CLng(INDEX_OBJ("Increasing Gross Margins")) + l
        DATA_ARR(1) = "": DATA_ARR(2) = ""
        DATA_ARR(3) = "": DATA_ARR(4) = ""
        r(1) = CLng(INDEX_OBJ("Gross Operating Profit")) + l
        r(2) = CLng(INDEX_OBJ("Operating Revenue")) + l
        If ((TEMP_MATRIX(r(1), j + 0) <> "") And (TEMP_MATRIX(r(1), j + 1) <> "") And _
            (TEMP_MATRIX(r(1), j + 2) <> "") And (TEMP_MATRIX(r(1), j + 3) <> "") And (TEMP_MATRIX(r(1), h) <> "") And _
            (TEMP_MATRIX(r(2), j + 0) <> "") And (TEMP_MATRIX(r(2), j + 1) <> "") And _
            (TEMP_MATRIX(r(2), j + 2) <> "") And (TEMP_MATRIX(r(2), j + 3) <> "") And (TEMP_MATRIX(r(2), h) <> "")) Then 'Increasing Gross Margins
            DATA_ARR(1) = TEMP_MATRIX(r(1), j + 0) + TEMP_MATRIX(r(1), j + 1) + TEMP_MATRIX(r(1), j + 2) + TEMP_MATRIX(r(1), j + 3)
            DATA_ARR(2) = TEMP_MATRIX(r(1), h)
            DATA_ARR(3) = TEMP_MATRIX(r(2), j + 0) + TEMP_MATRIX(r(2), j + 1) + TEMP_MATRIX(r(2), j + 2) + TEMP_MATRIX(r(2), j + 3)
            DATA_ARR(4) = TEMP_MATRIX(r(2), h)
            If ((DATA_ARR(3) <> 0) And (DATA_ARR(4) <> 0)) Then
                TEMP_MATRIX(jj, j) = IIf((DATA_ARR(1) / DATA_ARR(3)) > (DATA_ARR(2) / DATA_ARR(4)), 1, 0)
                TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            End If
        End If
        
        jj = CLng(INDEX_OBJ("Increasing Asset Turnover")) + l
        DATA_ARR(1) = "": DATA_ARR(2) = ""
        DATA_ARR(3) = "": DATA_ARR(4) = ""
        r(1) = CLng(INDEX_OBJ("Operating Revenue")) + l
        r(2) = CLng(INDEX_OBJ("Total Assets")) + l
        If ((TEMP_MATRIX(r(1), j + 0) <> "") And (TEMP_MATRIX(r(1), j + 1) <> "") And _
            (TEMP_MATRIX(r(1), j + 2) <> "") And (TEMP_MATRIX(r(1), j + 3) <> "") And (TEMP_MATRIX(r(1), h) <> "") And _
            (TEMP_MATRIX(r(2), j + 0) <> "") And (TEMP_MATRIX(r(2), h) <> "")) Then 'Increasing Asset Turnover
            DATA_ARR(1) = TEMP_MATRIX(r(1), j + 0) + TEMP_MATRIX(r(1), j + 1) + TEMP_MATRIX(r(1), j + 2) + TEMP_MATRIX(r(1), j + 3)
            DATA_ARR(2) = TEMP_MATRIX(r(1), h)
            DATA_ARR(3) = TEMP_MATRIX(r(2), j + 0)
            DATA_ARR(4) = TEMP_MATRIX(r(2), h)
            If ((DATA_ARR(3) <> 0) And (DATA_ARR(4) <> 0)) Then
                TEMP_MATRIX(jj, j) = IIf((DATA_ARR(1) / DATA_ARR(3)) > (DATA_ARR(2) / DATA_ARR(4)), 1, 0)
                TEMP_MATRIX(ii, j) = TEMP_MATRIX(ii, j) + TEMP_MATRIX(jj, j)
            End If
        End If
        j = j + 1
    '----------------------------------------------------------------------------------------------------------------------------
    Next k
    '----------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------------------------
Next l
'-------------------------------------------------------------------------------------------------------------------------------------


HISTORICAL_PIOTROSKI_FSCORE_FUNC = TEMP_MATRIX

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
INDEX_LINE:
'---------------------------------------------------------------------------------------------------------------------------------
    Call INDEX_OBJ.Add(CStr(i), HEADING_STR)
'---------------------------------------------------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
HISTORICAL_PIOTROSKI_FSCORE_FUNC = Err.number
End Function

