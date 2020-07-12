Attribute VB_Name = "FINAN_FUNDAM_NNWC_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Sub PRINT_HISTORICAL_NNWC_FUNC()
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

FINDEX_ARR = Array(4, 4, 6, 2, 2, _
                2, 2, 2, 2, 2, _
                2, 2, 2, 2, 3)

TICKERS_VECTOR = Excel.Application.InputBox("Symbol(s)", "Net Net Working Capital (NNWC)", , , , , , 8)
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

TEMP_MATRIX = HISTORICAL_NNWC_FUNC(TICKERS_VECTOR)
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

Private Function HISTORICAL_NNWC_FUNC(ByRef TICKERS_RNG As Variant)

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

Dim DATE_VAL As Variant
Dim DATA_ARR(1 To 7) As Variant
Dim INDEX_OBJ As New Collection
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

Const BV_RECEIVABLES_MULTIPLIER_VAL = 0.75
Const BV_INVENTORIES_MULTIPLIER_VAL = 0.5

On Error GoTo ERROR_LABEL

ERROR_STR = "-"
TICKERS_VECTOR = TICKERS_RNG
If UBound(TICKERS_VECTOR, 1) = 1 Then
    TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
End If
NROWS = UBound(TICKERS_VECTOR, 1)
NCOLUMNS = 10 + 20 + 1

V2 = 0
NSIZE = (NROWS + 1) * 15
ReDim TEMP_MATRIX(1 To NSIZE, 1 To NCOLUMNS)
For i = 1 To NSIZE: For j = 1 To NCOLUMNS: TEMP_MATRIX(i, j) = "": Next j: Next i

i = 1: HEADING_STR = "End Period": Y_VAL = 5196: Q_VAL = 8006: GoSub LOAD2_LINE: GoSub INDEX_LINE

i = i + NROWS + 1: HEADING_STR = "Data Loaded": Y_VAL = 5206: Q_VAL = 8026: GoSub LOAD2_LINE: GoSub INDEX_LINE
i = i + NROWS + 1: HEADING_STR = "Ending Quarter": Y_VAL = 0: Q_VAL = 8066: GoSub LOAD2_LINE: GoSub INDEX_LINE

' Cash & Equivalents
i = i + NROWS + 1: HEADING_STR = "BV Cash & Equivalents": Y_VAL = 5946: Q_VAL = 9506: GoSub LOAD2_LINE: GoSub INDEX_LINE
' Receivables
i = i + NROWS + 1: HEADING_STR = "BV Receivables": Y_VAL = 6006: Q_VAL = 9626: GoSub LOAD2_LINE: GoSub INDEX_LINE
V2 = 0: V1 = 1: r(1) = CLng(INDEX_OBJ("BV Receivables")): i = i + NROWS + 1: HEADING_STR = "NNV Receivables": GoSub LOAD3_LINE: GoSub INDEX_LINE

' Inventories --Total
i = i + NROWS + 1: HEADING_STR = "BV Inventory": Y_VAL = 6076: Q_VAL = 9766: GoSub LOAD2_LINE: GoSub INDEX_LINE
V2 = 0: V1 = 2: r(1) = CLng(INDEX_OBJ("BV Inventory")): i = i + NROWS + 1: HEADING_STR = "NNV Inventory": GoSub LOAD3_LINE: GoSub INDEX_LINE

' Total Liabilities
i = i + NROWS + 1: HEADING_STR = "BV Total Liabilities": Y_VAL = 6456: Q_VAL = 10526: GoSub LOAD2_LINE: GoSub INDEX_LINE

' Shares Outstanding
i = i + NROWS + 1: HEADING_STR = "Weight Common Shares": Y_VAL = 6666: Q_VAL = 10946: GoSub LOAD2_LINE: GoSub INDEX_LINE


V2 = 2: r(1) = CLng(INDEX_OBJ("BV Cash & Equivalents")): r(2) = CLng(INDEX_OBJ("NNV Receivables"))
r(3) = CLng(INDEX_OBJ("NNV Inventory")): r(4) = CLng(INDEX_OBJ("BV Total Liabilities"))
i = i + NROWS + 1: HEADING_STR = "Net Net Working Capital (NNWC)": GoSub LOAD3_LINE: GoSub INDEX_LINE

V2 = 1: V1 = 3: r(1) = CLng(INDEX_OBJ("Net Net Working Capital (NNWC)")): r(2) = CLng(INDEX_OBJ("Weight Common Shares"))
i = i + NROWS + 1: HEADING_STR = "NNWC per share": GoSub LOAD3_LINE: GoSub INDEX_LINE

p = i: i = i + NROWS + 1
r(1) = CLng(INDEX_OBJ("Data Loaded")): i = i + NROWS + 1: HEADING_STR = "Closing Stock Price": GoSub LOAD1_LINE: GoSub INDEX_LINE
i = p: V2 = 1: V1 = 4: r(2) = CLng(INDEX_OBJ("Weight Common Shares")): r(1) = CLng(INDEX_OBJ("Closing Stock Price")): i = i + NROWS + 1: HEADING_STR = "Market Cap": GoSub LOAD3_LINE: GoSub INDEX_LINE 'Weight Common Shares x Stock Price Data Loaded

i = i + NROWS + 1
V2 = 1: V1 = 3: r(1) = CLng(INDEX_OBJ("NNWC per share")): r(2) = CLng(INDEX_OBJ("Closing Stock Price"))
i = i + NROWS + 1: HEADING_STR = "Net Net Working Capital / Price": GoSub LOAD3_LINE: GoSub INDEX_LINE

HISTORICAL_NNWC_FUNC = TEMP_MATRIX

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
'---------------------------------------------------------------------------------------------------------------------------------
    Select Case V2
'---------------------------------------------------------------------------------------------------------------------------------
    Case 0
'---------------------------------------------------------------------------------------------------------------------------------
        For k = 1 To NROWS
            TICKER_STR = TICKERS_VECTOR(k, 1): TEMP_MATRIX(i + k, 1) = TICKER_STR
            ii = k + r(1)
            j = 2
            For m = 1 To 2
                If m = 1 Then n = 9 Else n = 19
                For l = 0 To n
                    If TEMP_MATRIX(ii, j) = "" Then: GoTo 1981
                    DATA_ARR(1) = TEMP_MATRIX(ii, j): DATA_ARR(2) = "": GoSub CALC_LINE
                    If DATA_ARR(2) <> "" Then: TEMP_MATRIX(i + k, j) = DATA_ARR(2)
1981:
                    j = j + 1
                Next l
            Next m
                
        Next k
'---------------------------------------------------------------------------------------------------------------------------------
    Case 1
'---------------------------------------------------------------------------------------------------------------------------------
        For k = 1 To NROWS
            TICKER_STR = TICKERS_VECTOR(k, 1): TEMP_MATRIX(i + k, 1) = TICKER_STR
            ii = k + r(1): jj = k + r(2)
            j = 2
            For m = 1 To 2
                If m = 1 Then n = 9 Else n = 19
                For l = 0 To n
                    If TEMP_MATRIX(ii, j) = "" Then: GoTo 1982
                    If TEMP_MATRIX(jj, j) = "" Then: GoTo 1982
                    DATA_ARR(1) = TEMP_MATRIX(ii, j)
                    DATA_ARR(2) = TEMP_MATRIX(jj, j)
                    DATA_ARR(3) = "": GoSub CALC_LINE
                    If DATA_ARR(3) <> "" Then: TEMP_MATRIX(i + k, j) = DATA_ARR(3)
1982:
                    j = j + 1
                Next l
            Next m
        Next k
'---------------------------------------------------------------------------------------------------------------------------------
    Case Else
'---------------------------------------------------------------------------------------------------------------------------------
        For k = 1 To NROWS
            TICKER_STR = TICKERS_VECTOR(k, 1): TEMP_MATRIX(i + k, 1) = TICKER_STR
            j = 2
            For m = 1 To 2
                If m = 1 Then n = 9 Else n = 19
                For l = 0 To n
                    For o = 1 To 4
                        DATA_ARR(o) = 0
                        If TEMP_MATRIX(k + r(o), j) = "" Then: GoTo 1983
                        DATA_ARR(o) = TEMP_MATRIX(k + r(o), j)
                    Next o
                    TEMP_MATRIX(i + k, j) = DATA_ARR(1) + DATA_ARR(2) + DATA_ARR(3) - DATA_ARR(4)
1983:
                    j = j + 1
                Next l
            Next m
        Next k
'---------------------------------------------------------------------------------------------------------------------------------
    End Select
'---------------------------------------------------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------------------------------------------------
CALC_LINE:
'---------------------------------------------------------------------------------------------------------------------------------
    Select Case V1
    Case 1
        DATA_ARR(2) = DATA_ARR(1) * BV_RECEIVABLES_MULTIPLIER_VAL
    Case 2
        DATA_ARR(2) = DATA_ARR(1) * BV_INVENTORIES_MULTIPLIER_VAL
    Case 3
        If DATA_ARR(2) <> 0 Then: DATA_ARR(3) = DATA_ARR(1) / DATA_ARR(2)
    Case 4
        DATA_ARR(3) = DATA_ARR(1) * DATA_ARR(2)
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
HISTORICAL_NNWC_FUNC = Err.number
End Function
