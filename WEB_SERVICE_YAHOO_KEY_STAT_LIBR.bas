Attribute VB_Name = "WEB_SERVICE_YAHOO_KEY_STAT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Written by Nico for the EUM Class 2009

Function YAHOO_KEY_STATISTICS_FUNC(ByVal TICKERS_RNG As Variant)

'1000 x PEG x VOLAT^2"
' 100 x EV/EBITDA x VOLAT^2 / Earnings Growth"

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_STR As String
Dim TEMP2_STR As String
Dim LINE_STR As String
Dim DELIM_STR As String
Dim DATA_STR As String
Dim TICKER_STR As String
Dim SRC_URL_STR As String

Dim TEMP_GROUP() As String
Dim DATA_GROUP As Variant
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
NCOLUMNS = UBound(TICKERS_VECTOR, 1)
NROWS = 58 'As of March 29, 2011
DELIM_STR = ","
TEMP1_STR = "yfnc_tablehead1"
TEMP2_STR = "yfnc_tabledata1"

ReDim DATA_GROUP(1 To NCOLUMNS)

For h = 1 To NCOLUMNS
    TICKER_STR = TICKERS_VECTOR(h, 1)
    SRC_URL_STR = "http://finance.yahoo.com/q/ks?s=" & TICKER_STR
    DATA_STR = SAVE_WEB_DATA_PAGE_FUNC(SRC_URL_STR, 0, False, 0, False)
    If DATA_STR = PUB_WEB_DATA_SYSTEM_ERROR_STR Then: GoTo 1983
    
    DATA_STR = Replace(DATA_STR, Chr(10), "")
    DATA_STR = Replace(DATA_STR, ":", "")
    DATA_STR = Replace(DATA_STR, DELIM_STR, "")
    DATA_STR = Replace(DATA_STR, "&amp;", "&")

        
    ReDim TEMP_GROUP(1 To 1)
    i = 0
    For j = 1 To 1
        i = InStr(i + 1, DATA_STR, TEMP1_STR)
        If i = 0 Then: GoTo 1983
    Next j
    k = 0
    Do Until i = 0
        i = i + Len(TEMP1_STR): i = i + 1
        i = InStr(i, DATA_STR, ">"): i = i + 1
        
        j = InStr(i, DATA_STR, "<")
        If j = 0 Then: Exit Do
        
        LINE_STR = Mid(DATA_STR, i, j - i) & DELIM_STR
        i = InStr(j, DATA_STR, TEMP2_STR)
        i = i + Len(TEMP2_STR)
        i = i + 1
        If Mid(DATA_STR, i, 6) = "><span" Then: i = i + 6
        i = InStr(i, DATA_STR, ">") + 1
        
        j = InStr(i, DATA_STR, "<")
        If j = 0 Then: Exit Do
        
        LINE_STR = LINE_STR & Mid(DATA_STR, i, j - i)
        k = k + 1
        ReDim Preserve TEMP_GROUP(1 To k)
        TEMP_GROUP(k) = LINE_STR
        If k >= NROWS Then: Exit Do
        i = InStr(j, DATA_STR, TEMP1_STR)
    Loop

    DATA_GROUP(h) = TEMP_GROUP
1983:
Next h

'--------------------------------------------------------------------------------------
If NCOLUMNS = 1 Then
'--------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To UBound(TEMP_GROUP), 1 To 2)
    For i = 1 To UBound(TEMP_GROUP)
        TEMP1_STR = TEMP_GROUP(i)
        If TEMP1_STR = "" Then: GoTo 1984
        j = InStr(1, TEMP1_STR, DELIM_STR)
        TEMP2_STR = Trim(Mid(TEMP1_STR, 1, j - 1))
        TEMP_MATRIX(i, 1) = TEMP2_STR
        
        TEMP2_STR = Trim(Mid(TEMP1_STR, j + 1, Len(TEMP1_STR) - j))
        If TEMP2_STR = "N/A" Or TEMP2_STR = "NaN%" Or TEMP2_STR = "NaN" Then
            TEMP_MATRIX(i, 2) = 0
        Else
            TEMP_MATRIX(i, 2) = TEMP2_STR
        End If
1984:
    Next i
'--------------------------------------------------------------------------------------
Else
'--------------------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NCOLUMNS + 1, 1 To NROWS + 1)
    TEMP_MATRIX(1, 1) = "Symbol"
    
    For h = 1 To NCOLUMNS
        TEMP_MATRIX(1 + h, 1) = TICKERS_VECTOR(h, 1)
        If IsArray(DATA_GROUP(h)) = False Then: GoTo 1986
        TEMP_GROUP = DATA_GROUP(h)
        For i = 1 To NROWS
            If i > UBound(TEMP_GROUP) Then: Exit For
            
            TEMP1_STR = TEMP_GROUP(i)
            If TEMP1_STR = "" Then: GoTo 1985
            
            j = InStr(1, TEMP1_STR, DELIM_STR)
            If j = 0 Then: GoTo 1985
            k = InStr(1, TEMP1_STR, "(")
            If k = 0 Then: k = j
            
            TEMP_MATRIX(1, i + 1) = Trim(Mid(TEMP1_STR, 1, k - 1))
            TEMP2_STR = Trim(Mid(TEMP1_STR, j + 1, Len(TEMP1_STR) - j))
            
            If TEMP2_STR = "N/A" Or TEMP2_STR = "NaN%" Or TEMP2_STR = "NaN" Then
                TEMP_MATRIX(h + 1, i + 1) = 0
            Else
                TEMP_MATRIX(h + 1, i + 1) = TEMP2_STR
            End If
1985:
        Next i
1986:
    Next h
'--------------------------------------------------------------------------------------
End If
'--------------------------------------------------------------------------------------
            
YAHOO_KEY_STATISTICS_FUNC = CONVERT_STRING_NUMBER_FUNC(TEMP_MATRIX)

Exit Function
ERROR_LABEL:
YAHOO_KEY_STATISTICS_FUNC = Err.number
End Function

Sub PRINT_YAHOO_KEY_STATISTICS()

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DST_RNG As Excel.Range
Dim TEMP_RNG As Excel.Range
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

TICKERS_VECTOR = Excel.Application.InputBox("Symbol(s)", "Yahoo Finance", , , , , , 8)
Call EXCEL_TURN_OFF_EVENTS_FUNC
        
Set DST_RNG = WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC("_"), ActiveWorkbook).Cells(3, 3)
TEMP_MATRIX = YAHOO_KEY_STATISTICS_FUNC(TICKERS_VECTOR)
If IsArray(TEMP_MATRIX) = False Then: GoTo 1983
        
SROW = LBound(TEMP_MATRIX, 1)
NROWS = UBound(TEMP_MATRIX, 1)
            
SCOLUMN = LBound(TEMP_MATRIX, 2)
NCOLUMNS = UBound(TEMP_MATRIX, 2)
            
Set TEMP_RNG = Range(DST_RNG.Cells(SROW, SCOLUMN), DST_RNG.Cells(NROWS, NCOLUMNS))
TEMP_RNG.value = TEMP_MATRIX
GoSub FORMAT_LINE

1983:
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
        With IIf(NCOLUMNS = 2, .Columns(1), .Rows(1))
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
