Attribute VB_Name = "FINAN_ASSET_MOMENTS_WEEKLY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function ASSETS_52_WEEKS_ROC_FUNC(ByRef TICKERS_RNG As Variant, _
Optional ByVal END_DATE As Date = 0)

Dim i As Long

Dim NROWS As Long

Dim VAL_1 As Double
Dim VAL_4 As Double
Dim VAL_13 As Double
Dim VAL_26 As Double
Dim VAL_52 As Double

Dim START_DATE As Date

Dim TICKER_STR As String
Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If END_DATE = 0 Then: END_DATE = Now
START_DATE = DateSerial(Year(END_DATE) - 1, Month(END_DATE) - 1, Day(END_DATE))

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

ReDim TEMP_MATRIX(0 To NROWS, 0 To 6)

TEMP_MATRIX(0, 0) = "SYMBOL"
TEMP_MATRIX(0, 1) = "WEIGHTED ROI"
TEMP_MATRIX(0, 2) = "52 WEEKS"
TEMP_MATRIX(0, 3) = "26 WEEKS"
TEMP_MATRIX(0, 4) = "13 WEEKS"
TEMP_MATRIX(0, 5) = "4 WEEKS"
TEMP_MATRIX(0, 6) = "1 WEEK"

'-------------------------------------------------------------------------
For i = 1 To NROWS
'-------------------------------------------------------------------------
    VAL_52 = 0: VAL_26 = 0: VAL_13 = 0: VAL_4 = 0: VAL_1 = 0
    TICKER_STR = TICKERS_VECTOR(i, 1)
    TEMP_MATRIX(i, 0) = TICKER_STR
    DATA_VECTOR = YAHOO_HISTORICAL_DATA_SERIE_FUNC(TICKER_STR, START_DATE, END_DATE, "WEEKLY", "A", False, False, False)
    If IsArray(DATA_VECTOR) = False Then: GoTo 1983
    If UBound(DATA_VECTOR, 1) <= 26 + 1 Then: GoTo 1983
    If UBound(DATA_VECTOR, 1) <= 52 + 1 Then: GoTo 1982
'-------------------------------------------------------------------------
    VAL_52 = DATA_VECTOR(52 + 1, 1)
1982:
    VAL_26 = DATA_VECTOR(26 + 1, 1)
    If VAL_26 = 0 Then: GoTo 1983
    
    VAL_13 = DATA_VECTOR(13 + 1, 1)
    If VAL_13 = 0 Then: GoTo 1983
    
    VAL_4 = DATA_VECTOR(4 + 1, 1)
    If VAL_4 = 0 Then: GoTo 1983
    
    VAL_1 = DATA_VECTOR(1, 1)
    If VAL_1 = 0 Then: GoTo 1983
'-------------------------------------------------------------------------
    TEMP_MATRIX(i, 1) = 0.4 * (VAL_1 / VAL_4 - 1) + 0.33 * (VAL_4 / VAL_13 - 1) + 0.27 * (VAL_13 / VAL_26 - 1)
'-------------------------------------------------------------------------
    TEMP_MATRIX(i, 2) = VAL_52
    TEMP_MATRIX(i, 3) = VAL_26
    TEMP_MATRIX(i, 4) = VAL_13
    TEMP_MATRIX(i, 5) = VAL_4
    TEMP_MATRIX(i, 6) = VAL_1
'-------------------------------------------------------------------------
1983:
Next i
'-------------------------------------------------------------------------

ASSETS_52_WEEKS_ROC_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_52_WEEKS_ROC_FUNC = Err.number
End Function


'-------------------------------------------------------------------------------
'Throwing darts to pick Al's penny stocks
'-------------------------------------------------------------------------------

Function ASSET_52HIGH_52LOW_FUNC(ByRef TICKERS_RNG As Variant, _
Optional ByVal REFRESH_CALLER As Variant, _
Optional ByVal SERVER_STR As String = "UNITED STATES", _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim epsilon As Double

Dim MAX_PRICE_52HIGH_VAL As Double
Dim MAX_PRICE_52HIGH_STR As String

Dim MIN_PRICE_52HIGH_VAL As Double
Dim MIN_PRICE_52HIGH_STR As String

Dim MAX_52HIGH_52LOW_VAL As Double
Dim MAX_52HIGH_52LOW_STR As String

Dim MAX_PRICE_52LOW_VAL As Double
Dim MAX_PRICE_52LOW_STR As String

Dim MAX_PRICE_LOW_PERC_VAL As Double
Dim MAX_PRICE_LOW_PERC_STR As String

Dim MAX_PRICE_OPEN_PERC_VAL As Double
Dim MAX_PRICE_OPEN_PERC_STR As String

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim DATA_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

'-------------------------------------------------------------------------------
'I happened to be analyzing the successful strategy with penny stocks of
'my bestfriend, and decided to analyze a list of hot penny stocks - penny
'stocks are stocks that sell for less than a buck ... or, at least,
'less than $5.. I analyzed them based on (Current Price) / (52-week Low)
'and (52-week High) / (52-week Low)

'Consider this:

'1) There is more stored energy in Canadian coal than all the country's
'oil, natural gas, and oil sands combined.

'2) The price of coal is expected to double by the end of 2009.

'3) Canadian hard coking coal producers have secured a provisional price
'for fiscal 2008 of US$225 a tonne, a gigantic increase from US$98 in 2007.

'4) Recent prices in the spot market have topped US$300 a tonne.

'5) Over 50% of the U.S. electricity production is generated from coal.

'6) GCE (Cache Coal, Alberta) has increased its production by 50% over the
'past year.

'7) The average GCE cost of production during the quarter was $50 per tonne,
'down from $62 per tonne in the previous quarter.

'8) On April 10, 2007 GCE stock closed at 46 cents.
'Yesterday (April 1, 2008) it closed at $4.70.
'An increase by more than a factor of 10 !!!
'-------------------------------------------------------------------------------

epsilon = 0.00001
If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
    End If
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If

NROWS = UBound(TICKERS_VECTOR)

ReDim TEMP_VECTOR(1 To 1, 1 To 9)
TEMP_VECTOR(1, 1) = "Name"
TEMP_VECTOR(1, 2) = "time of last trade"
TEMP_VECTOR(1, 3) = "Volume"
TEMP_VECTOR(1, 4) = "Last Trade"
TEMP_VECTOR(1, 5) = "Open"
TEMP_VECTOR(1, 6) = "High"
TEMP_VECTOR(1, 7) = "Low"
TEMP_VECTOR(1, 8) = "52-week High"
TEMP_VECTOR(1, 9) = "52-week Low"

DATA_MATRIX = YAHOO_QUOTES_FUNC(TICKERS_VECTOR, TEMP_VECTOR, REFRESH_CALLER, False, SERVER_STR)

If IsArray(DATA_MATRIX) = False Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To NROWS, 1 To 14)
For j = 1 To 14: For i = 0 To NROWS: _
TEMP_MATRIX(i, j) = "": Next i: Next j

TEMP_MATRIX(0, 1) = "Company Name"
TEMP_MATRIX(0, 2) = "Time of Last Trade"
TEMP_MATRIX(0, 3) = "Volume"
TEMP_MATRIX(0, 4) = "Price"
TEMP_MATRIX(0, 5) = "Open"
TEMP_MATRIX(0, 6) = "High"
TEMP_MATRIX(0, 7) = "Low"
TEMP_MATRIX(0, 8) = "52 High"
TEMP_MATRIX(0, 9) = "52 Low"
TEMP_MATRIX(0, 10) = "Price/52Hi"
TEMP_MATRIX(0, 11) = "52Hi/52Lo"
TEMP_MATRIX(0, 12) = "Price/52Lo"
TEMP_MATRIX(0, 13) = "Price/Low%"
TEMP_MATRIX(0, 14) = "Price/Open%"

'-------------------------------------------------------------------------------

MAX_PRICE_52HIGH_VAL = -2 ^ 52
MIN_PRICE_52HIGH_VAL = 2 ^ 52
MAX_52HIGH_52LOW_VAL = -2 ^ 52
MAX_PRICE_52LOW_VAL = -2 ^ 52
MAX_PRICE_LOW_PERC_VAL = -2 ^ 52
MAX_PRICE_OPEN_PERC_VAL = -2 ^ 52

MAX_PRICE_52HIGH_STR = ""
MIN_PRICE_52HIGH_STR = ""
MAX_52HIGH_52LOW_STR = ""
MAX_PRICE_52LOW_STR = ""
MAX_PRICE_LOW_PERC_STR = ""
MAX_PRICE_OPEN_PERC_STR = ""

'-------------------------------------------------------------------------------
For i = 1 To NROWS
'-------------------------------------------------------------------------------
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1) & " (" & TICKERS_VECTOR(i, 1) & ")"
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
    TEMP_MATRIX(i, 3) = DATA_MATRIX(i, 3)
    TEMP_MATRIX(i, 4) = DATA_MATRIX(i, 4)
    TEMP_MATRIX(i, 5) = DATA_MATRIX(i, 5)
    TEMP_MATRIX(i, 6) = DATA_MATRIX(i, 6)
    TEMP_MATRIX(i, 7) = DATA_MATRIX(i, 7)
    TEMP_MATRIX(i, 8) = DATA_MATRIX(i, 8)
    TEMP_MATRIX(i, 9) = DATA_MATRIX(i, 9)
'-------------------------------------------------------------------------------
    If TEMP_MATRIX(i, 4) <> "" And _
       TEMP_MATRIX(i, 4) > epsilon Then
'-------------------------------------------------------------------------------
            If TEMP_MATRIX(i, 8) <> "" And _
               TEMP_MATRIX(i, 8) > epsilon Then
                    TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 4) / TEMP_MATRIX(i, 8)
                    If TEMP_MATRIX(i, 10) <> "" Then
                        If TEMP_MATRIX(i, 10) > MAX_PRICE_52HIGH_VAL Then
                            MAX_PRICE_52HIGH_VAL = TEMP_MATRIX(i, 10)
                            MAX_PRICE_52HIGH_STR = TEMP_MATRIX(i, 1)
                        End If
                        If TEMP_MATRIX(i, 10) < MIN_PRICE_52HIGH_VAL Then
                            MIN_PRICE_52HIGH_VAL = TEMP_MATRIX(i, 10)
                            MIN_PRICE_52HIGH_STR = TEMP_MATRIX(i, 1)
                        End If
                    End If
                    If TEMP_MATRIX(i, 9) <> "" And _
                       TEMP_MATRIX(i, 9) > epsilon Then
                        TEMP_MATRIX(i, 11) = TEMP_MATRIX(i, 8) / TEMP_MATRIX(i, 9)
                        If TEMP_MATRIX(i, 11) <> "" Then
                            If TEMP_MATRIX(i, 11) > MAX_52HIGH_52LOW_VAL Then
                                MAX_52HIGH_52LOW_VAL = TEMP_MATRIX(i, 11)
                                MAX_52HIGH_52LOW_STR = TEMP_MATRIX(i, 1)
                            End If
                        End If
                    Else
                            TEMP_MATRIX(i, 11) = ""
                    End If
            Else
                    TEMP_MATRIX(i, 10) = ""
                    TEMP_MATRIX(i, 11) = ""
            End If
'-------------------------------------------------------------------------------
            If TEMP_MATRIX(i, 9) <> "" And _
               TEMP_MATRIX(i, 9) > epsilon Then
                    TEMP_MATRIX(i, 12) = TEMP_MATRIX(i, 4) / TEMP_MATRIX(i, 9)
                    If TEMP_MATRIX(i, 12) <> "" Then
                        If TEMP_MATRIX(i, 12) > MAX_PRICE_52LOW_VAL Then
                            MAX_PRICE_52LOW_VAL = TEMP_MATRIX(i, 12)
                            MAX_PRICE_52LOW_STR = TEMP_MATRIX(i, 1)
                        End If
                    End If
            Else
                    TEMP_MATRIX(i, 12) = ""
            End If
'-------------------------------------------------------------------------------
            If TEMP_MATRIX(i, 7) <> "" And _
               TEMP_MATRIX(i, 7) > epsilon Then
                    TEMP_MATRIX(i, 13) = TEMP_MATRIX(i, 4) / TEMP_MATRIX(i, 7) - 1
                    If TEMP_MATRIX(i, 13) = 1 Then: TEMP_MATRIX(i, 13) = ""
                    If TEMP_MATRIX(i, 13) <> "" Then
                        If TEMP_MATRIX(i, 13) > MAX_PRICE_LOW_PERC_VAL Then
                            MAX_PRICE_LOW_PERC_VAL = TEMP_MATRIX(i, 13)
                            MAX_PRICE_LOW_PERC_STR = TEMP_MATRIX(i, 1)
                        End If
                    End If
            Else
                    TEMP_MATRIX(i, 13) = ""
            End If
'-------------------------------------------------------------------------------
            If TEMP_MATRIX(i, 5) <> "" And _
               TEMP_MATRIX(i, 5) > epsilon Then
                    TEMP_MATRIX(i, 14) = TEMP_MATRIX(i, 4) / TEMP_MATRIX(i, 5) - 1
                    If TEMP_MATRIX(i, 14) = 1 Then: TEMP_MATRIX(i, 14) = ""
                    If TEMP_MATRIX(i, 14) <> "" Then
                        If TEMP_MATRIX(i, 14) > MAX_PRICE_OPEN_PERC_VAL Then
                            MAX_PRICE_OPEN_PERC_VAL = TEMP_MATRIX(i, 14)
                            MAX_PRICE_OPEN_PERC_STR = TEMP_MATRIX(i, 1)
                        End If
                    End If
            Else
                    TEMP_MATRIX(i, 14) = ""
            End If
'-------------------------------------------------------------------------------
    Else
        TEMP_MATRIX(i, 10) = ""
        TEMP_MATRIX(i, 11) = ""
        TEMP_MATRIX(i, 12) = ""
        TEMP_MATRIX(i, 13) = ""
        TEMP_MATRIX(i, 14) = ""
    End If
Next i
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
Select Case OUTPUT
'-------------------------------------------------------------------------------
Case 0
'-------------------------------------------------------------------------------
    ASSET_52HIGH_52LOW_FUNC = TEMP_MATRIX
'-------------------------------------------------------------------------------
Case Else
'-------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 6, 1 To 3)
    
    TEMP_VECTOR(1, 1) = "Max of Price/52High"
    TEMP_VECTOR(1, 2) = MAX_PRICE_52HIGH_STR
    TEMP_VECTOR(1, 3) = MAX_PRICE_52HIGH_VAL
    
    TEMP_VECTOR(2, 1) = "Min of Price/52High"
    TEMP_VECTOR(2, 2) = MIN_PRICE_52HIGH_STR
    TEMP_VECTOR(2, 3) = MIN_PRICE_52HIGH_VAL
    
    TEMP_VECTOR(3, 1) = "Max of 52Hi/52Low"
    TEMP_VECTOR(3, 2) = MAX_52HIGH_52LOW_STR
    TEMP_VECTOR(3, 3) = MAX_52HIGH_52LOW_VAL
    
    TEMP_VECTOR(4, 1) = "Max of Price/52Low"
    TEMP_VECTOR(4, 2) = MAX_PRICE_52LOW_STR
    TEMP_VECTOR(4, 3) = MAX_PRICE_52LOW_VAL
    
    TEMP_VECTOR(5, 1) = "Max of Price/Low%"
    TEMP_VECTOR(5, 2) = MAX_PRICE_LOW_PERC_STR
    TEMP_VECTOR(5, 3) = MAX_PRICE_LOW_PERC_VAL
    
    TEMP_VECTOR(6, 1) = "Max of Price/Open%"
    TEMP_VECTOR(6, 2) = MAX_PRICE_OPEN_PERC_STR
    TEMP_VECTOR(6, 3) = MAX_PRICE_OPEN_PERC_VAL

    If OUTPUT = 1 Then
        ASSET_52HIGH_52LOW_FUNC = TEMP_VECTOR
    Else
        ASSET_52HIGH_52LOW_FUNC = Array(TEMP_MATRIX, TEMP_VECTOR)
    End If
'-------------------------------------------------------------------------------
End Select
'-------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
ASSET_52HIGH_52LOW_FUNC = Err.number
End Function


Sub PRINT_52HIGH_52LOW_ANALYSIS()

Dim SROW As Long
Dim SCOLUMN As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DATA_RNG As Excel.Range
Dim DST_RNG As Excel.Range
Dim TEMP_RNG As Excel.Range

Dim SWITCH_FLAG As Boolean
Dim TEMP_GROUP As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

Set DATA_RNG = Excel.Application.InputBox("Symbols", "52 Weeks H/L Analysis", , , , , , 8)
If DATA_RNG Is Nothing Then: Exit Sub

Call EXCEL_TURN_OFF_EVENTS_FUNC

Set DST_RNG = _
WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC("_"), _
ActiveWorkbook).Cells(3, 3)

SWITCH_FLAG = True
TEMP_GROUP = ASSET_52HIGH_52LOW_FUNC(DATA_RNG, , , 2)
TEMP_MATRIX = TEMP_GROUP(UBound(TEMP_GROUP))
If IsArray(TEMP_MATRIX) = False Then: GoTo 1983
        
SROW = LBound(TEMP_MATRIX, 1)
NROWS = UBound(TEMP_MATRIX, 1)
            
SCOLUMN = LBound(TEMP_MATRIX, 2)
NCOLUMNS = UBound(TEMP_MATRIX, 2)
            
Set TEMP_RNG = Range(DST_RNG.Cells(SROW, SCOLUMN), _
DST_RNG.Cells(NROWS, NCOLUMNS))
TEMP_RNG.value = TEMP_MATRIX
GoSub FORMAT_LINE

Set DST_RNG = DST_RNG.Cells(NROWS + 3, 1)

SWITCH_FLAG = False
TEMP_MATRIX = TEMP_GROUP(LBound(TEMP_GROUP))
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
        With IIf(SWITCH_FLAG = False, .Rows(1), .Columns(1))
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

