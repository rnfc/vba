Attribute VB_Name = "FINAN_FUNDAM_PE_PRICE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'--------------------------------------------------------------------------------
'Rule #1 MOS Price
'http://www.philtown.typepad.com/phil_towns_blog/aboutruleone.html
'--------------------------------------------------------------------------------
Function MOS_PRICE_FUNC(ByVal HIGH_PE_VAL As Double, _
ByVal LOW_PE_VAL As Double, _
ByVal CURRENT_EPS_VAL As Double, _
ByVal PROJ_GROWTH_RATE_VAL As Double, _
Optional ByVal DISCOUNT_RATE_VAL As Double = 0.15, _
Optional ByVal FORWARD_PERIODS_VAL As Long = 10)

Dim FV_VAL As Double

On Error GoTo ERROR_LABEL

'MSN % Growth Rate -- P/E Ratio 5-Year High -- Company
'HIGH_PE_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER1_STR, 99, ERROR_STR)' 5-Year High P/E

'MSN % Growth Rate -- P/E Ratio 5-Year Low -- Company
'LOW_PE_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER1_STR, 102, ERROR_STR)' 5-Year Low P/E

'MSN Financial Highlights -- Earnings/Share
'CURRENT_EPS_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER1_STR, 44, ERROR_STR)' Current EPS

'MSN Earnings Growth Rates -- Company -- 4
'PROJ_GROWTH_RATE_VAL = RETRIEVE_WEB_DATA_ELEMENT_FUNC(TICKER1_STR, 442, ERROR_STR)' 5-Year Projected Growth Rate

'If HIGH_PE_VAL > 50 Then HIGH_PE_VAL = 50
FV_VAL = CURRENT_EPS_VAL * (1 + PROJ_GROWTH_RATE_VAL) ^ FORWARD_PERIODS_VAL
'FV(PROJ_GROWTH_RATE_VAL, FORWARD_PERIODS_VAL, 0, -CURRENT_EPS_VAL)
                            
MOS_PRICE_FUNC = ((FV_VAL * (HIGH_PE_VAL + LOW_PE_VAL) / 2) / (1 + DISCOUNT_RATE_VAL) ^ FORWARD_PERIODS_VAL) / 2
'PV(DISCOUNT_RATE_VAL, FORWARD_PERIODS_VAL, 0, -FV_VAL * (HIGH_PE_VAL + LOW_PE_VAL) / 2) / 2
                            
Exit Function
ERROR_LABEL:
MOS_PRICE_FUNC = Err.number
End Function


'P/E  versus  (Price - 52 Week Low)/(52 week High- 52 Week Low)
'The logic behind this model is simple. We look at the 52-week Hi
'and Low for each stock on and determine where the current stock
'price is, within this range. For example, if Company X was 2.2%
'above the 52-week Low, within the 52-week range. It's Price/Earnings
'ratio is 19.1 so we stick a point at (2.2%, 19.1) for X ... and do
'this for all stocks.

Function ASSET_MIN_MAX_PE_FUNC(ByRef TICKERS_RNG As Variant, _
Optional ByVal REFRESH_CALLER As Variant, _
Optional ByVal SERVER_STR As String = "UNITED STATES", _
Optional ByVal OUTPUT As Integer = 1)

'X: PRICE_RATIO
'Y: PE_RATIO

'Column 1 of Data Rng Must BE: Reference Name for each Stock
'Column 2 of Data Rng Must BE: Reference Date
'Column 3 --> 52-week Range
'Column 4 --> P/E Ratio
'Column 5 --> Last Trade (Price Only)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_STR As String
Dim TEMP_VAL As Variant
Dim DECIMAL_STR As String

Dim MIN_PE_VAL As Double
Dim MIN_PE_STR As String

Dim MAX_PE_VAL As Double
Dim MAX_PE_STR As String

Dim MIN_RATIO_VAL As Double
Dim MIN_RATIO_STR As String

Dim MAX_RATIO_VAL As Double
Dim MAX_RATIO_STR As String

Dim MIN_ORIGIN_VAL As Double 'Ratio
Dim MIN_ORIGIN_STR As String

Dim MAX_ORIGIN_VAL As Double 'Ratio
Dim MAX_ORIGIN_STR As String

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim TICKERS_VECTOR As Variant

On Error GoTo ERROR_LABEL

If IsArray(TICKERS_RNG) = True Then
    TICKERS_VECTOR = TICKERS_RNG
    If UBound(TICKERS_VECTOR, 1) = 1 Then: _
        TICKERS_VECTOR = MATRIX_TRANSPOSE_FUNC(TICKERS_VECTOR)
Else
    ReDim TICKERS_VECTOR(1 To 1, 1 To 1)
    TICKERS_VECTOR(1, 1) = TICKERS_RNG
End If

ReDim TEMP_VECTOR(1 To 1, 1 To 5)
TEMP_VECTOR(1, 1) = "Name"
TEMP_VECTOR(1, 2) = "time of last trade"
TEMP_VECTOR(1, 3) = "52-week Range"
TEMP_VECTOR(1, 4) = "P/E Ratio"
TEMP_VECTOR(1, 5) = "Last Trade"

DATA_MATRIX = YAHOO_QUOTES_FUNC(TICKERS_VECTOR, TEMP_VECTOR, REFRESH_CALLER, False, SERVER_STR)
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 10)

TEMP_MATRIX(0, 1) = ("NAME")
TEMP_MATRIX(0, 2) = ("TIME OF LAST TRADE")
TEMP_MATRIX(0, 3) = ("52-WEEK RANGE")
TEMP_MATRIX(0, 4) = ("P/E RATIO")
TEMP_MATRIX(0, 5) = ("LAST TRADE (PRICE ONLY)")
TEMP_MATRIX(0, 6) = ("LOW PRICE")
TEMP_MATRIX(0, 7) = ("HIGH PRICE")
TEMP_MATRIX(0, 8) = ("[ PRICE - 52 WEEK LOW ] / [ 52 WEEK HIGH - 52 WEEK LOW ]")
TEMP_MATRIX(0, 9) = ("CLOSE TO ORIGIN: MIN")
TEMP_MATRIX(0, 10) = ("CLOSE TO ORIGIN: MAX")


DECIMAL_STR = DECIMAL_SEPARATOR_FUNC()
'----------------------------FIRST PASS: PRICE RATIO CALCULATIONS-------------------

For i = 1 To NROWS
    
    TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1) & " (" & TICKERS_VECTOR(i, 1) & ")"
    TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2) 'Time of Last Trade
    
    If IS_NUMERIC_FUNC(DATA_MATRIX(i, 4), DECIMAL_STR) = True Then
        TEMP_MATRIX(i, 4) = DATA_MATRIX(i, 4) 'P/E Ratio
    Else
        TEMP_MATRIX(i, 4) = 0 'P/E Ratio
    End If
    
    If TEMP_MATRIX(i, 4) = 0 Then: j = j + 1
    
    If IS_NUMERIC_FUNC(DATA_MATRIX(i, 5), DECIMAL_STR) = True Then
        TEMP_MATRIX(i, 5) = DATA_MATRIX(i, 5) 'Last Trade (Price Only)
    Else
        TEMP_MATRIX(i, 5) = 0 'Last Trade (Price Only)
    End If
            
    TEMP_MATRIX(i, 3) = DATA_MATRIX(i, 3) '52-Week Range
    
    l = Len(TEMP_MATRIX(i, 3))
    k = InStr(1, TEMP_MATRIX(i, 3), "-")

    TEMP_STR = TEMP_MATRIX(i, 3)
    TEMP_STR = Trim(Mid(TEMP_STR, 1, k - 1)) 'Low Price
    
    TEMP_VAL = CONVERT_STRING_NUMBER_FUNC(TEMP_STR)
    
    If IS_NUMERIC_FUNC(TEMP_VAL, DECIMAL_STR) = True Then
        TEMP_MATRIX(i, 6) = TEMP_VAL
    Else
        TEMP_MATRIX(i, 6) = 0
    End If
    
    TEMP_STR = TEMP_MATRIX(i, 3)
    TEMP_STR = Trim(Mid(TEMP_STR, 1 + k, l - k + 1)) 'High Price
    TEMP_VAL = CONVERT_STRING_NUMBER_FUNC(TEMP_STR)
    
    If IS_NUMERIC_FUNC(TEMP_VAL, DECIMAL_STR) = True Then
        TEMP_MATRIX(i, 7) = TEMP_VAL
    Else
        TEMP_MATRIX(i, 7) = 0
    End If
    
    If Abs(TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 6)) > 0 Then
        TEMP_MATRIX(i, 8) = (TEMP_MATRIX(i, 5) - TEMP_MATRIX(i, 6)) / _
                            (TEMP_MATRIX(i, 7) - TEMP_MATRIX(i, 6))
                            'Price Ratio
    Else
        TEMP_MATRIX(i, 8) = 0
    End If
    
'-----------------------------------------------------------------------------------
    If i = 1 Then
'-----------------------------------------------------------------------------------
        If TEMP_MATRIX(1, 4) <> 0 Then
            MIN_PE_VAL = TEMP_MATRIX(1, 4)
            MAX_PE_VAL = TEMP_MATRIX(1, 4)
        Else
            MIN_PE_VAL = 2 ^ 52
            MAX_PE_VAL = -2 ^ 52
        End If

        If TEMP_MATRIX(1, 8) <> 0 Then
            MIN_RATIO_VAL = TEMP_MATRIX(1, 8)
            MAX_RATIO_VAL = TEMP_MATRIX(1, 8)
        Else
            MIN_RATIO_VAL = 2 ^ 52
            MAX_RATIO_VAL = -2 ^ 52
        End If
    
        MIN_PE_STR = TEMP_MATRIX(1, 1)
        MAX_PE_STR = TEMP_MATRIX(1, 1)
        
        MIN_RATIO_STR = TEMP_MATRIX(1, 1)
        MAX_RATIO_STR = TEMP_MATRIX(1, 1)
'-----------------------------------------------------------------------------------
    Else 'If i <> 1 Then
'-----------------------------------------------------------------------------------
        If TEMP_MATRIX(i, 4) <> 0 Then
            
            MIN_PE_VAL = MINIMUM_FUNC(TEMP_MATRIX(i, 4), MIN_PE_VAL)
            If MIN_PE_VAL = TEMP_MATRIX(i, 4) Then
                MIN_PE_STR = TEMP_MATRIX(i, 1)
            End If
            
            MAX_PE_VAL = MAXIMUM_FUNC(TEMP_MATRIX(i, 4), MAX_PE_VAL)
            If MAX_PE_VAL = TEMP_MATRIX(i, 4) Then
                MAX_PE_STR = TEMP_MATRIX(i, 1)
            End If
        End If
        
        If TEMP_MATRIX(i, 8) <> 0 Then
            
            MIN_RATIO_VAL = MINIMUM_FUNC(TEMP_MATRIX(i, 8), MIN_RATIO_VAL)
            If Abs(MIN_RATIO_VAL - TEMP_MATRIX(i, 8)) <= 10 ^ -15 Then
                MIN_RATIO_STR = TEMP_MATRIX(i, 1)
            End If
            
            MAX_RATIO_VAL = MAXIMUM_FUNC(TEMP_MATRIX(i, 8), MAX_RATIO_VAL)
            If Abs(MAX_RATIO_VAL - TEMP_MATRIX(i, 8)) <= 10 ^ -15 Then
                MAX_RATIO_STR = TEMP_MATRIX(i, 1)
            End If
        End If
'-----------------------------------------------------------------------------------
    End If
'-----------------------------------------------------------------------------------
Next i

'SECOND PASS:-----Close to Origin PE_RATIO: Z = ( A2 + 100 x B2 )^0.5---------------------
'-----------------------Where: A = P/E Ratio - Min(P/E Ratio)-----------------------------
'--------------------------------------B = Ratio - Min(Ratio)-----------------------------

For i = 1 To NROWS
    If (TEMP_MATRIX(i, 4) <> 0) And (TEMP_MATRIX(i, 8) <> 0) Then
        TEMP_MATRIX(i, 9) = Sqr((TEMP_MATRIX(i, 4) - MIN_PE_VAL) ^ 2 + (100 * (TEMP_MATRIX(i, 8) - MIN_RATIO_VAL)) ^ 2)
    Else
        TEMP_MATRIX(i, 9) = 0
    End If
'-----------------------------------------------------------------------------------
    If i = 1 Then
'-----------------------------------------------------------------------------------
        If TEMP_MATRIX(i, 9) <> 0 Then
            MIN_ORIGIN_VAL = TEMP_MATRIX(i, 9)
        Else
            MIN_ORIGIN_VAL = 2 ^ 52
        End If

        MIN_ORIGIN_STR = TEMP_MATRIX(1, 1)
'-----------------------------------------------------------------------------------
    Else 'If i <> 1 Then
'-----------------------------------------------------------------------------------
        If TEMP_MATRIX(i, 9) <> 0 Then
            MIN_ORIGIN_VAL = MINIMUM_FUNC(TEMP_MATRIX(i, 9), MIN_ORIGIN_VAL)
                If MIN_ORIGIN_VAL = TEMP_MATRIX(i, 9) Then
                    MIN_ORIGIN_STR = TEMP_MATRIX(i, 1)
                End If
        End If
'-----------------------------------------------------------------------------------
    End If
'-----------------------------------------------------------------------------------
Next i
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
For i = 1 To NROWS
'-----------------------------------------------------------------------------------
    If (TEMP_MATRIX(i, 9) <> 0) And (MIN_ORIGIN_VAL <> TEMP_MATRIX(i, 9)) _
        And (TEMP_MATRIX(i, 4) <> MIN_RATIO_VAL) Then
        TEMP_MATRIX(i, 10) = TEMP_MATRIX(i, 9)
    Else
        TEMP_MATRIX(i, 10) = 0
    End If
    
'-----------------------------------------------------------------------------------
    If i = 1 Then
'-----------------------------------------------------------------------------------
        If TEMP_MATRIX(i, 10) <> 0 Then
            MAX_ORIGIN_VAL = TEMP_MATRIX(i, 10)
        Else
            MAX_ORIGIN_VAL = 2 ^ 52
        End If
        MAX_ORIGIN_STR = TEMP_MATRIX(1, 1)
'-----------------------------------------------------------------------------------
    Else 'If i <> 1 Then
'-----------------------------------------------------------------------------------
        If TEMP_MATRIX(i, 10) <> 0 Then
            MAX_ORIGIN_VAL = MINIMUM_FUNC(TEMP_MATRIX(i, 10), MAX_ORIGIN_VAL)
            If MAX_ORIGIN_VAL = TEMP_MATRIX(i, 10) Then
                MAX_ORIGIN_STR = TEMP_MATRIX(i, 1)
            End If
        End If
'-----------------------------------------------------------------------------------
    End If
'-----------------------------------------------------------------------------------
Next i
'-----------------------------------------------------------------------------------

ReDim TEMP_VECTOR(0 To 6, 1 To 3)
        
TEMP_VECTOR(0, 1) = Format(j, "0") & " STOCKS HAVE P/E <= 0"
TEMP_VECTOR(0, 2) = ""
TEMP_VECTOR(0, 3) = ""

TEMP_VECTOR(1, 1) = "MAXIMUM P/E RATIO"
TEMP_VECTOR(1, 2) = MAX_PE_STR
TEMP_VECTOR(1, 3) = MAX_PE_VAL

TEMP_VECTOR(2, 1) = "MINIMUM P/E RATIO"
TEMP_VECTOR(2, 2) = MIN_PE_STR
TEMP_VECTOR(2, 3) = MIN_PE_VAL

TEMP_VECTOR(3, 1) = "MAXIMUM PRICE RATIO"
TEMP_VECTOR(3, 2) = MAX_RATIO_STR
TEMP_VECTOR(3, 3) = MAX_RATIO_VAL

TEMP_VECTOR(4, 1) = "MINIMUM PRICE-RATIO"
TEMP_VECTOR(4, 2) = MIN_RATIO_STR
TEMP_VECTOR(4, 3) = MIN_RATIO_VAL

TEMP_VECTOR(5, 1) = "CLOSE TO ORIGIN: MAX"
TEMP_VECTOR(5, 2) = MAX_ORIGIN_STR
TEMP_VECTOR(5, 3) = MAX_ORIGIN_VAL

TEMP_VECTOR(6, 1) = "CLOSE TO ORIGIN: MIN"
TEMP_VECTOR(6, 2) = MIN_ORIGIN_STR
TEMP_VECTOR(6, 3) = MIN_ORIGIN_VAL

Select Case OUTPUT
Case 0
    ASSET_MIN_MAX_PE_FUNC = TEMP_MATRIX
Case 1
    ASSET_MIN_MAX_PE_FUNC = TEMP_VECTOR
Case Else
    ASSET_MIN_MAX_PE_FUNC = Array(TEMP_MATRIX, TEMP_VECTOR)
End Select

Exit Function
ERROR_LABEL:
ASSET_MIN_MAX_PE_FUNC = Err.number
End Function


Sub PRINT_MIN_MAX_PE_ANALYSIS()

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

Set DATA_RNG = Excel.Application.InputBox("Symbols", "Min-Max PE Analysis", , , , , , 8)
If DATA_RNG Is Nothing Then: Exit Sub

Call EXCEL_TURN_OFF_EVENTS_FUNC

Set DST_RNG = WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC("_"), ActiveWorkbook).Cells(3, 3)

SWITCH_FLAG = True
TEMP_GROUP = ASSET_MIN_MAX_PE_FUNC(DATA_RNG, , , 2)
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

