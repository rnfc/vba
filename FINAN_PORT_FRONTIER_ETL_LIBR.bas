Attribute VB_Name = "FINAN_PORT_FRONTIER_ETL_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'**********************************************************************************
'**********************************************************************************
'FUNCTION      : PORT_ETL_OPTIMIZER_FUNC

'DESCRIPTION   : 'Volatility is the industry standard risk measure. However volatility has
'two different deficiencies, namely it is a symmetric measure that penalizes
'upside performance and it does not adequately capture extreme downside risk.
'Furthermore, volatility does not have the properties of a coherent risk
'measure*.

'Another widely used risk measure is Value-at-Risk (VaR).  For a given lower
'tail probability , is defined as the p-quantile of the portfolio returns
'distribution.  As such, it is an asymmetric downside risk measure. However,
'VaR has a number of serious limitations: (a) it provides a very limited kind
'of information about extreme losses, (b) it is not a coherent risk measure,
'and (c) it is a much too rough non-convex objective function for portfolio
'optimization purposes.

'In recognition of these facts, the following function offers next-generation
'portfolio optimization using a new downside risk measure known as Expected
'Tail Loss (PORT_EXPECTED_TAIL_LOSS_FUNC).  For a given lower tail probability
'p, e.g., or , is the average or expected loss conditioned on the loss being
'greater than the corresponding .

'As compared with , has the following advantages:  (a) it is a highly
'informative measure of extreme downside losses, (b) it is a coherent risk
'measure, and (c) its use in portfolio optimization leads to a relatively smooth
'convex optimization problem that can be solved by linear programming (LP)
'methods*.

'This portfolio optimization technique yields superior risk adjusted returns
'relative to conventional Markowitz portfolios at equivalent
'PORT_EXPECTED_TAIL_LOSS_FUNC risks, often dramatically so.
  
'References
'* Artzner, P., Delbaen, F., Deber, J., and Heath, D. (1999).  `“Coherent
'measures of risk“, Mathematical Finance, 9, 3, 203-228.

'** Rockafellar, R. T. and Uryasev, S. (2000).  “Optimization of
'conditional-value-at-risk”, Journal of Risk, 2, 21-41.

'Rachev, S. T., Menn, C. and Fabozzi, F. J. (2005).  Fat-Tailed and Skewed
'Asset Return Distributions: Implications for Risk Management, Portfolio
'Selection and Option Pricing, Wiley.

'Rachev, S. T., Stoyan S. and  Fabozzi F.J. (2006).   Probability Metrics and
'Quantitative Finance: Applications to Risk Measurement and  Portfolio
'Optimization, Wiley ( to appear).


'LIBRARY       : PORTFOLIO
'GROUP         : OPTIMIZER
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'**********************************************************************************
'**********************************************************************************

Public Function RNG_PORT_ETL_FRONTIER_FUNC(ByRef DST_RNG As Excel.Range, _
ByRef DATA_RNG As Excel.Range, _
Optional ByRef WEIGHTS_RNG As Excel.Range, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal NTRIALS As Integer = 4, _
Optional ByVal CONFIDENCE_VAL As Double = 0.975, _
Optional ByVal TARGET_PORT_RETURN As Double = 0.01, _
Optional ByVal TOTAL_EXPOSURE As Double = 1, _
Optional ByVal WEIGHTS_LOWER_BOUND As Double = 0, _
Optional ByVal WEIGHTS_UPPER_BOUND As Double = 1, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

'DATA_RNG --> Header, Dates, Prices/Returns

'nTRIALS --> Number of Portfolios between minimum risk &
'            maximum return portfolio to be calculated

'CONFIDENCE --> Used in calculation of expected loss

Dim i As Long
Dim j As Long
Dim k As Integer

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim DELTA_RETURN As Double

'-------------------------------------------------------------------
Dim TARGET_RNG As Excel.Range
Dim CHG_CELLS_RNG As Excel.Range
Dim REFER_CELLS_RNG() As Excel.Range
Dim CONST_CELLS_RNG() As Excel.Range
Dim RELATION_ARR() As Integer
'-------------------------------------------------------------------
Dim TEMP_RNG As Excel.Range
Dim CONTROL_RNG As Excel.Range
Dim HEADERS_RNG As Excel.Range
Dim RETURNS_RNG As Excel.Range
Dim POSITION_RNG As Excel.Range
Dim SAMPLING_RNG As Excel.Range
Dim PORT_RETURNS_RNG As Excel.Range
'-------------------------------------------------------------------
Dim SOLVER_FLAG As Boolean
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

RNG_PORT_ETL_FRONTIER_FUNC = False
'If CHECK_EXCEL_SOLVER_FUNC() = False Then: GoTo ERROR_LABEL

'-------------------------------------------------------------------
With Excel.Application
    .Calculation = xlCalculationAutomatic
    .EnableEvents = False
    .ScreenUpdating = False
    .Cursor = xlWait
    .StatusBar = False
End With
'-------------------------------------------------------------------

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

Set POSITION_RNG = DST_RNG.Cells(5, 1)
Set SAMPLING_RNG = Range(POSITION_RNG.Cells(1, 1), _
                         POSITION_RNG.Cells(NTRIALS + 2, 7 + NCOLUMNS))
'-------------------------------------------------------------------------------------
Set POSITION_RNG = DST_RNG.Cells(9 + NTRIALS, 2)
Set DATA_RNG = Range(POSITION_RNG.Cells(1, 1), POSITION_RNG.Cells(NROWS, NCOLUMNS))
DATA_RNG.value = DATA_MATRIX
'-------------------------------------------------------------------------------------
Set RETURNS_RNG = Range(POSITION_RNG.Cells(2, 2), POSITION_RNG.Cells(NROWS, NCOLUMNS))
If IsArray(WEIGHTS_RNG) = True Then
    DATA_MATRIX = WEIGHTS_RNG
    If UBound(DATA_MATRIX, 2) = 1 Then: DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)
Else
    ReDim DATA_MATRIX(1 To 1, 1 To NCOLUMNS)
    For j = 1 To NCOLUMNS
        DATA_MATRIX(1, j) = 1 / (NCOLUMNS - 1) 'Equal Weight to each asset
    Next j
End If
'---------------------------------------------------------------------------------------
Set WEIGHTS_RNG = Range(POSITION_RNG.Cells(0, 2), POSITION_RNG.Cells(0, NCOLUMNS))
Set HEADERS_RNG = Range(POSITION_RNG.Cells(1, 2), POSITION_RNG.Cells(1, NCOLUMNS))
WEIGHTS_RNG.value = DATA_MATRIX
'---------------------------------------------------------------------------------------
POSITION_RNG.Cells(0, 1) = "Weights"
POSITION_RNG.Cells(1, 0) = "Portfolio Returns"
Set PORT_RETURNS_RNG = Range(POSITION_RNG.Cells(2, 0), POSITION_RNG.Cells(NROWS, 0))
PORT_RETURNS_RNG.FormulaArray = "=TRANSPOSE(MMULT(" & WEIGHTS_RNG.Address & _
                                ",TRANSPOSE(" & RETURNS_RNG.Address & ")))"
'-------------------------------------------------------------------------------------
Set POSITION_RNG = DST_RNG.Cells(1, 1)
Set CONTROL_RNG = Range(POSITION_RNG.Cells(1, 1), POSITION_RNG.Cells(2, 14))
CONTROL_RNG.Rows(1).value = Array("Confidence Level", "No. Trials", _
                                  "Target Return", "Exposure", "Return", _
                                  "StDev", "VaR", "ETL", "MAD", _
                                  "Max Loss", "", "Exposure Limit", _
                                  "Exposure Lower Bound", "Exposure Upper Bound")
'---------------------------------------------------------------------------------------
CONTROL_RNG.Cells(2, 1).value = CONFIDENCE_VAL
CONTROL_RNG.Cells(2, 2).value = NTRIALS
CONTROL_RNG.Cells(2, 3).value = TARGET_PORT_RETURN
CONTROL_RNG.Cells(2, 4).formula = "=SUM(" & WEIGHTS_RNG.Address & ")"
CONTROL_RNG.Cells(2, 5).formula = "=AVERAGE(" & PORT_RETURNS_RNG.Address & ")"
CONTROL_RNG.Cells(2, 6).formula = "=STDEVP(" & PORT_RETURNS_RNG.Address & ")"
CONTROL_RNG.Cells(2, 7).formula = "=PERCENTILE(" & PORT_RETURNS_RNG.Address & ",1-" & _
                                    CONTROL_RNG.Cells(2, 1).Address & ")"
'---------------------------------------------------------------------------------------
CONTROL_RNG.Cells(2, 8).formula = "=PORT_EXPECTED_TAIL_LOSS_FUNC(" & _
                                    PORT_RETURNS_RNG.Address & ",," & _
                                    CONTROL_RNG.Cells(2, 1).Address & ",0,0,0)"
'---------------------------------------------------------------------------------------
CONTROL_RNG.Cells(2, 9).formula = "=AVEDEV(" & PORT_RETURNS_RNG.Address & ")"
CONTROL_RNG.Cells(2, 10).formula = "=MIN(" & PORT_RETURNS_RNG.Address & ")"
'---------------------------------------------------------------------------------------
CONTROL_RNG.Cells(2, 12).value = TOTAL_EXPOSURE
CONTROL_RNG.Cells(2, 13).value = WEIGHTS_LOWER_BOUND
CONTROL_RNG.Cells(2, 14).value = WEIGHTS_UPPER_BOUND
'-------------------------------------------------------------------------------------
k = IIf(VERSION = 0, 1, 2)
ReDim DATA_MATRIX(1 To NTRIALS + 2, 1 To 7 + NCOLUMNS)
'-------------------------------------------------------------------------------------

ReDim RELATION_ARR(1 To 3)
ReDim REFER_CELLS_RNG(1 To 3)
ReDim CONST_CELLS_RNG(1 To 3)

Set REFER_CELLS_RNG(1) = WEIGHTS_RNG
Set CONST_CELLS_RNG(1) = CONTROL_RNG.Cells(2, 13)
RELATION_ARR(1) = 3

Set REFER_CELLS_RNG(2) = WEIGHTS_RNG
Set CONST_CELLS_RNG(2) = CONTROL_RNG.Cells(2, 14)
RELATION_ARR(2) = 1

Set REFER_CELLS_RNG(3) = CONTROL_RNG.Cells(2, 4)
Set CONST_CELLS_RNG(3) = CONTROL_RNG.Cells(2, 12)
RELATION_ARR(3) = 2

If VERSION = 0 Then 'Max Portfolio Returns
    Set TARGET_RNG = CONTROL_RNG.Cells(2, 8)
    Set CHG_CELLS_RNG = WEIGHTS_RNG
Else 'Min Portfolio Standard Deviation
    Set TARGET_RNG = CONTROL_RNG.Cells(2, 6)
    Set CHG_CELLS_RNG = WEIGHTS_RNG
End If

SOLVER_FLAG = CALL_EXCEL_SOLVER_FUNC(TARGET_RNG, CHG_CELLS_RNG, _
     REFER_CELLS_RNG(), CONST_CELLS_RNG(), RELATION_ARR(), _
     k, 0, 100, 100, 0.000001, _
     False, False, 1, 1, 1, 5, False, 0.0001, True)
If SOLVER_FLAG = False Then: GoTo ERROR_LABEL

DATA_MATRIX(1, 1) = TARGET_PORT_RETURN
DATA_MATRIX(1, 2) = CONTROL_RNG.Cells(2, 5).value
DATA_MATRIX(1, 3) = CONTROL_RNG.Cells(2, 6).value
DATA_MATRIX(1, 4) = CONTROL_RNG.Cells(2, 7).value
DATA_MATRIX(1, 5) = Abs(CONTROL_RNG.Cells(2, 8).value)
DATA_MATRIX(1, 6) = CONTROL_RNG.Cells(2, 9).value
DATA_MATRIX(1, 7) = Abs(CONTROL_RNG.Cells(2, 10).value)

For j = 1 To NCOLUMNS
    DATA_MATRIX(1, 7 + j) = WEIGHTS_RNG.Cells(1, j)
Next j

'///////////////////////////// Maximum Return Portfolio\\\\\\\\\\\\\\\\\\\\\\\\\\\

Set REFER_CELLS_RNG(1) = WEIGHTS_RNG
Set CONST_CELLS_RNG(1) = CONTROL_RNG.Cells(2, 13)
RELATION_ARR(1) = 3

Set REFER_CELLS_RNG(2) = WEIGHTS_RNG
Set CONST_CELLS_RNG(2) = CONTROL_RNG.Cells(2, 14)
RELATION_ARR(2) = 1

Set REFER_CELLS_RNG(3) = CONTROL_RNG.Cells(2, 4)
Set CONST_CELLS_RNG(3) = CONTROL_RNG.Cells(2, 12)
RELATION_ARR(3) = 2

Set TARGET_RNG = CONTROL_RNG.Cells(2, 5)
Set CHG_CELLS_RNG = WEIGHTS_RNG

SOLVER_FLAG = CALL_EXCEL_SOLVER_FUNC(TARGET_RNG, CHG_CELLS_RNG, _
     REFER_CELLS_RNG(), CONST_CELLS_RNG(), RELATION_ARR(), _
     1, 0, 100, 100, 0.000001, _
     False, False, 1, 1, 1, 5, False, 0.0001, True)
If SOLVER_FLAG = False Then: GoTo ERROR_LABEL

DATA_MATRIX(NTRIALS + 2, 1) = TARGET_PORT_RETURN
DATA_MATRIX(NTRIALS + 2, 2) = CONTROL_RNG.Cells(2, 5).value
DATA_MATRIX(NTRIALS + 2, 3) = CONTROL_RNG.Cells(2, 6).value
DATA_MATRIX(NTRIALS + 2, 4) = CONTROL_RNG.Cells(2, 7).value
DATA_MATRIX(NTRIALS + 2, 5) = Abs(CONTROL_RNG.Cells(2, 8).value)
DATA_MATRIX(NTRIALS + 2, 6) = CONTROL_RNG.Cells(2, 9).value
DATA_MATRIX(NTRIALS + 2, 7) = Abs(CONTROL_RNG.Cells(2, 10).value)

For j = 1 To NCOLUMNS: DATA_MATRIX(NTRIALS + 2, 7 + j) = WEIGHTS_RNG.Cells(1, j): Next j
DELTA_RETURN = (DATA_MATRIX(NTRIALS + 2, 2) - DATA_MATRIX(1, 2)) / (NTRIALS + 1)

'--------------------------------------------------------------------------------
ReDim REFER_CELLS_RNG(1 To 4)
ReDim CONST_CELLS_RNG(1 To 4)
ReDim RELATION_ARR(1 To 4)

Set REFER_CELLS_RNG(1) = WEIGHTS_RNG
Set CONST_CELLS_RNG(1) = CONTROL_RNG.Cells(2, 13)
RELATION_ARR(1) = 3
    
Set REFER_CELLS_RNG(2) = WEIGHTS_RNG
Set CONST_CELLS_RNG(2) = CONTROL_RNG.Cells(2, 14)
RELATION_ARR(2) = 1
    
Set REFER_CELLS_RNG(3) = CONTROL_RNG.Cells(2, 4)
Set CONST_CELLS_RNG(3) = CONTROL_RNG.Cells(2, 12)
RELATION_ARR(3) = 2
    
Set REFER_CELLS_RNG(4) = CONTROL_RNG.Cells(2, 5)
Set CONST_CELLS_RNG(4) = CONTROL_RNG.Cells(2, 3)
RELATION_ARR(4) = 3

'--------------------------------------------------------------------------------
For i = 1 To NTRIALS
'--------------------------------------------------------------------------------
    CONTROL_RNG.Cells(2, 3).value = DATA_MATRIX(1, 2) + i * DELTA_RETURN
    
    If VERSION = 0 Then 'Max Portfolio Returns
        Set TARGET_RNG = CONTROL_RNG.Cells(2, 8)
        Set CHG_CELLS_RNG = WEIGHTS_RNG
    Else 'Min Portfolio Standard Deviation
        Set TARGET_RNG = CONTROL_RNG.Cells(2, 6)
        Set CHG_CELLS_RNG = WEIGHTS_RNG
    End If
    
    SOLVER_FLAG = CALL_EXCEL_SOLVER_FUNC(TARGET_RNG, CHG_CELLS_RNG, _
         REFER_CELLS_RNG(), CONST_CELLS_RNG(), RELATION_ARR(), _
         k, 0, 100, 100, 0.000001, _
         False, False, 1, 1, 1, 5, False, 0.0001, True)
    If SOLVER_FLAG = False Then: GoTo ERROR_LABEL

    DATA_MATRIX(i + 1, 1) = DATA_MATRIX(1, 2) + i * DELTA_RETURN
    DATA_MATRIX(i + 1, 2) = CONTROL_RNG.Cells(2, 5).value
    DATA_MATRIX(i + 1, 3) = CONTROL_RNG.Cells(2, 6).value
    DATA_MATRIX(i + 1, 4) = CONTROL_RNG.Cells(2, 7).value
    DATA_MATRIX(i + 1, 5) = Abs(CONTROL_RNG.Cells(2, 8).value)
    DATA_MATRIX(i + 1, 6) = CONTROL_RNG.Cells(2, 9).value
    DATA_MATRIX(i + 1, 7) = Abs(CONTROL_RNG.Cells(2, 10).value)

    For j = 1 To NCOLUMNS
        DATA_MATRIX(i + 1, 7 + j) = WEIGHTS_RNG.Cells(1, j)
    Next j
'--------------------------------------------------------------------------------
Next i
'--------------------------------------------------------------------------------

CONTROL_RNG.Cells(2, 3).value = DATA_MATRIX(NTRIALS + 1, 1)

Range(SAMPLING_RNG.Cells(0, 1), SAMPLING_RNG.Cells(0, 7)).value = _
Array("Target Return", "Actual Return", "StDev", "VaR", "ETL", "MAD", "Max Loss")
SAMPLING_RNG.value = DATA_MATRIX

Range(SAMPLING_RNG.Cells(0, 8), SAMPLING_RNG.Cells(0, 6 + NCOLUMNS)).FormulaArray = _
"=" & HEADERS_RNG.Address
'On Error Resume Next
GoSub FORMAT_FONTS
GoSub FORMAT_CONTROL_CELLS
'---------------------------------------------------------------------------------
With Excel.Application
    .EnableEvents = True
    .ScreenUpdating = True
    .Cursor = xlDefault
    .StatusBar = False
End With
'---------------------------------------------------------------------------------
RNG_PORT_ETL_FRONTIER_FUNC = True


Exit Function
'--------------------------------------------------------------------------------------
FORMAT_FONTS:
'--------------------------------------------------------------------------------------
    Set TEMP_RNG = Union(CONTROL_RNG, RETURNS_RNG.CurrentRegion, _
                     SAMPLING_RNG.CurrentRegion)
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
        With .Font
            .name = "Courier New"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
        .NumberFormat = "0.00%"
        .RowHeight = 15
        .ColumnWidth = 15
    End With
    RETURNS_RNG.CurrentRegion.Columns(2).NumberFormat = "dd-mmm-yyyy"
'--------------------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------------------
FORMAT_CONTROL_CELLS:
    Set TEMP_RNG = Union(CONTROL_RNG, RETURNS_RNG.CurrentRegion, _
                     SAMPLING_RNG.CurrentRegion)
    With TEMP_RNG
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

    End With
        
    CONTROL_RNG.Rows(1).Interior.Color = 15773696 'Azul
    With SAMPLING_RNG.CurrentRegion
        With .Rows(1)
            With Range(.Cells(1, 1), .Cells(1, 7))
                .Interior.Color = 49407
            End With
            With Range(.Cells(1, 8), .Cells(1, 6 + NCOLUMNS))
                .Interior.Color = 5296274
            End With
        End With
    End With
    With RETURNS_RNG.CurrentRegion
        .Rows(2).Interior.Color = 65535
        With Range(.Cells(1, 1), .Cells(1, 2))
            .Interior.Color = 5296274
        End With
    End With
    
    With CONTROL_RNG
        With .Columns(11)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeTop).LineStyle = xlNone
            .Borders(xlEdgeBottom).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
            With .Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End With
        .Cells(2, 2).NumberFormat = "0"
        Range(.Cells(2, 1), .Cells(2, 3)).Font.Color = -4165632
        Range(.Cells(2, 12), .Cells(2, 14)).Font.Color = -4165632
    End With
    WEIGHTS_RNG.Font.Color = -4165632
    RETURNS_RNG.Font.Color = -4165632
    RETURNS_RNG.Columns(1).Offset(0, -1).Font.Color = -4165632
    SAMPLING_RNG.Font.Color = -4165632
    HEADERS_RNG.Font.Color = -16776961
'--------------------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------------------
ERROR_LABEL:
'---------------------------------------------------------------------------------
With Excel.Application
    .EnableEvents = True
    .ScreenUpdating = True
    .Cursor = xlDefault
    .StatusBar = False
End With
'---------------------------------------------------------------------------------
RNG_PORT_ETL_FRONTIER_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_EXPECTED_TAIL_LOSS_FUNC
'DESCRIPTION   : Port Expected-Loss
'LIBRARY       : PORTFOLIO
'GROUP         : FRONTIER_ETL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_EXPECTED_TAIL_LOSS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef WEIGHTS_RNG As Variant, _
Optional ByVal CONFIDENCE_VAL As Double = 0.975, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

'CONFIDENCE: Used in calculation of expected loss

Dim i As Long
Dim j As Long

Dim VAR_VAL As Double
Dim ALPHA_VAL As Double
Dim TEMP_SUM As Double

Dim DATA_MATRIX As Variant
Dim WEIGHTS_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If IsArray(WEIGHTS_RNG) = True Then
    DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)
    WEIGHTS_VECTOR = WEIGHTS_RNG
    If UBound(WEIGHTS_VECTOR, 2) = 1 Then
        WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
    End If
    If UBound(WEIGHTS_VECTOR, 2) <> UBound(DATA_MATRIX, 2) Then: GoTo ERROR_LABEL
    
    DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(MMULT_FUNC(WEIGHTS_VECTOR, MATRIX_TRANSPOSE_FUNC(DATA_MATRIX), 70))
End If

ALPHA_VAL = 1 - CONFIDENCE_VAL
VAR_VAL = HISTOGRAM_PERCENTILE_FUNC(DATA_MATRIX, ALPHA_VAL, 1)
If OUTPUT <> 0 Then
    PORT_EXPECTED_TAIL_LOSS_FUNC = VAR_VAL
    Exit Function
End If
TEMP_SUM = 0
j = 0
For i = 1 To UBound(DATA_MATRIX, 1)
    If DATA_MATRIX(i, 1) <= VAR_VAL Then
        j = j + 1
        TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, 1)
    End If
Next i
PORT_EXPECTED_TAIL_LOSS_FUNC = TEMP_SUM / j

Exit Function
ERROR_LABEL:
PORT_EXPECTED_TAIL_LOSS_FUNC = Err.number
End Function
