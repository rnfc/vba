Attribute VB_Name = "WEB_SERVICE_YAHOO_FORMS_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private PUB_YAHOO_STREAMING_RUN_VAL As Double
'Private PUB_YAHOO_STREAMING_ADDRESS_STR As String
Private Const PUB_YAHOO_STREAMING_SECONDS_VAL = 5
Private Const PUB_YAHOO_STREAMING_RNG_STR = "YAHOO_STREAMING_RND_RNG"
Private Const PUB_YAHOO_STREAMING_FUNC_NAME = "YAHOO_STREAMING_FUNC"

'Streaming
Private Sub YAHOO_STREAMING_START_TIMER_FUNC()
PUB_YAHOO_STREAMING_RUN_VAL = Now + TimeSerial(0, 0, PUB_YAHOO_STREAMING_SECONDS_VAL)
Excel.Application.OnTime earliesttime:=PUB_YAHOO_STREAMING_RUN_VAL, Procedure:=PUB_YAHOO_STREAMING_FUNC_NAME, schedule:=True
End Sub
Private Sub YAHOO_STREAMING_STOP_TIMER_FUNC()
On Error Resume Next
Excel.Application.OnTime earliesttime:=PUB_YAHOO_STREAMING_RUN_VAL, Procedure:=PUB_YAHOO_STREAMING_FUNC_NAME, schedule:=False
ThisWorkbook.Names(PUB_YAHOO_STREAMING_RNG_STR).Delete
Debug.Print ThisWorkbook.Names.COUNT
'   Debug.Print EXCEL_MACRO_DELETE_HIDDEN_FUNC(PUB_YAHOO_STREAMING_RNG_STR)
'   Debug.Print EXCEL_MACRO_CHECK_HIDDEN_FUNC(PUB_YAHOO_STREAMING_RNG_STR)
End Sub
Private Sub YAHOO_STREAMING_FUNC()
ThisWorkbook.Names.Add PUB_YAHOO_STREAMING_RNG_STR, CStr(Rnd) 'Name goes to the black box
'   ActiveSheet.Cells(.Cells.Rows.Count, .Cells.Columns.Count).Value = Rnd
'   Debug.Print EXCEL_MACRO_VALUE_HIDDEN_FUNC(PUB_YAHOO_STREAMING_RNG_STR, CStr(Rnd), True)
'   Debug.Print EXCEL_MACRO_CHECK_HIDDEN_FUNC(PUB_YAHOO_STREAMING_RNG_STR)
Call YAHOO_STREAMING_START_TIMER_FUNC
End Sub

Public Function SHOW_YAHOO_HISTORICAL_DATA_FORM_FUNC()
frmYahoo.show
End Function

Public Function SHOW_YAHOO_QUOTES_FORM_FUNC()

Dim NROWS As Variant
Dim NCOLUMNS As Variant

On Error Resume Next

NROWS = Excel.Application.InputBox("No. Tickers", "Yahoo Finance")
If NROWS < 2 Or NROWS = "" Then: NROWS = 2

NCOLUMNS = Excel.Application.InputBox("No. Elements", "Yahoo Finance")
If NCOLUMNS < 5 Or NCOLUMNS = "" Then: NCOLUMNS = 5

Call EXCEL_TURN_OFF_EVENTS_FUNC
Call EXCEL_DISPLAY_ALERTS_FUNC(False)

Call WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC(), ActiveSheet.Parent)
Call RNG_YAHOO_QUOTES_FUNC(Cells(5, 2), 0, CInt(NROWS), _
CInt(NCOLUMNS), ActiveSheet.Parent)

Call EXCEL_TURN_ON_EVENTS_FUNC
Call EXCEL_DISPLAY_ALERTS_FUNC(True)

End Function
Public Function SHOW_YAHOO_INDEX_QUOTES_FORM_FUNC()

Dim NROWS As Variant
Dim NCOLUMNS As Variant

On Error Resume Next

NROWS = Excel.Application.InputBox("No. Tickers", "Yahoo Finance")
If NROWS < 5 Or NROWS = "" Or NROWS > 500 Then: NROWS = 100

NCOLUMNS = Excel.Application.InputBox("No. Elements", "Yahoo Finance")
If NCOLUMNS < 5 Or NCOLUMNS = "" Or NCOLUMNS > 52 Then: NCOLUMNS = 52

Call EXCEL_TURN_OFF_EVENTS_FUNC
Call EXCEL_DISPLAY_ALERTS_FUNC(False)

Call WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC(), ActiveSheet.Parent)
Call RNG_YAHOO_QUOTES_FUNC(Cells(5, 2), 1, CInt(NROWS), _
CInt(NCOLUMNS), ActiveSheet.Parent)

Call EXCEL_TURN_ON_EVENTS_FUNC
Call EXCEL_DISPLAY_ALERTS_FUNC(True)

End Function
Public Function SHOW_YAHOO_FX_QUOTES_FORM_FUNC()

Dim NROWS As Variant

On Error Resume Next

NROWS = Excel.Application.InputBox("No. Pairs", "Yahoo Finance")
If NROWS < 2 Or NROWS = "" Or NROWS > 100 Then: NROWS = 50

Call EXCEL_TURN_OFF_EVENTS_FUNC
Call EXCEL_DISPLAY_ALERTS_FUNC(False)

Call WSHEET_ADD_FUNC(PARSE_CURRENT_TIME_FUNC(), ActiveSheet.Parent)
Call RNG_YAHOO_FX_QUOTES_FUNC(Cells(5, 2), CInt(NROWS), ActiveSheet.Parent)

Call EXCEL_TURN_ON_EVENTS_FUNC
Call EXCEL_DISPLAY_ALERTS_FUNC(True)

End Function


Function RNG_YAHOO_QUOTES_FUNC(ByRef DST_RNG As Excel.Range, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal NROWS As Integer = 100, _
Optional ByVal NCOLUMNS As Integer = 10, _
Optional ByVal DST_WBOOK As Excel.Workbook)

'VERSION: 0/QUOTES, 1/INDEX
'NROWS --> No. Assets
'NCOLUMNS --> No. Elements

Dim i As Integer
Dim j As Integer

Dim ELEMENTS_ARR As Variant
Dim ELEMENTS_RNG As Excel.Range
Dim TEMP_RNG As Excel.Range

Dim DST_SHAPE As Excel.Shape

On Error GoTo ERROR_LABEL

RNG_YAHOO_QUOTES_FUNC = False

If NCOLUMNS <= 5 Then: NCOLUMNS = 5
If NCOLUMNS > 53 Then: NCOLUMNS = 52
If DST_WBOOK Is Nothing Then: Set DST_WBOOK = ActiveWorkbook

ELEMENTS_ARR = Split(PUB_YAHOO_QUOTES_CODES_STR, ",")
Set ELEMENTS_RNG = DST_RNG.Cells(1, NCOLUMNS + 3)
j = 1
For i = LBound(ELEMENTS_ARR) To UBound(ELEMENTS_ARR) - 1 Step 2
    ELEMENTS_RNG.Cells(j, 1) = ELEMENTS_ARR(i + 1)
    j = j + 1
Next i

Set ELEMENTS_RNG = Range(ELEMENTS_RNG.Cells(1, 1), ELEMENTS_RNG.Cells(j - 1, 1))
Set DST_RNG = Range(DST_RNG, DST_RNG.Cells(NROWS + 1, NCOLUMNS + 1))

'PUB_YAHOO_STREAMING_ADDRESS_STR = Range(DST_RNG.Cells(2, 1), DST_RNG.Cells(NROWS + 1, 1)).Address

With DST_RNG
    If VERSION = 0 Then
        For i = 1 To NROWS
            .Cells(i + 1, 1).value = "MSFT"
        Next i
    Else
        .Cells(1, 1).value = "name"
    End If
    
    j = LBound(ELEMENTS_ARR) + 1
    For i = 1 To NCOLUMNS
        .Cells(1, i + 1).value = ELEMENTS_ARR(j)
        j = j + 2
    Next i
    
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
    .RowHeight = 15
    .ColumnWidth = 20

    Set TEMP_RNG = DST_RNG
    GoSub FORM_BOX_LINE

    With .Columns(1)
        .Font.ColorIndex = 5
        With .Interior
            .ColorIndex = 36
            .Pattern = xlSolid
        End With
    End With

    With .Rows(1)
        .Font.ColorIndex = 3
        With .Interior
            .ColorIndex = 35
            .Pattern = xlSolid
        End With
        With .Validation
            .Delete
            .Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="=" & ELEMENTS_RNG.Address
        End With
    End With
    .Cells(1, 1).Validation.Delete
End With


'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
If VERSION = 0 Then
    Set TEMP_RNG = Range(DST_RNG.Cells(-2, 1), DST_RNG.Cells(-1, 2))
Else
    Set TEMP_RNG = Range(DST_RNG.Cells(-3, 1), DST_RNG.Cells(-1, 2))
End If

TEMP_RNG.HorizontalAlignment = xlCenter
TEMP_RNG.VerticalAlignment = xlBottom
GoSub FORM_BOX_LINE
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

With TEMP_RNG
    With .Columns(1)
        With .Cells(1, 1)
            .value = "SERVER"
        End With
        With .Cells(1, 2)
            .value = "UNITED STATES"
            .Font.ColorIndex = 3
            With .Validation
                .Delete
                .Add Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:="ARGENTINA,AUSTRALIA,BRAZIL,CANADA,FRANCE," & _
                "DENMARK,HONG KONG,INDIA,ITALY,KOREA,SINGAPORE,SPAIN," & _
                "UNITED KINGDOM, UNITED STATES"
            End With
        End With
        With .Cells(2, 1)
            .value = "CONNECTION"
        End With
        With .Cells(2, 2)
            .value = False
            .Font.ColorIndex = 3
            With .Validation
                .Delete
                .Add Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:="TRUE,FALSE"
            End With
        End With
        
        If VERSION = 1 Then
            With .Cells(3, 1)
                .value = "SYMBOL"
            End With
            With .Cells(3, 2)
                .value = "^DJI"
                .Font.ColorIndex = 3
            End With
        End If
        With .Interior
            .ColorIndex = 35
            .Pattern = xlSolid
        End With
    End With
    With .Columns(2)
        With .Interior
            .ColorIndex = 36
            .Pattern = xlSolid
        End With
    End With
End With

Set TEMP_RNG = TEMP_RNG.Cells(1, 4)
For i = 1 To 2
    Set DST_SHAPE = _
    DST_RNG.Worksheet.Shapes.AddShape(15, TEMP_RNG.Left, TEMP_RNG.Top, _
        TEMP_RNG.ColumnWidth * 8, TEMP_RNG.RowHeight * 2)
    DST_SHAPE.Select
    With Selection
        If i = 1 Then
            DST_SHAPE.OnAction = "YAHOO_STREAMING_START_TIMER_FUNC"
            .Characters.Text = "<START-STREAMING>"
            .Font.Color = 0
        ElseIf i = 2 Then
            DST_SHAPE.OnAction = "YAHOO_STREAMING_STOP_TIMER_FUNC"
            .Characters.Text = "<STOP-STREAMING>"
            .Font.Color = 0
        End If
    End With
    DST_SHAPE.Fill.TwoColorGradient msoGradientHorizontal, 1
    DST_SHAPE.Fill.ForeColor.SchemeColor = 5
    Set TEMP_RNG = TEMP_RNG.Cells(1, 3)
Next i

With DST_RNG
    If VERSION = 0 Then
        Range(.Cells(2, 2), _
          .Cells(1 + NROWS, 1 + NCOLUMNS)).FormulaArray = _
        "=IF(" & .Cells(-1, 2).Address & "=FALSE," & """" & "-" & """" & "," & _
            "YAHOO_QUOTES_FUNC(" & _
            Range(.Cells(2, 1), .Cells(1 + NROWS, 1)).Address & "," & _
            Range(.Cells(1, 2), .Cells(1, 1 + NCOLUMNS)).Address & _
            ",Now(),FALSE," & .Cells(-2, 2).Address & "))"
    Else
        Range(.Cells(2, 1), _
              .Cells(1 + NROWS, 1 + NCOLUMNS)).FormulaArray = _
            "=IF(" & .Cells(-2, 2).Address & "=FALSE," & """" & "-" & """" & "," & _
                "YAHOO_INDEX_QUOTES_FUNC(" & _
                .Cells(1, 1).Offset(-2, 1).Address & "," & _
                Range(.Cells(1, 2), .Cells(1, 1 + NCOLUMNS)).Address & _
                ",,FALSE,Now()," & .Cells(-3, 2).Address & "))"
    End If
    .Cells(1, 1).Select
End With

RNG_YAHOO_QUOTES_FUNC = True

Exit Function

'-------------------------------------------------------------------------------------
FORM_BOX_LINE:
'-------------------------------------------------------------------------------------
    With TEMP_RNG
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With
    End With
'-------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
RNG_YAHOO_QUOTES_FUNC = False
End Function


Function RNG_YAHOO_FX_QUOTES_FUNC(ByRef DST_RNG As Excel.Range, _
Optional ByVal NROWS As Integer = 100, _
Optional ByVal DST_WBOOK As Excel.Workbook)

'NROWS --> No. Assets
'5 --> No. Elements

Dim i As Integer
Dim j As Integer

Dim ELEMENTS_ARR As Variant
Dim ELEMENTS_RNG As Excel.Range
Dim TEMP_RNG As Excel.Range

Dim DST_SHAPE As Excel.Shape

On Error GoTo ERROR_LABEL

If NROWS > 100 Then: NROWS = 100
RNG_YAHOO_FX_QUOTES_FUNC = False

If DST_WBOOK Is Nothing Then: Set DST_WBOOK = ActiveWorkbook

'Debug.Print PUB_YAHOO_FX_CODES_STR
If PUB_YAHOO_FX_CODES_STR = "" Then: Call YAHOO_FX_CODES_DESCRIPTION_FUNC("USD")
ELEMENTS_ARR = ARRAY_SHELL_SORT_FUNC(Split(PUB_YAHOO_FX_CODES_STR, ","))

Set ELEMENTS_RNG = DST_RNG.Cells(1, 10)
j = 1
For i = LBound(ELEMENTS_ARR) + 1 To UBound(ELEMENTS_ARR) 'Skip Blank
    ELEMENTS_RNG.Cells(j, 1) = ELEMENTS_ARR(i)
    j = j + 1
Next i

Set ELEMENTS_RNG = Range(ELEMENTS_RNG.Cells(1, 1), ELEMENTS_RNG.Cells(j - 1, 1))
Set DST_RNG = Range(DST_RNG.Cells(1, 1), DST_RNG.Cells(NROWS + 1, 6 + 2))
'PUB_YAHOO_STREAMING_ADDRESS_STR = Range(DST_RNG.Cells(2, 3), DST_RNG.Cells(NROWS + 1, 3)).Address

With DST_RNG
    .Cells(1, 1).value = "Base Currency"
    .Cells(1, 2).value = "Quote Currency"

    For i = 1 To NROWS
        .Cells(i + 1, 1).value = "United States Dollar (USD)"
        .Cells(i + 1, 2).value = ELEMENTS_ARR(LBound(ELEMENTS_ARR) + i + 1) 'Skip Blank
    Next i
            
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
    .RowHeight = 15
    .ColumnWidth = 20

    Set TEMP_RNG = DST_RNG
    GoSub FORM_BOX_LINE

    With Range(.Cells(2, 1), .Cells(NROWS + 1, 2))
        .Font.ColorIndex = 5
        With .Interior
            .ColorIndex = 36
            .Pattern = xlSolid
        End With
        With .Validation
            .Delete
            .Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:="=" & ELEMENTS_RNG.Address
        End With
    End With
    With .Rows(1)
        .Font.ColorIndex = 3
        With .Interior
            .ColorIndex = 35
            .Pattern = xlSolid
        End With
    End With
End With

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
Set TEMP_RNG = Range(DST_RNG.Cells(-1, 1), DST_RNG.Cells(-1, 2))
TEMP_RNG.HorizontalAlignment = xlCenter
TEMP_RNG.VerticalAlignment = xlBottom
GoSub FORM_BOX_LINE
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

With TEMP_RNG
    With .Columns(1)
        With .Cells(1, 1)
            .value = "CONNECTION"
        End With
        With .Cells(1, 2)
            .value = False
            .Font.ColorIndex = 3
            With .Validation
                .Delete
                .Add Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:="TRUE,FALSE"
            End With
        End With
        With .Interior
            .ColorIndex = 35
            .Pattern = xlSolid
        End With
    End With
    With .Columns(2)
        With .Interior
            .ColorIndex = 36
            .Pattern = xlSolid
        End With
    End With
End With

Set TEMP_RNG = TEMP_RNG.Cells(0, 4)
For i = 1 To 2
    Set DST_SHAPE = _
    DST_RNG.Worksheet.Shapes.AddShape(15, TEMP_RNG.Left, TEMP_RNG.Top, _
        TEMP_RNG.ColumnWidth * 8, TEMP_RNG.RowHeight * 2)
    DST_SHAPE.Select
    With Selection
        If i = 1 Then
            DST_SHAPE.OnAction = "YAHOO_STREAMING_START_TIMER_FUNC"
            .Characters.Text = "<START-STREAMING>"
            .Font.Color = 0
        ElseIf i = 2 Then
            DST_SHAPE.OnAction = "YAHOO_STREAMING_STOP_TIMER_FUNC"
            .Characters.Text = "<STOP-STREAMING>"
            .Font.Color = 0
        End If
    End With
    DST_SHAPE.Fill.TwoColorGradient msoGradientHorizontal, 1
    DST_SHAPE.Fill.ForeColor.SchemeColor = 5
    Set TEMP_RNG = TEMP_RNG.Cells(1, 3)
Next i

With DST_RNG
    Range(.Cells(1, 3), _
          .Cells(1 + NROWS, 8)).FormulaArray = _
        "=IF(" & .Cells(-1, 2).Address & _
        "=FALSE," & """" & "-" & """" & "," & _
            "YAHOO_FX_QUOTES_FUNC(" & _
            Range(.Cells(2, 1), .Cells(1 + NROWS, 1)).Address & "," & _
            Range(.Cells(2, 2), .Cells(1 + NROWS, 2)).Address & _
            ",Now(),TRUE))"
    .Cells(1, 1).Select
    '.Columns(8).NumberFormat = "mmm dd, yyyy"
End With

RNG_YAHOO_FX_QUOTES_FUNC = True

Exit Function

'-------------------------------------------------------------------------------------
FORM_BOX_LINE:
'-------------------------------------------------------------------------------------
    With TEMP_RNG
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .WEIGHT = xlThin
            .ColorIndex = xlAutomatic
        End With
    End With
'-------------------------------------------------------------------------------------
Return
'-------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
RNG_YAHOO_FX_QUOTES_FUNC = False
End Function

Function PRINT_YAHOO_HISTORICAL_DATA_SERIES_FUNC(ByRef DST_RNG As Excel.Range, _
ByRef TICKERS_RNG As Variant, _
ByRef START_DATE As Date, _
ByRef END_DATE As Date, _
Optional ByVal PERIOD_STR As String = "Daily", _
Optional ByVal ELEMENT_STR As String = "DOHLCVA", _
Optional ByVal HEADERS_FLAG As Boolean = False, _
Optional ByVal ADJUST_FLAG As Boolean = False, _
Optional ByVal RESORT_FLAG As Boolean = True, _
Optional ByVal VALID_FLAG As Boolean = False)

Dim SROW As Integer
Dim NROWS As Integer

Dim RNG_ARR() As Excel.Range
Dim TEMP_FLAG As Boolean

Dim SCOLUMN As Integer
Dim NCOLUMNS As Integer
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PRINT_YAHOO_HISTORICAL_DATA_SERIES_FUNC = False

PERIOD_STR = YAHOO_HISTORICAL_DATA_PERIOD_STRING_FUNC(PERIOD_STR)
ELEMENT_STR = Right(YAHOO_HISTORICAL_DATA_ELEMENT_STRING_FUNC(ELEMENT_STR), 2)

TEMP_MATRIX = YAHOO_HISTORICAL_DATA_SERIES2_FUNC(TICKERS_RNG, _
              START_DATE, END_DATE, PERIOD_STR, ELEMENT_STR, _
              HEADERS_FLAG, ADJUST_FLAG, RESORT_FLAG)

If IsArray(TEMP_MATRIX) = False Then: GoTo ERROR_LABEL

SROW = LBound(TEMP_MATRIX, 1)
NROWS = UBound(TEMP_MATRIX, 1)
SCOLUMN = LBound(TEMP_MATRIX, 2)
NCOLUMNS = UBound(TEMP_MATRIX, 2)

Set DST_RNG = Range(DST_RNG.Cells(SROW, SCOLUMN), _
                    DST_RNG.Cells(NROWS, NCOLUMNS))

DST_RNG.value = TEMP_MATRIX

Select Case HEADERS_FLAG
Case False
    Call RNG_FORMAT_DATES_VECTOR_FUNC(DST_RNG, False, 0)
    Set DST_RNG = RNG_RESIZE_RNG_FUNC(DST_RNG.Offset(0, 1), 0, 1)
Case True
    Call RNG_FORMAT_DATES_VECTOR_FUNC(DST_RNG, True, 0)
        If VALID_FLAG = True Then
            Excel.Application.ScreenUpdating = True
            TEMP_FLAG = RNG_FILL_SET_ARR_FUNC(DST_RNG, RNG_ARR())
            If TEMP_FLAG = False Then: GoTo 1983
            Call RNG_CHECK_BLANKS_FUNC(RNG_ARR(), 2, 1, 2, 1, 2)
        End If
1983:
    Set DST_RNG = RNG_RESIZE_RNG_FUNC(DST_RNG.Offset(1, 1), 1, 1)
End Select

PRINT_YAHOO_HISTORICAL_DATA_SERIES_FUNC = True

Exit Function
ERROR_LABEL:
PRINT_YAHOO_HISTORICAL_DATA_SERIES_FUNC = False
End Function

                    
Function RNG_YAHOO_HISTORICAL_DATA_SERIES_FUNC(ByVal TICKER_STR As String, _
ByVal START_DATE As Variant, _
ByVal END_DATE As Variant, _
Optional ByRef PERIOD_STR As String = "d", _
Optional ByRef ELEMENT_STR As String = "DOHLCVA", _
Optional ByVal HEADERS_FLAG As Boolean = False, _
Optional ByVal ADJUST_FLAG As Boolean = False, _
Optional ByVal RESORT_FLAG As Boolean = True)

Dim NROWS As Integer
Dim NCOLUMNS As Integer

Dim START_YEAR_INT As Integer
Dim START_MONTH_INT As Integer
Dim START_DAY_INT As Integer
Dim END_YEAR_INT As Integer
Dim END_MONTH_INT As Integer
Dim END_DAY_INT As Integer

On Error GoTo ERROR_LABEL

START_YEAR_INT = Year(START_DATE)
START_MONTH_INT = Month(START_DATE)
START_DAY_INT = Day(START_DATE)

END_YEAR_INT = Year(END_DATE)
END_MONTH_INT = Month(END_DATE)
END_DAY_INT = Day(END_DATE)

NROWS = Excel.Application.Caller.Rows.COUNT
NCOLUMNS = Excel.Application.Caller.Columns.COUNT

RNG_YAHOO_HISTORICAL_DATA_SERIES_FUNC = _
YAHOO_HISTORICAL_DATA_MATRIX_FUNC(TICKER_STR, START_YEAR_INT, START_MONTH_INT, _
START_DAY_INT, END_YEAR_INT, END_MONTH_INT, END_DAY_INT, PERIOD_STR, _
ELEMENT_STR, HEADERS_FLAG, ADJUST_FLAG, RESORT_FLAG, NROWS, NCOLUMNS)

Exit Function
ERROR_LABEL:
RNG_YAHOO_HISTORICAL_DATA_SERIES_FUNC = Err.number
End Function

