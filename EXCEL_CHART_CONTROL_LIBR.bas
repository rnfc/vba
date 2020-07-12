Attribute VB_Name = "EXCEL_CHART_CONTROL_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_CONTROL_CREATE_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         :
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

'make_control_chart

Function EXCEL_CHART_CONTROL_CREATE_FUNC(ByRef DATA_RNG As Excel.Range, _
Optional ByRef LABELS_RNG As Excel.Range, _
Optional ByRef DST_WSHEET As Excel.Worksheet)
'Debug.Print EXCEL_CHART_CONTROL_CREATE_FUNC(Range("NICO"), Range("NICO").Offset(0, -1))

'DATA_RNG: Please select the range containing the DATA POINTS

Dim i As Integer
Dim j As Integer

Dim DATA_STR As String
Dim MEAN_STR As String
Dim LABEL_STR As String

Dim LABEL_FLAG As Boolean
Dim NAME_OBJ As Excel.name
Dim SERIES1_OBJ As Excel.Series
Dim SERIES2_OBJ As Excel.Series
'Dim SERIES3_OBJ As Excel.Series
Dim CHART_OBJ As Excel.ChartObject

On Error GoTo ERROR_LABEL 'GOTO THE END OF THE PROGRAM

EXCEL_CHART_CONTROL_CREATE_FUNC = False

If DST_WSHEET Is Nothing Then: Set DST_WSHEET = ActiveSheet
LABEL_FLAG = True   ' True to re-activate the input range
If IsArray(LABELS_RNG) = False Then: LABEL_FLAG = False

'LETS CREATE THE CHART NOW
Set CHART_OBJ = DST_WSHEET.ChartObjects.Add(Left:=100, Width:=400, Top:=25, Height:=300)
CHART_OBJ.Chart.ChartType = xlLineMarkers

'REMOVE ALL UNWANTED SERIES FROM CHART, IF ANY
For Each SERIES2_OBJ In CHART_OBJ.Chart.SeriesCollection ' CHART_OBJ.Chart.SeriesCollection
    SERIES2_OBJ.Delete
Next SERIES2_OBJ
Set SERIES2_OBJ = Nothing

If LABEL_FLAG Then 'IF WE HAVE THE LABEL RANGE
    'ADD NEW SERIES
    Set SERIES2_OBJ = CHART_OBJ.Chart.SeriesCollection.NewSeries
    With SERIES2_OBJ
        .name = "PLOT"
        .Values = DATA_RNG
        .XValues = LABELS_RNG
    End With
Else
    Set SERIES2_OBJ = CHART_OBJ.Chart.SeriesCollection.NewSeries
    With SERIES2_OBJ
        .name = "PLOT"
        .Values = DATA_RNG
    End With
End If

'FORMAT THE PLOT SERIES
Set SERIES1_OBJ = SERIES2_OBJ
With SERIES2_OBJ
    .Border.ColorIndex = 1
    .MarkerBackgroundColorIndex = 2
    .MarkerForegroundColorIndex = xlAutomatic
    .MarkerStyle = xlCircle
    .Smooth = False
    .MarkerSize = 5
    .Shadow = False
End With
Set SERIES2_OBJ = Nothing

'CREATE NAMED RANGE FOR THE DATA VALUES, AVERAGE, LOWER AND UPPER CONTROL LIMITS
DATA_STR = Excel.Application.WorksheetFunction.Substitute(CHART_OBJ.name, " ", "") & "_DATA_RNG"
MEAN_STR = "roundup(average(" & Excel.Application.WorksheetFunction.Substitute(CHART_OBJ.name, " ", "") & "_DATA_RNG" & "),2)"

With DST_WSHEET.Parent
    .Names.Add name:=Excel.Application.WorksheetFunction.Substitute(CHART_OBJ.name, " ", "") & "_DATA_RNG", RefersToR1C1:=DATA_RNG
    .Names.Add name:=Excel.Application.WorksheetFunction.Substitute(CHART_OBJ.name, " ", "") & "_AVG", RefersToR1C1:="=" & MEAN_STR & ""
    .Names.Add name:=Excel.Application.WorksheetFunction.Substitute(CHART_OBJ.name, " ", "") & "_LCL1", RefersToR1C1:="=" & MEAN_STR & "- roundup(1*stdev(" & DATA_STR & "),2)"
    .Names.Add name:=Excel.Application.WorksheetFunction.Substitute(CHART_OBJ.name, " ", "") & "_LCL2", RefersToR1C1:="=" & MEAN_STR & "- roundup(2*stdev(" & DATA_STR & "),2)"
    .Names.Add name:=Excel.Application.WorksheetFunction.Substitute(CHART_OBJ.name, " ", "") & "_LCL3", RefersToR1C1:="=" & MEAN_STR & "- roundup(3*stdev(" & DATA_STR & "),2)"
    .Names.Add name:=Excel.Application.WorksheetFunction.Substitute(CHART_OBJ.name, " ", "") & "_UCL1", RefersToR1C1:="=" & MEAN_STR & "+ roundup(1*stdev(" & DATA_STR & "),2)"
    .Names.Add name:=Excel.Application.WorksheetFunction.Substitute(CHART_OBJ.name, " ", "") & "_UCL2", RefersToR1C1:="=" & MEAN_STR & "+ roundup(2*stdev(" & DATA_STR & "),2)"
    .Names.Add name:=Excel.Application.WorksheetFunction.Substitute(CHART_OBJ.name, " ", "") & "_UCL3", RefersToR1C1:="=" & MEAN_STR & "+ roundup(3*stdev(" & DATA_STR & "),2)"
End With

'ADD THE LINE FOR AVERAGE
Set SERIES2_OBJ = CHART_OBJ.Chart.SeriesCollection.NewSeries

With SERIES2_OBJ
    .name = "AVG = "
    .Values = "='" & DST_WSHEET.name & "'!" & Excel.Application.WorksheetFunction.Substitute(CHART_OBJ.name, " ", "") & "_AVG"
    .ChartType = xlXYScatter
    '.ErrorBar Direction:=xlX, Include:=xlNone, Type:=xlFixedValue, Amount:=10000
    '.ErrorBar Direction:=xlX, Include:=xlUp, Type:=xlFixedValue, Amount:=20
    .ErrorBar direction:=xlX, Include:=xlPlusValues, Type:=xlFixedValue, amount:=DATA_RNG.Cells.COUNT
    .MarkerBackgroundColorIndex = xlAutomatic
    .MarkerForegroundColorIndex = xlAutomatic
    .MarkerStyle = xlNone
    .Smooth = False
    .MarkerSize = 5
    .Shadow = False
    With .Border
        .WEIGHT = xlHairline
        .LineStyle = xlNone
    End With
    'With .ErrorBars.Border
    '    .LineStyle = xlContinuous
    '    .ColorIndex = 3
    '    .Weight = xlThin
    'End With
End With

Set SERIES2_OBJ = Nothing

'ADD UPPER AND LOWER CONTROL LIMITS
For i = 1 To 3
    For j = -1 To 1 Step 2
        Select Case j:
        Case -1
            LABEL_STR = "LCL"
        Case 1
            LABEL_STR = "UCL"
        End Select
        
        Set SERIES2_OBJ = CHART_OBJ.Chart.SeriesCollection.NewSeries
        With SERIES2_OBJ
            .name = LABEL_STR & i & " ="
            .Values = "='" & DST_WSHEET.name & "'!" & Excel.Application.WorksheetFunction.Substitute(CHART_OBJ.name, " ", "") & "_" & LABEL_STR & i
            .ChartType = xlXYScatter
            .ErrorBar direction:=xlX, Include:=xlPlusValues, Type:=xlFixedValue, amount:=DATA_RNG.Cells.COUNT
        End With
        
        SERIES2_OBJ.ErrorBar direction:=xlX, Include:=xlPlusValues, Type:=xlFixedValue, amount:=DATA_RNG.Cells.COUNT
        Select Case i
        Case 1
            With SERIES2_OBJ.ErrorBars.Border
            .LineStyle = xlGray25
            .ColorIndex = 15
            .WEIGHT = xlHairline
            End With
        Case 2
            With SERIES2_OBJ.ErrorBars.Border
            .LineStyle = xlGray25
            .ColorIndex = 57
            .WEIGHT = xlHairline
            End With
        Case 3
            With SERIES2_OBJ.ErrorBars.Border
            .LineStyle = xlGray75
            .ColorIndex = 3
            .WEIGHT = xlHairline
            End With
        End Select
    
        SERIES2_OBJ.ErrorBars.EndStyle = xlNoCap
    
        With SERIES2_OBJ
            With .Border
                .WEIGHT = xlHairline
                .LineStyle = xlNone
            End With
            .MarkerBackgroundColorIndex = xlAutomatic
            .MarkerForegroundColorIndex = xlAutomatic
            .MarkerStyle = xlNone
            .Smooth = False
            .MarkerSize = 5
            .Shadow = False
        End With
        Set SERIES2_OBJ = Nothing
    Next j
Next i

CHART_OBJ.Chart.ApplyDataLabels AutoText:=True, LegendKey:=False, _
HasLeaderLines:=False, ShowSeriesName:=True, ShowCategoryName:=False, _
ShowValue:=True, ShowPercentage:=False, ShowBubbleSize:=False, Separator:=" "

'OFFSET THE LABELS
'If LABEL_FLAG Then
'    For Each SERIES3_OBJ In CHART_OBJ.Chart.SeriesCollection
'        SERIES3_OBJ.Points(1).DataLabel.Left = 400
'    Next SERIES3_OBJ
'End If

'LETS FORMAT THE CHART
With CHART_OBJ
    With .Chart.Axes(xlCategory)
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNextToAxis
    End With
    With .Chart.Axes(xlValue)
        .MajorTickMark = xlOutside
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNextToAxis
    End With
    With .Chart.ChartArea.Border
        .WEIGHT = 1
        .LineStyle = 0
    End With
    With .Chart.PlotArea.Border
        .ColorIndex = 1
        .WEIGHT = xlThin
        .LineStyle = xlContinuous
    End With
    With .Chart.PlotArea.Interior
        .ColorIndex = 2
        .PatternColorIndex = 1
        .Pattern = xlSolid
    End With
    With .Chart.ChartArea.Font
        .name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .Background = xlAutomatic
    End With
    With .Chart
        .HasTitle = False
        .Axes(xlCategory, xlPrimary).HasTitle = False
        .Axes(xlValue, xlPrimary).HasTitle = True
        .HasTitle = True
        .ChartTitle.Characters.Text = "Control Chart"
        .ChartTitle.Left = 134
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Observations"
    End With
    With .Chart.Axes(xlCategory).TickLabels
        .Alignment = xlCenter
        .Offset = 100
        .ReadingOrder = xlContext
        .Orientation = xlHorizontal
    End With
End With


With CHART_OBJ.Chart
    .Legend.Delete
    .PlotArea.Width = 310
    .Axes(xlValue).MajorGridlines.Delete
    .Axes(xlValue).CrossesAt = CHART_OBJ.Chart.Axes(xlValue).MinimumScale
    .ChartArea.Interior.ColorIndex = xlAutomatic
    .ChartArea.AutoScaleFont = True
End With

'DELETE THE LABELS FOR THE ACTUAL DATA SERIES
SERIES1_OBJ.DataLabels.Delete
Set SERIES1_OBJ = Nothing

EXCEL_CHART_CONTROL_CREATE_FUNC = True

Exit Function
ERROR_LABEL:
If Err.number Then
    For Each NAME_OBJ In DST_WSHEET.Parent.Names
        If Left(NAME_OBJ.name, 5) = "Chart" Then: NAME_OBJ.Delete
    Next NAME_OBJ
    EXCEL_CHART_CONTROL_CREATE_FUNC = False
End If
End Function
