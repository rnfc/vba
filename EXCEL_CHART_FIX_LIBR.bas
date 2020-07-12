Attribute VB_Name = "EXCEL_CHART_FIX_LIBR"

'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

Function EXCEL_CHART_FIX_FUNC(ByRef CHART_OBJ As Excel.Chart, _
Optional ByVal MARKER_SIZE_VAL As Integer = 1, _
Optional ByVal FIX_CHART_FONT_SIZE_FLAG As Boolean = False, _
Optional ByVal FIX_CHART_AXIS_FLAG As Boolean = False, _
Optional ByVal FIX_CHART_3D_FLAG As Boolean = False, _
Optional ByVal FIX_CHART_JUNK_FLAG As Boolean = False, _
Optional ByVal FIX_CHART_COLORS_FLAG As Boolean = False)

Dim i As Integer
Dim j As Integer
Dim SERIES_OBJ As Excel.Series

On Error GoTo ERROR_LABEL

With CHART_OBJ
    ' remove 3d from all charts
    If FIX_CHART_3D_FLAG Then
        Select Case .ChartType
        Case xlPyramidBarStacked
            .ChartType = xlBarStacked
        Case xlPyramidCol
            .ChartType = xlColumn
        Case xlPyramidColClustered
            .ChartType = xlColumnClustered
        Case xlPyramidColStacked
            .ChartType = xlColumnStacked
        Case xlPyramidColStacked100
            .ChartType = xlColumnStacked100
        Case xlConeBarStacked
            .ChartType = xlBarStacked
        Case xlConeCol
            .ChartType = xlColumn
        Case -4111 ' cone?
            .ChartType = xlColumnClustered
        Case xlConeColClustered
            .ChartType = xlColumnClustered
        Case xlConeColStacked
            .ChartType = xlColumnStacked
        Case xlConeColStacked100
            .ChartType = xlColumnStacked100
        Case xlSurface
            .ChartType = xlSurfaceTopView
        Case xlXYScatterSmooth
            .ChartType = xlXYScatterLines
        Case xlXYScatterSmoothNoMarkers
            .ChartType = xlXYScatterLinesNoMarkers
        Case xl3DArea
            .ChartType = xlArea
        Case xl3DAreaStacked
            .ChartType = xlAreaStacked
        Case xl3DAreaStacked100
            .ChartType = xlAreaStacked100
        Case xl3DBarClustered
            .ChartType = xlBarClustered
        Case xl3DBarStacked
            .ChartType = xlBarStacked
        Case xl3DBarStacked100
            .ChartType = xlBarStacked100
        Case xl3DColumn
            .ChartType = xlColumnClustered
        Case xl3DColumnClustered
            .ChartType = xlColumnClustered
        Case xl3DColumnStacked
            .ChartType = xlColumnStacked
        Case xl3DColumnStacked100
            .ChartType = xlColumnStacked100
        Case xl3DLine
            .ChartType = xlLine
        Case xl3DPie
            .ChartType = xlPie
        Case xl3DPieExploded
            .ChartType = xlPieExploded
        Case xlBubble3DEffect
            .ChartType = xlBubble
        Case xlConeBarClustered
            .ChartType = xlBar
        Case xlConeBarStacked
            .ChartType = xlBarStacked
        Case xlConeBarStacked100
            .ChartType = xlBarStacked100
        Case xlConeCol
            .ChartType = xlColumn
        Case xlConeColClustered
            .ChartType = xlColumnClustered
        Case xlConeColStacked
            .ChartType = xlColumnStacked
        Case xlConeColStacked100
            .ChartType = xlColumnStacked100
        Case xlCylinderBarClustered
            .ChartType = xlBarClustered
        Case xlCylinderBarStacked
            .ChartType = xlBarStacked
        Case xlCylinderBarStacked100
            .ChartType = xlBarStacked100
        Case xlCylinderCol
            .ChartType = xlColumn
        Case xlCylinderColClustered
            .ChartType = xlColumnClustered
            .ChartType = xlColumnClustered
        Case xlCylinderColStacked
            .ChartType = xlColumnStacked
        Case xlCylinderColStacked100
            .ChartType = xlColumnStacked100
        Case xlPyramidBarClustered
            .ChartType = xlBarClustered
        Case xlPyramidBarStacked100
            .ChartType = xlBarStacked100
        End Select
    End If
    
    Select Case .ChartType
    Case xlPyramidBarStacked
        .ChartType = xlBarStacked
    Case xlPyramidCol
        .ChartType = xlColumn
    Case xlPyramidColClustered
        .ChartType = xlColumnClustered
    Case xlPyramidColStacked
        .ChartType = xlColumnStacked
    Case xlPyramidColStacked100
        .ChartType = xlColumnStacked100
    Case xlSurface
        .ChartType = xlSurfaceTopView
    Case xlXYScatterSmooth
        .ChartType = xlXYScatterLines
    Case xlXYScatterSmoothNoMarkers
        .ChartType = xlXYScatterLinesNoMarkers
    Case xlConeBarClustered
        .ChartType = xlBar
    Case xlConeBarStacked
        .ChartType = xlBarStacked
    Case xlConeBarStacked100
        .ChartType = xlBarStacked100
    Case xlConeCol
        .ChartType = xlColumn
    Case xlConeColClustered
        .ChartType = xlColumnClustered
    Case xlConeColStacked
        .ChartType = xlColumnStacked
    Case xlConeColStacked100
        .ChartType = xlColumnStacked100
    Case xlCylinderBarClustered
        .ChartType = xlBarClustered
    Case xlCylinderBarStacked
        .ChartType = xlBarStacked
    Case xlCylinderBarStacked100
        .ChartType = xlBarStacked100
    Case xlCylinderCol
        .ChartType = xlColumn
    Case xlCylinderColClustered
        .ChartType = xlColumnClustered
    Case xlCylinderColStacked
        .ChartType = xlColumnStacked
    Case xlCylinderColStacked100
        .ChartType = xlColumnStacked100
    Case xlPyramidBarClustered
        .ChartType = xlBarClustered
    Case xlPyramidBarStacked100
        .ChartType = xlBarStacked100
    End Select
End With
'///////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////
If FIX_CHART_JUNK_FLAG Then 'bFixChartJunk
    With CHART_OBJ
        .PlotArea.Border.LineStyle = xlNone
        .PlotArea.Interior.Pattern = 1
        .ChartArea.Interior.Pattern = 1
        .PlotArea.Interior.ColorIndex = 2
        .ChartArea.Interior.ColorIndex = 2
        ' fix the legend
        .Legend.Interior.ColorIndex = 2
        .Legend.Interior.PatternColorIndex = 1
        .Legend.Interior.Pattern = xlSolid
        .Legend.Border.WEIGHT = xlHairline
        .Legend.Border.LineStyle = xlAutomatic
        
        .Axes(xlValue).Border.LineStyle = xlNone
        .Axes(xlValue).MajorGridlines.Border.ColorIndex = 15
        .Axes(xlValue).MajorGridlines.Border.WEIGHT = xlHairline
        .Axes(xlValue).MajorGridlines.Border.LineStyle = xlContinuous
        
        .Axes(xlValue, xlSecondary).Border.LineStyle = xlNone
        .Axes(xlValue, xlSecondary).MajorGridlines.Border.ColorIndex = 15
        .Axes(xlValue, xlSecondary).MajorGridlines.Border.WEIGHT = xlHairline
        .Axes(xlValue, xlSecondary).MajorGridlines.Border.LineStyle = xlContinuous
    End With
End If
'///////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////
If FIX_CHART_AXIS_FLAG Then ' fix the problem with a percentage chart going up to 1.2
    If CHART_OBJ.Axes(xlValue).MaximumScaleIsAuto = True And CHART_OBJ.Axes(xlValue).MaximumScale = 1.2 Then
        CHART_OBJ.Axes(xlValue).MaximumScaleIsAuto = False
        CHART_OBJ.Axes(xlValue).MaximumScale = 1#
    End If
' fix the value axis format
    scalediff = CHART_OBJ.Axes(xlValue).MaximumScale - CHART_OBJ.Axes(xlValue).MinimumScale
    If scalediff > 10000000 Then
        CHART_OBJ.Axes(xlValue).TickLabels.NumberFormat = "#,##0,,""mil"""
    ElseIf scalediff > 100000 Then
        CHART_OBJ.Axes(xlValue).TickLabels.NumberFormat = "#,##0,""k"""
    ElseIf scalediff > 1000 Then
        CHART_OBJ.Axes(xlValue).TickLabels.NumberFormat = "#,##0"
    ElseIf scalediff > 10 Then
        CHART_OBJ.Axes(xlValue).TickLabels.NumberFormat = "0"
    ElseIf scalediff > 0.5 Then
        CHART_OBJ.Axes(xlValue).TickLabels.NumberFormat = "0.0"
    Else
        CHART_OBJ.Axes(xlValue).TickLabels.NumberFormat = "0.00"
    End If
    ' fixes an excel bug where stacked columns show ranges from 0-100
    If CHART_OBJ.ChartType = xlColumnStacked100 Or CHART_OBJ.ChartType = xlBarStacked100 Then
        CHART_OBJ.Axes(xlValue).TickLabels.NumberFormat = "0%"
    End If
End If
'///////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////
If FIX_CHART_FONT_SIZE_FLAG Then ' set the font size
    CHART_OBJ.ChartArea.AutoScaleFont = True
    CHART_OBJ.ChartArea.Font.Size = CInt(2 * MARKER_SIZE_VAL)
    CHART_OBJ.ChartArea.Font.Size = Application.WorksheetFunction.max(6, CHART_OBJ.ChartArea.Font.Size)
    CHART_OBJ.ChartArea.Font.Size = Application.WorksheetFunction.Min(14, CHART_OBJ.ChartArea.Font.Size)
End If
'///////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////

If FIX_CHART_COLORS_FLAG Then
    j = CHART_OBJ.SeriesCollection.COUNT
    Set COLORS_ARR = get_colors_collection(j)
    ' set marker sizes and series colors
    If CHART_OBJ.ChartType = xlLine Or _
        CHART_OBJ.ChartType = xlLineMarkers Or _
        CHART_OBJ.ChartType = xlLineMarkersStacked Or _
        CHART_OBJ.ChartType = xlLineMarkersStacked100 Or _
        CHART_OBJ.ChartType = xlLineStacked Or _
        CHART_OBJ.ChartType = xlLineStacked100 Or _
        CHART_OBJ.ChartType = xlXYScatterLines Or _
        CHART_OBJ.ChartType = xlXYScatter Or _
        CHART_OBJ.ChartType = xlXYScatterLinesNoMarkers Or _
        CHART_OBJ.ChartType = xlXYScatterSmooth Or _
        CHART_OBJ.ChartType = xlXYScatterSmoothNoMarkers Then
        i = 1
        For Each SERIES_OBJ In CHART_OBJ.SeriesCollection ' handle lines
            If SERIES_OBJ.Border.LineStyle <> xlNone Then
                SERIES_OBJ.Border.WEIGHT = xlMedium
                ' only set the first 7 colors
                If i <= 7 Then SERIES_OBJ.Border.ColorIndex = COLORS_ARR(i)
            End If
            ' misc junk
            SERIES_OBJ.Smooth = False
            SERIES_OBJ.Shadow = False
            'SERIES_OBJ.Interior.ColorIndex = xlAutomatic
            SERIES_OBJ.Interior.Pattern = xlSolid
            ' markers
            If SERIES_OBJ.MarkerStyle <> xlNone Then
                If SERIES_OBJ.MarkerStyle = xlPlus Or _
                    SERIES_OBJ.MarkerStyle = xlX Or _
                    SERIES_OBJ.MarkerStyle = xlStar Or _
                    SERIES_OBJ.MarkerStyle = 1 Or _
                    SERIES_OBJ.MarkerStyle = xlMarkerStylePlus Or _
                    SERIES_OBJ.MarkerStyle = xlMarkerStyleX Or _
                    SERIES_OBJ.MarkerStyle = xlMarkerStyleStar Then
                    
                    SERIES_OBJ.MarkerBackgroundColorIndex = 2
                Else
                    'MsgBox SERIES_OBJ.MarkerStyle
                    ' doesn't catch automatic
                    SERIES_OBJ.MarkerBackgroundColorIndex = COLORS_ARR(i)
                End If
                SERIES_OBJ.MarkerForegroundColorIndex = COLORS_ARR(i)
                SERIES_OBJ.MarkerSize = MARKER_SIZE_VAL
            End If
            i = i + 1
        Next SERIES_OBJ
    ' don't change pie charts or doughnut charts
    ElseIf CHART_OBJ.ChartType <> xlPie And _
        CHART_OBJ.ChartType <> xlPieExploded And _
        CHART_OBJ.ChartType <> xlPieOfPie And _
        CHART_OBJ.ChartType <> xlDoughnut And _
        CHART_OBJ.ChartType <> xlDoughnutExploded Then
        i = 1
        For Each SERIES_OBJ In CHART_OBJ.SeriesCollection
            SERIES_OBJ.Border.WEIGHT = xlMedium
            SERIES_OBJ.Smooth = False
            SERIES_OBJ.Shadow = False
            SERIES_OBJ.Interior.ColorIndex = xlAutomatic
            SERIES_OBJ.Interior.Pattern = xlSolid
            ' only set colors on the first 7 series
            If i <= 7 Then
                SERIES_OBJ.Interior.ColorIndex = COLORS_ARR(i)
                SERIES_OBJ.Border.LineStyle = xlNone
            End If
            i = i + 1
        Next SERIES_OBJ
    Else:
    End If
End If

EXCEL_CHART_FIX_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_FIX_FUNC = False
End Function
