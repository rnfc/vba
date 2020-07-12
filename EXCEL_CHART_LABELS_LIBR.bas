Attribute VB_Name = "EXCEL_CHART_LABELS_LIBR"

'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_APPLY_DATA_LABELS_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         :
'ID            :
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_APPLY_DATA_LABELS_FUNC(ByRef SERIE_OBJ As Excel.Series, _
ByRef NAMES_ARR As Variant)
      
'Call EXCEL_CHART_APPLY_DATA_LABELS_FUNC(ActiveSheet.ChartObjects("Chart 7").Chart.SeriesCollection(1), _
'Array("TD", "KA"))
      
Dim i As Long
Dim j As Long

On Error GoTo ERROR_LABEL

EXCEL_CHART_APPLY_DATA_LABELS_FUNC = False

SERIE_OBJ.ApplyDataLabels Type:=xlDataLabelsShowValue, AutoText:=True, _
LegendKey:=False

'SERIE_OBJ.ApplyDataLabels AutoText:=True, LegendKey:=False, _
    HasLeaderLines:=False, ShowSeriesName:=True, ShowCategoryName:=False, _
    ShowValue:=True, ShowPercentage:=False, ShowBubbleSize:=False, Separator:=" "

j = 1
For i = LBound(NAMES_ARR) To UBound(NAMES_ARR)
    SERIE_OBJ.Points(j).DataLabel.Characters.Text = NAMES_ARR(i)
    j = j + 1
Next i

With SERIE_OBJ.DataLabels.Font
     .name = "Helvetica-Narrow"
     .FontStyle = "Regular"
     .Size = 10
     .Strikethrough = False
     .Superscript = False
     .Subscript = False
     .OutlineFont = False
     .Shadow = False
     .Underline = xlNone
     .ColorIndex = xlAutomatic
     .Background = xlAutomatic
End With

EXCEL_CHART_APPLY_DATA_LABELS_FUNC = True
      
Exit Function
ERROR_LABEL:
EXCEL_CHART_APPLY_DATA_LABELS_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_INSERT_DATA_LABELS_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         :
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

'The comparable chart with labels for each series shown on right.
'Debug.Print EXCEL_CHART_INSERT_DATA_LABELS_FUNC(ActiveChart)

Function EXCEL_CHART_INSERT_DATA_LABELS_FUNC(ByRef CHART_OBJ As Excel.Chart)

Dim i As Long
Dim SRC_SERIES As Excel.Series

On Error GoTo ERROR_LABEL

EXCEL_CHART_INSERT_DATA_LABELS_FUNC = False

For Each SRC_SERIES In CHART_OBJ.SeriesCollection
    With SRC_SERIES
        i = .Points.COUNT
        SRC_SERIES.Points(i).ApplyDataLabels Type:=xlDataLabelsShowValue, _
        AutoText:=True, LegendKey:=False
        SRC_SERIES.Points(i).DataLabel.Text = SRC_SERIES.name
        ' MsgBox (SRC_SERIES.Name)
        With SRC_SERIES.DataLabels
            .AutoScaleFont = True
            With .Font
                .name = "Arial"
                .FontStyle = "Bold"
                .Size = 8
            End With
            With .Border
                .WEIGHT = xlHairline
                .LineStyle = xlNone
            End With
            .Shadow = False
            With .Interior
                .ColorIndex = 2
                .PatternColorIndex = 1
                .Pattern = xlSolid
            End With
        End With
    End With
Next SRC_SERIES

EXCEL_CHART_INSERT_DATA_LABELS_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_INSERT_DATA_LABELS_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_OFFSET_DATA_LABELS_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         :
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_OFFSET_DATA_LABELS_FUNC(ByRef DATA_ARR As Variant, _
ByRef CHART_OBJ As Excel.Chart)

Dim i As Long
Dim j As Long '1
Dim LEFT_VAL As Double '400
Dim SERIES_OBJ As Excel.Series

On Error GoTo ERROR_LABEL

EXCEL_CHART_OFFSET_DATA_LABELS_FUNC = True

For Each SERIES_OBJ In CHART_OBJ.SeriesCollection
    For i = LBound(DATA_ARR) To UBound(DATA_ARR)
        j = DATA_ARR(i, 1)
        LEFT_VAL = DATA_ARR(i, 2)
        With SERIES_OBJ.Points(j).DataLabel
            .Left = LEFT_VAL
        End With
    Next i
Next SERIES_OBJ

EXCEL_CHART_OFFSET_DATA_LABELS_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_OFFSET_DATA_LABELS_FUNC = False
End Function
