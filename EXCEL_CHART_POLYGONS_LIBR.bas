Attribute VB_Name = "EXCEL_CHART_POLYGONS_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_POLYGONS_XADD_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         :
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_POLYGONS_XADD_FUNC(ByVal CHART_OBJ As Excel.Chart, _
ByRef PARAM_VECTOR As Variant)

'PARAM_VECTOR:
'X-Axis Values /Color Code/% Fill
'1980/8421631/30.0%
'1984/16764108/45.0%
'2000/6697881/50.0%
'2005///

'Added ability to read trend chart change points and color codes,
'add polygons to highlihgt chart range between change points

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim NSIZE As Long
Dim NPOLY As Long

Dim X_STR As String
Dim Y_STR As String

Dim COLOUR_VAL As Double
Dim PERCENT_VAL As Double

Dim XNODE_VAL As Double
Dim YNODE_VAL As Double
Dim XMIN_VAL As Double
Dim XMAX_VAL As Double
Dim YMIN_VAL As Double
Dim YMAX_VAL As Double
Dim XLEFT_VAL As Double
Dim YTOP_VAL As Double
Dim XWIDTH_VAL As Double
Dim YHEIGHT_VAL As Double

Dim SHAPE_OBJ As Excel.Shape
Dim SERIES_OBJ As Excel.Series
Dim BUILDER_OBJ As Excel.FreeformBuilder

'Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

EXCEL_CHART_POLYGONS_XADD_FUNC = False
'PARAM_VECTOR = PARAM_RNG

With CHART_OBJ
    For Each SHAPE_OBJ In .Shapes: SHAPE_OBJ.Delete: Next SHAPE_OBJ
' Establish Plot Area Dimensions
    XLEFT_VAL = .PlotArea.InsideLeft
    XWIDTH_VAL = .PlotArea.InsideWidth
    YTOP_VAL = .PlotArea.InsideTop
    YHEIGHT_VAL = .PlotArea.InsideHeight
    XMIN_VAL = .Axes(1).MinimumScale
    XMAX_VAL = .Axes(1).MaximumScale
    YMIN_VAL = .Axes(2).MinimumScale
    YMAX_VAL = .Axes(2).MaximumScale

' Determine number of polygons to be added
    NPOLY = UBound(PARAM_VECTOR, 1)
' Clear previous polygon related series
    NSIZE = .SeriesCollection.COUNT
    ' MsgBox "No series = " & NSIZE
    If NSIZE > 1 Then
        On Error Resume Next
        For i = 2 To NSIZE
            .SeriesCollection(2).Delete             ' Use 2 - xl renums series to 2
        Next i
    End If
    ' Cycle through change points to add new polygon series
    l = 1
    For i = 2 To NPOLY + 1
        If PARAM_VECTOR(l, 1) = "" Or PARAM_VECTOR(l + 1, 1) = "" Then: Exit For
        COLOUR_VAL = PARAM_VECTOR(l, 2)
        PERCENT_VAL = PARAM_VECTOR(l, 3) 'fill percent
        ' Add new series
        .SeriesCollection.NewSeries
        X_STR = "{" & PARAM_VECTOR(l, 1) & ", " & PARAM_VECTOR(l + 1, 1) & "," & _
                      PARAM_VECTOR(l + 1, 1) & "," & PARAM_VECTOR(l, 1) & "}"
                      ' Construct line X values
        Y_STR = "{" & YMIN_VAL & " ," & YMIN_VAL & " ," & _
                      YMAX_VAL & ", " & YMAX_VAL & "}"    ' Construct Y values
        .SeriesCollection(i).XValues = X_STR
        .SeriesCollection(i).Values = Y_STR
        Set SERIES_OBJ = .SeriesCollection(i)
        j = SERIES_OBJ.Points.COUNT ' first point
        XNODE_VAL = XLEFT_VAL + (SERIES_OBJ.XValues(j) - XMIN_VAL) * XWIDTH_VAL / (XMAX_VAL - XMIN_VAL)
        YNODE_VAL = YTOP_VAL + (YMAX_VAL - SERIES_OBJ.Values(j)) * YHEIGHT_VAL / (YMAX_VAL - YMIN_VAL)
        Set BUILDER_OBJ = .Shapes.BuildFreeform(msoEditingAuto, XNODE_VAL, YNODE_VAL)
        ' remaining points
        For k = 1 To j
            XNODE_VAL = XLEFT_VAL + (SERIES_OBJ.XValues(k) - XMIN_VAL) * XWIDTH_VAL / (XMAX_VAL - XMIN_VAL)
            YNODE_VAL = YTOP_VAL + (YMAX_VAL - SERIES_OBJ.Values(k)) * YHEIGHT_VAL / (YMAX_VAL - YMIN_VAL)
            BUILDER_OBJ.AddNodes msoSegmentLine, msoEditingAuto, XNODE_VAL, YNODE_VAL
        Next k
        Set SHAPE_OBJ = BUILDER_OBJ.ConvertToShape
        With SHAPE_OBJ
            .Fill.ForeColor.RGB = COLOUR_VAL
            .Fill.Transparency = PERCENT_VAL
            .Line.Visible = msoFalse
            .ZOrder msoSendToBack
        End With
        'MsgBox i & vbTab & COLOUR_VAL & vbTab & PERCENT_VAL
        l = l + 1
    Next i
    ' Clear polygon related series
    NSIZE = .SeriesCollection.COUNT
    If NSIZE > 1 Then
        On Error Resume Next
        For i = 2 To NSIZE
            .SeriesCollection(2).Delete             ' Use 2 - xl renums series to 2
        Next i
    End If
End With

EXCEL_CHART_POLYGONS_XADD_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_POLYGONS_XADD_FUNC = False
End Function
