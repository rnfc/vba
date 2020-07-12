Attribute VB_Name = "EXCEL_CHART_PARAMETERS_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_PARAMETERS_FUNC
'DESCRIPTION   : RETURN FEATURES OF CHART
'LIBRARY       : EXCEL_CHART
'GROUP         : PARAMETERS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_PARAMETERS_FUNC(ByVal CHART_NAME_STR As Variant, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

Dim CHART_OBJ As Excel.ChartObject
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

If EXCEL_CHART_LOOK_FUNC(CHART_NAME_STR, 0, SRC_WSHEET) = True Then
    Set CHART_OBJ = SRC_WSHEET.ChartObjects(CHART_NAME_STR)
Else: GoTo ERROR_LABEL
End If

ReDim TEMP_VECTOR(1 To 11, 1 To 2)

TEMP_VECTOR(1, 1) = "Chart name: "
TEMP_VECTOR(1, 2) = CHART_OBJ.Chart.name

TEMP_VECTOR(2, 1) = "Left property: "
TEMP_VECTOR(2, 2) = CHART_OBJ.Left

TEMP_VECTOR(3, 1) = "Top property: "
TEMP_VECTOR(3, 2) = CHART_OBJ.Top

TEMP_VECTOR(4, 1) = "Height property: "
TEMP_VECTOR(4, 2) = CHART_OBJ.Height

TEMP_VECTOR(5, 1) = "WIDTH property: "
TEMP_VECTOR(5, 2) = CHART_OBJ.Width

TEMP_VECTOR(6, 1) = "Chart type: "
TEMP_VECTOR(6, 2) = CHART_OBJ.Chart.ChartType

TEMP_VECTOR(7, 1) = "HasLegend property: "
TEMP_VECTOR(7, 2) = CHART_OBJ.Chart.HasLegend

TEMP_VECTOR(8, 1) = "HasTitle property: "
TEMP_VECTOR(8, 2) = CHART_OBJ.Chart.HasTitle

TEMP_VECTOR(9, 1) = "Title: "
If TEMP_VECTOR(8, 2) = True Then
    TEMP_VECTOR(9, 2) = CHART_OBJ.Chart.ChartTitle.Text
Else: TEMP_VECTOR(9, 2) = ""
End If

TEMP_VECTOR(10, 1) = "Number of series plotted: "
TEMP_VECTOR(10, 2) = CHART_OBJ.Chart.SeriesCollection.COUNT

TEMP_VECTOR(11, 1) = "Horizontal Axis: Format of tick labels "
TEMP_VECTOR(11, 2) = CHART_OBJ.Chart.Axes(xlCategory).TickLabels.NumberFormat

EXCEL_CHART_PARAMETERS_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
EXCEL_CHART_PARAMETERS_FUNC = Err.number
End Function


