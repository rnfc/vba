Attribute VB_Name = "EXCEL_CHART_TYPE_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

Function EXCEL_CHART_TYPE_CHOOSER_FUNC(ByRef CHART_OBJ As Excel.Chart, _
Optional ByVal CHART_TYPE_INT As Integer = 0)

On Error GoTo ERROR_LABEL

EXCEL_CHART_TYPE_CHOOSER_FUNC = True

Select Case CHART_TYPE_INT
Case 0
    CHART_OBJ.ChartType = xlLineMarkers
Case 1
    CHART_OBJ.ChartType = xlLine
Case 2
    CHART_OBJ.ChartType = xlPie
Case 3
    CHART_OBJ.ChartType = xlRadar
Case 4
    CHART_OBJ.ChartType = xlArea
Case Else '5
    CHART_OBJ.ChartType = xlColumnClustered
End Select

EXCEL_CHART_TYPE_CHOOSER_FUNC = True
    
Exit Function
ERROR_LABEL:
EXCEL_CHART_TYPE_CHOOSER_FUNC = False
End Function
