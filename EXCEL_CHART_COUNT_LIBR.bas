Attribute VB_Name = "EXCEL_CHART_COUNT_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHARTS_COUNT_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : COUNT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHARTS_COUNT_FUNC(Optional ByRef SRC_WSHEET As Excel.Worksheet)

Dim i As Long
Dim CHART_OBJ As Excel.ChartObject

On Error GoTo ERROR_LABEL

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

i = 0
For Each CHART_OBJ In SRC_WSHEET.ChartObjects
    i = i + 1
Next CHART_OBJ

EXCEL_CHARTS_COUNT_FUNC = i

Exit Function
ERROR_LABEL:
EXCEL_CHARTS_COUNT_FUNC = Err.number
End Function
