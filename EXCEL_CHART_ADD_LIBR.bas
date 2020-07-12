Attribute VB_Name = "EXCEL_CHART_ADD_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_ADD_CHART_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : ADD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'REFERENCE     : http://www.databison.com/index.php/vba-chart/
'************************************************************************************
'************************************************************************************

Function EXCEL_ADD_CHART_FUNC(ByVal CHART_NAME_STR As Variant, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByVal LEFT_VAL As Double = 205, _
Optional ByVal WIDTH_VAL As Double = 100, _
Optional ByVal TOP_VAL As Double = 100, _
Optional ByVal HEIGHT_VAL As Double = 100, _
Optional ByVal DST_WSHEET As Excel.Worksheet) As Excel.ChartObject

'Left:=100, Width:=400, Top:=25, Height:=300

Dim CHART_OBJ As Excel.ChartObject
 
On Error Resume Next
    If OUTPUT = 0 Then
        If DST_WSHEET Is Nothing Then: Set DST_WSHEET = ActiveSheet
        DST_WSHEET.ChartObjects(CHART_NAME_STR).Delete
    Else
        Excel.Application.DisplayAlerts = False
        ActiveWorkbook.Sheets(CHART_NAME_STR).Delete
        Excel.Application.DisplayAlerts = True
        Set DST_WSHEET = ActiveWorkbook.Worksheets(1)
    End If
    If Err.number <> 0 Then: Err.Clear
'On Error GoTo 0
On Error GoTo ERROR_LABEL

Set CHART_OBJ = DST_WSHEET.ChartObjects.Add(Left:=LEFT_VAL, Width:=WIDTH_VAL, Top:=TOP_VAL, Height:=HEIGHT_VAL)
Call EXCEL_CHART_REMOVE_SERIES_FUNC(CHART_OBJ)
If CHART_NAME_STR <> "" Then: CHART_OBJ.name = CHART_NAME_STR
If OUTPUT = 0 Then
    CHART_OBJ.Chart.Location where:=xlLocationAsObject, name:=DST_WSHEET.name
Else
    CHART_OBJ.Chart.Location where:=xlLocationAsNewSheet, name:=CHART_NAME_STR
End If

Set EXCEL_ADD_CHART_FUNC = CHART_OBJ
Set CHART_OBJ = Nothing
Exit Function
ERROR_LABEL:
Set CHART_OBJ = Nothing
Set EXCEL_ADD_CHART_FUNC = Nothing
End Function
