Attribute VB_Name = "EXCEL_CHART_ERROR_BAR_LIBR"

'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_ERROR_BAR_ADD_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : ERROR_BAR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_ERROR_BAR_ADD_FUNC(ByRef CHART_OBJ As Excel.Chart, _
ByVal SERIE_NAME_STR As Variant, _
ByVal RNG_ADDRESS_STR As String, _
ByVal AMOUNT_VAL As Long)

Dim SERIES_OBJ As Excel.Series

On Error GoTo ERROR_LABEL

CHART_FORMAT_FRAME_FUNC = False

Set SERIES_OBJ = CHART_OBJ.SeriesCollection.NewSeries

With SERIES_OBJ
    .name = SERIE_NAME_STR
    .Values = RNG_ADDRESS_STR
    .ChartType = xlXYScatter
    '.ErrorBar Direction:=xlX, Include:=xlNone, Type:=xlFixedValue, Amount:=10000
    '.ErrorBar Direction:=xlX, Include:=xlUp, Type:=xlFixedValue, Amount:=20
    .ErrorBar direction:=xlX, Include:=xlPlusValues, _
    Type:=xlFixedValue, amount:=AMOUNT_VAL 'data_values.Rows.Count
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
    With .ErrorBars.Border
        .LineStyle = xlContinuous 'xlGray25
        .ColorIndex = 3 '15 /57 /3
        .WEIGHT = xlThin 'xlHairline
    'End With
    .ErrorBars.EndStyle = xlNoCap
End With
Set SERIES_OBJ = Nothing

EXCEL_CHART_ERROR_BAR_ADD_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_ERROR_BAR_ADD_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_ERROR_BAR_SET_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : ERROR_BAR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_ERROR_BAR_SET_FUNC(ByRef SERIE_OBJ As Excel.Series, _
ByRef AMOUNT_ARR As Variant, _
ByRef MINUS_ARR As Variant)
      
On Error GoTo ERROR_LABEL
      
EXCEL_CHART_ERROR_BAR_SET_FUNC = False
SERIE_OBJ.ErrorBar direction:=xlY, Include:=xlBoth, _
      Type:=xlCustom, amount:=AMOUNT_ARR, _
      MinusValues:=MINUS_ARR
EXCEL_CHART_ERROR_BAR_SET_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_ERROR_BAR_SET_FUNC = False
End Function

