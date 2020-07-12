Attribute VB_Name = "EXCEL_CHART_LOOK_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHARTS_LIST_FUNC
'DESCRIPTION   : LOOK FOR A CHART IN A WORKSHEET
'LIBRARY       : EXCEL_CHART
'GROUP         : LOOK
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHARTS_LIST_FUNC(Optional ByVal SRC_WSHEET As Excel.Worksheet)

Dim i As Long
Dim j As Long
Dim CHART_OBJ As Excel.Chart
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

j = SRC_WSHEET.ChartObjects.COUNT
If j = 0 Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To j, 1 To 2)
TEMP_MATRIX(0, 1) = "CHART NAME"
TEMP_MATRIX(0, 2) = "CHART TYPE"

For i = 1 To j
    Set CHART_OBJ = SRC_WSHEET.ChartObjects(i).Chart
    TEMP_MATRIX(i, 1) = CHART_OBJ.name
    TEMP_MATRIX(i, 2) = CHART_OBJ.ChartType
Next i

EXCEL_CHARTS_LIST_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
EXCEL_CHARTS_LIST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_LOOK_FUNC
'DESCRIPTION   : LOOK FOR A CHART IN A WORKSHEET
'LIBRARY       : EXCEL_CHART
'GROUP         : LOOK
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_LOOK_FUNC(ByVal CHART_NAME_STR As String, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

Dim i As Long
Dim MATCH_FLAG As Boolean
Dim CHART_OBJ As Excel.ChartObject

On Error GoTo ERROR_LABEL

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

MATCH_FLAG = False
i = 0
For Each CHART_OBJ In SRC_WSHEET.ChartObjects
    i = i + 1
    If CHART_OBJ.name = CHART_NAME_STR Then
        MATCH_FLAG = True
        Exit For
    End If
Next CHART_OBJ

If MATCH_FLAG = True Then
    Select Case OUTPUT
    Case 0
        EXCEL_CHART_LOOK_FUNC = True
    Case Else
        EXCEL_CHART_LOOK_FUNC = i
    End Select
Else
    EXCEL_CHART_LOOK_FUNC = False
End If

Exit Function
ERROR_LABEL:
EXCEL_CHART_LOOK_FUNC = Err.number
End Function

