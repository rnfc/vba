Attribute VB_Name = "EXCEL_CHART_RANGE_LIBR"

'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------



'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_RANGE_SELECT_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : LOAD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_RANGE_SELECT_FUNC() As Excel.Range

On Error GoTo ERROR_LABEL

ACTIVATE_FLAG = False   ' True to re-activate the input range
Set DATA_RNG = EXCEL_CHART_RANGE_MSG_FUNC("Please select the range containing the DATA POINTS" & Chr(13) & "(press select a single column)")
If EXCEL_CHART_RANGE_CHECK_FUNC(DATA_RNG) Then
    MsgBox "Incorrect Input Data !"
    End
ElseIf Not (EXCEL_CHART_NUMERIC_CHECK_FUNC(DATA_RNG)) Then
    MsgBox "Incorrect Input Data !"
    End
End If

Exit Function
ERROR_LABEL:
Set EXCEL_CHART_RANGE_SELECT_FUNC = Nothing
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_RANGE_MSG_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : LOAD
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Private Function EXCEL_CHART_RANGE_MSG_FUNC(ByVal BOX_MSG_STR As String) As Excel.Range

On Error GoTo ERROR_LABEL

Set EXCEL_CHART_RANGE_MSG_FUNC = Nothing
Set EXCEL_CHART_RANGE_MSG_FUNC = Application.InputBox(BOX_MSG_STR, "Select Range", Selection.Address, , , , , 8)

Exit Function
ERROR_LABEL:
Set EXCEL_CHART_RANGE_MSG_FUNC = Nothing
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_RANGE_CHECK_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : LOAD
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Private Function EXCEL_CHART_RANGE_CHECK_FUNC(ByRef SRC_RNG As Excel.Range) As Boolean

Dim CELL_RNG As Excel.Range

On Error GoTo ERROR_LABEL

EXCEL_CHART_RANGE_CHECK_FUNC = True
If SRC_RNG.Rows.COUNT > 0 And SRC_RNG.Columns.COUNT = 1 Then
    EXCEL_CHART_RANGE_CHECK_FUNC = False
    Exit Function
End If
For Each CELL_RNG In SRC_RNG.Cells
    If Not (Application.WorksheetFunction.IsNumber(CELL_RNG.value)) Then
        EXCEL_CHART_RANGE_CHECK_FUNC = False
        Exit Function
    End If
Next CELL_RNG

Exit Function
ERROR_LABEL:
EXCEL_CHART_RANGE_CHECK_FUNC = True
End Function
