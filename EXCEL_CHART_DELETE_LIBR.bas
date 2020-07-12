Attribute VB_Name = "EXCEL_CHART_DELETE_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHARTS_DELETE_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : DELETE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHARTS_DELETE_FUNC(ByRef REF_VECTOR As Variant, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

Dim i As Long
Dim CHART_OBJ As Excel.Chart
Dim MATCH_FLAG As Boolean
        
On Error GoTo ERROR_LABEL
        
If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

EXCEL_CHARTS_DELETE_FUNC = False
For Each CHART_OBJ In SRC_WSHEET.ChartObjects
      MATCH_FLAG = False
        For i = LBound(REF_VECTOR) To UBound(REF_VECTOR)
            If REF_VECTOR(i) = CHART_OBJ.name Then
                MATCH_FLAG = True
                Exit For
            End If
        Next i
      If MATCH_FLAG = False Then: CHART_OBJ.Delete
Next CHART_OBJ

EXCEL_CHARTS_DELETE_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHARTS_DELETE_FUNC = False
End Function



