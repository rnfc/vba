Attribute VB_Name = "EXCEL_CHART_LEGEND_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_LEGEND_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : LEGEND
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_LEGEND_FUNC(ByRef CHART_OBJ As Excel.Chart)

On Error GoTo ERROR_LABEL
    
EXCEL_CHART_LEGEND_FUNC = False

With CHART_OBJ
    .HasLegend = True
    .Legend.Left = 621
    .Legend.Top = 320
    With .Legend.Font
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
End With

EXCEL_CHART_LEGEND_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_LEGEND_FUNC = False
End Function


