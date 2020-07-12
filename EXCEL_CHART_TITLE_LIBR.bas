Attribute VB_Name = "EXCEL_CHART_TITLE_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_TITLE_FUNC
'DESCRIPTION   : SETUP CHART TITLE
'LIBRARY       : EXCEL_CHART
'GROUP         : TITLE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_TITLE_FUNC(ByRef CHART_OBJ As Excel.ChartObject, _
ByVal TITLE_STR As String)

On Error GoTo ERROR_LABEL

EXCEL_CHART_TITLE_FUNC = False
       
With CHART_OBJ.Chart
     .HasTitle = True
     .ChartTitle.Characters.Text = TITLE_STR
    With .ChartTitle.Font
        .name = "Helvetica-Narrow"
        .FontStyle = "Bold"
        .Size = 16
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

EXCEL_CHART_TITLE_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_TITLE_FUNC = False
End Function



