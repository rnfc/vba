Attribute VB_Name = "EXCEL_CHART_BORDER_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : CHART_BORDER_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : BORDER
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function CHART_BORDER_FUNC(ByRef CHART_OBJ As Excel.Chart)
        
On Error GoTo ERROR_LABEL

CHART_BORDER_FUNC = False

With CHART_OBJ
    With .PlotArea.Border
        .WEIGHT = xlThin 'xlHairline
        .LineStyle = xlNone 'xlDash
    End With
    .PlotArea.Interior.ColorIndex = xlNone
    .PlotArea.Width = 700
End With

CHART_BORDER_FUNC = True

Exit Function
ERROR_LABEL:
CHART_BORDER_FUNC = False
End Function
