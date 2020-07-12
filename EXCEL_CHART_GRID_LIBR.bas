Attribute VB_Name = "EXCEL_CHART_GRID_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : CHART_GRID_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : GRID
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_GRID_FUNC(ByRef CHART_OBJ As Excel.Chart)
        
On Error GoTo ERROR_LABEL

EXCEL_CHART_GRID_FUNC = False

With CHART_OBJ
    With .Axes(xlCategory)
        .HasMajorGridlines = True
       ' .HasMinorGridlines = True
    End With
    With .Axes(xlCategory).MajorGridlines.Border
        .ColorIndex = 1
        .WEIGHT = xlHairline
        .LineStyle = xlDot
    End With
    With .Axes(xlValue)
        .HasMajorGridlines = True
       ' .HasMinorGridlines = True
    End With
    With .Axes(xlValue).MajorGridlines.Border
        .ColorIndex = 1
        .WEIGHT = xlHairline
        .LineStyle = xlDot
    End With
End With
EXCEL_CHART_GRID_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_GRID_FUNC = False
End Function
