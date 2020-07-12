Attribute VB_Name = "EXCEL_CHART_MOVING_AVG_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_MOVING_AVERAGE_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         :
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

' Procedure to show moving avg for user specified period
' Cycles through all charts, applies same criteria to each chart

'///////////////////////////////////////////////////////////////////
'User choices:
'1.Show data series?
'2.Moving Avg Period? If 0, no moving avg
'///////////////////////////////////////////////////////////////////

Function EXCEL_CHART_MOVING_AVERAGE_FUNC(ByRef CHART_OBJ As Excel.Chart, _
Optional ByVal MA_PERIOD As Long = 2, _
Optional ByVal DISPLAY_FLAG As Boolean = True)
'DISPLAY_FLAG --> Data Display On/Off

Dim i As Long
Dim j As Long
Dim k As Long
Dim NSIZE As Long

On Error GoTo ERROR_LABEL

EXCEL_CHART_MOVING_AVERAGE_FUNC = False

'Set moving Averages
With CHART_OBJ
    NSIZE = .SeriesCollection.COUNT
    ' Remove previous trendlines
    ' Check to see how many trend lines
    For i = 1 To NSIZE
        j = .SeriesCollection(i).Trendlines.COUNT
        If j > 0 Then
            For k = 1 To j: .SeriesCollection(i).Trendlines(k).Delete: Next k
        End If
        If MA_PERIOD >= 2 Then ' Check to see if period > 0; if yes, add trend line
            With .SeriesCollection(i)
                k = .Points.COUNT
                With .Trendlines.Add(Type:=xlMovingAvg, PERIOD:=MA_PERIOD, FORWARD:=0, Backward:=0, _
                    DisplayEquation:=False, DisplayRSquared:=False)
                    With .Border
                        .ColorIndex = 10
                        .WEIGHT = xlMedium
                        .LineStyle = xlHairline
                    End With
                End With
            End With
        End If
        If DISPLAY_FLAG = False Then ' Check to see if data series to be plotted
            With .SeriesCollection(i).Border
                .WEIGHT = xlHairline
                .LineStyle = xlNone
            End With
        Else
            With .SeriesCollection(i).Border
                .WEIGHT = xlThin
                .LineStyle = xlAutomatic
            End With
        End If
    Next i
End With

EXCEL_CHART_MOVING_AVERAGE_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_MOVING_AVERAGE_FUNC = False
End Function
