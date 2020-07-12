Attribute VB_Name = "EXCEL_CHART_AXES_LIBR"

'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_AXES_CATEGORY_SCALE_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : AXES
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_AXES_CATEGORY_SCALE_FUNC(ByRef CHART_OBJ As Excel.Chart, _
Optional ByVal CROSS_VAL As Long = 5, _
Optional ByVal LABEL_VAL As Long = 2, _
Optional ByVal MARK_VAL As Long = 3, _
Optional ByVal BETWEEN_FLAG As Boolean = True, _
Optional ByVal REVERSE_FLAG As Boolean = False)

On Error GoTo ERROR_LABEL

EXCEL_CHART_AXES_CATEGORY_SCALE_FUNC = False
With CHART_OBJ.Axes(xlCategory)
    .CrossesAt = CROSS_VAL
    .TickLabelSpacing = LABEL_VAL
    .TickMarkSpacing = MARK_VAL
    .AxisBetweenCategories = BETWEEN_FLAG
    .ReversePlotOrder = REVERSE_FLAG
End With

EXCEL_CHART_AXES_CATEGORY_SCALE_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_AXES_CATEGORY_SCALE_FUNC = False
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_AXES_VALUE_CATEGORY_SCALE1_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : AXES
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_AXES_VALUE_CATEGORY_SCALE1_FUNC(ByRef CHART_OBJ As Excel.Chart, _
Optional ByVal YAUTO_FLAG As Boolean = True, _
Optional ByVal XAUTO_FLAG As Boolean = True, _
Optional ByVal YMIN_VAL As Double = 0, _
Optional ByVal YMAX_VAL As Double = 0, _
Optional ByVal XMIN_VAL As Double = 0, _
Optional ByVal XMAX_VAL As Double = 0)

On Error GoTo ERROR_LABEL

EXCEL_CHART_AXES_VALUE_CATEGORY_SCALE1_FUNC = False
If YAUTO_FLAG = False Then ', xlSecondary
    With CHART_OBJ.Axes(xlValue)
        .MinimumScale = YMIN_VAL
        .MaximumScale = YMAX_VAL
    End With
ElseIf YAUTO_FLAG = True Then
    With CHART_OBJ.Axes(xlValue)
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
    End With
End If
      
If XAUTO_FLAG = False Then
    With CHART_OBJ.Axes(xlCategory)
        .MinimumScale = XMIN_VAL
        .MaximumScale = XMAX_VAL
    End With
ElseIf XAUTO_FLAG = True Then
    With CHART_OBJ.Axes(xlCategory)
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
    End With
End If

EXCEL_CHART_AXES_VALUE_CATEGORY_SCALE1_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_AXES_VALUE_CATEGORY_SCALE1_FUNC = False
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_AXES_VALUE_CATEGORY_SCALE1_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : AXES
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'REFERENCE     : Procedure to establish Y axis chart scales
' Based on Chap 15 - Advanced Charting Techniques (Bullen, Bovey, Green)
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_AXES_VALUE_CATEGORY_SCALE2_FUNC(ByRef CHART_OBJ As Excel.Chart, _
ByVal YMIN_VAL As Double, _
ByVal YMAX_VAL As Double, _
ByVal XMIN_VAL As Double, _
ByVal XMAX_VAL As Double, _
Optional ByVal GROUP_VAL As Variant = xlPrimary)

Dim DMIN_VAL As Double
Dim DMAX_VAL As Double
Dim POWER_VAL As Double
Dim DSCALE_VAL As Double

On Error GoTo ERROR_LABEL

With CHART_OBJ.Axes(xlCategory, GROUP_VAL)
.MinimumScale = XMIN_VAL
.MaximumScale = XMAX_VAL
End With

' Get Max & Min
DMAX_VAL = YMAX_VAL
DMIN_VAL = YMIN_VAL
' Check to see if max and min are the same
If DMAX_VAL = DMIN_VAL Then
   DSCALE_VAL = DMAX_VAL
   DMAX_VAL = DMAX_VAL * 1.01
   DMIN_VAL = DMIN_VAL * 0.99
End If
' Check to see if DMAX_VAL is bigger than DMIN_VAL, swap if not
If DMAX_VAL < DMIN_VAL Then
   DSCALE_VAL = DMAX_VAL
   DMAX_VAL = DMIN_VAL
   DMIN_VAL = DSCALE_VAL
End If
' Make DMAX_VAL a little bigger and DMIN_VAL a little smaller
If DMAX_VAL > 0 Then
   DMAX_VAL = DMAX_VAL + (DMAX_VAL - DMIN_VAL) * 0.01
Else
   DMAX_VAL = DMAX_VAL - (DMAX_VAL - DMIN_VAL) * 0.01
End If

If DMIN_VAL > 0 Then
   DMIN_VAL = DMIN_VAL - (DMAX_VAL - DMIN_VAL) * 0.01
Else
   DMIN_VAL = DMIN_VAL + (DMAX_VAL - DMIN_VAL) * 0.01
End If
'What if the y are both 0?
If (DMAX_VAL = 0) And (DMIN_VAL = 0) Then DMAX_VAL = 1
'Round max & min to reasonalbe values to chart
'Find range of values to chart
POWER_VAL = Log(DMAX_VAL - DMIN_VAL) / Log(10)
DSCALE_VAL = 10 ^ (POWER_VAL - Int(POWER_VAL))
'Find scaling factor
Select Case DSCALE_VAL
Case 0 To 2.5
    DSCALE_VAL = 0.2
Case 2.5 To 5
    DSCALE_VAL = 0.5
Case 5 To 7.5
    DSCALE_VAL = 1
Case Else
    DSCALE_VAL = 2
End Select
'Calculate the scaling factor (major unit)
DSCALE_VAL = DSCALE_VAL * 10 ^ Int(POWER_VAL)
'Round Axis values to nearest scaling factor
DMIN_VAL = DSCALE_VAL * Int(DMIN_VAL / DSCALE_VAL)
DMAX_VAL = DSCALE_VAL * (Int(DMAX_VAL / DSCALE_VAL) + 1)
'Set Chart Y axis scale
With CHART_OBJ.Axes(xlValue, GROUP_VAL)
    .MinimumScale = DMIN_VAL
    .MaximumScale = DMAX_VAL
    .MajorUnit = DSCALE_VAL
End With

EXCEL_CHART_AXES_VALUE_CATEGORY_SCALE2_FUNC = Array(DMIN_VAL, DMAX_VAL, DSCALE_VAL)

Exit Function
ERROR_LABEL:
EXCEL_CHART_AXES_VALUE_CATEGORY_SCALE2_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_AXES_DATES_SCALE_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : AXES
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_AXES_DATES_SCALE_FUNC(ByVal CHART_OBJ As Excel.Chart, _
ByVal XMIN_VAL As Double, _
ByVal XMAX_VAL As Double, _
ByVal XMAJOR_VAL As Double, _
Optional ByVal GROUP_VAL As Variant = xlPrimary, _
Optional ByVal VERSION As Integer = 0)

' This procedure establishes X axis min & max values and major unit
' User Controls: start & end dates, major unit and date format
'Establish chart spec sheet, be careful of name - DO NOT Use Target ID in name
'Establish ranges for: START, END, MAJOR UNIT , Date_format
'Establish control box for date format so that that user can select data format option
'VERSION: Date formats
'0:m/d/yy
'1:m/d
'2:d
'3:m 'yy
'4:mmm
'5:mmm,yy

On Error GoTo ERROR_LABEL

EXCEL_CHART_AXES_DATES_SCALE_FUNC = False

With CHART_OBJ.Axes(xlCategory, GROUP_VAL)
    .MinimumScale = XMIN_VAL
    .MaximumScale = XMAX_VAL
    .MajorUnit = XMAJOR_VAL
    Select Case VERSION
    Case 0
       .TickLabels.NumberFormat = "m/d/yy"
    Case 1
       .TickLabels.NumberFormat = "m/d"
    Case 2
       .TickLabels.NumberFormat = "d"
    Case 3
       .TickLabels.NumberFormat = "m 'yy"
    Case 4
       .TickLabels.NumberFormat = "mmm"
    Case Else
       .TickLabels.NumberFormat = "mmm, yy"
    End Select
End With

EXCEL_CHART_AXES_DATES_SCALE_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_AXES_DATES_SCALE_FUNC = False
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_AXES_CATEGORY_VALUE_FORMAT_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : AXES
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_AXES_CATEGORY_VALUE_FORMAT_FUNC(ByRef CHART_OBJ As Excel.Chart)
'ByRef CHART_OBJ As Excel.ChartObject --> CHART_OBJ.Chart.Axies.....

On Error GoTo ERROR_LABEL

EXCEL_CHART_AXES_CATEGORY_VALUE_FORMAT_FUNC = False

'LETS FORMAT THE CHART
With CHART_OBJ
    With .Axes(xlCategory)
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNextToAxis
    End With
    With .Axes(xlValue)
        .MajorTickMark = xlOutside
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNextToAxis
    End With
    With .ChartArea.Border
        .WEIGHT = 1
        .LineStyle = 0
    End With
    With .PlotArea.Border
        .ColorIndex = 1
        .WEIGHT = xlThin
        .LineStyle = xlContinuous
    End With
    With .PlotArea.Interior
        .ColorIndex = 2
        .PatternColorIndex = 1
        .Pattern = xlSolid
    End With
    With .ChartArea.Font
        .name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .Background = xlAutomatic
    End With
    .HasTitle = False
    .Axes(xlCategory, xlPrimary).HasTitle = False
    .Axes(xlValue, xlPrimary).HasTitle = True
    .HasTitle = True
    .ChartTitle.Characters.Text = "Control Chart"
    .ChartTitle.Left = 134
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Observations"
    With .Axes(xlCategory).TickLabels
        .Alignment = xlCenter
        .Offset = 100
        .ReadingOrder = xlContext
        .Orientation = xlHorizontal
    End With
    .Legend.Delete
    .PlotArea.Width = 310
    .Axes(xlValue).MajorGridlines.Delete
    .Axes(xlValue).CrossesAt = .Chart.Axes(xlValue).MinimumScale
    .ChartArea.Interior.ColorIndex = xlAutomatic
    .ChartArea.AutoScaleFont = True
End With

EXCEL_CHART_AXES_CATEGORY_VALUE_FORMAT_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_AXES_CATEGORY_VALUE_FORMAT_FUNC = False
End Function



'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_AXES_CATEGORY_FORMAT_FUNC
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         : AXES
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_AXES_CATEGORY_FORMAT_FUNC(ByRef CHART_OBJ As Excel.Chart)
     
On Error GoTo ERROR_LABEL

EXCEL_CHART_AXES_CATEGORY_FORMAT_FUNC = False

With CHART_OBJ
    .Axes(xlCategory).TickLabels.NumberFormat = "0"
    With .Axes(xlCategory).TickLabels.Font
       .name = "Helvetica-Narrow"
       .FontStyle = "Regular"
       .Size = 12
       .Strikethrough = False
       .Superscript = False
       .Subscript = False
       .OutlineFont = False
       .Shadow = False
       .Underline = xlNone
       .ColorIndex = xlAutomatic
       .Background = xlAutomatic
    End With
    
    With .Axes(xlCategory).Border
       .ColorIndex = 1
       .WEIGHT = xlMedium
       .LineStyle = xlContinuous
    End With
    
    With .Axes(xlCategory)
       .MajorTickMark = xlOutside
       .MinorTickMark = xlNone
       .TickLabelPosition = xlNextToAxis
    End With
    
    With .Axes(xlCategory).AxisTitle.Font
       .name = "Helvetica-Narrow"
       .FontStyle = "Bold"
       .Size = 14
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

EXCEL_CHART_AXES_CATEGORY_FORMAT_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_AXES_CATEGORY_FORMAT_FUNC = False
End Function


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_AXES_VALUE_FORMAT_FUNC
'DESCRIPTION   : Y AXIS FORMAT FRAME
'LIBRARY       : EXCEL_CHART
'GROUP         : AXES
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_AXES_VALUE_FORMAT_FUNC(ByRef CHART_OBJ As Excel.Chart)
     
On Error GoTo ERROR_LABEL

EXCEL_CHART_AXES_VALUE_FORMAT_FUNC = False

With CHART_OBJ
    With .Axes(xlValue)
       .MinimumScale = 1
       .MaximumScaleIsAuto = True
       .MinorUnitIsAuto = True
       .MajorUnitIsAuto = True
       .Crosses = xlAutomatic
       .ReversePlotOrder = False
       .ScaleType = False
    End With
    
    With .Axes(xlValue).TickLabels.Font
       .name = "Helvetica-Narrow"
       .FontStyle = "Regular"
       .Size = 12
       .Strikethrough = False
       .Superscript = False
       .Subscript = False
       .OutlineFont = False
       .Shadow = False
       .Underline = xlNone
       .ColorIndex = xlAutomatic
       .Background = xlAutomatic
    End With
    
    With .Axes(xlValue).Border
       .ColorIndex = 1
       .WEIGHT = xlMedium
       .LineStyle = xlContinuous
    End With
    
    With .Axes(xlValue)
       .MajorTickMark = xlOutside
       .MinorTickMark = xlNone
       .TickLabelPosition = xlNextToAxis
    End With
    
    '------------------------------------Y AXES TITLES
    
    With .Axes(xlValue).AxisTitle.Font
       .name = "Helvetica-Narrow"
       .FontStyle = "Bold"
       .Size = 14
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

EXCEL_CHART_AXES_VALUE_FORMAT_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_AXES_VALUE_FORMAT_FUNC = False
End Function


