Attribute VB_Name = "FINAN_FI_BOND_MONOTOME_LIBR"

'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MONOTOME_CONVEX_INTERPOLATE_FUNC

'Reference:

'Patrick S. Hagan and Graeme West. Methods for constructing a yield curve. Wilmott
'magazine, p 70-81, May 2008. preprint pdf.

'References:
'http://www.finmod.co.za/Hagan_West_curves_AMF.pdf
'http://www.finmod.co.za/interpreview.pdf
'http://www.finmod.co.za/interpolationsummaryglossy.pdf
'http://www.finmod.co.za/WestYieldCurvesOIS.pdf

'LIBRARY       : BOND
'GROUP         : MONOTOME
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function MONOTOME_CONVEX_INTERPOLATE_FUNC(ByVal REFER_VAL As Double, _
ByRef DATA_RNG As Variant, _
Optional ByVal LAMBDA_VAL As Double = 0.5, _
Optional ByVal FORWARDS_FLAG As Boolean = False, _
Optional ByVal NEGATIVE_FLAG As Boolean = False, _
Optional ByVal COUNT_BASIS As Double = 365, _
Optional ByVal OUTPUT As Integer = 1) 'As Double

'COUNT_BASIS --> TermsAre
'ESTIMATES_FLAG --> forward_are_calced
'FORWARDS_FLAG --> Inputs are Forwards
'NEGATIVE_FLAG --> Negative Forwards Allowed
'Lambda --> 0 for unameliorated; must lie in the interval [0,1]. Adjust accordingly.

Dim i As Long
Dim NROWS As Long
Dim TEMP_VAL As Double
Dim XDATA_ARR() As Double
Dim YDATA_ARR() As Double
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
ReDim XDATA_ARR(-1 To NROWS)
ReDim YDATA_ARR(0 To NROWS)

For i = 1 To NROWS
  TEMP_VAL = DATA_MATRIX(i, 1) / COUNT_BASIS
  If TEMP_VAL <> XDATA_ARR(i) Then: XDATA_ARR(i) = TEMP_VAL
  TEMP_VAL = DATA_MATRIX(i, 2)
  If TEMP_VAL <> YDATA_ARR(i) Then: YDATA_ARR(i) = TEMP_VAL
Next i
 
Select Case OUTPUT
Case 0
    MONOTOME_CONVEX_INTERPOLATE_FUNC = MONOTOME_CONVEX_RATE_FUNC(REFER_VAL / COUNT_BASIS, XDATA_ARR(), YDATA_ARR(), True, , , FORWARDS_FLAG, NEGATIVE_FLAG, LAMBDA_VAL)
Case Else
    MONOTOME_CONVEX_INTERPOLATE_FUNC = MONOTOME_CONVEX_FORWARD_FUNC(REFER_VAL / COUNT_BASIS, XDATA_ARR(), YDATA_ARR(), True, , , FORWARDS_FLAG, NEGATIVE_FLAG, LAMBDA_VAL)
End Select

Exit Function
ERROR_LABEL:
MONOTOME_CONVEX_INTERPOLATE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : MONOTOME_CONVEX_BOOTSTRAP_FUNC
'DESCRIPTION   :
'LIBRARY       : BOND
'GROUP         : MONOTOME
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function MONOTOME_CONVEX_BOOTSTRAP_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal LAMBDA_VAL As Double = 0.5, _
Optional ByVal FORWARDS_FLAG As Boolean = False, _
Optional ByVal NEGATIVE_FLAG As Boolean = False, _
Optional ByVal DELTA_VAL As Double = 0.01, _
Optional ByVal COUNT_BASIS As Double = 365)

'COUNT_BASIS --> TermsAre
'ESTIMATES_FLAG --> forward_are_calced
'FORWARDS_FLAG --> Inputs are Forwards
'NEGATIVE_FLAG --> Negative Forwards Allowed
'Lambda --> 0 for unameliorated; must lie in the interval [0,1]. Adjust accordingly.

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NSIZE As Long

Dim TEMP_VAL As Double
Dim XDATA_ARR() As Double
Dim YDATA_ARR() As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

k = 10
l = 100
DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)

ReDim XDATA_ARR(-1 To NROWS)
ReDim YDATA_ARR(0 To NROWS)

For i = 1 To NROWS
'-------------------------------------------------------------
    DATA_MATRIX(i, 1) = DATA_MATRIX(i, 1) / COUNT_BASIS
'-------------------------------------------------------------
    TEMP_VAL = DATA_MATRIX(i, 1)
    If TEMP_VAL <> XDATA_ARR(i) Then: XDATA_ARR(i) = TEMP_VAL
    
    TEMP_VAL = DATA_MATRIX(i, 2)
    If TEMP_VAL <> YDATA_ARR(i) Then: YDATA_ARR(i) = TEMP_VAL
Next i

'-------------------------------------------------------------
If FORWARDS_FLAG = False Then
'-------------------------------------------------------------
    i = CInt(l * DATA_MATRIX(NROWS, 1) + k * 2) / l
    j = (i) / DELTA_VAL + 1
    ReDim TEMP_MATRIX(0 To j, 1 To 3)
    TEMP_MATRIX(0, 1) = "TN"
    TEMP_MATRIX(0, 2) = "CURVE"
    TEMP_MATRIX(0, 3) = "FORWARD"
    NSIZE = CInt(l * DATA_MATRIX(NROWS, 1) + k * 2) / l
    j = 1
    For TEMP_VAL = 0 To NSIZE Step DELTA_VAL
      TEMP_MATRIX(j, 1) = TEMP_VAL * COUNT_BASIS
      TEMP_MATRIX(j, 2) = MONOTOME_CONVEX_RATE_FUNC(TEMP_VAL, XDATA_ARR(), YDATA_ARR(), True, , , FORWARDS_FLAG, NEGATIVE_FLAG, LAMBDA_VAL)
      TEMP_MATRIX(j, 3) = MONOTOME_CONVEX_FORWARD_FUNC(TEMP_VAL, XDATA_ARR(), YDATA_ARR(), True, , , FORWARDS_FLAG, NEGATIVE_FLAG, LAMBDA_VAL)
      j = j + 1
    Next TEMP_VAL
'-------------------------------------------------------------
Else
'-------------------------------------------------------------
    i = CInt(l * DATA_MATRIX(NROWS, 1) + k) / l
    j = (i) / DELTA_VAL + 1
    
    ReDim TEMP_MATRIX(0 To j, 1 To 2)
    TEMP_MATRIX(0, 1) = "TN"
    TEMP_MATRIX(0, 2) = "FORWARD RATES"
    
    NSIZE = CInt(l * DATA_MATRIX(NROWS, 1) + k) / l
    
    j = 1
    For TEMP_VAL = 0 To NSIZE Step DELTA_VAL
      TEMP_MATRIX(j, 1) = TEMP_VAL * COUNT_BASIS
      TEMP_MATRIX(j, 2) = MONOTOME_CONVEX_FORWARD_FUNC(TEMP_VAL, XDATA_ARR(), YDATA_ARR(), True, , , FORWARDS_FLAG, NEGATIVE_FLAG, LAMBDA_VAL)
      j = j + 1
    Next TEMP_VAL
'-------------------------------------------------------------
End If
'-------------------------------------------------------------
  
MONOTOME_CONVEX_BOOTSTRAP_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MONOTOME_CONVEX_BOOTSTRAP_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_MONOTOME_CONVEX_CURVE_FUNC
'DESCRIPTION   :
'LIBRARY       : BOND
'GROUP         : MONOTOME
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function RNG_MONOTOME_CONVEX_CURVE_FUNC(ByRef DATA_RNG As Excel.Range, _
ByVal CHART_NAME_STR As String, _
Optional ByRef DST_WBOOK As Excel.Workbook)

Dim NROWS As Long
Dim NCOLUMNS As Long
Dim XMIN_VAL As Double
Dim XMAX_VAL As Double

Dim YMIN_VAL As Double
Dim YMAX_VAL As Double

Dim DST_CHART As Excel.Chart

On Error GoTo ERROR_LABEL

RNG_MONOTOME_CONVEX_CURVE_FUNC = False

If DST_WBOOK Is Nothing Then Set DST_WBOOK = ActiveWorkbook

DST_WBOOK.Charts.Add
Set DST_CHART = ActiveChart

With DST_CHART
    .ChartType = xlXYScatterSmoothNoMarkers
    .SetSourceData source:=DATA_RNG, PlotBy:=xlColumns
    .Location where:=xlLocationAsNewSheet, name:=CHART_NAME_STR
    With .Axes(xlCategory)
        .HasMajorGridlines = False
        .HasMinorGridlines = False
    End With
    With .Axes(xlValue)
        .HasMajorGridlines = False
        .HasMinorGridlines = False
    End With
    
    .HasLegend = True
    .Legend.Position = xlBottom
    With .Axes(xlCategory)
        .MinimumScale = 0
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        .ScaleType = xlLinear
        .DisplayUnit = xlNone
        .TickLabels.AutoScaleFont = True
        With .TickLabels.Font
            .name = "Arial"
            .FontStyle = "Regular"
            .Size = 9
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
    End With
    
    With .PlotArea
        With .Border
            .ColorIndex = 16
            .WEIGHT = xlThin
            .LineStyle = xlContinuous
        End With
        .Interior.ColorIndex = xlNone
    End With
    
    With .Axes(xlValue)
        .TickLabels.AutoScaleFont = True
        With .TickLabels.Font
            .name = "Arial"
            .FontStyle = "Regular"
            .Size = 9
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        .TickLabels.NumberFormat = "0%"
    End With
    
    On Error Resume Next
        With .SeriesCollection(1)
            .Border.ColorIndex = 3
            .Border.WEIGHT = xlMedium
            .Border.LineStyle = xlContinuous
            .MarkerBackgroundColorIndex = xlNone
            .MarkerForegroundColorIndex = xlNone
            .MarkerStyle = xlNone
            .Smooth = True
            .MarkerSize = 3
            .Shadow = False
        End With
        With .SeriesCollection(2)
            .Border.ColorIndex = 5
            .Border.WEIGHT = xlMedium
            .Border.LineStyle = xlContinuous
            .MarkerBackgroundColorIndex = xlNone
            .MarkerForegroundColorIndex = xlNone
            .MarkerStyle = xlNone
            .Smooth = True
            .MarkerSize = 3
            .Shadow = False
        End With
    On Error GoTo ERROR_LABEL
    
    .SizeWithWindow = True
    .Deselect
End With

NROWS = DATA_RNG.Rows.COUNT
NCOLUMNS = DATA_RNG.Columns.COUNT
Set DATA_RNG = Range(DATA_RNG.Cells(2, 1), DATA_RNG.Cells(NROWS, NCOLUMNS))

If NCOLUMNS > 2 Then 'Rates Inputs
    XMIN_VAL = WorksheetFunction.Min(DATA_RNG.Columns(2))
    XMAX_VAL = WorksheetFunction.max(DATA_RNG.Columns(2))
    
    YMIN_VAL = WorksheetFunction.Min(DATA_RNG.Columns(3))
    YMAX_VAL = WorksheetFunction.max(DATA_RNG.Columns(3))
    
    If XMIN_VAL < YMIN_VAL Then: YMIN_VAL = XMIN_VAL
    If XMAX_VAL > YMAX_VAL Then: YMAX_VAL = XMAX_VAL
Else 'Forward Rates Inputs
    YMIN_VAL = WorksheetFunction.Min(DATA_RNG.Columns(2))
    YMAX_VAL = WorksheetFunction.max(DATA_RNG.Columns(2))
End If
XMIN_VAL = WorksheetFunction.Min(DATA_RNG.Columns(1))
XMAX_VAL = WorksheetFunction.max(DATA_RNG.Columns(1))

Call EXCEL_CHART_AXES_VALUE_CATEGORY_SCALE2_FUNC(DST_CHART, YMIN_VAL, YMAX_VAL, XMIN_VAL, XMAX_VAL, 1)

RNG_MONOTOME_CONVEX_CURVE_FUNC = True

Exit Function
ERROR_LABEL:
RNG_MONOTOME_CONVEX_CURVE_FUNC = False
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : MONOTOME_CONVEX_RATE_FUNC
'DESCRIPTION   : value of r(t) for any t.
'LIBRARY       : BOND
'GROUP         : MONOTOME
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Private Function MONOTOME_CONVEX_RATE_FUNC(ByVal REFER_VAL As Double, _
ByRef XDATA_ARR() As Double, _
ByRef YDATA_ARR() As Double, _
Optional ByVal ESTIMATES_FLAG As Boolean = True, _
Optional ByRef FORWARD_ARR As Variant, _
Optional ByRef DISCRETE_ARR As Variant, _
Optional ByVal FORWARDS_FLAG As Boolean = False, _
Optional ByVal NEGATIVE_FLAG As Boolean = False, _
Optional ByVal LAMBDA As Double = 0) 'As Double

'COUNT_BASIS --> TermsAre
'ESTIMATES_FLAG --> forward_are_calced
'FORWARDS_FLAG --> Inputs are Forwards
'NEGATIVE_FLAG --> Negative Forwards Allowed
'Lambda --> 0 for unameliorated; must lie in the interval [0,1]. Adjust accordingly.

Dim i As Long
Dim NROWS As Long

Dim A_VAL As Double
Dim G_VAL As Double
Dim L_VAL As Double
Dim X_VAL As Double

Dim G0_VAL As Double
Dim G1_VAL As Double

Dim ETA_VAL As Double

On Error GoTo ERROR_LABEL

NROWS = UBound(XDATA_ARR)
'numbering refers to Wilmott paper
If ESTIMATES_FLAG = True Then
  ESTIMATES_FLAG = MONOTOME_CONVEX_ESTIMATES_FUNC(XDATA_ARR(), YDATA_ARR(), FORWARD_ARR, DISCRETE_ARR, FORWARDS_FLAG, NEGATIVE_FLAG, LAMBDA)
  If ESTIMATES_FLAG = False Then: GoTo ERROR_LABEL
End If

If REFER_VAL <= 0 Then
  MONOTOME_CONVEX_RATE_FUNC = FORWARD_ARR(0)
ElseIf REFER_VAL > XDATA_ARR(NROWS) Then
  MONOTOME_CONVEX_RATE_FUNC = MONOTOME_CONVEX_RATE_FUNC(XDATA_ARR(NROWS), XDATA_ARR(), YDATA_ARR(), ESTIMATES_FLAG, FORWARD_ARR, DISCRETE_ARR, FORWARDS_FLAG, NEGATIVE_FLAG, LAMBDA) * XDATA_ARR(NROWS) / REFER_VAL + MONOTOME_CONVEX_FORWARD_FUNC(XDATA_ARR(NROWS), XDATA_ARR(), YDATA_ARR(), ESTIMATES_FLAG, FORWARD_ARR, DISCRETE_ARR, FORWARDS_FLAG, NEGATIVE_FLAG, LAMBDA) * (1 - XDATA_ARR(NROWS) / REFER_VAL)
Else
  i = ARRAY_LAST_INDEX_FUNC(XDATA_ARR, REFER_VAL)
  L_VAL = XDATA_ARR(i + 1) - XDATA_ARR(i)
  'the X_VAL in (25)
  X_VAL = (REFER_VAL - XDATA_ARR(i)) / L_VAL
  G0_VAL = FORWARD_ARR(i) - DISCRETE_ARR(i + 1)
  G1_VAL = FORWARD_ARR(i + 1) - DISCRETE_ARR(i + 1)
  
  If X_VAL = 0 Or X_VAL = 1 Then
    G_VAL = 0
  ElseIf (G0_VAL < 0 And -0.5 * G0_VAL <= G1_VAL And G1_VAL <= -2 * G0_VAL) Or (G0_VAL > 0 And -0.5 * G0_VAL >= G1_VAL And G1_VAL >= -2 * G0_VAL) Then
    'zone (i)
    G_VAL = L_VAL * (G0_VAL * (X_VAL - 2 * X_VAL ^ 2 + X_VAL ^ 3) + G1_VAL * (-X_VAL ^ 2 + X_VAL ^ 3))
  ElseIf (G0_VAL < 0 And G1_VAL > -2 * G0_VAL) Or (G0_VAL > 0 And G1_VAL < -2 * G0_VAL) Then
    'zone (ii)
    '(29)
    ETA_VAL = (G1_VAL + 2 * G0_VAL) / (G1_VAL - G0_VAL)
    '(28)
    If X_VAL <= ETA_VAL Then
      G_VAL = G0_VAL * (REFER_VAL - XDATA_ARR(i))
    Else
      G_VAL = G0_VAL * (REFER_VAL - XDATA_ARR(i)) + (G1_VAL - G0_VAL) * (X_VAL - ETA_VAL) ^ 3 / (1 - ETA_VAL) ^ 2 / 3 * L_VAL
    End If
  ElseIf (G0_VAL > 0 And 0 > G1_VAL And G1_VAL > -0.5 * G0_VAL) Or (G0_VAL < 0 And 0 < G1_VAL And G1_VAL < -0.5 * G0_VAL) Then
    'zone (iii)
    '(31)
    ETA_VAL = 3 * G1_VAL / (G1_VAL - G0_VAL)
    '(30)
    If X_VAL < ETA_VAL Then
      G_VAL = L_VAL * (G1_VAL * X_VAL - 1 / 3 * (G0_VAL - G1_VAL) * ((ETA_VAL - X_VAL) ^ 3 / ETA_VAL ^ 2 - ETA_VAL))
    Else
      G_VAL = L_VAL * (2 / 3 * G1_VAL + 1 / 3 * G0_VAL) * ETA_VAL + G1_VAL * (X_VAL - ETA_VAL) * L_VAL
    End If
  ElseIf G0_VAL = 0 And G1_VAL = 0 Then
    G_VAL = 0
  Else
    'zone (iv)
    '(33)
    ETA_VAL = G1_VAL / (G1_VAL + G0_VAL)
    '(34)
    A_VAL = -G0_VAL * G1_VAL / (G0_VAL + G1_VAL)
    '(32)
    If X_VAL <= ETA_VAL Then
      G_VAL = L_VAL * (A_VAL * X_VAL - 1 / 3 * (G0_VAL - A_VAL) * ((ETA_VAL - X_VAL) ^ 3 / ETA_VAL ^ 2 - ETA_VAL))
    Else
      G_VAL = L_VAL * (2 / 3 * A_VAL + 1 / 3 * G0_VAL) * ETA_VAL + L_VAL * (A_VAL * (X_VAL - ETA_VAL) + (G1_VAL - A_VAL) / 3 * (X_VAL - ETA_VAL) ^ 3 / (1 - ETA_VAL) ^ 2)
    End If
  End If
  '(12)
  MONOTOME_CONVEX_RATE_FUNC = 1 / REFER_VAL * (XDATA_ARR(i) * YDATA_ARR(i) + DISCRETE_ARR(i + 1) * (REFER_VAL - XDATA_ARR(i)) + G_VAL)
End If

Exit Function
ERROR_LABEL:
MONOTOME_CONVEX_RATE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : MONOTOME_CONVEX_FORWARD_FUNC
'DESCRIPTION   :
'LIBRARY       : BOND
'GROUP         : MONOTOME
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Private Function MONOTOME_CONVEX_FORWARD_FUNC(ByVal REFER_VAL As Double, _
ByRef XDATA_ARR() As Double, _
ByRef YDATA_ARR() As Double, _
Optional ByVal ESTIMATES_FLAG As Boolean = True, _
Optional ByRef FORWARD_ARR As Variant, _
Optional ByRef DISCRETE_ARR As Variant, _
Optional ByVal FORWARDS_FLAG As Boolean = False, _
Optional ByVal NEGATIVE_FLAG As Boolean = False, _
Optional ByVal LAMBDA As Double = 0) 'As Double

'COUNT_BASIS --> TermsAre
'ESTIMATES_FLAG --> forward_are_calced
'FORWARDS_FLAG --> Inputs are Forwards
'NEGATIVE_FLAG --> Negative Forwards Allowed
'Lambda --> 0 for unameliorated; must lie in the interval [0,1]. Adjust accordingly.

Dim i As Long
Dim NROWS As Long

Dim A_VAL As Double
Dim G_VAL As Double
Dim X_VAL As Double

Dim G0_VAL As Double
Dim G1_VAL As Double

Dim ETA_VAL As Double

On Error GoTo ERROR_LABEL

NROWS = UBound(XDATA_ARR)
'numbering refers to Wilmott paper
If ESTIMATES_FLAG = True Then
  ESTIMATES_FLAG = MONOTOME_CONVEX_ESTIMATES_FUNC(XDATA_ARR(), YDATA_ARR(), FORWARD_ARR, DISCRETE_ARR, FORWARDS_FLAG, NEGATIVE_FLAG, LAMBDA)
  If ESTIMATES_FLAG = False Then: GoTo ERROR_LABEL
End If

If REFER_VAL <= 0 Then
  MONOTOME_CONVEX_FORWARD_FUNC = FORWARD_ARR(0)
ElseIf REFER_VAL > XDATA_ARR(NROWS) Then
  MONOTOME_CONVEX_FORWARD_FUNC = MONOTOME_CONVEX_FORWARD_FUNC(XDATA_ARR(NROWS), XDATA_ARR(), YDATA_ARR(), ESTIMATES_FLAG, FORWARD_ARR, DISCRETE_ARR, FORWARDS_FLAG, NEGATIVE_FLAG, LAMBDA)
Else
  i = ARRAY_LAST_INDEX_FUNC(XDATA_ARR, REFER_VAL)
  'the X_VAL in (25)
  X_VAL = (REFER_VAL - XDATA_ARR(i)) / (XDATA_ARR(i + 1) - XDATA_ARR(i))
  G0_VAL = FORWARD_ARR(i) - DISCRETE_ARR(i + 1)
  G1_VAL = FORWARD_ARR(i + 1) - DISCRETE_ARR(i + 1)
  If X_VAL = 0 Then
    G_VAL = G0_VAL
  ElseIf X_VAL = 1 Then
    G_VAL = G1_VAL
  ElseIf (G0_VAL < 0 And -0.5 * G0_VAL <= G1_VAL And G1_VAL <= -2 * G0_VAL) Or (G0_VAL > 0 And -0.5 * G0_VAL >= G1_VAL And G1_VAL >= -2 * G0_VAL) Then
    'zone (i)
    G_VAL = G0_VAL * (1 - 4 * X_VAL + 3 * X_VAL ^ 2) + G1_VAL * (-2 * X_VAL + 3 * X_VAL ^ 2)
  ElseIf (G0_VAL < 0 And G1_VAL > -2 * G0_VAL) Or (G0_VAL > 0 And G1_VAL < -2 * G0_VAL) Then
    'zone (ii)
    '(29)
    ETA_VAL = (G1_VAL + 2 * G0_VAL) / (G1_VAL - G0_VAL)
    '(28)
    If X_VAL <= ETA_VAL Then
      G_VAL = G0_VAL
    Else
      G_VAL = G0_VAL + (G1_VAL - G0_VAL) * ((X_VAL - ETA_VAL) / (1 - ETA_VAL)) ^ 2
    End If
  ElseIf (G0_VAL > 0 And 0 > G1_VAL And G1_VAL > -0.5 * G0_VAL) Or (G0_VAL < 0 And 0 < G1_VAL And G1_VAL < -0.5 * G0_VAL) Then
    'zone (iii)
    '(31)
    ETA_VAL = 3 * G1_VAL / (G1_VAL - G0_VAL)
    '(30)
    If X_VAL < ETA_VAL Then
      G_VAL = G1_VAL + (G0_VAL - G1_VAL) * ((ETA_VAL - X_VAL) / ETA_VAL) ^ 2
    Else
      G_VAL = G1_VAL
    End If
  ElseIf G0_VAL = 0 And G1_VAL = 0 Then
    G_VAL = 0
  Else
    'zone (iv)
    '(33)
    ETA_VAL = G1_VAL / (G1_VAL + G0_VAL)
    '(34)
    A_VAL = -G0_VAL * G1_VAL / (G0_VAL + G1_VAL)
    '(32)
    If X_VAL <= ETA_VAL Then
      G_VAL = A_VAL + (G0_VAL - A_VAL) * ((ETA_VAL - X_VAL) / ETA_VAL) ^ 2
    Else
      G_VAL = A_VAL + (G1_VAL - A_VAL) * ((ETA_VAL - X_VAL) / (1 - ETA_VAL)) ^ 2
    End If
  End If
  '(26)
  MONOTOME_CONVEX_FORWARD_FUNC = G_VAL + DISCRETE_ARR(i + 1)
End If

Exit Function
ERROR_LABEL:
MONOTOME_CONVEX_FORWARD_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : MONOTOME_CONVEX_ESTIMATES_FUNC
'DESCRIPTION   :
'LIBRARY       : BOND
'GROUP         : MONOTOME
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Private Function MONOTOME_CONVEX_ESTIMATES_FUNC( _
ByRef XDATA_ARR() As Double, _
ByRef YDATA_ARR() As Double, _
ByRef FORWARD_ARR As Variant, _
ByRef DISCRETE_ARR As Variant, _
Optional ByVal FORWARDS_FLAG As Boolean = False, _
Optional ByVal NEGATIVE_FLAG As Boolean = False, _
Optional ByVal LAMBDA As Double = 0)

'COUNT_BASIS --> TermsAre
'ESTIMATES_FLAG --> forward_are_calced
'FORWARDS_FLAG --> Inputs are Forwards
'NEGATIVE_FLAG --> Negative Forwards Allowed
'Lambda --> 0 for unameliorated; must lie in the interval [0,1]. Adjust accordingly.

Dim i As Long
Dim NROWS As Long
Dim THETA_MATRIX() As Double
Dim TERMS_FALSE_ARR() As Double
Dim MIN_MAX_ARR() As Double

On Error GoTo ERROR_LABEL

MONOTOME_CONVEX_ESTIMATES_FUNC = False

NROWS = UBound(XDATA_ARR)
ReDim FORWARD_ARR(0 To NROWS)
ReDim DISCRETE_ARR(0 To NROWS)

'extend the curve to time 0, for the purpose of calculating forward at time 1
XDATA_ARR(0) = 0
YDATA_ARR(0) = YDATA_ARR(1)
'step 1
If FORWARDS_FLAG = False Then
    For i = 1 To NROWS
        DISCRETE_ARR(i) = (XDATA_ARR(i) * YDATA_ARR(i) - XDATA_ARR(i - 1) * YDATA_ARR(i - 1)) / (XDATA_ARR(i) - XDATA_ARR(i - 1))
    Next i
Else
    For i = 1 To NROWS
        DISCRETE_ARR(i) = YDATA_ARR(i)
    Next i
End If

If LAMBDA = 0 Then
'forward_i estimation under the unameliorated method
'numbering refers to Wilmott paper
'step 2
    For i = 1 To NROWS - 1
        FORWARD_ARR(i) = (XDATA_ARR(i) - XDATA_ARR(i - 1)) / (XDATA_ARR(i + 1) - XDATA_ARR(i - 1)) * DISCRETE_ARR(i + 1) + (XDATA_ARR(i + 1) - XDATA_ARR(i)) / (XDATA_ARR(i + 1) - XDATA_ARR(i - 1)) * DISCRETE_ARR(i)
    Next i
'step 3
    If NEGATIVE_FLAG = False Then
        FORWARD_ARR(0) = COLLAR_FUNC(0, DISCRETE_ARR(1) - 0.5 * (FORWARD_ARR(1) - DISCRETE_ARR(1)), 2 * DISCRETE_ARR(1))
        FORWARD_ARR(NROWS) = COLLAR_FUNC(0, DISCRETE_ARR(NROWS) - 0.5 * (FORWARD_ARR(NROWS - 1) - DISCRETE_ARR(NROWS)), 2 * DISCRETE_ARR(NROWS))
        For i = 1 To NROWS - 1
            FORWARD_ARR(i) = COLLAR_FUNC(0, FORWARD_ARR(i), 2 * MINIMUM_FUNC(DISCRETE_ARR(i), DISCRETE_ARR(i + 1)))
        Next i
    End If
Else
'forward_i estimation under the ameliorated method numbering refers to AMF paper
    ReDim THETA_MATRIX(-1 To 1, -1 To NROWS + 1)
    ReDim TERMS_FALSE_ARR(-1 To NROWS + 1)
    ReDim Preserve DISCRETE_ARR(0 To NROWS + 1)
    ReDim MIN_MAX_ARR(-1 To 1, 0 To 2, 0 To NROWS)
    For i = 0 To NROWS
        TERMS_FALSE_ARR(i) = XDATA_ARR(i)
    Next i
    '(72) and (73)
    TERMS_FALSE_ARR(-1) = -TERMS_FALSE_ARR(1)
    
    TERMS_FALSE_ARR(NROWS + 1) = 2 * TERMS_FALSE_ARR(NROWS) - TERMS_FALSE_ARR(NROWS - 1)
    
    DISCRETE_ARR(0) = DISCRETE_ARR(1) - (TERMS_FALSE_ARR(1) - TERMS_FALSE_ARR(0)) / (TERMS_FALSE_ARR(2) - TERMS_FALSE_ARR(0)) * (DISCRETE_ARR(2) - DISCRETE_ARR(1))
    DISCRETE_ARR(NROWS + 1) = DISCRETE_ARR(NROWS) + (TERMS_FALSE_ARR(NROWS) - TERMS_FALSE_ARR(NROWS - 1)) / (TERMS_FALSE_ARR(NROWS) - TERMS_FALSE_ARR(NROWS - 2)) * (DISCRETE_ARR(NROWS) - DISCRETE_ARR(NROWS - 1))
    '(74)
    For i = 0 To NROWS
        MIN_MAX_ARR(0, 0, i) = (TERMS_FALSE_ARR(i) - TERMS_FALSE_ARR(i - 1)) / (TERMS_FALSE_ARR(i + 1) - TERMS_FALSE_ARR(i - 1)) * DISCRETE_ARR(i + 1) + (TERMS_FALSE_ARR(i + 1) - TERMS_FALSE_ARR(i)) / (TERMS_FALSE_ARR(i + 1) - TERMS_FALSE_ARR(i - 1)) * DISCRETE_ARR(i)
    Next i
    '(68)
    For i = 1 To NROWS + 1
        THETA_MATRIX(-1, i) = (TERMS_FALSE_ARR(i) - TERMS_FALSE_ARR(i - 1)) / (TERMS_FALSE_ARR(i) - TERMS_FALSE_ARR(i - 2)) * (DISCRETE_ARR(i) - DISCRETE_ARR(i - 1))
    Next i
    '(71)
    For i = -1 To NROWS - 1
        THETA_MATRIX(1, i) = (TERMS_FALSE_ARR(i + 1) - TERMS_FALSE_ARR(i)) / (TERMS_FALSE_ARR(i + 2) - TERMS_FALSE_ARR(i)) * (DISCRETE_ARR(i + 2) - DISCRETE_ARR(i + 1))
    Next i
    '(67)
    For i = 1 To NROWS
        If DISCRETE_ARR(i - 1) < DISCRETE_ARR(i) And DISCRETE_ARR(i) <= DISCRETE_ARR(i + 1) Then
            MIN_MAX_ARR(-1, 1, i) = MINIMUM_FUNC(DISCRETE_ARR(i) + 0.5 * THETA_MATRIX(-1, i), DISCRETE_ARR(i + 1))
            MIN_MAX_ARR(1, 1, i) = MINIMUM_FUNC(DISCRETE_ARR(i) + 2 * THETA_MATRIX(-1, i), DISCRETE_ARR(i + 1))
        ElseIf DISCRETE_ARR(i - 1) < DISCRETE_ARR(i) And DISCRETE_ARR(i) > DISCRETE_ARR(i + 1) Then
            MIN_MAX_ARR(-1, 1, i) = MAXIMUM_FUNC(DISCRETE_ARR(i) - 0.5 * LAMBDA * THETA_MATRIX(-1, i), DISCRETE_ARR(i + 1))
            MIN_MAX_ARR(1, 1, i) = DISCRETE_ARR(i)
        ElseIf DISCRETE_ARR(i - 1) >= DISCRETE_ARR(i) And DISCRETE_ARR(i) <= DISCRETE_ARR(i + 1) Then
            MIN_MAX_ARR(-1, 1, i) = DISCRETE_ARR(i)
            MIN_MAX_ARR(1, 1, i) = MINIMUM_FUNC(DISCRETE_ARR(i) - 0.5 * LAMBDA * THETA_MATRIX(-1, i), DISCRETE_ARR(i + 1))
        ElseIf DISCRETE_ARR(i - 1) >= DISCRETE_ARR(i) And DISCRETE_ARR(i) > DISCRETE_ARR(i + 1) Then
            MIN_MAX_ARR(-1, 1, i) = MAXIMUM_FUNC(DISCRETE_ARR(i) + 2 * THETA_MATRIX(-1, i), DISCRETE_ARR(i + 1))
            MIN_MAX_ARR(1, 1, i) = MAXIMUM_FUNC(DISCRETE_ARR(i) + 0.5 * THETA_MATRIX(-1, i), DISCRETE_ARR(i + 1))
        End If
    Next i
    '(70)
    For i = 0 To NROWS - 1
        If DISCRETE_ARR(i) < DISCRETE_ARR(i + 1) And DISCRETE_ARR(i + 1) <= DISCRETE_ARR(i + 2) Then
            MIN_MAX_ARR(-1, 2, i) = MAXIMUM_FUNC(DISCRETE_ARR(i + 1) - 2 * THETA_MATRIX(1, i), DISCRETE_ARR(i))
            MIN_MAX_ARR(1, 2, i) = MAXIMUM_FUNC(DISCRETE_ARR(i + 1) - 0.5 * THETA_MATRIX(1, i), DISCRETE_ARR(i))
        ElseIf DISCRETE_ARR(i) < DISCRETE_ARR(i + 1) And DISCRETE_ARR(i + 1) > DISCRETE_ARR(i + 2) Then
            MIN_MAX_ARR(-1, 2, i) = MAXIMUM_FUNC(DISCRETE_ARR(i + 1) + 0.5 * LAMBDA * THETA_MATRIX(1, i), DISCRETE_ARR(i))
            MIN_MAX_ARR(1, 2, i) = DISCRETE_ARR(i + 1)
        ElseIf DISCRETE_ARR(i) >= DISCRETE_ARR(i + 1) And DISCRETE_ARR(i + 1) < DISCRETE_ARR(i + 2) Then
            MIN_MAX_ARR(-1, 2, i) = DISCRETE_ARR(i + 1)
            MIN_MAX_ARR(1, 2, i) = MINIMUM_FUNC(DISCRETE_ARR(i + 1) + 0.5 * LAMBDA * THETA_MATRIX(1, i), DISCRETE_ARR(i))
        ElseIf DISCRETE_ARR(i) >= DISCRETE_ARR(i + 1) And DISCRETE_ARR(i + 1) >= DISCRETE_ARR(i + 2) Then
            MIN_MAX_ARR(-1, 2, i) = MINIMUM_FUNC(DISCRETE_ARR(i + 1) - 0.5 * THETA_MATRIX(1, i), DISCRETE_ARR(i))
            MIN_MAX_ARR(1, 2, i) = MINIMUM_FUNC(DISCRETE_ARR(i + 1) - 2 * THETA_MATRIX(1, i), DISCRETE_ARR(i))
        End If
    Next i
    For i = 1 To NROWS - 1
        If MAXIMUM_FUNC(MIN_MAX_ARR(-1, 1, i), MIN_MAX_ARR(-1, 2, i)) <= MINIMUM_FUNC(MIN_MAX_ARR(1, 1, i), MIN_MAX_ARR(1, 2, i)) Then
        '(75, 76)
            MIN_MAX_ARR(0, 0, i) = COLLAR_FUNC(MAXIMUM_FUNC(MIN_MAX_ARR(-1, 1, i), MIN_MAX_ARR(-1, 2, i)), MIN_MAX_ARR(0, 0, i), MINIMUM_FUNC(MIN_MAX_ARR(1, 1, i), MIN_MAX_ARR(1, 2, i)))
        Else
        '(78)
            MIN_MAX_ARR(0, 0, i) = COLLAR_FUNC(MINIMUM_FUNC(MIN_MAX_ARR(1, 1, i), MIN_MAX_ARR(1, 2, i)), MIN_MAX_ARR(0, 0, i), MAXIMUM_FUNC(MIN_MAX_ARR(-1, 1, i), MIN_MAX_ARR(-1, 2, i)))
        End If
    Next i
    '(79)
    If Abs(MIN_MAX_ARR(0, 0, 0) - DISCRETE_ARR(0)) > 0.5 * Abs(MIN_MAX_ARR(0, 0, 1) - DISCRETE_ARR(0)) Then MIN_MAX_ARR(0, 0, 0) = DISCRETE_ARR(1) - 0.5 * (MIN_MAX_ARR(0, 0, 1) - DISCRETE_ARR(0))
    '(80)
    If Abs(MIN_MAX_ARR(0, 0, NROWS) - DISCRETE_ARR(NROWS)) > 0.5 * Abs(MIN_MAX_ARR(0, 0, NROWS - 1) - DISCRETE_ARR(NROWS)) Then MIN_MAX_ARR(0, 0, NROWS) = DISCRETE_ARR(NROWS) - 0.5 * (MIN_MAX_ARR(0, 0, NROWS - 1) - DISCRETE_ARR(NROWS))
    If NEGATIVE_FLAG = False Then
    '(60)
        MIN_MAX_ARR(0, 0, 0) = COLLAR_FUNC(0, MIN_MAX_ARR(0, 0, 0), 2 * DISCRETE_ARR(1))
    '(61)
        For i = 1 To NROWS - 1
            MIN_MAX_ARR(0, 0, i) = COLLAR_FUNC(0, MIN_MAX_ARR(0, 0, i), 2 * MINIMUM_FUNC(DISCRETE_ARR(i), DISCRETE_ARR(i + 1)))
        Next i
    '(62)
        MIN_MAX_ARR(0, 0, NROWS) = COLLAR_FUNC(0, MIN_MAX_ARR(0, 0, NROWS), 2 * DISCRETE_ARR(NROWS))
    End If
    'finish, so populate the FORWARD_ARR array
    For i = 0 To NROWS
        FORWARD_ARR(i) = MIN_MAX_ARR(0, 0, i)
    Next i
End If

MONOTOME_CONVEX_ESTIMATES_FUNC = True

Exit Function
ERROR_LABEL:
MONOTOME_CONVEX_ESTIMATES_FUNC = False
End Function


'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : ARRAY_LAST_INDEX_FUNC
'DESCRIPTION   : determines the unique value of i
'LIBRARY       : BOND
'GROUP         : MONOTOME
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Private Function ARRAY_LAST_INDEX_FUNC(ByRef DATA_ARR As Variant, _
ByVal REFER_VAL As Variant) 'AS Long
  
  Dim k As Long 'iLastIndex
  
On Error GoTo ERROR_LABEL

k = CInt(COLLAR_FUNC(LBound(DATA_ARR), CDbl(k), UBound(DATA_ARR)))
Do
    If REFER_VAL >= DATA_ARR(k) Then
        If k = UBound(DATA_ARR) Then
            If REFER_VAL = DATA_ARR(k) Then
                ARRAY_LAST_INDEX_FUNC = UBound(DATA_ARR) - 1
                Exit Function
            Else
                ARRAY_LAST_INDEX_FUNC = UBound(DATA_ARR)
                Exit Function
            End If
        Else
            If REFER_VAL >= DATA_ARR(k + 1) Then
                k = k + 1
            Else
                ARRAY_LAST_INDEX_FUNC = k
                Exit Function
            End If
        End If
    Else
        If k = LBound(DATA_ARR) Then
            ARRAY_LAST_INDEX_FUNC = LBound(DATA_ARR) - 1
            Exit Function
        Else
            If REFER_VAL >= DATA_ARR(k - 1) Then
                ARRAY_LAST_INDEX_FUNC = k - 1
                Exit Function
            Else
                k = k - 1
            End If
        End If
    End If
Loop

Exit Function
ERROR_LABEL:
ARRAY_LAST_INDEX_FUNC = Err.number
End Function
