Attribute VB_Name = "EXCEL_CHART_SHAPES_LIBR"


'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------

' Subroutine to insert a picture within a cell based on

Function SHAPE_ADD_RNG_FUNC(ByVal SRC_PICTURE_STR As String, _
Optional ByVal OUTPUT As Integer = 0)

Dim TOP_VAL As Double
Dim LEFT_VAL As Double
Dim WIDTH_VAL As Double
Dim HEIGHT_VAL As Double

Dim SRC_SHAPE As Excel.Shape
Dim CALLER_RNG As Excel.Range
Dim VALID_FLAG As Boolean

On Error GoTo ERROR_LABEL

VALID_FLAG = SHAPE_CALLER_RNG_FUNC(CALLER_RNG, HEIGHT_VAL, WIDTH_VAL, LEFT_VAL, TOP_VAL)
If VALID_FLAG = False Then: GoTo ERROR_LABEL
With CALLER_RNG.Worksheet.Shapes
    Set SRC_SHAPE = .AddShape(msoShapeRectangle, LEFT_VAL + 5, TOP_VAL + 5, WIDTH_VAL - 10, HEIGHT_VAL - 10)
End With

With SRC_SHAPE
    .Fill.UserPicture SRC_PICTURE_STR
    .AlternativeText = SRC_PICTURE_STR
    .OnAction = "SHAPE_TOGGLE_FUNC"
End With

Select Case OUTPUT
Case 0
    SHAPE_ADD_RNG_FUNC = SRC_PICTURE_STR
Case 1
    SHAPE_ADD_RNG_FUNC = True
Case Else
    Set SHAPE_ADD_RNG_FUNC = SRC_SHAPE
End Select

Exit Function
ERROR_LABEL:
Select Case OUTPUT
Case 0, 1
    SHAPE_ADD_RNG_FUNC = False
Case Else
    Set SHAPE_ADD_RNG_FUNC = Nothing
End Select
End Function


'Subroutine to toggle shape (i.e. zoom in/zoom out) within a cell

Sub SHAPE_TOGGLE_FUNC()
    
Dim TOP_VAL As Double
Dim LEFT_VAL As Double
Dim WIDTH_VAL As Double
Dim HEIGHT_VAL As Double

Dim SRC_SHAPE As Excel.Shape
Dim SRC_RANGE As Excel.Range
Dim SRC_WSHEET As Excel.Worksheet

On Error GoTo ERROR_LABEL

Set SRC_WSHEET = ActiveSheet
Set SRC_SHAPE = SRC_WSHEET.Shapes(Excel.Application.Caller)
Set SRC_RANGE = SRC_SHAPE.BottomRightCell
    
TOP_VAL = SRC_SHAPE.Top
LEFT_VAL = SRC_SHAPE.Left
WIDTH_VAL = SRC_SHAPE.Width
HEIGHT_VAL = SRC_SHAPE.Height

If TOP_VAL = SRC_RANGE.Top + 5 And _
   LEFT_VAL = SRC_RANGE.Left + 5 And _
   WIDTH_VAL = SRC_RANGE.Width - 10 And _
   HEIGHT_VAL = SRC_RANGE.Height - 10 Then
        SRC_SHAPE.Width = SRC_RANGE.Width * 0.25
        SRC_SHAPE.Height = SRC_RANGE.Height * 0.25
Else
    SRC_SHAPE.Top = SRC_RANGE.Top + 5
    SRC_SHAPE.Left = SRC_RANGE.Left + 5
    SRC_SHAPE.Width = SRC_RANGE.Width - 10
    SRC_SHAPE.Height = SRC_RANGE.Height - 10
End If

Exit Sub
ERROR_LABEL:
'ADD MSG HERE; Err.Description
End Sub


'Function to create "in cell" charts -- line charts, bar charts, or
'slope of linear regression

Function SHAPE_TREND_RNG_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef VERSION As Integer = 2, _
Optional ByRef COLOR_INDEX As Integer = 203, _
Optional ByRef MARGIN_VAL As Integer = 2, _
Optional ByRef GAP_VAL As Integer = 1)
    
'GAP_VAL --> Size of gap to use between bar charts
'MARGIN_VAL -- >A margin to buffer the usable cell area
    
Dim i As Integer
Dim NSIZE As Integer

Dim X1_VAL As Double
Dim Y1_VAL As Double

Dim X2_VAL As Double
Dim Y2_VAL As Double

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim MULT_VAL As Double

Dim ATOP_VAL As Double
Dim BTOP_VAL As Double

Dim ALEFT_VAL As Double
Dim BLEFT_VAL As Double

Dim AWIDTH_VAL As Double
Dim BWIDTH_VAL As Double

Dim AHEIGHT_VAL As Double
Dim BHEIGHT_VAL As Double
Dim CALLER_RNG As Excel.Range ' The calling range for the function
 
Dim VALID_FLAG As Boolean
Dim TREND_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

SHAPE_TREND_RNG_FUNC = False

VALID_FLAG = SHAPE_CALLER_RNG_FUNC(CALLER_RNG, _
             AHEIGHT_VAL, AWIDTH_VAL, ALEFT_VAL, ATOP_VAL)
If VALID_FLAG = False Then: GoTo ERROR_LABEL

With CALLER_RNG.Font
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = 0
End With

'Copy input range/array to standard array area
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: _
DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
NSIZE = UBound(DATA_VECTOR, 1)
    
'Determine type of chart to create
    
'----------------------------------------------------------------------------------
Select Case VERSION
'----------------------------------------------------------------------------------
Case 0 ' Create a Bar Chart
'----------------------------------------------------------------------------------
    'Determine minimum and maximum chartable values
    MIN_VAL = Excel.Application.WorksheetFunction.Min(DATA_VECTOR)
    MAX_VAL = Excel.Application.WorksheetFunction.max(DATA_VECTOR)
    If MIN_VAL > 0 Then MIN_VAL = 0
    If MIN_VAL = MAX_VAL Then
       MIN_VAL = MIN_VAL - 1
       MAX_VAL = MAX_VAL + 1
    End If

    'Draw the bar for each data point
    With CALLER_RNG.Worksheet.Shapes
         For i = 0 To NSIZE - 1
             MULT_VAL = (AHEIGHT_VAL - (MARGIN_VAL * 2)) / (MAX_VAL - MIN_VAL)
             BLEFT_VAL = MARGIN_VAL + GAP_VAL + ALEFT_VAL + _
                        (i * (AWIDTH_VAL - (MARGIN_VAL * 2)) / NSIZE)
             BTOP_VAL = MARGIN_VAL + ATOP_VAL + (MAX_VAL - _
                        IIf(DATA_VECTOR(i + 1, 1) < 0, 0, DATA_VECTOR(i + 1, 1))) * MULT_VAL
             BWIDTH_VAL = (AWIDTH_VAL - (MARGIN_VAL * 2)) / NSIZE - (GAP_VAL * 2)
             BHEIGHT_VAL = Abs(DATA_VECTOR(i + 1, 1)) * MULT_VAL
             With .AddShape(msoShapeRectangle, BLEFT_VAL, BTOP_VAL, _
                    BWIDTH_VAL, BHEIGHT_VAL)
                  If COLOR_INDEX > 0 Then _
                  .Fill.ForeColor.RGB = COLOR_INDEX _
                  Else .Fill.ForeColor.SchemeColor = -COLOR_INDEX
             End With
         Next i
    End With
       
'----------------------------------------------------------------------------------
Case 1 'Create a line chart
'----------------------------------------------------------------------------------
    'Determine minimum and maximum chartable values
    MIN_VAL = Excel.Application.WorksheetFunction.Min(DATA_VECTOR)
    MAX_VAL = Excel.Application.WorksheetFunction.max(DATA_VECTOR)
    If MIN_VAL = MAX_VAL Then
       MIN_VAL = MIN_VAL - 1
       MAX_VAL = MAX_VAL + 1
    End If
    
    'Draw the lines for each pair of data points
    With CALLER_RNG.Worksheet.Shapes
         For i = 0 To NSIZE - 2
             X1_VAL = MARGIN_VAL + ALEFT_VAL + (i * (AWIDTH_VAL - _
                    (MARGIN_VAL * 2)) / (NSIZE - 1))
             Y1_VAL = MARGIN_VAL + ATOP_VAL + (MAX_VAL - _
                    DATA_VECTOR(i + 1, 1)) * (AHEIGHT_VAL - _
                    (MARGIN_VAL * 2)) / (MAX_VAL - MIN_VAL)
             X2_VAL = MARGIN_VAL + ALEFT_VAL + ((i + 1) * _
                    (AWIDTH_VAL - (MARGIN_VAL * 2)) / (NSIZE - 1))
             Y2_VAL = MARGIN_VAL + ATOP_VAL + (MAX_VAL - _
                    DATA_VECTOR(i + 2, 1)) * (AHEIGHT_VAL - _
                    (MARGIN_VAL * 2)) / (MAX_VAL - MIN_VAL)
             With .AddLine(X1_VAL, Y1_VAL, X2_VAL, Y2_VAL)
                  If COLOR_INDEX > 0 Then _
                  .Line.ForeColor.RGB = COLOR_INDEX Else _
                  .Line.ForeColor.SchemeColor = -COLOR_INDEX
             End With
          Next i
    End With
       
'----------------------------------------------------------------------------------
Case Else 'Create a chart of a linear regression slope line
'----------------------------------------------------------------------------------
    'Create linear regression trend line
    TREND_VECTOR = Excel.Application.WorksheetFunction.Trend(DATA_VECTOR)
    'Determine minimum and maximum chartable values
    MIN_VAL = Excel.Application.WorksheetFunction.Min(DATA_VECTOR, TREND_VECTOR)
    MAX_VAL = Excel.Application.WorksheetFunction.max(DATA_VECTOR, TREND_VECTOR)
    If MIN_VAL = MAX_VAL Then
       MIN_VAL = MIN_VAL - 1
       MAX_VAL = MAX_VAL + 1
    End If
    
    'Draw the regression line
    With CALLER_RNG.Worksheet.Shapes
         X1_VAL = MARGIN_VAL + ALEFT_VAL
         Y1_VAL = MARGIN_VAL + ATOP_VAL + (MAX_VAL - _
                TREND_VECTOR(1, 1)) * (AHEIGHT_VAL - _
                (MARGIN_VAL * 2)) / (MAX_VAL - MIN_VAL)
         X2_VAL = ALEFT_VAL + AWIDTH_VAL - MARGIN_VAL
         Y2_VAL = MARGIN_VAL + ATOP_VAL + _
                (MAX_VAL - TREND_VECTOR(NSIZE, 1)) * _
                (AHEIGHT_VAL - (MARGIN_VAL * 2)) / (MAX_VAL - MIN_VAL)
         With .AddLine(X1_VAL, Y1_VAL, X2_VAL, Y2_VAL)
              If COLOR_INDEX > 0 Then _
                .Line.ForeColor.RGB = COLOR_INDEX Else _
                .Line.ForeColor.SchemeColor = -COLOR_INDEX
                .Line.BeginArrowheadStyle = msoArrowheadOval
                .Line.BeginArrowheadLength = msoArrowheadShort
                .Line.BeginArrowheadWidth = msoArrowheadNarrow
                .Line.EndArrowheadStyle = msoArrowheadStealth
         End With
    End With

'----------------------------------------------------------------------------------
End Select
'----------------------------------------------------------------------------------

SHAPE_TREND_RNG_FUNC = True
Exit Function
ERROR_LABEL:
SHAPE_TREND_RNG_FUNC = False
End Function

Function SHAPE_CALLER_RNG_FUNC(ByRef CALLER_RNG As Excel.Range, _
Optional ByRef HEIGHT_VAL As Double, _
Optional ByRef WIDTH_VAL As Double, _
Optional ByRef LEFT_VAL As Double, _
Optional ByRef TOP_VAL As Double)
    
Dim TEMP_RNG As Excel.Range
Dim SHAPE_OBJ As Excel.Shape
    
On Error GoTo ERROR_LABEL
    
'Identify the calling range
    
SHAPE_CALLER_RNG_FUNC = False

Set CALLER_RNG = Excel.Application.Caller
HEIGHT_VAL = CALLER_RNG.MergeArea.Height
WIDTH_VAL = CALLER_RNG.MergeArea.Width
LEFT_VAL = CALLER_RNG.MergeArea.Left
TOP_VAL = CALLER_RNG.MergeArea.Top
 
 'Delete existing shapes in the calling range

On Error Resume Next
    
For Each SHAPE_OBJ In CALLER_RNG.Worksheet.Shapes
    Set TEMP_RNG = Intersect(Range(SHAPE_OBJ.TopLeftCell, _
                   SHAPE_OBJ.BottomRightCell), CALLER_RNG.MergeArea)
    If Not TEMP_RNG Is Nothing Then
        If TEMP_RNG.Address = Range(SHAPE_OBJ.TopLeftCell, _
           SHAPE_OBJ.BottomRightCell).Address Then SHAPE_OBJ.Delete
    End If
Next SHAPE_OBJ

SHAPE_CALLER_RNG_FUNC = True

Exit Function
ERROR_LABEL:
SHAPE_CALLER_RNG_FUNC = False
End Function

Function SHAPE_DRAW_POINT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal FACTOR_VAL As Double = 1, _
Optional ByRef DST_WSHEET As Excel.Worksheet)

Dim i As Long
Dim NSIZE As Long

Dim NAME_STR As String

Dim D_VAL As Double
Dim U_VAL As Double
Dim V_VAL As Double

Dim XSCALE_VAL As Double
Dim YSCALE_VAL As Double

Dim ASHAPE_OBJ As Object
Dim BSHAPE_OBJ As Object

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG

If DST_WSHEET Is Nothing Then: Set DST_WSHEET = ActiveSheet
NSIZE = UBound(DATA_VECTOR)

With DST_WSHEET.Shapes
    XSCALE_VAL = DATA_VECTOR(1, 1) - FACTOR_VAL / 2
    YSCALE_VAL = DATA_VECTOR(1, 2) - FACTOR_VAL / 2
    Set ASHAPE_OBJ = .AddShape(msoShapeOval, CSng(XSCALE_VAL), _
                        CSng(YSCALE_VAL), CSng(FACTOR_VAL), CSng(FACTOR_VAL))
    ASHAPE_OBJ.Fill.ForeColor.SchemeColor = 8
    NAME_STR = ASHAPE_OBJ.name
    U_VAL = DATA_VECTOR(1, 1)
    V_VAL = DATA_VECTOR(1, 2)
    For i = 2 To NSIZE
        D_VAL = Sqr((DATA_VECTOR(i, 1) - U_VAL) ^ 2 + (DATA_VECTOR(i, 2) - V_VAL) ^ 2)
        If D_VAL > 1.5 Then
            XSCALE_VAL = DATA_VECTOR(i, 1) - FACTOR_VAL / 2
            YSCALE_VAL = DATA_VECTOR(i, 2) - FACTOR_VAL / 2
            Set ASHAPE_OBJ = .AddShape(msoShapeOval, _
                            CSng(XSCALE_VAL), CSng(YSCALE_VAL), _
                            CSng(FACTOR_VAL), CSng(FACTOR_VAL))
            ASHAPE_OBJ.Fill.ForeColor.SchemeColor = 8
            DST_WSHEET.Shapes.Range(Array(NAME_STR, _
                    ASHAPE_OBJ.name)).Select
            Set BSHAPE_OBJ = Selection.ShapeRange.Group
            NAME_STR = BSHAPE_OBJ.name
            'NAME_STR = Selection.ShapeRange.Name

            U_VAL = DATA_VECTOR(i, 1)
            V_VAL = DATA_VECTOR(i, 2)
        End If
    Next i
End With

SHAPE_DRAW_POINT_FUNC = NAME_STR

Exit Function
ERROR_LABEL:
SHAPE_DRAW_POINT_FUNC = Err.number
End Function

Function SHAPE_DRAW_MESH_FUNC(ByRef DATA_RNG As Variant, _
ByVal NROWS As Long, _
ByVal NCOLUMNS As Long, _
Optional ByVal MATCH_FLAG As Variant = msoTrue, _
Optional ByRef DST_WSHEET As Excel.Worksheet)

Dim i As Long
Dim j As Long

Dim U_VAL As Double
Dim V_VAL As Double

Dim XSCALE_VAL As Double
Dim YSCALE_VAL As Double

Dim NAME_STR As String

Dim ASHAPE_OBJ As Object
Dim BSHAPE_OBJ As Object
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If DST_WSHEET Is Nothing Then: Set DST_WSHEET = ActiveSheet

XSCALE_VAL = (DATA_VECTOR(2, 1) - DATA_VECTOR(1, 1)) / NROWS
YSCALE_VAL = (DATA_VECTOR(2, 2) - DATA_VECTOR(1, 2)) / NCOLUMNS

With DST_WSHEET.Shapes
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            U_VAL = (i - 1) * XSCALE_VAL + DATA_VECTOR(1, 1)
            V_VAL = (j - 1) * YSCALE_VAL + DATA_VECTOR(1, 2)
            
            Set ASHAPE_OBJ = .AddShape(msoShapeRectangle, CSng(U_VAL), _
                    CSng(V_VAL), CSng(XSCALE_VAL), CSng(YSCALE_VAL))

            If MATCH_FLAG = msoFalse Then ASHAPE_OBJ.Line.Visible = msoFalse
            If i = 1 And j = 1 Then
                NAME_STR = ASHAPE_OBJ.name
            Else
                DST_WSHEET.Shapes.Range(Array(NAME_STR, _
                    ASHAPE_OBJ.name)).Select
                Set BSHAPE_OBJ = Selection.ShapeRange.Group
                NAME_STR = BSHAPE_OBJ.name
            End If
        Next j
    Next i
End With

'SHAPE_DRAW_MESH_FUNC = NAME_STR
SHAPE_DRAW_MESH_FUNC = U_VAL

Exit Function
ERROR_LABEL:
SHAPE_DRAW_MESH_FUNC = Err.number
End Function

Function SHAPE_DRAW_CURVE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef DST_WSHEET As Excel.Worksheet)

Dim i As Long
Dim NSIZE As Long

Dim D_VAL As Double
Dim U_VAL As Double
Dim V_VAL As Double

Dim SHAPE_OBJ As Object
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If DST_WSHEET Is Nothing Then: Set DST_WSHEET = ActiveSheet
NSIZE = UBound(DATA_VECTOR)

With DST_WSHEET.Shapes.BuildFreeform(msoEditingAuto, DATA_VECTOR(1, 1), DATA_VECTOR(1, 2))
    U_VAL = DATA_VECTOR(1, 1)
    V_VAL = DATA_VECTOR(1, 2)
    For i = 2 To NSIZE
        D_VAL = Sqr((DATA_VECTOR(i, 1) - U_VAL) ^ 2 + _
                (DATA_VECTOR(i, 2) - V_VAL) ^ 2)
        
        If D_VAL > 1.5 Then
            .AddNodes msoSegmentLine, msoEditingAuto, _
                DATA_VECTOR(i, 1), DATA_VECTOR(i, 2)
            U_VAL = DATA_VECTOR(i, 1)
            V_VAL = DATA_VECTOR(i, 2)
        End If
    Next i

    Set SHAPE_OBJ = .ConvertToShape
    SHAPE_OBJ.Line.ForeColor.RGB = RGB(119, 119, 119)
    SHAPE_OBJ.Fill.Visible = msoFalse
End With

SHAPE_DRAW_CURVE_FUNC = SHAPE_OBJ.name

Exit Function
ERROR_LABEL:
SHAPE_DRAW_CURVE_FUNC = Err.number
End Function

Function SHAPE_DRAW_INTENSITY_FUNC(ByRef DATA_RNG As Variant, _
ByVal RSHAPE_OBJ As Object, _
ByVal MIN_VAL As Double, _
ByVal MAX_VAL As Double, _
Optional ByRef DST_WSHEET As Excel.Worksheet)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim RGB_VAL As Double
Dim SHAPE_OBJ As Object 'Shape
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

SHAPE_DRAW_INTENSITY_FUNC = False

DATA_VECTOR = DATA_RNG
If DST_WSHEET Is Nothing Then: Set DST_WSHEET = ActiveSheet

Set SHAPE_OBJ = DST_WSHEET.Shapes(RSHAPE_OBJ)

NSIZE = SHAPE_OBJ.GroupItems.COUNT
NROWS = UBound(DATA_VECTOR, 1)
NCOLUMNS = UBound(DATA_VECTOR, 2)
    
For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        k = (i - 1) * NCOLUMNS + j
        RGB_VAL = 255 * DATA_VECTOR(i, j) / (MAX_VAL - MIN_VAL)
        SHAPE_OBJ.GroupItems(k).Fill.ForeColor.RGB = RGB(RGB_VAL, RGB_VAL, RGB_VAL)
    Next j
Next i

SHAPE_DRAW_INTENSITY_FUNC = True

Exit Function
ERROR_LABEL:
SHAPE_DRAW_INTENSITY_FUNC = False
End Function

Function SHAPE_XAXES_DRAW_FUNC(ByRef BOX_RNG As Variant, _
ByRef PLANE_RNG As Variant, _
Optional ByRef DST_WSHEET As Excel.Worksheet)

Dim X_VAL As Double
Dim Y_VAL As Double
Dim NAME_STR As String

Dim BOX_VECTOR As Variant
Dim PLANE_VECTOR As Variant

Dim DATA_VECTOR As Variant
Dim SHAPE_OBJ As Object

On Error GoTo ERROR_LABEL

BOX_VECTOR = BOX_RNG
PLANE_VECTOR = PLANE_RNG

If DST_WSHEET Is Nothing Then: Set DST_WSHEET = ActiveSheet

If PLANE_VECTOR(1, 2) < 0 And PLANE_VECTOR(2, 2) > 0 Then
    
    ReDim DATA_VECTOR(1 To 2, 1 To 2)
    Y_VAL = 0 ' '(PLANE_VECTOR(1, 2) + PLANE_VECTOR(2, 2)) / 2
    X_VAL = PLANE_VECTOR(1, 1)
    
    Call SHAPE_TRANSFORM_VECTOR_FUNC(X_VAL, Y_VAL, _
                DATA_VECTOR(1, 1), DATA_VECTOR(1, 2), _
                PLANE_VECTOR, BOX_VECTOR)
    X_VAL = PLANE_VECTOR(2, 1)
    
    Call SHAPE_TRANSFORM_VECTOR_FUNC(X_VAL, Y_VAL, _
                DATA_VECTOR(2, 1), DATA_VECTOR(2, 2), _
                PLANE_VECTOR, BOX_VECTOR)
    
    With DST_WSHEET.Shapes 'draw a line between point p1-p2
        Set SHAPE_OBJ = .AddLine(DATA_VECTOR(1, 1), _
                        DATA_VECTOR(1, 2), DATA_VECTOR(2, 1), _
                        DATA_VECTOR(2, 2))
        SHAPE_OBJ.Fill.ForeColor.SchemeColor = 8
    End With
    NAME_STR = SHAPE_OBJ.name
    
    DST_WSHEET.Shapes(NAME_STR).Line.EndArrowheadStyle = msoArrowheadTriangle
    SHAPE_XAXES_DRAW_FUNC = NAME_STR
Else
    SHAPE_XAXES_DRAW_FUNC = ""
End If

Exit Function
ERROR_LABEL:
SHAPE_XAXES_DRAW_FUNC = Err.number
End Function


Function SHAPE_Y_AXES_DRAW_FUNC(ByRef BOX_RNG As Variant, _
ByRef PLANE_RNG As Variant, _
Optional ByRef DST_WSHEET As Excel.Worksheet)

Dim X_VAL As Double
Dim Y_VAL As Double

Dim NAME_STR As String

Dim SHAPE_OBJ As Object

Dim BOX_VECTOR As Variant
Dim PLANE_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

BOX_VECTOR = BOX_RNG
PLANE_VECTOR = PLANE_RNG
If DST_WSHEET Is Nothing Then: Set DST_WSHEET = ActiveSheet

If PLANE_VECTOR(1, 1) < 0 And PLANE_VECTOR(2, 1) > 0 Then
    
    ReDim DATA_VECTOR(1 To 2, 1 To 2)
    
    X_VAL = 0 ' (PLANE_VECTOR(1, 1) + PLANE_VECTOR(2, 1)) / 2
    Y_VAL = PLANE_VECTOR(1, 2)
    
    Call SHAPE_TRANSFORM_VECTOR_FUNC(X_VAL, Y_VAL, DATA_VECTOR(1, 1), DATA_VECTOR(1, 2), PLANE_VECTOR, BOX_VECTOR)
    Y_VAL = PLANE_VECTOR(2, 2)
    Call SHAPE_TRANSFORM_VECTOR_FUNC(X_VAL, Y_VAL, DATA_VECTOR(2, 1), DATA_VECTOR(2, 2), PLANE_VECTOR, BOX_VECTOR)
    With DST_WSHEET.Shapes 'draw a line between point p1-p2
        Set SHAPE_OBJ = .AddLine(DATA_VECTOR(1, 1), DATA_VECTOR(1, 2), DATA_VECTOR(2, 1), DATA_VECTOR(2, 2))
        SHAPE_OBJ.Fill.ForeColor.SchemeColor = 8
    End With
    NAME_STR = SHAPE_OBJ.name
    DST_WSHEET.Shapes(NAME_STR).Line.EndArrowheadStyle = msoArrowheadTriangle
    SHAPE_Y_AXES_DRAW_FUNC = NAME_STR
Else
    SHAPE_Y_AXES_DRAW_FUNC = ""
End If
Exit Function
ERROR_LABEL:
SHAPE_Y_AXES_DRAW_FUNC = Err.number
End Function

Function SHAPE_Z_AXES_DRAW_FUNC(ByRef BOX_RNG As Variant, _
ByRef PLANE_RNG As Variant, _
Optional ByRef DST_WSHEET As Excel.Worksheet)

Dim X_VAL As Double
Dim Y_VAL As Double

Dim NAME_STR As String

Dim BOX_VECTOR As Variant
Dim PLANE_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

BOX_VECTOR = BOX_RNG
PLANE_VECTOR = PLANE_RNG

'X_VAL axe
If PLANE_VECTOR(1, 2) < 0 And PLANE_VECTOR(2, 2) > 0 Then
    ReDim DATA_VECTOR(1 To 3, 1 To 2)
    Y_VAL = 0 ' '(PLANE_VECTOR(1, 2) + PLANE_VECTOR(2, 2)) / 2
    X_VAL = PLANE_VECTOR(1, 1)
    Call SHAPE_TRANSFORM_VECTOR_FUNC(X_VAL, Y_VAL, DATA_VECTOR(1, 1), DATA_VECTOR(1, 2), PLANE_VECTOR, BOX_VECTOR)
    X_VAL = (PLANE_VECTOR(1, 1) + PLANE_VECTOR(2, 1)) / 2
    Call SHAPE_TRANSFORM_VECTOR_FUNC(X_VAL, Y_VAL, DATA_VECTOR(2, 1), DATA_VECTOR(2, 2), PLANE_VECTOR, BOX_VECTOR)
    X_VAL = PLANE_VECTOR(2, 1)
    Call SHAPE_TRANSFORM_VECTOR_FUNC(X_VAL, Y_VAL, DATA_VECTOR(3, 1), DATA_VECTOR(3, 2), PLANE_VECTOR, BOX_VECTOR)
    NAME_STR = SHAPE_DRAW_CURVE_FUNC(DATA_VECTOR, DST_WSHEET)
    SHAPE_Z_AXES_DRAW_FUNC = NAME_STR
Else
    SHAPE_Z_AXES_DRAW_FUNC = ""
End If

Exit Function
ERROR_LABEL:
SHAPE_Z_AXES_DRAW_FUNC = Err.number
End Function


Function SHAPE_TRANSFORM_VECTOR_FUNC(ByVal X_VAL As Double, _
ByVal Y_VAL As Double, _
ByRef U_VAL As Double, _
ByRef V_VAL As Double, _
ByRef DATA1_ARR As Variant, _
ByRef DATA2_ARR As Variant)

On Error GoTo ERROR_LABEL

SHAPE_TRANSFORM_VECTOR_FUNC = False

U_VAL = (X_VAL - DATA1_ARR(1, 1)) / (DATA1_ARR(2, 1) - DATA1_ARR(1, 1)) * (DATA2_ARR(2, 1) - DATA2_ARR(1, 1)) + DATA2_ARR(1, 1)
V_VAL = DATA2_ARR(2, 2) - (Y_VAL - DATA1_ARR(1, 2)) / (DATA1_ARR(2, 2) - DATA1_ARR(1, 2)) * (DATA2_ARR(2, 2) - DATA2_ARR(1, 2))

'BOX_VECTOR constraining
If U_VAL < DATA2_ARR(1, 1) Then U_VAL = DATA2_ARR(1, 1)
If U_VAL > DATA2_ARR(2, 1) Then U_VAL = DATA2_ARR(2, 1)
If V_VAL < DATA2_ARR(1, 2) Then V_VAL = DATA2_ARR(1, 2)
If V_VAL > DATA2_ARR(2, 2) Then V_VAL = DATA2_ARR(2, 2)

SHAPE_TRANSFORM_VECTOR_FUNC = True

Exit Function
ERROR_LABEL:
SHAPE_TRANSFORM_VECTOR_FUNC = False
End Function


'MsoAutoShapeType can be one of these MsoAutoShapeType constants:

'msoShape16pointStar
'msoShape24pointStar
'msoShape32pointStar
'msoShape4pointStar
'msoShape5pointStar
'msoShape8pointStar
'msoShapeActionButtonBackorPrevious
'msoShapeActionButtonBeginning
'msoShapeActionButtonCustom
'msoShapeActionButtonDocument
'msoShapeActionButtonEnd
'msoShapeActionButtonForwardorNext
'msoShapeActionButtonHelp
'msoShapeActionButtonHome
'msoShapeActionButtonInformation
'msoShapeActionButtonMovie
'msoShapeActionButtonReturn
'msoShapeActionButtonSound
'msoShapeArc
'msoShapeBalloon
'msoShapeBentArrow
'msoShapeBentUpArrow
'msoShapeBevel
'msoShapeBlockArc
'msoShapeCan
'msoShapeChevron
'msoShapeCircularArrow
'msoShapeCloudCallout
'msoShapeCross
'msoShapeCube
'msoShapeCurvedDownArrow
'msoShapeCurvedDownRibbon
'msoShapeCurvedLeftArrow
'msoShapeCurvedRightArrow
'msoShapeCurvedUpArrow
'msoShapeCurvedUpRibbon
'msoShapeDiamond
'msoShapeDonut
'msoShapeDoubleBrace
'msoShapeDoubleBracket
'msoShapeDoubleWave
'msoShapeDownArrow
'msoShapeDownArrowCallout
'msoShapeDownRibbon
'msoShapeExplosion1
'msoShapeExplosion2
'msoShapeFlowchartAlternateProcess
'msoShapeFlowchartCard
'msoShapeFlowchartCollate
'msoShapeFlowchartConnector
'msoShapeFlowchartData
'msoShapeFlowchartDecision
'msoShapeFlowchartDelay
'msoShapeFlowchartDirectAccessStorage
'msoShapeFlowchartDisplay
'msoShapeFlowchartDocument
'msoShapeFlowchartExtract
'msoShapeFlowchartInternalStorage
'msoShapeFlowchartMagneticDisk
'msoShapeFlowchartManualInput
'msoShapeFlowchartManualOperation
'msoShapeFlowchartMerge
'msoShapeFlowchartMultidocument
'msoShapeFlowchartOffpageConnector
'msoShapeFlowchartOr
'msoShapeFlowchartPredefinedProcess
'msoShapeFlowchartPreparation
'msoShapeFlowchartProcess
'msoShapeFlowchartPunchedTape
'msoShapeFlowchartSequentialAccessStorage
'msoShapeFlowchartSort
'msoShapeFlowchartStoredData
'msoShapeFlowchartSummingJunction
'msoShapeFlowchartTerminator
'msoShapeFoldedCorner
'msoShapeHeart
'msoShapeHexagon
'msoShapeHorizontalScroll
'msoShapeIsoscelesTriangle
'msoShapeLeftArrow
'msoShapeLeftArrowCallout
'msoShapeLeftBrace
'msoShapeLeftBracket
'msoShapeLeftRightArrow
'msoShapeLeftRightArrowCallout
'msoShapeLeftRightUpArrow
'msoShapeLeftUpArrow
'msoShapeLightningBolt
'msoShapeLineCallout1
'msoShapeLineCallout1AccentBar
'msoShapeLineCallout1BorderandAccentBar
'msoShapeLineCallout1NoBorder
'msoShapeLineCallout2
'msoShapeLineCallout2AccentBar
'msoShapeLineCallout2BorderandAccentBar
'msoShapeLineCallout2NoBorder
'msoShapeLineCallout3
'msoShapeLineCallout3AccentBar
'msoShapeLineCallout3BorderandAccentBar
'msoShapeLineCallout3NoBorder
'msoShapeLineCallout4
'msoShapeLineCallout4AccentBar
'msoShapeLineCallout4BorderandAccentBar
'msoShapeLineCallout4NoBorder
'msoShapeMixed
'msoShapeMoon
'msoShapeNoSymbol
'msoShapeNotchedRightArrow
'msoShapeNotPrimitive
'msoShapeOctagon
'msoShapeOval
'msoShapeOvalCallout
'msoShapeParallelogram
'msoShapePentagon
'msoShapePlaque
'msoShapeQuadArrow
'msoShapeQuadArrowCallout
'msoShapeRectangle
'msoShapeRectangularCallout
'msoShapeRegularPentagon
'msoShapeRightArrow
'msoShapeRightArrowCallout
'msoShapeRightBrace
'msoShapeRightBracket
'msoShapeRightTriangle
'msoShapeRoundedRectangle
'msoShapeRoundedRectangularCallout
'msoShapeSmileyFace
'msoShapeStripedRightArrow
'msoShapeSun
'msoShapeTrapezoid
'msoShapeUpArrow
'msoShapeUpArrowCallout
'msoShapeUpDownArrow
'msoShapeUpDownArrowCallout
'msoShapeUpRibbon
'msoShapeUTurnArrow
'msoShapeVerticalScroll
'msoShapeWave
