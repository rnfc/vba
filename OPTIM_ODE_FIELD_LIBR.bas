Attribute VB_Name = "OPTIM_ODE_FIELD_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : ODE_FIELD_GENERIC_FUNC

'DESCRIPTION   : This routine generate and visualize the slope y'(x) of a
'1st ODE solution over a rectangular domain. It is didactically useful for
'studying the 1st order differential equation.

'The algorithm generates the graphical object showing the slope y' of each point of
'the grid. Every solution y(x) of the differential equation follows the direction
'of the slope field. Therefore the field gives a global view of the solutions
'crossing the given domain.

'LIBRARY       : OPTIMIZATION
'GROUP         : ODE_FIELD
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ODE_FIELD_GENERIC_FUNC(ByVal FUNC_NAME_STR As String, _
Optional ByVal XMIN_VAL As Double = -2, _
Optional ByVal XMAX_VAL As Double = 2, _
Optional ByVal YMIN_VAL As Double = -2, _
Optional ByVal YMAX_VAL As Double = 2, _
Optional ByVal NSIZE As Long = 15, _
Optional ByVal AXES_FLAG As Boolean = True, _
Optional ByVal DOT_FLAG As Boolean = True, _
Optional ByVal DST_WSHEET As Excel.Worksheet)

'NSIZE = Grid number, from 4 to 24, set the density of the grid.
'DOT_FLAG =  adds the grid points to the plot
'AXES_FLAG = adds the x-y axes

Dim i As Long
Dim j As Long
Dim k As Long

Dim X1_VAL As Double
Dim X2_VAL As Double

Dim Y1_VAL As Double
Dim Y2_VAL As Double

Dim DX_VAL As Double
Dim DY_VAL As Double

Dim HX_VAL As Double
Dim HY_VAL As Double

Dim TETA_VAL As Double
Dim DELTA_VAL As Double

Dim XTEMP_VAL As Double
Dim YTEMP_VAL As Double

Dim FUNC_VAL As Variant

Dim TEMP1_STR As String
Dim TEMP2_STR As String

Dim TEMP_GROUP As Variant
Dim TEMP_MATRIX As Variant
Dim PARAM_VECTOR As Variant

Dim CONST_BOX As Variant
Dim PLANE_BOX As Variant

On Error GoTo ERROR_LABEL

ODE_FIELD_GENERIC_FUNC = False

If DST_WSHEET Is Nothing Then: Set DST_WSHEET = ActiveSheet

ReDim CONST_BOX(1 To 2, 1 To 2)
ReDim PLANE_BOX(1 To 2, 1 To 2)

If NSIZE > 25 Or NSIZE < 4 Then: GoTo ERROR_LABEL
'grid range must be > 4 and < 25

CONST_BOX(1, 1) = 10
CONST_BOX(1, 2) = 10
CONST_BOX(2, 1) = 310
CONST_BOX(2, 2) = 310

Set TEMP_GROUP = DST_WSHEET.Shapes.AddShape(msoShapeRectangle, _
                CONST_BOX(1, 1), CONST_BOX(1, 2), _
                CONST_BOX(2, 1) - CONST_BOX(1, 1), _
                CONST_BOX(2, 2) - CONST_BOX(1, 2)) 'Add Box
TEMP1_STR = TEMP_GROUP.name


PLANE_BOX(1, 1) = XMIN_VAL
PLANE_BOX(1, 2) = YMIN_VAL
PLANE_BOX(2, 1) = XMAX_VAL
PLANE_BOX(2, 2) = YMAX_VAL

If AXES_FLAG = True Then 'add axes
    TEMP2_STR = SHAPE_X_AXES_DRAW_FUNC(CONST_BOX, PLANE_BOX, DST_WSHEET)
    If TEMP2_STR <> "" Then
        DST_WSHEET.Shapes.Range(Array(TEMP1_STR, TEMP2_STR)).Select
        Set TEMP_GROUP = Selection.ShapeRange.Group
        TEMP1_STR = TEMP_GROUP.name
    End If
    TEMP2_STR = SHAPE_Y_AXES_DRAW_FUNC(CONST_BOX, PLANE_BOX, DST_WSHEET)
    If TEMP2_STR <> "" Then
        DST_WSHEET.Shapes.Range(Array(TEMP1_STR, TEMP2_STR)).Select
        Set TEMP_GROUP = Selection.ShapeRange.Group
        TEMP1_STR = TEMP_GROUP.name
    End If
End If

ReDim TEMP_MATRIX(1 To NSIZE ^ 2, 1 To 2)
ReDim PARAM_VECTOR(1 To NSIZE ^ 2, 1 To 2)

HX_VAL = (XMAX_VAL - XMIN_VAL) / (NSIZE + 1)
HY_VAL = (YMAX_VAL - YMIN_VAL) / (NSIZE + 1)

k = 0
For i = 1 To NSIZE
    For j = 1 To NSIZE
        k = k + 1
        XTEMP_VAL = i * HX_VAL + XMIN_VAL
        YTEMP_VAL = j * HY_VAL + YMIN_VAL
        Call SHAPE_TRANSFORM_VECTOR_FUNC(XTEMP_VAL, YTEMP_VAL, PARAM_VECTOR(k, 1), _
                        PARAM_VECTOR(k, 2), PLANE_BOX, CONST_BOX)
        TEMP_MATRIX(k, 1) = XTEMP_VAL
        TEMP_MATRIX(k, 2) = YTEMP_VAL
    Next j
Next i

If DOT_FLAG = True Then
    TEMP2_STR = SHAPE_DRAW_POINT_FUNC(PARAM_VECTOR, 3, DST_WSHEET)
    DST_WSHEET.Shapes.Range(Array(TEMP1_STR, TEMP2_STR)).Select
    Set TEMP_GROUP = Selection.ShapeRange.Group
    TEMP1_STR = TEMP_GROUP.name
End If

DELTA_VAL = (HX_VAL + HY_VAL) / 4  '<<

For k = 1 To UBound(TEMP_MATRIX)
    ReDim PARAM_VECTOR(1 To 3, 1 To 2)
    XTEMP_VAL = TEMP_MATRIX(k, 1)
    YTEMP_VAL = TEMP_MATRIX(k, 2)
    Call SHAPE_TRANSFORM_VECTOR_FUNC(XTEMP_VAL, YTEMP_VAL, PARAM_VECTOR(2, 1), PARAM_VECTOR(2, 2), PLANE_BOX, CONST_BOX)
    FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, XTEMP_VAL, YTEMP_VAL)
'   FUNC_VAL = -2 * XTEMP_VAL * YTEMP_VAL

    If Not IsError(FUNC_VAL) Then
        TETA_VAL = Atn(FUNC_VAL)
        DX_VAL = 0.5 * HX_VAL * Cos(TETA_VAL)
        DY_VAL = 0.5 * HY_VAL * Sin(TETA_VAL)
        'compute the segment bounds
        X1_VAL = XTEMP_VAL + DX_VAL
        Y1_VAL = YTEMP_VAL + DY_VAL
        Call SHAPE_TRANSFORM_VECTOR_FUNC(X1_VAL, Y1_VAL, PARAM_VECTOR(1, 1), PARAM_VECTOR(1, 2), PLANE_BOX, CONST_BOX)
        X2_VAL = XTEMP_VAL - DX_VAL
        Y2_VAL = YTEMP_VAL - DY_VAL
        Call SHAPE_TRANSFORM_VECTOR_FUNC(X2_VAL, Y2_VAL, PARAM_VECTOR(3, 1), PARAM_VECTOR(3, 2), PLANE_BOX, CONST_BOX)
        TEMP2_STR = SHAPE_DRAW_CURVE_FUNC(PARAM_VECTOR, DST_WSHEET)
        DST_WSHEET.Shapes.Range(Array(TEMP1_STR, TEMP2_STR)).Select
        Set TEMP_GROUP = Selection.ShapeRange.Group
        TEMP1_STR = TEMP_GROUP.name
    End If
Next k

ODE_FIELD_GENERIC_FUNC = True

Exit Function
ERROR_LABEL:
ODE_FIELD_GENERIC_FUNC = False
End Function
