Attribute VB_Name = "MATRIX_GRID_LIBR"

'-----------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-----------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_DRAW_GRID_FUNC

'DESCRIPTION   : Matrix Flow - graph; the graph consists of nodes and branches
'the number of nodes is equal to the dimension of the matrix; the nodes, numbered
'from 1 to N, represent the elements of; the first diagonal aii for all elements
'aij we draw an oriented branch (arrow) from node-i to node-j

'LIBRARY       : MATRIX
'GROUP         : EXTRACTION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/19/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_DRAW_GRID_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal LABEL_FLAG As Boolean = True, _
Optional ByVal LINK_FLAG As Boolean = True, _
Optional ByVal DST_WSHEET As Excel.Worksheet)

Dim i As Long
Dim j As Long

Dim ii As Variant
Dim jj As Variant

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim X0_VAL As Double
Dim X_VAL As Double
Dim DX_VAL As Double
Dim DDX_VAL As Double

Dim Y0_VAL As Double
Dim Y_VAL As Double
Dim DY_VAL As Double
Dim DDY_VAL As Double

Dim TETA_VAL As Double
Dim RAD_VAL As Double
Dim LIMIT_VAL As Double
Dim DIAM_VAL As Double
Dim HEIGHT_VAL As Double

Dim ATEMP_STR As String
Dim BTEMP_STR As String

Dim NODE_VECTOR As Variant
Dim DATA_MATRIX As Variant

Dim SHAPE_OBJ As Object

On Error GoTo ERROR_LABEL

MATRIX_DRAW_GRID_FUNC = False
If DST_WSHEET Is Nothing Then: Set DST_WSHEET = ActiveSheet
DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

If NROWS * NCOLUMNS > 400 Then: GoTo ERROR_LABEL 'Matrix too large

NSIZE = NROWS
If NCOLUMNS > NSIZE Then NSIZE = NCOLUMNS

ReDim NODE_VECTOR(1 To NSIZE, 1 To 3)
'Column 1 --> Name
'Column 2 --> PosLeft
'Column 3 --> PosTop

HEIGHT_VAL = 15
LIMIT_VAL = 25
DIAM_VAL = 23
RAD_VAL = 15 * NSIZE + 10

X0_VAL = 350
If NROWS < 11 Then
    Y0_VAL = 190
Else
    Y0_VAL = 15 * NSIZE + 40
End If
Y0_VAL = Y0_VAL + HEIGHT_VAL

TETA_VAL = 2 * 3.14159265358979 / NSIZE

For i = 1 To NSIZE
    X_VAL = X0_VAL + RAD_VAL * Cos(TETA_VAL * i)
    Y_VAL = Y0_VAL - RAD_VAL * Sin(TETA_VAL * i)
    ATEMP_STR = CStr(i)
    
    Set SHAPE_OBJ = DST_WSHEET.Shapes.AddShape(msoShapeOval, X_VAL, Y_VAL, DIAM_VAL, DIAM_VAL)
    SHAPE_OBJ.DrawingObject.Text = ATEMP_STR
    SHAPE_OBJ.TextFrame.HorizontalAlignment = xlHAlignCenter
    BTEMP_STR = SHAPE_OBJ.name
    Set SHAPE_OBJ = Nothing
    
    NODE_VECTOR(i, 1) = BTEMP_STR
    NODE_VECTOR(i, 2) = X_VAL
    NODE_VECTOR(i, 3) = Y_VAL
Next i

For i = 1 To NROWS
    For j = 1 To NCOLUMNS
        If DATA_MATRIX(i, j) <> 0 And DATA_MATRIX(i, j) <> "" Then
            If i <> j Then
                Set SHAPE_OBJ = _
                    DST_WSHEET.Shapes.AddConnector(msoConnectorStraight, _
                    282#, 196.5, 1.5, 26.25)
                SHAPE_OBJ.Line.EndArrowheadStyle = msoArrowheadTriangle
                With SHAPE_OBJ.ConnectorFormat
                    .BeginConnect _
                        ConnectedShape:=DST_WSHEET.Shapes(NODE_VECTOR(i, 1)), _
                        ConnectionSite:=5
                    .EndConnect _
                        ConnectedShape:=DST_WSHEET.Shapes(NODE_VECTOR(j, 1)), _
                        ConnectionSite:=1
                    SHAPE_OBJ.RerouteConnections
                End With
                Set SHAPE_OBJ = Nothing
                
                If LABEL_FLAG = True Then
                    ATEMP_STR = CStr(DATA_MATRIX(i, j))
                    Set SHAPE_OBJ = _
                        DST_WSHEET.Shapes.AddShape(msoTextOrientationHorizontal, _
                        (2 * NODE_VECTOR(i, 2) + NODE_VECTOR(j, 2)) / 3, _
                        (2 * NODE_VECTOR(i, 3) + NODE_VECTOR(j, 3)) / 3, 26, 15)
                        SHAPE_OBJ.DrawingObject.Text = ATEMP_STR
                        SHAPE_OBJ.DrawingObject.Font.Size = 8
                        SHAPE_OBJ.TextFrame.HorizontalAlignment = xlHAlignCenter
                        SHAPE_OBJ.TextFrame.VerticalAlignment = xlHAlignCenter
                        If Len(ATEMP_STR) > 4 Then: _
                            SHAPE_OBJ.TextFrame.AutoSize = True
                        BTEMP_STR = SHAPE_OBJ.name
                        Set SHAPE_OBJ = Nothing
                End If
            End If
        
            If i = j And LINK_FLAG = True Then
            
                X_VAL = DST_WSHEET.Shapes(NODE_VECTOR(i, 1)).Left
                Y_VAL = DST_WSHEET.Shapes(NODE_VECTOR(i, 1)).Top
                DX_VAL = X_VAL - X0_VAL
                DY_VAL = Y_VAL - Y0_VAL
                
                If DX_VAL < 0 Then
                    If DY_VAL > 0 Then
                        ii = 3
                        jj = 5
                    Else
                        ii = 1
                        jj = 3
                    End If
                Else
                    If DY_VAL > 0 Then
                          ii = 5
                          jj = 7
                    Else
                          ii = 7
                          jj = 1
                    End If
                End If

                Set SHAPE_OBJ = DST_WSHEET.Shapes.AddConnector(msoConnectorCurve, 282#, 196.5, 1.5, 26.25)
                SHAPE_OBJ.Line.EndArrowheadStyle = msoArrowheadTriangle
                With SHAPE_OBJ.ConnectorFormat
                    .BeginConnect _
                        ConnectedShape:=DST_WSHEET.Shapes(NODE_VECTOR(i, 1)), _
                        ConnectionSite:=ii
                    .EndConnect _
                        ConnectedShape:=DST_WSHEET.Shapes(NODE_VECTOR(i, 1)), _
                        ConnectionSite:=jj
                    'SHAPE_OBJ.RerouteConnections
                End With
                Set SHAPE_OBJ = Nothing
            
                If LABEL_FLAG = True Then
                    ATEMP_STR = CStr(DATA_MATRIX(i, j))
            
                        X_VAL = DST_WSHEET.Shapes(NODE_VECTOR(i, 1)).Left
                        Y_VAL = DST_WSHEET.Shapes(NODE_VECTOR(i, 1)).Top
                        
                        DX_VAL = X_VAL - X0_VAL
                        DY_VAL = Y_VAL - Y0_VAL
                        
                        DDX_VAL = LIMIT_VAL
                        DDY_VAL = LIMIT_VAL
                        
                        If DX_VAL < 0 Then DDX_VAL = -DDX_VAL
                        If DY_VAL < 0 Then DDY_VAL = -DDY_VAL
            
                        Set SHAPE_OBJ = _
                            DST_WSHEET.Shapes.AddShape(msoTextOrientationHorizontal, _
                                X_VAL + DDX_VAL, Y_VAL + DDY_VAL, 26, 15)
                            SHAPE_OBJ.DrawingObject.Text = ATEMP_STR
                            SHAPE_OBJ.DrawingObject.Font.Size = 8
                            SHAPE_OBJ.TextFrame.HorizontalAlignment = xlHAlignCenter
                            SHAPE_OBJ.TextFrame.VerticalAlignment = xlHAlignCenter
                            If Len(ATEMP_STR) > 4 Then: _
                            SHAPE_OBJ.TextFrame.AutoSize = True
                            ATEMP_STR = SHAPE_OBJ.name
                            Set SHAPE_OBJ = Nothing
                End If
            End If
        End If
    Next j
Next i

MATRIX_DRAW_GRID_FUNC = True

Exit Function
ERROR_LABEL:
MATRIX_DRAW_GRID_FUNC = False
End Function
