Attribute VB_Name = "EXCEL_CHART_SHAPE_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      :
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         :
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_SHAPE_LOOK_FUNC(ByVal SHAPE_NAME_STR As Variant, _
Optional ByVal OUTPUT As Integer = 0, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)

Dim i As Long
Dim SHAPE_OBJ As Excel.Shapes
Dim MATCH_FLAG As Boolean

On Error GoTo ERROR_LABEL

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

MATCH_FLAG = False
i = 0
For Each SHAPE_OBJ In SRC_WSHEET.Shapes
    i = i + 1
    If SHAPE_OBJ.name = SHAPE_NAME_STR Then
        MATCH_FLAG = True
        Exit For
    End If
Next SHAPE_OBJ

If MATCH_FLAG = True Then
    Select Case OUTPUT
    Case 0
        EXCEL_CHART_SHAPE_LOOK_FUNC = True
    Case Else
        EXCEL_CHART_SHAPE_LOOK_FUNC = i
    End Select
Else
    EXCEL_CHART_SHAPE_LOOK_FUNC = False
End If

Exit Function
ERROR_LABEL:
EXCEL_CHART_SHAPE_LOOK_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      :
'DESCRIPTION   :
'LIBRARY       : EXCEL_CHART
'GROUP         :
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_SHAPE_SCALE_FUNC(ByVal CHART_NAME_STR As Variant, _
Optional ByVal WIDTH_VAL As Variant = Null, _
Optional ByVal HEIGHT_VAL As Variant = Null, _
Optional ByVal VERSION As Integer = 0, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)
            
Dim SHAPE_OBJ As Object

On Error GoTo ERROR_LABEL

EXCEL_CHART_SHAPE_SCALE_FUNC = False
If (VarType(WIDTH_VAL) = vbNull) And (VarType(HEIGHT_VAL) = vbNull) Then: GoTo ERROR_LABEL

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

If EXCEL_CHART_LOOK_FUNC(CHART_NAME_STR, 0, SRC_WSHEET) = True Then
   Set SHAPE_OBJ = SRC_WSHEET.Shapes(CHART_NAME_STR)
Else
   GoTo ERROR_LABEL
End If

Select Case VERSION
Case 0
    SHAPE_OBJ.ScaleWidth WIDTH_VAL, msoFalse, msoScaleFromTopLeft
Case Else
    SHAPE_OBJ.ScaleHeight HEIGHT_VAL, msoFalse, msoScaleFromBottomRight
End Select

EXCEL_CHART_SHAPE_SCALE_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_SHAPE_SCALE_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : EXCEL_CHART_SHAPE_HIDE_FUNC
'DESCRIPTION   :
'LIBRARY       :
'GROUP         :
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************


Function EXCEL_CHART_SHAPE_HIDE_FUNC(ByVal SHAPE_NAME_STR As Variant, _
ByVal HROW_VAL As Double, _
ByVal HCOLUMN_VAL As Double, _
ByVal SROW_VAL As Double, _
ByVal SCOLUMN_VAL As Double, _
Optional ByVal VERSION As Integer = 0, _
Optional ByRef SRC_WSHEET As Excel.Worksheet)
            
Dim SHAPE_OBJ As Object

On Error GoTo ERROR_LABEL

EXCEL_CHART_SHAPE_HIDE_FUNC = False

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

If EXCEL_CHART_SHAPE_LOOK_FUNC(SHAPE_NAME_STR, 0, SRC_WSHEET) = True Then
    Set SHAPE_OBJ = SRC_WSHEET.Shapes(SHAPE_NAME_STR)
Else
    GoTo ERROR_LABEL
End If

Select Case VERSION
'---------------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------------
    SHAPE_OBJ.IncrementLeft SCOLUMN_VAL
    SHAPE_OBJ.IncrementTop SROW_VAL
    SHAPE_OBJ.Characters.Text = "SHOW GRAPH"
'---------------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------------
    SHAPE_OBJ.IncrementLeft HCOLUMN_VAL
    SHAPE_OBJ.IncrementTop HROW_VAL
    SHAPE_OBJ.Characters.Text = "HIDE GRAPH"
'    SHAPE_OBJ.Select
'    Selection.Characters.Text = "HIDE GRAPH"
'---------------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------------
 
EXCEL_CHART_SHAPE_HIDE_FUNC = True
   
Exit Function
ERROR_LABEL:
EXCEL_CHART_SHAPE_HIDE_FUNC = False
End Function
