Attribute VB_Name = "EXCEL_PICTURES_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private Const PUB_TOGGLE_PICTURE_RNG As String = "RNG_TOGGLE_PICTURE_FUNC"

'Subroutine to toggle image (i.e. zoom in/zoom out) within a cell
'First use requires subroutine be executed while a cell with URL of picture is selected.  After that,
'clicking on the image zooms it to normal size or back down to normal cell size.
   
Public Sub RNG_TOGGLE_PICTURE_FUNC()

Dim TOP_VAL As Double
Dim LEFT_VAL As Double

Dim OLD_URL_STR As String
Dim NEW_URL_STR As String
Dim TEMP_URL_STR As String

Dim TEMP_SHAPE As Excel.Shape
Dim SRC_WSHEET As Excel.Worksheet

On Error GoTo ERROR_LABEL

Set SRC_WSHEET = ActiveSheet
Set TEMP_SHAPE = SRC_WSHEET.Shapes(Excel.Application.Caller)
    
OLD_URL_STR = TEMP_SHAPE.AlternativeText
NEW_URL_STR = TEMP_SHAPE.TopLeftCell.Text

If OLD_URL_STR = NEW_URL_STR Then
    With TEMP_SHAPE
         If Abs(.Height - .TopLeftCell.Height) < 1 Then
            .ScaleHeight 1, msoTrue
            .ScaleWidth 1, msoTrue
         Else
            .Height = .TopLeftCell.Height
         End If
    End With
Else
   LEFT_VAL = TEMP_SHAPE.Left
   TOP_VAL = TEMP_SHAPE.Top
   TEMP_SHAPE.Delete
   Set TEMP_SHAPE = SRC_WSHEET.Pictures.Insert(NEW_URL_STR)
   TEMP_SHAPE.name = Excel.Application.Caller
   TEMP_SHAPE.OnAction = PUB_TOGGLE_PICTURE_RNG
   TEMP_SHAPE.Left = LEFT_VAL
   TEMP_SHAPE.Top = TOP_VAL
   SRC_WSHEET.Shapes(Excel.Application.Caller).AlternativeText = NEW_URL_STR
End If

On Error Resume Next
TEMP_SHAPE.ZOrder msoBringToFront

Exit Sub
ERROR_LABEL:
TEMP_URL_STR = Selection.Text
Set TEMP_SHAPE = SRC_WSHEET.Pictures.Insert(TEMP_URL_STR)
TEMP_SHAPE.OnAction = PUB_TOGGLE_PICTURE_RNG
TEMP_SHAPE.Left = Selection.Left
TEMP_SHAPE.Top = Selection.Top
SRC_WSHEET.Shapes(TEMP_SHAPE.name).AlternativeText = TEMP_URL_STR
End Sub
