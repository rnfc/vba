Attribute VB_Name = "EXCEL_RNG_HEAT_LIBR"

'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------
'Making Heatcharts in Excel
'-----------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------


Function RNG_HEAT_MAP1_COLOUR_FUNC(ByRef DATA_RNG As Excel.Range, _
ByRef RGB_RNG As Excel.Range, _
Optional ByRef SRC_WBOOK As Excel.Workbook)

Dim i As Long 'column
Dim j As Long 'column

Dim X_VAL As Single

Dim MIN_VAL As Single
Dim MAX_VAL As Single

Dim RMIN_VAL As Long
Dim GMIN_VAL As Long
Dim BMIN_VAL As Long

Dim RMAX_VAL As Long
Dim GMAX_VAL As Long
Dim BMAX_VAL As Long

Dim TEMP_CELL As Excel.Range
Dim RGB_MATRIX As Variant

On Error GoTo ERROR_LABEL

RNG_HEAT_MAP1_COLOUR_FUNC = False

If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook

RGB_MATRIX = RGB_RNG
'You set the three colours above by specifying RGB values. You need to
'do it this way because you can't use the existing palette.
'To find out the RGB values, select Tools, Options, then the Color tab.
'Click any colour, then the Modify button. The Custom tab will give you
'RGB values. You then create your table on this sheet, select the cells
'you want to colour, and press the button above. Then copy your table to
'Powerpoint, Word or whatever.

'--------------------------------------
'        Low   Middle  High
'--------------------------------------
'R       247    255    152
'G       129    255    182
'B       117    255    212
'--------------------------------------
'Value   -5      0      5
'--------------------------------------

RMIN_VAL = RGB_MATRIX(1, 1)
RMAX_VAL = RGB_MATRIX(1, 2)

GMIN_VAL = RGB_MATRIX(2, 1)
GMAX_VAL = RGB_MATRIX(2, 2)

BMIN_VAL = RGB_MATRIX(3, 1)
BMAX_VAL = RGB_MATRIX(3, 2)

MIN_VAL = RGB_MATRIX(4, 1)
MAX_VAL = RGB_MATRIX(4, 2)

For i = 0 To 55
  X_VAL = i / 55
  SRC_WBOOK.Colors(i + 1) = RGB(RMIN_VAL + (RMAX_VAL - RMIN_VAL) * X_VAL, _
                                   GMIN_VAL + (GMAX_VAL - GMIN_VAL) * X_VAL, _
                                   BMIN_VAL + (BMAX_VAL - BMIN_VAL) * X_VAL)
Next i


On Error Resume Next

For Each TEMP_CELL In DATA_RNG.Cells
  With TEMP_CELL
    If TEMP_CELL <> "" Then
      X_VAL = .value
      j = Int(56 * (X_VAL - MIN_VAL) / (MAX_VAL - MIN_VAL)) + 1
      If j < 1 Then j = 1 Else If j > 56 Then j = 56
      .Interior.ColorIndex = j
      If j > 0.6 * 16777215 Then
        '.Font.Color = 16777215
      Else
        '.Font.Color = 0
      End If
    Else
      .Interior.ColorIndex = xlNone
    End If
  End With
Next TEMP_CELL

RNG_HEAT_MAP1_COLOUR_FUNC = True

Exit Function
ERROR_LABEL:
RNG_HEAT_MAP1_COLOUR_FUNC = False
End Function



Function RNG_HEAT_MAP2_COLOUR_FUNC(ByRef DATA_RNG As Excel.Range, _
ByRef RGB_RNG As Excel.Range, _
Optional ByRef SRC_WBOOK As Excel.Workbook)

Dim i As Long 'column
Dim j As Long 'column

Dim X_VAL As Single

Dim MIN_VAL As Single
Dim MID_VAL As Single
Dim MAX_VAL As Single

Dim RMIN_VAL As Long
Dim GMIN_VAL As Long
Dim BMIN_VAL As Long

Dim RMID_VAL As Long
Dim GMID_VAL As Long
Dim BMID_VAL As Long

Dim RMAX_VAL As Long
Dim GMAX_VAL As Long
Dim BMAX_VAL As Long

Dim TEMP_CELL As Excel.Range
Dim RGB_MATRIX As Variant

On Error GoTo ERROR_LABEL

RNG_HEAT_MAP2_COLOUR_FUNC = False

If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ActiveWorkbook

RGB_MATRIX = RGB_RNG
'You set the three colours above by specifying RGB values. You need to
'do it this way because you can't use the existing palette.
'To find out the RGB values, select Tools, Options, then the Color tab.
'Click any colour, then the Modify button. The Custom tab will give you
'RGB values. You then create your table on this sheet, select the cells
'you want to colour, and press the button above. Then copy your table to
'Powerpoint, Word or whatever.

'--------------------------------------
'        Low   Middle  High
'--------------------------------------
'R       247    255    152
'G       129    255    182
'B       117    255    212
'--------------------------------------
'Value   -5      0      5
'--------------------------------------

RMIN_VAL = RGB_MATRIX(1, 1)
RMID_VAL = RGB_MATRIX(1, 2)
RMAX_VAL = RGB_MATRIX(1, 3)

GMIN_VAL = RGB_MATRIX(2, 1)
GMID_VAL = RGB_MATRIX(2, 2)
GMAX_VAL = RGB_MATRIX(2, 3)

BMIN_VAL = RGB_MATRIX(3, 1)
BMID_VAL = RGB_MATRIX(3, 2)
BMAX_VAL = RGB_MATRIX(3, 3)

MIN_VAL = RGB_MATRIX(4, 1)
MID_VAL = RGB_MATRIX(4, 2)
MAX_VAL = RGB_MATRIX(4, 3)

For i = 0 To 27
  X_VAL = i / 27
  SRC_WBOOK.Colors(i + 1) = RGB(RMIN_VAL + (RMID_VAL - RMIN_VAL) * X_VAL, _
                                   GMIN_VAL + (GMID_VAL - GMIN_VAL) * X_VAL, _
                                   BMIN_VAL + (BMID_VAL - BMIN_VAL) * X_VAL)
Next i

For i = 0 To 27
  X_VAL = i / 27
  SRC_WBOOK.Colors(i + 29) = RGB(RMID_VAL + (RMAX_VAL - RMID_VAL) * X_VAL, _
                                    GMID_VAL + (GMAX_VAL - GMID_VAL) * X_VAL, _
                                    BMID_VAL + (BMAX_VAL - BMID_VAL) * X_VAL)
Next i

On Error Resume Next

For Each TEMP_CELL In DATA_RNG.Cells
  With TEMP_CELL
    If TEMP_CELL <> "" Then
      X_VAL = .value
      j = Int(56 * (X_VAL - MIN_VAL) / (MAX_VAL - MIN_VAL)) + 1
      If j < 1 Then j = 1 Else If j > 56 Then j = 56
      .Interior.ColorIndex = j
      If j > 0.6 * 16777215 Then
        '.Font.Color = 16777215
      Else
        '.Font.Color = 0
      End If
    Else
      .Interior.ColorIndex = xlNone
    End If
  End With
Next TEMP_CELL

RNG_HEAT_MAP2_COLOUR_FUNC = True

Exit Function
ERROR_LABEL:
RNG_HEAT_MAP2_COLOUR_FUNC = False
End Function
