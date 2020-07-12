Attribute VB_Name = "EXCEL_CHART_PICTURE_LIBR"
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'--------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : CHART_SNAPSHOT_FUNC
'DESCRIPTION   : TAKE A PICTURE OF A CHART
'LIBRARY       : EXCEL_CHART
'GROUP         : PICTURE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'************************************************************************************
'************************************************************************************

Function EXCEL_CHART_SNAPSHOT_FUNC(ByRef CHART_OBJ As Excel.Chart, _
ByRef DST_RNG As Excel.Range)

On Error GoTo ERROR_LABEL

EXCEL_CHART_SNAPSHOT_FUNC = False

CHART_OBJ.CopyPicture Appearance:=xlScreen, Format:=xlPicture
DST_RNG.Activate
DST_RNG.Worksheet.Paste

EXCEL_CHART_SNAPSHOT_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_SNAPSHOT_FUNC = False
End Function

Function EXCEL_CHART_COPY_PASTE_IMAGE_FUNC(ByRef IMAGE_OBJ As Image, _
ByRef CHART_OBJ As Excel.Chart)

Dim k As Long 'lPicType

On Error GoTo ERROR_LABEL

EXCEL_CHART_COPY_PASTE_IMAGE_FUNC = False

'Do we want a metafile or a bitmap?
'If doing a 1 to 1 copy, xlBitmap will give a 'truer' rendition.
'If scaling the image, xlPicture will give better results
k = xlPicture
'Update the chart type and copy it to the clipboard, as seen on screen
CHART_OBJ.CopyPicture xlScreen, k, xlScreen
'Paste the picture from the clipboard into our image control
Set IMAGE_OBJ.Picture = PastePicture(k)

EXCEL_CHART_COPY_PASTE_IMAGE_FUNC = True

Exit Function
ERROR_LABEL:
EXCEL_CHART_COPY_PASTE_IMAGE_FUNC = False
End Function
