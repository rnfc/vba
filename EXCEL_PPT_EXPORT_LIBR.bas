Attribute VB_Name = "EXCEL_PPT_EXPORT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function EXCEL_PPT_EXPORT_FUNC(ByRef SRC_RNG() As Excel.Range, _
Optional ByVal FILE_NAME_STR As Variant = "", _
Optional ByVal TITLE_STR As String = "", _
Optional ByVal REPORT_DATE_STR As String = "", _
Optional ByVal HEIGHT_VAL As Double = 0.75, _
Optional ByVal WIDTH_VAL As Double = 0.75)

'HEIGHT/WIDTH: Scale Factor PPT Export

'Dim FILE_NAME_STR As String
'Dim REPORT_DATE_STR As String
'REPORT_DATE_STR = Range("myReportDate").Value & " / Week " & _
                  Range("myReportWeek").Value & " - "
'FILE_NAME_STR = ThisWorkbook.Path & Application.PathSeparator & "NICO.ppt"
'Dim SRC_RNG(1 To 3) As Excel.Range
'Set SRC_RNG(1) = Range("myDashboard01")
'Set SRC_RNG(2) = Range("myDashboard02")
'Set SRC_RNG(3) = Range("myDashboard03")
'Debug.Print EXCEL_PPT_EXPORT_FUNC(SRC_RNG(), FILE_NAME_STR, _
     Range("myInputStartTitles").Offset(1, 0), REPORT_DATE_STR, 1, 1)

Dim i As Long
Dim j As Long
Dim k As Long
Dim PP_OBJ As Object
Dim FILE_OBJ As Object
Dim SLIDE_OBJ As Object

On Error GoTo ERROR_LABEL

If REPORT_DATE_STR = "" Then: REPORT_DATE_STR = Format(Now, "DD-MM-YY")
Set PP_OBJ = CreateObject("Powerpoint.Application")

If FILE_NAME_STR = "" Then
    FILE_NAME_STR = _
    Application.GetOpenFilename("Microsoft PowerPoint-Files (*.ppt), *.ppt")
End If

If FILE_NAME_STR = False Then
    PP_OBJ.Activate
    PP_OBJ.Presentations.Add
    Set FILE_OBJ = PP_OBJ.ActivePresentation
Else
    PP_OBJ.Activate
    Set FILE_OBJ = PP_OBJ.Presentations.Open(FILE_NAME_STR)
End If
PP_OBJ.Visible = True

For k = LBound(SRC_RNG()) To UBound(SRC_RNG())
    SRC_RNG(k).CopyPicture Appearance:=xlScreen, Format:=xlPicture
    i = 11
    PP_OBJ.ActivePresentation.Slides.Add PP_OBJ.ActivePresentation.Slides.COUNT + 1, i
    
    Set SLIDE_OBJ = FILE_OBJ.Slides(PP_OBJ.ActivePresentation.Slides.COUNT)
    SLIDE_OBJ.Shapes.Title.TextFrame.TextRange.Text = REPORT_DATE_STR & ": " & TITLE_STR
    j = SLIDE_OBJ.Shapes.COUNT + 1
    i = 2
    With SLIDE_OBJ
        .Shapes.PasteSpecial i
        .Shapes(j).ScaleHeight HEIGHT_VAL, 1
        .Shapes(j).ScaleWidth WIDTH_VAL, 1
        .Shapes(j).Left = FILE_OBJ.PageSetup.SlideWidth \ 2 - _
                          SLIDE_OBJ.Shapes(j).Width \ 2
        .Shapes(j).Top = 90
    End With
Next k

Set SLIDE_OBJ = Nothing
Set FILE_OBJ = Nothing
Set PP_OBJ = Nothing
'Worksheets(1).Activate

EXCEL_PPT_EXPORT_FUNC = True

Exit Function
ERROR_LABEL:
Set SLIDE_OBJ = Nothing
Set FILE_OBJ = Nothing
Set PP_OBJ = Nothing
'MsgBox "Error No.: " & Err.Number & vbNewLine & vbNewLine & "Description: " & _
Err.Description, vbCritical, "Error"
EXCEL_PPT_EXPORT_FUNC = False
End Function
