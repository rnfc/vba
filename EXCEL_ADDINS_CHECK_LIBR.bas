Attribute VB_Name = "EXCEL_ADDINS_CHECK_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private PUB_ADDIN_ADDRESS_STR As String
Private PUB_ADDIN_INSTALLED_PROPERLY_FLAG As Boolean

Private Sub ADDIN_INSTALLATION_FLAG_FUNC()
'Private Sub Workbook_AddinInstall()
    PUB_ADDIN_INSTALLED_PROPERLY_FLAG = True
End Sub

Private Sub ADDIN_INSTALLATION_CHECK_FUNC()
'Private Sub Workbook_Open()
Dim MATCH_FLAG As Boolean
Dim SRC_ADDIN As Excel.AddIn
Dim SRC_WBOOK As Excel.Workbook

On Error GoTo ERROR_LABEL

Set SRC_WBOOK = ThisWorkbook
If Not SRC_WBOOK.IsAddin Then Exit Sub
If Not PUB_ADDIN_INSTALLED_PROPERLY_FLAG Then
'       Add it to the AddIns collection
    MATCH_FLAG = False
    For Each SRC_ADDIN In AddIns
        If SRC_ADDIN.name = SRC_WBOOK.name Then: MATCH_FLAG = True
    Next SRC_ADDIN
    If Not MATCH_FLAG Then AddIns.Add FileName:=SRC_WBOOK.FullName
'       Install it
    TITLE_STR = ""
    For Each SRC_ADDIN In AddIns
        If SRC_ADDIN.name = SRC_WBOOK.name Then: TITLE_STR = SRC_ADDIN.Title
    Next SRC_ADDIN
    Excel.Application.EnableEvents = False
    AddIns(TITLE_STR).Installed = True
    Excel.Application.EnableEvents = True
'       Inform user
    MSG_STR = SRC_WBOOK.name & " has been installed as an add-in. "
    MSG_STR = MSG_STR & "Use the Tools Add-Ins command to uninstall it."
    MsgBox MSG_STR, vbInformation, TITLE_STR
End If

Exit Sub
ERROR_LABEL:
End Sub

Private Sub ADDIN_FIX_LINKS_WBOOK_FUNC(Optional ByRef SRC_WBOOK As Workbook)
Dim DSHEET As Worksheet
On Error GoTo ERROR_LABEL
If SRC_WBOOK Is Nothing Then: Set SRC_WBOOK = ThisWorkbook
If PUB_ADDIN_ADDRESS_STR = "" Then: PUB_ADDIN_ADDRESS_STR = SRC_WBOOK.Path & "\" & SRC_WBOOK.name
For Each DSHEET In SRC_WBOOK.Worksheets
    DSHEET.Cells.Replace _
        What:=PUB_ADDIN_ADDRESS_STR, _
        Replacement:="", _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        MatchCase:=False, _
        SearchFormat:=False, _
        ReplaceFormat:=False
Next DSHEET
Exit Sub
ERROR_LABEL:
End Sub



