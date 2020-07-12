VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIEBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim PUB_IE_APP_OBJ As InternetExplorer
Attribute PUB_IE_APP_OBJ.VB_VarHelpID = -1

Private Sub Class_Initialize()
On Error Resume Next
PUB_IE_APP_OBJ.Visible = False
If (PUB_IE_APP_OBJ Is Nothing) Or (Err.number <> 0) Then
    Set PUB_IE_APP_OBJ = New InternetExplorer
    PUB_IE_APP_OBJ.Visible = False
End If
End Sub

Public Sub IENavigate(ByVal SRC_URL_STR As String, _
ByVal FUNC_NAME_STR As String)
   With PUB_IE_APP_OBJ
       .navigate SRC_URL_STR
       .Visible = False
       Do While .Busy: DoEvents: Loop
       Do While .readyState <> 4: DoEvents: Loop
   End With
   Call Excel.Application.Run(PUB_IE_APP_OBJ.document)
   'PUB_IE_APP_OBJ.document.DocumentElement.outerHTML
End Sub
