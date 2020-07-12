Attribute VB_Name = "WEB_IE_LOGIN_LIBR"
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function IE_WEB_LOGIN_FUNC(ByVal SRC_URL_STR As String, _
ByVal USERNAME_STR As String, _
ByVal PASSWORD_STR As String, _
Optional ByVal SHOW_FLAG As Boolean = False)

Dim i As Integer
Dim j As Integer

Dim LogonForm As HTMLFormElement

Dim CONNECT_OBJ As InternetExplorer

On Error GoTo ERROR_LABEL

IE_WEB_LOGIN_FUNC = False

Set CONNECT_OBJ = New InternetExplorer

With CONNECT_OBJ
    .navigate SRC_URL_STR
    .Visible = SHOW_FLAG
    Do While .Busy: DoEvents: Loop
    Do While .readyState <> 4: DoEvents: Loop
    With .document.forms
        For i = 0 To .length - 1
            Set LogonForm = .Item(i)
            With LogonForm
                For j = 0 To .length - 1
                    If .Item(j).Type = "password" Then
                        LogonForm.UserName.value = USERNAME_STR
                        LogonForm.password.value = PASSWORD_STR
                        LogonForm.submit
                        Do Until CONNECT_OBJ.readyState = READYSTATE_COMPLETE
                            DoEvents
                        Loop
                        If Not CONNECT_OBJ Is Nothing Then CONNECT_OBJ.Quit
                        Set CONNECT_OBJ = Nothing
                        IE_WEB_LOGIN_FUNC = True
                        Exit Function
                    End If
                Next j
            End With
        Next i
    End With
End With

Exit Function
ERROR_LABEL:
On Error Resume Next
If Not CONNECT_OBJ Is Nothing Then CONNECT_OBJ.Quit
Set CONNECT_OBJ = Nothing
IE_WEB_LOGIN_FUNC = False
End Function


