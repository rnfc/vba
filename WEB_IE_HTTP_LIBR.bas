Attribute VB_Name = "WEB_IE_HTTP_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Need a reference to Microsoft Internet Controls
'Need a reference to Microsoft HTML Object Library

Function IE_HTTP_SYNCHRONOUS_FUNC(ByRef SRC_URL_STR As String, _
Optional ByVal VERSION As Integer = 0)

Dim DATA_STR As String
Dim CONNECT_OBJ As InternetExplorer 'As Object
'CreateObject("InternetExplorer.Application")

On Error GoTo ERROR_LABEL

Set CONNECT_OBJ = New InternetExplorer

'------------------------------------------------------------------------------
Select Case VERSION
'------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------
    With CONNECT_OBJ
        .Visible = False
        .navigate SRC_URL_STR
        Do While .Busy: DoEvents: Loop
        'IE runs asynchronously with VBA
        Do While .readyState <> 4: DoEvents: Loop
        DATA_STR = .document.DocumentElement.outerHTML
        .Quit
    End With
'------------------------------------------------------------------------------
Case Else
'------------------------------------------------------------------------------
    With CONNECT_OBJ
        .Visible = False
        .navigate SRC_URL_STR
        Do While .Busy: DoEvents: Loop
        'IE runs asynchronously with VBA
        Do While .readyState <> 4: DoEvents: Loop
        .ExecWB 17, 0 '--> Select all
        'Prompt the user for input or not, whichever is
        'the default behavior.
        .ExecWB 12, 2 '--> Copy
        'Execute the command without prompting the user.
        .ExecWB 18, 2 'Clear Selection
        'References:
            'http://msdn2.microsoft.com/en-us/library/ms691264.aspx
            'http://msdn2.microsoft.com/en-us/library/ms683930.aspx
        DATA_STR = .document.DocumentElement.outerHTML
    End With
'------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------
If Not CONNECT_OBJ Is Nothing Then CONNECT_OBJ.Quit
Set CONNECT_OBJ = Nothing
IE_HTTP_SYNCHRONOUS_FUNC = DATA_STR

Exit Function
ERROR_LABEL:
On Error Resume Next
If Not CONNECT_OBJ Is Nothing Then CONNECT_OBJ.Quit
Set CONNECT_OBJ = Nothing
IE_HTTP_SYNCHRONOUS_FUNC = Err.number
End Function

Function IE_HTTP_ASYNCHRONOUS_FUNC(ByVal FUNC_NAME_STR As String, _
ByVal SRC_URL_STR As String, _
Optional ByRef IE_CLASS_OBJ As clsIEBrowser)
    On Error GoTo ERROR_LABEL
IE_HTTP_ASYNCHRONOUS_FUNC = False
    If IE_CLASS_OBJ Is Nothing Then: Set IE_CLASS_OBJ = New clsIEBrowser
    Call IE_CLASS_OBJ.IENavigate(SRC_URL_STR, FUNC_NAME_STR)
IE_HTTP_ASYNCHRONOUS_FUNC = True
Exit Function
ERROR_LABEL:
IE_HTTP_ASYNCHRONOUS_FUNC = Err.number
End Function
