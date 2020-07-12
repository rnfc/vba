Attribute VB_Name = "WEB_XML_POST_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function XML_WEB_PAGE_POST_FUNC(ByVal SRC_URL_STR As String, _
Optional ByVal REFER_STR As String = "username=rafael_nicolas&password=hello")

Dim RESPONSE_TEXT As String
Dim SEND_MSG_STR() As Byte
Dim CONNECT_OBJ As MSXML2.XMLHTTP60

On Error GoTo ERROR_LABEL

Set CONNECT_OBJ = New MSXML2.XMLHTTP60

SEND_MSG_STR = StrConv(REFER_STR, vbFromUnicode)

With CONNECT_OBJ
    .Open "POST", SRC_URL_STR
    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    .send SEND_MSG_STR
    RESPONSE_TEXT = .ResponseText
End With

If Not CONNECT_OBJ Is Nothing Then Set CONNECT_OBJ = Nothing
XML_WEB_PAGE_POST_FUNC = IIf(RESPONSE_TEXT = "", False, RESPONSE_TEXT)

Exit Function
ERROR_LABEL:
On Error Resume Next
If Not CONNECT_OBJ Is Nothing Then Set CONNECT_OBJ = Nothing
XML_WEB_PAGE_POST_FUNC = False
End Function
