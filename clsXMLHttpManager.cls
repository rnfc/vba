VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsXMLHttpManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim PUB_XML_HTTP_MON() As clsXMLHttpMonitor


Public Sub XMLHttpCall( _
ByVal METHOD_STR As String, _
ByVal SRC_URL_STR As String, _
ByRef RESPONSE_TEXT As String, _
ByVal FUNCTION_STR As String, _
Optional ByVal COMMAND_STR As String, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal ASYNC_CALL_FLAG As Boolean = True, _
Optional ByVal SEND_MSG_STR As String = "")

Dim PUB_XML_HTTP_MON As clsXMLHttpMonitor

Set PUB_XML_HTTP_MON = findAvailMon()

PUB_XML_HTTP_MON.XMLHttpCall METHOD_STR, SRC_URL_STR, _
                             FUNCTION_STR, COMMAND_STR, PARAM_RNG, _
                             ASYNC_CALL_FLAG, SEND_MSG_STR

RESPONSE_TEXT = PUB_XML_HTTP_MON.ResponseText

End Sub

Private Function findAvailMon() As clsXMLHttpMonitor
Dim i As Integer
Dim EXIT_FLAG As Boolean
i = LBound(PUB_XML_HTTP_MON)
Do
    If PUB_XML_HTTP_MON(i) Is Nothing Then
        EXIT_FLAG = True
    ElseIf PUB_XML_HTTP_MON(i).XMLAvailable Then
        EXIT_FLAG = True
    Else
        i = i + 1
        EXIT_FLAG = i > UBound(PUB_XML_HTTP_MON)
    End If
Loop Until EXIT_FLAG

If i > UBound(PUB_XML_HTTP_MON) Then
    ReDim Preserve PUB_XML_HTTP_MON(0 To UBound(PUB_XML_HTTP_MON) + 1)
End If
Set PUB_XML_HTTP_MON(i) = New clsXMLHttpMonitor
Set findAvailMon = PUB_XML_HTTP_MON(i)
End Function

Private Sub Class_Initialize()
    ReDim PUB_XML_HTTP_MON(0 To 0)
End Sub

