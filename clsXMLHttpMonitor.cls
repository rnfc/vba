VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsXMLHttpMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ResponseText As String

Dim PUB_AVAILABLE_FLAG As Boolean
Dim PUB_FUNCTION_STR As String
Dim PUB_PARAM_RNG As Variant
Dim PUB_COMMAND_STR As String
Dim PUB_SRC_URL_STR As String
Dim PUB_SEND_MSG_STR As String
Dim PUB_XML_HTTP_REQUEST As MSXML2.XMLHTTP60

'It’s always a good idea to encapsulate such code and in this case it is
'relatively easy to do. But before we do that there is one more item to
'consider. Once we have all the code in a class module, we must provide
'some way for the consumer of the class to “tell” us how to process the
'response.  Obviously, we don’t want to change the code in the class
'module each time someone wants to call a new routine to process the
'response. So, we implement a new “ResponseProcessor” variable and use
'the Application.Run method to call it when the request is completed.
'A sample subroutine that acts as a response processor is below.
'Note that it takes three parameters, the XMLHttp object, the original
'URL, and the original message.

'For Microsoft’s take on the different ways to asynchronously process
'data with the XMLHttp see
'http://msdn2.microsoft.com/en-us/library/ms757030.aspx

'To implement all of the above issues, the new clsXMLHttpMonitor will (1)
'have a read-only flag to indicate when it is available for a new task,
'(2) take care of all XMLHttpReq processing, and (3) call a user
'provided “ResponseProcessor” subroutine.  Note that the onreadystatechange
'property is now set to Me, i.e., the instantiation of the class.

'While we are at it, we might as well take care of one more item.
'Since we are encapsulating all the XMLHttp related code in a single
'class, we might as well take over responsibility for synchronous requests.

'There’s one more task to take care of.  We need code to properly
'instantiate a new request taking into account that existing requests
'may still be pending.  Rather than make this the responsibility of
'every consumer of our class, we will create a ‘Manager’ class.
'The task of this new class module, named clsXMLHttpManager, is the
'following: It either finds the first available XMLHttpMon object or
'creates a new one. Next, it passes along to this object the request it
'received.  So, once it finds (or creates) a new monitor, it acts as
'just a pass through routine.

Property Get XMLAvailable() As Boolean
    XMLAvailable = PUB_AVAILABLE_FLAG
End Property

Sub ReadyStateChangeHandler()
Attribute ReadyStateChangeHandler.VB_UserMemId = 0
'-------------------------------------------------------------------------
'As noted above, we still have to designate
'ReadyStateChangeHandler as the default method for
'this class.  To make that happen, do the following:
'1)      Export and remove this class module.
'2)      Open the exported file in a text editor like Notepad.
'        After the ‘Sub ReadyStateChangeHandler()’ line add a
'        new line: ‘Attribute Value.VB_UserMemId = 0’
'3)      Save the file and import it back into the VBA project.
'-------------------------------------------------------------------------
If PUB_XML_HTTP_REQUEST.readyState = 4 Then
'XMLHTTP.readyState
'The read-only readyState represents the state of the request.
'When the request is complete, readyState will have a value of 4.
'-------------------------------------------------------------------------
    If PUB_XML_HTTP_REQUEST.Status = 200 Then
'XMLHTTP.Status
'The read-only status represents the HTTP status code returned by
'the request. A value of 200 indicates it was processed successfully.
'For other codes see the link below.
'-------------------------------------------------------------------------
        'Process the response here
        If PUB_FUNCTION_STR <> "" Then
            Call Excel.Application.Run( _
                 PUB_FUNCTION_STR, _
                 PUB_SRC_URL_STR, _
                 PUB_XML_HTTP_REQUEST.ResponseText, _
                 PUB_COMMAND_STR, _
                 PUB_PARAM_RNG)
        End If
        ResponseText = PUB_XML_HTTP_REQUEST.ResponseText
        PUB_AVAILABLE_FLAG = True
    Else
      '  Debug.Print PUB_XML_HTTP_REQUEST.Status & ", " & _
                    PUB_XML_HTTP_REQUEST.ResponseText
        PUB_AVAILABLE_FLAG = False
    End If
'-------------------------------------------------------------------------
'XMLHTTP.ResponseText
'The read-only responseText property contains the text of the
'response as a string.
'-------------------------------------------------------------------------
End If
End Sub

Public Sub XMLHttpCall(ByVal METHOD_STR As String, _
ByVal SRC_URL_STR As String, _
ByVal FUNCTION_STR As String, _
Optional ByVal COMMAND_STR As String, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByVal ASYNC_CALL_FLAG As Boolean = True, _
Optional ByVal SEND_MSG_STR As String = "")

Set PUB_XML_HTTP_REQUEST = New MSXML2.XMLHTTP60

If ASYNC_CALL_FLAG Then
    PUB_SRC_URL_STR = SRC_URL_STR
    PUB_SEND_MSG_STR = SEND_MSG_STR
    PUB_FUNCTION_STR = FUNCTION_STR
    PUB_COMMAND_STR = COMMAND_STR
    PUB_PARAM_RNG = PARAM_RNG
'-------------------------------------------------------------------------
    PUB_XML_HTTP_REQUEST.OnReadyStateChange = Me
'XMLHTTP.OnReadyStateChange = funcMyHandler
'If the XMLHttp response is to be processed
'asynchronously, the onreadystatechange property
'specifies the event handler for the readystatechange
'event. In VB this requires some special handling
'that we will discuss in a later section
'-------------------------------------------------------------------------

End If

With PUB_XML_HTTP_REQUEST
    .Open METHOD_STR, SRC_URL_STR, ASYNC_CALL_FLAG
    '.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    .send SEND_MSG_STR
End With

If Not ASYNC_CALL_FLAG Then
    If FUNCTION_STR <> "" Then
        Call Excel.Application.Run(FUNCTION_STR, _
                       SRC_URL_STR, _
                       PUB_XML_HTTP_REQUEST.ResponseText, _
                       PUB_COMMAND_STR, _
                       PARAM_RNG)
    End If
    ResponseText = PUB_XML_HTTP_REQUEST.ResponseText
End If
End Sub
