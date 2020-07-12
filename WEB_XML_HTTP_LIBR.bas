Attribute VB_Name = "WEB_XML_HTTP_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.



Public Function XML_HTTP_SYNCHRONOUS_FUNC(ByVal SRC_URL_STR As String, _
Optional ByVal METHOD_STR As String = "GET", _
Optional ByVal SEND_MSG_STR As String = "")

'Using XMLHttp to retrieve data in a synchronous manner
'This, the first use of XMLHttp, essentially duplicates the
'InternetExplorer capability we saw above. The code to use the
'ActiveX object is below. Notice how much simpler it is compared to
'the InternetExplorer version. The Open method’s 3rd argument is False
'to indicate the request should be processed synchronously. This will
'make our code wait at the Send until the request completes processing.
'Once that happens, the webpage’s HTML is in the responseText string
'property.  To use the HTMLDocument object, we simply create one and
'assign its body’s innerHTML property to responseText. Now, we can reuse
'the analyzer modules from earlier.

Dim RESPONSE_TEXT As String
Static CONNECT_OBJ As MSXML2.XMLHTTP60

On Error GoTo ERROR_LABEL

'Call REMOVE_WEB_PAGE_CACHE_FUNC(SRC_URL_STR)

If CONNECT_OBJ Is Nothing Then
    Set CONNECT_OBJ = New MSXML2.XMLHTTP60
End If

'There are instances when one needs to use the POST method to
'initiate a XMLHttp request. The important thing to remember is
'that when using POST the arguments passed to the web server
'must be in the message and not in the URL itself

With CONNECT_OBJ
'--------------------------------------------------------------------------
    .Open METHOD_STR, SRC_URL_STR, False
'--------------------------------------------------------------------------
'XMLHttp.open(strMethod, strUrl, varAsync, strUser, strPassword)
'The open method opens a connection to the server but it doesn’t
'actually send any information across. strMethod specifies the
'GET or POST method strUrl is the URL of interest, varAsync
'indicates whether the response will be process asynchronously
'(varAsync=TRUE) or synchronously, and strUser and strPassword
'are for authentication.
'--------------------------------------------------------------------------
    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'--------------------------------------------------------------------------
    .send SEND_MSG_STR
'XMLHttp.send(varBody);
'The send method actually sends the message to the server. varBody
'is an optional message text. It contains the parameters for the POST method.
'--------------------------------------------------------------------------
    RESPONSE_TEXT = .ResponseText
'--------------------------------------------------------------------------
End With

'       Case 0: XML_HTTP_SYNCHRONOUS_FUNC = RESPONSE_TEXT
'       Case 200: XML_HTTP_SYNCHRONOUS_FUNC = RESPONSE_TEXT

If CONNECT_OBJ.Status = 200 Then
    XML_HTTP_SYNCHRONOUS_FUNC = RESPONSE_TEXT
Else
    GoTo ERROR_LABEL
End If

Exit Function
ERROR_LABEL:
XML_HTTP_SYNCHRONOUS_FUNC = Err.number
End Function

Function XML_HTTP_COLLECTION_DRIVER_FUNC(ByVal SRC_URL_STR As String)
   
Dim RESPONSE_TEXT As String

Static XML_HTTP_OBJ As MSXML2.XMLHTTP60
Static XML_HTP_COLLECTION_OBJ As Collection

On Error GoTo ERROR_LABEL
 
If XML_HTTP_OBJ Is Nothing Then
  Set XML_HTTP_OBJ = New MSXML2.XMLHTTP60
End If

If XML_HTP_COLLECTION_OBJ Is Nothing Then
  Set XML_HTP_COLLECTION_OBJ = New Collection
End If

On Error Resume Next

RESPONSE_TEXT = XML_HTP_COLLECTION_OBJ(SRC_URL_STR)

If Err.number <> 0 Then 'no error, we have a match
  XML_HTTP_OBJ.Open "GET", SRC_URL_STR, False
  XML_HTTP_OBJ.send
  
  If XML_HTTP_OBJ.Status = 200 Then
   RESPONSE_TEXT = XML_HTTP_OBJ.ResponseText
  Else
   GoTo ERROR_LABEL
  End If
  Call XML_HTP_COLLECTION_OBJ.Add(RESPONSE_TEXT, SRC_URL_STR)
  Err.Clear
End If

'  Debug.Print XML_HTP_COLLECTION_OBJ.Count
 
XML_HTTP_COLLECTION_DRIVER_FUNC = RESPONSE_TEXT

Exit Function
ERROR_LABEL:
XML_HTTP_COLLECTION_DRIVER_FUNC = Err.number
End Function

'----------------------------------------------------------------------------
'Using XMLHttp to retrieve data asynchronously

Public Function XML_HTTP_ASYNCHRONOUS_FUNC( _
ByVal SRC_URL_STR As String, _
Optional ByVal METHOD_STR As String = "GET", _
Optional ByVal SEND_MSG_STR As String = "", _
Optional ByVal FUNCTION_STR As String = "", _
Optional ByVal COMMAND_STR As String, _
Optional ByRef PARAM_RNG As Variant, _
Optional ByRef XML_HTTP_MANAGER_OBJ As clsXMLHttpManager)

'--------------------------------------------------------------------------
'NOTE: For METHOD_STR = "POST"
'--------------------------------------------------------------------------
'There are instances when one needs to use the POST method to
'initiate a XMLHTTP60 request.  The important thing to remember
'is that when using POST the arguments passed to the web server
'must be in the message and not in the URL itself.
'--------------------------------------------------------------------------

Dim METHOD_FLAG As Boolean
Dim RESPONSE_TEXT As String

On Error GoTo ERROR_LABEL
    
If XML_HTTP_MANAGER_OBJ Is Nothing Then
    Set XML_HTTP_MANAGER_OBJ = New clsXMLHttpManager
    METHOD_FLAG = False 'Sync
Else
    METHOD_FLAG = True 'ASync
End If

Call XML_HTTP_MANAGER_OBJ.XMLHttpCall( _
     METHOD_STR, SRC_URL_STR, RESPONSE_TEXT, _
     FUNCTION_STR, COMMAND_STR, PARAM_RNG, _
     METHOD_FLAG, SEND_MSG_STR)

XML_HTTP_ASYNCHRONOUS_FUNC = RESPONSE_TEXT

Exit Function
ERROR_LABEL:
XML_HTTP_ASYNCHRONOUS_FUNC = Err.number
End Function

Public Function XML_CHECK_HTTP_CONNECTION_FUNC() As Boolean
Dim MSG_STR As String
On Error Resume Next
XML_CHECK_HTTP_CONNECTION_FUNC = False
Dim CONNECT_OBJ As MSXML2.XMLHTTP60
With CONNECT_OBJ
    .Open "GET", "http://www.google.com"
    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    .send MSG_STR
End With
If Err.number = 0 Then: XML_CHECK_HTTP_CONNECTION_FUNC = True
End Function
