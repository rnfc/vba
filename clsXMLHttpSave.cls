VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsXMLHttpSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'----------------------------------------------------------------------------

Option Explicit     'Requires that all variables to be declared explicitly.
'Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'----------------------------------------------------------------------------
'There are many instances – almost all, some might say – when it does
'not make sense to wait for the web service to complete its response.
'Rather, the code may move on to some other task in response to another
'user action and we may want to process the web service response when it
'actually completes.

'In this case, our request to the server will not cause our code to wait
'at the Send method but simply continue on. Obviously, we cannot process
'the result “in line” as in the last section.  Instead, we create a
'subroutine that is called as a callback function by our request. This
'callback function should track the state of the request and process the
'result only when the request is completed successfully.

'The XMLHttp object was designed to work with scripting languages and it
'is very easy to use it in that context. For example, in JavaScript,
'given a function named funcMyHandler, one would simply use
'XMLHTTP.OnReadyStateChange = funcMyHandler
'Of course, VBA doesn’t support the above construct in which we are doing
'the equivalent of the AddressOf operator, which is what we would use
'to pass the address of a callback function to an API.  But, we cannot
'use AddressOf in this context since the following will generate a
'compile-time error XMLHttp.onreadystatechange = AddressOf funcMyHandler
'So, what does one do if the AddressOf operator doesn’t work and the
'scripting technique doesn’t either?  It turns out that the
'onReadyStateChange property will accept any object and at run time call
'its default method.  So, we create a custom class, add a method to it and
'designate that method as the class’s default method.  But, wait.
'We can’t assign a default anything to a class in VBA!  So, what next?
'It turns out that while we cannot designate a default for a class through
'the Visual Basic Editor, there’s a workaround that we can use.
'Start with the code to create the asynchronous call. The XMLHttpMonitor is
'the class instance that we will use to implement our workaround. There
'are only two changes in the interaction with the XMLHttp object: the
'specification of the onreadystatechange property and the change from False
'to True in the value of the third argument to the Open method.
'----------------------------------------------------------------------------

Dim PUB_XML_HTTP_REQ As MSXML2.XMLHTTP60
Dim PUB_FILE_NAME_STR As String
Dim PUB_SAVE_METHOD As Integer
Public ResponseText As String

Public Sub Initialize(ByRef XMLHttpRequest As MSXML2.XMLHTTP60)
   Set PUB_XML_HTTP_REQ = XMLHttpRequest
End Sub

Sub ReadyStateChangeHandler()
Attribute ReadyStateChangeHandler.VB_UserMemId = 0
   Dim j As Long
    
'As noted above, we still have to designate ReadyStateChangeHandler as
'the default method for this class.  To make that happen, do the
'following:
'1)      Export and remove this class module.
'2)      Open the exported file in a text editor like Notepad.
'After the ‘Sub ReadyStateChangeHandler()’ line add a new line:
'‘Attribute Value.VB_UserMemId = 0’
'3)      Save the file and import it back into the VBA project.
   
'-------------------------------------------------------------------------
   If PUB_XML_HTTP_REQ.readyState = 4 Then
'-------------------------------------------------------------------------
      If PUB_XML_HTTP_REQ.Status = 200 Then
'-------------------------------------------------------------------------
          ResponseText = PUB_XML_HTTP_REQ.ResponseText
'-------------------------------------------------------------------------
          Select Case PUB_SAVE_METHOD
'-------------------------------------------------------------------------
          Case 0
'-------------------------------------------------------------------------
            j = FreeFile
            Open PUB_FILE_NAME_STR For Output As #j
            Print #j, ResponseText;
            Close #j
'-------------------------------------------------------------------------
          Case Else
'-------------------------------------------------------------------------
            Dim BYTES_ARR() As Byte
            j = FreeFile
            Open PUB_FILE_NAME_STR For Binary As #j
                BYTES_ARR() = ResponseText
                Put #j, 1, BYTES_ARR()
            Close #j
'-------------------------------------------------------------------------
          End Select
'-------------------------------------------------------------------------
      Else
'-------------------------------------------------------------------------
         'Debug.Print PUB_XML_HTTP_REQ.Status & ", " & _
                     PUB_XML_HTTP_REQ.ResponseText
'-------------------------------------------------------------------------
      End If
'-------------------------------------------------------------------------
   End If
'-------------------------------------------------------------------------
End Sub

Public Sub SaveAPage(ByVal SRC_URL_STR As String, _
ByVal DST_PATH_NAME As String, _
ByVal METHOD As Integer)
   PUB_SAVE_METHOD = METHOD
   PUB_FILE_NAME_STR = DST_PATH_NAME
   PUB_XML_HTTP_REQ.Open "GET", SRC_URL_STR, True
   PUB_XML_HTTP_REQ.send ""
End Sub

'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
'Properly encapsulating XMLHttp requests

'The method described here works fine when one tests it.
'However, in a “production” environment it is very likely to fail.
'The reason is that the caller (the consumer of the class) may initiate
'some task that reuses the XMLHttpMonitor before it is done processing
'an earlier request. The first diagram below illustrates when the above code
'works safely. Some user action triggers the code that initiates the
'asynchronous XMLHttp request. Since the completed asynchronous response
'is processed before the next user action, the code works as intended.

'In the next sequence of events, shown below, a second user action triggers
'the asynchronous request before the first has finished processing, which
'means that the global variables XMLHttpReq and XMLHttpMon are still in
'use.  If we reuse them, we would lose the ability to process the response
'to the first request!

'The other issue we should address is that the code that initiates the
'XMLHttp request must declare and support two globals: XMLHttpReq and
'XMLHttpMon. In addition, the code and the objects are in the consumer
'domain, as it were, and we are dependent on individual developers to get
'correctly use the XMLHttpReq and XMLHttpMon objects.
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
