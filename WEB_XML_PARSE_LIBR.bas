Attribute VB_Name = "WEB_XML_PARSE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function PARSE_XML_JSON_FUNC(ByVal SRC_URL_STR As String, _
ByVal XSL_ENCODE_FILE_PATH As String)

'Debug.Print PARSE_XML_JSON_FUNC( _
"http://search.yahooapis.com/WebSearchService/V1/webSearch?appid=YahooDemo&query=tushar+mehta&results=20", _
"C:\Documents and Settings\HOME\Desktop\EUM\PART_1\WEB_SESSION\yahoo_result_xls.xsl")

Static XML_DOC_OBJ As MSXML2.DOMDocument
Static XSL_DOC_OBJ As MSXML2.DOMDocument
Static CONNECT_OBJ As MSXML2.XMLHTTP60

On Error GoTo ERROR_LABEL

Set CONNECT_OBJ = New MSXML2.XMLHTTP60
Set XML_DOC_OBJ = New DOMDocument
Set XSL_DOC_OBJ = New DOMDocument

XML_DOC_OBJ.async = False
XSL_DOC_OBJ.async = False

XSL_DOC_OBJ.Load XSL_ENCODE_FILE_PATH
    
With CONNECT_OBJ
    .Open "GET", SRC_URL_STR, False
    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    .send ""
End With

If CONNECT_OBJ.Status = 200 Then
    Set XML_DOC_OBJ = CONNECT_OBJ.responseXML
    'The read-only responseXML property returns an object containing the
    'parsed XML document.
Else
    GoTo ERROR_LABEL
End If

PARSE_XML_JSON_FUNC = XML_DOC_OBJ.transformNode(XSL_DOC_OBJ)

Exit Function
ERROR_LABEL:
PARSE_XML_JSON_FUNC = Err.number
End Function


Public Function PARSE_XML_HEADERS_FUNC(ByVal SRC_URL_STR As String)

    Dim i As Long
    Dim j As Long
    
    Dim TEMP_ARR() As String
    Dim TEMP_VECTOR As Variant
    Static CONNECT_OBJ As MSXML2.XMLHTTP60
    
    On Error GoTo ERROR_LABEL
    
    If CONNECT_OBJ Is Nothing Then
        Set CONNECT_OBJ = New MSXML2.XMLHTTP60
    End If
    
    With CONNECT_OBJ
        .Open "GET", SRC_URL_STR, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send ""
        TEMP_VECTOR = .getAllResponseHeaders
    End With
    TEMP_VECTOR = Replace(TEMP_VECTOR, Chr(13), "")
    TEMP_VECTOR = Split(TEMP_VECTOR, Chr(10), -1)
'-------------------------------------------------------------------------------
    j = 1
    ReDim TEMP_ARR(1 To j)
    For i = LBound(TEMP_VECTOR) To UBound(TEMP_VECTOR)
        If TEMP_VECTOR(i) <> "" Then
            ReDim Preserve TEMP_ARR(1 To j)
            TEMP_ARR(j) = TEMP_VECTOR(i)
            j = j + 1
        End If
    Next i
    ReDim TEMP_VECTOR(1 To j - 1, 1 To 1)
    For i = 1 To j - 1
        TEMP_VECTOR(i, 1) = TEMP_ARR(i)
    Next i
'-------------------------------------------------------------------------------
    If CONNECT_OBJ.Status = 200 Then
        PARSE_XML_HEADERS_FUNC = TEMP_VECTOR
    Else
        GoTo ERROR_LABEL
    End If

Exit Function
ERROR_LABEL:
PARSE_XML_HEADERS_FUNC = Err.number
End Function

