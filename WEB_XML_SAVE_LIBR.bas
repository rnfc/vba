Attribute VB_Name = "WEB_XML_SAVE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Public Function XML_SAVE_WEB_PAGES_FUNC(ByRef SRC_URL_RNG() As String, _
ByRef DST_FILE_PATH_RNG() As String)
   
   Dim i As Long
   Dim NROWS As Long
   
   Dim SRC_URL_STR As String
   Dim FILE_PATH_NAME As String
   
   Dim XML_HTTP_SAVE_CLASS As clsXMLHttpSave
   Dim CONNECT_OBJ As MSXML2.XMLHTTP60
   
   On Error GoTo ERROR_LABEL
      
   XML_SAVE_WEB_PAGES_FUNC = False
   
   If UBound(SRC_URL_RNG, 1) <> UBound(DST_FILE_PATH_RNG, 1) Then
    GoTo ERROR_LABEL
   End If
   NROWS = UBound(SRC_URL_RNG, 1)
   
'------------------------------------------------------------------------------
   For i = 1 To NROWS
'------------------------------------------------------------------------------
      SRC_URL_STR = SRC_URL_RNG(i, 1)
      FILE_PATH_NAME = DST_FILE_PATH_RNG(i, 1)
      Set CONNECT_OBJ = New MSXML2.XMLHTTP60
      Set XML_HTTP_SAVE_CLASS = New clsXMLHttpSave
      XML_HTTP_SAVE_CLASS.Initialize CONNECT_OBJ
      CONNECT_OBJ.OnReadyStateChange = XML_HTTP_SAVE_CLASS
      XML_HTTP_SAVE_CLASS.SaveAPage SRC_URL_STR, FILE_PATH_NAME, 0
'------------------------------------------------------------------------------
   Next i
'------------------------------------------------------------------------------
   
   XML_SAVE_WEB_PAGES_FUNC = True

Exit Function
ERROR_LABEL:
XML_SAVE_WEB_PAGES_FUNC = Err.number
End Function

