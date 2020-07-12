Attribute VB_Name = "WEB_XML_TEST_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'YAHOO_HISTORICAL_DATA_URL_FUNC

Sub TEST_XML_ASYNC()

Dim i As Long
Dim NROWS As Long

Dim END_DATE As Date
Dim START_DATE As Date

Dim TICKER_STR As String
Dim DATA_VECTOR As Variant

Dim SRC_URL_STR As String
Dim XML_HTTP_MANAGER_OBJ As New clsXMLHttpManager

On Error GoTo ERROR_LABEL

END_DATE = Now()
START_DATE = END_DATE - 365 * 10
DATA_VECTOR = Range("NICO")
NROWS = UBound(DATA_VECTOR)

For i = 1 To NROWS
   TICKER_STR = DATA_VECTOR(i, 1)
   SRC_URL_STR = YAHOO_HISTORICAL_DATA_URL_FUNC(TICKER_STR, _
                 Year(START_DATE), Month(START_DATE), Day(START_DATE), _
                 Year(END_DATE), Month(END_DATE), Day(END_DATE), "d")
   Call XML_HTTP_ASYNCHRONOUS_FUNC(SRC_URL_STR, "GET", "", _
        "TEST_XML_PROCEDURE_A_FUNC", , TICKER_STR, XML_HTTP_MANAGER_OBJ)
Next i
   
Exit Sub
ERROR_LABEL:
'ADD MSG HERE; Err.Description
End Sub

Public Function TEST_XML_PROCEDURE_A_FUNC(ByVal SRC_URL_STR As String, _
ByRef RESPONSE_TEXT As String, _
Optional ByVal INDEX_STR As String, _
Optional ByRef PARAM_RNG As Variant)
        
Dim DST_FILE_PATH As String
On Error GoTo ERROR_LABEL

TEST_XML_PROCEDURE_A_FUNC = False

DST_FILE_PATH = "C:\TEST\" & PARAM_RNG & ".txt"
Call WRITE_DATA_TEXT_FILE_FUNC(DST_FILE_PATH, RESPONSE_TEXT, 3)
TEST_XML_PROCEDURE_A_FUNC = True

Exit Function
ERROR_LABEL:
TEST_XML_PROCEDURE_A_FUNC = False
End Function

Sub TEST_XML_SAVE()

Dim i As Long
Dim j As Long
Dim ii As Long
Dim jj As Long

Dim NSIZE As Long

On Error GoTo ERROR_LABEL

ii = 14000
jj = 14010

NSIZE = (jj - ii + 1)

ReDim URL_VECTOR(1 To NSIZE, 1 To 1) As Variant
ReDim DST_VECTOR(1 To NSIZE, 1 To 1) As Variant

i = 1
For j = ii To jj
   URL_VECTOR(i, 1) = _
   "http://www.city-data.com/zips/" & Format(j, "00000") & ".html"
   
   DST_VECTOR(i, 1) = _
   "C:\zips\" & Format(j, "00000") & ".html"
   i = i + 1
Next j
Debug.Print XML_SAVE_WEB_PAGES_FUNC(URL_VECTOR, DST_VECTOR)
   
Exit Sub
ERROR_LABEL:
'ADD MSG HERE; Err.Description
End Sub
