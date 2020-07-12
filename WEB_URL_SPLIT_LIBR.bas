Attribute VB_Name = "WEB_URL_SPLIT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Shlwapi.dll version 5.00 or greater, Windows XP, 2000,
'Windows NT4 with IE 5 or later, Windows 98, or Windows 95
'with IE 5 or later.


Private Const MAX_PATH As Long = 260
Private Const ERROR_SUCCESS As Long = 0

Private Const URL_PART_SCHEME As Long = 1
Private Const URL_PART_HOSTNAME As Long = 2
Private Const URL_PART_USERNAME As Long = 3
Private Const URL_PART_PASSWORD As Long = 4
Private Const URL_PART_PORT As Long = 5
Private Const URL_PART_QUERY As Long = 6

Private Const URL_PARTFLAG_KEEPSCHEME As Long = &H1

Private Declare PtrSafe Function UrlGetPart Lib "shlwapi" _
   Alias "UrlGetPartA" _
  (ByVal pszIn As String, _
   ByVal pszOut As String, _
   pcchOut As Long, _
   ByVal dwPart As Long, _
   ByVal dwFlags As Long) As Long


Function URL_SPLIT_FUNC(ByVal SRC_URL_STR As String)

'Inputs: The full url including http://
' Two variables that will be changed

'Returns: Splits the SRC_URL_STR$ var into the server name
' and the ATEMP_STR path

  Dim i As Long
  Dim ATEMP_STR As String
  Dim BTEMP_STR As String
  
  On Error GoTo ERROR_LABEL

  i = InStr(SRC_URL_STR$, "/")
  BTEMP_STR$ = Mid(SRC_URL_STR$, i + 2, Len(SRC_URL_STR$) - (i + 1))
  i = InStr(BTEMP_STR$, "/")
  ATEMP_STR$ = Mid(BTEMP_STR$, i, Len(BTEMP_STR$) + 1 - i)
  BTEMP_STR$ = Left$(BTEMP_STR$, i - 1)

URL_SPLIT_FUNC = Array(ATEMP_STR, BTEMP_STR)

Exit Function
ERROR_LABEL:
URL_SPLIT_FUNC = Err.number
End Function

Function URL_PARTS_FUNC(ByVal SRC_URL_STR As String, _
ByRef INDEX_NO As Long, _
Optional ByRef FLAG_NO As Long = URL_PART_HOSTNAME)

   Dim ii As Long
   Dim TEMP_STR As String
   
   On Error GoTo ERROR_LABEL
   
   If Len(SRC_URL_STR) > 0 Then
   
      TEMP_STR = Space$(MAX_PATH)
      ii = Len(TEMP_STR)
     
      If UrlGetPart(SRC_URL_STR, _
                    TEMP_STR, _
                    ii, _
                    INDEX_NO, _
                    FLAG_NO) = ERROR_SUCCESS Then
   
         URL_PARTS_FUNC = Left$(TEMP_STR, ii)
         
      End If  'If UrlGetPart
   Else
      GoTo ERROR_LABEL
   End If 'If Len(SRC_URL_STR) > 0

Exit Function
ERROR_LABEL:
URL_PARTS_FUNC = Err.number
End Function
