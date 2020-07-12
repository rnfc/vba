Attribute VB_Name = "WEB_HTML_BROWSER_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private Declare PtrSafe Function FindExecutable Lib "shell32" _
   Alias "FindExecutableA" _
  (ByVal lpFile As String, _
   ByVal lpDirectory As String, _
   ByVal sResult As String) As Long

Private Declare PtrSafe Function GetTempPath Lib "kernel32" _
   Alias "GetTempPathA" _
  (ByVal i As Long, _
   ByVal lpBuffer As String) As Long

Private Const MAX_PATH As Long = 260
Private Const ERROR_FILE_NO_ASSOCIATION As Long = 31
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_PATH_NOT_FOUND As Long = 3
Private Const ERROR_FILE_SUCCESS As Long = 32 'my constant
Private Const ERROR_BAD_FORMAT As Long = 11


Function RNG_ADD_BROWSER_FUNC(ByVal SRC_URL_STR As String, _
ByRef DST_RNG As Excel.Range)

Dim TOP_VAL As Double
Dim LEFT_VAL As Double
Dim WIDTH_VAL As Double
Dim HEIGHT_VAL As Double

Dim TEMP_OBJ As OLEObject

RNG_ADD_BROWSER_FUNC = False

HEIGHT_VAL = DST_RNG.Height
WIDTH_VAL = DST_RNG.Width
LEFT_VAL = DST_RNG.Left
TOP_VAL = DST_RNG.Top

For Each TEMP_OBJ In DST_RNG.Worksheet.OLEObjects
    If TypeName(TEMP_OBJ.Object) = "WebBrowser" Then
        TEMP_OBJ.Delete
    End If
Next TEMP_OBJ

Set TEMP_OBJ = _
        DST_RNG.Worksheet.OLEObjects.Add(ClassType:="Shell.Explorer.2", _
        link:=False, _
        DisplayAsIcon:=False, _
        Left:=LEFT_VAL, _
        Top:=TOP_VAL, _
        Width:=WIDTH_VAL, _
        Height:=HEIGHT_VAL)

TEMP_OBJ.Object.Navigate2 SRC_URL_STR
TEMP_OBJ.Object.Visible = True

TEMP_OBJ.Height = HEIGHT_VAL
TEMP_OBJ.Width = WIDTH_VAL
TEMP_OBJ.Left = LEFT_VAL
TEMP_OBJ.Top = TOP_VAL

TEMP_OBJ.Activate

RNG_ADD_BROWSER_FUNC = True

Exit Function
ERROR_LABEL:
RNG_ADD_BROWSER_FUNC = False
End Function


Function OPEN_WEB_BROWSER_FUNC(Optional ByVal URL_STR_NAME As String = "")

Dim i As Long
Dim j As Long
On Error GoTo ERROR_LABEL
OPEN_WEB_BROWSER_FUNC = False
j = Shell(WEB_BROWSER_NAME_FUNC(i) & " " & URL_STR_NAME, 1)
OPEN_WEB_BROWSER_FUNC = True

Exit Function
ERROR_LABEL:
OPEN_WEB_BROWSER_FUNC = False
End Function



'The routine works by first determining the system temp folder, creating
'a dummy filename there with the extension .html, then passing that
'information to FindExecutable. The API returns the path and application
'associated with the passed file .html extension, and the temporary file
'is deleted. Also passed to the function is a flags parameter which is
'filled with the return code of the API call. The return value can be
'handled as needed by the app.


Function WEB_BROWSER_TEMP_DIR_FUNC()

Dim i As Long
Dim TEMP_STR As String

On Error GoTo ERROR_LABEL
    
TEMP_STR = Space$(MAX_PATH)
i = Len(TEMP_STR)
Call GetTempPath(i, TEMP_STR)
WEB_BROWSER_TEMP_DIR_FUNC = WEB_BROWSER_TRIM_NULL_FUNC(TEMP_STR)

Exit Function
ERROR_LABEL:
WEB_BROWSER_TEMP_DIR_FUNC = Err.number
End Function
 
 
 'FindExecutable: Obtain Exe of the Default Browser
' CreateProcess: Start Separate Instances of the Default Browser
'FindExecutable: Find Exe Associated with a Registered Extension
'RegSetValueEx: Create a Registered File Association
 
'The routine here will return the full drive, path and filename of the
'user's default application associated with html files. Typically,
'this will always be the user's default browser.

Function WEB_BROWSER_NAME_FUNC(ByRef kk As Long) As String

   Dim i As Long
   
   Dim ATEMP_STR As String 'sResult
   Dim BTEMP_STR As String 'sTempFolder
   
   On Error GoTo ERROR_LABEL
        
  'get the user's temp folder
   BTEMP_STR = WEB_BROWSER_TEMP_DIR_FUNC()
   
  'create a dummy html file in the temp dir
   i = FreeFile
      Open BTEMP_STR & "dummy.html" For Output As #i
   Close #i

  'get the file path & name associated with the file
   ATEMP_STR = Space$(MAX_PATH)
   kk = FindExecutable("dummy.html", BTEMP_STR, ATEMP_STR)
  
  'clean up
   Kill BTEMP_STR & "dummy.html"
'Note however that this code can not guarantee the associated application
'returned is a browser, only that *some* application is associated with
'the specified extension. While in most cases this will be the default
'browser, it is possible that another browser or application has hijacked
'the html extension.  One reported case indicated MS Word was associated
'with the html file extension.  Therefore it may be justified to check the
'returned string from this call for "iexplore.exe" or the name of the
'browser of your choice, and if not the expected value to prompt the user
'to select their default browser, which you'd then save as your own setting
'for future reference. Another alternative would be to present the user with
'the list of installed browsers contained under the registry key:
'    HKEY_LOCAL_MACHINE\SOFTWARE\Clients\StartMenuInternet
'... or to retrieve the application related to the http:// command listed under:
'    HKEY_LOCAL_MACHINE\SOFTWARE\Classes\http\shell\open\command
   
  'return result
   WEB_BROWSER_NAME_FUNC = WEB_BROWSER_TRIM_NULL_FUNC(ATEMP_STR)

Exit Function
ERROR_LABEL:
WEB_BROWSER_NAME_FUNC = Err.number
End Function


Private Function WEB_BROWSER_TRIM_NULL_FUNC(ByRef DATA_STR As String)
    
    Dim j As Integer 'pos
    
    On Error GoTo ERROR_LABEL
    
    j = InStr(DATA_STR, Chr$(0))
    If j Then
       WEB_BROWSER_TRIM_NULL_FUNC = Left$(DATA_STR, j - 1)
    Else
       WEB_BROWSER_TRIM_NULL_FUNC = DATA_STR
    End If

Exit Function
ERROR_LABEL:
WEB_BROWSER_TRIM_NULL_FUNC = Err.number
End Function
