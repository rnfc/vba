Attribute VB_Name = "WEB_TEXT_FILE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function WRITE_DATA_TEXT_FILE_FUNC(ByVal FILE_PATH_STR As Variant, _
ByVal DATA_STR As String, _
Optional ByVal VERSION As Integer = 0)

Dim j As Long
    
On Error GoTo ERROR_LABEL
    
WRITE_DATA_TEXT_FILE_FUNC = False

j = FreeFile() 'Returns an Integer value representing the
'next file number available for use by the FileOpen function.

'---------------------------------------------------------------------------
Select Case VERSION
'---------------------------------------------------------------------------
Case 0
'---------------------------------------------------------------------------
    Open FILE_PATH_STR For Output As #j
    Write #j, DATA_STR 'Output Text --> Clean the entire file
'---------------------------------------------------------------------------
Case 1
'---------------------------------------------------------------------------
    Open FILE_PATH_STR For Append As #j
    Print #j, DATA_STR 'Write Text --> In the next line available
'---------------------------------------------------------------------------
Case 2
'---------------------------------------------------------------------------
    Open FILE_PATH_STR For Output As #j
    Print #j, DATA_STR;
    Close #j
'---------------------------------------------------------------------------
Case Else
'---------------------------------------------------------------------------
    Dim BYTES_ARR() As Byte
    j = FreeFile
    Open FILE_PATH_STR For Binary As #j
        BYTES_ARR() = DATA_STR
        Put #j, 1, BYTES_ARR()
    Close #j
'---------------------------------------------------------------------------
End Select
'---------------------------------------------------------------------------

Close #j
WRITE_DATA_TEXT_FILE_FUNC = True

Exit Function
ERROR_LABEL:
On Error Resume Next
Close #j
WRITE_DATA_TEXT_FILE_FUNC = False
End Function

Function WRITE_TEMP_HTML_TEXT_FILE_FUNC( _
Optional ByRef FILE_NAME As String = "wpage.html", _
Optional ByRef TOP_STR As String = "", _
Optional ByVal BODY_STR As String = "")

   Dim k As Integer
   Dim TEMP_FLAG As Boolean
   
   On Error GoTo ERROR_LABEL
        
  'get the user's temp folder
   WRITE_TEMP_HTML_TEXT_FILE_FUNC = False
   FILE_NAME = WEB_BROWSER_TEMP_DIR_FUNC() & FILE_NAME
   
  'create a dummy html file in the temp dir
   k = FreeFile
      Open FILE_NAME For Output As #k
          If TOP_STR <> "" Then: Print #k, TOP_STR;
          If BODY_STR <> "" Then: Print #k, BODY_STR;
   Close #k
  
  TEMP_FLAG = OPEN_WEB_BROWSER_FUNC(FILE_NAME)
  If TEMP_FLAG = True Then: WRITE_TEMP_HTML_TEXT_FILE_FUNC = True 'return result

Exit Function
ERROR_LABEL:
WRITE_TEMP_HTML_TEXT_FILE_FUNC = False
End Function

Function CONVERT_MATRIX_TEXT_FILE_FUNC(ByVal FILE_PATH_STR As String, _
ByRef DATA_RNG As Variant, _
Optional ByVal DELIM_CHR As String = "|", _
Optional ByVal VERSION As Integer = 0)

'VERSION = 0 Then: with consecutive delimiters, without the
'empty string

Dim i As Long
Dim j As Long
Dim k As Long

Dim SROW As Long
Dim NROWS As Long

Dim SCOLUMN As Long
Dim NCOLUMNS As Long

Dim TEMP_STR As String
Dim TEMP_CHR As String
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

CONVERT_MATRIX_TEXT_FILE_FUNC = False

k = FreeFile()
DATA_MATRIX = DATA_RNG

SROW = LBound(DATA_MATRIX, 1)
NROWS = UBound(DATA_MATRIX, 1)

SCOLUMN = LBound(DATA_MATRIX, 2)
NCOLUMNS = UBound(DATA_MATRIX, 2)

Open FILE_PATH_STR For Output Access Write As #k

For i = SROW To NROWS
    TEMP_STR = ""
    For j = SCOLUMN To NCOLUMNS
        If DATA_MATRIX(i, j) = "" Then
            TEMP_CHR = IIf(VERSION = 0, Chr(34) & Chr(34), " ")
        Else
            TEMP_CHR = DATA_MATRIX(i, j)
        End If
        TEMP_STR = TEMP_STR & TEMP_CHR & DELIM_CHR
    Next j
    TEMP_STR = Left(TEMP_STR, Len(TEMP_STR) - Len(DELIM_CHR))
    Print #k, TEMP_STR '& DELIM_CHR
Next i

Close #k

CONVERT_MATRIX_TEXT_FILE_FUNC = True
Exit Function
ERROR_LABEL:
On Error Resume Next
Close #k
CONVERT_MATRIX_TEXT_FILE_FUNC = False
End Function


Function CONVERT_TEXT_FILE_MATRIX_FUNC(ByVal FILE_PATH_STR As String, _
Optional ByVal NROWS As Variant = Null, _
Optional ByVal NCOLUMNS As Variant = Null, _
Optional ByVal DELIM_CHR As String = ",")

Dim i As Long
Dim j As Long
Dim k As Long

Dim ii As Long
Dim jj As Long

Dim TEMP_OBJ As Object
Dim TEMP_LINE As String

Dim TEMP_MATRIX() As Variant
Dim TEMP_GROUP() As Variant

On Error GoTo ERROR_LABEL

Set TEMP_OBJ = CreateObject("Scripting.FileSystemObject")
Set TEMP_OBJ = TEMP_OBJ.GetFile(FILE_PATH_STR)
Set TEMP_OBJ = TEMP_OBJ.OpenAsTextStream(1, -2) '1: Read Opt; -2: Trist Opt

'----------------------------------------------------------------------------
ii = 1 'No. Rows
jj = 1 'No. Columns
'----------------------------------------------------------------------------

    ReDim TEMP_GROUP(1 To ii)
    ReDim TEMP_MATRIX(1 To jj)

    Do While TEMP_OBJ.AtEndOfStream <> True
        TEMP_LINE = TEMP_OBJ.ReadLine
        i = 1: j = 1: k = 1
        i = InStr(i, TEMP_LINE, DELIM_CHR, 1)
        If i = 0 Then: GoTo 1983
        Do 'Getting Columns
            ReDim Preserve TEMP_MATRIX(1 To k)
            If k = 1 Then
                TEMP_MATRIX(k) = Mid(TEMP_LINE, j, i - j)
            Else
                TEMP_MATRIX(k) = Mid(TEMP_LINE, j + 1, i - j - 1)
            End If
            k = k + 1
            j = i
            i = InStr(j + 1, TEMP_LINE, DELIM_CHR, 1)
        Loop While i > 0
        ReDim Preserve TEMP_GROUP(1 To ii)
        TEMP_GROUP(ii) = TEMP_MATRIX
        If UBound(TEMP_GROUP(ii)) >= jj Then: jj = UBound(TEMP_GROUP(ii))
        ii = ii + 1
1983:
    Loop

    TEMP_OBJ.Close

'----------------------------------------------------------------------------
If VarType(NROWS) = vbNull And VarType(NCOLUMNS) = vbNull Then
'----------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To UBound(TEMP_GROUP), 1 To jj)
    For i = 1 To UBound(TEMP_GROUP)
        For j = 1 To UBound(TEMP_GROUP(i))
            TEMP_MATRIX(i, j) = TEMP_GROUP(i)(j)
        Next j
    Next i
'----------------------------------------------------------------------------
Else
'----------------------------------------------------------------------------
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    For i = 1 To NROWS
        If i > UBound(TEMP_GROUP) Then: Exit For
        For j = 1 To NCOLUMNS
            If j > UBound(TEMP_GROUP(i)) Then: Exit For
            TEMP_MATRIX(i, j) = TEMP_GROUP(i)(j)
        Next j
    Next i
'----------------------------------------------------------------------------
End If
'----------------------------------------------------------------------------

CONVERT_TEXT_FILE_MATRIX_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
On Error Resume Next
TEMP_OBJ.Close
CONVERT_TEXT_FILE_MATRIX_FUNC = Err.number
End Function


Function FIND_REPLACE_TEXT_FILE_STRING_FUNC(ByVal FILE_PATH_STR As String, _
ByVal FIND_STR As String, _
ByVal REPLACE_STR As String)

Dim i As Long
Dim ODATA_STR As String
Dim NDATA_STR As String

Const TYPE_STR As String = "txt csv html xml vbs"
Const DELIM_STR As String = ","
Dim TYPES_ARR() As String
Dim EXTENSION_STR As String

On Error GoTo ERROR_LABEL
FIND_REPLACE_TEXT_FILE_STRING_FUNC = False
If Len(Dir(FILE_PATH_STR)) = 0 Then: GoTo ERROR_LABEL

' split the string into an array, using delimiter
TYPES_ARR = Split(Replace(TYPE_STR, " ", DELIM_STR), DELIM_STR)
EXTENSION_STR = LCase$(Right$(FILE_PATH_STR, 3))

If UBound(Filter(TYPES_ARR, EXTENSION_STR)) = -1 Then: GoTo ERROR_LABEL
' open file and read contents
i = FreeFile
Open FILE_PATH_STR For Input As #i
ODATA_STR = Input$(LOF(i), #i)
Close #i
' replace old char with new char
NDATA_STR = Replace(ODATA_STR, FIND_STR, REPLACE_STR)
' reopen file and write new contents
i = FreeFile
Open FILE_PATH_STR For Output As #i
Print #i, NDATA_STR
Close #i
FIND_REPLACE_TEXT_FILE_STRING_FUNC = True

Exit Function
ERROR_LABEL:
FIND_REPLACE_TEXT_FILE_STRING_FUNC = False
End Function

'this function splits text files

Function SPLIT_TEXT_FILE_FUNC(ByVal FILE_PATH_STR As String, _
ByVal CHUNK_VAL As Single, _
Optional ByVal VALIDATE_FLAG As Boolean = False)

'Debug.Print SPLIT_TEXT_FILE_FUNC( _
"C:\Documents and Settings\HOME\Desktop\NICO.txt", 13, True)
'Rafael Nicolas Fermin Cota --> 26 Characters
'First Text File: Rafael Nicola
'Second Text File: s Fermin Cota


Dim i As Long
Dim j As Long
Dim k As Long

Dim TEXT_STR As String
Dim TEMP_PATH As String
Dim SPLIT_FILE_PATH_STR As String

Dim FINISHED_FLAG As Boolean

On Error GoTo ERROR_LABEL

'get user input

SPLIT_TEXT_FILE_FUNC = False

'check source file exists
On Error Resume Next
If Dir(FILE_PATH_STR) = "" Then
'I can't find the file. Please specify the full path and name
  Exit Function
End If
On Error GoTo 0

'initialise
TEMP_PATH = Left(FILE_PATH_STR, InStrRev(FILE_PATH_STR, "\"))
k = FileLen(FILE_PATH_STR) '+ 1

'delete any existing split files
i = 0
Do
  i = i + 1
  SPLIT_FILE_PATH_STR = TEMP_PATH & "Split" & i & ".txt"
  If Dir(SPLIT_FILE_PATH_STR) <> "" Then
    Kill SPLIT_FILE_PATH_STR
  Else
    Exit Do
  End If
Loop


'do the split
i = 0
Open FILE_PATH_STR For Binary As #1
Do
  'the last chunk will be what's left over...
  j = k - i * CHUNK_VAL
  If j > CHUNK_VAL Then
    j = CHUNK_VAL
  Else 'last chunk, mark as FINISHED_FLAG
    FINISHED_FLAG = True
  End If
  'set up dummy string of correct length & read in data
  TEXT_STR = String(j, " ")
  Get #1, , TEXT_STR
  'write to file
  i = i + 1
  SPLIT_FILE_PATH_STR = TEMP_PATH & "Split" & i & ".txt"
  Open SPLIT_FILE_PATH_STR For Binary As #2
  Put #2, , TEXT_STR
  Close #2
Loop While Not FINISHED_FLAG
Close #1

'test result - True parameter tells Excel this is a test only

If VALIDATE_FLAG = True Then
    If COMBINE_TEXT_FILE_SPLITS_FUNC(FILE_PATH_STR, True) = False Then
        GoTo ERROR_LABEL
        'The file hasn't been split successfully
    End If
End If

SPLIT_TEXT_FILE_FUNC = True

Exit Function
ERROR_LABEL:
SPLIT_TEXT_FILE_FUNC = False
End Function

Function COMBINE_TEXT_FILE_SPLITS_FUNC(ByVal FILE_PATH_STR As String, _
Optional ByVal TESTING_FLAG As Boolean = False)

Dim i As Long
Dim l As Long

Dim TEXT_STR As String
Dim TEMP_PATH As String

Dim COMBINE_FILE_PATH_STR As String

On Error GoTo ERROR_LABEL

COMBINE_TEXT_FILE_SPLITS_FUNC = False

'if testing, create a dummy filename to put the results into
'-------------------------------------------------------------------
If TESTING_FLAG Then
'-------------------------------------------------------------------
  COMBINE_FILE_PATH_STR = FILE_PATH_STR 'actual original filename
  i = InStr(COMBINE_FILE_PATH_STR, ".") - 1
  FILE_PATH_STR = Left(COMBINE_FILE_PATH_STR, i) & "_a" & _
  Mid(COMBINE_FILE_PATH_STR, i + 1) 'dummy filename
  
'-------------------------------------------------------------------
Else 'we're combining for real, watch out if file exists already
'-------------------------------------------------------------------
  If Dir(FILE_PATH_STR) <> "" Then
    If MsgBox( _
"The filename you have specified exists already. Do you want to overwrite it", _
vbQuestion + vbYesNo) <> vbYes Then
      COMBINE_TEXT_FILE_SPLITS_FUNC = False
      Exit Function
    End If
  End If
'-------------------------------------------------------------------
End If
'-------------------------------------------------------------------

TEMP_PATH = Left(FILE_PATH_STR, InStrRev(FILE_PATH_STR, "\"))
'do the combining
i = 0
Open FILE_PATH_STR For Binary As #1
Do
  i = i + 1
  TEXT_STR = Dir(TEMP_PATH & "Split" & i & ".txt")
  If TEXT_STR = "" Then Exit Do
  TEXT_STR = String(FileLen(TEMP_PATH & "Split" & i & ".txt"), " ")
  Open TEMP_PATH & "Split" & i & ".txt" For Binary As #2
  Get #2, , TEXT_STR
  Close #2
  Put #1, , TEXT_STR
Loop
Close #1

If i = 1 Then
  Debug.Print "I didn't find any files to combine", vbExclamation
  COMBINE_TEXT_FILE_SPLITS_FUNC = False
  Exit Function
End If

'if testing, compare dummy result with original, bye by byte
'-------------------------------------------------------------------
If TESTING_FLAG Then
'-------------------------------------------------------------------
  Dim ABYTE_ARR() As Byte
  Dim BBYTE_ARR() As Byte
  
  l = FreeFile 'Open file
  Open COMBINE_FILE_PATH_STR For Binary Access Read As #l
  'Size the array to hold the file contents
  ReDim ABYTE_ARR(1 To LOF(l))
  Get #l, , ABYTE_ARR
  Close #l
  
  l = FreeFile 'Open file
  Open FILE_PATH_STR For Binary Access Read As #l
  'Size the array to hold the file contents
  ReDim BBYTE_ARR(1 To LOF(l))
  Get #l, , BBYTE_ARR
  Close #l
  
  Kill FILE_PATH_STR 'kill dummy file
  
  For i = 1 To UBound(ABYTE_ARR)
    If ABYTE_ARR(i) <> BBYTE_ARR(i) Then
      Debug.Print "The split failed at byte " & i
      COMBINE_TEXT_FILE_SPLITS_FUNC = False
      Exit Function
    End If
  Next i
  COMBINE_TEXT_FILE_SPLITS_FUNC = True
'-------------------------------------------------------------------
Else
'-------------------------------------------------------------------
  Debug.Print "The files have been recombined"
  COMBINE_TEXT_FILE_SPLITS_FUNC = True
'-------------------------------------------------------------------
End If
'-------------------------------------------------------------------

Exit Function
ERROR_LABEL:
COMBINE_TEXT_FILE_SPLITS_FUNC = False
'-------------------------------------------------------------------
End Function
'-------------------------------------------------------------------
