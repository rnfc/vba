Attribute VB_Name = "EXCEL_WSHEET_SECURITY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Sub SET_WSHEET_HASH_FUNC(ByVal SIGNATURE_STR As String, _
Optional ByRef SRC_WSHEET As Worksheet)
    
    On Error Resume Next
        
    If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet
    
    With SRC_WSHEET
        .CustomProperties(1).Delete
        .CustomProperties.Add name:="Signature", value:=SIGNATURE_STR
        .CustomProperties.Add name:="Hash", value:=CALC_WSHEET_HASH_FUNC
    End With

Exit Sub
ERROR_LABEL:
End Sub

Sub CHECK_WSHEET_HASH_FUNC(Optional ByRef SRC_WSHEET As Worksheet)

Dim USER_STR As String
Dim HASH_STR As String

On Error Resume Next

If SRC_WSHEET Is Nothing Then: Set SRC_WSHEET = ActiveSheet

USER_STR = SRC_WSHEET.CustomProperties.Item(2)
If Err.number = 9 Then
  MsgBox "This sheet has not been reviewed"
  Exit Sub
End If
HASH_STR = SRC_WSHEET.CustomProperties.Item(1)

If HASH_STR = CALC_WSHEET_HASH_FUNC() Then
  MsgBox "This sheet has been signed by " & USER_STR, vbInformation
Else
  MsgBox "This sheet has been signed by " & USER_STR & _
    " But the sheet has changed and the signature is no longer valid", _
    vbExclamation
End If
End Sub

Function ADD_WBOOK_SIGNATURE_FUNC(ByVal FILE_STR_NAME As String, _
ByVal SALT_STR As String)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim DATA_STR As String
Dim DUMMY_STR As String
Dim HASH_STR As String
Dim WBOOK_NAME As String
Dim PATH_SEPAR_STR As String
Dim SRC_WBOOK As Excel.Workbook

On Error GoTo ERROR_LABEL

ADD_WBOOK_SIGNATURE_FUNC = False

If Dir(FILE_STR_NAME) = "" Then: GoTo ERROR_LABEL
'MsgBox "k cannot find the file " & FILE_STR_NAME, vbExclamation

PATH_SEPAR_STR = Excel.Application.PathSeparator
j = Len(FILE_STR_NAME)
i = j
Do While Mid(FILE_STR_NAME, i, 1) <> PATH_SEPAR_STR
    i = i - 1
Loop
i = i + 1
WBOOK_NAME = Mid(FILE_STR_NAME, i, j - i + 1)
DUMMY_STR = String(84, 65)
'open target file
Workbooks.Open FILE_STR_NAME
Set SRC_WBOOK = Excel.Application.Workbooks(WBOOK_NAME)

'insert dummy text as custom document property
On Error Resume Next
SRC_WBOOK.CustomDocumentProperties("ID").Delete

On Error Resume Next
SRC_WBOOK.CustomDocumentProperties.Add name:="ID", _
LinkToContent:=False, Type:=msoPropertyTypeString, value:=DUMMY_STR
SRC_WBOOK.Close (True)

On Error GoTo ERROR_LABEL

'read file into memory --> calculating & storing hash
l = FreeFile()
Open FILE_STR_NAME For Binary As l
DATA_STR = String(LOF(1), 0)
Get l, , DATA_STR
k = InStr(1, DATA_STR, DUMMY_STR)

If k > 0 Then
  'calculate hash of document
  HASH_STR = SHA256_ENCRYPTION_FUNC(SALT_STR & DATA_STR & SALT_STR)
  Put l, k + 20, HASH_STR
Else 'error
  GoTo ERROR_LABEL
  'MsgBox "k couldn't insert the hash value", vbExclamation
End If
Close l

ADD_WBOOK_SIGNATURE_FUNC = True

Exit Function
ERROR_LABEL:
On Error Resume Next
Close l
ADD_WBOOK_SIGNATURE_FUNC = False
End Function

Function VIEW_WBOOK_SIGNATURE_FUNC(ByVal FILE_STR_NAME As String, _
Optional ByVal SALT_STR As String = "")

Dim k As Long
Dim l As Long

Dim DATA_STR As String
Dim AHASH_STR As String
Dim BHASH_STR As String
Dim DUMMY_STR As String

On Error GoTo ERROR_LABEL

VIEW_WBOOK_SIGNATURE_FUNC = False

If Dir(FILE_STR_NAME) = "" Then: GoTo ERROR_LABEL
'MsgBox "k cannot find the file " & FILE_STR_NAME, vbExclamation

DUMMY_STR = String(20, 65)

'read file into memory
l = FreeFile()
Open FILE_STR_NAME For Binary As l
DATA_STR = String(LOF(1), 0)
Get l, , DATA_STR
Close l

k = InStr(1, DATA_STR, DUMMY_STR)
If k > 0 Then 'Calculating and comparing hash
  AHASH_STR = Mid$(DATA_STR, k + 20, 64)
  Mid$(DATA_STR, k + 20, 64) = String(64, 65)
  BHASH_STR = SHA256_ENCRYPTION_FUNC(SALT_STR & DATA_STR & SALT_STR)
Else
  GoTo ERROR_LABEL
  'MsgBox "k couldn't retrieve the hash value", vbExclamation
End If

VIEW_WBOOK_SIGNATURE_FUNC = (BHASH_STR = AHASH_STR)
'is signed/unsigned

Exit Function
ERROR_LABEL:
On Error Resume Next
Close l
VIEW_WBOOK_SIGNATURE_FUNC = False
End Function

Private Function CALC_WSHEET_HASH_FUNC()
Const PUB_AUTHOR_NAME As String = "Rafael Nicolas Fermin Cota"
On Error GoTo ERROR_LABEL
'----------------------------------------------------------------------------------
CALC_WSHEET_HASH_FUNC = SHA256_ENCRYPTION_FUNC("Written by " & PUB_AUTHOR_NAME)
'SHA-256 is an algorithm specified by the US National Institute of
'Technology (NIST) for producing a secure hash for a text string (or
'file). Hashes are short strings which can be used to show whether a
'document has been altered, because it is practically impossible to
'alter the text string and still give the same hash.
'http://csrc.nist.gov/publications/fips/fips180-2/fips180-2.pdf
'for the full specification
'----------------------------------------------------------------------------------
Exit Function
ERROR_LABEL:
CALC_WSHEET_HASH_FUNC = Err.number
End Function
