Attribute VB_Name = "CIPHER_CRYPTO_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private Const PUB_CHUNK_VAL = 100000
Private CRYPTO_CLASS_OBJ As New clsCrypto

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'// These algorithms encodes/decodes at about 200k per second
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------

Function ENCRYPT_TEXT_MSG_FUNC(ByVal DATA_STR As String, _
ByVal PASSWORD_STR As String)

On Error GoTo ERROR_LABEL

CRYPTO_CLASS_OBJ.password = PASSWORD_STR
CRYPTO_CLASS_OBJ.OutBuffer = ""
CRYPTO_CLASS_OBJ.InBuffer = DATA_STR
CRYPTO_CLASS_OBJ.Encrypt
ENCRYPT_TEXT_MSG_FUNC = CRYPTO_CLASS_OBJ.OutBuffer

Exit Function
ERROR_LABEL:
ENCRYPT_TEXT_MSG_FUNC = Err.number
End Function

Function DECRYPT_TEXT_MSG_FUNC(ByVal DATA_STR As String, _
ByVal PASSWORD_STR As String)

On Error GoTo ERROR_LABEL

CRYPTO_CLASS_OBJ.password = PASSWORD_STR
'CRYPTO_CLASS_OBJ.OutBuffer = ""
CRYPTO_CLASS_OBJ.InBuffer = DATA_STR
CRYPTO_CLASS_OBJ.Decrypt
DECRYPT_TEXT_MSG_FUNC = CRYPTO_CLASS_OBJ.OutBuffer

Exit Function
ERROR_LABEL:
DECRYPT_TEXT_MSG_FUNC = Err.number
End Function


Function ENCRYPT_FILE_FUNC(ByVal ORIGINAL_FILE_STR As String, _
ByVal ENCRYPTED_FILE_STR As String, _
ByVal PASSWORD_STR As String)

'Debug.Print ENCRYPT_FILE_FUNC( _
"C:\Documents and Settings\HOME\Desktop\NICO.xls", _
"C:\Documents and Settings\HOME\Desktop\NICO.DATA_STR", "Rafael")

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long

Dim EOF_FLAG As Boolean
Dim DATA_STR As String

  On Error GoTo ERROR_LABEL

  ENCRYPT_FILE_FUNC = False
    
  CRYPTO_CLASS_OBJ.password = PASSWORD_STR

  CRYPTO_CLASS_OBJ.OutBuffer = ""
  
  'Application.StatusBar = "Reading file"
  
  NSIZE = FileLen(ORIGINAL_FILE_STR) / PUB_CHUNK_VAL
  DATA_STR = Space$(PUB_CHUNK_VAL)
  
  j = FreeFile
  k = j + 1

  Open ENCRYPTED_FILE_STR For Output As j: Close j
  
  Open ORIGINAL_FILE_STR For Binary As j
  Open ENCRYPTED_FILE_STR For Binary As k
  EOF_FLAG = False

  Do While EOF_FLAG = False

    Get j, , DATA_STR
    If EOF(j) Then
      DATA_STR = Space$(FileLen(ORIGINAL_FILE_STR) - i * PUB_CHUNK_VAL)
      Get j, i * PUB_CHUNK_VAL + 1, DATA_STR
      EOF_FLAG = True
    End If
    
    CRYPTO_CLASS_OBJ.InBuffer = DATA_STR
    CRYPTO_CLASS_OBJ.Encrypt
    DATA_STR = CRYPTO_CLASS_OBJ.OutBuffer
    
    Put k, , DATA_STR

    i = i + 1: If i > NSIZE Then i = NSIZE
    'Application.StatusBar = "Encrypting.. " & Format(i / NSIZE, "0%")
  Loop
  Close j
  Close k

  ENCRYPT_FILE_FUNC = True

Exit Function
ERROR_LABEL:
On Error Resume Next
Close j
Close k
ENCRYPT_FILE_FUNC = False
End Function

Function DECRYPT_FILE_FUNC(ByVal DECRYPTED_FILE_STR As String, _
ByVal ENCRYPTED_FILE_STR As String, _
ByVal PASSWORD_STR As String)

'Debug.Print DECRYPT_FILE_FUNC( _
"C:\Documents and Settings\HOME\Desktop\NICO.xls", _
"C:\Documents and Settings\HOME\Desktop\NICO.DATA_STR", _
"Rafael")

Dim i As Long
Dim j As Long
Dim k As Long
Dim NSIZE As Long

Dim DATA_STR As String
Dim EOF_FLAG As Boolean

  On Error GoTo ERROR_LABEL
  
  DECRYPT_FILE_FUNC = False
  
  CRYPTO_CLASS_OBJ.password = PASSWORD_STR
  
  'Application.StatusBar = "Reading encrypted data"
  NSIZE = FileLen(ENCRYPTED_FILE_STR) / PUB_CHUNK_VAL
  DATA_STR = Space$(PUB_CHUNK_VAL)
  
  j = FreeFile
  k = j + 1
  Open DECRYPTED_FILE_STR For Output As j: Close j
  Open ENCRYPTED_FILE_STR For Binary As j
  Open DECRYPTED_FILE_STR For Binary As k
  EOF_FLAG = False

  Do While EOF_FLAG = False

    Get j, , DATA_STR
    If EOF(j) Then
      DATA_STR = Space$(FileLen(ENCRYPTED_FILE_STR) - i * PUB_CHUNK_VAL)
      Get j, i * PUB_CHUNK_VAL + 1, DATA_STR
      EOF_FLAG = True
    End If
    
    CRYPTO_CLASS_OBJ.InBuffer = DATA_STR
    CRYPTO_CLASS_OBJ.Decrypt

    DATA_STR = CRYPTO_CLASS_OBJ.OutBuffer
  
    Put k, , DATA_STR

    i = i + 1: If i > NSIZE Then i = NSIZE
    'Application.StatusBar = "Decrypting.. " & Format(i / NSIZE, "0%")
  Loop
  Close j
  Close k

  DECRYPT_FILE_FUNC = True

Exit Function
ERROR_LABEL:
On Error Resume Next
Close j
Close k
DECRYPT_FILE_FUNC = False
End Function
