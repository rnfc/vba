Attribute VB_Name = "CIPHER_RIJNDAEL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 0       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
Private Const PUB_KEY_LENGTH = 256 'or 128 or 192

'Rijndael is the American Encryption Standard

Public Function RIJNDAEL_ENCRYPT_STRING_FUNC(ByVal DATA_STR As String, _
ByVal PASS_STR As String, _
Optional ByRef RIJNDAEL_CLASS_OBJ As clsRijnDael) As String

Dim i As Long
Dim KEY_ARR(0 To 31) As Byte
Dim BLOCK_ARR(0 To 31) As Byte
Dim ENCODED_STR As String

On Error GoTo ERROR_LABEL

If RIJNDAEL_CLASS_OBJ Is Nothing Then
    Set RIJNDAEL_CLASS_OBJ = New clsRijnDael
End If

RIJNDAEL_CLASS_OBJ.gentables
   
For i = 0 To Len(PASS_STR) - 1
  KEY_ARR(i) = Asc(Mid$(PASS_STR, i + 1, 1))
Next i
RIJNDAEL_CLASS_OBJ.gkey PUB_KEY_LENGTH / 32, PUB_KEY_LENGTH / 32, KEY_ARR

Do While DATA_STR <> ""
  If Len(DATA_STR) < 32 Then DATA_STR = Left(DATA_STR & "                               ", 32)

  For i = 0 To 31
    BLOCK_ARR(i) = Asc(Mid(DATA_STR, i + 1, 1))
  Next

  RIJNDAEL_CLASS_OBJ.Encrypt BLOCK_ARR
  For i = 0 To 31
    ENCODED_STR = ENCODED_STR & Right("0" & Hex(BLOCK_ARR(i)), 2)
  Next

  DATA_STR = Mid(DATA_STR, 33)
Loop

RIJNDAEL_ENCRYPT_STRING_FUNC = ENCODED_STR

Exit Function
ERROR_LABEL:
RIJNDAEL_ENCRYPT_STRING_FUNC = Err.number
End Function

Public Function RIJNDAEL_DECRYPT_STRING_FUNC(ByVal ENCODED_STR As String, _
ByVal PASS_STR As String, _
Optional ByRef RIJNDAEL_CLASS_OBJ As clsRijnDael) As String

Dim i As Long
Dim KEY_ARR(0 To 31) As Byte
Dim BLOCK_ARR(0 To 31) As Byte
Dim TEMP_STR As String

On Error GoTo ERROR_LABEL

If RIJNDAEL_CLASS_OBJ Is Nothing Then
    Set RIJNDAEL_CLASS_OBJ = New clsRijnDael
End If

RIJNDAEL_CLASS_OBJ.gentables

For i = 0 To Len(PASS_STR) - 1
  KEY_ARR(i) = Asc(Mid$(PASS_STR, i + 1, 1))
Next i
RIJNDAEL_CLASS_OBJ.gkey PUB_KEY_LENGTH / 32, PUB_KEY_LENGTH / 32, KEY_ARR

TEMP_STR = ""
Do While ENCODED_STR <> ""
 For i = 0 To 31
   BLOCK_ARR(i) = CInt("&H" & Mid$(ENCODED_STR, i * 2 + 1, 2))
 Next i
 RIJNDAEL_CLASS_OBJ.Decrypt BLOCK_ARR
 
 For i = 0 To 31
   TEMP_STR = TEMP_STR & Chr$(BLOCK_ARR(i))
 Next i

 ENCODED_STR = Mid(ENCODED_STR, 65)
Loop
TEMP_STR = Trim(TEMP_STR)

RIJNDAEL_DECRYPT_STRING_FUNC = TEMP_STR

Exit Function
ERROR_LABEL:
RIJNDAEL_DECRYPT_STRING_FUNC = Err.number
End Function

Function RIJNDAEL_ENCRYPT_FUNC(ByRef DATA_RNG As Variant)

'This routine encrypts the plaintext records. The PIN number
'is used to encrypt the PIN number in the first column, then the
'PIN number is used to encrypt the data string, and the two encrypted
'strings are concatenated and stored in the array.

Dim i As Long

Dim NROWS As Long

Dim PIN_STR As String
Dim TEMP_STR As String
Dim DECODE_STR As String
Dim TEMP_VECTOR() As String
Dim DATA_VECTOR As Variant
Dim RIJNDAEL_CLASS_OBJ As New clsRijnDael

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS
  PIN_STR = DATA_VECTOR(i, 1)
  
  TEMP_STR = RIJNDAEL_ENCRYPT_STRING_FUNC(PIN_STR, PIN_STR, _
             RIJNDAEL_CLASS_OBJ)

  DECODE_STR = DATA_VECTOR(i, 2)
  TEMP_VECTOR(i, 1) = TEMP_STR & RIJNDAEL_ENCRYPT_STRING_FUNC(DECODE_STR, _
                      PIN_STR, RIJNDAEL_CLASS_OBJ)
'  Excel.Application.StatusBar = "Encrypting data: " & _
                            Format(i) & " records completed"
'  DoEvents

Next i

RIJNDAEL_ENCRYPT_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
RIJNDAEL_ENCRYPT_FUNC = Err.number
End Function

'This routine decrypts all the data in tne EncodedData

Function RIJNDAEL_DECRYPT_FUNC(ByRef PIN_RNG As Variant, _
ByRef ENCODE_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long 'Key Length

Dim NROWS As Long

Dim PIN_STR As String
Dim ENCODE_STR As String
Dim DECODE_STR As String
Dim TEMP_STR As String
Dim TEMP_VECTOR() As String
Dim PIN_VECTOR As Variant
Dim ENCODE_VECTOR As Variant
Dim RIJNDAEL_CLASS_OBJ As New clsRijnDael

On Error GoTo ERROR_LABEL

PIN_VECTOR = PIN_RNG
ENCODE_VECTOR = ENCODE_RNG
If UBound(PIN_VECTOR, 1) <> UBound(ENCODE_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(PIN_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS
  TEMP_VECTOR(i, 1) = ""
  PIN_STR = PIN_VECTOR(i, 1)
  
  TEMP_STR = RIJNDAEL_ENCRYPT_STRING_FUNC(PIN_STR, PIN_STR, _
             RIJNDAEL_CLASS_OBJ)
  
  k = Len(TEMP_STR) 'Key Length
  
  ENCODE_STR = ENCODE_VECTOR(i, 1)

  If TEMP_STR <> Left$(ENCODE_STR, k) Then
    j = j + 1
    GoTo 1983
  Else
    DECODE_STR = Mid$(ENCODE_STR, k + 1)
    
    DECODE_STR = Trim$(RIJNDAEL_DECRYPT_STRING_FUNC(DECODE_STR, PIN_STR, _
               RIJNDAEL_CLASS_OBJ))
    
    'If DECODE_STR <> DATA_VECTOR(i, 2) Then j = j + 1
    'compares it with the PlainText data, and reports the total number of errors.
  End If
  
  TEMP_VECTOR(i, 1) = DECODE_STR
1983:
  'Excel.Application.StatusBar = "Decrypting data: " & _
                        Format(tRow) & _
                        " records completed (" & Format(j) & " errors)"
  'DoEvents
Next i

RIJNDAEL_DECRYPT_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
RIJNDAEL_DECRYPT_FUNC = Err.number
End Function
