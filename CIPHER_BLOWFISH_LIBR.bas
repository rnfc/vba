Attribute VB_Name = "CIPHER_BLOWFISH_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Function BLOWFISH_ENCRYPT_STRING_FUNC(ByVal DATA_STR As String, _
ByVal PASS_STR As String, _
Optional ByRef BLOWFISH_CLASS_OBJ As clsBlowFish) As String

Dim i As Integer
Dim j As Integer
Dim TEMP_STR As String
Dim ENCODED_STR As String

On Error GoTo ERROR_LABEL

If BLOWFISH_CLASS_OBJ Is Nothing Then
    Set BLOWFISH_CLASS_OBJ = New clsBlowFish
End If

BLOWFISH_CLASS_OBJ.password PASS_STR
TEMP_STR = BLOWFISH_CLASS_OBJ.EncryptString(DATA_STR)

j = Len(TEMP_STR)
For i = 1 To j
  ENCODED_STR = ENCODED_STR & Right("0" & Hex(Asc(Mid$(TEMP_STR, i, 1))), 2)
Next i
  
BLOWFISH_ENCRYPT_STRING_FUNC = ENCODED_STR

Exit Function
ERROR_LABEL:
BLOWFISH_ENCRYPT_STRING_FUNC = Err.number
End Function

Function BLOWFISH_DECRYPT_STRING_FUNC(ByVal ENCODED_STR As String, _
ByVal PASS_STR As String, _
Optional ByRef BLOWFISH_CLASS_OBJ As clsBlowFish) As String

Dim i As Integer
Dim CHR_STR As String

On Error GoTo ERROR_LABEL

If BLOWFISH_CLASS_OBJ Is Nothing Then
    Set BLOWFISH_CLASS_OBJ = New clsBlowFish
End If

BLOWFISH_CLASS_OBJ.password PASS_STR

For i = 1 To Len(ENCODED_STR) / 2
  CHR_STR = CHR_STR & Chr$(CInt("&H" & Mid$(ENCODED_STR, i * 2 - 1, 2)))
Next i

BLOWFISH_DECRYPT_STRING_FUNC = Trim$(BLOWFISH_CLASS_OBJ.DecryptString(CHR_STR))
  
Exit Function
ERROR_LABEL:
BLOWFISH_DECRYPT_STRING_FUNC = Err.number
End Function



Function BLOWFISH_ENCRYPT_FUNC(ByRef DATA_RNG As Variant)

'This routine encrypts the plaintext records. The PIN number
'is used to encrypt the PIN number in the first column, then the
'PIN number is used to encrypt the data string, and the two encrypted
'strings are concatenated and stored in the array.
'(Note that the encrypted string is stored in hex, because some characters
'cannot be stored).

Dim i As Long
Dim NROWS As Long

Dim PIN_STR As String
Dim TEMP_STR As String
Dim DECODE_STR As String
Dim TEMP_VECTOR() As String
Dim DATA_VECTOR As Variant
Dim BLOWFISH_CLASS_OBJ As New clsBlowFish

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS
  PIN_STR = DATA_VECTOR(i, 1)
  BLOWFISH_CLASS_OBJ.password PIN_STR
  
  TEMP_STR = BLOWFISH_CLASS_OBJ.EncryptString(PIN_STR, True)
  DECODE_STR = DATA_VECTOR(i, 2)
  TEMP_VECTOR(i, 1) = TEMP_STR & _
                      BLOWFISH_CLASS_OBJ.EncryptString(DECODE_STR, True)
'  Excel.Application.StatusBar = "Encrypting data: " & _
                            Format(i) & " records completed"
'  DoEvents

Next i

BLOWFISH_ENCRYPT_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
BLOWFISH_ENCRYPT_FUNC = Err.number
End Function

'This routine decrypts all the data in tne EncodedData

Function BLOWFISH_DECRYPT_FUNC(ByRef PIN_RNG As Variant, _
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
Dim BLOWFISH_CLASS_OBJ As New clsBlowFish

On Error GoTo ERROR_LABEL

PIN_VECTOR = PIN_RNG
ENCODE_VECTOR = ENCODE_RNG
If UBound(PIN_VECTOR, 1) <> UBound(ENCODE_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(PIN_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS
  TEMP_VECTOR(i, 1) = ""
  PIN_STR = PIN_VECTOR(i, 1)
  BLOWFISH_CLASS_OBJ.password PIN_STR
  TEMP_STR = BLOWFISH_CLASS_OBJ.EncryptString(PIN_STR, True)
  k = Len(TEMP_STR) 'Key Length
  
  ENCODE_STR = ENCODE_VECTOR(i, 1)

  If TEMP_STR <> Left$(ENCODE_STR, k) Then
    j = j + 1
    GoTo 1983
  Else
    DECODE_STR = Mid$(ENCODE_STR, k + 1)
    DECODE_STR = Trim$(BLOWFISH_CLASS_OBJ.DecryptString(DECODE_STR, True))
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

BLOWFISH_DECRYPT_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
BLOWFISH_DECRYPT_FUNC = Err.number
End Function
