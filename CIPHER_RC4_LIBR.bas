Attribute VB_Name = "CIPHER_RC4_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'RC4-encryption function
'RC4 is a stream cipher designed by Rivest for RSA Security
'Data String must be without spaces

Function RC_4_ENCRYPTION_FUNC(ByVal DATA_STR As String, _
ByVal PASSWORD_STR As String)

Dim i As Long
Dim j As Long
Dim k As Long
Dim KEY_ARR() As Byte
Dim BYTE_ARR() As Byte
Dim TEMP_BYTE As Byte
Dim TEMP_ARR(0 To 255) As Integer

On Error GoTo ERROR_LABEL

If Len(PASSWORD_STR) = 0 Then: GoTo ERROR_LABEL
If Len(DATA_STR) = 0 Then: GoTo ERROR_LABEL

On Error Resume Next

If Len(PASSWORD_STR) > 256 Then
    KEY_ARR() = StrConv(Left$(PASSWORD_STR, 256), vbFromUnicode)
Else
    KEY_ARR() = StrConv(PASSWORD_STR, vbFromUnicode)
End If
For i = 0 To 255
    TEMP_ARR(i) = i
Next i

i = 0: j = 0: k = 0
For i = 0 To 255
    j = (j + TEMP_ARR(i) + KEY_ARR(i Mod Len(PASSWORD_STR))) Mod 256
    TEMP_BYTE = TEMP_ARR(i)
    TEMP_ARR(i) = TEMP_ARR(j)
    TEMP_ARR(j) = TEMP_BYTE
Next i

i = 0: j = 0: k = 0
BYTE_ARR() = StrConv(DATA_STR, vbFromUnicode)
For i = 0 To Len(DATA_STR)
    j = (j + 1) Mod 256
    k = (k + TEMP_ARR(j)) Mod 256
    TEMP_BYTE = TEMP_ARR(j)
    TEMP_ARR(j) = TEMP_ARR(k)
    TEMP_ARR(k) = TEMP_BYTE
    BYTE_ARR(i) = BYTE_ARR(i) Xor (TEMP_ARR((TEMP_ARR(j) + TEMP_ARR(k)) Mod 256))
Next i
RC_4_ENCRYPTION_FUNC = StrConv(BYTE_ARR, vbUnicode)

Exit Function
ERROR_LABEL:
RC_4_ENCRYPTION_FUNC = Err.number
End Function
