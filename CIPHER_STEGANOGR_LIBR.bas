Attribute VB_Name = "CIPHER_STEGANOGR_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'Steganography is the art of hiding secret messages in plain sight.
'A general once tattooed a message on a messenger's head, let the
'hair grow back, and sent him on his way. The enemy couldn't find
'the message and let him through.

Function STEGANOGRAPHY_ENCODE_FUNC(ByRef DATA_RNG As Variant, _
ByVal MSG_STR As String, _
Optional ByVal DEPTH_VAL As Integer = 7)

'DEPTH_VAL: Bury it in the Xth digit of each number

Dim i As Integer
Dim j As Integer

Dim k As Long

Dim ii As Long
Dim jj As Long

Dim TEMP_CHR As String
Dim DATA_MATRIX As Variant
Dim DIGITS_ARR() As String

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
'we need 3 numbers per character of the message and 10 at the end
ReDim DIGITS_ARR(1 To Len(MSG_STR) * 3 + 10)
k = UBound(DIGITS_ARR)

j = 1
'loop through message and create number array
For i = 1 To Len(MSG_STR)
  'get ASCII char, pad with 0 so there are always 3 numbers
  TEMP_CHR = Right$("000" & CStr(Asc(Mid$(MSG_STR, i, 1))), 3)
  DIGITS_ARR(j + 2) = Left$(TEMP_CHR, 1)
  DIGITS_ARR(j + 1) = Mid$(TEMP_CHR, 2, 1)
  DIGITS_ARR(j) = Mid$(TEMP_CHR, 3, 1)
  j = j + 3
Next i

'put 10 nines on the end to signal end of message
For i = 1 To 10
  DIGITS_ARR(j) = "9"
  j = j + 1
Next i

'now let's hide the numbers
j = 0
'only hide in constants, no use hiding in cells with formulae
For jj = LBound(DATA_MATRIX, 2) To UBound(DATA_MATRIX, 2)
    For ii = LBound(DATA_MATRIX, 1) To UBound(DATA_MATRIX, 1) 'exit when finished
        If j = k Then: Exit For 'cell has to have decimal places
        i = InStr(DATA_MATRIX(ii, jj), ".")
        If i > 0 Then '.....lots of decimal places
          If IsNumeric(DATA_MATRIX(ii, jj)) And _
            Len(DATA_MATRIX(ii, jj)) > i + DEPTH_VAL Then
            j = j + 1 'insert our decimal
            DATA_MATRIX(ii, jj) = _
                Left(DATA_MATRIX(ii, jj), i + DEPTH_VAL - 1) & _
                DIGITS_ARR(j) & _
                Mid(DATA_MATRIX(ii, jj), i + DEPTH_VAL + 1)
          End If
        End If
    Next ii
Next jj
'tell user if we ran out of places to hide numbers
If j < k Then
  
  STEGANOGRAPHY_ENCODE_FUNC = "I could only place " & j & _
        " digits out of " & UBound(DIGITS_ARR) & ". You need more data!"
Else
  STEGANOGRAPHY_ENCODE_FUNC = DATA_MATRIX
End If

Exit Function
ERROR_LABEL:
STEGANOGRAPHY_ENCODE_FUNC = Err.number
End Function

'this decodes the message
'it works exactly like the encoder, in reverse
Function STEGANOGRAPHY_DECODE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DEPTH_VAL As Integer = 7)

Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim ii As Long
Dim jj As Long

Dim MSG_STR As String
Dim DATA_MATRIX As Variant
Dim DIGITS_ARR() As String

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
ReDim DIGITS_ARR(1 To 1)

j = 0
k = 0 'counts the number of sequential 9's. Get to 10
'and that's the end of message
For jj = LBound(DATA_MATRIX, 2) To UBound(DATA_MATRIX, 2)
    For ii = LBound(DATA_MATRIX, 1) To UBound(DATA_MATRIX, 1)
        If k = 10 Then Exit For 'exit if 10 nines in a row
        'see STEGANOGRAPHY_ENCODE_FUNC comments on how the next bit works
        i = InStr(DATA_MATRIX(ii, jj), ".")
        If i > 0 Then
          If IsNumeric(DATA_MATRIX(ii, jj)) And _
                   Len(DATA_MATRIX(ii, jj)) > i + DEPTH_VAL Then
            j = j + 1
            ReDim Preserve DIGITS_ARR(1 To j)
            DIGITS_ARR(j) = Mid$(DATA_MATRIX(ii, jj), i + DEPTH_VAL, 1)
            'look out for nines, count them
            If DIGITS_ARR(j) = "9" Then
              k = k + 1
            Else 'reset counter if not a nine
              k = 0
            End If
          End If
        End If
    Next ii
Next jj

'turn number string into characters
MSG_STR = ""
For i = 1 To j - 10 Step 3 'skip the last 10, they were just an
'end of message signal
  MSG_STR = MSG_STR & Chr$(DIGITS_ARR(i + 2) * 100 + _
                 DIGITS_ARR(i + 1) * 10 + DIGITS_ARR(i))
Next i

'put the answer in the sheet
STEGANOGRAPHY_DECODE_FUNC = MSG_STR

Exit Function
ERROR_LABEL:
STEGANOGRAPHY_DECODE_FUNC = Err.number
End Function


Function CONFUSE_STRING_FUNC(ByVal DATA_STR As String, _
Optional ByVal VERSION As Integer = 0)
   
Dim i As Long
Dim j As Long
Dim TEMP_STR As String

On Error GoTo ERROR_LABEL

Select Case VERSION
Case 0
   Rnd -4
   For i = 1 To Len(DATA_STR)
        TEMP_STR = TEMP_STR & _
                Chr$(Asc(Mid$(DATA_STR, i)) Xor Rnd * 99)
   Next i
Case Else
   For i = 0 To Len(DATA_STR) \ 4
      For j = 1 To 4
         TEMP_STR = TEMP_STR & Mid$(DATA_STR, (4 * i) + 5 - j, 1)
      Next j
   Next i
End Select

CONFUSE_STRING_FUNC = TEMP_STR

Exit Function
ERROR_LABEL:
CONFUSE_STRING_FUNC = Err.number
End Function


