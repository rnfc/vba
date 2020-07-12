Attribute VB_Name = "CIPHER_PIN_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'This routine creates the PIN numbers and also concatenates the data as a
'comma-separated string. The results are put in the array, with
'the PIN number in the first column.

Function CREATE_PIN_TEXT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal PIN_STR As String = "123456789ABCDEFGHJKLMNPQRSTUVWXYZ", _
Optional ByVal KEY_LENGTH As Integer = 8)

'----------------------------------------------------------------------------------
'Statistics
'Len(PIN_STR) --> 33  characters long
'KEY_LENGTH * Len(PIN_STR) --> 1.40641E+12 Combinations
'----------------------------------------------------------------------------------

'PIN_STR: This PIN string contains the characters which can
'be used in the PIN number. The more of them, the harder it
'will be for someone to guess a PIN.

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_STR As String
Dim TEMP_MATRIX() As String
Dim PIN_NUMBER_OBJ As New clsPIN
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

PIN_NUMBER_OBJ.InitialisePIN PIN_STR, KEY_LENGTH
For i = 1 To NROWS
    TEMP_MATRIX(i, 1) = PIN_NUMBER_OBJ.CreatePIN
    TEMP_STR = ""
    For j = 1 To NCOLUMNS
        TEMP_STR = TEMP_STR & Trim$(DATA_MATRIX(i, j)) & ","
    Next j
    TEMP_MATRIX(i, 2) = Left$(TEMP_STR, Len(TEMP_STR) - 1)
Next i

PIN_NUMBER_OBJ.ClosePIN
CREATE_PIN_TEXT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CREATE_PIN_TEXT_FUNC = Err.number
End Function

'Randomize Timer
'For i = 1 To 15
  'RANDOM_PASSWORD_FUNC = RANDOM_PASSWORD_FUNC & Chr(Rnd * 92 + 34)
'Next i

