VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const PUB_MAX_VAL = 256

Dim PUB_PIN_TEXT_STR As String
Dim PUB_PIN_LEN_VAL As Integer
Dim PUB_PIN_TEXT_LEN_VAL As Integer
Dim PUB_PIN_COUNT_VAL As Integer
Dim PUB_PIN_ARR() As String

Public Sub InitialisePIN(ByVal PIN_TEXT_STR As String, _
ByVal PIN_LENGTH_VAL As Integer)

PUB_PIN_TEXT_STR = PIN_TEXT_STR
PUB_PIN_TEXT_LEN_VAL = Len(PIN_TEXT_STR)
PUB_PIN_LEN_VAL = PIN_LENGTH_VAL
Randomize Timer
ReDim PUB_PIN_ARR(0 To 100)
PUB_PIN_COUNT_VAL = 0

End Sub

Public Function CreatePIN() As String

Dim i As Integer
Dim j As Integer
Dim PIN_STR As String

1983:

PIN_STR = ""

'---------------------------------------------------------------
For i = 1 To PUB_PIN_LEN_VAL
'---------------------------------------------------------------
  j = Rnd * PUB_PIN_TEXT_LEN_VAL + 1
'---------------------------------------------------------------
  If j < 1 Then
    j = 1
  ElseIf j > PUB_PIN_TEXT_LEN_VAL Then
    j = PUB_PIN_TEXT_LEN_VAL
  End If
'---------------------------------------------------------------
  PIN_STR = PIN_STR & Mid(PUB_PIN_TEXT_STR, j, 1)
'---------------------------------------------------------------
Next i
'---------------------------------------------------------------

'---------------------------------------------------------------
PUB_PIN_COUNT_VAL = PUB_PIN_COUNT_VAL + 1
If PUB_PIN_COUNT_VAL > UBound(PUB_PIN_ARR) Then
  ReDim Preserve PUB_PIN_ARR(0 To PUB_PIN_COUNT_VAL + 50)
End If
PUB_PIN_ARR(PUB_PIN_COUNT_VAL) = PIN_STR
'---------------------------------------------------------------
For i = 1 To PUB_PIN_COUNT_VAL - 1
  If PIN_STR = PUB_PIN_ARR(i) Then Exit For
Next i
'---------------------------------------------------------------
If i < PUB_PIN_COUNT_VAL Then
  PUB_PIN_COUNT_VAL = PUB_PIN_COUNT_VAL - 1
  GoTo 1983
End If
'---------------------------------------------------------------
CreatePIN = PIN_STR
'---------------------------------------------------------------

End Function

Public Sub ClosePIN()
Erase PUB_PIN_ARR
End Sub
