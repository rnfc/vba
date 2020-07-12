VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEcuyer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 0       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Const MAXGEN = 100
Const h = 32768
Const M32 = 4294967296#
Const M16 = 65536

Dim aw(0 To 3) As Long
Dim avw(0 To 3) As Long
Dim A(0 To 3) As Long
Dim m(0 To 3) As Long

Dim Lg(4, 101) As Long
Dim Ig(4, 101) As Long
Dim Cg(4, 101) As Long


Sub Initialize(Optional ByVal V_VALUE As Long = 31, _
Optional ByVal W_VALUE As Long = 41)
Call Init(V_VALUE, W_VALUE)
End Sub

'always start with this to initialise the RNG
'Init 31, 41
'call Rnd as often as you like
'you can have up to 100 different RNGs going at once
'you specify which you want with the parameter you pass through here
'if you only want one set, just pass 1 every time
'MsgBox Format(Rnd(1), "0.000000000000000")
  

Public Function Rnd(g As Long) As Double

Dim k As Long
Dim s As Long
Dim U As Double

s = Cg(0, g)
k = s \ 46693
s = 45991 * (s - k * 46693) - k * 25584
If s < 0 Then s = s + 2147483647
Cg(0, g) = s
U = U + s * 4.6566128752458E-10
If U >= 1 Then U = U - 1

s = Cg(1, g)
k = s \ 10339
s = 207707 * (s - k * 10339) - k * 870
If s < 0 Then s = s + 2147483543
Cg(1, g) = s
U = U - s * 4.65661310075986E-10
If U < 0 Then U = U + 1

s = Cg(2, g)
k = s \ 15499
s = 138556 * (s - k * 15499) - k * 3979
If s < 0 Then s = s + 2147483423
Cg(2, g) = s
U = U + s * 4.65661336096842E-10
If U >= 1 Then U = U - 1

s = Cg(3, g)
k = s \ 43218
s = 49689 * (s - k * 43218) - k * 24121
If s < 0 Then s = s + 2147483323
Cg(3, g) = s
U = U - s * 4.65661357780891E-10
If U < 0 Then U = U + 1

Rnd = U

End Function

Private Function MultModM(s As Long, _
t As Long, _
m As Long) As Long

Dim r As Long
Dim S0 As Long
Dim s1 As Long
Dim q As Long
Dim qh As Long
Dim rh As Long
Dim k As Long

If s < 0 Then s = s + m
If t < 0 Then t = t + m
If s < h Then
  S0 = s
  r = 0
Else
  s1 = s \ h
  S0 = s - h * s1
  qh = m \ h
  rh = m - qh * h
  
  If s1 >= h Then
    s1 = s1 - h
    k = t \ qh
    r = h * (t - k * qh) - k * rh
    Do While r < 0
      r = r + m
    Loop
  Else
    r = 0
  End If
  
  If s1 <> 0 Then
    q = m \ s1
    k = t \ q
    r = r - k * (m - s1 * q)
    If r > 0 Then r = r - m
    r = r + s1 * (t - k * q)
    Do While r < 0
      r = r + m
    Loop
  End If
  
  k = r \ qh
  r = h * (r - k * qh) - k * rh
  Do While r < 0
    r = r + m
  Loop
  
End If

If S0 <> 0 Then
  q = m \ S0
  k = t \ q
  r = r - k * (m - S0 * q)
  If r > 0 Then r = r - m
  r = r + S0 * (t - k * q)
  Do While r < 0
    r = r + m
  Loop
End If

MultModM = r

End Function

Private Sub SetArrays()

If A(0) > 0 Then Exit Sub

A(0) = 45991
A(1) = 207707
A(2) = 138556
A(3) = 49689
m(0) = 2147483647
m(1) = 2147483543
m(2) = 2147483423
m(3) = 2147483323

End Sub

Private Sub InitGenerator(g As Long, _
SeedType As Long)

Dim j As Long

SetArrays

For j = 0 To 3
  
  Select Case SeedType
  Case 1 'initial seed
    Lg(j, g) = Ig(j, g)
  Case 2 'new seed
    Lg(j, g) = MultModM(aw(j), Lg(j, g), m(j))
  Case 3 'last seed
  
  End Select
  
  Cg(j, g) = Lg(j, g)
  
Next j

End Sub

Private Sub Init(V As Long, _
W As Long)

Dim j As Long
Dim i As Long
Dim sd(0 To 3) As Long

SetArrays

For j = 0 To 3
  sd(j) = 11111111 * (j + 1)
  aw(j) = A(j)
  For i = 1 To W
    aw(j) = MultModM(aw(j), aw(j), m(j))
  Next i
  avw(j) = aw(j)
  For i = 1 To V
    avw(j) = MultModM(avw(j), avw(j), m(j))
  Next i
Next j

SetInitialSeed sd()

End Sub


Private Sub SetSeed(g As Long, _
s() As Long)

Dim j As Long
For j = 0 To 3
  Ig(j, g) = s(j)
  InitGenerator g, 1
Next j

End Sub

Private Sub SetInitialSeed(s() As Long)

Dim g As Long
Dim j As Long

For j = 0 To 3
  Ig(j, 0) = s(j)
Next j

InitGenerator 0, 1

For g = 1 To MAXGEN
  For j = 0 To 3
    Ig(j, g) = MultModM(avw(j), Ig(j, g - 1), m(j))
    InitGenerator g, 1
  Next j
Next g

End Sub
