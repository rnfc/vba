Attribute VB_Name = "STAT_DIST_NORMAL_UNIVAR_LIBR"

'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------


'************************************************************************************
'************************************************************************************
'FUNCTION      : UNIVAR_CUMUL_NORM_FUNC
'DESCRIPTION   : Univariate Normal Distribution Functions
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_UNIVAR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function UNIVAR_CUMUL_NORM_FUNC(ByVal UNIVAR_TYPE As String, _
ByVal Z_VAL As Double) As Double
   
On Error GoTo ERROR_LABEL

  Select Case UNIVAR_TYPE
'------------------------------------------------------------------------------------
  Case "ab & steg"
    UNIVAR_CUMUL_NORM_FUNC = ABRAM_STEG_CUMUL_NORM_FUNC(Z_VAL)
    Exit Function
'------------------------------------------------------------------------------------
  Case "ab & steg fix"
    UNIVAR_CUMUL_NORM_FUNC = FIX_CUMUL_NORM_FUNC(Z_VAL)
    Exit Function
'------------------------------------------------------------------------------------
' For dim=1 a series approach by Marsaglia and the solution given at Genz
' are compared. It turns out that for small values an interpolation should
' be used (cdfN_Hart), for medium size Marsaglia cdfN_Marsaglia is worth its
' cost and for larger ones an asymptotic is the choice.

  Case "hart"
1983:
    UNIVAR_CUMUL_NORM_FUNC = HART_CUMUL_NORM_FUNC(Z_VAL)
    Exit Function
'------------------------------------------------------------------------------------
  Case "Marsaglia_0"
    UNIVAR_CUMUL_NORM_FUNC = MARSAG1_CUMUL_NORM_FUNC(Z_VAL)
    Exit Function
'------------------------------------------------------------------------------------
  Case "Marsaglia"
1984:
    UNIVAR_CUMUL_NORM_FUNC = MARSAG2_CUMUL_NORM_FUNC(Z_VAL)
    Exit Function
'------------------------------------------------------------------------------------
  Case "asymptotic"
1985:
    UNIVAR_CUMUL_NORM_FUNC = ASYMP_CUMUL_NORM_FUNC(Z_VAL, 1000)
    Exit Function
'------------------------------------------------------------------------------------
  Case Else
      If Abs(Z_VAL) < 4 Then: GoTo 1983
      If Abs(Z_VAL) < 7.4 Then: GoTo 1984
      GoTo 1985
  End Select
'------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
UNIVAR_CUMUL_NORM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : TEST_UNIVAR_CUMUL_NORM_FUNC
'DESCRIPTION   : Test the performance of univariate normal distribution functions
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_UNIVAR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Sub TEST_UNIVAR_CUMUL_NORM_FUNC()

Dim i As Long
Dim nLOOPS As Long

Dim START_TIME As Double

Dim TEMP_VAL As Double
Dim TEMP_DELTA As Double

On Error GoTo ERROR_LABEL

'ab & steg: 24.21875
'ab & steg fix: 26.5625
'hart: 18.75
'Marsaglia: 60.15625
'Marsaglia_0: 113.28125
'gmNorm: 778.90625
'myN: 1158.59375
'Excel    : 100.78125

TEMP_DELTA = 7         ' start in TEMP_DELTA = TEMP_DELTA
nLOOPS = 100000 ' end in TEMP_DELTA + 1 after 100000 steps
'UNIVAR_CUMUL_NORM_FUNC

START_TIME = Timer
For i = 0 To nLOOPS
  TEMP_VAL = UNIVAR_CUMUL_NORM_FUNC("ab & steg", -i / 100000 + TEMP_DELTA)
Next i
Debug.Print "ab & steg: " & (Timer - START_TIME) * 100

START_TIME = Timer
For i = 0 To nLOOPS
  TEMP_VAL = UNIVAR_CUMUL_NORM_FUNC("ab & steg fix", -i / 100000 + TEMP_DELTA)
Next i
Debug.Print "ab & steg fix: " & (Timer - START_TIME) * 100

START_TIME = Timer
For i = 0 To nLOOPS
  TEMP_VAL = UNIVAR_CUMUL_NORM_FUNC("hart", -i / 100000 + TEMP_DELTA)
Next i
Debug.Print "hart: " & (Timer - START_TIME) * 100

START_TIME = Timer
For i = 0 To nLOOPS
  TEMP_VAL = UNIVAR_CUMUL_NORM_FUNC("Marsaglia", -i / 100000 + TEMP_DELTA)
Next i
Debug.Print "Marsaglia: " & (Timer - START_TIME) * 100

START_TIME = Timer
For i = 0 To nLOOPS
  TEMP_VAL = UNIVAR_CUMUL_NORM_FUNC("Marsaglia_0", -i / 100000 + TEMP_DELTA)
Next i
Debug.Print "Marsaglia_0: " & (Timer - START_TIME) * 100

START_TIME = Timer
For i = 0 To nLOOPS
  TEMP_VAL = UNIVAR_CUMUL_NORM_FUNC("asymptotic", -i / 100000 + TEMP_DELTA)
Next i
Debug.Print "Asymtotic: " & (Timer - START_TIME) * 100

START_TIME = Timer
For i = 0 To nLOOPS
  TEMP_VAL = Excel.Application.NormSDist(-i / 100000 + TEMP_DELTA)
Next i
Debug.Print "Excel    : " & (Timer - START_TIME) * 100

Exit Sub
ERROR_LABEL:
'ADD MSG HERE; Err.Description
End Sub


'************************************************************************************
'************************************************************************************
'FUNCTION      : ABRAM_STEG_CUMUL_NORM_FUNC

'DESCRIPTION   : Abramowitz and Stegun - 6dp accuracy - Univariate normal
'distribution function

'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_UNIVAR
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function ABRAM_STEG_CUMUL_NORM_FUNC(ByVal Z_VAL As Double) As Double
  
Dim TEMP_SUM As Double
Dim TEMP_FACTOR As Double

On Error GoTo ERROR_LABEL

TEMP_FACTOR = 1 / (1 + 0.2316419 * Abs(Z_VAL))

TEMP_SUM = 0.31938153 * TEMP_FACTOR - 0.356563782 * _
            TEMP_FACTOR ^ 2 + 1.781477937 * TEMP_FACTOR ^ 3 - _
            1.821255978 * TEMP_FACTOR ^ 4 + 1.330274429 * TEMP_FACTOR ^ 5

If Z_VAL < 0 Then
   ABRAM_STEG_CUMUL_NORM_FUNC = 0.39894228 * _
                            Exp(-Z_VAL ^ 2 / 2) * TEMP_SUM
Else
  ABRAM_STEG_CUMUL_NORM_FUNC = 1 - 0.39894228 * _
                            Exp(-Z_VAL ^ 2 / 2) * TEMP_SUM
End If

Exit Function
ERROR_LABEL:
ABRAM_STEG_CUMUL_NORM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : FIX_CUMUL_NORM_FUNC
'DESCRIPTION   : Fix univariate normal distribution function
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_UNIVAR
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function FIX_CUMUL_NORM_FUNC(ByVal Z_VAL As Double) As Double
  
Dim TEMP_SUM As Double
Dim TEMP_FACTOR As Double
  
On Error GoTo ERROR_LABEL

If Z_VAL = 0 Then
  FIX_CUMUL_NORM_FUNC = 0.5
Else
  TEMP_FACTOR = 1 / (1 + 0.2316419 * Abs(Z_VAL))
  TEMP_SUM = 0.31938153 * TEMP_FACTOR - 0.356563782 * _
             TEMP_FACTOR ^ 2 + 1.781477937 * TEMP_FACTOR ^ 3 - _
             1.821255978 * TEMP_FACTOR ^ 4 + 1.330274429 * TEMP_FACTOR ^ 5
  
  If Z_VAL < 0 Then
     FIX_CUMUL_NORM_FUNC = 0.39894228 * Exp(-Z_VAL ^ 2 / 2) * TEMP_SUM
  Else
    FIX_CUMUL_NORM_FUNC = 1 - 0.39894228 * Exp(-Z_VAL ^ 2 / 2) * TEMP_SUM
  End If
End If
  
Exit Function
ERROR_LABEL:
FIX_CUMUL_NORM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HART_CUMUL_NORM_FUNC
'DESCRIPTION   : Hart Univariate Normal Distribution function

' A series approach by Marsaglia and the solution given at Genz
' are compared. It turns out that for small values an interpolation should
' be used (cdfN_Hart), for medium size Marsaglia cdfN_Marsaglia is worth its
' cost and for larger ones an asymptotic is the choice.

'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_UNIVAR
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HART_CUMUL_NORM_FUNC(ByVal Z_VAL As Double) As Double

Dim TEMP_ABS As Double
Dim TEMP_EXPON As Double
Dim TEMP_BUILD As Double
  
On Error GoTo ERROR_LABEL

TEMP_ABS = Abs(Z_VAL)
If TEMP_ABS > 37 Then
  HART_CUMUL_NORM_FUNC = 0
Else
  TEMP_EXPON = Exp(-TEMP_ABS ^ 2 / 2)
  If TEMP_ABS < 7.07106781186547 Then
    TEMP_BUILD = 3.52624965998911E-02 * TEMP_ABS + 0.700383064443688
    TEMP_BUILD = TEMP_BUILD * TEMP_ABS + 6.37396220353165
    TEMP_BUILD = TEMP_BUILD * TEMP_ABS + 33.912866078383
    TEMP_BUILD = TEMP_BUILD * TEMP_ABS + 112.079291497871
    TEMP_BUILD = TEMP_BUILD * TEMP_ABS + 221.213596169931
    TEMP_BUILD = TEMP_BUILD * TEMP_ABS + 220.206867912376
    
    HART_CUMUL_NORM_FUNC = TEMP_EXPON * TEMP_BUILD
    
    TEMP_BUILD = 8.83883476483184E-02 * TEMP_ABS + 1.75566716318264
    TEMP_BUILD = TEMP_BUILD * TEMP_ABS + 16.064177579207
    TEMP_BUILD = TEMP_BUILD * TEMP_ABS + 86.7807322029461
    TEMP_BUILD = TEMP_BUILD * TEMP_ABS + 296.564248779674
    TEMP_BUILD = TEMP_BUILD * TEMP_ABS + 637.333633378831
    TEMP_BUILD = TEMP_BUILD * TEMP_ABS + 793.826512519948
    TEMP_BUILD = TEMP_BUILD * TEMP_ABS + 440.413735824752
    
    HART_CUMUL_NORM_FUNC = HART_CUMUL_NORM_FUNC / TEMP_BUILD
  Else
    TEMP_BUILD = TEMP_ABS + 0.65
    TEMP_BUILD = TEMP_ABS + 4 / TEMP_BUILD
    TEMP_BUILD = TEMP_ABS + 3 / TEMP_BUILD
    TEMP_BUILD = TEMP_ABS + 2 / TEMP_BUILD
    TEMP_BUILD = TEMP_ABS + 1 / TEMP_BUILD
    
    HART_CUMUL_NORM_FUNC = TEMP_EXPON / TEMP_BUILD / 2.506628274631
  End If
End If
If Z_VAL > 0 Then HART_CUMUL_NORM_FUNC = 1 - HART_CUMUL_NORM_FUNC

Exit Function
ERROR_LABEL:
HART_CUMUL_NORM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASYMP_CUMUL_NORM_FUNC
'DESCRIPTION   : Continued fraction for erfc, cdfN(-x) = erfc( x/Sqr(2) ) / 2
'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_UNIVAR
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function ASYMP_CUMUL_NORM_FUNC(ByVal Z_VAL As Double, _
Optional ByVal nLOOPS As Long = 1000) As Double

Dim i As Long

Dim A_VAL As Double
Dim B_VAL As Double
Dim C_VAL As Double
Dim D_VAL As Double

Dim X_VAL As Double
Dim Y_VAL As Double
Dim W_VAL As Double

Dim SQR_2_VAL As Double
Dim SQR_PI_VAL As Double

On Error GoTo ERROR_LABEL

SQR_2_VAL = 0.707106781186548
SQR_PI_VAL = 0.5723649429247

X_VAL = -Abs(Z_VAL) * SQR_2_VAL

A_VAL = 0#
B_VAL = 1#
C_VAL = 1#
D_VAL = X_VAL

i = 0
Do While B_VAL * C_VAL <> A_VAL * D_VAL
  i = i + 1
  If i > nLOOPS Then: GoTo ERROR_LABEL

  Y_VAL = CDbl(i) - 0.5
  A_VAL = X_VAL * B_VAL + Y_VAL * A_VAL
  C_VAL = X_VAL * D_VAL + Y_VAL * C_VAL
  B_VAL = X_VAL * A_VAL + CDbl(i) * B_VAL
  D_VAL = X_VAL * C_VAL + CDbl(i) * D_VAL
Loop

W_VAL = B_VAL / D_VAL * Exp(-X_VAL * _
            X_VAL - SQR_PI_VAL) ' erfc(X_VAL)
W_VAL = -W_VAL * 0.5

If 0 < Z_VAL Then: W_VAL = 1 - W_VAL

ASYMP_CUMUL_NORM_FUNC = W_VAL

Exit Function
ERROR_LABEL:
ASYMP_CUMUL_NORM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MARSAG1_CUMUL_NORM_FUNC
'DESCRIPTION   : Taylor series around 0 for the cumulative normal due to
' George Marsaglia should give 15 digits, cut off at 7.1 < abs(x)

'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_UNIVAR
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MARSAG1_CUMUL_NORM_FUNC(ByVal Z_VAL As Double) As Double

Dim i As Integer

Dim X_VAL As Double
Dim Y_VAL As Double
Dim W_VAL As Double

Dim PWR_VAL As Double
Dim SQR_VAL As Double

On Error GoTo ERROR_LABEL

If Z_VAL = 0# Then
  MARSAG1_CUMUL_NORM_FUNC = 0.5
  Exit Function
End If

If Abs(Z_VAL) < 7.1 Then ' use Marsaglia
  SQR_VAL = Z_VAL * Z_VAL
  Y_VAL = 0
  X_VAL = 1
  W_VAL = Z_VAL
  PWR_VAL = Z_VAL

  For i = 2 To 200 Step 2
    X_VAL = X_VAL / (i + 1)
    PWR_VAL = PWR_VAL * SQR_VAL
    Y_VAL = W_VAL
    W_VAL = W_VAL + PWR_VAL * X_VAL
    If W_VAL = Y_VAL Then
      Exit For
    End If
  Next

  MARSAG1_CUMUL_NORM_FUNC = 0.5 + W_VAL * Exp(-0.5 * Z_VAL * Z_VAL - 0.918938533204673)
  Exit Function

ElseIf Abs(Z_VAL) < 37 Then    ' use asymptotics
  MARSAG1_CUMUL_NORM_FUNC = -Z_VAL / (1 + Z_VAL * Z_VAL) * _
                            Exp(-0.5 * Z_VAL * Z_VAL - 0.918938533204673)
  
ElseIf (Abs(37) <= Z_VAL) Then ' avoid numerical Nirvana
  MARSAG1_CUMUL_NORM_FUNC = 0
End If

If (0 < Z_VAL) Then: MARSAG1_CUMUL_NORM_FUNC = MARSAG1_CUMUL_NORM_FUNC + 1

Exit Function
ERROR_LABEL:
MARSAG1_CUMUL_NORM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MARSAG2_CUMUL_NORM_FUNC
'DESCRIPTION   : Taylor series for the cumulative normal around various integers cut
' off at 7.1 < abs(x) and should give 15 digits, due to George Marsaglia

'LIBRARY       : STATISTICS
'GROUP         : DIST_NORMAL_UNIVAR
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MARSAG2_CUMUL_NORM_FUNC(ByVal Z_VAL As Double) As Double

Dim i As Integer
Dim j As Integer

Dim A_VAL As Double
Dim B_VAL As Double
Dim W_VAL As Double
Dim V_VAL As Double
Dim X_VAL As Double
Dim Y_VAL As Double

Dim SQR_VAL As Double
Dim TEMP_SUM As Double
Dim TEMP_PWR As Double

Dim TEMP_ARR(0 To 8) As Double

On Error GoTo ERROR_LABEL

TEMP_ARR(0) = 0#
TEMP_ARR(1) = 0.655679542418798 + 4.7154E-16
TEMP_ARR(2) = 0.421369229288054 + 4.7322E-16
TEMP_ARR(3) = 0.304590298710103 + 2.9573E-16
TEMP_ARR(4) = 0.23665238291356 + 6.7062E-16
TEMP_ARR(5) = 0.192808104715315 + 7.6488E-16
TEMP_ARR(6) = 0.162377660896867 + 4.6182E-16
TEMP_ARR(7) = 0.14010418345305 + 2.416E-16
TEMP_ARR(8) = 0.123131963257932 + 2.9628E-16

V_VAL = Z_VAL
If (Z_VAL < 0) Then: V_VAL = -V_VAL
If Z_VAL = 0# Then
  MARSAG2_CUMUL_NORM_FUNC = 0.5
  Exit Function
End If

If Abs(Z_VAL) < 7.1 Then ' use Marsaglia
  j = CInt(V_VAL + 1)
  X_VAL = CDbl(j)
  Y_VAL = V_VAL - X_VAL
  A_VAL = TEMP_ARR(j)
  B_VAL = X_VAL * A_VAL - 1
  TEMP_PWR = 1
  TEMP_SUM = A_VAL + Y_VAL * B_VAL

  SQR_VAL = Y_VAL * Y_VAL
  For i = 2 To 64 Step 2
    W_VAL = TEMP_SUM
    A_VAL = (A_VAL + X_VAL * B_VAL) / i
    B_VAL = (B_VAL + X_VAL * A_VAL) / (i + 1)
    TEMP_PWR = TEMP_PWR * SQR_VAL
    TEMP_SUM = TEMP_SUM + TEMP_PWR * (A_VAL + Y_VAL * B_VAL)
    If TEMP_SUM = W_VAL Then: Exit For
  Next i

  TEMP_SUM = TEMP_SUM * Exp(-0.5 * Z_VAL * Z_VAL - 0.918938533204673)
  If (0 < Z_VAL) Then
    TEMP_SUM = 1 - TEMP_SUM
  End If
  
  MARSAG2_CUMUL_NORM_FUNC = TEMP_SUM
  Exit Function

ElseIf Abs(Z_VAL) < 37 Then    ' use asymptotics
  MARSAG2_CUMUL_NORM_FUNC = -Z_VAL / (1 + Z_VAL * _
        Z_VAL) * Exp(-0.5 * Z_VAL * Z_VAL - 0.918938533204673)
  
ElseIf (Abs(37) <= Z_VAL) Then ' avoid numerical Nirvana
  MARSAG2_CUMUL_NORM_FUNC = 0
End If

If (0 < Z_VAL) Then: MARSAG2_CUMUL_NORM_FUNC = MARSAG2_CUMUL_NORM_FUNC + 1

Exit Function
ERROR_LABEL:
MARSAG2_CUMUL_NORM_FUNC = Err.number
End Function
