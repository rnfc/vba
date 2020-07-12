Attribute VB_Name = "FINAN_DERIV_BINOM_LR_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : LR_EXTRAPOLATION_FUNC
'DESCRIPTION   : LEISEN_REIMER: A usual 2 pt Richardson extrapolation gives not
'much advantage ...but a 3 pt method does (for calls, does not hold for 'puts' as
'these methods are of order 1)

'LIBRARY       : DERIVATIVES
'GROUP         : BINOM_LR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function LR_EXTRAPOLATION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
ByVal nSTEPS As Long, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal VERSION As Integer = 0)

On Error GoTo ERROR_LABEL

Dim ii As Single
Dim jj As Single

On Error GoTo ERROR_LABEL

ii = Round(nSTEPS / 2, 0)
jj = Round(ii / 2, 0)

Select Case VERSION
    Case 0 '2 pt
        LR_EXTRAPOLATION_FUNC = 2 * LR_BINOMIAL_OPTION_FUNC(SPOT, STRIKE, _
                            TENOR, RATE, DIVD, SIGMA, 2 * 2 * jj + 1, OPTION_FLAG) - _
                            LR_BINOMIAL_OPTION_FUNC(SPOT, STRIKE, TENOR, RATE, DIVD, _
                            SIGMA, 1 * 2 * jj + 1, OPTION_FLAG)
    Case 1 'averaged
        LR_EXTRAPOLATION_FUNC = ((2 * LR_BINOMIAL_OPTION_FUNC(SPOT, STRIKE, _
                            TENOR, RATE, DIVD, SIGMA, 2 * 2 * jj + 1, OPTION_FLAG) - _
                            LR_BINOMIAL_OPTION_FUNC(SPOT, STRIKE, TENOR, RATE, DIVD, _
                            SIGMA, 1 * 2 * jj + 1, OPTION_FLAG)) + (LR_BINOMIAL_OPTION_FUNC(SPOT, _
                            STRIKE, TENOR, RATE, DIVD, SIGMA, nSTEPS, OPTION_FLAG))) / 2
    Case 2 '3 pt
        LR_EXTRAPOLATION_FUNC = 1 / 2 * LR_BINOMIAL_OPTION_FUNC(SPOT, STRIKE, TENOR, _
                            RATE, DIVD, SIGMA, 1 * 2 * jj + 1, OPTION_FLAG) _
                            - 4 * LR_BINOMIAL_OPTION_FUNC(SPOT, STRIKE, TENOR, _
                            RATE, DIVD, SIGMA, 2 * 2 * jj + 1, OPTION_FLAG) _
                            + 9 / 2 * LR_BINOMIAL_OPTION_FUNC(SPOT, STRIKE, TENOR, _
                            RATE, DIVD, SIGMA, 3 * 2 * jj + 1, OPTION_FLAG)
    Case Else '4 pt
        LR_EXTRAPOLATION_FUNC = -1 / 6 * LR_BINOMIAL_OPTION_FUNC(SPOT, STRIKE, TENOR, _
                            RATE, DIVD, SIGMA, 1 * 2 * jj + 1, OPTION_FLAG) _
                            + 4 * LR_BINOMIAL_OPTION_FUNC(SPOT, STRIKE, TENOR, _
                            RATE, DIVD, SIGMA, 2 * 2 * jj + 1, OPTION_FLAG) _
                            - 27 / 2 * LR_BINOMIAL_OPTION_FUNC(SPOT, STRIKE, TENOR, _
                            RATE, DIVD, SIGMA, 3 * 2 * jj + 1, OPTION_FLAG) _
                            + 32 / 3 * LR_BINOMIAL_OPTION_FUNC(SPOT, STRIKE, TENOR, _
                            RATE, DIVD, SIGMA, 4 * 2 * jj + 1, OPTION_FLAG)
End Select

Exit Function
ERROR_LABEL:
LR_EXTRAPOLATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : LR_BINOMIAL_OPTION_FUNC

'DESCRIPTION   : Binomial Option Value for American Options through
' Leisen Reimer tree with reduced computational costs

' Speed is about 880 prices per second for a 257 step tree
' and a 4-point Richardson extrapolation
' gives an exactness of 6 - 7 digits for the european case
' (starting with 65 steps), while extrapolation
' is not really helpfull in the case of early exercise.

'LIBRARY       : DERIVATIVES
'GROUP         : BINOM_LR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function LR_BINOMIAL_OPTION_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
ByVal nSTEPS As Long, _
Optional ByVal OPTION_FLAG As Integer = 1)

'OPTION_FLAG --> 1; CALL
'OPTION_FLAG --> -1; PUT

Dim i As Long '
Dim j As Long '
Dim NSIZE As Long '

Dim hh As Double
Dim ii As Double 'Terms
Dim jj As Double
Dim kk As Double 'Factors

Dim D1_VAL As Double
Dim D2_VAL As Double

Dim UP_VAL As Double
Dim DOWN_VAL As Double

Dim TEMP_VAL As Double
Dim DASH_VAL As Double
Dim STAR_VAL As Double

Dim DELTA_VAL As Double
Dim FACT_VAL As Double
Dim MULT_VAL As Double
Dim INV_FACT_VAL As Double

Dim X_VAL As Double
Dim PROB_VAL As Double

Dim TEMP_VECTOR As Variant ' working array for Leisen Reimer tree

On Error GoTo ERROR_LABEL

ReDim TEMP_VECTOR(0 To 2000)

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1 'Put

If (0 < nSTEPS And 0 < SPOT And 0 < STRIKE And 0 < TENOR And 0 < SIGMA) Then
Else
  LR_BINOMIAL_OPTION_FUNC = -2
  Exit Function
End If

NSIZE = 2 * CInt((nSTEPS - 1) / 2) + 1
X_VAL = STRIKE / SPOT

DELTA_VAL = TENOR / NSIZE
FACT_VAL = Exp(RATE * DELTA_VAL)
INV_FACT_VAL = 1 / FACT_VAL
MULT_VAL = Exp((RATE - DIVD) * DELTA_VAL)

' parameters for Leisen-Reimer
D2_VAL = ((Log(1 / X_VAL) + (RATE - DIVD) * _
TENOR) / (SIGMA * Sqr(TENOR)) + 0.5 * SIGMA * Sqr(TENOR)) - SIGMA * Sqr(TENOR)

D1_VAL = (Log(1 / X_VAL) + (RATE - DIVD) * _
TENOR) / (SIGMA * Sqr(TENOR)) + 0.5 * SIGMA * Sqr(TENOR)

PROB_VAL = 0.5 + Sgn(D2_VAL) * Sqr((0.25 * (1 - (Exp(-((D2_VAL / _
                    ((2 * CInt((NSIZE - 1) / 2) + 1) + 1 / 3 + 0.1 / _
                    ((2 * CInt((NSIZE - 1) / 2) + 1) + 1))) ^ 2) * _
                    ((2 * CInt((NSIZE - 1) / 2) + 1) + 1 / 6)))))) _
                    ' Preizer-Pratt Inversion (for odd NSIZE)

STAR_VAL = 1 - PROB_VAL
DASH_VAL = 0.5 + Sgn(D1_VAL) * Sqr((0.25 * (1 - (Exp(-((D1_VAL / _
                    ((2 * CInt((NSIZE - 1) / 2) + 1) + 1 / 3 + 0.1 / _
                    ((2 * CInt((NSIZE - 1) / 2) + 1) + 1))) ^ 2) * _
                    ((2 * CInt((NSIZE - 1) / 2) + 1) + 1 / 6)))))) _
                    ' Preizer-Pratt Inversion (for odd NSIZE)

UP_VAL = MULT_VAL * DASH_VAL / PROB_VAL
DOWN_VAL = (MULT_VAL - PROB_VAL * UP_VAL) / (1 - PROB_VAL)

jj = (DASH_VAL / (-1 + DASH_VAL) * (-1 + PROB_VAL) / PROB_VAL)
kk = (-1 + PROB_VAL) / (-1 + DASH_VAL) / MULT_VAL

hh = (MULT_VAL * (-1 + DASH_VAL) / (-1 + PROB_VAL)) ^ NSIZE
For i = 0 To NSIZE
  TEMP_VAL = OPTION_FLAG * (hh - X_VAL)
  If 0 <= TEMP_VAL Then ' pay off
        TEMP_VECTOR(i) = TEMP_VAL
  Else: TEMP_VECTOR(i) = 0#
  End If
  hh = jj * hh
Next i

ii = (1 / kk) ^ (NSIZE - 1)
For j = NSIZE - 1 To 0 Step -1
  hh = ii
  For i = 0 To j
    TEMP_VECTOR(i) = (PROB_VAL * TEMP_VECTOR(i + 1) + _
                      STAR_VAL * TEMP_VECTOR(i)) * INV_FACT_VAL
    TEMP_VAL = OPTION_FLAG * (hh - X_VAL)
    If TEMP_VECTOR(i) <= TEMP_VAL Then: TEMP_VECTOR(i) = TEMP_VAL
    hh = jj * hh
  Next i
  ii = kk * ii
Next j

LR_BINOMIAL_OPTION_FUNC = SPOT * TEMP_VECTOR(0)

Exit Function
ERROR_LABEL:
LR_BINOMIAL_OPTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LR_BINOMIAL_TREE_FUNC

'DESCRIPTION   : A binomial Leisen-Reimer tree avoids the oscillations of
'the usual binomial trees through a proper choice of the tree parameters. Besides that the
'geometry is the same as for CRR the speed is improved by a factor better than 40

'LIBRARY       : DERIVATIVES
'GROUP         : BINOM_LR
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function LR_BINOMIAL_TREE_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
ByVal nSTEPS As Long, _
Optional ByVal OPTION_FLAG As Integer = 1, _
Optional ByVal EXERCISE_TYPE As Integer = 1, _
Optional ByVal TREE_TYPE As Integer = 2)

'----------------------------------------------------------------------
'----------------------------------------------------------------------
' Returns Binomial Option Value
'----------------------------------------------------------------------
'OPTION_FLAG --> 1; CALL
'OPTION_FLAG --> -1; PUT
'EXERCISE_TYPE = 0 for European, else for American
'TREE_TYPE = 0 for CRR, else for Leisen Reimer;
'----------------------------------------------------------------------
'----------------------------------------------------------------------

Dim i As Double
Dim j As Double

Dim UP_VAL As Double
Dim DOWN_VAL As Double
Dim D1_VAL As Double
Dim D2_VAL As Double

Dim DASH_VAL As Double
Dim STAR_VAL As Double

Dim DELTA_VAL As Double
Dim FACT_VAL As Double
Dim MULT_VAL As Double

Dim PROB_VAL As Double

Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_VECTOR(0 To 2000)

If OPTION_FLAG <> 1 Then: OPTION_FLAG = -1 'Put

If TREE_TYPE <> 0 Then
    nSTEPS = IIf(nSTEPS Mod 2 <> 0 = False, nSTEPS + 1, nSTEPS)
'Same as --> Excel.Application.Odd(NSTEPS)
End If

ReDim TEMP_VECTOR(0 To nSTEPS)

If SPOT > 0 And STRIKE > 0 And TENOR > 0 And SIGMA > 0 Then
Else
  LR_BINOMIAL_TREE_FUNC = -1
  Exit Function
End If

DELTA_VAL = TENOR / nSTEPS
FACT_VAL = Exp(RATE * DELTA_VAL)
MULT_VAL = Exp((RATE - DIVD) * DELTA_VAL)

' Choice between TreeMode=0 (Cox,Ross&Rubinstein)
' and TreeMode=else (Leisen&Reimer)

If TREE_TYPE = 0 Then
  UP_VAL = Exp(SIGMA * Sqr(DELTA_VAL))
  DOWN_VAL = 1 / UP_VAL
  PROB_VAL = (MULT_VAL - DOWN_VAL) / (UP_VAL - DOWN_VAL)
  STAR_VAL = 1 - PROB_VAL
Else
  D2_VAL = ((Log(SPOT / STRIKE) + (RATE - DIVD) * _
            TENOR) / (SIGMA * Sqr(TENOR)) + 0.5 * _
            SIGMA * Sqr(TENOR)) - SIGMA * Sqr(TENOR)
  
  D1_VAL = (Log(SPOT / STRIKE) + (RATE - DIVD) * _
            TENOR) / (SIGMA * Sqr(TENOR)) + 0.5 * SIGMA * Sqr(TENOR)

  PROB_VAL = 0.5 + Sgn(D2_VAL) * Sqr((0.25 * (1 - (Exp(-((D2_VAL / _
                    ((2 * CInt((nSTEPS - 1) / 2) + 1) + 1 / 3 + 0.1 / _
                    ((2 * CInt((nSTEPS - 1) / 2) + 1) + 1))) ^ 2) * _
                    ((2 * CInt((nSTEPS - 1) / 2) + 1) + 1 / 6)))))) _
                    ' Preizer-Pratt Inversion (for odd NSTEPS)
  
  STAR_VAL = 1 - PROB_VAL
  DASH_VAL = 0.5 + Sgn(D1_VAL) * Sqr((0.25 * (1 - (Exp(-((D1_VAL / _
                    ((2 * CInt((nSTEPS - 1) / 2) + 1) + 1 / 3 + 0.1 / _
                    ((2 * CInt((nSTEPS - 1) / 2) + 1) + 1))) ^ 2) * _
                    ((2 * CInt((nSTEPS - 1) / 2) + 1) + 1 / 6)))))) _
                    ' Preizer-Pratt Inversion (for odd NSTEPS)

  UP_VAL = MULT_VAL * DASH_VAL / PROB_VAL
  DOWN_VAL = (MULT_VAL - PROB_VAL * UP_VAL) / (1 - PROB_VAL)
End If

For i = 0 To nSTEPS
  TEMP_VECTOR(i) = MAXIMUM_FUNC(OPTION_FLAG * (SPOT * (UP_VAL ^ i) _
                    * (DOWN_VAL ^ (nSTEPS - i)) - STRIKE), 0)
Next i

For j = nSTEPS - 1 To 0 Step -1
  For i = 0 To j
      TEMP_VECTOR(i) = (PROB_VAL * TEMP_VECTOR(i + 1) + _
                    STAR_VAL * TEMP_VECTOR(i)) / FACT_VAL
      If EXERCISE_TYPE <> 0 Then
        TEMP_VECTOR(i) = MAXIMUM_FUNC(TEMP_VECTOR(i), OPTION_FLAG * _
                    (SPOT * (UP_VAL ^ i) * (DOWN_VAL ^ _
                    (j - i)) - STRIKE))
      End If
  Next i
Next j

LR_BINOMIAL_TREE_FUNC = TEMP_VECTOR(0)

Exit Function
ERROR_LABEL:
LR_BINOMIAL_TREE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : LR_DELTA_FUNC
'DESCRIPTION   : Leisen-Reimer Delta by central difference
'LIBRARY       : DERIVATIVES
'GROUP         : BINOM_LR
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/29/2009
'************************************************************************************
'************************************************************************************

Function LR_DELTA_FUNC(ByVal SPOT As Double, _
ByVal STRIKE As Double, _
ByVal TENOR As Double, _
ByVal RATE As Double, _
ByVal DIVD As Double, _
ByVal SIGMA As Double, _
ByVal nSTEPS As Long, _
Optional ByVal OPTION_FLAG As Integer = 1)

Dim FACTOR_VAL As Double

On Error GoTo ERROR_LABEL

FACTOR_VAL = 10000
        
LR_DELTA_FUNC = (LR_BINOMIAL_OPTION_FUNC(SPOT + SPOT / FACTOR_VAL, _
            STRIKE, TENOR, RATE, DIVD, SIGMA, nSTEPS, OPTION_FLAG) _
            - LR_BINOMIAL_OPTION_FUNC(SPOT - SPOT / FACTOR_VAL, STRIKE, TENOR, _
            RATE, DIVD, SIGMA, nSTEPS, OPTION_FLAG)) / (2 * SPOT / FACTOR_VAL)

Exit Function
ERROR_LABEL:
LR_DELTA_FUNC = Err.number
End Function
