Attribute VB_Name = "NUMBER_REAL_TRIGONOMETRIC_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : SIN_FUNC
'DESCRIPTION   : Returns the sine of the given angle
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function SIN_FUNC(ByVal X_VAL As Double)
'NEED FINAL ADJUSTMENT (error <= 0.000000001)

Dim i As Long
Dim j As Long

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double
Dim FTEMP_VAL As Double
Dim GTEMP_VAL As Double

Dim PI_VAL As Double
Dim QUARTER_PI As Double
Dim epsilon As Double
Dim TEMP_ARR As Variant

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979
epsilon = 10 ^ -30

If X_VAL = 0 Then SIN_FUNC = 0: Exit Function
ATEMP_VAL = 0: DTEMP_VAL = 0

TEMP_ARR = REDUCE_DEGREES_FUNC(X_VAL)
i = LBound(TEMP_ARR)
ATEMP_VAL = TEMP_ARR(i)
DTEMP_VAL = TEMP_ARR(i + 1)

QUARTER_PI = PI_VAL / 4
If QUARTER_PI >= ATEMP_VAL Then
    ETEMP_VAL = CDec(ATEMP_VAL)
    FTEMP_VAL = ETEMP_VAL
    GTEMP_VAL = ETEMP_VAL * ETEMP_VAL
    For i = 3 To 1000 Step 2
        j = i * (i - 1)
        ETEMP_VAL = ETEMP_VAL * GTEMP_VAL
        ETEMP_VAL = -ETEMP_VAL / j
        FTEMP_VAL = FTEMP_VAL + ETEMP_VAL
        If Abs(ETEMP_VAL) <= epsilon Then Exit For
    Next i
    CTEMP_VAL = FTEMP_VAL
Else
    BTEMP_VAL = CDec(QUARTER_PI * 2 - ATEMP_VAL)
    GTEMP_VAL = BTEMP_VAL
    GTEMP_VAL = GTEMP_VAL * GTEMP_VAL
    ETEMP_VAL = 1
    FTEMP_VAL = ETEMP_VAL
    For i = 2 To 1000 Step 2
        j = i * (i - 1)
        ETEMP_VAL = ETEMP_VAL * GTEMP_VAL
        ETEMP_VAL = -ETEMP_VAL / j
        FTEMP_VAL = FTEMP_VAL + ETEMP_VAL
        If Abs(ETEMP_VAL) <= epsilon Then Exit For
    Next i
    CTEMP_VAL = FTEMP_VAL
End If
If DTEMP_VAL = 1 Or DTEMP_VAL = 2 Then
    'nothing to do
Else: CTEMP_VAL = -CTEMP_VAL
End If

If Abs(CTEMP_VAL) < epsilon Then CTEMP_VAL = 0
SIN_FUNC = CTEMP_VAL

Exit Function
ERROR_LABEL:
SIN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COS_FUNC
'DESCRIPTION   : Returns the cosine of the given angle
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COS_FUNC(ByVal X_VAL As Double)
'NEED FINAL ADJUSTMENT (error <= 0.000000001)

Dim i As Long
Dim j As Long

Dim PI_VAL As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double
Dim FTEMP_VAL As Double
Dim GTEMP_VAL As Double

Dim TEMP_ARR As Variant
Dim QUARTER_PI As Double

Dim epsilon As Double

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979
epsilon = 10 ^ -30

If X_VAL = 0 Then COS_FUNC = 1: Exit Function
ATEMP_VAL = 0: DTEMP_VAL = 0

TEMP_ARR = REDUCE_DEGREES_FUNC(X_VAL)
i = LBound(TEMP_ARR)
ATEMP_VAL = TEMP_ARR(i)
DTEMP_VAL = TEMP_ARR(i + 1)

QUARTER_PI = PI_VAL / 4

If QUARTER_PI >= ATEMP_VAL Then
    GTEMP_VAL = CDec(ATEMP_VAL)
    GTEMP_VAL = GTEMP_VAL * GTEMP_VAL
    ETEMP_VAL = 1
    FTEMP_VAL = ETEMP_VAL
    For i = 2 To 1000 Step 2
        j = i * (i - 1)
        ETEMP_VAL = ETEMP_VAL * GTEMP_VAL
        ETEMP_VAL = -ETEMP_VAL / j
        FTEMP_VAL = FTEMP_VAL + ETEMP_VAL
        If Abs(ETEMP_VAL) <= epsilon Then Exit For
    Next i
    CTEMP_VAL = FTEMP_VAL
Else
    BTEMP_VAL = (QUARTER_PI * 2 - ATEMP_VAL)
    ETEMP_VAL = CDec(BTEMP_VAL)
    FTEMP_VAL = ETEMP_VAL
    GTEMP_VAL = ETEMP_VAL * ETEMP_VAL
    For i = 3 To 1000 Step 2
        j = i * (i - 1)
        ETEMP_VAL = ETEMP_VAL * GTEMP_VAL
        ETEMP_VAL = -ETEMP_VAL / j
        FTEMP_VAL = FTEMP_VAL + ETEMP_VAL
        If Abs(ETEMP_VAL) <= epsilon Then Exit For
    Next i
    CTEMP_VAL = FTEMP_VAL
End If

If DTEMP_VAL = 1 Or DTEMP_VAL = 4 Then
Else: CTEMP_VAL = -CTEMP_VAL
End If

If Abs(CTEMP_VAL) < epsilon Then CTEMP_VAL = 0
COS_FUNC = CTEMP_VAL

Exit Function
ERROR_LABEL:
COS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : TAN_FUNC
'DESCRIPTION   : Returns the tangent of the given angle
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function TAN_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
TAN_FUNC = SIN_FUNC(X_VAL) / COS_FUNC(X_VAL)
Exit Function
ERROR_LABEL:
TAN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : SEC_FUNC
'DESCRIPTION   : Returns the secant of the given angle
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function SEC_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL

SEC_FUNC = 1 / COS_FUNC(X_VAL)   'Secant:
Exit Function
ERROR_LABEL:
SEC_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COSEC_FUNC
'DESCRIPTION   : Returns the cosecant of the given angle
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COSEC_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
COSEC_FUNC = 1 / SIN_FUNC(X_VAL)   'Cosecant:
Exit Function
ERROR_LABEL:
COSEC_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COTAN_FUNC
'DESCRIPTION   : Returns the cotangent of the given angle
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COTAN_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
COTAN_FUNC = 1 / TAN_FUNC(X_VAL)   'Cotangent:
Exit Function
ERROR_LABEL:
COTAN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ACOS_FUNC
'DESCRIPTION   : Returns the arccosine, or inverse cosine, of a number. The
'arccosine is the angle whose cosine is number. The returned
'angle is given in radians in the range 0 (zero) to pi.
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ACOS_FUNC(ByVal X_VAL As Double)

Dim PI_VAL As Double

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979

If X_VAL = 1 Then
    ACOS_FUNC = 0
ElseIf X_VAL = -1 Then
    ACOS_FUNC = PI_VAL
Else
    ACOS_FUNC = Atn(-X_VAL / Sqr(-X_VAL * X_VAL + 1)) + 2 * Atn(1)
End If

Exit Function
ERROR_LABEL:
ACOS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASIN_FUNC
'DESCRIPTION   : Returns the arcsine, or inverse sine, of a number.
'The arcsine is the angle whose sine is number. The returned angle
'is given in radians in the range -pi/2 to pi/2.
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ASIN_FUNC(ByVal X_VAL As Double)
  
Dim PI_VAL As Double
  
On Error GoTo ERROR_LABEL
  
PI_VAL = 3.14159265358979
  
If Abs(X_VAL) = 1 Then
    ASIN_FUNC = Sgn(X_VAL) * PI_VAL / 2
Else
    ASIN_FUNC = Atn(X_VAL / Sqr(-X_VAL * X_VAL + 1))
End If

Exit Function
ERROR_LABEL:
ASIN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ATAN_FUNC
'DESCRIPTION   : Computes the arctangent to about 13.7 decimal digits
'of accuracy using a simple rational polynomial. The problem with this
'routine is, arctan is a very non-linear function, and resistent to quick
'polynomial evaluation with reasonable accuracy.
'--------------------------------------------------------------------------------------
'The max absolute error in [0,1] is about 3.2e-7 which is still
'slightly better than 32bit floating-point representation of
'arctan(X_VAL) in [0,1].
'--------------------------------------------------------------------------------------
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ATAN_FUNC(ByVal X_VAL As Double)

'---------------------------------------------------------------------------------
'Derivation of the Polynomial:
'Tanh(X_VAL) = (Exp(X_VAL) - Exp(-X_VAL)) /
'(Exp(X_VAL) + Exp(-X_VAL))
'        = 2/(exp(-2*X_VAL) + 1) - 1
'both arctan() and tanh() are odd symmetry functions that have asymptotic
'behavior at +/- inf.
'that pre-polynomial had to approximate:

'p(X_VAL) = Log((1 - 2 / pi * arctan(X_VAL)) / (1 + 2 / pi *
'arctan(X_VAL)))
'so arctan(X_VAL) = pi / 2 * Tanh(-p(X_VAL) / 2)
'or arctan(X_VAL) = pi / (Exp(p(X_VAL)) + 1) - pi / 2
'---------------------------------------------------------------------------------

Dim C1_VAL As Double
Dim C2_VAL As Double
Dim C3_VAL As Double
Dim C4_VAL As Double
Dim C5_VAL As Double
Dim C6_VAL As Double

Dim PI_VAL As Double
Dim Y_VAL As Double

Dim COMP_FLAG As Boolean
Dim REGION_FLAG As Boolean
Dim SIGN_FLAG As Boolean

On Error GoTo ERROR_LABEL

COMP_FLAG = False             ' true if arg was >1
REGION_FLAG = False                 ' true depending on REGION_FLAG arg is in
SIGN_FLAG = False                   ' true if arg was < 0

PI_VAL = 3.14159265358979

C1_VAL = 48.701070044049
C2_VAL = 49.5326263772254
C3_VAL = 9.40604244231624
C4_VAL = 48.70107004405
C5_VAL = 65.7663163908956
C6_VAL = 21.5879340670203

If (X_VAL < 0) Then
    X_VAL = -X_VAL
    SIGN_FLAG = True                    ' arctan(-X_VAL)=-arctan(X_VAL)
End If

If (X_VAL > 1#) Then
    X_VAL = 1# / X_VAL                     ' keep arg between 0 and 1
    COMP_FLAG = True
End If

If (X_VAL > (TAN_FUNC(PI_VAL / 12#))) Then
    X_VAL = (X_VAL - TAN_FUNC(PI_VAL / 6#)) / (1 + TAN_FUNC(PI_VAL / 6#) * _
    X_VAL) ' reduce arg to under TAN_FUNC(pi/12)
    REGION_FLAG = True
End If

Y_VAL = (X_VAL * (C1_VAL + (X_VAL * X_VAL) * _
    (C2_VAL + (X_VAL * X_VAL) * C3_VAL)) / (C4_VAL + (X_VAL * X_VAL) * _
    (C5_VAL + (X_VAL * X_VAL) * (C6_VAL + (X_VAL * X_VAL)))))



' run the approximation
If (REGION_FLAG) Then: Y_VAL = Y_VAL + (PI_VAL / 6#)
' correct for REGION_FLAG we're in
If (COMP_FLAG) Then: Y_VAL = (PI_VAL / 2#) - Y_VAL
' correct for 1/X_VAL if we did that
If (SIGN_FLAG) Then: Y_VAL = -Y_VAL
' correct for negative arg

ATAN_FUNC = Y_VAL

Exit Function
ERROR_LABEL:
ATAN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ATAN2_FUNC
'DESCRIPTION   : Returns the arctangent, or inverse tangent, of the
'specified x- and y-coordinates. The arctangent is the angle from
'the x-axis to a line containing the origin (0, 0) and a point with
'coordinates (x_num, y_num). The angle is given in radians between
'-pi and pi, excluding -pi.
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ATAN2_FUNC(ByVal YX_VAL As Double, _
ByVal XX_VAL As Double)

'The underlying idea is that you are going to rotate an
'(XX_VAL,YX_VAL) vector to align
'with the positive x-axis through a series of known
'shifts that are computationally efficient to apply.
'The final result will be the negative of
'the sum of shifts applied to get as close to the x-axis
'as you wanted. Generally, each additional phase bit output
'costs a sign check, 2 shifts and 2 accumulates (conditionally
'add or subtract), plus any addressing overhead.

'There is a known scaling required (multiply) that can be at the end,
'beginning, or could be distributed.

'There is a discussion of it in an appendix in Frerking. Besides
'a typo, he needs to extend his index range to one lower to be
'able to compute ATAN2_FUNC on any angle in the [0,90] degree range
'he claims. You need to rotate the vector into the [0,90] degree range
'using XX_VAL,YX_VAL exchanges and sign adjustments, accumulating the
'corresponding phase rotations in the process,
'before you launch into his algorithm.

Dim PI_VAL As Double
Dim HALF_PI_VAL As Double

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979

HALF_PI_VAL = PI_VAL / 2

If (Abs(YX_VAL) >= Abs(XX_VAL)) Then
  ATAN2_FUNC = Sgn(YX_VAL) * HALF_PI_VAL - Atn(XX_VAL / YX_VAL)
Else
  If (XX_VAL > 0) Then
    ATAN2_FUNC = Atn(YX_VAL / XX_VAL)
  Else
    If (YX_VAL >= 0) Then
      ATAN2_FUNC = PI_VAL + Atn(YX_VAL / XX_VAL)
    Else
      ATAN2_FUNC = -PI_VAL + Atn(YX_VAL / XX_VAL)
    End If
  End If
End If

Exit Function
ERROR_LABEL:
ATAN2_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ATN_FUNC
'DESCRIPTION   : Returns the arctangent of a number
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ATN_FUNC(ByVal X_VAL As Double) 'As Double
  
Dim PI_VAL As Double

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979

If Abs(X_VAL) = 1 Then
    ATN_FUNC = Sgn(X_VAL) * PI_VAL / 2
Else
    ATN_FUNC = Atn(X_VAL / Sqr(1 - X_VAL ^ 2))
End If

Exit Function
ERROR_LABEL:
ATN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASEC_FUNC
'DESCRIPTION   : Returns the arcsecant, or inverse secant, of a number
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008

'************************************************************************************
'************************************************************************************

Function ASEC_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
ASEC_FUNC = 2 * Atn(1) - Atn(Sgn(X_VAL) / Sqr(X_VAL * X_VAL - 1))
'Inverse Secant
Exit Function
ERROR_LABEL:
ASEC_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ACOSEC_FUNC
'DESCRIPTION   : Returns the arccosecant, or inverse cosecant, of a number
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ACOSEC_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
ACOSEC_FUNC = Atn(Sgn(X_VAL) / Sqr(X_VAL * X_VAL - 1)) 'Inverse Cosecant:
Exit Function
ERROR_LABEL:
ACOSEC_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ACOTAN_FUNC
'DESCRIPTION   : Returns the arccotangent, or inverse cotangent, of a number
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ACOTAN_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
ACOTAN_FUNC = 2 * Atn(1) - Atn(X_VAL)   'Inverse Cotangent:
Exit Function
ERROR_LABEL:
ACOTAN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : COSH_FUNC
'DESCRIPTION   : Returns the hyperbolic cosine of a number
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COSH_FUNC(ByVal X_VAL As Double)
  
On Error GoTo ERROR_LABEL
  
COSH_FUNC = (Exp(X_VAL) + Exp(-X_VAL)) / 2

Exit Function
ERROR_LABEL:
COSH_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SINSH_FUNC
'DESCRIPTION   : Returns the hyperbolic sine of a number
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function SINSH_FUNC(ByVal X_VAL As Double)
  
On Error GoTo ERROR_LABEL
  
SINSH_FUNC = (Exp(X_VAL) - Exp(-X_VAL)) / 2

Exit Function
ERROR_LABEL:
SINSH_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : TANH_FUNC
'DESCRIPTION   : Returns the hyperbolic tangent of a number
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function TANH_FUNC(ByVal X_VAL As Double)
  
On Error GoTo ERROR_LABEL
  
TANH_FUNC = (Exp(X_VAL) - Exp(-X_VAL)) / (Exp(X_VAL) + Exp(-X_VAL))

Exit Function
ERROR_LABEL:
TANH_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SECH_FUNC
'DESCRIPTION   : Returns the hyperbolic Secant of a number
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function SECH_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
SECH_FUNC = 2 / (Exp(X_VAL) + Exp(-X_VAL))
'Hyperbolic Secant:
Exit Function
ERROR_LABEL:
SECH_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COSECH_FUNC
'DESCRIPTION   : Returns the hyperbolic Cosecant of a number
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 018
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function COSECH_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
COSECH_FUNC = 2 / (Exp(X_VAL) - Exp(-X_VAL))
'Hyperbolic Cosecant:
Exit Function
ERROR_LABEL:
COSECH_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : COTANH_FUNC
'DESCRIPTION   : Returns the hyperbolic Cotangent of a number
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 019
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function COTANH_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
COTANH_FUNC = (Exp(X_VAL) + Exp(-X_VAL)) / (Exp(X_VAL) - Exp(-X_VAL))
'Hyperbolic Cotangent:
Exit Function
ERROR_LABEL:
COTANH_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ACOSH_FUNC
'DESCRIPTION   : Returns the inverse hyperbolic cosine of a number.
'Number must be greater than or equal to 1. The inverse hyperbolic cosine
'is the value whose hyperbolic cosine is number, so
'ACOSH_FUNC(COSH_FUNC(number)) equals number.
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 020
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ACOSH_FUNC(ByVal X_VAL As Double)
  
On Error GoTo ERROR_LABEL
  
ACOSH_FUNC = Log(X_VAL + Sqr(X_VAL * X_VAL - 1))

Exit Function
ERROR_LABEL:
ACOSH_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASINH_FUNC
'DESCRIPTION   : Returns the inverse hyperbolic sine of a number. The
'inverse hyperbolic sine is the value whose hyperbolic sine is number,
'so ASINH_FUNC(SINH_FUNC(number)) equals number.
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 021
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function ASINH_FUNC(ByVal X_VAL As Double)
  
On Error GoTo ERROR_LABEL
  
ASINH_FUNC = Log(X_VAL + Sqr(X_VAL * X_VAL + 1))

Exit Function
ERROR_LABEL:
ASINH_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ATANH_FUNC
'DESCRIPTION   : Returns the inverse hyperbolic tangent of a number.
'Number must be between -1 and 1 (excluding -1 and 1). The inverse
'hyperbolic tangent is the value whose hyperbolic tangent is number, so
'ATANH_FUNC(TANH_FUNC(number)) equals number.
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 022
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ATANH_FUNC(ByVal X_VAL As Double)

On Error GoTo ERROR_LABEL

ATANH_FUNC = Log((1 + X_VAL) / (1 - X_VAL)) / 2

Exit Function
ERROR_LABEL:
ATANH_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASECH_FUNC
'DESCRIPTION   : Returns the inverse hyperbolic secant of a number
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 023
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ASECH_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
ASECH_FUNC = Log((Sqr(-X_VAL * X_VAL + 1) + 1) / X_VAL)
'Inverse Hyperbolic Secant:
Exit Function
ERROR_LABEL:
ASECH_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ACOSECH_FUNC
'DESCRIPTION   : Returns the inverse hyperbolic cosecant of a number
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 024
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ACOSECH_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
ACOSECH_FUNC = Log((Sgn(X_VAL) * Sqr(X_VAL * _
X_VAL + 1) + 1) / X_VAL)  'Inverse Hyperbolic Cosecant:
Exit Function
ERROR_LABEL:
ACOSECH_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ACOTANH_FUNC
'DESCRIPTION   : Returns the inverse hyperbolic cotangent of a number
'LIBRARY       : NUMBER_REAL
'GROUP         : TRIGONOMETRIC
'ID            : 025
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function ACOTANH_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
ACOTANH_FUNC = Log((X_VAL + 1) / (X_VAL - 1)) / 2
'Inverse Hyperbolic Cotangent:
Exit Function
ERROR_LABEL:
ACOTANH_FUNC = Err.number
End Function
