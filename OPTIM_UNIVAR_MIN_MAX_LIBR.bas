Attribute VB_Name = "OPTIM_UNIVAR_MIN_MAX_LIBR"

'-------------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
'-------------------------------------------------------------------------------------

'************************************************************************************
'************************************************************************************
'FUNCTION      : UNIVAR_MIN_BRENT_FUNC

'DESCRIPTION   : The algorithm finds a local minimum of F(X_MIN_VALUE) function in
'the interval [LOWER,UPPER].

'This simple optimization function use the golden section method. It has
'guaranteed linear convergence independent of the function whose minimum
'is to be found. An interval is divided into 2 parts, and on each step
'the algorithm chooses the left or right interval without taking into
'account the behavior of the function. Independently of the value of function
'obtained, if we select left interval, we'll calculate the following function
'value in a point A, otherwise, the next point will be B.

'Besides, we can take into account that in the neighborhood of the minimum
'function it could be described by a parabola. The most obvious solution is
'to build a parabola by three known values, and find its minimum. Then we
'select three more points near the extremum found, and build a new parabola,
'having the next interpolation.

'The convergence of such method is quadratic, but this method isn't applicable.
'First, it is only the smooth function that can be described by a parabola in
'the neighborhood of minimum, but the method can fail when working with
'non-smooth functions. Second, we must have a rather small initial interval
'containing the minimum. Outside of such an interval the function behavior is
'arbitrary, and we could converge to the maximum instead of the minimum, or
'get outside of the search domain while finding a minimum of the recurrent parabola.

'The Brent method combines some lines of a golden section method and some of a
'parabolic interpolation method, which is proposed above. This method is characterized
'by quadratic convergence in case of smooth functions and guaranteed linear
'convergence in case of nonsmooth or sophisticated functions.

'Few words about the principles of work. The algorithm keeps in memory five
'points: points a and b limit interval [LOWER, UPPER], where the minimum is localised,
'and three points with the best-case value of f(X_MIN_VALUE) are kept in points
'X_MIN_VALUE, WTEMP_VAL, VTEMP_VAL.

'On the basis of these points, the recurrent parabola is formulated. The next
'value of the function is calculated in the point of the minimum of this parabola.
'If the assumed minimum is outside of the interval [a, b] or if some conditions
'are fulfilled, the next step is performed on the basis of the golden section
'method. These conditions are: the presence of signs of nonsmoothness or a rather
'big distance to extremum.

'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_MIN_MAX
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function UNIVAR_MIN_BRENT_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef CONST_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal epsilon As Double = 0.000000000000001)

Dim i As Long
    
Dim INIT_VAL As Double
Dim FIRST_VAL As Double
Dim SECOND_VAL As Double
Dim THIRD_VAL As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double

Dim ETEMP_VAL As Double
Dim FTEMP_VAL As Double
Dim GTEMP_VAL As Double

Dim PTEMP_VAL As Double
Dim QTEMP_VAL As Double
Dim RTEMP_VAL As Double
Dim UTEMP_VAL As Double
Dim VTEMP_VAL As Double
Dim WTEMP_VAL As Double

Dim LOWER_BOUND As Double
Dim UPPER_BOUND As Double

Dim TEMP_DELTA As Double
Dim GOLD_CONST As Double
    
Dim CONST_DATA As Variant

Dim X_MIN_VALUE As Double
Dim FUNC_MIN_VAL As Double
Dim MIN_MAX_MULT As Integer

On Error GoTo ERROR_LABEL

If MIN_FLAG = True Then 'Minimization
    MIN_MAX_MULT = 1
Else
    MIN_MAX_MULT = -1
End If

CONST_DATA = CONST_RNG

If UBound(CONST_DATA, 2) > 1 Then
    LOWER_BOUND = CONST_DATA(1, 1)       'Xmin
    UPPER_BOUND = CONST_DATA(1, 2)       'Xmax
Else
    LOWER_BOUND = CONST_DATA(1, 1)       'Xmin
    UPPER_BOUND = CONST_DATA(2, 1)       'Xmax
End If

GOLD_CONST = 0.381966
INIT_VAL = 0.5 * (LOWER_BOUND + UPPER_BOUND)
If LOWER_BOUND < UPPER_BOUND Then
      FIRST_VAL = LOWER_BOUND
Else: FIRST_VAL = UPPER_BOUND
End If

If LOWER_BOUND > UPPER_BOUND Then
      SECOND_VAL = LOWER_BOUND
Else: SECOND_VAL = UPPER_BOUND
End If

VTEMP_VAL = INIT_VAL
WTEMP_VAL = VTEMP_VAL
X_MIN_VALUE = VTEMP_VAL
FTEMP_VAL = 0#
BTEMP_VAL = Excel.Application.Run(FUNC_NAME_STR, X_MIN_VALUE) * MIN_MAX_MULT
THIRD_VAL = X_MIN_VALUE
CTEMP_VAL = BTEMP_VAL
DTEMP_VAL = BTEMP_VAL

COUNTER = 0
For i = 1 To nLOOPS Step 1
    TEMP_DELTA = 0.5 * (FIRST_VAL + SECOND_VAL)
    If Abs(X_MIN_VALUE - TEMP_DELTA) <= epsilon * 2# - 0.5 * _
            (SECOND_VAL - FIRST_VAL) Then
        Exit For
    End If
    If Abs(FTEMP_VAL) > epsilon Then
        RTEMP_VAL = (X_MIN_VALUE - WTEMP_VAL) * (BTEMP_VAL - CTEMP_VAL)
        QTEMP_VAL = (X_MIN_VALUE - VTEMP_VAL) * (BTEMP_VAL - DTEMP_VAL)
        PTEMP_VAL = (X_MIN_VALUE - VTEMP_VAL) * QTEMP_VAL - _
                (X_MIN_VALUE - WTEMP_VAL) * RTEMP_VAL
        QTEMP_VAL = 2# * (QTEMP_VAL - RTEMP_VAL)
        If QTEMP_VAL > 0# Then
            PTEMP_VAL = -PTEMP_VAL
        End If
        QTEMP_VAL = Abs(QTEMP_VAL)
        ETEMP_VAL = FTEMP_VAL
        FTEMP_VAL = GTEMP_VAL
        If Not (Abs(PTEMP_VAL) >= Abs(0.5 * QTEMP_VAL * _
                ETEMP_VAL) Or PTEMP_VAL <= QTEMP_VAL * _
                (FIRST_VAL - X_MIN_VALUE) Or PTEMP_VAL >= QTEMP_VAL * _
                        (SECOND_VAL - X_MIN_VALUE)) Then
            GTEMP_VAL = PTEMP_VAL / QTEMP_VAL
            UTEMP_VAL = X_MIN_VALUE + GTEMP_VAL
            If UTEMP_VAL - FIRST_VAL < epsilon * 2# Or SECOND_VAL - _
                    UTEMP_VAL < epsilon * 2# Then
                GTEMP_VAL = IIf((TEMP_DELTA - X_MIN_VALUE) > 0#, _
                        Abs(epsilon), -Abs(epsilon))
            End If
        Else
            If X_MIN_VALUE >= TEMP_DELTA Then
                FTEMP_VAL = FIRST_VAL - X_MIN_VALUE
            Else
                FTEMP_VAL = SECOND_VAL - X_MIN_VALUE
            End If
            GTEMP_VAL = GOLD_CONST * FTEMP_VAL
        End If
    Else
        If X_MIN_VALUE >= TEMP_DELTA Then
            FTEMP_VAL = FIRST_VAL - X_MIN_VALUE
        Else
            FTEMP_VAL = SECOND_VAL - X_MIN_VALUE
        End If
        GTEMP_VAL = GOLD_CONST * FTEMP_VAL
    End If
    If Abs(GTEMP_VAL) >= epsilon Then
          UTEMP_VAL = X_MIN_VALUE + GTEMP_VAL
    Else: UTEMP_VAL = X_MIN_VALUE + IIf(GTEMP_VAL > 0#, _
                Abs(epsilon), -Abs(epsilon))
    End If
    ATEMP_VAL = Excel.Application.Run(FUNC_NAME_STR, UTEMP_VAL) * MIN_MAX_MULT
    If ATEMP_VAL <= BTEMP_VAL Then
        If UTEMP_VAL >= X_MIN_VALUE Then
              FIRST_VAL = X_MIN_VALUE
        Else: SECOND_VAL = X_MIN_VALUE
        End If
        VTEMP_VAL = WTEMP_VAL
        CTEMP_VAL = DTEMP_VAL
        WTEMP_VAL = X_MIN_VALUE
        DTEMP_VAL = BTEMP_VAL
        X_MIN_VALUE = UTEMP_VAL
        BTEMP_VAL = ATEMP_VAL
    Else
        If UTEMP_VAL < X_MIN_VALUE Then
            FIRST_VAL = UTEMP_VAL
        Else
            SECOND_VAL = UTEMP_VAL
        End If
        If ATEMP_VAL <= DTEMP_VAL Or WTEMP_VAL = X_MIN_VALUE Then
            VTEMP_VAL = WTEMP_VAL
            CTEMP_VAL = DTEMP_VAL
            WTEMP_VAL = UTEMP_VAL
            DTEMP_VAL = ATEMP_VAL
        Else
            If ATEMP_VAL <= CTEMP_VAL Or VTEMP_VAL = X_MIN_VALUE Or _
                VTEMP_VAL = 2# Then
                VTEMP_VAL = UTEMP_VAL
                CTEMP_VAL = ATEMP_VAL
            End If
        End If
    End If
    FUNC_MIN_VAL = BTEMP_VAL '--> Function min value
    COUNTER = COUNTER + 1
Next i

UNIVAR_MIN_BRENT_FUNC = X_MIN_VALUE

Exit Function
ERROR_LABEL:
UNIVAR_MIN_BRENT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : UNIVAR_MIN_GOLD_FUNC

'DESCRIPTION   : Returns an estimate for the minimum location with accuracy
'  3SQRT_EPSILONabs(x) + tol.

'  The function always obtains a local minimum which coincides with
'  the global one only if a function under investigation being
'  unimodular.
'  If a function being examined possesses no local minimum within
'  the given range, Fminbr returns 'a' (if f(a) < f(b)), otherwise
'  it returns the right range boundary value b.
'
'  Algorithm
'  G.Forsythe, M.Malcolm, C.Moler, Computer methods for mathematical
'  computations. M., Mir, 1980, p.202 of the Russian edition
'
'  The function makes use of the "gold section" procedure combined with
'  the parabolic interpolation.
'  At every step program operates three abscissae - x,v, and w.
'  x - the last and the best approximation to the minimum location,
'      i.e. f(x) <= f(a) or/and f(x) <= f(b)
'      (if the function f has a local minimum in (a,b), then the both
'      conditions are fulfiled after one or two steps).
'  v,w are previous approximations to the minimum location. They may
'  coincide with a, b, or x (although the algorithm tries to make all
'  u, v, and w distinct). Points x, v, and w are used to construct
'  interpolating parabola whose minimum will be treated as a new
'  approximation to the minimum location if the former falls within
'  [a,b] and reduces the range enveloping minimum more efficient than
'  the gold section procedure.
'  When f(x) has a second derivative positive at the minimum location
'  (not coinciding with a or b) the procedure converges superlinearly
'  at a rate order about 1.324

'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_MIN_MAX
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function UNIVAR_MIN_GOLD_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef CONST_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 200, _
Optional ByVal epsilon As Double = 0.000000000000001)

Dim i As Long

Dim MID_RNG As Double
Dim NEW_STEP As Double
Dim TEMP_RNG As Double

Dim LOWER_BOUND As Double
Dim UPPER_BOUND As Double

Dim TEMP_DELTA As Double

Dim FIRST_POINT As Double
Dim SECOND_POINT As Double

Dim TEMP_POINT As Double
Dim TEMP_FUNC As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double

Dim DELTA_FUNC As Double
Dim FIRST_FUNC As Double
Dim SECOND_FUNC As Double

Dim CONST_DATA As Variant
Dim MIN_MAX_MULT As Integer

Dim tolerance As Double

On Error GoTo ERROR_LABEL

If MIN_FLAG = True Then 'Minimization
    MIN_MAX_MULT = 1
Else
    MIN_MAX_MULT = -1
End If

CONST_DATA = CONST_RNG

If UBound(CONST_DATA, 2) > 1 Then
    LOWER_BOUND = CONST_DATA(1, 1)       'Xmin
    UPPER_BOUND = CONST_DATA(1, 2)       'Xmax
Else
    LOWER_BOUND = CONST_DATA(1, 1)       'Xmin
    UPPER_BOUND = CONST_DATA(2, 1)       'Xmax
End If

TEMP_DELTA = (3# - Sqr(5#)) / 2

'assert( epsilon > 0 && UPPER_BOUND > LOWER_BOUND );
If (0 < epsilon And LOWER_BOUND < UPPER_BOUND) Then
Else
    GoTo ERROR_LABEL
End If

TEMP_POINT = LOWER_BOUND + TEMP_DELTA * (UPPER_BOUND - LOWER_BOUND)
FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT) * MIN_MAX_MULT
FIRST_POINT = TEMP_POINT
SECOND_POINT = TEMP_POINT
DELTA_FUNC = FIRST_FUNC
SECOND_FUNC = FIRST_FUNC


COUNTER = 0
For i = 1 To nLOOPS
  
  TEMP_RNG = UPPER_BOUND - LOWER_BOUND
  MID_RNG = (LOWER_BOUND + UPPER_BOUND) / 2
  tolerance = (1E-150) * Abs(FIRST_POINT) + epsilon / 3
'  tolerance = 0.000000000000001

  If (Abs(FIRST_POINT - MID_RNG) + TEMP_RNG / 2 <= 2 * tolerance) Then
    Exit For
  End If

  If (FIRST_POINT < MID_RNG) Then
    DTEMP_VAL = UPPER_BOUND - FIRST_POINT
  Else
    DTEMP_VAL = LOWER_BOUND - FIRST_POINT
  End If
  NEW_STEP = TEMP_DELTA * DTEMP_VAL

  If (Abs(FIRST_POINT - SECOND_POINT) >= tolerance) Then
    ATEMP_VAL = (FIRST_POINT - SECOND_POINT) * (DELTA_FUNC - FIRST_FUNC)
    CTEMP_VAL = (FIRST_POINT - TEMP_POINT) * (DELTA_FUNC - SECOND_FUNC)
    BTEMP_VAL = (FIRST_POINT - TEMP_POINT) * CTEMP_VAL - _
                (FIRST_POINT - SECOND_POINT) * ATEMP_VAL
    CTEMP_VAL = 2 * (CTEMP_VAL - ATEMP_VAL)

    If (CTEMP_VAL > 0) Then
      BTEMP_VAL = -BTEMP_VAL
    Else
      CTEMP_VAL = -CTEMP_VAL
    End If

    If (Abs(BTEMP_VAL) < Abs(NEW_STEP * CTEMP_VAL) And _
      BTEMP_VAL > CTEMP_VAL * (LOWER_BOUND - FIRST_POINT + 2 * tolerance) And _
      BTEMP_VAL < CTEMP_VAL * (UPPER_BOUND - FIRST_POINT - 2 * tolerance)) Then
      NEW_STEP = BTEMP_VAL / CTEMP_VAL
    End If

  End If

  If (Abs(NEW_STEP) < tolerance) Then
    If (NEW_STEP > 0) Then
      NEW_STEP = tolerance
    Else
      NEW_STEP = -tolerance
    End If
  End If

  ATEMP_VAL = FIRST_POINT + NEW_STEP
  TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, ATEMP_VAL) * MIN_MAX_MULT
  If (TEMP_FUNC <= DELTA_FUNC) Then
    If (ATEMP_VAL < FIRST_POINT) Then
      UPPER_BOUND = FIRST_POINT
    Else
      LOWER_BOUND = FIRST_POINT
    End If
    TEMP_POINT = SECOND_POINT
    SECOND_POINT = FIRST_POINT
    FIRST_POINT = ATEMP_VAL
    FIRST_FUNC = SECOND_FUNC
    SECOND_FUNC = DELTA_FUNC
    DELTA_FUNC = TEMP_FUNC
  Else
    If (ATEMP_VAL < FIRST_POINT) Then
      LOWER_BOUND = ATEMP_VAL
    Else
      UPPER_BOUND = ATEMP_VAL
    End If
    If (TEMP_FUNC <= SECOND_FUNC Or SECOND_POINT = FIRST_POINT) Then
      TEMP_POINT = SECOND_POINT
      SECOND_POINT = ATEMP_VAL
      FIRST_FUNC = SECOND_FUNC
      SECOND_FUNC = TEMP_FUNC
    ElseIf (TEMP_FUNC <= FIRST_FUNC Or TEMP_POINT = FIRST_POINT Or _
            TEMP_POINT = SECOND_POINT) Then
      TEMP_POINT = ATEMP_VAL
      FIRST_FUNC = TEMP_FUNC
    End If
  End If
  COUNTER = COUNTER + 1
Next i

UNIVAR_MIN_GOLD_FUNC = FIRST_POINT

Exit Function
ERROR_LABEL:
UNIVAR_MIN_GOLD_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : UNIVAR_MIN_PARABOLIC_FUNC

'DESCRIPTION   : Find of local extreme (max or min) with the parabolic method
'For univariate functions only. This algorithm uses a parabolic interpolation to
'find any local extreme (maximum or minimum). It is very efficient and fast
'with smooth-derivable functions. The starting point is simply a segment of
'parameter space bracketing the extreme (local or not) that we want to find.
'The condition is that the extreme must be within the stated segment.

'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_MIN_MAX
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function UNIVAR_MIN_PARABOLIC_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef CONST_RNG As Variant, _
Optional ByRef MIN_FLAG As Boolean, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal epsilon As Double = 0.000000000000001)

Dim X1_VAL As Double
Dim X2_VAL As Double
Dim X3_VAL As Double

Dim Y1_VAL As Double
Dim Y2_VAL As Double
Dim Y3_VAL As Double

Dim TEMP_BOUND As Double
Dim LOWER_BOUND As Double
Dim UPPER_BOUND As Double

Dim CONST_DATA As Variant
Dim CONST_BOX As Variant

Dim XTEMP_ERR As Double
Dim XTEMP_VALUE As Double

Dim X_MIN_VALUE As Double
Dim FUNC_MIN_VAL As Double

On Error GoTo ERROR_LABEL

ReDim CONST_BOX(1 To 2, 1 To 2)

CONST_DATA = CONST_RNG

If UBound(CONST_DATA, 2) > 1 Then
    CONST_BOX(1, 1) = CONST_DATA(1, 1)       'Xmin
    CONST_BOX(1, 2) = CONST_DATA(1, 2)       'Xmax
Else
    CONST_BOX(1, 1) = CONST_DATA(1, 1)       'Xmin
    CONST_BOX(1, 2) = CONST_DATA(2, 1)       'Xmax
End If

LOWER_BOUND = CONST_BOX(1, 1)
UPPER_BOUND = CONST_BOX(1, 2)

X1_VAL = LOWER_BOUND
X2_VAL = (LOWER_BOUND + UPPER_BOUND) / 2
X3_VAL = UPPER_BOUND
XTEMP_VALUE = X1_VAL
Y1_VAL = Excel.Application.Run(FUNC_NAME_STR, XTEMP_VALUE)
XTEMP_VALUE = X2_VAL
Y2_VAL = Excel.Application.Run(FUNC_NAME_STR, XTEMP_VALUE)
XTEMP_VALUE = X3_VAL
Y3_VAL = Excel.Application.Run(FUNC_NAME_STR, XTEMP_VALUE)

If Y2_VAL > (Y1_VAL + Y2_VAL) / 2 Then
    MIN_FLAG = False   'search for max
    Y1_VAL = -Y1_VAL
    Y2_VAL = -Y2_VAL
    Y3_VAL = -Y3_VAL
Else
    MIN_FLAG = True    'search for min
End If

COUNTER = 0

Do
    
    If ((X2_VAL - X1_VAL) * (Y2_VAL - Y3_VAL) - _
        (X2_VAL - X3_VAL) * (Y2_VAL - Y1_VAL)) <> 0 Then
            
        X_MIN_VALUE = X2_VAL - ((X2_VAL - X1_VAL) ^ 2 * (Y2_VAL - Y3_VAL) - _
             (X2_VAL - X3_VAL) ^ 2 * (Y2_VAL - Y1_VAL)) / ((X2_VAL - X1_VAL) * _
             (Y2_VAL - Y3_VAL) - (X2_VAL - X3_VAL) * (Y2_VAL - Y1_VAL)) / 2
             'find of local extreme (max or min) with the parabolic method
    Else
        If (X2_VAL - X1_VAL) <> 0 Then
              X_MIN_VALUE = X1_VAL - (Y2_VAL - Y1_VAL) / (X2_VAL - X1_VAL)
        Else
              X_MIN_VALUE = X1_VAL
        End If
    End If
    

    XTEMP_VALUE = X_MIN_VALUE
  
    FUNC_MIN_VAL = Excel.Application.Run(FUNC_NAME_STR, XTEMP_VALUE)
    If MIN_FLAG = False Then FUNC_MIN_VAL = -FUNC_MIN_VAL
    
    If Y1_VAL > Y2_VAL Then 'Exchange
        TEMP_BOUND = X2_VAL
        X2_VAL = X1_VAL
        X1_VAL = TEMP_BOUND
        TEMP_BOUND = Y2_VAL
        Y2_VAL = Y1_VAL
        Y1_VAL = TEMP_BOUND
    End If
    
    If Y1_VAL > Y3_VAL Then 'Exchange
        TEMP_BOUND = X3_VAL
        X3_VAL = X1_VAL
        X1_VAL = TEMP_BOUND
        TEMP_BOUND = Y3_VAL
        Y3_VAL = Y1_VAL
        Y1_VAL = TEMP_BOUND
    End If
    
    If Y2_VAL > Y3_VAL Then 'Exchange
        TEMP_BOUND = X3_VAL
        X3_VAL = X2_VAL
        X2_VAL = TEMP_BOUND
        TEMP_BOUND = Y3_VAL
        Y3_VAL = Y2_VAL
        Y2_VAL = TEMP_BOUND
    End If
    
    XTEMP_ERR = Abs(X_MIN_VALUE - X3_VAL)
    If Abs(X_MIN_VALUE) > 1 Then XTEMP_ERR = XTEMP_ERR / Abs(X_MIN_VALUE)
    
    Y3_VAL = FUNC_MIN_VAL
    X3_VAL = X_MIN_VALUE
    
    COUNTER = COUNTER + 1

Loop Until COUNTER > nLOOPS Or XTEMP_ERR < epsilon

UNIVAR_MIN_PARABOLIC_FUNC = X_MIN_VALUE

Exit Function
ERROR_LABEL:
UNIVAR_MIN_PARABOLIC_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : UNIVAR_MIN_DIVIDE_CONQUER_FUNC

'DESCRIPTION   : Divide-Conquer 1D; This is another very robust, derivative free
'univariate algorithm. It is simply a modified version of
'the bisection algorithm. It can be adapted to every function,
'smooth or discontinuous. It converges over a very
'large segment of parameter space.

'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_MIN_MAX
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function UNIVAR_MIN_DIVIDE_CONQUER_FUNC(ByVal FUNC_NAME_STR As String, _
ByRef CONST_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal epsilon As Double = 0.000000000000001)

'MIN_FLAG = 1 --> Min
'MIN_FLAG = Else --> Max

'For univariate functions only. It's another very robust, derivative free
'algorithm. It is simply a modified version of the bisection algorithm.
'It can be adapted to every function, smooth or discontinuous. It converges
'over a very large segment of parameter space.

Dim i As Long
Dim j As Long

Dim TEMP_MIN As Double
Dim TEMP_ERR As Double
Dim TEMP_DELTA As Double
Dim TEMP_FACTOR As Double

Dim CONST_DATA As Variant
Dim CONST_BOX As Variant

Dim LOWER_BOUND As Double
Dim UPPER_BOUND As Double

Dim X_MIN_VALUE As Double
Dim FUNC_MIN_VAL As Double

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

Dim MIN_MAX_MULT As Integer

On Error GoTo ERROR_LABEL

If MIN_FLAG = True Then 'Minimization
    MIN_MAX_MULT = 1
Else
    MIN_MAX_MULT = -1
End If

ReDim CONST_BOX(1 To 2, 1 To 2)

CONST_DATA = CONST_RNG

If UBound(CONST_DATA, 2) > 1 Then
    CONST_BOX(1, 1) = CONST_DATA(1, 1)       'Xmin
    CONST_BOX(1, 2) = CONST_DATA(1, 2)       'Xmax
Else
    CONST_BOX(1, 1) = CONST_DATA(1, 1)       'Xmin
    CONST_BOX(1, 2) = CONST_DATA(2, 1)       'Xmax
End If

LOWER_BOUND = CONST_BOX(1, 1)
UPPER_BOUND = CONST_BOX(1, 2)

j = Int(nLOOPS / 30)

ReDim ATEMP_VECTOR(1 To j, 1 To 1)
ReDim BTEMP_VECTOR(1 To j, 1 To 1)

COUNTER = 0

Do
    COUNTER = COUNTER + j 'take equispaced samples
    TEMP_FACTOR = (UPPER_BOUND - LOWER_BOUND) / (j - 1)
    For i = 1 To j
        ATEMP_VECTOR(i, 1) = LOWER_BOUND + (i - 1) * TEMP_FACTOR
        BTEMP_VECTOR(i, 1) = Excel.Application.Run(FUNC_NAME_STR, _
                            ATEMP_VECTOR(i, 1)) * MIN_MAX_MULT
    Next i
    
    TEMP_MIN = 1
    FUNC_MIN_VAL = BTEMP_VECTOR(TEMP_MIN, 1)
    For i = 2 To j
        If BTEMP_VECTOR(i, 1) < FUNC_MIN_VAL Then
            TEMP_MIN = i
            FUNC_MIN_VAL = BTEMP_VECTOR(TEMP_MIN, 1)
        End If
    Next i
    'choose the other bound
    If TEMP_MIN = j Then
        LOWER_BOUND = ATEMP_VECTOR(TEMP_MIN - 1, 1)
        UPPER_BOUND = ATEMP_VECTOR(TEMP_MIN, 1)
        TEMP_DELTA = Abs(BTEMP_VECTOR(TEMP_MIN - 1, 1) _
                    - BTEMP_VECTOR(TEMP_MIN, 1))
    ElseIf TEMP_MIN = 1 Then
        LOWER_BOUND = ATEMP_VECTOR(TEMP_MIN, 1)
        UPPER_BOUND = ATEMP_VECTOR(TEMP_MIN + 1, 1)
        TEMP_DELTA = Abs(BTEMP_VECTOR(TEMP_MIN, 1) _
                    - BTEMP_VECTOR(TEMP_MIN + 1, 1))
    Else
        LOWER_BOUND = ATEMP_VECTOR(TEMP_MIN - 1, 1)
        UPPER_BOUND = ATEMP_VECTOR(TEMP_MIN + 1, 1)
        TEMP_DELTA = Abs(BTEMP_VECTOR(TEMP_MIN - 1, 1) _
                    - BTEMP_VECTOR(TEMP_MIN + 1, 1))
    End If
    TEMP_ERR = Abs(LOWER_BOUND - UPPER_BOUND)
    X_MIN_VALUE = (LOWER_BOUND + UPPER_BOUND) / 2
    If Abs(X_MIN_VALUE) > 1 Then TEMP_ERR = TEMP_ERR / Abs(X_MIN_VALUE)
Loop Until TEMP_ERR < epsilon Or COUNTER > nLOOPS

UNIVAR_MIN_DIVIDE_CONQUER_FUNC = X_MIN_VALUE   'point of minimum

Exit Function
ERROR_LABEL:
UNIVAR_MIN_DIVIDE_CONQUER_FUNC = Err.number
End Function

'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'Input parameters:

'    * LOWER CONSTRAINT BOUND - the left boundary of an interval to search minimum in
'    * UPPER CONSTRAINT BOUND - the right boundary
'    * Epsilon - absolute error of the value of the function minimum.

'------------------------------------------------------------------------------------------
'---------------------------Univariate optimization algorithms-----------------------------
'------------------------------------------------------------------------------------------

