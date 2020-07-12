Attribute VB_Name = "OPTIM_UNIVAR_ZERO_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    
Private Const PUB_EPSILON As Double = 2 ^ 52

'************************************************************************************
'************************************************************************************
'FUNCTION      : BRENT_ZERO_FUNC

'DESCRIPTION   :  Implements the Brent zero finder.
'  Parameters:
'    Input/output, real X, X1.  On input, two points defining the interval
'    in which the search will take place.  F(X) and F(X1) should have opposite
'    signs.  On output, X is the best estimate for the root, and X1 is
'    a recently computed point such that F changes sign in the interval
'    between X and X1.
'  Reference:
'    Richard Brent,
'    Algorithms for Minimization without Derivatives,
'    Prentice Hall, 1973.

'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function BRENT_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef GUESS_VAL As Double = 0, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal tolerance As Double = 0.000000000000001)

'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
Dim LOWER_BOUND As Double
LOWER_BOUND = 0

Dim UPPER_BOUND As Double
UPPER_BOUND = 0

Dim LOWER_ENFORCED As Boolean
LOWER_ENFORCED = False

Dim UPPER_ENFORCED As Boolean
UPPER_ENFORCED = False

If GUESS_VAL = 0 Then
    GUESS_VAL = UPPER_VAL - tolerance
ElseIf GUESS_VAL = UPPER_VAL Then
    GUESS_VAL = UPPER_VAL - tolerance
ElseIf GUESS_VAL = LOWER_VAL Then
    GUESS_VAL = LOWER_VAL + tolerance
End If
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double
Dim ETEMP_VAL As Double
Dim FTEMP_VAL As Double

'Dim CONVERG_STR As String
  
Dim TEMP_MIN As Double
Dim TEMP_MAX As Double
Dim TEMP_ROOT As Double
Dim TEMP_MULT As Double
Dim TEMP_VAL As Double

Dim TEMP_ACC As Double
Dim TEMP_MID As Double

Dim FIRST_MIN As Double
Dim SECOND_MIN As Double

Dim FUNC_ROOT As Double
Dim FUNC_LOWER_VAL As Double
Dim FUNC_UPPER_VAL As Double

Dim epsilon As Double
            
On Error GoTo ERROR_LABEL

CONVERG_VAL = 0
epsilon = 2 ^ -52
TEMP_MULT = tolerance
            
ATEMP_VAL = 0
FTEMP_VAL = 0

TEMP_MIN = LOWER_VAL
TEMP_MAX = UPPER_VAL

If Not TEMP_MIN < TEMP_MAX Then
'    CONVERG_STR = "invalid range: TEMP_MIN (" & TEMP_MIN _
                & ") >= TEMP_MAX (" & TEMP_MAX
    CONVERG_VAL = 1
    GoTo ERROR_LABEL
End If

If Not (LOWER_ENFORCED = 0 Or TEMP_MIN >= LOWER_BOUND) Then
'    CONVERG_STR = "TEMP_MIN (" & TEMP_MIN & ") < enforced low bound (" & _
                LOWER_BOUND & ")"
    CONVERG_VAL = 2
    GoTo ERROR_LABEL
End If

If Not (UPPER_ENFORCED = 0 Or TEMP_MAX <= UPPER_BOUND) Then
'    CONVERG_STR = "TEMP_MAX (" & TEMP_MAX & _
            ") > enforced hi bound (" & UPPER_BOUND & ")"
    CONVERG_VAL = 3
    GoTo ERROR_LABEL
End If
          
FUNC_LOWER_VAL = Excel.Application.Run(FUNC_NAME_STR, TEMP_MIN)
If (Abs(FUNC_LOWER_VAL) < tolerance) Then
    CONVERG_VAL = 4
    BRENT_ZERO_FUNC = TEMP_MIN
    Exit Function
End If

FUNC_UPPER_VAL = Excel.Application.Run(FUNC_NAME_STR, TEMP_MAX)
If (Abs(FUNC_UPPER_VAL) < tolerance) Then
    CONVERG_VAL = 5
    BRENT_ZERO_FUNC = TEMP_MAX
    Exit Function
End If

COUNTER = 2
            
If Not FUNC_LOWER_VAL * FUNC_UPPER_VAL < 0 Then
    'CONVERG_STR = "root not bracketed: f[" & TEMP_MIN & "," & _
                (TEMP_MAX) & "] -> [" & (FUNC_LOWER_VAL) & "," & _
                (FUNC_UPPER_VAL) & "]"
    CONVERG_VAL = 6
    GoTo ERROR_LABEL
End If

If Not GUESS_VAL > TEMP_MIN Then
    'CONVERG_STR = "Solver1D: GUESS_VAL (" & GUESS_VAL & _
                ") < TEMP_MIN (" & TEMP_MIN & ")"
    CONVERG_VAL = 7
    GoTo ERROR_LABEL
End If

If Not GUESS_VAL < TEMP_MAX Then
    'CONVERG_STR = "Solver1D: GUESS_VAL (" & GUESS_VAL & _
            ") > TEMP_MAX (" & TEMP_MAX & ")"
    CONVERG_VAL = 8
    GoTo ERROR_LABEL
End If

TEMP_ROOT = GUESS_VAL
TEMP_ROOT = TEMP_MAX
FUNC_ROOT = FUNC_UPPER_VAL

Do While (COUNTER <= nLOOPS)
    If ((FUNC_ROOT > 0# And FUNC_UPPER_VAL > 0#) Or _
        (FUNC_ROOT < 0# And FUNC_UPPER_VAL < 0#)) Then
        ' Rename TEMP_MIN, TEMP_ROOT, TEMP_MAX and adjust bounds
        TEMP_MAX = TEMP_MIN
        FUNC_UPPER_VAL = FUNC_LOWER_VAL
        ATEMP_VAL = TEMP_ROOT - TEMP_MIN
        FTEMP_VAL = ATEMP_VAL
    End If
    
    If (Abs(FUNC_UPPER_VAL) < Abs(FUNC_ROOT)) Then
        TEMP_MIN = TEMP_ROOT
        TEMP_ROOT = TEMP_MAX
        TEMP_MAX = TEMP_MIN
        FUNC_LOWER_VAL = FUNC_ROOT
        FUNC_ROOT = FUNC_UPPER_VAL
        FUNC_UPPER_VAL = FUNC_LOWER_VAL
    End If ' Convergence check
    
    TEMP_ACC = 2 * epsilon * Abs(TEMP_ROOT) + 0.5 * TEMP_MULT
    TEMP_MID = (TEMP_MAX - TEMP_ROOT) / 2
    
    If (Abs(TEMP_MID) <= TEMP_ACC Or FUNC_ROOT = 0) Then
        BRENT_ZERO_FUNC = TEMP_ROOT
        Exit Function
    End If
    
    If (Abs(FTEMP_VAL) >= TEMP_ACC And Abs(FUNC_LOWER_VAL) > _
        Abs(FUNC_ROOT)) Then
        ' Attempt inverse quadratic interpolation
        ETEMP_VAL = FUNC_ROOT / FUNC_LOWER_VAL
        If (TEMP_MIN = TEMP_MAX) Then
            BTEMP_VAL = 2# * TEMP_MID * ETEMP_VAL
            CTEMP_VAL = 1# - ETEMP_VAL
        Else
            CTEMP_VAL = FUNC_LOWER_VAL / FUNC_UPPER_VAL
            DTEMP_VAL = FUNC_ROOT / FUNC_UPPER_VAL
            BTEMP_VAL = ETEMP_VAL * (2# * TEMP_MID * CTEMP_VAL * (CTEMP_VAL - DTEMP_VAL) - _
            (TEMP_ROOT - TEMP_MIN) * (DTEMP_VAL - 1#))
            CTEMP_VAL = (CTEMP_VAL - 1#) * (DTEMP_VAL - 1#) * (ETEMP_VAL - 1#)
        End If
        If (BTEMP_VAL > 0#) Then
            CTEMP_VAL = -CTEMP_VAL   ' Check whether in bounds
        End If
        
        BTEMP_VAL = Abs(BTEMP_VAL)
        FIRST_MIN = 3# * TEMP_MID * CTEMP_VAL - Abs(TEMP_ACC * CTEMP_VAL)
        SECOND_MIN = Abs(FTEMP_VAL * CTEMP_VAL)
        If FIRST_MIN < SECOND_MIN Then
              TEMP_VAL = FIRST_MIN
        Else
              TEMP_VAL = SECOND_MIN
        End If
        
        If (2# * BTEMP_VAL < TEMP_VAL) Then
            FTEMP_VAL = ATEMP_VAL              ' Accept interpolation
            ATEMP_VAL = BTEMP_VAL / CTEMP_VAL
        Else
            ATEMP_VAL = TEMP_MID ' Interpolation failed, use bisection
            FTEMP_VAL = ATEMP_VAL
        End If
    
    Else ' Bounds decreasing too slowly, use bisection
        ATEMP_VAL = TEMP_MID
        FTEMP_VAL = ATEMP_VAL
    End If
    TEMP_MIN = TEMP_ROOT
    FUNC_LOWER_VAL = FUNC_ROOT
    If (Abs(ATEMP_VAL) > TEMP_ACC) Then
        TEMP_ROOT = TEMP_ROOT + ATEMP_VAL
    Else
        If TEMP_MID >= 0 Then
            TEMP_ROOT = TEMP_ROOT + Abs(TEMP_ACC)
        Else
            TEMP_ROOT = TEMP_ROOT + -Abs(TEMP_ACC)
        End If
    End If
    FUNC_ROOT = Excel.Application.Run(FUNC_NAME_STR, TEMP_ROOT)
    COUNTER = COUNTER + 1
Loop
    

Exit Function
ERROR_LABEL:
    BRENT_ZERO_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : QUADRATIC_ZERO_FUNC

'DESCRIPTION   : Returns an estimate for the root with accuracy
'  4EPSILONabs(x) + tol
'
'  Algorithm
'  G.Forsythe, M.Malcolm, C.Moler, Computer methods for mathematical
'  computations. M., Mir, 1980, p.180 of the Russian edition
'
'  The function makes use of the bissection procedure combined with
'  the linear or quadric inverse interpolation.
'  At every step program operates on three abscissae - a, b, and c.
'  b - the last and the best approximation to the root
'  a - the last but one approximation
'  c - the last but one or even earlier approximation than a that
'    1) |f(b)| <= |f(c)|
'    2) f(b) and f(c) have opposite signs, i.e. b and c confine
'       the root
'  At every step Zeroin selects one of the two new approximations, the
'  former being obtained by the bissection procedure and the latter
'  resulting in the interpolation (if a,b, and c are all different
'  the quadric interpolation is utilized, otherwise the linear one).
'  If the latter (i.e. obtained by the interpolation) point is
'  reasonable (i.e. lies within the current interval [b,c] not being
'  too close to the boundaries) it is accepted. The bissection result
'  is used in the other case. Therefore, the range of uncertainty is
'  ensured to be reduced at least by the factor 1.6

'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function QUADRATIC_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal tolerance As Double = 0.000000000000001)

Dim i As Long

Dim PREV_STEP As Double
Dim NEW_STEP  As Double

Dim FIRST_DELTA As Double
Dim SECOND_DELTA As Double

Dim FIRST_FUNC  As Double
Dim SECOND_FUNC  As Double
Dim THIRD_FUNC  As Double

Dim LOWER_BOUND As Double
Dim UPPER_BOUND As Double
Dim MID_BOUND As Double

Dim TEMP_FUNC As Double
Dim TEMP_LOW As Double
Dim TEMP_MID As Double
Dim TEMP_HIGH As Double

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double

Dim ATEMP_EPS As Double
Dim BTEMP_EPS As Double

Dim epsilon As Double

On Error GoTo ERROR_LABEL

CONVERG_VAL = 0
ATEMP_EPS = 0.000000000000001
BTEMP_EPS = 1E-30
    
LOWER_BOUND = LOWER_VAL       'Xmin
UPPER_BOUND = UPPER_VAL       'Xmax

FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_BOUND)
SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_BOUND)
MID_BOUND = LOWER_BOUND
THIRD_FUNC = FIRST_FUNC

If ( _
  (FIRST_FUNC < 0# And SECOND_FUNC < 0#) Or _
  (FIRST_FUNC > 0# And SECOND_FUNC > 0#)) Then
  CONVERG_VAL = 1
  QUADRATIC_ZERO_FUNC = MID_BOUND
  Exit Function 'Root must be bracketed in zbrent
End If

COUNTER = 0
For i = 1 To nLOOPS
  PREV_STEP = UPPER_BOUND - LOWER_BOUND

  If (Abs(THIRD_FUNC) < Abs(SECOND_FUNC)) Then
    LOWER_BOUND = UPPER_BOUND
    UPPER_BOUND = MID_BOUND
    MID_BOUND = LOWER_BOUND
    FIRST_FUNC = SECOND_FUNC
    SECOND_FUNC = THIRD_FUNC
    THIRD_FUNC = FIRST_FUNC
  End If

  epsilon = ATEMP_EPS
  NEW_STEP = (MID_BOUND - UPPER_BOUND) / 2

  If (Abs(NEW_STEP) <= ATEMP_EPS Or Abs(SECOND_FUNC) <= BTEMP_EPS) Then
    Exit For
  End If

  If (Abs(PREV_STEP) >= epsilon And _
    Abs(FIRST_FUNC) > Abs(SECOND_FUNC)) Then
    CTEMP_VAL = MID_BOUND - UPPER_BOUND
    If (LOWER_BOUND = MID_BOUND) Then
      FIRST_DELTA = SECOND_FUNC / FIRST_FUNC
      ATEMP_VAL = CTEMP_VAL * FIRST_DELTA
      BTEMP_VAL = 1# - FIRST_DELTA
    Else
      BTEMP_VAL = FIRST_FUNC / THIRD_FUNC
      FIRST_DELTA = SECOND_FUNC / THIRD_FUNC
      SECOND_DELTA = SECOND_FUNC / FIRST_FUNC
      ATEMP_VAL = SECOND_DELTA * _
            (CTEMP_VAL * BTEMP_VAL * (BTEMP_VAL - FIRST_DELTA) - _
            (UPPER_BOUND - LOWER_BOUND) * (FIRST_DELTA - 1#))
      BTEMP_VAL = (BTEMP_VAL - 1#) * (FIRST_DELTA - 1#) * (SECOND_DELTA - 1#)
    End If

    If (ATEMP_VAL > 0#) Then
      BTEMP_VAL = -BTEMP_VAL
    Else
      ATEMP_VAL = -ATEMP_VAL
    End If

    If (ATEMP_VAL < (0.75 * CTEMP_VAL * _
        BTEMP_VAL - Abs(epsilon * BTEMP_VAL) / 2) And _
      ATEMP_VAL < Abs(PREV_STEP * BTEMP_VAL / 2)) Then
      NEW_STEP = ATEMP_VAL / BTEMP_VAL
    End If
  End If

  If (Abs(NEW_STEP) < epsilon) Then
    If (NEW_STEP > 0#) Then
      NEW_STEP = epsilon
    Else
      NEW_STEP = -epsilon
    End If
  End If

  LOWER_BOUND = UPPER_BOUND
  FIRST_FUNC = SECOND_FUNC
  UPPER_BOUND = UPPER_BOUND + NEW_STEP
  SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_BOUND)

  If ((SECOND_FUNC > 0 And THIRD_FUNC > 0) Or (SECOND_FUNC < 0 _
  And THIRD_FUNC < 0)) Then
    MID_BOUND = LOWER_BOUND
    THIRD_FUNC = FIRST_FUNC
  End If
  'Debug.Print i & ":   " & Abs(NEW_STEP)
  COUNTER = COUNTER + 1
Next i

' continue and refine by bisection method; assumes already that
' f has changing sign
    TEMP_LOW = LOWER_BOUND - tolerance
    TEMP_HIGH = UPPER_BOUND + tolerance
    For i = 0 To nLOOPS
      TEMP_MID = (TEMP_LOW + TEMP_HIGH) / 2
      TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_MID)
          If (Abs(TEMP_FUNC) < BTEMP_EPS) Or _
            (Abs(TEMP_HIGH - TEMP_LOW) < ATEMP_EPS) Then
            Exit For
          End If
      If 0 <= Sgn(TEMP_FUNC) Then
            TEMP_HIGH = TEMP_MID
      Else
            TEMP_LOW = TEMP_MID
      End If
      COUNTER = COUNTER + 1
    Next i
    
    If (COUNTER > nLOOPS) Then: CONVERG_VAL = 2

    QUADRATIC_ZERO_FUNC = TEMP_MID

Exit Function
ERROR_LABEL:
    QUADRATIC_ZERO_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : NEWTON_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements Newton's method (with 2nd derivative)
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function NEWTON_ZERO_FUNC(ByVal GUESS_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByVal GRAD_STR_NAME As String = "", _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 0.000000000000001)

'-------------------------------------------------------------------------
'Newton-Ralphson method - or simply the Newton's method is one of the
'most commonly used numerical searching method for solving equations.
'Usually Newton's method converges well and quickly, but the convergence
'is not guaranteed.
'-------------------------------------------------------------------------

'Newton's method requires an initial value (GUESS_VAL).  This
'values can determine the way the search is converged. The major challenge
'to using this method is that the first differential (first derivative)
'of the equation is required as an input for the search procedure.
'Sometimes, it may be difficult or impossible to derive that. In this case,
'you must try the finite difference method to approximate the gradient
'-------------------------------------------------------------------------

Dim FUNC_VAL As Double
Dim GRAD_VAL As Double
Dim PARAM_VAL As Double
Dim DELTA_VAL As Double
'Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

    
  CONVERG_VAL = 0
  COUNTER = 0
  PARAM_VAL = GUESS_VAL
  FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, PARAM_VAL)
  
'  ReDim TEMP_VECTOR(0 To COUNTER)
  
  Do '  If the error tolerance is satisfied, then exit.
    If (Abs(FUNC_VAL) <= tolerance) Then
      NEWTON_ZERO_FUNC = PARAM_VAL
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      NEWTON_ZERO_FUNC = PARAM_VAL
      Exit Function
    End If

    If GRAD_STR_NAME <> "" Then
        GRAD_VAL = Excel.Application.Run(GRAD_STR_NAME, PARAM_VAL)
    Else
        GRAD_VAL = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, PARAM_VAL)
    End If
    
    If (GRAD_VAL = 0) Then
      CONVERG_VAL = 3
      NEWTON_ZERO_FUNC = PARAM_VAL
      Exit Function
    End If
    
    DELTA_VAL = -1 * FUNC_VAL / GRAD_VAL '  Set the increment
    PARAM_VAL = PARAM_VAL + DELTA_VAL '  Update the iterate
    'and function values
    FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, PARAM_VAL)
    
  '  ReDim Preserve TEMP_VECTOR(0 To COUNTER)
 '   TEMP_VECTOR(COUNTER) = PARAM_VAL
  Loop
  
'Select Case OUTPUT
 '   Case 0
        NEWTON_ZERO_FUNC = PARAM_VAL
  '  Case 1
   '     NEWTON_ZERO_FUNC = TEMP_VECTOR
'End Select

Exit Function
ERROR_LABEL:
NEWTON_ZERO_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HALLEY_GRAD_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements Halley's method (with 2nd derivative)
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function HALLEY_GRAD_ZERO_FUNC(ByVal GUESS_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
ByVal GRAD_STR_NAME As String, _
ByVal SECOND_GRAD_STR_NAME As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal tolerance As Double = 0.000000000000001)

'---------------------------------------------------------------------------
' HALLEY_GRAD_ZERO_FUNC; implements Halley's method (with 2nd derivative)
'  Parameters:
'    Input/output, real X.
'    On input, an estimate for the root of the equation.
'    On output, if CONVERG_VAL = 0, X is an approximate root for which
'    abs ( F(X) ) <= tolerance.
'---------------------------------------------------------------------------

Dim PARAM_VAL As Double
Dim FUNC_VAL As Double
Dim GRAD_VAL As Double
Dim DERIV_2_VAL As Double

Dim TEMP_GRAD As Double
Dim TEMP_DELTA As Double

Dim epsilon As Double

On Error GoTo ERROR_LABEL

'  Initialization.
'
  CONVERG_VAL = 0
  COUNTER = 0
  
  PARAM_VAL = GUESS_VAL
  FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, PARAM_VAL)
  GRAD_VAL = Excel.Application.Run(GRAD_STR_NAME, PARAM_VAL)
  DERIV_2_VAL = Excel.Application.Run(SECOND_GRAD_STR_NAME, PARAM_VAL)
  epsilon = 5 * tolerance

'  Iteration loop:

  Do
'
'  If the error tolerance is satisfied, then exit.
'
    If (Abs(FUNC_VAL) <= epsilon) Then
      HALLEY_GRAD_ZERO_FUNC = PARAM_VAL
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      HALLEY_GRAD_ZERO_FUNC = PARAM_VAL
      Exit Function
    End If

    If (GRAD_VAL = 0) Then
      CONVERG_VAL = 3
      HALLEY_GRAD_ZERO_FUNC = PARAM_VAL
      Exit Function
    End If

    TEMP_DELTA = DERIV_2_VAL * FUNC_VAL / GRAD_VAL ^ 2

    If (2 - TEMP_DELTA = 0) Then
      CONVERG_VAL = 4
      HALLEY_GRAD_ZERO_FUNC = PARAM_VAL
      Exit Function
    End If

'  Set the increment.
    TEMP_GRAD = -(FUNC_VAL / GRAD_VAL) / (1 - 0.5 * TEMP_DELTA)
'  Update the iterate and function values.
    PARAM_VAL = PARAM_VAL + TEMP_GRAD
   
   FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, PARAM_VAL)
   GRAD_VAL = Excel.Application.Run(GRAD_STR_NAME, PARAM_VAL)
   DERIV_2_VAL = Excel.Application.Run(SECOND_GRAD_STR_NAME, PARAM_VAL)

  Loop

Exit Function
ERROR_LABEL:
    HALLEY_GRAD_ZERO_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HALLEY_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements Halley's method
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function HALLEY_ZERO_FUNC(ByVal GUESS_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
ByVal GRAD_STR_NAME As String, _
ByVal SECOND_GRAD_STR_NAME As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal tolerance As Double = 0.000000000000001)

Dim TEMP_FUNC As Double
Dim FIRST_DERIV As Double
Dim SECOND_DERIV As Double

Dim TEMP_POINT As Double
Dim TEMP_GRAD As Double
Dim DELTA_VAL As Double

Dim epsilon As Double
  
'-------------------------------------------------------------------------------------
' HALLEY1 implements Halley's method.
'  Parameters:
'    Input/output, real X.
'    On input, an estimate for the root of the equation.
'    On output, if IERROR = 0, X is an approximate root for which
'    abs ( F(X) ) <= ABSERR.
'-------------------------------------------------------------------------------------

On Error GoTo ERROR_LABEL
  
  TEMP_POINT = GUESS_VAL
  
  CONVERG_VAL = 0
  COUNTER = 0
  TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)
  FIRST_DERIV = Excel.Application.Run(GRAD_STR_NAME, TEMP_POINT)
  SECOND_DERIV = Excel.Application.Run(SECOND_GRAD_STR_NAME, TEMP_POINT)
  epsilon = 5 * tolerance
'
'  Iteration loop:
'
  Do
'
'  If the error tolerance is satisfied, then exit.
'
    If (Abs(TEMP_FUNC) <= epsilon) Then
      HALLEY_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      HALLEY_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    If (FIRST_DERIV = 0) Then
      CONVERG_VAL = 3
      HALLEY_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    DELTA_VAL = SECOND_DERIV * TEMP_FUNC / FIRST_DERIV ^ 2

    If (2 - DELTA_VAL = 0) Then
      CONVERG_VAL = 4
      HALLEY_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

'  Set the increment.
'
    TEMP_GRAD = -(TEMP_FUNC / FIRST_DERIV) / (1 - 0.5 * DELTA_VAL)
'
'  Update the iterate and function values.
'
    TEMP_POINT = TEMP_POINT + TEMP_GRAD
    TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)
    FIRST_DERIV = Excel.Application.Run(GRAD_STR_NAME, TEMP_POINT)
    SECOND_DERIV = Excel.Application.Run(SECOND_GRAD_STR_NAME, TEMP_POINT)

  Loop

Exit Function
ERROR_LABEL:
    HALLEY_ZERO_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : WDB_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the Wijngaarden-Dekker-Brent zero finder
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function WDB_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal tolerance As Double = 0.000000000000001)

'    Wijngaarden-Dekker-Brent zero finder.
'    Input/output, real X, X1.  On input, two points defining the interval
'    in which the search will take place.  F(X) and F(X1) should have opposite
'    signs.  On output, X is the best estimate for the root, and X1 is
'    a recently computed point such that F changes sign in the interval
'    between X and X1.
'  Reference:
'    Richard Brent,
'    Algorithms for Minimization without Derivatives,
'    Prentice Hall, 1973.

Dim h As Double
Dim i As Double
Dim j As Double
Dim k As Double

Dim hh As Double
Dim ii As Double
Dim jj As Double
Dim kk As Double

Dim TEMP_VAL As Double
Dim GRAD_VAL As Double

Dim MIN_FUNC_VAL As Double
Dim TEMP_POINT_VAL As Double

Dim FIRST_FUNC_VAL As Double
Dim SEC_FUNC_VAL As Double

Dim epsilon As Double

On Error GoTo ERROR_LABEL

'  Initialization.
'
  CONVERG_VAL = 0
  COUNTER = 0
  FIRST_FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  MIN_FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)
  
  TEMP_POINT_VAL = UPPER_VAL
  SEC_FUNC_VAL = MIN_FUNC_VAL
  epsilon = tolerance

'  Iteration loop:

  Do

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      WDB_ZERO_FUNC = UPPER_VAL
      Exit Function
    End If
    If (MIN_FUNC_VAL > 0 And SEC_FUNC_VAL > 0) Or _
                        (MIN_FUNC_VAL < 0 And _
                        SEC_FUNC_VAL < 0) Then
      TEMP_POINT_VAL = LOWER_VAL
      SEC_FUNC_VAL = FIRST_FUNC_VAL
      i = UPPER_VAL - LOWER_VAL
      j = UPPER_VAL - LOWER_VAL
    End If

    If (Abs(SEC_FUNC_VAL) < Abs(MIN_FUNC_VAL)) Then
      TEMP_VAL = UPPER_VAL
      UPPER_VAL = TEMP_POINT_VAL
      TEMP_POINT_VAL = TEMP_VAL
      
      TEMP_VAL = MIN_FUNC_VAL
      MIN_FUNC_VAL = SEC_FUNC_VAL
      SEC_FUNC_VAL = TEMP_VAL
    End If
    k = 0.5 * (TEMP_POINT_VAL - UPPER_VAL)
    If (Abs(k) <= epsilon) Or (Abs(MIN_FUNC_VAL) < epsilon) Then
      WDB_ZERO_FUNC = UPPER_VAL
      Exit Function
    End If
    If (Abs(j) < epsilon) Or (Abs(FIRST_FUNC_VAL) <= Abs(MIN_FUNC_VAL)) Then
      i = k
      j = k
    Else
      ii = MIN_FUNC_VAL / FIRST_FUNC_VAL
      If (LOWER_VAL = TEMP_POINT_VAL) Then
        jj = 2 * k * ii
        kk = 1 - ii
'        Debug.Print "secante"
      Else
        hh = FIRST_FUNC_VAL / SEC_FUNC_VAL
        h = MIN_FUNC_VAL / SEC_FUNC_VAL
        jj = ii * (2 * k * hh * (hh - h) - (UPPER_VAL - LOWER_VAL) * (h - 1))
        kk = (hh - 1) * (h - 1) * (ii - 1)
'        Debug.Print "inverse quadratic"
      End If

      If (jj > 0) Then
        kk = -kk
      Else
        jj = -jj
      End If

      ii = j
      j = i

      If ((2 * jj < 3 * k * kk - Abs(epsilon * kk)) Or _
                (jj < Abs(0.5 * ii * kk))) Then
        i = jj / kk       'Accept interpolation.
      Else
        i = k          'Interpolation failed, use bisection.
        j = k
'        Debug.Print "Interpolation failed"
      End If
    End If
'  Set the increment.
    If (Abs(i) > epsilon) Then
      GRAD_VAL = i
    ElseIf (k > 0) Then
      GRAD_VAL = epsilon
    Else
      GRAD_VAL = -epsilon
    End If
'  Remember current data for next step.
    LOWER_VAL = UPPER_VAL
    FIRST_FUNC_VAL = MIN_FUNC_VAL
'  Update the iterate and function values.
    UPPER_VAL = UPPER_VAL + GRAD_VAL
    MIN_FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)
  Loop

Exit Function
ERROR_LABEL:
    WDB_ZERO_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : BISEC_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the the bisection method
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function BISEC_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal tolerance As Double = 0.000000000000001)


'---------------------------------------------------------------------------------
'Bisec method, unlike the Newton-Ralphson method, does not require the
'main function in question. However, it does requires an initial min and max value.
'These values can determine the way the search is converged. The major challenge
'to using this method is that the first differential (first derivative)
'of the equation is required as an input for the search procedure.
'Sometimes, it may be difficult or impossible to derive that.
'---------------------------------------------------------------------------------

'*******************************************************************************
' BISECT carries out the bisection method.
'  Parameters:
'    Input/output, real X, X1, two points defining the interval in which the
'    search will take place.  F(X) and F(X1) should have opposite signs.
'    On output, X is the best estimate for the root, and X1 is
'    a recently computed point such that F changes sign in the interval
'    between X and X1.
'*******************************************************************************
  
  Dim FIRST_FUNC_VAL As Double
  Dim SEC_FUNC_VAL As Double
  
  Dim MIN_FUNC_VAL As Double
  Dim PARAM_VAL As Double
  
  Dim TEMP_VAL As Double
  
  On Error GoTo ERROR_LABEL

  CONVERG_VAL = 0
  COUNTER = 0

'  Evaluate the function at the starting points.

  FIRST_FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL) ‘ = CALL_IRR_OBJ_FUNC(X_VAL)
  SEC_FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)

'  Set XPOS and XNEG to the LOWER_VAL values for which F(LOWER_VAL) is positive
'  and negative, respectively.

  If (FIRST_FUNC_VAL >= 0# And SEC_FUNC_VAL <= 0#) Then

  ElseIf (FIRST_FUNC_VAL <= 0# And SEC_FUNC_VAL >= 0#) Then
      TEMP_VAL = LOWER_VAL
      LOWER_VAL = UPPER_VAL
      UPPER_VAL = TEMP_VAL
    
      TEMP_VAL = FIRST_FUNC_VAL
      FIRST_FUNC_VAL = SEC_FUNC_VAL
      SEC_FUNC_VAL = TEMP_VAL
  Else
    CONVERG_VAL = 1
        BISEC_ZERO_FUNC = PARAM_VAL
        Exit Function
  End If

'  Iteration loop:

  Do
'
'  If the error tolerance is satisfied, then exit.
'
    If (Abs(LOWER_VAL - UPPER_VAL) <= tolerance) Then
        BISEC_ZERO_FUNC = PARAM_VAL
        Exit Function
    End If

    If (Abs(FIRST_FUNC_VAL) <= tolerance) Then
        BISEC_ZERO_FUNC = PARAM_VAL
        Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
        BISEC_ZERO_FUNC = PARAM_VAL
        Exit Function
    End If

'  Update the iterate and function values.
    PARAM_VAL = (UPPER_VAL + LOWER_VAL) / 2
    MIN_FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, PARAM_VAL)

    If (MIN_FUNC_VAL >= 0) Then
      LOWER_VAL = PARAM_VAL
      FIRST_FUNC_VAL = MIN_FUNC_VAL
    Else
      UPPER_VAL = PARAM_VAL
      SEC_FUNC_VAL = MIN_FUNC_VAL
    End If

  Loop


Exit Function
ERROR_LABEL:
    BISEC_ZERO_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SECANT_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the Secant method
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function SECANT_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal tolerance As Double = 0.000000000000001)


'------------------------------------------------------------------------------
'Secant method, unlike the Newton-Ralphson method, does not require the
'differentiation of the equation in question.  Because of that, it can be
'used to solve complex equations without the difficulty that one might have
'to encounter in trying to differentiate the equations.  Secant method requires
'two initial values.  Test shows that this method converge little bit slower _
'than the Newton-Ralphson method.
'------------------------------------------------------------------------------


Dim GRAD_VAL As Double
Dim MIN_FUNC_VAL As Double
Dim TEMP_FUNC_VAL As Double

' SECANT_ZERO_FUNC implements the SECANT_ZERO_FUNC method.
'    Input/output, real X, X1.
'    On input, two distinct points that start the method.
'    On output, X is an approximation to a root of the equation
'    which satisfies abs ( F(X) ) < ABSERR, and X1 is the previous
'    estimate.
'  Reference:
'    Joseph Traub,
'    Iterative Methods for the Solution of Equations,
'    Prentice Hall, 1964.
'    Input, integer KMAX, the maximum number of iterations allowed.

On Error GoTo ERROR_LABEL

'  Initialization.
'
  CONVERG_VAL = 0
  COUNTER = 0
  MIN_FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  TEMP_FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)
'
'  Iteration loop:
'
  Do

'  If the error tolerance is satisfied, then exit.

    If (Abs(MIN_FUNC_VAL) <= tolerance) Then
      SECANT_ZERO_FUNC = LOWER_VAL
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      SECANT_ZERO_FUNC = LOWER_VAL
      Exit Function
    End If

    If ((MIN_FUNC_VAL - TEMP_FUNC_VAL) = 0#) Then
      CONVERG_VAL = 3
      SECANT_ZERO_FUNC = LOWER_VAL
      Exit Function
    End If

'  Set the increment.

    GRAD_VAL = -MIN_FUNC_VAL * (LOWER_VAL - UPPER_VAL) / _
                (MIN_FUNC_VAL - TEMP_FUNC_VAL)

'  Remember current data for next step.
    
    UPPER_VAL = LOWER_VAL
    TEMP_FUNC_VAL = MIN_FUNC_VAL

'  Update the iterate and function values.
    
    LOWER_VAL = LOWER_VAL + GRAD_VAL
    MIN_FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)

  Loop

Exit Function
ERROR_LABEL:
    SECANT_ZERO_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PEGASO_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'determines the zero by applying the Pegasus-method
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PEGASO_ZERO_FUNC(ByVal LOWER_BOUND As Double, _
ByVal UPPER_BOUND As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal tolerance As Double = 0.000000000000001)

Dim i As Long

Dim PARAM_VAL As Double
Dim TEMP_DELTA As Double

Dim LOWER_VAL As Double
Dim UPPER_VAL As Double
Dim TEMP_POINT_VAL As Double

Dim FIRST_FUNC_VAL As Double
Dim SECOND_FUNC_VAL As Double
Dim MIN_FUNC_VAL As Double

Dim epsilon As Double

'-----------------------------------------------------------------------------------
' PEGASUS Dowell-Jarrat method (a modified regula falsi method)
' Zero_Pegaso determines the zero by applying the Pegasus-method
'-----------------------------------------------------------------------------------
'  Parameters:
'  a,b   : endpoints of the interval containing a zero.
'  MaxIt : maximum number of iteration steps.
'  xsi   : zero or approximate value for the zero
'          of the function Funct.
'  x1,x2 : the last two iterates, so that [X1,X2] contains a
'          zero of Funct.
'  Numit : number of iteration steps executed.
'  Ierr  : = 0 OK
'          = 3, the maximum number of iteration steps was
'               reached without meeting the break-off criterion.
'  abserr : absolute error tolerance
'
'  Derived from the original FORTRAN version by                                                              *
'  author     : Gisela Engeln-Muellges
'  date       : 09.02.1985
'  source     : ESFR FORTRAN 77 Library
'-----------------------------------------------------------------------------------

On Error GoTo ERROR_LABEL

      COUNTER = 0
      epsilon = 4 * tolerance
      LOWER_VAL = LOWER_BOUND
      UPPER_VAL = UPPER_BOUND

'  calculating the functional values at the endpoints LOWER_VAL and UPPER_VAL.
      
      FIRST_FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
      SECOND_FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)

'  testing for alternate signs of Funct at LOWER_VAL and UPPER_VAL:
' Excel.Application.Run(FUNC_NAME_STR,LOWER_VAL)* _
  Excel.Application.Run(FUNC_NAME_STR,UPPER_VAL) < 0.0 .
      
      If (FIRST_FUNC_VAL * SECOND_FUNC_VAL > 0) Then
         CONVERG_VAL = -1
         PEGASO_ZERO_FUNC = LOWER_VAL
         Exit Function
      ElseIf (FIRST_FUNC_VAL * SECOND_FUNC_VAL = 0) Then
         CONVERG_VAL = 0
         PEGASO_ZERO_FUNC = UPPER_VAL
         Exit Function
      End If
'
'  executing the Pegasus-method.
      For i = 1 To nLOOPS
         COUNTER = i
'     testing whether the value of SECOND_FUNC_VAL is less than four times
         If (Abs(SECOND_FUNC_VAL) < epsilon) Then
            TEMP_POINT_VAL = UPPER_VAL
            CONVERG_VAL = 0
            PEGASO_ZERO_FUNC = TEMP_POINT_VAL
            Exit Function
'     testing for the break-off criterion.
         ElseIf (Abs(UPPER_VAL - LOWER_VAL) <= Abs(UPPER_VAL) _
                    * tolerance) Then
            TEMP_POINT_VAL = UPPER_VAL
            If (Abs(FIRST_FUNC_VAL) < Abs(SECOND_FUNC_VAL)) Then _
                    TEMP_POINT_VAL = LOWER_VAL
            CONVERG_VAL = 0
            PEGASO_ZERO_FUNC = TEMP_POINT_VAL
            Exit Function
         Else
'     calculating the secant slope.
            TEMP_DELTA = (SECOND_FUNC_VAL - FIRST_FUNC_VAL) / _
                    (UPPER_VAL - LOWER_VAL)
'     calculating the secant intercept PARAM_VAL with the x-axis.
            PARAM_VAL = UPPER_VAL - SECOND_FUNC_VAL / TEMP_DELTA
'     calculating LOWER_VAL new functional value at PARAM_VAL.
            MIN_FUNC_VAL = Excel.Application.Run(FUNC_NAME_STR, PARAM_VAL)
'     definition of the endpoints of LOWER_VAL smaller inclusion interval.
            If (SECOND_FUNC_VAL * MIN_FUNC_VAL <= 0) Then
               LOWER_VAL = UPPER_VAL
               FIRST_FUNC_VAL = SECOND_FUNC_VAL
            Else
               FIRST_FUNC_VAL = FIRST_FUNC_VAL * SECOND_FUNC_VAL / _
                                (SECOND_FUNC_VAL + MIN_FUNC_VAL)
            End If
            UPPER_VAL = PARAM_VAL
            SECOND_FUNC_VAL = MIN_FUNC_VAL
         End If
      Next i
      
      CONVERG_VAL = -3
      PEGASO_ZERO_FUNC = TEMP_POINT_VAL

Exit Function
ERROR_LABEL:
    PEGASO_ZERO_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : TRAUB_PHI_21_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the Traub capital PHI(2,1) function
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function TRAUB_PHI_21_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal tolerance As Double = 0.000000000000001)

Dim TEMP_VAL As Double
Dim TEMP_MID As Double
Dim TEMP_MULT As Double
Dim TEMP_DET As Double
Dim TEMP_GRAD As Double

Dim TEMP_FUNC As Double
Dim FIRST_FUNC As Double
Dim SECOND_FUNC As Double

Dim FIRST_DELTA As Double
Dim SECOND_DELTA As Double

On Error GoTo ERROR_LABEL

'-------------------------------------------------------------------------------
'' CAP_PHI_21 implements the Traub capital PHI(2,1) function.
'
'  Reference:
'
'    Joseph Traub,
'    Iterative Methods for the Solution of Equations,
'    Prentice Hall, 1964.
'
'  Modified:
'
'    20 July 2000
'
'  Author:
'
'    John Burkardt
'
'  Parameters:
'
'    Input/output, real X, X1, X2.
'    On input, three distinct points that start the method.
'    On output, X is an approximation to a root of the equation
'    which satisfies abs ( F(X) ) < ABSERR, and X1 and X2 are the
'    previous estimates.
'
'    Input, real ABSERR, an error tolerance.
'
'    Input, real external F, the name of the routine that evaluates the
'    function or its derivatives, of the form
'      function f ( x, ider )
'
'    Output, integer IERROR, error indicator.
'    0, no error occurred.
'    nonzero, an error occurred, and the iteration was halted.
'
'    Output, integer K, the number of steps taken.
'
'    Input, integer KMAX, the maximum number of iterations allowed.
'-------------------------------------------------------------------------------

'  Initialization.
   TEMP_MID = (LOWER_VAL + UPPER_VAL) / 2
  CONVERG_VAL = 0
  COUNTER = -2
  SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)
  COUNTER = -1
  FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  COUNTER = 0
  TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_MID)

  If (LOWER_VAL = UPPER_VAL) Then
    CONVERG_VAL = 3
    TRAUB_PHI_21_ZERO_FUNC = TEMP_MID
    Exit Function
  End If

  FIRST_DELTA = (FIRST_FUNC - SECOND_FUNC) / (LOWER_VAL - UPPER_VAL)
'
'  Iteration loop:
'
  Do
    ' write partial results <<<<<<<<<<<<
        TEMP_VAL = Abs(TEMP_FUNC)
        If TEMP_VAL < tolerance Then TEMP_VAL = tolerance
    ' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'
'  If the error tolerance is satisfied, then exit.
'
    If (Abs(TEMP_FUNC) <= tolerance) Then
      TRAUB_PHI_21_ZERO_FUNC = TEMP_MID
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      TRAUB_PHI_21_ZERO_FUNC = TEMP_MID
      Exit Function
    End If

    SECOND_DELTA = FIRST_DELTA

    If (TEMP_MID = LOWER_VAL) Then
      CONVERG_VAL = 3
      TRAUB_PHI_21_ZERO_FUNC = TEMP_MID
      Exit Function
    End If

    FIRST_DELTA = (TEMP_FUNC - FIRST_FUNC) / (TEMP_MID - LOWER_VAL)

    If (TEMP_MID = UPPER_VAL) Then
      CONVERG_VAL = 3
      TRAUB_PHI_21_ZERO_FUNC = TEMP_MID
      Exit Function
    End If

    SECOND_DELTA = (FIRST_DELTA - SECOND_DELTA) / (TEMP_MID - UPPER_VAL)
    TEMP_MULT = FIRST_DELTA + (TEMP_MID - LOWER_VAL) * SECOND_DELTA
    TEMP_DET = TEMP_MULT ^ 2 - 4 * TEMP_FUNC * SECOND_DELTA
    
    If TEMP_DET < 0 Then TEMP_DET = 0
'
'  Set the increment. parabola interpolation
'
    TEMP_GRAD = -2 * TEMP_FUNC / (TEMP_MULT + Sqr(TEMP_DET))
'
'  Remember current data for next step.
'
    UPPER_VAL = LOWER_VAL
    SECOND_FUNC = FIRST_FUNC

    LOWER_VAL = TEMP_MID
    FIRST_FUNC = TEMP_FUNC
'
'  Update the iterate and function values.
'
    TEMP_MID = TEMP_MID + TEMP_GRAD
    TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_MID)

  Loop

Exit Function
ERROR_LABEL:
    TRAUB_PHI_21_ZERO_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : TRAUB_E_21_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the Traub perp E 21 method
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function TRAUB_E_21_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal tolerance As Double = 0.000000000000001)

Dim TEMP_VAL As Double
Dim TEMP_MID As Double
Dim TEMP_FUNC As Double
Dim TEMP_GRAD As Double
Dim TEMP_FRAC As Double

Dim FIRST_FUNC As Double
Dim SECOND_FUNC As Double

Dim FIRST_DELTA As Double
Dim SECOND_DELTA As Double

On Error GoTo ERROR_LABEL

'---------------------------------------------------------------------------
'' PERP_E_21 implements the Traub perp E 21 method.
'  Reference:
'    Joseph Traub,
'    Iterative Methods for the Solution of Equations,
'    Prentice Hall, 1964, page 233.
'  Parameters:
'    Input/output, real X, X1, X2.
'    On input, three distinct points that start the method.
'    On output, X is an approximation to a root of the equation
'    which satisfies abs ( F(X) ) < ABSERR, and X1 and X2 are the
'    previous estimates.
'
'    Input, real ABSERR, an error tolerance.
'
'    Input, real external F, the name of the routine that evaluates the
'    function or its derivatives, of the form
'      function f ( x, ider )
'
'    Output, integer IERROR, error indicator.
'    0, no error occurred.
'    nonzero, an error occurred, and the iteration was halted.
'
'    Output, integer K, the number of steps taken.
'
'    Input, integer KMAX, the maximum number of iterations allowed.
'
' 1 function evaluation + 13 operations
'----------------------------------------------------------------------------

  CONVERG_VAL = 0
  COUNTER = 0
  TEMP_MID = (LOWER_VAL + UPPER_VAL) / 2
  TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_MID)
  FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)

  If (LOWER_VAL = UPPER_VAL) Then
    CONVERG_VAL = 3
    TRAUB_E_21_ZERO_FUNC = TEMP_MID
    Exit Function
  End If

  FIRST_DELTA = (FIRST_FUNC - SECOND_FUNC) / (LOWER_VAL - UPPER_VAL)
'
'  Iteration loop:
'
  Do
        TEMP_VAL = Abs(TEMP_FUNC)
        If TEMP_VAL < tolerance Then TEMP_VAL = tolerance

'  If the error tolerance is satisfied, then exit.
'
    If (Abs(TEMP_FUNC) <= tolerance) Then
      TRAUB_E_21_ZERO_FUNC = TEMP_MID
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      TRAUB_E_21_ZERO_FUNC = TEMP_MID
      Exit Function
    End If

    SECOND_DELTA = FIRST_DELTA

    If (TEMP_MID = LOWER_VAL) Then
      CONVERG_VAL = 3
      TRAUB_E_21_ZERO_FUNC = TEMP_MID
      Exit Function
    End If

    If (TEMP_MID = UPPER_VAL) Then
      CONVERG_VAL = 3
      TRAUB_E_21_ZERO_FUNC = TEMP_MID
      Exit Function
    End If

    FIRST_DELTA = (TEMP_FUNC - FIRST_FUNC) / (TEMP_MID - LOWER_VAL)
    TEMP_FRAC = (TEMP_FUNC - SECOND_FUNC) / (TEMP_MID - UPPER_VAL)

'  Set the increment.
    TEMP_GRAD = -TEMP_FUNC * (1 / FIRST_DELTA + _
            1 / TEMP_FRAC - 1 / SECOND_DELTA)
'  Remember current data for next step.
    UPPER_VAL = LOWER_VAL
    SECOND_FUNC = FIRST_FUNC
    LOWER_VAL = TEMP_MID
    FIRST_FUNC = TEMP_FUNC
'  Update the iterate and function values.
    TEMP_MID = TEMP_MID + TEMP_GRAD
    TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_MID)
Loop

Exit Function
ERROR_LABEL:
    TRAUB_E_21_ZERO_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MULLER_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the Muller's method
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function MULLER_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal tolerance As Double = 0.000000000000001)

Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double
Dim DTEMP_VAL As Double

Dim TEMP_MID As Double
Dim TEMP_MULT As Double
Dim TEMP_GRAD As Double

Dim TEMP_FUNC As Double
Dim FIRST_FUNC As Double
Dim SECOND_FUNC As Double

On Error GoTo ERROR_LABEL
'-------------------------------------------------------------------------------
' MULLER implements Muller's method
'
'  Parameters:
'
'    Input/output, real X, X1, X2.
'    On input, three distinct points that start the method.
'    On output, X is an approximation to a root of the equation
'    which satisfies abs ( F(X) ) < ABSERR, and X1 and X2 are the
'    previous estimates.
'-------------------------------------------------------------------------------

  CONVERG_VAL = 0
  COUNTER = 0
  TEMP_MID = (LOWER_VAL + UPPER_VAL) / 2
  
  SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)
  FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_MID)

'  Iteration loop:

  Do
'
'  If the error tolerance is satisfied, then exit.
'
    If (Abs(TEMP_FUNC) <= tolerance) Then
      MULLER_ZERO_FUNC = TEMP_MID
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      MULLER_ZERO_FUNC = TEMP_MID
      Exit Function
    End If

    DTEMP_VAL = (TEMP_MID - LOWER_VAL) / (LOWER_VAL - UPPER_VAL)
    'variabile normalizzata   0 < DTEMP < 1

    ATEMP_VAL = DTEMP_VAL * TEMP_FUNC - DTEMP_VAL * (1 + DTEMP_VAL) * FIRST_FUNC + _
            DTEMP_VAL ^ 2 * SECOND_FUNC
    BTEMP_VAL = (2 * DTEMP_VAL + 1) * TEMP_FUNC - (1 + DTEMP_VAL) ^ 2 * FIRST_FUNC + _
        DTEMP_VAL ^ 2 * SECOND_FUNC
    
    CTEMP_VAL = (1 + DTEMP_VAL) * TEMP_FUNC

    TEMP_MULT = BTEMP_VAL ^ 2 - 4 * ATEMP_VAL * CTEMP_VAL
    If TEMP_MULT < 0 Then TEMP_MULT = 0
    TEMP_MULT = Sqr(TEMP_MULT)
    If (BTEMP_VAL < 0) Then: TEMP_MULT = -TEMP_MULT

'  Set the increment.
'
    TEMP_GRAD = -(TEMP_MID - LOWER_VAL) * 2 * CTEMP_VAL / (BTEMP_VAL + TEMP_MULT)
'
'  Remember current data for next step.
'
    UPPER_VAL = LOWER_VAL
    SECOND_FUNC = FIRST_FUNC
    LOWER_VAL = TEMP_MID
    FIRST_FUNC = TEMP_FUNC
'
'  Update the iterate and function values.
'
    TEMP_MID = TEMP_MID + TEMP_GRAD
    TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_MID)

  Loop

Exit Function
ERROR_LABEL:
    MULLER_ZERO_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RHEIN1_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the Rheinboldt bisection - secant method
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function RHEIN1_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal tolerance As Double = 0.000000000000001)


  Dim i As Long
  
  Dim ATEMP_VAL As Double
  Dim BTEMP_VAL As Double
  Dim CTEMP_VAL As Double
    
  Dim TEMP_VAL As Double
  Dim TEMP_FUNC As Double
  
  Dim FIRST_FUNC As Double
  Dim SECOND_FUNC As Double
  
  Dim TEMP_DELTA As Double
  Dim TEMP_POINT As Double
  Dim TEMP_GRAD As Double
  Dim TEMP_SWAP As Double
  
  Dim FORCE_FLAG As Boolean

  On Error GoTo ERROR_LABEL
'------------------------------------------------------------------------
'' RHEIN1 implements the Rheinboldt bisection - secant method.
'
'  Reference:
'
'    W C Rheinboldt,
'    Algorithms for finding zeros of a function
'    UMAP Journal,
'    Volume 2, 1, 1981, pages 43 - 72.
'
'  Parameters:
'
'    Input/output, real X, X1.
'    On input, two distinct points that start the method.
'    On output, X is an approximation to a root of the equation
'    which satisfies abs ( F(X) ) < ABSERR, and X1 is the previous
'    estimate.
'
'    Input, real ABSERR, an error tolerance.
'
'    Input, real external F, the name of the routine that evaluates the
'    function or its derivatives, of the form
'      function f ( x, ider )
'
'    Output, integer IERROR, error indicator.
'    0, no error occurred.
'    nonzero, an error occurred, and the iteration was halted.
'
'    Output, integer CTEMP, the number of steps taken.
'
'    Input, integer KMAX, the maximum number of iterations allowed.
'------------------------------------------------------------------------

'  Initialization.

  i = 0
  CONVERG_VAL = 0
  COUNTER = 0
  TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)

  If (Abs(TEMP_FUNC) > Abs(FIRST_FUNC)) Then
      TEMP_SWAP = LOWER_VAL
      LOWER_VAL = UPPER_VAL
      UPPER_VAL = TEMP_SWAP
    
      TEMP_SWAP = TEMP_FUNC
      TEMP_FUNC = FIRST_FUNC
      FIRST_FUNC = TEMP_SWAP
  End If

  TEMP_POINT = UPPER_VAL
  SECOND_FUNC = FIRST_FUNC
  TEMP_DELTA = 0.5 * Abs(LOWER_VAL - UPPER_VAL)

'  Iteration loop:

  Do ' write partial results <<<<<<<<<<<<
        TEMP_VAL = Abs(TEMP_FUNC)
        If TEMP_VAL < tolerance Then TEMP_VAL = tolerance
'  If the error tolerance is satisfied, then exit.
    
    If (Abs(TEMP_FUNC) <= tolerance) Then
      RHEIN1_ZERO_FUNC = LOWER_VAL
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      RHEIN1_ZERO_FUNC = LOWER_VAL
      Exit Function
    End If
'
'  Force ABS ( TEMP_FUNC ) <= ABS ( FIRST_FUNC ).
'
    If (Abs(TEMP_FUNC) > Abs(FIRST_FUNC)) Then

      TEMP_POINT = LOWER_VAL
      LOWER_VAL = UPPER_VAL
      UPPER_VAL = TEMP_POINT

      SECOND_FUNC = TEMP_FUNC
      TEMP_FUNC = FIRST_FUNC
      FIRST_FUNC = SECOND_FUNC

    End If

    ATEMP_VAL = 0.5 * (UPPER_VAL - LOWER_VAL)
'
'  Compute the numerator and denominator for secant step.
'
    BTEMP_VAL = (LOWER_VAL - TEMP_POINT) * TEMP_FUNC
    CTEMP_VAL = SECOND_FUNC - TEMP_FUNC

    If (BTEMP_VAL < 0) Then
      BTEMP_VAL = -BTEMP_VAL
      CTEMP_VAL = -CTEMP_VAL
    End If
'
'  Save the old minimum residual point.
'
    TEMP_POINT = LOWER_VAL
    SECOND_FUNC = TEMP_FUNC
'
'  Test for FORCE_FLAG bisection.
'
    i = i + 1
    FORCE_FLAG = False

    If (i > 3) Then
      If (8 * Abs(ATEMP_VAL) > TEMP_DELTA) Then
        FORCE_FLAG = True
      Else
        i = 0
        TEMP_DELTA = ATEMP_VAL
      End If
    End If
'
'  Set the increment.

    If (FORCE_FLAG) Then
      TEMP_GRAD = ATEMP_VAL
    ElseIf (BTEMP_VAL <= Abs(CTEMP_VAL) * tolerance) Then
      TEMP_GRAD = IIf(ATEMP_VAL > 0, Abs(1), -Abs(1)) * tolerance
    ElseIf (BTEMP_VAL < CTEMP_VAL * ATEMP_VAL) Then
      TEMP_GRAD = BTEMP_VAL / CTEMP_VAL
    Else
      TEMP_GRAD = ATEMP_VAL
    End If
'
'  Update the iterate and function values.
'
    LOWER_VAL = LOWER_VAL + TEMP_GRAD
    TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
'
'  Preserve the change of sign interval.
'
    If IIf(TEMP_FUNC > 0, Abs(1), -Abs(1)) = _
       IIf(FIRST_FUNC > 0, Abs(1), -Abs(1)) Then
      UPPER_VAL = TEMP_POINT
      FIRST_FUNC = SECOND_FUNC
    End If
    

  Loop

Exit Function
ERROR_LABEL:
    RHEIN1_ZERO_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RHEIN2_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the Rheinboldt bisection - secant - inverse quadratic method.
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function RHEIN2_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal tolerance As Double = 0.000000000000001)

  Dim i As Long
  
  Dim DTEMP_VAL As Double
  Dim UTEMP_VAL As Double
  Dim VTEMP_VAL As Double
  Dim WTEMP_VAL As Double

  Dim TEMP_PS As Double
  Dim TEMP_PI As Double
  Dim TEMP_QI As Double
  Dim TEMP_QS As Double
  
  Dim TEMP_SWAP As Double
  Dim TEMP_GRAD As Double
  
  Dim TEMP_FUNC As Double
  Dim FIRST_FUNC As Double
  Dim SECOND_FUNC As Double
  
  Dim TEMP_STEP As Double
  
  Dim FORCE_FLAG As Boolean
  
  Dim TEMP_POINT As Double
  Dim TEMP_DELTA As Double

  On Error GoTo ERROR_LABEL
  
'-------------------------------------------------------------------------
' RHEIN2 implements the Rheinboldt bisection - secant -
' inverse quadratic method.
'
'  Parameters:
'
'    Input/output, real X, X1.
'    On input, two distinct points that start the method.
'    On output, X is an approximation to a root of the equation
'    which satisfies abs ( F(X) ) < ABSERR, and X1 is the previous
'    estimate.
'
'  Reference:
'
'    W C Rheinboldt,
'    Algorithms for Finding Zeros of a Function,
'    UMAP Journal,
'    Volume 2, 1, 1981, pages 43 - 72.
'--------------------------------------------------------------------------
  
'  Initialization.

  i = 0
  CONVERG_VAL = 0
  COUNTER = 0
  TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)

  If (Abs(TEMP_FUNC) > Abs(FIRST_FUNC)) Then
    TEMP_SWAP = LOWER_VAL
    LOWER_VAL = UPPER_VAL
    UPPER_VAL = TEMP_SWAP
    
    TEMP_SWAP = TEMP_FUNC
    TEMP_FUNC = FIRST_FUNC
    FIRST_FUNC = TEMP_SWAP
    
  End If

  TEMP_POINT = UPPER_VAL
  SECOND_FUNC = FIRST_FUNC
  COUNTER = 0
  TEMP_DELTA = 0.5 * Abs(LOWER_VAL - UPPER_VAL)
'
'  Iteration loop:
'
  Do
'
'  If the error tolerance is satisfied, then exit.
'
    If (Abs(TEMP_FUNC) <= 5 * tolerance) Then
      RHEIN2_ZERO_FUNC = LOWER_VAL
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      RHEIN2_ZERO_FUNC = LOWER_VAL
      Exit Function
    End If

    i = i + 1

    DTEMP_VAL = 0.5 * (UPPER_VAL - LOWER_VAL)
'
'  Compute the numerator and denominator for secant step.
'
    If (2 * Abs(TEMP_POINT - LOWER_VAL) < _
            Abs(UPPER_VAL - LOWER_VAL)) Then
      TEMP_PS = (LOWER_VAL - TEMP_POINT) * TEMP_FUNC
      TEMP_QS = SECOND_FUNC - TEMP_FUNC
    Else
      TEMP_PS = (LOWER_VAL - UPPER_VAL) * TEMP_FUNC
      TEMP_QS = FIRST_FUNC - TEMP_FUNC
    End If

    If (TEMP_PS < 0) Then
      TEMP_PS = -TEMP_PS
      TEMP_QS = -TEMP_QS
    End If
'
'  Compute the numerator and denominator for inverse quadratic.
'
    TEMP_PI = 0
    TEMP_QI = 0

    If (UPPER_VAL <> TEMP_POINT) Then
      UTEMP_VAL = TEMP_FUNC / SECOND_FUNC
      VTEMP_VAL = SECOND_FUNC / FIRST_FUNC
      WTEMP_VAL = TEMP_FUNC / FIRST_FUNC
      
      TEMP_PI = UTEMP_VAL * (2 * DTEMP_VAL * VTEMP_VAL * (VTEMP_VAL - WTEMP_VAL) - (LOWER_VAL - _
            TEMP_POINT) * (WTEMP_VAL - 1))
      TEMP_QI = (UTEMP_VAL - 1) * (VTEMP_VAL - 1) * (WTEMP_VAL - 1)

      If (TEMP_PI > 0#) Then: TEMP_QI = -TEMP_QI

      TEMP_PI = Abs(TEMP_PI)
    End If
'
'  Save the old minimum residual point.
'
    TEMP_POINT = LOWER_VAL
    SECOND_FUNC = TEMP_FUNC

    TEMP_STEP = (Abs(LOWER_VAL) + Abs(DTEMP_VAL) + 1) * tolerance
'
'  Choose bisection, secant or inverse quadratic step.
'
    FORCE_FLAG = False

    If (i > 3) Then
      If (8 * Abs(DTEMP_VAL) > TEMP_DELTA) Then
        FORCE_FLAG = True
      Else
        i = 0
        TEMP_DELTA = DTEMP_VAL
      End If
    End If
'
'  Set the increment.
'
    If (FORCE_FLAG) Then
      TEMP_GRAD = DTEMP_VAL
    ElseIf (TEMP_PI < 1.5 * DTEMP_VAL * TEMP_QI) And _
            (Abs(TEMP_PI) > Abs(TEMP_QI) * TEMP_STEP) Then
      TEMP_GRAD = TEMP_PI / TEMP_QI
    ElseIf (TEMP_PS < TEMP_QS * DTEMP_VAL) And (Abs(TEMP_PS) > _
            Abs(TEMP_QS) * TEMP_STEP) Then
      TEMP_GRAD = TEMP_PS / TEMP_QS
    Else
      TEMP_GRAD = DTEMP_VAL
    End If
'
'  Update the iterate and function values.
'
    LOWER_VAL = LOWER_VAL + TEMP_GRAD
    TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
'
'  Set the new UPPER_VAL as either UPPER_VAL or TEMP_POINT,
'  depending on whether F(UPPER_VAL) or F(TEMP_POINT) has the
'  opposite sign from F(LOWER_VAL).
    
    If IIf(TEMP_FUNC > 0, Abs(1), -Abs(1)) = _
       IIf(FIRST_FUNC > 0, Abs(1), -Abs(1)) Then
      UPPER_VAL = TEMP_POINT
      FIRST_FUNC = SECOND_FUNC
    End If

'  Force ABS ( TEMP_FUNC ) <= ABS ( FIRST_FUNC ).
'
    If (Abs(TEMP_FUNC) > Abs(FIRST_FUNC)) Then
      
        TEMP_SWAP = LOWER_VAL
        LOWER_VAL = UPPER_VAL
        UPPER_VAL = TEMP_SWAP
      
        TEMP_SWAP = TEMP_FUNC
        TEMP_FUNC = FIRST_FUNC
        FIRST_FUNC = TEMP_SWAP
    End If

  Loop

Exit Function
ERROR_LABEL:
    RHEIN2_ZERO_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CHEBY_FD_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the Chebyshev-Householder method with finite differences
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function CHEBY_FD_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal tolerance As Double = 0.000000000000001)
  
  Dim TEMP_POINT As Double
  Dim FIRST_DELTA As Double
  Dim SECOND_DELTA As Double
  Dim TEMP_GRAD As Double
  Dim TEMP_FUNC As Double
  Dim FIRST_FUNC As Double
  Dim SECOND_FUNC As Double

  On Error GoTo ERROR_LABEL
  
'--------------------------------------------------------------------------
' Cheby_FD carries out the Chebyshev-Householder method with finite
' differences.
'  Parameters:
'    Input/output, real X, X1, X2.
'    On input, three distinct points that start the method.
'    On output, X is an approximation to a root of the equation
'    which satisfies abs ( F(X) ) < ABSERR, and X1 and X2 are two previous
'    estimates.
'--------------------------------------------------------------------------

  CONVERG_VAL = 0
  COUNTER = 0
  
  TEMP_POINT = (LOWER_VAL + UPPER_VAL) / 2
  
  FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)
  FIRST_DELTA = (FIRST_FUNC - SECOND_FUNC) / (LOWER_VAL - UPPER_VAL)

  If (FIRST_DELTA = 0) Then
    CONVERG_VAL = 3
    CHEBY_FD_ZERO_FUNC = TEMP_POINT
    Exit Function
  End If

  TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)

'  Iteration loop:

  Do

'  If the error tolerance is satisfied, then exit.

    If (Abs(TEMP_FUNC) <= tolerance) Then
      CHEBY_FD_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      CHEBY_FD_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    SECOND_DELTA = FIRST_DELTA

    If (TEMP_POINT = LOWER_VAL) Then
      CONVERG_VAL = 3
      CHEBY_FD_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    FIRST_DELTA = (TEMP_FUNC - FIRST_FUNC) / (TEMP_POINT - LOWER_VAL)

    If (FIRST_DELTA = 0) Then
      CONVERG_VAL = 3
      CHEBY_FD_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    If (SECOND_FUNC = FIRST_FUNC) Then
      CONVERG_VAL = 3
      CHEBY_FD_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If
'
'  Set the increment.
'
    TEMP_GRAD = -TEMP_FUNC / FIRST_DELTA + (TEMP_FUNC * _
                FIRST_FUNC / (TEMP_FUNC - SECOND_FUNC)) * _
                (1 / FIRST_DELTA - 1 / SECOND_DELTA)
'
'  Remember current data for next step.
'
    UPPER_VAL = LOWER_VAL
    SECOND_FUNC = FIRST_FUNC
    LOWER_VAL = TEMP_POINT
    FIRST_FUNC = TEMP_FUNC
'
'  Update the iterate and function values.
'
    TEMP_POINT = TEMP_POINT + TEMP_GRAD
    TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)
  Loop

Exit Function
ERROR_LABEL:
    CHEBY_FD_ZERO_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : TRAUB_SE_21_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the Traub *E21 method
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function TRAUB_SE_21_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal tolerance As Double = 0.000000000000001)

'--------------------------------------------------------------------------
' STAR_E21 implements the Traub *E21 method.
'  Parameters:
'    Input/output, real X, X1, X2.
'    On input, three distinct points that start the method.
'    On output, X is an approximation to a root of the equation
'    which satisfies abs ( F(X) ) < ABSERR, and X1 and X2 are the
'    previous estimates.
'  Reference:
'    Joseph Traub,
'    Iterative Methods for the Solution of Equations,
'    Prentice Hall, 1964, page 234.
'--------------------------------------------------------------------------

  Dim TEMP_FUNC  As Double
  Dim TEMP_POINT As Double
  Dim FIRST_FUNC As Double
  Dim SECOND_FUNC As Double
  Dim TEMP_DELTA As Double
  Dim FIRST_DELTA As Double
  Dim SECOND_DELTA As Double
  Dim TEMP_FACT As Double
  Dim TEMP_GRAD As Double

  On Error GoTo ERROR_LABEL

  CONVERG_VAL = 0
  COUNTER = 0
  TEMP_POINT = (LOWER_VAL + UPPER_VAL) / 2
  TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)
  FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)
  FIRST_DELTA = (FIRST_FUNC - SECOND_FUNC) / (LOWER_VAL - UPPER_VAL)
'
'  Iteration loop:
'
  Do
'
'  If the error tolerance is satisfied, then exit.
'
    If (Abs(TEMP_FUNC) <= tolerance) Then
      TRAUB_SE_21_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      TRAUB_SE_21_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    SECOND_DELTA = FIRST_DELTA

    If (TEMP_POINT = LOWER_VAL) Then
      CONVERG_VAL = 3
      TRAUB_SE_21_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    FIRST_DELTA = (TEMP_FUNC - FIRST_FUNC) / (TEMP_POINT - LOWER_VAL)

    If (TEMP_POINT = UPPER_VAL) Then
      CONVERG_VAL = 3
      TRAUB_SE_21_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    TEMP_DELTA = (TEMP_FUNC - SECOND_FUNC) / (TEMP_POINT - UPPER_VAL)
    TEMP_FACT = FIRST_DELTA + TEMP_DELTA - SECOND_DELTA
    
    If (TEMP_FACT = 0) Then
      CONVERG_VAL = 3
      TRAUB_SE_21_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If
'
'  Set the increment.
'
    TEMP_GRAD = -TEMP_FUNC / TEMP_FACT
'
'  Remember current data for next step.
'
    UPPER_VAL = LOWER_VAL
    SECOND_FUNC = FIRST_FUNC

    LOWER_VAL = TEMP_POINT
    FIRST_FUNC = TEMP_FUNC
'
'  Update the iterate and function values.
'
    TEMP_POINT = TEMP_POINT + TEMP_GRAD
    TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)
  Loop

Exit Function
ERROR_LABEL:
    TRAUB_SE_21_ZERO_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : STEFF_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the Steffenson's method
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function STEFF_ZERO_FUNC(ByVal GUESS_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal tolerance As Double = 0.000000000000001)
'-------------------------------------------------------------------------------
' STEFFENSON implements Steffenson's method.
'  Parameters:
'    Input/output, real X.
'    On input, the point that starts the method.
'    On output, X is an approximation to a root of the equation
'    which satisfies abs ( F(X) ) < ABSERR.
'  Reference:
'    Joseph Traub,
'    Iterative Methods for the Solution of Equations,
'    Prentice Hall, 1964, page 178.
'-------------------------------------------------------------------------------

Dim TEMP_FUNC As Double
Dim TEMP_GRAD As Double
Dim TEMP_DERIV As Double
Dim TEMP_VAL As Double
Dim TEMP_POINT As Double

On Error GoTo ERROR_LABEL
  
  TEMP_POINT = GUESS_VAL
  CONVERG_VAL = 0
  COUNTER = 0

  TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)

'  Iteration loop:

  Do
'
'  If the error tolerance is satisfied, then exit.
'
    If (Abs(TEMP_FUNC) <= tolerance) Then
      STEFF_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      STEFF_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If
    TEMP_VAL = TEMP_POINT + TEMP_FUNC
    TEMP_GRAD = (Excel.Application.Run(FUNC_NAME_STR, _
                TEMP_VAL) - TEMP_FUNC) / TEMP_FUNC

    If (TEMP_GRAD = 0) Then
      CONVERG_VAL = 3
      STEFF_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If
'
'  Set the increment.
'
    TEMP_DERIV = -TEMP_FUNC / TEMP_GRAD
'
'  Update the iterate and function values.
'
    TEMP_POINT = TEMP_POINT + TEMP_DERIV
    TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)

  Loop

Exit Function
ERROR_LABEL:
    STEFF_ZERO_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HALLEY_FD_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the Halley's method, with finite differences
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 018
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function HALLEY_FD_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal tolerance As Double = 0.000000000000001)

'-------------------------------------------------------------------------------
' HALLEY_FD implements Halley's method, with finite differences.
'  Parameters:
'    Input/output, real X, X1, X2.
'    On input, three distinct points that start the method.
'    On output, X is an approximation to a root of the equation
'    which satisfies abs ( F(X) ) < ABSERR, and X1 and X2 are the
'    previous estimates.
'-------------------------------------------------------------------------------

Dim TEMP_DERIV As Double
Dim TEMP_FUNC As Double
Dim FIRST_FUNC As Double
Dim SECOND_FUNC As Double
Dim FIRST_DELTA As Double
Dim SECOND_DELTA As Double
Dim TEMP_DELTA As Double
Dim TEMP_POINT As Double

On Error GoTo ERROR_LABEL
  
  CONVERG_VAL = 0
  COUNTER = 0
  TEMP_POINT = (LOWER_VAL + UPPER_VAL) / 2
  TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)
  FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)

  FIRST_DELTA = (FIRST_FUNC - SECOND_FUNC) / (LOWER_VAL - UPPER_VAL)
'
'  Iteration loop:
'
  Do
'
'  If the error tolerance is satisfied, then exit.
'
    If (Abs(TEMP_FUNC) <= tolerance) Then
      HALLEY_FD_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      HALLEY_FD_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    SECOND_DELTA = FIRST_DELTA

    If (TEMP_POINT = LOWER_VAL) Then
      CONVERG_VAL = 3
      HALLEY_FD_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    FIRST_DELTA = (TEMP_FUNC - FIRST_FUNC) / (TEMP_POINT - LOWER_VAL)

    If (TEMP_POINT = UPPER_VAL) Then
      CONVERG_VAL = 3
      HALLEY_FD_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    SECOND_DELTA = (FIRST_DELTA - SECOND_DELTA) / (TEMP_POINT - UPPER_VAL)

    If (FIRST_DELTA = 0) Then
      CONVERG_VAL = 3
      HALLEY_FD_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    TEMP_DELTA = FIRST_DELTA - FIRST_FUNC * SECOND_DELTA / FIRST_DELTA

    If (TEMP_DELTA = 0) Then
      CONVERG_VAL = 3
      HALLEY_FD_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If
'
'  Set the increment.
'
    TEMP_DERIV = -TEMP_FUNC / TEMP_DELTA
'
'  Remember current data for next step.
'
    UPPER_VAL = LOWER_VAL
    SECOND_FUNC = FIRST_FUNC
    LOWER_VAL = TEMP_POINT
    FIRST_FUNC = TEMP_FUNC
'
'  Update the iterate and function values.
'
    TEMP_POINT = TEMP_POINT + TEMP_DERIV
    TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)

  Loop

Exit Function
ERROR_LABEL:
    HALLEY_FD_ZERO_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PARAB_INV_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the parabola inverse interpolation method
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 019
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PARAB_INV_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal tolerance As Double = 0.000000000000001)
'-------------------------------------------------------------------------------
' parabola_inv implements the parabola inverse interpolation method.
'  Parameters:
'    Input/output, real X, X1, X2.
'    On input, three distinct points that start the method.
'    On output, X is an approximation to a root of the equation
'    which satisfies abs ( F(X) ) < ABSERR, and X1 and X2 are the
'    previous estimates.
'-------------------------------------------------------------------------------

Dim TEMP_VAL As Double
Dim FIRST_FUNC As Double
Dim SECOND_FUNC As Double

Dim TEMP_DELTA As Double
Dim TEMP_POINT As Double

On Error GoTo ERROR_LABEL
  
  CONVERG_VAL = 0
  COUNTER = 0
  TEMP_POINT = (LOWER_VAL + UPPER_VAL) / 2
  TEMP_VAL = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)
  FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)
  
'
'  Iteration loop:
'
  Do
'
'  If the error tolerance is satisfied, then exit.
'
    If (Abs(TEMP_VAL) <= tolerance) Then
      PARAB_INV_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      PARAB_INV_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    If Abs(TEMP_VAL - FIRST_FUNC) <= tolerance Then
      CONVERG_VAL = 3
      PARAB_INV_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If
'
'  Set the increment.

TEMP_DELTA = (FIRST_FUNC * (UPPER_VAL * TEMP_VAL - _
            TEMP_POINT * SECOND_FUNC) / (SECOND_FUNC - _
            TEMP_VAL) - TEMP_VAL * (LOWER_VAL * SECOND_FUNC - _
            UPPER_VAL * FIRST_FUNC) / (FIRST_FUNC - _
            SECOND_FUNC)) / (TEMP_VAL - FIRST_FUNC)
'
'  Remember current data for next step.
'
    LOWER_VAL = UPPER_VAL
    FIRST_FUNC = SECOND_FUNC

    UPPER_VAL = TEMP_POINT
    SECOND_FUNC = TEMP_VAL
'
'  Update the iterate and function values.
'
    TEMP_POINT = TEMP_DELTA
    TEMP_VAL = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)


  Loop

Exit Function
ERROR_LABEL:
PARAB_INV_ZERO_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PARAB_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the vertical parabola interpolation method
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 020
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PARAB_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 100, _
Optional ByVal tolerance As Double = 0.000000000000001)

'-------------------------------------------------------------------------------
' parabola implements the vertical parabola interpolation method.
'  Parameters:
'    Input/output, real X, X1, X2
'    On input, three distinct points that start the method.
'    On output, X is an approximation to a root of the equation
'    which satisfies abs ( F(X) ) < ABSERR, and X1 and X2 are the
'    previous estimates.
'-------------------------------------------------------------------------------


Dim ATEMP_VAL As Double
Dim BTEMP_VAL As Double
Dim CTEMP_VAL As Double

Dim TEMP_DELTA As Double
Dim TEMP_DERIV As Double

Dim FIRST_DELTA As Double
Dim SECOND_DELTA As Double

Dim TEMP_FUNC As Double
Dim FIRST_FUNC As Double
Dim SECOND_FUNC As Double

Dim LOWER_POINT As Double
Dim UPPER_POINT As Double
Dim DELTA_POINT As Double

Dim TEMP_POINT As Double

On Error GoTo ERROR_LABEL

'
'  Initialization.
'
  CONVERG_VAL = 0
  COUNTER = 0
  DELTA_POINT = LOWER_VAL
  LOWER_VAL = UPPER_VAL
  UPPER_VAL = (LOWER_VAL + DELTA_POINT) / 2
  SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)
  FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, DELTA_POINT)
'
'  Iteration loop:
'
  Do
'
'  If the error tolerance is satisfied, then exit.
'
    If (Abs(SECOND_FUNC) <= tolerance) Then
        TEMP_POINT = UPPER_VAL
        PARAB_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If
    
    If (Abs(UPPER_VAL - LOWER_VAL) <= tolerance) Then
        TEMP_POINT = UPPER_VAL
        PARAB_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If
    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
        PARAB_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    ATEMP_VAL = LOWER_VAL - DELTA_POINT
    BTEMP_VAL = LOWER_VAL - UPPER_VAL
    CTEMP_VAL = DELTA_POINT - UPPER_VAL
    
    SECOND_DELTA = (FIRST_FUNC - SECOND_FUNC) / BTEMP_VAL
    FIRST_DELTA = (TEMP_FUNC - SECOND_FUNC) / CTEMP_VAL
    
    LOWER_POINT = (SECOND_DELTA - FIRST_DELTA) / ATEMP_VAL
    UPPER_POINT = (BTEMP_VAL * FIRST_DELTA - _
                   CTEMP_VAL * SECOND_DELTA) / ATEMP_VAL
    
    TEMP_DELTA = UPPER_POINT ^ 2 - 4 * LOWER_POINT * SECOND_FUNC
    
    If TEMP_DELTA < 0 Then TEMP_DELTA = 0
    TEMP_DERIV = -2 * SECOND_FUNC / (UPPER_POINT + _
                    Sgn(UPPER_POINT) * Sqr(TEMP_DELTA))

'  Remember current data for next step.

    DELTA_POINT = LOWER_VAL
    TEMP_FUNC = FIRST_FUNC
    LOWER_VAL = UPPER_VAL
    FIRST_FUNC = SECOND_FUNC
'
'  Update the iterate and function values.
'
    UPPER_VAL = UPPER_VAL + TEMP_DERIV
    SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)

  Loop
    
Exit Function
ERROR_LABEL:
    PARAB_ZERO_FUNC = PUB_EPSILON
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : REGULA_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the Regula Falsi method
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 021
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function REGULA_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal tolerance As Double = 0.000000000000001)

'-------------------------------------------------------------------------------------
' REGULA implements the Regula Falsi method.
'  Parameters:
'    Input/output, real X, X1.
'    On input, two distinct points that start the method.
'    On output, X is an approximation to a root of the equation
'    which satisfies abs ( F(X) ) < ABSERR, and X1 is the previous
'    estimate.
'-------------------------------------------------------------------------------------

Dim FIRST_FUNC As Double
Dim SECOND_FUNC As Double
Dim TEMP_FUNC As Double

Dim TEMP_GRAD As Double
Dim TEMP_POINT As Double
Dim TEMP_VAL As Double

On Error GoTo ERROR_LABEL

'  Initialization
  CONVERG_VAL = 0
  COUNTER = 0

  FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)

  If (FIRST_FUNC < 0#) Then
      TEMP_VAL = LOWER_VAL
      LOWER_VAL = UPPER_VAL
      UPPER_VAL = TEMP_VAL
    
      TEMP_VAL = FIRST_FUNC
      FIRST_FUNC = SECOND_FUNC
      SECOND_FUNC = TEMP_VAL
  End If

TEMP_GRAD = 1
  
  Do
'  If the error tolerance is satisfied, then exit.
    If (Abs(FIRST_FUNC) <= tolerance) Then
        LOWER_VAL = TEMP_POINT
        REGULA_ZERO_FUNC = LOWER_VAL
      Exit Function
    End If
    
    If (Abs(TEMP_GRAD) <= tolerance) Then
        LOWER_VAL = TEMP_POINT
        REGULA_ZERO_FUNC = LOWER_VAL
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
        REGULA_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If
'  Set the increment.
    TEMP_GRAD = -SECOND_FUNC * (LOWER_VAL - _
                UPPER_VAL) / (FIRST_FUNC - SECOND_FUNC)
'  Update the iterate and function values.
    TEMP_POINT = UPPER_VAL + TEMP_GRAD
    TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)

    If (TEMP_FUNC >= 0#) Then
      LOWER_VAL = TEMP_POINT
      FIRST_FUNC = TEMP_FUNC
    Else
      UPPER_VAL = TEMP_POINT
      SECOND_FUNC = TEMP_FUNC
    End If

  Loop

Exit Function
ERROR_LABEL:
    REGULA_ZERO_FUNC = PUB_EPSILON
End Function



'************************************************************************************
'************************************************************************************
'FUNCTION      : SECANT_BACK_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the safe secant method with back-step
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 022
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function SECANT_BACK_ZERO_FUNC(ByVal LOWER_BOUND As Double, _
ByVal UPPER_BOUND As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal tolerance As Double = 0.000000000000001)


'-------------------------------------------------------------------------------------
' SECANT implements the safe secant method with back-step
'  Parameters:
'    Input/output, real X, a, b.
'     where F(a) * F(b) < 0
'-------------------------------------------------------------------------------------

Dim DELTA_FUNC As Double
Dim FIRST_FUNC As Double
Dim SECOND_FUNC As Double

Dim TEMP_GRAD As Double

Dim DELTA_POINT As Double
Dim LOWER_VAL As Double
Dim UPPER_VAL As Double
Dim TEMP_POINT As Double

Dim FUNC_ERR As Double

On Error GoTo ERROR_LABEL

'  Initialization.
'
  CONVERG_VAL = 0
  COUNTER = 0
  LOWER_VAL = LOWER_BOUND
  UPPER_VAL = UPPER_BOUND
  
  FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)

'  Iteration loop:

  Do
'  If the error tolerance is satisfied, then exit.
    
    If Abs(FIRST_FUNC) < Abs(SECOND_FUNC) Then
       FUNC_ERR = Abs(FIRST_FUNC)
    Else
       FUNC_ERR = Abs(SECOND_FUNC)
    End If
    
    If (FUNC_ERR <= tolerance) Then
        TEMP_POINT = UPPER_VAL
        SECANT_BACK_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
        SECANT_BACK_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    If ((FIRST_FUNC - SECOND_FUNC) = 0#) Then
      CONVERG_VAL = 3
        SECANT_BACK_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If
'
'  Set the increment.
'
    TEMP_GRAD = -SECOND_FUNC * (UPPER_VAL - _
                LOWER_VAL) / (SECOND_FUNC - FIRST_FUNC)
    TEMP_POINT = UPPER_VAL + TEMP_GRAD
'
    If LOWER_BOUND < TEMP_POINT And TEMP_POINT < UPPER_BOUND Then
        'accept the point TEMP_POINT
        DELTA_POINT = LOWER_VAL
        LOWER_VAL = UPPER_VAL
        UPPER_VAL = TEMP_POINT
        DELTA_FUNC = FIRST_FUNC
        FIRST_FUNC = SECOND_FUNC
        SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)
    Else
        'reject the point TEMP_POINT
        If DELTA_POINT = LOWER_VAL Then
            CONVERG_VAL = 4
            SECANT_BACK_ZERO_FUNC = TEMP_POINT
            Exit Function
        End If
        LOWER_VAL = DELTA_POINT
        FIRST_FUNC = DELTA_FUNC
    End If
  Loop

Exit Function
ERROR_LABEL:
    SECANT_BACK_ZERO_FUNC = PUB_EPSILON
End Function

'// PERFECT

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : RIDDER_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the Ridder's method
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 023
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function RIDDER_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal tolerance As Double = 0.000000000000001)

'-------------------------------------------------------------------------------------
'  RIDDER implements the Ridder's method.
'  Parameters:
'    Input/output, real X, X1, X2.
'    On input, three distinct points that start the method.
'    On output, X is an approximation to a root of the equation
'    which satisfies abs ( F(X) ) < ABSERR, and X1 and X2 are the
'    previous estimates.
'    Input, real ABSERR, an error tolerance.
'    Input, real external F, the name of the routine that evaluates the
'    function or its derivatives, of the form
'      function f ( x, ider )
'    Output, integer IERROR, error indicator.
'    0, no error occurred.
'    nonzero, an error occurred, and the iteration was halted.
'    Output, integer K, the number of steps taken.
'    Input, integer KMAX, the maximum number of iterations allowed.
'-------------------------------------------------------------------------------------

Dim TEMP_FUNC As Double
Dim FIRST_FUNC As Double
Dim SECOND_FUNC As Double

Dim TEMP_GRAD As Double
Dim DELTA_VAL As Double
Dim TEMP_POINT As Double

On Error GoTo ERROR_LABEL

  CONVERG_VAL = 0
  COUNTER = 0
  TEMP_POINT = (LOWER_VAL + UPPER_VAL) / 2
  TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)
  FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)
  
'  Iteration loop:
Do
'  If the error tolerance is satisfied, then exit.
    If (Abs(TEMP_FUNC) <= tolerance) Then
      RIDDER_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      RIDDER_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    If Abs(TEMP_POINT - LOWER_VAL) <= tolerance Then
      CONVERG_VAL = 3
      RIDDER_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If
'
'  Set the increment.
    DELTA_VAL = TEMP_FUNC ^ 2 - FIRST_FUNC * SECOND_FUNC
    If DELTA_VAL < 0 Then DELTA_VAL = 1
'
    TEMP_GRAD = (TEMP_POINT - LOWER_VAL) * _
                Sgn(FIRST_FUNC - SECOND_FUNC) * _
                TEMP_FUNC / Sqr(DELTA_VAL)
'
'  Remember current data for next step.
'
    LOWER_VAL = UPPER_VAL
    FIRST_FUNC = SECOND_FUNC

    UPPER_VAL = TEMP_POINT
    SECOND_FUNC = TEMP_FUNC
'
'  Update the iterate and function values.
'
    TEMP_POINT = TEMP_POINT + TEMP_GRAD
    TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)
  
  Loop

Exit Function
ERROR_LABEL:
    RIDDER_ZERO_FUNC = PUB_EPSILON
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FRACTION_ZERO_FUNC
'DESCRIPTION   : Calculates the values necessary to achieve a specific goal -
'implements the linear fraction interpolation method
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_ZERO
'ID            : 024
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function FRACTION_ZERO_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef CONVERG_VAL As Integer, _
Optional ByRef COUNTER As Long, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal tolerance As Double = 0.000000000000001)

Dim TEMP_FUNC As Double
Dim FIRST_FUNC As Double
Dim SECOND_FUNC As Double

Dim DELTA_VAL As Double
Dim TEMP_GRAD As Double
Dim TEMP_POINT As Double

'-------------------------------------------------------------------------------------
' fraction implements the linear fraction interpolation method.
'  Parameters:
'    Input/output, real X, X1, X2.
'    On input, three distinct points that start the method.
'    On output, X is an approximation to a root of the equation
'    which satisfies abs ( F(X) ) < ABSERR, and X1 and X2 are the
'    previous estimates.
'-------------------------------------------------------------------------------------

On Error GoTo ERROR_LABEL

'  Initialization.
'
  CONVERG_VAL = 0
  COUNTER = 0
  TEMP_POINT = (LOWER_VAL + UPPER_VAL) / 2
  TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)
  FIRST_FUNC = Excel.Application.Run(FUNC_NAME_STR, LOWER_VAL)
  SECOND_FUNC = Excel.Application.Run(FUNC_NAME_STR, UPPER_VAL)
'
'  Iteration loop:
'
  Do
'
'  If the error tolerance is satisfied, then exit.
'
    If (Abs(TEMP_FUNC) <= tolerance) Then
      FRACTION_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If

    COUNTER = COUNTER + 1

    If (COUNTER > nLOOPS) Then
      CONVERG_VAL = 2
      FRACTION_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If
    
    DELTA_VAL = (FIRST_FUNC - TEMP_FUNC) / _
        (LOWER_VAL - TEMP_POINT) * _
        SECOND_FUNC - (SECOND_FUNC - TEMP_FUNC) / _
        (UPPER_VAL - TEMP_POINT) * FIRST_FUNC

    If Abs(DELTA_VAL) <= tolerance Then
      CONVERG_VAL = 3
      FRACTION_ZERO_FUNC = TEMP_POINT
      Exit Function
    End If
'
'  Set the increment.

    TEMP_GRAD = (FIRST_FUNC - SECOND_FUNC) * TEMP_FUNC / DELTA_VAL
'
'  Remember current data for next step.
'
    LOWER_VAL = UPPER_VAL
    FIRST_FUNC = SECOND_FUNC

    UPPER_VAL = TEMP_POINT
    SECOND_FUNC = TEMP_FUNC
'
'  Update the iterate and function values.
'
    TEMP_POINT = TEMP_POINT + TEMP_GRAD
    TEMP_FUNC = Excel.Application.Run(FUNC_NAME_STR, TEMP_POINT)
  Loop
  
Exit Function
ERROR_LABEL:
FRACTION_ZERO_FUNC = PUB_EPSILON
End Function

'----------------------------------------------------------------------------------
'The following function have:
'   Input, real tolerance, an error tolerance.
'   Output, integer CONVERG_VAL, error indicator.
'   0, no error occurred.
'   nonzero, an error occurred, and the iteration was halted.
'   Output, integer COUNTER, the number of steps taken.
'   Input, integer nLOOPS, the maximum number of iterations allowed.
'---------------------------------------------------------------------------------
