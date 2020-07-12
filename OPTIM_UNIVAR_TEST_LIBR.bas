Attribute VB_Name = "OPTIM_UNIVAR_TEST_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_TEST_ZERO_FRAME_FUNC
'DESCRIPTION   : Testing Univariate Optimization functions
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_TEST
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function CALL_TEST_ZERO_FRAME_FUNC(ByVal LOWER_VAL As Double, _
ByVal UPPER_VAL As Double, _
ByVal FUNC_NAME_STR As String, _
Optional ByVal nLOOPS As Long = 500, _
Optional ByVal tolerance As Double = 10 ^ -15)

Dim XTEMP_VAL As Double
Dim COUNTER As Long
Dim CONVERG_VAL As Integer
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

ReDim TEMP_MATRIX(0 To 22, 1 To 6)

TEMP_MATRIX(0, 1) = "ALGORITHM"
TEMP_MATRIX(0, 2) = "X_VAL"
TEMP_MATRIX(0, 3) = "Y_VAL"
TEMP_MATRIX(0, 4) = "GRADIENT FD APPROX"
TEMP_MATRIX(0, 5) = "COUNTER"
TEMP_MATRIX(0, 6) = "CONVERG_VAL"

'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(1, 1) = "BRENT"
XTEMP_VAL = BRENT_ZERO_FUNC(LOWER_VAL, UPPER_VAL, FUNC_NAME_STR, _
            UPPER_VAL, CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(1, 2) = XTEMP_VAL
    TEMP_MATRIX(1, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(1, 2))
    TEMP_MATRIX(1, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(1, 2))
    TEMP_MATRIX(1, 5) = COUNTER
Else
    TEMP_MATRIX(1, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(2, 1) = "NEWTON"
XTEMP_VAL = NEWTON_ZERO_FUNC(UPPER_VAL, _
                    FUNC_NAME_STR, "", CONVERG_VAL, _
                    COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(2, 2) = XTEMP_VAL
    TEMP_MATRIX(2, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(2, 2))
    TEMP_MATRIX(2, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(2, 2))
    TEMP_MATRIX(2, 5) = COUNTER
Else
    TEMP_MATRIX(2, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(3, 1) = "WEB"
XTEMP_VAL = WDB_ZERO_FUNC(LOWER_VAL, UPPER_VAL, _
            FUNC_NAME_STR, CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then

    TEMP_MATRIX(3, 2) = XTEMP_VAL
    TEMP_MATRIX(3, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(3, 2))
    TEMP_MATRIX(3, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(3, 2))
    TEMP_MATRIX(3, 5) = COUNTER
Else
    TEMP_MATRIX(3, 6) = CONVERG_VAL
End If
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(4, 1) = "BISEC"
XTEMP_VAL = BISEC_ZERO_FUNC(LOWER_VAL, UPPER_VAL, _
            FUNC_NAME_STR, CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(4, 2) = XTEMP_VAL
    TEMP_MATRIX(4, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(4, 2))
    TEMP_MATRIX(4, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(4, 2))
    TEMP_MATRIX(4, 5) = COUNTER
Else
    TEMP_MATRIX(4, 6) = CONVERG_VAL
End If
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(5, 1) = "SECANT"
XTEMP_VAL = SECANT_ZERO_FUNC(LOWER_VAL, UPPER_VAL, _
            FUNC_NAME_STR, CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(5, 2) = XTEMP_VAL
    TEMP_MATRIX(5, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(5, 2))
    TEMP_MATRIX(5, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(5, 2))
    TEMP_MATRIX(5, 5) = COUNTER
Else
    TEMP_MATRIX(5, 6) = CONVERG_VAL
End If
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(6, 1) = "PEGASO"
XTEMP_VAL = PEGASO_ZERO_FUNC(LOWER_VAL, UPPER_VAL, _
            FUNC_NAME_STR, CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(6, 2) = XTEMP_VAL
    TEMP_MATRIX(6, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(6, 2))
    TEMP_MATRIX(6, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(6, 2))
    TEMP_MATRIX(6, 5) = COUNTER
Else
    TEMP_MATRIX(6, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(7, 1) = "PHI-21"
XTEMP_VAL = TRAUB_PHI_21_ZERO_FUNC(LOWER_VAL, UPPER_VAL, _
            FUNC_NAME_STR, CONVERG_VAL, COUNTER, nLOOPS, tolerance)

If CONVERG_VAL = 0 Then
    TEMP_MATRIX(7, 2) = XTEMP_VAL
    TEMP_MATRIX(7, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(7, 2))
    TEMP_MATRIX(7, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(7, 2))
    TEMP_MATRIX(7, 5) = COUNTER
Else
    TEMP_MATRIX(7, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------
'
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(8, 1) = "E-21"
XTEMP_VAL = TRAUB_E_21_ZERO_FUNC(LOWER_VAL, UPPER_VAL, _
            FUNC_NAME_STR, CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(8, 2) = XTEMP_VAL
    TEMP_MATRIX(8, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(8, 2))
    TEMP_MATRIX(8, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(8, 2))
    TEMP_MATRIX(8, 5) = COUNTER
Else
    TEMP_MATRIX(8, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(9, 1) = "MULLER"
XTEMP_VAL = MULLER_ZERO_FUNC(LOWER_VAL, UPPER_VAL, _
            FUNC_NAME_STR, CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(9, 2) = XTEMP_VAL
    TEMP_MATRIX(9, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(9, 2))
    TEMP_MATRIX(9, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(9, 2))
    TEMP_MATRIX(9, 5) = COUNTER
Else
    TEMP_MATRIX(9, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(10, 1) = "RHEIN1"
XTEMP_VAL = RHEIN1_ZERO_FUNC(LOWER_VAL, UPPER_VAL, _
            FUNC_NAME_STR, CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(10, 2) = XTEMP_VAL
    TEMP_MATRIX(10, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(10, 2))
    TEMP_MATRIX(10, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(10, 2))
    TEMP_MATRIX(10, 5) = COUNTER
Else
    TEMP_MATRIX(10, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(11, 1) = "RHEIN2"
XTEMP_VAL = RHEIN2_ZERO_FUNC(LOWER_VAL, UPPER_VAL, _
            FUNC_NAME_STR, CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(11, 2) = XTEMP_VAL
    TEMP_MATRIX(11, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(11, 2))
    TEMP_MATRIX(11, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(11, 2))
    TEMP_MATRIX(11, 5) = COUNTER
Else
    TEMP_MATRIX(11, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(12, 1) = "CHEBY-FD"
XTEMP_VAL = CHEBY_FD_ZERO_FUNC(LOWER_VAL, UPPER_VAL, _
            FUNC_NAME_STR, CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(12, 2) = XTEMP_VAL
    TEMP_MATRIX(12, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(12, 2))
    TEMP_MATRIX(12, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(12, 2))
    TEMP_MATRIX(12, 5) = COUNTER
Else
    TEMP_MATRIX(12, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------
'

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(13, 1) = "*E-21"
XTEMP_VAL = TRAUB_SE_21_ZERO_FUNC(LOWER_VAL, UPPER_VAL, _
            FUNC_NAME_STR, CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(13, 2) = XTEMP_VAL
    TEMP_MATRIX(13, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(13, 2))
    TEMP_MATRIX(13, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(13, 2))
    TEMP_MATRIX(13, 5) = COUNTER
Else
    TEMP_MATRIX(13, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(14, 1) = "STEFF"
XTEMP_VAL = STEFF_ZERO_FUNC(UPPER_VAL, FUNC_NAME_STR, _
            CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(14, 2) = XTEMP_VAL
    TEMP_MATRIX(14, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(14, 2))
    TEMP_MATRIX(14, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(14, 2))
    TEMP_MATRIX(14, 5) = COUNTER
Else
    TEMP_MATRIX(14, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(15, 1) = "HALL"
XTEMP_VAL = HALLEY_FD_ZERO_FUNC(LOWER_VAL, UPPER_VAL, FUNC_NAME_STR, _
            CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(15, 2) = XTEMP_VAL
    TEMP_MATRIX(15, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(15, 2))
    TEMP_MATRIX(15, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(15, 2))
    TEMP_MATRIX(15, 5) = COUNTER
Else
    TEMP_MATRIX(15, 6) = CONVERG_VAL
End If


'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(16, 1) = "PARAB-INV"
XTEMP_VAL = PARAB_INV_ZERO_FUNC(LOWER_VAL, UPPER_VAL, FUNC_NAME_STR, _
            CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(16, 2) = XTEMP_VAL
    TEMP_MATRIX(16, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(16, 2))
    TEMP_MATRIX(16, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(16, 2))
    TEMP_MATRIX(16, 5) = COUNTER
Else
    TEMP_MATRIX(16, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(17, 1) = "PARAB"
XTEMP_VAL = PARAB_ZERO_FUNC(LOWER_VAL, UPPER_VAL, FUNC_NAME_STR, _
            CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(17, 2) = XTEMP_VAL
    TEMP_MATRIX(17, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(17, 2))
    TEMP_MATRIX(17, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(17, 2))
    TEMP_MATRIX(17, 5) = COUNTER
Else
    TEMP_MATRIX(17, 6) = CONVERG_VAL
End If


'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(18, 1) = "REGULA"
XTEMP_VAL = REGULA_ZERO_FUNC(LOWER_VAL, UPPER_VAL, FUNC_NAME_STR, _
            CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(18, 2) = XTEMP_VAL
    TEMP_MATRIX(18, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(18, 2))
    TEMP_MATRIX(18, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(18, 2))
    TEMP_MATRIX(18, 5) = COUNTER
Else
    TEMP_MATRIX(18, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(19, 1) = "SECANT-BACK"
XTEMP_VAL = SECANT_BACK_ZERO_FUNC(LOWER_VAL, UPPER_VAL, FUNC_NAME_STR, _
            CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(19, 2) = XTEMP_VAL
    TEMP_MATRIX(19, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(19, 2))
    TEMP_MATRIX(19, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(19, 2))
    TEMP_MATRIX(19, 5) = COUNTER
Else
    TEMP_MATRIX(19, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(20, 1) = "RIDDER"
XTEMP_VAL = RIDDER_ZERO_FUNC(LOWER_VAL, UPPER_VAL, FUNC_NAME_STR, _
            CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(20, 2) = XTEMP_VAL
    TEMP_MATRIX(20, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(20, 2))
    TEMP_MATRIX(20, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(20, 2))
    TEMP_MATRIX(20, 5) = COUNTER
Else
    TEMP_MATRIX(20, 6) = CONVERG_VAL
End If


'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(21, 1) = "FRACTION"
XTEMP_VAL = FRACTION_ZERO_FUNC(LOWER_VAL, UPPER_VAL, FUNC_NAME_STR, _
            CONVERG_VAL, COUNTER, nLOOPS, tolerance)
If CONVERG_VAL = 0 Then
    TEMP_MATRIX(21, 2) = XTEMP_VAL
    TEMP_MATRIX(21, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(21, 2))
    TEMP_MATRIX(21, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(21, 2))
    TEMP_MATRIX(21, 5) = COUNTER
Else
    TEMP_MATRIX(21, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
COUNTER = 0
TEMP_MATRIX(22, 1) = "QUADRATIC"
XTEMP_VAL = QUADRATIC_ZERO_FUNC(LOWER_VAL, UPPER_VAL, _
            FUNC_NAME_STR, CONVERG_VAL, COUNTER, nLOOPS, tolerance)

If CONVERG_VAL = 0 Then
    TEMP_MATRIX(22, 2) = XTEMP_VAL
    TEMP_MATRIX(22, 3) = Excel.Application.Run(FUNC_NAME_STR, _
                        TEMP_MATRIX(22, 2))
    TEMP_MATRIX(22, 4) = UNIVAR_FD_GRAD_APPROX_FUNC(FUNC_NAME_STR, _
                        TEMP_MATRIX(22, 2))
    TEMP_MATRIX(22, 5) = COUNTER
Else
    TEMP_MATRIX(22, 6) = CONVERG_VAL
End If

'-------------------------------------------------------------------------------



CALL_TEST_ZERO_FRAME_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CALL_TEST_ZERO_FRAME_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_ZERO_OBJ_1_FUNC
'DESCRIPTION   : X^0.5-Sin(x)-1
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_TEST
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Private Function CALL_ZERO_OBJ_1_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL

CALL_ZERO_OBJ_1_FUNC = X_VAL ^ 0.5 - Sin(X_VAL) - 1

Exit Function
ERROR_LABEL:
CALL_ZERO_OBJ_1_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************



'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_ZERO_OBJ_2_FUNC
'DESCRIPTION   :e^(-4x) + e^(-x-3)-x^6
'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_TEST
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Private Function CALL_ZERO_OBJ_2_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
CALL_ZERO_OBJ_2_FUNC = Exp(-4 * X_VAL) + _
                          Exp(-1 * (X_VAL + 3)) - _
                          X_VAL ^ 6

Exit Function
ERROR_LABEL:
CALL_ZERO_OBJ_2_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_UNIVAR_OBJ_1_FUNC

'DESCRIPTION   : Smooth function: The search for an extreme in a uni-variate smooth
'function is quite simple and almost all algorithms usually work. We have only to
'plot the function to locate immediately the extreme.

'Assume for example, a problem of finding a local maximum and minimum of the
'following function in the range 0 < x < 10

'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_TEST
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function CALL_UNIVAR_OBJ_1_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL
CALL_UNIVAR_OBJ_1_FUNC = Sin(X_VAL) / (1 + X_VAL ^ 2) ^ 0.5 'This function has three
'local extreme points within the range of 0 to 10.

'Interval   -- Extreme
'0 < x < 2  -- local max that is also the absolute max
'2 < x < 6  -- local min that is also the absolute min
'6 < x < 10 -- local max

'In order to approximate the extreme points we can use the parabolic interpolation
'macro. This algorithm converges to the extreme within each specified interval, no
'matter if it is a maximum or a minimum.

'-----------------------------------------------------------------------------------
'[a b]      x          f(x)
'[0 2]  1.109293102  0.599522164
'[2 6]  4.503864793 -0.21205764
'[6 10] 7.727383943  0.127312641
'-----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
CALL_UNIVAR_OBJ_1_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_UNIVAR_OBJ_2_FUNC

'DESCRIPTION   : Many local minima
'An optimization algorithm may give a wrong result when there are too many local
'extremes (minimum and maximum) near the desired absolute maximum or minimum.
'In that case they can be trapped into one of the local extremes.

'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_TEST
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function CALL_UNIVAR_OBJ_2_FUNC(ByVal X_VAL As Double)

'In the range 0 < x < 5 there are many local minimum and maximum points. We could
'bracket the absolute maximum within a small interval before starting the searching
'algorithm. But we want to show the evidence that the Divide-and-Conquer algorithm,
'(thanks to its intrinsically robustness) can escape the local "traps" and give
'the correct absolute max and min

On Error GoTo ERROR_LABEL
'For example assume to have to find the maximum and minimum of the following
'function within the range, 0 < x < 5
CALL_UNIVAR_OBJ_2_FUNC = X_VAL * Exp(-X_VAL) * Cos(6 * X_VAL)

'----------------------------------------------------------------------------------
'    [a b]    x     f(x)
'----------------------------------------------------------------------------------
'max [0 5]  1.0459   0.367
'min [0 5]  1.5608  -0.327
'----------------------------------------------------------------------------------

'As we can see the algorithm ignores the other local extremes and converges to the
'true absolute maximum. But of course this is a didactic extreme case. Generally
'speaking, it is always better to isolate the desired extreme within a sufficiently close
'segment before attempting to find the absolute maximum (minimum). If it's impossible
'and there are many local extremes, you may increase the number of points limit from
'600 (default) to 1000 or 2000.

Exit Function
ERROR_LABEL:
CALL_UNIVAR_OBJ_2_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_UNIVAR_OBJ_3_FUNC

'DESCRIPTION   : The Saw Ramp
'This example illustrates a case quite difficult for
'many optimization algorithms, even if it is quite
'simple to find the maximum and minimum by
'inspection of a plot of the function.


'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_TEST
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function CALL_UNIVAR_OBJ_3_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL

'Assume we have to search the maximum and
'minimum for a function shown in the plot, in the
'Range 0 < x < 4

CALL_UNIVAR_OBJ_3_FUNC = Abs(X_VAL) + 4 * Abs(Int(X_VAL + 0.5) - X_VAL) + 1

'The optimization macro will converge to the point (0, 1) for the minimum and to (3.5,
'6.5) for the maximum

Exit Function
ERROR_LABEL:
CALL_UNIVAR_OBJ_3_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_UNIVAR_OBJ_4_FUNC
'DESCRIPTION   : Stiff Function
'This example shows another difficult function. The concept of stiff functions is similar
'to the differential equation problem: We are speaking of stiff problems when the
'function evolves smoothly and slowly in a large interval except in one or more small
'intervals where the evolution is more rapid. Usually this kind of function needs two or
'more plots with different scales.

'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_TEST
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function CALL_UNIVAR_OBJ_4_FUNC(ByVal X_VAL As Double)
On Error GoTo ERROR_LABEL

'Given the following function for x >= 0 , find the absolute max and min

'-----------------------------------------------------------------------------------
'One plot covers the wide range 0 < x < 100 and another covers the smaller region
'0 < x < 10 where the function shows a narrow maximum. (Note that there are other
'local extremes within the global interval.) The absolute minimum is located within the
'interval 10 < x < 30#
'-----------------------------------------------------------------------------------
'For finding the maximum and minimum with the best accuracy, we can use the divideand-
'conquer algorithm (robust convergence)
'-----------------------------------------------------------------------------------

CALL_UNIVAR_OBJ_4_FUNC = Sin(X_VAL ^ 0.5) / (1 + X_VAL ^ 2) ^ 0.5

'For finding the maximum and minimum with the best accuracy, we can use the divideand-
'conquer algorithm (robust convergence), obtaining the following result:
'-----------------------------------------------------------------------------------
'[a b]        x              f(x)
'[0 100]  18.2845204  -- -0.0494926
'[0 100] 0.760360454  --  0.6094426
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
'Note that the parabolic algorithm has some convergence difficulty in finding the
'maximum near 0, if the interval is not sufficiently close to zero. On the contrary,
'there is no problem for the minimum
'-----------------------------------------------------------------------------------

'[a b]           x          f(x)

'[0 3]      6.869565509   0.07165213
'[0 2]     -1.668589413    #NUM!

'[0.5 1.5]  0.760360472   0.6094426
'[10   30]  18.28452054  -0.0494926

'This behavior for the ranges near zero, can be explained by observing that the
'function does not have a derivative value at x = 0
'-----------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
CALL_UNIVAR_OBJ_4_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_UNIVAR_OBJ_5_FUNC_A
'DESCRIPTION   : The Orbits
'Two satellites follow two plane elliptic orbits described by the following parametric
'equations with respect to the earth.

'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_TEST
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function CALL_UNIVAR_OBJ_5_FUNC_A(ByVal X_VAL As Double, _
Optional ByVal VERSION As Integer = 0)

On Error GoTo ERROR_LABEL

'For time t = 0 the two satellites stay at
'positions (2, -1), (1, 0) respectively

Select Case VERSION
'------------------------------------------------------------------------------------
    Case 0 'X Sat 1
        CALL_UNIVAR_OBJ_5_FUNC_A = 2 * Cos(X_VAL) + 3 * Sin(X_VAL)
'------------------------------------------------------------------------------------
    Case 1 'X Sat 2
        CALL_UNIVAR_OBJ_5_FUNC_A = Cos(X_VAL) + 2 * Sin(X_VAL)
'------------------------------------------------------------------------------------
    Case 2 'Y Sat 1
        CALL_UNIVAR_OBJ_5_FUNC_A = 4 * Sin(X_VAL) - Cos(X_VAL)
'------------------------------------------------------------------------------------
    Case Else 'Y Sat 2
        CALL_UNIVAR_OBJ_5_FUNC_A = Sin(X_VAL)
'------------------------------------------------------------------------------------
End Select

'We want to find when the two satellites have a minimum distance from each other (In
'order to transmit messages with the lowest noise possible). We want to also find the
'position of each satellite at the minimum distance. (Note that, in general, this position
'does not coincide with the static minimum distance between the orbits.)
'This problem can be regarded as a minimization problem having one parameter (the
'time "t") and one objective function (the distance "d")

Exit Function
ERROR_LABEL:
CALL_UNIVAR_OBJ_5_FUNC_A = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_UNIVAR_OBJ_5_FUNC_B

'DESCRIPTION   : This problem can be regarded as a minimization problem having one parameter (the
'time "t") and one objective function (the distance "d")
'The distance on a plane between two points is

'LIBRARY       : OPTIMIZATION
'GROUP         : UNIVAR_TEST
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function CALL_UNIVAR_OBJ_5_FUNC_B(ByVal X_VAL As Double)

On Error GoTo ERROR_LABEL

'---------------------------------------------------------------------------------------------
't min     t max    time error
'---------------------------------------------------------------------------------------------
'0 0.3 0.231823805  7.074E-12
'3 3.5 3.373416458  1.934E-12
'The final values are accurate better than 1E-11
'---------------------------------------------------------------------------------------------

CALL_UNIVAR_OBJ_5_FUNC_B = ( _
(CALL_UNIVAR_OBJ_5_FUNC_A(X_VAL, 0) - CALL_UNIVAR_OBJ_5_FUNC_A(X_VAL, 1)) ^ 2 + _
(CALL_UNIVAR_OBJ_5_FUNC_A(X_VAL, 2) - CALL_UNIVAR_OBJ_5_FUNC_A(X_VAL, 3)) ^ 2 _
) ^ 0.5

'First of all we note that both orbits are periodic of the same period T = ƒÎ . 6.28
'So we can study the problem for 0 < t < 6.28

'Note that when you change the parameter "t" , for example giving a set of sequential
'values (0, 1, 2, 3, 4, 5) we get immediately the orbit coordinates

'We observe that the condition of "minimum distance" happens two times: in the
'intervals (0, 3) and (3, 6).

'---------------------------------------------------------------------------------------------
'Starting the macro "1D-divide and Conquer" with the following constrain conditions:
'(tmin = 0, tmax = 3) and (tmin = 3, tmax = 6.28), returns the following values.
'Constraints SAT 1 SAT 2
'---------------------------------------------------------------------------------------------
't min --> t max --> time     --> distance --> x        -->        y -->        x --> y
'---------------------------------------------------------------------------------------------
'0     --> 3     --> 0.231824 --> 1.236068 --> 2.635757 --> -0.05424 --> 1.432755 --> 0.229753
'3     --> 6.28  --> 3.373416 --> 1.236068 --> -2.63576 --> 0.054237 --> -1.43275 --> -0.22975
'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------

'Improving accuracy
'The optimum values of the parameter "t" was calculated with a good accuracy of about
'1E-8. If we want to improve the accuracy we may try to use the parabolic algorithm.
'However with the parabolic algorithm on this application, we have to pay attention to
'bracketing the minimum in a narrow interval about the desired minimum point. For
'example we can use the segment 0 < t < 0.3 for the first minimum and 3 < t < 3.5 for
'the second one.
'---------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
CALL_UNIVAR_OBJ_5_FUNC_B = Err.number
End Function
