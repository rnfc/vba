Attribute VB_Name = "OPTIM_MULTVAR_TEST_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private PUB_FUNC_NAME_STR As String


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_TEST_MULTVAR_FRAME_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_TEST_MULTVAR_FRAME_FUNC(ByVal FUNC_STR_NAME As String, _
ByRef PARAM_RNG As Variant, _
ByRef CONST_RNG As Variant, _
Optional ByVal GRAD_STR_NAME As String = "", _
Optional ByVal CONST_STR_NAME As String = "", _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByVal nLOOPS As Long = 600, _
Optional ByVal epsilon As Double = 0.000000000000001)

Dim i As Long
Dim NSIZE As Long

Dim CONST_BOX As Variant
Dim TEMP_MATRIX As Variant
Dim XTEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PUB_FUNC_NAME_STR = FUNC_STR_NAME
CONST_BOX = CONST_RNG
PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)
NSIZE = UBound(PARAM_VECTOR, 1)
'--------------------------------------------------------------------------------
ReDim TEMP_MATRIX(0 To 8, 0 To NSIZE + 1)
'--------------------------------------------------------------------------------
TEMP_MATRIX(0, 0) = "--"
TEMP_MATRIX(0, NSIZE + 1) = "Y_VAR"
For i = 1 To NSIZE
    TEMP_MATRIX(0, i) = "X_VAR_" & i
Next i
    
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
TEMP_MATRIX(1, 0) = "NELDER-MEAD"
XTEMP_VECTOR = NELDER_MEAD_OPTIMIZATION_FRAME_FUNC(FUNC_STR_NAME, PARAM_VECTOR, CONST_BOX, MIN_FLAG, 0, nLOOPS, epsilon)
For i = 1 To NSIZE
    TEMP_MATRIX(1, i) = XTEMP_VECTOR(i, 1)
Next i
'--------------------------------------------------------------------------------
TEMP_MATRIX(1, NSIZE + 1) = MULTVAR_CALL_OBJ_FUNC(FUNC_STR_NAME, _
                            XTEMP_VECTOR, "", 1) 'Thanks to David Lifchitz
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
TEMP_MATRIX(2, 0) = "RND"
XTEMP_VECTOR = MULTVAR_RESIZE_OPTIM_FUNC(FUNC_STR_NAME, _
            CONST_BOX, GRAD_STR_NAME, PARAM_VECTOR, _
            MIN_FLAG, 0, nLOOPS, epsilon)
For i = 1 To NSIZE
    TEMP_MATRIX(2, i) = XTEMP_VECTOR(i, 1)
Next i
'--------------------------------------------------------------------------------
TEMP_MATRIX(2, NSIZE + 1) = MULTVAR_CALL_OBJ_FUNC(FUNC_STR_NAME, _
                            XTEMP_VECTOR, "", 1) 'Thanks to David Lifchitz
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
TEMP_MATRIX(3, 0) = "GRADIENT"
XTEMP_VECTOR = MULTVAR_RESIZE_OPTIM_FUNC(FUNC_STR_NAME, _
            CONST_BOX, GRAD_STR_NAME, PARAM_VECTOR, _
            MIN_FLAG, 1, nLOOPS, epsilon)
For i = 1 To NSIZE
    TEMP_MATRIX(3, i) = XTEMP_VECTOR(i, 1)
Next i
'--------------------------------------------------------------------------------
TEMP_MATRIX(3, NSIZE + 1) = MULTVAR_CALL_OBJ_FUNC(FUNC_STR_NAME, _
                            XTEMP_VECTOR, "", 1) 'Thanks to David Lifchitz
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
TEMP_MATRIX(4, 0) = "CONJUGATE-GRADIENT"
XTEMP_VECTOR = MULTVAR_RESIZE_OPTIM_FUNC(FUNC_STR_NAME, _
            CONST_BOX, GRAD_STR_NAME, PARAM_VECTOR, _
            MIN_FLAG, 2, nLOOPS, epsilon)
For i = 1 To NSIZE
    TEMP_MATRIX(4, i) = XTEMP_VECTOR(i, 1)
Next i
'--------------------------------------------------------------------------------
TEMP_MATRIX(4, NSIZE + 1) = MULTVAR_CALL_OBJ_FUNC(FUNC_STR_NAME, _
                            XTEMP_VECTOR, "", 1) 'Thanks to David Lifchitz
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
TEMP_MATRIX(5, 0) = "DFP"
XTEMP_VECTOR = MULTVAR_RESIZE_OPTIM_FUNC(FUNC_STR_NAME, _
            CONST_BOX, GRAD_STR_NAME, PARAM_VECTOR, _
            MIN_FLAG, 3, nLOOPS, epsilon)
For i = 1 To NSIZE
    TEMP_MATRIX(5, i) = XTEMP_VECTOR(i, 1)
Next i
'--------------------------------------------------------------------------------
TEMP_MATRIX(5, NSIZE + 1) = MULTVAR_CALL_OBJ_FUNC(FUNC_STR_NAME, _
                            XTEMP_VECTOR, "", 1) 'Thanks to David Lifchitz
'--------------------------------------------------------------------------------


'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
TEMP_MATRIX(6, 0) = "RND + DFP"
XTEMP_VECTOR = MULTVAR_RESIZE_OPTIM_FUNC(FUNC_STR_NAME, _
            CONST_BOX, GRAD_STR_NAME, PARAM_VECTOR, _
            MIN_FLAG, 0, nLOOPS, epsilon)

XTEMP_VECTOR = MULTVAR_RESIZE_OPTIM_FUNC(FUNC_STR_NAME, _
            CONST_BOX, GRAD_STR_NAME, XTEMP_VECTOR, _
            MIN_FLAG, 3, nLOOPS, epsilon)

For i = 1 To NSIZE
    TEMP_MATRIX(6, i) = XTEMP_VECTOR(i, 1)
Next i
'--------------------------------------------------------------------------------
TEMP_MATRIX(6, NSIZE + 1) = MULTVAR_CALL_OBJ_FUNC(FUNC_STR_NAME, _
                            XTEMP_VECTOR, "", 1) 'Thanks to David Lifchitz
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
TEMP_MATRIX(7, 0) = "SIMPLEX"
If CONST_STR_NAME <> "" Then
    Dim TEMP_FACTOR As Variant
    TEMP_FACTOR = -1 'Maximum
    If MIN_FLAG = True Then: TEMP_FACTOR = 1
    XTEMP_VECTOR = SIMPLEX_MINIMUM_OPTIMIZATION_FUNC(FUNC_STR_NAME, _
                CONST_STR_NAME, PARAM_VECTOR, 0.01, 100000, epsilon)
    If IsArray(XTEMP_VECTOR) = False Then: GoTo 1983
    For i = 1 To NSIZE
        XTEMP_VECTOR(i, 1) = XTEMP_VECTOR(i, 1) * TEMP_FACTOR
        TEMP_MATRIX(7, i) = XTEMP_VECTOR(i, 1)
    Next i
'--------------------------------------------------------------------------------
    TEMP_MATRIX(7, NSIZE + 1) = MULTVAR_CALL_OBJ_FUNC(FUNC_STR_NAME, _
                            XTEMP_VECTOR, "", 1) 'Thanks to David Lifchitz
'--------------------------------------------------------------------------------
End If
'--------------------------------------------------------------------------------
1983:

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
TEMP_MATRIX(8, 0) = "PIKAIA"
XTEMP_VECTOR = PIKAIA_OPTIMIZATION_FUNC(FUNC_STR_NAME, _
                            CONST_BOX, False)
For i = 1 To NSIZE
    XTEMP_VECTOR(i, 1) = XTEMP_VECTOR(i, 1) * IIf(MIN_FLAG = True, -1, 1)
    TEMP_MATRIX(8, i) = XTEMP_VECTOR(i, 1)
Next i
'--------------------------------------------------------------------------------
TEMP_MATRIX(8, NSIZE + 1) = MULTVAR_CALL_OBJ_FUNC(FUNC_STR_NAME, _
                            XTEMP_VECTOR, "", 1) 'Thanks to David Lifchitz
'--------------------------------------------------------------------------------



CALL_TEST_MULTVAR_FRAME_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CALL_TEST_MULTVAR_FRAME_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_JACOBI_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_JACOBI_FUNC(ByRef PARAM_RNG As Variant)

Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PUB_FUNC_NAME_STR = "CALL_MULTVAR_OBJ_1_FUNC"

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

'CALL_MULTVAR_JACOBI_FUNC = MATRIX_TRANSPOSE_FUNC(JACOBI_CENTRAL_FUNC(PUB_FUNC_NAME_STR, _
                          PARAM_VECTOR, 0.00001))

CALL_MULTVAR_JACOBI_FUNC = MATRIX_TRANSPOSE_FUNC(JACOBI_FORWARD_FUNC(PUB_FUNC_NAME_STR, _
                          PARAM_VECTOR, 0.00001))
Exit Function
ERROR_LABEL:
CALL_MULTVAR_JACOBI_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_OBJ_1_FUNC
'DESCRIPTION   : Peak and Pit minimum function
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_OBJ_1_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

CALL_MULTVAR_OBJ_1_FUNC = (X_VAL + Y_VAL) / _
    (X_VAL ^ 4 - 2 * X_VAL ^ 2 + Y_VAL ^ 4 - 2 * Y_VAL ^ 2 + 3)

Exit Function
ERROR_LABEL:
CALL_MULTVAR_OBJ_1_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_CONST_1_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_CONST_1_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim TEMP_FLAG As Boolean
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = True
PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

If X_VAL < -2.5 Then: TEMP_FLAG = False
If X_VAL > 2.5 Then: TEMP_FLAG = False
If Y_VAL < -2.5 Then: TEMP_FLAG = False
If Y_VAL > 2.5 Then: TEMP_FLAG = False

'In the plot we clearly observe the presence of a maximum and a minimum in
'the domain -2.5 < x < 2.5 , -2.5 < y < 2.5

'CONST_BOX(1, 1) = -2.5 'X-Min
'CONST_BOX(1, 2) = -2.5 'Y-Min

'CONST_BOX(2, 1) = 2.5 'X-Max
'CONST_BOX(2, 2) = 2.5 'Y-Max

'The maximum is located in the area { x, y | x>0 , y>0 } and the minimum is located in
'the area { x, y | x<0 , y<0 }. The point (0, 0) is at the middle of the maximum and
'minimum points so we can use it as starting point for both searches.

CALL_MULTVAR_CONST_1_FUNC = TEMP_FLAG

Exit Function
ERROR_LABEL:
CALL_MULTVAR_CONST_1_FUNC = TEMP_FLAG
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_OBJ_2_FUNC
'DESCRIPTION   : Parabolic surface
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_OBJ_2_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)
CALL_MULTVAR_OBJ_2_FUNC = 2 * X_VAL ^ 2 + 4 * X_VAL * Y_VAL _
                                - 12 * X_VAL + 8 * Y_VAL ^ 2 - 36 * Y_VAL + 43

'We see that the minimum is located in the region 0 < x < 2, 0 < y < 4.
'Because the gradient is simple, we can also insert the derivative function

Exit Function
ERROR_LABEL:
CALL_MULTVAR_OBJ_2_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_GRAD_2_FUNC
'DESCRIPTION   :
'As we can see, for smooth functions like polynomials, the exact
'derivatives are useful, since with the Newton-Raphson (NR) final step,
'the global accuracy of the solution can be improved.

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_GRAD_2_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim TEMP_ARR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG

X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

ReDim TEMP_ARR(1 To 2, 1 To 1)

TEMP_ARR(1, 1) = 4 * X_VAL + 4 * Y_VAL - 12
TEMP_ARR(2, 1) = 4 * X_VAL + 16 * Y_VAL - 36

CALL_MULTVAR_GRAD_2_FUNC = TEMP_ARR

Exit Function
ERROR_LABEL:
CALL_MULTVAR_GRAD_2_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_CONST_2_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_CONST_2_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim TEMP_FLAG As Boolean
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = True
PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

If X_VAL < 0 Then: TEMP_FLAG = False
If X_VAL > 2 Then: TEMP_FLAG = False
If Y_VAL < 0 Then: TEMP_FLAG = False
If Y_VAL > 4 Then: TEMP_FLAG = False

CALL_MULTVAR_CONST_2_FUNC = TEMP_FLAG

Exit Function
ERROR_LABEL:
CALL_MULTVAR_CONST_2_FUNC = TEMP_FLAG
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_OBJ_3_FUNC

'DESCRIPTION   : Super parabolic surface minimum function
'This example shows another case in which the optimization algorithms that do not
'require external derivative equations are sometimes superior to those that
'require external derivative equations.

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_OBJ_3_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

CALL_MULTVAR_OBJ_3_FUNC = X_VAL ^ 2 - _
    (3 / 5) * X_VAL + _
    Y_VAL ^ 4 - _
    2 * Y_VAL ^ 3 + _
    (3 / 2) * Y_VAL ^ 2 - _
    (1 / 2) * Y_VAL + _
    (61 / 100)
    
'Note that the last root of the second gradient has a multiplicity of 3.
'This means that it is a root also for the 2nd derivative df/dy.


Exit Function
ERROR_LABEL:
CALL_MULTVAR_OBJ_3_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_GRAD_3_FUNC
'DESCRIPTION   : Surprisingly the methods that are derivative-free like Random
'and Simplex, show the best results.

'This happens because they are not affected by the cancellation of the 2nd
'derivatives. On the contrary the other method, also with NR refinement step,
'cannot reduce the error less then 1E-3. Note that the bigger error happens
'over the y variable. This is not strange because, as we have demonstrated,
'the y variable annihilates its 2nd derivatives at the point y = 0.5.

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_GRAD_3_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim TEMP_ARR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG

X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

ReDim TEMP_ARR(1 To 2, 1 To 1)

TEMP_ARR(1, 1) = 2 * X_VAL - (3 / 5)

TEMP_ARR(2, 1) = 4 * Y_VAL ^ 3 - _
                 6 * Y_VAL ^ 2 + _
                 3 * Y_VAL - (1 / 2)

CALL_MULTVAR_GRAD_3_FUNC = TEMP_ARR

Exit Function
ERROR_LABEL:
CALL_MULTVAR_GRAD_3_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_CONST_3_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_CONST_3_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim TEMP_FLAG As Boolean
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = True
PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

If X_VAL < 0 Then: TEMP_FLAG = False
If X_VAL > 1 Then: TEMP_FLAG = False
If Y_VAL < 0 Then: TEMP_FLAG = False
If Y_VAL > 1 Then: TEMP_FLAG = False

CALL_MULTVAR_CONST_3_FUNC = TEMP_FLAG

Exit Function
ERROR_LABEL:
CALL_MULTVAR_CONST_3_FUNC = TEMP_FLAG
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_OBJ_4_FUNC
'DESCRIPTION   : The trap
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_OBJ_4_FUNC(ByRef PARAM_RNG As Variant)

'As we can see, all algorithms, except one, fail to converge at the true
'minimum. They all fall into the false central minimum. Only the random
'algorithm has escaped from the "trap", giving the true minimum with a
'good accuracy (1E-5). Random algorithms are in general, suitable for
'finding a narrow global optimum where there are surrounding
'local optimums..

'convergence region

'It 's reasonable that for the other algorithms there will be some starting
'points, from which the algorithm will converge to the true minimum B. There
'will be other starting points that the algorithm will end up at the false
'minimum (0,0). The set of "good" starting points constitutes the convergence
'region.

'A larger convergence region means a robust algorithm.
'We want to investigate the convergence regions for each of these algorithms.
'We repeat the above minimum searching with many starting points inside the
'domain -2 < x < 2 and -2 < y < 2. For each trial we note if the algorithm has
'failed or not.

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)
CALL_MULTVAR_OBJ_4_FUNC = 1 - _
    0.5 * Exp(-1 * (X_VAL ^ 2 + Y_VAL ^ 2)) - _
    Exp(-10 * (X_VAL ^ 2 + Y_VAL ^ 2 - 2 * X_VAL - 2 * Y_VAL + 2))

'This interesting example shows how to avoid a situation that would trap the most
'sophisticated algorithms. This happens when there are one or more local extremes
'near the true absolute minimum.

'-------------------------------------------------------------------------------
'The random algorithm, of course has a convergence region coincident with
'the black square. As we can see, from the point of view of convergence,
'the most robust algorithm is the Random, followed by the Downhill-Simplex
'and then by the CG and DFP

'The Downhill-Simplex has a sufficiently large convergence region and is
'considered a robust algorithm. The CG, DFP and the NR algorithms have, on
'the contrary a poor global convergence characteristic.

'Mixed METHOD

'But of course we could use a "mix of algorithms" to reach the best results
'For example if we start with the random method, we can find a sufficiently
'accurate starting point for the DFP algorithm. Following this mixed method,
'we can find the optimum with a very high accuracy (2E-9), no matter what
'the starting point was.
'-------------------------------------------------------------------------------


Exit Function
ERROR_LABEL:
CALL_MULTVAR_OBJ_4_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_CONST_4_FUNC
'DESCRIPTION   : the contours-plot shows the presence of two extreme points: one
'in the center (0, 0), called A, and another one in a more narrow region near
'the point (1, 1), called B

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_CONST_4_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim TEMP_FLAG As Boolean
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = True
PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

'it 's interesting to also draw the 3D plot. We note the presence of the
'larger local minimum at the center point A (0, 0) and the true narrow
'minimum near the point B. If we choose a starting point like
'(-1,-1) the algorithm path would likely cross into the point (0, 0) region and
'will be trapped at this local minimum. On the contrary, if we start from the
'point (2, 2) it is reasonable to guess that we would find the true minimum.
'But what will happen if we start from a point like (0, 1) ? Let's see.
'We try to find the minimum with all the methods, starting from the point
'(0, 1), in the domains of - 2 < x < 2 And -2 < y < 2#

If X_VAL < -2 Then: TEMP_FLAG = False
If X_VAL > 2 Then: TEMP_FLAG = False
If Y_VAL < -2 Then: TEMP_FLAG = False
If Y_VAL > 2 Then: TEMP_FLAG = False

CALL_MULTVAR_CONST_4_FUNC = TEMP_FLAG

Exit Function
ERROR_LABEL:
CALL_MULTVAR_CONST_4_FUNC = TEMP_FLAG
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_OBJ_5_FUNC
'DESCRIPTION   : The eye
'Derivative discontinuity in general can give problems to those algorithms using
'gradient information. But this not always true.

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_OBJ_5_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

CALL_MULTVAR_OBJ_5_FUNC = Abs(X_VAL - 2) ^ 2 + Abs(Y_VAL - 1)

'The contour-plot takes on an "eye" pattern for the individual contours. The
'plot shows that the minimum is clearly the point (2, 1). Note from the 3D
'plot that the gradient in the minimum is not continuous.

Exit Function
ERROR_LABEL:
CALL_MULTVAR_OBJ_5_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_CONST_5_FUNC
'DESCRIPTION   : let's see how the algorithms works, in the domain box 0 < x < 4 and
'0 < y < 2 , starting from the point (0, 0) [0,0]

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_CONST_5_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim TEMP_FLAG As Boolean
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = True
PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

'The random algorithm, of course has a convergence region coincident with
'the black square. As we can see, from the point of view of convergence,
'the most robust algorithm is the Random, followed by the Downhill-Simplex
'and then by the CG and DFP

'The Downhill-Simplex has a sufficiently large convergence region and is
'considered a robust algorithm. The CG, DFP and the NR algorithms have, on
'the contrary a poor global convergence characteristic.

'Mixed METHOD

'But of course we could use a "mix of algorithms" to reach the best results
'For example if we start with the random method, we can find a sufficiently
'accurate starting point for the DFP algorithm. Following this mixed method,
'we can find the optimum with a very high accuracy (2E-9), no matter what
'the starting point was.

If X_VAL < 0 Then: TEMP_FLAG = False
If X_VAL > 4 Then: TEMP_FLAG = False
If Y_VAL < 0 Then: TEMP_FLAG = False
If Y_VAL > 2 Then: TEMP_FLAG = False

CALL_MULTVAR_CONST_5_FUNC = TEMP_FLAG

Exit Function
ERROR_LABEL:
CALL_MULTVAR_CONST_5_FUNC = TEMP_FLAG
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_OBJ_6_FUNC

'DESCRIPTION   : Four Hill: many times the function to be optimized is symmetric
'to one or both axes

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_OBJ_6_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

CALL_MULTVAR_OBJ_6_FUNC = 1 / _
(X_VAL ^ 4 + Y_VAL ^ 4 - 2 * X_VAL ^ 2 - 2 * Y_VAL ^ 2 + 3)

'Both variables appear only with even powers. So the function is symmetric to
'both x and y axes. This means that if the function has a maximum in the 1st
'region {x, y | x>0 , y>0 }, it will have also three other maximum extremes
'in all other regions.

'The optimization function cannot give in one pass, all four maximum points
'(within the designated region) so one of them is chosen randomly. To avoid
'this little indecision we must give the initial starting point nearer one
'of these points or, resizing the convergence region.

Exit Function
ERROR_LABEL:
CALL_MULTVAR_OBJ_6_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_CONST_6_FUNC

'DESCRIPTION   : No too clear? Never mind. Let's see the following plot in the
'symmetric region [2,2]

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_CONST_6_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim TEMP_FLAG As Boolean
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = True
PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

If X_VAL < 0 Then: TEMP_FLAG = False
If X_VAL > 2 Then: TEMP_FLAG = False
If Y_VAL < 0 Then: TEMP_FLAG = False
If Y_VAL > 2 Then: TEMP_FLAG = False

CALL_MULTVAR_CONST_6_FUNC = TEMP_FLAG

'It 's clear that the function has four symmetric maximums in every region of the
'selected interval. We can restrict our study to the 1st region 0 < x < 2 ,
'0 < y < 2. In this region, starting from a point like (2, 2) all algorithms
'work fine in reaching the true maximum extreme (1, 1) with good accuracy.

Exit Function
ERROR_LABEL:
CALL_MULTVAR_CONST_6_FUNC = TEMP_FLAG
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_OBJ_7_FUNC
'DESCRIPTION   : Rosenbrock's parabolic valley
'This family of test functions is well know to be a minimization problem of
'high difficulty.

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_OBJ_7_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

CALL_MULTVAR_OBJ_7_FUNC = 100 * (Y_VAL - X_VAL ^ 2) ^ 2 + (1 - X_VAL) ^ 2

'The parameter "100" changes the level of difficulty. A high m value
'means high difficulty in searching for a minimum. The reason is that
'the minimum is located in a large flat region with a very low slope.

'The function is always positive except at the point (1, 1) where it's 0.
'Taking the Gradient it 's simple to demonstrate the only extreme is at
'the point (1, 1), which is the absolute minimum of the
'function.

Exit Function
ERROR_LABEL:
CALL_MULTVAR_OBJ_7_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_CONST_7_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 018
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_CONST_7_FUNC(ByRef PARAM_RNG As Variant)

'We note a general loss of accuracy, because all algorithms seem to have
'difficulty in locating the exact minimum. They seem to get "stuck in the
'mud" of the valley. Also the random algorithm seems to have a greater
'difficulty in finding the minimum. The reason is that, when the random
'algorithm samples a quasi-flat area, all points have similar heights so
'it has difficulty in discovering where the true minimum is located.
'The only exception is the Downhill-Simplex algorithm. Its path, rolling
'into the valley, is both fast and accurate. Why? I have to admit that we
'cannot explain it... but it works!

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim TEMP_FLAG As Boolean
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = True
PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

If X_VAL < 0 Then: TEMP_FLAG = False
If X_VAL > 0 Then: TEMP_FLAG = False
If Y_VAL < 1 Then: TEMP_FLAG = False
If Y_VAL > 1 Then: TEMP_FLAG = False

CALL_MULTVAR_CONST_7_FUNC = TEMP_FLAG

Exit Function
ERROR_LABEL:
CALL_MULTVAR_CONST_7_FUNC = TEMP_FLAG
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_OBJ_8_FUNC

'DESCRIPTION   : nonlinear Regression with Absolute Sums
'This example explains how to perform a nonlinear regression with objective
'function different from the Least Squared. In this example we adopt the
'Absolutes Sum. We choose the exponential model:
'f (x,a,k) = a · e^(-k·x)

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 019
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_OBJ_8_FUNC(ByRef PARAM_RNG As Variant)

Dim i As Variant
Dim X_VAL As Double
Dim Y_VAL As Double
Dim TEMP_SUM As Double
Dim TEMP_FACTOR As Double
Dim PARAM_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG

X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

ReDim DATA_VECTOR(1 To 11, 1 To 2)

DATA_VECTOR(1, 1) = 0.1
DATA_VECTOR(1, 2) = 0.8187308

DATA_VECTOR(2, 1) = 0.2
DATA_VECTOR(2, 2) = 0.67032

DATA_VECTOR(3, 1) = 0.3
DATA_VECTOR(3, 2) = 0.5488116

DATA_VECTOR(4, 1) = 0.4
DATA_VECTOR(4, 2) = 0.449329

DATA_VECTOR(5, 1) = 0.5
DATA_VECTOR(5, 2) = 0.3678794

DATA_VECTOR(6, 1) = 0.6
DATA_VECTOR(6, 2) = 0.3011942

DATA_VECTOR(7, 1) = 0.7
DATA_VECTOR(7, 2) = 0.246597

DATA_VECTOR(8, 1) = 0.8
DATA_VECTOR(8, 2) = 0.2018965

DATA_VECTOR(9, 1) = 0.9
DATA_VECTOR(9, 2) = 0.1652989

DATA_VECTOR(10, 1) = 1
DATA_VECTOR(10, 2) = 0.1353353

DATA_VECTOR(11, 1) = 1.1
DATA_VECTOR(11, 2) = 0.1108032

TEMP_SUM = 0

For i = 1 To 11
    TEMP_FACTOR = (X_VAL * Exp(Y_VAL * DATA_VECTOR(i, 1)))
    TEMP_SUM = TEMP_SUM + Abs(DATA_VECTOR(i, 2) - TEMP_FACTOR)
Next i

CALL_MULTVAR_OBJ_8_FUNC = TEMP_SUM

'The goal of the regression is to find the best couple of PARAM_RNG values
'(a, k) that minimize the sum of the absolute value of the difference between
'the regression model and the given data set.
'AS = Sum| y - f (x,a,k) |

'The objective function AS depends only on the PARAM_RNG a and k. By minimizing AS,
'with our optimization algorithms, we hope to solve the regression problem.


'Starting from the point (1, 0) you will see the cells changing
'quickly until the optimization algorithm stops itself, leaving
'the following "best" fitting parameter values of the regression y*
'Best fitting PARAM_RNG {a k} {1 -2}

Exit Function
ERROR_LABEL:
CALL_MULTVAR_OBJ_8_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_CONST_8_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 020
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_CONST_8_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim TEMP_FLAG As Boolean
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = True
PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

If X_VAL < 0 Then: TEMP_FLAG = False
If X_VAL > 1 Then: TEMP_FLAG = False
If Y_VAL < -2 Then: TEMP_FLAG = False
If Y_VAL > 0 Then: TEMP_FLAG = False

CALL_MULTVAR_CONST_8_FUNC = TEMP_FLAG

Exit Function
ERROR_LABEL:
CALL_MULTVAR_CONST_8_FUNC = TEMP_FLAG
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_OBJ_9_FUNC
'DESCRIPTION   : the ground fault: assume the minimum of the following function
'is to be found
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 021
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_OBJ_9_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

CALL_MULTVAR_OBJ_9_FUNC = 2 - _
        (1 / (1 + (X_VAL - Y_VAL - 1) ^ 2)) - _
        (1 / (1 + (X_VAL + Y_VAL - 3) ^ 2))

'--------------------------------------------------------------------------
'Both plots indicate clearly a narrow minimum near the point (2, 1).
'Nevertheless this function may create some difficulty because the
'narrow minimum is hidden at the cross of two long valleys (like a
'ground fault).
'--------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
CALL_MULTVAR_OBJ_9_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_CONST_9_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 022
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_CONST_9_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim TEMP_FLAG As Boolean
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = True
PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

If X_VAL < -10 Then: TEMP_FLAG = False
If X_VAL > 10 Then: TEMP_FLAG = False
If Y_VAL < -10 Then: TEMP_FLAG = False
If Y_VAL > 10 Then: TEMP_FLAG = False

CALL_MULTVAR_CONST_9_FUNC = TEMP_FLAG


'shows an extremely flat valley, bordered at the corners with high walls.
'This plot is quite useless for locating the minimum

Exit Function
ERROR_LABEL:
CALL_MULTVAR_CONST_9_FUNC = TEMP_FLAG
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_OBJ_10_FUNC
'DESCRIPTION   : Brown bad scaled function: this function is often used as a
'benchmark for testing the scaling ability of algorithms

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 023
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_OBJ_10_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

CALL_MULTVAR_OBJ_10_FUNC = (X_VAL - 10 ^ 6) ^ 2 + _
(Y_VAL - 10 ^ -6) ^ 2 + (X_VAL * Y_VAL - 2) ^ 2

'This function is always positive and it is zero only at the point
'(10^6 , 2 10^-6 ). At this point, the abscissa is very high and the
'ordinate is very low. It is hard to generate good plots of this
'function. We also have no idea where the extremes are located. This
'situation, is not very common indeed, but if this happens, the only
'thing that we can do is to run the Downhill-Simplex algorithm trusting
'in its intrinsically robustness.

'Fortunately, in this case, the algorithm converges quickly to
'the exact minimun with a very high accuracy

Exit Function
ERROR_LABEL:
CALL_MULTVAR_OBJ_10_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_CONST_10_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 024
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_CONST_10_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim TEMP_FLAG As Boolean
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = True
PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

If X_VAL < 0 Then: TEMP_FLAG = False
If X_VAL > 1000000 Then: TEMP_FLAG = False
If Y_VAL < 0 Then: TEMP_FLAG = False
If Y_VAL > 10 Then: TEMP_FLAG = False

CALL_MULTVAR_CONST_10_FUNC = TEMP_FLAG


'shows an extremely flat valley, bordered at the corners with high walls.
'This plot is quite useless for locating the minimum

Exit Function
ERROR_LABEL:
CALL_MULTVAR_CONST_10_FUNC = TEMP_FLAG
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_OBJ_11_FUNC

'DESCRIPTION   : Beale function: another test function that is very difficult
'to study:

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 025
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************


Function CALL_MULTVAR_OBJ_11_FUNC(ByRef PARAM_RNG As Variant)

'As we can see, the final accuracy is a thousand times less then the previous one.
'Clearly the time spent for choosing a suitable starting point is useful (This is
'in general true, when it's possible).

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

CALL_MULTVAR_OBJ_11_FUNC = (1.5 - X_VAL * (1 - Y_VAL)) ^ 2 + _
    (2.25 - X_VAL * (1 - Y_VAL ^ 2)) ^ 2 + _
    (2.625 - X_VAL * (1 - Y_VAL ^ 3)) ^ 2

'It is always positive, being a sum of three square terms. So the minimum,
'if exists, must be positive or 0.

Exit Function
ERROR_LABEL:
CALL_MULTVAR_OBJ_11_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_CONST_11_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 026
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_CONST_11_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim TEMP_FLAG As Boolean
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = True
PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)

If X_VAL < -4 Then: TEMP_FLAG = False
If X_VAL > 4 Then: TEMP_FLAG = False
If Y_VAL < -4 Then: TEMP_FLAG = False
If Y_VAL > 4 Then: TEMP_FLAG = False

CALL_MULTVAR_CONST_11_FUNC = TEMP_FLAG

'shows an extremely flat valley, bordered at the corners with high walls.
'This plot is quite useless for locating the minimum

Exit Function
ERROR_LABEL:
CALL_MULTVAR_CONST_11_FUNC = TEMP_FLAG
End Function

'//////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////
'Examples of multivariate functions
'//////////////////////////////////////////////////////////////////////////////////////
'The searching of extremes of a multivariable function, apart from elementary
'examples, can be very difficult, This is because, in general, we cannot use
'the graphic method illustrated in the previous examples with one and two
'variable functions. Sometimes graphic methods may still be applied for
'particular kinds of functions.
'//////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_OBJ_12_FUNC

'DESCRIPTION   :
'Splitting function method: assume the maximum and the minimum of this function
'function is to be found

'Sometime the function can be split into parts, each having a separate set of
'variables. If each part contains no more than two variables, we can apply the
'graphic method for each part. This example explain this concept.


'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 027
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_OBJ_12_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim Z_VAL As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)
Z_VAL = PARAM_VECTOR(3, 1)

'First of all, we observe that the function has no maximum; so it could
'have only the minimum.

CALL_MULTVAR_OBJ_12_FUNC = X_VAL ^ 2 + 4 * Y_VAL ^ 2 + 2 * _
            Abs(Z_VAL) ^ (3 / 2) + _
            X_VAL * Abs(Y_VAL) ^ (1 / 2) + _
            X_VAL + Z_VAL

'It is always positive, being a sum of three square terms. So the minimum,
'if exists, must be positive or 0.

Exit Function
ERROR_LABEL:
CALL_MULTVAR_OBJ_12_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_CONST_12_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 028
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_CONST_12_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim Z_VAL As Variant
Dim TEMP_FLAG As Boolean
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = True
PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)
Z_VAL = PARAM_VECTOR(3, 1)

If X_VAL < -2 Then: TEMP_FLAG = False
If X_VAL > 0 Then: TEMP_FLAG = False
If Y_VAL < -1 Then: TEMP_FLAG = False
If Y_VAL > 1 Then: TEMP_FLAG = False
If Z_VAL < -0.2 Then: TEMP_FLAG = False
If Z_VAL > 0 Then: TEMP_FLAG = False

CALL_MULTVAR_CONST_12_FUNC = TEMP_FLAG

'From the first contours-plot we deduce that the minimum is located in the region of -2
'< x < 0 and -1 < y < 1. From the second plot we have the region -0.2 < z < 0. Now
'we have a constraints box for searching for the minimum of f(x,y,z).
'[-0.6 0.02 -0.1]

Exit Function
ERROR_LABEL:
CALL_MULTVAR_CONST_12_FUNC = TEMP_FLAG
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_OBJ_13_FUNC
'DESCRIPTION   : The gradient condition
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 029
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_OBJ_13_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim Z_VAL As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)
Z_VAL = PARAM_VECTOR(3, 1)

CALL_MULTVAR_OBJ_13_FUNC = (X_VAL - Z_VAL) ^ 2 + Y_VAL * _
(Y_VAL - X_VAL) + Z_VAL ^ 2


Exit Function
ERROR_LABEL:
CALL_MULTVAR_OBJ_13_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_GRAD_13_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 030
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_GRAD_13_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim Z_VAL As Variant
Dim TEMP_ARR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG

X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)
Z_VAL = PARAM_VECTOR(3, 1)

ReDim TEMP_ARR(1 To 3, 1 To 1)
'This function is top-unlimited

TEMP_ARR(1, 1) = 2 * X_VAL - Y_VAL - 2 * Z_VAL
TEMP_ARR(2, 1) = 2 * Y_VAL - X_VAL
TEMP_ARR(3, 1) = 4 * Z_VAL - 2 * X_VAL

CALL_MULTVAR_GRAD_13_FUNC = TEMP_ARR

Exit Function
ERROR_LABEL:
CALL_MULTVAR_GRAD_13_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_CONST_13_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 031
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************


Function CALL_MULTVAR_CONST_13_FUNC(ByRef PARAM_RNG As Variant)

'We have seen that this function has no upper limit. This is true if
'the variables are unconstrained. But surely the maximum exists if the
'variables are limited by a specific range. Assume now that each variable
'must be limited in the range [-2, 2].

'In this way the maximum surely will belong to the
'surface of the square box centered around the origin
'and having length of 4.

'But where in this box will the maximum be located?
'It may lie on a face, or on edge, or even at a corner
'of the box. Let's discover it

'[2,-1,-1]

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim Z_VAL As Variant
Dim TEMP_FLAG As Boolean
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = True
PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)
Z_VAL = PARAM_VECTOR(3, 1)

If X_VAL < -2 Then: TEMP_FLAG = False
If X_VAL > 2 Then: TEMP_FLAG = False
If Y_VAL < -2 Then: TEMP_FLAG = False
If Y_VAL > 2 Then: TEMP_FLAG = False
If Z_VAL < -2 Then: TEMP_FLAG = False
If Z_VAL > 2 Then: TEMP_FLAG = False

CALL_MULTVAR_CONST_13_FUNC = TEMP_FLAG


'-----------------------------------------------------------------------------
'We can restart the function searching for the max in the given box or we can
'also use the CG macro starting from any internal point like for example (1, 1, 1)
'We see that the max, f = 28, is located in the corner (-2, 2, 2)

'We have to observe that the function is symmetric with respect to the origin.
'f(x, y, z) = f(-x, -y, -z)

'So there must be another maximum point at the symmetrical point (2, -2, -2).
'To test for it, simply restart the CG macro, this time choosing the starting
'point (2, -1, -1). It will converge exactly to the second maximum point.

'-----------------------------------------------------------------------------
'We see that the only point for the minimum is (0, 0, 0). Starting from any
'point around the origin, every algorithm will converge to the origin.

Exit Function
ERROR_LABEL:
CALL_MULTVAR_CONST_13_FUNC = TEMP_FLAG
End Function



'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_OBJ_14_FUNC

'DESCRIPTION   :
'Production: this example shows how to tune the production of several
'products to maximize profit. The function model here is the Cobb-Douglas
'production function for three products

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 032
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_OBJ_14_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim Z_VAL As Variant
Dim PRICE_VAL As Variant
Dim COST_VAL As Variant
Dim PARAM_VECTOR As Variant

Dim tolerance As Variant

On Error GoTo ERROR_LABEL

tolerance = 0.000000000000001
PARAM_VECTOR = PARAM_RNG
'Where z, y, x are the quantities of each product (input)
X_VAL = IIf(PARAM_VECTOR(1, 1) < tolerance, 0, PARAM_VECTOR(1, 1))
Y_VAL = IIf(PARAM_VECTOR(2, 1) < tolerance, 0, PARAM_VECTOR(2, 1))
Z_VAL = IIf(PARAM_VECTOR(3, 1) < tolerance, 0, PARAM_VECTOR(3, 1))

PRICE_VAL = (X_VAL ^ 0.1) * (Y_VAL ^ 0.2) * (Z_VAL ^ 0.3)
'PRICE_VAL is an arbitrary unit-less measure of value of the ouput products.

COST_VAL = 0.3 * X_VAL + 0.1 * Y_VAL + 0.2 * Z_VAL + 2

CALL_MULTVAR_OBJ_14_FUNC = 2 * PRICE_VAL - COST_VAL
'The cost for each item plus the fixed cost and the sale cost factor

Exit Function
ERROR_LABEL:
CALL_MULTVAR_OBJ_14_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_CONST_14_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 033
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_CONST_14_FUNC(ByRef PARAM_RNG As Variant)
'[5,25,25]

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim Z_VAL As Variant
Dim TEMP_FLAG As Boolean
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = True
PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)
Z_VAL = PARAM_VECTOR(3, 1)

If X_VAL < 0 Then: TEMP_FLAG = False
If X_VAL > 10 Then: TEMP_FLAG = False
If Y_VAL < 0 Then: TEMP_FLAG = False
If Y_VAL > 50 Then: TEMP_FLAG = False
If Z_VAL < 0 Then: TEMP_FLAG = False
If Z_VAL > 50 Then: TEMP_FLAG = False

CALL_MULTVAR_CONST_14_FUNC = TEMP_FLAG

Exit Function
ERROR_LABEL:
CALL_MULTVAR_CONST_14_FUNC = TEMP_FLAG
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_OBJ_15_FUNC
'DESCRIPTION   : Parabolic 3D
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 034
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_OBJ_15_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim Z_VAL As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)
Z_VAL = PARAM_VECTOR(3, 1)

CALL_MULTVAR_OBJ_15_FUNC = 2 * X_VAL ^ 2 + _
            10 * Y_VAL ^ 2 + _
            5 * Z_VAL ^ 2 + _
            6 * X_VAL * Y_VAL - _
            2 * X_VAL * Z_VAL + _
            4 * Y_VAL * Z_VAL - _
            6 * X_VAL - _
            14 * Y_VAL - _
            2 * Z_VAL + 6


Exit Function
ERROR_LABEL:
CALL_MULTVAR_OBJ_15_FUNC = Err.number
End Function



'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_GRAD_15_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 035
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_GRAD_15_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim Z_VAL As Variant
Dim TEMP_ARR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG

X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)
Z_VAL = PARAM_VECTOR(3, 1)

ReDim TEMP_ARR(1 To 3, 1 To 1)
'This function is top-unlimited

TEMP_ARR(1, 1) = 4 * X_VAL + 6 * Y_VAL - 2 * Z_VAL - 6
TEMP_ARR(2, 1) = 6 * X_VAL + 20 * Y_VAL + 4 * Z_VAL - 14
TEMP_ARR(3, 1) = -2 * X_VAL + 4 * Y_VAL + 10 * Z_VAL - 2

CALL_MULTVAR_GRAD_15_FUNC = TEMP_ARR

Exit Function
ERROR_LABEL:
CALL_MULTVAR_GRAD_15_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_MULTVAR_CONST_15_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_TEST
'ID            : 036
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_MULTVAR_CONST_15_FUNC(ByRef PARAM_RNG As Variant)

Dim X_VAL As Variant
Dim Y_VAL As Variant
Dim Z_VAL As Variant
Dim TEMP_FLAG As Boolean
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

TEMP_FLAG = True
PARAM_VECTOR = PARAM_RNG
X_VAL = PARAM_VECTOR(1, 1)
Y_VAL = PARAM_VECTOR(2, 1)
Z_VAL = PARAM_VECTOR(3, 1)

If X_VAL < -10 Then: TEMP_FLAG = False
If X_VAL > 10 Then: TEMP_FLAG = False
If Y_VAL < -10 Then: TEMP_FLAG = False
If Y_VAL > 10 Then: TEMP_FLAG = False
If Z_VAL < -10 Then: TEMP_FLAG = False
If Z_VAL > 10 Then: TEMP_FLAG = False

CALL_MULTVAR_CONST_15_FUNC = TEMP_FLAG

Exit Function
ERROR_LABEL:
CALL_MULTVAR_CONST_15_FUNC = TEMP_FLAG
End Function
