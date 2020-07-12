Attribute VB_Name = "OPTIM_NLE_FRAMES_LIBR"


'// PERFECT

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_OBJ_1_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_OBJ_1_FUNC(ByRef PARAM_RNG As Variant)


Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

'The intersection is located in the range: 0 < x < 0.005 and 200 < y < 400
'--------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 2, 1 To 1)
    
    TEMP_VECTOR(1, 1) = PARAM_VECTOR(1, 1) - _
                        (PARAM_VECTOR(2, 1) / 840) ^ 6 - _
                        PARAM_VECTOR(2, 1) / 200000
    
    TEMP_VECTOR(2, 1) = PARAM_VECTOR(1, 1) - _
                        1 / PARAM_VECTOR(2, 1)

CALL_NLE_OBJ_1_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_OBJ_1_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_JACOBI_1_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_JACOBI_1_FUNC(ByRef PARAM_RNG As Variant)

Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

TEMP_VECTOR = JACOBI_FORWARD_FUNC("CALL_NLE_OBJ_1_FUNC", PARAM_VECTOR, 0.00001)
CALL_NLE_JACOBI_1_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_JACOBI_1_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_OBJ_2_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_OBJ_2_FUNC(ByRef PARAM_RNG As Variant)

Dim PI_VAL As Double
Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

PI_VAL = 3.14159265358979

'The plot method shows clearly that there are two intersection points
'between the curves. We estimate the first solution near (0.5, 0.3) and the
'second solution near (0.3, -0.2). They are a raw estimation but they should be
'sufficient close to start the Newton algorithm with a good chance
'--------------------------------------------------------------------------------

    ReDim TEMP_VECTOR(1 To 2, 1 To 1)
    TEMP_VECTOR(1, 1) = 5 * PARAM_VECTOR(1, 1) ^ 2 - _
                        6 * PARAM_VECTOR(1, 1) * PARAM_VECTOR(2, 1) + _
                        5 * PARAM_VECTOR(2, 1) ^ 2 - 1
    TEMP_VECTOR(2, 1) = 2 ^ (-1 * PARAM_VECTOR(1, 1)) - _
                        Cos(PI_VAL * PARAM_VECTOR(2, 1))

CALL_NLE_OBJ_2_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_OBJ_2_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_JACOBI_2_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_JACOBI_2_FUNC(ByRef PARAM_RNG As Variant)

Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

TEMP_VECTOR = JACOBI_FORWARD_FUNC("CALL_NLE_OBJ_2_FUNC", PARAM_VECTOR, 0.00001)
CALL_NLE_JACOBI_2_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_JACOBI_2_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_OBJ_3_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_OBJ_3_FUNC(ByRef PARAM_RNG As Variant)

Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

'The 4 intersections are symmetric respect to the origin.
'One of these is located in the range: 1 < x < 2 and 1 < y < 2
'--------------------------------------------------------------------------------

    ReDim TEMP_VECTOR(1 To 2, 1 To 1)
    TEMP_VECTOR(1, 1) = PARAM_VECTOR(1, 1) ^ 4 + _
                        2 * PARAM_VECTOR(2, 1) ^ 4 - _
                        16
    TEMP_VECTOR(2, 1) = PARAM_VECTOR(1, 1) ^ 2 + _
                        PARAM_VECTOR(2, 1) ^ 2 - _
                        4

CALL_NLE_OBJ_3_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_OBJ_3_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_JACOBI_3_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_JACOBI_3_FUNC(ByRef PARAM_RNG As Variant)

Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

TEMP_VECTOR = JACOBI_FORWARD_FUNC("CALL_NLE_OBJ_3_FUNC", PARAM_VECTOR, 0.00001)
CALL_NLE_JACOBI_3_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_JACOBI_3_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_OBJ_4_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_OBJ_4_FUNC(ByRef PARAM_RNG As Variant)


Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

'The plot method shows clearly that there are only one intersection point
'between the curves. We estimate the solution near (0.5, 0.5). It should be
'sufficient close to start the Broyden algorithm with a good chance.
'--------------------------------------------------------------------------------

    ReDim TEMP_VECTOR(1 To 2, 1 To 1)
    TEMP_VECTOR(1, 1) = PARAM_VECTOR(1, 1) * PARAM_VECTOR(2, 1) + _
                        PARAM_VECTOR(1, 1) ^ 1.2 + _
                        PARAM_VECTOR(2, 1) ^ 0.5 - _
                        1
    TEMP_VECTOR(2, 1) = Exp(-2 * PARAM_VECTOR(1, 1)) + _
                        PARAM_VECTOR(2, 1) - 1
                        

CALL_NLE_OBJ_4_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_OBJ_4_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_JACOBI_4_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_JACOBI_4_FUNC(ByRef PARAM_RNG As Variant)

Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

TEMP_VECTOR = JACOBI_FORWARD_FUNC("CALL_NLE_OBJ_4_FUNC", PARAM_VECTOR, 0.00001)
CALL_NLE_JACOBI_4_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_JACOBI_4_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_OBJ_5_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_OBJ_5_FUNC(ByRef PARAM_RNG As Variant)


Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

'There are 4 intersections. One of these is located in the range:
'1 < x < 3 and 2 < y < 4 and we may choose the starting point
'x0= 2, y0 = 4. Repeating with the starting points (x0,y0) = (3, 0),
'(x0,y0) = (-2, 0.1), (x0,y0) = (-2, -2) we get all the system solutions
'--------------------------------------------------------------------------------

    
    ReDim TEMP_VECTOR(1 To 2, 1 To 1)
    TEMP_VECTOR(1, 1) = PARAM_VECTOR(1, 1) ^ 4 + _
                        3 * PARAM_VECTOR(2, 1) ^ 2 - _
                        8 * PARAM_VECTOR(1, 1) + _
                        2 * PARAM_VECTOR(2, 1) - 33

    TEMP_VECTOR(2, 1) = PARAM_VECTOR(1, 1) ^ 3 + _
                        3 * PARAM_VECTOR(2, 1) ^ 2 - _
                        8 * PARAM_VECTOR(1, 1) * PARAM_VECTOR(2, 1) - _
                        12 * PARAM_VECTOR(1, 1) ^ 2 + _
                        PARAM_VECTOR(1, 1) + 59


CALL_NLE_OBJ_5_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_OBJ_5_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_JACOBI_5_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_JACOBI_5_FUNC(ByRef PARAM_RNG As Variant)

Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

TEMP_VECTOR = JACOBI_FORWARD_FUNC("CALL_NLE_OBJ_5_FUNC", PARAM_VECTOR, 0.00001)
CALL_NLE_JACOBI_5_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_JACOBI_5_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_OBJ_6_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_OBJ_6_FUNC(ByRef PARAM_RNG As Variant)


Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

'The plot method shows clearly that there are only one intersection point
'between the curves. We estimate the solution near (0.5, 0.1). It should be
'sufficient close to start the Brown algorithm with a good chance.
'--------------------------------------------------------------------------------

    ReDim TEMP_VECTOR(1 To 2, 1 To 1)
    TEMP_VECTOR(1, 1) = 2 * PARAM_VECTOR(1, 1) + _
                        Cos(PARAM_VECTOR(2, 1)) - 2

    TEMP_VECTOR(2, 1) = Cos(PARAM_VECTOR(1, 1)) + _
                        2 * PARAM_VECTOR(2, 1) - 1

'--------------------------------------------------------------------------------

CALL_NLE_OBJ_6_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_OBJ_6_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_JACOBI_6_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_JACOBI_6_FUNC(ByRef PARAM_RNG As Variant)

Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

TEMP_VECTOR = JACOBI_FORWARD_FUNC("CALL_NLE_OBJ_6_FUNC", PARAM_VECTOR, 0.00001)
CALL_NLE_JACOBI_6_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_JACOBI_6_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_OBJ_7_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_OBJ_7_FUNC(ByRef PARAM_RNG As Variant)


Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

'There are 4 intersections located in the box range
'- 4 <= x <= 4 ; - 4 <= y <= 4
'--------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 2, 1 To 1)
    TEMP_VECTOR(1, 1) = PARAM_VECTOR(1, 1) ^ 2 - _
                        PARAM_VECTOR(2, 1) ^ 2 + _
                        PARAM_VECTOR(1, 1) * PARAM_VECTOR(2, 1) - 1

    TEMP_VECTOR(2, 1) = PARAM_VECTOR(1, 1) ^ 4 + _
                        2 * PARAM_VECTOR(2, 1) ^ 4 + _
                        PARAM_VECTOR(1, 1) * PARAM_VECTOR(2, 1) - 16

CALL_NLE_OBJ_7_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_OBJ_7_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_JACOBI_7_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_JACOBI_7_FUNC(ByRef PARAM_RNG As Variant)

Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

TEMP_VECTOR = JACOBI_FORWARD_FUNC("CALL_NLE_OBJ_7_FUNC", PARAM_VECTOR, 0.00001)
CALL_NLE_JACOBI_7_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_JACOBI_7_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_OBJ_8_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_OBJ_8_FUNC(ByRef PARAM_RNG As Variant)

Dim PI_VAL As Double
Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

'Plotting the equation we note that the roots are located in
'the range - 5 <= x <= 5
'--------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 1, 1 To 1)
    TEMP_VECTOR(1, 1) = Cos(PI_VAL * PARAM_VECTOR(1, 1)) + 1 / 4 * PARAM_VECTOR(1, 1)


CALL_NLE_OBJ_8_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_OBJ_8_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_JACOBI_8_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_JACOBI_8_FUNC(ByRef PARAM_RNG As Variant)

Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

TEMP_VECTOR = JACOBI_FORWARD_FUNC("CALL_NLE_OBJ_8_FUNC", PARAM_VECTOR, 0.00001)
CALL_NLE_JACOBI_8_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_JACOBI_8_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_OBJ_9_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_OBJ_9_FUNC(ByRef PARAM_RNG As Variant)


Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

'Of course the bounding could be more tight, for example -0.2 <= x <= 2.
'--------------------------------------------------------------------------------

    ReDim TEMP_VECTOR(1 To 3, 1 To 1)
    TEMP_VECTOR(1, 1) = 2 * PARAM_VECTOR(1, 1) + _
                        Cos(PARAM_VECTOR(2, 1)) + _
                        Cos(PARAM_VECTOR(3, 1)) - _
                        1.9

    TEMP_VECTOR(2, 1) = Cos(PARAM_VECTOR(1, 1)) + _
                        2 * PARAM_VECTOR(2, 1) + _
                        Cos(PARAM_VECTOR(3, 1)) - _
                        1.8

    TEMP_VECTOR(3, 1) = Cos(PARAM_VECTOR(1, 1)) + _
                        Cos(PARAM_VECTOR(2, 1)) + _
                        2 * PARAM_VECTOR(3, 1) - _
                        1.7
                        
CALL_NLE_OBJ_9_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_OBJ_9_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_JACOBI_9_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 018
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_JACOBI_9_FUNC(ByRef PARAM_RNG As Variant)

Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

TEMP_VECTOR = JACOBI_FORWARD_FUNC("CALL_NLE_OBJ_9_FUNC", PARAM_VECTOR, 0.00001)
CALL_NLE_JACOBI_9_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_JACOBI_9_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_OBJ_10_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 019
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_OBJ_10_FUNC(ByRef PARAM_RNG As Variant)


Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

'We can choose the lower limit -2 and the upper limit 5 for each
'variable. This is a sort of "bracketing" of the system roots.
'--------------------------------------------------------------------------------

    ReDim TEMP_VECTOR(1 To 3, 1 To 1)
    TEMP_VECTOR(1, 1) = Exp(-PARAM_VECTOR(1, 1)) - _
                        PARAM_VECTOR(1, 1) * PARAM_VECTOR(2, 1) + _
                        PARAM_VECTOR(2, 1) ^ 3

    TEMP_VECTOR(2, 1) = 3 * PARAM_VECTOR(2, 1) ^ 2 + _
                        2 * PARAM_VECTOR(3, 1) ^ 2 - 4

    TEMP_VECTOR(3, 1) = PARAM_VECTOR(1, 1) ^ 2 - _
                        2 * PARAM_VECTOR(1, 1) + _
                        2 * PARAM_VECTOR(3, 1) ^ 2 + _
                        PARAM_VECTOR(1, 1) * PARAM_VECTOR(3, 1) - _
                        6

CALL_NLE_OBJ_10_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_OBJ_10_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CALL_NLE_JACOBI_10_FUNC
'DESCRIPTION   :
'LIBRARY       : OPTIMIZATION
'GROUP         : NLE_FRAMES
'ID            : 020
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function CALL_NLE_JACOBI_10_FUNC(ByRef PARAM_RNG As Variant)

Dim TEMP_VECTOR As Variant
Dim PARAM_VECTOR As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
If UBound(PARAM_VECTOR, 1) = 1 Then: PARAM_VECTOR = MATRIX_TRANSPOSE_FUNC(PARAM_VECTOR)

TEMP_VECTOR = JACOBI_FORWARD_FUNC("CALL_NLE_OBJ_10_FUNC", PARAM_VECTOR, 0.00001)
CALL_NLE_JACOBI_10_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
CALL_NLE_JACOBI_10_FUNC = Err.number
End Function
