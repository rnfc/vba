Attribute VB_Name = "OPTIM_MULTVAR_SEGMENT_LIBR"
'// PERFECT

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : SEGMENT_OPTIMIZATION_FUNC

'DESCRIPTION   : This minimize a function along a segment defined by two
'points a, b; a = starting point of the segment; b = ending point of the segment

'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_SEGMENT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Function SEGMENT_OPTIMIZATION_FUNC(ByRef LOWER_RNG As Variant, _
ByRef UPPER_RNG As Variant, _
ByRef MIN_FUNC_VAL As Double, _
ByRef NTRIALS As Long, _
ByVal FUNC_NAME_STR As String, _
Optional ByRef SCALE_RNG As Variant, _
Optional ByVal MIN_FLAG As Boolean = True, _
Optional ByVal NO_POINTS As Long = 20, _
Optional ByVal nLOOPS As Long = 1000)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NO_VAR As Long
Dim COUNTER As Long

Dim TEMP_SUM As Double
Dim TEMP_MAX As Double
Dim TEMP_FOSC As Double
Dim TEMP_LBA As Double
Dim TEMP_POINT As Double

Dim XTEMP_ABS As Double
Dim XTEMP_ERR As Double
Dim YTEMP_ERR As Double

Dim FIRST_DELTA As Double
Dim SECOND_DELTA As Double

Dim LOWER_VECTOR As Variant
Dim UPPER_VECTOR As Variant

Dim SCALE_VECTOR As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

Dim XTEMP_VECTOR As Variant
Dim YTEMP_VECTOR As Variant

Dim DELTA_VECTOR As Variant
Dim PARAM_VECTOR As Variant


Dim TEMP_MATRIX As Variant

Dim epsilon As Double

On Error GoTo ERROR_LABEL

epsilon = 0.000000000000001

LOWER_VECTOR = LOWER_RNG
If UBound(LOWER_VECTOR, 1) = 1 Then: LOWER_VECTOR = MATRIX_TRANSPOSE_FUNC(LOWER_VECTOR)

UPPER_VECTOR = UPPER_RNG
If UBound(UPPER_VECTOR, 1) = 1 Then: UPPER_VECTOR = MATRIX_TRANSPOSE_FUNC(UPPER_VECTOR)

NO_VAR = UBound(LOWER_VECTOR)

If IsArray(SCALE_RNG) = True Then
    SCALE_VECTOR = SCALE_RNG
    If UBound(SCALE_VECTOR, 1) = 1 Then: SCALE_VECTOR = MATRIX_TRANSPOSE_FUNC(SCALE_VECTOR)
Else
    ReDim SCALE_VECTOR(1 To NO_VAR, 1 To 1)
    For i = 1 To NO_VAR
        SCALE_VECTOR(i, 1) = 1
    Next i
End If

ReDim TEMP_MATRIX(1 To NO_POINTS, 1 To NO_VAR)
ReDim YTEMP_VECTOR(1 To NO_POINTS, 1 To 1)

ReDim XTEMP_VECTOR(1 To NO_VAR, 1 To 1)
ReDim PARAM_VECTOR(1 To NO_VAR, 1 To 1)
ReDim DELTA_VECTOR(1 To NO_VAR, 1 To 1)

ATEMP_VECTOR = LOWER_VECTOR
BTEMP_VECTOR = UPPER_VECTOR
''compute the direction versor
For i = 1 To NO_VAR
    DELTA_VECTOR(i, 1) = BTEMP_VECTOR(i, 1) - ATEMP_VECTOR(i, 1)
Next i

TEMP_SUM = 0
For i = 1 To NO_VAR 'Normalize |d|=1
    TEMP_SUM = TEMP_SUM + DELTA_VECTOR(i, 1) ^ 2
Next i

TEMP_SUM = (TEMP_SUM) ^ 0.5
If TEMP_SUM > epsilon Then
    For i = 1 To NO_VAR
        DELTA_VECTOR(i, 1) = DELTA_VECTOR(i, 1) / TEMP_SUM
    Next i
End If

COUNTER = 0
NTRIALS = 0

Do
    COUNTER = COUNTER + 1  'loop nTRIALS
    TEMP_LBA = 0
    For i = 1 To NO_VAR
        TEMP_LBA = TEMP_LBA + (BTEMP_VECTOR(i, 1) - ATEMP_VECTOR(i, 1)) ^ 2
    Next i
    TEMP_LBA = Sqr(TEMP_LBA)
    'take equispaced points along direction DELTA_VECTOR
    FIRST_DELTA = TEMP_LBA / (NO_POINTS - 1)
    For i = 1 To NO_POINTS
        For j = 1 To NO_VAR
            TEMP_MATRIX(i, j) = ATEMP_VECTOR(j, 1) + (i - 1) * _
                                FIRST_DELTA * DELTA_VECTOR(j, 1)
        Next j
    Next i
    
    YTEMP_VECTOR = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, _
                TEMP_MATRIX, SCALE_VECTOR, MIN_FLAG)
    'search for the lowest function value
    k = 1
    MIN_FUNC_VAL = YTEMP_VECTOR(1, 1)
    TEMP_MAX = YTEMP_VECTOR(1, 1)
    For i = 2 To NO_POINTS
        If YTEMP_VECTOR(i, 1) < MIN_FUNC_VAL Then
            k = i
            MIN_FUNC_VAL = YTEMP_VECTOR(k, 1)
        End If
        If YTEMP_VECTOR(i, 1) > TEMP_MAX Then TEMP_MAX = YTEMP_VECTOR(i, 1)
    Next i
    'choose the other bound
    If k = NO_POINTS Then
        i = k - 1
        j = k
    ElseIf k = 1 Then
        i = k
        j = k + 1
    Else
        i = k - 1
        j = k + 1
    End If

    YTEMP_ERR = 0
    XTEMP_ABS = 0
    For l = 1 To NO_VAR
        ATEMP_VECTOR(l, 1) = TEMP_MATRIX(i, l)
        BTEMP_VECTOR(l, 1) = TEMP_MATRIX(j, l)
        YTEMP_ERR = YTEMP_ERR + Abs(ATEMP_VECTOR(l, 1) - BTEMP_VECTOR(l, 1))
        XTEMP_ABS = XTEMP_ABS + (Abs(ATEMP_VECTOR(l, 1)) + _
                    Abs(BTEMP_VECTOR(l, 1))) / 2
    Next l
    XTEMP_ERR = YTEMP_ERR / (XTEMP_ABS + 100 * epsilon) 'relative error
    'compute the relative function oscillation
    
    TEMP_FOSC = (TEMP_MAX - MIN_FUNC_VAL) / (Abs(TEMP_MAX) + _
                Abs(MIN_FUNC_VAL) + 1000 * epsilon)
    
    If TEMP_FOSC > 10 ^ -5 Then SECOND_DELTA = TEMP_LBA
    NTRIALS = NTRIALS + NO_POINTS

Loop Until XTEMP_ERR < epsilon Or NTRIALS > nLOOPS

For j = 1 To NO_VAR
    XTEMP_VECTOR(j, 1) = (ATEMP_VECTOR(j, 1) + BTEMP_VECTOR(j, 1)) / 2
Next j

MIN_FUNC_VAL = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, XTEMP_VECTOR, _
                SCALE_VECTOR, MIN_FLAG)
'' attempt to refine the solution with the parabolic interpolation
If SECOND_DELTA > 0.3 Then: GoTo 1983
For j = 1 To NO_VAR
    ATEMP_VECTOR(j, 1) = XTEMP_VECTOR(j, 1) - SECOND_DELTA * DELTA_VECTOR(j, 1)
    BTEMP_VECTOR(j, 1) = XTEMP_VECTOR(j, 1) + SECOND_DELTA * DELTA_VECTOR(j, 1)
Next j

ReDim YTEMP_VECTOR(1 To 3, 1 To 1)

YTEMP_VECTOR(1, 1) = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, ATEMP_VECTOR, _
                    SCALE_VECTOR, MIN_FLAG) - MIN_FUNC_VAL
YTEMP_VECTOR(2, 1) = 0
YTEMP_VECTOR(3, 1) = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, BTEMP_VECTOR, _
                    SCALE_VECTOR, MIN_FLAG) - MIN_FUNC_VAL

TEMP_POINT = SEGMENT_PARABOLIC_FUNC(-SECOND_DELTA, 0, _
             SECOND_DELTA, YTEMP_VECTOR(1, 1), _
             YTEMP_VECTOR(2, 1), YTEMP_VECTOR(3, 1))

If TEMP_POINT > -SECOND_DELTA And TEMP_POINT < SECOND_DELTA Then
    For j = 1 To NO_VAR
        PARAM_VECTOR(j, 1) = XTEMP_VECTOR(j, 1) + TEMP_POINT * DELTA_VECTOR(j, 1)
    Next j
    'compute the function in the new point
    
    YTEMP_VECTOR(2, 1) = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, PARAM_VECTOR, _
                        SCALE_VECTOR, MIN_FLAG)
    YTEMP_VECTOR(3, 1) = MIN_FUNC_VAL * (1 + 4 * epsilon)
    If YTEMP_VECTOR(2, 1) > YTEMP_VECTOR(3, 1) Then
        'reject new PARAM_VECTOR
        
        MIN_FUNC_VAL = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, XTEMP_VECTOR, _
                        SCALE_VECTOR, MIN_FLAG)
    Else
        'accept the parabolic approximation
        For j = 1 To NO_VAR
            XTEMP_VECTOR(j, 1) = PARAM_VECTOR(j, 1)
        Next j
    End If
Else
    'reject new PARAM_VECTOR
    MIN_FUNC_VAL = MULTVAR_CALL_OBJ_FUNC(FUNC_NAME_STR, XTEMP_VECTOR, _
                    SCALE_VECTOR, MIN_FLAG)
End If

1983:
SEGMENT_OPTIMIZATION_FUNC = XTEMP_VECTOR

Exit Function
ERROR_LABEL:
SEGMENT_OPTIMIZATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : SEGMENT_PARABOLIC_FUNC
'DESCRIPTION   : Find of local extreme (max or min) with the parabolic method
'LIBRARY       : OPTIMIZATION
'GROUP         : MULTVAR_SEGMENT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 02/13/2009
'************************************************************************************
'************************************************************************************

Private Function SEGMENT_PARABOLIC_FUNC(ByVal X1_VAL As Double, _
ByVal X2_VAL As Double, _
ByVal X3_VAL As Double, _
ByVal Y1_VAL As Double, _
ByVal Y2_VAL As Double, _
ByVal Y3_VAL As Double)

Dim XTEMP_A As Double
Dim XTEMP_B As Double

Dim YTEMP_A As Double
Dim YTEMP_B As Double

Dim TEMP_DIFF As Double

On Error GoTo ERROR_LABEL
'-----------------------------------------------------------------------------
XTEMP_A = X2_VAL - X1_VAL
XTEMP_B = X2_VAL - X3_VAL
YTEMP_A = Y2_VAL - Y1_VAL
YTEMP_B = Y2_VAL - Y3_VAL
TEMP_DIFF = XTEMP_A * YTEMP_B - XTEMP_B * YTEMP_A
'-----------------------------------------------------------------------------

If TEMP_DIFF <> 0 Then
    SEGMENT_PARABOLIC_FUNC = X2_VAL - (XTEMP_A ^ 2 * YTEMP_B - _
                               XTEMP_B ^ 2 * YTEMP_A) / TEMP_DIFF / 2
Else
    If XTEMP_A <> 0 Then
        SEGMENT_PARABOLIC_FUNC = X1_VAL - YTEMP_A / XTEMP_A
    Else
        SEGMENT_PARABOLIC_FUNC = X1_VAL
    End If
End If

Exit Function
ERROR_LABEL:
SEGMENT_PARABOLIC_FUNC = Err.number
End Function
