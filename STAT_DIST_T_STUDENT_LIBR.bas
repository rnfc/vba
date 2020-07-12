Attribute VB_Name = "STAT_DIST_T_STUDENT_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : FIT_TDIST_FUNC
'DESCRIPTION   : Fit a student t distribution to a given return vector.
'LIBRARY       : STATISTICS
'GROUP         : DIST_T_STUDENT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function FIT_TDIST_FUNC(ByRef DATA_RNG As Variant, _
ByVal nDEGREES As Long)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long

Dim MIN_VAL As Double
Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim TEMP_SUM As Double
Dim TEMP_VAL As Double

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

k = 2
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)

' standarddize & sort returns
MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
SIGMA_VAL = MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1)

DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
ReDim TEMP1_VECTOR(1 To NROWS, 1)
For i = 1 To NROWS
    TEMP1_VECTOR(i, 1) = (DATA_VECTOR(i, 1) - MEAN_VAL) / SIGMA_VAL
Next i

TEMP_SUM = 0
For i = 1 To NROWS 'standard normal distribution
    TEMP_SUM = TEMP_SUM + (TEMP1_VECTOR(i, 1) - NORMSINV_FUNC(i / (NROWS + 1), 0, 1, 0)) ^ k
Next i
TEMP_SUM = (TEMP_SUM / NROWS) ^ (1 / k)

ReDim TEMP2_VECTOR(1 To nDEGREES, 1 To 1)
For j = 1 To nDEGREES 'series of DEGREES to student t distributiob
    For i = 1 To NROWS
        TEMP_VAL = i / (NROWS + 1)
        If TEMP_VAL < 0.5 Then 'Inverse CDF T-Dist
            TEMP_VAL = INVERSE_TDIST_FUNC(TEMP_VAL, j)
        Else
            TEMP_VAL = INVERSE_TDIST_FUNC((1 - TEMP_VAL), j) * -1
        End If
        TEMP2_VECTOR(j, 1) = TEMP2_VECTOR(j, 1) + (TEMP1_VECTOR(i, 1) - TEMP_VAL) ^ k
    Next i
    TEMP2_VECTOR(j, 1) = (TEMP2_VECTOR(j, 1) / NROWS) ^ (1 / k)
Next j

ReDim TEMP_MATRIX(1 To 2, 1 To 1)
k = 0: MIN_VAL = 2 ^ 52
For j = 1 To nDEGREES
    If TEMP2_VECTOR(j, 1) < MIN_VAL Then
        MIN_VAL = TEMP2_VECTOR(j, 1)
        k = j
    End If
Next j
'k --> Best-Fit nDEGREES of Freedom
'MIN_VAL < TEMP_SUM --> Is Fit of T Distr better than Normal Distr?

FIT_TDIST_FUNC = Array(MIN_VAL, k, MIN_VAL < TEMP_SUM)

Exit Function
ERROR_LABEL:
FIT_TDIST_FUNC = Err.number
End Function
