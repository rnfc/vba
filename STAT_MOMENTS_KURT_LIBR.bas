Attribute VB_Name = "STAT_MOMENTS_KURT_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_KURT_FUNC
'DESCRIPTION   : Returns the kurtosis of a data set. Kurtosis characterizes the relative
'peakedness or flatness of a distribution compared with the normal distribution. Positive
'kurtosis indicates a relatively peaked distribution. Negative kurtosis indicates a
'relatively flat distribution.
'LIBRARY       : STATISTICS
'GROUP         : MOMENTS_KURT_KURT
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 08/29/2010
'************************************************************************************
'************************************************************************************

Function MATRIX_KURT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim j As Long
Dim i As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim KURT_VAL As Double
Dim MEAN_VAL As Double
Dim STDEV_VAL As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To 1, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    TEMP1_SUM = 0
    For i = 1 To NROWS
        TEMP1_SUM = TEMP1_SUM + DATA_MATRIX(i, j)
    Next i
    MEAN_VAL = TEMP1_SUM / NROWS
    TEMP1_SUM = 0: TEMP2_SUM = 0
    For i = 1 To NROWS
        TEMP1_SUM = TEMP1_SUM + (DATA_MATRIX(i, j) - MEAN_VAL) ^ 2
        TEMP2_SUM = TEMP2_SUM + (DATA_MATRIX(i, j) - MEAN_VAL)
    Next i
    STDEV_VAL = Sqr((TEMP1_SUM - TEMP2_SUM * TEMP2_SUM / NROWS) / (NROWS - 1)) 'Rounding error corrected var formula.
    KURT_VAL = 0
    For i = 1 To NROWS
        KURT_VAL = KURT_VAL + ((DATA_MATRIX(i, j) - MEAN_VAL) / STDEV_VAL) ^ 4
    Next i
    If VERSION <> 0 Then 'calcs with first two moments
        TEMP_MATRIX(1, j) = KURT_VAL / NROWS - 3
    Else 'Excel Definition
        TEMP_MATRIX(1, j) = (KURT_VAL * (NROWS * (NROWS + 1) / ((NROWS - 1) * (NROWS - 2) * (NROWS - 3)))) - ((3 * (NROWS - 1) ^ 2 / ((NROWS - 2) * (NROWS - 3)))) 'Excel Definition
    End If
Next j

MATRIX_KURT_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_KURT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : KURT_MOORE_FUNC
'DESCRIPTION   : Robust Moment Estimator: Moore's kurtosis measure
'LIBRARY       : STATISTICS
'GROUP         : MOMENTS_KURT
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function KURT_MOORE_FUNC(ByRef DATA_RNG As Variant)

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
KURT_MOORE_FUNC = ((HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.875, 0) - HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.625, 0)) - (HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.375, 0) - HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.125, 0))) / (HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.75, 0) - HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.25, 0))

Exit Function
ERROR_LABEL:
KURT_MOORE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : KURT_CROW_FUNC
'DESCRIPTION   : Robust Moment Estimator: Crow/Siddiqui measure
'LIBRARY       : STATISTICS
'GROUP         : MOMENTS_KURT
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function KURT_CROW_FUNC(ByRef DATA_RNG As Variant)

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
KURT_CROW_FUNC = (HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.975, 0) + HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.025, 0)) / (HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.75, 0) - HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.25, 0))

Exit Function
ERROR_LABEL:
KURT_CROW_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PEAKEDNESS_FUNC
'DESCRIPTION   : Robust Moment Estimator: measure for peakedness
'LIBRARY       : STATISTICS
'GROUP         : MOMENTS_KURT
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function PEAKEDNESS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal FACTOR1_VAL As Double = 0.125, _
Optional ByVal FACTOR2_VAL As Double = 0.25)

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
PEAKEDNESS_FUNC = (HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 1 - FACTOR1_VAL, 0) - HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, FACTOR1_VAL, 0)) / (HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 1 - FACTOR2_VAL, 0) - HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, FACTOR2_VAL, 0))
Exit Function
ERROR_LABEL:
PEAKEDNESS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : TAIL_WEIGHT_FUNC
'DESCRIPTION   : Robust Moment Estimator: measure for tail weights
'LIBRARY       : STATISTICS
'GROUP         : MOMENTS_KURT
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function TAIL_WEIGHT_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal FACTOR1_VAL As Double = 0.025, _
Optional ByVal FACTOR2_VAL As Double = 0.125)

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
TAIL_WEIGHT_FUNC = (HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 1 - FACTOR1_VAL, 0) - HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, FACTOR1_VAL, 0)) / (HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 1 - FACTOR2_VAL, 0) - HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, FACTOR2_VAL, 0))

Exit Function
ERROR_LABEL:
TAIL_WEIGHT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PEAK_TAIL_FUNC
'DESCRIPTION   : Robust Moment Estimator: Kurtosis L = Peakedness
'FACTOR1_VAL * Tails T
'LIBRARY       : STATISTICS
'GROUP         : MOMENTS_KURT
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function PEAK_TAIL_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal FACTOR1_VAL As Double = 0.025, _
Optional ByVal FACTOR2_VAL As Double = 0.125, _
Optional ByVal FACTOR3_VAL As Double = 0.25)
Dim DATA_VECTOR As Variant
On Error GoTo ERROR_LABEL
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
PEAK_TAIL_FUNC = PEAKEDNESS_FUNC(DATA_VECTOR, FACTOR2_VAL, FACTOR3_VAL) * TAIL_WEIGHT_FUNC(DATA_VECTOR, FACTOR1_VAL, FACTOR2_VAL)
Exit Function
ERROR_LABEL:
PEAK_TAIL_FUNC = Err.number
End Function
