Attribute VB_Name = "STAT_MOMENTS_SKEW_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_SKEW_FUNC

'DESCRIPTION   : Returns the skewness of a distribution. Skewness characterizes the
'degree of asymmetry of a distribution around its mean. Positive skewness indicates a
'distribution with an asymmetric tail extending toward more positive values. Negative
'skewness indicates a distribution with an asymmetric tail extending toward more
'negative values.

'LIBRARY       : STATISTICS
'GROUP         : MOMENTS_SKEW
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 08/29/2010
'************************************************************************************
'************************************************************************************

Function MATRIX_SKEW_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim j As Long
Dim i As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim SKEW_VAL As Double
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
    For i = 1 To NROWS: TEMP1_SUM = TEMP1_SUM + DATA_MATRIX(i, j): Next i
    MEAN_VAL = TEMP1_SUM / NROWS
    TEMP1_SUM = 0: TEMP2_SUM = 0
    For i = 1 To NROWS
        TEMP1_SUM = TEMP1_SUM + (DATA_MATRIX(i, j) - MEAN_VAL) ^ 2
        TEMP2_SUM = TEMP2_SUM + (DATA_MATRIX(i, j) - MEAN_VAL)
    Next i
    STDEV_VAL = Sqr((TEMP1_SUM - TEMP2_SUM * TEMP2_SUM / NROWS) / (NROWS - 1)) 'Rounding error corrected var formula.
    SKEW_VAL = 0
    For i = 1 To NROWS
        SKEW_VAL = SKEW_VAL + ((DATA_MATRIX(i, j) - MEAN_VAL) / STDEV_VAL) ^ 3
    Next i
    If VERSION <> 0 Then 'calcs with first two moments
        TEMP_MATRIX(1, j) = SKEW_VAL / NROWS
    Else 'Excel Definition
        TEMP_MATRIX(1, j) = SKEW_VAL * (NROWS / ((NROWS - 1) * (NROWS - 2))) 'Excel Definition
    End If
Next j

MATRIX_SKEW_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_SKEW_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SKEW_BOWLEY_FUNC
'DESCRIPTION   : Robust Moment Estimator: Bowley's Skewness Coefficient
'LIBRARY       : STATISTICS
'GROUP         : MOMENTS_SKEW
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 08/29/2010
'************************************************************************************
'************************************************************************************

Function SKEW_BOWLEY_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim Q1_VAL As Double
Dim Q2_VAL As Double
Dim Q3_VAL As Double

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)

Q1_VAL = HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.25, 0)
Q2_VAL = HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.5, 0)
Q3_VAL = HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.75, 0)

Select Case OUTPUT
Case 0
    SKEW_BOWLEY_FUNC = (Q3_VAL - 2 * Q2_VAL + Q1_VAL) / (Q3_VAL - Q1_VAL)
Case 1
    SKEW_BOWLEY_FUNC = Q3_VAL - Q1_VAL
Case Else
    SKEW_BOWLEY_FUNC = Array((Q3_VAL - 2 * Q2_VAL + Q1_VAL) / (Q3_VAL - Q1_VAL), Q3_VAL - Q1_VAL)
End Select

Exit Function
ERROR_LABEL:
SKEW_BOWLEY_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : SKEW_PEARSON_FUNC
'DESCRIPTION   : Pearson's Skewness Coefficient (= "Skewness Sharpe Ratio")
'LIBRARY       : STATISTICS
'GROUP         : MOMENTS_SKEW
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 08/29/2010
'************************************************************************************
'************************************************************************************

Function SKEW_PEARSON_FUNC(ByRef DATA_RNG As Variant)

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim MEDIAN_VAL As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
MEDIAN_VAL = HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.5, 1)
SIGMA_VAL = MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1)

SKEW_PEARSON_FUNC = (MEAN_VAL - MEDIAN_VAL) / SIGMA_VAL

Exit Function
ERROR_LABEL:
SKEW_PEARSON_FUNC = Err.number
End Function
