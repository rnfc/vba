Attribute VB_Name = "STAT_MOMENTS_OUTLIER_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_ZCORE_FUNC
'DESCRIPTION   : Computes z scores of a data vector
'LIBRARY       : STATISTICS
'GROUP         : OUTLIERS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function VECTOR_ZCORE_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim NROWS As Long

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

MEAN_VAL = MATRIX_MEAN_FUNC(STRIP_NUMERICS_FUNC(DATA_VECTOR))(1, 1)
SIGMA_VAL = MATRIX_STDEV_FUNC(STRIP_NUMERICS_FUNC(DATA_VECTOR))(1, 1)

NROWS = UBound(DATA_VECTOR, 1)
ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    If IsNumeric(DATA_VECTOR(i, 1)) And Not IsEmpty(DATA_VECTOR(i, 1)) Then
        TEMP_VECTOR(i, 1) = (DATA_VECTOR(i, 1) - MEAN_VAL) / SIGMA_VAL
    Else
        TEMP_VECTOR(i, 1) = 0
    End If
Next i
'=IF(ABS(C19:C53)>=3,"red",IF(ABS(C19:C53)>=2,"yellow","green"))
VECTOR_ZCORE_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_ZCORE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_MODIFIED_ZSCORE_FUNC
'DESCRIPTION   : Computes modified z scores of a data vector
'LIBRARY       : STATISTICS
'GROUP         : OUTLIERS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************


Function VECTOR_MODIFIED_ZSCORE_FUNC(ByRef DATA_RNG As Variant)
    
Dim i As Long
Dim NROWS As Long

Dim VAR_VAL As Double
Dim MEDIAN_VAL As Double

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

MEDIAN_VAL = HISTOGRAM_PERCENTILE_FUNC(STRIP_NUMERICS_FUNC(DATA_VECTOR), 0.5, 1)
VAR_VAL = MATRIX_ABSOLUTE_DEVIATION_FUNC(DATA_VECTOR)

NROWS = UBound(DATA_VECTOR, 1)
ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
    If IsNumeric(DATA_VECTOR(i, 1)) And Not IsEmpty(DATA_VECTOR(i, 1)) Then
        TEMP_VECTOR(i, 1) = (DATA_VECTOR(i, 1) - MEDIAN_VAL) / VAR_VAL
    Else
        TEMP_VECTOR(i, 1) = CVErr(xlErrNA)
    End If
Next i
'=IF(ABS(F19:F53)>=3,"red",IF(ABS(F19:F53)>=2,"yellow","green"))
VECTOR_MODIFIED_ZSCORE_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
VECTOR_MODIFIED_ZSCORE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GRUBBS_TABLE_FUNC
'DESCRIPTION   : Grubbs' Table (one-sided and double-sided versions) to detect
'outliers
'http://en.wikipedia.org/wiki/Grubbs%27_test_for_outliers
'http://elsmar.com/Forums/showthread.php?t=6918
'LIBRARY       : STATISTICS
'GROUP         : OUTLIERS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function GRUBBS_TABLE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal ALPHA As Double = 0.05, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

'------------------------------------------------------------------------------------
'Grubbs test
'------------------------------------------------------------------------------------
'H0: No outliers
'Reject H0: if G > z(a)
'------------------------------------------------------------------------------------
'Data Must Be sorted in Ascending Order
'------------------------------------------------------------------------------------

Dim i As Long
Dim NROWS As Long

Dim A_VAL As Double
Dim B_VAL As Double
Dim C_VAL As Double

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL
    
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then
    DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
    DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
End If

NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_MATRIX(1 To 12, 1 To 4)

TEMP_MATRIX(12, 1) = "STD(X)"
TEMP_MATRIX(12, 2) = MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1)
TEMP_MATRIX(12, 3) = ""
TEMP_MATRIX(12, 4) = ""

TEMP_MATRIX(11, 1) = "MEAN(X)"
TEMP_MATRIX(11, 2) = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
TEMP_MATRIX(11, 3) = ""
TEMP_MATRIX(11, 4) = ""

TEMP_MATRIX(10, 1) = "A"
TEMP_MATRIX(10, 2) = ALPHA
TEMP_MATRIX(10, 3) = ""
TEMP_MATRIX(10, 4) = ""

TEMP_MATRIX(9, 1) = "NOBS"
TEMP_MATRIX(9, 2) = UBound(DATA_VECTOR, 1)
TEMP_MATRIX(9, 3) = ""
TEMP_MATRIX(9, 4) = ""

TEMP_MATRIX(8, 1) = "DEGREE_FREEDOM"
TEMP_MATRIX(8, 2) = TEMP_MATRIX(9, 2) - 2
TEMP_MATRIX(8, 3) = ""
TEMP_MATRIX(8, 4) = ""

ReDim TEMP_VECTOR(1 To NROWS, 1 To 3)
A_VAL = 0: B_VAL = 0: C_VAL = 0
For i = 1 To NROWS
    TEMP_VECTOR(i, 1) = Abs(DATA_VECTOR(i, 1) - TEMP_MATRIX(11, 2))
    If TEMP_VECTOR(i, 1) > A_VAL Then: A_VAL = TEMP_VECTOR(i, 1)
    
    TEMP_VECTOR(i, 2) = TEMP_MATRIX(11, 2) - DATA_VECTOR(i, 1)
    If TEMP_VECTOR(i, 2) > B_VAL Then: B_VAL = TEMP_VECTOR(i, 2)
    
    TEMP_VECTOR(i, 3) = DATA_VECTOR(i, 1) - TEMP_MATRIX(11, 2)
    If TEMP_VECTOR(i, 3) > C_VAL Then: C_VAL = TEMP_VECTOR(i, 3)
Next i


TEMP_MATRIX(7, 1) = "G"
TEMP_MATRIX(7, 2) = A_VAL / TEMP_MATRIX(12, 2)
TEMP_MATRIX(7, 3) = B_VAL / TEMP_MATRIX(12, 2)
TEMP_MATRIX(7, 4) = C_VAL / TEMP_MATRIX(12, 2)

TEMP_MATRIX(6, 1) = "B"
TEMP_MATRIX(6, 2) = TEMP_MATRIX(10, 2) / (2 * TEMP_MATRIX(9, 2))
TEMP_MATRIX(6, 3) = TEMP_MATRIX(10, 2) / TEMP_MATRIX(9, 2)
TEMP_MATRIX(6, 4) = TEMP_MATRIX(10, 2) / TEMP_MATRIX(9, 2)

TEMP_MATRIX(5, 1) = "T"
TEMP_MATRIX(5, 2) = INVERSE_TDIST_FUNC(TEMP_MATRIX(6, 2), TEMP_MATRIX(8, 2)) * -1
TEMP_MATRIX(5, 3) = INVERSE_TDIST_FUNC(TEMP_MATRIX(6, 3), TEMP_MATRIX(8, 2)) * -1
TEMP_MATRIX(5, 4) = INVERSE_TDIST_FUNC(TEMP_MATRIX(6, 4), TEMP_MATRIX(8, 2)) * -1

TEMP_MATRIX(4, 1) = "GRUBBS"
TEMP_MATRIX(4, 2) = GRUBBS_TEST_FUNC(DATA_VECTOR, TEMP_MATRIX(10, 2))
TEMP_MATRIX(4, 3) = GRUBBS_NEGATIVE_OUTLIERS_TEST_FUNC(DATA_VECTOR, TEMP_MATRIX(10, 2))
TEMP_MATRIX(4, 4) = GRUBBS_POSITIVE_OUTLIERS_TEST_FUNC(DATA_VECTOR, TEMP_MATRIX(10, 2))

TEMP_MATRIX(3, 1) = "Z(a)"
TEMP_MATRIX(3, 2) = ((TEMP_MATRIX(9, 2) - 1) / Sqr(TEMP_MATRIX(9, 2))) * Sqr((TEMP_MATRIX(5, 2) ^ 2) / (TEMP_MATRIX(9, 2) - 2 + TEMP_MATRIX(5, 2) ^ 2))
TEMP_MATRIX(3, 3) = ((TEMP_MATRIX(9, 2) - 1) / Sqr(TEMP_MATRIX(9, 2))) * Sqr((TEMP_MATRIX(5, 3) ^ 2) / (TEMP_MATRIX(9, 2) - 2 + TEMP_MATRIX(5, 3) ^ 2))
TEMP_MATRIX(3, 4) = ((TEMP_MATRIX(9, 2) - 1) / Sqr(TEMP_MATRIX(9, 2))) * Sqr((TEMP_MATRIX(5, 4) ^ 2) / (TEMP_MATRIX(9, 2) - 2 + TEMP_MATRIX(5, 4) ^ 2))

TEMP_MATRIX(2, 1) = "G>Z(a)"
TEMP_MATRIX(2, 2) = TEMP_MATRIX(7, 2) > TEMP_MATRIX(3, 2)
TEMP_MATRIX(2, 3) = TEMP_MATRIX(7, 3) > TEMP_MATRIX(3, 3)
TEMP_MATRIX(2, 4) = TEMP_MATRIX(7, 4) > TEMP_MATRIX(3, 4)

TEMP_MATRIX(1, 1) = "-"
TEMP_MATRIX(1, 2) = "OUTLIERS"
TEMP_MATRIX(1, 3) = "NEG_OUTLIERS"
TEMP_MATRIX(1, 4) = "POS_OUTLIERS"

GRUBBS_TABLE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
GRUBBS_TABLE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GRUBBS_TEST_FUNC
'DESCRIPTION   : Computes Grubbs' test
'LIBRARY       : STATISTICS
'GROUP         : OUTLIERS
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function GRUBBS_TEST_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal ALPHA As Double = 0.05)

Dim i As Long
Dim NROWS As Long

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim DELTA_VAL As Double

Dim TEMP_VAL As Double
Dim TEMP_FACTOR As Double

Dim TDIST_VAL As Double

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
NROWS = UBound(DATA_VECTOR, 1)

DATA_VECTOR = STRIP_NUMERICS_FUNC(DATA_VECTOR)
MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
SIGMA_VAL = MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1)
ReDim TEMP_VECTOR(1 To NROWS, 1 To 1)
For i = 1 To NROWS
     TEMP_VECTOR(i, 1) = Abs(DATA_VECTOR(i, 1) - MEAN_VAL)
Next i
DELTA_VAL = MATRIX_ELEMENTS_MAX_FUNC(TEMP_VECTOR, 0) / SIGMA_VAL

TEMP_FACTOR = (2 * ALPHA) / (2 * NROWS)
TDIST_VAL = INVERSE_TDIST_FUNC(TEMP_FACTOR / 2, NROWS - 2) * -1
TEMP_VAL = ((NROWS - 1) / (NROWS ^ 0.5)) * ((TDIST_VAL ^ 2) / (NROWS - 2 + TDIST_VAL ^ 2)) ^ 0.5

If DELTA_VAL > TEMP_VAL Then
    GRUBBS_TEST_FUNC = "Outliers"
Else
    GRUBBS_TEST_FUNC = "No Outliers"
End If

Exit Function
ERROR_LABEL:
GRUBBS_TEST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GRUBBS_NEGATIVE_OUTLIERS_TEST_FUNC
'DESCRIPTION   : Computes one-sided Grubbs' test for negative outliers
'LIBRARY       : STATISTICS
'GROUP         : OUTLIERS
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function GRUBBS_NEGATIVE_OUTLIERS_TEST_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal ALPHA As Double = 0.05)

Dim NROWS As Long

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim DELTA_VAL As Double

Dim TEMP_VAL As Double
Dim TEMP_FACTOR As Double

Dim TDIST_VAL As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
NROWS = UBound(DATA_VECTOR, 1)

DATA_VECTOR = STRIP_NUMERICS_FUNC(DATA_VECTOR)
MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
SIGMA_VAL = MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1)
DELTA_VAL = (MEAN_VAL - MATRIX_ELEMENTS_MIN_FUNC(DATA_VECTOR, 0)) / SIGMA_VAL

TEMP_FACTOR = (2 * ALPHA) / (NROWS)
TDIST_VAL = INVERSE_TDIST_FUNC(TEMP_FACTOR / 2, NROWS - 2) * -1
TEMP_VAL = ((NROWS - 1) / (NROWS ^ 0.5)) * ((TDIST_VAL ^ 2) / (NROWS - 2 + TDIST_VAL ^ 2)) ^ 0.5
If DELTA_VAL > TEMP_VAL Then
    GRUBBS_NEGATIVE_OUTLIERS_TEST_FUNC = "Negative Outliers"
Else
    GRUBBS_NEGATIVE_OUTLIERS_TEST_FUNC = "No Negative Outliers"
End If

Exit Function
ERROR_LABEL:
GRUBBS_NEGATIVE_OUTLIERS_TEST_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : GRUBBS_POSITIVE_OUTLIERS_TEST_FUNC
'DESCRIPTION   : Computes one-sided Grubbs' test for positive outliers
'LIBRARY       : STATISTICS
'GROUP         : OUTLIERS
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function GRUBBS_POSITIVE_OUTLIERS_TEST_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal ALPHA As Double = 0.05)

Dim NROWS As Long

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim DELTA_VAL As Double

Dim TEMP_VAL As Double
Dim TEMP_FACTOR As Double

Dim TDIST_VAL As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
NROWS = UBound(DATA_VECTOR, 1)

DATA_VECTOR = STRIP_NUMERICS_FUNC(DATA_VECTOR)
MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
SIGMA_VAL = MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1)
DELTA_VAL = (MATRIX_ELEMENTS_MAX_FUNC(DATA_VECTOR, 0) - MEAN_VAL) / SIGMA_VAL
TEMP_FACTOR = (2 * ALPHA) / (NROWS)

TDIST_VAL = INVERSE_TDIST_FUNC(TEMP_FACTOR / 2, NROWS - 2) * -1

TEMP_VAL = ((NROWS - 1) / (NROWS ^ 0.5)) * ((TDIST_VAL ^ 2) / _
            (NROWS - 2 + TDIST_VAL ^ 2)) ^ 0.5
If DELTA_VAL > TEMP_VAL Then
    GRUBBS_POSITIVE_OUTLIERS_TEST_FUNC = "Positive Outliers"
Else
    GRUBBS_POSITIVE_OUTLIERS_TEST_FUNC = "No Positive Outliers"
End If

Exit Function
ERROR_LABEL:
GRUBBS_POSITIVE_OUTLIERS_TEST_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_BIVAR_ZSCORE_FUNC
'DESCRIPTION   : Bivariate Z-Score Analysis

'LIBRARY       : STATISTICS
'GROUP         : OUTLIERS
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************


Function VECTOR_BIVAR_ZSCORE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal WEIGHT_VAL As Double = 0.5, _
Optional ByVal ZSCORE_FLAG As Boolean = False, _
Optional ByVal MODIFIED_FLAG As Boolean = False)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim ZFACTOR_VAL As Variant

Dim ATEMP_VECTOR As Variant
Dim BTEMP_VECTOR As Variant

Dim CTEMP_VECTOR As Variant
Dim DTEMP_VECTOR As Variant

Dim ETEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
If NCOLUMNS = 1 Then: DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)

ReDim ATEMP_VECTOR(1 To NROWS, 1 To 1)
ReDim BTEMP_VECTOR(1 To NROWS, 1 To 1)

If MODIFIED_FLAG Then
    For i = 1 To NROWS
        ReDim ETEMP_VECTOR(1 To 1, 1 To NCOLUMNS - 1 + 1)
        For j = 0 To NCOLUMNS - 1
            ETEMP_VECTOR(1, j + 1) = DATA_MATRIX(i, 1 + j)
        Next j
        ATEMP_VECTOR(i, 1) = HISTOGRAM_PERCENTILE_FUNC(STRIP_NUMERICS_FUNC(ETEMP_VECTOR), 0.5, 1)
        BTEMP_VECTOR(i, 1) = MATRIX_ABSOLUTE_DEVIATION_FUNC(ETEMP_VECTOR)
    Next i
Else
    For i = 1 To NROWS
        ReDim ETEMP_VECTOR(1 To 1, 1 To NCOLUMNS - 1 + 1)
        For j = 0 To NCOLUMNS - 1
            ETEMP_VECTOR(1, j + 1) = DATA_MATRIX(i, 1 + j)
        Next j
        ATEMP_VECTOR(i, 1) = MATRIX_MEAN_FUNC(STRIP_NUMERICS_FUNC(ETEMP_VECTOR))(1, 1)
        BTEMP_VECTOR(i, 1) = MATRIX_STDEV_FUNC(STRIP_NUMERICS_FUNC(ETEMP_VECTOR))(1, 1)
    Next i
End If

ReDim CTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)
ReDim DTEMP_VECTOR(1 To NCOLUMNS, 1 To 1)

If MODIFIED_FLAG Then
    For i = 1 To NCOLUMNS
        ReDim ETEMP_VECTOR(1 To NROWS - 1 + 1, 1 To 1)
        For j = 0 To NROWS - 1
            ETEMP_VECTOR(j + 1, 1) = DATA_MATRIX(1 + j, i)
        Next j
        CTEMP_VECTOR(i, 1) = HISTOGRAM_PERCENTILE_FUNC(STRIP_NUMERICS_FUNC(ETEMP_VECTOR), 0.5, 1)
        DTEMP_VECTOR(i, 1) = MATRIX_ABSOLUTE_DEVIATION_FUNC(ETEMP_VECTOR)
    Next i
Else
    For i = 1 To NCOLUMNS
        ReDim ETEMP_VECTOR(1 To NROWS - 1 + 1, 1 To 1)
        For j = 0 To NROWS - 1
            ETEMP_VECTOR(j + 1, 1) = DATA_MATRIX(1 + j, i)
        Next j
        CTEMP_VECTOR(i, 1) = MATRIX_MEAN_FUNC(STRIP_NUMERICS_FUNC(ETEMP_VECTOR))(1, 1)
        DTEMP_VECTOR(i, 1) = MATRIX_STDEV_FUNC(STRIP_NUMERICS_FUNC(ETEMP_VECTOR))(1, 1)
    Next i
End If

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

If ZSCORE_FLAG Then
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            If IsNumeric(DATA_MATRIX(i, j)) And Not IsEmpty(DATA_MATRIX(i, j)) Then
                TEMP_MATRIX(i, j) = WEIGHT_VAL * (DATA_MATRIX(i, j) - ATEMP_VECTOR(i, 1)) / BTEMP_VECTOR(i, 1) + (1 - WEIGHT_VAL) * (DATA_MATRIX(i, j) - CTEMP_VECTOR(j, 1)) / DTEMP_VECTOR(j, 1)
            Else
                TEMP_MATRIX(i, j) = CVErr(xlErrNA)
            End If
        Next j
    Next i
    
Else
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            If IsNumeric(DATA_MATRIX(i, j)) And Not IsEmpty(DATA_MATRIX(i, j)) Then
                ZFACTOR_VAL = WEIGHT_VAL * (DATA_MATRIX(i, j) - ATEMP_VECTOR(i, 1)) / BTEMP_VECTOR(i, 1) + (1 - WEIGHT_VAL) * (DATA_MATRIX(i, j) - CTEMP_VECTOR(j, 1)) / DTEMP_VECTOR(j, 1)
                If Abs(ZFACTOR_VAL) > 3 Then
                    TEMP_MATRIX(i, j) = "***"
                Else
                    If Abs(ZFACTOR_VAL) > 2 Then
                        TEMP_MATRIX(i, j) = "**"
                    Else
                        If Abs(ZFACTOR_VAL) > 1 Then
                            TEMP_MATRIX(i, j) = "*"
                        Else
                            TEMP_MATRIX(i, j) = ""
                        End If
                    End If
                End If
            Else
                TEMP_MATRIX(i, j) = CVErr(xlErrNA)
            End If
        Next j
    Next i
End If
    
VECTOR_BIVAR_ZSCORE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
VECTOR_BIVAR_ZSCORE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_BIVAR_ZSCORE_FUNC
'DESCRIPTION   : BIVARIATE GAUSSIAN OUTLIER DETECTION

'If series X and Y are bivariate Gaussian distributed, a confidence region
'in the form an ellipse in X-Y space can be drawn.

'Observations outside the confidence region are unlikely to stem from a
'bivariate Gaussian distribution.

'Interestingly, bivariate outliers are not always also univariate outliers.


'LIBRARY       : STATISTICS
'GROUP         : OUTLIERS
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 14/02/2008
'************************************************************************************
'************************************************************************************

Function MATRIX_GAUSS_OUTLIERS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CONFIDENCE_VAL As Double = 0.9)
'DATA_RNG -->  X /Y

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim PI_VAL As Double

Dim C_VAL As Double
Dim D_VAL As Double
Dim E_VAL As Double
Dim F_VAL As Double

Dim Q_VAL As Double
Dim R_VAL As Double
Dim V_VAL As Double
Dim PHI_VAL As Double

Dim DATA_MATRIX As Variant
Dim MEAN_VECTOR As Variant
Dim VOLATILITY_VECTOR As Variant
Dim INVERSE_MATRIX As Variant
Dim CORRELATION_MATRIX As Variant
Dim COVARIANCE_MATRIX As Variant

On Error GoTo ERROR_LABEL

PI_VAL = 3.14159265358979
DATA_MATRIX = DATA_RNG
C_VAL = INVERSE_CHI_SQUARED_DIST_FUNC(1 - CONFIDENCE_VAL, 2, False) ' critical value
E_VAL = Abs(NORMSINV_FUNC((1 - CONFIDENCE_VAL) / 2, 0, 1, 0))
NROWS = UBound(DATA_MATRIX, 1) ' nObs
NCOLUMNS = UBound(DATA_MATRIX, 2) ' nSeries
F_VAL = 360 / (NROWS - 1) ' angle step

COVARIANCE_MATRIX = MATRIX_COVARIANCE_FRAME3_FUNC(DATA_MATRIX) ' Inverse covariance matrix, correlation matrix
INVERSE_MATRIX = MATRIX_INVERSE_FUNC(COVARIANCE_MATRIX, 2)
CORRELATION_MATRIX = MATRIX_CORRELATION_COVARIANCE_FUNC(COVARIANCE_MATRIX)
MEAN_VECTOR = MATRIX_TRANSPOSE_FUNC(MATRIX_MEAN_FUNC(DATA_MATRIX)) ' Means & Stdevs
VOLATILITY_VECTOR = MATRIX_VOLATILITY_COVARIANCE_FUNC(COVARIANCE_MATRIX)

ReDim TEMP_VECTOR(1 To 1, 1 To NCOLUMNS)
ReDim TEMP_MATRIX(0 To NROWS, 1 To 8)

TEMP_MATRIX(0, 1) = "Bivariate Outlier"
TEMP_MATRIX(0, 2) = ""
TEMP_MATRIX(0, 3) = "Univariate Outlier Marginal X"
TEMP_MATRIX(0, 4) = ""
TEMP_MATRIX(0, 5) = "Univariate Outlier Marginal Y"
TEMP_MATRIX(0, 6) = ""
TEMP_MATRIX(0, 7) = "Confidence Region"
TEMP_MATRIX(0, 8) = ""

For i = 1 To NROWS
    ' get demeaned row vector
    For j = 1 To NCOLUMNS
        TEMP_VECTOR(1, j) = DATA_MATRIX(i, j) - MEAN_VECTOR(j, 1)
    Next j
    ' calculate distance D_VAL
    D_VAL = MMULT_FUNC(MMULT_FUNC(TEMP_VECTOR, INVERSE_MATRIX, 70), MATRIX_TRANSPOSE_FUNC(TEMP_VECTOR), 70)(1, 1)
    ' determine whether bivariate outliers
    If D_VAL > C_VAL Then
        TEMP_MATRIX(i, 1) = DATA_MATRIX(i, 1)
        TEMP_MATRIX(i, 2) = DATA_MATRIX(i, 2)
    Else
        TEMP_MATRIX(i, 1) = CVErr(xlErrNA)
        TEMP_MATRIX(i, 2) = CVErr(xlErrNA)
    End If
    ' determine whether univariate outliers
    If Abs((DATA_MATRIX(i, 1) - MEAN_VECTOR(1, 1)) / VOLATILITY_VECTOR(1, 1)) > E_VAL Then
        TEMP_MATRIX(i, 3) = DATA_MATRIX(i, 1)
        TEMP_MATRIX(i, 4) = DATA_MATRIX(i, 2)
    Else
        TEMP_MATRIX(i, 3) = CVErr(xlErrNA)
        TEMP_MATRIX(i, 4) = CVErr(xlErrNA)
    End If
    If Abs((DATA_MATRIX(i, 2) - MEAN_VECTOR(2, 1)) / VOLATILITY_VECTOR(2, 1)) > E_VAL Then
        TEMP_MATRIX(i, 5) = DATA_MATRIX(i, 1)
        TEMP_MATRIX(i, 6) = DATA_MATRIX(i, 2)
    Else
        TEMP_MATRIX(i, 5) = CVErr(xlErrNA)
        TEMP_MATRIX(i, 6) = CVErr(xlErrNA)
    End If
    
    ' calculate ellipse
    PHI_VAL = -180 + (i - 1) * F_VAL
    Q_VAL = Tan(PHI_VAL * PI_VAL / 180)
    V_VAL = Sqr(C_VAL) / (Sqr(1 / (1 - CORRELATION_MATRIX(1, 2) ^ 2) * (1 / COVARIANCE_MATRIX(1, 1) + Q_VAL ^ 2 / COVARIANCE_MATRIX(2, 2) - 2 * CORRELATION_MATRIX(1, 2) * Q_VAL / Sqr(COVARIANCE_MATRIX(1, 1) * COVARIANCE_MATRIX(2, 2)))))
    R_VAL = V_VAL * Sqr(Q_VAL ^ 2 + 1)
    
    TEMP_MATRIX(i, 7) = R_VAL * Cos(PHI_VAL * PI_VAL / 180) + MEAN_VECTOR(1, 1)
    TEMP_MATRIX(i, 8) = R_VAL * Sin(PHI_VAL * PI_VAL / 180) + MEAN_VECTOR(2, 1)
Next i

MATRIX_GAUSS_OUTLIERS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_GAUSS_OUTLIERS_FUNC = Err.number
End Function
