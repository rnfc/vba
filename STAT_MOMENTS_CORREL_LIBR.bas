Attribute VB_Name = "STAT_MOMENTS_CORREL_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : CORRELATION_UPDATE_FUNC
'DESCRIPTION   : UPDATE CORRELATION BASED ON NEW SPOT PRICES
'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CORRELATION_UPDATE_FUNC(ByVal OLD_CORREL_VAL As Double, _
ByVal OLD_SIGMA1_VAL As Double, _
ByVal OLD_SIGMA2_VAL As Double, _
ByVal OLD_SPOT1_VAL As Double, _
ByVal OLD_SPOT2_VAL As Double, _
ByVal NEW_SPOT1_VAL As Double, _
ByVal NEW_SPOT2_VAL As Double, _
Optional ByVal LAMBDA_VAL As Double = 0.99, _
Optional ByVal DAYS_PER_YEAR_VAL As Double = 250)
 
Dim NEW_SIGMA1_VAL As Double
Dim NEW_SIGMA2_VAL As Double

Dim OLD_COVAR_VAL As Double
Dim NEW_COVAR_VAL As Double

On Error GoTo ERROR_LABEL

NEW_SIGMA1_VAL = UPDATE_VOLATILITY_FUNC(OLD_SIGMA1_VAL, OLD_SPOT1_VAL, NEW_SPOT1_VAL, LAMBDA_VAL, DAYS_PER_YEAR_VAL)
NEW_SIGMA2_VAL = UPDATE_VOLATILITY_FUNC(OLD_SIGMA2_VAL, OLD_SPOT2_VAL, NEW_SPOT2_VAL, LAMBDA_VAL, DAYS_PER_YEAR_VAL)

OLD_COVAR_VAL = OLD_CORREL_VAL * OLD_SIGMA1_VAL * OLD_SIGMA2_VAL
NEW_COVAR_VAL = LAMBDA_VAL * OLD_COVAR_VAL + (1 - LAMBDA_VAL) * Log(NEW_SPOT1_VAL / OLD_SPOT1_VAL) * Log(NEW_SPOT2_VAL / OLD_SPOT2_VAL) * DAYS_PER_YEAR_VAL

If NEW_SIGMA1_VAL > 0 And NEW_SIGMA2_VAL > 0 Then
    CORRELATION_UPDATE_FUNC = NEW_COVAR_VAL / (NEW_SIGMA1_VAL * NEW_SIGMA2_VAL)
Else
    CORRELATION_UPDATE_FUNC = 0
End If

Exit Function
ERROR_LABEL:
CORRELATION_UPDATE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CORRELATION_FUNC
'DESCRIPTION   : Returns Correlation Coefficient between two vectors
'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CORRELATION_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim MEAN1_VAL As Double
Dim MEAN2_VAL As Double

Dim VAR1_VAL As Double
Dim VAR2_VAL As Double

Dim DEV1_VAL As Double
Dim DEV2_VAL As Double

Dim COVAR_VAL As Double

Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA1_VECTOR = DATA1_RNG
If UBound(DATA1_VECTOR, 1) = 1 Then
    DATA1_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA1_VECTOR)
End If
If DATA_TYPE <> 0 Then DATA1_VECTOR = MATRIX_PERCENT_FUNC(DATA1_VECTOR, LOG_SCALE)
DATA2_VECTOR = DATA2_RNG
If UBound(DATA2_VECTOR, 1) = 1 Then
    DATA2_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA2_VECTOR)
End If
If DATA_TYPE <> 0 Then DATA2_VECTOR = MATRIX_PERCENT_FUNC(DATA2_VECTOR, LOG_SCALE)
NROWS = UBound(DATA1_VECTOR, 1)
If UBound(DATA1_VECTOR, 1) < 2 Or UBound(DATA1_VECTOR, 1) <> UBound(DATA2_VECTOR, 1) Then: GoTo ERROR_LABEL
' First pass to determine averages
TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 1 To NROWS
     TEMP1_SUM = TEMP1_SUM + DATA1_VECTOR(i, 1)
     TEMP2_SUM = TEMP2_SUM + DATA2_VECTOR(i, 1)
Next i
MEAN1_VAL = TEMP1_SUM / NROWS
MEAN2_VAL = TEMP2_SUM / NROWS
' Second pass to determine covariance, stdeviations
DEV1_VAL = 0: DEV2_VAL = 0
VAR1_VAL = 0: VAR2_VAL = 0
COVAR_VAL = 0
For i = 1 To NROWS
     DEV1_VAL = DATA1_VECTOR(i, 1) - MEAN1_VAL
     DEV2_VAL = DATA2_VECTOR(i, 1) - MEAN2_VAL
     VAR1_VAL = VAR1_VAL + DEV1_VAL ^ 2
     VAR2_VAL = VAR2_VAL + DEV2_VAL ^ 2
     COVAR_VAL = COVAR_VAL + DEV1_VAL * DEV2_VAL
Next i

TEMP1_SUM = 0: TEMP2_SUM = 0
VAR1_VAL = (VAR1_VAL - TEMP1_SUM * TEMP1_SUM / NROWS) / NROWS 'rounding error corrected var
VAR2_VAL = (VAR2_VAL - TEMP2_SUM * TEMP2_SUM / NROWS) / NROWS 'rounding error corrected var
CORRELATION_FUNC = COVAR_VAL / (NROWS * Sqr(VAR1_VAL) * Sqr(VAR2_VAL)) 'COVAR_VAL / NROWS

Exit Function
ERROR_LABEL:
CORRELATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CORRELATION_SPEARMAN_FUNC
'DESCRIPTION   : Returns the Spearman Correlation between two vectors
'http://www.gummy-stuff.org/spearman-correlation.htm

'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CORRELATION_SPEARMAN_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim MEAN1_VAL As Double
Dim MEAN2_VAL As Double

Dim DEV1_VAL As Double
Dim DEV2_VAL As Double

Dim VAR1_VAL As Double
Dim VAR2_VAL As Double

Dim COVAR_VAL As Double

Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA1_VECTOR = DATA1_RNG
If UBound(DATA1_VECTOR, 1) = 1 Then
    DATA1_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA1_VECTOR)
End If
DATA2_VECTOR = DATA2_RNG
If UBound(DATA2_VECTOR, 1) = 1 Then
    DATA2_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA2_VECTOR)
End If
If DATA_TYPE <> 0 Then DATA1_VECTOR = MATRIX_PERCENT_FUNC(DATA1_VECTOR, LOG_SCALE)
If DATA_TYPE <> 0 Then DATA2_VECTOR = MATRIX_PERCENT_FUNC(DATA2_VECTOR, LOG_SCALE)
NROWS = UBound(DATA1_VECTOR, 1)
If UBound(DATA1_VECTOR, 1) < 2 Or UBound(DATA1_VECTOR, 1) <> UBound(DATA2_VECTOR, 1) Then: GoTo ERROR_LABEL
TEMP1_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA1_VECTOR, 1, 0)
TEMP2_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA2_VECTOR, 1, 0)
For j = 1 To NROWS
    TEMP1_VAL = TEMP1_VECTOR(j, 1)
    TEMP2_VAL = TEMP2_VECTOR(j, 1)
    For i = 1 To NROWS
        If TEMP1_VAL = DATA1_VECTOR(i, 1) Then
             DATA1_VECTOR(i, 1) = j
             Exit For
        End If
    Next i
    For i = 1 To NROWS
        If TEMP2_VAL = DATA2_VECTOR(i, 1) Then
             DATA2_VECTOR(i, 1) = j
             Exit For
        End If
    Next i
Next j
' First pass to determine averages
TEMP1_SUM = 0: TEMP2_SUM = 0
For i = 1 To NROWS
     TEMP1_SUM = TEMP1_SUM + DATA1_VECTOR(i, 1)
     TEMP2_SUM = TEMP2_SUM + DATA2_VECTOR(i, 1)
Next i
MEAN1_VAL = TEMP1_SUM / NROWS
MEAN2_VAL = TEMP2_SUM / NROWS
' Second pass to determine covariance, stdeviations
DEV1_VAL = 0: DEV2_VAL = 0
VAR1_VAL = 0: VAR2_VAL = 0
COVAR_VAL = 0
For i = 1 To NROWS
     DEV1_VAL = DATA1_VECTOR(i, 1) - MEAN1_VAL
     DEV2_VAL = DATA2_VECTOR(i, 1) - MEAN2_VAL
     VAR1_VAL = VAR1_VAL + DEV1_VAL ^ 2
     VAR2_VAL = VAR2_VAL + DEV2_VAL ^ 2
     COVAR_VAL = COVAR_VAL + DEV1_VAL * DEV2_VAL
Next i
TEMP1_SUM = 0: TEMP2_SUM = 0
VAR1_VAL = (VAR1_VAL - TEMP1_SUM * TEMP1_SUM / NROWS) / NROWS
'rounding error corrected var
VAR2_VAL = (VAR2_VAL - TEMP2_SUM * TEMP2_SUM / NROWS) / NROWS
'rounding error corrected var
CORRELATION_SPEARMAN_FUNC = COVAR_VAL / (NROWS * Sqr(VAR1_VAL) * Sqr(VAR2_VAL))
Exit Function
ERROR_LABEL:
CORRELATION_SPEARMAN_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CORRELATION_MOVING_AVERAGE_FUNC
'DESCRIPTION   : MOVING AVERAGE CORRELATION VECTOR
'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CORRELATION_MOVING_AVERAGE_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByVal MA_PERIODS As Long = 3, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant
Dim TEMP3_VECTOR As Variant

Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA1_VECTOR = DATA1_RNG
If UBound(DATA1_VECTOR, 1) = 1 Then
    DATA1_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA1_VECTOR)
End If

DATA2_VECTOR = DATA2_RNG
If UBound(DATA2_VECTOR, 1) = 1 Then
    DATA2_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA2_VECTOR)
End If

If DATA_TYPE <> 0 Then DATA1_VECTOR = MATRIX_PERCENT_FUNC(DATA1_VECTOR, LOG_SCALE)
If DATA_TYPE <> 0 Then DATA2_VECTOR = MATRIX_PERCENT_FUNC(DATA2_VECTOR, LOG_SCALE)

If UBound(DATA1_VECTOR, 1) <> UBound(DATA2_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(DATA1_VECTOR, 1)
If (MA_PERIODS = 1) Or (MA_PERIODS = 0) Or (MA_PERIODS = 2) Then: GoTo ERROR_LABEL
If MA_PERIODS > NROWS Then MA_PERIODS = NROWS

ReDim TEMP3_VECTOR(1 To NROWS - 1, 1 To 1)
For i = 2 To MA_PERIODS
    ReDim TEMP1_VECTOR(1 To i, 1 To 1)
    ReDim TEMP2_VECTOR(1 To i, 1 To 1)
    k = 1
    For j = i To 1 Step -1
        TEMP1_VECTOR(k, 1) = DATA1_VECTOR(j, 1)
        TEMP2_VECTOR(k, 1) = DATA2_VECTOR(j, 1)
        k = k + 1
    Next j
    TEMP3_VECTOR(i - 1, 1) = CORRELATION_FUNC(TEMP1_VECTOR, TEMP2_VECTOR, 0, 0)
Next i

ReDim TEMP1_VECTOR(1 To MA_PERIODS, 1 To 1)
ReDim TEMP2_VECTOR(1 To MA_PERIODS, 1 To 1)

For i = MA_PERIODS + 1 To NROWS
    k = 1
    For j = i To i - MA_PERIODS + 1 Step -1
        TEMP1_VECTOR(k, 1) = DATA1_VECTOR(j, 1)
        TEMP2_VECTOR(k, 1) = DATA2_VECTOR(j, 1)
        k = k + 1
    Next j
    TEMP3_VECTOR(i - 1, 1) = CORRELATION_FUNC(TEMP1_VECTOR, TEMP2_VECTOR, 0, 0)
Next i

CORRELATION_MOVING_AVERAGE_FUNC = TEMP3_VECTOR

Exit Function
ERROR_LABEL:
CORRELATION_MOVING_AVERAGE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CORRELATION_EXCEED_FUNC

'DESCRIPTION   : Compute exceedence correlation with observations to the left
'of a quantile in the analysis of asymmetric /
'non-gaussian / non-linear dependence structures.

'REFERENCES:
'EXTREME CORRELATION OF INTERNATIONAL EQUITY MARKETS:
'    François Longin and Bruno Solnik, Journal of Finance

'Dependence Patterns across Financial Markets: Methods and Evidence, Ling Hu
'On the Out-of-Sample Importance of Skewness and Asymmetric Dependence for
'Asset Allocation, ANDREW J. PATTON, London School of Economics

'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CORRELATION_EXCEED_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByVal QUANTILE As Double = 0.2, _
Optional ByVal THRESHOLD As Long = 8, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

'THRESHOLD: Minimum number of observations

Dim i As Long
Dim k As Long

Dim NROWS As Long

Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant
Dim TEMP3_VECTOR As Variant
Dim TEMP4_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA1_VECTOR = DATA1_RNG
If UBound(DATA1_VECTOR, 1) = 1 Then: DATA1_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA1_VECTOR)

DATA2_VECTOR = DATA2_RNG
If UBound(DATA2_VECTOR, 1) = 1 Then: DATA2_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA2_VECTOR)
If UBound(DATA1_VECTOR, 1) <> UBound(DATA2_VECTOR, 1) Then: GoTo ERROR_LABEL

If DATA_TYPE <> 0 Then: DATA1_VECTOR = MATRIX_PERCENT_FUNC(DATA1_VECTOR, LOG_SCALE)
If DATA_TYPE <> 0 Then: DATA2_VECTOR = MATRIX_PERCENT_FUNC(DATA2_VECTOR, LOG_SCALE)

NROWS = UBound(DATA1_VECTOR, 1)
' determine quantiles - not efficient, but can be used for other purposes
TEMP1_VECTOR = VECTOR_RANK_PERCENTILE_FUNC(DATA1_VECTOR, 1)
TEMP2_VECTOR = VECTOR_RANK_PERCENTILE_FUNC(DATA2_VECTOR, 1)

' count exceeding observations
k = 0
For i = 1 To NROWS
    If QUANTILE <= 0.5 Then
        If TEMP1_VECTOR(i, 1) <= QUANTILE And TEMP2_VECTOR(i, 1) <= QUANTILE Then
            k = k + 1
        End If
    Else
        If TEMP1_VECTOR(i, 1) > QUANTILE And TEMP2_VECTOR(i, 1) > QUANTILE Then
            k = k + 1
        End If
    End If
Next i

If k < 2 Or k < THRESHOLD Then
    CORRELATION_EXCEED_FUNC = "N/A"
    Exit Function
End If

' calculate correlation
ReDim TEMP3_VECTOR(1 To k, 1 To 1)
ReDim TEMP4_VECTOR(1 To k, 1 To 1)

k = 1
For i = 1 To NROWS
    If QUANTILE <= 0.5 Then
        If TEMP1_VECTOR(i, 1) <= QUANTILE And TEMP2_VECTOR(i, 1) <= QUANTILE Then
            TEMP3_VECTOR(k, 1) = DATA1_VECTOR(i, 1)
            TEMP4_VECTOR(k, 1) = DATA2_VECTOR(i, 1)
            k = k + 1
        End If
    Else
        If TEMP1_VECTOR(i, 1) > QUANTILE And TEMP2_VECTOR(i, 1) > QUANTILE Then
            TEMP3_VECTOR(k, 1) = DATA1_VECTOR(i, 1)
            TEMP4_VECTOR(k, 1) = DATA2_VECTOR(i, 1)
            k = k + 1
        End If
    End If
Next i

CORRELATION_EXCEED_FUNC = CORRELATION_FUNC(TEMP3_VECTOR, TEMP4_VECTOR, 0, 0)

Exit Function
ERROR_LABEL:
CORRELATION_EXCEED_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CORRELATION_EXCEED_QUANTILE_FUNC

'DESCRIPTION   : Exceedance Correlations in the analysis of asymmetric /
'non-gaussian / non-linear dependence structures.

'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CORRELATION_EXCEED_QUANTILE_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByVal THRESHOLD As Long = 8, _
Optional ByVal MIN_QUANT As Double = 0.15, _
Optional ByVal MAX_QUANT As Double = 0.85, _
Optional ByVal DELTA_QUANT As Double = 0.025, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

'THRESHOLD: Minimum number of observations

Dim i As Long
Dim NROWS As Long

Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA1_VECTOR = DATA1_RNG
If UBound(DATA1_VECTOR, 1) = 1 Then: DATA1_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA1_VECTOR)

DATA2_VECTOR = DATA2_RNG
If UBound(DATA2_VECTOR, 1) = 1 Then: DATA2_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA2_VECTOR)

If UBound(DATA1_VECTOR, 1) <> UBound(DATA2_VECTOR, 1) Then: GoTo ERROR_LABEL

If DATA_TYPE <> 0 Then: DATA1_VECTOR = MATRIX_PERCENT_FUNC(DATA1_VECTOR, LOG_SCALE)
If DATA_TYPE <> 0 Then: DATA2_VECTOR = MATRIX_PERCENT_FUNC(DATA2_VECTOR, LOG_SCALE)

NROWS = Int((MAX_QUANT - MIN_QUANT) / DELTA_QUANT) + 1

ReDim TEMP_MATRIX(0 To NROWS, 1 To 3)

TEMP_MATRIX(0, 1) = "QUANTILE"
TEMP_MATRIX(0, 2) = "EXCEED CORREL"
TEMP_MATRIX(0, 3) = "NO OBS"

For i = 1 To NROWS
    If i = 1 Then
        TEMP_MATRIX(i, 1) = MIN_QUANT
    Else
        TEMP_MATRIX(i, 1) = TEMP_MATRIX(i - 1, 1) + DELTA_QUANT
    End If
    TEMP_MATRIX(i, 1) = Round(TEMP_MATRIX(i, 1), 3)
    TEMP_MATRIX(i, 2) = CORRELATION_EXCEED_FUNC(DATA1_VECTOR, DATA2_VECTOR, TEMP_MATRIX(i, 1), THRESHOLD, 0, 0)
    TEMP_MATRIX(i, 3) = CORRELATION_EXCEED_COUNT_FUNC(DATA1_VECTOR, DATA2_VECTOR, TEMP_MATRIX(i, 1), 0, 0)
Next i

CORRELATION_EXCEED_QUANTILE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
CORRELATION_EXCEED_QUANTILE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : CORRELATION_EXCEED_COUNT_FUNC
'DESCRIPTION   : Count observations to the left of a quantile
'LIBRARY       : STATISTICS
'GROUP         : COUNT
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function CORRELATION_EXCEED_COUNT_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByVal QUANTILE As Double = 0.25, _
Optional ByVal DATA_TYPE As Variant = 0, _
Optional ByVal LOG_SCALE As Variant = 0)

Dim i As Long
Dim k As Long

Dim NROWS As Long

Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA1_VECTOR = DATA1_RNG
If UBound(DATA1_VECTOR, 1) = 1 Then
    DATA1_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA1_VECTOR)
End If

DATA2_VECTOR = DATA2_RNG
If UBound(DATA2_VECTOR, 1) = 1 Then
    DATA2_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA2_VECTOR)
End If

If DATA_TYPE <> 0 Then: DATA1_VECTOR = MATRIX_PERCENT_FUNC(DATA1_VECTOR, LOG_SCALE)
If DATA_TYPE <> 0 Then: DATA2_VECTOR = MATRIX_PERCENT_FUNC(DATA2_VECTOR, LOG_SCALE)

NROWS = UBound(DATA1_VECTOR, 1)

' determine quantiles - not efficient, but can be used for other purposes
DATA1_VECTOR = VECTOR_RANK_PERCENTILE_FUNC(DATA1_VECTOR, 1)
DATA2_VECTOR = VECTOR_RANK_PERCENTILE_FUNC(DATA2_VECTOR, 1)

' count exceeding observations
k = 0
For i = 1 To NROWS
    If QUANTILE <= 0.5 Then
        If DATA1_VECTOR(i, 1) <= QUANTILE And DATA2_VECTOR(i, 1) <= QUANTILE Then
            k = k + 1
        End If
    Else
        If DATA1_VECTOR(i, 1) > QUANTILE And DATA2_VECTOR(i, 1) > QUANTILE Then
            k = k + 1
        End If
    End If
Next i

CORRELATION_EXCEED_COUNT_FUNC = k

Exit Function
ERROR_LABEL:
CORRELATION_EXCEED_COUNT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CORRELATION_KENDALL_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CORRELATION_KENDALL_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim TEMP_SUM As Double

Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA1_VECTOR = DATA1_RNG
If UBound(DATA1_VECTOR, 1) = 1 Then: DATA1_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA1_VECTOR)

DATA2_VECTOR = DATA2_RNG
If UBound(DATA2_VECTOR, 1) = 1 Then: DATA2_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA2_VECTOR)

If UBound(DATA1_VECTOR, 1) <> UBound(DATA2_VECTOR, 1) Then: GoTo ERROR_LABEL

If DATA_TYPE <> 0 Then: DATA1_VECTOR = MATRIX_PERCENT_FUNC(DATA1_VECTOR, LOG_SCALE)
If DATA_TYPE <> 0 Then: DATA2_VECTOR = MATRIX_PERCENT_FUNC(DATA2_VECTOR, LOG_SCALE)

NROWS = UBound(DATA1_VECTOR, 1)

For i = 1 To NROWS
    For j = i To NROWS
        If j > i Then
            TEMP_SUM = TEMP_SUM + Sgn((DATA1_VECTOR(i, 1) - DATA1_VECTOR(j, 1)) * (DATA2_VECTOR(i, 1) - DATA2_VECTOR(j, 1)))
        End If
    Next j
Next i

CORRELATION_KENDALL_FUNC = (COMBINATIONS_FUNC(NROWS, 2) ^ -1) * TEMP_SUM

Exit Function
ERROR_LABEL:
CORRELATION_KENDALL_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CORRELATION_PEARSON_MOMENT_FUNC
'DESCRIPTION   :


'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 009
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CORRELATION_PEARSON_MOMENT_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim MEAN1_VAL As Double
Dim MEAN2_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA1_VECTOR = DATA1_RNG
If UBound(DATA1_VECTOR, 1) = 1 Then
    DATA1_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA1_VECTOR)
End If
DATA2_VECTOR = DATA2_RNG
If UBound(DATA2_VECTOR, 1) = 1 Then
    DATA2_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA2_VECTOR)
End If
If UBound(DATA1_VECTOR, 1) <> UBound(DATA2_VECTOR, 1) Then: GoTo ERROR_LABEL
NROWS = UBound(DATA1_VECTOR, 1)

'------------------------------------------------------------------------------------------
Select Case VERSION
'------------------------------------------------------------------------------------------
Case 0
'------------------------------------------------------------------------------------------
    For i = 1 To NROWS
        TEMP1_SUM = TEMP1_SUM + DATA1_VECTOR(i, 1) * DATA1_VECTOR(i, 1)
        TEMP2_SUM = TEMP2_SUM + DATA2_VECTOR(i, 1) * DATA2_VECTOR(i, 1)
        TEMP3_SUM = TEMP3_SUM + DATA1_VECTOR(i, 1) * DATA2_VECTOR(i, 1)
    Next i
    CORRELATION_PEARSON_MOMENT_FUNC = TEMP3_SUM / (TEMP1_SUM) ^ 0.5 / (TEMP2_SUM) ^ 0.5
'------------------------------------------------------------------------------------------
Case Else 'Pearson’s Product Moment Correlation Coefficient measures the direction and
'strength of the linear relationship between two variables
'------------------------------------------------------------------------------------------
    TEMP1_SUM = 0: TEMP2_SUM = 0
    For i = 1 To NROWS
        TEMP1_SUM = TEMP1_SUM + DATA1_VECTOR(i, 1)
        TEMP2_SUM = TEMP2_SUM + DATA2_VECTOR(i, 1)
    Next i
    MEAN1_VAL = TEMP1_SUM / NROWS
    MEAN2_VAL = TEMP2_SUM / NROWS
    TEMP1_SUM = 0: TEMP2_SUM = 0: TEMP3_SUM = 0
    For i = 1 To NROWS
        TEMP1_SUM = TEMP1_SUM + ((DATA1_VECTOR(i, 1) - MEAN1_VAL) * (DATA2_VECTOR(i, 1) - MEAN2_VAL))
        TEMP2_SUM = TEMP2_SUM + (DATA1_VECTOR(i, 1) - MEAN1_VAL) ^ 2
        TEMP3_SUM = TEMP3_SUM + (DATA2_VECTOR(i, 1) - MEAN2_VAL) ^ 2
    Next i
    CORRELATION_PEARSON_MOMENT_FUNC = TEMP1_SUM / (TEMP2_SUM * TEMP3_SUM) ^ 0.5
    'One of the advantages of Pearson’s Product Moment Correlation Coefficient is that it is a unitless
    'quantity that falls between –1 and +1.
'------------------------------------------------------------------------------------------
End Select
'------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
CORRELATION_PEARSON_MOMENT_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CORRELATION_FUNC
'DESCRIPTION   : Returns the correlation matrix
'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 010
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CORRELATION_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim tolerance As Double

On Error GoTo ERROR_LABEL

tolerance = 10 ^ -10
DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP1_VECTOR(1 To NCOLUMNS) 'average for each column
ReDim TEMP2_VECTOR(1 To NCOLUMNS) 'standard deviation for each column
ReDim TEMP_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS) 'compute average for column

For j = 1 To NCOLUMNS
    TEMP_SUM = 0
    For i = 1 To NROWS
        TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, j)
    Next i
    TEMP1_VECTOR(j) = TEMP_SUM / NROWS
Next j

'compute standard deviation for column
For j = 1 To NCOLUMNS
    TEMP_SUM = 0
    For i = 1 To NROWS
        TEMP_SUM = TEMP_SUM + (DATA_MATRIX(i, j) - TEMP1_VECTOR(j)) ^ 2
    Next i
    If TEMP_SUM <> 0 Then
        TEMP2_VECTOR(j) = Sqr(TEMP_SUM / NROWS)
    Else
        TEMP2_VECTOR(j) = tolerance
    End If
Next j
'normalize matrix
For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        DATA_MATRIX(i, j) = (DATA_MATRIX(i, j) - TEMP1_VECTOR(j)) / TEMP2_VECTOR(j)
    Next i
Next j
'compute the cross covariance matrix
For i = 1 To NCOLUMNS
    For j = 1 To NCOLUMNS
        If j < i Then
            TEMP_MATRIX(i, j) = TEMP_MATRIX(j, i)
        Else
            TEMP_SUM = 0
            For k = 1 To NROWS
                TEMP_SUM = TEMP_SUM + DATA_MATRIX(k, i) * DATA_MATRIX(k, j)
            Next k
            TEMP_MATRIX(i, j) = TEMP_SUM / NROWS
        End If
    Next j
Next i

MATRIX_CORRELATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_CORRELATION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PEARSON_FUNC
'DESCRIPTION   : RETURNS A MATRIX WITH THE PEARSON CORRELATION
'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CORRELATION_PEARSON_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)
 
'DATA TYPE = 0 using SPOT VALUES
'DATA TYPE = 0 using % CHANGE VALUES

Dim j As Long
Dim i As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MEAN1_VAL As Double
Dim MEAN2_VAL As Double

Dim STDEVP1_VAL As Double
Dim STDEVP2_VAL As Double

Dim DEV1_VAL As Double
Dim DEV2_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim TEMP_MATRIX As Variant

Dim DATA_MATRIX As Variant
Dim CORRELATION_MATRIX As Variant

On Error GoTo ERROR_LABEL

If DATA_TYPE = 1 Then
Select Case LOG_SCALE
    Case 0
        DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_RNG, 0)
    Case 1
        DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_RNG, 1)
    End Select
Else
    DATA_MATRIX = DATA_RNG
End If

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
    
ReDim CORRELATION_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)

For i = 1 To NCOLUMNS
    For j = 1 To i
    
        ReDim TEMP_MATRIX(1 To NROWS, 1 To 2)
        For k = 1 To NROWS
            TEMP_MATRIX(k, 1) = DATA_MATRIX(k, i)
            TEMP_MATRIX(k, 2) = DATA_MATRIX(k, j)
        Next k
         
        'calculate the mean and remove it
        For k = 1 To NROWS
            TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(k, 1)
            TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(k, 2)
        Next k
        MEAN1_VAL = TEMP1_SUM / (k - 1)
        MEAN2_VAL = TEMP2_SUM / (k - 1)
        TEMP1_SUM = 0
        TEMP2_SUM = 0
    
        For k = 1 To NROWS
             TEMP_MATRIX(k, 1) = TEMP_MATRIX(k, 1) - MEAN1_VAL
             DEV1_VAL = DEV1_VAL + TEMP_MATRIX(k, 1)
             TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(k, 1) ^ 2
             
             TEMP_MATRIX(k, 2) = TEMP_MATRIX(k, 2) - MEAN2_VAL
             DEV2_VAL = DEV2_VAL + TEMP_MATRIX(k, 2)
             TEMP2_SUM = TEMP2_SUM + (TEMP_MATRIX(k, 2)) ^ 2
        Next k
        
        STDEVP1_VAL = Sqr((TEMP1_SUM - DEV1_VAL * DEV1_VAL / k) / (k - 1))
        STDEVP2_VAL = Sqr((TEMP2_SUM - DEV2_VAL * DEV2_VAL / k) / (k - 1))
                                
        'Calculate PEARSON_FUNC
        For k = 1 To NROWS
            TEMP3_SUM = TEMP_MATRIX(k, 1) * TEMP_MATRIX(k, 2) + TEMP3_SUM
        Next k
        
        CORRELATION_MATRIX(i, j) = TEMP3_SUM / ((k - 1) * STDEVP1_VAL * STDEVP2_VAL)
        CORRELATION_MATRIX(j, i) = CORRELATION_MATRIX(i, j)
        
        TEMP1_SUM = 0
        TEMP2_SUM = 0
        TEMP3_SUM = 0
    Next j
Next i

MATRIX_CORRELATION_PEARSON_FUNC = CORRELATION_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_CORRELATION_PEARSON_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CORRELATION_RANK_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CORRELATION_RANK_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)

For i = 1 To NCOLUMNS
    TEMP_MATRIX(i, i) = 1
    For j = i + 1 To NCOLUMNS
        TEMP_MATRIX(i, j) = CORRELATION_SPEARMAN_FUNC(MATRIX_GET_COLUMN_FUNC(DATA_MATRIX, i, 1), MATRIX_GET_COLUMN_FUNC(DATA_MATRIX, j, 1), 0, 0)
        TEMP_MATRIX(j, i) = TEMP_MATRIX(i, j)
    Next j
Next i

MATRIX_CORRELATION_RANK_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_CORRELATION_RANK_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CORRELATION_SHRINK_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 013
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CORRELATION_SHRINK_FUNC(ByRef CORRELATION_RNG As Variant, _
ByVal SHRINKAGE_FACTOR As Double)

Dim i As Long
Dim j As Long
Dim NSIZE As Long

Dim PAIRWISE_VAL As Double
Dim CORRELATION_MATRIX As Variant

On Error GoTo ERROR_LABEL

CORRELATION_MATRIX = CORRELATION_RNG
If UBound(CORRELATION_MATRIX, 1) <> UBound(CORRELATION_MATRIX, 2) Then: GoTo ERROR_LABEL
PAIRWISE_VAL = MATRIX_CORRELATION_AVERAGE_PAIRWISE_COVARIANCE_FUNC(CORRELATION_MATRIX)
NSIZE = UBound(CORRELATION_MATRIX, 2)
For i = 1 To NSIZE
    For j = 1 To NSIZE
        If i <> j Then
            CORRELATION_MATRIX(i, j) = CORRELATION_MATRIX(i, j) + SHRINKAGE_FACTOR * (PAIRWISE_VAL - CORRELATION_MATRIX(i, j))
        End If
    Next j
Next i
MATRIX_CORRELATION_SHRINK_FUNC = CORRELATION_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_CORRELATION_SHRINK_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CORRELATION_SHRINK_FUNC
'DESCRIPTION   : Compute correlation coefficient matrix from covariance matrix
'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 014
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CORRELATION_COVARIANCE_FUNC(ByRef COVARIANCE_RNG As Variant)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim FACTOR_VAL As Double
Dim COVARIANCE_MATRIX As Variant

On Error GoTo ERROR_LABEL

COVARIANCE_MATRIX = COVARIANCE_RNG
NROWS = UBound(COVARIANCE_MATRIX, 1)
NCOLUMNS = UBound(COVARIANCE_MATRIX, 2)

If NROWS <> NCOLUMNS Then: GoTo ERROR_LABEL

ReDim TEMP_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)

For i = 1 To NCOLUMNS
    For j = 1 To NCOLUMNS
        If IsNumeric(COVARIANCE_MATRIX(i, i)) And Not IsEmpty(COVARIANCE_MATRIX(i, i)) And _
           IsNumeric(COVARIANCE_MATRIX(j, j)) And Not IsEmpty(COVARIANCE_MATRIX(j, j)) Then
            FACTOR_VAL = COVARIANCE_MATRIX(i, i) * COVARIANCE_MATRIX(j, j)
            If FACTOR_VAL > 0 Then
                TEMP_MATRIX(i, j) = COVARIANCE_MATRIX(i, j) / (FACTOR_VAL ^ 0.5)
            Else
                TEMP_MATRIX(i, j) = 0
            End If
        Else
            TEMP_MATRIX(i, j) = "N/A"
        End If
    Next j
Next i

MATRIX_CORRELATION_COVARIANCE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_CORRELATION_COVARIANCE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CORRELATION_AVERAGE_PAIRWISE_FUNC

'DESCRIPTION   : Computes average pairwise correlation coefficient from a
'matrix of returns

'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 015
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CORRELATION_AVERAGE_PAIRWISE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)
DATA_MATRIX = MATRIX_CORRELATION_COVARIANCE_FUNC(MATRIX_COVARIANCE_FRAME3_FUNC(DATA_MATRIX, 0, 0))

NCOLUMNS = UBound(DATA_MATRIX, 2)
TEMP_SUM = 0
For i = 1 To NCOLUMNS - 1
    For j = i + 1 To NCOLUMNS
        TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, j)
    Next j
Next i
TEMP_SUM = TEMP_SUM * 2 / (NCOLUMNS * (NCOLUMNS - 1))

MATRIX_CORRELATION_AVERAGE_PAIRWISE_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
MATRIX_CORRELATION_AVERAGE_PAIRWISE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CORRELATION_AVERAGE_PAIRWISE_COVARIANCE_FUNC
'DESCRIPTION   : Computes average pairwise correlation coefficient from a
'covariance matrix
'LIBRARY       : STATISTICS
'GROUP         : COVARIANCE
'ID            : 016
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CORRELATION_AVERAGE_PAIRWISE_COVARIANCE_FUNC( _
ByRef COVARIANCE_RNG As Variant)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long
Dim FACTOR_VAL As Double
Dim TEMP_SUM As Double
Dim CORRELATION_MATRIX As Variant
Dim COVARIANCE_MATRIX As Variant

On Error GoTo ERROR_LABEL

COVARIANCE_MATRIX = COVARIANCE_RNG

NROWS = UBound(CORRELATION_MATRIX, 1)
NCOLUMNS = UBound(CORRELATION_MATRIX, 2)

If NROWS <> NCOLUMNS Then: GoTo ERROR_LABEL

ReDim CORRELATION_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
' compute correlation coefficient matrix from covariance matrix

For i = 1 To NCOLUMNS
    For j = 1 To NCOLUMNS
        FACTOR_VAL = COVARIANCE_MATRIX(i, i) * COVARIANCE_MATRIX(j, j)
        If FACTOR_VAL > 0 Then
            CORRELATION_MATRIX(i, j) = COVARIANCE_MATRIX(i, j) / (FACTOR_VAL ^ 0.5)
        Else
            CORRELATION_MATRIX(i, j) = 0
        End If
    Next j
Next i

TEMP_SUM = 0
For i = 1 To NCOLUMNS - 1
    For j = i + 1 To NCOLUMNS
        TEMP_SUM = TEMP_SUM + CORRELATION_MATRIX(i, j)
    Next j
Next i

MATRIX_CORRELATION_AVERAGE_PAIRWISE_COVARIANCE_FUNC = TEMP_SUM * 2 / (NCOLUMNS * (NCOLUMNS - 1))

Exit Function
ERROR_LABEL:
MATRIX_CORRELATION_AVERAGE_PAIRWISE_COVARIANCE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CORRELATION_AVERAGE_PAIRWISE_COVARIANCE2_FUNC

'DESCRIPTION   : Computes average pairwise correlation coefficient
' matrix from covariance matrix. This is a very basic estimator for
' generating robust inputs in mean/variance optimization.

'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 017
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CORRELATION_AVERAGE_PAIRWISE_COVARIANCE2_FUNC( _
ByRef CORREL_RNG As Variant, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long

Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim MEAN_VAL As Double

Dim TEMP_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL
    
DATA_MATRIX = CORREL_RNG
If UBound(DATA_MATRIX, 1) <> UBound(DATA_MATRIX, 2) Then: GoTo ERROR_LABEL
NCOLUMNS = UBound(DATA_MATRIX, 1)

'------------------------------------------------------------------------------
If VERSION = 0 Then
'------------------------------------------------------------------------------
    TEMP_SUM = 0
    For i = 1 To NCOLUMNS - 1
        For j = i + 1 To NCOLUMNS
            TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, j)
        Next j
    Next i
    MEAN_VAL = TEMP_SUM * 2 / (NCOLUMNS * (NCOLUMNS - 1))
    For i = 1 To NCOLUMNS
        For j = 1 To NCOLUMNS
            If i <> j Then
                DATA_MATRIX(i, j) = MEAN_VAL
            End If
        Next j
    Next i
'------------------------------------------------------------------------------
    MATRIX_CORRELATION_AVERAGE_PAIRWISE_COVARIANCE2_FUNC = DATA_MATRIX
    Exit Function
'------------------------------------------------------------------------------
Else
'------------------------------------------------------------------------------
    ReDim TEMP_VECTOR(1 To 4, 1 To NCOLUMNS)

    TEMP_VECTOR(1, 1) = "COLUMN SUM OF PAIRWISE CORREL"
    TEMP_VECTOR(2, 1) = "TOTAL SUM OF PAIRWISE CORREL"
    TEMP_VECTOR(3, 1) = "NUMBER OF PAIRWISE CORREL"
    TEMP_VECTOR(4, 1) = "AVERAGE PAIRWISE CORREL"
    
    For j = 1 To NCOLUMNS - 1
        TEMP_SUM = 0
        For i = j + 1 To NCOLUMNS
            TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, j)
        Next i
        TEMP_VECTOR(1, j + 1) = TEMP_SUM
        MEAN_VAL = MEAN_VAL + TEMP_VECTOR(1, j + 1)
        TEMP_VECTOR(2, j + 1) = ""
        TEMP_VECTOR(3, j + 1) = ""
        TEMP_VECTOR(4, j + 1) = ""
    Next j
    
    TEMP_VECTOR(2, 2) = MEAN_VAL
    TEMP_VECTOR(3, 2) = NCOLUMNS * (NCOLUMNS - 1) / 2
    TEMP_VECTOR(4, 2) = TEMP_VECTOR(2, 2) / TEMP_VECTOR(3, 2)

    MATRIX_CORRELATION_AVERAGE_PAIRWISE_COVARIANCE2_FUNC = TEMP_VECTOR
'------------------------------------------------------------------------------
End If
'------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
MATRIX_CORRELATION_AVERAGE_PAIRWISE_COVARIANCE2_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_CORRELATION_VALIDATE_FUNC

'DESCRIPTION   : Given lower and upper bounds on each correlation value, this
'routine generates a valid correlation matrix. This is useful for correlation
'stress testing / scenario analysis.

'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 018
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_CORRELATION_VALIDATE_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Const nLOOPS As Long = 10000
Dim NSIZE As Long
Dim TEMP_SUM As Double

Dim DATA_MATRIX As Variant

Dim TEMP1_MATRIX() As Double
Dim TEMP2_MATRIX() As Double
Dim ERROR_FLAG As Boolean

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG 'Restriction Matrix
'…the upper diagonal matrix contains the upper bounds
'…the lower diagonal matrix contains the lower bounds

NSIZE = UBound(DATA_MATRIX, 1)

ReDim TEMP1_MATRIX(1 To NSIZE, 1 To NSIZE)

l = 0
'------------------------------------------------------------------------------
Do
'------------------------------------------------------------------------------
    
    For i = 1 To NSIZE
        For j = i To NSIZE
            If i = j Then
                TEMP1_MATRIX(i, i) = 1
            Else
                If DATA_MATRIX(j, i) > DATA_MATRIX(i, j) Then: GoTo ERROR_LABEL
                    'MATRIX_CORRELATION_VALIDATE_FUNC = CVErr(xlErrNA)
                    'Exit Function
                'End If
                TEMP1_MATRIX(i, j) = DATA_MATRIX(j, i) + (DATA_MATRIX(i, j) - DATA_MATRIX(j, i)) * Rnd()
                TEMP1_MATRIX(j, i) = TEMP1_MATRIX(i, j)
            End If
        Next j
    Next i

    '---------------------------------Cholesky-----------------------------------
    ReDim TEMP2_MATRIX(1 To NSIZE, 1 To NSIZE)
    
    For j = 1 To NSIZE
        TEMP_SUM = 0
        For k = 1 To j - 1
            TEMP_SUM = TEMP_SUM + TEMP2_MATRIX(j, k) ^ 2
        Next k
        TEMP2_MATRIX(j, j) = TEMP1_MATRIX(j, j) - TEMP_SUM
        ' Matrix is not semi-positive definite, no solution exists
        If TEMP2_MATRIX(j, j) < 0 Then
            ERROR_FLAG = True
            GoTo 1983
        Else
            ERROR_FLAG = False
        End If
        TEMP2_MATRIX(j, j) = Sqr(TEMP2_MATRIX(j, j))
        For i = j + 1 To NSIZE
            TEMP_SUM = 0
            For k = 1 To j - 1
                TEMP_SUM = TEMP_SUM + TEMP2_MATRIX(i, k) * TEMP2_MATRIX(j, k)
            Next k
            If TEMP2_MATRIX(j, j) = 0 Then
               TEMP2_MATRIX(i, j) = 0
            Else
               TEMP2_MATRIX(i, j) = (TEMP1_MATRIX(i, j) - TEMP_SUM) / TEMP2_MATRIX(j, j)
            End If
        Next i
    Next j
1983:

'------------------------------------------------------------------------------
    If ERROR_FLAG = True Then
        l = l + 1
        If l > nLOOPS Then: Exit Do
    End If
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
Loop Until ERROR_FLAG = False
'------------------------------------------------------------------------------

'The simulated valid correlation matrix with cmin(i,j) < c(i,j) <cmax(i,j)
'Remember to check the determinant of the simulated correlation matrix
'------------------------------------------------------------------------------
MATRIX_CORRELATION_VALIDATE_FUNC = TEMP1_MATRIX
'You can also perform the Cholesky decomposition of the simulated correlation matrix

Exit Function
ERROR_LABEL:
MATRIX_CORRELATION_VALIDATE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : CORRELATION_PROBABILITY_PLOT_FUNC
'DESCRIPTION   : Assessing fit with a normal distribution with the help of
'the probability plot correlation coefficient.
'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 019
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function CORRELATION_PROBABILITY_PLOT_FUNC(ByRef DATA_RNG As Variant)

'One simple method is to look at the histogram or - a bit more refined - at
'the "probability plot", which plots the empirical probabilties against
'the theortical probabilties. Connecting the dots should result in a linear
'relationship. The better the linear fit, the more likely that the observations
'are from a normal distribution. Looking at the charts usually says a lot more
'about the distribution (presence or absence of fat tails, skewness etc.). To
'quantify the "degree of normality", one can calculate the correlation of the
'data in a probability plot, the so called probability plot correlation (PPC)

'Probability plots have the advantage that observed values can be tested against
'any other distributional assumption. Further, the results are rather easy to
'communicate and can be summarized with an already well known measure, the
'correlation coefficient. It's also easy to rank the degree of normality of
'comparable return series (for example a portfolio versus a benchmark,
'peers etc.)

Dim i As Long
Dim NROWS As Long

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

NROWS = UBound(DATA_VECTOR, 1)
DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
MEAN_VAL = MATRIX_MEAN_FUNC(DATA_VECTOR)(1, 1)
SIGMA_VAL = MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1)

ReDim TEMP1_VECTOR(1 To NROWS, 1 To 1)
ReDim TEMP2_VECTOR(1 To NROWS, 1 To 1)

For i = 1 To NROWS
    TEMP1_VECTOR(i, 1) = NORMSDIST_FUNC(DATA_VECTOR(i, 1), MEAN_VAL, SIGMA_VAL, 0)
    TEMP2_VECTOR(i, 1) = i / NROWS
Next i

CORRELATION_PROBABILITY_PLOT_FUNC = CORRELATION_FUNC(TEMP1_VECTOR, TEMP2_VECTOR, 0, 0)

Exit Function
ERROR_LABEL:
CORRELATION_PROBABILITY_PLOT_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_CORRELATION_VOLATILITY_FUNC

'DESCRIPTION   : Returns a diagonal matrix with the correlation outputs
' of the simulation

'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 020
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function RNG_CORRELATION_VOLATILITY_FUNC(ByRef STARTING_POS As Excel.Range, _
ByRef SIMUL_RNG As Variant)

Dim j As Long
Dim k As Long

Dim NSIZE As Long

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

RNG_CORRELATION_VOLATILITY_FUNC = False

TEMP_MATRIX = SIMUL_RNG
NSIZE = UBound(TEMP_MATRIX, 2)

STARTING_POS.Offset(-1, 0) = "Sigma-Correlation Matrix of Simulated Data"

For j = 1 To NSIZE
    STARTING_POS.Offset(j, j) = MATRIX_STDEV_FUNC(MATRIX_GET_COLUMN_FUNC(TEMP_MATRIX, j, 1))
    For k = j + 1 To NSIZE
        STARTING_POS.Offset(j, k) = CORRELATION_FUNC(MATRIX_GET_COLUMN_FUNC(TEMP_MATRIX, j, 1), MATRIX_GET_COLUMN_FUNC(TEMP_MATRIX, k, 1))
    Next k
Next j

RNG_CORRELATION_VOLATILITY_FUNC = True

Exit Function
ERROR_LABEL:
RNG_CORRELATION_VOLATILITY_FUNC = False
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : RNG_CORRELATION_COVARIANCE_FUNC
'DESCRIPTION   : SetUp Correlation, Covariance and CHOLESKY Matrix
'LIBRARY       : STATISTICS
'GROUP         : CORRELATION
'ID            : 021
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function RNG_CORRELATION_COVARIANCE_FUNC(ByRef DST_RNG As Excel.Range, _
ByVal NASSETS As Long, _
Optional ByVal ADD_RNG_NAME As Boolean = False)

Dim i As Long
Dim j As Long
Dim k As Long

Dim CORREL_POS_RNG As Excel.Range
Dim COVAR_POS_RNG As Excel.Range
Dim CHOL_POS_RNG As Excel.Range

Dim CORREL_RNG As Excel.Range
Dim COVAR_RNG As Excel.Range
Dim CHOL_RNG As Excel.Range

On Error GoTo ERROR_LABEL

RNG_CORRELATION_COVARIANCE_FUNC = False

k = 4
Set CORREL_POS_RNG = DST_RNG

'-------------------------SETTING UP CORRELATION MATRIX-------------
With CORREL_POS_RNG
    Set CORREL_RNG = Range(.Offset(NASSETS, 1), .Offset(1, NASSETS))
    If ADD_RNG_NAME = True Then: CORREL_RNG.name = "CORREL_MAT"
    For i = 1 To NASSETS
        With .Offset(0, i)
           .value = "Asset " & CStr(i)
           .Font.ColorIndex = 3
        End With
        With .Offset(i, 0)
            .formula = "=offset(" & CORREL_POS_RNG.Address & _
            ",0," & CStr(i) & ")"
        End With
    Next i
    With .Offset(-1, 0)
        .value = "SIGMA-CORRELATION MATRIX"
        .Font.Bold = True
    End With
End With
With CORREL_RNG
    For i = 1 To NASSETS 'NAssets - 1
         For j = i To NASSETS  'For j = i + 1
             With .Cells(i, j)
                 .value = 0
                 .Font.ColorIndex = 5
             End With
         Next j
         .Cells(i, i) = 1
    Next i
    For i = 2 To NASSETS
         For j = 1 To i - 1
             With .Cells(i, j)
                 .formula = "=offset(" & CORREL_POS_RNG.Address & "," & CStr(j) & "," & CStr(i) & ")"
             End With
         Next j
    Next i
End With
'---------------------------SETTING UP COVARIANCE MATRIX-------------------------------
Set COVAR_POS_RNG = CORREL_POS_RNG.Offset(NASSETS + k, 0)

With COVAR_POS_RNG
    Set COVAR_RNG = Range(.Offset(NASSETS, 1), .Offset(1, NASSETS))
    If ADD_RNG_NAME = True Then: COVAR_RNG.name = "VARCOV_MAT"
    For i = 1 To NASSETS
        With .Offset(0, i)
            .formula = "=offset(" & CORREL_POS_RNG.Address & ",0," & CStr(i) & ")"
        End With
        With .Offset(i, 0)
            .formula = "=offset(" & CORREL_POS_RNG.Address & ",0," & CStr(i) & ")"
        End With
    Next i
    With .Offset(-1, 0)
      .value = "VAR-COV MATRIX"
      .Font.Bold = True
    End With
End With
    
With COVAR_RNG
    For i = 1 To NASSETS - 1
        For j = i + 1 To NASSETS
            With .Cells(i, j)
                .formula = "=OFFSET(" & CORREL_POS_RNG.Address & _
                "," & CStr(i) & _
                "," & CStr(j) & ")*Index(" & CORREL_RNG.Address & _
                "," & CStr(j) & _
                "," & CStr(j) & ")*Index(" & CORREL_RNG.Address & _
                "," & CStr(i) & _
                "," & CStr(i) & ")"
            End With
        Next j
    Next i
    For i = 1 To NASSETS
        With .Cells(i, i)
            .formula = "=OFFSET(" & CORREL_POS_RNG.Address & "," & CStr(i) & "," & CStr(i) & ")^2"
        End With
    Next i
    For i = 2 To NASSETS
        For j = 1 To i - 1
            With .Cells(i, j)
                .formula = "=offset(" & COVAR_POS_RNG.Address & "," & CStr(j) & "," & CStr(i) & ")"
            End With
        Next j
    Next i
End With

'----------------------------SETTING UP CHOLESKI MATRIX-----------------------------
Set CHOL_POS_RNG = COVAR_POS_RNG.Offset(NASSETS + k, 0)
With CHOL_POS_RNG
    Set CHOL_RNG = Range(.Offset(NASSETS, 1), .Offset(1, NASSETS))
    If ADD_RNG_NAME = True Then: CHOL_RNG.name = "CHOL_MATRIX"
    For i = 1 To NASSETS
          With .Offset(0, i)
              .formula = "=offset(" & CORREL_POS_RNG.Address & ",0," & CStr(i) & ")"
          End With
          With .Offset(i, 0)
              .formula = "=offset(" & CORREL_POS_RNG.Address & ",0," & CStr(i) & ")"
          End With
    Next i
    With .Offset(-1, 0)
      .value = "CHOLESKI MATRIX"
      .Font.Bold = True
    End With
End With

CHOL_RNG.FormulaArray = "=MATRIX_CHOLESKY_FUNC(" & CORREL_RNG.Address & ")"

RNG_CORRELATION_COVARIANCE_FUNC = True

Exit Function
ERROR_LABEL:
RNG_CORRELATION_COVARIANCE_FUNC = False
End Function
