Attribute VB_Name = "FINAN_ASSET_TA_ROLLING_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_ROLLING_MEAN_FUNC
'DESCRIPTION   : Rolling Returns Forward
'LIBRARY       : FINAN_ASSET
'GROUP         : TA_ROLLING
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function ASSET_ROLLING_MEAN_FUNC(ByRef DATE_RNG As Variant, _
ByRef DATA_RNG As Variant, _
Optional ByVal MA_PERIODS As Long = 6, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal FACTOR_VAL As Double = 100)

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long

Dim DATE_VECTOR As Variant
Dim DATA_VECTOR As Variant

Dim TEMP_MULT As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 1) = 1 Then: DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

If DATA_TYPE <> 0 Then
    DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, 0)
    NROWS = UBound(DATE_VECTOR, 1) - 1
    k = 1
Else
    k = 0
    NROWS = UBound(DATE_VECTOR, 1)
End If

ReDim TEMP_MATRIX(0 To NROWS, 1 To 6)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PERIOD RETURN"
TEMP_MATRIX(0, 3) = "MULTIPLIER"
TEMP_MATRIX(0, 4) = "CUMUL RETURN"
TEMP_MATRIX(0, 5) = "ROLLING_RETURN: " & Format(MA_PERIODS, "0") & " PERIODS"
TEMP_MATRIX(0, 6) = "HPR" ' Holding Period Return

TEMP_MULT = 1

i = 1
TEMP_MATRIX(i, 1) = DATE_VECTOR(i + k, 1)
TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
TEMP_MATRIX(i, 3) = (1 + TEMP_MATRIX(i, 2) / FACTOR_VAL)
TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2)
TEMP_MATRIX(i, 5) = ""

For i = 2 To NROWS
    TEMP_MATRIX(i, 1) = DATE_VECTOR(i + k, 1)
    TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = (1 + TEMP_MATRIX(i, 2) / FACTOR_VAL)
    TEMP_MATRIX(i, 4) = FACTOR_VAL * ((1 + TEMP_MATRIX(i - 1, 4) / FACTOR_VAL) * (1 + TEMP_MATRIX(i, 2) / FACTOR_VAL) - 1)
    If i >= MA_PERIODS Then
        TEMP_MULT = 1
        For j = 0 To (MA_PERIODS - 1)
            TEMP_MULT = TEMP_MULT * TEMP_MATRIX(i - j, 3)
        Next j
        TEMP_MATRIX(i, 5) = FACTOR_VAL * (TEMP_MULT - 1)
    Else
        TEMP_MATRIX(i, 5) = ""
    End If
Next i

i = NROWS
TEMP_MATRIX(i, 6) = TEMP_MATRIX(i, 2)
For i = NROWS - 1 To 1 Step -1
    TEMP_MATRIX(i, 6) = FACTOR_VAL * ((1 + TEMP_MATRIX(i + 1, 6) / FACTOR_VAL) * (1 + TEMP_MATRIX(i, 2) / FACTOR_VAL) - 1)
Next i

ASSET_ROLLING_MEAN_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_ROLLING_MEAN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_ROLLING_SIGMA_FUNC
'DESCRIPTION   : Rolling Forward Volatilities
'LIBRARY       : FINAN_ASSET
'GROUP         : TA_ROLLING
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function ASSET_ROLLING_SIGMA_FUNC(ByRef DATE_RNG As Variant, _
ByRef DATA_RNG As Variant, _
Optional ByVal MA_PERIODS As Long = 12, _
Optional ByVal LAMBDA_VAL As Double = 0.9, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim NROWS As Long

Dim DATE_VECTOR As Variant
Dim DATA_VECTOR As Variant

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double
Dim TEMP4_SUM As Double

Dim SIGMA_VAL As Double

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

If MA_PERIODS < 2 Then: GoTo ERROR_LABEL

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 1) = 1 Then
    DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
End If

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If DATA_TYPE <> 0 Then
    DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
    NROWS = UBound(DATE_VECTOR, 1) - 1
    k = 1
Else
    k = 0
    NROWS = UBound(DATE_VECTOR, 1)
End If

ReDim TEMP_MATRIX(0 To NROWS, 1 To 7)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PERIOD_RETURN"
TEMP_MATRIX(0, 3) = "EWMA SIGMA (zero mean)"
TEMP_MATRIX(0, 4) = "EWMA SIGMA (non-zero mean)"
TEMP_MATRIX(0, 5) = "FULL PERIOD_SIGMA"
TEMP_MATRIX(0, 6) = "ROLLING SIGMA: " & Format(MA_PERIODS, "0") & " PERIODS"
TEMP_MATRIX(0, 7) = "PERIOD SIGMA"

TEMP1_SUM = 0
For i = 1 To NROWS
    TEMP1_SUM = TEMP1_SUM + DATA_VECTOR(i, 1)
Next i
TEMP1_SUM = TEMP1_SUM / NROWS

TEMP2_SUM = 0
For i = 1 To NROWS
    TEMP2_SUM = TEMP2_SUM + (DATA_VECTOR(i, 1) - TEMP1_SUM) ^ 2
Next i
TEMP2_SUM = (TEMP2_SUM / (NROWS - 1)) ^ 0.5
SIGMA_VAL = TEMP2_SUM

TEMP1_SUM = 0: TEMP3_SUM = 0
For i = 1 To NROWS
    
    TEMP_MATRIX(i, 1) = DATE_VECTOR(i + k, 1)
    TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
    
    ReDim TEMP_VECTOR(1 To i, 1 To 1)
    For j = 1 To i
        TEMP_VECTOR(j, 1) = DATA_VECTOR(j, 1)
    Next j
    TEMP1_SUM = TEMP1_SUM + DATA_VECTOR(i, 1)

    TEMP_MATRIX(i, 3) = VECTOR_VOLATILITY_FORECAST_FUNC(TEMP_VECTOR, LAMBDA_VAL, 1, 0, 0)(i, 1)
    TEMP_MATRIX(i, 4) = VECTOR_VOLATILITY_FORECAST_FUNC(TEMP_VECTOR, LAMBDA_VAL, 0, 0, 0)(i, 1)
    TEMP_MATRIX(i, 5) = SIGMA_VAL

    If i > 1 Then
        TEMP2_SUM = 0
        For j = 1 To i
            TEMP2_SUM = TEMP2_SUM + (DATA_VECTOR(j, 1) - (TEMP1_SUM / i)) ^ 2
        Next j
        TEMP_MATRIX(i, 7) = (TEMP2_SUM / (i - 1)) ^ 0.5
    Else
        TEMP_MATRIX(i, 7) = ""
    End If
    
    If i >= MA_PERIODS Then
        TEMP3_SUM = TEMP3_SUM + TEMP_MATRIX(i, 2)
        l = i - MA_PERIODS + 1
        TEMP4_SUM = 0
        For j = i To l Step -1
            TEMP4_SUM = TEMP4_SUM + (TEMP_MATRIX(j, 2) - (TEMP3_SUM / MA_PERIODS)) ^ 2
        Next j
        TEMP_MATRIX(i, 6) = (TEMP4_SUM / (MA_PERIODS - 1)) ^ 0.5
        TEMP3_SUM = TEMP3_SUM - TEMP_MATRIX(l, 2)
    Else
        TEMP3_SUM = TEMP3_SUM + TEMP_MATRIX(i, 2)
        TEMP_MATRIX(i, 6) = ""
    End If

Next i

ASSET_ROLLING_SIGMA_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_ROLLING_SIGMA_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : ASSET_ROLLING_SIGMA_CONFIDENCE_FUNC
'DESCRIPTION   : Confidence Intervals for Volatilities
'LIBRARY       : FINAN_ASSET
'GROUP         : TA_ROLLING
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function ASSET_ROLLING_SIGMA_CONFIDENCE_FUNC(ByRef DATE_RNG As Variant, _
ByRef DATA_RNG As Variant, _
Optional ByVal MA_PERIODS As Long = 12, _
Optional ByVal ALPHA_VAL As Double = 0.05, _
Optional ByVal COUNT_BASIS As Double = 12, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

'Volatility CONFIDENCE Intervals

'In the case of volatility, the estimator is constructed from a sum of
'squared normally-distributed random variables. For this reason, the
'sampling distribution is closely related to the Chi-square
'distribution.


Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim NROWS As Long

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double
Dim TEMP4_SUM As Double

Dim DATE_VECTOR As Variant
Dim DATA_VECTOR As Variant

Dim SIGMA_VAL As Double
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

If MA_PERIODS < 2 Then: GoTo ERROR_LABEL

DATE_VECTOR = DATE_RNG
If UBound(DATE_VECTOR, 1) = 1 Then: DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)

If DATA_TYPE <> 0 Then
    DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
    NROWS = UBound(DATE_VECTOR, 1) - 1
    k = 1
Else
    k = 0
    NROWS = UBound(DATE_VECTOR, 1)
End If

ReDim TEMP_MATRIX(0 To NROWS, 1 To 9)

TEMP_MATRIX(0, 1) = "DATE"
TEMP_MATRIX(0, 2) = "PERIOD RETURN"
TEMP_MATRIX(0, 3) = "ROLLING SIGMA: " & Format(MA_PERIODS, "0") & " PERIODS"
TEMP_MATRIX(0, 4) = "ROLLING LOWER CONFIDENCE"
TEMP_MATRIX(0, 5) = "ROLLING UPPER CONFIDENCE"
TEMP_MATRIX(0, 6) = "CUMUL SIGMA"
TEMP_MATRIX(0, 7) = "SIGMA"
TEMP_MATRIX(0, 8) = "LOWER CONFIDENCE"
TEMP_MATRIX(0, 9) = "UPPER CONFIDENCE"

TEMP1_SUM = 0
For i = 1 To NROWS
    TEMP1_SUM = TEMP1_SUM + DATA_VECTOR(i, 1)
Next i
TEMP1_SUM = TEMP1_SUM / NROWS

TEMP2_SUM = 0
For i = 1 To NROWS
    TEMP2_SUM = TEMP2_SUM + (DATA_VECTOR(i, 1) - TEMP1_SUM) ^ 2
Next i
TEMP2_SUM = (TEMP2_SUM / (NROWS - 1)) ^ 0.5
SIGMA_VAL = TEMP2_SUM * Sqr(COUNT_BASIS)

TEMP1_SUM = 0: TEMP3_SUM = 0

For i = 1 To NROWS
    
    TEMP_MATRIX(i, 1) = DATE_VECTOR(i + k, 1)
    TEMP_MATRIX(i, 2) = DATA_VECTOR(i, 1)
    
    
    If i >= MA_PERIODS Then
        
        TEMP3_SUM = TEMP3_SUM + TEMP_MATRIX(i, 2)
        l = i - MA_PERIODS + 1
        TEMP4_SUM = 0
        For j = i To l Step -1
            TEMP4_SUM = TEMP4_SUM + (TEMP_MATRIX(j, 2) - (TEMP3_SUM / MA_PERIODS)) ^ 2
        Next j
        TEMP_MATRIX(i, 3) = (TEMP4_SUM / (MA_PERIODS - 1)) ^ 0.5 * Sqr(COUNT_BASIS)
        TEMP3_SUM = TEMP3_SUM - TEMP_MATRIX(l, 2)
        TEMP_MATRIX(i, 4) = Sqr((COUNT_BASIS - 1) * (TEMP_MATRIX(i, 3) ^ 2) / INVERSE_CHI_SQUARED_DIST_FUNC(ALPHA_VAL / 2, COUNT_BASIS - 1, False))
        TEMP_MATRIX(i, 5) = Sqr((COUNT_BASIS - 1) * (TEMP_MATRIX(i, 3) ^ 2) / INVERSE_CHI_SQUARED_DIST_FUNC(1 - ALPHA_VAL / 2, COUNT_BASIS - 1, False))
    Else
        TEMP3_SUM = TEMP3_SUM + TEMP_MATRIX(i, 2)
        TEMP_MATRIX(i, 3) = ""
        TEMP_MATRIX(i, 4) = ""
        TEMP_MATRIX(i, 5) = ""
    End If
    
    TEMP1_SUM = TEMP1_SUM + DATA_VECTOR(i, 1)
    If i > 1 Then
        TEMP2_SUM = 0
        For j = 1 To i
            TEMP2_SUM = TEMP2_SUM + (DATA_VECTOR(j, 1) - (TEMP1_SUM / i)) ^ 2
        Next j
        TEMP_MATRIX(i, 6) = (TEMP2_SUM / (i - 1)) ^ 0.5 * Sqr(COUNT_BASIS)
    Else
        TEMP_MATRIX(i, 6) = ""
    End If
    TEMP_MATRIX(i, 7) = SIGMA_VAL
    
    If i > 1 Then
        TEMP_MATRIX(i, 8) = Sqr((i - 1) * (TEMP_MATRIX(i, 6) ^ 2) / INVERSE_CHI_SQUARED_DIST_FUNC(ALPHA_VAL / 2, i - 1, False))
    Else
        TEMP_MATRIX(i, 8) = ""
    End If

    If i > 2 Then
        TEMP_MATRIX(i, 9) = Sqr((i - 1) * (TEMP_MATRIX(i, 6) ^ 2) / INVERSE_CHI_SQUARED_DIST_FUNC(1 - ALPHA_VAL / 2, i - 1, False))
    Else
        TEMP_MATRIX(i, 9) = ""
    End If

Next i

ASSET_ROLLING_SIGMA_CONFIDENCE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSET_ROLLING_SIGMA_CONFIDENCE_FUNC = Err.number
End Function
