Attribute VB_Name = "STAT_DIST_EMPIRICAL_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : FIT_EMPIRICAL_KERNEL_DISTRIBUTION_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : DIST_EMPIRICAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function FIT_EMPIRICAL_KERNEL_DISTRIBUTION_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NBINS As Long = 100)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim MIN_VAL As Double
Dim TEMP_VAL As Double

Dim DATA_VECTOR As Variant

Dim FACTOR1_VAL As Double
Dim FACTOR2_VAL As Double

Dim INV_SQR2_PI_VAL As Double

On Error GoTo ERROR_LABEL

INV_SQR2_PI_VAL = 1 / Sqr(2 * 3.14159265359)
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)

MIN_VAL = 2 ^ 50: MAX_VAL = -2 ^ 50
For i = 1 To NROWS
  If DATA_VECTOR(i, 1) < MIN_VAL Then: MIN_VAL = DATA_VECTOR(i, 1)
  If DATA_VECTOR(i, 1) > MAX_VAL Then: MAX_VAL = DATA_VECTOR(i, 1)
Next i

FACTOR1_VAL = Abs(MAX_VAL - MIN_VAL) / (NBINS - 1)
FACTOR2_VAL = 1.06 * MINIMUM_FUNC(MATRIX_STDEV_FUNC(DATA_VECTOR)(1, 1), HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.75, 1) - HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 0.25, 1)) * NROWS ^ (-1 / 5)
ReDim TEMP_MATRIX(1 To NBINS, 1 To 2) 'x   pdf
For i = 1 To NBINS
    TEMP_MATRIX(i, 1) = MIN_VAL + (i - 1) * FACTOR1_VAL
    TEMP_VAL = 0
    For j = 1 To NROWS
        TEMP_VAL = TEMP_VAL + INV_SQR2_PI_VAL * Exp(-0.5 * ((TEMP_MATRIX(i, 1) - DATA_VECTOR(j, 1)) / FACTOR2_VAL) ^ 2)
    Next j
    TEMP_MATRIX(i, 2) = TEMP_VAL / (NROWS * FACTOR2_VAL * NBINS)
Next i

FIT_EMPIRICAL_KERNEL_DISTRIBUTION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
FIT_EMPIRICAL_KERNEL_DISTRIBUTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FIT_EMPIRICAL_SMOOTH_DISTRIBUTION_FUNC
'DESCRIPTION   :

' Empirical distribution PDF and CDF with Epanechnikov kernel smoothing

' Epanechnikov kernel is optimum choice for smoothing because
' it minimizes asymptotic mean integrated squared error (AMISE).
' In fact all other kernels efficiency like gaussian, triangular,
' uniform is measured against this kernel

' Following is algorith/formula for the Epanechnikov kernel:
'
'   estimate =(3/4)(1-u2) for -1<u<1
'   estimate = 0 for u outside that range.
'
'   where
'   FACTOR_VAL:  Window Width
'   xi : absicassa/ values of the independent variables in the data
'   u = (x - xi) / FACTOR_VAL
'
' To use this function, set DATA_VECTOR to the data series, and
' nBins to the no. of desired x-axis points
'
' It returns a matrix with 3 columns.
' 1.x-axis abscissa
' 2.probability density function value at the x-point
' 3.cumulative density function value at the x-point

'LIBRARY       : STATISTICS
'GROUP         : DIST_EMPIRICAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function FIT_EMPIRICAL_SMOOTH_DISTRIBUTION_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal NBINS As Long = 30)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim K_VAL As Double
Dim X_VAL As Double
Dim Z_VAL As Double

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim DELTA_VAL As Double
Dim ROOT_VAL As Double

Dim TEMP_SUM As Double
Dim TEMP_VAL As Double
Dim FACTOR_VAL As Double

Dim PDF_VECTOR As Variant
Dim CDF_VECTOR As Variant

Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NBINS, 1 To 3)
TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_SUM = TEMP_SUM + DATA_VECTOR(i, 1)
Next i
MEAN_VAL = TEMP_SUM / NROWS

TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_VAL = DATA_VECTOR(i, 1) - MEAN_VAL
    TEMP_SUM = TEMP_SUM + TEMP_VAL * TEMP_VAL
Next i
SIGMA_VAL = TEMP_SUM / (NROWS - 1)
SIGMA_VAL = SIGMA_VAL ^ 0.5
FACTOR_VAL = 1.06 * SIGMA_VAL * NROWS ^ (-1 / 5)

ReDim TEMP_VECTOR(1 To NBINS, 1 To 1)

MIN_VAL = 2 ^ 50: MAX_VAL = -2 ^ 50
For i = 1 To NROWS
    If DATA_VECTOR(i, 1) < MIN_VAL Then: MIN_VAL = DATA_VECTOR(i, 1)
    If DATA_VECTOR(i, 1) > MAX_VAL Then: MAX_VAL = DATA_VECTOR(i, 1)
Next i
  
DELTA_VAL = (MAX_VAL - MIN_VAL) / NBINS
TEMP_VECTOR(1, 1) = MIN_VAL
For i = 2 To NBINS
    TEMP_VECTOR(i, 1) = TEMP_VECTOR(i - 1, 1) + DELTA_VAL
Next i

ReDim PDF_VECTOR(1 To NBINS, 1 To 1)

ROOT_VAL = 5 ^ 0.5
TEMP_SUM = 0
TEMP_VAL = (3 / (4 * ROOT_VAL))
For i = 1 To NBINS
    X_VAL = TEMP_VECTOR(i, 1)
    TEMP_SUM = 0
    For j = 1 To NROWS
        Z_VAL = (X_VAL - DATA_VECTOR(j, 1)) / FACTOR_VAL
        If Abs(Z_VAL) <= ROOT_VAL Then
            K_VAL = TEMP_VAL * (1 - (1 / 5) * Z_VAL * Z_VAL)
        Else
            K_VAL = 0
        End If
        TEMP_SUM = TEMP_SUM + K_VAL
    Next j
    PDF_VECTOR(i, 1) = TEMP_SUM / (FACTOR_VAL * NROWS)
Next i

ReDim CDF_VECTOR(1 To NBINS, 1 To 1)
TEMP_SUM = 0
For i = 1 To NBINS
    TEMP_SUM = TEMP_SUM + PDF_VECTOR(i, 1) * DELTA_VAL
    CDF_VECTOR(i, 1) = TEMP_SUM
Next i

TEMP_MATRIX(0, 1) = "X"
TEMP_MATRIX(0, 2) = "PDF"
TEMP_MATRIX(0, 3) = "CDF"

For i = 1 To NBINS
    TEMP_MATRIX(i, 1) = TEMP_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = PDF_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = CDF_VECTOR(i, 1)
Next i

FIT_EMPIRICAL_SMOOTH_DISTRIBUTION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
FIT_EMPIRICAL_SMOOTH_DISTRIBUTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FIT_EMPIRICAL_NORMAL_DISTRIBUTION_FUNC
'DESCRIPTION   : Prob. Dist & Cumul. Dist Functions
'LIBRARY       : STATISTICS
'GROUP         : DIST_EMPIRICAL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function FIT_EMPIRICAL_NORMAL_DISTRIBUTION_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 1, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal NBINS As Long = 20)

'CUMUL_CHART: X-AXIS: BIN_VECTOR; Y-AXIS: CDF_VECTOR
'PROB_HIST: X-AXIS: MID_POINTS; Y-AXIS: PDF_VECTOR

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim NOBS As Double
Dim MIN_VAL As Double
Dim MAX_VAL As Double
Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double
Dim WIDTH_VAL As Double

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
NROWS = UBound(DATA_VECTOR, 1)

TEMP_VECTOR = DATA_BASIC_MOMENTS_FUNC(DATA_VECTOR, 0, 0, 0.05, 0)

NOBS = TEMP_VECTOR(1, 1)
MIN_VAL = TEMP_VECTOR(1, 2)
MAX_VAL = TEMP_VECTOR(1, 3)
MEAN_VAL = TEMP_VECTOR(1, 4)
SIGMA_VAL = TEMP_VECTOR(1, 7)
WIDTH_VAL = (MAX_VAL - MIN_VAL) / NBINS

DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
ReDim TEMP1_MATRIX(1 To NROWS, 1 To 2)
For i = 1 To NROWS
    TEMP1_MATRIX(i, 1) = DATA_VECTOR(i, 1)
    TEMP1_MATRIX(i, 2) = i / NROWS
Next i

ReDim TEMP2_MATRIX(0 To NBINS + 1, 1 To 6)

TEMP2_MATRIX(0, 1) = "INDEX"
TEMP2_MATRIX(0, 2) = "BIN"
TEMP2_MATRIX(0, 3) = "MID-POINT"
TEMP2_MATRIX(0, 4) = "PDF"
TEMP2_MATRIX(0, 5) = "CDF"
TEMP2_MATRIX(0, 6) = "CND"

i = NBINS
TEMP2_MATRIX(i + 1, 1) = i
TEMP2_MATRIX(i + 1, 2) = MIN_VAL + WIDTH_VAL * TEMP2_MATRIX(i + 1, 1)
TEMP2_MATRIX(i + 1, 3) = ""
j = NROWS
Do While TEMP2_MATRIX(i + 1, 2) < TEMP1_MATRIX(j, 1): j = j - 1: Loop
TEMP2_MATRIX(i + 1, 5) = TEMP1_MATRIX(j, 2)
TEMP2_MATRIX(i + 1, 6) = NORMSDIST_FUNC(TEMP2_MATRIX(i + 1, 2), MEAN_VAL, SIGMA_VAL, 0)
TEMP2_MATRIX(i + 1, 4) = ""

For i = NBINS - 1 To 1 Step -1
    TEMP2_MATRIX(i + 1, 1) = i
    TEMP2_MATRIX(i + 1, 2) = MIN_VAL + WIDTH_VAL * TEMP2_MATRIX(i + 1, 1)
    TEMP2_MATRIX(i + 1, 3) = (TEMP2_MATRIX(i + 1, 2) + TEMP2_MATRIX(i + 2, 2)) / 2
    j = NROWS
    Do While TEMP2_MATRIX(i + 1, 2) < TEMP1_MATRIX(j, 1): j = j - 1: Loop
    TEMP2_MATRIX(i + 1, 5) = TEMP1_MATRIX(j, 2)
    TEMP2_MATRIX(i + 1, 6) = NORMSDIST_FUNC(TEMP2_MATRIX(i + 1, 2), MEAN_VAL, SIGMA_VAL, 0)
    TEMP2_MATRIX(i + 1, 4) = (TEMP2_MATRIX(i + 2, 5) - TEMP2_MATRIX(i + 1, 5))
Next i

i = 0
TEMP2_MATRIX(i + 1, 1) = i
TEMP2_MATRIX(i + 1, 2) = MIN_VAL + WIDTH_VAL * TEMP2_MATRIX(i + 1, 1)
TEMP2_MATRIX(i + 1, 3) = (TEMP2_MATRIX(i + 1, 2) + TEMP2_MATRIX(i + 2, 2)) / 2
TEMP2_MATRIX(i + 1, 5) = 0
TEMP2_MATRIX(i + 1, 6) = 0
TEMP2_MATRIX(i + 1, 4) = 0

FIT_EMPIRICAL_NORMAL_DISTRIBUTION_FUNC = TEMP2_MATRIX

Exit Function
ERROR_LABEL:
FIT_EMPIRICAL_NORMAL_DISTRIBUTION_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : FIT_EMPIRICAL_LAPLACE_DISTRIBUTION_FUNC
'DESCRIPTION   : Hist Analysis --> Prob. Dist & Cumul. Dist Functions
'LIBRARY       : STATISTICS
'GROUP         : DIST_EMPIRICAL
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function FIT_EMPIRICAL_LAPLACE_DISTRIBUTION_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal FACTOR_VAL As Double = 1)

Dim i As Long
Dim NROWS As Long
Dim NBINS As Long

Dim BIN_MIN As Double
Dim BIN_WIDTH As Double

Dim MEAN_VAL As Double
Dim SIGMA_VAL As Double

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double

Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant
Dim FREQUENCY_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
NROWS = UBound(DATA_VECTOR, 1)

FREQUENCY_VECTOR = DATA_BASIC_MOMENTS_FUNC(DATA_VECTOR, 0, 0, 0.05, 0)

MIN_VAL = FREQUENCY_VECTOR(1, 2)
MAX_VAL = FREQUENCY_VECTOR(1, 3)
MEAN_VAL = FREQUENCY_VECTOR(1, 4)
SIGMA_VAL = FREQUENCY_VECTOR(1, 7)

FREQUENCY_VECTOR = HISTOGRAM_BIN_LIMITS_FUNC(MIN_VAL, MAX_VAL, NROWS, 3)
BIN_WIDTH = FREQUENCY_VECTOR(LBound(FREQUENCY_VECTOR))
BIN_MIN = FREQUENCY_VECTOR(LBound(FREQUENCY_VECTOR) + 1)
NBINS = FREQUENCY_VECTOR(LBound(FREQUENCY_VECTOR) + 2)
FREQUENCY_VECTOR = HISTOGRAM_FREQUENCY_FUNC(DATA_VECTOR, NBINS, BIN_MIN, BIN_WIDTH, 1)
NBINS = UBound(FREQUENCY_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NBINS, 1 To 8)

TEMP_MATRIX(0, 1) = "BINS"
TEMP_MATRIX(0, 2) = "FREQ"
TEMP_MATRIX(0, 3) = "ACTUAL: PDF"
TEMP_MATRIX(0, 4) = "ACTUAL: CDF"
TEMP_MATRIX(0, 5) = "NORMAL: PDF"
TEMP_MATRIX(0, 6) = "NORMAL: CDF"
TEMP_MATRIX(0, 7) = "LAPLACE: PDF"
TEMP_MATRIX(0, 8) = "LAPLACE: CDF"

TEMP1_SUM = 0
For i = 1 To NBINS: TEMP1_SUM = TEMP1_SUM + FREQUENCY_VECTOR(i, 2): Next i
For i = 1 To NBINS
    
    TEMP_MATRIX(i, 1) = FREQUENCY_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = FREQUENCY_VECTOR(i, 2)
    TEMP_MATRIX(i, 3) = TEMP_MATRIX(i, 2) / TEMP1_SUM
        
    TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(i, 2)
    
    TEMP_MATRIX(i, 4) = TEMP2_SUM / TEMP1_SUM
    TEMP_MATRIX(i, 5) = NORMDIST_FUNC(TEMP_MATRIX(i, 1), MEAN_VAL, SIGMA_VAL, 0) * FACTOR_VAL
        
    TEMP_MATRIX(i, 6) = NORMSDIST_FUNC(TEMP_MATRIX(i, 1), MEAN_VAL, SIGMA_VAL, 0)
        
    TEMP_MATRIX(i, 7) = LAPLACE_DIST_FUNC(TEMP_MATRIX(i, 1), MEAN_VAL, SIGMA_VAL, FACTOR_VAL, False)
    TEMP_MATRIX(i, 8) = LAPLACE_DIST_FUNC(TEMP_MATRIX(i, 1), MEAN_VAL, SIGMA_VAL, FACTOR_VAL, True)
Next i

FIT_EMPIRICAL_LAPLACE_DISTRIBUTION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
FIT_EMPIRICAL_LAPLACE_DISTRIBUTION_FUNC = Err.number
End Function
