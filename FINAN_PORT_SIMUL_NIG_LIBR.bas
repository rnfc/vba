Attribute VB_Name = "FINAN_PORT_SIMUL_NIG_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_NIG_SIMULATION_FUNC

'DESCRIPTION   : Simulation of Correlated NIG Distributions
'Starting with the first four moments of n Stocks and a rank correlation matrix between them,
'weighted portfolio returns are simulated assuming correlated NIG distributions.
'Expected Tail Loss (=Conditional Value-At-Risk) and VaR are calculated for the portfolio.


'http://en.wikipedia.org/wiki/Normal-inverse_Gaussian_distribution
'http://digilander.libero.it/foxes/poly/Moments_Regression.pdf

'LIBRARY       : PORTFOLIO
'GROUP         : SIMULATION_NIG
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 22/08/2010
'************************************************************************************
'************************************************************************************

Function PORT_NIG_SIMULATION_FUNC(ByRef RETURNS_RNG As Variant, _
ByRef VARIANCES_RNG As Variant, _
ByRef SKEWNESS_RNG As Variant, _
ByRef KURTOSIS_RNG As Variant, _
ByRef CORREL_RNG As Variant, _
Optional ByRef WEIGHTS_RNG As Variant, _
Optional ByVal CONFIDENCE_VAL As Double = 0.95, _
Optional ByVal nLOOPS As Long = 5000)

'ExpectedReturns: RETURNS_RNG

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long

Dim TEMP_SUM As Double
Dim VAR_VAL As Double
Dim RETURN_VAL As Double

Dim PARAMS_ARR As Variant 'Empirical Moments

Dim ALPHA_ARR() As Double
Dim BETA_ARR() As Double
Dim MU_ARR() As Double
Dim DELTA_ARR() As Double

Dim SN_MATRIX() As Double
Dim IG_MATRIX() As Double

Dim TEMP0_MATRIX() As Double
Dim TEMP1_MATRIX() As Double

Dim WEIGHTS_VECTOR As Variant
Dim RETURNS_VECTOR As Variant
Dim VARIANCES_VECTOR As Variant
Dim SKEWNESS_VECTOR As Variant
Dim KURTOSIS_VECTOR As Variant
Dim CORREL_MATRIX As Variant

On Error GoTo ERROR_LABEL

RETURNS_VECTOR = RETURNS_RNG
If UBound(RETURNS_VECTOR, 1) = 1 Then
    RETURNS_VECTOR = MATRIX_TRANSPOSE_FUNC(RETURNS_VECTOR)
End If
NSIZE = UBound(RETURNS_VECTOR, 1)

VARIANCES_VECTOR = VARIANCES_RNG
If UBound(VARIANCES_VECTOR, 1) = 1 Then
    VARIANCES_VECTOR = MATRIX_TRANSPOSE_FUNC(VARIANCES_VECTOR)
End If
If NSIZE <> UBound(VARIANCES_VECTOR, 1) Then: GoTo ERROR_LABEL

SKEWNESS_VECTOR = SKEWNESS_RNG
If UBound(SKEWNESS_VECTOR, 1) = 1 Then
    SKEWNESS_VECTOR = MATRIX_TRANSPOSE_FUNC(SKEWNESS_VECTOR)
End If
If NSIZE <> UBound(SKEWNESS_VECTOR, 1) Then: GoTo ERROR_LABEL

KURTOSIS_VECTOR = KURTOSIS_RNG
If UBound(KURTOSIS_VECTOR, 1) = 1 Then
    KURTOSIS_VECTOR = MATRIX_TRANSPOSE_FUNC(KURTOSIS_VECTOR)
End If
If NSIZE <> UBound(KURTOSIS_VECTOR, 1) Then: GoTo ERROR_LABEL

CORREL_MATRIX = CORREL_RNG
If NSIZE <> UBound(CORREL_MATRIX, 1) Then: GoTo ERROR_LABEL
If NSIZE <> UBound(CORREL_MATRIX, 2) Then: GoTo ERROR_LABEL

GoSub NIG_LINE
'---------------------------------------------------------------------------------------------------------------------------
If IsArray(WEIGHTS_RNG) = False Then
'---------------------------------------------------------------------------------------------------------------------------
    PORT_NIG_SIMULATION_FUNC = TEMP1_MATRIX 'Portfolio Constituents
'---------------------------------------------------------------------------------------------------------------------------
Else 'Simulated Portfolio Characteristics
'---------------------------------------------------------------------------------------------------------------------------
    WEIGHTS_VECTOR = WEIGHTS_RNG
    If UBound(WEIGHTS_VECTOR, 1) = 1 Then
        WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
    End If
    If NSIZE <> UBound(WEIGHTS_VECTOR, 1) Then: GoTo ERROR_LABEL
    ReDim TEMP0_MATRIX(1 To nLOOPS, 1 To 1)
    For i = 1 To nLOOPS
        RETURN_VAL = 0
        For j = 1 To NSIZE
            RETURN_VAL = RETURN_VAL + WEIGHTS_VECTOR(j, 1) * TEMP1_MATRIX(i, j)
        Next j
        TEMP0_MATRIX(i, 1) = RETURN_VAL 'Simulated Returns
    Next i
    Erase TEMP1_MATRIX
    VAR_VAL = HISTOGRAM_PERCENTILE_FUNC(TEMP0_MATRIX, 1 - CONFIDENCE_VAL, 1) 'value -At - Risk
    k = 0: TEMP_SUM = 0
    For i = 1 To nLOOPS
        RETURN_VAL = TEMP0_MATRIX(i, 1)
        If RETURN_VAL <= VAR_VAL Then 'Expected Tail Loss
            k = k + 1
            TEMP_SUM = TEMP_SUM + RETURN_VAL
        End If
    Next i
    PORT_NIG_SIMULATION_FUNC = Array(TEMP_SUM / k, VAR_VAL)
'---------------------------------------------------------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------------------------------------------------------

Exit Function
'---------------------------------------------------------------------------------------------------------------------------
NIG_LINE:
'---------------------------------------------------------------------------------------------------------------------------
    ' generate NIG parameters
    ReDim ALPHA_ARR(1 To 1, 1 To NSIZE)
    ReDim BETA_ARR(1 To 1, 1 To NSIZE)
    ReDim MU_ARR(1 To 1, 1 To NSIZE)
    ReDim DELTA_ARR(1 To 1, 1 To NSIZE)
    For j = 1 To NSIZE
        PARAMS_ARR = NIG_MLE_PARAMETERS_FUNC(RETURNS_VECTOR(j, 1), VARIANCES_VECTOR(j, 1), SKEWNESS_VECTOR(j, 1), KURTOSIS_VECTOR(j, 1))
        ALPHA_ARR(1, j) = PARAMS_ARR(1, 1)
        BETA_ARR(1, j) = PARAMS_ARR(2, 1)
        MU_ARR(1, j) = PARAMS_ARR(3, 1)
        DELTA_ARR(1, j) = PARAMS_ARR(4, 1)
    Next j
    ' generate independent standard normal matrix & IG matrix
    ReDim SN_MATRIX(1 To nLOOPS, 1 To NSIZE)
    ReDim IG_MATRIX(1 To nLOOPS, 1 To NSIZE)
    For i = 1 To nLOOPS
        For j = 1 To NSIZE
            SN_MATRIX(i, j) = RANDOM_NORMAL_FUNC(0, 1, 0)
            IG_MATRIX(i, j) = IG_RANDOM_FUNC(DELTA_ARR(1, j), (ALPHA_ARR(1, j) ^ 2 - BETA_ARR(1, j) ^ 2) ^ 0.5)
        Next j
    Next i
    
    ' generate Chol matrix
    ReDim TEMP0_MATRIX(1 To NSIZE, 1 To NSIZE)
    For j = 1 To NSIZE
        TEMP_SUM = 0
        For k = 1 To j - 1
            TEMP_SUM = TEMP_SUM + TEMP0_MATRIX(j, k) ^ 2
        Next k
        TEMP0_MATRIX(j, j) = CORREL_MATRIX(j, j) - TEMP_SUM
        ' Matrix is not semi-positive definite, no solution exists
        If TEMP0_MATRIX(j, j) < 0 Then: GoTo ERROR_LABEL
        TEMP0_MATRIX(j, j) = (MAXIMUM_FUNC(0, TEMP0_MATRIX(j, j))) ^ 0.5
        For i = j + 1 To NSIZE
            TEMP_SUM = 0
            For k = 1 To j - 1
                TEMP_SUM = TEMP_SUM + TEMP0_MATRIX(i, k) * TEMP0_MATRIX(j, k)
            Next k
            If TEMP0_MATRIX(j, j) = 0 Then
               TEMP0_MATRIX(i, j) = 0
            Else
               TEMP0_MATRIX(i, j) = (CORREL_MATRIX(i, j) - TEMP_SUM) / TEMP0_MATRIX(j, j)
            End If
        Next i
    Next j
    CORREL_MATRIX = TEMP0_MATRIX
    ' generate correlated standard normals: .transpose(.mmult(CORREL_MATRIX, .transpose(SN_MATRIX)))
    ReDim TEMP0_MATRIX(1 To nLOOPS, 1 To NSIZE)
    For i = 1 To NSIZE
        For j = 1 To nLOOPS
            TEMP_SUM = 0
            For k = 1 To NSIZE
                TEMP_SUM = TEMP_SUM + CORREL_MATRIX(i, k) * SN_MATRIX(j, k)
            Next k
            TEMP0_MATRIX(j, i) = TEMP_SUM
        Next j
    Next i
    Erase CORREL_MATRIX
    ' generate normals with given volatilities and expected returns
    ReDim TEMP1_MATRIX(1 To nLOOPS, 1 To NSIZE)
    For i = 1 To nLOOPS
        For j = 1 To NSIZE
            TEMP1_MATRIX(i, j) = MU_ARR(1, j) + IG_MATRIX(i, j) * BETA_ARR(1, j) + TEMP0_MATRIX(i, j) * IG_MATRIX(i, j) ^ 0.5
        Next j
    Next i
    Erase TEMP0_MATRIX
'---------------------------------------------------------------------------------------------------------------------------
Return
'---------------------------------------------------------------------------------------------------------------------------
ERROR_LABEL:
PORT_NIG_SIMULATION_FUNC = Err.number
End Function


'# Maximum Likelihood Estimation of Univariate NIG parameters & Analysis
'of NIG VaR: Comparison of Value-At-Risk figures for daily SMI stock return
'series calculated as a) VaR based on NIG moments, b) VaR based on NIG MLE
'estimation, c) VaR based on the Gaussian distribution, d) Historical VaR
'and e) Cornish-Fisher / Modified VaR

Function ASSETS_NIG_VAR_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CONFIDENCE_VAL As Double = 0.99, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

'MM: Moment Matching
'MLE: Maximum Likelihood Estimation

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim TEMP_VAR As Double
Dim TEMP_DEV As Double

Dim MEAN_VAL As Double
Dim VAR_VAL As Double
Dim VOLAT_VAL As Double
Dim SKEW_VAL As Double
Dim KURT_VAL As Double

Dim HEADINGS_STR As String

Dim DATA_MATRIX As Variant
Dim DATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant
Dim MLE_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then: DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)

NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_MATRIX(0 To NCOLUMNS, 1 To 22)
HEADINGS_STR = "MM: E(X),MM: V(X),MM: S(X),MM: K(X),MLE: E(X),MLE: V(X),MLE: S(X),MLE: K(X),MM: ALPHA,MM: BETA,MM: MU,MM: DELTA,MLE: ALPHA,MLE: BETA,MLE: MU,MLE: DELTA,VALUE-AT-RISK,NIG MOMENTS,NIG MLE,NORMAL,HISTORICAL,CORNISH-FISHER,"
i = 1
For k = 1 To 22
    j = InStr(i, HEADINGS_STR, ",")
    TEMP_MATRIX(0, k) = Mid(HEADINGS_STR, i, j - i)
    i = j + 1
Next k

ReDim DATA_VECTOR(1 To NROWS, 1 To 1)
For j = 1 To NCOLUMNS
    TEMP_SUM = 0
    For i = 1 To NROWS
        DATA_VECTOR(i, 1) = DATA_MATRIX(i, j)
        TEMP_SUM = TEMP_SUM + DATA_VECTOR(i, 1)
    Next i
    MEAN_VAL = TEMP_SUM / NROWS
    TEMP_DEV = 0: TEMP_SUM = 0: TEMP_VAR = 0
    For i = 1 To NROWS
        TEMP_DEV = (DATA_VECTOR(i, 1) - MEAN_VAL)
        TEMP_SUM = TEMP_SUM + TEMP_DEV
        TEMP_VAR = TEMP_DEV * TEMP_DEV + TEMP_VAR
    Next i
    VAR_VAL = (TEMP_VAR - TEMP_SUM * TEMP_SUM / NROWS) / (NROWS - 1)
    'Variance: Corrected two-pass formula.
    'VOLAT_VAL = Sqr(NROWS / (NROWS - 1)) * VOLAT_VAL 'Population
    VOLAT_VAL = Sqr(VAR_VAL) 'Sample Standard Deviation
    ' Calculate 3rd and 4th moments
    SKEW_VAL = 0: KURT_VAL = 0
    For i = 1 To NROWS
        SKEW_VAL = SKEW_VAL + ((DATA_VECTOR(i, 1) - MEAN_VAL) / VOLAT_VAL) ^ 3
        KURT_VAL = KURT_VAL + ((DATA_VECTOR(i, 1) - MEAN_VAL) / VOLAT_VAL) ^ 4
    Next i
    '-----------------------------------------------------------------------------------------------------------
    'calcs with moments
    'SKEW_VAL = SKEW_VAL / NROWS
    'KURT_VAL = (KURT_VAL / NROWS) '- 3
    '-----------------------------------------------------------------------------------------------------------
    
    '-----------------------------------------------------------------------------------------------------------
    'calcs with Excel's definition of Skew & Kurt
    KURT_VAL = (KURT_VAL * (NROWS * (NROWS + 1) / ((NROWS - 1) * (NROWS - 2) * (NROWS - 3)))) - ((3 * (NROWS - 1) ^ 2 / ((NROWS - 2) * (NROWS - 3)))) 'Excel Definition
    KURT_VAL = KURT_VAL + 3
    SKEW_VAL = SKEW_VAL * (NROWS / ((NROWS - 1) * (NROWS - 2))) 'Excel Definition
    '-----------------------------------------------------------------------------------------------------------
        
    TEMP_MATRIX(j, 1) = MEAN_VAL: TEMP_MATRIX(j, 2) = VAR_VAL
    TEMP_MATRIX(j, 3) = SKEW_VAL: TEMP_MATRIX(j, 4) = KURT_VAL
    
    PARAM_VECTOR = NIG_MLE_PARAMETERS_FUNC(MEAN_VAL, VAR_VAL, SKEW_VAL, KURT_VAL)
    If IsArray(PARAM_VECTOR) = False Then: GoTo 1983
    
    TEMP_MATRIX(j, 9) = PARAM_VECTOR(1, 1)
    TEMP_MATRIX(j, 10) = PARAM_VECTOR(2, 1)
    TEMP_MATRIX(j, 11) = PARAM_VECTOR(3, 1)
    TEMP_MATRIX(j, 12) = PARAM_VECTOR(4, 1)

    MLE_VECTOR = NIG_MLE_SOLVER_FUNC(DATA_VECTOR, PARAM_VECTOR)
    If IsArray(MLE_VECTOR) = False Then: GoTo 1983
    
    TEMP_MATRIX(j, 13) = MLE_VECTOR(1, 1)
    TEMP_MATRIX(j, 14) = MLE_VECTOR(2, 1)
    TEMP_MATRIX(j, 15) = MLE_VECTOR(3, 1)
    TEMP_MATRIX(j, 16) = MLE_VECTOR(4, 1)

    TEMP_MATRIX(j, 17) = CONFIDENCE_VAL
    TEMP_MATRIX(j, 18) = NIG_INV_CDF_FUNC(1 - CONFIDENCE_VAL, PARAM_VECTOR(1, 1), PARAM_VECTOR(2, 1), PARAM_VECTOR(3, 1), PARAM_VECTOR(4, 1), 0, 1, 500, 20000, 0.0000001)
    TEMP_MATRIX(j, 19) = NIG_INV_CDF_FUNC(1 - CONFIDENCE_VAL, MLE_VECTOR(1, 1), MLE_VECTOR(2, 1), MLE_VECTOR(3, 1), MLE_VECTOR(4, 1), 0, 1, 500, 20000, 0.0000001)
    TEMP_MATRIX(j, 20) = MEAN_VAL + NORMSINV_FUNC(1 - CONFIDENCE_VAL, 0, 1, 0) * Sqr(VAR_VAL)
    TEMP_MATRIX(j, 21) = HISTOGRAM_PERCENTILE_FUNC(DATA_VECTOR, 1 - CONFIDENCE_VAL, 1)
    TEMP_MATRIX(j, 22) = MEAN_VAL + (NORMSINV_FUNC(1 - CONFIDENCE_VAL, 0, 1, 0) + (1 / 6) * (NORMSINV_FUNC(1 - CONFIDENCE_VAL, 0, 1, 0) ^ 2 - 1) * SKEW_VAL + (1 / 24) * (NORMSINV_FUNC(1 - CONFIDENCE_VAL, 0, 1, 0) ^ 3 - 3 * NORMSINV_FUNC(1 - CONFIDENCE_VAL, 0, 1, 0)) * (KURT_VAL - 3) - (1 / 36) * (2 * NORMSINV_FUNC(1 - CONFIDENCE_VAL, 0, 1, 0) ^ 3 - 5 * NORMSINV_FUNC(1 - CONFIDENCE_VAL, 0, 1, 0)) * SKEW_VAL ^ 2) * Sqr(VAR_VAL)
    
    PARAM_VECTOR = NIG_MLE_MOMENTS_FUNC(MLE_VECTOR(1, 1), MLE_VECTOR(2, 1), MLE_VECTOR(3, 1), MLE_VECTOR(4, 1))
    If IsArray(PARAM_VECTOR) = False Then: GoTo 1983
    TEMP_MATRIX(j, 5) = PARAM_VECTOR(1, 1)
    TEMP_MATRIX(j, 6) = PARAM_VECTOR(2, 1)
    TEMP_MATRIX(j, 7) = PARAM_VECTOR(3, 1)
    TEMP_MATRIX(j, 8) = PARAM_VECTOR(4, 1)
1983:
Next j

ASSETS_NIG_VAR_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
ASSETS_NIG_VAR_FUNC = Err.number
End Function

