Attribute VB_Name = "FINAN_PORT_FRONTIER_CO_LPM_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_CO_LPM_FRONTIER_FUNC

'DESCRIPTION   : Shortfall Risk Optimization: Various Lower Partial Moment (LPM) functions
'and generation of efficient frontiers based on the traditional covariance,
'the asymmetrical LPM and symmetrical LPM matrices.

'Literature:
'"Optimal Algorithms and Lower Partial Moment: Ex Post Results", Nawrocki D.,
'Applied Econometrics, 1991, 23

'"Asset Allocation Based on Shortfall Risk", PhD Thesis Dipl.-Kffr.
'Denisa ?umova, Fakultät für
'Wirtschaftswissenschaften der Technischen Universität Chemnitz, 2004

'This thesis presents an innovative portfolio model appropriate for a large group of
'investors which are not content with the asset allocation with the traditional, mean
'return-variance based portfolio model above all in term of its rather specific definition
'of the risk and value decision parameters, risk diversification, related utility function
'and its restrictions imposed on the asset universe. Its modifiable risk measure ? shortfall
'risk ? expresses variable risk preferences below the return benchmark. The upside return
'deviations from the benchmark are not minimized as in case of the mean return-variance
'portfolio model or considered risk neutral as in the mean return-shortfall risk portfolio
'model, but employs variable degrees of the chance potential (upper partial moments) in order
'to provide investors with broader range of utility choices and so reflect arbitrary preferences.
'The elimination of the assumption of normally distributed returns in the chance potential-shortfall
'risk model allows correctallocation of assets with non-normally distributed returns as e.g. financial
'derivatives, equities, real estates, fixed return assets, commodities where the mean-variance portfolio
'model tends to inferior asset allocation decisions. The computational issues of the optimization
'algorithm developed for the mean-variance, mean-shortfall risk and chance potential-shortfall risk
'portfolio selection are described to ease their practical application. Additionally, the application
'of the chance potential-shortfall risk model is shown on the asset universe containing stocks, covered
'calls and protective puts.

'http://archiv.tu-chemnitz.de/pub/2005/0085/data/diss_monarch.pdf

'LIBRARY       : PORTFOLIO
'GROUP         : FRONTIER_CO_LPM
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_CO_LPM_FRONTIER_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef TOLERANCE_RNG As Variant = 10, _
Optional ByRef LPM_DEGREE_VAL As Long = 2, _
Optional ByVal TARGET_RETURN As Double = 0.07, _
Optional ByVal BUDGET_VAL As Double = 1, _
Optional ByRef LOWER_RNG As Variant = 0, _
Optional ByRef UPPER_RNG As Variant = 1, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal VERSION As Integer = 0)

Dim i As Long
Dim j As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim SIGMA_VAL As Double

Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant

Dim LOWER_VECTOR As Variant
Dim UPPER_VECTOR As Variant

Dim RETURNS_VECTOR As Variant
Dim TOLERANCE_VECTOR As Variant
Dim SYMMETRIC_MATRIX As Variant
Dim ASYMMETRIC_MATRIX As Variant
Dim COVARIANCE_MATRIX As Variant

On Error GoTo ERROR_LABEL
'---------------------------------------------------------------------------------
DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then
    DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)
End If
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
'---------------------------------------------------------------------------------
If IsArray(TOLERANCE_RNG) = True Then
'---------------------------------------------------------------------------------
    TOLERANCE_VECTOR = TOLERANCE_RNG
    If UBound(TOLERANCE_VECTOR, 1) = 1 Then
        TOLERANCE_VECTOR = MATRIX_TRANSPOSE_FUNC(TOLERANCE_VECTOR)
    End If
    NSIZE = UBound(TOLERANCE_VECTOR, 1)
'---------------------------------------------------------------------------------
Else
'---------------------------------------------------------------------------------
    NSIZE = TOLERANCE_RNG
    ReDim TOLERANCE_VECTOR(1 To NSIZE, 1 To 1)
    For i = 1 To NSIZE
        TOLERANCE_VECTOR(i, 1) = 0.01 * Exp(i - 1)
    Next i
'---------------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------------
If IsArray(LOWER_RNG) = True Then
'---------------------------------------------------------------------------------
    LOWER_VECTOR = LOWER_RNG
    If UBound(LOWER_VECTOR, 1) = 1 Then
        LOWER_VECTOR = MATRIX_TRANSPOSE_FUNC(LOWER_VECTOR)
    End If
    If NCOLUMNS <> UBound(LOWER_VECTOR, 1) Then: GoTo ERROR_LABEL
'---------------------------------------------------------------------------------
Else
'---------------------------------------------------------------------------------
    ReDim LOWER_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NCOLUMNS
        LOWER_VECTOR(i, 1) = LOWER_RNG
    Next i
'---------------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------------
If IsArray(UPPER_RNG) = True Then
'---------------------------------------------------------------------------------
    UPPER_VECTOR = UPPER_RNG
    If UBound(UPPER_VECTOR, 1) = 1 Then
        UPPER_VECTOR = MATRIX_TRANSPOSE_FUNC(UPPER_VECTOR)
    End If
    If NCOLUMNS <> UBound(UPPER_VECTOR, 1) Then: GoTo ERROR_LABEL
'---------------------------------------------------------------------------------
Else
'---------------------------------------------------------------------------------
    ReDim UPPER_VECTOR(1 To NCOLUMNS, 1 To 1)
    For i = 1 To NCOLUMNS
        UPPER_VECTOR(i, 1) = UPPER_RNG
    Next i
'---------------------------------------------------------------------------------
End If
'---------------------------------------------------------------------------------
RETURNS_VECTOR = MATRIX_MEAN_FUNC(DATA_MATRIX)
SYMMETRIC_MATRIX = PORT_SCLPMM_FUNC(TARGET_RETURN, DATA_MATRIX, LPM_DEGREE_VAL)
ASYMMETRIC_MATRIX = PORT_ACLPMM_FUNC(TARGET_RETURN, DATA_MATRIX, LPM_DEGREE_VAL)
COVARIANCE_MATRIX = MATRIX_COVARIANCE_FRAME3_FUNC(DATA_MATRIX, 0, 0)

ReDim TEMP_MATRIX(1 To NCOLUMNS + 6, 1 To NSIZE + 1)
TEMP_MATRIX(1, 1) = "RISK_TOLERANCE"
For i = 1 To NCOLUMNS
    TEMP_MATRIX(1 + i, 1) = "ASSET_WEIGHT" & CStr(i)
Next i
TEMP_MATRIX(NCOLUMNS + 2, 1) = "SUM_WEIGHT"
TEMP_MATRIX(NCOLUMNS + 3, 1) = "PORT_MEAN"
TEMP_MATRIX(NCOLUMNS + 4, 1) = "PORT_VOLATILITY"
TEMP_MATRIX(NCOLUMNS + 5, 1) = "ASYMLPM^(1/N)"
TEMP_MATRIX(NCOLUMNS + 6, 1) = "SYMLPM^(1/N)"

For j = 2 To NSIZE + 1
    TEMP_MATRIX(1, j) = TOLERANCE_VECTOR(j - 1, 1)
'---------------------------------------------------------------------------
    Select Case VERSION
'---------------------------------------------------------------------------
    Case 0 'symmetrical Co-LPM Frontier
'---------------------------------------------------------------------------
        DATA_VECTOR = _
                PORT_SHARPE_WEIGHTS_OPTIMIZER_FUNC(BUDGET_VAL, _
                TEMP_MATRIX(1, j), RETURNS_VECTOR, _
                SYMMETRIC_MATRIX, LOWER_VECTOR, UPPER_VECTOR)
'---------------------------------------------------------------------------
    Case 1 'asymmetrical Co-LPM Frontier --> PERFECT
'---------------------------------------------------------------------------
        DATA_VECTOR = _
                PORT_SHARPE_WEIGHTS_OPTIMIZER_FUNC(BUDGET_VAL, _
                TEMP_MATRIX(1, j), RETURNS_VECTOR, _
                ASYMMETRIC_MATRIX, LOWER_VECTOR, UPPER_VECTOR)
'---------------------------------------------------------------------------
    Case Else 'Historical covariance matrix --> PERFECT
'---------------------------------------------------------------------------
        DATA_VECTOR = _
                PORT_SHARPE_WEIGHTS_OPTIMIZER_FUNC(BUDGET_VAL, _
                TEMP_MATRIX(1, j), RETURNS_VECTOR, _
                COVARIANCE_MATRIX, LOWER_VECTOR, UPPER_VECTOR)
'---------------------------------------------------------------------------
    End Select
'---------------------------------------------------------------------------
    TEMP_SUM = 0
    For i = 1 To NCOLUMNS
        TEMP_MATRIX(1 + i, j) = DATA_VECTOR(i, 1)
        TEMP_SUM = TEMP_SUM + DATA_VECTOR(i, 1)
    Next i
    TEMP_MATRIX(NCOLUMNS + 2, j) = TEMP_SUM
    TEMP_MATRIX(NCOLUMNS + 3, j) = PORT_WEIGHTED_RETURN2_FUNC(RETURNS_VECTOR, DATA_VECTOR)
        
    SIGMA_VAL = PORT_WEIGHTED_SIGMA_COVAR_FUNC(COVARIANCE_MATRIX, DATA_VECTOR)
    TEMP_MATRIX(NCOLUMNS + 4, j) = SIGMA_VAL
        
    SIGMA_VAL = PORT_WEIGHTED_SIGMA_COVAR_FUNC(ASYMMETRIC_MATRIX, DATA_VECTOR)
    TEMP_MATRIX(NCOLUMNS + 5, j) = (SIGMA_VAL ^ 2) ^ (1 / LPM_DEGREE_VAL)
        
    SIGMA_VAL = PORT_WEIGHTED_SIGMA_COVAR_FUNC(SYMMETRIC_MATRIX, DATA_VECTOR)
    TEMP_MATRIX(NCOLUMNS + 6, j) = (SIGMA_VAL ^ 2) ^ (1 / LPM_DEGREE_VAL)
Next j

PORT_CO_LPM_FRONTIER_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_CO_LPM_FRONTIER_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_ACLPMM_FUNC
'DESCRIPTION   : Asymmetric Co-Lower Partial Moment Matrix
'LIBRARY       : PORTFOLIO
'GROUP         : FRONTIER_CO_LPM
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Private Function PORT_ACLPMM_FUNC(ByVal TARGET_RETURN As Double, _
ByRef DATA_RNG As Variant, _
ByVal LPM_DEGREE_VAL As Long, _
Optional ByVal VERSION As Integer = 1)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP1_VECTOR(1 To NROWS, 1 To 1)
ReDim TEMP2_VECTOR(1 To NROWS, 1 To 1)
ReDim TEMP_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)

For i = 1 To NCOLUMNS
    For j = i To NCOLUMNS
        For k = 1 To NROWS
            TEMP1_VECTOR(k, 1) = DATA_MATRIX(k, i)
            TEMP2_VECTOR(k, 1) = DATA_MATRIX(k, j)
        Next k
        If i = j Then
            TEMP_MATRIX(i, j) = PORT_LPM_FUNC(TARGET_RETURN, TEMP1_VECTOR, LPM_DEGREE_VAL)
        Else
            TEMP_MATRIX(i, j) = PORT_CLPM_FUNC(TARGET_RETURN, TEMP1_VECTOR, TEMP2_VECTOR, LPM_DEGREE_VAL, VERSION)
            TEMP_MATRIX(j, i) = TEMP_MATRIX(i, j)
        End If
    Next j
Next i
PORT_ACLPMM_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_ACLPMM_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_SCLPMM_FUNC
'DESCRIPTION   : Symmetric Co-Lower Partial Moment Matrix
'LIBRARY       : PORTFOLIO
'GROUP         : FRONTIER_CO_LPM
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Private Function PORT_SCLPMM_FUNC(ByVal TARGET_RETURN As Double, _
ByRef DATA_RNG As Variant, _
ByVal LPM_DEGREE_VAL As Long)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP1_VECTOR As Variant
Dim TEMP2_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP1_VECTOR(1 To NROWS, 1 To 1)
ReDim TEMP2_VECTOR(1 To NROWS, 1 To 1)

ReDim TEMP_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)

For i = 1 To NCOLUMNS
    For j = i To NCOLUMNS
        For k = 1 To NROWS
            TEMP1_VECTOR(k, 1) = DATA_MATRIX(k, i)
            TEMP2_VECTOR(k, 1) = DATA_MATRIX(k, j)
        Next k
        TEMP_MATRIX(i, j) = CORRELATION_FUNC(TEMP1_VECTOR, TEMP2_VECTOR, 0, 0) * (PORT_LPM_FUNC(TARGET_RETURN, TEMP1_VECTOR, LPM_DEGREE_VAL) ^ (1 / LPM_DEGREE_VAL)) * (PORT_LPM_FUNC(TARGET_RETURN, TEMP2_VECTOR, LPM_DEGREE_VAL) ^ (1 / LPM_DEGREE_VAL))
        TEMP_MATRIX(j, i) = TEMP_MATRIX(i, j)
    Next j
Next i
PORT_SCLPMM_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_SCLPMM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_CLPM_FUNC
'DESCRIPTION   : Co-Lower Partial Moment
'LIBRARY       : PORTFOLIO
'GROUP         : FRONTIER_CO_LPM
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Private Function PORT_CLPM_FUNC(ByVal TARGET_RETURN As Double, _
ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
ByVal LPM_DEGREE_VAL As Long, _
Optional ByVal VERSION As Integer = 1)

Dim i As Long
Dim NROWS As Long
Dim TEMP_SUM As Double
Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA1_VECTOR = DATA1_RNG 'ReturnSeries 1
If UBound(DATA1_VECTOR, 1) = 1 Then
    DATA1_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA1_VECTOR)
End If

DATA2_VECTOR = DATA2_RNG 'ReturnSeries 2
If UBound(DATA2_VECTOR, 1) = 1 Then
    DATA2_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA2_VECTOR)
End If

NROWS = UBound(DATA1_VECTOR, 1)
TEMP_SUM = 0
'--------------------------------------------------------------------------
For i = 1 To NROWS
'--------------------------------------------------------------------------
    Select Case VERSION
'--------------------------------------------------------------------------
    Case 0 ' Traditional formula, requires n>1, i.e. risk aversion
'--------------------------------------------------------------------------
        TEMP_SUM = TEMP_SUM + (MAXIMUM_FUNC(0, TARGET_RETURN - DATA1_VECTOR(i, 1)) ^ (LPM_DEGREE_VAL - 1)) * (TARGET_RETURN - DATA2_VECTOR(i, 1))
'--------------------------------------------------------------------------
    Case Else ' Enhanced formula
'--------------------------------------------------------------------------
        TEMP_SUM = TEMP_SUM + (MAXIMUM_FUNC(0, TARGET_RETURN - DATA1_VECTOR(i, 1)) ^ (LPM_DEGREE_VAL / 2)) * ((Abs(TARGET_RETURN - DATA2_VECTOR(i, 1)) ^ (LPM_DEGREE_VAL / 2)) * Sgn(TARGET_RETURN - DATA2_VECTOR(i, 1)))
'--------------------------------------------------------------------------
    End Select
'--------------------------------------------------------------------------
Next i
'--------------------------------------------------------------------------
PORT_CLPM_FUNC = TEMP_SUM / NROWS

Exit Function
ERROR_LABEL:
PORT_CLPM_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_LPM_FUNC
'DESCRIPTION   : Lower Partial Moment
'LIBRARY       : PORTFOLIO
'GROUP         : FRONTIER_CO_LPM
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Private Function PORT_LPM_FUNC(ByVal TARGET_RETURN As Double, _
ByRef DATA_RNG As Variant, _
ByVal LPM_DEGREE_VAL As Long)

Dim i As Long
Dim NROWS As Long
Dim TEMP_SUM As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If

NROWS = UBound(DATA_VECTOR, 1)
TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_SUM = TEMP_SUM + MAXIMUM_FUNC(0, TARGET_RETURN - DATA_VECTOR(i, 1)) ^ LPM_DEGREE_VAL
Next i

PORT_LPM_FUNC = TEMP_SUM / NROWS

Exit Function
ERROR_LABEL:
PORT_LPM_FUNC = Err.number
End Function
