Attribute VB_Name = "FINAN_PORT_RISK_PCA_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_PRINCIPAL_COMPONENTS_FUNC

'DESCRIPTION   : Statistical Factor Modeling (based on principal components)
'Implementation of a model based statistical factors derived from principal
'component analysis (eigenvalues, eigenvectores). Also known as "implicit factor
'model" in the literature. It computs total asset variances from betas, residual
'variances and factor covariances

'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_FACTORS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 07/07/2010
'************************************************************************************
'************************************************************************************

Function PORT_PRINCIPAL_COMPONENTS_FUNC(ByRef DATA_RNG As Variant, _
ByRef WEIGHTS_RNG As Variant, _
Optional ByVal NO_FACTORS As Long = 7)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long
Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

NROWS = NO_FACTORS + 7
NCOLUMNS = NO_FACTORS + 1
ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS: For i = 1 To NROWS: TEMP_MATRIX(i, j) = "": Next i: Next j

TEMP_MATRIX(1, 1) = "% OF VARIANCE EXPLAINED"
TEMP_MATRIX(2, 1) = "IDX"
TEMP_MATRIX(3, 1) = "EIGENVALUE"
TEMP_MATRIX(4, 1) = "FACTOR ALPHAS"
For i = 1 To NO_FACTORS
    TEMP_MATRIX(4 + i, 1) = "FACTOR BETAS " & CStr(i)
Next i

TEMP_MATRIX(4 + NO_FACTORS + 1, 1) = "VARIANCES"
TEMP_MATRIX(4 + NO_FACTORS + 2, 1) = "PORT VARIANCE"
TEMP_MATRIX(4 + NO_FACTORS + 3, 1) = "PORT STDEV"

DATA_VECTOR = MATRIX_CORRELATION_FUNC(DATA_RNG)
DATA_VECTOR = PORT_FACTOR_VARIANCE_EXPLAINED_FUNC(DATA_VECTOR, 0)
For i = 2 To NCOLUMNS
    j = i - 1
    TEMP_MATRIX(1, i) = DATA_VECTOR(1, j)
    TEMP_MATRIX(2, i) = DATA_VECTOR(2, j)
    TEMP_MATRIX(3, i) = DATA_VECTOR(3, j)
Next i
DATA_VECTOR = PORT_FACTOR_BUILD_FACTOR_ALPHAS_FUNC(DATA_RNG, NO_FACTORS)
For i = 2 To NCOLUMNS
    j = i - 1
    TEMP_MATRIX(4, i) = DATA_VECTOR(1, j)
Next i
DATA_VECTOR = PORT_FACTOR_BUILD_FACTOR_BETAS_FUNC(DATA_RNG, NO_FACTORS)
For i = 2 To NCOLUMNS
    For j = 1 To NO_FACTORS
        TEMP_MATRIX(4 + j, i) = DATA_VECTOR(j, i - 1)
    Next j
Next i
DATA_VECTOR = MATRIX_STDEVP_FUNC(DATA_RNG)
For i = 2 To NCOLUMNS
    j = i - 1
    TEMP_MATRIX(4 + NO_FACTORS + 1, i) = DATA_VECTOR(1, j) ^ 2
Next i
TEMP_MATRIX(5 + NO_FACTORS + 1, 2) = PORT_FACTOR_VARIANCE_FUNC(DATA_RNG, NO_FACTORS, WEIGHTS_RNG)(1, 1)
TEMP_MATRIX(6 + NO_FACTORS + 1, 2) = TEMP_MATRIX(5 + NO_FACTORS + 1, 2) ^ 0.5

PORT_PRINCIPAL_COMPONENTS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_PRINCIPAL_COMPONENTS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_FACTOR_VARIANCE_EXPLAINED_FUNC
'DESCRIPTION   : Computes total asset variances from betas
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_FACTORS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 07/07/2010
'************************************************************************************
'************************************************************************************

Function PORT_FACTOR_VARIANCE_EXPLAINED_FUNC( _
ByRef CORREL_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant
Dim CORREL_MATRIX As Variant
Dim EIGENVALUES_VECTOR As Variant
Dim SORTED_EIGENVALUES_VECTOR As Variant

On Error GoTo ERROR_LABEL

CORREL_MATRIX = CORREL_RNG
NROWS = UBound(CORREL_MATRIX, 1)
' Calculate vectorized Eigenvalues
EIGENVALUES_VECTOR = MATRIX_PCA_FUNC(CORREL_MATRIX, False, 1)

' Calculate percentages
TEMP_SUM = 0
For i = 1 To NROWS
    TEMP_SUM = TEMP_SUM + EIGENVALUES_VECTOR(i, 1)
Next i
For i = 1 To NROWS
    EIGENVALUES_VECTOR(i, 1) = EIGENVALUES_VECTOR(i, 1) / TEMP_SUM
Next i
SORTED_EIGENVALUES_VECTOR = MATRIX_QUICK_SORT_FUNC(EIGENVALUES_VECTOR, 1, 0)
ReDim TEMP_MATRIX(1 To 3, 1 To NROWS)
For i = 1 To NROWS
    TEMP_MATRIX(1, i) = SORTED_EIGENVALUES_VECTOR(i, 1)
    For j = 1 To NROWS
        If EIGENVALUES_VECTOR(j, 1) = TEMP_MATRIX(1, i) Then
            TEMP_MATRIX(2, i) = j
            TEMP_MATRIX(3, i) = EIGENVALUES_VECTOR(j, 1) * TEMP_SUM
            Exit For
        End If
    Next j
Next i

Select Case OUTPUT
Case 0
    PORT_FACTOR_VARIANCE_EXPLAINED_FUNC = TEMP_MATRIX
Case 1
    PORT_FACTOR_VARIANCE_EXPLAINED_FUNC = MATRIX_QUICK_SORT_FUNC(MATRIX_TRANSPOSE_FUNC(TEMP_MATRIX), 2, 1)
Case Else
    PORT_FACTOR_VARIANCE_EXPLAINED_FUNC = SORTED_EIGENVALUES_VECTOR
End Select

Exit Function
ERROR_LABEL:
PORT_FACTOR_VARIANCE_EXPLAINED_FUNC = Err.number
End Function

Function PORT_FACTOR_WEIGHTS_FUNC(ByRef CORREL_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim NROWS As Long
Dim TEMP1_SUM As Double
Dim V_VECTOR As Variant 'EigenVectors

On Error GoTo ERROR_LABEL

V_VECTOR = MATRIX_PCA_FUNC(CORREL_RNG, True, 0)
NROWS = UBound(V_VECTOR, 1)
ReDim Preserve V_VECTOR(1 To NROWS, 1 To 1)

Select Case OUTPUT
Case 0 'Weights
    TEMP1_SUM = 0
    For i = 1 To NROWS: TEMP1_SUM = TEMP1_SUM + V_VECTOR(i, 1): Next i
    For i = 1 To NROWS
        V_VECTOR(i, 1) = V_VECTOR(i, 1) / TEMP1_SUM
    Next i
    PORT_FACTOR_WEIGHTS_FUNC = V_VECTOR
Case Else
    If OUTPUT = 1 Then
        PORT_FACTOR_WEIGHTS_FUNC = V_VECTOR
        Exit Function
    End If
    Dim TEMP2_SUM As Double
    Dim TEMP3_SUM As Double
    Dim AV_VECTOR As Variant
    AV_VECTOR = MMULT_FUNC(CORREL_RNG, V_VECTOR, 70)
    If OUTPUT = 2 Then
        PORT_FACTOR_WEIGHTS_FUNC = AV_VECTOR
    ElseIf OUTPUT = 3 Then 'Max. Eigenvalue
        TEMP2_SUM = 0
        For i = 1 To NROWS: TEMP2_SUM = TEMP2_SUM + AV_VECTOR(i, 1) ^ 2: Next i
        PORT_FACTOR_WEIGHTS_FUNC = TEMP2_SUM ^ 0.5
    Else
        TEMP1_SUM = 0: TEMP2_SUM = 0: TEMP3_SUM = 0
        For i = 1 To NROWS
            TEMP1_SUM = TEMP1_SUM + V_VECTOR(i, 1) * AV_VECTOR(i, 1)
            TEMP2_SUM = TEMP2_SUM + V_VECTOR(i, 1) ^ 2
            TEMP3_SUM = TEMP3_SUM + AV_VECTOR(i, 1) ^ 2
        Next i
        PORT_FACTOR_WEIGHTS_FUNC = ACOS_FUNC(TEMP1_SUM / TEMP2_SUM ^ 0.5 / TEMP3_SUM ^ 0.5)
    End If
End Select

Exit Function
ERROR_LABEL:
PORT_FACTOR_WEIGHTS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_FACTOR_BUILD_FACTORS_FUNC
'DESCRIPTION   : computes total asset factors
'factor covariances
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_FACTORS
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 07/07/2010
'************************************************************************************
'************************************************************************************

Function PORT_FACTOR_BUILD_FACTORS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NFACTORS As Long = 4)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim DATA_VECTOR As Variant
Dim EIGENVECTORS_MATRIX As Variant

On Error GoTo ERROR_LABEL

' Get sorted factor indices
DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
DATA_VECTOR = MATRIX_CORRELATION_FUNC(DATA_MATRIX)
EIGENVECTORS_MATRIX = MATRIX_PCA_FUNC(DATA_VECTOR, False, 0)
DATA_VECTOR = PORT_FACTOR_VARIANCE_EXPLAINED_FUNC(DATA_VECTOR, 0)
' Calculate matrix with eigenvalue vectors

NCOLUMNS = UBound(EIGENVECTORS_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NROWS, 1 To NFACTORS)
For i = 1 To NROWS
    For j = 1 To NFACTORS
        For k = 1 To NCOLUMNS
            l = DATA_VECTOR(2, j)
            TEMP_MATRIX(i, j) = TEMP_MATRIX(i, j) + DATA_MATRIX(i, k) * EIGENVECTORS_MATRIX(k, l)
        Next k
    Next j
Next i
PORT_FACTOR_BUILD_FACTORS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_FACTOR_BUILD_FACTORS_FUNC = TEMP_MATRIX
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_FACTOR_BUILD_FACTOR_BETAS_FUNC
'DESCRIPTION   : computes total asset betas
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_FACTORS
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 07/07/2010
'************************************************************************************
'************************************************************************************

Function PORT_FACTOR_BUILD_FACTOR_BETAS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NFACTORS As Long = 4)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim FF_MATRIX As Variant
Dim YDATA_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim OLS_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim FACTOR_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
FACTOR_MATRIX = PORT_FACTOR_BUILD_FACTORS_FUNC(DATA_MATRIX, NFACTORS)

ReDim FF_MATRIX(1 To NROWS, 1 To NFACTORS + 1)
For i = 1 To NROWS
    FF_MATRIX(i, 1) = 1
    For j = 1 To NFACTORS
        FF_MATRIX(i, 1 + j) = FACTOR_MATRIX(i, j)
    Next j
Next i

ReDim YDATA_VECTOR(1 To NROWS, 1 To 1)
ReDim TEMP_MATRIX(1 To NFACTORS, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        YDATA_VECTOR(i, 1) = DATA_MATRIX(i, j)
    Next i
    
    OLS_MATRIX = MMULT_FUNC(MMULT_FUNC(MATRIX_INVERSE_FUNC(MMULT_FUNC(MATRIX_TRANSPOSE_FUNC(FF_MATRIX), FF_MATRIX, 70), 0), MATRIX_TRANSPOSE_FUNC(FF_MATRIX), 70), YDATA_VECTOR, 70)
    For k = 1 To NFACTORS
        TEMP_MATRIX(k, j) = OLS_MATRIX(k + 1, 1)
    Next k
Next j
PORT_FACTOR_BUILD_FACTOR_BETAS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_FACTOR_BUILD_FACTOR_BETAS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_FACTOR_BUILD_FACTOR_ALPHAS_FUNC
'DESCRIPTION   : computes total asset alphas
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_FACTORS
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 07/07/2010
'************************************************************************************
'************************************************************************************

Function PORT_FACTOR_BUILD_FACTOR_ALPHAS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NFACTORS As Long = 4)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim FF_MATRIX As Variant
Dim YDATA_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim OLS_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim FACTOR_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
FACTOR_MATRIX = PORT_FACTOR_BUILD_FACTORS_FUNC(DATA_MATRIX, NFACTORS)

ReDim FF_MATRIX(1 To NROWS, 1 To NFACTORS + 1)
For i = 1 To NROWS
    FF_MATRIX(i, 1) = 1
    For j = 1 To NFACTORS
        FF_MATRIX(i, 1 + j) = FACTOR_MATRIX(i, j)
    Next j
Next i

ReDim YDATA_VECTOR(1 To NROWS, 1 To 1)
ReDim TEMP_MATRIX(1 To 1, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        YDATA_VECTOR(i, 1) = DATA_MATRIX(i, j)
    Next i
    OLS_MATRIX = MMULT_FUNC(MMULT_FUNC(MATRIX_INVERSE_FUNC(MMULT_FUNC(MATRIX_TRANSPOSE_FUNC(FF_MATRIX), FF_MATRIX, 70), 0), MATRIX_TRANSPOSE_FUNC(FF_MATRIX), 70), YDATA_VECTOR, 70)
    TEMP_MATRIX(1, j) = OLS_MATRIX(1, 1)
Next j
PORT_FACTOR_BUILD_FACTOR_ALPHAS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_FACTOR_BUILD_FACTOR_ALPHAS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_FACTOR_BUILD_FACTOR_RESIDUALS_FUNC
'DESCRIPTION   : computes asset returns unexplained by factors (eta of a
'regression if one asset on factors)
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_FACTORS
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 07/07/2010
'************************************************************************************
'************************************************************************************

Function PORT_FACTOR_BUILD_FACTOR_RESIDUALS_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NFACTORS As Long = 4)

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim FF_MATRIX As Variant
Dim YDATA_VECTOR As Variant
Dim YHAT_VECTOR As Variant
Dim DATA_MATRIX As Variant
Dim OLS_MATRIX As Variant
Dim TEMP_MATRIX As Variant
Dim FACTOR_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)
FACTOR_MATRIX = PORT_FACTOR_BUILD_FACTORS_FUNC(DATA_MATRIX, NFACTORS)

ReDim FF_MATRIX(1 To NROWS, 1 To NFACTORS + 1)
For i = 1 To NROWS
    FF_MATRIX(i, 1) = 1
    For j = 1 To NFACTORS
        FF_MATRIX(i, 1 + j) = FACTOR_MATRIX(i, j)
    Next j
Next i

ReDim YDATA_VECTOR(1 To NROWS, 1 To 1)
ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
For j = 1 To NCOLUMNS
    For i = 1 To NROWS
        YDATA_VECTOR(i, 1) = DATA_MATRIX(i, j)
    Next i
    OLS_MATRIX = MMULT_FUNC(MMULT_FUNC(MATRIX_INVERSE_FUNC(MMULT_FUNC(MATRIX_TRANSPOSE_FUNC(FF_MATRIX), FF_MATRIX, 70), 0), MATRIX_TRANSPOSE_FUNC(FF_MATRIX), 70), YDATA_VECTOR, 70)
    YHAT_VECTOR = MMULT_FUNC(FF_MATRIX, OLS_MATRIX, 70)
    For i = 1 To NROWS
        TEMP_MATRIX(i, j) = YHAT_VECTOR(i, 1) - YDATA_VECTOR(i, 1)
    Next i
Next j

PORT_FACTOR_BUILD_FACTOR_RESIDUALS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_FACTOR_BUILD_FACTOR_RESIDUALS_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_FACTOR_BUILD_FACTOR_WEIGHTS_FUNC
'DESCRIPTION   : Build factor weights
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_FACTORS
'ID            : 007
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 07/07/2010
'************************************************************************************
'************************************************************************************

Function PORT_FACTOR_BUILD_FACTOR_WEIGHTS_FUNC(ByRef DATA_RNG As Variant, _
ByVal NFACTORS As Long, _
ByRef WEIGHTS_RNG As Variant)

Dim WEIGHTS_VECTOR As Variant

On Error GoTo ERROR_LABEL

WEIGHTS_VECTOR = WEIGHTS_RNG
If UBound(WEIGHTS_VECTOR, 2) = 1 Then
    WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
End If
PORT_FACTOR_BUILD_FACTOR_WEIGHTS_FUNC = MATRIX_TRANSPOSE_FUNC(MMULT_FUNC(PORT_FACTOR_BUILD_FACTOR_BETAS_FUNC(DATA_RNG, NFACTORS), MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR), 70))

Exit Function
ERROR_LABEL:
PORT_FACTOR_BUILD_FACTOR_WEIGHTS_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_FACTOR_VARIANCE_FUNC
'DESCRIPTION   : Build factor variance
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_FACTORS
'ID            : 008
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 07/07/2010
'************************************************************************************
'************************************************************************************

Function PORT_FACTOR_VARIANCE_FUNC(ByRef DATA_RNG As Variant, _
ByVal NFACTORS As Long, _
ByRef WEIGHTS_RNG As Variant)

Dim FCOVAR_MATRIX As Variant
Dim SCOVAR_MATRIX As Variant
Dim FACTOR_MATRIX As Variant
Dim WEIGHTS_VECTOR As Variant

On Error GoTo ERROR_LABEL

WEIGHTS_VECTOR = WEIGHTS_RNG
If UBound(WEIGHTS_VECTOR, 2) = 1 Then
    WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
End If

FACTOR_MATRIX = PORT_FACTOR_BUILD_FACTOR_BETAS_FUNC(DATA_RNG, NFACTORS)
FCOVAR_MATRIX = MATRIX_COVARIANCE_FRAME1_FUNC(PORT_FACTOR_BUILD_FACTORS_FUNC(DATA_RNG, NFACTORS), 0, 0)
'Covariance Matrix of Factors
SCOVAR_MATRIX = MATRIX_COVARIANCE_FRAME1_FUNC(PORT_FACTOR_BUILD_FACTOR_RESIDUALS_FUNC(DATA_RNG, NFACTORS), 0, 0)
'Covariance Matrix of Residuals

PORT_FACTOR_VARIANCE_FUNC = MMULT_FUNC(MMULT_FUNC(WEIGHTS_VECTOR, MATRIX_ELEMENTS_ADD_FUNC(MMULT_FUNC(MMULT_FUNC(MATRIX_TRANSPOSE_FUNC(FACTOR_MATRIX), FCOVAR_MATRIX, 70), FACTOR_MATRIX, 70), SCOVAR_MATRIX, 1, 1), 70), MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR), 70)
'Same as A + B
'A = VAR(SUM_PRODUCT(PORT_FACTOR_BUILD_FACTORS_FUNC,PORT_FACTOR_BUILD_FACTOR_WEIGHTS_FUNC))
'B = VAR(SUM_PRODUCT(PORT_FACTOR_BUILD_FACTOR_RESIDUALS_FUNC,INITIAL_WEIGHTS))

Exit Function
ERROR_LABEL:
PORT_FACTOR_VARIANCE_FUNC = Err.number
End Function
