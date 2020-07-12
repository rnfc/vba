Attribute VB_Name = "FINAN_PORT_RISK_RESIDUAL_LIBR"

'---------------------------------------------------------------------------------
'---------------------------------------------------------------------------------
Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_RESIDUAL_COVARIANCE_FUNC
'DESCRIPTION   : Computes volatility vector from betas, residual variances and
'factor covariances
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_RESIDUAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PORT_RESIDUAL_COVARIANCE_FUNC(ByRef BETA_RNG As Variant, _
ByRef RESIDUAL_VARIANCE_RNG As Variant, _
ByRef FACTOR_COVARIANCE_RNG As Variant)

Dim i As Long
Dim NASSETS As Long
Dim NFACTORS As Long

Dim TEMP_MATRIX As Variant
Dim BETA_VECTOR As Variant
Dim VARIANCE_VECTOR As Variant

Dim COVARIANCE_MATRIX As Variant

On Error GoTo ERROR_LABEL

If IsArray(FACTOR_COVARIANCE_RNG) = True Then 'n-Factor
    COVARIANCE_MATRIX = FACTOR_COVARIANCE_RNG
Else 'One Factor
    ReDim COVARIANCE_MATRIX(1 To 1, 1 To 1)
    COVARIANCE_MATRIX(1, 1) = FACTOR_COVARIANCE_RNG
End If

BETA_VECTOR = BETA_RNG ' number of assets
NASSETS = UBound(BETA_VECTOR, 1) ' number of factors
NFACTORS = UBound(BETA_VECTOR, 2) ' calculate systematic covariances

VARIANCE_VECTOR = RESIDUAL_VARIANCE_RNG
If UBound(VARIANCE_VECTOR, 1) = 1 Then: _
VARIANCE_VECTOR = MATRIX_TRANSPOSE_FUNC(VARIANCE_VECTOR)

TEMP_MATRIX = MMULT_FUNC(MMULT_FUNC(BETA_VECTOR, COVARIANCE_MATRIX, 70), MATRIX_TRANSPOSE_FUNC(BETA_VECTOR), 70)
For i = 1 To NASSETS ' add residual risk
    TEMP_MATRIX(i, i) = TEMP_MATRIX(i, i) + VARIANCE_VECTOR(i, 1)
Next i

PORT_RESIDUAL_COVARIANCE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_RESIDUAL_COVARIANCE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_RESIDUAL_VARIANCE_FUNC
'DESCRIPTION   : Computes total asset variances from betas, residual variances
' and factor covariances
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_RESIDUAL
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PORT_RESIDUAL_VARIANCE_FUNC(ByRef BETA_RNG As Variant, _
ByRef RESIDUAL_VARIANCE_RNG As Variant, _
ByRef FACTOR_COVARIANCE_RNG As Variant)

Dim i As Long
Dim NASSETS As Long
Dim TEMP_MATRIX As Variant
Dim FACTOR_COVARIANCE_MATRIX As Variant

On Error GoTo ERROR_LABEL

' Calc full covariance matrix
FACTOR_COVARIANCE_MATRIX = PORT_RESIDUAL_COVARIANCE_FUNC(BETA_RNG, RESIDUAL_VARIANCE_RNG, FACTOR_COVARIANCE_RNG)
' number of assets
NASSETS = UBound(FACTOR_COVARIANCE_MATRIX, 1)

' get diagonal of covariance matrix
ReDim TEMP_MATRIX(1 To NASSETS, 1 To 1)
For i = 1 To NASSETS
    TEMP_MATRIX(i, 1) = FACTOR_COVARIANCE_MATRIX(i, i)
Next i

PORT_RESIDUAL_VARIANCE_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_RESIDUAL_VARIANCE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_RESIDUAL_CORRELATION_FUNC
'DESCRIPTION   : Computes asset correlations from betas, residual variances
' and factor covariances
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_RESIDUAL
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PORT_RESIDUAL_CORRELATION_FUNC(ByRef BETA_RNG As Variant, _
ByRef RESIDUAL_VARIANCE_RNG As Variant, _
ByRef FACTOR_COVARIANCE_RNG As Variant)

Dim i As Long
Dim j As Long
Dim NASSETS As Long

Dim TEMP_MATRIX As Variant
Dim VARIANCE_VECTOR As Variant
Dim FACTOR_COVARIANCE_MATRIX As Variant

On Error GoTo ERROR_LABEL

' Calc full covariance matrix
FACTOR_COVARIANCE_MATRIX = PORT_RESIDUAL_COVARIANCE_FUNC(BETA_RNG, RESIDUAL_VARIANCE_RNG, FACTOR_COVARIANCE_RNG)
' Calc variance vector
VARIANCE_VECTOR = PORT_RESIDUAL_VARIANCE_FUNC(BETA_RNG, RESIDUAL_VARIANCE_RNG, FACTOR_COVARIANCE_RNG)
' number of assets
NASSETS = UBound(FACTOR_COVARIANCE_MATRIX, 1)
' get diagonal of covariance matrix
ReDim TEMP_MATRIX(1 To NASSETS, 1 To NASSETS)
For i = 1 To NASSETS
    For j = 1 To NASSETS
        TEMP_MATRIX(i, j) = FACTOR_COVARIANCE_MATRIX(i, j) / Sqr(VARIANCE_VECTOR(i, 1) * VARIANCE_VECTOR(j, 1))
    Next j
Next i

PORT_RESIDUAL_CORRELATION_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_RESIDUAL_CORRELATION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_RESIDUAL_RETURNS_FUNC
'DESCRIPTION   : Computes expected asset returns from betas, residual returns
' and factor returns
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_RESIDUAL
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PORT_RESIDUAL_RETURNS_FUNC(ByRef BETA_RNG As Variant, _
ByRef RESIDUAL_RETURN_RNG As Variant, _
ByRef FACTOR_EXPECTED_RETURN_RNG As Variant)

Dim i As Long
Dim NASSETS As Long

Dim TEMP_MATRIX As Variant
Dim BETA_VECTOR As Variant
Dim RESIDUAL_RETURN_VECTOR As Variant
Dim FACTOR_EXPECTED_RETURN_VECTOR As Variant

On Error GoTo ERROR_LABEL

BETA_VECTOR = BETA_RNG

If IsArray(FACTOR_EXPECTED_RETURN_RNG) = True Then
    FACTOR_EXPECTED_RETURN_VECTOR = FACTOR_EXPECTED_RETURN_RNG
    If UBound(FACTOR_EXPECTED_RETURN_VECTOR, 1) = 1 Then
        FACTOR_EXPECTED_RETURN_VECTOR = _
        MATRIX_TRANSPOSE_FUNC(FACTOR_EXPECTED_RETURN_VECTOR)
    End If
Else
    ReDim FACTOR_EXPECTED_RETURN_VECTOR(1 To 1, 1 To 1)
    FACTOR_EXPECTED_RETURN_VECTOR(1, 1) = FACTOR_EXPECTED_RETURN_RNG
End If

RESIDUAL_RETURN_VECTOR = RESIDUAL_RETURN_RNG
If UBound(RESIDUAL_RETURN_VECTOR, 1) = 1 Then
    RESIDUAL_RETURN_VECTOR = MATRIX_TRANSPOSE_FUNC(RESIDUAL_RETURN_VECTOR)
End If

' number of assets
NASSETS = UBound(BETA_VECTOR, 1)

' calculate systematic return
TEMP_MATRIX = MMULT_FUNC(BETA_VECTOR, FACTOR_EXPECTED_RETURN_VECTOR, 70)

For i = 1 To NASSETS ' add residual return
    TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 1) + RESIDUAL_RETURN_VECTOR(i, 1)
Next i

PORT_RESIDUAL_RETURNS_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_RESIDUAL_RETURNS_FUNC = Err.number
End Function
