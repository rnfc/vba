Attribute VB_Name = "FINAN_PORT_FRONTIER_MARKOW_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_MARKOWITZ_FRONTIER_FUNC

'DESCRIPTION   : Markowitz 's celebrated mean-variance portfolio
'optimization theory assumes that the means and covariances of the
'underlying asset returns are known. In practice, they are unknown
'and have to be estimated from historical data. Plug them into the
'efficient frontier that assumes know parameters leads to the
'so-called "Markowitz enigma", which states that portfolio with
'the "plug-in" efficient frontier can behave badly and be
'counter-intuitive.

'LIBRARY       : PORTFOLIO
'GROUP         : FRONTIER_MARKOWITZ
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************
'PORT_MARKOWITZ_EXPECTED_ALLOCATION_FUNC

Function PORT_MARKOWITZ_FRONTIER_FUNC(ByVal BUDGET_VAL As Variant, _
ByVal RISK_TOLER_RNG As Variant, _
ByRef EXPECTED_RNG As Variant, _
ByRef COVAR_RNG As Variant, _
Optional ByRef LOWER_RNG As Variant = 0, _
Optional ByRef UPPER_RNG As Variant = 1, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim COVARIANCES_MATRIX As Variant
Dim LOWER_VECTOR As Variant
Dim UPPER_VECTOR As Variant
Dim EXPECTED_VECTOR As Variant
Dim RISK_TOLER_VECTOR As Variant

Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

EXPECTED_VECTOR = EXPECTED_RNG
If UBound(EXPECTED_VECTOR, 1) = 1 Then: _
    EXPECTED_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPECTED_VECTOR)
NROWS = UBound(EXPECTED_VECTOR, 1)

COVARIANCES_MATRIX = COVAR_RNG
If UBound(COVARIANCES_MATRIX, 1) <> UBound(COVARIANCES_MATRIX, 2) Then: GoTo ERROR_LABEL
If UBound(COVARIANCES_MATRIX, 1) <> UBound(EXPECTED_VECTOR, 1) Then: GoTo ERROR_LABEL


If IsArray(RISK_TOLER_RNG) = True Then
    RISK_TOLER_VECTOR = RISK_TOLER_RNG
    If UBound(RISK_TOLER_VECTOR, 1) = 1 Then: RISK_TOLER_VECTOR = MATRIX_TRANSPOSE_FUNC(RISK_TOLER_VECTOR)
    If UBound(COVARIANCES_MATRIX, 1) <> UBound(RISK_TOLER_VECTOR, 1) Then: GoTo ERROR_LABEL
Else
    ReDim RISK_TOLER_VECTOR(1 To 1, 1 To 1)
    RISK_TOLER_VECTOR(1, 1) = RISK_TOLER_RNG
End If

NCOLUMNS = UBound(RISK_TOLER_VECTOR, 1)

If IsArray(LOWER_RNG) = True Then
    LOWER_VECTOR = LOWER_RNG
    If UBound(LOWER_VECTOR, 1) = 1 Then: _
        LOWER_VECTOR = MATRIX_TRANSPOSE_FUNC(LOWER_VECTOR)
    If UBound(COVARIANCES_MATRIX, 1) <> UBound(LOWER_VECTOR, 1) Then: GoTo ERROR_LABEL
Else
    ReDim LOWER_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        LOWER_VECTOR(i, 1) = LOWER_RNG
    Next i
End If

If IsArray(UPPER_RNG) = True Then
    UPPER_VECTOR = UPPER_RNG
    If UBound(UPPER_VECTOR, 1) = 1 Then: _
        UPPER_VECTOR = MATRIX_TRANSPOSE_FUNC(UPPER_VECTOR)
    If UBound(COVARIANCES_MATRIX, 1) <> UBound(UPPER_VECTOR, 1) Then: GoTo ERROR_LABEL
Else
    ReDim UPPER_VECTOR(1 To NROWS, 1 To 1)
    For i = 1 To NROWS
        UPPER_VECTOR(i, 1) = UPPER_RNG
    Next i
End If

ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    TEMP_VECTOR = PORT_SHARPE_WEIGHTS_OPTIMIZER_FUNC(BUDGET_VAL, RISK_TOLER_VECTOR(j, 1), EXPECTED_VECTOR, COVARIANCES_MATRIX, LOWER_VECTOR, UPPER_VECTOR)
    For i = 1 To NROWS
        TEMP_MATRIX(i, j) = TEMP_VECTOR(i, 1)
    Next i
Next j

If OUTPUT = 1 Then
    PORT_MARKOWITZ_FRONTIER_FUNC = TEMP_MATRIX
    Exit Function
End If
        
ReDim TEMP_VECTOR(1 To 2, 1 To NROWS)
For i = 1 To NROWS
    TEMP_VECTOR(1, i) = PORT_WEIGHTED_RETURN2_FUNC(EXPECTED_VECTOR, MATRIX_GET_COLUMN_FUNC(TEMP_MATRIX, i, 1)) 'Optimal Portfolio Return
    TEMP_VECTOR(2, i) = PORT_WEIGHTED_SIGMA_COVAR_FUNC(COVARIANCES_MATRIX, MATRIX_GET_COLUMN_FUNC(TEMP_MATRIX, i, 1)) 'Optimal Portfolio StDev
Next i

If OUTPUT = 0 Then
    PORT_MARKOWITZ_FRONTIER_FUNC = TEMP_VECTOR
Else
    PORT_MARKOWITZ_FRONTIER_FUNC = Array(TEMP_MATRIX, TEMP_VECTOR)
    'TEMP_MATRIX --> Optimal Weights
End Select

Exit Function
ERROR_LABEL:
PORT_MARKOWITZ_FRONTIER_FUNC = Err.number
End Function
