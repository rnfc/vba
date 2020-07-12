Attribute VB_Name = "FINAN_ASSET_PAIR_FRONTIER_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_COMBINATORIAL_FRONTIER_FUNC
'DESCRIPTION   : A combinatorics approach to asset allocation - Two Asset Portfolios.
'LIBRARY       : PORTFOLIO
'GROUP         : FRONTIER_COMBINATORIAL
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_COMBINATORIAL_FRONTIER_FUNC(ByVal NO_ALLOCS As Long, _
ByRef EXPECTED_RNG As Variant, _
ByRef COVAR_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 0)
'NO_ALLOCS - Number of Allocations
'COVAR_RNG - Variance-Covariance Matrix

Dim H1 As Long
Dim H2 As Long
Dim i As Long
Dim j As Long
Dim k As Long

Dim hh As Long
Dim ii As Long
Dim jj As Long
Dim kk As Long

Dim NROWS As Long
Dim NO_ASSETS As Long

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant

Dim EXPECTED_VECTOR As Variant
Dim COVARIANCES_MATRIX As Variant

On Error GoTo ERROR_LABEL

EXPECTED_VECTOR = EXPECTED_RNG
If UBound(EXPECTED_VECTOR, 1) = 1 Then
    EXPECTED_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPECTED_VECTOR)
End If
COVARIANCES_MATRIX = COVAR_RNG
If UBound(EXPECTED_VECTOR, 1) <> UBound(COVARIANCES_MATRIX, 2) Then: GoTo ERROR_LABEL
NO_ASSETS = UBound(EXPECTED_VECTOR, 1)

hh = 2: ii = COMBINATIONS_FUNC(NO_ASSETS, hh)
jj = NO_ALLOCS: kk = 1

ReDim TEMP_MATRIX(1 To ii * jj, 1 To NO_ASSETS)

TEMP_VECTOR = COMBINATIONS_ELEMENTS_FUNC(NO_ASSETS, hh)
' Variable number of allocations for pairs of two; works for all lb = 0 and all ub = 1, with sum of all
' weights equal to 1
For i = 1 To ii Step 1
    For j = 1 To jj
        H1 = TEMP_VECTOR(i, 1)
        TEMP_MATRIX(i + ii * (j - 1), H1) = kk / (jj + 1) * j
        H2 = TEMP_VECTOR(i, 2)
        TEMP_MATRIX(i + ii * (j - 1), H2) = kk - TEMP_MATRIX(i + ii * (j - 1), H1)
    Next j
Next i
NROWS = UBound(TEMP_MATRIX, 1)
ReDim TEMP_VECTOR(0 To NROWS, 1 To 2)
TEMP_VECTOR(0, 1) = "RETURNS"
TEMP_VECTOR(0, 2) = "STDEV"

For i = 1 To NROWS
    For j = 1 To NO_ASSETS
        TEMP_VECTOR(i, 1) = TEMP_VECTOR(i, 1) + TEMP_MATRIX(i, j) * EXPECTED_VECTOR(j, 1)
    Next j
    For j = 1 To NO_ASSETS
        For k = 1 To NO_ASSETS
            TEMP_VECTOR(i, 2) = TEMP_VECTOR(i, 2) + TEMP_MATRIX(i, j) * TEMP_MATRIX(i, k) * COVARIANCES_MATRIX(j, k)
        Next k
    Next j
    TEMP_VECTOR(i, 2) = TEMP_VECTOR(i, 2) ^ 0.5
Next i

Select Case OUTPUT
Case 0 'Returns/Sigma
    PORT_COMBINATORIAL_FRONTIER_FUNC = TEMP_VECTOR
Case 1 'Allocations - Two Asset Portfolios
    PORT_COMBINATORIAL_FRONTIER_FUNC = TEMP_MATRIX
Case Else
    PORT_COMBINATORIAL_FRONTIER_FUNC = Array(TEMP_VECTOR, TEMP_MATRIX)
End Select

Exit Function
ERROR_LABEL:
PORT_COMBINATORIAL_FRONTIER_FUNC = Err.number
End Function
