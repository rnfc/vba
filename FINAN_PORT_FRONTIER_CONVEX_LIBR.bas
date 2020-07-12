Attribute VB_Name = "FINAN_PORT_FRONTIER_CONVEX_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_CONVEX_HULL_FRONTIER_FUNC

'DESCRIPTION   : Random Feasible Portfolios: A Monte Carlo approach to asset
'allocation. Simulates random feasible portfolios and draws the efficient
'frontier as a convex hull.

'LIBRARY       : PORTFOLIO
'GROUP         : FRONTIER_CONVEX
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_CONVEX_HULL_FRONTIER_FUNC(ByRef EXPECTED_RNG As Variant, _
ByRef COVAR_RNG As Variant, _
Optional ByRef LOWER_RNG As Variant, _
Optional ByRef UPPER_RNG As Variant, _
Optional ByVal BUDGET_VAL As Double = 1, _
Optional ByVal nLOOPS As Long = 10000, _
Optional ByVal RANDOM_TYPE As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim LOWER_SUM As Double
Dim UPPER_SUM As Double

Dim LOWER_CUMUL As Double
Dim UPPER_CUMUL As Double

Dim TEMP_SUM As Double

Dim TEMP_MATRIX As Variant
Dim TEMP_VECTOR As Variant

Dim COVAR_MATRIX As Variant

Dim LOWER_VECTOR As Variant
Dim UPPER_VECTOR As Variant

Dim WEIGHT_MATRIX As Variant
Dim EXPECTED_VECTOR As Variant

On Error GoTo ERROR_LABEL

EXPECTED_VECTOR = EXPECTED_RNG
If UBound(EXPECTED_VECTOR, 1) = 1 Then
    EXPECTED_VECTOR = MATRIX_TRANSPOSE_FUNC(EXPECTED_VECTOR)
End If

COVAR_MATRIX = COVAR_RNG
If UBound(COVAR_MATRIX, 1) <> UBound(COVAR_MATRIX, 2) Then: GoTo ERROR_LABEL

If UBound(COVAR_MATRIX, 1) <> UBound(EXPECTED_VECTOR, 1) Then: GoTo ERROR_LABEL

If IsArray(LOWER_RNG) = True Then
    LOWER_VECTOR = LOWER_RNG
    If UBound(LOWER_VECTOR, 1) = 1 Then: LOWER_VECTOR = MATRIX_TRANSPOSE_FUNC(LOWER_VECTOR)
    If UBound(COVAR_MATRIX, 1) <> UBound(LOWER_VECTOR, 1) Then: GoTo ERROR_LABEL
Else
    ReDim LOWER_VECTOR(1 To NSIZE, 1 To 1)
    For i = 1 To NSIZE
        LOWER_VECTOR(i, 1) = LOWER_RNG
    Next i
End If

If IsArray(UPPER_RNG) = True Then
    UPPER_VECTOR = UPPER_RNG
    If UBound(UPPER_VECTOR, 1) = 1 Then: UPPER_VECTOR = MATRIX_TRANSPOSE_FUNC(UPPER_VECTOR)
    If UBound(COVAR_MATRIX, 1) <> UBound(UPPER_VECTOR, 1) Then: GoTo ERROR_LABEL
Else
    ReDim UPPER_VECTOR(1 To NSIZE, 1 To 1)
    For i = 1 To NSIZE
        UPPER_VECTOR(i, 1) = UPPER_RNG
    Next i
End If

NSIZE = UBound(EXPECTED_VECTOR, 1)

LOWER_SUM = MATRIX_SUM_FUNC(LOWER_VECTOR, 0)(1, 1)
UPPER_SUM = MATRIX_SUM_FUNC(UPPER_VECTOR, 0)(1, 1)

ReDim WEIGHT_MATRIX(1 To NSIZE, 1 To nLOOPS)

For j = 1 To nLOOPS
    TEMP_VECTOR = VECTOR_RANDOM_INDEX_FUNC(NSIZE, RANDOM_TYPE)
    LOWER_CUMUL = LOWER_SUM
    UPPER_CUMUL = UPPER_SUM
    TEMP_SUM = 0
    For i = 1 To NSIZE
        k = TEMP_VECTOR(i, 1)
        LOWER_CUMUL = LOWER_CUMUL - LOWER_VECTOR(k, 1)
        UPPER_CUMUL = UPPER_CUMUL - UPPER_VECTOR(k, 1)
        MAX_VAL = MAXIMUM_FUNC(BUDGET_VAL - TEMP_SUM - UPPER_CUMUL, LOWER_VECTOR(k, 1))
        MIN_VAL = MINIMUM_FUNC(BUDGET_VAL - TEMP_SUM - LOWER_CUMUL, UPPER_VECTOR(k, 1))
        WEIGHT_MATRIX(k, j) = MAX_VAL + HALTON_SEQUENCE_FUNC(Int(1000 * PSEUDO_RANDOM_FUNC(RANDOM_TYPE))) * (MIN_VAL - MAX_VAL)
        TEMP_SUM = TEMP_SUM + WEIGHT_MATRIX(k, j)
    Next i
Next j

ReDim SIGMA_VECTOR(1 To nLOOPS, 1 To 1)
ReDim RETURN_VECTOR(1 To nLOOPS, 1 To 1)

For i = 1 To nLOOPS
    For j = 1 To NSIZE
        RETURN_VECTOR(i, 1) = RETURN_VECTOR(i, 1) + WEIGHT_MATRIX(j, i) * EXPECTED_VECTOR(j, 1)
    Next j
    For j = 1 To NSIZE
        For k = 1 To NSIZE
            SIGMA_VECTOR(i, 1) = SIGMA_VECTOR(i, 1) + WEIGHT_MATRIX(j, i) * WEIGHT_MATRIX(k, i) * COVAR_MATRIX(j, k)
        Next k
    Next j
    SIGMA_VECTOR(i, 1) = SIGMA_VECTOR(i, 1) ^ 0.5
Next i

TEMP_VECTOR = CONVEX_HULL_FUNC(RETURN_VECTOR, SIGMA_VECTOR)
j = UBound(TEMP_VECTOR, 1)

ReDim TEMP_MATRIX(0 To j, 1 To 3)
TEMP_MATRIX(0, 1) = "INDEX"
TEMP_MATRIX(0, 2) = "RETURN"
TEMP_MATRIX(0, 3) = "SIGMA"

For i = 1 To j
    k = TEMP_VECTOR(i, 1)
    TEMP_MATRIX(i, 1) = k
    TEMP_MATRIX(i, 2) = RETURN_VECTOR(k, 1)
    TEMP_MATRIX(i, 3) = SIGMA_VECTOR(k, 1)
Next i

Select Case OUTPUT
    Case 0
        PORT_CONVEX_HULL_FRONTIER_FUNC = TEMP_MATRIX
    Case 1
        PORT_CONVEX_HULL_FRONTIER_FUNC = RETURN_VECTOR
    Case 2
        PORT_CONVEX_HULL_FRONTIER_FUNC = SIGMA_VECTOR
    Case 3
        PORT_CONVEX_HULL_FRONTIER_FUNC = WEIGHT_MATRIX
    Case Else
        PORT_CONVEX_HULL_FRONTIER_FUNC = Array(TEMP_MATRIX, RETURN_VECTOR, SIGMA_VECTOR, WEIGHT_MATRIX)
End Select

Exit Function
ERROR_LABEL:
PORT_CONVEX_HULL_FRONTIER_FUNC = Err.number
End Function
