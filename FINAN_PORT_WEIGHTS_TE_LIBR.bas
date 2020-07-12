Attribute VB_Name = "FINAN_PORT_WEIGHTS_TE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_TRACKING_ERROR_SIMULATION_FUNC

'DESCRIPTION   : Given a strategic benchmark and a feasible
'active weight budget, the expected distribution of tracking
'errors is calculated.

'LIBRARY       : PORTFOLIO
'GROUP         : SIMULATION_TE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************

Function PORT_TRACKING_ERROR_SIMULATION_FUNC(ByRef COVAR_RNG As Variant, _
ByVal BENCH_RNG As Variant, _
ByRef LOWER_RNG As Variant, _
ByRef UPPER_RNG As Variant, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal RANDOM_TYPE As Integer = 0, _
Optional ByVal NBINS As Long = 50, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim NO_ALLOCS As Long
Dim BUDGET_VAL As Double ' Total Exposure

Dim MIN_VAL As Double
Dim MAX_VAL As Double
Dim TEMP_VAL As Double
Dim TEMP_SUM As Double

Dim LOWER_SUM As Double
Dim LOWER_CUMUL As Double

Dim UPPER_SUM As Double
Dim UPPER_CUMUL As Double

Dim LOWER_VECTOR As Variant
Dim UPPER_VECTOR As Variant
Dim BENCH_VECTOR As Variant
Dim COVAR_MATRIX As Variant

Dim TEMP1_MATRIX As Variant
Dim TEMP2_MATRIX As Variant
Dim WEIGHT_MATRIX As Variant

On Error GoTo ERROR_LABEL

BENCH_VECTOR = BENCH_RNG
If UBound(BENCH_VECTOR, 1) = 1 Then
    BENCH_VECTOR = MATRIX_TRANSPOSE_FUNC(BENCH_VECTOR)
End If
COVAR_MATRIX = COVAR_RNG
If UBound(COVAR_MATRIX, 1) <> UBound(COVAR_MATRIX, 2) Then: GoTo ERROR_LABEL
If UBound(COVAR_MATRIX, 1) <> UBound(BENCH_VECTOR, 1) Then: GoTo ERROR_LABEL

LOWER_VECTOR = LOWER_RNG
If UBound(LOWER_VECTOR, 1) = 1 Then: LOWER_VECTOR = MATRIX_TRANSPOSE_FUNC(LOWER_VECTOR)
If UBound(COVAR_MATRIX, 1) <> UBound(LOWER_VECTOR, 1) Then: GoTo ERROR_LABEL

UPPER_VECTOR = UPPER_RNG
If UBound(UPPER_VECTOR, 1) = 1 Then: UPPER_VECTOR = MATRIX_TRANSPOSE_FUNC(UPPER_VECTOR)
If UBound(COVAR_MATRIX, 1) <> UBound(UPPER_VECTOR, 1) Then: GoTo ERROR_LABEL

NO_ALLOCS = UBound(BENCH_VECTOR, 1)
BUDGET_VAL = MATRIX_SUM_FUNC(BENCH_VECTOR, 0)(1, 1)

NSIZE = UBound(LOWER_VECTOR, 1)
LOWER_SUM = MATRIX_SUM_FUNC(LOWER_VECTOR, 0)(1, 1)
UPPER_SUM = MATRIX_SUM_FUNC(UPPER_VECTOR, 0)(1, 1)

ReDim WEIGHT_MATRIX(1 To NSIZE, 1 To nLOOPS)
ReDim TEMP1_MATRIX(1 To NSIZE, 1)
For i = 1 To NSIZE: TEMP1_MATRIX(i, 1) = i: Next i

For j = 1 To nLOOPS
    TEMP2_MATRIX = TEMP1_MATRIX
    For i = 1 To NSIZE
        k = Int(NSIZE * PSEUDO_RANDOM_FUNC(RANDOM_TYPE) + 1)
        TEMP_VAL = TEMP2_MATRIX(i, 1)
        TEMP2_MATRIX(i, 1) = TEMP2_MATRIX(k, 1)
        TEMP2_MATRIX(k, 1) = TEMP_VAL
    Next i
    
    LOWER_CUMUL = LOWER_SUM
    UPPER_CUMUL = UPPER_SUM
    
    TEMP_SUM = 0
    For i = 1 To NSIZE
        k = TEMP2_MATRIX(i, 1)
        
        LOWER_CUMUL = LOWER_CUMUL - LOWER_VECTOR(k, 1)
        UPPER_CUMUL = UPPER_CUMUL - UPPER_VECTOR(k, 1)
        
        MAX_VAL = MAXIMUM_FUNC(BUDGET_VAL - TEMP_SUM - UPPER_CUMUL, LOWER_VECTOR(k, 1))
        MIN_VAL = MINIMUM_FUNC(BUDGET_VAL - TEMP_SUM - LOWER_CUMUL, UPPER_VECTOR(k, 1))
        
        WEIGHT_MATRIX(k, j) = MAX_VAL + PSEUDO_RANDOM_FUNC(RANDOM_TYPE) * (MIN_VAL - MAX_VAL)
        TEMP_SUM = TEMP_SUM + WEIGHT_MATRIX(k, j)
    Next i
Next j

ReDim TEMP2_MATRIX(1 To NO_ALLOCS, 1 To nLOOPS)
For i = 1 To NO_ALLOCS
    For j = 1 To nLOOPS
        TEMP2_MATRIX(i, j) = WEIGHT_MATRIX(i, j) - BENCH_VECTOR(i, 1)
    Next j
Next i

ReDim TEMP1_MATRIX(1 To nLOOPS, 1 To 1)
For i = 1 To nLOOPS
    For j = 1 To NO_ALLOCS
        For k = 1 To NO_ALLOCS
            TEMP1_MATRIX(i, 1) = TEMP1_MATRIX(i, 1) + TEMP2_MATRIX(j, i) * TEMP2_MATRIX(k, i) * COVAR_MATRIX(j, k)
        Next k
    Next j
Next i
Erase TEMP2_MATRIX
TEMP1_MATRIX = MATRIX_QUICK_SORT_FUNC(TEMP1_MATRIX, 1, 1)

Select Case OUTPUT
Case 0
    PORT_TRACKING_ERROR_SIMULATION_FUNC = TEMP1_MATRIX
Case Else
    PORT_TRACKING_ERROR_SIMULATION_FUNC = HISTOGRAM_DYNAMIC_FREQUENCY_FUNC(TEMP1_MATRIX, NBINS, 0, 0)
End Select

Exit Function
ERROR_LABEL:
PORT_TRACKING_ERROR_SIMULATION_FUNC = Err.number
End Function
