Attribute VB_Name = "FINAN_PORT_RISK_STYLE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_DYNAMIC_FACTOR_FUNC

'DESCRIPTION   : This function uses "Flexible Least Squares" to estimate time-varying
'parameters in index models. The same technique can
'be used to estimate time-varying weights in Style Analysis.

'Model: y(t) = b(t) * x(t) + e(t)
'Estimation: Choose a(t) and b(t) such that the loss function
'(y(t) - b(t)*x(t))^2 + L*(b(t)-b(t-1))^2 is minimized for an L > 0

'Reference: Kalaba, R. and Tesfatsion, L., 1989. "Time Varying Linear Regression via
'Flexible Least Squares", Computers and Mathematics with Applications, Vol. 17, pp.
'1215-1245

'A non.technical version of the above paper can be downloaded here:
'http://ideas.repec.org/a/eee/dyncon/v12y1988i1p43-48.html

'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_INDEX
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 21/01/2009
'************************************************************************************
'************************************************************************************
'PORT_STYLE_ONE_FACTOR_FUNC

Function PORT_DYNAMIC_FACTOR_FUNC(ByRef ASSET_RNG As Variant, _
ByRef BENCH_RNG As Variant, _
ByRef BETA_RNG As Variant, _
Optional ByVal LAMBDA As Double = 40, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim NROWS As Long

Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant

Dim BETA_VECTOR As Variant
Dim ASSET_VECTOR As Variant
Dim BENCH_VECTOR As Variant

On Error GoTo ERROR_LABEL

BETA_VECTOR = BETA_RNG 'Dynamic Beta
If UBound(BETA_VECTOR, 1) = 1 Then: BETA_VECTOR = MATRIX_TRANSPOSE_FUNC(BETA_VECTOR)

ASSET_VECTOR = ASSET_RNG
If UBound(ASSET_VECTOR, 1) = 1 Then: ASSET_VECTOR = MATRIX_TRANSPOSE_FUNC(ASSET_VECTOR)

BENCH_VECTOR = BENCH_RNG
If UBound(BENCH_VECTOR, 1) = 1 Then: BENCH_VECTOR = MATRIX_TRANSPOSE_FUNC(BENCH_VECTOR)

If UBound(BETA_VECTOR, 1) <> UBound(ASSET_VECTOR, 1) Then: GoTo ERROR_LABEL
If UBound(BETA_VECTOR, 1) <> UBound(BENCH_VECTOR, 1) Then: GoTo ERROR_LABEL

NROWS = UBound(BETA_VECTOR, 1)

ReDim TEMP_MATRIX(0 To NROWS, 1 To 7)

TEMP_MATRIX(0, 1) = "ASSET"
TEMP_MATRIX(0, 2) = "BENCHMARK"
TEMP_MATRIX(0, 3) = "DYNAMIC BETA"
TEMP_MATRIX(0, 4) = "CALCULATED FUND"
TEMP_MATRIX(0, 5) = "SQR ERROR"
TEMP_MATRIX(0, 6) = "DYNAMIC SQR ERROR"
TEMP_MATRIX(0, 7) = "SQR ERROR"

TEMP_MATRIX(1, 1) = ASSET_VECTOR(1, 1)
TEMP_MATRIX(1, 2) = BENCH_VECTOR(1, 1)
TEMP_MATRIX(1, 3) = BETA_VECTOR(1, 1)
TEMP_MATRIX(1, 4) = TEMP_MATRIX(1, 2) * TEMP_MATRIX(1, 3)
TEMP_MATRIX(1, 5) = (TEMP_MATRIX(1, 4) - TEMP_MATRIX(1, 1)) ^ 2
TEMP_MATRIX(1, 6) = 0
TEMP_MATRIX(1, 7) = TEMP_MATRIX(1, 5) + TEMP_MATRIX(1, 6) * LAMBDA

TEMP_SUM = TEMP_MATRIX(1, 7)

'----------------------------------------------------------------------------------
For i = 2 To NROWS
'----------------------------------------------------------------------------------

    TEMP_MATRIX(i, 1) = ASSET_VECTOR(i, 1)
    TEMP_MATRIX(i, 2) = BENCH_VECTOR(i, 1)
    TEMP_MATRIX(i, 3) = BETA_VECTOR(i, 1)
    TEMP_MATRIX(i, 4) = TEMP_MATRIX(i, 2) * TEMP_MATRIX(i, 3)
    TEMP_MATRIX(i, 5) = (TEMP_MATRIX(i, 4) - TEMP_MATRIX(i, 1)) ^ 2
    TEMP_MATRIX(i, 6) = (TEMP_MATRIX(i - 1, 3) - TEMP_MATRIX(i, 3)) ^ 2
    TEMP_MATRIX(i, 7) = TEMP_MATRIX(i, 5) + TEMP_MATRIX(i, 6) * LAMBDA
    
    TEMP_SUM = TEMP_SUM + TEMP_MATRIX(i, 7)

'----------------------------------------------------------------------------------
Next i
'----------------------------------------------------------------------------------

Select Case OUTPUT
    Case 0
        PORT_DYNAMIC_FACTOR_FUNC = TEMP_MATRIX
    Case Else 'Loss Function
        PORT_DYNAMIC_FACTOR_FUNC = TEMP_SUM
        'Using Excel' Solver to estimate a factor model with time-varying
        'factor exposures with the "flexible least squares approach".
        'MIN_FUNCTION --> TEMP_SUM
        'CHANGING CELLS --> BETA_RNG
        'CONSTRAINTS --> none
End Select

Exit Function
ERROR_LABEL:
PORT_DYNAMIC_FACTOR_FUNC = Err.number
End Function
