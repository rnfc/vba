Attribute VB_Name = "FINAN_PORT_MOMENTS_COVAR_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_COVAR_FUNC
'DESCRIPTION   : Returns the portfolio covariance, using two weight vectors
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_COVAR
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PORT_COVAR_FUNC(ByRef DATA_RNG As Variant, _
ByRef WEIGHTS1_RNG As Variant, _
ByRef WEIGHTS2_RNG As Variant, _
Optional ByVal COUNT_BASIS As Double = 252, _
Optional ByVal TRANS_OPT As Integer = 0, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)
    
Dim TEMP_MATRIX As Variant
Dim COVAR_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim WEIGHTS1_VECTOR As Variant
Dim WEIGHTS2_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG

If TRANS_OPT <> 0 Then DATA_MATRIX = MATRIX_REVERSE_FUNC(DATA_MATRIX)
If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)

COVAR_MATRIX = MATRIX_COVARIANCE_FRAME1_FUNC(DATA_MATRIX, 0, 0)

WEIGHTS1_VECTOR = WEIGHTS1_RNG
WEIGHTS2_VECTOR = WEIGHTS2_RNG

If UBound(WEIGHTS1_VECTOR, 1) <> 1 Then WEIGHTS1_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS1_VECTOR)
If UBound(WEIGHTS2_VECTOR, 1) <> 1 Then WEIGHTS2_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS2_VECTOR)

TEMP_MATRIX = MMULT_FUNC(WEIGHTS1_VECTOR, COVAR_MATRIX)
TEMP_MATRIX = MMULT_FUNC(TEMP_MATRIX, MATRIX_TRANSPOSE_FUNC(WEIGHTS2_VECTOR))

PORT_COVAR_FUNC = TEMP_MATRIX(1, 1) * COUNT_BASIS

Exit Function
ERROR_LABEL:
PORT_COVAR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_COVAR_SIGMA_FUNC
'DESCRIPTION   : Computes covariance matrix from sigma and correlation
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_COVAR
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PORT_COVAR_SIGMA_FUNC(ByRef SIGMA_RNG As Variant, _
ByRef CORREL_RNG As Variant)

Dim i As Long
Dim j As Long
Dim NSIZE As Long

Dim CORREL_MATRIX As Variant
Dim SIGMA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

SIGMA_VECTOR = SIGMA_RNG
If UBound(SIGMA_VECTOR, 1) = 1 Then: SIGMA_VECTOR = MATRIX_TRANSPOSE_FUNC(SIGMA_VECTOR)

CORREL_MATRIX = CORREL_RNG
If UBound(CORREL_MATRIX, 1) <> UBound(CORREL_MATRIX, 2) Then: GoTo ERROR_LABEL
If UBound(CORREL_MATRIX, 2) <> UBound(SIGMA_VECTOR, 1) Then: GoTo ERROR_LABEL

NSIZE = UBound(CORREL_MATRIX, 1)

ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)

For i = 1 To NSIZE
    For j = 1 To NSIZE
        TEMP_MATRIX(i, j) = CORREL_MATRIX(i, j) * SIGMA_VECTOR(i, 1) * SIGMA_VECTOR(j, 1)
    Next j
Next i

PORT_COVAR_SIGMA_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_COVAR_SIGMA_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_SIGMA_COVAR_FUNC

'DESCRIPTION   : Compute standard deviation matrix from covariance matrix
' using scale factor sf

'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_COVAR
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************


Function PORT_SIGMA_COVAR_FUNC(ByRef COVAR_RNG As Variant)

Dim i As Long
Dim NSIZE As Long

Dim TEMP_VECTOR As Variant
Dim COVAR_MATRIX As Variant

On Error GoTo ERROR_LABEL

COVAR_MATRIX = COVAR_RNG
If UBound(COVAR_MATRIX, 1) <> UBound(COVAR_MATRIX, 2) Then: GoTo ERROR_LABEL

NSIZE = UBound(COVAR_MATRIX, 2)
ReDim TEMP_VECTOR(1 To NSIZE, 1 To 1)
For i = 1 To NSIZE
    TEMP_VECTOR(i, 1) = COVAR_MATRIX(i, i) ^ 0.5
Next i

PORT_SIGMA_COVAR_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
PORT_SIGMA_COVAR_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_WEIGHTED_SIGMA_COVAR_FUNC
'DESCRIPTION   : Computes standard deviation of portfolio weights from covariance
'matrix
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_COVAR
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PORT_WEIGHTED_SIGMA_COVAR_FUNC(ByRef COVAR_RNG As Variant, _
ByRef WEIGHTS_RNG As Variant)

Dim i As Long
Dim j As Long
Dim NROWS As Long

Dim TEMP_SUM As Double

Dim WEIGHTS_VECTOR As Variant
Dim COVAR_MATRIX As Variant

On Error GoTo ERROR_LABEL

WEIGHTS_VECTOR = WEIGHTS_RNG
If UBound(WEIGHTS_VECTOR, 1) = 1 Then: _
    WEIGHTS_VECTOR = MATRIX_TRANSPOSE_FUNC(WEIGHTS_VECTOR)
NROWS = UBound(WEIGHTS_VECTOR, 1)

COVAR_MATRIX = COVAR_RNG

TEMP_SUM = 0
For i = 1 To NROWS
    For j = 1 To NROWS
        TEMP_SUM = TEMP_SUM + WEIGHTS_VECTOR(i, 1) * WEIGHTS_VECTOR(j, 1) * COVAR_MATRIX(i, j)
    Next j
Next i

PORT_WEIGHTED_SIGMA_COVAR_FUNC = TEMP_SUM ^ 0.5

Exit Function
ERROR_LABEL:
PORT_WEIGHTED_SIGMA_COVAR_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PORT_CORREL_COVAR_FUNC
'DESCRIPTION   : Compute correlation coefficient matrix from covariance matrix
'LIBRARY       : PORTFOLIO
'GROUP         : MOMENTS_COVAR
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function PORT_CORREL_COVAR_FUNC(ByRef COVAR_RNG As Variant)

Dim i As Long
Dim j As Long
Dim NSIZE As Long

Dim TEMP_VAL As Double
Dim TEMP_MATRIX As Variant
Dim COVAR_MATRIX As Variant

On Error GoTo ERROR_LABEL

COVAR_MATRIX = COVAR_RNG
If UBound(COVAR_MATRIX, 1) <> UBound(COVAR_MATRIX, 2) Then: GoTo ERROR_LABEL

NSIZE = UBound(COVAR_MATRIX, 2)

ReDim TEMP_MATRIX(1 To NSIZE, 1 To NSIZE)

For i = 1 To NSIZE
    For j = 1 To NSIZE
        TEMP_VAL = COVAR_MATRIX(i, i) * COVAR_MATRIX(j, j)
        If TEMP_VAL > 0 Then
            TEMP_MATRIX(i, j) = COVAR_MATRIX(i, j) / (TEMP_VAL ^ 0.5)
        Else
            TEMP_MATRIX(i, j) = 0
        End If
    Next j
Next i

PORT_CORREL_COVAR_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
PORT_CORREL_COVAR_FUNC = Err.number
End Function
