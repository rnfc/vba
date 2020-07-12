Attribute VB_Name = "STAT_MOMENTS_COVAR_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_COVARIANCE_FRAME1_FUNC
'DESCRIPTION   : Compute covariance matrix for historical data (A)
'LIBRARY       : STATISTICS
'GROUP         : COVARIANCE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_COVARIANCE_FRAME1_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)
 
'DATA TYPE = 0 using SPOT VALUES
'DATA TYPE = 0 using % CHANGE VALUES
 
Dim j As Long
Dim i As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim MEAN1_VAL As Double
Dim MEAN2_VAL As Double

Dim TEMP1_SUM As Double
Dim TEMP2_SUM As Double
Dim TEMP3_SUM As Double

Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim COVAR_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim COVAR_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)

For i = 1 To NCOLUMNS
    For j = 1 To i
        ReDim TEMP_MATRIX(1 To NROWS, 1 To 2)
        For k = 1 To NROWS
            TEMP_MATRIX(k, 1) = DATA_MATRIX(k, i)
            TEMP_MATRIX(k, 2) = DATA_MATRIX(k, j)
        Next k
        'calculate the mean and remove it
        For k = 1 To NROWS
            TEMP1_SUM = TEMP1_SUM + TEMP_MATRIX(k, 1)
            TEMP2_SUM = TEMP2_SUM + TEMP_MATRIX(k, 2)
        Next k
        MEAN1_VAL = TEMP1_SUM / (k - 1)
        MEAN2_VAL = TEMP2_SUM / (k - 1)
        TEMP1_SUM = 0: TEMP2_SUM = 0
        For k = 1 To NROWS
            TEMP_MATRIX(k, 1) = TEMP_MATRIX(k, 1) - MEAN1_VAL
            TEMP_MATRIX(k, 2) = TEMP_MATRIX(k, 2) - MEAN2_VAL
        Next k
        For k = 1 To NROWS
            TEMP3_SUM = TEMP_MATRIX(k, 1) * TEMP_MATRIX(k, 2) + TEMP3_SUM
        Next k
        COVAR_MATRIX(i, j) = TEMP3_SUM / (k - 1)
        COVAR_MATRIX(j, i) = COVAR_MATRIX(i, j)
        TEMP3_SUM = 0
    Next j
Next i
    
MATRIX_COVARIANCE_FRAME1_FUNC = COVAR_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_COVARIANCE_FRAME1_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_COVARIANCE_FRAME2_FUNC
'DESCRIPTION   : Compute covariance matrix for historical data (B)
'LIBRARY       : STATISTICS
'GROUP         : COVARIANCE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_COVARIANCE_FRAME2_FUNC(ByRef DATA_RNG As Variant)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim TEMP_VECTOR As Variant
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

ReDim TEMP_VECTOR(1 To NCOLUMNS) 'average for each column
ReDim TEMP_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)
'compute the average
For j = 1 To NCOLUMNS
    TEMP_SUM = 0
    For i = 1 To NROWS
        TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, j)
    Next i
    TEMP_VECTOR(j) = TEMP_SUM / NROWS
Next j
'compute the cross covariance matrix

For i = 1 To NCOLUMNS
    For j = 1 To NCOLUMNS
        If j < i Then
            TEMP_MATRIX(i, j) = TEMP_MATRIX(j, i)
        Else
            TEMP_SUM = 0
            For k = 1 To NROWS
                TEMP_SUM = TEMP_SUM + (DATA_MATRIX(k, i) - TEMP_VECTOR(i)) * (DATA_MATRIX(k, j) - TEMP_VECTOR(j))
            Next k
            TEMP_MATRIX(i, j) = TEMP_SUM / NROWS
        End If
    Next j
Next i
MATRIX_COVARIANCE_FRAME2_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_COVARIANCE_FRAME2_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_COVARIANCE_FRAME3_FUNC
'DESCRIPTION   : Compute covariance matrix for historical data (c)
'LIBRARY       : STATISTICS
'GROUP         : COVARIANCE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_COVARIANCE_FRAME3_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim TEMP_MATRIX As Variant
Dim DATA_MATRIX As Variant
Dim RETURNS_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

RETURNS_VECTOR = MATRIX_AVERAGE_RETURNS_FUNC(DATA_MATRIX, 0, 0)

ReDim TEMP_MATRIX(1 To NCOLUMNS, 1 To NCOLUMNS)

For j = 1 To NCOLUMNS
    For k = 1 To j
        TEMP_SUM = 0
        l = 0
        For i = 1 To NROWS
            If IsNumeric(DATA_MATRIX(i, j)) And Not IsEmpty(DATA_MATRIX(i, j)) And IsNumeric(DATA_MATRIX(i, k)) And Not IsEmpty(DATA_MATRIX(i, k)) Then
                TEMP_SUM = TEMP_SUM + (DATA_MATRIX(i, j) - RETURNS_VECTOR(j, 1)) * (DATA_MATRIX(i, k) - RETURNS_VECTOR(k, 1))
                l = l + 1
            End If
        Next i
        If l > 0 Then
            TEMP_MATRIX(j, k) = TEMP_SUM / NROWS
        Else
            TEMP_MATRIX(j, k) = "N/A"
        End If
    Next k
Next j

For j = 1 To NCOLUMNS
    For k = j + 1 To NCOLUMNS
        TEMP_MATRIX(j, k) = TEMP_MATRIX(k, j)
    Next k
Next j

MATRIX_COVARIANCE_FRAME3_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_COVARIANCE_FRAME3_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_COVARIANCE_AVERAGE_PAIRWISE_FUNC
'DESCRIPTION   : Computes average pairwise covariance a matrix of returns
'LIBRARY       : STATISTICS
'GROUP         : COVARIANCE
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_COVARIANCE_AVERAGE_PAIRWISE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)

Dim i As Long
Dim j As Long
Dim NCOLUMNS As Long

Dim TEMP_SUM As Double
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If DATA_TYPE <> 0 Then DATA_MATRIX = MATRIX_PERCENT_FUNC(DATA_MATRIX, LOG_SCALE)
DATA_MATRIX = MATRIX_COVARIANCE_FRAME3_FUNC(DATA_MATRIX, 0, 0)

NCOLUMNS = UBound(DATA_MATRIX, 2)
TEMP_SUM = 0
For i = 1 To NCOLUMNS - 1
    For j = i + 1 To NCOLUMNS
        TEMP_SUM = TEMP_SUM + DATA_MATRIX(i, j)
    Next j
Next i
TEMP_SUM = TEMP_SUM * 2 / (NCOLUMNS * (NCOLUMNS - 1))
MATRIX_COVARIANCE_AVERAGE_PAIRWISE_FUNC = TEMP_SUM

Exit Function
ERROR_LABEL:
MATRIX_COVARIANCE_AVERAGE_PAIRWISE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : MATRIX_COVARIANCE_AVERAGE_PAIRWISE_FUNC
'DESCRIPTION   : Computes covariance matrix from volatility vector and
'correlation matrix
'LIBRARY       : STATISTICS
'GROUP         : COVARIANCE
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function MATRIX_COVARIANCE_VOLATILITY_CORRELATION_FUNC( _
ByRef VOLATILITY_RNG As Variant, _
ByRef CORRELATION_RNG As Variant)

Dim i As Long
Dim j As Long
Dim NSIZE As Long

Dim COVARIANCE_MATRIX As Variant
Dim VOLATILITY_VECTOR As Variant
Dim CORRELATION_MATRIX As Variant

On Error GoTo ERROR_LABEL

VOLATILITY_VECTOR = VOLATILITY_RNG
If UBound(VOLATILITY_VECTOR, 1) = 1 Then
    VOLATILITY_VECTOR = MATRIX_TRANSPOSE_FUNC(VOLATILITY_VECTOR)
End If
CORRELATION_MATRIX = CORRELATION_RNG
If UBound(CORRELATION_MATRIX, 1) <> UBound(CORRELATION_MATRIX, 2) Then: GoTo ERROR_LABEL

NSIZE = UBound(CORRELATION_MATRIX, 1)
ReDim COVARIANCE_MATRIX(1 To NSIZE, 1 To NSIZE)
For i = 1 To NSIZE
    For j = 1 To NSIZE
        COVARIANCE_MATRIX(i, j) = CORRELATION_MATRIX(i, j) * VOLATILITY_VECTOR(i, 1) * VOLATILITY_VECTOR(j, 1)
    Next j
Next i
MATRIX_COVARIANCE_VOLATILITY_CORRELATION_FUNC = COVARIANCE_MATRIX

Exit Function
ERROR_LABEL:
MATRIX_COVARIANCE_VOLATILITY_CORRELATION_FUNC = Err.number
End Function
