Attribute VB_Name = "STAT_DIST_FAT_TAILS_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

Private PUB_DX_VAL As Double
Private PUB_M_VAL As Double
Private PUB_S_VAL As Double
Private PUB_NSIZE As Long

'************************************************************************************
'************************************************************************************
'FUNCTION      : FIT_FAT_TAIL_DISTRIBUTION_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : FAT_TAILS
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function FIT_FAT_TAIL_DISTRIBUTION_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal K_VAL As Double = 1500, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal nLOOPS As Long = 1000, _
Optional ByVal tolerance As Double = 10 ^ -10, _
Optional ByVal epsilon As Double = 10 ^ -6, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long

Dim NROWS As Long
Dim NBINS As Long

Dim A_VAL As Double
Dim B_VAL As Double

Dim TEMP_VAL As Double
Dim TEMP_SUM As Double

Dim BIN_MIN As Double
Dim BIN_WIDTH As Double

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant
Dim DATA_VECTOR As Variant
Dim PARAM_VECTOR As Variant
Dim FREQUENCY_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
NROWS = UBound(DATA_VECTOR, 1)


FREQUENCY_VECTOR = DATA_BASIC_MOMENTS_FUNC(DATA_VECTOR, 0, 0, 0.05, 0)
PUB_M_VAL = FREQUENCY_VECTOR(1, 4)
PUB_S_VAL = FREQUENCY_VECTOR(1, 7)
FREQUENCY_VECTOR = HISTOGRAM_BIN_LIMITS_FUNC(FREQUENCY_VECTOR(1, 2), FREQUENCY_VECTOR(1, 3), NROWS, 3)

BIN_WIDTH = FREQUENCY_VECTOR(LBound(FREQUENCY_VECTOR))
PUB_DX_VAL = BIN_WIDTH
BIN_MIN = FREQUENCY_VECTOR(LBound(FREQUENCY_VECTOR) + 1)
NBINS = FREQUENCY_VECTOR(LBound(FREQUENCY_VECTOR) + 2)
FREQUENCY_VECTOR = HISTOGRAM_FREQUENCY_FUNC(DATA_VECTOR, NBINS, BIN_MIN, BIN_WIDTH, 1)
NBINS = UBound(FREQUENCY_VECTOR, 1)

PUB_NSIZE = 0
For i = 1 To NBINS
    PUB_NSIZE = PUB_NSIZE + FREQUENCY_VECTOR(i, 2)
Next i

'-----------------------------------------------------------------------------------------------
Select Case OUTPUT
'-----------------------------------------------------------------------------------------------
Case 0
'-----------------------------------------------------------------------------------------------
    ReDim PARAM_VECTOR(1 To 1, 1 To 1)
    PARAM_VECTOR(1, 1) = K_VAL
    
    ReDim XDATA_VECTOR(1 To NBINS, 1 To 1)
    ReDim YDATA_VECTOR(1 To NBINS, 1 To 1)
    For i = 1 To NBINS
        XDATA_VECTOR(i, 1) = FREQUENCY_VECTOR(i, 1)
        YDATA_VECTOR(i, 1) = FREQUENCY_VECTOR(i, 2)
    Next i
    FIT_FAT_TAIL_DISTRIBUTION_FUNC = _
    LEVENBERG_MARQUARDT_OPTIMIZATION2_FUNC(XDATA_VECTOR, YDATA_VECTOR, _
    PARAM_VECTOR, "FAT_TAIL_ERROR_FUNC", "", 0, nLOOPS, tolerance, epsilon)
'-----------------------------------------------------------------------------------------------
Case Else
'-----------------------------------------------------------------------------------------------
    ReDim DATA_VECTOR(0 To NBINS, 1 To 5)
    DATA_VECTOR(0, 1) = "BINS"
    DATA_VECTOR(0, 2) = "FREQ"
    DATA_VECTOR(0, 3) = "FAT PDF"
    DATA_VECTOR(0, 4) = "FAT CDF"
    
    B_VAL = 1 / K_VAL / PUB_S_VAL ^ 2
    A_VAL = -2 ^ 52
    TEMP_SUM = 0
    For i = 1 To NBINS
        TEMP_VAL = Exp(-B_VAL * Sqr(1 + (FREQUENCY_VECTOR(i, 1) - PUB_M_VAL) ^ 2 / PUB_S_VAL ^ 2)) * PUB_DX_VAL
        TEMP_SUM = TEMP_SUM + TEMP_VAL
        If TEMP_SUM > A_VAL Then: A_VAL = TEMP_SUM
    Next i
    A_VAL = 1 / A_VAL
    
    For i = 1 To NBINS
        DATA_VECTOR(i, 1) = FREQUENCY_VECTOR(i, 1)
        DATA_VECTOR(i, 2) = FREQUENCY_VECTOR(i, 2)
        
        DATA_VECTOR(i, 3) = A_VAL * Exp(-B_VAL * Sqr(1 + (DATA_VECTOR(i, 1) - _
                            PUB_M_VAL) ^ 2 / PUB_S_VAL ^ 2)) * PUB_DX_VAL
        If i > 1 Then
            DATA_VECTOR(i, 4) = DATA_VECTOR(i - 1, 4) + DATA_VECTOR(i, 3)
        Else
            DATA_VECTOR(i, 4) = DATA_VECTOR(i, 3)
        End If
        DATA_VECTOR(i, 5) = DATA_VECTOR(i, 3) * PUB_NSIZE
    Next i
    DATA_VECTOR(0, 5) = "FAT FREQ"
    FIT_FAT_TAIL_DISTRIBUTION_FUNC = DATA_VECTOR
'-----------------------------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
FIT_FAT_TAIL_DISTRIBUTION_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : FAT_TAIL_ERROR_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : FAT_TAILS
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Private Function FAT_TAIL_ERROR_FUNC(ByRef XDATA_RNG As Variant, _
ByRef PARAM_RNG As Variant)

Dim i As Long
Dim NBINS As Long

Dim A_VAL As Double
Dim B_VAL As Double
Dim K_VAL As Double

Dim TEMP_SUM As Double
Dim TEMP_VAL As Double

Dim PARAM_VECTOR As Variant
Dim XDATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

PARAM_VECTOR = PARAM_RNG
XDATA_VECTOR = XDATA_RNG
If UBound(XDATA_VECTOR, 1) = 1 Then: XDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(XDATA_VECTOR)

NBINS = UBound(XDATA_VECTOR, 1)

ReDim TEMP_MATRIX(1 To NBINS, 1 To 1)

K_VAL = PARAM_VECTOR(1, 1)
B_VAL = 1 / K_VAL / PUB_S_VAL ^ 2
A_VAL = -2 ^ 52
TEMP_SUM = 0
For i = 1 To NBINS
    TEMP_VAL = Exp(-B_VAL * Sqr(1 + (XDATA_VECTOR(i, 1) - PUB_M_VAL) ^ 2 / PUB_S_VAL ^ 2)) * PUB_DX_VAL
    TEMP_SUM = TEMP_SUM + TEMP_VAL
    If TEMP_SUM > A_VAL Then: A_VAL = TEMP_SUM
Next i
A_VAL = 1 / A_VAL

For i = 1 To NBINS 'p1*EXP(-p2*SQRT(1+(x-m)^2/s^2))*dx
    TEMP_MATRIX(i, 1) = A_VAL * Exp(-B_VAL * Sqr(1 + (XDATA_VECTOR(i, 1) - PUB_M_VAL) ^ 2 / _
                        PUB_S_VAL ^ 2)) * PUB_DX_VAL 'pdf
    TEMP_MATRIX(i, 1) = TEMP_MATRIX(i, 1) * PUB_NSIZE 'Number of Returns
Next i

FAT_TAIL_ERROR_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
FAT_TAIL_ERROR_FUNC = Err.number
End Function
