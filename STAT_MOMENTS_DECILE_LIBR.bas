Attribute VB_Name = "STAT_MOMENTS_DECILE_LIBR"


Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : DECILE_MATRIX_FUNC

'DESCRIPTION   : Algorithm to plot the decile matrix of two return series.

'LIBRARY       : STATISTICS
'GROUP         : DECILE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 12/08/2008
'************************************************************************************
'************************************************************************************

Function DECILE_MATRIX_FUNC(ByRef DATA1_RNG As Variant, _
ByRef DATA2_RNG As Variant, _
Optional ByVal NBINS As Long = 10, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0, _
Optional ByVal VERSION As Integer = 0)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim NROWS As Long

Dim MIN1_VAL As Double
Dim MAX1_VAL As Double

Dim MIN2_VAL As Double
Dim MAX2_VAL As Double

Dim DELTA1_VAL As Double
Dim DELTA2_VAL As Double

Dim DATA1_VECTOR As Variant
Dim DATA2_VECTOR As Variant

Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA1_VECTOR = DATA1_RNG
If UBound(DATA1_VECTOR, 1) = 1 Then: DATA1_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA1_VECTOR)

DATA2_VECTOR = DATA2_RNG
If UBound(DATA2_VECTOR, 1) = 1 Then: DATA2_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA2_VECTOR)
If UBound(DATA1_VECTOR, 1) <> UBound(DATA2_VECTOR, 1) Then: GoTo ERROR_LABEL

If DATA_TYPE <> 0 Then: DATA1_VECTOR = MATRIX_PERCENT_FUNC(DATA1_VECTOR, LOG_SCALE)
If DATA_TYPE <> 0 Then: DATA2_VECTOR = MATRIX_PERCENT_FUNC(DATA2_VECTOR, LOG_SCALE)

NROWS = UBound(DATA1_VECTOR, 1)

MIN1_VAL = 2 ^ 52: MAX1_VAL = -2 ^ 52
MIN2_VAL = MIN1_VAL: MAX2_VAL = MAX1_VAL
For h = 1 To NROWS
    If DATA1_VECTOR(h, 1) < MIN1_VAL Then: MIN1_VAL = DATA1_VECTOR(h, 1)
    If DATA1_VECTOR(h, 1) > MAX1_VAL Then: MAX1_VAL = DATA1_VECTOR(h, 1)
    
    If DATA2_VECTOR(h, 1) < MIN2_VAL Then: MIN2_VAL = DATA2_VECTOR(h, 1)
    If DATA2_VECTOR(h, 1) > MAX2_VAL Then: MAX2_VAL = DATA2_VECTOR(h, 1)
Next h

If VERSION = 0 Then
    ReDim TEMP_MATRIX(0 To NBINS, 0 To NBINS)

    TEMP_MATRIX(0, 0) = "DECILE_TABLE"
    For h = 1 To NBINS
        TEMP_MATRIX(0, h) = "A" & h
        TEMP_MATRIX(h, 0) = "A" & NBINS - h + 1
    Next h
Else
    ReDim TEMP_MATRIX(1 To NBINS, 1 To NBINS)
End If

For h = 1 To NROWS
    DELTA1_VAL = NBINS * (DATA1_VECTOR(h, 1) - MIN1_VAL) / (MAX1_VAL - MIN1_VAL)
    DELTA1_VAL = CEILING_FUNC(DELTA1_VAL, 1)
    If DELTA1_VAL < 1 Then: DELTA1_VAL = 1
    
    DELTA2_VAL = NBINS * (DATA2_VECTOR(h, 1) - MIN2_VAL) / (MAX2_VAL - MIN2_VAL)
    DELTA2_VAL = CEILING_FUNC(DELTA2_VAL, 1)
    If DELTA2_VAL < 1 Then: DELTA2_VAL = 1
    
    i = (NBINS + 1) - DELTA1_VAL: j = DELTA2_VAL
    k = (NBINS + 1) - DELTA1_VAL: l = DELTA2_VAL
    TEMP_MATRIX(i, j) = TEMP_MATRIX(k, l) + 1
Next h


DECILE_MATRIX_FUNC = TEMP_MATRIX

Exit Function
ERROR_LABEL:
DECILE_MATRIX_FUNC = Err.number
End Function
