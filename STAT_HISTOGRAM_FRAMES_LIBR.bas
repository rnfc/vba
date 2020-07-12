Attribute VB_Name = "STAT_HISTOGRAM_FRAMES_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_DISCRETE_FUNC
'DESCRIPTION   : This function is designed to use the regular histogram routine
'to draw histograms based on scatter diagrams for discrete values
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_DISTRIB
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_DISCRETE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal SORT_OPT As Integer = 1, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim A_VAL As Double
Dim B_VAL As Double
Dim C_VAL As Double

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim TEMP_MIN As Double
Dim BIN_WIDTH As Double

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim DATA_VECTOR As Variant
Dim TEMP_MATRIX As Variant

Dim BIN_LEFT_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: _
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
NROWS = UBound(DATA_VECTOR, 1)

If (SORT_OPT = 1) Then: _
    DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
ReDim TEMP_MATRIX(1 To NROWS - 1, 1 To 1)

For i = 1 To NROWS - 1
    TEMP_MATRIX(i, 1) = DATA_VECTOR(i + 1, 1) - DATA_VECTOR(i, 1)
Next i

i = 2

A_VAL = TEMP_MATRIX(1, 1)
TEMP_MIN = A_VAL

Do While i < NROWS
    B_VAL = TEMP_MATRIX(i, 1)
    C_VAL = EUCLIDEAN_FUNC(A_VAL, B_VAL)
    A_VAL = C_VAL
    If TEMP_MATRIX(i, 1) < TEMP_MIN Then TEMP_MIN = TEMP_MATRIX(i, 1)
    i = i + 1
Loop


'-------------Make ATEMP the width and have histogram begin
'-------------at DATA_VECTOR(1,1) - ATEMP and end at
'-------------DATA_VECTOR(NROWS, 1) + ATEMP

If (A_VAL > 1) And (A_VAL < TEMP_MIN) Then
    BIN_WIDTH = A_VAL
    MIN_VAL = DATA_VECTOR(1, 1) - A_VAL
    MAX_VAL = DATA_VECTOR(NROWS, 1) + A_VAL
Else
    BIN_WIDTH = TEMP_MIN / 2
    MIN_VAL = DATA_VECTOR(1, 1) - BIN_WIDTH
    MAX_VAL = DATA_VECTOR(NROWS, 1) + BIN_WIDTH
End If

ReDim BIN_LEFT_VECTOR(1 To NROWS, 1 To 1)
    
For i = 1 To NROWS
    BIN_LEFT_VECTOR(i, 1) = DATA_VECTOR(i, 1) - BIN_WIDTH
Next i

ReDim XDATA_VECTOR(1 To 5 * NROWS, 1 To 1)
ReDim YDATA_VECTOR(1 To 5 * NROWS, 1 To 1)

For i = 1 To NROWS
    j = 5 * i - 4
        
    XDATA_VECTOR(j, 1) = BIN_LEFT_VECTOR(i, 1)
    YDATA_VECTOR(j, 1) = 0
        
    XDATA_VECTOR(j + 1, 1) = BIN_LEFT_VECTOR(i, 1)
    YDATA_VECTOR(j + 1, 1) = DATA_VECTOR(i, 2)
        
    XDATA_VECTOR(j + 2, 1) = BIN_LEFT_VECTOR(i, 1) + BIN_WIDTH
    YDATA_VECTOR(j + 2, 1) = DATA_VECTOR(i, 2)
        
    XDATA_VECTOR(j + 3, 1) = BIN_LEFT_VECTOR(i, 1) + 2 * BIN_WIDTH
    YDATA_VECTOR(j + 3, 1) = DATA_VECTOR(i, 2)
        
    XDATA_VECTOR(j + 4, 1) = BIN_LEFT_VECTOR(i, 1) + 2 * BIN_WIDTH
    YDATA_VECTOR(j + 4, 1) = 0
    
Next i
    
ReDim TEMP_MATRIX(1 To 5 * NROWS, 1 To 2)
    
For j = 1 To 5 * NROWS
    TEMP_MATRIX(j, 1) = XDATA_VECTOR(j, 1)
    TEMP_MATRIX(j, 2) = YDATA_VECTOR(j, 1)
Next j

Select Case OUTPUT
Case 0
    HISTOGRAM_DISCRETE_FUNC = TEMP_MATRIX  'Coordinates Histogram
Case 1
    HISTOGRAM_DISCRETE_FUNC = DATA_VECTOR
Case Else
    HISTOGRAM_DISCRETE_FUNC = BIN_LEFT_VECTOR
End Select

Exit Function
ERROR_LABEL:
HISTOGRAM_DISCRETE_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_FRAME1_FUNC
'DESCRIPTION   : HISTOGRAM FRAME FOR ONE DATASET
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_DISTRIB
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_FRAME1_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NBINS As Long = 0, _
Optional ByVal SORT_OPT As Integer = 1, _
Optional ByVal VERSION As Integer = 0, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim MIN_VAL As Double
Dim MAX_VAL As Double

Dim RNG_VAL As Double
Dim BIN_MIN As Double
Dim BIN_WIDTH As Double

Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim BIN_LEFT_VECTOR As Variant
Dim BIN_COUNT_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If SORT_OPT <> 0 Then: DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)

NROWS = UBound(DATA_VECTOR, 1) 'NROWS = NROWS or LOOPS or DRAWS or REPETITIONS


'---------------------DATA_VECTOR (DATA_RNG) MUST BE SORTED
MIN_VAL = DATA_VECTOR(1, 1)
MAX_VAL = DATA_VECTOR(NROWS, 1)
'--------------------------------------------------------------------------

BIN_WIDTH = HISTOGRAM_BIN_LIMITS_FUNC(MIN_VAL, MAX_VAL, NROWS, 0)

'--------------------------------------------------------------------------
Select Case VERSION
'--------------------------------------------------------------------------
Case 0
'--------------------------------------------------------------------------
    If NBINS = 0 Then
        If BIN_WIDTH = 0 Then
            NBINS = 1
            BIN_MIN = MAX_VAL
        Else
            BIN_MIN = HISTOGRAM_BIN_MIN_PRECISION_FUNC(MIN_VAL, BIN_WIDTH)
            RNG_VAL = MAX_VAL - BIN_MIN
            NBINS = Int(RNG_VAL / BIN_WIDTH) + 1
        End If
    Else
        If BIN_WIDTH = 0 Then
            NBINS = 1
            BIN_MIN = MAX_VAL
        Else
            BIN_MIN = HISTOGRAM_BIN_MIN_PRECISION_FUNC(MIN_VAL, BIN_WIDTH)
        End If
    End If
'--------------------------------------------------------------------------
Case Else
'--------------------------------------------------------------------------
    If NBINS = 0 Then
        If BIN_WIDTH = 0 Then
            NBINS = 1
            BIN_MIN = MAX_VAL
        Else
            BIN_MIN = HISTOGRAM_BIN_LIMITS_FUNC(MIN_VAL, MAX_VAL, NROWS, 1)
            NBINS = HISTOGRAM_BIN_LIMITS_FUNC(MIN_VAL, MAX_VAL, NROWS, 2)
        End If
    Else
        If BIN_WIDTH = 0 Then
            NBINS = 1
            BIN_MIN = MAX_VAL
        Else
            BIN_MIN = HISTOGRAM_BIN_LIMITS_FUNC(MIN_VAL, MAX_VAL, NROWS, 1)
        End If
    End If
'--------------------------------------------------------------------------
End Select
'--------------------------------------------------------------------------

ReDim BIN_LEFT_VECTOR(0 To NBINS, 1 To 1)
ReDim BIN_COUNT_VECTOR(1 To NBINS, 1 To 1)

ReDim XDATA_VECTOR(0 To 2 * NBINS + 2, 1 To 1)
ReDim YDATA_VECTOR(0 To 2 * NBINS + 2, 1 To 1)
ReDim TEMP_MATRIX(1 To NBINS * 2 + 2, 1 To 2)

BIN_LEFT_VECTOR(0, 1) = BIN_MIN
For i = 1 To NBINS
    BIN_LEFT_VECTOR(i, 1) = BIN_LEFT_VECTOR(0, 1) + BIN_WIDTH * i
Next i

For j = 1 To NROWS
    i = 1
    Do While BIN_LEFT_VECTOR(i, 1) < DATA_VECTOR(j, 1)
        If i = NBINS Then Exit Do
        i = i + 1
    Loop
    BIN_COUNT_VECTOR(i, 1) = BIN_COUNT_VECTOR(i, 1) + 1
Next j

XDATA_VECTOR(0, 1) = BIN_LEFT_VECTOR(0, 1)
YDATA_VECTOR(0, 1) = 0

For i = 1 To NBINS
    j = 2 * i - 1
    XDATA_VECTOR(j, 1) = BIN_LEFT_VECTOR(i - 1, 1)
    YDATA_VECTOR(j, 1) = BIN_COUNT_VECTOR(i, 1)
    XDATA_VECTOR(j + 1, 1) = BIN_LEFT_VECTOR(i, 1)
    YDATA_VECTOR(j + 1, 1) = BIN_COUNT_VECTOR(i, 1)
Next i
        
XDATA_VECTOR(NBINS * 2 + 1, 1) = BIN_LEFT_VECTOR(NBINS, 1)
YDATA_VECTOR(NBINS * 2 + 1, 1) = 0
For j = 0 To NBINS * 2 + 1
    TEMP_MATRIX(j + 1, 1) = XDATA_VECTOR(j, 1)
    TEMP_MATRIX(j + 1, 2) = YDATA_VECTOR(j, 1)
Next j

Select Case OUTPUT
    Case 0
        HISTOGRAM_FRAME1_FUNC = TEMP_MATRIX  'Coordinates Histogram
    Case 1
        HISTOGRAM_FRAME1_FUNC = BIN_LEFT_VECTOR
    Case Else
        HISTOGRAM_FRAME1_FUNC = BIN_COUNT_VECTOR
End Select

Exit Function
ERROR_LABEL:
HISTOGRAM_FRAME1_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_FRAME2_FUNC
'DESCRIPTION   : HISTOGRAM FRAME FOR TWO DATASETS
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_DISTRIB
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_FRAME2_FUNC(ByRef ADATA_RNG As Variant, _
ByRef BDATA_RNG As Variant, _
Optional ByVal BIN_WIDTH_A As Double = 0, _
Optional ByVal BIN_WIDTH_B As Double = 0, _
Optional ByVal SORT_OPT As Integer = 1, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS_A As Long
Dim NROWS_B As Long

Dim NBINS As Long
Dim NBINS_A As Long
Dim NBINS_B As Long

Dim RNG_VAL As Double

Dim MIN_VAL As Double
Dim MIN_VAL_A As Double
Dim MIN_VAL_B As Double

Dim MAX_VAL As Double
Dim MAX_VAL_A As Double
Dim MAX_VAL_B As Double

Dim BIN_MIN As Double
Dim BIN_MIN_A As Double
Dim BIN_MIN_B As Double

Dim BIN_WIDTH As Double
'Dim TEMP_BIN As double

Dim ADATA_VECTOR As Variant
Dim BDATA_VECTOR As Variant

Dim ATEMP_MATRIX As Variant
Dim BTEMP_MATRIX As Variant

'---------------------------------------------------------------------
Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim BIN_LEFT_VECTOR As Variant
Dim BIN_COUNT_VECTOR As Variant
'---------------------------------------------------------------------

On Error GoTo ERROR_LABEL

ADATA_VECTOR = ADATA_RNG
If UBound(ADATA_VECTOR, 1) = 1 Then: _
    ADATA_VECTOR = MATRIX_TRANSPOSE_FUNC(ADATA_VECTOR)

BDATA_VECTOR = BDATA_RNG
If UBound(BDATA_VECTOR, 1) = 1 Then: _
    BDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(BDATA_VECTOR)

NROWS_A = UBound(ADATA_VECTOR, 1)
NROWS_B = UBound(BDATA_VECTOR, 1)

If (SORT_OPT = 1) Then
    ADATA_VECTOR = MATRIX_QUICK_SORT_FUNC(ADATA_VECTOR, 1, 1)
    BDATA_VECTOR = MATRIX_QUICK_SORT_FUNC(BDATA_VECTOR, 1, 1)
End If
  
MIN_VAL_A = ADATA_VECTOR(1, 1)
MIN_VAL_B = BDATA_VECTOR(1, 1)

MAX_VAL_A = ADATA_VECTOR(NROWS_A, 1)
MAX_VAL_B = BDATA_VECTOR(NROWS_B, 1)

'--------------------------------------------------------------------------
MIN_VAL = MIN_VAL_A
    If MIN_VAL_B < MIN_VAL_A Then: MIN_VAL = MIN_VAL_B

MAX_VAL = MAX_VAL_A
    If MAX_VAL_B > MAX_VAL_A Then: MAX_VAL = MAX_VAL_B
'--------------------------------------------------------------------------

If (BIN_WIDTH_A = 0) Then
    BIN_WIDTH_A = HISTOGRAM_BIN_LIMITS_FUNC(MIN_VAL_A, MAX_VAL_A, NROWS_A, 0)
    'BIN_WIDTH_A = HISTOGRAM_BIN_WIDTH_PRECISION_FUNC(BIN_WIDTH_A)
Else
    BIN_WIDTH_A = HISTOGRAM_BIN_WIDTH_PRECISION_FUNC(BIN_WIDTH_A)
End If

If BIN_WIDTH_A = 0 Then
    BIN_WIDTH_A = 1
    BIN_MIN_A = MAX_VAL_A
Else
    BIN_MIN_A = HISTOGRAM_BIN_MIN_PRECISION_FUNC(MIN_VAL_A, BIN_WIDTH_A)
End If
'--------------------------------------------------------------------------

If (BIN_WIDTH_B = 0) Then
    BIN_WIDTH_B = HISTOGRAM_BIN_LIMITS_FUNC(MIN_VAL_B, MAX_VAL_B, NROWS_B, 0)
    'BIN_WIDTH_B = HISTOGRAM_BIN_WIDTH_PRECISION_FUNC(BIN_WIDTH_B)
Else
    BIN_WIDTH_B = HISTOGRAM_BIN_WIDTH_PRECISION_FUNC(BIN_WIDTH_B)
End If

If BIN_WIDTH_B = 0 Then
    BIN_WIDTH_B = 1
    BIN_MIN_B = MAX_VAL_B
Else
    BIN_MIN_B = HISTOGRAM_BIN_MIN_PRECISION_FUNC(MIN_VAL_B, BIN_WIDTH_B)
End If
'--------------------------------------------------------------------------
' --------------Choose the larger Width as the class interval Width----------------

If BIN_WIDTH_A > BIN_WIDTH_B Then
    BIN_WIDTH = BIN_WIDTH_A
    k = 1
Else
    BIN_WIDTH = BIN_WIDTH_B
    k = 2
End If

If MIN_VAL_A = MAX_VAL_A Then
    BIN_WIDTH = BIN_WIDTH_B
    k = 2
End If

If MIN_VAL_B = MAX_VAL_B Then
    BIN_WIDTH = BIN_WIDTH_A
    k = 1
End If

'----------------------This gives us parameters for the first histogram-------------

If BIN_MIN_A < BIN_MIN_B Then
    If k = 1 Then
        BIN_MIN = BIN_MIN_A
    Else
        NBINS = Int((BIN_MIN_B - BIN_MIN_A) / BIN_WIDTH)
        BIN_MIN = BIN_MIN_B - NBINS * BIN_WIDTH
        If BIN_MIN > BIN_MIN_A Then: BIN_MIN = BIN_MIN - BIN_WIDTH
    End If
Else
    If k = 2 Then
        BIN_MIN = BIN_MIN_B
    Else
        NBINS = Int((BIN_MIN_A - BIN_MIN_B) / BIN_WIDTH)
        BIN_MIN = BIN_MIN_A - NBINS * BIN_WIDTH
        If BIN_MIN > BIN_MIN_B Then: BIN_MIN = BIN_MIN - BIN_WIDTH
    End If
End If

If BIN_WIDTH = 0 Then
    If RNG_VAL > 0 Then
        BIN_WIDTH = (MAX_VAL - MIN_VAL) / 2
    Else
        BIN_WIDTH = 1
    End If
End If

RNG_VAL = MAX_VAL_A - BIN_MIN
NBINS_A = Int(RNG_VAL / BIN_WIDTH) + 1

ReDim BIN_LEFT_VECTOR(0 To NBINS_A, 1 To 1)
ReDim BIN_COUNT_VECTOR(1 To NBINS_A, 1 To 1)

ReDim XDATA_VECTOR(0 To 2 * NBINS_A + 2, 1 To 1)
ReDim YDATA_VECTOR(0 To 2 * NBINS_A + 2, 1 To 1)
ReDim ATEMP_MATRIX(1 To NBINS_A * 2 + 2, 1 To 2)

BIN_LEFT_VECTOR(0, 1) = BIN_MIN
For i = 1 To NBINS_A
    BIN_LEFT_VECTOR(i, 1) = BIN_LEFT_VECTOR(0, 1) + BIN_WIDTH * i
Next i

For j = 1 To NROWS_A
    i = 1
    Do While BIN_LEFT_VECTOR(i, 1) < ADATA_VECTOR(j, 1)
        If i = NBINS_A Then Exit Do
        i = i + 1
    Loop
    BIN_COUNT_VECTOR(i, 1) = BIN_COUNT_VECTOR(i, 1) + 1
Next j

XDATA_VECTOR(0, 1) = BIN_LEFT_VECTOR(0, 1)
YDATA_VECTOR(0, 1) = 0

For i = 1 To NBINS_A
    j = 2 * i - 1
    XDATA_VECTOR(j, 1) = BIN_LEFT_VECTOR(i - 1, 1)
    YDATA_VECTOR(j, 1) = BIN_COUNT_VECTOR(i, 1)
    XDATA_VECTOR(j + 1, 1) = BIN_LEFT_VECTOR(i, 1)
    YDATA_VECTOR(j + 1, 1) = BIN_COUNT_VECTOR(i, 1)
Next i
        
XDATA_VECTOR(NBINS_A * 2 + 1, 1) = BIN_LEFT_VECTOR(NBINS_A, 1)
YDATA_VECTOR(NBINS_A * 2 + 1, 1) = 0
For j = 0 To NBINS_A * 2 + 1
    ATEMP_MATRIX(j + 1, 1) = XDATA_VECTOR(j, 1)
    ATEMP_MATRIX(j + 1, 2) = YDATA_VECTOR(j, 1)
Next j

'TEMP_BIN = BIN_LEFT_VECTOR(NBINS_A, 1)

'----------------------This gives us parameters for the second histogram-------------

RNG_VAL = MAX_VAL_B - BIN_MIN
NBINS_B = Int(RNG_VAL / BIN_WIDTH) + 1

ReDim BIN_LEFT_VECTOR(0 To NBINS_B, 1 To 1)
ReDim BIN_COUNT_VECTOR(1 To NBINS_B, 1 To 1)

ReDim XDATA_VECTOR(0 To 2 * NBINS_B + 2, 1 To 1)
ReDim YDATA_VECTOR(0 To 2 * NBINS_B + 2, 1 To 1)
ReDim BTEMP_MATRIX(1 To NBINS_B * 2 + 2, 1 To 2)

BIN_LEFT_VECTOR(0, 1) = BIN_MIN
For i = 1 To NBINS_B
    BIN_LEFT_VECTOR(i, 1) = BIN_LEFT_VECTOR(0, 1) + BIN_WIDTH * i
Next i

For j = 1 To NROWS_B
    i = 1
    Do While BIN_LEFT_VECTOR(i, 1) < BDATA_VECTOR(j, 1)
        If i = NBINS_B Then Exit Do
        i = i + 1
    Loop
    BIN_COUNT_VECTOR(i, 1) = BIN_COUNT_VECTOR(i, 1) + 1
Next j

XDATA_VECTOR(0, 1) = BIN_LEFT_VECTOR(0, 1)
YDATA_VECTOR(0, 1) = 0

For i = 1 To NBINS_B
    j = 2 * i - 1
    XDATA_VECTOR(j, 1) = BIN_LEFT_VECTOR(i - 1, 1)
    YDATA_VECTOR(j, 1) = BIN_COUNT_VECTOR(i, 1)
    XDATA_VECTOR(j + 1, 1) = BIN_LEFT_VECTOR(i, 1)
    YDATA_VECTOR(j + 1, 1) = BIN_COUNT_VECTOR(i, 1)
Next i
        
XDATA_VECTOR(NBINS_B * 2 + 1, 1) = BIN_LEFT_VECTOR(NBINS_B, 1)
YDATA_VECTOR(NBINS_B * 2 + 1, 1) = 0
For j = 0 To NBINS_B * 2 + 1
    BTEMP_MATRIX(j + 1, 1) = XDATA_VECTOR(j, 1)
    BTEMP_MATRIX(j + 1, 2) = YDATA_VECTOR(j, 1) * (NROWS_A / NROWS_B) _
            '* (BIN_WIDTH_A / BIN_WIDTH_B)
Next j

'If BIN_LEFT_VECTOR(NBINS_B, 1) > TEMP_BIN Then: TEMP_BIN = BIN_LEFT_VECTOR(NBINS_B, 1)

Select Case OUTPUT
    Case 0
        HISTOGRAM_FRAME2_FUNC = BTEMP_MATRIX
    Case Else
        HISTOGRAM_FRAME2_FUNC = ATEMP_MATRIX
End Select

Exit Function
ERROR_LABEL:
HISTOGRAM_FRAME2_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_SCALED_FRAME1_FUNC
'DESCRIPTION   : This frame of the histogram scales the horizontal axis
'so that it's centered on the BIN_CENTER and has endpoints which
'are specified.
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_DISTRIB
'ID            : 004
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_SCALED_FRAME1_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NBINS As Long = 20, _
Optional ByVal BIN_CENTER As Double = 0, _
Optional ByVal MIN_VAL As Double = 0, _
Optional ByVal MAX_VAL As Double = 0, _
Optional ByVal SORT_OPT As Integer = 1, _
Optional ByVal WIDTH_OPT As Integer = 1, _
Optional ByVal OUTPUT As Integer = 0)

' If WIDTH_OPT = 0, then use BIN_CENTER
' If WIDTH_OPT = 1, then use MIN_VAL and MAX_VAL

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim RNG_VAL As Double

Dim MIN_COMP As Double
Dim MAX_COMP As Double
Dim BIN_WIDTH As Double

Dim TEMP_MIN As Double
Dim TEMP_MAX As Double
Dim TEMP_LARGE As Double

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim BIN_LEFT_VECTOR As Variant
Dim BIN_COUNT_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: _
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
NROWS = UBound(DATA_VECTOR, 1)

If (SORT_OPT = 1) Then
    DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
End If

ReDim BIN_LEFT_VECTOR(0 To NBINS, 1 To 1)
ReDim BIN_COUNT_VECTOR(1 To NBINS, 1 To 1)
    
ReDim XDATA_VECTOR(0 To 2 * NBINS + 2, 1 To 1)
ReDim YDATA_VECTOR(0 To 2 * NBINS + 2, 1 To 1)
ReDim TEMP_MATRIX(1 To NBINS * 2 + 2, 1 To 2)

Select Case WIDTH_OPT
Case 0
    RNG_VAL = MAX_VAL - MIN_VAL
    BIN_LEFT_VECTOR(0, 1) = MIN_VAL
Case Else

' Find the extremes and use them to assign bins, with BIN_CENTER value in
' BIN_CENTER of middle bin or on borderline of two middle bins
' Need a rounding routine to get the bin boundaries to be nice

    MIN_COMP = DATA_VECTOR(1, 1)
    MAX_COMP = DATA_VECTOR(1, 1)
        
    For i = 2 To NROWS
        If DATA_VECTOR(i, 1) < MIN_COMP Then
            MIN_COMP = DATA_VECTOR(i, 1)
        End If
        If DATA_VECTOR(i, 1) > MAX_COMP Then
            MAX_COMP = DATA_VECTOR(i, 1)
        End If
    Next i
        
    TEMP_MIN = BIN_CENTER - MIN_COMP
    TEMP_MAX = MAX_COMP - BIN_CENTER
        
    If TEMP_MIN > TEMP_MAX Then
        TEMP_LARGE = TEMP_MIN
    Else: TEMP_LARGE = TEMP_MAX
    End If
        
    RNG_VAL = 2 * TEMP_LARGE
    BIN_LEFT_VECTOR(0, 1) = BIN_CENTER - TEMP_LARGE
End Select

BIN_WIDTH = RNG_VAL / NBINS
For i = 1 To NBINS
    BIN_LEFT_VECTOR(i, 1) = BIN_LEFT_VECTOR(0, 1) + BIN_WIDTH * i
Next i

For j = 1 To NROWS
    i = 1
    Do While BIN_LEFT_VECTOR(i, 1) < DATA_VECTOR(j, 1)
        If i = NBINS Then Exit Do
        i = i + 1
    Loop
    BIN_COUNT_VECTOR(i, 1) = BIN_COUNT_VECTOR(i, 1) + 1
Next j

XDATA_VECTOR(0, 1) = BIN_LEFT_VECTOR(0, 1)
YDATA_VECTOR(0, 1) = 0

For i = 1 To NBINS
    j = 2 * i - 1
    XDATA_VECTOR(j, 1) = BIN_LEFT_VECTOR(i - 1, 1)
    YDATA_VECTOR(j, 1) = BIN_COUNT_VECTOR(i, 1)
    XDATA_VECTOR(j + 1, 1) = BIN_LEFT_VECTOR(i, 1)
    YDATA_VECTOR(j + 1, 1) = BIN_COUNT_VECTOR(i, 1)
Next i
            
XDATA_VECTOR(NBINS * 2 + 1, 1) = BIN_LEFT_VECTOR(NBINS, 1)
YDATA_VECTOR(NBINS * 2 + 1, 1) = 0
For j = 0 To NBINS * 2 + 1
    TEMP_MATRIX(j + 1, 1) = XDATA_VECTOR(j, 1)
    TEMP_MATRIX(j + 1, 2) = YDATA_VECTOR(j, 1)
Next j

Select Case OUTPUT
Case 0
    HISTOGRAM_SCALED_FRAME1_FUNC = TEMP_MATRIX  'Coordinates Histogram
Case 1
    HISTOGRAM_SCALED_FRAME1_FUNC = BIN_LEFT_VECTOR
Case Else
    HISTOGRAM_SCALED_FRAME1_FUNC = BIN_COUNT_VECTOR
End Select

Exit Function
ERROR_LABEL:
HISTOGRAM_SCALED_FRAME1_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_SCALED_FRAME2_FUNC
'DESCRIPTION   : This is a simple frame of the histogram scales.  In this
'function you can choose the percentile bins, so if NBINS is going
'to be 20 then you want 5 percentile bins.
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_DISTRIB
'ID            : 005
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_SCALED_FRAME2_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal NBINS As Long = 20, _
Optional ByVal SORT_OPT As Integer = 1, _
Optional ByVal OUTPUT As Integer = 0)

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim A_VAL As Double
Dim B_VAL As Double
Dim TEMP_WIDTH As Double

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim BIN_LEFT_VECTOR As Variant
Dim BIN_COUNT_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: _
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
NROWS = UBound(DATA_VECTOR, 1)

If (SORT_OPT = 1) Then
    DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
End If

ReDim BIN_LEFT_VECTOR(0 To NBINS, 1 To 1)
ReDim BIN_COUNT_VECTOR(1 To NBINS, 1 To 1)
    
ReDim XDATA_VECTOR(0 To 2 * NBINS + 2, 1 To 1)
ReDim YDATA_VECTOR(0 To 2 * NBINS + 2, 1 To 1)
ReDim TEMP_MATRIX(1 To NBINS * 2 + 2, 1 To 2)

B_VAL = 0
BIN_LEFT_VECTOR(0, 1) = DATA_VECTOR(1, 1)
        
For i = 1 To NBINS
    A_VAL = (i / NBINS) * NROWS
    BIN_LEFT_VECTOR(i, 1) = DATA_VECTOR(A_VAL, 1)
    BIN_COUNT_VECTOR(i, 1) = A_VAL - B_VAL
    B_VAL = A_VAL
Next i
        
XDATA_VECTOR(0, 1) = BIN_LEFT_VECTOR(0, 1)
YDATA_VECTOR(0, 1) = 0

For i = 1 To NBINS
    j = 2 * i - 1
    TEMP_WIDTH = BIN_LEFT_VECTOR(i, 1) - BIN_LEFT_VECTOR(i - 1, 1)
    If TEMP_WIDTH = 0 Then: TEMP_WIDTH = 1
    XDATA_VECTOR(j, 1) = BIN_LEFT_VECTOR(i - 1, 1)
    YDATA_VECTOR(j, 1) = BIN_COUNT_VECTOR(i, 1) / (TEMP_WIDTH)
    XDATA_VECTOR(j + 1, 1) = BIN_LEFT_VECTOR(i, 1)
    YDATA_VECTOR(j + 1, 1) = BIN_COUNT_VECTOR(i, 1) / (TEMP_WIDTH)
Next i
    
XDATA_VECTOR(NBINS * 2 + 1, 1) = BIN_LEFT_VECTOR(NBINS, 1)
YDATA_VECTOR(NBINS * 2 + 1, 1) = 0

For j = 0 To NBINS * 2 + 1
    TEMP_MATRIX(j + 1, 1) = XDATA_VECTOR(j, 1)
    TEMP_MATRIX(j + 1, 2) = YDATA_VECTOR(j, 1)
Next j

Select Case OUTPUT
Case 0
    HISTOGRAM_SCALED_FRAME2_FUNC = TEMP_MATRIX  'Coordinates Histogram
Case 1
    HISTOGRAM_SCALED_FRAME2_FUNC = BIN_LEFT_VECTOR
Case Else
    HISTOGRAM_SCALED_FRAME2_FUNC = BIN_COUNT_VECTOR
End Select

Exit Function
ERROR_LABEL:
HISTOGRAM_SCALED_FRAME2_FUNC = Err.number
End Function


'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_SCALED_FRAME3_FUNC
'DESCRIPTION   : This frame of the histogram scales for two datasets in the same
'histogram
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_DISTRIB
'ID            : 006
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_SCALED_FRAME3_FUNC(ByRef ADATA_RNG As Variant, _
ByRef BDATA_RNG As Variant, _
Optional ByVal NBINS As Long = 20, _
Optional ByVal BIN_CENTER As Double = 0, _
Optional ByVal MIN_VAL As Double = 0, _
Optional ByVal MAX_VAL As Double = 0, _
Optional ByVal SORT_OPT As Integer = 1, _
Optional ByVal WIDTH_OPT As Integer = 1, _
Optional ByVal OUTPUT As Integer = 0)

' If WIDTH_OPT = 0, then use BIN_CENTER
' If WIDTH_OPT = 1, then use MIN_VAL and MAX_VAL
' REF_VALUE = 4 * Approx Standard Error

Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim RNG_VAL As Double
Dim BIN_WIDTH As Double

Dim TEMP_MIN As Double
Dim TEMP_MAX As Double

Dim MIN_COMP As Double
Dim MAX_COMP As Double

Dim TEMP_LARGE As Double

Dim XDATA_VECTOR As Variant
Dim YDATA_VECTOR As Variant

Dim BIN_LEFT_VECTOR As Variant
Dim BIN_COUNT_VECTOR As Variant

Dim TEMP_MATRIX As Variant
Dim ADATA_VECTOR As Variant
Dim BDATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

ADATA_VECTOR = ADATA_RNG
If UBound(ADATA_VECTOR, 1) = 1 Then: _
    ADATA_VECTOR = MATRIX_TRANSPOSE_FUNC(ADATA_VECTOR)

BDATA_VECTOR = BDATA_RNG
If UBound(BDATA_VECTOR, 1) = 1 Then: _
    BDATA_VECTOR = MATRIX_TRANSPOSE_FUNC(BDATA_VECTOR)

NROWS = UBound(ADATA_VECTOR, 1)
If NROWS <> UBound(BDATA_VECTOR, 1) Then: GoTo ERROR_LABEL

If (SORT_OPT = 1) Then
    ADATA_VECTOR = MATRIX_QUICK_SORT_FUNC(ADATA_VECTOR, 1, 1)
    BDATA_VECTOR = MATRIX_QUICK_SORT_FUNC(BDATA_VECTOR, 1, 1)
End If
   
ReDim BIN_LEFT_VECTOR(0 To NBINS, 1 To 2)
ReDim BIN_COUNT_VECTOR(1 To NBINS, 1 To 2)
    
ReDim XDATA_VECTOR(0 To 2 * NBINS + 2, 1 To 2)
ReDim YDATA_VECTOR(0 To 2 * NBINS + 2, 1 To 2)
ReDim TEMP_MATRIX(1 To NBINS * 2 + 2, 1 To 4)

Select Case WIDTH_OPT
Case 0
    RNG_VAL = MAX_VAL - MIN_VAL
    BIN_LEFT_VECTOR(0, 1) = MIN_VAL
Case Else

' Find the extremes and use them to assign bins, with BIN_CENTER value in
' BIN_CENTER of middle bin or on borderline of two middle bins
' Need a rounding routine to get the bin boundaries to be nice

    MIN_COMP = ADATA_VECTOR(1, 1)
    MAX_COMP = ADATA_VECTOR(1, 1)
        
    For i = 2 To NROWS
        If ADATA_VECTOR(i, 1) < MIN_COMP Then
            MIN_COMP = ADATA_VECTOR(i, 1)
        End If
        If ADATA_VECTOR(i, 1) > MAX_COMP Then
            MAX_COMP = ADATA_VECTOR(i, 1)
        End If
    Next i
        
    For i = 1 To NROWS
        If BDATA_VECTOR(i, 1) < MIN_COMP Then
            MIN_COMP = BDATA_VECTOR(i, 1)
        End If
        If BDATA_VECTOR(i, 1) > MAX_COMP Then
            MAX_COMP = BDATA_VECTOR(i, 1)
        End If
    Next i
        
    TEMP_MIN = BIN_CENTER - MIN_COMP
    TEMP_MAX = MAX_COMP - BIN_CENTER
        
    If TEMP_MIN > TEMP_MAX Then
        TEMP_LARGE = TEMP_MIN
    Else: TEMP_LARGE = TEMP_MAX
    End If
        
    RNG_VAL = 2 * TEMP_LARGE
    BIN_LEFT_VECTOR(0, 1) = BIN_CENTER - TEMP_LARGE
End Select

'----------------------------SETTING FIRST_VECTOR----------------------------

BIN_WIDTH = RNG_VAL / NBINS
For i = 1 To NBINS
    BIN_LEFT_VECTOR(i, 1) = BIN_LEFT_VECTOR(0, 1) + BIN_WIDTH * i
Next i

For j = 1 To NROWS
    i = 1
    Do While BIN_LEFT_VECTOR(i, 1) < ADATA_VECTOR(j, 1)
        If i = NBINS Then Exit Do
        i = i + 1
    Loop
    BIN_COUNT_VECTOR(i, 1) = BIN_COUNT_VECTOR(i, 1) + 1
Next j

XDATA_VECTOR(0, 1) = BIN_LEFT_VECTOR(0, 1)
YDATA_VECTOR(0, 1) = 0

For i = 1 To NBINS
    j = 2 * i - 1
    XDATA_VECTOR(j, 1) = BIN_LEFT_VECTOR(i - 1, 1)
    YDATA_VECTOR(j, 1) = BIN_COUNT_VECTOR(i, 1)
    XDATA_VECTOR(j + 1, 1) = BIN_LEFT_VECTOR(i, 1)
    YDATA_VECTOR(j + 1, 1) = BIN_COUNT_VECTOR(i, 1)
Next i
            
XDATA_VECTOR(NBINS * 2 + 1, 1) = BIN_LEFT_VECTOR(NBINS, 1)
YDATA_VECTOR(NBINS * 2 + 1, 1) = 0
For j = 0 To NBINS * 2 + 1
    TEMP_MATRIX(j + 1, 1) = XDATA_VECTOR(j, 1)
    TEMP_MATRIX(j + 1, 2) = YDATA_VECTOR(j, 1)
Next j
'-----------------------------SETTING SECOND_VECTOR-------------------------
        
For i = 0 To NBINS
    BIN_LEFT_VECTOR(i, 2) = BIN_LEFT_VECTOR(i, 1)
Next i

For j = 1 To NROWS
    i = 1
    Do While BIN_LEFT_VECTOR(i, 2) < BDATA_VECTOR(j, 1)
        If i = NBINS Then Exit Do
        i = i + 1
    Loop
    BIN_COUNT_VECTOR(i, 2) = BIN_COUNT_VECTOR(i, 2) + 1
Next j

XDATA_VECTOR(0, 2) = BIN_LEFT_VECTOR(0, 2)
YDATA_VECTOR(0, 2) = 0

For i = 1 To NBINS
    j = 2 * i - 1
    XDATA_VECTOR(j, 2) = BIN_LEFT_VECTOR(i - 1, 2)
    YDATA_VECTOR(j, 2) = BIN_COUNT_VECTOR(i, 2)
    XDATA_VECTOR(j + 1, 2) = BIN_LEFT_VECTOR(i, 2)
    YDATA_VECTOR(j + 1, 2) = BIN_COUNT_VECTOR(i, 2)
Next i
            
XDATA_VECTOR(NBINS * 2 + 1, 2) = BIN_LEFT_VECTOR(NBINS, 2)
YDATA_VECTOR(NBINS * 2 + 1, 1) = 0
            
'-----------------------------SETTING RESULT VECTOR-------------------------
For j = 0 To NBINS * 2 + 1
    TEMP_MATRIX(j + 1, 1) = XDATA_VECTOR(j, 1)
    TEMP_MATRIX(j + 1, 2) = YDATA_VECTOR(j, 1)
    TEMP_MATRIX(j + 1, 3) = XDATA_VECTOR(j, 2)
    TEMP_MATRIX(j + 1, 4) = YDATA_VECTOR(j, 2)
Next j
'---------------------------------------------------------------------------

Select Case OUTPUT
Case 0
    HISTOGRAM_SCALED_FRAME3_FUNC = TEMP_MATRIX  'Coordinates Histogram
Case 1
    HISTOGRAM_SCALED_FRAME3_FUNC = BIN_LEFT_VECTOR
Case Else
    HISTOGRAM_SCALED_FRAME3_FUNC = BIN_COUNT_VECTOR
End Select

Exit Function
ERROR_LABEL:
HISTOGRAM_SCALED_FRAME3_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_PLOT_CHART_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_DISTRIB
'ID            : 011
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_PLOT_CHART_FUNC(ByRef CHART_OBJ As Excel.Chart, _
ByRef DATA_RNG As Excel.Range, _
ByVal BIN_MIN As Double, _
ByVal BIN_MAX As Double, _
Optional ByVal TITLE_NAME As String = "")

On Error GoTo ERROR_LABEL

HISTOGRAM_PLOT_CHART_FUNC = False
CHART_OBJ.ChartType = xlXYScatterLinesNoMarkers
CHART_OBJ.SetSourceData source:=DATA_RNG, PlotBy:=xlColumns

With CHART_OBJ.Axes(xlCategory)
    .HasMajorGridlines = False
    .HasMinorGridlines = False
End With

With CHART_OBJ.Axes(xlValue)
    .HasMajorGridlines = False
    .HasMinorGridlines = False
End With

CHART_OBJ.HasLegend = True

CHART_OBJ.PlotArea.Interior.ColorIndex = xlNone
CHART_OBJ.PlotArea.ClearFormats

With CHART_OBJ
    .HasTitle = True
    .ChartTitle.Characters.Text = "HISTOGRAM OF " & TITLE_NAME
End With

With CHART_OBJ.SeriesCollection(1).Border
    .ColorIndex = 7
    .WEIGHT = xlThin
    .LineStyle = xlContinuous
End With

With CHART_OBJ.SeriesCollection(2).Border
    .ColorIndex = 5
    .WEIGHT = xlThin
    .LineStyle = xlContinuous
End With

With CHART_OBJ.Axes(xlCategory)
    .MinimumScale = BIN_MIN
    .MaximumScale = BIN_MAX
    '.MajorUnit = BIN_WIDTH
End With

With CHART_OBJ.Axes(xlValue).Border
    .WEIGHT = xlHairline
    .LineStyle = xlNone
End With

With CHART_OBJ.Axes(xlValue)
    .MajorTickMark = xlNone
    .MinorTickMark = xlNone
    .TickLabelPosition = xlNone
End With

HISTOGRAM_PLOT_CHART_FUNC = True

Exit Function
ERROR_LABEL:
HISTOGRAM_PLOT_CHART_FUNC = False
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_CUMULATIVE_CHART_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_DISTRIB
'ID            : 012
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_CUMULATIVE_CHART_FUNC(ByRef CHART_OBJ As Excel.Chart, _
ByRef DATA_RNG As Excel.Range)

On Error GoTo ERROR_LABEL

HISTOGRAM_CUMULATIVE_CHART_FUNC = False

CHART_OBJ.ChartType = xlXYScatter
CHART_OBJ.SetSourceData source:=DATA_RNG, PlotBy:= _
    xlColumns

With CHART_OBJ
    .HasTitle = True
    .ChartTitle.Characters.Text = "CUMULATIVE_FREQUENCY"
    .Axes(xlCategory, xlPrimary).HasTitle = False
    .Axes(xlValue, xlPrimary).HasTitle = False
End With

CHART_OBJ.Axes(xlCategory, xlPrimary).CategoryType = xlAutomatic
With CHART_OBJ.Axes(xlCategory)
    .HasMajorGridlines = False
    .HasMinorGridlines = False
End With
With CHART_OBJ.Axes(xlValue)
    .HasMajorGridlines = False
    .HasMinorGridlines = False
End With

With CHART_OBJ.PlotArea.Border
    .ColorIndex = 16
    .WEIGHT = xlThin
    .LineStyle = xlContinuous
End With

CHART_OBJ.PlotArea.Interior.ColorIndex = xlNone
With CHART_OBJ.SeriesCollection(1)
    With .Border
        .ColorIndex = 3
        .WEIGHT = xlHairline
        .LineStyle = xlContinuous
    End With
        .MarkerBackgroundColorIndex = xlAutomatic
        .MarkerForegroundColorIndex = xlAutomatic
        .MarkerStyle = xlAutomatic
        .Smooth = False
        .MarkerSize = 3
        .Shadow = False
End With

HISTOGRAM_CUMULATIVE_CHART_FUNC = True

Exit Function
ERROR_LABEL:
HISTOGRAM_CUMULATIVE_CHART_FUNC = False
End Function
