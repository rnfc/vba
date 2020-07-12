Attribute VB_Name = "STAT_HISTOGRAM_FREQUENCY_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.
                    

'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_FREQUENCY_FUNC

'DESCRIPTION   : Calculates how often values occur within a range of values,
'and then returns a vertical array of numbers. For example, use FREQUENCY
'to count the number of test scores that fall within ranges of scores
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_FREQUENCY
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************
Function HISTOGRAM_FREQUENCY_FUNC(ByRef DATA_RNG As Variant, _
ByVal NBINS As Long, _
ByVal BIN_MIN As Double, _
ByVal BIN_WIDTH As Double, _
Optional ByVal SORT_OPT As Integer = 1)

'Logic behind Frequency

'-21.3%: <-21.3%
'-14.6%: (-21.3%,-14.6%)
'-7.9%: (-14.6%,-7.9%)
'-1.2%: (-7.9%,-1.2%)
'5.5%: (-1.2%,5.5%)
'12.1%: (5.5%,12.1%)
'18.8%: (12.1%,18.8%)
'25.5%: (18.8%,25.5%)
'32.2%: (25.5%,32.2%)
'38.9%: (32.2%,38.9%)
'45.6%: (38.9%,45.6%)
'0.0%: >45.6%

Dim i As Long
Dim j As Long

Dim NROWS As Long
Dim DATA_VECTOR As Variant
Dim TEMP_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If SORT_OPT <> 0 Then: DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)

NROWS = UBound(DATA_VECTOR, 1)

ReDim TEMP_VECTOR(1 To NBINS + 1, 1 To 2)

TEMP_VECTOR(1, 1) = BIN_MIN
TEMP_VECTOR(1, 2) = 0
For i = 2 To NBINS + 1
    TEMP_VECTOR(i, 1) = TEMP_VECTOR(1, 1) + BIN_WIDTH * (i - 1)
    TEMP_VECTOR(i, 2) = 0
Next i

'-------------------------------------------------------------
For j = 1 To NROWS
    i = 1 '---> YOU CAN CHANGE THIS
    Do While TEMP_VECTOR(i, 1) < DATA_VECTOR(j, 1) '--> YOU CAN ALSO USE "<="
        If i = NBINS + 1 Then: Exit Do
        i = i + 1
    Loop
    TEMP_VECTOR(i, 2) = TEMP_VECTOR(i, 2) + 1
Next j
'-------------------------------------------------------------

HISTOGRAM_FREQUENCY_FUNC = TEMP_VECTOR

Exit Function
ERROR_LABEL:
HISTOGRAM_FREQUENCY_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_DYNAMIC_FREQUENCY_FUNC
'DESCRIPTION   : Function to generate automated histograms

'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_FREQUENCY
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************


Function HISTOGRAM_DYNAMIC_FREQUENCY_FUNC(ByRef DATA_RNG As Variant, _
ByVal NBINS As Long, _
Optional ByVal DATA_TYPE As Integer = 0, _
Optional ByVal LOG_SCALE As Integer = 0)
    
Dim i As Long
Dim j As Long

Dim NROWS As Long

Dim LENGTH_VAL As Double

Dim FREQUENCY_ARR As Variant
Dim BREAKS_ARR As Variant
Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then: DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
If DATA_TYPE <> 0 Then: DATA_VECTOR = MATRIX_PERCENT_FUNC(DATA_VECTOR, LOG_SCALE)
NROWS = UBound(DATA_VECTOR, 1)
DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)

ReDim BREAKS_ARR(1 To NBINS)
ReDim FREQUENCY_ARR(1 To NBINS)

'---------------------------------------------------------------------------------------------------------
LENGTH_VAL = (DATA_VECTOR(NROWS, 1) - DATA_VECTOR(1, 1)) / NBINS
For i = 1 To NBINS
    BREAKS_ARR(i) = DATA_VECTOR(1, 1) + LENGTH_VAL * i
Next i
'---------------------------------------------------------------------------------------------------------
For i = 1 To NROWS
    If (DATA_VECTOR(i, 1) <= BREAKS_ARR(1)) Then FREQUENCY_ARR(1) = FREQUENCY_ARR(1) + 1
    If (DATA_VECTOR(i, 1) >= BREAKS_ARR(NBINS - 1)) Then FREQUENCY_ARR(NBINS) = FREQUENCY_ARR(NBINS) + 1
    For j = 2 To NBINS - 1
        If (DATA_VECTOR(i, 1) > BREAKS_ARR(j - 1) And DATA_VECTOR(i, 1) <= BREAKS_ARR(j)) Then
            FREQUENCY_ARR(j) = FREQUENCY_ARR(j) + 1
        End If
    Next j
Next i
'---------------------------------------------------------------------------------------------------------
ReDim TEMP_VECTOR(1 To NBINS, 1 To 2)
For i = 1 To NBINS
    TEMP_VECTOR(i, 1) = BREAKS_ARR(i)
    TEMP_VECTOR(i, 2) = FREQUENCY_ARR(i)
Next i
'---------------------------------------------------------------------------------------------------------
HISTOGRAM_DYNAMIC_FREQUENCY_FUNC = TEMP_VECTOR
'---------------------------------------------------------------------------------------------------------

Exit Function
ERROR_LABEL:
HISTOGRAM_DYNAMIC_FREQUENCY_FUNC = Err.number
End Function
'---------------------------------------------------------------------------------------------------------

