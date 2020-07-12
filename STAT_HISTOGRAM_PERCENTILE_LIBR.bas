Attribute VB_Name = "STAT_HISTOGRAM_PERCENTILE_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_PERCENTILE_FUNC
'DESCRIPTION   : Returns the k-th percentile of values in a range. You can use
'this function to establish a threshold of acceptance
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_PERCENTILE
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_PERCENTILE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal CONFIDENCE_VAL As Double = 0.5, _
Optional ByVal SORT_OPT As Integer = 1)
    
Dim i As Long
Dim T_VAL As Double
Dim NROWS As Long
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL
    
DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)
If SORT_OPT <> 0 Then: DATA_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
    
If CONFIDENCE_VAL = 0# Then
    HISTOGRAM_PERCENTILE_FUNC = DATA_VECTOR(1, 1)
    Exit Function
End If
If CONFIDENCE_VAL = 1# Then
    HISTOGRAM_PERCENTILE_FUNC = DATA_VECTOR(NROWS, 1)
    Exit Function
End If

T_VAL = CONFIDENCE_VAL * (NROWS - 1#)
i = Int(T_VAL)
T_VAL = T_VAL - Int(T_VAL)
HISTOGRAM_PERCENTILE_FUNC = DATA_VECTOR(i + 1, 1) * (1# - T_VAL) + DATA_VECTOR(i + 2, 1) * T_VAL

Exit Function
ERROR_LABEL:
HISTOGRAM_PERCENTILE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_MEDIAN_FUNC
'DESCRIPTION   : Returns the median of the given numbers. The median is the
'number in the middle of a set of numbers
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_PERCENTILE
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_MEDIAN_FUNC(ByRef DATA_RNG As Variant)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long

Dim NROWS As Long

Dim TEMP1_VAL As Double
Dim TEMP2_VAL As Double
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
NROWS = UBound(DATA_VECTOR)

' Some degenerate cases

HISTOGRAM_MEDIAN_FUNC = 0#
If NROWS <= 0# Then
    Exit Function
End If
If NROWS = 1# Then
    HISTOGRAM_MEDIAN_FUNC = DATA_VECTOR(1, 1)
    Exit Function
End If
If NROWS = 2# Then
    HISTOGRAM_MEDIAN_FUNC = 0.5 * (DATA_VECTOR(1, 1) + DATA_VECTOR(2, 1))
    Exit Function
End If

' Common case, N>=3.
' Choose DATA_VECTOR[(NROWS-1)/2+1,1]

l = 0#
h = NROWS - 1#
k = (NROWS - 1#) \ 2#
Do While True
    If h <= l + 1# Then ' 1 or 2 elements in partition
        If h = l + 1# And DATA_VECTOR(h + 1, 1) < DATA_VECTOR(l + 1, 1) Then
            TEMP2_VAL = DATA_VECTOR(l + 1, 1)
            DATA_VECTOR(l + 1, 1) = DATA_VECTOR(h + 1, 1)
            DATA_VECTOR(h + 1, 1) = TEMP2_VAL
        End If
        Exit Do
    Else
        m = (l + h) \ 2#
        TEMP2_VAL = DATA_VECTOR(m + 1, 1)
        DATA_VECTOR(m + 1, 1) = DATA_VECTOR(l + 2, 1)
        
        DATA_VECTOR(l + 2, 1) = TEMP2_VAL
        If DATA_VECTOR(l + 1, 1) > DATA_VECTOR(h + 1, 1) Then
            TEMP2_VAL = DATA_VECTOR(l + 1, 1)
            DATA_VECTOR(l + 1, 1) = DATA_VECTOR(h + 1, 1)
            DATA_VECTOR(h + 1, 1) = TEMP2_VAL
        End If
        
        If DATA_VECTOR(l + 2, 1) > DATA_VECTOR(h + 1, 1) Then
            TEMP2_VAL = DATA_VECTOR(l + 2, 1)
            DATA_VECTOR(l + 2, 1) = DATA_VECTOR(h + 1, 1)
            DATA_VECTOR(h + 1, 1) = TEMP2_VAL
        End If
        If DATA_VECTOR(l + 1, 1) > DATA_VECTOR(l + 2, 1) Then
            TEMP2_VAL = DATA_VECTOR(l + 1, 1)
            DATA_VECTOR(l + 1, 1) = DATA_VECTOR(l + 2, 1)
            DATA_VECTOR(l + 2, 1) = TEMP2_VAL
        End If
        i = l + 1#
        j = h
        TEMP1_VAL = DATA_VECTOR(l + 2, 1)
        Do While True
            Do: i = i + 1#: Loop Until DATA_VECTOR(i + 1, 1) >= TEMP1_VAL
            Do: j = j - 1#: Loop Until DATA_VECTOR(j + 1, 1) <= TEMP1_VAL
            If j < i Then: Exit Do
            TEMP2_VAL = DATA_VECTOR(i + 1, 1)
            DATA_VECTOR(i + 1, 1) = DATA_VECTOR(j + 1, 1)
            DATA_VECTOR(j + 1, 1) = TEMP2_VAL
        Loop
        DATA_VECTOR(l + 2, 1) = DATA_VECTOR(j + 1, 1)
        DATA_VECTOR(j + 1, 1) = TEMP1_VAL
        If j >= k Then: h = j - 1#
        If j <= k Then: l = i
    End If
Loop

'
' If NROWS is odd, return result
'
If NROWS Mod 2# = 1# Then
    HISTOGRAM_MEDIAN_FUNC = DATA_VECTOR(k + 1, 1)
    Exit Function
End If
TEMP1_VAL = DATA_VECTOR(NROWS, 1)
For i = k + 1# To NROWS - 1# Step 1
    If DATA_VECTOR(i + 1, 1) < TEMP1_VAL Then
        TEMP1_VAL = DATA_VECTOR(i + 1, 1)
    End If
Next i
HISTOGRAM_MEDIAN_FUNC = 0.5 * (DATA_VECTOR(k + 1, 1) + TEMP1_VAL)

Exit Function
ERROR_LABEL:
HISTOGRAM_MEDIAN_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : HISTOGRAM_PERCENTILE_TABLE_FUNC
'DESCRIPTION   :
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_PERCENTILE
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 01/22/2009
'************************************************************************************
'************************************************************************************

Function HISTOGRAM_PERCENTILE_TABLE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByRef DATE_RNG As Variant, _
Optional ByRef REFERENCE_RNG As Variant)

Dim h As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim INDEX_ARR As Variant
Dim TEMP_MATRIX As Variant
Dim DATE_VECTOR As Variant
Dim DATA_VECTOR As Variant
Dim REFERENCE_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)

'--------------------------------------------------------------------------------------
If IsArray(REFERENCE_RNG) = True And IsArray(DATE_RNG) = True Then
'--------------------------------------------------------------------------------------
    DATE_VECTOR = DATE_RNG
    If UBound(DATE_VECTOR, 1) = 1 Then
        DATE_VECTOR = MATRIX_TRANSPOSE_FUNC(DATE_VECTOR)
    End If
    If NROWS <> UBound(DATE_VECTOR, 1) Then: GoTo ERROR_LABEL
    
    REFERENCE_VECTOR = REFERENCE_RNG
    If UBound(REFERENCE_VECTOR, 1) = 1 Then
        REFERENCE_VECTOR = MATRIX_TRANSPOSE_FUNC(REFERENCE_VECTOR)
    End If
    NCOLUMNS = UBound(REFERENCE_VECTOR, 1)
    REFERENCE_VECTOR = MATRIX_QUICK_SORT_FUNC(REFERENCE_VECTOR, 1, 1)
'--------------------------------------------------------------------------------------
    ReDim INDEX_ARR(1 To NCOLUMNS, 1 To 2)
    ReDim TEMP_MATRIX(1 To 7, 1 To NCOLUMNS)
'--------------------------------------------------------------------------------------
    TEMP_MATRIX(2, 1) = "MINIMUM"
    TEMP_MATRIX(3, 1) = "25TH PERCENTILE"
    TEMP_MATRIX(4, 1) = "50TH PERCENTILE"
    TEMP_MATRIX(5, 1) = "MEAN"
    TEMP_MATRIX(6, 1) = "75TH PERCENTILE"
    TEMP_MATRIX(7, 1) = "MAXIMUM"
'--------------------------------------------------------------------------------------
    TEMP_MATRIX(1, 1) = REFERENCE_VECTOR(1, 1)
'--------------------------------------------------------------------------------------
    k = 1
    For j = 2 To NCOLUMNS
        TEMP_MATRIX(1, j) = REFERENCE_VECTOR(j, 1)
        INDEX_ARR(j - 1, 1) = 1
        INDEX_ARR(j - 1, 2) = NROWS
        For i = k To NROWS
            If DATE_VECTOR(i, 1) = REFERENCE_VECTOR(j - 1, 1) Then
                INDEX_ARR(j - 1, 1) = i
            ElseIf DATE_VECTOR(i, 1) < REFERENCE_VECTOR(j, 1) Then
                INDEX_ARR(j - 1, 2) = i
            Else
                k = i
                Exit For
            End If
        Next i
        h = INDEX_ARR(j - 1, 1) 'start row
        k = INDEX_ARR(j - 1, 2) 'end row
        l = k - h + 1
        If l <= 0 Then: GoTo 1983
        ReDim TEMP_VECTOR(1 To l, 1 To 1)
        l = 1
        For i = h To k
            TEMP_VECTOR(l, 1) = DATA_VECTOR(i, 1)
            l = l + 1
        Next i
        i = 1: GoSub CALCS_LINE
1983:
    Next j
'--------------------------------------------------------------------------------------
Else
'--------------------------------------------------------------------------------------
    NCOLUMNS = UBound(DATA_VECTOR, 2)
    ReDim TEMP_MATRIX(1 To 7, 1 To NCOLUMNS + 1)
    j = 1
    TEMP_MATRIX(1, j) = "DATASET"
    
    TEMP_MATRIX(2, j) = "MINIMUM"
    TEMP_MATRIX(3, j) = "25TH PERCENTILE"
    TEMP_MATRIX(4, j) = "50TH PERCENTILE"
    TEMP_MATRIX(5, j) = "MEAN"
    TEMP_MATRIX(6, j) = "75TH PERCENTILE"
    TEMP_MATRIX(7, j) = "MAXIMUM"
    
    For k = 1 To NCOLUMNS
        i = 1: j = k + 1
        TEMP_MATRIX(1, j) = k
        TEMP_VECTOR = VECTOR_TRIM_FUNC(MATRIX_GET_COLUMN_FUNC(DATA_VECTOR, k, 1), "")
        GoSub CALCS_LINE
    Next k
'--------------------------------------------------------------------------------------
End If
'--------------------------------------------------------------------------------------

HISTOGRAM_PERCENTILE_TABLE_FUNC = TEMP_MATRIX

Exit Function
'--------------------------------------------------------------------------------------
CALCS_LINE:
'--------------------------------------------------------------------------------------
    TEMP_VECTOR = MATRIX_QUICK_SORT_FUNC(TEMP_VECTOR, 1, 1)
    
    TEMP_MATRIX(1 + i, j) = TEMP_VECTOR(1, 1)
    TEMP_MATRIX(2 + i, j) = HISTOGRAM_PERCENTILE_FUNC(TEMP_VECTOR, 0.25, 0)
    TEMP_MATRIX(3 + i, j) = HISTOGRAM_PERCENTILE_FUNC(TEMP_VECTOR, 0.5, 0)
    TEMP_MATRIX(4 + i, j) = MATRIX_MEAN_FUNC(TEMP_VECTOR)(1, 1)
    TEMP_MATRIX(5 + i, j) = HISTOGRAM_PERCENTILE_FUNC(TEMP_VECTOR, 0.75, 0)
    TEMP_MATRIX(6 + i, j) = TEMP_VECTOR(UBound(TEMP_VECTOR, 1), 1)
'--------------------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------------------
ERROR_LABEL:
HISTOGRAM_PERCENTILE_TABLE_FUNC = Err.number
End Function

