Attribute VB_Name = "STAT_HISTOGRAM_RANK_LIBR"

Option Explicit     'Requires that all variables to be declared explicitly.
Option Base 1       'The "Option Base" statement allows to specify 0 or 1 as the
                    'default first index of arrays.

'************************************************************************************
'************************************************************************************
'FUNCTION      : VECTOR_RANK_PERCENTILE_FUNC
'DESCRIPTION   : Returns the rank of a value in a data set as a percentage
'of the data set. This function can be used to evaluate the relative
'standing of a value within a data set.
'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_RANK
'ID            : 001
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 09/22/2010
'************************************************************************************
'************************************************************************************

Function VECTOR_RANK_PERCENTILE_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal OUTPUT As Integer = 1)

Dim i As Long
Dim j As Long
Dim k As Long
Dim NROWS As Long

Dim TEMP_VECTOR As Variant
Dim DATA_VECTOR As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
If UBound(DATA_VECTOR, 1) = 1 Then
    DATA_VECTOR = MATRIX_TRANSPOSE_FUNC(DATA_VECTOR)
End If
NROWS = UBound(DATA_VECTOR, 1)
ReDim Preserve DATA_VECTOR(1 To NROWS, 1 To 2)
For i = 1 To NROWS: DATA_VECTOR(i, 2) = i: Next i
TEMP_VECTOR = MATRIX_QUICK_SORT_FUNC(DATA_VECTOR, 1, 1)
ReDim Preserve TEMP_VECTOR(1 To NROWS, 1 To 3)
TEMP_VECTOR(1, 3) = 1
For i = 2 To NROWS
    If TEMP_VECTOR(i, 1) = TEMP_VECTOR(i - 1, 1) Then
        TEMP_VECTOR(i, 3) = TEMP_VECTOR(i - 1, 3) + 1
    Else
        TEMP_VECTOR(i, 3) = 1
    End If
Next i
For i = 1 To NROWS
    j = TEMP_VECTOR(i, 2)
    DATA_VECTOR(j, 1) = TEMP_VECTOR(i, 1)
    k = i - TEMP_VECTOR(i, 3)
    DATA_VECTOR(j, 2) = k / (NROWS - 1)
Next i
Erase TEMP_VECTOR

Select Case OUTPUT
Case 0
    VECTOR_RANK_PERCENTILE_FUNC = DATA_VECTOR
Case Else
    VECTOR_RANK_PERCENTILE_FUNC = MATRIX_GET_COLUMN_FUNC(DATA_VECTOR, 2, 1)
End Select

Exit Function
ERROR_LABEL:
VECTOR_RANK_PERCENTILE_FUNC = Err.number
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : PERCENTRANK_FUNC
'DESCRIPTION   : Calculate the PERCENTRANK_FUNC(array, x)

'If X matches one of the values in the array, this function is
'equivalent to the Excel formula =(RANK(x)-1)/(N-1) where N is
'the number of data points.

'If X does not match one of the values, then the PERCENTRANK_FUNC
'function interpolates.

'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_RANK
'ID            : 002
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 09/22/2010
'************************************************************************************
'************************************************************************************

Function PERCENTRANK_FUNC(ByRef DATA_RNG As Variant, _
ByVal X0_VAL As Double, _
Optional ByVal DIGITS_VAL As Integer = 3)

Dim i As Long
Dim j As Long
Dim k As Long

Dim NROWS As Long
Dim NCOLUMNS As Long

Dim X1_VAL As Variant
Dim X2_VAL As Variant

Dim Y1_VAL As Variant
Dim Y2_VAL As Variant
Dim Y0_VAL As Variant

Dim DATA_VECTOR As Variant
Dim DATA_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_VECTOR = DATA_RNG
NROWS = UBound(DATA_VECTOR, 1)
NCOLUMNS = UBound(DATA_VECTOR, 2)

If NCOLUMNS = 1 Then
    DATA_MATRIX = DATA_VECTOR
Else
    ReDim DATA_MATRIX(1 To NROWS * NCOLUMNS, 1 To 1)
    k = 1
    For i = 1 To NROWS
        For j = 1 To NCOLUMNS
            DATA_MATRIX(k, 1) = DATA_VECTOR(i, j)
            k = k + 1
        Next j
    Next i
    NROWS = NROWS * NCOLUMNS
End If

DATA_MATRIX = MATRIX_QUICK_SORT_FUNC(DATA_MATRIX, 1, 1)
ReDim Preserve DATA_MATRIX(1 To NROWS, 1 To 2)
i = 1: DATA_MATRIX(i, 2) = 1: GoSub CHECK_LINE
For i = 2 To NROWS
    If DATA_MATRIX(i, 1) = DATA_MATRIX(i - 1, 1) Then
        DATA_MATRIX(i, 2) = DATA_MATRIX(i - 1, 2) + 1
    Else
        DATA_MATRIX(i, 2) = 1
    End If
    GoSub CHECK_LINE
Next i
GoSub INTERP_LINE
1983:
If DIGITS_VAL > 0 Then
    PERCENTRANK_FUNC = ADOWN_DIG_FUNC(Y0_VAL, DIGITS_VAL)
Else
    PERCENTRANK_FUNC = Y0_VAL
End If
Exit Function
'--------------------------------------------------------------------------------
CHECK_LINE:
'--------------------------------------------------------------------------------
    If DATA_MATRIX(i, 1) = X0_VAL Then
        k = i - DATA_MATRIX(i, 2)
        Y0_VAL = k / (NROWS - 1)
        GoTo 1983
    End If
'--------------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------------
INTERP_LINE:
'--------------------------------------------------------------------------------
    For i = 1 To NROWS
        If X0_VAL < DATA_MATRIX(i, 1) Then
            j = i - 1
            Exit For
        ElseIf X0_VAL = DATA_MATRIX(i, 1) Then
            j = i
            Exit For
        End If
    Next i
    k = i - DATA_MATRIX(j, 2)
    Y1_VAL = k / (NROWS - 1)
    
    k = i - DATA_MATRIX(j + 1, 2)
    Y2_VAL = k / (NROWS - 1)
    
    X1_VAL = DATA_MATRIX(j, 1)
    X2_VAL = DATA_MATRIX(j + 1, 1)
    If X0_VAL <= X1_VAL Then
        Y0_VAL = Y1_VAL
    ElseIf X0_VAL >= X2_VAL Then
        Y0_VAL = Y2_VAL
    Else
        Y0_VAL = (X0_VAL - X1_VAL) / (X2_VAL - X1_VAL) * Y2_VAL + (X2_VAL - X0_VAL) / (X2_VAL - X1_VAL) * Y1_VAL
    End If
'--------------------------------------------------------------------------------
Return
'--------------------------------------------------------------------------------
ERROR_LABEL:
PERCENTRANK_FUNC = CVErr(xlErrNA)
End Function

'************************************************************************************
'************************************************************************************
'FUNCTION      : RANK_ARRAY_FUNC
'DESCRIPTION   : Returns the rank of a number in a list of numbers. The rank of
'a number is its size relative to other values in a list. (If you were to sort the
'list, the rank of the number would be its position.)

'Order: is a number specifying how to rank number.
'If order is 0 (zero) or omitted, Microsoft Excel ranks number as if ref were a list
'sorted in descending order. If order is any nonzero value, Microsoft Excel ranks
'number as if ref were a list sorted in ascending order.

'LIBRARY       : STATISTICS
'GROUP         : HISTOGRAM_RANK
'ID            : 003
'AUTHOR        : RAFAEL NICOLAS FERMIN COTA
'LAST UPDATE   : 09/22/2010
'************************************************************************************
'************************************************************************************

Function RANK_ARRAY_FUNC(ByRef DATA_RNG As Variant, _
Optional ByVal ORDER_TYPE As Integer = 0)

'DATA_RNG: An array of, or a reference to, a list of numbers. Nonnumeric values
'in ref are ignored.

'If ORDER_TYPE is 0 (zero) or omitted, the list is sorted in
'descending order.

'If ORDER_TYPE is any nonzero value, the list is sorted in
'ascending order.

Dim i As Long
Dim j As Long
Dim k As Long

Dim NSIZE As Long
Dim NROWS As Long
Dim NCOLUMNS As Long

Dim TEMP_VAL As Variant
Dim DATA_MATRIX As Variant
Dim TEMP_MATRIX As Variant

On Error GoTo ERROR_LABEL

DATA_MATRIX = DATA_RNG
If UBound(DATA_MATRIX, 1) = 1 Then
    DATA_MATRIX = MATRIX_TRANSPOSE_FUNC(DATA_MATRIX)
End If
NROWS = UBound(DATA_MATRIX, 1)
NCOLUMNS = UBound(DATA_MATRIX, 2)

If NCOLUMNS > 1 Then
    NSIZE = NROWS * NCOLUMNS
    ReDim TEMP_MATRIX(1 To NSIZE, 1 To 2)
    k = 1
    For j = 1 To NCOLUMNS
        For i = 1 To NROWS
            TEMP_MATRIX(k, 1) = DATA_MATRIX(i, j)
            TEMP_MATRIX(k, 2) = 0
            k = k + 1
        Next i
    Next j
    DATA_MATRIX = TEMP_MATRIX
    TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(DATA_MATRIX, 1, ORDER_TYPE)
Else
    NSIZE = NROWS
    TEMP_MATRIX = MATRIX_QUICK_SORT_FUNC(DATA_MATRIX, 1, ORDER_TYPE)
    ReDim Preserve TEMP_MATRIX(1 To NSIZE, 1 To 2)
End If

i = 1: j = 1
TEMP_MATRIX(i, 2) = j
GoSub LOOK_LINE
For i = 2 To NSIZE
    j = j + 1
    If TEMP_MATRIX(i - 1, 1) <> TEMP_MATRIX(i, 1) Then
        TEMP_MATRIX(i, 2) = j
    Else
        TEMP_MATRIX(i, 2) = TEMP_MATRIX(i - 1, 2)
    End If
'This routine gives duplicate numbers the same rank. However, the presence of duplicate numbers affects the
'ranks of subsequent numbers. For example, in a list of integers sorted in ascending order, if the
'number 10 appears twice and has a rank of 5, then 11 would have a rank of 7 (no number would have
'a rank of 6).
    GoSub LOOK_LINE
Next i

'For some purposes one might want to use a definition of rank that takes ties into account. In the
'previous example, one would want a revised rank of 5.5 for the number 10. This can be done by adding
'the following correction factor to the value returned by RANK. This correction factor is appropriate
'both for the case where rank is computed in descending order (order = 0 or omitted) or ascending order
'(order = nonzero value).

'Correction factor for tied ranks=[COUNT(ref) + 1 – RANK(number, ref, 0) – RANK(number, ref, 1)]/2.
'In the following example, RANK(A2,A1:A5,1) equals 3. The correction factor is (5 + 1 – 2 – 3)/2 = 0.5
'and the revised rank that takes ties into account is 3 + 0.5 = 3.5. If number occurs only once in ref,
'the correction factor will be 0, since RANK would not have to be adjusted for a tie.

If NCOLUMNS > 1 Then
    ReDim TEMP_MATRIX(1 To NROWS, 1 To NCOLUMNS)
    k = 1
    For j = 1 To NCOLUMNS
        For i = 1 To NROWS
            TEMP_MATRIX(i, j) = DATA_MATRIX(k, 1)
            k = k + 1
        Next i
    Next j
    RANK_ARRAY_FUNC = TEMP_MATRIX
Else
    RANK_ARRAY_FUNC = DATA_MATRIX
End If

'-----------------------------------------------------------------------------------------
Exit Function
'-----------------------------------------------------------------------------------------
LOOK_LINE:
'-----------------------------------------------------------------------------------------
    TEMP_VAL = TEMP_MATRIX(i, 1)
    For k = 1 To NSIZE
        If TEMP_VAL = DATA_MATRIX(k, 1) Then
             DATA_MATRIX(k, 1) = TEMP_MATRIX(i, 2)
             Exit For
        End If
    Next k
'-----------------------------------------------------------------------------------------
Return
'-----------------------------------------------------------------------------------------
ERROR_LABEL:
RANK_ARRAY_FUNC = Err.number
End Function


